/**
 * calculerDebitsProfils_v3.js — PATCH BAIE_TRAVERSE_H RECOUPÉE + MONTANTS FILANTS + BONUS CLOSOIRS AUTO
 *
 * RÈGLE :
 * - Si BAIE_TRAVERSE_H non recoupée => calcul NOMENCLATURE (comme avant)
 * - Si BAIE_TRAVERSE_H recoupée => IGNORER NOMENCLATURE pour BAIE_TRAVERSE_H et calculer chaque segment AU CLAIR
 *   (1 ligne par segment, puisque tu as déjà 1 profilchassis par segment).
 *
 * AJOUTS :
 * - Pour les traverses BAIE segmentées : déduction des épaisseurs int sur les coupes internes (x=Lg, x=Lg+Lp)
 *   même si le montant est stocké en zone FIXE_G / PASSAGE / FIXE_D (recherche "AnyZone").
 * - CAS METIER : sur BAIE_TRAVERSE_H segment PASSAGE en montants filants :
 *     * on ne déduit pas les 1/2 épaisseurs des montants aux limites du passage,
 *     * on ajoute automatiquement le bonus = épaisseurs int des closoirs PASSAGE_CLOSOIR_G/D qui s'arrêtent sous la traverse.
 */

// ==== CHOICES VALIDÉS ====
const ROLE = {
  CADRE: 745350000,
  MONTANT: 745350001,
  TRAVERSE: 745350002,
  PARCLOSE: 745350003
};

const COTE = {
  G: 745350000,
  D: 745350001,
  B: 745350002,
  H: 745350003
};

// Tolérance matching mm
const TOL = 0.5;

/**
 * IMPORTANT : libellés exacts (FormattedValue) du Choice crbee_position, en MAJUSCULES.
 * Sert uniquement aux piles de rive gauche/droite.
 */
const RIVE_POSITIONS = {
  "BAIE": {
    G: ["BAIE_MONTANT_G", "BAIE_CLOSOIR_G"],
    D: ["BAIE_MONTANT_D", "BAIE_CLOSOIR_D"]
  },
  "PASSAGE": {
    G: ["PASSAGE_MONTANT_G", "PASSAGE_CLOSOIR_G"],
    D: ["PASSAGE_MONTANT_D", "PASSAGE_CLOSOIR_D"]
  },
  "FIXE_G": {
    G: ["FIXE_G_MONTANT_G", "FIXE_G_CLOSOIR_G"],
    D: ["FIXE_G_MONTANT_D", "FIXE_G_CLOSOIR_D"]
  },
  "FIXE_D": {
    G: ["FIXE_D_MONTANT_G", "FIXE_D_CLOSOIR_G"],
    D: ["FIXE_D_MONTANT_D", "FIXE_D_CLOSOIR_D"]
  }
};

// ---------------- Helpers ----------------
function getChoiceLabel(entity, fieldName) {
  return entity[`${fieldName}@OData.Community.Display.V1.FormattedValue`] || "";
}

function normLabel(s, fallback = "") {
  return (s || fallback).toUpperCase().trim();
}

// mm entiers (jamais de virgule)
function mm(x) {
  return Math.floor(Math.max(0, x));
}

async function updateByChunks(entityName, updates, chunkSize = 20) {
  for (let i = 0; i < updates.length; i += chunkSize) {
    const chunk = updates.slice(i, i + chunkSize);
    await Promise.all(chunk.map(u => Xrm.WebApi.updateRecord(entityName, u.id, u.data)));
  }
}

function largeurZone(contexte, zoneLabel) {
  switch (zoneLabel) {
    case "FIXE_G": return contexte.LARGEUR_FIXE_G || 0;
    case "FIXE_D": return contexte.LARGEUR_FIXE_D || 0;
    case "PASSAGE": return contexte.LARGEUR_PASSAGE || 0;
    default: return contexte.LARGEUR_BAIE || 0; // BAIE
  }
}

function approxEq(a, b, tol = 1) {
  return Math.abs((a ?? 0) - (b ?? 0)) <= tol;
}

/**
 * Cache produit -> crbee_largeurfaceparclose (face intérieure / parclose)
 */
function createProduitFaceCache() {
  const cache = {};
  return async function getFaceParcloseProduit(produitId) {
    if (!produitId) return 0;
    if (cache[produitId] !== undefined) return cache[produitId];
    const prod = await Xrm.WebApi.retrieveRecord("crbee_produit", produitId, "?$select=crbee_largeurfaceparclose");
    const w = prod.crbee_largeurfaceparclose || 0;
    cache[produitId] = w;
    return w;
  };
}

/**
 * Cache produit -> crbee_largeurfaceexterieur (face extérieure)
 */
function createProduitExtCache() {
  const cache = {};
  return async function getEpExtProduit(produitId) {
    if (!produitId) return 0;
    if (cache[produitId] !== undefined) return cache[produitId];
    const prod = await Xrm.WebApi.retrieveRecord("crbee_produit", produitId, "?$select=crbee_largeurfaceexterieur");
    const w = prod.crbee_largeurfaceexterieur || 0;
    cache[produitId] = w;
    return w;
  };
}

/**
 * Épaisseur du cadre horizontal bas/haut calculée par convention de libellé :
 * - tout profil CADRE dont crbee_position (FormattedValue) finit par "_B" est bas
 * - tout profil CADRE dont crbee_position finit par "_H" est haut
 * => somme en face parclose
 */
async function epCadreHorizontalDepuisSuffixes(cadreProfils, zoneLabel, suffix, getFaceParcloseProduit) {
  const z = normLabel(zoneLabel, "BAIE");
  let sum = 0;

  for (const p of cadreProfils) {
    const pZone = normLabel(getChoiceLabel(p, "crbee_zone"), "BAIE");
    if (pZone !== z) continue;

    const pos = normLabel(getChoiceLabel(p, "crbee_position"), "");
    if (!pos.endsWith(suffix)) continue;

    sum += await getFaceParcloseProduit(p._crbee_profil_value);
  }

  return sum;
}

// ---------------- 0) MONTANTS INTERMÉDIAIRES ----------------
async function calculerMontantsIntermediairesDepuisProfils(profils, contexte) {
  const getEpIntProduit = createProduitFaceCache(); // face parclose
  const getEpExtProduit = createProduitExtCache();  // face ext

  const montants = profils.filter(p => p.crbee_role === ROLE.MONTANT);
  if (!montants.length) return;

  const cadre = profils.filter(p => p.crbee_role === ROLE.CADRE);
  const traverses = profils.filter(p => p.crbee_role === ROLE.TRAVERSE);

  const Hbaie = contexte.HAUTEUR_BAIE || 0;

  // Cache cadre horizontal bas/haut (face parclose)
  const cacheCadreHor = new Map();
  async function getCadreHor(zoneLabel, suffix) {
    const z = normLabel(zoneLabel, "BAIE");
    const key = `${z}|${suffix}`;
    if (cacheCadreHor.has(key)) return cacheCadreHor.get(key);
    const v = await epCadreHorizontalDepuisSuffixes(cadre, z, suffix, getEpIntProduit);
    cacheCadreHor.set(key, v);
    return v;
  }

  // Index traverses par zone
  const traversesByZone = new Map();
  for (const t of traverses) {
    const z = normLabel(getChoiceLabel(t, "crbee_zone"), "BAIE");
    if (!traversesByZone.has(z)) traversesByZone.set(z, []);
    traversesByZone.get(z).push(t);
  }

  function hasValidPortee(debut, fin) {
    return debut !== null && debut !== undefined &&
      fin !== null && fin !== undefined &&
      isFinite(debut) && isFinite(fin) &&
      fin > debut + TOL;
  }

  async function findTraversePassage(zoneLabel, yDessous, x) {
    const TOL_Y = 5; // mm
    const z0 = normLabel(zoneLabel, "BAIE");
    const zonesToTry = [z0, "PASSAGE", "BAIE"];

    let best = null;
    let bestDist = 1e9;

    for (const z of zonesToTry) {
      const arr = traversesByZone.get(z) || [];
      for (const t of arr) {
        const deb = t.crbee_positionmmdebut ?? 0;
        const fin = (t.crbee_positionmmfin ?? Number.POSITIVE_INFINITY);
        if (!(deb <= x + TOL && fin >= x - TOL)) continue;

        const epExt = await getEpExtProduit(t._crbee_profil_value);
        const yCentreAttendu = yDessous + epExt / 2;
        const yPos = (t.crbee_positionmm ?? 0);

        const dCentre = Math.abs(yPos - yCentreAttendu);
        const dDessous = Math.abs(yPos - yDessous);
        const d = Math.min(dCentre, dDessous);

        if (d <= TOL_Y && d < bestDist) {
          best = t;
          bestDist = d;
        }
      }
      if (best) return best;
    }

    return null;
  }

  const updates = [];

  for (const m of montants) {
    const zoneLabel = normLabel(getChoiceLabel(m, "crbee_zone"), "BAIE");
    const isRenforce = (m.crbee_typeprofil === 3);

    const x = m.crbee_positionmm ?? 0;
    const y0 = m.crbee_positionmmdebut;
    const y1 = m.crbee_positionmmfin;

    let L = 0;

    if (isRenforce) {
      L = mm(Hbaie);
    } else if (hasValidPortee(y0, y1)) {
      // Montant d’imposte (FACE PARCLOSE)
      let start;
      if (Math.abs(y0 - 0) <= TOL) {
        start = await getCadreHor(zoneLabel, "_B");
      } else {
        const tPassage = await findTraversePassage(zoneLabel, y0, x);
        const epExtT = tPassage ? await getEpExtProduit(tPassage._crbee_profil_value) : 0;
        const epIntT = tPassage ? await getEpIntProduit(tPassage._crbee_profil_value) : 0;
        start = y0 + (epExtT / 2) + (epIntT / 2);
      }

      const epIntCadreHaut = await getCadreHor(zoneLabel, "_H");
      let end;

      if (Math.abs(y1 - Hbaie) <= TOL) end = Hbaie - epIntCadreHaut;
      else if (Math.abs(y1 - (Hbaie - epIntCadreHaut)) <= TOL) end = y1;
      else end = y1;

      L = mm(end - start);
    } else {
      // Montant standard filant au clair (face parclose)
      const epBas = await getCadreHor(zoneLabel, "_B");
      const epHaut = await getCadreHor(zoneLabel, "_H");
      L = mm(Hbaie - epBas - epHaut);
    }

    updates.push({
      id: m.crbee_profilchassisid,
      data: { crbee_longueur: L }
    });
  }

  await updateByChunks("crbee_profilchassis", updates, 20);
}

// ---------------- BAIE_TRAVERSE_H : DETECTION RECOUPEE ----------------
function isBaieTraverseHRecoupee(profils) {
  const hits = profils.filter(p => {
    const pos = normLabel(getChoiceLabel(p, "crbee_position"), "");
    if (pos !== "BAIE_TRAVERSE_H") return false;

    const deb = p.crbee_positionmmdebut;
    const fin = p.crbee_positionmmfin;

    const hasSeg = (deb !== null && deb !== undefined &&
                    fin !== null && fin !== undefined &&
                    isFinite(deb) && isFinite(fin) &&
                    Math.abs(fin - deb) > TOL);
    return hasSeg;
  });

  return hits.length >= 2;
}

// ---------------- 1) TRAVERSES AU CLAIR ENTRE PILES (+ BAIE_TRAVERSE_H RECOUPÉE) ----------------
async function calculerTraversesAuClairDepuisProfils(profils, contexte, opts = {}) {
  const getEpIntProduit = createProduitFaceCache(); // face parclose
  const baieTraverseHRecoupee = !!opts.baieTraverseHRecoupee;

  // Traverses "normales" + segments BAIE_TRAVERSE_H (role CADRE) quand recoupée
  const traverses = profils.filter(p => {
    if (p.crbee_role === ROLE.TRAVERSE) return true;

    if (!baieTraverseHRecoupee) return false;

    const pos = normLabel(getChoiceLabel(p, "crbee_position"), "");
    const isBaieH = (pos === "BAIE_TRAVERSE_H");
    const hasSeg = (p.crbee_positionmmdebut !== null && p.crbee_positionmmdebut !== undefined &&
                    p.crbee_positionmmfin !== null && p.crbee_positionmmfin !== undefined);
    return (p.crbee_role === ROLE.CADRE && isBaieH && hasSeg);
  });

  if (!traverses.length) return;

  const montantsInter = profils.filter(p => p.crbee_role === ROLE.MONTANT);

  // Montants verticaux CADRE comme piles internes (jonctions fixes/passage)
  const montantsCadreVert = profils.filter(p => {
    if (p.crbee_role !== ROLE.CADRE) return false;
    const pos = normLabel(getChoiceLabel(p, "crbee_position"), "");
    return pos.includes("MONTANT_");
  });

  const montants = montantsInter.concat(montantsCadreVert);

  const montantsByZone = new Map();
  for (const m of montants) {
    const z = normLabel(getChoiceLabel(m, "crbee_zone"), "BAIE");
    if (!montantsByZone.has(z)) montantsByZone.set(z, []);
    montantsByZone.get(z).push(m);
  }

  function montantCouvreY(m, y) {
    const y0 = m.crbee_positionmmdebut;
    const y1 = m.crbee_positionmmfin;

    // Pas de portée => filant
    if (y0 === null || y0 === undefined || y1 === null || y1 === undefined) return true;

    return (y >= y0 - TOL) && (y <= y1 + TOL);
  }

  function findMontantAtInZone(zoneLabel, x, yTraverse) {
    const z = normLabel(zoneLabel, "BAIE");
    const arr = montantsByZone.get(z) || [];
    let best = null;
    let bestDist = 1e9;

    for (const m of arr) {
      if (!montantCouvreY(m, yTraverse)) continue;

      const dx = Math.abs((m.crbee_positionmm ?? 0) - x);
      if (dx <= TOL && dx < bestDist) {
        best = m;
        bestDist = dx;
      }
    }
    return best;
  }

  // Pour BAIE : chercher le montant sur TOUTES les zones (FIXE_G, PASSAGE, FIXE_D, BAIE)
  function findMontantAtAnyZone(x, yTraverse) {
    let best = null;
    let bestDist = 1e9;

    for (const [, arr] of montantsByZone.entries()) {
      for (const m of arr) {
        if (!montantCouvreY(m, yTraverse)) continue;

        const dx = Math.abs((m.crbee_positionmm ?? 0) - x);
        if (dx <= TOL && dx < bestDist) {
          best = m;
          bestDist = dx;
        }
      }
    }
    return best;
  }

  const cadre = profils.filter(p => p.crbee_role === ROLE.CADRE);

  function isBaieTraverseH(p) {
    return normLabel(getChoiceLabel(p, "crbee_position"), "") === "BAIE_TRAVERSE_H";
  }

  async function epPileRive(zoneLabel, side /*G|D*/) {
    const z = normLabel(zoneLabel, "BAIE");
    const cfg = RIVE_POSITIONS[z] || RIVE_POSITIONS["BAIE"];
    const wanted = cfg?.[side] || [];
    if (!wanted.length) return 0;

    let sum = 0;
    for (const p of cadre) {
      const pZone = normLabel(getChoiceLabel(p, "crbee_zone"), "BAIE");
      if (pZone !== z) continue;

      const pos = normLabel(getChoiceLabel(p, "crbee_position"), "");
      if (!wanted.includes(pos)) continue;

      sum += await getEpIntProduit(p._crbee_profil_value);
    }
    return sum;
  }

  const pileCache = new Map();
  async function getPile(zoneLabel, side) {
    const z = normLabel(zoneLabel, "BAIE");
    const key = `${z}|${side}`;
    if (pileCache.has(key)) return pileCache.get(key);
    const v = await epPileRive(z, side);
    pileCache.set(key, v);
    return v;
  }

  // BONUS closoirs automatique (PASSAGE_CLOSOIR_G/D) s'ils s'arrêtent sous la traverse haute
  async function bonusClosoirsPassageSousTraverseH(getEpIntProduitFn, cadreProfils, yTraverse) {
    let bonus = 0;

    for (const p of cadreProfils) {
      const pos = normLabel(getChoiceLabel(p, "crbee_position"), "");
      if (pos !== "PASSAGE_CLOSOIR_G" && pos !== "PASSAGE_CLOSOIR_D") continue;

      // Si portée verticale renseignée : il doit s'arrêter sous la traverse haute pour compter
      const yFin = p.crbee_positionmmfin;
      if (yFin !== null && yFin !== undefined && isFinite(yFin)) {
        if (yFin >= yTraverse - TOL) continue;
      }

      bonus += await getEpIntProduitFn(p._crbee_profil_value);
    }

    return bonus;
  }

  const updates = [];

  for (const t of traverses) {
    const zoneLabel = normLabel(getChoiceLabel(t, "crbee_zone"), "BAIE");
    const Lzone = largeurZone(contexte, zoneLabel);
    if (!Lzone) continue;

    const yTraverse = t.crbee_positionmm ?? 0;

    const debut = t.crbee_positionmmdebut ?? 0;
    const fin = t.crbee_positionmmfin ?? 0;
    const base = fin - debut;

    const isZoneBaie = (normLabel(zoneLabel, "BAIE") === "BAIE");

    // CAS METIER : BAIE_TRAVERSE_H segment "PASSAGE" en montants filants
    const Lg = contexte.LARGEUR_FIXE_G || 0;
    const Lp = contexte.LARGEUR_PASSAGE || 0;

    const estSegmentPassage =
      isBaieTraverseH(t) &&
      isZoneBaie &&
      approxEq(debut, Lg, 1) &&
      approxEq(fin, Lg + Lp, 1) &&
      (contexte.MONTANTS_FILANTS === 1);

    const bonusClosoirs = estSegmentPassage
      ? await bonusClosoirsPassageSousTraverseH(getEpIntProduit, cadre, yTraverse)
      : 0;

    // épaisseur gauche
    let epG = 0;
    if (Math.abs(debut - 0) <= TOL) {
      epG = await getPile(zoneLabel, "G");
    } else {
      const m = isZoneBaie ? findMontantAtAnyZone(debut, yTraverse)
                           : findMontantAtInZone(zoneLabel, debut, yTraverse);
      epG = m ? (await getEpIntProduit(m._crbee_profil_value)) / 2 : 0;
    }

    // épaisseur droite
    let epD = 0;
    if (Math.abs(fin - Lzone) <= TOL) {
      epD = await getPile(zoneLabel, "D");
    } else {
      const m = isZoneBaie ? findMontantAtAnyZone(fin, yTraverse)
                           : findMontantAtInZone(zoneLabel, fin, yTraverse);
      epD = m ? (await getEpIntProduit(m._crbee_profil_value)) / 2 : 0;
    }

    // CAS METIER : sur le segment PASSAGE de BAIE_TRAVERSE_H en filant, pas de déduction des piles aux limites du passage
    if (estSegmentPassage) {
      epG = 0;
      epD = 0;
    }

    const L = mm(base - epG - epD + bonusClosoirs);

    updates.push({
      id: t.crbee_profilchassisid,
      data: { crbee_longueur: L }
    });
  }

  await updateByChunks("crbee_profilchassis", updates, 20);
}

// ---------------- 2) PARCLOSES (AMÉLIORÉES) ----------------
async function calculerParclosesDepuisProfils(profils, contexte) {
  const getEpIntProduit = createProduitFaceCache(); // face parclose

  const parcloses = profils.filter(p => p.crbee_role === ROLE.PARCLOSE);
  if (!parcloses.length) return;

  const cadre = profils.filter(p => p.crbee_role === ROLE.CADRE);
  const montants = profils.filter(p => p.crbee_role === ROLE.MONTANT);
  const traverses = profils.filter(p => p.crbee_role === ROLE.TRAVERSE);

  // ---- index montants par zone
  const montantsByZone = new Map();
  for (const m of montants) {
    const z = normLabel(getChoiceLabel(m, "crbee_zone"), "BAIE");
    if (!montantsByZone.has(z)) montantsByZone.set(z, []);
    montantsByZone.get(z).push(m);
  }

  function findMontantAt(zoneLabel, x) {
    const z = normLabel(zoneLabel, "BAIE");
    const arr = montantsByZone.get(z) || [];
    let best = null, bestDist = 1e9;
    for (const m of arr) {
      const dx = Math.abs((m.crbee_positionmm ?? 0) - x);
      if (dx <= TOL && dx < bestDist) {
        best = m; bestDist = dx;
      }
    }
    return best;
  }

  function findMontantAtAnyZone(x) {
    let best = null;
    let bestDist = 1e9;
    for (const arr of montantsByZone.values()) {
      for (const m of arr) {
        const dx = Math.abs((m.crbee_positionmm ?? 0) - x);
        if (dx <= TOL && dx < bestDist) {
          best = m;
          bestDist = dx;
        }
      }
    }
    return best;
  }

  // ---- index traverses par zone (pour trouver traverse au y0/y1 qui couvre x)
  const traversesByZone = new Map();
  for (const t of traverses) {
    const z = normLabel(getChoiceLabel(t, "crbee_zone"), "BAIE");
    if (!traversesByZone.has(z)) traversesByZone.set(z, []);
    traversesByZone.get(z).push(t);
  }

  function findTraverseAt(zoneLabel, y, x) {
    const z = normLabel(zoneLabel, "BAIE");
    const arr = traversesByZone.get(z) || [];
    let best = null, bestDist = 1e9;

    for (const t of arr) {
      const dy = Math.abs((t.crbee_positionmm ?? 0) - y);
      if (dy > TOL) continue;

      const deb = t.crbee_positionmmdebut ?? 0;
      const fin = (t.crbee_positionmmfin ?? 0);

      if (deb <= x + TOL && fin >= x - TOL) {
        if (dy < bestDist) {
          best = t;
          bestDist = dy;
        }
      }
    }
    return best;
  }

  function findTraverseAtAnyZone(y, x) {
    let best = null;
    let bestDist = 1e9;

    for (const arr of traversesByZone.values()) {
      for (const t of arr) {
        const dy = Math.abs((t.crbee_positionmm ?? 0) - y);
        if (dy > TOL) continue;

        const deb = t.crbee_positionmmdebut ?? 0;
        const fin = (t.crbee_positionmmfin ?? 0);

        if (deb <= x + TOL && fin >= x - TOL) {
          if (dy < bestDist) {
            best = t;
            bestDist = dy;
          }
        }
      }
    }
    return best;
  }


  // ---- pile rive (montant+closoir) pour horizontales
  async function epPileRive(zoneLabel, side /*G|D*/) {
    const z = normLabel(zoneLabel, "BAIE");
    const cfg = RIVE_POSITIONS[z] || RIVE_POSITIONS["BAIE"];
    const wanted = cfg?.[side] || [];
    if (!wanted.length) return 0;

    let sum = 0;
    for (const p of cadre) {
      const pZone = normLabel(getChoiceLabel(p, "crbee_zone"), "BAIE");
      if (pZone !== z) continue;

      const pos = normLabel(getChoiceLabel(p, "crbee_position"), "");
      if (!wanted.includes(pos)) continue;

      sum += await getEpIntProduit(p._crbee_profil_value);
    }
    return sum;
  }

  const pileCache = new Map();
  async function getPile(zoneLabel, side) {
    const z = normLabel(zoneLabel, "BAIE");
    const key = `${z}|${side}`;
    if (pileCache.has(key)) return pileCache.get(key);
    const v = await epPileRive(z, side);
    pileCache.set(key, v);
    return v;
  }

  // ---- cache cadre horizontal bas/haut (_B/_H) en face parclose
  const cadreHorCache = new Map();
  async function getCadreHor(zoneLabel, suffix /* "_B" | "_H" */) {
    const z = normLabel(zoneLabel, "BAIE");
    const key = `${z}|${suffix}`;
    if (cadreHorCache.has(key)) return cadreHorCache.get(key);
    const v = await epCadreHorizontalDepuisSuffixes(cadre, z, suffix, getEpIntProduit);
    cadreHorCache.set(key, v);
    return v;
  }

  // ---- group par cellulekey
  const byCell = new Map();
  for (const p of parcloses) {
    const key = p.crbee_cellulekey || "__NO_KEY__";
    if (!byCell.has(key)) byCell.set(key, []);
    byCell.get(key).push(p);
  }

  const updates = [];

  for (const [cellKey, items] of byCell.entries()) {
    const pG = items.find(x => x.crbee_cote === COTE.G);
    const pD = items.find(x => x.crbee_cote === COTE.D);
    const pB = items.find(x => x.crbee_cote === COTE.B);
    const pH = items.find(x => x.crbee_cote === COTE.H);

    const any = pB || pH || pG || pD;
    if (!any) continue;

    const Wbaie = contexte.LARGEUR_BAIE || 0;
    const Hzone = contexte.HAUTEUR_BAIE || 0;

    // (A) Horizontales
    async function calcHoriz(pHor) {
      if (!pHor) return null;
      const deb = pHor.crbee_positionmmdebut ?? 0;
      const fin = pHor.crbee_positionmmfin ?? 0;
      const base = fin - deb;

      let epG = 0;
      if (Math.abs(deb - 0) <= TOL) {
        epG = await getPile("BAIE", "G");
      } else {
        const m = findMontantAtAnyZone(deb);
        epG = m ? (await getEpIntProduit(m._crbee_profil_value)) / 2 : 0;
      }

      let epD = 0;
      if (Math.abs(fin - Wbaie) <= TOL) {
        epD = await getPile("BAIE", "D");
      } else {
        const m = findMontantAtAnyZone(fin);
        epD = m ? (await getEpIntProduit(m._crbee_profil_value)) / 2 : 0;
      }

      return mm(base - epG - epD);
    }

    const Lb = await calcHoriz(pB);
    const Lh = await calcHoriz(pH);

    if (pB && Lb !== null) updates.push({ id: pB.crbee_profilchassisid, data: { crbee_longueur: Lb } });
    if (pH && Lh !== null) updates.push({ id: pH.crbee_profilchassisid, data: { crbee_longueur: Lh } });

    const Wb = pB ? await getEpIntProduit(pB._crbee_profil_value) : 0;
    const Wh = pH ? await getEpIntProduit(pH._crbee_profil_value) : 0;

    // (B) Verticales
    async function calcVert(pVer) {
      if (!pVer) return null;

      const x = pVer.crbee_positionmm ?? 0;
      const y0 = pVer.crbee_positionmmdebut ?? 0;
      const y1 = pVer.crbee_positionmmfin ?? 0;

      let base = (y1 - y0);

      if (Math.abs(y0 - 0) <= TOL) {
        base -= await getCadreHor("BAIE", "_B");
      } else {
        const t0 = findTraverseAtAnyZone(y0, x);
        if (t0) base -= (await getEpIntProduit(t0._crbee_profil_value)) / 2;
      }

      if (Math.abs(y1 - Hzone) <= TOL) {
        base -= await getCadreHor("BAIE", "_H");
      } else {
        const t1 = findTraverseAtAnyZone(y1, x);
        if (t1) base -= (await getEpIntProduit(t1._crbee_profil_value)) / 2;
      }

      base -= (Wb + Wh);

      return mm(base);
    }

    const Lg = await calcVert(pG);
    const Ld = await calcVert(pD);

    if (pG && Lg !== null) updates.push({ id: pG.crbee_profilchassisid, data: { crbee_longueur: Lg } });
    if (pD && Ld !== null) updates.push({ id: pD.crbee_profilchassisid, data: { crbee_longueur: Ld } });
  }

  await updateByChunks("crbee_profilchassis", updates, 20);
}

// ---------------- 3) PROFILDATA POUR FORMULES NOMENCLATURE ----------------
async function chargerProfilData(profils) {
  const profilData = {};

  for (let profil of profils) {
    const positionRaw = profil["crbee_position@OData.Community.Display.V1.FormattedValue"];
    const position = normLabel(positionRaw, "");

    if (!position) continue;
    if (profilData[position]) continue;

    const produitId = profil._crbee_profil_value;
    if (!produitId) {
      profilData[position] = { ep_ext: 0, ep_int: 0 };
      continue;
    }

    const produit = await Xrm.WebApi.retrieveRecord(
      "crbee_produit",
      produitId,
      "?$select=crbee_largeurfaceexterieur,crbee_largeurfaceparclose"
    );

    profilData[position] = {
      ep_ext: produit.crbee_largeurfaceexterieur || 0,
      ep_int: produit.crbee_largeurfaceparclose || 0
    };
  }

  return profilData;
}

// ---------------- 4) EVALUATION FORMULE (MM ENTIER) ----------------
function evaluerFormule(formule, contexte, profilData) {
  let expression = formule;

  for (const cle in contexte) {
    expression = expression.replace(new RegExp(`\\b${cle}\\b`, "g"), contexte[cle]);
  }

  expression = expression.replace(/EPAISSEUR_EXT\(([^)]+)\)/g,
    (_, pos) => (profilData[normLabel(pos.trim(), "")]?.ep_ext ?? 0)
  );

  expression = expression.replace(/EPAISSEUR_INT\(([^)]+)\)/g,
    (_, pos) => (profilData[normLabel(pos.trim(), "")]?.ep_int ?? 0)
  );

  try {
    const v = eval(expression);
    return mm(v);
  } catch (e) {
    console.error("Erreur formule :", formule, e);
    return null;
  }
}

// ================== FONCTION PRINCIPALE (BOUTON CALCUL) ==================
async function calculerDebitsProfils(formContext) {
  const chassisId = formContext.data.entity.getId().replace("{", "").replace("}", "");

  // Dimensions châssis
  const largeurBaie = formContext.getAttribute("crbee_largeurdelabaie")?.getValue();
  const hauteurBaie = formContext.getAttribute("crbee_hauteurdelabaie")?.getValue();
  const largeurFixeG = formContext.getAttribute("crbee_largeurchassisgauche")?.getValue();
  const largeurFixeD = formContext.getAttribute("crbee_largeurchassisdroite")?.getValue();
  const largeurPassage = formContext.getAttribute("crbee_largeurdepassage")?.getValue();
  const hauteurPassage = formContext.getAttribute("crbee_hauteurdepassage")?.getValue();
  const montantsFilants = formContext.getAttribute("crbee_montantsfilants")?.getValue();

  if (!largeurBaie || !hauteurBaie) {
    Xrm.Navigation.openAlertDialog({ text: "Les dimensions du châssis sont incomplètes." });
    return;
  }

  // Charger profils châssis
  const profilsResp = await Xrm.WebApi.retrieveMultipleRecords(
    "crbee_profilchassis",
    `?$filter=_crbee_chassis_value eq ${chassisId}`
  );
  const profils = profilsResp.entities;

  if (!profils.length) {
    Xrm.Navigation.openAlertDialog({ text: "Aucun profil à calculer." });
    return;
  }

  // détecter si BAIE_TRAVERSE_H est recoupée
  const baieTraverseHRecoupee = isBaieTraverseHRecoupee(profils);

  // Contexte pour formules
  const contexte = {
    LARGEUR_BAIE: largeurBaie,
    HAUTEUR_BAIE: hauteurBaie,
    LARGEUR_FIXE_G: largeurFixeG,
    LARGEUR_FIXE_D: largeurFixeD,
    LARGEUR_PASSAGE: largeurPassage,
    HAUTEUR_PASSAGE: hauteurPassage || 0,
    MONTANTS_FILANTS: montantsFilants ? 1 : 0
  };

  // Charger profilData (épaisseurs) pour formules
  const profilData = await chargerProfilData(profils);

  // Charger nomenclature modèle
  const modele = formContext.getAttribute("crbee_modele")?.getValue();
  if (!modele || modele.length === 0) {
    Xrm.Navigation.openAlertDialog({ text: "Aucun modèle sélectionné sur le châssis." });
    return;
  }
  const modeleId = modele[0].id.replace("{", "").replace("}", "");

  const nomenclatureResp = await Xrm.WebApi.retrieveMultipleRecords(
    "crbee_nomenclature",
    `?$filter=_crbee_modele_value eq ${modeleId}`
  );
  const nomenclature = nomenclatureResp.entities;

  // Index formule nomenclature par POSITION
  const mapNomenclature = {};
  nomenclature.forEach(n => {
    const posTxt = normLabel(n["crbee_position@OData.Community.Display.V1.FormattedValue"], "");
    mapNomenclature[posTxt] = n.crbee_regledecalcul;
  });

  // 1) Calcul profils pilotés par nomenclature (cadre etc.)
  // SAUF BAIE_TRAVERSE_H si recoupée => on la laisse au calcul "au clair"
  for (const profil of profils) {
    const position = normLabel(profil["crbee_position@OData.Community.Display.V1.FormattedValue"], "");

    if (baieTraverseHRecoupee && position === "BAIE_TRAVERSE_H") {
      continue;
    }

    const formule = mapNomenclature[position];
    if (!formule) continue;

    const longueur = evaluerFormule(formule, contexte, profilData);
    if (longueur === null || longueur === undefined) continue;

    await Xrm.WebApi.updateRecord(
      "crbee_profilchassis",
      profil.crbee_profilchassisid,
      { crbee_longueur: longueur }
    );
  }

  // 2) Montants intermédiaires (incluant imposte)
  await calculerMontantsIntermediairesDepuisProfils(profils, contexte);

  // 3) Traverses au clair (inclut BAIE_TRAVERSE_H recoupée segment par segment)
  await calculerTraversesAuClairDepuisProfils(profils, contexte, { baieTraverseHRecoupee });

  // 4) Parcloses
  await calculerParclosesDepuisProfils(profils, contexte);

  Xrm.Navigation.openAlertDialog({
    text: `Calcul terminé. BAIE_TRAVERSE_H recoupée = ${baieTraverseHRecoupee ? "OUI" : "NON"}`
  });

  formContext.data.refresh();
}
