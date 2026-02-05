/**
 * genererProfilsDepuisNomenclature_v6 — VERSION IMPOSTE (FIXES)
 *
 * FIX #1 : Si montantsFilants = true, les montants CADRE doivent couper BAIE_TRAVERSE_H (traverse haute).
 *          => On segmente BAIE_TRAVERSE_H à la création, sur les X des montants CADRE internes.
 *
 * FIX #2 : Les cellules vitrages (parcloses) ne doivent pas dépendre uniquement des montants INTERMEDIAIRES.
 *          => allMontantsAbs = montants INTERMEDIAIRES + montants CADRE (toujours).
 *
 * RÈGLE MÉTIER :
 * - Montant FILANT => coupe traverses + participe aux découpes vitrages (sur toute la hauteur)
 * - Montant NON filant => ne coupe pas les traverses (pas de segmentation) et ne crée pas de vitrage supplémentaire
 */

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

const ORI = { H: 1, V: 2 };

// Tolérance mm
const TOL = 0.5;


// ===================== PATCH PARCLOSES (épaisseurs intérieures) =====================
// Chez toi : l'épaisseur intérieure des profils = champ crbee_largeurfaceparclose
const FIELD_EPAISSEUR_INTERIEURE_PRODUIT = "crbee_largeurfaceparclose";

// util : charger plusieurs produits et leurs champs (avec chunk pour éviter URL trop longue)
async function chargerProduitsProps(ids, selectFields) {
  const uniq = [...new Set((ids || []).filter(Boolean))];
  const map = new Map();
  if (!uniq.length) return map;

  const chunkSize = 20;
  for (let i = 0; i < uniq.length; i += chunkSize) {
    const chunk = uniq.slice(i, i + chunkSize);
    // (crbee_produitid eq guid) OR ...
    const filter = chunk.map(id => `crbee_produitid eq ${id}`).join(" or ");
    const q = `?$select=crbee_produitid,${selectFields.join(",")}&$filter=${filter}`;
    const resp = await Xrm.WebApi.retrieveMultipleRecords("crbee_produit", q);
    for (const p of (resp.entities || [])) map.set(p.crbee_produitid, p);
  }
  return map;
}
// ===================== FIN PATCH PARCLOSES =====================


async function supprimerTousLesProfilsDuChassis(chassisId) {
  // On supprime tout ce qui est dans crbee_profilchassis pour ce châssis
  // (Option 1 : reset complet)
  let next = true;
  let page = null;

  while (next) {
    const query =
      `?$select=crbee_profilchassisid` +
      `&$filter=_crbee_chassis_value eq ${chassisId}` +
      `&$top=5000` +
      (page ? `&$skiptoken=${encodeURIComponent(page)}` : "");

    const resp = await Xrm.WebApi.retrieveMultipleRecords("crbee_profilchassis", query);

    const ids = resp.entities.map(e => e.crbee_profilchassisid);
    if (!ids.length) break;

    // suppression en parallèle (chunk)
    const chunkSize = 20;
    for (let i = 0; i < ids.length; i += chunkSize) {
      const chunk = ids.slice(i, i + chunkSize);
      await Promise.all(chunk.map(id => Xrm.WebApi.deleteRecord("crbee_profilchassis", id)));
    }

    // pagination
    const nextLink = resp["@odata.nextLink"];
    if (nextLink) {
      const m = nextLink.match(/\$skiptoken=([^&]+)/);
      page = m ? decodeURIComponent(m[1]) : null;
      next = !!page;
    } else {
      next = false;
    }
  }
}


function getChoiceLabel(entity, fieldName) {
  return entity[`${fieldName}@OData.Community.Display.V1.FormattedValue`] || "";
}

function uniqSorted(arr) {
  return [...new Set((arr || []).filter(v => v !== null && v !== undefined))].sort((a, b) => a - b);
}

/**
 * IMPORTANT :
 * crbee_zone = zone de largeur (BAIE/FIXE_G/PASSAGE/FIXE_D)
 * "IMPOSTE" n’est pas une zone de largeur -> on la remappe.
 */
function normalizeZoneLabelForWidth(zLabel) {
  const z = (zLabel || "BAIE").toUpperCase().trim();
  if (z === "IMPOSTE") return "PASSAGE";
  return z;
}

function zoneLargeur(contexte, zoneLabel) {
  switch ((zoneLabel || "BAIE").toUpperCase()) {
    case "FIXE_G": return contexte.LARGEUR_FIXE_G || 0;
    case "FIXE_D": return contexte.LARGEUR_FIXE_D || 0;
    case "PASSAGE": return contexte.LARGEUR_PASSAGE || 0;
    default: return contexte.LARGEUR_BAIE || 0;
  }
}

function zoneOffsetX(contexte, zoneLabel) {
  const Lg = contexte.LARGEUR_FIXE_G || 0;
  const Lp = contexte.LARGEUR_PASSAGE || 0;

  switch ((zoneLabel || "BAIE").toUpperCase()) {
    case "FIXE_G":  return 0;
    case "PASSAGE": return Lg;
    case "FIXE_D":  return Lg + Lp;
    default:        return 0;
  }
}

// -------------------------
// RÈGLES COUPE FILANT / NON FILANT
// -------------------------

function isMontantFilant(m, H) {
  const y0 = m.crbee_porteedebutmm;
  const y1 = m.crbee_porteefinmm;

  // Pas de portée => on considère filant
  if (y0 === null || y0 === undefined || y1 === null || y1 === undefined) return true;

  // Filant si couvre ~[0..H]
  return (y0 <= 0 + TOL) && (y1 >= H - TOL);
}

// Coupe une traverse à yTraverse ?
// - FILANT => OUI
// - NON FILANT => OUI seulement si yTraverse est strictement à l'intérieur (pas en appui)
function coupeTraverse(m, yTraverse, H) {
  if (isMontantFilant(m, H)) return true;

  const y0 = m.crbee_porteedebutmm ?? 0;
  const y1 = m.crbee_porteefinmm ?? 999999;

  return (yTraverse > y0 + TOL) && (yTraverse < y1 - TOL);
}

// Coupe vitrage dans une bande (y0..y1) ?
// - FILANT => OUI
// - NON FILANT => OUI seulement s'il traverse la bande (strict)
function coupeVitrageBande(m, y0, y1, H) {
  if (isMontantFilant(m, H)) return true;

  const a = m.crbee_porteedebutmm ?? m.y0 ?? 0;
  const b = m.crbee_porteefinmm ?? m.y1 ?? 999999;

  // ✅ Le montant coupe si sa portée recouvre la bande [y0..y1]
  // Important : on accepte a == y0 (montant qui démarre sur la traverse d’imposte)
  return (a <= y0 + TOL) && (b >= y1 - TOL);
}



// -------------------------
// CADRE viewer enrichment
// -------------------------
function enrichirCadrePourViewer(ligne, contexte, opts = {}) {
  const hasImposte = !!opts.hasImposte;
  const montantsFilants = !!opts.montantsFilants;
  const hauteurPassage = opts.hauteurPassage;

  const posLabel = (getChoiceLabel(ligne, "crbee_position") || "").toUpperCase();
  const zoneLabel = normalizeZoneLabelForWidth(getChoiceLabel(ligne, "crbee_zone") || "BAIE");

  const offX = zoneOffsetX(contexte, zoneLabel);
  const Lz = zoneLargeur(contexte, zoneLabel);
  const xStart = offX;
  const xEnd = offX + Lz;

  const H = contexte.HAUTEUR_BAIE || 0;

  if (posLabel.includes("TRAVERSE_H")) {
    return {
      crbee_orientation: ORI.H,
      crbee_typeprofil: 2,
      crbee_positionmm: H,
      crbee_positionmmdebut: xStart,
      crbee_positionmmfin: xEnd
    };
  }

  if (posLabel.includes("TRAVERSE_B")) {
    return {
      crbee_orientation: ORI.H,
      crbee_typeprofil: 2,
      crbee_positionmm: 0,
      crbee_positionmmdebut: xStart,
      crbee_positionmmfin: xEnd
    };
  }

  if (posLabel.includes("MONTANT_G")) {
    // Exemple conservé : certains montants internes peuvent s’arrêter à hauteurPassage selon options
    const estMontantInterneNonFilant = (
      hasImposte && !montantsFilants && hauteurPassage !== null && hauteurPassage !== undefined &&
      (zoneLabel === "FIXE_D" && posLabel.includes("MONTANT_G"))
    );
    const yEnd = estMontantInterneNonFilant ? hauteurPassage : H;
    return {
      crbee_orientation: ORI.V,
      crbee_typeprofil: 1,
      crbee_positionmm: xStart,
      crbee_positionmmdebut: 0,
      crbee_positionmmfin: yEnd
    };
  }

  if (posLabel.includes("MONTANT_D")) {
    const estMontantInterneNonFilant = (
      hasImposte && !montantsFilants && hauteurPassage !== null && hauteurPassage !== undefined &&
      (zoneLabel === "FIXE_G" && posLabel.includes("MONTANT_D"))
    );
    const yEnd = estMontantInterneNonFilant ? hauteurPassage : H;
    return {
      crbee_orientation: ORI.V,
      crbee_typeprofil: 1,
      crbee_positionmm: xEnd,
      crbee_positionmmdebut: 0,
      crbee_positionmmfin: yEnd
    };
  }

  return {};
}

/**
 * Ajoute/Met à jour une traverse d’imposte "obligatoire" dans crbee_profilintermediaire (zone BAIE)
 */
async function ajouterTraverseImposteSiBesoin(chassisId, modeleId, contexte, zoneValueBaie) {
  const modele = await Xrm.WebApi.retrieveRecord(
    "crbee_modele",
    modeleId,
    "?$select=crbee_hasimposte,_crbee_traverseimposte_value"
  );
  if (!modele.crbee_hasimposte) return;

  const produitId = modele._crbee_traverseimposte_value;
  if (!produitId) {
    Xrm.Navigation.openAlertDialog({ text: "Modèle: imposte activée mais crbee_traverseimposte non renseigné." });
    return;
  }

  const ch = await Xrm.WebApi.retrieveRecord(
    "crbee_chassisfabrique",
    chassisId,
    "?$select=crbee_hauteurdepassage"
  );

  const y = ch.crbee_hauteurdepassage;
  if (y === null || y === undefined) {
    Xrm.Navigation.openAlertDialog({ text: "Châssis: imposte activée mais crbee_hauteurdepassage non renseignée." });
    return;
  }

  const x0 = 0;
  const x1 = contexte.LARGEUR_BAIE || 0;
  if (!x1 || x1 <= 0) {
    Xrm.Navigation.openAlertDialog({ text: "Châssis: largeur baie manquante pour générer la traverse d’imposte." });
    return;
  }

  if (zoneValueBaie === undefined || zoneValueBaie === null) {
    Xrm.Navigation.openAlertDialog({ text: "Impossible de déterminer la valeur du Choice ZONE=BAIE (zoneValueBaie manquant)." });
    return;
  }

  // Eviter doublon (même zone, même y)
  const existing = await Xrm.WebApi.retrieveMultipleRecords(
    "crbee_profilintermediaire",
    `?$select=crbee_profilintermediaireid
     &$filter=_crbee_chassis_value eq ${chassisId}
       and crbee_type eq 2
       and crbee_zone eq ${zoneValueBaie}
       and crbee_positionmm eq ${y}`
  );

  if (existing.entities.length > 0) {
    const id = existing.entities[0].crbee_profilintermediaireid;
    await Xrm.WebApi.updateRecord("crbee_profilintermediaire", id, {
      "crbee_Produit@odata.bind": `/crbee_produits(${produitId})`,
      crbee_porteedebutmm: x0,
      crbee_porteefinmm: x1
    });
    return;
  }

  await Xrm.WebApi.createRecord("crbee_profilintermediaire", {
    "crbee_Chassis@odata.bind": `/crbee_chassisfabriques(${chassisId})`,
    "crbee_Produit@odata.bind": `/crbee_produits(${produitId})`,
    crbee_type: 2,
    crbee_zone: zoneValueBaie, // ✅ BAIE
    crbee_positionmm: y,
    crbee_porteedebutmm: x0,
    crbee_porteefinmm: x1
  });
}

/**
 * Segmentation de BAIE_TRAVERSE_H (cadre) sur les X des montants CADRE.
 * Retourne une liste de segments [{debut, fin}...]
 */
function segmenterBaieTraverseHSurMontantsCadre(xStart, xEnd, cadreMontantsAbs) {
  const xCuts = uniqSorted(
    (cadreMontantsAbs || [])
      .map(m => m.xAbs)
      .filter(x => x > xStart + TOL && x < xEnd - TOL)
  );

  const X = [xStart, ...xCuts, xEnd].sort((a, b) => a - b);
  const segs = [];
  for (let i = 0; i < X.length - 1; i++) {
    const a = X[i];
    const b = X[i + 1];
    if (b > a + TOL) segs.push({ debut: a, fin: b });
  }
  return segs;
}

async function genererProfilsDepuisNomenclature_v6(primaryControl) {
  const formContext = primaryControl;
  const chassisId = formContext.data.entity.getId().replace("{", "").replace("}", "");
  const modele = formContext.getAttribute("crbee_modele")?.getValue();
  
  // 🔹 Récupération de la commande de production liée au châssis
	const chassisRecWithCommande = await Xrm.WebApi.retrieveRecord(
	  "crbee_chassisfabrique",
	  chassisId,
	  "?$select=_crbee_commandeproduction_value"
	);

const commandeProductionId = chassisRecWithCommande._crbee_commandeproduction_value || null;

  const contexte = {
    LARGEUR_BAIE: formContext.getAttribute("crbee_largeurdelabaie")?.getValue(),
    HAUTEUR_BAIE: formContext.getAttribute("crbee_hauteurdelabaie")?.getValue(),
    LARGEUR_FIXE_G: formContext.getAttribute("crbee_largeurchassisgauche")?.getValue(),
    LARGEUR_FIXE_D: formContext.getAttribute("crbee_largeurchassisdroite")?.getValue(),
    LARGEUR_PASSAGE: formContext.getAttribute("crbee_largeurdepassage")?.getValue()
  };

  if (!contexte.LARGEUR_BAIE || !contexte.HAUTEUR_BAIE) {
    Xrm.Navigation.openAlertDialog({ text: "Dimensions châssis incomplètes (largeur/hauteur baie)." });
    return;
  }

  if (!modele) {
    Xrm.Navigation.openAlertDialog({ text: "Aucun modèle lié à ce châssis." });
    return;
  }
  const modeleId = modele[0].id.replace("{", "").replace("}", "");

  // ✅ Option 1 : reset complet avant regen
  const confirm = await Xrm.Navigation.openConfirmDialog({
	  title: "Régénérer les profils",
	  text: "Cette action supprime tous les profils existants du châssis puis les regénère. Continuer ?"
	});
	if (!confirm.confirmed) return;

	await supprimerTousLesProfilsDuChassis(chassisId);


	  // Paramètres châssis + PARCLOSE depuis châssis
	const chassisRec = await Xrm.WebApi.retrieveRecord(
	  "crbee_chassisfabrique",
	  chassisId,
	  "?$select=crbee_montantsfilants,crbee_hauteurdepassage,_crbee_parclose_value"
	);
	const montantsFilants = !!chassisRec.crbee_montantsfilants;
	const hauteurPassage = chassisRec.crbee_hauteurdepassage;
	const produitParcloseId = chassisRec?._crbee_parclose_value || null;

	if (!produitParcloseId) {
	  await Xrm.Navigation.openAlertDialog({
		text: "Le châssis fabriqué n'a pas de parclose (crbee_parclose)."
	  });
	  return;
	}

	// Modèle (imposte uniquement)
	const modeleRec = await Xrm.WebApi.retrieveRecord(
	  "crbee_modele",
	  modeleId,
	  "?$select=crbee_hasimposte,_crbee_traverseimposte_value"
	);
	const hasImposte = !!modeleRec?.crbee_hasimposte;


  // Map zoneLabel -> zoneValue (numérique)
  const zoneValueByLabel = {};
  function rememberZone(entity) {
    const val = entity.crbee_zone;
    const label = getChoiceLabel(entity, "crbee_zone");
    if (val !== null && val !== undefined && label) zoneValueByLabel[label.toUpperCase()] = val;
  }

  // -------------------------------------------------------------------
  // 1) CADRE depuis nomenclature (2-PASS) + FIX BAIE_TRAVERSE_H segmentée
  // -------------------------------------------------------------------
  const nomResp = await Xrm.WebApi.retrieveMultipleRecords(
    "crbee_nomenclature",
    `?$select=crbee_nomenclatureid,crbee_zone,crbee_position,crbee_coupe,crbee_quantite,_crbee_profil_value
     &$filter=_crbee_modele_value eq ${modeleId} and crbee_typedelement ne 745350007 and crbee_quantite gt 0`
  );

  const lignesCadre = [];
  const cadreMontantsAbs = []; // [{xAbs,y0,y1}]
  const cadreTraversesAbs = []; // [{yAbs,x0,x1,produitId}]


  // Pass 1: enrich + collecte montants CADRE
  for (const ligne of nomResp.entities) {
    rememberZone(ligne);

    const enrich = enrichirCadrePourViewer(ligne, contexte, {
      hasImposte,
      montantsFilants,
      hauteurPassage
    });

    const posLabel = (getChoiceLabel(ligne, "crbee_position") || "").toUpperCase();

    // Collecte montants CADRE (séparations FIXE/PASSAGE/FIXE_D inclus)
    if (enrich.crbee_orientation === ORI.V) {
		cadreMontantsAbs.push({
			xAbs: enrich.crbee_positionmm,
			y0: enrich.crbee_positionmmdebut ?? 0,
			y1: enrich.crbee_positionmmfin ?? (contexte.HAUTEUR_BAIE || 999999),
			produitId: ligne._crbee_profil_value || null
		});
	}
	
	if (enrich.crbee_orientation === ORI.H) {
		cadreTraversesAbs.push({
			yAbs: enrich.crbee_positionmm,
			x0: enrich.crbee_positionmmdebut ?? 0,
			x1: enrich.crbee_positionmmfin ?? (contexte.LARGEUR_BAIE || 999999),
			produitId: ligne._crbee_profil_value || null
		});
	}



    lignesCadre.push({ ligne, enrich, posLabel });
  }

  // Pass 2: création CADRE (avec segmentation BAIE_TRAVERSE_H si montantsFilants=true)
  for (const item of lignesCadre) {
    const { ligne, enrich, posLabel } = item;

    // Si traverse haute de BAIE et montants filants => découper sur montants CADRE internes
    if (
      montantsFilants &&
      enrich.crbee_orientation === ORI.H &&
      posLabel.includes("BAIE_TRAVERSE_H")
    ) {
      const segs = segmenterBaieTraverseHSurMontantsCadre(
        enrich.crbee_positionmmdebut,
        enrich.crbee_positionmmfin,
        cadreMontantsAbs
      );

      for (const s of segs) {
        await Xrm.WebApi.createRecord("crbee_profilchassis", {
          "crbee_Chassis@odata.bind": `/crbee_chassisfabriques(${chassisId})`,
          "crbee_Nomenclature@odata.bind": `/crbee_nomenclatures(${ligne.crbee_nomenclatureid})`,
          "crbee_Profil@odata.bind": `/crbee_produits(${ligne._crbee_profil_value})`,
		  // 🔹 Lien commande de production
		  "crbee_Commandeproduction@odata.bind":`/crbee_commandedachats(${commandeProductionId})`,
 
          "crbee_zone": ligne.crbee_zone || null,
          "crbee_position": ligne.crbee_position || null,
          "crbee_coupe": ligne.crbee_coupe,
          "crbee_quantite": 1,
          "crbee_source": 1,
          "crbee_role": ROLE.CADRE,
          ...enrich,
          crbee_positionmmdebut: s.debut,
          crbee_positionmmfin: s.fin
        });
      }

      continue; // ne pas créer le profil non segmenté
    }

    // Normal
    await Xrm.WebApi.createRecord("crbee_profilchassis", {
      "crbee_Chassis@odata.bind": `/crbee_chassisfabriques(${chassisId})`,
      "crbee_Nomenclature@odata.bind": `/crbee_nomenclatures(${ligne.crbee_nomenclatureid})`,
      "crbee_Profil@odata.bind": `/crbee_produits(${ligne._crbee_profil_value})`,
	  "crbee_Commandeproduction@odata.bind":`/crbee_commandedachats(${commandeProductionId})`,
      "crbee_zone": ligne.crbee_zone || null,
      "crbee_position": ligne.crbee_position || null,
      "crbee_coupe": ligne.crbee_coupe,
      "crbee_quantite": ligne.crbee_quantite,
      "crbee_source": 1,
      "crbee_role": ROLE.CADRE,
      ...enrich
    });
  }

  // 1bis) Ajout traverse d’imposte (en BAIE)
  const zoneValueBaie = zoneValueByLabel["BAIE"];
  await ajouterTraverseImposteSiBesoin(chassisId, modeleId, contexte, zoneValueBaie);

  // -------------------------------------------------------------------
  // 2) Charger intermediaires (avec portée)
  // -------------------------------------------------------------------
  const interResp = await Xrm.WebApi.retrieveMultipleRecords(
    "crbee_profilintermediaire",
    `?$select=crbee_type,crbee_positionmm,crbee_zone,crbee_porteedebutmm,crbee_porteefinmm,_crbee_produit_value
     &$filter=_crbee_chassis_value eq ${chassisId}`
  );
  const inter = interResp.entities;

  // 3) Répartir intermediaires par zone (labels normalisés)
  const montantsEntitiesByZoneLabel = {};
  const traversesEntitiesByZoneLabel = {};

  for (const p of inter) {
    rememberZone(p);
    const rawZ = (getChoiceLabel(p, "crbee_zone") || "BAIE");
    const zLabel = normalizeZoneLabelForWidth(rawZ);

    if (p.crbee_type === 1 || p.crbee_type === 3) {
      montantsEntitiesByZoneLabel[zLabel] ||= [];
      montantsEntitiesByZoneLabel[zLabel].push(p);
    } else if (p.crbee_type === 2) {
      traversesEntitiesByZoneLabel[zLabel] ||= [];
      traversesEntitiesByZoneLabel[zLabel].push(p);
    }
  }

  // 4) Créer montants (avec portée Y) dans profilchassis (X ABSOLU)
  const H = contexte.HAUTEUR_BAIE || 0;

  for (const zoneLabel of Object.keys(montantsEntitiesByZoneLabel)) {
    const zoneLabelU = normalizeZoneLabelForWidth(zoneLabel || "BAIE");
    const zoneValue = zoneValueByLabel[zoneLabelU] ?? null;
    const offX = zoneOffsetX(contexte, zoneLabelU);

    for (const m of montantsEntitiesByZoneLabel[zoneLabel]) {
      const y0 = (m.crbee_porteedebutmm ?? 0);
      const y1 = (m.crbee_porteefinmm ?? H);

      await Xrm.WebApi.createRecord("crbee_profilchassis", {
        "crbee_Chassis@odata.bind": `/crbee_chassisfabriques(${chassisId})`,
        "crbee_Profil@odata.bind": `/crbee_produits(${m._crbee_produit_value})`,
		"crbee_Commandeproduction@odata.bind":`/crbee_commandedachats(${commandeProductionId})`,
        "crbee_source": 2,
        "crbee_typeprofil": m.crbee_type, // 1 montant, 3 montant renforcé
        "crbee_role": ROLE.MONTANT,
        "crbee_orientation": ORI.V,
        "crbee_zone": zoneValue ?? m.crbee_zone ?? null,
        "crbee_positionmm": (m.crbee_positionmm || 0) + offX, // X ABSOLU
        "crbee_positionmmdebut": y0, // portée Y
        "crbee_positionmmfin": y1,   // portée Y
        "crbee_quantite": 1,
        "crbee_coupe": 745350001
      });
    }
  }

  // -------------------------------------------------------------------
  // 5) allMontantsAbs = INTERMEDIAIRES + CADRE (TOUJOURS)  ✅ FIX vitrages
  // -------------------------------------------------------------------
  const allMontantsAbs = []
    // intermediaires
    .concat(montantsEntitiesByZoneLabel["FIXE_G"] || [])
    .concat(montantsEntitiesByZoneLabel["PASSAGE"] || [])
    .concat(montantsEntitiesByZoneLabel["FIXE_D"] || [])
    .concat(montantsEntitiesByZoneLabel["BAIE"] || [])
    .map(m => {
      const zLab = normalizeZoneLabelForWidth(getChoiceLabel(m, "crbee_zone") || "BAIE");
      const off = zoneOffsetX(contexte, zLab);
      return {
		xAbs: (m.crbee_positionmm || 0) + off,
		y0: m.crbee_porteedebutmm ?? 0,
		y1: m.crbee_porteefinmm ?? 999999,
		produitId: m._crbee_produit_value || null
	 };

    })
    // + cadre (séparations fixes/passage/etc.) => indispensable pour découper les vitrages
    .concat(
      (cadreMontantsAbs || []).map(m => ({
        xAbs: m.xAbs,
        y0: m.y0 ?? 0,
        y1: m.y1 ?? 999999
      }))
    )
    .map(m => {
      // Si montantsFilants = true => on force filant partout pour les découpes
      if (montantsFilants) return { ...m, y0: 0, y1: H };
      return m;
    })
    .sort((a, b) => a.xAbs - b.xAbs);

  // -------------------------------------------------------------------
  // 6) Créer traverses segmentées (respect portée X + couverture Y)
  // -------------------------------------------------------------------
  for (const zoneLabel of Object.keys(traversesEntitiesByZoneLabel)) {
    const zoneLabelU = normalizeZoneLabelForWidth(zoneLabel || "BAIE");
    const zoneValue = zoneValueByLabel[zoneLabelU] ?? null;

    const offX = zoneOffsetX(contexte, zoneLabelU);
    const Lz = zoneLargeur(contexte, zoneLabelU);

    // Montants inter de cette zone (COORD LOCALES)
    const montantsInterZone = (montantsEntitiesByZoneLabel[zoneLabelU] || []).slice()
      .sort((a, b) => (a.crbee_positionmm || 0) - (b.crbee_positionmm || 0));

    for (const t of traversesEntitiesByZoneLabel[zoneLabel]) {
      const yTraverse = t.crbee_positionmm;

      // portée X LOCALE de la traverse
      const x0Local = (t.crbee_porteedebutmm ?? 0);
      const x1Local = (t.crbee_porteefinmm ?? Lz);

      // borne ABS (pour écrire dans profilchassis)
      const x0Abs = x0Local + offX;
      const x1Abs = x1Local + offX;

      // Liste des X de coupe
      let xCutsAbs = [];

      if (zoneLabelU === "BAIE") {
        // Traverse BAIE : couper avec TOUS les montants (cadre + inter), en ABS
        xCutsAbs = allMontantsAbs
          .filter(m => m.xAbs > x0Abs + TOL && m.xAbs < x1Abs - TOL)
          .filter(m => coupeTraverse(
            { crbee_porteedebutmm: m.y0, crbee_porteefinmm: m.y1 },
            yTraverse,
            H
          ))
          .map(m => m.xAbs);
      } else {
        // Traverse non-BAIE : couper seulement avec montants inter de la zone
        xCutsAbs = montantsInterZone
          .filter(m =>
            (m.crbee_positionmm ?? 0) > x0Local + TOL &&
            (m.crbee_positionmm ?? 0) < x1Local - TOL &&
            coupeTraverse(m, yTraverse, H)
          )
          .map(m => (m.crbee_positionmm ?? 0) + offX);
      }

      xCutsAbs = uniqSorted(xCutsAbs);

      // Segments
      let debutAbs = x0Abs;

      for (let i = 0; i <= xCutsAbs.length; i++) {
        const finAbs = (i < xCutsAbs.length) ? xCutsAbs[i] : x1Abs;
        if (finAbs <= debutAbs + TOL) {
          debutAbs = finAbs;
          continue;
        }

        await Xrm.WebApi.createRecord("crbee_profilchassis", {
          "crbee_Chassis@odata.bind": `/crbee_chassisfabriques(${chassisId})`,
          "crbee_Profil@odata.bind": `/crbee_produits(${t._crbee_produit_value})`,
		  "crbee_Commandeproduction@odata.bind":`/crbee_commandedachats(${commandeProductionId})`,
          "crbee_source": 2,
          "crbee_typeprofil": 2,
          "crbee_role": ROLE.TRAVERSE,
          "crbee_orientation": ORI.H,
          "crbee_zone": zoneValue ?? t.crbee_zone ?? null,
          "crbee_positionmmdebut": debutAbs,
          "crbee_positionmmfin": finAbs,
          "crbee_positionmm": yTraverse,
          "crbee_quantite": 1,
          "crbee_coupe": 745350001
        });

        debutAbs = finAbs;
      }
    }
  }

// 7) PARCLOSES — HYBRIDE : bas = FIXE_G + FIXE_D, haut (imposte) = BAIE pleine largeur
// -------------------------------------------------------------------
{
  const H = contexte.HAUTEUR_BAIE || 0;
  const W = contexte.LARGEUR_BAIE || 0;
  const Lg = contexte.LARGEUR_FIXE_G || 0;
  const Lp = contexte.LARGEUR_PASSAGE || 0;

  if (W <= 0 || H <= 0) throw new Error("Dimensions BAIE invalides pour génération parcloses.");

  const zoneValueBaie = zoneValueByLabel["BAIE"];
  if (zoneValueBaie === undefined || zoneValueBaie === null) {
    throw new Error("ZoneValue BAIE introuvable (zoneValueByLabel['BAIE']).");
  }

  // --- 1) yImposte (si traverse BAIE existe)
  const traversesBaie = (traversesEntitiesByZoneLabel["BAIE"] || [])
    .map(t => t.crbee_positionmm)
    .filter(y => typeof y === "number" && y > 0 + TOL && y < H - TOL)
    .sort((a, b) => a - b);

  const yImposte = traversesBaie.length ? traversesBaie[0] : null;

  // --- 2) Y-cuts = toutes les traverses + bords
  const Ys = [0, H];
  for (const z of Object.keys(traversesEntitiesByZoneLabel)) {
    for (const t of (traversesEntitiesByZoneLabel[z] || [])) {
      if (t && typeof t.crbee_positionmm === "number") Ys.push(t.crbee_positionmm);
    }
  }
  const Y = uniqSorted(Ys).filter(y => y >= 0 - TOL && y <= H + TOL);

  // --- 3) Intervalles X selon bande
  function intervalsForBand(y0, y1) {
    const W  = contexte.LARGEUR_BAIE || 0;
    const Lg = contexte.LARGEUR_FIXE_G || 0;
    const Lp = contexte.LARGEUR_PASSAGE || 0;

    // ✅ 1) Si on est dans la zone imposte (au-dessus de la traverse imposte) :
    // vitrage pleine largeur BAIE (comme avant)
    if (yImposte !== null && y0 >= yImposte - TOL) {
      return [{ x0: 0, x1: W }];
    }

    // ✅ 2) Sinon (en dessous de l’imposte, ou pas d’imposte) :
    // vitrages uniquement sur les FIXES (jamais dans PASSAGE)
    const xA0 = 0,       xA1 = Lg;     // FIXE_G
    const xB0 = Lg + Lp, xB1 = W;      // FIXE_D

    const res = [];
    if (xA1 > xA0 + TOL) res.push({ x0: xA0, x1: xA1 }); // FIXE_G si existe
    if (xB1 > xB0 + TOL) res.push({ x0: xB0, x1: xB1 }); // FIXE_D si existe
    return res;
  }

  let idx = 0;

  for (let j = 0; j < Y.length - 1; j++) {
    const y0 = Y[j];
    const y1 = Y[j + 1];
    if (y1 <= y0 + TOL) continue;

    const intervals = intervalsForBand(y0, y1);
    if (!intervals.length) continue;

    for (const interval of intervals) {
      const ix0 = interval.x0;
      const ix1 = interval.x1;

      // X-cuts = montants qui traversent la bande ET qui sont dans l’intervalle
      const xCuts = uniqSorted(
        allMontantsAbs
          .filter(m => m.xAbs > ix0 + TOL && m.xAbs < ix1 - TOL)
          .filter(m => coupeVitrageBande(
            { crbee_porteedebutmm: m.y0, crbee_porteefinmm: m.y1 },
            y0,
            y1,
            H
          ))
          .map(m => m.xAbs)
      );

      const X = [ix0, ...xCuts, ix1].sort((a, b) => a - b);
      if (X.length < 2) continue;

      for (let i = 0; i < X.length - 1; i++) {
        const x0b = X[i];
        const x1b = X[i + 1];
        if (x1b <= x0b + TOL) continue;

        const celluleKey = `BAIE|C${i}|B${j}|#${idx++}`;

        // Parclose G (V)
        await Xrm.WebApi.createRecord("crbee_profilchassis", {
          "crbee_Chassis@odata.bind": `/crbee_chassisfabriques(${chassisId})`,
          "crbee_Profil@odata.bind": `/crbee_produits(${produitParcloseId})`,
		  "crbee_Commandeproduction@odata.bind":`/crbee_commandedachats(${commandeProductionId})`,
          "crbee_role": ROLE.PARCLOSE,
          "crbee_zone": zoneValueBaie,
          "crbee_orientation": ORI.V,
          "crbee_cote": COTE.G,
          "crbee_cellulekey": celluleKey,
          "crbee_positionmm": x0b,
          "crbee_positionmmdebut": y0,
          "crbee_positionmmfin": y1,
          "crbee_quantite": 1,
          "crbee_coupe": 745350001
        });

        // Parclose D (V)
        await Xrm.WebApi.createRecord("crbee_profilchassis", {
          "crbee_Chassis@odata.bind": `/crbee_chassisfabriques(${chassisId})`,
          "crbee_Profil@odata.bind": `/crbee_produits(${produitParcloseId})`,
		  "crbee_Commandeproduction@odata.bind":`/crbee_commandedachats(${commandeProductionId})`,
          "crbee_role": ROLE.PARCLOSE,
          "crbee_zone": zoneValueBaie,
          "crbee_orientation": ORI.V,
          "crbee_cote": COTE.D,
          "crbee_cellulekey": celluleKey,
          "crbee_positionmm": x1b,
          "crbee_positionmmdebut": y0,
          "crbee_positionmmfin": y1,
          "crbee_quantite": 1,
          "crbee_coupe": 745350001
        });

        // Parclose B (H)
        await Xrm.WebApi.createRecord("crbee_profilchassis", {
          "crbee_Chassis@odata.bind": `/crbee_chassisfabriques(${chassisId})`,
          "crbee_Profil@odata.bind": `/crbee_produits(${produitParcloseId})`,
		  "crbee_Commandeproduction@odata.bind":`/crbee_commandedachats(${commandeProductionId})`,
          "crbee_role": ROLE.PARCLOSE,
          "crbee_zone": zoneValueBaie,
          "crbee_orientation": ORI.H,
          "crbee_cote": COTE.B,
          "crbee_cellulekey": celluleKey,
          "crbee_positionmm": y0,
          "crbee_positionmmdebut": x0b,
          "crbee_positionmmfin": x1b,
          "crbee_quantite": 1,
          "crbee_coupe": 745350001
        });

        // Parclose H (H)
        await Xrm.WebApi.createRecord("crbee_profilchassis", {
          "crbee_Chassis@odata.bind": `/crbee_chassisfabriques(${chassisId})`,
          "crbee_Profil@odata.bind": `/crbee_produits(${produitParcloseId})`,
		  "crbee_Commandeproduction@odata.bind":`/crbee_commandedachats(${commandeProductionId})`,
          "crbee_role": ROLE.PARCLOSE,
          "crbee_zone": zoneValueBaie,
          "crbee_orientation": ORI.H,
          "crbee_cote": COTE.H,
          "crbee_cellulekey": celluleKey,
          "crbee_positionmm": y1,
          "crbee_positionmmdebut": x0b,
          "crbee_positionmmfin": x1b,
          "crbee_quantite": 1,
          "crbee_coupe": 745350001
        });
      }
    }
  }
}




  Xrm.Navigation.openAlertDialog({
    text: "OK : Profils + parcloses générés. FIX montants filants: BAIE_TRAVERSE_H coupée + vitrages recoupés comme avant."
  });

  formContext.data.refresh();
}
