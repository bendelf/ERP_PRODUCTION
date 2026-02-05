// Benjamin. D - 27/01/2026
// ERP MIROITERIE DIGNOISE
// Fonction qui permet créer un châssis à fabriquer depuis un matériel
// MODIF : création chassis uniquement si produit Assemblé, sinon création directe profilchassis avec longueur

function creerChassisDepuisMateriel(primaryControl) {

    // 🔑 primaryControl = formContext
    var formContext = primaryControl;

    // 🔹 Récupérer le type de commande depuis le formulaire
    var typeCommande = formContext.getAttribute("crbee_typedecommande")?.getValue();

    // 🔒 Si ce n'est pas le type 745350001 → on ne fait rien
    if (typeCommande !== 745350001) {
        Xrm.Navigation.openAlertDialog({
            text: "Le châssis ne peut être créé que pour une commande de type 'En production'."
        });
        return;
    }

    // Constantes type de produit (crbee_produit.crbee_typedeproduit)
    const TYPE_PRODUIT_PROFIL_SEUL = 745350000;
    const TYPE_PRODUIT_ASSEMBLE = 745350003;

    var subgridMateriel = formContext.getControl("Subgrid_materiel");
    var subgridChassisFab = formContext.getControl("Subgrid_chassis_fabrique");

    if (!subgridMateriel || !subgridMateriel.getGrid()) {
        return;
    }

    var selectedRows = subgridMateriel.getGrid().getSelectedRows();

    if (selectedRows.getLength() === 0) {
        Xrm.Navigation.openAlertDialog({
            text: "Veuillez sélectionner au moins un matériel."
        });
        return;
    }

    selectedRows.forEach(function (row) {

        // 🔹 1️⃣ ID du matériel (seule info venant de la vue)
        var materielId = row.getData().getEntity().getId().replace(/[{}]/g, "");

        // 🔹 2️⃣ Lecture directe Dataverse
        // Ajout : crbee_longueur
        Xrm.WebApi.retrieveRecord(
            "crbee_commande", // table matériel
            materielId,
            "?$select=crbee_chassisgenere,crbee_longueur,_crbee_commandedachat_value,_crbee_produits_value"
        ).then(function (materiel) {

            // 🔒 Déjà généré → on ne fait rien
            if (materiel.crbee_chassisgenere === true) {
                return null;
            }

            // 🔒 Pas de produit → on ne peut pas déterminer le type
            if (!materiel._crbee_produits_value) {
                console.warn("Matériel sans produit, id materiel =", materielId);
                return null;
            }

            // 🔹 2bis) Lecture du produit pour connaître le type
            return Xrm.WebApi.retrieveRecord(
                "crbee_produit",
                materiel._crbee_produits_value,
                "?$select=crbee_typedeproduit"
            ).then(function (produitRec) {

                return { materiel: materiel, typeProduit: produitRec.crbee_typedeproduit };
            });

        }).then(function (ctx) {

            if (!ctx) return null;

            var materiel = ctx.materiel;
            var typeProduit = ctx.typeProduit;

            // =========================
            // CAS 1 : PRODUIT ASSEMBLÉ
            // =========================
            if (typeProduit === TYPE_PRODUIT_ASSEMBLE) {

                // 🔹 3️⃣ Préparation du châssis
                var chassis = {
                    "crbee_Materiel@odata.bind": "/crbee_commandes(" + materielId + ")",
                    "crbee_statut": 745350000
                };

                if (materiel._crbee_commandedachat_value) {
                    chassis["crbee_Commandeproduction@odata.bind"] =
                        "/crbee_commandedachats(" + materiel._crbee_commandedachat_value + ")";
                }

                if (materiel._crbee_produits_value) {
                    chassis["crbee_Produit@odata.bind"] =
                        "/crbee_produits(" + materiel._crbee_produits_value + ")";
                }

                // 🔹 4️⃣ Création du châssis
                return Xrm.WebApi.createRecord("crbee_chassisfabrique", chassis)
                    .then(function (result) {
                        if (!result) return null;

                        var chassisId = result.id.replace(/[{}]/g, "");

                        // 🔹 5️⃣ Mise à jour du matériel (lien vers châssis)
                        return Xrm.WebApi.updateRecord("crbee_commande", materielId, {
                            "crbee_Chassisfabrique@odata.bind": "/crbee_chassisfabriques(" + chassisId + ")",
                            crbee_chassisgenere: true,
                            crbee_statut: 745350001
                        });
                    });
            }

            // =========================
            // CAS 2 : PROFIL SEUL
            // =========================
            if (typeProduit === TYPE_PRODUIT_PROFIL_SEUL) {

                // 🔹 Création directe profilchassis
                var profil = {};

                if (materiel._crbee_commandedachat_value) {
                    profil["crbee_Commandeproduction@odata.bind"] =
                        "/crbee_commandedachats(" + materiel._crbee_commandedachat_value + ")";
                }

                profil["crbee_Profil@odata.bind"] =
                    "/crbee_produits(" + materiel._crbee_produits_value + ")";

                // ✅ Copie longueur matériel -> profil
                profil["crbee_longueur"] = materiel.crbee_longueur;
				profil["crbee_quantite"] = materiel.crbee_quantite;
                profil["crbee_coloris"] = materiel.crbee_coloris;
				
				// Champs renseigné
				profil["crbee_coupe"] = 745350001;

                return Xrm.WebApi.createRecord("crbee_profilchassis", profil)
                    .then(function (resProfil) {

                        // 🔹 Mise à jour du matériel (pas de lien châssis dans ce cas)
                        return Xrm.WebApi.updateRecord("crbee_commande", materielId, {
                            crbee_chassisgenere: true,
                            crbee_statut: 745350001
                        });
                    });
            }

            // Type non géré
            console.warn("Type produit non géré :", typeProduit, "pour materielId =", materielId);
            return null;

        }).catch(function (error) {
            console.error(error.message);
        });
    });

    // 🔄 Rafraîchissement des sous-grilles
    subgridMateriel.refresh();
    subgridChassisFab.refresh();
    subgridMateriel.refresh();
}
