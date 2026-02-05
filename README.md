# ERP_PRODUCTION

Scripts Power Apps / Dataverse pour la gestion de la production d'aluminium.

## Organisation

```
.
├── README.md
└── scripts
    ├── calculs
    │   └── calculerDebitsProfils_v4_run.js
    └── generation
        ├── creerChassisDepuisMateriel_v3.js
        └── genererProfilsDepuisNomenclature_v6_run.js
```

## Rôle des scripts

- `scripts/calculs/calculerDebitsProfils_v4_run.js`
  - Calcul des longueurs (débits) des profils, traverses, montants et parcloses à partir des profils du châssis et de la nomenclature.
- `scripts/generation/creerChassisDepuisMateriel_v3.js`
  - Création d'un châssis fabriqué (ou d'un profilchassis direct) depuis un matériel, selon le type de produit.
- `scripts/generation/genererProfilsDepuisNomenclature_v6_run.js`
  - Génération complète des profils d'un châssis à partir de la nomenclature (cadre, montants, traverses, parcloses).

## Utilisation (Power Apps Model-Driven)

Ces scripts sont conçus pour être attachés à des boutons ou des événements de formulaire via les bibliothèques JavaScript Dataverse (Xrm). Assure-toi d'importer le fichier correspondant dans la bibliothèque du formulaire et d'appeler la fonction principale exposée par chaque script.
