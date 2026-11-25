# METRO JP 1 — suivi des étalonnages ALCO

Ce dossier permet de suivre la métrologie des outillages d'ALCO : calendrier des étalonnages,
alertes de préemption à 30 jours, fiches de suivi par instrument, rapports d'étalonnage,
certificats et documents techniques.

Tous les fichiers sont directement éditables (Markdown, CSV, XLSX) : on les complète
sur le support lui-même, aucune conversion ni PDF figé.

## Arborescence

```
METRO JP 1/
├── README.md                                  ← ce fichier
├── Calendrier des étalonnages.xlsx            ← calendrier source (à mettre à jour)
├── Calendrier des étalonnages - ALERTES 30j.xlsx  ← régénéré : colonnes « Jours restants » et « Alerte », lignes colorées
├── calendrier_etalonnages.csv                 ← données normalisées (dates ISO, statut d'alerte)
├── registre_suivi_etalonnages.csv             ← journal global à compléter à chaque étalonnage
├── alertes/
│   └── alertes_preemption_30j.md              ← rapport d'alertes régénéré à chaque lancement
├── instruments/
│   ├── README.md                              ← index de tous les instruments avec leur statut
│   └── <instrument>/
│       ├── fiche_suivi.md                     ← fiche suiveuse (section auto + historique à compléter)
│       ├── rapport_etalonnage.md              ← rapport d'étalonnage à compléter à chaque campagne
│       ├── (fiches FO 202 / FO 317 d'origine déposées à la racine, ex. palmer-atelier)
│       ├── documents_techniques/              ← manuels, notices constructeur, fiches techniques
│       └── annexe/                            ← COPIE AUDIT : fiche + rapport + relevés FO + manuels + certificats
└── ANNEXES_AUDIT.md                           ← index des annexes (une par instrument)
├── certificats/                               ← certificats d'étalonnage (PDF)
├── procedures/                                ← procédure FO 201 et procédures qualité
└── scripts/
    └── alerte_etalonnage.py                    ← moteur de calcul des alertes (Python standard, sans dépendance)
```

## Règle d'alerte (préemption à 30 jours)

- La **date de validité** du calendrier est la date de préemption : date limite avant laquelle l'instrument doit être réétalonné.
- **ALERTE** : préemption dans **30 jours ou moins** (lignes orange dans le xlsx, section ⚠️ du rapport).
- **DÉPASSÉ** : date de validité dépassée, étalonnage en retard (lignes rouges).
- **OK** : au-delà de 30 jours (lignes vertes).
- Le seuil est paramétrable : `--seuil 45` par exemple.

## Utilisation

Après chaque modification du calendrier source (`Calendrier des étalonnages.xlsx`), lancer :

```bash
python3 scripts/alerte_etalonnage.py
```

## Modifier le calendrier (dates, ajout / retrait de matériel)

Le fichier à éditer est **`Calendrier des étalonnages.xlsx`** (calendrier source). Le fichier « - ALERTES 30j.xlsx » est régénéré automatiquement : ne pas l'éditer à la main.

- **Modifier une date** : changez « date dernier étalonnage » / « date de validité » sur la ligne concernée. Formats acceptés : `JJ/MM/AAAA`, `MM/AAAA` (fin de mois), `Mars 2030` (fin de mois), ou numéro de série Excel.
- **Ajouter un matériel** : ajoutez une ligne (désignation, fabricant, réf., n° de série, périodicité, deux dates). Au prochain lancement du script, le dossier `instruments/<instrument>/` est créé automatiquement (fiche, rapport, documents_techniques, annexe) et la ligne apparaît avec son alerte et son lien cliquable.
- **Retirer un matériel** : supprimez sa ligne. Au prochain lancement, il disparaît du calendrier / CSV / rapport d'alertes, mais son dossier `instruments/<...>/` est **conservé** (historique d'audit) et listé « Hors calendrier » dans `instruments/README.md`.

Chaque ligne du calendrier « - ALERTES 30j.xlsx » comporte un lien cliquable (colonne **« Fiche / annexe (audit) »**) vers `instruments/<instrument>/annexe/fiche_suivi.md` : fiche complète (suivi interne, rapport d'étalonnage, relevés FO 202/FO 317, manuels, certificats).

Options : `--seuil 30` (défaut), `--date 2026-12-31` (date de référence, défaut = aujourd'hui),
`--calendrier <chemin>` et `--dossier <chemin>` si le dossier est déplacé.

Le script régénère alors : le xlsx avec alertes colorées, le CSV, le rapport
`alertes/alertes_preemption_30j.md`, l'index des instruments et la **section automatique**
de chaque `fiche_suivi.md` (entre les balises `<!-- DEBUT_AUTO -->` / `<!-- FIN_AUTO -->`).
Tout ce que vous écrivez **hors de ces balises** (historique, remarques, rapports) est conservé.

## À chaque étalonnage

1. Compléter `instruments/<instrument>/rapport_etalonnage.md` (résultats, conformité, prochaine échéance).
2. Ajouter une ligne dans le tableau « Historique des étalonnages » de la `fiche_suivi.md`.
3. Ajouter une ligne dans `registre_suivi_etalonnages.csv`.
4. Déposer le certificat PDF dans `certificats/` (nom conseillé : `AAAA-MM-JJ_<instrument>_<n° certificat>.pdf`).
5. Mettre à jour le calendrier source (dernier étalonnage + date de validité) puis relancer le script.

## Conventions de dates dans le calendrier source

Sont acceptées : `JJ/MM/AAAA`, `MM/AAAA` (interprété en fin de mois), `Mars 2030` (fin de mois),
ou un numéro de série Excel (ex. 45985). Les dates normalisées sont en ISO `AAAA-MM-JJ`.

## Ajouter un instrument

Ajouter une ligne dans `Calendrier des étalonnages.xlsx` puis relancer le script :
le dossier `instruments/<instrument>/` (fiche, rapport, documents techniques) est créé automatiquement.
