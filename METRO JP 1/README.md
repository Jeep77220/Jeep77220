# METRO JP 1 — suivi des étalonnages ALCO

Un seul calendrier, un dossier par instrument (nommé par son **n° de série**), une annexe audit par instrument.
Tous les fichiers sont directement éditables (Markdown, CSV, XLSX).

## Arborescence

```
METRO JP 1/
├── README.md                            ← ce fichier
├── Calendrier des étalonnages.xlsx      ← LE SEUL calendrier (source éditable + présentation) : éditez-le
│                                          (dates, ajout/suppression de lignes). Colonnes H/I/J (jours restants,
│                                          alerte ≤ 30 j colorée, lien cliquable « Fiche / annexe ») régénérées
│                                          automatiquement par le script à chaque lancement.
├── registre_suivi_etalonnages.xlsx      ← LE SEUL registre (journal GLOBAL de toutes les interventions, source
│                                          éditable) : une ligne par étalonnage/vérification (date, appareil, n° série,
│                                          prestataire, n° certificat, fichier, validité, conformité).
│                                          Lignes colorées : vert = conforme, orange = conforme avec restriction,
│                                          rouge = non conforme.
├── alertes/
│   └── alertes_preemption_30j.md        ← rapport d'alertes généré : qui est en ALERTE / DÉPASSÉ à J-30
├── instruments/
│   ├── README.md                        ← index des instruments (trié par échéance)
│   ├── ANNEXES_AUDIT.md                 ← index des annexes audit
│   └── <N° de série>/                   ← un dossier par instrument, ex. instruments/BB01684/
│       ├── fiche_suivi.md               ← fiche de suivi (FO 202) : identité + historique à compléter
│       ├── rapport_etalonnage.md        ← rapport d'étalonnage (FO 317 / labo) à compléter
│       ├── documents_techniques/        ← manuels et documentation de l'appareil
│       ├── certificats/                 ← certificats de CET instrument (labo COFRAC / constructeur)
│       └── annexe/                      ← COPIE AUDIT, dans l'ordre :
│           ├── 00 - LISEZMOI.txt
│           ├── 01 - Fiche de suivi FO 202/          (fiche_suivi.md + fiches FO 202 d'origine)
│           ├── 02 - Relevés étalonnage FO 317 et certificats laboratoire/  (rapport_etalonnage.md + certificats)
│           └── 03 - Documentation technique de l'appareil/  (manuels)
├── procedures/                          ← procédures qualité applicables : FO 201 (suivi étalonnages),
│                                          IN36 (contrôle dimensionnel), modèles de formulaires FO 202 / FO 317
└── scripts/
    └── alerte_etalonnage.py             ← met à jour le calendrier (alertes + liens) et régénère le dossier
```

## Règle d'alerte (préemption à 30 jours)

- La **date de validité** est la date de préemption : date limite avant laquelle l'instrument doit être réétalonné.
- **ALERTE** : préemption dans **≤ 30 jours** (ligne orange dans le calendrier, section ⚠️ du rapport `alertes/`).
- **DÉPASSÉ** : date de validité dépassée (ligne rouge).
- **OK** : au-delà (ligne verte). Seuil modifiable (`--seuil 45`).
- Chaque ligne du calendrier a un **lien cliquable** (colonne J « Fiche / annexe (audit) ») vers `instruments/<N° série>/annexe/fiche_suivi.md`.

## Utilisation

Après chaque modification du calendrier, lancer :

```bash
python3 scripts/alerte_etalonnage.py
```

Options : `--seuil 30` (défaut), `--date 2026-12-31` (date de référence), `--calendrier` / `--dossier` si déplacé.

Le script met à jour **le calendrier lui-même** (colonnes H/I/J + liens) et régénère : CSV, rapport `alertes/`, index, fiches (section auto entre balises `DEBUT_AUTO`/`FIN_AUTO` — le reste est conservé) et annexes.

## Modifier le calendrier (dates, ajout / retrait de matériel)

Éditez **`Calendrier des étalonnages.xlsx`** (colonnes A–G) :

- **Modifier une date** : changez « date dernier étalonnage » / « date de validité ». Formats acceptés : `JJ/MM/AAAA`, `AAAA-MM-JJ`, `MM/AAAA` (fin de mois), `Mars 2030` (fin de mois), ou n° de série Excel.
- **Ajouter un matériel** : ajoutez une ligne → son dossier `instruments/<n° série>/` est créé automatiquement au prochain lancement.
- **Retirer un matériel** : supprimez sa ligne → il disparaît du calendrier/rapport, mais son dossier est **conservé** et listé « Hors calendrier » dans `instruments/README.md`.

## À chaque étalonnage

1. Compléter `instruments/<n° série>/rapport_etalonnage.md`.
2. Ajouter une ligne dans l'historique de la `fiche_suivi.md`.
3. Ajouter une ligne dans `registre_suivi_etalonnages.xlsx` (journal global, colonne « conforme » colorée automatiquement).
4. Déposer le certificat PDF dans `instruments/<n° série>/certificats/` (nom : `AAAA-MM-JJ_<n° série>_<n° certificat>.pdf`).
5. Mettre à jour les dates du calendrier, puis relancer le script.

## Ajouter un instrument

Nouvelle ligne dans le calendrier → relancer le script : dossier, fiche, rapport, certificats/, documents_techniques/ et annexe sont créés automatiquement.
