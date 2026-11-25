# Fiches de suivi FO 202 & calendrier des étalonnages — ALCO inox

## Contenu

- **`calendrier des étalonnages - avec alertes.xlsx`** — le calendrier original + un système d'alerte automatique sur les dates de validité (péremption) :
  - colonne « Date réelle » + colonne « Alerte » (formule avec `TODAY()`, se met à jour toute seule à l'ouverture),
  - mise en forme conditionnelle : **rouge = DÉPASSÉ**, **orange = À PLANIFIER (≤ 60 jours)**, **vert = OK**,
  - onglet **« Alertes en cours »** : liste figée des appareils à planifier / dépassés.
- **`fiches pré-remplies/`** :
  - `FO 202 - fiches de suivi pré-remplies.xlsx` — classeur récapitulatif : onglet « Récapitulatif » + **un onglet par instrument** (30 onglets au total),
  - **26 fiches FO 202 individuelles** (`.docx`), pré-remplies : identité de l'appareil + interventions fusionnées depuis le calendrier, vos FO 202 existantes, les rapports d'étalonnage (FO 317 / rapports) et le certificat des cales étalons INSIZE n° 61138,
  - `suppléments - hors calendrier/` — 4 appareils présents dans vos documents mais **absents du calendrier actuel** (pied à coulisse Roch 9 2360, pied à coulisse Filetta J87847, mètre ruban M4, mètre laser SKIL 886492175) : à réintégrer au calendrier si besoin,
  - `00 - récapitulatif.txt` — synthèse texte de toutes les fiches.

## Comment télécharger

Sur la page du dépôt : bouton **« Code »** (vert) en haut à droite → **« Download ZIP »**.

## METRO JP 1 — dossier de suivi métrologie complet

Le dossier **`METRO JP 1/`** regroupe tout le suivi métrologie ALCO :
- `Calendrier des étalonnages - ALERTES 30j.xlsx` — calendrier + colonnes « Jours restants » / « Alerte (préemption ≤ 30 j) », lignes colorées (orange = alerte, rouge = dépassé, vert = OK) ;
- `alertes/alertes_preemption_30j.md` + `calendrier_etalonnages.csv` — rapport et données générés (seuil 30 jours, modifiable) ;
- `instruments/<instrument>/` — 26 instruments, chacun avec sa `fiche_suivi.md` (historique rempli depuis les FO 202/FO 317), son `rapport_etalonnage.md` et son dossier `documents_techniques/` (manuels déposés) ;
- `certificats/` — certificats d'étalonnage (INSIZE 2007, BEA A22EE08141C, BEA A26EJ02353, certificat usine TROTEC) ;
- `registre_suivi_etalonnages.csv` — journal des interventions ;
- `scripts/alerte_etalonnage.py` — régénère alertes/fiches après chaque mise à jour du calendrier (`python3 scripts/alerte_etalonnage.py`).

**Télécharger** : bouton « Code » → « Download ZIP » (sur cette branche), ou le fichier `METRO JP 1.zip` dans l'espace partagé.
