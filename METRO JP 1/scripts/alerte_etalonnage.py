#!/usr/bin/env python3
import argparse
import csv
import re
import unicodedata
import zipfile
from datetime import date, timedelta
from pathlib import Path
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'

AUTO_BEGIN = '<!-- DEBUT_AUTO -->'
AUTO_END = '<!-- FIN_AUTO -->'

MOIS = {
    'janvier': 1, 'janv': 1, 'jan': 1,
    'fevrier': 2, 'fevr': 2, 'feb': 2,
    'mars': 3, 'mar': 3,
    'avril': 4, 'avr': 4, 'apr': 4,
    'mai': 5,
    'juin': 6, 'jun': 6,
    'juillet': 7, 'juil': 7, 'jul': 7,
    'aout': 8,
    'septembre': 9, 'sept': 9, 'sep': 9,
    'octobre': 10, 'oct': 10,
    'novembre': 11, 'nov': 11,
    'decembre': 12, 'dec': 12,
}


def last_day(y, m):
    if m == 12:
        return date(y, 12, 31)
    return date(y, m + 1, 1) - timedelta(days=1)


def parse_date(raw):
    if raw is None:
        return None, ''
    s = str(raw).strip()
    if not s:
        return None, ''
    if re.fullmatch(r'\d+(\.\d+)?', s):
        return date(1899, 12, 30) + timedelta(days=int(float(s))), s
    m = re.fullmatch(r'(\d{1,2})/(\d{1,2})/(\d{4})', s)
    if m:
        d, mo, y = map(int, m.groups())
        try:
            return date(y, mo, d), s
        except ValueError:
            return None, s
    m = re.fullmatch(r'(\d{1,2})/(\d{4})', s)
    if m:
        mo, y = int(m.group(1)), int(m.group(2))
        if 1 <= mo <= 12:
            return last_day(y, mo), s
        return None, s
    m = re.fullmatch(r'([a-zàâäéèêëîïôöùûüç]+)\.?\s+(\d{4})', s, re.IGNORECASE)
    if m and m.group(1).lower() in MOIS:
        return last_day(int(m.group(2)), MOIS[m.group(1).lower()]), s
    return None, s


def col_index(ref):
    m = re.match(r'([A-Z]+)(\d+)', ref)
    c = 0
    for ch in m.group(1):
        c = c * 26 + ord(ch) - 64
    return c - 1, int(m.group(2)) - 1


def read_calendar(path):
    z = zipfile.ZipFile(path)
    ss = []
    if 'xl/sharedStrings.xml' in z.namelist():
        root = ET.fromstring(z.read('xl/sharedStrings.xml'))
        for si in root.findall(NS + 'si'):
            ss.append(''.join(t.text or '' for t in si.iter(NS + 't')))
    sheet = ET.fromstring(z.read('xl/worksheets/sheet1.xml'))
    grid = {}
    for c in sheet.iter(NS + 'c'):
        ci, ri = col_index(c.attrib['r'])
        t = c.attrib.get('t')
        v = c.find(NS + 'v')
        is_el = c.find(NS + 'is')
        if t == 's' and v is not None:
            val = ss[int(v.text)]
        elif t == 'inlineStr' and is_el is not None:
            val = ''.join(x.text or '' for x in is_el.iter(NS + 't'))
        elif v is not None:
            val = v.text
        else:
            val = None
        grid.setdefault(ri, {})[ci] = val
    items = []
    for ri in sorted(grid):
        if ri < 2:
            continue
        row = grid[ri]
        des = str(row.get(0) or '').strip()
        if not des:
            continue
        items.append({
            'designation': des,
            'fabricant': str(row.get(1) or '').strip(),
            'ref': str(row.get(2) or '').strip(),
            'serie': str(row.get(3) or '').strip(),
            'periodicite': str(row.get(4) or '').strip(),
            'dernier_raw': str(row.get(5) or '').strip(),
            'validite_raw': str(row.get(6) or '').strip(),
        })
    return items


def slugify(s):
    s = unicodedata.normalize('NFKD', s.lower())
    s = ''.join(c for c in s if not unicodedata.combining(c))
    s = re.sub(r'[^a-z0-9]+', '-', s).strip('-')
    return s or 'instrument'


def statut(jours, seuil):
    if jours is None:
        return 'INDÉTERMINÉ', 'date de validité illisible'
    if jours < 0:
        return 'DÉPASSÉ', f'préemption dépassée de {-jours} j'
    if jours <= seuil:
        return 'ALERTE', f'préemption dans {jours} j'
    return 'OK', f'préemption dans {jours} j'


def fmt_date(d, raw):
    if d is not None:
        return d.isoformat()
    return raw if raw else '—'


def col_letter(i):
    s = ''
    i += 1
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


def xcell(ref, value, s=0):
    if value is None or value == '':
        return f'<c r="{ref}" s="{s}"/>'
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return f'<c r="{ref}" s="{s}"><v>{value}</v></c>'
    return f'<c r="{ref}" s="{s}" t="inlineStr"><is><t xml:space="preserve">{escape(str(value))}</t></is></c>'


STYLES_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    '<fonts count="5">'
    '<font><sz val="11"/><name val="Calibri"/></font>'
    '<font><b/><sz val="11"/><name val="Calibri"/></font>'
    '<font><b/><color rgb="FF9C0006"/><sz val="11"/><name val="Calibri"/></font>'
    '<font><b/><color rgb="FF9C6500"/><sz val="11"/><name val="Calibri"/></font>'
    '<font><color rgb="FF006100"/><sz val="11"/><name val="Calibri"/></font>'
    '</fonts>'
    '<fills count="6">'
    '<fill><patternFill patternType="none"/></fill>'
    '<fill><patternFill patternType="gray125"/></fill>'
    '<fill><patternFill patternType="solid"><fgColor rgb="FFFFC7CE"/><bgColor rgb="FFFFC7CE"/></patternFill></fill>'
    '<fill><patternFill patternType="solid"><fgColor rgb="FFFFEB9C"/><bgColor rgb="FFFFEB9C"/></patternFill></fill>'
    '<fill><patternFill patternType="solid"><fgColor rgb="FFC6EFCE"/><bgColor rgb="FFC6EFCE"/></patternFill></fill>'
    '<fill><patternFill patternType="solid"><fgColor rgb="FFD9D9D9"/><bgColor rgb="FFD9D9D9"/></patternFill></fill>'
    '</fills>'
    '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    '<cellXfs count="5">'
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
    '<xf numFmtId="0" fontId="1" fillId="5" borderId="0" xfId="0" applyFont="1" applyFill="1"/>'
    '<xf numFmtId="0" fontId="2" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>'
    '<xf numFmtId="0" fontId="3" fillId="3" borderId="0" xfId="0" applyFont="1" applyFill="1"/>'
    '<xf numFmtId="0" fontId="4" fillId="4" borderId="0" xfId="0" applyFont="1" applyFill="1"/>'
    '</cellXfs>'
    '</styleSheet>'
)

STYLE = {'DÉPASSÉ': 2, 'ALERTE': 3, 'OK': 4, 'INDÉTERMINÉ': 0}


def write_xlsx(path, title, header, rows, row_styles):
    def row_xml(r, values, styles):
        cells = ''.join(
            xcell(f'{col_letter(ci)}{r}', v, s=styles[ci])
            for ci, v in enumerate(values)
        )
        return f'<row r="{r}">{cells}</row>'

    body = [row_xml(1, [title], [1])]
    body.append(row_xml(2, header, [1] * len(header)))
    for i, (values, st) in enumerate(zip(rows, row_styles), start=3):
        styles = [0] * len(values)
        for ci in range(6, len(values)):
            styles[ci] = st
        body.append(row_xml(i, values, styles))
    sheet = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f'<sheetData>{"".join(body)}</sheetData>'
        '</worksheet>'
    )
    parts = {
        '[Content_Types].xml': (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
            '</Types>'
        ),
        '_rels/.rels': (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            '</Relationships>'
        ),
        'xl/workbook.xml': (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<sheets><sheet name="Calendrier + alertes" sheetId="1" r:id="rId1"/></sheets>'
            '</workbook>'
        ),
        'xl/_rels/workbook.xml.rels': (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
            '</Relationships>'
        ),
        'xl/styles.xml': STYLES_XML,
        'xl/worksheets/sheet1.xml': sheet,
    }
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        for name, data in parts.items():
            z.writestr(name, data)


def auto_block(it, today):
    jours = 'indéterminé' if it['jours'] is None else str(it['jours'])
    return (
        f'{AUTO_BEGIN}\n'
        '*Section générée automatiquement par `scripts/alerte_etalonnage.py` — ne pas modifier entre les balises.*\n'
        '\n'
        f"- **Fabricant** : {it['fabricant']}\n"
        f"- **Référence fabricant** : {it['ref']}\n"
        f"- **Numéro de série** : {it['serie']}\n"
        f"- **Périodicité d'étalonnage** : {it['periodicite']}\n"
        f"- **Dernier étalonnage** : {fmt_date(it['dernier'], it['dernier_raw'])}\n"
        f"- **Date de validité (préemption)** : {fmt_date(it['validite'], it['validite_raw'])}\n"
        f"- **Jours restants au {today.isoformat()}** : {jours}\n"
        f"- **Statut alerte** : **{it['statut']}** ({it['detail']})\n"
        f'{AUTO_END}'
    )


def fiche_template(it, today):
    dernier = fmt_date(it['dernier'], it['dernier_raw'])
    validite = fmt_date(it['validite'], it['validite_raw'])
    return f"""# Fiche de suivi — {it['designation']}

{auto_block(it, today)}

## Historique des étalonnages
À compléter directement dans ce fichier à chaque étalonnage (journal global : `../../registre_suivi_etalonnages.csv`).

| Date | Prestataire / laboratoire | N° certificat | Certificat (dossier `certificats/`) | Prochaine échéance | Conforme | Remarques |
|---|---|---|---|---|---|---|
| {dernier} |  |  |  | {validite} |  |  |

## Rapport d'étalonnage
- [Dernier rapport d'étalonnage](rapport_etalonnage.md) — à compléter directement à chaque étalonnage.

## Documents techniques et manuels
- Dossier [`documents_techniques/`](documents_techniques/) : manuels, notices constructeur, fiches techniques, schémas (à déposer ici).

## Remarques

"""


def update_fiche(path, it, today):
    block = auto_block(it, today)
    if path.exists():
        content = path.read_text(encoding='utf-8')
        if AUTO_BEGIN in content and AUTO_END in content:
            new = re.sub(
                re.escape(AUTO_BEGIN) + '.*?' + re.escape(AUTO_END),
                lambda m: block,
                content,
                flags=re.S,
            )
            path.write_text(new, encoding='utf-8')
            return
    path.write_text(fiche_template(it, today), encoding='utf-8')


def rapport_template(it):
    return f"""# Rapport d'étalonnage — {it['designation']}

À compléter directement dans ce fichier à chaque étalonnage.

- **Instrument** : {it['designation']}
- **Fabricant / référence** : {it['fabricant']} / {it['ref']}
- **Numéro de série** : {it['serie']}
- **Date d'étalonnage** :
- **Prestataire / laboratoire d'étalonnage** :
- **N° de certificat d'étalonnage** :
- **Certificat (PDF)** : à déposer dans `../../certificats/`
- **Date de validité après étalonnage** :
- **Points de contrôle / résultats obtenus** :
- **Conformité** : conforme / non conforme
- **Écarts constatés et actions correctives** :
- **Prochaine échéance** :

"""


DOC_TECH_LISEZMOI = (
    "Déposer ici les documents techniques de l'instrument : manuel d'utilisation, "
    "notice constructeur, fiche technique, schéma, procédure de vérification interne, etc.\n"
)

REGISTRE_HEADER = (
    'date_etalonnage;instrument;numero_serie;prestataire;numero_certificat;'
    'fichier_certificat;date_validite;prochaine_echeance;conforme;remarques'
)


def write_report(path, items, today, seuil):
    alerts = [it for it in items if it['statut'] == 'ALERTE']
    dep = [it for it in items if it['statut'] == 'DÉPASSÉ']
    ok = [it for it in items if it['statut'] == 'OK']
    ind = [it for it in items if it['statut'] == 'INDÉTERMINÉ']
    ordered = sorted(items, key=lambda it: (it['validite'] is None, it['validite'] or date.max))

    def table(sub):
        lines = [
            '| Instrument | N° série | Périodicité | Dernier étalonnage | Date de validité | Jours restants | Statut |',
            '|---|---|---|---|---|---|---|',
        ]
        for it in sub:
            jours = '—' if it['jours'] is None else str(it['jours'])
            lines.append(
                f"| {it['designation']} | {it['serie']} | {it['periodicite']} | "
                f"{fmt_date(it['dernier'], it['dernier_raw'])} | "
                f"{fmt_date(it['validite'], it['validite_raw'])} | {jours} | **{it['statut']}** |"
            )
        return lines

    out = []
    out.append(f'# Alertes de préemption — échéances à {seuil} jours ou moins')
    out.append('')
    out.append(f'Généré le **{today.isoformat()}** — seuil d\'alerte : **{seuil} jours** (option `--seuil` du script).')
    out.append('Source : `Calendrier des étalonnages.xlsx`, script : `scripts/alerte_etalonnage.py`.')
    out.append('')
    out.append('## Résumé')
    out.append('')
    out.append(f'- **{len(alerts)} ALERTE(S)** : préemption dans ≤ {seuil} jours')
    out.append(f'- **{len(dep)} DÉPASSÉ(S)** : date de validité dépassée — étalonnage en retard')
    out.append(f'- **{len(ok)} OK**')
    if ind:
        out.append(f'- **{len(ind)} INDÉTERMINÉ(S)** : date de validité à corriger dans le calendrier')
    out.append('')
    out.append(f'## ⚠️ Alertes (préemption ≤ {seuil} jours)')
    out.append('')
    if alerts:
        out.extend(table(alerts))
    else:
        out.append(f'Aucun instrument en alerte à cette date (aujourd\'hui : {today.isoformat()}).')
    out.append('')
    out.append('## 🚨 Validités dépassées')
    out.append('')
    if dep:
        out.extend(table(dep))
    else:
        out.append('Aucune validité dépassée.')
    out.append('')
    out.append('## Toutes les échéances (triées par date de validité)')
    out.append('')
    out.extend(table(ordered))
    out.append('')
    path.write_text('\n'.join(out), encoding='utf-8')


def write_instruments_index(path, items, slugs, today, seuil):
    out = []
    out.append('# Index des instruments — fiches de suivi')
    out.append('')
    out.append(f'Généré le {today.isoformat()} — seuil d\'alerte : {seuil} jours. Chaque instrument a sa fiche, son rapport et son dossier de documents techniques.')
    out.append('')
    out.append('| Instrument | N° série | Date de validité | Jours restants | Statut | Fiche |')
    out.append('|---|---|---|---|---|---|')
    ordered = sorted(items, key=lambda it: (it['validite'] is None, it['validite'] or date.max))
    for it in ordered:
        slug = slugs[id(it)]
        jours = '—' if it['jours'] is None else str(it['jours'])
        out.append(
            f"| {it['designation']} | {it['serie']} | {fmt_date(it['validite'], it['validite_raw'])} | "
            f"{jours} | **{it['statut']}** | [fiche_suivi.md]({slug}/fiche_suivi.md) |"
        )
    out.append('')
    path.write_text('\n'.join(out), encoding='utf-8')


def main():
    ap = argparse.ArgumentParser(
        description='Calcule les alertes de préemption du calendrier d\'étalonnage ALCO et régénère le dossier de suivi.'
    )
    ap.add_argument('--calendrier', default=None, help='Chemin du xlsx source (défaut : Calendrier des étalonnages.xlsx à la racine du dossier)')
    ap.add_argument('--dossier', default=None, help='Racine du dossier de suivi (défaut : dossier parent de scripts/)')
    ap.add_argument('--seuil', type=int, default=30, help='Seuil d\'alerte en jours (défaut : 30)')
    ap.add_argument('--date', default=None, help='Date de référence AAAA-MM-JJ (défaut : aujourd\'hui)')
    args = ap.parse_args()

    dossier = Path(args.dossier).resolve() if args.dossier else Path(__file__).resolve().parent.parent
    cal = Path(args.calendrier).resolve() if args.calendrier else dossier / 'Calendrier des étalonnages.xlsx'
    today = date.fromisoformat(args.date) if args.date else date.today()

    items = read_calendar(cal)
    for it in items:
        it['dernier'], _ = parse_date(it['dernier_raw'])
        it['validite'], _ = parse_date(it['validite_raw'])
        it['jours'] = (it['validite'] - today).days if it['validite'] else None
        it['statut'], it['detail'] = statut(it['jours'], args.seuil)

    slugs = {}
    used = set()
    for it in items:
        base = slugify(it['designation'])
        slug = base
        n = 2
        while slug in used:
            slug = f'{base}-{n}'
            n += 1
        used.add(slug)
        slugs[id(it)] = slug

    (dossier / 'alertes').mkdir(exist_ok=True)
    (dossier / 'instruments').mkdir(exist_ok=True)

    csv_path = dossier / 'calendrier_etalonnages.csv'
    with csv_path.open('w', newline='', encoding='utf-8-sig') as f:
        w = csv.writer(f, delimiter=';')
        w.writerow([
            'designation', 'fabricant', 'reference_fabricant', 'numero_serie', 'periodicite',
            'dernier_etalonnage', 'date_validite', 'jours_restants', 'statut', 'detail',
            'fiche_suivi',
        ])
        for it in items:
            w.writerow([
                it['designation'], it['fabricant'], it['ref'], it['serie'], it['periodicite'],
                fmt_date(it['dernier'], it['dernier_raw']), fmt_date(it['validite'], it['validite_raw']),
                '' if it['jours'] is None else it['jours'], it['statut'], it['detail'],
                f"instruments/{slugs[id(it)]}/fiche_suivi.md",
            ])

    header = [
        "Désignation de l'instrument", 'Fabricant', 'Référence fabricant (P/N…)',
        'Numéro de série', "Périodicité d'étalonnage", 'Date dernier étalonnage',
        'Date de validité (préemption)', f'Jours restants (au {today.isoformat()})',
        f'Alerte (préemption ≤ {args.seuil} j)',
    ]
    rows = []
    row_styles = []
    for it in items:
        rows.append([
            it['designation'], it['fabricant'], it['ref'], it['serie'], it['periodicite'],
            fmt_date(it['dernier'], it['dernier_raw']), fmt_date(it['validite'], it['validite_raw']),
            '' if it['jours'] is None else it['jours'], f"{it['statut']} — {it['detail']}",
        ])
        row_styles.append(STYLE[it['statut']])
    write_xlsx(
        dossier / f'Calendrier des étalonnages - ALERTES {args.seuil}j.xlsx',
        f"Calendrier de suivi des dates d'étalonnage et de validité des outillages au sein d'ALCO — alertes de préemption ≤ {args.seuil} jours (généré le {today.isoformat()})",
        header, rows, row_styles,
    )

    write_report(dossier / 'alertes' / f'alertes_preemption_{args.seuil}j.md', items, today, args.seuil)

    registre = dossier / 'registre_suivi_etalonnages.csv'
    if not registre.exists():
        registre.write_text(REGISTRE_HEADER + '\n', encoding='utf-8-sig')

    for it in items:
        slug = slugs[id(it)]
        idir = dossier / 'instruments' / slug
        (idir / 'documents_techniques').mkdir(parents=True, exist_ok=True)
        update_fiche(idir / 'fiche_suivi.md', it, today)
        rapport = idir / 'rapport_etalonnage.md'
        if not rapport.exists():
            rapport.write_text(rapport_template(it), encoding='utf-8')
        lisezmoi = idir / 'documents_techniques' / 'LISEZMOI.txt'
        if not lisezmoi.exists():
            lisezmoi.write_text(DOC_TECH_LISEZMOI, encoding='utf-8')

    write_instruments_index(dossier / 'instruments' / 'README.md', items, slugs, today, args.seuil)

    alerts = [it for it in items if it['statut'] in ('ALERTE', 'DÉPASSÉ')]
    print(f'Dossier : {dossier}')
    print(f'Date de référence : {today.isoformat()} — seuil d\'alerte : {args.seuil} jours — {len(items)} instrument(s)')
    print('')
    for it in sorted(items, key=lambda x: (x['validite'] is None, x['validite'] or date.max)):
        jours = '—' if it['jours'] is None else it['jours']
        print(f"[{it['statut']:>12}] {it['designation']:<45} validité {fmt_date(it['validite'], it['validite_raw']):<12} J-{jours}")
    print('')
    if alerts:
        print(f'ALERTES ({len(alerts)}) :')
        for it in alerts:
            print(f"  - {it['designation']} : {it['detail']} (validité {fmt_date(it['validite'], it['validite_raw'])})")
    else:
        print(f'Aucune alerte à {args.seuil} jours à cette date.')
    print('')
    print('Fichiers régénérés : calendrier_etalonnages.csv, Calendrier des étalonnages - ALERTES.xlsx,')
    print(f'alertes/alertes_preemption_{args.seuil}j.md, instruments/*/fiche_suivi.md, instruments/README.md')


if __name__ == '__main__':
    main()
