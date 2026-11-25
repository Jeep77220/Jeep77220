#!/usr/bin/env python3
import argparse
import csv
import os
import re
import shutil
import unicodedata
import zipfile
from urllib.parse import quote
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
    m = re.fullmatch(r'(\d{4})-(\d{1,2})-(\d{1,2})', s)
    if m:
        y, mo, d = map(int, m.groups())
        try:
            return date(y, mo, d), s
        except ValueError:
            return None, s
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


def read_grid(path):
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
    rows = []
    for ri in sorted(grid):
        row = grid[ri]
        rows.append([str(row.get(ci) or '') for ci in range(max(row) + 1)] if row else [])
    return rows


def read_calendar(path):
    items = []
    for row in read_grid(path)[2:]:
        row += [''] * (7 - len(row))
        des = row[0].strip()
        if not des:
            continue
        items.append({
            'designation': des,
            'fabricant': row[1].strip(),
            'ref': row[2].strip(),
            'serie': row[3].strip(),
            'periodicite': row[4].strip(),
            'dernier_raw': row[5].strip(),
            'validite_raw': row[6].strip(),
        })
    return items


def slugify(s):
    s = unicodedata.normalize('NFKD', s.lower())
    s = ''.join(c for c in s if not unicodedata.combining(c))
    s = re.sub(r'[^a-z0-9]+', '-', s).strip('-')
    return s or 'instrument'


def folder_name(s):
    s = unicodedata.normalize('NFKC', s)
    s = re.sub(r'[\\/:*?"<>|\s]+', '-', s).strip('-')
    return s or 'instrument'


def ident_of(it):
    for cand in (it['serie'], it['ref']):
        if cand and cand.upper() != 'N/A':
            return folder_name(cand)
    return slugify(it['designation'])


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
    '<fonts count="6">'
    '<font><sz val="11"/><name val="Calibri"/></font>'
    '<font><b/><sz val="11"/><name val="Calibri"/></font>'
    '<font><b/><color rgb="FF9C0006"/><sz val="11"/><name val="Calibri"/></font>'
    '<font><b/><color rgb="FF9C6500"/><sz val="11"/><name val="Calibri"/></font>'
    '<font><color rgb="FF006100"/><sz val="11"/><name val="Calibri"/></font>'
    '<font><u/><color rgb="FF0563C1"/><sz val="11"/><name val="Calibri"/></font>'
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
    '<cellXfs count="6">'
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
    '<xf numFmtId="0" fontId="1" fillId="5" borderId="0" xfId="0" applyFont="1" applyFill="1"/>'
    '<xf numFmtId="0" fontId="2" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>'
    '<xf numFmtId="0" fontId="3" fillId="3" borderId="0" xfId="0" applyFont="1" applyFill="1"/>'
    '<xf numFmtId="0" fontId="4" fillId="4" borderId="0" xfId="0" applyFont="1" applyFill="1"/>'
    '<xf numFmtId="0" fontId="5" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
    '</cellXfs>'
    '</styleSheet>'
)

STYLE = {'DÉPASSÉ': 2, 'ALERTE': 3, 'OK': 4, 'INDÉTERMINÉ': 0}


def write_xlsx(path, title, header, rows, row_styles, hyperlinks=None, color_from=6):
    hyperlinks = hyperlinks or []

    def row_xml(r, values, styles):
        cells = ''.join(
            xcell(f'{col_letter(ci)}{r}', v, s=styles[ci])
            for ci, v in enumerate(values)
        )
        return f'<row r="{r}">{cells}</row>'

    body = [row_xml(1, [title], [1])]
    body.append(row_xml(2, header, [1] * len(header)))
    last_data = len(rows[0]) - (1 if hyperlinks else 0) if rows else 0
    for i, (values, st) in enumerate(zip(rows, row_styles), start=3):
        styles = [0] * len(values)
        for ci in range(color_from, last_data):
            styles[ci] = st
        if hyperlinks:
            styles[-1] = 5
        body.append(row_xml(i, values, styles))
    hyper = ''
    if hyperlinks:
        hyper = '<hyperlinks>' + ''.join(
            f'<hyperlink ref="{ref}" r:id="{rid}"/>' for rid, ref, _ in hyperlinks
        ) + '</hyperlinks>'
    sheet = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'<sheetData>{"".join(body)}</sheetData>{hyper}'
        '</worksheet>'
    )
    sheet_rels = ''
    if hyperlinks:
        sheet_rels = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            + ''.join(
                f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="{escape(quote(target, safe="/"), {chr(34): "&quot;"})}" TargetMode="External"/>'
                for rid, _, target in hyperlinks
            )
            + '</Relationships>'
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
    if sheet_rels:
        parts['xl/worksheets/_rels/sheet1.xml.rels'] = sheet_rels
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        for name, data in parts.items():
            z.writestr(name, data)


def auto_block(it, today):
    jours = 'indéterminé' if it['jours'] is None else str(it['jours'])
    return (
        f'{AUTO_BEGIN}\n'
        '*Section générée automatiquement par `scripts/alerte_etalonnage.py` — ne pas modifier entre les balises.*\n'
        '\n'
        f"- **Désignation** : {it['designation']}\n"
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
    return f"""# Fiche de suivi — {it['designation']} (n° {it['serie'] or it['ref']})

{auto_block(it, today)}

## Historique des étalonnages
À compléter directement dans ce fichier à chaque étalonnage (journal global : `../../registre_suivi_etalonnages.xlsx`).

| Date | Prestataire / laboratoire | N° certificat | Certificat (dossier `certificats/`) | Prochaine échéance | Conforme | Remarques |
|---|---|---|---|---|---|---|
| {dernier} |  |  |  | {validite} |  |  |

## Rapport d'étalonnage
- [Dernier rapport d'étalonnage](rapport_etalonnage.md) — à compléter directement à chaque étalonnage.

## Documents techniques et manuels
- Dossier [`documents_techniques/`](documents_techniques/) : manuels, notices constructeur, fiches techniques, schémas (à déposer ici).

## Certificats
- Dossier [`certificats/`](certificats/) : certificats d'étalonnage de cet instrument (laboratoire COFRAC ou constructeur).

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
    return f"""# Rapport d'étalonnage — {it['designation']} (n° {it['serie'] or it['ref']})

À compléter directement dans ce fichier à chaque étalonnage.

- **Instrument** : {it['designation']}
- **Fabricant / référence** : {it['fabricant']} / {it['ref']}
- **Numéro de série** : {it['serie']}
- **Date d'étalonnage** :
- **Prestataire / laboratoire d'étalonnage** :
- **N° de certificat d'étalonnage** :
- **Certificat (PDF)** : à déposer dans `certificats/`
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
    out.append('Source : `Calendrier des étalonnages.xlsx` (colonne « Alerte »), script : `scripts/alerte_etalonnage.py`.')
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
    out.append(f'Généré le {today.isoformat()} — seuil d\'alerte : {seuil} jours. Dossiers nommés par n° de série (ou réf. fabricant).')
    out.append('')
    out.append('| N° série / réf | Instrument | Date de validité | Jours restants | Statut | Dossier |')
    out.append('|---|---|---|---|---|---|')
    ordered = sorted(items, key=lambda it: (it['validite'] is None, it['validite'] or date.max))
    for it in ordered:
        slug = slugs[id(it)]
        jours = '—' if it['jours'] is None else str(it['jours'])
        ident = it['serie'] if it['serie'] and it['serie'].upper() != 'N/A' else it['ref']
        out.append(
            f"| {ident} | {it['designation']} | {fmt_date(it['validite'], it['validite_raw'])} | "
            f"{jours} | **{it['statut']}** | [instruments/{slug}/]({slug}/) |"
        )
    out.append('')
    current = set(slugs.values())
    stale = sorted(
        p.name for p in path.parent.iterdir()
        if p.is_dir() and p.name not in current and (p / 'fiche_suivi.md').exists()
    )
    if stale:
        out.append("## Hors calendrier (conservés pour l'historique)")
        out.append('')
        for s in stale:
            out.append(f'- [{s}]({s}/fiche_suivi.md) — retiré du calendrier, archivé ici (annexe : [{s}/annexe/]({s}/annexe/))')
        out.append('')
    path.write_text('\n'.join(out), encoding='utf-8')


def migrate_folders(instruments_dir, items, old_names, new_names):
    mapping = {}
    for old, new in zip(old_names, new_names):
        mapping[old] = new
        src = instruments_dir / old
        dst = instruments_dir / new
        if src.exists() and not dst.exists():
            os.rename(src, dst)
    return mapping


def relocate_certificats(dossier, items, slugs):
    cert_dir = dossier / 'certificats'
    if not cert_dir.exists():
        return
    registre = dossier / 'registre_suivi_etalonnages.csv'
    reg_text = registre.read_text(encoding='utf-8-sig') if registre.exists() else ''
    for cert in sorted(cert_dir.glob('*.pdf')):
        targets = []
        for it in items:
            s = it['serie']
            if s and s.upper() != 'N/A' and re.search(r'(?<![0-9A-Za-z])' + re.escape(s) + r'(?![0-9A-Za-z])', cert.name):
                targets.append(slugs[id(it)])
        if not targets:
            token = cert.name.split('_')[1] if '_' in cert.name else ''
            for it in items:
                des = slugify(it['designation'])
                if token and (token in des or des in token):
                    targets.append(slugs[id(it)])
        if not targets:
            targets = [slugs[id(items[0])]]
        for slug in dict.fromkeys(targets):
            cdir = dossier / 'instruments' / slug / 'certificats'
            cdir.mkdir(parents=True, exist_ok=True)
            shutil.copy2(cert, cdir / cert.name)
        if reg_text:
            reg_text = reg_text.replace(f'certificats/{cert.name}', f'instruments/{targets[0]}/certificats/{cert.name}')
        cert.unlink()
    if reg_text and registre.exists():
        registre.write_text(reg_text, encoding='utf-8-sig')
    lisezmoi = cert_dir / 'LISEZMOI.txt'
    if lisezmoi.exists():
        lisezmoi.unlink()
    try:
        cert_dir.rmdir()
    except OSError:
        pass


def rewrite_paths(dossier, mapping):
    registre = dossier / 'registre_suivi_etalonnages.csv'
    if registre.exists():
        text = registre.read_text(encoding='utf-8-sig')
        for old, new in mapping.items():
            text = text.replace(f'instruments/{old}/', f'instruments/{new}/')
        registre.write_text(text, encoding='utf-8-sig')
    instruments = dossier / 'instruments'
    if instruments.exists():
        for md in instruments.rglob('*.md'):
            text = md.read_text(encoding='utf-8')
            new_text = text.replace('](<../../certificats/', '](<certificats/').replace('`../../certificats/`', '`certificats/`')
            new_text = new_text.replace('registre_suivi_etalonnages.csv', 'registre_suivi_etalonnages.xlsx')
            for old, new in mapping.items():
                new_text = new_text.replace(f'instruments/{old}/', f'instruments/{new}/')
            if new_text != text:
                md.write_text(new_text, encoding='utf-8')


def build_annexes(dossier, slugs):
    instruments = dossier / 'instruments'
    index = [
        '# Annexes audit — une annexe par instrument',
        '',
        "Chaque dossier `instruments/<N° de série>/annexe/` regroupe tout ce qui concerne l'instrument pour un audit :",
        "fiche de suivi, rapport d'étalonnage, relevés FO 202 / FO 317, documents techniques (manuels) et certificats.",
        '',
        'Régénéré automatiquement par `scripts/alerte_etalonnage.py` (copies à jour des fichiers vivants).',
        '',
        '| N° série / réf | Instrument | Annexe |',
        '|---|---|---|',
    ]
    for idir in sorted(p for p in instruments.iterdir() if p.is_dir() and (p / 'fiche_suivi.md').exists()):
        slug = idir.name
        annexe = idir / 'annexe'
        if annexe.exists():
            shutil.rmtree(annexe)
        d01 = annexe / '01 - Fiche de suivi FO 202'
        d02 = annexe / '02 - Relevés étalonnage FO 317 et certificats laboratoire'
        d03 = annexe / '03 - Documentation technique de l\'appareil'
        d01.mkdir(parents=True)
        (d02 / 'certificats').mkdir(parents=True)
        d03.mkdir(parents=True)
        (annexe / '00 - LISEZMOI.txt').write_text(
            "Annexe audit de l'instrument. Ordre : 01 fiche de suivi FO 202, "
            "02 relevés d'étalonnage FO 317 / rapports de laboratoire et certificats, "
            "03 documentation technique (manuels). Copie régénérée automatiquement par scripts/alerte_etalonnage.py.\n",
            encoding='utf-8',
        )
        shutil.copy2(idir / 'fiche_suivi.md', d01 / 'fiche_suivi.md')
        for pdf in idir.glob('*.pdf'):
            shutil.copy2(pdf, d01 / pdf.name)
        rapport = idir / 'rapport_etalonnage.md'
        if rapport.exists():
            shutil.copy2(rapport, d02 / 'rapport_etalonnage.md')
        cdir = idir / 'certificats'
        if cdir.exists():
            for f in cdir.glob('*.pdf'):
                shutil.copy2(f, d02 / 'certificats' / f.name)
        dt = idir / 'documents_techniques'
        if dt.exists():
            for f in dt.iterdir():
                if f.is_file() and f.name != 'LISEZMOI.txt':
                    shutil.copy2(f, d03 / f.name)
        titre = slug
        for line in (idir / 'fiche_suivi.md').read_text(encoding='utf-8').splitlines():
            if line.startswith('# '):
                titre = line.lstrip('# ').strip()
                break
        index.append(f'| {slug} | {titre} | [annexe/]({slug}/annexe/) |')
    index.append('')
    (instruments / 'ANNEXES_AUDIT.md').write_text('\n'.join(index), encoding='utf-8')


REGISTRE_TITRE = 'Registre global des étalonnages ALCO — vert = conforme, orange = conforme avec restriction, rouge = non conforme'


def write_registre_xlsx(path):
    rows = read_grid(path)
    header = rows[1] if len(rows) > 1 else REGISTRE_HEADER.split(';')
    data = rows[2:] if len(rows) > 2 else []
    styles = []
    for r in data:
        r += [''] * (9 - len(r))
        val = r[8].lower()
        if 'non conforme' in val:
            styles.append(2)
        elif 'restriction' in val:
            styles.append(3)
        elif 'conforme' in val:
            styles.append(4)
        else:
            styles.append(0)
    write_xlsx(path, REGISTRE_TITRE, header, data, styles, color_from=0)


def main():
    ap = argparse.ArgumentParser(
        description='Met à jour le calendrier unique (alertes de préemption + liens vers les annexes) et régénère le dossier de suivi.'
    )
    ap.add_argument('--calendrier', default=None, help='Chemin du xlsx (défaut : Calendrier des étalonnages.xlsx à la racine du dossier)')
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

    old_names = []
    used_old = set()
    for it in items:
        base = slugify(it['designation'])
        name = base
        n = 2
        while name in used_old:
            name = f'{base}-{n}'
            n += 1
        used_old.add(name)
        old_names.append(name)

    new_names = []
    used_new = set()
    for it in items:
        base = ident_of(it)
        name = base
        n = 2
        while name in used_new:
            name = f'{base}-{n}'
            n += 1
        used_new.add(name)
        new_names.append(name)

    slugs = {id(it): new_names[i] for i, it in enumerate(items)}

    instruments_dir = dossier / 'instruments'
    instruments_dir.mkdir(exist_ok=True)
    mapping = migrate_folders(instruments_dir, items, old_names, new_names)
    relocate_certificats(dossier, items, slugs)
    rewrite_paths(dossier, mapping)

    (dossier / 'alertes').mkdir(exist_ok=True)

    legacy_csv = dossier / 'calendrier_etalonnages.csv'
    if legacy_csv.exists():
        legacy_csv.unlink()

    header = [
        "Désignation de l'instrument", 'Fabricant', 'Référence fabricant (P/N…)',
        'Numéro de série', "Périodicité d'étalonnage", 'Date dernier étalonnage',
        'Date de validité (préemption)', f'Jours restants (au {today.isoformat()})',
        f'Alerte (préemption ≤ {args.seuil} j)', 'Fiche / annexe (audit)',
    ]
    rows = []
    row_styles = []
    hyperlinks = []
    for i, it in enumerate(items):
        slug = slugs[id(it)]
        ri = i + 3
        rows.append([
            it['designation'], it['fabricant'], it['ref'], it['serie'], it['periodicite'],
            fmt_date(it['dernier'], it['dernier_raw']), fmt_date(it['validite'], it['validite_raw']),
            '' if it['jours'] is None else it['jours'], f"{it['statut']} — {it['detail']}",
            f'Annexe {slug}',
        ])
        row_styles.append(STYLE[it['statut']])
        hyperlinks.append((f'hlink{i + 1}', f'J{ri}', f'instruments/{slug}/annexe/'))
    write_xlsx(
        cal,
        f"Calendrier de suivi des dates d'étalonnage et de validité des outillages au sein d'ALCO — alertes de préemption ≤ {args.seuil} jours (généré le {today.isoformat()})",
        header, rows, row_styles, hyperlinks,
    )
    for legacy in dossier.glob('Calendrier des étalonnages - ALERTES *.xlsx'):
        legacy.unlink()

    write_report(dossier / 'alertes' / f'alertes_preemption_{args.seuil}j.md', items, today, args.seuil)

    registre_xlsx = dossier / 'registre_suivi_etalonnages.xlsx'
    registre_csv = dossier / 'registre_suivi_etalonnages.csv'
    if not registre_xlsx.exists():
        if registre_csv.exists():
            with registre_csv.open(newline='', encoding='utf-8-sig') as f:
                rows = [r for r in csv.reader(f, delimiter=';') if r]
            write_xlsx(registre_xlsx, REGISTRE_TITRE, rows[0], rows[1:], [0] * max(len(rows) - 1, 0), color_from=0)
        else:
            write_xlsx(registre_xlsx, REGISTRE_TITRE, REGISTRE_HEADER.split(';'), [], [], color_from=0)
    if registre_csv.exists():
        registre_csv.unlink()
    write_registre_xlsx(registre_xlsx)

    for it in items:
        slug = slugs[id(it)]
        idir = instruments_dir / slug
        (idir / 'documents_techniques').mkdir(parents=True, exist_ok=True)
        (idir / 'certificats').mkdir(exist_ok=True)
        update_fiche(idir / 'fiche_suivi.md', it, today)
        rapport = idir / 'rapport_etalonnage.md'
        if not rapport.exists():
            rapport.write_text(rapport_template(it), encoding='utf-8')
        lisezmoi = idir / 'documents_techniques' / 'LISEZMOI.txt'
        if not lisezmoi.exists():
            lisezmoi.write_text(DOC_TECH_LISEZMOI, encoding='utf-8')

    write_instruments_index(instruments_dir / 'README.md', items, slugs, today, args.seuil)
    build_annexes(dossier, slugs)

    alerts = [it for it in items if it['statut'] in ('ALERTE', 'DÉPASSÉ')]
    print(f'Dossier : {dossier}')
    print(f'Date de référence : {today.isoformat()} — seuil d\'alerte : {args.seuil} jours — {len(items)} instrument(s)')
    print('')
    for it in sorted(items, key=lambda x: (x['validite'] is None, x['validite'] or date.max)):
        jours = '—' if it['jours'] is None else it['jours']
        ident = it['serie'] if it['serie'] and it['serie'].upper() != 'N/A' else it['ref']
        print(f"[{it['statut']:>12}] {ident:<22} {it['designation']:<40} validité {fmt_date(it['validite'], it['validite_raw']):<12} J-{jours}")
    print('')
    if alerts:
        print(f'ALERTES ({len(alerts)}) :')
        for it in alerts:
            print(f"  - {it['designation']} : {it['detail']} (validité {fmt_date(it['validite'], it['validite_raw'])})")
    else:
        print(f'Aucune alerte à {args.seuil} jours à cette date.')
    print('')
    print('Calendrier unique mis à jour : Calendrier des étalonnages.xlsx (alertes + liens cliquables colonne J)')
    print('Registre unique mis à jour : registre_suivi_etalonnages.xlsx (couleurs conformité)')
    print(f'Régénérés : alertes/alertes_preemption_{args.seuil}j.md, instruments/*/fiche_suivi.md, instruments/README.md, ANNEXES_AUDIT.md, annexes/')


if __name__ == '__main__':
    main()
