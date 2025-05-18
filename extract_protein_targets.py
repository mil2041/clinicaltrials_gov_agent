import zipfile
import xml.etree.ElementTree as ET
import re

path='data/ctg-studies_CAR.parsed.xlsx'
protein_pattern=re.compile(r'(bcma|mesothelin|claudin(?:\s*18\.2)?|her2|psma|cd\d{2,3})',re.I)

zf=zipfile.ZipFile(path)
strings=[si.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t').text if si.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t') is not None else '' for si in ET.fromstring(zf.read('xl/sharedStrings.xml'))]
root=ET.fromstring(zf.read('xl/worksheets/sheet1.xml'))
ns='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
rows=root.find(ns+'sheetData').findall(ns+'row')
headers=[strings[int(c.find(ns+'v').text)] for c in rows[0].findall(ns+'c')]
col_map={h:i for i,h in enumerate(headers)}

output=[]
for row in rows[1:]:
    cells=row.findall(ns+'c')
    values=[strings[int(c.find(ns+'v').text)] if c.get('t')=='s' else (c.find(ns+'v').text if c.find(ns+'v') is not None else '') for c in cells]
    nct_id=values[col_map['nctId']]
    ec_text=values[col_map['eligibilityCriteria']]
    inc=ec_text.split('Exclusion')[0].split('EXCLUSION')[0].split('exclusion')[0]
    matches=protein_pattern.findall(inc)
    unique_targets=sorted(set(m.lower() for m in matches))
    output.append((nct_id,';'.join(unique_targets)))

with open('protein_targets.tsv','w') as f:
    f.write('NCTID\tprotein_target\n')
    for nct, tgt in output:
        f.write(f"{nct}\t{tgt}\n")
