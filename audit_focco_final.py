#!/usr/bin/env python3
"""
Focco ERP Price Audit vs Mother Tables (v2)
Compares prices in Focco 'Itens - Regras de Preço' against 3 mother price tables.
"""

import openpyxl
import re
from collections import defaultdict

# ============================================================
# CONFIG
# ============================================================
FOCCO_PATH = '/home/user/claude/PIU_MOBILE_HDA_CORRIGIDO7.xlsx'
TABLE_PIU_MOBILE = '/root/.claude/uploads/311cc8c5-e636-5ebe-9ac9-ac0347d05c63/adf55a4c-TABELA_DE_PRE_O__PIU_MOBILE_2026__HD.xlsx'
TABLE_PIURB      = '/root/.claude/uploads/311cc8c5-e636-5ebe-9ac9-ac0347d05c63/8042adc0-TABELA_DE_PRE_O___PIURB_2026__HD.xlsx'
TABLE_ITALO      = '/root/.claude/uploads/311cc8c5-e636-5ebe-9ac9-ac0347d05c63/9618cc7b-TABELA_DE_PRE_O_ITALO_TROPICAL__102025__PIU_MOBILE.xlsx'

# Items to skip (already corrected)
SKIP_CODES = {
    '22362', '17075', '17076',
    '22370', '22371', '22372', '22373', '22374', '22375', '22376', '22377',  # Xique Xique
    '22066', '22124', '22239', '22312', '22313', '22314', '22315', '22325',  # FORNECIDO bugs
    '22213',  # ALOFI
    '17073', '17074',  # MESC
}

# ============================================================
# GRADE NORMALIZATION
# All fabric grade variants -> canonical name
# ============================================================
GRADE_NORMALIZE = {
    'FX FORNECIDO': 'FX FORNECIDO', 'FORNECIDO': 'FX FORNECIDO', 'FORN': 'FX FORNECIDO',
    'FX FORN RB': 'FX FORNECIDO', 'FX FORN. RB': 'FX FORNECIDO', 'TC FORN. RB': 'FX FORNECIDO',
    'COURO FORN.': 'COURO FORN', 'COUR FORN. ': 'COURO FORN', 'COURO FORN': 'COURO FORN',
    'TECIDO FORNECIDO': 'FX FORNECIDO', 'TECIDO_FORNECIDO': 'FX FORNECIDO',
    'TEC_FORNECIDO': 'FX FORNECIDO',
    'FX A': 'FX A', 'TECIDO A': 'FX A', 'TEC A': 'FX A', 'TECIDO_A': 'FX A',
    'FX B': 'FX B', 'TECIDO B': 'FX B', 'TEC B': 'FX B', 'TECIDO_B': 'FX B',
    'FX C': 'FX C', 'TECIDO C': 'FX C', 'TEC C': 'FX C', 'TECIDO_C': 'FX C',
    'FX D': 'FX D', 'TECIDO D': 'FX D', 'TEC D': 'FX D', 'TECIDO_D': 'FX D', 'TECIDO D ': 'FX D',
    'FX E': 'FX E', 'TECIDO E': 'FX E', 'TEC E': 'FX E', 'TECIDO_E': 'FX E',
    'FX F': 'FX F', 'TECIDO F': 'FX F', 'TEC F': 'FX F', 'TECIDO_F': 'FX F', 'F': 'FX F',
    'FX G': 'FX G', 'TECIDO G': 'FX G', 'TEC G': 'FX G', 'TECIDO_G': 'FX G', 'G': 'FX G',
    'FX H': 'FX H', 'TECIDO H': 'FX H', 'TEC H': 'FX H', 'TECIDO_H': 'FX H',
    'FX I': 'FX I', 'TECIDO I': 'FX I', 'TEC I': 'FX I', 'TECIDO_I': 'FX I',
    'FX J': 'FX J', 'TECIDO J': 'FX J', 'TEC J': 'FX J', 'TECIDO_J': 'FX J',
    'FX J / N': 'FX J', 'FX J/N': 'FX J', 'TECIDO J/N': 'FX J', 'TECIDO J ': 'FX J',
    'J/N': 'FX J', 'TECIDO_N': 'FX J',
    'FX R': 'FX R', 'TECIDO R': 'FX R', 'TECIDO_R': 'FX R', 'R': 'FX R',
    'FX V': 'FX V', 'TECIDO V': 'FX V', 'TECIDO_V': 'FX V', 'V': 'FX V',
    'FX KNIT': 'FX KNIT', 'TECIDO KNIT': 'FX KNIT',
    'CR VQ': 'CR VQ', 'CR FZ': 'CR FZ',
}


# ============================================================
# NAME UTILITIES
# ============================================================
PREFIXES = ['SOFA CAMA ', 'SOFA ', 'POLTRONA ', 'PUFF MESA ', 'PUFF ',
            'CHAISE ', 'CAMA ', 'BANCO ', 'CANTO ', 'MESA LATERAL ', 'MESA ',
            'CADEIRA ', 'CABECEIRA ', 'BASE ', 'RECAMIER ', 'APOIO ', 'BANDEJA ',
            'ENCOSTO ']

# Type abbreviations used in mother table product names
MOTHER_TYPE_HINTS = {
    'POL': ['POLTRONA'],
    'PUFF': ['PUFF'],
    'CH': ['CHAISE'],
    'CAMA': ['CAMA'],
    'BANCO': ['BANCO'],
    'SB': ['SOFA', 'CHAISE'],   # sem braço
    'CTO': ['SOFA'],            # canto
    'CABECEIRA': ['CABECEIRA'],
    'BASE': ['BASE'],
}

def strip_focco_prefix(name):
    """Strip furniture type prefix from focco name to get core product name."""
    name = name.upper().strip()
    for p in PREFIXES:
        if name.startswith(p):
            return name[len(p):].strip(), p.strip()
    return name, ''

def focco_type(name):
    """Determine the furniture type from focco name."""
    name = name.upper().strip()
    if name.startswith('SOFA CAMA '): return 'SOFA CAMA'
    if name.startswith('SOFA '): return 'SOFA'
    if name.startswith('POLTRONA '): return 'POL'
    if name.startswith('PUFF MESA '): return 'PUFF'
    if name.startswith('PUFF '): return 'PUFF'
    if name.startswith('CHAISE '): return 'CH'
    if name.startswith('CAMA '): return 'CAMA'
    if name.startswith('BANCO '): return 'BANCO'
    if name.startswith('CANTO '): return 'CANTO'
    if name.startswith('CABECEIRA '): return 'CABECEIRA'
    if name.startswith('BASE '): return 'BASE'
    return ''

def name_core_words(name):
    """Get the first 1-2 significant words of a name."""
    words = name.split()
    sig = [w for w in words if re.search(r'[A-Z]', w)]
    return sig[:2] if len(sig) >= 2 else sig


def names_match(focco_name, mother_name):
    """
    Check if focco item matches a mother table product.
    Strategy: strip focco prefix, check if mother starts with the same words.
    Also tries to match product type (POL, PUFF, CH, etc.)
    """
    focco_core, focco_prefix = strip_focco_prefix(focco_name)
    mother_upper = mother_name.upper().strip()
    ftype = focco_type(focco_name)

    # Check if mother name starts with focco core
    if mother_upper.startswith(focco_core):
        return True

    # Check first 1-2 words match
    fw = name_core_words(focco_core)
    mw = name_core_words(mother_upper)
    if len(fw) >= 1 and len(mw) >= 1:
        if fw[0] == mw[0]:
            if len(fw) >= 2 and len(mw) >= 2:
                return fw[1] == mw[1]
            return True  # single word match

    return False


def names_match_strict(focco_name, mother_name):
    """
    Stricter matching that also considers product type abbreviation.
    Returns (match_score, matches) where score=2 means type also matches.
    """
    if not names_match(focco_name, mother_name):
        return 0

    ftype = focco_type(focco_name)
    mother_upper = mother_name.upper().strip()

    # Check if type hint in mother name matches focco type
    for type_abbr, focco_types in MOTHER_TYPE_HINTS.items():
        if type_abbr in mother_upper.split() and ftype in focco_types:
            return 2  # Strong match - type matches too
        if type_abbr in mother_upper and ftype in focco_types:
            return 2

    # Check special type matches
    if ftype == 'PUFF' and 'PUFF' in mother_upper:
        return 2
    if ftype == 'POL' and ('POL ' in mother_upper or mother_upper.endswith('POL')):
        return 2
    if ftype == 'CH' and 'CH ' in mother_upper:
        return 2
    if ftype == 'CAMA' and 'CAMA' in mother_upper:
        return 2
    if ftype == 'CABECEIRA' and 'CABECEIRA' in mother_upper:
        return 2
    if ftype == 'BASE' and 'BASE' in mother_upper:
        return 2
    if ftype in ('SOFA', 'SOFA CAMA') and not any(t in mother_upper for t in ['POL', 'PUFF ', 'CH ', 'CAMA', 'BANCO']):
        return 2  # sofa matches if no other type indicator

    return 1  # Weak match - name matches but type unclear


# ============================================================
# DIMENSION UTILITIES
# ============================================================
def dim_from_str(s):
    """
    Parse dimension string handling both formats:
    - Meter format: '2,47X1,10X0,85' -> (2.47, 1.10, 0.85)
    - Cm format: '247X110X85' -> (2.47, 1.10, 0.85) [auto-detected by values >10]
    - Mixed: '2,47X1,10X0,85 (ESPECIAL)' - ignore parenthesized content
    """
    if not s:
        return None
    # Remove parenthesized content and trailing H (height indicator)
    s = re.sub(r'\s*\([^)]*\)', '', str(s)).strip()
    s = re.sub(r'[Hh]$', '', s)
    # Only take the first dimension sequence (before space or slash)
    s = re.split(r'[\s/]', s)[0]
    parts = re.split(r'[Xx]', s.replace(',', '.'))
    try:
        floats = []
        for p in parts:
            p = p.strip()
            if not p or p in ('-',):
                continue
            try:
                floats.append(float(p))
            except:
                pass
        if not floats:
            return None
        # Auto-detect cm vs meters: if all major dims > 10, assume cm
        if len(floats) >= 2 and all(v > 10 for v in floats[:2]):
            floats = [round(v/100, 2) for v in floats]
        else:
            floats = [round(v, 2) for v in floats]
        return tuple(floats)
    except:
        pass
    return None


def dims_match(focco_dim, mother_dim, tolerance=0.02):
    """
    focco_dim: (C, P, A) tuple
    mother_dim: (C, P) or (C, P, A) tuple
    Match on first 2 values (C and P), ignore A.
    """
    if not focco_dim or not mother_dim:
        return False
    if len(focco_dim) < 2 or len(mother_dim) < 2:
        return False
    return (abs(focco_dim[0] - mother_dim[0]) <= tolerance and
            abs(focco_dim[1] - mother_dim[1]) <= tolerance)


# ============================================================
# LOAD PIU MOBILE TABLE
# Sheets have product names in col[1], C=col[2], P=col[3], A=col[4]
# Grade cols: FORN=col[8], FX A=col[9] ... FX KNIT=col[21]
# ============================================================
def load_piu_mobile():
    wb = openpyxl.load_workbook(TABLE_PIU_MOBILE, data_only=True)
    products = {}  # name_upper -> { dim_tuple -> { grade_normalized -> price } }

    PIU_GRADE_COLS = {
        8: 'FX FORNECIDO', 9: 'FX A', 10: 'FX B', 11: 'FX C', 12: 'FX D',
        13: 'FX E', 14: 'FX F', 15: 'FX G', 16: 'FX H', 17: 'FX I',
        18: 'FX J', 19: 'FX R', 20: 'FX V', 21: 'FX KNIT'
    }

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        for row in ws.iter_rows(values_only=True):
            name = row[1] if len(row) > 1 else None
            if not name or not isinstance(name, str):
                continue
            name_up = name.strip().upper()
            if not name_up or name_up == 'PRODUTO':
                continue
            c_val = row[2] if len(row) > 2 else None
            p_val = row[3] if len(row) > 3 else None
            if not isinstance(c_val, (int, float)):
                continue
            dim = (round(float(c_val), 2),
                   round(float(p_val), 2) if isinstance(p_val, (int, float)) else 0.0)
            a_val = row[4] if len(row) > 4 else None
            if isinstance(a_val, (int, float)):
                dim = dim + (round(float(a_val), 2),)

            if name_up not in products:
                products[name_up] = {}
            if dim not in products[name_up]:
                products[name_up][dim] = {}

            for ci, grade in PIU_GRADE_COLS.items():
                if len(row) > ci and isinstance(row[ci], (int, float)) and row[ci] > 0:
                    products[name_up][dim][grade] = float(row[ci])

    return products


# ============================================================
# LOAD PIURB TABLE
# Products: col[0]=name, col[1]=C, col[2]=P, col[3]=A
# Grade cols vary by section; typical: TC FORN RB=col[7], FX G=col[8]...FX V=col[13], CR VQ=col[14], CR FZ=col[15], COUR FORN=col[16]
# ============================================================
def load_piurb():
    wb = openpyxl.load_workbook(TABLE_PIURB, data_only=True)
    products = {}

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        # Track current header mapping
        cur_grade_cols = {}  # col_index -> grade_normalized

        for row in ws.iter_rows(values_only=True):
            name = row[0] if len(row) > 0 else None
            if not name:
                continue

            if isinstance(name, str):
                name_up = name.strip().upper()
                if not name_up:
                    continue

                # Detect header row
                if name_up == 'PRODUTO':
                    cur_grade_cols = {}
                    for ci, val in enumerate(row):
                        if val and isinstance(val, str):
                            v = val.strip().rstrip()
                            ng = GRADE_NORMALIZE.get(v)
                            if ng is None:
                                # Try alternate spellings
                                ng = GRADE_NORMALIZE.get(v.strip())
                            if ng:
                                # Check for duplicates, prefer first occurrence
                                if ng not in cur_grade_cols.values():
                                    cur_grade_cols[ci] = ng
                    continue

                # Product row
                c_val = row[1] if len(row) > 1 else None
                p_val = row[2] if len(row) > 2 else None
                if not isinstance(c_val, (int, float)):
                    continue

                dim = (round(float(c_val), 2),
                       round(float(p_val), 2) if isinstance(p_val, (int, float)) else 0.0)
                a_val = row[3] if len(row) > 3 else None
                if isinstance(a_val, (int, float)):
                    dim = dim + (round(float(a_val), 2),)

                if name_up not in products:
                    products[name_up] = {}
                if dim not in products[name_up]:
                    products[name_up][dim] = {}

                for ci, grade in cur_grade_cols.items():
                    if len(row) > ci and isinstance(row[ci], (int, float)) and row[ci] > 0:
                        products[name_up][dim][grade] = float(row[ci])

    return products


# ============================================================
# LOAD ITALO TABLE
# Products: col[0]=name, col[1]=C, col[2]=P, col[3]=A
# Grade cols: TC FORN RB=col[7], FX A=col[8]..FX F=col[13], FX G=col[14]..FX V=col[19],
#             CR VQ=col[20], CR FZ=col[21], COUR FORN=col[22]
# ============================================================
def load_italo():
    wb = openpyxl.load_workbook(TABLE_ITALO, data_only=True)
    products = {}

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        cur_grade_cols = {}

        for row in ws.iter_rows(values_only=True):
            name = row[0] if len(row) > 0 else None
            if not name:
                continue

            if isinstance(name, str):
                name_up = name.strip().upper()
                if not name_up:
                    continue

                if name_up == 'PRODUTO':
                    cur_grade_cols = {}
                    for ci, val in enumerate(row):
                        if val and isinstance(val, str):
                            v = val.strip().rstrip()
                            ng = GRADE_NORMALIZE.get(v)
                            if ng is None:
                                ng = GRADE_NORMALIZE.get(v.strip())
                            if ng and ng not in cur_grade_cols.values():
                                cur_grade_cols[ci] = ng
                    continue

                c_val = row[1] if len(row) > 1 else None
                p_val = row[2] if len(row) > 2 else None
                if not isinstance(c_val, (int, float)):
                    continue

                dim = (round(float(c_val), 2),
                       round(float(p_val), 2) if isinstance(p_val, (int, float)) else 0.0)
                a_val = row[3] if len(row) > 3 else None
                if isinstance(a_val, (int, float)):
                    dim = dim + (round(float(a_val), 2),)

                if name_up not in products:
                    products[name_up] = {}
                if dim not in products[name_up]:
                    products[name_up][dim] = {}

                for ci, grade in cur_grade_cols.items():
                    if len(row) > ci and isinstance(row[ci], (int, float)) and row[ci] > 0:
                        products[name_up][dim][grade] = float(row[ci])

    return products


# ============================================================
# PARSE FOCCO
# ============================================================
def parse_focco():
    wb = openpyxl.load_workbook(FOCCO_PATH, data_only=True)

    ws_items = wb['Itens']
    items = {}
    for i, row in enumerate(ws_items.iter_rows(values_only=True)):
        if i == 0: continue
        code = str(row[0]).strip() if row[0] else ''
        name = str(row[1]).strip() if row[1] else ''
        if code:
            items[code] = name

    ws_rules = wb['Itens - Regras de Preço']
    rules = defaultdict(list)

    for i, row in enumerate(ws_rules.iter_rows(values_only=True)):
        if i == 0: continue
        code = str(row[0]).strip() if row[0] else ''
        cond = str(row[2]).strip() if row[2] else ''
        price = row[5]

        if not code or not cond or price is None:
            continue
        if not isinstance(price, (int, float)):
            continue

        dim_match = re.search(r'DIMENSAO_\w+="([^"]+)"', cond)
        dim_str = dim_match.group(1) if dim_match else None

        grade = None
        grade_match = re.search(r'(?:CAT_TECIDO\d*|TECIDO_\w+)="([^"]+)"', cond)
        if grade_match:
            grade = grade_match.group(1).strip()

        if grade is None:
            continue

        rules[code].append((dim_str, grade, float(price)))

    return items, rules


# ============================================================
# MAIN AUDIT
# ============================================================
def main():
    print("Loading Focco data...")
    items, rules = parse_focco()

    print("Loading mother tables...")
    table_piu = load_piu_mobile()
    table_piurb = load_piurb()
    table_italo = load_italo()

    tables = {
        'PIU MOBILE': table_piu,
        'PIURB': table_piurb,
        'ITALO': table_italo,
    }

    print(f"PIU MOBILE products: {len(table_piu)}")
    print(f"PIURB products: {len(table_piurb)}")
    print(f"ITALO products: {len(table_italo)}")

    # Results
    discrepancies = []
    no_match = []
    no_dim_match = []  # matched by name but couldn't find dim

    numeric_codes = sorted([c for c in items if c.isdigit() and c not in SKIP_CODES])

    for code in numeric_codes:
        name = items[code]
        item_rules = rules.get(code, [])

        if not item_rules:
            continue

        # Find all matching mother products by name, sorted by match quality
        name_matches_raw = []  # list of (score, table_name, product_name_in_table)
        for tname, tdata in tables.items():
            for pname in tdata:
                score = names_match_strict(name, pname)
                if score > 0:
                    name_matches_raw.append((score, tname, pname))
        # Sort by score descending - prefer type-matched products
        name_matches_raw.sort(key=lambda x: -x[0])
        name_matches = [(t, p) for _, t, p in name_matches_raw]

        if not name_matches:
            no_match.append((code, name, len(item_rules)))
            continue

        # For each focco rule, compare price
        for (dim_str, grade, focco_price) in item_rules:
            focco_dim = dim_from_str(dim_str)
            norm_grade = GRADE_NORMALIZE.get(grade, grade)

            mother_price = None
            mother_source = None
            mother_prod = None
            found_dim = False

            for (tname, pname) in name_matches:
                tdata = tables[tname]
                prod_dims = tdata.get(pname, {})

                for mdim, grade_prices in prod_dims.items():
                    if focco_dim is None or dims_match(focco_dim, mdim):
                        found_dim = True
                        # Try to find the grade
                        price = grade_prices.get(norm_grade)
                        if price is None and grade != norm_grade:
                            price = grade_prices.get(grade)
                        if price is not None:
                            mother_price = price
                            mother_source = tname
                            mother_prod = pname
                            break
                if mother_price is not None:
                    break

            if mother_price is None:
                if found_dim:
                    no_dim_match.append({
                        'code': code, 'name': name, 'dim': dim_str,
                        'grade': grade, 'norm_grade': norm_grade,
                        'focco_price': focco_price,
                        'matches': name_matches[:2]
                    })
                else:
                    no_dim_match.append({
                        'code': code, 'name': name, 'dim': dim_str,
                        'grade': grade, 'norm_grade': norm_grade,
                        'focco_price': focco_price,
                        'matches': name_matches[:2],
                        'reason': 'no_dim'
                    })
            else:
                diff = focco_price - mother_price
                if abs(diff) > 1:
                    discrepancies.append({
                        'code': code, 'name': name,
                        'dim': dim_str, 'grade': grade,
                        'focco_price': focco_price,
                        'mother_price': mother_price,
                        'diff': diff,
                        'source': mother_source,
                        'mother_prod': mother_prod
                    })

    # ============================================================
    # REPORT
    # ============================================================
    print("\n" + "="*110)
    print("AUDIT REPORT - FOCCO vs MOTHER TABLES")
    print("="*110)

    # SECTION 1: No match
    print(f"\n{'='*70}")
    print(f"SECTION 1: ITEMS WITH NO MATCH IN ANY MOTHER TABLE ({len(no_match)} items)")
    print(f"{'='*70}")
    for code, name, rule_count in no_match:
        print(f"  {code}: {name}  ({rule_count} price rules)")

    # SECTION 2: Price discrepancies
    print(f"\n{'='*70}")
    print(f"SECTION 2: PRICE DISCREPANCIES ({len(discrepancies)} rules with |diff| > R$1)")
    print(f"{'='*70}")
    if discrepancies:
        print(f"  {'CODE':<8} {'NAME':<35} {'DIM':<28} {'GRADE':<18} {'FOCCO':>8} {'MOTHER':>8} {'DIFF':>8}  SOURCE")
        print("  " + "-"*120)
        for d in sorted(discrepancies, key=lambda x: (x['code'], x['dim'] or '', x['grade'])):
            diff_str = f"{int(d['diff']):+d}"
            print(f"  {d['code']:<8} {d['name'][:35]:<35} {str(d['dim'] or '')[:28]:<28} {d['grade'][:18]:<18} "
                  f"{int(d['focco_price']):>8} {int(d['mother_price']):>8} {diff_str:>8}  {d['source']}")
    else:
        print("  No price discrepancies found!")

    # SECTION 3: Unmatched grades/dims - summarize by item
    print(f"\n{'='*70}")
    print(f"SECTION 3: UNMATCHED GRADES OR DIMENSIONS ({len(no_dim_match)} rules)")
    print(f"(Name matched in mother table but grade/dim not found)")
    print(f"{'='*70}")
    by_code = defaultdict(list)
    for m in no_dim_match:
        by_code[m['code']].append(m)
    for code in sorted(by_code.keys()):
        entries = by_code[code]
        name = entries[0]['name']
        grades_missing = sorted(set(e['grade'] for e in entries))
        dims_missing = sorted(set(e['dim'] for e in entries if e['dim']))
        reasons = set(e.get('reason', 'grade_not_found') for e in entries)
        matched_to = entries[0]['matches']
        print(f"\n  {code}: {name}")
        print(f"    Matched to: {[p for _, p in matched_to]}")
        print(f"    Missing grades: {grades_missing}")
        if 'no_dim' in reasons:
            print(f"    Some dims not found: {dims_missing[:5]}")
        print(f"    Count: {len(entries)} unmatched rules")

    # SUMMARY
    print(f"\n{'='*70}")
    print("SUMMARY")
    print(f"{'='*70}")
    print(f"  Total Focco items with price rules:  {len([c for c in numeric_codes if rules.get(c)])}")
    print(f"  Skipped (already corrected):         {len(SKIP_CODES)}")
    print(f"  Items with no name match:            {len(no_match)}")
    print(f"  Price discrepancies (|diff|>R$1):    {len(discrepancies)}")
    print(f"  Unmatched grade/dim rules:           {len(no_dim_match)}")
    if discrepancies:
        total_diff = sum(abs(d['diff']) for d in discrepancies)
        max_disc = max(discrepancies, key=lambda d: abs(d['diff']))
        print(f"  Total absolute difference:           R$ {total_diff:,.0f}")
        print(f"  Largest discrepancy:                 {max_disc['code']} {max_disc['name']} "
              f"{max_disc['dim']} {max_disc['grade']} -> R$ {max_disc['diff']:+,.0f}")

    return discrepancies, no_match, no_dim_match


if __name__ == '__main__':
    main()
