#!/usr/bin/env python3
"""
Focco ERP Price Audit vs Mother Tables
Compares prices in Focco 'Itens - Regras de Preço' against 3 mother price tables.
"""

import openpyxl
import re
from collections import defaultdict

# ============================================================
# CONFIG
# ============================================================
FOCCO_PATH = '/home/user/claude/PIU_MOBILE_HDA_CORRIGIDO3.xlsx'
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

# Non-numeric item codes to skip
SKIP_NON_NUMERIC = True

# ============================================================
# FABRIC GRADE MAPPING
# Focco CAT_TECIDO/TECIDO_xxx values → mother table column index
# Mother table header row (0-indexed from col 0):
# col1=Produto, col2=C, col3=P, col4=A, col5=Braço, col6=Tec.Liso, col7=Couro FORN
# col8=FORN, col9=FX A, col10=FX B, col11=FX C, col12=FX D, col13=FX E,
# col14=FX F, col15=FX G, col16=FX H, col17=FX I, col18=FX J, col19=FX R,
# col20=FX V, col21=FX KNIT
# ============================================================

GRADE_TO_COL = {
    # Standard FX grades (PIU MOBILE and ITALO tables)
    'FX FORNECIDO': 8,
    'FORNECIDO': 8,
    'FORN': 8,
    'FX FORN RB': 8,
    'FX FORN. RB': 8,
    'COURO FORN.': 8,
    'COUR FORN. ': 8,
    'TC FORN. RB': 8,
    'TECIDO FORNECIDO': 8,
    'TECIDO_FORNECIDO': 8,
    'TEC_FORNECIDO': 8,
    'FX A': 9, 'TECIDO A': 9, 'TEC A': 9, 'TECIDO_A': 9,
    'FX B': 10, 'TECIDO B': 10, 'TEC B': 10, 'TECIDO_B': 10,
    'FX C': 11, 'TECIDO C': 11, 'TEC C': 11, 'TECIDO_C': 11,
    'FX D': 12, 'TECIDO D': 12, 'TEC D': 12, 'TECIDO_D': 12, 'TECIDO D ': 12,
    'FX E': 13, 'TECIDO E': 13, 'TEC E': 13, 'TECIDO_E': 13,
    'FX F': 14, 'TECIDO F': 14, 'TEC F': 14, 'TECIDO_F': 14,
    'FX G': 15, 'TECIDO G': 15, 'TEC G': 15, 'TECIDO_G': 15,
    'FX H': 16, 'TECIDO H': 16, 'TEC H': 16, 'TECIDO_H': 16,
    'FX I': 17, 'TECIDO I': 17, 'TEC I': 17, 'TECIDO_I': 17,
    'FX J': 18, 'TECIDO J': 18, 'TEC J': 18, 'TECIDO_J': 18,
    'FX J / N': 18, 'FX J/N': 18, 'TECIDO J/N': 18, 'TECIDO J ': 18, 'J/N': 18,
    'TECIDO_N': 18,  # legacy "N" maps to J/N column
    'TECIDO_J': 18,
    'FX R': 19, 'TECIDO R': 19, 'TEC_FORNECIDO_V_8001': 19,
    'TECIDO_R': 19, 'R': 19,
    'FX V': 20, 'TECIDO V': 20, 'TECIDO_V': 20, 'V': 20,
    'FX KNIT': 21, 'TECIDO KNIT': 21,
    # PIURB extra grades
    'CR VQ': 20,   # Couro VQ similar to FX V column position in PIURB
    'CR FZ': 21,
    # Some old format
    'TECIDO A': 9, 'G': 15, 'F': 14,
}

# PIURB table has different column positions
# Produto(0), C(1), P(2), A(3), Braço(4), TEC mt L(5), Qtd Couro FORN(6), TC FORN RB(7),
# FX G(8), FX H(9), FX I(10), FX J(11), FX R(12), FX V(13), CR VQ(14), CR FZ(15), COUR FORN(16)
PIURB_GRADE_TO_COL = {
    'TC FORN. RB': 7, 'FX FORN. RB': 7, 'FORNECIDO': 7, 'FX FORNECIDO': 7,
    'TECIDO FORNECIDO': 7, 'COURO FORN.': 16, 'COUR FORN. ': 16,
    'FX G': 8, 'TECIDO G': 8, 'TECIDO_G': 8,
    'FX H': 9, 'TECIDO H': 9, 'TECIDO_H': 9,
    'FX I': 10, 'TECIDO I': 10, 'TECIDO_I': 10,
    'FX J': 11, 'FX J / N': 11, 'FX J/N': 11, 'TECIDO_J': 11, 'TECIDO_N': 11,
    'FX R': 12, 'TECIDO R': 12, 'TECIDO_R': 12, 'R': 12,
    'FX V': 13, 'TECIDO V': 13, 'TECIDO_V': 13, 'V': 13,
    'CR VQ': 14,
    'CR FZ': 15,
}

# ITALO table column positions (from exploration):
# col0=Produto, col1=C, col2=P, col3=A, col4=Braço, col5=TEC mt L, col6=Qtd Couro FORN
# col7=TC FORN RB, col8=FX A, col9=FX B, col10=FX C, col11=FX D, col12=FX E, col13=FX F,
# col14=FX G, col15=FX H, col16=FX I, col17=FX J/N, col18=FX R, col19=FX V,
# col20=CR VQ, col21=CR FZ, col22=COUR FORN
ITALO_GRADE_TO_COL = {
    'TC FORN. RB': 7, 'FX FORN. RB': 7, 'FORNECIDO': 7, 'FX FORNECIDO': 7,
    'TECIDO FORNECIDO': 7, 'COURO FORN.': 22, 'COUR FORN. ': 22,
    'FX A': 8, 'TECIDO A': 8, 'TECIDO_A': 8,
    'FX B': 9, 'TECIDO B': 9, 'TECIDO_B': 9,
    'FX C': 10, 'TECIDO C': 10, 'TECIDO_C': 10,
    'FX D': 11, 'TECIDO D': 11, 'TECIDO_D': 11,
    'FX E': 12, 'TECIDO E': 12, 'TECIDO_E': 12,
    'FX F': 13, 'TECIDO F': 13, 'TECIDO_F': 13,
    'FX G': 14, 'TECIDO G': 14, 'TECIDO_G': 14,
    'FX H': 15, 'TECIDO H': 15, 'TECIDO_H': 15,
    'FX I': 16, 'TECIDO I': 16, 'TECIDO_I': 16,
    'FX J': 17, 'FX J / N': 17, 'FX J/N': 17, 'TECIDO_J': 17, 'TECIDO_N': 17,
    'FX R': 18, 'TECIDO R': 18, 'TECIDO_R': 18, 'R': 18,
    'FX V': 19, 'TECIDO V': 19, 'TECIDO_V': 19, 'V': 19,
    'CR VQ': 20,
    'CR FZ': 21,
}

# ============================================================
# HELPER: normalize dimension string for comparison
# "200X102X085" -> "2,00X1,02X0,85"
# "2,00X1,02X0,85" -> "2,00X1,02X0,85"
# ============================================================
def normalize_dim(s):
    """Convert dimension string to canonical float-tuple."""
    if not s:
        return None
    s = str(s).strip()
    # Try format with commas: "2,47X1,10X0,85"
    parts_comma = re.split(r'[Xx]', s.replace(',', '.'))
    try:
        floats = [round(float(p), 2) for p in parts_comma if p.strip()]
        if floats:
            return tuple(floats)
    except:
        pass
    return None

def dim_key_from_floats(dims):
    """Create a comparable key from dimension floats."""
    if not dims:
        return None
    return tuple(round(d, 2) for d in dims)


# ============================================================
# LOAD MOTHER TABLES
# Returns dict: {normalized_name -> {dim_tuple -> {grade -> price}}}
# dim_tuple = (C, P, A) floats
# ============================================================
def load_mother_table_piu_mobile():
    """Load PIU MOBILE mother table - all 4 sheets."""
    wb = openpyxl.load_workbook(TABLE_PIU_MOBILE, data_only=True)
    products = {}  # name -> {dim_tuple -> {grade -> price}}

    # PIU MOBILE: sheets have header rows with grade labels
    # Header row structure: col1=Produto, col2=C, col3=P, col4=A, col5=Braço,
    # col6=Tec.Liso, col7=Couro FORN, col8=FORN, col9=FX A ... col21=FX KNIT

    GRADE_COLS_PIU = {
        8: 'FX FORNECIDO', 9: 'FX A', 10: 'FX B', 11: 'FX C', 12: 'FX D',
        13: 'FX E', 14: 'FX F', 15: 'FX G', 16: 'FX H', 17: 'FX I',
        18: 'FX J', 19: 'FX R', 20: 'FX V', 21: 'FX KNIT'
    }

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))
        for row in rows:
            # Product rows: col1 has name, col2=C (float), col3=P (float), col4=A
            name = row[1] if len(row) > 1 else None
            if not name or not isinstance(name, str):
                continue
            name = name.strip().upper()
            if not name or name in ('PRODUTO', 'C', 'P', 'A'):
                continue
            # Check if col2, col3 are numeric (dimensions)
            c_val = row[2] if len(row) > 2 else None
            p_val = row[3] if len(row) > 3 else None
            a_val = row[4] if len(row) > 4 else None
            if not isinstance(c_val, (int, float)):
                continue
            dim = (round(float(c_val), 2), round(float(p_val), 2) if isinstance(p_val, (int, float)) else 0)
            if a_val and isinstance(a_val, (int, float)):
                dim = dim + (round(float(a_val), 2),)

            if name not in products:
                products[name] = {}
            if dim not in products[name]:
                products[name][dim] = {}

            for col_idx, grade in GRADE_COLS_PIU.items():
                if len(row) > col_idx:
                    price = row[col_idx]
                    if isinstance(price, (int, float)) and price > 0:
                        products[name][dim][grade] = price

    return products


def load_mother_table_piurb():
    """Load PIURB mother table."""
    wb = openpyxl.load_workbook(TABLE_PIURB, data_only=True)
    products = {}

    # PIURB: col0=Produto, col1=C, col2=P, col3=A, col4=Braço
    # Then grade columns vary by section header
    # From exploration: col7=TC FORN RB, col8=FX G, col9=FX H, col10=FX I,
    # col11=FX J, col12=FX R, col13=FX V, col14=CR VQ, col15=CR FZ, col16=COUR FORN
    GRADE_COLS_PIURB = {
        7: 'FX FORNECIDO', 8: 'FX G', 9: 'FX H', 10: 'FX I',
        11: 'FX J', 12: 'FX R', 13: 'FX V', 14: 'CR VQ', 15: 'CR FZ', 16: 'COURO FORN'
    }

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))
        current_header_cols = dict(GRADE_COLS_PIURB)  # default

        for row in rows:
            name = row[0] if len(row) > 0 else None
            if not name or not isinstance(name, str):
                continue
            name_upper = name.strip().upper()

            # Check if this is a header row
            if name_upper == 'PRODUTO':
                # Re-parse header to get actual grade columns
                current_header_cols = {}
                for ci, val in enumerate(row):
                    if val and isinstance(val, str):
                        v = val.strip()
                        if v in ('FX G', 'FX H', 'FX I', 'FX J', 'FX J / N', 'FX R', 'FX V',
                                 'CR VQ', 'CR FZ', 'TC FORN. RB', 'FX FORN. RB', 'COUR FORN. '):
                            current_header_cols[ci] = v
                continue

            if not name_upper:
                continue

            c_val = row[1] if len(row) > 1 else None
            p_val = row[2] if len(row) > 2 else None
            a_val = row[3] if len(row) > 3 else None
            if not isinstance(c_val, (int, float)):
                continue

            dim = (round(float(c_val), 2), round(float(p_val), 2) if isinstance(p_val, (int, float)) else 0)
            if a_val and isinstance(a_val, (int, float)):
                dim = dim + (round(float(a_val), 2),)

            if name_upper not in products:
                products[name_upper] = {}
            if dim not in products[name_upper]:
                products[name_upper][dim] = {}

            for col_idx, grade in current_header_cols.items():
                if len(row) > col_idx:
                    price = row[col_idx]
                    if isinstance(price, (int, float)) and price > 0:
                        products[name_upper][dim][grade] = price

    return products


def load_mother_table_italo():
    """Load ITALO mother table."""
    wb = openpyxl.load_workbook(TABLE_ITALO, data_only=True)
    products = {}

    # From exploration: col0=Produto, col1=C, col2=P, col3=A, col4=Braço
    # col5=TEC mt L, col6=Qtd Couro FORN, col7=TC FORN RB
    # col8=FX A...FX F(col13), FX G(col14)...FX V(col19), CR VQ(col20), CR FZ(col21), COUR FORN(col22)
    # Duplicate set starting at col24: TC FORN RB(col24), FX A(col25)...

    GRADE_COLS_ITALO = {
        7: 'FX FORNECIDO',
        8: 'FX A', 9: 'FX B', 10: 'FX C', 11: 'FX D', 12: 'FX E', 13: 'FX F',
        14: 'FX G', 15: 'FX H', 16: 'FX I', 17: 'FX J', 18: 'FX R', 19: 'FX V',
        20: 'CR VQ', 21: 'CR FZ', 22: 'COURO FORN'
    }

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))
        current_header_cols = dict(GRADE_COLS_ITALO)

        for row in rows:
            name = row[0] if len(row) > 0 else None
            if not name or not isinstance(name, str):
                continue
            name_upper = name.strip().upper()

            if name_upper == 'PRODUTO':
                # Re-parse columns
                current_header_cols = {}
                for ci, val in enumerate(row):
                    if val and isinstance(val, str):
                        v = val.strip().rstrip()
                        if v in ('FX A', 'FX B', 'FX C', 'FX D', 'FX E', 'FX F ',
                                 'FX F', 'FX G', 'FX H', 'FX I', 'FX J / N',
                                 'FX R', 'FX V', 'CR VQ', 'CR FZ',
                                 'TC FORN. RB', 'COUR FORN. '):
                            current_header_cols[ci] = v.strip()
                continue

            if not name_upper:
                continue

            c_val = row[1] if len(row) > 1 else None
            p_val = row[2] if len(row) > 2 else None
            a_val = row[3] if len(row) > 3 else None
            if not isinstance(c_val, (int, float)):
                continue

            dim = (round(float(c_val), 2), round(float(p_val), 2) if isinstance(p_val, (int, float)) else 0)
            if a_val and isinstance(a_val, (int, float)):
                dim = dim + (round(float(a_val), 2),)

            if name_upper not in products:
                products[name_upper] = {}
            if dim not in products[name_upper]:
                products[name_upper][dim] = {}

            for col_idx, grade in current_header_cols.items():
                if len(row) > col_idx:
                    price = row[col_idx]
                    if isinstance(price, (int, float)) and price > 0:
                        products[name_upper][dim][grade] = price

    return products


# ============================================================
# PARSE FOCCO RULES
# Returns: {code -> [(dim_str, grade, price)]}
# ============================================================
def parse_focco_rules():
    wb = openpyxl.load_workbook(FOCCO_PATH, data_only=True)

    # Get items
    ws_items = wb['Itens']
    items = {}  # code -> name
    for i, row in enumerate(ws_items.iter_rows(values_only=True)):
        if i == 0: continue
        code = str(row[0]).strip() if row[0] else ''
        name = str(row[1]).strip() if row[1] else ''
        if code:
            items[code] = name

    # Parse rules
    ws_rules = wb['Itens - Regras de Preço']
    rules = defaultdict(list)  # code -> [(dim_str, grade, price)]

    for i, row in enumerate(ws_rules.iter_rows(values_only=True)):
        if i == 0: continue
        code = str(row[0]).strip() if row[0] else ''
        cond = str(row[2]).strip() if row[2] else ''
        price = row[5]

        if not code or not cond or price is None:
            continue
        if not isinstance(price, (int, float)):
            continue

        # Extract dimension from condition
        # Format: DIMENSAO_xxx="2,47X1,10X0,85"
        dim_match = re.search(r'DIMENSAO_\w+="([^"]+)"', cond)
        dim_str = dim_match.group(1) if dim_match else None

        # Extract fabric grade
        # Can be: CAT_TECIDO="FX A" or TECIDO_22053="FX A" or CAT_TECIDO="FORNECIDO"
        grade = None
        grade_match = re.search(r'(?:CAT_TECIDO\d*|TECIDO_\w+)="([^"]+)"', cond)
        if grade_match:
            grade = grade_match.group(1).strip()

        if grade is None:
            continue

        rules[code].append((dim_str, grade, float(price)))

    return items, rules


# ============================================================
# NAME MATCHING
# Try to find product in mother tables by name
# ============================================================
def extract_product_keywords(name):
    """Extract key words from product name for matching."""
    name = name.upper().strip()
    # Remove common prefixes
    name = re.sub(r'^(SOFA|POLTRONA|PUFF|CAMA|CHAISE|BANCO|CANTO|MESA|CADEIRA|CABECEIRA|BASE)\s+', '', name)
    return name.split()


def find_in_mother_tables(focco_name, tables):
    """
    Try to match focco item name to a product in one of the mother tables.
    Returns list of (table_name, matched_name) tuples.
    """
    focco_upper = focco_name.upper().strip()
    matches = []
    for table_name, table_data in tables.items():
        for product_name in table_data.keys():
            # Exact substring match
            if focco_upper in product_name or product_name in focco_upper:
                matches.append((table_name, product_name))
    return matches


# ============================================================
# DIMENSION MATCHING
# Compare focco dimension string to mother table dimension key
# ============================================================
def dim_from_focco(dim_str):
    """Parse focco dimension like '2,47X1,10X0,85' -> (2.47, 1.10, 0.85)"""
    if not dim_str:
        return None
    parts = re.split(r'[Xx]', dim_str.replace(',', '.'))
    try:
        floats = [round(float(p), 2) for p in parts if p.strip()]
        return tuple(floats)
    except:
        return None


def dim_matches(focco_dim_tuple, mother_dim_key, tolerance=0.02):
    """Check if focco dim (C, P, A) matches mother table key (C, P) or (C, P, A)."""
    if not focco_dim_tuple or not mother_dim_key:
        return False
    # Mother key is (C, P) or (C, P, A)
    # Focco dim is usually (C, P, A)
    if len(focco_dim_tuple) >= 2 and len(mother_dim_key) >= 2:
        c_match = abs(focco_dim_tuple[0] - mother_dim_key[0]) <= tolerance
        p_match = abs(focco_dim_tuple[1] - mother_dim_key[1]) <= tolerance
        if c_match and p_match:
            return True
    return False


# ============================================================
# NORMALIZE GRADE for comparison
# ============================================================
GRADE_NORMALIZE = {
    'FX FORNECIDO': 'FX FORNECIDO',
    'FORNECIDO': 'FX FORNECIDO',
    'FORN': 'FX FORNECIDO',
    'FX FORN RB': 'FX FORNECIDO',
    'FX FORN. RB': 'FX FORNECIDO',
    'COURO FORN.': 'FX FORNECIDO',
    'COUR FORN. ': 'FX FORNECIDO',
    'TC FORN. RB': 'FX FORNECIDO',
    'TECIDO FORNECIDO': 'FX FORNECIDO',
    'TECIDO_FORNECIDO': 'FX FORNECIDO',
    'TEC_FORNECIDO': 'FX FORNECIDO',
    'FX A': 'FX A', 'TECIDO A': 'FX A', 'TEC A': 'FX A', 'TECIDO_A': 'FX A',
    'FX B': 'FX B', 'TECIDO B': 'FX B', 'TEC B': 'FX B', 'TECIDO_B': 'FX B',
    'FX C': 'FX C', 'TECIDO C': 'FX C', 'TEC C': 'FX C', 'TECIDO_C': 'FX C',
    'FX D': 'FX D', 'TECIDO D': 'FX D', 'TEC D': 'FX D', 'TECIDO_D': 'FX D', 'TECIDO D ': 'FX D',
    'FX E': 'FX E', 'TECIDO E': 'FX E', 'TEC E': 'FX E', 'TECIDO_E': 'FX E',
    'FX F': 'FX F', 'TECIDO F': 'FX F', 'TEC F': 'FX F', 'TECIDO_F': 'FX F',
    'FX G': 'FX G', 'TECIDO G': 'FX G', 'TEC G': 'FX G', 'TECIDO_G': 'FX G', 'G': 'FX G',
    'FX H': 'FX H', 'TECIDO H': 'FX H', 'TEC H': 'FX H', 'TECIDO_H': 'FX H',
    'FX I': 'FX I', 'TECIDO I': 'FX I', 'TEC I': 'FX I', 'TECIDO_I': 'FX I',
    'FX J': 'FX J', 'TECIDO J': 'FX J', 'TEC J': 'FX J', 'TECIDO_J': 'FX J',
    'FX J / N': 'FX J', 'FX J/N': 'FX J', 'TECIDO J/N': 'FX J', 'TECIDO J ': 'FX J',
    'J/N': 'FX J', 'TECIDO_N': 'FX J',
    'FX R': 'FX R', 'TECIDO R': 'FX R', 'TECIDO_R': 'FX R', 'R': 'FX R',
    'FX V': 'FX V', 'TECIDO V': 'FX V', 'TECIDO_V': 'FX V', 'V': 'FX V',
    'FX KNIT': 'FX KNIT', 'TECIDO KNIT': 'FX KNIT',
    'CR VQ': 'FX V',
    'CR FZ': 'FX KNIT',
    'F': 'FX F',
    'COURO FORN': 'FX FORNECIDO',
    'COURO FORN ': 'FX FORNECIDO',
}


# ============================================================
# MAIN AUDIT
# ============================================================
def main():
    print("Loading Focco data...")
    items, rules = parse_focco_rules()

    print("Loading mother tables...")
    table_piu = load_mother_table_piu_mobile()
    table_piurb = load_mother_table_piurb()
    table_italo = load_mother_table_italo()

    tables = {
        'PIU MOBILE': table_piu,
        'PIURB': table_piurb,
        'ITALO': table_italo,
    }

    # Merge all tables for name search
    all_table_products = {}
    for tname, tdata in tables.items():
        for pname in tdata:
            all_table_products.setdefault(pname, []).append(tname)

    print(f"\nFocco items: {len(items)}")
    print(f"PIU MOBILE products: {len(table_piu)}")
    print(f"PIURB products: {len(table_piurb)}")
    print(f"ITALO products: {len(table_italo)}")

    # ---- Reports ----
    discrepancies = []
    missing_grades = []
    no_match = []
    matched_no_dim = []

    numeric_codes = [c for c in items if c.isdigit() and c not in SKIP_CODES]
    numeric_codes.sort()

    for code in numeric_codes:
        name = items[code]
        item_rules = rules.get(code, [])

        if not item_rules:
            # No price rules at all
            continue

        # Try to find matching product in mother tables
        name_upper = name.upper().strip()
        found_matches = []

        # Search all table names for match
        for tname, tdata in tables.items():
            for pname in tdata:
                # Try various matching strategies
                # 1. Exact match
                if name_upper == pname:
                    found_matches.append((tname, pname))
                    continue
                # 2. Focco name contained in mother name or vice versa
                # Extract the "core" name (after SOFA/POLTRONA/etc)
                # e.g. "SOFA ALOFI 1BR" -> "ALOFI 1BR"
                # Mother: "ALOFI 1B 0,23 - 2,03X1,00 2AS" -> starts with "ALOFI 1B"
                # Check if the first words match
                focco_words = name_upper.split()
                mother_words = pname.split()
                if len(focco_words) >= 2 and len(mother_words) >= 2:
                    if focco_words[0] == mother_words[0] and focco_words[1] == mother_words[1]:
                        found_matches.append((tname, pname))

        if not found_matches:
            no_match.append((code, name, len(item_rules)))
            continue

        # For each focco rule, find the corresponding mother price
        item_discrepancies = []
        item_missing_grades = []

        for (dim_str, grade, focco_price) in item_rules:
            focco_dim = dim_from_focco(dim_str) if dim_str else None
            norm_grade = GRADE_NORMALIZE.get(grade, grade)

            # Find matching price in any mother table
            mother_price = None
            mother_source = None
            mother_prod = None

            for (tname, pname) in found_matches:
                tdata = tables[tname]
                prod_dims = tdata.get(pname, {})

                # Find matching dimension
                for mdim, grade_prices in prod_dims.items():
                    if focco_dim is None or dim_matches(focco_dim, mdim):
                        # Found matching dimension, now find grade
                        # Try normalized grade first
                        price = grade_prices.get(norm_grade)
                        if price is None:
                            # Try original grade
                            price = grade_prices.get(grade)
                        if price is None:
                            # Try all grade aliases
                            for g, p in grade_prices.items():
                                if GRADE_NORMALIZE.get(g) == norm_grade:
                                    price = p
                                    break

                        if price is not None:
                            mother_price = price
                            mother_source = tname
                            mother_prod = pname
                            break
                if mother_price is not None:
                    break

            if mother_price is None:
                item_missing_grades.append({
                    'code': code, 'name': name,
                    'dim': dim_str, 'grade': grade,
                    'focco_price': focco_price,
                    'matches': [(t, p) for t, p in found_matches[:2]]
                })
            else:
                diff = focco_price - mother_price
                if abs(diff) > 1:  # tolerance of 1 real
                    item_discrepancies.append({
                        'code': code, 'name': name,
                        'dim': dim_str, 'grade': grade,
                        'focco_price': focco_price,
                        'mother_price': mother_price,
                        'diff': diff,
                        'source': mother_source,
                        'mother_prod': mother_prod
                    })

        discrepancies.extend(item_discrepancies)
        missing_grades.extend(item_missing_grades)

    # ============================================================
    # PRINT REPORT
    # ============================================================
    print("\n" + "="*100)
    print("AUDIT REPORT - FOCCO vs MOTHER TABLES")
    print("="*100)

    # Items with no match
    print(f"\n{'='*60}")
    print(f"SECTION 1: ITEMS WITH NO MATCH IN ANY MOTHER TABLE ({len(no_match)} items)")
    print(f"{'='*60}")
    for code, name, rule_count in no_match:
        print(f"  {code}: {name}  ({rule_count} price rules)")

    # Price discrepancies
    print(f"\n{'='*60}")
    print(f"SECTION 2: PRICE DISCREPANCIES ({len(discrepancies)} rules)")
    print(f"{'='*60}")
    if discrepancies:
        print(f"{'CODE':<8} {'NAME':<35} {'DIM':<25} {'GRADE':<16} {'FOCCO':>8} {'MOTHER':>8} {'DIFF':>8} {'SOURCE'}")
        print("-"*130)
        for d in sorted(discrepancies, key=lambda x: (x['code'], x['dim'] or '', x['grade'])):
            print(f"  {d['code']:<8} {d['name']:<35} {str(d['dim'] or ''):<25} {d['grade']:<16} "
                  f"{d['focco_price']:>8.0f} {d['mother_price']:>8.0f} {d['diff']:>+8.0f}  {d['source']}")
    else:
        print("  No price discrepancies found!")

    # Missing grade mappings
    print(f"\n{'='*60}")
    print(f"SECTION 3: MISSING GRADE MAPPINGS ({len(missing_grades)} rules)")
    print(f"(Focco has a price but no corresponding grade found in mother table)")
    print(f"{'='*60}")
    if missing_grades:
        # Group by code
        by_code = defaultdict(list)
        for m in missing_grades:
            by_code[m['code']].append(m)
        for code in sorted(by_code.keys()):
            entries = by_code[code]
            name = entries[0]['name']
            grades_seen = set(e['grade'] for e in entries)
            dims_seen = set(e['dim'] for e in entries)
            print(f"  {code}: {name}")
            print(f"    Grades not found: {sorted(grades_seen)}")
            print(f"    Sample dims: {list(dims_seen)[:3]}")
            print(f"    Matched to: {entries[0]['matches']}")
            print()
    else:
        print("  No missing grade mappings!")

    # Summary
    print(f"\n{'='*60}")
    print("SUMMARY")
    print(f"{'='*60}")
    print(f"  Items skipped (already corrected): {len(SKIP_CODES)}")
    print(f"  Items with no mother table match:  {len(no_match)}")
    print(f"  Price discrepancies found:         {len(discrepancies)}")
    print(f"  Missing grade mappings:            {len(missing_grades)}")

    if discrepancies:
        total_diff = sum(abs(d['diff']) for d in discrepancies)
        print(f"  Total absolute price difference:   R$ {total_diff:,.0f}")


if __name__ == '__main__':
    main()
