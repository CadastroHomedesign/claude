"""
Corrige divergências encontradas na revisão do PIU_MOBILE_HD_CADASTRO.xlsx
Produtos novos com preços errados vs tabela mãe PIURB 2026.

Erros confirmados:
1. 17477 PUFF NONNA INOX: tinha preços do ACO, deve ter INOX
2. 17475 POL NONNA ACO: COUR FORN. errado
3. 17476 PUFF NONNA ACO: COUR FORN. errado
4. 17479 POL NAVE FIXA: todos os 10 grades errados
5. 17478 SOFA NAVE 2B: COUR FORN. errado em todas as 5 dims
6. 17480 POL NAVE SB: COUR FORN. errado em todas as 5 dims

Altera apenas coluna 6 (preço). B/D preservados.
"""
import openpyxl, shutil, re

SRC = "PIU_MOBILE_HD_CADASTRO.xlsx"
DST = "PIU_MOBILE_HD_CADASTRO.xlsx"  # corrige no mesmo arquivo
wb = openpyxl.load_workbook(DST, data_only=False)
ws = wb["Itens - Regras de Preço"]

n = 0
def setv(r, v, tag):
    global n
    cur = ws.cell(r, 6).value
    if cur != v:
        ws.cell(r, 6).value = v; n += 1
        print(f"  {tag}: {cur} -> {v}")

# ── NONNA ACO POL 17475 ──────────────────────────────────────────
# Mae PIURB r406: CR VQ=7863 CR FZ=9269 COUR FORN.=5626
# CR VQ e CR FZ já estão corretos; só COUR FORN. errado
for r in range(2, ws.max_row+1):
    if str(ws.cell(r,1).value) != '17475': continue
    cond = str(ws.cell(r,3).value or '')
    if 'COUR FORN.' in cond:
        setv(r, 5626, "17475 POL NONNA ACO COUR FORN.")

# ── NONNA ACO PUFF 17476 ─────────────────────────────────────────
# Mae r407: CR VQ=4160 CR FZ=4617 COUR FORN.=3433
for r in range(2, ws.max_row+1):
    if str(ws.cell(r,1).value) != '17476': continue
    cond = str(ws.cell(r,3).value or '')
    if 'COUR FORN.' in cond:
        setv(r, 3433, "17476 PUFF NONNA ACO COUR FORN.")

# ── NONNA INOX PUFF 17477 ────────────────────────────────────────
# Mae r413: CR VQ=5141 CR FZ=5598 COUR FORN.=4415
nonna_inox = {'CR VQ': 5141, 'CR FZ': 5598, 'COUR FORN.': 4415}
for r in range(2, ws.max_row+1):
    if str(ws.cell(r,1).value) != '17477': continue
    cond = str(ws.cell(r,3).value or '')
    for grade, val in nonna_inox.items():
        if f'"{grade}' in cond or f'"{grade} ' in cond:
            setv(r, val, f"17477 PUFF NONNA INOX {grade}")
            break

# ── NAVE POL FIXA 17479 ──────────────────────────────────────────
# Mae r396: FX FORN. RB=4724 FX G=5105 FX H=5187 FX I=5290
#           FX J=5547 FX R=5650 FX V=5856 CR VQ=8213 CR FZ=10115 COUR FORN.=5187
nave_fixa = {
    'FX FORN. RB': 4724, 'FX G': 5105, 'FX H': 5187, 'FX I': 5290,
    'FX J / N': 5547, 'FX R': 5650, 'FX V': 5856,
    'CR VQ': 8213, 'CR FZ': 10115, 'COUR FORN. ': 5187
}
for r in range(2, ws.max_row+1):
    if str(ws.cell(r,1).value) != '17479': continue
    cond = str(ws.cell(r,3).value or '')
    gm = re.search(r'TECIDO_\w+="([^"]+)"', cond)
    gr = gm.group(1) if gm else None
    if gr and gr in nave_fixa:
        setv(r, nave_fixa[gr], f"17479 NAVE FIXA {gr}")

# ── NAVE 2B 17478 – COUR FORN. por dimensão ──────────────────────
# Mae PIURB: dim 240→COUR=6654, 260→6890, 280→7166, 300→7501, 320→7798
nave_2b_cour = {
    '240': 6654, '260': 6890, '280': 7166, '300': 7501, '320': 7798
}
for r in range(2, ws.max_row+1):
    if str(ws.cell(r,1).value) != '17478': continue
    cond = str(ws.cell(r,3).value or '')
    if 'COUR FORN.' not in cond: continue
    dm = re.search(r'DIMENSAO_\w+="(\d+)X', cond)
    if dm:
        dim3 = dm.group(1)[:3]
        if dim3 in nave_2b_cour:
            setv(r, nave_2b_cour[dim3], f"17478 NAVE 2B {dm.group(1)} COUR FORN.")

# ── NAVE SB 17480 – COUR FORN. por dimensão ─────────────────────
# Mae PIURB: SB 80→3982, 90→4115, 100→4274, 110→4470, 120→4605
nave_sb_cour = {
    '80': 3982, '90': 4115, '100': 4274, '110': 4470, '120': 4605
}
for r in range(2, ws.max_row+1):
    if str(ws.cell(r,1).value) != '17480': continue
    cond = str(ws.cell(r,3).value or '')
    if 'COUR FORN.' not in cond: continue
    dm = re.search(r'DIMENSAO_\w+="(\d+)X', cond)
    if dm:
        dim2 = dm.group(1).lstrip('0') or '0'
        if dim2 in nave_sb_cour:
            setv(r, nave_sb_cour[dim2], f"17480 NAVE SB {dm.group(1)} COUR FORN.")

print(f"\nTotal correções: {n}")

def _ult(w, nc):
    u = w.max_row
    while u > 1 and all(w.cell(u,c).value is None for c in range(1,nc+1)): u -= 1
    return u
for sheet, tab, nc, L in [
    ("Lista de Opções","Tabela_ListaOpcoes",5,"E"),
    ("Características","Tabela_Caracteristicas",4,"D"),
    ("Itens","Tabela_Itens",14,"N"),
    ("Itens - Atributos","Tabela_ItensAtributos",3,"C"),
    ("Itens - Regras de Preço","Tabela_ItensRegrasPreco",7,"G")]:
    w = wb[sheet]; w.tables[tab].ref = f"A1:{L}{_ult(w,nc)}"

wb.save(DST)
print(f"✓ Saved {DST}")
