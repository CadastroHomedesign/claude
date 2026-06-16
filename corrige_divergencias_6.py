"""
Correção 6: alinha as divergências restantes à tabela mãe (valores verificados
linha-a-linha contra a linha/dimensão CORRETA, não a do relatório bruto).
Trabalha sobre PIU_MOBILE_HDA_CORRIGIDO5.xlsx -> PIU_MOBILE_HDA_CORRIGIDO6.xlsx

Altera APENAS a coluna 6 (preço); B/D (QUANDO/ENTÃO) e fills preservados.
"""

import openpyxl, shutil

SRC = "PIU_MOBILE_HDA_CORRIGIDO5.xlsx"
DST = "PIU_MOBILE_HDA_CORRIGIDO6.xlsx"
shutil.copy2(SRC, DST)
wb = openpyxl.load_workbook(DST, data_only=False)
ws = wb["Itens - Regras de Preço"]

# DONNA 3,12 -> mapeia TECIDO A..J para FX A..FX J/N
donna = {'A':6142,'B':6278,'C':6414,'D':6603,'E':6739,'F':6956,
         'G':7145,'H':7362,'I':7633,'J':8311}
# KIM FORNECIDO por dimensão (estava com valor da FX A)
kim = {'0,93X2,08':4662,'1,43X2,08':5097,'1,63X2,18':5283,'1,83X2,18':5445,'1,98X2,23':5571}
# SAPUCAIA FORNECIDO por dimensão
sap = {'0,98X2,22':4021,'1,48X2,22':4429,'1,68X2,32':4657,'1,88X2,32':4879,'2,03X2,37':5092}

n = 0
def setv(r, v, tag):
    global n
    if ws.cell(r,6).value != v:
        ws.cell(r,6).value = v; n += 1
        print(f"  {tag}: -> {v}")

for r in range(2, ws.max_row+1):
    cod = str(ws.cell(r,1).value or '')
    cond = str(ws.cell(r,3).value or '')

    if cod in ('22029','22030') and 'TECIDO_SOFA_DONNA=' in cond:
        for k,v in donna.items():
            if f'TECIDO_SOFA_DONNA="TECIDO {k}"' in cond:
                setv(r, v, f"{cod} DONNA TECIDO {k}"); break

    elif cod == '17077':
        if 'TECIDO_17070="FX J/N"' in cond: setv(r, 4798, "17077 SINHO FX J/N")
        elif 'TECIDO_17070="FX KNIT"' in cond: setv(r, 5655, "17077 SINHO FX KNIT")

    elif cod == '17068' and 'TECIDO_17067="G"' in cond:
        setv(r, 4503, "17068 SINHO FX G")

    elif cod == '17067' and 'TECIDO_17067="G"' in cond:
        setv(r, 2709, "17067 TULIPA GIR FX G")

    elif cod == '22006' and 'FORNECIDO' in cond:
        setv(r, 5350, "22006 REGIA POL FORNECIDO")

    elif cod == '22036' and 'FORNECIDO' in cond:
        setv(r, 7862, "22036 CRIO FORNECIDO")

    elif cod == '17066' and 'FORNECIDO' in cond:
        setv(r, 7587, "17066 CRIO ABA FORNECIDO")

    elif cod == '22196' and 'FX FORN RB' in cond:
        setv(r, 5204, "22196 MARAA PUFF FORN RB")

    elif cod == '22307' and 'FX FORN RB' in cond:
        setv(r, 5232, "22307 MARAA CTO FORN RB")

    elif cod == '22067' and 'FORNECIDO' in cond:
        for d,v in kim.items():
            if d in cond: setv(r, v, f"22067 KIM {d} FORN"); break

    elif cod == '22072' and 'FORNECIDO' in cond:
        for d,v in sap.items():
            if d in cond: setv(r, v, f"22072 SAPUCAIA {d} FORN"); break

print(f"\nTotal corrigido: {n}")

def _ult(w, ncols):
    u = w.max_row
    while u > 1 and all(w.cell(u,c).value is None for c in range(1,ncols+1)):
        u -= 1
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
