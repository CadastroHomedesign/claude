# -*- coding: utf-8 -*-
"""
Correção e complemento Xique Xique — PIU Mobile HDA
---
1. Corrige preços FORNECIDO errados no item 22362 (a IA anterior colocou
   metragem de tecido em vez do preço).
2. Completa regras do item 17075 (SOFA XIQUE-XIQUE SB 1 ASS) — apenas
   2 regras existiam; acrescenta as 68 faltantes e 3 novas opções de dimensão.
3. Completa regras do item 17076 (PUFF MESA XIQUE-XIQUE) — só FX E existia;
   acrescenta 13 faixas faltantes.
4. Cadastra 8 itens novos (22370–22377) ausentes na tabela mãe.
---
REGRAS INVIOLÁVEIS observadas:
  - QUANDO/ENTÃO gravados como TEXTO PURO + fill cinza em toda linha nova/modificada.
  - Preço Base = 0; CLASSIFICACAO_ITEM = MOVEIS para itens novos.
  - .ref das 5 tabelas expandido após inserção.
  - openpyxl.load_workbook (NUNCA pandas to_excel).
"""

import shutil
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.styles.colors import Color

SRC = '/root/.claude/uploads/311cc8c5-e636-5ebe-9ac9-ac0347d05c63/5931d5df-PIU_MOBILE_HDA.xlsx'
DST = '/home/user/claude/PIU_MOBILE_HDA_CORRIGIDO.xlsx'

FILL = PatternFill(patternType="solid", fgColor=Color(theme=0, tint=-0.249977111117893))

# ── ordem das faixas (igual Lista de Opções existente em TECIDO_17070) ─────────
FAIXAS = [
    'FORNECIDO', 'FX A', 'FX B', 'FX C', 'FX D', 'FX E',
    'FX F', 'FX G', 'FX H', 'FX I', 'FX J/N', 'FX R', 'FX V', 'FX KNIT'
]

# ── preços da tabela mãe (PIU - DESIGNERS, colunas 8-21) ──────────────────────
# XIQUE-XIQUE SB (tecido) — chave = medida arredondada em cm
PRECO_SB = {
    '215': [5880, 6186, 6288, 6390, 6533, 6635, 6799, 6942, 7105, 7309, 7820, 8024, 8432, 10372],
    '235': [6127, 6461, 6572, 6683, 6839, 6950, 7128, 7284, 7461, 7684, 8240, 8462, 8907, 11020],
    '255': [6416, 6779, 6899, 7020, 7189, 7310, 7503, 7672, 7866, 8107, 8711, 8953, 9436, 11731],
    '275': [6685, 7075, 7205, 7336, 7518, 7648, 7856, 8038, 8246, 8506, 9157, 9417, 9938, 12409],
    '295': [6975, 7395, 7535, 7675, 7871, 8011, 8234, 8430, 8654, 8934, 9633, 9913, 10473, 13131],
}

# XIQUE-XIQUE SB 2 MESAS (tecido)
PRECO_SB2M = {
    '283': [7073, 7401, 7510, 7619, 7772, 7881, 8056, 8209, 8383, 8602, 9148, 9366, 9803, 11878],
    '303': [7321, 7678, 7797, 7915, 8082, 8201, 8391, 8557, 8748, 8986, 9580, 9818, 10293, 12553],
    '323': [7612, 7997, 8125, 8253, 8432, 8560, 8765, 8945, 9150, 9406, 10046, 10302, 10815, 13249],
    '343': [7869, 8282, 8420, 8557, 8750, 8888, 9108, 9300, 9521, 9796, 10484, 10759, 11309, 13924],
    '363': [8142, 8585, 8732, 8880, 9087, 9234, 9470, 9677, 9913, 10208, 10946, 11242, 11832, 14636],
}

# XIQUE-XIQUE SB MESA (tecido)
PRECO_SBMESA = {
    '249': [6482, 6792, 6896, 6999, 7144, 7247, 7413, 7558, 7724, 7931, 8448, 8655, 9069, 11036],
    '269': [6716, 7055, 7168, 7281, 7439, 7552, 7733, 7891, 8072, 8298, 8863, 9089, 9541, 11688],
    '289': [7044, 7411, 7533, 7656, 7827, 7950, 8145, 8317, 8513, 8757, 9369, 9614, 10104, 12429],
    '309': [7291, 7687, 7819, 7951, 8136, 8267, 8478, 8663, 8874, 9138, 9797, 10061, 10589, 13095],
    '329': [7586, 8012, 8153, 8295, 8494, 8636, 8862, 9061, 9288, 9571, 10280, 10564, 11131, 13825],
}

# XIQUE-XIQUE PUFF 0,88X1,12 (tecido) — item novo 22370
PRECO_PUFF = [3447, 3586, 3632, 3678, 3743, 3789, 3863, 3928, 4001, 4094, 4325, 4417, 4602, 5480]

# XIQUE-XIQUE PUFF MESA 1,22X1,12 (tecido) — completa 17076 + novo 22377
PRECO_PUFF_MESA = [4128, 4267, 4313, 4359, 4424, 4470, 4544, 4609, 4683, 4775, 5006, 5098, 5283, 6161]

# LAMINADO - SB (5 tamanhos)
PRECO_LAM_SB = {
    '215': [6293, 6575, 6669, 6762, 6894, 6988, 7138, 7270, 7420, 7608, 8077, 8265, 8640, 10425],
    '235': [6584, 6898, 7002, 7107, 7254, 7358, 7526, 7673, 7840, 8049, 8573, 8782, 9201, 11190],
    '255': [6898, 7237, 7350, 7463, 7622, 7735, 7916, 8074, 8255, 8481, 9047, 9273, 9725, 11874],
    '275': [7224, 7591, 7714, 7836, 8007, 8130, 8326, 8497, 8693, 8938, 9550, 9794, 10284, 12610],
    '295': [7553, 7946, 8077, 8207, 8390, 8521, 8731, 8914, 9123, 9385, 10039, 10300, 10823, 13309],
}

# LAMINADO - SB 2 MESAS
PRECO_LAM_SB2M = {
    '283': [7790, 8072, 8166, 8260, 8391, 8485, 8636, 8767, 8917, 9105, 9575, 9762, 10138, 11922],
    '303': [8093, 8407, 8512, 8617, 8763, 8868, 9036, 9182, 9350, 9559, 10083, 10292, 10711, 12700],
    '323': [8408, 8747, 8860, 8973, 9132, 9245, 9426, 9584, 9765, 9991, 10557, 10783, 11235, 13384],
    '343': [8704, 9071, 9194, 9316, 9487, 9610, 9806, 9977, 10173, 10418, 11030, 11274, 11764, 14090],
    '363': [9013, 9405, 9536, 9667, 9850, 9981, 10190, 10373, 10582, 10844, 11498, 11760, 12283, 14768],
}

# LAMINADO - SB MESA
PRECO_LAM_SBMESA = {
    '249': [7106, 7388, 7482, 7576, 7707, 7801, 7952, 8083, 8233, 8421, 8891, 9078, 9454, 11238],
    '269': [7383, 7697, 7801, 7906, 8053, 8157, 8325, 8471, 8639, 8848, 9372, 9581, 10000, 11989],
    '289': [7749, 8088, 8201, 8315, 8473, 8586, 8767, 8925, 9106, 9332, 9898, 10124, 10577, 12725],
    '309': [8037, 8404, 8527, 8649, 8820, 8943, 9139, 9310, 9506, 9751, 10363, 10607, 11097, 13423],
    '329': [8370, 8762, 8893, 9024, 9207, 9337, 9547, 9730, 9939, 10201, 10855, 11116, 11640, 14125],
}

# LAMINADO - PUFF 0,88X1,12
PRECO_LAM_PUFF = [3670, 3760, 3791, 3821, 3863, 3894, 3942, 3985, 4033, 4094, 4245, 4306, 4427, 5003]

# LAMINADO - PUFF MESA 1,22X1,12
PRECO_LAM_PUFF_MESA = [4514, 4605, 4635, 4665, 4708, 4738, 4786, 4829, 4877, 4938, 5089, 5150, 5271, 5847]


def _ult(ws, ncols):
    u = ws.max_row
    while u > 1 and all(ws.cell(u, c).value is None for c in range(1, ncols + 1)):
        u -= 1
    return u


def add_regra(ws, r, cod, cond, valor):
    ws.cell(r, 1, cod)
    b = ws.cell(r, 2, "QUANDO"); b.fill = FILL
    ws.cell(r, 3, cond)
    d = ws.cell(r, 4, "ENTÃO"); d.fill = FILL
    ws.cell(r, 5, "CONCEDE ACRÉSCIMO DE")
    ws.cell(r, 6, valor)
    ws.cell(r, 7, "NO PREÇO BASE")
    return r + 1


def add_opcao(wsLO, r, carac, val):
    wsLO.cell(r, 1, carac)
    wsLO.cell(r, 2, val)
    wsLO.cell(r, 3, val)
    wsLO.cell(r, 4, "SIM")
    return r + 1


def run():
    shutil.copy(SRC, DST)
    wb = load_workbook(DST)

    wsLO = wb["Lista de Opções"]
    wsCA = wb["Características"]
    wsIT = wb["Itens"]
    wsAT = wb["Itens - Atributos"]
    wsRP = wb["Itens - Regras de Preço"]

    uLO = _ult(wsLO, 5)
    uCA = _ult(wsCA, 4)
    uIT = _ult(wsIT, 14)
    uAT = _ult(wsAT, 3)
    uRP = _ult(wsRP, 7)

    cars_def = {str(wsCA.cell(r, 1).value) for r in range(2, uCA + 1) if wsCA.cell(r, 1).value}
    lo_keys = set()
    for r in range(2, uLO + 1):
        n = wsLO.cell(r, 1).value
        v = wsLO.cell(r, 2).value
        if n and v:
            lo_keys.add((str(n), str(v)))

    print("=" * 60)
    print("ETAPA 1 — Corrigir preços FORNECIDO de 22362")
    print("=" * 60)
    # Mapa: dimensão → preço correto
    fornecido_fix = {
        '215X112X085': 5880,
        '235X112X085': 6127,
        '255X112X085': 6416,
        '275X112X085': 6685,
        '295X112X085': 6975,
    }
    fixes = 0
    for r in range(2, uRP + 2):
        cod = wsRP.cell(r, 1).value
        if str(cod) != '22362':
            continue
        cond = wsRP.cell(r, 3).value or ''
        if 'FX FORNECIDO' not in cond:
            continue
        # Encontra qual dimensão
        for dim, preco_correto in fornecido_fix.items():
            if f'DIMENSAO_22362="{dim}"' in cond:
                val_atual = wsRP.cell(r, 6).value
                wsRP.cell(r, 6, preco_correto)
                # Corrige B e D para texto puro (bug =QUANDO)
                b = wsRP.cell(r, 2, "QUANDO"); b.fill = FILL
                d = wsRP.cell(r, 4, "ENTÃO"); d.fill = FILL
                print(f"  22362 {dim} FX FORNECIDO: {val_atual} → {preco_correto}")
                fixes += 1
                break
    print(f"  {fixes} preços corrigidos")

    print()
    print("=" * 60)
    print("ETAPA 2 — Completar regras de 17075 (SB 1 ASS)")
    print("=" * 60)
    # Dimensões × faixas esperadas
    dims_17075 = [
        ('215X112X85',  PRECO_SB['215']),   # formato existente (sem 0)
        ('235X112X085', PRECO_SB['235']),
        ('255X112X085', PRECO_SB['255']),
        ('275X112X085', PRECO_SB['275']),
        ('295X112X085', PRECO_SB['295']),
    ]
    # Regras já existentes em 17075
    existentes_17075 = set()
    for r in range(2, uRP + 2):
        if str(wsRP.cell(r, 1).value) == '17075':
            existentes_17075.add(str(wsRP.cell(r, 3).value))

    # Adicionar dimensões novas à Lista de Opções
    for dim, _ in dims_17075:
        key = ('DIMENSAO_17075', dim)
        if key not in lo_keys:
            uLO = add_opcao(wsLO, uLO + 1, 'DIMENSAO_17075', dim) - 1
            lo_keys.add(key)
            print(f"  LO: DIMENSAO_17075 = {dim}")

    novas = 0
    for dim, precos in dims_17075:
        for i, faixa in enumerate(FAIXAS):
            cond = f'DIMENSAO_17075="{dim}" E TECIDO_17070="{faixa}"'
            if cond in existentes_17075:
                continue
            uRP = add_regra(wsRP, uRP + 1, '17075', cond, precos[i]) - 1
            novas += 1
    print(f"  {novas} regras adicionadas ao 17075")

    print()
    print("=" * 60)
    print("ETAPA 3 — Completar regras de 17076 (PUFF MESA XIQUE-XIQUE)")
    print("=" * 60)
    existentes_17076 = set()
    for r in range(2, uRP + 2):
        if str(wsRP.cell(r, 1).value) == '17076':
            existentes_17076.add(str(wsRP.cell(r, 3).value))

    novas = 0
    for i, faixa in enumerate(FAIXAS):
        cond = f'DIMENSAO_17076="122X112" E TECIDO_17070="{faixa}"'
        if cond in existentes_17076:
            continue
        uRP = add_regra(wsRP, uRP + 1, '17076', cond, PRECO_PUFF_MESA[i]) - 1
        novas += 1
    print(f"  {novas} regras adicionadas ao 17076")

    print()
    print("=" * 60)
    print("ETAPA 4 — Cadastrar itens novos 22370–22377")
    print("=" * 60)

    # Helper para cadastrar item completo com 1 ou N dimensões
    def cadastrar_item(cod, nome, dim_carac, dim_opcoes_precos, tec_carac, faixas_list, obs1='TEM_OBSERVACAO_TECIDO'):
        nonlocal uLO, uCA, uIT, uAT, uRP
        # Opções de dimensão
        for dim_val, _ in dim_opcoes_precos:
            key = (dim_carac, dim_val)
            if key not in lo_keys:
                uLO = add_opcao(wsLO, uLO + 1, dim_carac, dim_val) - 1
                lo_keys.add(key)
        # Opções de tecido (faixas completas)
        for faixa in faixas_list:
            key = (tec_carac, faixa)
            if key not in lo_keys:
                uLO = add_opcao(wsLO, uLO + 1, tec_carac, faixa) - 1
                lo_keys.add(key)
        # Característica dimensão
        if dim_carac not in cars_def:
            wsCA.cell(uCA + 1, 1, dim_carac)
            wsCA.cell(uCA + 1, 2, "DIMENSAO:")
            wsCA.cell(uCA + 1, 3, "ESCOLHA")
            wsCA.cell(uCA + 1, 4, dim_carac)
            uCA += 1
            cars_def.add(dim_carac)
        # Característica tecido
        if tec_carac not in cars_def:
            wsCA.cell(uCA + 1, 1, tec_carac)
            wsCA.cell(uCA + 1, 2, "TECIDO:")
            wsCA.cell(uCA + 1, 3, "ESCOLHA")
            wsCA.cell(uCA + 1, 4, tec_carac)
            uCA += 1
            cars_def.add(tec_carac)
        # Item
        wsIT.cell(uIT + 1, 1, cod)
        wsIT.cell(uIT + 1, 2, nome)
        wsIT.cell(uIT + 1, 3, 0)
        wsIT.cell(uIT + 1, 4, True)
        wsIT.cell(uIT + 1, 5, dim_carac)
        wsIT.cell(uIT + 1, 6, tec_carac)
        wsIT.cell(uIT + 1, 7, obs1)
        wsIT.cell(uIT + 1, 8, "TEM_OBSERVACAO")
        uIT += 1
        # Atributo
        wsAT.cell(uAT + 1, 1, cod)
        wsAT.cell(uAT + 1, 2, "CLASSIFICACAO_ITEM")
        wsAT.cell(uAT + 1, 3, "MOVEIS")
        uAT += 1
        # Regras
        nr = 0
        for dim_val, precos in dim_opcoes_precos:
            for i, faixa in enumerate(faixas_list):
                cond = f'{dim_carac}="{dim_val}" E {tec_carac}="{faixa}"'
                uRP = add_regra(wsRP, uRP + 1, cod, cond, precos[i]) - 1
                nr += 1
        print(f"  {cod}  {nome}  ({nr} regras)")

    # Dimensões para SB (mesmo para LAMINADO SB)
    dims_sb = [
        ('215X112X085', PRECO_SB['215']),
        ('235X112X085', PRECO_SB['235']),
        ('255X112X085', PRECO_SB['255']),
        ('275X112X085', PRECO_SB['275']),
        ('295X112X085', PRECO_SB['295']),
    ]
    dims_lam_sb = [
        ('215X112X085', PRECO_LAM_SB['215']),
        ('235X112X085', PRECO_LAM_SB['235']),
        ('255X112X085', PRECO_LAM_SB['255']),
        ('275X112X085', PRECO_LAM_SB['275']),
        ('295X112X085', PRECO_LAM_SB['295']),
    ]
    dims_sb2m = [
        ('283X112X085', PRECO_SB2M['283']),
        ('303X112X085', PRECO_SB2M['303']),
        ('323X112X085', PRECO_SB2M['323']),
        ('343X112X085', PRECO_SB2M['343']),
        ('363X112X085', PRECO_SB2M['363']),
    ]
    dims_lam_sb2m = [
        ('283X112X085', PRECO_LAM_SB2M['283']),
        ('303X112X085', PRECO_LAM_SB2M['303']),
        ('323X112X085', PRECO_LAM_SB2M['323']),
        ('343X112X085', PRECO_LAM_SB2M['343']),
        ('363X112X085', PRECO_LAM_SB2M['363']),
    ]
    dims_sbmesa = [
        ('249X112X085', PRECO_SBMESA['249']),
        ('269X112X085', PRECO_SBMESA['269']),
        ('289X112X085', PRECO_SBMESA['289']),
        ('309X112X085', PRECO_SBMESA['309']),
        ('329X112X085', PRECO_SBMESA['329']),
    ]
    dims_lam_sbmesa = [
        ('249X112X085', PRECO_LAM_SBMESA['249']),
        ('269X112X085', PRECO_LAM_SBMESA['269']),
        ('289X112X085', PRECO_LAM_SBMESA['289']),
        ('309X112X085', PRECO_LAM_SBMESA['309']),
        ('329X112X085', PRECO_LAM_SBMESA['329']),
    ]

    cadastrar_item('22370', 'PUFF XIQUE-XIQUE',
                   'DIMENSAO_22370', [('88X112', PRECO_PUFF)],
                   'TECIDO_22370', FAIXAS)

    cadastrar_item('22371', 'SOFA XIQUE XIQUE SB 2 MESAS',
                   'DIMENSAO_22371', dims_sb2m,
                   'TECIDO_22371', FAIXAS)

    cadastrar_item('22372', 'SOFA XIQUE XIQUE SB MESA',
                   'DIMENSAO_22372', dims_sbmesa,
                   'TECIDO_22372', FAIXAS)

    cadastrar_item('22373', 'SOFA XIQUE XIQUE LAMINADO SB',
                   'DIMENSAO_22373', dims_lam_sb,
                   'TECIDO_22373', FAIXAS)

    cadastrar_item('22374', 'SOFA XIQUE XIQUE LAMINADO SB 2 MESAS',
                   'DIMENSAO_22374', dims_lam_sb2m,
                   'TECIDO_22374', FAIXAS)

    cadastrar_item('22375', 'SOFA XIQUE XIQUE LAMINADO SB MESA',
                   'DIMENSAO_22375', dims_lam_sbmesa,
                   'TECIDO_22375', FAIXAS)

    cadastrar_item('22376', 'PUFF XIQUE-XIQUE LAMINADO',
                   'DIMENSAO_22376', [('88X112', PRECO_LAM_PUFF)],
                   'TECIDO_22376', FAIXAS)

    cadastrar_item('22377', 'PUFF MESA XIQUE-XIQUE LAMINADO',
                   'DIMENSAO_22377', [('122X112', PRECO_LAM_PUFF_MESA)],
                   'TECIDO_22377', FAIXAS)

    print()
    print("=" * 60)
    print("ETAPA 5 — Expandir .ref das 5 tabelas")
    print("=" * 60)
    TABELAS = [
        ("Lista de Opções",         "Tabela_ListaOpcoes",     5, "E"),
        ("Características",         "Tabela_Caracteristicas", 4, "D"),
        ("Itens",                   "Tabela_Itens",           14, "N"),
        ("Itens - Atributos",       "Tabela_ItensAtributos",  3, "C"),
        ("Itens - Regras de Preço", "Tabela_ItensRegrasPreco",7, "G"),
    ]
    for sheet, tab, nc, L in TABELAS:
        ws = wb[sheet]
        old = ws.tables[tab].ref
        novo = f"A1:{L}{_ult(ws, nc)}"
        ws.tables[tab].ref = novo
        print(f"  {tab}: {old} → {novo}")

    wb.save(DST)
    print()
    print(f"Arquivo salvo: {DST}")


if __name__ == '__main__':
    run()
