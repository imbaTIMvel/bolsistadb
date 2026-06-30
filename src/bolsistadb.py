"""
DADOS BANCÁRIOS — Automação de tratamento de dados bancários dos alunos bolsistas
Baseado na proposta de automação e seguindo a arquitetura do SmartPC.
"""

import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill
from copy import copy
import unicodedata
import re
import os
import sys
import subprocess
from datetime import datetime, date

from PySide6.QtWidgets import (
    QApplication, QWidget, QPushButton, QLabel, QFileDialog,
    QMessageBox, QVBoxLayout, QHBoxLayout, QFrame, QScrollArea,
    QCheckBox
)
from PySide6.QtGui import QIcon, QPixmap, QPainter, QColor
from PySide6.QtCore import Qt, QThread, Signal, QRectF


# ==============================
# BANCO DE DADOS FEBRABAN
# Código → Nome canônico
# Fonte: lista FEBRABAN (principais bancos BR)
# ==============================

BANCOS_FEBRABAN = {
    "001": "Banco do Brasil S.A.",
    "003": "Banco da Amazônia S.A.",
    "004": "Banco do Nordeste do Brasil S.A.",
    "021": "Banestes S.A.",
    "025": "Banco Alfa S.A.",
    "033": "Banco Santander (Brasil) S.A.",
    "036": "Banco Bradesco BBI S.A.",
    "037": "Banco do Estado do Pará S.A.",
    "041": "Banco do Estado do Rio Grande do Sul S.A.",
    "047": "Banco do Estado de Sergipe S.A.",
    "062": "Hipercard Banco Múltiplo S.A.",
    "069": "Banco Crefisa S.A.",
    "070": "Banco de Brasília S.A. — BRB",
    "077": "Banco Inter S.A.",
    "084": "Uniprime Norte do Paraná",
    "085": "Cooperativa Central de Crédito Urbano — Cecred",
    "089": "Cooperativa de Crédito Rural da Região da Mogiana",
    "097": "Cooperativa Central de Crédito Noroeste Brasileiro",
    "099": "Uniprime Central",
    "104": "Caixa Econômica Federal",
    "107": "Banco Bocom BBM S.A.",
    "114": "Central Cooperativa de Crédito no Estado do Espírito Santo",
    "133": "Cresol Confederação",
    "136": "Unicred do Brasil",
    "197": "Stone Pagamentos S.A.",
    "208": "Banco BTG Pactual S.A.",
    "212": "Banco Original S.A.",
    "218": "Banco BS2 S.A.",
    "224": "Banco Fibra S.A.",
    "237": "Banco Bradesco S.A.",
    "243": "Banco Máxima S.A.",
    "246": "Banco ABC Brasil S.A.",
    "260": "Nu Pagamentos S.A. — Nubank",
    "290": "Pagseguro Internet S.A.",
    "318": "Banco BMG S.A.",
    "320": "China Construction Bank (Brasil) Banco Múltiplo S.A.",
    "336": "Banco C6 S.A.",
    "341": "Itaú Unibanco S.A.",
    "348": "Banco XP S.A.",
    "364": "Gerencianet Pagamentos do Brasil Ltda.",
    "376": "Banco J. P. Morgan S.A.",
    "389": "Banco Mercantil do Brasil S.A.",
    "394": "Banco Bradesco Financiamentos S.A.",
    "399": "Kirton Bank S.A.",
    "422": "Banco Safra S.A.",
    "456": "Banco MUFG Brasil S.A.",
    "473": "Banco Caixa Geral — Brasil S.A.",
    "477": "Citibank N.A.",
    "487": "Deutsche Bank S.A.",
    "505": "Banco Credit Suisse (Brasil) S.A.",
    "600": "Banco Luso Brasileiro S.A.",
    "604": "Banco Industrial do Brasil S.A.",
    "610": "Banco VR S.A.",
    "611": "Banco Paulista S.A.",
    "613": "Omni Banco S.A.",
    "623": "Banco Pan S.A.",
    "626": "Banco C6 Consignado S.A.",
    "630": "Banco Smartbank S.A.",
    "633": "Banco Rendimento S.A.",
    "634": "Banco Triângulo S.A.",
    "637": "Banco Sofisa S.A.",
    "643": "Banco Pine S.A.",
    "652": "Itaú Unibanco Holding S.A.",
    "653": "Banco Indusval S.A.",
    "654": "Banco A.J.Renner S.A.",
    "655": "Banco Votorantim S.A.",
    "707": "Banco Daycoval S.A.",
    "719": "Banif — Banco Internacional do Funchal (Brasil) S.A.",
    "735": "Banco Neon S.A.",
    "739": "Banco Cetelem S.A.",
    "741": "Banco Ribeirão Preto S.A.",
    "743": "Banco Semear S.A.",
    "745": "Banco Citibank S.A.",
    "746": "Banco Modal S.A.",
    "747": "Banco Rabobank International Brasil S.A.",
    "748": "Sicredi",
    "751": "Scotiabank Brasil S.A.",
    "752": "Banco BNP Paribas Brasil S.A.",
    "753": "Novo Banco Continental S.A.",
    "754": "Banco Sistema S.A.",
    "755": "Bank of America Merrill Lynch Banco Múltiplo S.A.",
    "756": "Bancoob / Sicoob",
    "757": "Banco KEB Hana do Brasil S.A.",
    "323": "Mercado Pago Instituição de Pagamento Ltda.",
    "380": "PicPay Serviços S.A.",
}

# Alias de nomes comuns para facilitar o fuzzy match
ALIASES_BANCOS = {
    "bb": "001",
    "banco do brasil": "001",
    "banese": "047",
    "banpara": "037",
    "banrisul": "041",
    "bradesco": "237",
    "btg": "208",
    "btg pactual": "208",
    "c6": "336",
    "c6 bank": "336",
    "caixa": "104",
    "cef": "104",
    "caixa economica": "104",
    "caixa economica federal": "104",
    "citibank": "477",
    "cresol": "133",
    "daycoval": "707",
    "inter": "077",
    "banco inter": "077",
    "itau": "341",
    "itaú": "341",
    "itau unibanco": "341",
    "nubank": "260",
    "nu": "260",
    "neon": "735",
    "original": "212",
    "pagbank": "290",
    "pagseguro": "290",
    "pan": "623",
    "santander": "033",
    "sicoob": "756",
    "bancoob": "756",
    "sicredi": "748",
    "stone": "197",
    "uniprime": "084",
    "votorantim": "655",
    "xp": "348",
    "xp investimentos": "348",
    "mercado pago": "323",
    "mercadopago": "323",
    "picpay": "380",
    "pic pay": "380",
}


# ==============================
# UTILITÁRIOS DE TEXTO
# ==============================

def normalizar_texto(texto):
    if texto is None or (isinstance(texto, float) and pd.isna(texto)):
        return ""
    texto = str(texto).strip().lower()
    texto = unicodedata.normalize("NFKD", texto).encode("ASCII", "ignore").decode("ASCII")
    texto = texto.replace("º", "o").replace("°", "o")
    return texto


def normalizar_nome(nome):
    """Normaliza nome para comparação: strip, upper, remove espaços duplos."""
    if nome is None or (isinstance(nome, float) and pd.isna(nome)):
        return ""
    nome = str(nome).strip().upper()
    nome = re.sub(r'\s+', ' ', nome)
    return nome


def normalizar_colunas(df):
    df.columns = [normalizar_texto(c) for c in df.columns]
    return df


def encontrar_coluna_por_keywords(df, keywords, excluir_keywords=None, obrigatoria=True):
    """
    Busca colunas que contenham TODAS as keywords fornecidas.
    Retorna a primeira coluna (não vazia) que satisfaz a condição.
    """
    colunas_norm = {normalizar_texto(c): c for c in df.columns}
    candidatos = []

    for col_norm, col_orig in colunas_norm.items():
        if all(kw in col_norm for kw in keywords):
            if excluir_keywords:
                if any(ekw in col_norm for ekw in excluir_keywords):
                    continue
            candidatos.append(col_orig)

    # Prioriza a primeira com pelo menos um valor não-vazio
    for col in candidatos:
        if df[col].notna().any() and (df[col].astype(str).str.strip() != "").any():
            return col

    # Se todas estão vazias, retorna a primeira mesmo assim
    if candidatos:
        return candidatos[0]

    if obrigatoria:
        raise Exception(f"Coluna com keywords {keywords} não encontrada.")
    return None


def ler_excel(path, **kwargs):
    ext = os.path.splitext(path)[1].lower()
    engine = "xlrd" if ext == ".xls" else "openpyxl"
    return pd.read_excel(path, engine=engine, dtype=str, **kwargs)


# ==============================
# TRATAMENTO DE BANCO
# ==============================

def _similaridade_levenshtein(a, b):
    """Distância de Levenshtein simples para strings curtas."""
    if len(a) == 0: return len(b)
    if len(b) == 0: return len(a)
    matriz = [[0] * (len(b) + 1) for _ in range(len(a) + 1)]
    for i in range(len(a) + 1): matriz[i][0] = i
    for j in range(len(b) + 1): matriz[0][j] = j
    for i in range(1, len(a) + 1):
        for j in range(1, len(b) + 1):
            custo = 0 if a[i - 1] == b[j - 1] else 1
            matriz[i][j] = min(
                matriz[i - 1][j] + 1,
                matriz[i][j - 1] + 1,
                matriz[i - 1][j - 1] + custo
            )
    return matriz[len(a)][len(b)]


def _score_similaridade(a, b):
    """Retorna score 0..1 (1 = idêntico) usando Levenshtein."""
    dist = _similaridade_levenshtein(a, b)
    max_len = max(len(a), len(b), 1)
    return 1.0 - dist / max_len


def resolver_banco(valor_bruto):
    """
    Recebe entrada bruta do campo banco/código.
    Retorna (codigo, nome_canonico, confianca, alerta).

    Estratégias em cascata — priorizando identificação pelo NOME:
    1. Alias exato (nomes comuns e abreviações conhecidas)
    2. Fuzzy match contra nomes canônicos FEBRABAN (Levenshtein)
    3. Código numérico extraído da string (último recurso)
    4. Falha com confiança = 0

    A priorização por nome evita que um código numérico presente na string
    (mas referente a outra versão do formulário ou digitação errônea) prevaleça
    sobre o nome claramente legível do banco.
    """
    if not valor_bruto or normalizar_texto(valor_bruto) == "":
        return ("", "", 0.0, "Campo banco em branco")

    entrada = normalizar_texto(valor_bruto)

    # Palavras genéricas a ignorar no fuzzy
    _GENERICAS = r'\b(banco|s\.?a\.?|ltda|do|da|de|e|brasil|instituicao|pagamento|cooperativo|cooperativa)\b'

    # 1. Alias exato (inclui abreviações e nomes populares)
    for alias, codigo in ALIASES_BANCOS.items():
        if alias in entrada:
            return (codigo, BANCOS_FEBRABAN[codigo], 1.0, "")

    # 2. Fuzzy match contra nomes canônicos FEBRABAN
    entrada_limpa = re.sub(_GENERICAS, '', entrada).strip() or entrada
    melhor_score  = 0.0
    melhor_codigo = ""
    melhor_nome   = ""

    for codigo, nome in BANCOS_FEBRABAN.items():
        nome_limpo = re.sub(_GENERICAS, '', normalizar_texto(nome)).strip()
        score = _score_similaridade(entrada_limpa, nome_limpo)
        if score > melhor_score:
            melhor_score  = score
            melhor_codigo = codigo
            melhor_nome   = nome

    THRESHOLD_FUZZY = 0.60
    if melhor_score >= THRESHOLD_FUZZY:
        alerta = (f"Banco identificado por similaridade ({melhor_score:.0%}): verificar"
                  if melhor_score < 0.95 else "")
        return (melhor_codigo, melhor_nome, melhor_score, alerta)

    # 3. Código numérico (último recurso — apenas quando nome não identificou)
    codigos_encontrados = re.findall(r'\b(\d{3,4})\b', entrada)
    for codigo_raw in codigos_encontrados:
        for codigo in [codigo_raw.zfill(3), codigo_raw.lstrip('0').zfill(3)]:
            if codigo in BANCOS_FEBRABAN:
                return (codigo, BANCOS_FEBRABAN[codigo], 0.80,
                        "Identificado apenas pelo código numérico — confirmar nome")

    return ("", str(valor_bruto).strip(), 0.0,
            f"Banco não identificado na tabela FEBRABAN: '{valor_bruto}'")


# ==============================
# TRATAMENTO DE AGÊNCIA E CONTA
# ==============================

# ==============================
# SPECS BANCÁRIOS
# Fonte: documentação oficial de cada banco + análise dos dados reais
# ag_digits  → comprimento canônico do número da agência (sem dígito)
# ag_has_dig → se a agência tem dígito verificador próprio
# cc_digits  → comprimento canônico do número da conta (sem dígito)
# cc_has_dig → se a conta tem dígito verificador próprio
# ==============================

# ─────────────────────────────────────────────────────────────────
# SPECS_BANCARIOS
# Fonte: "Regras de Protocolação e Validação Bancária" (documento oficial)
#        + análise dos dados reais do formulário.
#
# Estrutura: código → (ag_digits, ag_dv_fixo, cc_digits, cc_dv_fixo, nivel)
#
#   ag_digits  : nº de dígitos do número da agência (sem DV). None = variável.
#   ag_dv_fixo : valor fixo do DV da agência (ex: "0", "9") ou None (variável).
#                "FIXO_AG" = agência inteira é fixa (ex: "0001").
#   cc_digits  : nº de dígitos do número da conta (sem DV). None = variável.
#   cc_dv_fixo : valor fixo do DV da conta ou None (variável).
#   nivel      : 1=regra específica (validação completa)
#                2=regra genérica   (validação parcial)
#                3=exceção          (sem regra — conferência manual)
#
# Nota sobre o Banrisul (041): contas chegam no formato "35.051947.0-1"
# onde ".0" é sufixo de produto (descartado). Ag tem DV de 2 dígitos (DD).
#
# "ag_dv_fixo = None" significa DV variável (qualquer dígito 0-9).
# "ag_fixo"   indica agência com número fixo conforme o PDF (ex: 0001).
# ─────────────────────────────────────────────────────────────────

# Sentinel para indicar que a agência inteira é fixa (número + DV)
_AG_FIXA = "AG_FIXA"

SPECS_BANCARIOS = {
    # cód: (ag_digits, ag_dv_fixo, cc_digits, cc_dv_fixo, nivel)
    # ── Nível 1: regra específica do PDF ─────────────────────────
    "001": (4, None,    8, None,   1),  # BB          Ag:DDDD-D  CC:DDDDDDDD-D
    "033": (4, "0",     8, None,   1),  # Santander   Ag:DDDD-0  CC:DDDDDDDD-D
    "041": (4, None,    9, None,   1),  # Banrisul    Ag:DDDD-DD CC:DDDDDDDDD-D (DV ag = 2 chars)
    "070": (4, "0",     6, None,   1),  # BRB         Ag:DDDD-0  CC:DDDDDD-D
    "077": (4, "9",     7, None,   1),  # Inter       Ag:0001-9  CC:DDDDDDD-D  (ag fixa 0001)
    "104": (4, "0",     8, None,   1),  # Caixa       Ag:DDDD-0  CC:DDDDDDDD-D
    "133": (4, "0",     7, None,   1),  # Cresol      Ag:DDDD-0  CC:DDDDDDD-D
    "136": (4, "0",     7, None,   1),  # Unicred     Ag:DDDD-0  CC:DDDDDDD-D
    "197": (4, "0",     8, None,   1),  # Stone       Ag:0001-0  CC:DDDDDDDD-D (ag fixa 0001)
    "208": (4, "0",     8, None,   1),  # BTG         Ag:DDDD-0  CC:DDDDDDDD-D
    "212": (4, "0",     8, None,   1),  # Original    Ag:0001-0  CC:DDDDDDDD-D (ag fixa 0001)
    "237": (4, None,    7, None,   1),  # Bradesco    Ag:DDDD-D  CC:DDDDDDD-D
    "260": (4, "0",     8, None,   1),  # Nubank      Ag:0001-0  CC:DDDDDDDD-D (ag fixa 0001)
    "290": (4, "0",     8, None,   1),  # PagBank     Ag:0001-0  CC:DDDDDDDD-D (ag fixa 0001)
    "318": (4, "0",     7, None,   1),  # BMG         Ag:DDDD-0  CC:DDDDDDD-D
    "323": (4, "0",     8, None,   1),  # MercadoPago Ag:0001-0  CC:DDDDDDDD-D (ag fixa 0001)
    "336": (4, "0",     6, None,   1),  # C6          Ag:0001-0  CC:DDDDDD-D   (ag fixa 0001)
    "341": (4, "0",     5, None,   1),  # Itaú        Ag:DDDD-0  CC:DDDDD-D
    "348": (4, "0",     8, None,   1),  # XP          Ag:0001-0  CC:DDDDDDDD-D (ag fixa 0001)
    "364": (4, "0",     8, None,   1),  # Efí         Ag:0001-0  CC:DDDDDDDD-D (ag fixa 0001)
    "380": (4, "0",     8, None,   1),  # PicPay      Ag:0001-0  CC:DDDDDDDD-D (ag fixa 0001)
    "422": (4, "0",     7, None,   1),  # Safra       Ag:DDDD-0  CC:DDDDDDD-D
    "623": (4, "0",     8, None,   1),  # PAN         Ag:DDDD-0  CC:DDDDDDDD-D
    "748": (4, "0",     7, None,   1),  # Sicredi     Ag:DDDD-0  CC:DDDDDDD-D
    "756": (4, "0",     7, None,   1),  # Sicoob      Ag:DDDD-0  CC:DDDDDDD-D
    # ── Nível 2: bancos conhecidos sem regra consolidada no PDF ──
    "004": (4, None,    8, None,   2),  # BNB
    "037": (4, None,    9, None,   2),  # Banpará
    "047": (4, None,    9, None,   2),  # Banese
    "084": (4, None,    8, None,   2),  # Uniprime
    "085": (4, None,    8, None,   2),  # Cecred
    "197": (4, "0",     8, None,   1),  # Stone (já em N1, mantido por completude)
    "290": (4, "0",     8, None,   1),  # PagBank (idem)
    "318": (4, "0",     7, None,   1),  # BMG (idem)
    "633": (4, None,    8, None,   2),  # Banco Rendimento
    "637": (4, None,    8, None,   2),  # Sofisa
    "655": (4, None,    8, None,   2),  # Votorantim
    "707": (4, None,    8, None,   2),  # Daycoval
    "735": (4, None,    8, None,   2),  # Neon
    "739": (4, None,    8, None,   2),  # Cetelem
    "746": (4, None,    8, None,   2),  # Modal
}

# Agências fixas por banco (conforme PDF: 0001 para bancos digitais)
# Comprimento do DV da agência quando não é fixo e pode ter mais de 1 dígito.
# Ausente = 1 dígito (padrão).
AG_DV_LEN = {
    "041": 2,   # Banrisul: Ag. DDDD-DD
}


AG_FIXA = {
    "077": ("0001", "9"),   # Inter
    "197": ("0001", "0"),   # Stone
    "212": ("0001", "0"),   # Original
    "260": ("0001", "0"),   # Nubank
    "290": ("0001", "0"),   # PagBank
    "323": ("0001", "0"),   # Mercado Pago
    "336": ("0001", "0"),   # C6
    "348": ("0001", "0"),   # XP
    "364": ("0001", "0"),   # Efí
    "380": ("0001", "0"),   # PicPay
}


def validar_bancario(ag_num, ag_dv, cc_num, cc_dv, codigo_banco):
    """
    Aplica as regras de protocolação bancária conforme o documento oficial.

    Retorna:
        ag_num_out  : número da agência padronizado
        ag_dv_out   : DV da agência (corrigido se fixo)
        cc_num_out  : número da conta padronizado
        cc_dv_out   : DV da conta
        status      : 0=sem dados, 1=validado específico, 2=validado genérico, 3=exceção
        excecao     : string descritiva da exceção (vazia se OK)
    """
    # Sem dados bancários
    if not ag_num and not cc_num:
        return ag_num, ag_dv, cc_num, cc_dv, 0, ""

    spec = SPECS_BANCARIOS.get(codigo_banco)

    # ── Nível 3: banco sem regra cadastrada ──────────────────────
    if spec is None:
        excecao = (f"Banco {codigo_banco} sem regra de validação cadastrada — "
                   "conferência manual necessária")
        return ag_num, ag_dv, cc_num, cc_dv, 3, excecao

    ag_digits, ag_dv_fixo, cc_digits, cc_dv_fixo, nivel = spec
    problemas = []

    # ── Agência fixa (bancos digitais) ───────────────────────────
    ag_num_fixo, ag_dv_fixo_val = AG_FIXA.get(codigo_banco, (None, None))
    ag_foi_sobrescrita = False
    if ag_num_fixo:
        # Independente do que o respondente digitou, a agência é sempre fixa
        ag_num = ag_num_fixo
        ag_dv  = ag_dv_fixo_val
        ag_foi_sobrescrita = True

    else:
        # ── Validar comprimento da agência ────────────────────────
        if ag_num and ag_digits:
            ag_limpo = ag_num.lstrip("0") or "0"
            ag_num = ag_limpo.zfill(ag_digits)
            if len(ag_num) > ag_digits:
                problemas.append(
                    f"Agência com {len(ag_num)} dígitos (esperado {ag_digits}): '{ag_num}'"
                )

        # ── DV da agência fixo ────────────────────────────────────
        if ag_dv_fixo is not None and nivel == 1:
            if ag_dv and ag_dv != ag_dv_fixo:
                problemas.append(
                    f"DV da agência deveria ser '{ag_dv_fixo}' (informado: '{ag_dv}')"
                )
            ag_dv = ag_dv_fixo   # corrige para o valor fixo
        else:
            # DV variável — aplicar zfill se o banco usa DV com mais de 1 dígito
            dv_len = AG_DV_LEN.get(codigo_banco, 1)
            if ag_dv and dv_len > 1:
                ag_dv = ag_dv.zfill(dv_len)

    # ── Validar e padronizar comprimento da conta ────────────────
    if cc_num and cc_digits:
        n_bruto = len(cc_num)
        if n_bruto == cc_digits:
            # Comprimento exato — zfill para garantir zeros à esquerda
            cc_num = cc_num.zfill(cc_digits)
        elif n_bruto == cc_digits + 1 and not cc_dv:
            # Um dígito a mais E sem DV informado pelo respondente:
            # inferir DV embutido no último dígito e reportar como inferência.
            cc_dv  = cc_num[-1]   # último dígito do número bruto
            cc_num = cc_num[:-1]  # número sem o dígito extra
            problemas.append(
                f"DV inferido do último dígito da conta (sem DV informado pelo respondente): "
                f"'{cc_dv}' — confirmar"
            )
        elif n_bruto == cc_digits + 1 and cc_dv:
            # Um dígito a mais E DV já informado pelo respondente:
            # preservar DV informado, remover o excedente do número.
            cc_num = cc_num[:-1]
        elif n_bruto < cc_digits:
            # Curta — zfill sem lstrip para preservar zeros significativos
            cc_num = cc_num.zfill(cc_digits)
        else:
            # Mais de 1 dígito além do esperado — reportar
            problemas.append(
                f"Conta com {n_bruto} dígitos (esperado {cc_digits}): '{cc_num}'"
            )

    # ── DV da conta com mais algarismos que a convenção ─────────
    # Convenção: DV sempre 1 char (exceto Banrisul=2). Se informado com mais,
    # preservar como está e reportar — não truncar nem alterar.
    dv_len_esperado = AG_DV_LEN.get(codigo_banco, 1)   # reutiliza o mapa de DV len
    # Para conta, a convenção é sempre 1 char de DV (independente do banco)
    dv_cc_len_esperado = 1
    if cc_dv and len(cc_dv) > dv_cc_len_esperado:
        problemas.append(
            f"DV da conta com {len(cc_dv)} algarismo(s) (esperado {dv_cc_len_esperado}): "
            f"'{cc_dv}' — copiado como informado"
        )
        # cc_dv já está preservado como informado — não alteramos

    # ── Sanidade: DV não pode ser igual ao número inteiro ────────
    # Compara sem zeros à esquerda para evitar falso-positivo
    # Compara número significativo (sem zeros à esquerda); ignora se ambos são "0"
    ag_sig  = ag_num.lstrip("0") if ag_num else ""
    dv_ag_sig = ag_dv.lstrip("0") if ag_dv else ""
    if ag_sig and dv_ag_sig and dv_ag_sig == ag_sig:
        problemas.append(
            f"DV da agência igual ao número da agência ('{ag_dv}') — dado suspeito"
        )
    cc_sig  = cc_num.lstrip("0") if cc_num else ""
    dv_cc_sig = cc_dv.lstrip("0") if cc_dv else ""
    if cc_sig and dv_cc_sig and dv_cc_sig == cc_sig:
        problemas.append(
            f"DV da conta igual ao número da conta ('{cc_dv}') — dado suspeito"
        )

    # ── Validar se DV da conta existe quando deveria ──────────────
    if nivel == 1 and not cc_dv and cc_num:
        problemas.append("DV da conta não informado — verificar manualmente")

    # ── Nível 2: validação parcial (sem DV fixo) ──────────────────
    if nivel == 2:
        ag_sig = ag_num.lstrip("0") or "0"
        cc_sig = cc_num.lstrip("0") or "0" if cc_num else ""
        if ag_num and (len(ag_sig) < 4 or len(ag_sig) > 6):
            problemas.append(f"Agência fora do intervalo esperado (4-6 dígitos): '{ag_num}'")
        if cc_num and (len(cc_sig) < 5 or len(cc_sig) > 15):
            problemas.append(f"Conta fora do intervalo esperado (5-15 dígitos): '{cc_num}'")

    # Montar string de exceção e definir status final
    # Quando a agência foi sobrescrita pelo valor fixo, os problemas de DV
    # de agência já foram resolvidos — não reportar como exceção.
    if ag_foi_sobrescrita:
        problemas = [p for p in problemas if "agência" not in p.lower() or "conta" in p.lower()]

    # Avisos "suaves" (apenas DV ou DV não informado) não rebaixam para status 3
    # "conflitante" é grave mesmo que mencione "DV"; "suspeito" idem
    def _e_grave(p):
        if "conflitante" in p or "suspeito" in p:
            return True
        return "DV" not in p and "não informado" not in p
    problemas_graves  = [p for p in problemas if _e_grave(p)]
    problemas_avisos  = [p for p in problemas if not _e_grave(p)]

    excecao = "; ".join(problemas)
    if problemas_graves:
        status_final = 3
    elif problemas_avisos:
        status_final = nivel   # aviso registrado mas dados estruturalmente válidos
    else:
        status_final = nivel

    return ag_num, ag_dv, cc_num, cc_dv, status_final, excecao


def separar_numero_digito(valor_bruto):
    """
    Separa número e dígito de agência ou conta a partir do valor bruto.

    Corte automático no delimitador explícito (hífen, barra, ponto) — o que
    estiver à direita do último delimitador é tratado como dígito verificador.
    Sem delimitador: retorna (numero, "") e deixa a padronização posterior
    decidir com base no spec do banco.
    """
    if not valor_bruto or normalizar_texto(valor_bruto) in ("", "nan"):
        return ("", "")

    valor = str(valor_bruto).strip()

    # Remove lixo textual antes do número (ex: "Ag.3527-0", "AGÊNCIA 1899")
    valor = re.sub(r'(?i)ag[eê]ncia\.?\s*', '', valor).strip()
    valor = re.sub(r'(?i)cc?\.?\s*', '', valor).strip()

    # Banrisul: formato "35.073774.0-4", "35.051947.0", "39050756.0"
    # Estrutura: SEG1.SEG2[.SEG3][-DV]
    # Todos os segmentos separados por ponto formam o número completo
    # (ex: "35.073774.0" → "350737740", 9 dígitos).
    # O DV verificador vem após hífen; se não há hífen, o ponto final pode
    # indicar DV ausente (ex: "35.234971." → número "35234971", DV "").
    # Detectado por: ≥1 ponto, só dígitos/pontos/hífen.
    if '.' in valor and re.match(r'^[\d.\-]+$', valor):
        digito = ""
        base = valor
        # 1. Extrai DV verificador do hífen (se houver)
        if '-' in base:
            base, dig_part = base.rsplit('-', 1)
            if re.match(r'^[\dX]{1,2}$', dig_part.strip(), re.IGNORECASE):
                digito = dig_part.strip().upper()
        # 2. Concatena todos os segmentos como número (remove pontos e espaços)
        numero = re.sub(r'\D', '', base)
        if numero:
            return (numero, digito)

    # Delimitador explícito — split da DIREITA (último hífen/barra é o separador)
    for sep in ["-", "/"]:
        if sep in valor:
            partes = valor.rsplit(sep, 1)
            numero = re.sub(r'\D', '', partes[0])
            digito = partes[1].strip().upper()
            # Rejeita dígito com mais de 2 chars (provavelmente não é dígito)
            if len(re.sub(r'\D', '', digito)) <= 2 and len(digito) <= 2:
                return (numero, digito)
            # Se o suposto dígito é longo, trata tudo como número sem dígito
            return (re.sub(r'\D', '', valor), "")

    # Sem separador: retorna apenas dígitos, sem dígito inferido
    return (re.sub(r'\D', '', valor), "")


def padronizar_agencia(valor_bruto, codigo_banco=""):
    """
    Padroniza número de agência.
    - Corta o dígito via separar_numero_digito (hífen automático).
    - Preenche com zeros à esquerda até o comprimento canônico do banco.
    - Alerta se o comprimento não bate após a padronização.
    Retorna (agencia_str, digito_str, alerta).
    """
    numero, digito = separar_numero_digito(valor_bruto)
    alerta = ""

    if not numero:
        return ("", "", "Agência em branco ou ilegível")

    if not re.match(r'^\d+$', numero):
        alerta = f"Agência com caracteres não numéricos: '{numero}'"
        numero = re.sub(r'\D', '', numero)

    if digito and not re.match(r'^[\dXx]{1,2}$', digito):
        alerta = f"Dígito de agência inválido: '{digito}'"
        digito = ""

    # Zfill e validação de comprimento são feitos em validar_bancario()
    return (numero, digito.upper(), alerta)


def padronizar_conta(valor_bruto, codigo_banco=""):
    """
    Padroniza número de conta corrente.
    - Corta o dígito via separar_numero_digito (hífen automático).
    - Preenche com zeros à esquerda até o comprimento canônico do banco.
    - Alerta se o comprimento não bate após a padronização.
    Retorna (conta_str, digito_str, alerta).
    """
    numero, digito = separar_numero_digito(valor_bruto)
    alerta = ""

    if not numero:
        return ("", "", "Conta em branco ou ilegível")

    if not re.match(r'^\d+$', numero):
        alerta = f"Conta com caracteres não numéricos: '{numero}'"
        numero = re.sub(r'\D', '', numero)

    if digito and not re.match(r'^[\dXx]{1,2}$', digito):
        alerta = f"Dígito de conta inválido: '{digito}'"
        digito = ""

    # Zfill e validação de comprimento são feitos em validar_bancario()
    return (numero, digito.upper(), alerta)


# ==============================
# TRATAMENTO DE E-MAIL
# ==============================

def padronizar_email(valor):
    if not valor or normalizar_texto(valor) in ("", "nan"):
        return ("", "")

    email = str(valor).strip().lower()

    # Domínios não-corporativos conhecidos
    DOMINIOS_PESSOAIS = {"gmail", "hotmail", "outlook", "yahoo", "bol", "uol", "live", "icloud"}

    # Remove espaços internos
    email = re.sub(r'\s+', '', email)

    # Verifica formato básico
    match = re.match(r'^([^@]+)@([^@]+)$', email)
    if not match:
        return (email, f"E-mail com formato inválido: '{email}'")

    usuario, dominio = match.group(1), match.group(2)

    # Verifica TLD: precisa de pelo menos um ponto
    partes_dominio = dominio.split('.')
    if len(partes_dominio) < 2 or any(p == '' for p in partes_dominio):
        # Tenta corrigir se for um domínio pessoal sem TLD
        dominio_base = partes_dominio[0]
        if dominio_base in DOMINIOS_PESSOAIS:
            email_corrigido = f"{usuario}@{dominio_base}.com"
            return (email_corrigido, f"E-mail incompleto corrigido: '{valor}' → '{email_corrigido}'")
        return (email, f"Domínio de e-mail possivelmente incompleto: '{dominio}'")

    # Avisa sobre domínio pessoal (mas não rejeita)
    dominio_base = partes_dominio[0]
    if dominio_base in DOMINIOS_PESSOAIS:
        return (email, f"E-mail não institucional (pessoal): '{email}'")

    return (email, "")


# ==============================
# TRATAMENTO DE CPF
# ==============================

def padronizar_cpf(valor):
    if not valor or normalizar_texto(valor) in ("", "nan"):
        return ("", "")
    cpf = re.sub(r'\D', '', str(valor))
    if len(cpf) != 11:
        return (cpf, f"CPF com número de dígitos incorreto: '{valor}'")
    return (cpf, "")


# ==============================
# TRATAMENTO DE DATA
# ==============================

def padronizar_data(valor):
    if not valor or normalizar_texto(valor) in ("", "nan"):
        return (None, "")
    if isinstance(valor, (datetime, date)):
        return (valor if isinstance(valor, datetime) else datetime.combine(valor, datetime.min.time()), "")
    try:
        return (pd.to_datetime(str(valor), dayfirst=True), "")
    except Exception:
        return (None, f"Data não reconhecida: '{valor}'")


def calcular_maior_idade(data_nascimento):
    if data_nascimento is None:
        return None
    hoje = datetime.today()
    anos = hoje.year - data_nascimento.year - (
        (hoje.month, hoje.day) < (data_nascimento.month, data_nascimento.day)
    )
    return anos >= 18


# ==============================
# TRATAMENTO DE CEP
# ==============================

def padronizar_cep(valor):
    if not valor or normalizar_texto(valor) in ("", "nan"):
        return ("", "")
    cep = re.sub(r'\D', '', str(valor))
    if len(cep) != 8:
        return (cep, f"CEP com número de dígitos incorreto: '{valor}'")
    return (cep, "")


# ==============================
# CORRESPONDÊNCIA DE NOMES
# ==============================

def nome_esta_contido(nome_aluno, nome_formulario):
    """
    Verifica se o nome do aluno (planilha alunos.xlsx) está contido
    no nome do formulário (dados_bancarios.xlsx), após normalização.
    A proposta exige que o primeiro esteja contido no segundo.
    """
    a = normalizar_nome(nome_aluno)
    b = normalizar_nome(nome_formulario)
    if not a or not b:
        return False
    # Tokeniza e verifica se todos os tokens do nome do aluno aparecem no formulário
    tokens_a = set(a.split())
    tokens_b = set(b.split())
    return tokens_a.issubset(tokens_b)


# ==============================
# LEITURA DAS PLANILHAS
# ==============================

def ler_alunos(path):
    df = ler_excel(path)
    df = normalizar_colunas(df)
    return df


def ler_dados_bancarios(path):
    df = ler_excel(path)
    df = normalizar_colunas(df)
    return df


def _todas_cols(df, keywords, excluir=None):
    """Retorna todas as colunas que contêm todas as keywords."""
    resultado = []
    for col in df.columns:
        if all(kw in col for kw in keywords):
            if excluir and any(ek in col for ek in excluir):
                continue
            resultado.append(col)
    return resultado


def detectar_conjuntos_bancarios(df):
    """
    Detecta os conjuntos de colunas bancárias no formulário.

    Estratégia posicional direta:
    - Localiza todas as colunas com "corrente" (número de conta corrente)
      como âncora, pois esse termo é inequívoco e aparece uma vez por conjunto.
    - Para cada coluna-conta encontrada, os campos do conjunto são as 3
      colunas imediatamente anteriores (banco, agência, dígito agência)
      e a coluna seguinte (dígito conta).

    Essa abordagem é insensível ao conteúdo dos títulos e funciona mesmo
    quando os cabeçalhos são longos, ambíguos ou em português com variações.

    Sobre a coluna "INDICAR O DÍGITO DA AGÊNCIA" (col[18] neste formulário):
    Apesar do título mencionar "agência", na prática essa coluna guarda o
    dígito da CONTA para usuários de bancos digitais no set1 — e está
    posicionalmente APÓS a conta (col[17]), confirmando ser dig_cc do set1.
    Como não existe coluna de dígito da agência separada no set1, dig_ag
    fica None para esse conjunto.
    """
    colunas = list(df.columns)

    # Encontra todas as colunas de conta corrente (âncora do conjunto)
    ancoras_conta = []
    for i, col in enumerate(colunas):
        if "corrente" in col and "responsavel" not in col and "digito" not in col:
            ancoras_conta.append(i)

    conjuntos = []
    usadas = set()

    for i_conta in ancoras_conta:
        if i_conta in usadas:
            continue

        col_conta = colunas[i_conta]

        # As 3 colunas antes da conta: banco, agencia, dig_ag
        # (podem não existir se a conta estiver no início)
        def col_antes(offset):
            idx = i_conta - offset
            if idx >= 0 and idx not in usadas:
                return idx, colunas[idx]
            return None, None

        i_dag, col_dag   = col_antes(1)   # 1 antes da conta = dígito agência
        i_ag,  col_ag    = col_antes(2)   # 2 antes = agência
        i_bk,  col_banco = col_antes(3)   # 3 antes = banco

        # Coluna de dígito da conta: a SEGUINTE à conta
        i_dcc  = i_conta + 1
        col_dcc = colunas[i_dcc] if i_dcc < len(colunas) else None
        # dig_cc não pode ser uma coluna sem "digito" (não pode ser outra conta)
        if col_dcc and "digito" not in col_dcc and "corrente" in col_dcc:
            col_dcc = None
            i_dcc   = None

        # Marca como usadas para não reutilizar em outro conjunto
        for idx in [i_bk, i_ag, i_dag, i_conta, i_dcc]:
            if idx is not None:
                usadas.add(idx)

        conjuntos.append({
            "banco":   col_banco,
            "agencia": col_ag,
            "dig_ag":  col_dag,
            "conta":   col_conta,
            "dig_cc":  col_dcc,
        })

    if not conjuntos:
        conjuntos = [{"banco": None, "agencia": None, "dig_ag": None,
                      "conta": None, "dig_cc": None}]

    return conjuntos


def _valor_nao_vazio(row, col):
    """Retorna True se a célula da coluna col na linha row tem conteúdo real."""
    if col is None:
        return False
    v = row.get(col, None)
    if v is None:
        return False
    return str(v).strip().lower() not in ("", "nan", "none")


def _escolher_conjunto(row, conjuntos):
    """
    Para uma linha de dados, escolhe o conjunto bancário que tem mais campos
    preenchidos. Em caso de empate, prefere o de índice mais alto (conjuntos
    mais recentes tendem a ser versões corrigidas do formulário).
    """
    melhor_idx = 0
    melhor_score = -1
    for i, c in enumerate(conjuntos):
        score = sum(_valor_nao_vazio(row, c[k])
                    for k in ("banco", "agencia", "dig_ag", "conta", "dig_cc"))
        if score >= melhor_score:
            melhor_score = score
            melhor_idx = i
    return conjuntos[melhor_idx]


# ==============================
# NÚCLEO DO ALGORITMO
# ==============================

def processar(path_alunos, path_bancarios, path_template=None):
    """
    Executa o processamento completo.
    Retorna (df_saida, logs_erros, caminho_template_usado).
    """
    logs = []

    def log(linha, linha_planilha=None):
        prefixo = f"[Linha {linha_planilha}] " if linha_planilha else ""
        logs.append(prefixo + linha)

    # --- Leitura ---
    df_alunos = ler_alunos(path_alunos)
    df_banco = ler_dados_bancarios(path_bancarios)

    # --- Detectar colunas em alunos.xlsx ---
    col_turma   = encontrar_coluna_por_keywords(df_alunos, ["turma"])
    col_nome_al = encontrar_coluna_por_keywords(df_alunos, ["nome"])
    col_email   = encontrar_coluna_por_keywords(df_alunos, ["email"])
    col_cpf_al  = encontrar_coluna_por_keywords(df_alunos, ["cpf"])
    col_rg      = encontrar_coluna_por_keywords(df_alunos, ["rg"], obrigatoria=False)
    col_nasc_al = encontrar_coluna_por_keywords(df_alunos, ["nasc"], obrigatoria=False) or \
                  encontrar_coluna_por_keywords(df_alunos, ["data"], obrigatoria=False)
    col_end     = encontrar_coluna_por_keywords(df_alunos, ["endereco"], obrigatoria=False) or \
                  encontrar_coluna_por_keywords(df_alunos, ["logradouro"], obrigatoria=False)
    col_bairro  = encontrar_coluna_por_keywords(df_alunos, ["bairro"], obrigatoria=False)
    col_cidade  = encontrar_coluna_por_keywords(df_alunos, ["cidade"], obrigatoria=False)
    col_estado  = encontrar_coluna_por_keywords(df_alunos, ["estado"], obrigatoria=False) or \
                  encontrar_coluna_por_keywords(df_alunos, ["uf"], obrigatoria=False)
    col_cep     = encontrar_coluna_por_keywords(df_alunos, ["cep"], obrigatoria=False)

    # --- Detectar colunas em dados_bancarios.xlsx ---
    col_nome_bk  = encontrar_coluna_por_keywords(df_banco, ["nome"], excluir_keywords=["responsavel"])
    col_cpf_bk   = encontrar_coluna_por_keywords(df_banco, ["cpf"], obrigatoria=False)
    col_nasc_bk  = encontrar_coluna_por_keywords(df_banco, ["nasc"], obrigatoria=False) or \
                   encontrar_coluna_por_keywords(df_banco, ["data", "nascimento"], obrigatoria=False)
    conjuntos_bancarios = detectar_conjuntos_bancarios(df_banco)

    # Responsável legal
    col_nome_resp   = encontrar_coluna_por_keywords(df_banco, ["responsavel", "nome"], obrigatoria=False)
    col_cpf_resp    = encontrar_coluna_por_keywords(df_banco, ["responsavel", "cpf"], obrigatoria=False)
    col_email_resp  = (encontrar_coluna_por_keywords(df_banco, ["responsavel", "email"],   obrigatoria=False) or
                       encontrar_coluna_por_keywords(df_banco, ["email",       "responsavel"], obrigatoria=False) or
                       encontrar_coluna_por_keywords(df_banco, ["responsavel", "e-mail"],      obrigatoria=False) or
                       encontrar_coluna_por_keywords(df_banco, ["e-mail",      "responsavel"], obrigatoria=False))
    col_cont_resp   = encontrar_coluna_por_keywords(df_banco, ["responsavel", "contato"], obrigatoria=False) or \
                      encontrar_coluna_por_keywords(df_banco, ["celular", "responsavel"], obrigatoria=False)

    # Ordenar turmas alfabeticamente, alunos alfabeticamente dentro de cada turma
    df_alunos["_turma_norm"] = df_alunos[col_turma].apply(normalizar_texto)
    df_alunos["_nome_norm"]  = df_alunos[col_nome_al].apply(normalizar_nome)
    df_alunos = df_alunos.sort_values(["_turma_norm", "_nome_norm"]).reset_index(drop=True)

    # Índice de busca no formulário (nome normalizado → linha(s))
    df_banco["_nome_norm"] = df_banco[col_nome_bk].apply(normalizar_nome)

    linhas_saida = []
    id_por_turma = {}

    for idx_al, row_al in df_alunos.iterrows():
        linha_planilha = len(linhas_saida) + 2  # +2: cabeçalho + 1-index

        nome_al = str(row_al[col_nome_al]).strip() if pd.notna(row_al[col_nome_al]) else ""
        turma   = str(row_al[col_turma]).strip() if pd.notna(row_al[col_turma]) else ""

        # --- ID por turma ---
        if turma not in id_por_turma:
            id_por_turma[turma] = 0
        id_por_turma[turma] += 1
        id_aluno = id_por_turma[turma]

        # --- Correspondência de nome ---
        matches = df_banco[df_banco["_nome_norm"].apply(
            lambda n: nome_esta_contido(nome_al, n)
        )]

        if len(matches) == 0:
            log(f"Sem correspondência no formulário para aluno: '{nome_al}' (turma: {turma})", linha_planilha)
            row_bk = None
        elif len(matches) > 1:
            log(f"Múltiplas correspondências no formulário para aluno: '{nome_al}' — usando primeira ocorrência", linha_planilha)
            row_bk = matches.iloc[0]
        else:
            row_bk = matches.iloc[0]

        # --- E-mail ---
        email_raw = str(row_al[col_email]).strip() if col_email and pd.notna(row_al.get(col_email)) else ""
        email, alerta_email = padronizar_email(email_raw)
        if alerta_email:
            log(alerta_email, linha_planilha)

        # --- CPF ---
        cpf_al_raw = str(row_al[col_cpf_al]).strip() if pd.notna(row_al.get(col_cpf_al)) else ""
        cpf, alerta_cpf = padronizar_cpf(cpf_al_raw)
        if alerta_cpf:
            log(alerta_cpf, linha_planilha)
        if row_bk is not None and col_cpf_bk:
            cpf_bk_raw = str(row_bk.get(col_cpf_bk, "")).strip()
            cpf_bk, _ = padronizar_cpf(cpf_bk_raw)
            if cpf and cpf_bk and cpf != cpf_bk:
                log(f"CPF divergente para '{nome_al}': alunos='{cpf}' / formulário='{cpf_bk}'", linha_planilha)

        # --- RG ---
        rg = str(row_al[col_rg]).strip() if col_rg and pd.notna(row_al.get(col_rg)) else ""

        # --- Data de Nascimento ---
        nasc_raw = str(row_al[col_nasc_al]).strip() if col_nasc_al and pd.notna(row_al.get(col_nasc_al)) else ""
        data_nasc, alerta_nasc = padronizar_data(nasc_raw)
        if alerta_nasc:
            log(alerta_nasc, linha_planilha)
        if row_bk is not None and col_nasc_bk:
            nasc_bk_raw = str(row_bk.get(col_nasc_bk, "")).strip()
            data_nasc_bk, _ = padronizar_data(nasc_bk_raw)
            if data_nasc and data_nasc_bk and data_nasc.date() != data_nasc_bk.date():
                log(f"Data de nascimento divergente para '{nome_al}': alunos='{data_nasc.date()}' / formulário='{data_nasc_bk.date()}'", linha_planilha)

        maior = calcular_maior_idade(data_nasc)

        # --- Endereço ---
        partes_end = []
        for col_e in [col_end, col_bairro, col_cidade, col_estado]:
            if col_e and pd.notna(row_al.get(col_e)):
                v = str(row_al[col_e]).strip()
                if v and v.lower() not in ("nan", ""):
                    partes_end.append(v)
        endereco_completo = " - ".join(partes_end)
        # Remove hífens duplicados que possam surgir se algum campo já contiver " - "
        endereco_completo = re.sub(r'(\s*-\s*){2,}', ' - ', endereco_completo).strip(" -")

        # --- CEP ---
        cep_raw = str(row_al[col_cep]).strip() if col_cep and pd.notna(row_al.get(col_cep)) else ""
        cep, alerta_cep = padronizar_cep(cep_raw)
        if alerta_cep:
            log(alerta_cep, linha_planilha)

        # --- Responsável legal ---
        nome_resp = cpf_resp = email_resp = cont_resp = ""
        if row_bk is not None:
            if maior is False:  # menor de idade
                nome_resp  = str(row_bk.get(col_nome_resp, "")).strip() if col_nome_resp else ""
                cpf_resp   = str(row_bk.get(col_cpf_resp, "")).strip() if col_cpf_resp else ""
                email_resp = str(row_bk.get(col_email_resp, "")).strip() if col_email_resp else ""
                cont_resp  = str(row_bk.get(col_cont_resp, "")).strip() if col_cont_resp else ""
                if not nome_resp:
                    log(f"Menor de idade sem responsável informado: '{nome_al}'", linha_planilha)
            elif maior is True:
                # Maior → campos de responsável ficam em branco
                pass
            else:
                log(f"Não foi possível determinar maioridade de '{nome_al}' (data ausente/inválida)", linha_planilha)

        # --- Dados bancários ---
        # Escolhe o conjunto de colunas com mais campos preenchidos para esta linha
        nome_banco = codigo_banco = agencia = dig_ag = conta = dig_cc = ""
        orig_banco = orig_agencia = orig_dig_ag = orig_conta = orig_dig_cc = ""
        status_banco  = 0
        excecao_banco = ""
        if row_bk is not None:
            cj = _escolher_conjunto(row_bk, conjuntos_bancarios)
            # Guarda valores originais (sem nenhum tratamento) para as colunas de auditoria
            orig_banco   = str(row_bk.get(cj["banco"],   "")).strip() if cj["banco"]   else ""
            orig_agencia = str(row_bk.get(cj["agencia"], "")).strip() if cj["agencia"] else ""
            orig_dig_ag  = str(row_bk.get(cj["dig_ag"],  "")).strip() if cj["dig_ag"]  else ""
            orig_conta   = str(row_bk.get(cj["conta"],   "")).strip() if cj["conta"]   else ""
            orig_dig_cc  = str(row_bk.get(cj["dig_cc"],  "")).strip() if cj["dig_cc"]  else ""
            for v in [orig_banco, orig_agencia, orig_dig_ag, orig_conta, orig_dig_cc]:
                if v.lower() == "nan": v = ""

            # Banco
            banco_raw = str(row_bk.get(cj["banco"], "")).strip() if cj["banco"] else ""
            codigo_banco, nome_banco, confianca, alerta_banco = resolver_banco(banco_raw)
            if alerta_banco:
                log(alerta_banco, linha_planilha)

            # Agência
            ag_raw     = str(row_bk.get(cj["agencia"], "")).strip() if cj["agencia"] else ""
            dig_ag_raw = str(row_bk.get(cj["dig_ag"],  "")).strip() if cj["dig_ag"]  else ""

            agencia, dig_ag_inline, alerta_ag = padronizar_agencia(ag_raw, codigo_banco)
            if alerta_ag:
                log(alerta_ag, linha_planilha)

            if dig_ag_raw and dig_ag_raw.lower() not in ("nan", ""):
                dig_ag = dig_ag_raw.strip().upper()
            elif dig_ag_inline:
                dig_ag = dig_ag_inline

            # Conta
            cc_raw     = str(row_bk.get(cj["conta"],   "")).strip() if cj["conta"]   else ""
            dig_cc_raw = str(row_bk.get(cj["dig_cc"],  "")).strip() if cj["dig_cc"]  else ""

            conta, dig_cc_inline, alerta_cc = padronizar_conta(cc_raw, codigo_banco)
            if alerta_cc:
                log(alerta_cc, linha_planilha)
            # Se a conta ficou vazia após padronização, o dígito inline também é inválido
            if not conta:
                dig_cc_inline = ""

            if dig_cc_raw and dig_cc_raw.lower() not in ("nan", ""):
                # O valor informado pelo respondente é soberano.
                # Limpeza mínima: remover hífen inicial isolado ("-5" → "5")
                # e espaços, mas preservar o conteúdo.
                candidato = dig_cc_raw.strip().upper()
                m_hifen = re.match(r'^-([0-9X]{1,2})$', candidato)
                if m_hifen:
                    dig_cc = m_hifen.group(1)          # "-5" → "5"
                else:
                    dig_cc = candidato                  # copiar como está
            elif dig_cc_inline:
                # Coluna de DV vazia — usar DV extraído via separador explícito
                # (hífen na string da conta, ex: "38041-5"). Não é inferência:
                # o respondente embutiu o DV no campo da conta com hífen.
                dig_cc = dig_cc_inline

            # ── Validação e padronização final conforme regras bancárias ──
            agencia, dig_ag, conta, dig_cc, status_banco, excecao_banco = validar_bancario(
                agencia, dig_ag, conta, dig_cc, codigo_banco
            )

        else:
            # Aluno sem correspondência no formulário: sem dados bancários
            status_banco  = 0
            excecao_banco = ""

        # --- Montar linha de saída ---
        # Datas como datetime nativo para que o numfmt mm-dd-yy do template funcione
        linhas_saida.append({
            "No.":           len(linhas_saida) + 1,
            "ID":            id_aluno,
            "TURMA":         turma,
            "NOME COMPLETO": nome_al,
            "Data Inicio":   None,
            "Data Final":    None,
            "E-MAIL":        email,
            "CPF":           cpf,
            "RG":            rg,
            "DATA NASCIMENTO": data_nasc if data_nasc else None,
            "ENDERECO COMPLETO (Logradouro-Bairro-Cidade)": endereco_completo,
            "CEP":           cep,
            "Nome Responsavel": nome_resp,
            "CPF - Respon.": cpf_resp,
            "E-mail Respon.": email_resp,
            "Contato Respon.": cont_resp,
            "Nome do Banco": (f"{codigo_banco.zfill(3)} - {nome_banco}"
                             if codigo_banco else nome_banco),
            "Agencia":       agencia,
            "Dig.Ag":        dig_ag,
            "Conta":         conta,
            "Dig.C/C":       dig_cc,
            # Validação bancária
            "Status":        status_banco,
            "Excecoes":      excecao_banco,
            # Colunas de auditoria — dados brutos originais do formulário
            "Orig.Banco":    orig_banco   if orig_banco   not in ("nan", "")  else "",
            "Orig.Agencia":  orig_agencia if orig_agencia not in ("nan", "")  else "",
            "Orig.Dig.Ag":   orig_dig_ag  if orig_dig_ag  not in ("nan", "")  else "",
            "Orig.Conta":    orig_conta   if orig_conta   not in ("nan", "")  else "",
            "Orig.Dig.C/C":  orig_dig_cc  if orig_dig_cc  not in ("nan", "")  else "",
        })

    df_saida = pd.DataFrame(linhas_saida)
    return df_saida, logs


# ==============================
# EXPORTAÇÃO PARA XLSX
# ==============================

# Mapeamento: chave do dicionário de saída → nome da coluna na Tabela1
# (preserva os nomes exatos do template, incluindo acentos)
MAPA_COLUNAS = {
    "No.":           "No.",
    "ID":            "ID",
    "TURMA":         "TURMA",
    "NOME COMPLETO": "NOME COMPLETO",
    "Data Inicio":   "Data Início",
    "Data Final":    "Data Final",
    "E-MAIL":        "E-MAIL",
    "CPF":           "CPF",
    "RG":            "RG",
    "DATA NASCIMENTO": "DATA NASCIMENTO",
    "ENDERECO COMPLETO (Logradouro-Bairro-Cidade)": "ENDEREÇO COMPLETO (Logradouro-Bairro-Cidade)",
    "CEP":           "CEP",
    "Nome Responsavel": "Nome Responsável",
    "CPF - Respon.": "CPF - Respon.",
    "E-mail Respon.": "E-mail Respon.",
    "Contato Respon.": "Contato Respon.",
    "Nome do Banco": "Nome do Banco",
    "Agencia":       "Agência",
    "Dig.Ag":        "Díg.Ag",
    "Conta":         "Conta",
    "Dig.C/C":       "Díg.C/C",
    # Colunas de validação e auditoria (fora da Tabela1, logo após a última coluna)
    "Status":        "Status",
    "Excecoes":      "Exceções do algoritmo",
    "Orig.Banco":    "Banco (original)",
    "Orig.Agencia":  "Agência (original)",
    "Orig.Dig.Ag":   "Díg.Ag (original)",
    "Orig.Conta":    "Conta (original)",
    "Orig.Dig.C/C":  "Díg.C/C (original)",
}


def exportar_xlsx(df_saida, path_saida, path_template=None):
    """
    Copia o template, limpa a linha-modelo de dados (linha 2), insere as
    linhas novas clonando os estilos dessa linha-modelo, e expande o ref
    da Tabela1 para cobrir todas as linhas inseridas.

    Se o template não existir, emite um xlsx simples sem formatação.
    """
    import shutil

    n = len(df_saida)

    # ── COM TEMPLATE ────────────────────────────────────────────────────────
    if path_template and os.path.exists(path_template):
        shutil.copy2(path_template, path_saida)
        wb = load_workbook(path_saida)
        ws = wb.active

        # Localiza a Tabela1 e descobre a linha do cabeçalho
        tabela = ws.tables.get("Tabela1")
        if tabela is None:
            raise Exception("Tabela1 não encontrada no template.")

        # ref da tabela: ex. "A1:U2" → linha_header=1, col_ini=A, col_fim=U
        from openpyxl.utils import range_boundaries, get_column_letter
        col_min, row_header, col_max, _ = range_boundaries(tabela.ref)
        linha_modelo = row_header + 1          # linha 2: dados de exemplo
        linha_dados_ini = linha_modelo          # vamos escrever a partir daqui

        # Lê os estilos da linha-modelo antes de qualquer modificação
        estilos_modelo = {}
        for c_idx in range(col_min, col_max + 1):
            origem = ws.cell(row=linha_modelo, column=c_idx)
            estilos_modelo[c_idx] = {
                "font":          copy(origem.font),
                "fill":          copy(origem.fill),
                "border":        copy(origem.border),
                "alignment":     copy(origem.alignment),
                "number_format": origem.number_format,
            }

        # Monta mapa: nome_coluna_template → índice de coluna (1-based)
        cabecalho_para_col = {}
        for c_idx in range(col_min, col_max + 1):
            nome = ws.cell(row=row_header, column=c_idx).value
            if nome:
                cabecalho_para_col[nome] = c_idx

        # Limpa o conteúdo da linha-modelo (valor e hyperlink), mantém estilos
        for c_idx in range(col_min, col_max + 1):
            cell = ws.cell(row=linha_modelo, column=c_idx)
            cell.value     = None
            cell.hyperlink = None

        # Insere linhas extras se n > 1 (a linha 2 já existe)
        if n > 1:
            ws.insert_rows(linha_modelo + 1, n - 1)

        # Chaves que vão dentro da Tabela1 e chaves de auditoria (fora)
        CHAVES_AUDITORIA = ["Status", "Excecoes",
                            "Orig.Banco", "Orig.Agencia", "Orig.Dig.Ag",
                            "Orig.Conta", "Orig.Dig.C/C"]
        MAPA_AUDITORIA   = {k: MAPA_COLUNAS[k] for k in CHAVES_AUDITORIA}
        MAPA_TABELA      = {k: v for k, v in MAPA_COLUNAS.items()
                            if k not in CHAVES_AUDITORIA}

        # Estilo para as colunas de auditoria (fundo cinza, borda fina, texto menor)
        fill_audit  = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
        font_audit  = Font(name="Calibri", size=9, italic=True, color="444444")
        border_audit = Border(
            left=Side(style="thin"), right=Side(style="thin"),
            top=Side(style="thin"),  bottom=Side(style="thin"),
        )
        align_audit = Alignment(horizontal="center", vertical="center", wrap_text=False)

        # Cabeçalhos das colunas de auditoria (linha 1, após col_max)
        col_audit_ini = col_max + 1
        for i, (chave_aud, nome_cab) in enumerate(MAPA_AUDITORIA.items()):
            c = col_audit_ini + i
            cell = ws.cell(row=row_header, column=c, value=nome_cab)
            cell.font      = Font(name="Calibri", size=9, bold=True, color="444444")
            cell.fill      = PatternFill(start_color="BFBFBF", end_color="BFBFBF", fill_type="solid")
            cell.border    = border_audit
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            # Exceções: coluna mais larga para texto descritivo
            ws.column_dimensions[get_column_letter(c)].width = (
                6 if chave_aud == "Status" else
                45 if chave_aud == "Excecoes" else 18
            )

        # Escreve dados: Tabela1 + auditoria
        for r_offset, (_, row_df) in enumerate(df_saida.iterrows()):
            r = linha_dados_ini + r_offset
            # Colunas da Tabela1
            for chave_df, nome_template in MAPA_TABELA.items():
                c_idx = cabecalho_para_col.get(nome_template)
                if c_idx is None:
                    continue
                val = row_df.get(chave_df, None)
                if isinstance(val, str) and val.strip() == "":
                    val = None
                cell = ws.cell(row=r, column=c_idx, value=val)
                est = estilos_modelo[c_idx]
                cell.font          = copy(est["font"])
                cell.fill          = copy(est["fill"])
                cell.border        = copy(est["border"])
                cell.alignment     = copy(est["alignment"])
                cell.number_format = est["number_format"]
            # Colunas de auditoria (fora da tabela)
            for i, chave_df in enumerate(MAPA_AUDITORIA.keys()):
                c = col_audit_ini + i
                val = row_df.get(chave_df, None)
                if isinstance(val, str) and val.strip() in ("", "nan"):
                    val = None
                # Status é numérico; Excecoes e originais são texto (@)
                if chave_df == "Status":
                    val = int(val) if val not in (None, "") else None
                cell = ws.cell(row=r, column=c, value=val)
                cell.font      = copy(font_audit)
                cell.fill      = copy(fill_audit)
                cell.border    = copy(border_audit)
                cell.alignment = copy(align_audit)
                cell.number_format = "0" if chave_df == "Status" else "@"

        # Atualiza o ref da Tabela1 (não inclui as colunas de auditoria)
        ultima_linha = linha_dados_ini + n - 1
        col_ini_letra = get_column_letter(col_min)
        col_fim_letra = get_column_letter(col_max)
        tabela.ref = f"{col_ini_letra}{row_header}:{col_fim_letra}{ultima_linha}"

        wb.save(path_saida)

    # ── SEM TEMPLATE (fallback) ─────────────────────────────────────────────
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "Dados Bancários"
        ws.sheet_view.showGridLines = False

        cabecalhos = list(MAPA_COLUNAS.values())
        borda = Border(
            left=Side(style='thin'), right=Side(style='thin'),
            top=Side(style='thin'), bottom=Side(style='thin')
        )

        for c_idx, nome in enumerate(cabecalhos, 1):
            cell = ws.cell(row=1, column=c_idx, value=nome)
            cell.font      = Font(name="Calibri", bold=True, size=11)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.fill      = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
            cell.font      = Font(name="Calibri", bold=True, size=11, color="FFFFFF")
            cell.border    = borda

        for r_offset, (_, row_df) in enumerate(df_saida.iterrows(), start=2):
            for c_idx, chave_df in enumerate(MAPA_COLUNAS.keys(), 1):
                val = row_df.get(chave_df, None)
                if isinstance(val, str) and val.strip() == "":
                    val = None
                cell = ws.cell(row=r_offset, column=c_idx, value=val)
                cell.font   = Font(name="Calibri", size=11)
                cell.border = borda

        wb.save(path_saida)


def salvar_logs(logs, path_saida_xlsx):
    """Salva error_logs.txt no mesmo diretório da planilha de saída."""
    if not logs:
        return None
    dir_saida = os.path.dirname(path_saida_xlsx)
    path_log = os.path.join(dir_saida, "error_logs.txt")
    with open(path_log, "w", encoding="utf-8") as f:
        f.write(f"Registro de exceções — {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")
        f.write("=" * 60 + "\n\n")
        for linha in logs:
            f.write(linha + "\n")
    return path_log


# ==============================
# WORKER THREAD (não trava a UI)
# ==============================

class WorkerThread(QThread):
    concluido = Signal(object, list)
    erro = Signal(str)

    def __init__(self, path_alunos, path_bancarios, path_template):
        super().__init__()
        self.path_alunos = path_alunos
        self.path_bancarios = path_bancarios
        self.path_template = path_template

    def run(self):
        try:
            df, logs = processar(self.path_alunos, self.path_bancarios, self.path_template)
            self.concluido.emit(df, logs)
        except Exception as e:
            self.erro.emit(str(e))


# ==============================
# INTERFACE GRÁFICA (PySide6)
# ==============================

def resource_path(rel):
    try:
        base = sys._MEIPASS
    except Exception:
        base = os.path.abspath(".")
    return os.path.join(base, rel)


BUTTON_STYLE = """
QPushButton {
    background-color: #f9b02e;
    color: black;
    border: none;
    border-radius: 8px;
    padding: 10px;
    font-weight: bold;
    font-size: 14px;
}
QPushButton:hover { background-color: #ffd166; }
QPushButton:pressed { background-color: #e69500; }
QPushButton:disabled { background-color: #666666; color: #aaaaaa; }
"""

REMOVE_STYLE = """
QPushButton {
    background-color: #444444;
    color: #cccccc;
    border: none;
    border-radius: 8px;
    padding: 10px;
    font-size: 12px;
}
QPushButton:hover { background-color: #666666; }
"""


class ToggleSwitch(QCheckBox):
    def __init__(self, text="", parent=None):
        super().__init__(text, parent)
        self.setFixedHeight(28)
        self.setCursor(Qt.CursorShape.PointingHandCursor)

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        track_rect = QRectF(0, 4, 44, 20)
        painter.setBrush(QColor("#f9b02e") if self.isChecked() else QColor("#3a3a3a"))
        painter.setPen(Qt.PenStyle.NoPen)
        painter.drawRoundedRect(track_rect, 10, 10)
        knob_x = 24 if self.isChecked() else 2
        painter.setBrush(QColor("white"))
        painter.drawEllipse(QRectF(knob_x, 6, 16, 16))
        painter.setPen(QColor("#dddddd"))
        painter.drawText(54, 19, self.text())


def criar_janela():
    app = QApplication(sys.argv)

    window = QWidget()
    window.setWindowTitle("BolsistaDB")
    window.setWindowIcon(QIcon(resource_path("bolsistadb.ico")))
    window.resize(960, 540)
    window.setMinimumSize(700, 460)
    window.setStyleSheet("""
        QWidget { background-color: #1e1e1e; font-family: "Segoe UI"; color: white; }
        QScrollArea { border: none; background: transparent; }
        QScrollBar:vertical { background: #2b2b2b; width: 10px; border-radius: 5px; }
        QScrollBar::handle:vertical { background: #f9b02e; border-radius: 5px; }
        QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0px; }
        QCheckBox { color: #dddddd; font-size: 12px; spacing: 12px; background: transparent; }
        QCheckBox::indicator { width: 42px; height: 22px; border-radius: 11px;
                               background-color: #3a3a3a; border: 1px solid #555555; }
        QCheckBox::indicator:checked { background-color: #f9b02e; border: 1px solid #f9b02e; }
    """)

    # Fundo bg_hbr.png (mesmo recurso do SmartPC)
    bg_label = QLabel(window)
    bg_pixmap = QPixmap(resource_path("bg_hbr.png"))
    bg_label.setPixmap(bg_pixmap)
    bg_label.setScaledContents(True)

    scroll = QScrollArea(window)
    scroll.setWidgetResizable(True)
    scroll.setStyleSheet("background: transparent;")

    container = QWidget()
    container.setStyleSheet("background: transparent;")
    scroll.setWidget(container)

    card = QFrame()
    card.setObjectName("card")
    card.setStyleSheet("""
        #card {
            background-color: rgba(30,30,30,220);
            border-radius: 18px;
            border: 1px solid rgba(255,255,255,25);
        }
    """)

    # Template resolvido automaticamente a partir da pasta do script/exe
    PATH_TEMPLATE = resource_path("template.xlsx")

    # Estado dos arquivos
    state = {
        "alunos": None,
        "bancarios": None,
        "worker": None,
        "df_resultado": None,
        "logs_resultado": None,
    }

    # --- Labels ---
    label_alunos    = QLabel("-")
    label_bancarios = QLabel("-")

    def estilo_label():
        return """
            background-color: #2b2b2b;
            border: 1px solid #3a3a3a;
            border-radius: 8px;
            padding: 8px;
            color: #cccccc;
            font-size: 11px;
        """

    for lbl in [label_alunos, label_bancarios]:
        lbl.setWordWrap(True)
        lbl.setStyleSheet(estilo_label())
        lbl.setMinimumHeight(24)
        lbl.setMaximumHeight(48)

    btn_executar = QPushButton("Gerar Planilha")
    btn_executar.setFixedSize(150, 44)
    btn_executar.setStyleSheet(BUTTON_STYLE)
    btn_executar.setEnabled(False)

    status_label = QLabel("")
    status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
    status_label.setStyleSheet("color: #aaaaaa; font-size: 11px; background: transparent;")

    def atualizar_status():
        ok = bool(state["alunos"]) and bool(state["bancarios"])
        btn_executar.setEnabled(ok)

    def criar_bloco(titulo, callback_sel, callback_rem, label_widget, aceita="Excel (*.xlsx *.xls)"):
        bloco = QFrame()
        bloco.setStyleSheet("""
            QFrame { background-color: rgba(255,255,255,8); border-radius: 12px; }
        """)
        layout = QVBoxLayout(bloco)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(8)

        titulo_lbl = QLabel(titulo)
        titulo_lbl.setStyleSheet("color: white; font-size: 13px; font-weight: bold; background: transparent;")
        layout.addWidget(titulo_lbl)

        btn_row = QHBoxLayout()
        btn_sel = QPushButton("Selecionar")
        btn_sel.setFixedSize(110, 32)
        btn_sel.setStyleSheet(BUTTON_STYLE)
        btn_sel.clicked.connect(callback_sel)

        btn_rem = QPushButton("Remover")
        btn_rem.setFixedSize(110, 32)
        btn_rem.setStyleSheet(REMOVE_STYLE)
        btn_rem.clicked.connect(callback_rem)

        btn_row.addWidget(btn_sel)
        btn_row.addWidget(btn_rem)
        btn_row.addStretch()
        layout.addLayout(btn_row)
        layout.addWidget(label_widget)
        return bloco

    def sel_alunos():
        f, _ = QFileDialog.getOpenFileName(window, "Selecionar lista de bolsistas", "", "Excel (*.xlsx *.xls)")
        if f:
            state["alunos"] = f
            label_alunos.setText(os.path.basename(f))
        atualizar_status()

    def rem_alunos():
        state["alunos"] = None
        label_alunos.setText("-")
        atualizar_status()

    def sel_bancarios():
        f, _ = QFileDialog.getOpenFileName(window, "Selecionar planilha de dados bancários", "", "Excel (*.xlsx *.xls)")
        if f:
            state["bancarios"] = f
            label_bancarios.setText(os.path.basename(f))
        atualizar_status()

    def rem_bancarios():
        state["bancarios"] = None
        label_bancarios.setText("-")
        atualizar_status()

    def ao_concluir(df, logs):
        state["df_resultado"]   = df
        state["logs_resultado"] = logs

        btn_executar.setEnabled(True)
        status_label.setText("")

        n_logs = len(logs)
        msg = f"Processamento concluído com {len(df)} alunos."
        if n_logs:
            msg += f"\n{n_logs} aviso(s) registrado(s) em error_logs.txt."

        path_saida, _ = QFileDialog.getSaveFileName(
            window, "Salvar planilha de saída", "saida_dados_bancarios.xlsx", "Excel (*.xlsx)"
        )
        if not path_saida:
            QMessageBox.warning(window, "Cancelado", "Arquivo não salvo.")
            return

        try:
            exportar_xlsx(df, path_saida, PATH_TEMPLATE)
            path_log = salvar_logs(logs, path_saida) if logs else None
            detalhe = f"Planilha salva em:\n{path_saida}"
            if path_log:
                detalhe += f"\n\nLog de exceções salvo em:\n{path_log}"
            QMessageBox.information(window, "Concluído", msg + "\n\n" + detalhe)
            if abrir_checkbox.isChecked():
                try:
                    os.startfile(path_saida)
                except Exception:
                    subprocess.call(["open", path_saida])
        except Exception as e:
            QMessageBox.critical(window, "Erro ao salvar", str(e))

    def ao_erro(msg):
        btn_executar.setEnabled(True)
        status_label.setText("")
        QMessageBox.critical(window, "Erro no processamento", msg)

    def executar():
        btn_executar.setEnabled(False)
        status_label.setText("Processando…")
        worker = WorkerThread(state["alunos"], state["bancarios"], PATH_TEMPLATE)
        worker.concluido.connect(ao_concluir)
        worker.erro.connect(ao_erro)
        state["worker"] = worker
        worker.start()

    btn_executar.clicked.connect(executar)

    # --- Layout do card ---
    card_layout = QVBoxLayout(card)
    card_layout.setContentsMargins(28, 28, 28, 28)
    card_layout.setSpacing(14)

    titulo = QLabel("BolsistaDB")
    titulo.setAlignment(Qt.AlignmentFlag.AlignCenter)
    titulo.setStyleSheet("""
        font-family: "Bahnschrift Condensed";
        font-size: 38px;
        font-weight: bold;
        color: white;
        padding-bottom: 2px;
        background: transparent;
    """)
    card_layout.addWidget(titulo)

    # subtitulo = QLabel("Automação de tratamento de dados bancários — HBR")
    # subtitulo.setAlignment(Qt.AlignmentFlag.AlignCenter)
    # subtitulo.setStyleSheet("color: #888888; font-size: 11px; background: transparent; margin-bottom: 6px;")
    # card_layout.addWidget(subtitulo)

    card_layout.addWidget(criar_bloco("Lista de Bolsistas", sel_alunos, rem_alunos, label_alunos))
    card_layout.addWidget(criar_bloco("Planilha de Dados Bancários", sel_bancarios, rem_bancarios, label_bancarios))

    abrir_checkbox = ToggleSwitch("Abrir planilha quando estiver pronta")
    card_layout.addWidget(abrir_checkbox)

    exec_row = QHBoxLayout()
    exec_row.addStretch()
    exec_row.addWidget(btn_executar)
    exec_row.addStretch()
    card_layout.addLayout(exec_row)
    card_layout.addWidget(status_label)

    # --- Layout principal ---
    main_layout = QVBoxLayout(container)
    main_layout.addStretch()
    main_layout.addWidget(card, alignment=Qt.AlignmentFlag.AlignCenter)
    main_layout.addStretch()
    main_layout.setContentsMargins(30, 30, 30, 30)

    # --- Assinatura ---
    github_link = QLabel()
    github_link.setText(
        '<a href="https://github.com/imbaTIMvel/bolsistadb">'
        'BolsistaDB v0.1.0 - GitHub'
        '</a>'
        )
    github_link.setAlignment(Qt.AlignmentFlag.AlignCenter)
    github_link.setOpenExternalLinks(True)
    github_link.setStyleSheet("""
        QLabel {
            background-color: transparent;
            color: rgba(255,255,255,120);
            font-size: 11px;
        }
        QLabel:hover {
            color: #f9b02e;
        }
        """)
    footer = QLabel("Desenvolvido por: Diretoria Administrativa Financeira — DAF")
    footer.setAlignment(Qt.AlignmentFlag.AlignCenter)
    footer.setStyleSheet("""
        QLabel {
            background-color: transparent;
            color: rgba(255,255,255,120);
            font-size: 10px;
            padding-bottom: 4px;
        }
    """)
    main_layout.addWidget(github_link)
    main_layout.addWidget(footer)

    window_layout = QVBoxLayout(window)
    window_layout.setContentsMargins(0, 0, 0, 0)
    window_layout.addWidget(scroll)

    def resize_event(event):
        bg_label.resize(window.size())
        scroll.resize(window.size())
        card.setMaximumWidth(620)

    window.resizeEvent = resize_event

    window.show()
    sys.exit(app.exec())


# ==============================
# ENTRY POINT
# ==============================

if __name__ == "__main__":
    criar_janela()