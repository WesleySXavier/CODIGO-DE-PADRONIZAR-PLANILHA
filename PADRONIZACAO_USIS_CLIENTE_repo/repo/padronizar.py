#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Padroniza qualquer planilha "não tratada" (INPUT) e devolve no MODELO do TEMPLATE.

Uso:
    python padronizar.py --template TEMPLATE.xlsx --input base_nao_tratada.xlsx --output saida.xlsx

Se --input não for informado, o próprio TEMPLATE é tratado como base de dados
(útil para testar/validar as regras).

O que ele faz (todas as letras se referem às colunas do TEMPLATE):
    - CEP (Y)       -> #####-###
    - CPF (AP)      -> ###.###.###-##
    - RG  (AQ)      -> limpa; zera quando vier tudo 0; formata XX.XXX.XXX-X quando tiver 9 dígitos
    - Telefone (AT) -> (##) #####-#### (11 díg.) ou (##) ####-#### (10 díg.)
    - Cidade (T)    -> TUDO MAIÚSCULO
    - E-mail (AU)   -> minúsculo; apaga quando não for e-mail válido
    - Complemento (W) -> apaga "CASA" sozinho, zeros, pontos e valores puramente numéricos;
                         mantém "CASA 1", "CASA 2", "FUNDOS", "APTO 2", etc.
    - Código de cadastro interno (S) <- copiado de codigo_unico_cliente (CW); CW é preservada.
    - A saída mantém EXATAMENTE os cabeçalhos e a ordem das colunas do TEMPLATE.
"""

import argparse
import re
import unicodedata

import openpyxl
from openpyxl.utils import column_index_from_string, get_column_letter

# ----------------------------------------------------------------------------
# CONFIG — ajuste aqui se precisar incluir mais colunas no futuro.
# As chaves são letras de coluna NO TEMPLATE.
# ----------------------------------------------------------------------------
COL_CIDADE       = "T"
COL_NUMERO       = "U"    # Número do imóvel
COL_COD_INTERNO  = "S"    # destino (vazio no template)
COL_COMPLEMENTO  = "W"
COL_CEP          = "Y"
COL_CPF          = "AP"
COL_RG           = "AQ"
COL_TELEFONE     = "AT"
COL_EMAIL        = "AU"
COL_COD_CLIENTE  = "CW"   # origem do código único (codigo_unico_cliente)

# Prefixo do código (parte antes do "_") -> município correto.
# Se a cidade da linha não bater com o prefixo, a CIDADE é corrigida (o código NUNCA muda).
PREFIXO_MUNICIPIO = {
    "APIAI": "APIAÍ",
    "SCAJA": "CAJATI",
    "UCANA": "CANANÉIA",
    "SELDO": "ELDORADO",
    "BIPOR": "IPORANGA",
    "UITAP": "ITAPETININGA",
    "UJACU": "JACUPIRANGA",
    "UREGT": "REGISTRO",
    "USETB": "SETE BARRAS",
}

HEADER_ROW = 1            # linha dos cabeçalhos no template
FIRST_DATA_ROW = 2

# Valores em COMPLEMENTO que significam "sem complemento" e devem ser apagados.
COMPLEMENTO_SEM_INFO = {
    "CASA", "NAO POSSUI", "NÃO POSSUI", "SEM COMPLEMENTO", "SEM", "NENHUM",
    "NENHUMA", "N/A", "NA", "NAO TEM", "NÃO TEM", "SEM INFORMACAO",
    "SEM INFORMAÇÃO", "-", "--",
}

EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


# ----------------------------------------------------------------------------
# Helpers de normalização
# ----------------------------------------------------------------------------
def _so_digitos(valor):
    return re.sub(r"\D", "", str(valor)) if valor is not None else ""


def _norm_header(texto):
    """Normaliza cabeçalho para casar colunas entre input e template."""
    if texto is None:
        return ""
    t = unicodedata.normalize("NFKD", str(texto))
    t = t.encode("ascii", "ignore").decode("ascii")
    t = t.strip().lower().rstrip(":").strip()
    t = re.sub(r"\s+", " ", t)
    return t


def _vazio(valor):
    return valor is None or str(valor).strip() == ""


# ----------------------------------------------------------------------------
# Regras de padronização por campo
# ----------------------------------------------------------------------------
def fmt_cep(valor):
    if _vazio(valor):
        return None
    d = _so_digitos(valor)
    if d == "" or d.strip("0") == "":   # vazio ou só zeros
        return None
    if len(d) == 8:
        return f"{d[:5]}-{d[5:]}"
    return d  # tamanho inesperado: devolve só os dígitos


def fmt_cpf(valor):
    if _vazio(valor):
        return None
    d = _so_digitos(valor)
    if d == "" or d.strip("0") == "":
        return None
    if len(d) == 11:
        return f"{d[:3]}.{d[3:6]}.{d[6:9]}-{d[9:]}"
    return d


def fmt_rg(valor):
    if _vazio(valor):
        return None
    # RG pode ter letra/X; mantém alfanumérico
    bruto = re.sub(r"[^0-9A-Za-z]", "", str(valor)).upper()
    if bruto == "" or re.fullmatch(r"0+", bruto):   # vazio ou só zeros -> apaga
        return None
    # CIN (nova identidade): o "RG" é o próprio CPF (11 dígitos) -> formata como CPF
    if len(bruto) == 11 and bruto.isdigit():
        return f"{bruto[:3]}.{bruto[3:6]}.{bruto[6:9]}-{bruto[9:]}"
    if len(bruto) == 9 and bruto[:-1].isdigit():    # padrão SP XX.XXX.XXX-X
        return f"{bruto[:2]}.{bruto[2:5]}.{bruto[5:8]}-{bruto[8:]}"
    return bruto


def fmt_telefone(valor):
    if _vazio(valor):
        return None
    d = _so_digitos(valor)
    if d.startswith("55") and len(d) in (12, 13):   # tira DDI Brasil
        d = d[2:]
    if d == "" or d.strip("0") == "":
        return None
    if len(d) == 11:
        return f"({d[:2]}) {d[2:7]}-{d[7:]}"
    if len(d) == 10:
        return f"({d[:2]}) {d[2:6]}-{d[6:]}"
    return d


def fmt_numero(valor, codigo=None):
    """
    Número do imóvel:
      - Se estiver vazio, usa os dígitos do código interno (ex.: SCAJA_0869 -> 0869).
      - Tira os zeros à esquerda de TODO número (0086 -> 86, 0324 -> 324).
      - Se sobrar só zero (ou vazio), apaga.
      - Valor não numérico (ex.: 'S/N') é mantido como está.
    """
    txt = "" if _vazio(valor) else str(valor).strip()

    # "Sem número": vazio OU só zeros (0, 00, 000...).
    sem_numero = (txt == "") or (re.fullmatch(r"\d+", txt) and txt.lstrip("0") == "")

    # Sem número -> pega os numerais do código interno (parte depois do último '_').
    if sem_numero and not _vazio(codigo):
        seg = str(codigo).split("_")[-1]
        txt = re.sub(r"\D", "", seg) or re.sub(r"\D", "", str(codigo))
    elif sem_numero:
        return None   # sem número e sem código -> apaga

    # Número puro -> remove zeros à esquerda.
    if re.fullmatch(r"\d+", txt):
        sem_zeros = txt.lstrip("0")
        return sem_zeros or None

    return valor   # não numérico: mantém


def fmt_cidade(valor):
    if _vazio(valor):
        return valor
    return str(valor).strip().upper()


def municipio_por_codigo(cod):
    """Retorna o município correto a partir do prefixo do código (antes do '_'), ou None."""
    if _vazio(cod):
        return None
    prefixo = str(cod).strip().split("_")[0].upper()
    return PREFIXO_MUNICIPIO.get(prefixo)


def fmt_email(valor):
    if _vazio(valor):
        return None
    e = str(valor).strip().lower()
    return e if EMAIL_RE.match(e) else None   # inválido -> apaga


def fmt_complemento(valor):
    """Mantém só complemento 'em extenso' de verdade (ex.: CASA 1, FUNDOS, APTO 2)."""
    if _vazio(valor):
        return None
    txt = str(valor).replace(".", " ")          # remove pontos
    txt = re.sub(r"\s+", " ", txt).strip()
    if txt == "":
        return None
    chave = _unaccent_upper(txt)
    if chave in {_unaccent_upper(x) for x in COMPLEMENTO_SEM_INFO}:
        return None
    if re.fullmatch(r"[0-9]+", chave):           # puro número (inclui zeros) -> apaga
        return None
    return txt                                    # mantém (CASA 1, CASA 2, FUNDOS...)


def _unaccent_upper(texto):
    t = unicodedata.normalize("NFKD", str(texto))
    t = t.encode("ascii", "ignore").decode("ascii")
    return t.strip().upper()


# ----------------------------------------------------------------------------
# Núcleo: alinhar input ao template e aplicar regras
# ----------------------------------------------------------------------------
def _headers_ordenados(ws):
    """Lista [(indice_coluna, header_normalizado), ...] na ordem da planilha."""
    return [(c, _norm_header(ws.cell(HEADER_ROW, c).value))
            for c in range(1, ws.max_column + 1)]


def mapear_colunas(tpl_headers, in_headers):
    """
    Casa coluna do input -> coluna do template por nome, respeitando a ORDEM.
    Assim cabeçalhos repetidos (ex.: dois 'cidade') casam 1 a 1 na sequência,
    e colunas que só existem no template (ex.: 'Código de cadastro interno')
    ficam sem origem (serão preenchidas por regra).
    Retorna {indice_template: indice_input}.
    """
    usados = set()
    mapa = {}
    for tci, th in tpl_headers:
        if not th:
            continue
        for ici, ih in in_headers:
            if ici in usados or ih != th:
                continue
            mapa[tci] = ici
            usados.add(ici)
            break
    return mapa


def _linhas_com_dados(ws, cols):
    linhas = []
    for r in range(FIRST_DATA_ROW, ws.max_row + 1):
        if all(_vazio(ws.cell(r, c).value) for c in cols):
            continue
        linhas.append(r)
    return linhas


def _idx_colunas():
    return {k: column_index_from_string(v) for k, v in {
        "cidade": COL_CIDADE, "numero": COL_NUMERO, "cod_interno": COL_COD_INTERNO,
        "complemento": COL_COMPLEMENTO, "cep": COL_CEP, "cpf": COL_CPF,
        "rg": COL_RG, "tel": COL_TELEFONE, "email": COL_EMAIL,
        "cod_cliente": COL_COD_CLIENTE,
    }.items()}


def _escrever_e_tratar(ws_out, r, ws_in, r_in, mapa, idx):
    """Copia uma linha do input para a linha r da saída e aplica todas as regras."""
    # Copia os dados do input para a coluna correspondente do template.
    for tci, ici in mapa.items():
        ws_out.cell(r, tci, ws_in.cell(r_in, ici).value)

    # Código de cadastro interno (S) <- codigo_unico_cliente (CW). CW é mantida.
    cod = ws_out.cell(r, idx["cod_cliente"]).value
    if not _vazio(cod):
        ws_out.cell(r, idx["cod_interno"]).value = cod

    # Padronizações. (usa .value = para que None realmente APAGUE a célula)
    # Cidade: se o prefixo do código indicar um município, ele MANDA (corrige a cidade);
    # senão, apenas deixa em maiúsculo.
    muni = municipio_por_codigo(ws_out.cell(r, idx["cod_interno"]).value)
    ws_out.cell(r, idx["cidade"]).value      = muni or fmt_cidade(ws_out.cell(r, idx["cidade"]).value)
    ws_out.cell(r, idx["numero"]).value      = fmt_numero(ws_out.cell(r, idx["numero"]).value,
                                                           ws_out.cell(r, idx["cod_interno"]).value)
    ws_out.cell(r, idx["cep"]).value         = fmt_cep(ws_out.cell(r, idx["cep"]).value)
    ws_out.cell(r, idx["cpf"]).value         = fmt_cpf(ws_out.cell(r, idx["cpf"]).value)
    ws_out.cell(r, idx["rg"]).value          = fmt_rg(ws_out.cell(r, idx["rg"]).value)
    ws_out.cell(r, idx["tel"]).value         = fmt_telefone(ws_out.cell(r, idx["tel"]).value)
    ws_out.cell(r, idx["email"]).value       = fmt_email(ws_out.cell(r, idx["email"]).value)
    ws_out.cell(r, idx["complemento"]).value = fmt_complemento(ws_out.cell(r, idx["complemento"]).value)


def processar_varios(template_path, input_paths, output_path):
    """
    Unifica VÁRIAS planilhas cruas em um único arquivo no modelo do template,
    empilhando os registros e aplicando o tratamento em toda a base.
    """
    # 1) Template define a estrutura (cabeçalhos + ordem) e é a base da saída.
    wb_out = openpyxl.load_workbook(template_path)
    ws_out = wb_out.active
    tpl_headers = _headers_ordenados(ws_out)
    idx = _idx_colunas()

    # 2) Limpa as linhas de exemplo do template na saída.
    if ws_out.max_row >= FIRST_DATA_ROW:
        ws_out.delete_rows(FIRST_DATA_ROW, ws_out.max_row - HEADER_ROW)

    # 3) Para cada input, casa colunas e empilha os registros tratados.
    r = FIRST_DATA_ROW
    total_por_arquivo = []
    for caminho in input_paths:
        ws_in = openpyxl.load_workbook(caminho, data_only=True).active
        in_headers = _headers_ordenados(ws_in)
        mapa = mapear_colunas(tpl_headers, in_headers)
        linhas_in = _linhas_com_dados(ws_in, [c for c, _ in in_headers])
        for r_in in linhas_in:
            _escrever_e_tratar(ws_out, r, ws_in, r_in, mapa, idx)
            r += 1
        total_por_arquivo.append((caminho, len(linhas_in)))

    wb_out.save(output_path)
    return total_por_arquivo


def processar(template_path, input_path, output_path):
    """Trata uma única planilha (atalho para processar_varios com um input)."""
    return sum(n for _, n in processar_varios(template_path, [input_path], output_path))


import os
import glob

# Pastas padrão (ficam ao lado do script).
BASE_DIR     = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR    = os.path.join(BASE_DIR, "input")
OUTPUT_DIR   = os.path.join(BASE_DIR, "output")
TEMPLATE_PADRAO = os.path.join(BASE_DIR, "TEMPLATE.xlsx")
TEMPLATE_BASE_FINAL = os.path.join(BASE_DIR, "TEMPLATE_BASE_FINAL.xlsx")


def _planilhas(pasta):
    arqs = []
    for ext in ("*.xlsx", "*.xlsm"):
        arqs += glob.glob(os.path.join(pasta, ext))
    # ignora arquivos temporários do Excel (~$...)
    return sorted(a for a in arqs if not os.path.basename(a).startswith("~$"))


def _gerar_base_final(tratado_path):
    """Chama o base_final.py para produzir a base de ajuste final a partir do tratado."""
    if not os.path.exists(TEMPLATE_BASE_FINAL):
        print("(aviso) TEMPLATE_BASE_FINAL.xlsx não encontrado ao lado do script; "
              "pulei a base final.")
        return
    import sys
    if BASE_DIR not in sys.path:
        sys.path.insert(0, BASE_DIR)
    try:
        import base_final
    except ImportError:
        print("(aviso) base_final.py não encontrado ao lado do script; pulei a base final.")
        return
    saida_bf = os.path.join(OUTPUT_DIR, "BASE_FINAL.xlsx")
    cont, sobra = base_final.gerar(TEMPLATE_BASE_FINAL, tratado_path, saida_bf)
    print(f"OK: base final -> output/{os.path.basename(saida_bf)}")
    for aba, n in cont.items():
        print(f"   - {aba}: {n} registros")
    if sobra:
        print(f"   *** ATENÇÃO: {sobra} registro(s) sem aba de município "
              f"foram para a aba 'NAO_MAPEADO' (verificar).")


def rodar_por_pastas(template_path, unificar=True):
    """
    Lê TODA planilha da pasta input/.
    unificar=True  -> junta tudo em um único arquivo em output/ (padrão).
    unificar=False -> gera um arquivo tratado por planilha.
    """
    os.makedirs(INPUT_DIR, exist_ok=True)
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    if not os.path.exists(template_path):
        print(f"ERRO: template não encontrado em {template_path}")
        return

    arquivos = _planilhas(INPUT_DIR)
    if not arquivos:
        print(f"Nenhuma planilha encontrada em: {INPUT_DIR}")
        print("Coloque os arquivos crus (.xlsx) dentro da pasta 'input' e rode de novo.")
        return

    if unificar:
        saida = os.path.join(OUTPUT_DIR, "BASE_UNIFICADA_TRATADA.xlsx")
        try:
            detalhes = processar_varios(template_path, arquivos, saida)
        except Exception as e:
            print(f"ERRO na unificação: {e}")
            return
        total = sum(n for _, n in detalhes)
        print(f"Unificadas {len(arquivos)} planilha(s), {total} registros no total:")
        for caminho, n in detalhes:
            print(f"   - {os.path.basename(caminho)}: {n} registros")
        print(f"OK: tratado -> output/{os.path.basename(saida)}")

        # 2ª etapa: gera a BASE FINAL a partir do tratado (se o template existir).
        _gerar_base_final(saida)
    else:
        for caminho in arquivos:
            nome = os.path.splitext(os.path.basename(caminho))[0]
            saida = os.path.join(OUTPUT_DIR, f"{nome}_TRATADO.xlsx")
            try:
                n = processar(template_path, caminho, saida)
                print(f"OK: {os.path.basename(caminho)} -> output/{os.path.basename(saida)} ({n} registros)")
            except Exception as e:
                print(f"ERRO ao processar {os.path.basename(caminho)}: {e}")


def main():
    p = argparse.ArgumentParser(
        description="Padroniza planilha no modelo do template. "
                    "Sem argumentos: lê tudo de ./input e salva em ./output.")
    p.add_argument("--template", help="Arquivo modelo (.xlsx). Padrão: ./TEMPLATE.xlsx")
    p.add_argument("--input", help="Base crua específica (.xlsx). Se omitido, usa a pasta ./input")
    p.add_argument("--output", help="Arquivo de saída específico (.xlsx)")
    p.add_argument("--separado", action="store_true",
                   help="Gera um arquivo tratado por planilha, em vez de unificar tudo.")
    args = p.parse_args()

    template_path = args.template or TEMPLATE_PADRAO

    # Modo arquivo único (quando --input e --output são informados).
    if args.input and args.output:
        n = processar(template_path, args.input, args.output)
        print(f"OK: {n} registros padronizados -> {args.output}")
        return

    # Modo pastas (padrão): unifica tudo de ./input num só arquivo tratado.
    rodar_por_pastas(template_path, unificar=not args.separado)


if __name__ == "__main__":
    main()
