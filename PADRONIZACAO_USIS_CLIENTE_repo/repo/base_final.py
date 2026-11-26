#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Gera a BASE FINAL (de ajuste) a partir da base já TRATADA pelo padronizar.py.

- Pega os dados tratados (layout do TEMPLATE do padronizar) e joga nas colunas
  respectivas do TEMPLATE_BASE_FINAL.xlsx.
- Separa os registros por MUNICÍPIO, uma aba para cada (as abas já existem no template).
- Aplica as regras da base final (status da entrega, status DEEP, tipo de adesão,
  fossa séptica, coordenadas com vírgula).

Uso isolado:
    python base_final.py --template TEMPLATE_BASE_FINAL.xlsx --tratado BASE_UNIFICADA_TRATADA.xlsx --output BASE_FINAL.xlsx
Normalmente é chamado automaticamente pelo padronizar.py.
"""

import argparse
import re
import unicodedata
from datetime import date

import openpyxl
from openpyxl.utils import column_index_from_string as C

# ----------------------------------------------------------------------------
# MAPA: coluna da BASE FINAL  <-  coluna do TRATADO (layout do template padronizar)
# (só as colunas com origem clara; as demais ficam em branco)
# ----------------------------------------------------------------------------
MAPA_BF_PARA_TRATADO = {
    "A":  "S",    # CÓDIGO DE CADASTRO INTERNO  <- Código de cadastro interno
    "B":  "T",    # MUNICÍPIO                   <- Cidade (já corrigida pelo código)
    "C":  "X",    # BAIRRO                      <- Bairro:
    "D":  "V",    # RUA                         <- Logradouro
    "E":  "U",    # NÚMERO DO IMÓVEL            <- Número
    "F":  "W",    # COMPLEMENTO                 <- Complemento:
    "G":  "J",    # LATITUDE                    <- Latitude
    "H":  "K",    # LONGITUDE                   <- Longitude
    "I":  "AC",   # STATUS DO CADASTRO          <- Status do cadastro
    "J":  "AA",   # QUAL TIPO DE ADESÃO         <- Qual tipo de adesão?
    "N":  "AB",   # FOTO DA FACHADA             <- Foto da fachada do imóvel
    "Q":  "Y",    # CEP                         <- CEP
    "V":  "AN",   # NOME COMPLETO DO MORADOR    <- Nome completo do morador
    "W":  "AO",   # DATA DE NASCIMENTO          <- Data de nascimento
    "X":  "AP",   # NÚMERO CPF                  <- Numero do CPF
    "Y":  "AQ",   # NÚMERO DO RG                <- Número do RG
    "AB": "AT",   # NÚMERO DE TELEFONE          <- Número de telefone
    "AC": "AU",   # E-MAIL                      <- E-mail
    "AD": "AV",   # OS DADOS ACIMA SÃO          <- Os dados acima são:
    "AE": "BP",   # POSSUI DOCUMENTAÇÃO IMÓVEL  <- Possui a documentação do imóvel...
    "AF": "BQ",   # QUAL DOCUMENTO              <- Qual?
    "AG": "BR",   # FOTO DO DOCUMENTO IMÓVEL    <- Foto do documento do imóvel
    "AH": "BS",   # TEMPO DE MORADIA            <- Tempo de moradia no local
    "AO": "BV",   # FOTO DO HIDRÔMETRO          <- Foto do hidrômetro
    "AP": "BW",   # NÚMERO DO HIDRÔMETRO        <- Número do hidrômetro
    "AQ": "BX",   # PDE OU RGI                  <- PDE ou RGI
    "AR": "BY",   # FOTO DA CONTA DE ÁGUA       <- Foto da conta de água
    "BG": "CB",   # SITUAÇÃO DE ESGOTAMENTO     <- Solução de esgotamento atual
    "BT": "G",    # DATA DO CADASTRO            <- Data do registro
    "BW": "AD",   # OBSERVAÇÃO                  <- Observações
}

# Colunas de regras na BASE FINAL
BF_MUNICIPIO   = "B"
BF_LAT         = "G"
BF_LON         = "H"
BF_STATUS_CAD  = "I"
BF_TIPO_ADESAO = "J"
BF_ESGOTO      = "BG"
BF_STATUS_ENTREGA = "BX"
BF_STATUS_DEEP    = "BY"

# Origem (no TRATADO) da situação de backoffice p/ o STATUS DEEP
TRAT_SIT_BACKOFFICE = "M"   # 'Situação Backoffice - Visita'

HEADER_ROW = 1
FIRST_DATA_ROW = 2


def _na(v):
    return v is None or str(v).strip() == ""


def _norm(s):
    if s is None:
        return ""
    t = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"\s+", " ", t.strip().upper())


# ----------------------------------------------------------------------------
# Regras da base final
# ----------------------------------------------------------------------------
def coord_virgula(v):
    if _na(v):
        return v
    return str(v).replace(".", ",")


def tipo_adesao(v):
    """Água -> ICA | Esgoto -> ICE | Água e esgoto -> ICA E ICE."""
    if _na(v):
        return v
    n = _norm(v)
    tem_agua = "AGUA" in n
    tem_esg = "ESGOTO" in n
    if tem_agua and tem_esg:
        return "ICA E ICE"
    if tem_agua:
        return "ICA"
    if tem_esg:
        return "ICE"
    return v


def status_cadastro(valor_status, valor_esgoto):
    """Se for 'Cadastro realizado', troca pela regra da fossa séptica (usa a situação de esgoto)."""
    if _na(valor_status):
        return valor_status
    if _norm(valor_status) != "CADASTRO REALIZADO":
        return valor_status
    esg = _norm(valor_esgoto)
    if "FOSSA SEPTICA" in esg or "TANQUE SEPTICO" in esg:
        return "possui fossa séptica"
    return "não possui fossa séptica"


def status_deep(valor_backoffice):
    """Aprovado/Reprovado -> jip liberado | Aguardando ajustes/Pendente -> deep em conferência."""
    n = _norm(valor_backoffice)
    if n in ("APROVADO", "REPROVADO"):
        return "Liberado"
    if "AGUARDANDO" in n or "PENDENTE" in n:
        return "Em conferência"
    return None


# ----------------------------------------------------------------------------
# Geração
# ----------------------------------------------------------------------------
def gerar(bf_template, tratado_path, saida_path):
    wb = openpyxl.load_workbook(bf_template)
    # índice de abas por município normalizado
    abas = {_norm(ws.title): ws for ws in wb.worksheets}

    # aba de segurança p/ municípios sem aba correspondente (não perder dados)
    nome_sobra = "NAO_MAPEADO"
    ws_sobra = None

    # limpa linhas de exemplo de todas as abas
    for ws in wb.worksheets:
        if ws.max_row >= FIRST_DATA_ROW:
            ws.delete_rows(FIRST_DATA_ROW, ws.max_row - HEADER_ROW)

    # dados tratados
    ws_t = openpyxl.load_workbook(tratado_path, data_only=True).active

    mapa = {C(bf): C(tr) for bf, tr in MAPA_BF_PARA_TRATADO.items()}
    i_mun   = C(BF_MUNICIPIO)
    i_lat   = C(BF_LAT); i_lon = C(BF_LON)
    i_stat  = C(BF_STATUS_CAD); i_ades = C(BF_TIPO_ADESAO); i_esg = C(BF_ESGOTO)
    i_ent   = C(BF_STATUS_ENTREGA); i_deep = C(BF_STATUS_DEEP)
    i_sitbo = C(TRAT_SIT_BACKOFFICE)

    hoje = date.today().strftime("%d/%m/%Y")
    texto_entrega = f"incluído na base {hoje}"

    prox_linha = {}   # título aba -> próxima linha livre
    contagem = {}
    sobra = 0

    for r in range(FIRST_DATA_ROW, ws_t.max_row + 1):
        if all(_na(ws_t.cell(r, c).value) for c in range(1, ws_t.max_column + 1)):
            continue
        municipio = ws_t.cell(r, C(MAPA_BF_PARA_TRATADO[BF_MUNICIPIO])).value
        ws_dest = abas.get(_norm(municipio))
        if ws_dest is None:
            if ws_sobra is None:
                ws_sobra = wb.create_sheet(nome_sobra)
                # copia cabeçalho da primeira aba
                base = wb.worksheets[0]
                for c in range(1, base.max_column + 1):
                    ws_sobra.cell(1, c, base.cell(1, c).value)
            ws_dest = ws_sobra
            sobra += 1

        rr = prox_linha.get(ws_dest.title, FIRST_DATA_ROW)
        prox_linha[ws_dest.title] = rr + 1
        contagem[ws_dest.title] = contagem.get(ws_dest.title, 0) + 1

        # copia colunas mapeadas
        for bf_c, tr_c in mapa.items():
            ws_dest.cell(rr, bf_c, ws_t.cell(r, tr_c).value)

        # coordenadas com vírgula
        ws_dest.cell(rr, i_lat).value = coord_virgula(ws_dest.cell(rr, i_lat).value)
        ws_dest.cell(rr, i_lon).value = coord_virgula(ws_dest.cell(rr, i_lon).value)

        # tipo de adesão -> ICA / ICE
        ws_dest.cell(rr, i_ades).value = tipo_adesao(ws_dest.cell(rr, i_ades).value)

        # status do cadastro -> regra da fossa séptica
        ws_dest.cell(rr, i_stat).value = status_cadastro(
            ws_dest.cell(rr, i_stat).value, ws_dest.cell(rr, i_esg).value)

        # status da entrega -> "INCLUÍDO NA BASE <data de hoje>"
        ws_dest.cell(rr, i_ent).value = texto_entrega

        # status DEEP -> a partir da situação de backoffice do tratado
        ws_dest.cell(rr, i_deep).value = status_deep(ws_t.cell(r, i_sitbo).value)

    wb.save(saida_path)
    return contagem, sobra


import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))


def main():
    p = argparse.ArgumentParser(description="Gera a base final a partir da base tratada.")
    p.add_argument("--template", default=os.path.join(BASE_DIR, "TEMPLATE_BASE_FINAL.xlsx"))
    p.add_argument("--tratado", default=os.path.join(BASE_DIR, "output", "BASE_UNIFICADA_TRATADA.xlsx"))
    p.add_argument("--output", default=os.path.join(BASE_DIR, "output", "BASE_FINAL.xlsx"))
    a = p.parse_args()

    if not os.path.exists(a.tratado):
        print("Não encontrei o arquivo tratado em:", a.tratado)
        print(">> Rode primeiro o 'padronizar.py' (ele gera o tratado e já chama o base final).")
        return

    cont, sobra = gerar(a.template, a.tratado, a.output)
    print("Base final gerada:", a.output)
    for aba, n in cont.items():
        print(f"   {aba}: {n}")
    if sobra:
        print(f"   *** {sobra} registro(s) sem aba de município -> aba NAO_MAPEADO")


if __name__ == "__main__":
    main()
