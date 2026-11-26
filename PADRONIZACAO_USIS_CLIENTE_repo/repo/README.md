# Padronização USIS Cliente

Automação para **padronizar** bases cruas de cadastro e gerar a **base final de ajuste**
(separada por município), a partir de modelos (`TEMPLATE.xlsx` e `TEMPLATE_BASE_FINAL.xlsx`).

Numa única execução, o projeto:

1. Junta todas as planilhas cruas da pasta `input/` em uma só, no modelo do template.
2. Aplica o tratamento/padronização.
3. Gera a base final de ajuste, separada por município (uma aba por município).
4. Entrega **dois arquivos** em `output/`.

## Estrutura

```
padronizar.py            # programa principal (roda tudo)
base_final.py            # gera a base final (chamado automaticamente pelo padronizar)
TEMPLATE.xlsx            # modelo do tratamento (somente cabeçalhos neste repositório)
TEMPLATE_BASE_FINAL.xlsx # modelo da base final, com as abas por município
input/                   # coloque aqui as planilhas BRUTAS (.xlsx) — não versionadas
output/                  # resultados gerados — não versionados
requirements.txt
```

> **Importante (LGPD):** o `input/` e o `output/` estão no `.gitignore`.
> Nunca faça commit de planilhas com dados reais (CPF, nomes, e-mails).
> O `TEMPLATE.xlsx` deste repositório contém **apenas os cabeçalhos**, sem dados.

## Como usar

```bash
# 1) instalar dependência (uma vez)
pip install -r requirements.txt

# 2) colocar as planilhas brutas na pasta input/

# 3) rodar
python padronizar.py
```

Saem dois arquivos em `output/`:

- `BASE_UNIFICADA_TRATADA.xlsx` — base tratada/padronizada.
- `BASE_FINAL.xlsx` — base final de ajuste, separada por município em abas.

### Opções

```bash
python padronizar.py --separado        # um arquivo tratado por planilha (não unifica)
python padronizar.py --input a.xlsx --output s.xlsx   # tratar um arquivo específico
```

## O que o tratamento faz (padronizar.py)

- **Cidade**: tudo em maiúsculo.
- **Número do imóvel**: remove zeros à esquerda; se estiver vazio ou só zeros, usa os dígitos do código interno.
- **Complemento**: remove `casa` sozinho, pontos, zeros e valores puramente numéricos; mantém `casa 1`, `fundos`, etc.
- **CEP**: `#####-###`.
- **CPF**: `###.###.###-##`.
- **RG**: formata; se tiver 11 dígitos (CPF/CIN), formata como CPF.
- **Telefone**: `(##) #####-####` ou `(##) ####-####`.
- **E-mail**: minúsculo; apaga inválidos.
- **Código de cadastro interno**: copiado do `codigo_unico_cliente`.
- **Município pelo código**: o prefixo do código corrige o município quando divergente.

## O que a base final faz (base_final.py)

- Leva os dados tratados para as colunas correspondentes da base final.
- Separa por município (uma aba por município).
- Coordenadas (lat/long) com **vírgula** no lugar do ponto.
- **Tipo de adesão**: Água→`ICA`, Esgoto→`ICE`, Água e esgoto→`ICA E ICE`.
- **Status do cadastro**: onde for "Cadastro realizado", usa a situação de esgotamento —
  fossa/tanque séptico → `possui fossa séptica`, senão → `não possui fossa séptica`.
- **Status da entrega**: `incluído na base <data de hoje>`.
- **Status DEEP**: Aprovado/Reprovado → `Liberado`; Aguardando ajustes/Pendente → `Em conferência`.
- Registros de município sem aba vão para a aba `NAO_MAPEADO`.

## Requisitos

- Python 3.9+
- openpyxl (ver `requirements.txt`)
