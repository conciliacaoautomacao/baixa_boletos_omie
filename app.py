import streamlit as st
import pandas as pd
import fitz
import re
import io
import os
from datetime import date, datetime, timedelta
from supabase import create_client
from openpyxl import load_workbook
from copy import copy

# =============================
# CONFIG
# =============================
st.set_page_config(
    page_title="Baixa Boletos Omie",
    page_icon="📄",
    layout="wide"
)

SUPABASE_URL = st.secrets["SUPABASE_URL"]
SUPABASE_KEY = st.secrets["SUPABASE_KEY"]

supabase = create_client(SUPABASE_URL, SUPABASE_KEY)

MODELO_EXCEL = "modelo/Omie_Contas_Pagar_v1_1_5.xlsx"


# =============================
# FUNÇÕES
# =============================
def br_to_float(valor):
    if valor is None:
        return 0.0

    valor = str(valor).strip()
    valor = valor.replace("R$", "").strip()
    valor = valor.replace(".", "").replace(",", ".")

    try:
        return float(valor)
    except:
        return 0.0


def format_data_br(data):
    if pd.isna(data) or data is None:
        return ""
    if isinstance(data, str):
        return data
    return data.strftime("%d/%m/%Y")


def calcular_data_previsao(data_vencimento):
    if data_vencimento.weekday() == 5:  # sábado
        return data_vencimento + timedelta(days=2)

    if data_vencimento.weekday() == 6:  # domingo
        return data_vencimento + timedelta(days=1)

    return data_vencimento


def primeiro_dia_mes_atual():
    hoje = date.today()
    return date(hoje.year, hoje.month, 1)


def extrair_texto_pdf(arquivo_pdf):
    texto = ""

    with fitz.open(stream=arquivo_pdf.read(), filetype="pdf") as doc:
        for page in doc:
            texto += page.get_text("text") + "\n"

    return texto


def extrair_boleto(arquivo_pdf):
    nome_arquivo = arquivo_pdf.name
    texto = extrair_texto_pdf(arquivo_pdf)

    # Código de barras / linha digitável
    codigo_barras = ""
    match_codigo = re.search(
        r"(341\d{2}\.\d{5}\s+\d{5}\.\d{6}\s+\d{5}\.\d{6}\s+\d\s+\d{14})",
        texto
    )
    if match_codigo:
        codigo_barras = match_codigo.group(1).strip()

    # Datas
    datas = re.findall(r"\b\d{2}/\d{2}/\d{4}\b", texto)

    data_documento = None
    vencimento = None

    if len(datas) >= 2:
        vencimento = datetime.strptime(datas[0], "%d/%m/%Y").date()
        data_documento = datetime.strptime(datas[1], "%d/%m/%Y").date()

    # Valor do documento
    valor_documento = 0.0
    match_valor = re.search(r"R\$\s*([\d\.]+,\d{2})", texto)
    if match_valor:
        valor_documento = br_to_float(match_valor.group(1))
    else:
        valores = re.findall(r"\b\d{1,3}(?:\.\d{3})*,\d{2}\b", texto)
        if valores:
            valor_documento = br_to_float(valores[0])

    # Pagador
    pagador = ""
    
    match_pagador = re.search(
        r"Pagador:\s*([A-Za-zÀ-ÿ\s]+?)\s+CPF:",
        texto,
        re.IGNORECASE
    )
    
    if match_pagador:
        pagador = match_pagador.group(1).strip()
    else:
        # fallback: pega a linha anterior ao CPF
        linhas = [l.strip() for l in texto.splitlines() if l.strip()]
    
        for i, linha in enumerate(linhas):
            if "CPF:" in linha and i > 0:
                possivel_nome = linhas[i - 1].replace("Pagador:", "").strip()
    
                if possivel_nome and not possivel_nome.upper().startswith("ENDERE"):
                    pagador = possivel_nome
                    break    

    data_registro = primeiro_dia_mes_atual()
    data_previsao = calcular_data_previsao(vencimento) if vencimento else None
    data_pagamento = data_previsao

    observacoes = (
        "Pagamento referente ao acionamento do Sinistro do Seguro Prestamista - "
        + pagador
    )

    print("PAGADOR EXTRAÍDO:", pagador)

    return {
        "nome_arquivo": nome_arquivo,
        "data_documento": data_documento,
        "vencimento": vencimento,
        "valor_documento": valor_documento,
        "codigo_barras": codigo_barras,
        "pagador": pagador,
        "data_registro": data_registro,
        "data_previsao": data_previsao,
        "data_pagamento": data_pagamento,
        "observacoes": observacoes,
        "status": "extraido"
    }


def salvar_no_supabase(df):
    dados = []

    for _, row in df.iterrows():
        dados.append({
            "nome_arquivo": row["nome_arquivo"],
            "data_documento": row["data_documento"].isoformat() if row["data_documento"] else None,
            "vencimento": row["vencimento"].isoformat() if row["vencimento"] else None,
            "valor_documento": float(row["valor_documento"] or 0),
            "codigo_barras": row["codigo_barras"],
            "pagador": row["pagador"],
            "data_registro": row["data_registro"].isoformat() if row["data_registro"] else None,
            "data_previsao": row["data_previsao"].isoformat() if row["data_previsao"] else None,
            "data_pagamento": row["data_pagamento"].isoformat() if row["data_pagamento"] else None,
            "observacoes": row["observacoes"],
            "status": "salvo"
        })

    try:
        supabase.table("boletos_extraidos").insert(dados).execute()
        return True, "Dados salvos com sucesso."
    except Exception as e:
        return False, f"Erro ao salvar. Possível boleto duplicado pelo código de barras. Detalhe: {e}"

def normalizar_texto(txt):
    if txt is None:
        return ""

    return (
        str(txt)
        .strip()
        .replace("\n", " ")
        .replace("  ", " ")
        .upper()
    )


def localizar_colunas_por_cabecalho(ws, linha_cabecalho=5):
    mapa = {}

    for col in range(1, ws.max_column + 1):
        valor = ws.cell(linha_cabecalho, col).value

        if valor:
            mapa[normalizar_texto(valor)] = col

    return mapa


def escrever_por_coluna(ws, linha, mapa_colunas, nome_coluna, valor):
    chave = normalizar_texto(nome_coluna)

    if chave not in mapa_colunas:
        raise Exception(f"Coluna não encontrada na planilha: {nome_coluna}")

    ws.cell(linha, mapa_colunas[chave]).value = valor

def copiar_estilo_linha(ws, linha_modelo, linha_destino):
    ws.row_dimensions[linha_destino].height = ws.row_dimensions[linha_modelo].height

    for col in range(1, ws.max_column + 1):
        origem = ws.cell(linha_modelo, col)
        destino = ws.cell(linha_destino, col)

        if origem.has_style:
            destino._style = copy(origem._style)

        destino.font = copy(origem.font)
        destino.fill = copy(origem.fill)
        destino.border = copy(origem.border)
        destino.alignment = copy(origem.alignment)
        destino.number_format = origem.number_format
        destino.protection = copy(origem.protection)
        
def gerar_excel_omie(df):
    wb = load_workbook(MODELO_EXCEL)
    ws = wb["Omie_Contas_Pagar"]

    linha_cabecalho = 5
    linha_inicial = 6

    mapa_colunas = localizar_colunas_por_cabecalho(ws, linha_cabecalho)

    for idx, row in df.iterrows():
        linha = linha_inicial + idx
        copiar_estilo_linha(ws, linha_inicial, linha)

        escrever_por_coluna(
            ws, linha, mapa_colunas,
            "Fornecedor * (Razão Social, Nome Fantasia, CNPJ ou CPF)",
            "OPI GOOROO FUNDO DE INVESTIMENTO EM DIREITOS CREDITORIOS"
        )

        escrever_por_coluna(
            ws, linha, mapa_colunas,
            "Categoria *",
            "Adiantamento de Seguros - Money Plus"
        )

        escrever_por_coluna(
            ws, linha, mapa_colunas,
            "Conta Corrente *",
            "Santander"
        )

        escrever_por_coluna(
            ws, linha, mapa_colunas,
            "Data de Emissão",
            row["data_documento"]
        )

        escrever_por_coluna(
            ws, linha, mapa_colunas,
            "Data de Registro *",
            row["data_registro"]
        )

        escrever_por_coluna(
            ws, linha, mapa_colunas,
            "Data de Vencimento *",
            row["vencimento"]
        )

        escrever_por_coluna(
            ws, linha, mapa_colunas,
            "Data de Previsão",
            row["data_previsao"]
        )

        escrever_por_coluna(
            ws, linha, mapa_colunas,
            "Data do Pagamento",
            row["data_pagamento"]
        )

        escrever_por_coluna(
            ws, linha, mapa_colunas,
            "Valor do Pagamento",
            float(row["valor_documento"] or 0)
        )

        escrever_por_coluna(
            ws, linha, mapa_colunas,
            "Observações",
            row["observacoes"]
        )

        escrever_por_coluna(
            ws, linha, mapa_colunas,
            "Tipo de Documento",
            "Boleto"
        )

        escrever_por_coluna(
            ws, linha, mapa_colunas,
            "Forma de Pagamento",
            "Pagamento de Boleto"
        )

        escrever_por_coluna(
            ws, linha, mapa_colunas,
            "Código de Barras do Boleto",
            row["codigo_barras"]
        )

        escrever_por_coluna(
            ws, linha, mapa_colunas,
            "Departamento (100%)",
            "0006 - Crédito e Cobrança"
        )

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    return output

# =============================
# INTERFACE
# =============================
st.title("📄 Baixa Boletos Omie")
st.caption("Extração automática de boletos PDF e geração da planilha padrão Omie.")

st.markdown("---")

arquivos = st.file_uploader(
    "Selecione os boletos em PDF",
    type=["pdf"],
    accept_multiple_files=True
)

if arquivos:
    st.success(f"{len(arquivos)} arquivo(s) selecionado(s).")

    if st.button("🔎 Extrair informações dos boletos", use_container_width=True):
        registros = []

        progresso = st.progress(0)

        for i, arquivo in enumerate(arquivos):
            try:
                dados = extrair_boleto(arquivo)
                registros.append(dados)
            except Exception as e:
                st.error(f"Erro ao processar {arquivo.name}: {e}")

            progresso.progress((i + 1) / len(arquivos))

        st.session_state["df_boletos"] = pd.DataFrame(registros)

if "df_boletos" in st.session_state:
    st.markdown("### ✅ Conferência dos dados extraídos")

    df_editado = st.data_editor(
        st.session_state["df_boletos"],
        use_container_width=True,
        num_rows="dynamic"
    )

    st.session_state["df_boletos_editado"] = df_editado

    col1, col2 = st.columns(2)

    with col1:
        if st.button("💾 Salvar no Supabase", use_container_width=True):
            ok, msg = salvar_no_supabase(df_editado)
        
            if ok:
                st.success(msg)
            else:
                st.error(msg)

    with col2:
        excel = gerar_excel_omie(df_editado)

        st.download_button(
            label="📥 Baixar planilha Omie preenchida",
            data=excel,
            file_name="Omie_Contas_Pagar_Preenchida.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
