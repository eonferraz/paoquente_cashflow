import streamlit as st
import pandas as pd
import pyodbc
import io
from datetime import date

st.set_page_config(page_title="Relatório de Pagamentos", layout="wide")

# Função de conexão
def conectar_banco():
    return pyodbc.connect(
        'DRIVER={ODBC Driver 17 for SQL Server};'
        'SERVER=sx-global.database.windows.net;'
        'DATABASE=sx_comercial;'
        'UID=paulo.ferraz;'
        'PWD=Gs!^42j$G0f0^EI#ZjRv'
    )

@st.cache_data(ttl=600)
def buscar_dados():
    conn = conectar_banco()
    query = """
        SELECT 
            RAZAO_SOCIAL,
            TIPO_DOC,
            CATEGORIA,
            DESCRICAO,
            PARCELA,
            TOTAL_PARCELAS,
            DATA_LANCAMENTO,
            DATA_VENCIMENTO,
            DATA_INTENCAO,
            VALOR_NOMINAL,
            VALOR_ENCARGOS,
            VALOR_DESCONTOS
        FROM pq_lancamentos 
        WHERE DATA_CANCELAMENTO IS NULL
        AND TIPO = 'Contas à Pagar'
        AND DATA_PAGAMENTO IS NULL
    """
    df = pd.read_sql(query, conn)
    conn.close()
    return df

df_completo = buscar_dados()
df_completo["DATA_INTENCAO"] = pd.to_datetime(df_completo["DATA_INTENCAO"], errors="coerce")

data_inicio, data_fim = st.date_input("Filtrar por intervalo de Data de Intenção", [date.today(), date.today()])
df_filtrado = df_completo[(df_completo["DATA_INTENCAO"].dt.date >= data_inicio) & (df_completo["DATA_INTENCAO"].dt.date <= data_fim)].copy()

st.write("### Contas a Pagar")

if not df_filtrado.empty:
    for col in ['VALOR_NOMINAL', 'VALOR_ENCARGOS', 'VALOR_DESCONTOS']:
        df_filtrado[col] = df_filtrado[col].astype(str).str.replace(".", "", regex=False).str.replace(",", ".", regex=False)
        df_filtrado[col] = pd.to_numeric(df_filtrado[col], errors='coerce').fillna(0)

    df_filtrado['VALOR_TOTAL'] = (
        df_filtrado['VALOR_NOMINAL'] +
        df_filtrado['VALOR_ENCARGOS'] -
        df_filtrado['VALOR_DESCONTOS']
    )

    df_filtrado['PARCELA_TOTAL'] = df_filtrado['PARCELA'].astype(str) + "/" + df_filtrado['TOTAL_PARCELAS'].astype(str)

    colunas_visiveis = [
        'RAZAO_SOCIAL', 'TIPO_DOC', 'CATEGORIA', 'DESCRICAO',
        'PARCELA_TOTAL', 'DATA_LANCAMENTO', 'DATA_VENCIMENTO',
        'DATA_INTENCAO', 'VALOR_TOTAL'
    ]

    df_exibir = df_filtrado[colunas_visiveis].copy()
    df_exibir['Selecionar'] = False

    # Filtro de ordenação
    col_ordenacao = st.selectbox("Ordenar por coluna:", df_exibir.columns.drop("Selecionar"))
    crescente = st.checkbox("Ordem crescente", value=True)
    df_exibir = df_exibir.sort_values(by=col_ordenacao, ascending=crescente).reset_index(drop=True)

    st.write("Selecione as linhas desejadas:")
    edited_df = st.data_editor(
        df_exibir,
        use_container_width=True,
        num_rows="dynamic",
        column_config={
            "Selecionar": st.column_config.CheckboxColumn(label="Selecionar")
        },
        hide_index=True,
        key="data_editor_pagamentos",
        column_order=["Selecionar"] + list(df_exibir.columns.drop("Selecionar")),
        disabled=False
    )

    selecionados = edited_df[edited_df['Selecionar'] == True]
    total = selecionados['VALOR_TOTAL'].sum()
    st.markdown(f"### 💰 Total a Pagar Selecionado: R$ {total:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

    if not selecionados.empty:
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            selecionados.drop(columns=["Selecionar"]).to_excel(writer, index=False, sheet_name="Contas_a_Pagar")
        st.download_button(
            label="⬇️ Exportar Selecionados para Excel",
            data=output.getvalue(),
            file_name="contas_a_pagar.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.warning("Nenhum registro encontrado para o intervalo de datas selecionado.")
