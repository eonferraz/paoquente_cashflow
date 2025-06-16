import streamlit as st
import pandas as pd
import pyodbc
import io
from datetime import date

st.set_page_config(page_title="Relatório de Pagamentos", layout="wide")

# Função de conexão
@st.cache_data(ttl=600)
def conectar_banco():
    return pyodbc.connect(
        'DRIVER={ODBC Driver 17 for SQL Server};'
        'SERVER=sx-global.database.windows.net;'
        'DATABASE=sx_comercial;'
        'UID=paulo.ferraz;'
        'PWD=Gs!^42j$G0f0^EI#ZjRv'
    )

# Busca os dados SEM filtro de data
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
            DATA_PAGAMENTO,
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

# Carregar dados
df_completo = buscar_dados()

# Input de filtro de data
data_ref = st.date_input("Filtrar por Data de Vencimento", value=date.today())

# Aplicar filtro localmente no DataFrame
df_filtrado = df_completo[df_completo["DATA_VENCIMENTO"].dt.date == data_ref]

st.write("### Contas a Pagar")

if not df_filtrado.empty:
    df_filtrado['Selecionado'] = False
    checkboxes = []

    for i in range(len(df_filtrado)):
        col1, col2 = st.columns([0.05, 0.95])
        with col1:
            check = st.checkbox("", key=f"check_{i}")
            checkboxes.append(check)
        with col2:
            st.write(df_filtrado.iloc[i, :-1].to_frame().T)

    df_filtrado['Selecionado'] = checkboxes
    df_filtrado['VALOR_TOTAL'] = df_filtrado['VALOR_NOMINAL'] + df_filtrado['VALOR_ENCARGOS'] - df_filtrado['VALOR_DESCONTOS']

    total = df_filtrado[df_filtrado['Selecionado']]['VALOR_TOTAL'].sum()
    st.markdown(f"### 💰 Total a Pagar Selecionado: R$ {total:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

    df_exportar = df_filtrado[df_filtrado['Selecionado']].drop(columns=["Selecionado"])
    if not df_exportar.empty:
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_exportar.to_excel(writer, index=False, sheet_name="Contas_a_Pagar")
        st.download_button(
            label="⬇️ Exportar Selecionados para Excel",
            data=output.getvalue(),
            file_name="contas_a_pagar.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.warning("Nenhum registro encontrado para a data selecionada.")
