import streamlit as st
import pandas as pd
import pyodbc
import io
from datetime import date

st.set_page_config(page_title="Relatório de Pagamentos", layout="wide")

# Função de conexão (não cachear)
def conectar_banco():
    return pyodbc.connect(
        'DRIVER={ODBC Driver 17 for SQL Server};'
        'SERVER=sx-global.database.windows.net;'
        'DATABASE=sx_comercial;'
        'UID=paulo.ferraz;'
        'PWD=Gs!^42j$G0f0^EI#ZjRv'
    )

# Busca dados sem filtro de data
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

# Converter DATA_INTENCAO para datetime
df_completo["DATA_INTENCAO"] = pd.to_datetime(df_completo["DATA_INTENCAO"], errors="coerce")

# Input do filtro de intervalo de datas
data_inicio, data_fim = st.date_input("Filtrar por intervalo de Data de Intenção", [date.today(), date.today()])

# Aplica o filtro localmente
df_filtrado = df_completo[(df_completo["DATA_INTENCAO"].dt.date >= data_inicio) & (df_completo["DATA_INTENCAO"].dt.date <= data_fim)].copy()

st.write("### Contas a Pagar")

if not df_filtrado.empty:
    # Converter valores numéricos para float
    for col in ['VALOR_NOMINAL', 'VALOR_ENCARGOS', 'VALOR_DESCONTOS']:
        df_filtrado[col] = pd.to_numeric(df_filtrado[col], errors='coerce').fillna(0)

    df_filtrado['VALOR_TOTAL'] = (
        df_filtrado['VALOR_NOMINAL'] +
        df_filtrado['VALOR_ENCARGOS'] -
        df_filtrado['VALOR_DESCONTOS']
    )

    # Interface com checkboxes por linha na tabela
    st.write("Selecione as linhas desejadas:")
    df_filtrado['Selecionar'] = False
    edited_df = st.data_editor(
        df_filtrado,
        use_container_width=True,
        num_rows="dynamic",
        column_config={"Selecionar": st.column_config.CheckboxColumn(label="Selecionar")}
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
