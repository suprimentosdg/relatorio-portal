import streamlit as st
import pandas as pd
from pymongo import MongoClient
from io import BytesIO
from xlsxwriter import Workbook
from datetime import datetime, time

st.set_page_config(page_title="Relatórios")
st.title("Portal de Suprimentos")
st.subheader("Relatórios")

@st.cache_data
def loading_dadosCham():
    connectString = "mongodb+srv://suprimentosdglobo:suprimentosdg2023@cluster0.dx7yrgp.mongodb.net/?retryWrites=true&w=majority"
    client = MongoClient(connectString)
    db = client["confirmations"]
    mycolection = db.Cl02
    dados_mongodb = list(mycolection.find())
    dd=[r for r in dados_mongodb]
    df1 = pd.DataFrame(dd)
    return df1

@st.cache_data
def loading_dadosConfirm():
    connectString = "mongodb+srv://suprimentosdglobo:suprimentosdg2023@cluster0.dx7yrgp.mongodb.net/?retryWrites=true&w=majority"
    client = MongoClient(connectString)
    db = client["confirmations"]
    mycolection = db.Cl01
    dados_mongodb = list(mycolection.find())
    dd=[r for r in dados_mongodb]
    df2 = pd.DataFrame(dd)
    return df2

options = st.selectbox("Selecione o relatório desejado:", ["Aberturas de chamado", "Confirmações de entrega"])
with st.container():
    if  options == "Aberturas de chamado":
        df1 = loading_dadosCham() 

        st.dataframe(df1)

        if st.button("Exibir Gráficos"):
            st.subheader("Gráfico Geral de Solicitações de Toner:")
            tipo_item1 = "Solicitação de toner"
            df_filtrado1 = df1[df1['opcao'] == tipo_item1]
            contagem_solicitacoes = df_filtrado1['regional'].value_counts()
            contagem_df = pd.DataFrame({'regional': contagem_solicitacoes.index, 'contagem': contagem_solicitacoes.values})
            st.bar_chart(contagem_df.set_index('regional'))

            st.subheader("Gráfico Geral de Aberturas de Chamado:")
            tipo_item2 = "Assistência técnica"
            df_filtrado2 = df1[df1['opcao'] == tipo_item2]
            contagem_aberturas = df_filtrado2['regional'].value_counts()
            contagem_df2 = pd.DataFrame({'regional': contagem_aberturas.index, 'contagem': contagem_aberturas.values})
            st.bar_chart(contagem_df2.set_index('regional'))

        excel_buffer = BytesIO()
        with pd.ExcelWriter(excel_buffer, engine="xlsxwriter") as writer:
            df1.to_excel(writer, index=False, header=True)
        excel_bytes = excel_buffer.getvalue()
        st.download_button(
            label="Baixar Relatório Geral",
            data=excel_bytes,
            file_name=f"relatórioImpressoras.xlsx",
            key="download_button_geral",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        
        st.sidebar.markdown("Filtros")
        df1_data = pd.to_datetime(df1["timestamp"]).dt.date.drop_duplicates()
        min_date = min(df1_data)
        max_date = max(df1_data)

        regionais = df1['regional'].unique()
        regional_selecionada = st.sidebar.selectbox("Selecione a regional:", regionais)

        start_date = st.sidebar.text_input("Digite uma data de início", min_date)
        end_date = st.sidebar.text_input("Digite uma data final", max_date)

        start = pd.to_datetime(start_date)
        end = pd.to_datetime(end_date)

        if start > end:
            st.error("Data final deve ser **Maior** que data inicial")
        
        df1filtered = df1[(df1["regional"] == regional_selecionada) & (pd.to_datetime(df1["timestamp"]) >= start) & (pd.to_datetime(df1["timestamp"]) <= end)]

        st.subheader(f"Dados da Regional: {regional_selecionada}")
        st.dataframe(df1filtered)

        excel_buffer = BytesIO()
        with pd.ExcelWriter(excel_buffer, engine="xlsxwriter") as writer:
            df1filtered.to_excel(writer, index=False, header=True)
        excel_bytes = excel_buffer.getvalue()
        st.download_button(
            label=f"Baixar Relatório da regional **{regional_selecionada}**",
            data=excel_bytes,
            file_name=f"relatórioImpressoras.xlsx",
            key="download_button_regional",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        df2 = loading_dadosConfirm()
        colunasUteis = ["nome", "regional", "loja", "fornecedor", "data_recebimento", "nf", "timestamp"]
        df2 = df2[colunasUteis]
        st.dataframe(df2)

        countsRegions = df2['regional'].value_counts()
        countsRegions_df = pd.DataFrame({'regional': countsRegions.index, 'contagem': countsRegions.values})

        if st.button("Exibir Gráfico"):
            st.bar_chart(countsRegions_df.set_index('regional'))

        excel_buffer = BytesIO()
        with pd.ExcelWriter(excel_buffer, engine="xlsxwriter") as writer:
            df2.to_excel(writer, index=False, header=True)
        excel_bytes = excel_buffer.getvalue()
        st.download_button(
            label="Baixar Relatório",
            data=excel_bytes,
            file_name=f"relatórioEntregas.xlsx",
            key="download_button",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )