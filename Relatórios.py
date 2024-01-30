import streamlit as st
import pandas as pd
from pymongo import MongoClient
from io import BytesIO
from xlsxwriter import Workbook

st.set_page_config(page_title="Relatórios")
st.title("Portal de Suprimentos")
st.subheader("Relatórios")

options = st.selectbox("Selecione o relatório desejado:", ["Confirmações de entrega", "Aberturas de chamado"])
with st.container():
    if  options == "Confirmações de entrega":
        connectString = "mongodb+srv://suprimentosdglobo:suprimentosdg2023@cluster0.dx7yrgp.mongodb.net/?retryWrites=true&w=majority"
        client = MongoClient(connectString)
        db = client["confirmations"]
        mycolection = db.Cl01
        dados_mongodb = list(mycolection.find())
        dd=[r for r in dados_mongodb]
        df = pd.DataFrame(dd)
        st.dataframe(df)

        countsRegions = df['regional'].value_counts()
        countsRegions_df = pd.DataFrame({'regional': countsRegions.index, 'contagem': countsRegions.values})

        if st.button("Exibir Gráfico"):
            st.bar_chart(countsRegions_df.set_index('regional'))

        excel_buffer = BytesIO()
        with pd.ExcelWriter(excel_buffer, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, header=True)
        excel_bytes = excel_buffer.getvalue()
        st.download_button(
            label="Baixar Relatório",
            data=excel_bytes,
            file_name=f"relatórioEntregas.xlsx",
            key="download_button",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    else:
        connectString = "mongodb+srv://suprimentosdglobo:suprimentosdg2023@cluster0.dx7yrgp.mongodb.net/?retryWrites=true&w=majority"
        client = MongoClient(connectString)
        db = client["confirmations"]
        mycolection = db.Cl02
        dados_mongodb = list(mycolection.find())
        dd=[r for r in dados_mongodb]
        df = pd.DataFrame(dd)
        st.dataframe(df)

        countsRegions = df['regional'].value_counts()
        countsRegions_df = pd.DataFrame({'regional': countsRegions.index, 'contagem': countsRegions.values})

        if st.button("Exibir Gráfico Geral"):
            st.subheader("Gráfico Geral:")
            st.bar_chart(countsRegions_df.set_index('regional'))

        excel_buffer = BytesIO()
        with pd.ExcelWriter(excel_buffer, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, header=True)
        excel_bytes = excel_buffer.getvalue()
        st.download_button(
            label="Baixar Relatório Geral",
            data=excel_bytes,
            file_name=f"relatórioImpressoras.xlsx",
            key="download_button",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        
        if st.button(f"Filtrar por Regional"):
            regionais = df['regional'].unique()
            regional_selecionada = st.selectbox("Selecione a regional para filtrar as solicitações:", regionais)

            if regional_selecionada:
                df_filtrado = df[df['regional'] == regional_selecionada]
            st.subheader(f"Dados da Regional: {regional_selecionada}")
            st.dataframe(df_filtrado)

            contagem_regional_filtrada = df_filtrado['regional'].value_counts()
            contagem_regional_filtrada_df = pd.DataFrame({'regional': contagem_regional_filtrada.index, 'contagem': contagem_regional_filtrada.values})

            st.subheader(f"Gráfico da Regional {regional_selecionada}")
            st.bar_chart(contagem_regional_filtrada_df.set_index('regional'))

            excel_buffer = BytesIO()
            with pd.ExcelWriter(excel_buffer, engine="xlsxwriter") as writer:
                df_filtrado.to_excel(writer, index=False, header=True)
            excel_bytes = excel_buffer.getvalue()
            st.download_button(
                label=f"Baixar Relatório da regional {regional_selecionada}",
                data=excel_bytes,
                file_name=f"relatórioImpressoras.xlsx",
                key="download_button",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )