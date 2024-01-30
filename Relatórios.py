import streamlit as st
import pandas as pd
from pymongo import MongoClient
from io import BytesIO
from xlsxwriter import Workbook
import datetime

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
            key="download_button_geral",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        
        if st.button(f"Filtrar por Regional"):
            regionais = df['regional'].unique()
            regional_selecionada = st.selectbox("Selecione a regional para filtrar as solicitações:", regionais)

            st.markdown("Filtrar por data")
            data_inicio = st.date_input("Selecione a data inicial:")
            data_fim = st.date_input("Selecione a data final:")

            if data_inicio:
                data_inicio = datetime.combine(data_inicio, datetime.min.time())
            if data_fim:
                data_fim = datetime.combine(data_fim, datetime.max.time())
            
            df['timestamp'] = pd.to_datetime(df['timestamp'])

            if data_inicio and data_fim:
                df_filtrado = df[(df['timestamp'] >= data_inicio) & (df['timestamp'] <= data_fim)]
            elif data_inicio:
                df_filtrado = df[df['timestamp'] >= data_inicio]
            elif data_fim:
                df_filtrado = df[df['timestamp'] <= data_fim]
            else:
                df_filtrado = df

            if regional_selecionada:
                df_filtrado = df[df['regional'] == regional_selecionada]
            st.subheader(f"Dados da Regional: {regional_selecionada}")
            st.dataframe(df_filtrado)

            excel_buffer = BytesIO()
            with pd.ExcelWriter(excel_buffer, engine="xlsxwriter") as writer:
                df_filtrado.to_excel(writer, index=False, header=True)
            excel_bytes = excel_buffer.getvalue()
            st.download_button(
                label=f"Baixar Relatório da regional {regional_selecionada}",
                data=excel_bytes,
                file_name=f"relatórioImpressoras.xlsx",
                key="download_button_regional",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )