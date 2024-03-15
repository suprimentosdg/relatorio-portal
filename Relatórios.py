import streamlit as st
import pandas as pd
from pymongo import MongoClient
from io import BytesIO
from xlsxwriter import Workbook
from datetime import datetime, timedelta
from PIL import Image

st.set_page_config(page_title="Relatórios")
col1, col2 = st.columns([6, 1])
col1.title("Portal de Suprimentos")
image_path = "logo_globo.png"
image = col2.image(image_path, width=80)
st.subheader("Relatórios")
st.write('---')

col3, col4, col5 = st.columns([1, 1.4, 1])
options = col4.selectbox("Selecione o relatório desejado:", ["Gerenciamento de Impressoras", "Confirmações de Entregas"])
with st.container():
    if  options == "Gerenciamento de Impressoras":
        connectString = "mongodb+srv://suprimentosdglobo:suprimentosdg2023@cluster0.dx7yrgp.mongodb.net/?retryWrites=true&w=majority"
        client = MongoClient(connectString)
        db = client["confirmations"]
        mycolection = db.Cl02
        dados_mongodb = list(mycolection.find())
        dd=[r for r in dados_mongodb]
        df = pd.DataFrame(dd)
        df['timestamp'] = pd.to_datetime(df['timestamp'], format="%d/%m/%Y %H:%M:%S") - timedelta(hours=3)

        show_filters = st.checkbox("Exibir Relatório Geral")
        if show_filters:
            st.dataframe(df.drop(columns=['_id']))

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

        if st.button("Exibir Gráficos Gerais"):
            st.subheader("Gráfico Geral de Solicitações de Toner:")
            tipo_item1 = "Solicitação de toner"
            df_filtrado1 = df[df['opcao'] == tipo_item1]
            contagem_solicitacoes = df_filtrado1['regional'].value_counts()
            contagem_df = pd.DataFrame({'regional': contagem_solicitacoes.index, 'contagem': contagem_solicitacoes.values})
            st.bar_chart(contagem_df.set_index('regional'))

            st.subheader("Gráfico Geral de Aberturas de Chamado:")
            tipo_item2 = "Assistência técnica"
            df_filtrado2 = df[df['opcao'] == tipo_item2]
            contagem_aberturas = df_filtrado2['regional'].value_counts()
            contagem_df2 = pd.DataFrame({'regional': contagem_aberturas.index, 'contagem': contagem_aberturas.values})
            st.bar_chart(contagem_df2.set_index('regional'))

            st.subheader("Gráfico do Consumo de Toner por Impressora:")
            tipo_item3 = "Solicitação de toner"
            df_filtrado3 = df[df['opcao'] == tipo_item3]
            contagem_toners_impr = df_filtrado3['impressora'].value_counts()
            contagem_df3 = pd.DataFrame({'impressora': contagem_toners_impr.index, 'contagem': contagem_toners_impr.values})
            st.bar_chart(contagem_df3.set_index('impressora'))

            st.subheader("Gráfico de Assistência técnica por Impressora:")
            tipo_item4 = "Assistência técnica"
            df_filtrado4 = df[df['opcao'] == tipo_item4]
            contagem_abert_impr = df_filtrado3['impressora'].value_counts()
            contagem_df3 = pd.DataFrame({'impressora': contagem_abert_impr.index, 'contagem': contagem_abert_impr.values})
            st.bar_chart(contagem_df3.set_index('impressora'))

        show_filters2 = st.checkbox("Filtragem Geral")
        if show_filters2:
            st.sidebar.markdown("**Filtragem Geral**")
            df1_data = pd.to_datetime(df["timestamp"]).dt.date
            min_date = min(df1_data)
            max_date = max(df1_data)
            min_date = min_date.strftime('%d/%m/%Y')
            max_date = max_date.strftime('%d/%m/%Y')

            start_date = st.sidebar.text_input("Digite uma data de início", min_date)
            end_date = st.sidebar.text_input("Digite uma data final", max_date)
            regionais = df['regional'].unique()
            regional_selecionada = st.sidebar.selectbox("Selecione a regional:", regionais)  
            start = pd.to_datetime(start_date, format='%d/%m/%Y')
            end = pd.to_datetime(end_date, format='%d/%m/%Y') + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)

            if start > end:
                st.error("Data final deve ser **Maior** que data inicial")
            
            df1filtered = df[(df["regional"] == regional_selecionada) & (pd.to_datetime(df["timestamp"]) >= start) & (pd.to_datetime(df["timestamp"]) <= end)]

            df1filtered['timestamp'] = pd.to_datetime(df1filtered['timestamp'])

            df2 = df1filtered

            show_filters3 = st.sidebar.checkbox("Filtro por Solicitação")
            if show_filters3:              
                opcao = df['opcao'].unique()
                opcao_selecionada = st.sidebar.selectbox("Selecione uma opção:", opcao)
                df1filtered = df[(df["opcao"] == opcao_selecionada) & (pd.to_datetime(df["timestamp"]) >= start) & (pd.to_datetime(df["timestamp"]) <= end)]

                st.write("---")

                df1filtered['timestamp'] = pd.to_datetime(df1filtered['timestamp'])
                st.subheader(f"Dados de {opcao_selecionada} da Regional {regional_selecionada}")
                df3 = df1filtered
                st.dataframe(df3.drop(columns=['_id']))

                excel_buffer = BytesIO()
                with pd.ExcelWriter(excel_buffer, engine="xlsxwriter") as writer:
                    df1filtered.to_excel(writer, index=False, header=True)
                excel_bytes = excel_buffer.getvalue()
                st.download_button(
                    label=f"Baixar Relatório da Regional **{regional_selecionada}**",
                    data=excel_bytes,
                    file_name=f"relatórioImpressoras.xlsx",
                    key="download_button_regional",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            else:
                st.subheader(f"Dados filtrados da Regional {regional_selecionada}")
                st.dataframe(df2.drop(columns=['_id']))
                excel_buffer = BytesIO()
                with pd.ExcelWriter(excel_buffer, engine="xlsxwriter") as writer:
                    df2.to_excel(writer, index=False, header=True)
                excel_bytes = excel_buffer.getvalue()
                st.download_button(
                    label=f"Baixar Relatório da Regional **{regional_selecionada}**,
                    data=excel_bytes,
                    file_name=f"relatórioImpressoras.xlsx",
                    key="download_button_regional",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                if st.button(f"Exibir Gráficos da Regional {regional_selecionada}"):
                    st.subheader(f"Gráfico Geral da regional {regional_selecionada}:")
                    df_filtered_options = df1filtered[df1filtered["opcao"].isin(["Assistência técnica", "Solicitação de toner"])]
                    counts = df_filtered_options["opcao"].value_counts()
                    st.bar_chart(counts)

    else:
        connectString = "mongodb+srv://suprimentosdglobo:suprimentosdg2023@cluster0.dx7yrgp.mongodb.net/?retryWrites=true&w=majority"
        client = MongoClient(connectString)
        db = client["confirmations"]
        mycolection = db.Cl01
        dados_mongodb = list(mycolection.find())
        dd=[r for r in dados_mongodb]
        df = pd.DataFrame(dd)
        df['nf'] = df['nf'].astype(str)
        df['timestamp'] = pd.to_datetime(df['timestamp'], format="%d/%m/%Y %H:%M:%S") - timedelta(hours=3)
        show_filters = st.checkbox("Exibir Relatório Geral")
        if show_filters:
            st.dataframe(df.drop(columns=['_id']))

        countsRegions = df['regional'].value_counts()
        countsRegions_df = pd.DataFrame({'regional': countsRegions.index, 'contagem': countsRegions.values})

        excel_buffer = BytesIO()
        with pd.ExcelWriter(excel_buffer, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, header=True)
        excel_bytes = excel_buffer.getvalue()
        st.download_button(
            label="Baixar Relatório Geral",
            data=excel_bytes,
            file_name=f"relatórioConfirmações.xlsx",
            key="download_button",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        if st.button("Exibir Gráficos Gerais"):
            st.bar_chart(countsRegions_df.set_index('regional'))

        show_filters2 = st.checkbox("Filtragem Geral")    
        if show_filters2:
            st.sidebar.markdown("**Filtragem Geral**")
            df1_data = pd.to_datetime(df["timestamp"]).dt.date
            min_date = min(df1_data)
            max_date = max(df1_data)
            min_date = min_date.strftime('%d/%m/%Y')
            max_date = max_date.strftime('%d/%m/%Y')

            start_date = st.sidebar.text_input("Digite uma data de início", min_date)
            end_date = st.sidebar.text_input("Digite uma data final", max_date)

            start = pd.to_datetime(start_date, format='%d/%m/%Y')
            end = pd.to_datetime(end_date, format='%d/%m/%Y') + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)

            if start > end:
                st.error("Data final deve ser **Maior** que data inicial")
            
            df1filtered = df[(pd.to_datetime(df["timestamp"]) >= start) & (pd.to_datetime(df["timestamp"]) <= end)]

            df1filtered['timestamp'] = pd.to_datetime(df1filtered['timestamp'])

            df2 = df1filtered

            show_filters3 = st.sidebar.checkbox("Filtro da Regional")
            if show_filters3:
                regionais = df['regional'].unique()
                regional_selecionada = st.sidebar.selectbox("Selecione a regional:", regionais)
                df1filtered = df[(df["regional"] == regional_selecionada) & (pd.to_datetime(df["timestamp"]) >= start) & (pd.to_datetime(df["timestamp"]) <= end)]

                st.write("---")

                df1filtered['timestamp'] = pd.to_datetime(df1filtered['timestamp'])

                st.subheader(f"Dados da Regional: {regional_selecionada}")
                df3 = df1filtered
                st.dataframe(df3.drop(columns=['_id']))

                excel_buffer = BytesIO()
                with pd.ExcelWriter(excel_buffer, engine="xlsxwriter") as writer:
                    df1filtered.to_excel(writer, index=False, header=True)
                excel_bytes = excel_buffer.getvalue()
                st.download_button(
                    label=f"Baixar Relatório da Regional **{regional_selecionada}**",
                    data=excel_bytes,
                    file_name=f"relatórioConfirmações.xlsx",
                    key="download_button_regional",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                if st.button(f"Exibir Gráfico da Regional {regional_selecionada}"):
                    st.subheader(f"Gráfico Geral da Regional {regional_selecionada}:")
                    df_filtered_options = df1filtered[df1filtered["fornecedor"].isin(["Atlas Papelaria", "Atakadinho Bahia", "Brilhante", "Casa Norte", "Distribuidora Teresina", "Ecopaper", "E Pacheco", "KC Carvalho", "Macropack", "Nacional", "PL", "Supermercado São Jorge (JB)"])]
                    counts = df_filtered_options["fornecedor"].value_counts()
                    st.bar_chart(counts)

            else:
                st.subheader(f"Dados da Filtragem Geral")
                st.dataframe(df2.drop(columns=['_id']))
                excel_buffer = BytesIO()
                with pd.ExcelWriter(excel_buffer, engine="xlsxwriter") as writer:
                    df2.to_excel(writer, index=False, header=True)
                excel_bytes = excel_buffer.getvalue()
                st.download_button(
                    label=f"Baixar Relatório Filtrado",
                    data=excel_bytes,
                    file_name=f"relatórioConfirmações.xlsx",
                    key="download_button_regional",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )