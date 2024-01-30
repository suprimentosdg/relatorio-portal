import streamlit as st
import pandas as pd
from pymongo import MongoClient

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
        if st.button("Baixar relatório"):
            caminho_arquivo = "relatorioEntregas.xlsx"
            df.to_excel(caminho_arquivo, index=False, sheet_name="RelatorioEntregas")
            st.success(f"Relatório exportado com sucesso para {caminho_arquivo}")

            with open(caminho_arquivo, "rb") as arquivo:
                arquivo_bytes = arquivo.read()

            st.write(arquivo_bytes, unsafe_allow_html=True)
    else:
        connectString = "mongodb+srv://suprimentosdglobo:suprimentosdg2023@cluster0.dx7yrgp.mongodb.net/?retryWrites=true&w=majority"
        client = MongoClient(connectString)
        db = client["confirmations"]
        mycolection = db.Cl02
        dados_mongodb = list(mycolection.find())
        dd=[r for r in dados_mongodb]
        df = pd.DataFrame(dd)
        st.dataframe(df)
        if st.button("Baixar relatório"):
            caminho_arquivo = "relatorioImpressoras.xlsx"
            df.to_excel(caminho_arquivo, index=False, sheet_name="RelatorioImpressoras")
            st.success(f"Relatório exportado com sucesso para {caminho_arquivo}")

            with open(caminho_arquivo, "rb") as arquivo:
                arquivo_bytes = arquivo.read()

            st.write(arquivo_bytes, unsafe_allow_html=True)