import streamlit as st
import pandas as pd
from datetime import datetime

# Configurazione della pagina
st.set_page_config(
    page_title="ATS Liguria - Strumento Ispezione Sicurezza Antincendio",
    page_icon="🔥",
    layout="wide"
)

# Titolo principale
st.title("🔥 Strumento di Ispezione Sicurezza Antincendio - ATS Liguria")
st.markdown("Benvenuto nel portale web per la gestione e la compilazione dei verbali di ispezione antincendio.")

# Sidebar per la navigazione
st.sidebar.header("Navigazione Ispezione")
menu = st.sidebar.selectbox("Seleziona Sezione", ["Nuova Ispezione", "Visualizza Verbali Salvati", "Linee Guida ATS"])

if menu == "Nuova Ispezione":
    st.header("Compilazione Nuovo Verbale di Ispezione")
    
    with st.form("form_ispezione"):
        st.subheader("1. Dati Generali")
        col1, col2 = st.columns(2)
        with col1:
            nome_ispettore = st.text_input("Nome Ispettore")
            struttura = st.text_input("Nome Struttura / Azienda")
        with col2:
            data_ispezione = st.date_input("Data Ispezione", value=datetime.today())
            indirizzo = st.text_input("Indirizzo Struttura")
            
        st.subheader("2. Controlli Principali")
        estintori = st.selectbox("Controllo Estintori e Idranti", ["Conforme", "Non Conforme", "Non Applicabile"])
        uscite = st.selectbox("Vie di Esodo e Uscite di Sicurezza libere", ["Conforme", "Non Conforme", "Non Applicabile"])
        porte_tagliafuoco = st.selectbox("Integrità Porte Tagliafuoco", ["Conforme", "Non Conforme", "Non Applicabile"])
        impianto_allarme = st.selectbox("Funzionalità Impianto Allarme", ["Conforme", "Non Conforme", "Non Applicabile"])
        
        st.subheader("3. Note e Osservazioni")
        note = st.text_area("Inserire eventuali note o prescrizioni da rispettare:")
        
        submitted = st.form_submit_button("Salva e Genera Verbale")
        
        if submitted:
            if not nome_ispettore or not struttura:
                st.error("Per favore, compila almeno 'Nome Ispettore' e 'Nome Struttura'.")
            else:
                st.success("Verbale registrato con successo!")
                st.write(f"**Struttura:** {struttura}")
                st.write(f"**Ispettore:** {nome_ispettore}")
                st.write(f"**Data:** {data_ispezione}")
                st.write(f"**Stato Estintori:** {estintori}")
                st.write(f"**Stato Vie di Esodo:** {uscite}")
                st.write(f"**Note:** {note if note else 'Nessuna nota inserita.'}")

elif menu == "Visualizza Verbali Salvati":
    st.header("Archivio Verbali")
    st.info("Qui saranno elencati i verbali precedentemente salvati nel sistema.")

elif menu == "Linee Guida ATS":
    st.header("Linee Guida e Normative Antincendio - ATS Liguria")
    st.markdown("""
    * D.Lgs. 81/08 e s.m.i.
    * Codice di Prevenzione Incendi (D.M. 3 Agosto 2015 e aggiornamenti).
    * Procedure operative interne ATS Liguria per il controllo della sicurezza antincendio nelle strutture sanitarie e produttive.
    """)
