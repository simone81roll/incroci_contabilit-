import streamlit as st
import pandas as pd
import re
from io import BytesIO
from difflib import SequenceMatcher # <-- IMPORTANTE: Aggiunto questo

# --- 1. FUNZIONI DI PULIZIA E UTILITY (Invariate) ---

def pulisci_valuta(valore):
    """Converte una stringa di valuta in un numero float, gestendo vari formati."""
    if isinstance(valore, (int, float)):
        return float(valore)
    if isinstance(valore, str):
        # Rimuove lettere, spazi e punti delle migliaia, poi sostituisce la virgola
        valore_pulito = re.sub(r'[A-Za-z\s\.]', '', valore).replace(',', '.')
        if valore_pulito:
            try:
                return float(valore_pulito)
            except ValueError:
                return 0.0
    return 0.0

def normalizza_nome_cliente(descrizione):
    """Estrae e normalizza il nome del cliente dalla descrizione."""
    if not isinstance(descrizione, str): return "N/D"
    nome = descrizione.upper()
    # Rimuove codici e descrizioni comuni all'inizio
    nome = re.sub(r'^(BDS-\s*)?BON VITTORIA SIN\s*', '', nome)
    # Rimuove forme societarie e parole comuni
    parole_da_rimuovere = [
        'SNC', 'SAS', 'SRL', 'SPA', 'DI', '&', 'C', 'ESTINTORI', 
        'TRAPUNTIFICIO', 'ARREDAMENT'
    ]
    for parola in parole_da_rimuovere:
        nome = re.sub(r'\b' + re.escape(parola) + r'\b', '', nome)
        
    nome_pulito = re.sub(r'\s+', ' ', nome).strip()
    return nome_pulito or "N/D"

def to_excel(df):
    """Converte un DataFrame in un file Excel in memoria per il download."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Report')
    return output.getvalue()

# --- 2. CUORE DELLA LOGICA: LA RICONCILIAZIONE (MODIFICATA) ---

def riconcilia_transazioni(df, tolleranza, soglia_similarita):
    """
    Funzione principale che abbina le transazioni di Dare e Avere.
    Usa una soglia di similarità (fuzzy matching) per le descrizioni.
    """
    # Prepara i dati di lavoro
    df_lavoro = df.copy()
    df_lavoro['ID_Originale'] = df_lavoro.index
    df_lavoro['Importo_Unico'] = df_lavoro['Dare_Num'] + df_lavoro['Avere_Num']
    
    # Normalizza la descrizione per il matching
    df_lavoro['Nome_Norm'] = df_lavoro['Descrizione'].apply(normalizza_nome_cliente)
    
    # Separa crediti (Avere) e debiti (Dare)
    crediti = df_lavoro[df_lavoro['Avere_Num'] > 0].sort_values('Avere_Num').to_dict('records')
    debiti = df_lavoro[df_lavoro['Dare_Num'] > 0].sort_values('Dare_Num').to_dict('records')

    riconciliati = []
    id_usati_debito = set()
    id_usati_credito = set()

    for debito in debiti:
        if debito['ID_Originale'] in id_usati_debito:
            continue
        
        miglior_match = None
        miglior_similarita = -1.0  # Partiamo da un valore impossibile
        miglior_diff = float('inf')
        
        nome_debito = debito['Nome_Norm']

        for credito in crediti:
            if credito['ID_Originale'] in id_usati_credito:
                continue

            diff = abs(debito['Dare_Num'] - credito['Avere_Num'])

            # --- LOGICA DI MATCHING MODIFICATA ---
            # 1. Controlla se l'importo è entro la tolleranza
            if diff <= tolleranza:
                
                nome_credito = credito['Nome_Norm']
                
                # 2. Calcola la similarità (0.0 a 1.0)
                #    Ignoriamo il match se uno dei due è "N/D"
                if nome_debito == "N/D" or nome_credito == "N/D":
                    similarita = 0.0
                else:
                    similarita = SequenceMatcher(None, nome_debito, nome_credito).ratio()

                # 3. Controlla se la similarità è sopra la soglia
                if similarita >= soglia_similarita:
                    
                    # 4. Trovato un candidato. È il migliore finora?
                    #    Priorità 1: Massima Similarità
                    #    Priorità 2: Minima Differenza (a parità di similarità)
                    
                    if similarita > miglior_similarita:
                        # Trovato un match *più simile*
                        miglior_similarita = similarita
                        miglior_diff = diff
                        miglior_match = credito
                    elif similarita == miglior_similarita:
                        # Trovato un match *altrettanto simile*
                        # Scegliamo quello con la differenza di importo minore
                        if diff < miglior_diff:
                            miglior_diff = diff
                            miglior_match = credito
            
        # --- FINE LOGICA DI MATCHING ---

        if miglior_match:
            riconciliati.append({
                'Data_Dare': debito['Data_Reg'],
                'Descrizione_Dare': debito['Descrizione'],
                'Importo_Dare': debito['Dare_Num'],
                'Data_Avere': miglior_match['Data_Reg'],
                'Descrizione_Avere': miglior_match['Descrizione'],
                'Importo_Avere': miglior_match['Avere_Num'],
                'Differenza': miglior_diff,
                'Similarita_Desc': miglior_similarita # <-- Nuova colonna
            })
            id_usati_debito.add(debito['ID_Originale'])
            id_usati_credito.add(miglior_match['ID_Originale'])

    id_riconciliati = id_usati_debito.union(id_usati_credito)
    
    # Gestione DataFrame vuoto
    if not riconciliati:
        df_riconciliati = pd.DataFrame(columns=[
            'Data_Dare', 'Descrizione_Dare', 'Importo_Dare', 
            'Data_Avere', 'Descrizione_Avere', 'Importo_Avere', 
            'Differenza', 'Similarita_Desc'
        ])
    else:
        df_riconciliati = pd.DataFrame(riconciliati)

    df_residui = df[~df.index.isin(id_riconciliati)].copy()

    return df_riconciliati, df_residui

# --- 3. INTERFACCIA STREAMLIT (Modificata) ---

st.set_page_config(layout="wide", page_title="Riconciliazione Intelligente")

st.title("💡 Dashboard di Riconciliazione Intelligente")
st.markdown("Questa app abbina le transazioni in 'Dare' e 'Avere' basandosi su importo e similarità della descrizione.")

# Caricamento File
uploaded_file = st.file_uploader("Carica il tuo file Excel (.xlsx) o CSV (.csv)", type=["xlsx", "csv"])

if uploaded_file:
    # Lettura flessibile del file
    try:
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file, engine='openpyxl')
    except Exception as e:
        st.error(f"Errore nella lettura del file: {e}")
        st.stop()

    # --- SETUP E PRE-PROCESSING ---
    df.columns = [
        'Esercizio', 'Data_Reg', 'N_Reg', 'Sede', 'Descrizione', 'Data_Doc',
        'N_Doc', 'Prot', 'Dare', 'Avere', 'Col11', 'Col12', 'Col13', 'Col14',
        'Col15', 'Col16', 'Col17', 'Col18', 'Col19', 'Col20', 'Col21'
    ]
    df['Dare_Num'] = df['Dare'].apply(pulisci_valuta)
    df['Avere_Num'] = df['Avere'].apply(pulisci_valuta)
    
    # --- PANNELLO DI CONTROLLO (Modificato) ---
    st.sidebar.header("Impostazioni di Riconciliazione")
    
    tolleranza = st.sidebar.number_input(
        "Tolleranza di abbinamento (€)", 
        min_value=0.0, 
        max_value=10.0, 
        value=0.10,  # 10 centesimi di default
        step=0.01,
        help="La massima differenza in euro tra Dare e Avere per considerarli una coppia."
    )

    # --- NUOVO SLIDER PER SIMILARITÀ ---
    soglia_similarita = st.sidebar.slider(
        "Soglia di similarità descrizione",
        min_value=0.0,
        max_value=1.0,
        value=0.3, # Default 30%
        step=0.05,
        help="Quanto devono essere simili le descrizioni (dopo la pulizia) per essere abbinate? 1.0 = identiche, 0.0 = qualsiasi."
    )
    # -----------------------------------

    # --- ESECUZIONE E VISUALIZZAZIONE ---
    if st.sidebar.button("Avvia Riconciliazione", type="primary"):
        with st.spinner("Sto cercando gli abbinamenti..."):
            df_riconciliati, df_residui = riconcilia_transazioni(df, tolleranza, soglia_similarita)

        st.success(f"Analisi completata! Trovate **{len(df_riconciliati)}** coppie riconciliate.")

        # --- 1. TABELLA DELLE TRANSAZIONI RICONCILIATE (Modificata) ---
        st.header("✅ Transazioni Riconciliate")
        if not df_riconciliati.empty:
            st.dataframe(df_riconciliati.style.format({
                'Importo_Dare': '€{:.2f}', 
                'Importo_Avere': '€{:.2f}', 
                'Differenza': '{:.2f}',
                'Similarita_Desc': '{:.1%}' # Formatta come percentuale
            }))
            st.download_button(
                "📥 Scarica Report Riconciliati",
                to_excel(df_riconciliati),
                "report_riconciliati.xlsx"
            )
        else:
            st.info("Nessuna transazione è stata riconciliata con i filtri impostati.")
        st.warning(f"Analisi completata! Trovate **{len(df_residui)}** coppie da riconciliare.")

        # --- 2. TABELLA DELLE TRANSAZIONI RESIDUE ---
        st.header("⚠️ Transazioni Residue non Abbinate")
        if not df_residui.empty:
            st.dataframe(df_residui[['Data_Reg', 'Descrizione', 'Dare_Num', 'Avere_Num']])
            st.download_button(
                "📥 Scarica Report Residui",
                to_excel(df_residui[['Data_Reg', 'Descrizione', 'Dare_Num', 'Avere_Num']]),
                "report_residui.xlsx"
            )
        else:
            st.balloons()
            st.success("Fantastico! Tutte le transazioni sono state riconciliate!")
else:
    st.info("In attesa di un file per avviare l'analisi.")