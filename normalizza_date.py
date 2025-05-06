import streamlit as st
import pandas as pd
from datetime import datetime
import io
import numpy as np
from dateutil import parser

st.set_page_config(page_title="Normalizzazione Date in Excel", layout="wide")
st.title("Normalizzazione Date in Excel")
st.write("Carica un file Excel per convertire le date nel formato gg-mm-aaaa e ordinare i dati cronologicamente")

# Funzione per normalizzare le date
def normalizza_data(data, solo_formato=False):
    """
    Funzione che normalizza le date in vari formati.
    
    Args:
        data: Il valore da normalizzare
        solo_formato: Se True, restituisce solo la stringa formattata. 
                     Se False, restituisce anche l'oggetto datetime per l'ordinamento.
    
    Returns:
        Se solo_formato=True: stringa in formato 'dd-mm-yyyy'
        Se solo_formato=False: tupla (stringa formattata, oggetto datetime)
    """
    try:
        # Prova diverse interpretazioni della data
        formati = [
            '%Y-%m-%d', '%d/%m/%Y', '%m/%d/%Y', 
            '%d-%m-%Y', '%m-%d-%Y', '%Y/%m/%d',
            '%d.%m.%Y', '%m.%d.%Y', '%Y.%m.%d',
            '%d %b %Y', '%d %B %Y', '%b %d, %Y', '%B %d, %Y',
            '%Y%m%d', '%d-%b-%Y', '%d-%B-%Y',
            '%a, %d %b %Y', '%A, %d %b %Y', '%A, %d %B %Y',
            '%A %d %B %Y'  # Formato italiano: "giovedì 12 giugno 2025"
        ]
        
        dt_obj = None
        
        # Se è già un datetime o timestamp pandas, lo usiamo direttamente
        if isinstance(data, (pd.Timestamp, datetime)):
            dt_obj = data
            formatted = data.strftime('%d-%m-%Y')
        
        # Se è una stringa, proviamo a interpretarla
        elif isinstance(data, str):
            # Puliamo la stringa
            data = data.strip()
            
            # Prima proviamo con formati specifici
            for formato in formati:
                try:
                    dt_obj = datetime.strptime(data, formato)
                    formatted = dt_obj.strftime('%d-%m-%Y')
                    break
                except ValueError:
                    continue
            
            # Se non funziona, proviamo con dateutil.parser che è più flessibile
            if dt_obj is None:
                try:
                    # Per i formati italiani, sostituiamo i nomi dei mesi in inglese
                    data_temp = data
                    mesi_it_to_en = {
                        'gennaio': 'January', 'febbraio': 'February', 'marzo': 'March',
                        'aprile': 'April', 'maggio': 'May', 'giugno': 'June',
                        'luglio': 'July', 'agosto': 'August', 'settembre': 'September',
                        'ottobre': 'October', 'novembre': 'November', 'dicembre': 'December'
                    }
                    giorni_it_to_en = {
                        'lunedì': 'Monday', 'martedì': 'Tuesday', 'mercoledì': 'Wednesday',
                        'giovedì': 'Thursday', 'venerdì': 'Friday', 'sabato': 'Saturday',
                        'domenica': 'Sunday'
                    }
                    
                    # Sostituiamo i nomi dei mesi italiani con quelli inglesi
                    for mese_it, mese_en in mesi_it_to_en.items():
                        if mese_it in data_temp.lower():
                            data_temp = data_temp.lower().replace(mese_it, mese_en).capitalize()
                            break
                    
                    # Sostituiamo i nomi dei giorni italiani con quelli inglesi
                    for giorno_it, giorno_en in giorni_it_to_en.items():
                        if giorno_it in data_temp.lower():
                            data_temp = data_temp.lower().replace(giorno_it, giorno_en).capitalize()
                            break
                    
                    # Proviamo prima con la data modificata se è stata fatta una sostituzione
                    if data_temp != data:
                        try:
                            dt_obj = parser.parse(data_temp, dayfirst=True)
                            formatted = dt_obj.strftime('%d-%m-%Y')
                        except:
                            # Se fallisce, proviamo con la data originale
                            dt_obj = parser.parse(data, dayfirst=True)  # Assumiamo giorno prima del mese per ambiguità
                            formatted = dt_obj.strftime('%d-%m-%Y')
                    else:
                        # Se non ci sono state sostituzioni, usiamo la data originale
                        dt_obj = parser.parse(data, dayfirst=True)  # Assumiamo giorno prima del mese per ambiguità
                        formatted = dt_obj.strftime('%d-%m-%Y')
                except:
                    pass
        
        # Se è un numero, potrebbe essere un timestamp Excel
        elif isinstance(data, (int, float)) and not np.isnan(data):
            try:
                dt_obj = pd.Timestamp.fromordinal(pd.Timestamp('1899-12-30').to_ordinal() + int(data))
                formatted = dt_obj.strftime('%d-%m-%Y')
            except:
                try:
                    # Potrebbe essere un timestamp UNIX (in secondi)
                    dt_obj = datetime.fromtimestamp(data)
                    formatted = dt_obj.strftime('%d-%m-%Y')
                except:
                    pass
        
        # Se abbiamo trovato un oggetto datetime valido
        if dt_obj is not None:
            if solo_formato:
                return formatted
            else:
                return formatted, dt_obj
        
        # Se non siamo riusciti a interpretare, restituiamo il valore originale
        if solo_formato:
            return data
        else:
            return data, None
            
    except Exception as e:
        if solo_formato:
            return data
        else:
            return data, None

# Sidebar per le opzioni
with st.sidebar:
    st.header("Opzioni")
    st.write("Configura le opzioni per la normalizzazione delle date")
    
    # Opzione per scegliere se ordinare cronologicamente
    ordina_date = st.checkbox("Ordina per data", value=True, help="Ordina i dati in ordine cronologico")
    
    # Formato di output
    formato_output = st.selectbox(
        "Formato di visualizzazione delle date", 
        options=["gg-mm-aaaa", "gg/mm/aaaa", "aaaa-mm-gg"],
        index=0,
        help="Scegli il formato di visualizzazione delle date (internamente saranno comunque normalizzate)"
    )
    
    # Mappatura formati
    formati_output = {
        "gg-mm-aaaa": "%d-%m-%Y",
        "gg/mm/aaaa": "%d/%m/%Y",
        "aaaa-mm-gg": "%Y-%m-%d"
    }

# Upload del file
file = st.file_uploader("Carica un file Excel", type=["xlsx", "xls"])

if file is not None:
    try:
        # Leggiamo il file assicurandoci che la prima riga sia usata come intestazione
        df = pd.read_excel(file, header=0)
        
        st.write("Anteprima del file caricato:")
        st.dataframe(df.head())
        
        # Nome della prima colonna (colonna delle date)
        colonna_date = df.columns[0]
        
        # Mostriamo alcune date prima della normalizzazione
        st.write(f"Esempi di date prima della normalizzazione nella colonna '{colonna_date}':")
        st.write(df[colonna_date].head().tolist())
        
        # Verifica se l'etichetta della colonna è una data
        etichetta_colonna = colonna_date
        st.info(f"Colonna selezionata per la normalizzazione: '{etichetta_colonna}'")
        
        # Creiamo una colonna temporanea per gli oggetti datetime
        df_temp = df.copy()
        
        # Normalizzazione e conversione (solo sui dati, non sulle etichette)
        risultati = df_temp[colonna_date].apply(normalizza_data)
        
        # Separiamo la stringa formattata e l'oggetto datetime
        if len(df_temp) > 0 and isinstance(risultati.iloc[0], tuple):
            df_temp['_data_formattata'] = risultati.apply(lambda x: x[0] if isinstance(x, tuple) else x)
            df_temp['_data_oggetto'] = risultati.apply(lambda x: x[1] if isinstance(x, tuple) and x[1] is not None else pd.NaT)
            
            # Contiamo quanti valori sono stati convertiti correttamente
            num_convertiti = (~df_temp['_data_oggetto'].isna()).sum()
            perc_convertiti = (num_convertiti / len(df_temp)) * 100 if len(df_temp) > 0 else 0
            
            st.write(f"**Stato conversione:** {num_convertiti} su {len(df_temp)} date convertite correttamente ({perc_convertiti:.1f}%)")
            
            if perc_convertiti < 100:
                st.warning(f"Alcune date ({len(df_temp) - num_convertiti}) non sono state convertite correttamente. "
                          f"Potrebbero essere in formati non riconosciuti o non rappresentare date valide.")
                
                # Mostriamo i valori problematici
                problematici = df_temp[df_temp['_data_oggetto'].isna()][[colonna_date]]
                if not problematici.empty:
                    with st.expander(f"Mostra valori problematici ({len(problematici)} record)"):
                        # Aggiunge un indice per identificare le righe problematiche
                        problematici = problematici.reset_index().rename(columns={"index": "Riga nel file"})
                        st.write("**Date non riconosciute:**")
                        st.dataframe(problematici)
                        
                        # Mostra dettagli aggiuntivi sui valori problematici
                        st.write("**Dettagli dei valori problematici:**")
                        for idx, row in problematici.iterrows():
                            st.write(f"Riga {row['Riga nel file']}: '{row[colonna_date]}' (tipo: {type(row[colonna_date]).__name__})")
                        
                        # Aggiunge un'opzione per scaricare solo le righe problematiche
                        output_problematici = io.BytesIO()
                        with pd.ExcelWriter(output_problematici, engine='xlsxwriter') as writer:
                            problematici.to_excel(writer, index=False)
                            
                        st.download_button(
                            label="Scarica elenco date problematiche",
                            data=output_problematici.getvalue(),
                            file_name="date_problematiche.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
            
            # Ordinamento cronologico se richiesto
            if ordina_date and perc_convertiti > 0:
                st.write("Ordinamento dati in ordine cronologico...")
                df_temp = df_temp.sort_values(by='_data_oggetto', na_position='last')
            
            # Applichiamo il formato di output scelto
            formato_selezionato = formati_output[formato_output]
            df_temp[colonna_date] = df_temp['_data_oggetto'].apply(
                lambda x: x.strftime(formato_selezionato) if pd.notna(x) else df_temp.loc[df_temp['_data_oggetto'] == x, '_data_formattata'].values[0]
                if isinstance(df_temp.loc[df_temp['_data_oggetto'] == x, '_data_formattata'].values, np.ndarray) and len(df_temp.loc[df_temp['_data_oggetto'] == x, '_data_formattata'].values) > 0
                else str(x)
            )
            
            # Rimuoviamo le colonne temporanee
            df = df_temp.drop(columns=['_data_formattata', '_data_oggetto'])
        else:
            # Fallback alla vecchia logica
            df[colonna_date] = df[colonna_date].apply(lambda x: normalizza_data(x, True))
        
        # Mostriamo alcune date dopo la normalizzazione
        st.write(f"Esempi di date dopo la normalizzazione nella colonna '{colonna_date}':")
        st.write(df[colonna_date].head().tolist())
        
        # Visualizziamo il dataframe modificato
        st.write("Anteprima del file con date normalizzate:")
        st.dataframe(df.head(10))
        
        # Statistiche sulla colonna delle date
        with st.expander("Statistiche della colonna delle date"):
            # Controlla se abbiamo date valide per le statistiche
            date_valide = df_temp[pd.notna(df_temp['_data_oggetto'])]['_data_oggetto']
            
            if not date_valide.empty:
                col1, col2 = st.columns(2)
                
                with col1:
                    st.write("**Data più vecchia:**")
                    st.write(date_valide.min().strftime(formati_output[formato_output]))
                
                with col2:
                    st.write("**Data più recente:**")
                    st.write(date_valide.max().strftime(formati_output[formato_output]))
                
                # Calcoliamo l'intervallo di tempo totale
                delta = date_valide.max() - date_valide.min()
                st.write(f"**Intervallo temporale**: {delta.days} giorni")
            else:
                st.write("Non ci sono date valide per calcolare le statistiche.")
        
        # Opzione per scaricare il file modificato
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            # Prepariamo un dataframe con le date come oggetti datetime per l'export
            df_export = df.copy()
            
            # Convertiamo la colonna delle date in oggetti datetime per Excel
            if '_data_oggetto' in df_temp:
                # Utilizziamo gli oggetti datetime che abbiamo già generato
                df_export[colonna_date] = df_temp['_data_oggetto']
            else:
                # Tentiamo di convertire le stringhe di data in datetime
                try:
                    df_export[colonna_date] = pd.to_datetime(df_export[colonna_date], dayfirst=True)
                except:
                    pass  # Manteniamo il formato originale se la conversione fallisce
                
            # Salviamo il file con le date in formato Excel nativo
            df_export.to_excel(writer, index=False)
            
            # Otteniamo un riferimento al foglio di lavoro
            worksheet = writer.sheets['Sheet1']
            
            # Formattazione specifica per le date nella prima colonna
            date_format = writer.book.add_format({'num_format': 'dd/mm/yyyy'})
            worksheet.set_column(0, 0, 15, date_format)
        
        st.download_button(
            label="Scarica Excel con date normalizzate",
            data=output.getvalue(),
            file_name="date_normalizzate.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        # Aggiunge una nota informativa
        st.info("Le date sono state normalizzate e possono ora essere utilizzate per l'ordinamento cronologico.")
        
        # Aggiungiamo un footer con istruzioni
        st.markdown("---")
        st.markdown("**Suggerimenti per l'utilizzo:**")
        st.markdown("1. I dati sono stati ordinati cronologicamente se hai selezionato l'opzione.")
        st.markdown("2. Puoi cambiare il formato di visualizzazione nella barra laterale.")
        st.markdown("3. Se alcune date non sono state riconosciute, controlla i 'valori problematici'.")
        
    except Exception as e:
        st.error(f"Si è verificato un errore: {e}")
        st.error("Dettagli dell'errore per il debug:")
        st.exception(e)