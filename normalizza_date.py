import streamlit as st
import pandas as pd
from datetime import datetime
import io
import numpy as np
from dateutil import parser

st.set_page_config(page_title="Normalizzazione Date in Excel", layout="wide")
st.title("Normalizzazione Date in Excel")
st.write("Carica un file Excel per convertire le date nel formato desiderato e ordinare i dati cronologicamente")
st.write("✨ **Novità**: Puoi selezionare una o più colonne da normalizzare!")

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

def elabora_foglio(df, colonne_selezionate, colonna_ordinamento, ordina_date, formato_output, formati_output, nome_foglio=""):
    """
    Funzione per elaborare un singolo foglio di Excel
    
    Returns:
        df_elaborato, statistiche_conversione, df_temp_con_oggetti_data
    """
    # Creiamo una copia del dataframe per le modifiche
    df_temp = df.copy()
    
    # Dizionario per memorizzare le statistiche di conversione per ogni colonna
    statistiche_conversione = {}
    
    prefisso_nome = f" ({nome_foglio})" if nome_foglio else ""
    
    # Normalizzazione per ogni colonna selezionata
    for colonna_date in colonne_selezionate:
        if colonna_date not in df_temp.columns:
            st.warning(f"⚠️ Colonna '{colonna_date}' non trovata nel foglio{prefisso_nome}. Saltata.")
            continue
            
        if nome_foglio:
            st.write(f"### Normalizzazione colonna '{colonna_date}' - Foglio '{nome_foglio}'")
        else:
            st.write(f"### Normalizzazione colonna: '{colonna_date}'")
        
        # Normalizzazione e conversione (solo sui dati, non sulle etichette)
        risultati = df_temp[colonna_date].apply(normalizza_data)
        
        # Separiamo la stringa formattata e l'oggetto datetime
        if len(df_temp) > 0 and isinstance(risultati.iloc[0], tuple):
            df_temp[f'_data_formattata_{colonna_date}'] = risultati.apply(lambda x: x[0] if isinstance(x, tuple) else x)
            df_temp[f'_data_oggetto_{colonna_date}'] = risultati.apply(lambda x: x[1] if isinstance(x, tuple) and x[1] is not None else pd.NaT)
            
            # Contiamo quanti valori sono stati convertiti correttamente
            num_convertiti = (~df_temp[f'_data_oggetto_{colonna_date}'].isna()).sum()
            perc_convertiti = (num_convertiti / len(df_temp)) * 100 if len(df_temp) > 0 else 0
            
            # Salviamo le statistiche
            statistiche_conversione[colonna_date] = {
                'convertiti': num_convertiti,
                'totali': len(df_temp),
                'percentuale': perc_convertiti,
                'foglio': nome_foglio
            }
            
            st.write(f"**Stato conversione per '{colonna_date}'{prefisso_nome}:** {num_convertiti} su {len(df_temp)} date convertite correttamente ({perc_convertiti:.1f}%)")
            
            if perc_convertiti < 100:
                st.warning(f"Alcune date nella colonna '{colonna_date}'{prefisso_nome} ({len(df_temp) - num_convertiti}) non sono state convertite correttamente.")
                
                # Mostriamo i valori problematici per questa colonna
                problematici = df_temp[df_temp[f'_data_oggetto_{colonna_date}'].isna()][[colonna_date]]
                if not problematici.empty:
                    with st.expander(f"Mostra valori problematici per '{colonna_date}'{prefisso_nome} ({len(problematici)} record)"):
                        # Aggiunge un indice per identificare le righe problematiche
                        problematici = problematici.reset_index().rename(columns={"index": "Riga nel file"})
                        st.write(f"**Date non riconosciute nella colonna '{colonna_date}'{prefisso_nome}:**")
                        st.dataframe(problematici)
            
            # Applichiamo il formato di output scelto per questa colonna
            formato_selezionato = formati_output[formato_output]
            
            # Funzione helper per applicare il formato
            def applica_formato(row):
                data_obj = row[f'_data_oggetto_{colonna_date}']
                data_formattata = row[f'_data_formattata_{colonna_date}']
                
                if pd.notna(data_obj):
                    return data_obj.strftime(formato_selezionato)
                else:
                    # Se non abbiamo un oggetto datetime valido, usiamo la stringa formattata
                    return data_formattata if pd.notna(data_formattata) else str(row[colonna_date])
            
            df_temp[colonna_date] = df_temp.apply(applica_formato, axis=1)
        else:
            # Fallback alla vecchia logica
            df_temp[colonna_date] = df_temp[colonna_date].apply(lambda x: normalizza_data(x, True))
            statistiche_conversione[colonna_date] = {
                'convertiti': len(df_temp),
                'totali': len(df_temp),
                'percentuale': 100.0,
                'foglio': nome_foglio
            }
    
    # Ordinamento cronologico se richiesto (usa la colonna di ordinamento selezionata)
    if ordina_date and colonna_ordinamento and colonna_ordinamento in df_temp.columns:
        if f'_data_oggetto_{colonna_ordinamento}' in df_temp.columns:
            perc_convertiti_ordinamento = statistiche_conversione.get(colonna_ordinamento, {}).get('percentuale', 0)
            if perc_convertiti_ordinamento > 0:
                st.write(f"Ordinamento dati in ordine cronologico basato sulla colonna '{colonna_ordinamento}'{prefisso_nome}...")
                df_temp = df_temp.sort_values(by=f'_data_oggetto_{colonna_ordinamento}', na_position='last')
    
    # Rimuoviamo le colonne temporanee per il dataframe finale
    colonne_temp = [col for col in df_temp.columns if col.startswith('_data_formattata_') or col.startswith('_data_oggetto_')]
    df_elaborato = df_temp.drop(columns=colonne_temp)
    
    return df_elaborato, statistiche_conversione, df_temp

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
        # Prima leggiamo i nomi dei fogli disponibili
        excel_file = pd.ExcelFile(file)
        fogli_disponibili = excel_file.sheet_names
        
        st.write("### 📋 Seleziona il foglio di calcolo")
        st.write(f"**Fogli disponibili nel file:** {len(fogli_disponibili)}")
        
        # Se c'è solo un foglio, lo selezioniamo automaticamente
        if len(fogli_disponibili) == 1:
            foglio_selezionato = fogli_disponibili[0]
            st.info(f"📄 Foglio selezionato automaticamente: **{foglio_selezionato}**")
        else:
            # Se ci sono più fogli, permettiamo all'utente di scegliere
            foglio_selezionato = st.selectbox(
                "Scegli il foglio di calcolo da elaborare:",
                options=fogli_disponibili,
                index=0,
                help=f"Il file contiene {len(fogli_disponibili)} fogli. Seleziona quello che contiene le date da normalizzare."
            )
            st.write(f"**Foglio selezionato:** {foglio_selezionato}")
        
        # Se ci sono più fogli, offriamo anche l'opzione di elaborarli tutti
        elabora_tutti_fogli = False
        if len(fogli_disponibili) > 1:
            with st.expander("🔄 Opzioni avanzate per più fogli"):
                elabora_tutti_fogli = st.checkbox(
                    "Elabora tutti i fogli del file Excel",
                    value=False,
                    help="Se selezionato, tutte le colonne selezionate verranno normalizzate in tutti i fogli del file Excel"
                )
                
                if elabora_tutti_fogli:
                    st.warning("⚠️ **Attenzione**: Questa opzione elaborerà TUTTI i fogli del file. Assicurati che le colonne selezionate esistano in tutti i fogli.")
                    st.write(f"**Fogli che verranno elaborati:** {', '.join(fogli_disponibili)}")
        
        # Leggiamo il foglio selezionato (o il primo se elaboriamo tutti)
        df = pd.read_excel(file, sheet_name=foglio_selezionato, header=0)
        
        st.write("### 📊 Anteprima del foglio selezionato:")
        st.write(f"**Dimensioni:** {len(df)} righe × {len(df.columns)} colonne")
        st.dataframe(df.head())
        
        # Selezione delle colonne da normalizzare
        st.write("### Seleziona le colonne da normalizzare")
        colonne_disponibili = df.columns.tolist()
        
        # Widget per selezionare multiple colonne
        colonne_selezionate = st.multiselect(
            "Scegli una o più colonne contenenti date da normalizzare:",
            options=colonne_disponibili,
            default=[colonne_disponibili[0]] if colonne_disponibili else [],
            help="Puoi selezionare più colonne se il tuo file contiene date in colonne diverse"
        )
        
        if not colonne_selezionate:
            st.warning("⚠️ Seleziona almeno una colonna da normalizzare per continuare.")
            st.stop()
        
        st.write(f"**Colonne selezionate per la normalizzazione:** {', '.join(colonne_selezionate)}")
        
        # Se abbiamo più colonne selezionate, permettiamo di scegliere quale usare per l'ordinamento
        colonna_ordinamento = colonne_selezionate[0]  # Default alla prima colonna
        if len(colonne_selezionate) > 1 and ordina_date:
            st.write("### Opzioni di ordinamento")
            colonna_ordinamento = st.selectbox(
                "Seleziona la colonna da usare per l'ordinamento cronologico:",
                options=colonne_selezionate,
                index=0,
                help="Quando hai selezionato più colonne, scegli quale usare come riferimento per l'ordinamento"
            )
            st.write(f"**Ordinamento basato su:** '{colonna_ordinamento}'")
        
        # Mostriamo alcune date prima della normalizzazione per ogni colonna selezionata
        for colonna in colonne_selezionate:
            if colonna in df.columns:
                st.write(f"**Esempi di date nella colonna '{colonna}' prima della normalizzazione:**")
                st.write(df[colonna].head().tolist())
            else:
                st.warning(f"⚠️ Colonna '{colonna}' non trovata nel foglio selezionato!")
        
        # Variabili per raccogliere tutti i risultati
        tutti_df_elaborati = {}
        tutte_statistiche = {}
        tutti_df_temp = {}
        
        if elabora_tutti_fogli:
            # Elaboriamo tutti i fogli
            st.write("## 🔄 Elaborazione di tutti i fogli")
            for nome_foglio in fogli_disponibili:
                st.write(f"### 📄 Elaborazione foglio: {nome_foglio}")
                try:
                    df_foglio = pd.read_excel(file, sheet_name=nome_foglio, header=0)
                    
                    # Controlliamo se le colonne selezionate esistono in questo foglio
                    colonne_esistenti = [col for col in colonne_selezionate if col in df_foglio.columns]
                    colonne_mancanti = [col for col in colonne_selezionate if col not in df_foglio.columns]
                    
                    if colonne_mancanti:
                        st.warning(f"⚠️ Nel foglio '{nome_foglio}' mancano le colonne: {', '.join(colonne_mancanti)}")
                    
                    if colonne_esistenti:
                        # Elaboriamo solo le colonne che esistono
                        colonna_ord_foglio = colonna_ordinamento if colonna_ordinamento in colonne_esistenti else colonne_esistenti[0]
                        
                        df_elaborato, stats, df_temp = elabora_foglio(
                            df_foglio, colonne_esistenti, colonna_ord_foglio, 
                            ordina_date, formato_output, formati_output, nome_foglio
                        )
                        
                        tutti_df_elaborati[nome_foglio] = df_elaborato
                        tutte_statistiche.update({f"{k}_{nome_foglio}": v for k, v in stats.items()})
                        tutti_df_temp[nome_foglio] = df_temp
                        
                        # Mostriamo alcune date dopo la normalizzazione
                        for colonna in colonne_esistenti:
                            st.write(f"**Esempi di date nella colonna '{colonna}' dopo la normalizzazione (Foglio '{nome_foglio}'):**")
                            st.write(df_elaborato[colonna].head().tolist())
                    else:
                        st.error(f"❌ Nessuna delle colonne selezionate trovata nel foglio '{nome_foglio}'")
                        
                except Exception as e:
                    st.error(f"❌ Errore nell'elaborazione del foglio '{nome_foglio}': {e}")
            
            # Il dataframe principale sarà quello del primo foglio (per compatibilità)
            if tutti_df_elaborati:
                df = tutti_df_elaborati[list(tutti_df_elaborati.keys())[0]]
                statistiche_conversione = tutte_statistiche
                df_temp = tutti_df_temp[list(tutti_df_temp.keys())[0]]
            else:
                st.error("❌ Nessun foglio è stato elaborato con successo!")
                st.stop()
                
        else:
            # Elaboriamo solo il foglio selezionato
            df, statistiche_conversione, df_temp = elabora_foglio(
                df, colonne_selezionate, colonna_ordinamento, 
                ordina_date, formato_output, formati_output
            )
            
            # Mostriamo alcune date dopo la normalizzazione per ogni colonna
            for colonna in colonne_selezionate:
                if colonna in df.columns:
                    st.write(f"**Esempi di date nella colonna '{colonna}' dopo la normalizzazione:**")
                    st.write(df[colonna].head().tolist())
        
        # Visualizziamo il dataframe modificato
        st.write("Anteprima del file con date normalizzate:")
        st.dataframe(df.head(10))
        
        # Statistiche sulle colonne delle date
        with st.expander("Statistiche delle colonne normalizzate"):
            # Mostra un riepilogo generale
            st.write("### Riepilogo conversioni")
            
            if elabora_tutti_fogli:
                # Raggruppiamo le statistiche per foglio
                fogli_stats = {}
                for key, stats in statistiche_conversione.items():
                    # La chiave è nel formato "colonna_nomefoglio"
                    if '_' in key:
                        colonna = key.rsplit('_', 1)[0]
                        foglio = stats.get('foglio', 'Sconosciuto')
                    else:
                        colonna = key
                        foglio = stats.get('foglio', 'Principale')
                    
                    if foglio not in fogli_stats:
                        fogli_stats[foglio] = {}
                    fogli_stats[foglio][colonna] = stats
                
                # Mostriamo le statistiche per ogni foglio
                for foglio, stats_foglio in fogli_stats.items():
                    st.write(f"#### 📄 Foglio: {foglio}")
                    for colonna, stats in stats_foglio.items():
                        col1, col2, col3 = st.columns(3)
                        with col1:
                            st.metric(f"Colonna: {colonna}", f"{stats['convertiti']}/{stats['totali']}")
                        with col2:
                            st.metric("Percentuale successo", f"{stats['percentuale']:.1f}%")
                        with col3:
                            if stats['percentuale'] == 100:
                                st.success("✅ Completato")
                            elif stats['percentuale'] >= 80:
                                st.warning("⚠️ Parziale")
                            else:
                                st.error("❌ Problemi")
            else:
                # Singolo foglio - visualizzazione normale
                for colonna, stats in statistiche_conversione.items():
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric(f"Colonna: {colonna}", f"{stats['convertiti']}/{stats['totali']}")
                    with col2:
                        st.metric("Percentuale successo", f"{stats['percentuale']:.1f}%")
                    with col3:
                        if stats['percentuale'] == 100:
                            st.success("✅ Completato")
                        elif stats['percentuale'] >= 80:
                            st.warning("⚠️ Parziale")
                        else:
                            st.error("❌ Problemi")
            
            # Statistiche dettagliate per ogni colonna con date valide
            if elabora_tutti_fogli:
                for nome_foglio, df_temp_foglio in tutti_df_temp.items():
                    st.write(f"### 📄 Statistiche temporali - Foglio '{nome_foglio}'")
                    for colonna in colonne_selezionate:
                        if f'_data_oggetto_{colonna}' in df_temp_foglio.columns:
                            date_valide = df_temp_foglio[pd.notna(df_temp_foglio[f'_data_oggetto_{colonna}'])][f'_data_oggetto_{colonna}']
                            
                            if not date_valide.empty:
                                st.write(f"#### Colonna '{colonna}'")
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
                                st.write(f"#### Colonna '{colonna}': Non ci sono date valide per calcolare le statistiche.")
            else:
                # Singolo foglio
                for colonna in colonne_selezionate:
                    if f'_data_oggetto_{colonna}' in df_temp.columns:
                        date_valide = df_temp[pd.notna(df_temp[f'_data_oggetto_{colonna}'])][f'_data_oggetto_{colonna}']
                        
                        if not date_valide.empty:
                            st.write(f"### Statistiche temporali - Colonna '{colonna}'")
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
                            st.write(f"### Colonna '{colonna}': Non ci sono date valide per calcolare le statistiche.")
        
        # Opzione per scaricare file con errori (se ci sono)
        if elabora_tutti_fogli:
            # Per più fogli, controlliamo tutti i fogli
            fogli_con_errori = {}
            for nome_foglio, df_temp_foglio in tutti_df_temp.items():
                colonne_errori_foglio = []
                for colonna in colonne_selezionate:
                    if f'_data_oggetto_{colonna}' in df_temp_foglio.columns:
                        errori_colonna = df_temp_foglio[df_temp_foglio[f'_data_oggetto_{colonna}'].isna()]
                        if not errori_colonna.empty:
                            colonne_errori_foglio.append(colonna)
                
                if colonne_errori_foglio:
                    fogli_con_errori[nome_foglio] = colonne_errori_foglio
            
            if fogli_con_errori:
                with st.expander(f"📥 Scarica file con date problematiche (Più fogli)"):
                    for nome_foglio, colonne_errori in fogli_con_errori.items():
                        st.write(f"**Foglio '{nome_foglio}'**: Problemi nelle colonne {', '.join(colonne_errori)}")
                    
                    # Creiamo un file Excel con tutti i fogli che hanno errori
                    output_errori = io.BytesIO()
                    with pd.ExcelWriter(output_errori, engine='xlsxwriter') as writer:
                        for nome_foglio, colonne_errori in fogli_con_errori.items():
                            df_temp_foglio = tutti_df_temp[nome_foglio]
                            df_elaborato_foglio = tutti_df_elaborati[nome_foglio]
                            
                            # Creiamo una maschera per le righe con errori
                            mask_errori = pd.Series([False] * len(df_temp_foglio))
                            for colonna in colonne_errori:
                                if f'_data_oggetto_{colonna}' in df_temp_foglio.columns:
                                    mask_errori |= df_temp_foglio[f'_data_oggetto_{colonna}'].isna()
                            
                            df_errori_foglio = df_elaborato_foglio[mask_errori].copy()
                            if not df_errori_foglio.empty:
                                # Nome foglio limitato a 31 caratteri per Excel
                                nome_sheet = f"Errori_{nome_foglio}"[:31]
                                df_errori_foglio.to_excel(writer, sheet_name=nome_sheet, index=False)
                    
                    st.download_button(
                        label="📋 Scarica fogli con date problematiche",
                        data=output_errori.getvalue(),
                        file_name="date_problematiche_multifogli.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
        else:
            # Singolo foglio - logica originale
            colonne_errori = []
            for colonna in colonne_selezionate:
                if f'_data_oggetto_{colonna}' in df_temp.columns:
                    errori_colonna = df_temp[df_temp[f'_data_oggetto_{colonna}'].isna()]
                    if not errori_colonna.empty:
                        colonne_errori.append(colonna)
            
            if colonne_errori:
                with st.expander(f"📥 Scarica file con date problematiche"):
                    st.write(f"Sono state trovate date problematiche in {len(colonne_errori)} colonna/e: {', '.join(colonne_errori)}")
                    
                    # Creiamo un file con solo le righe problematiche
                    mask_errori = pd.Series([False] * len(df_temp))
                    for colonna in colonne_errori:
                        if f'_data_oggetto_{colonna}' in df_temp.columns:
                            mask_errori |= df_temp[f'_data_oggetto_{colonna}'].isna()
                    
                    df_errori = df[mask_errori].copy()
                    
                    if not df_errori.empty:
                        st.write(f"Numero di righe con problemi: {len(df_errori)}")
                        st.dataframe(df_errori.head())
                        
                        output_errori = io.BytesIO()
                        with pd.ExcelWriter(output_errori, engine='xlsxwriter') as writer:
                            df_errori.to_excel(writer, index=False)
                        
                        st.download_button(
                            label="📋 Scarica righe con date problematiche",
                            data=output_errori.getvalue(),
                            file_name="date_problematiche_dettagliate.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
        
        # Opzione per scaricare il file modificato
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            
            if elabora_tutti_fogli:
                # Scriviamo tutti i fogli elaborati
                for nome_foglio, df_elaborato in tutti_df_elaborati.items():
                    df_export = df_elaborato.copy()
                    df_temp_foglio = tutti_df_temp[nome_foglio]
                    
                    # Convertiamo le colonne selezionate in oggetti datetime per Excel
                    for colonna in colonne_selezionate:
                        if colonna in df_export.columns and f'_data_oggetto_{colonna}' in df_temp_foglio.columns:
                            # Utilizziamo gli oggetti datetime che abbiamo già generato
                            df_export[colonna] = df_temp_foglio[f'_data_oggetto_{colonna}']
                        elif colonna in df_export.columns:
                            # Tentiamo di convertire le stringhe di data in datetime
                            try:
                                df_export[colonna] = pd.to_datetime(df_export[colonna], dayfirst=True)
                            except:
                                pass  # Manteniamo il formato originale se la conversione fallisce
                    
                    # Salviamo il foglio con le date in formato Excel nativo
                    # Nome foglio limitato a 31 caratteri per Excel
                    nome_sheet = nome_foglio[:31]
                    df_export.to_excel(writer, sheet_name=nome_sheet, index=False)
                    
                    # Otteniamo un riferimento al foglio di lavoro
                    worksheet = writer.sheets[nome_sheet]
                    
                    # Formattazione specifica per le date nelle colonne selezionate
                    date_format = writer.book.add_format({'num_format': 'dd/mm/yyyy'})
                    
                    # Applichiamo la formattazione alle colonne contenenti date
                    for i, colonna in enumerate(df_export.columns):
                        if colonna in colonne_selezionate:
                            worksheet.set_column(i, i, 15, date_format)
            else:
                # Singolo foglio - logica originale
                df_export = df.copy()
                
                # Convertiamo le colonne selezionate in oggetti datetime per Excel
                for colonna in colonne_selezionate:
                    if f'_data_oggetto_{colonna}' in df_temp:
                        # Utilizziamo gli oggetti datetime che abbiamo già generato
                        df_export[colonna] = df_temp[f'_data_oggetto_{colonna}']
                    else:
                        # Tentiamo di convertire le stringhe di data in datetime
                        try:
                            df_export[colonna] = pd.to_datetime(df_export[colonna], dayfirst=True)
                        except:
                            pass  # Manteniamo il formato originale se la conversione fallisce
                    
                # Salviamo il file con le date in formato Excel nativo
                df_export.to_excel(writer, index=False)
                
                # Otteniamo un riferimento al foglio di lavoro
                worksheet = writer.sheets['Sheet1']
                
                # Formattazione specifica per le date nelle colonne selezionate
                date_format = writer.book.add_format({'num_format': 'dd/mm/yyyy'})
                
                # Applichiamo la formattazione alle colonne contenenti date
                for i, colonna in enumerate(df_export.columns):
                    if colonna in colonne_selezionate:
                        worksheet.set_column(i, i, 15, date_format)
        
        # Informazioni sul download
        st.write("### 📥 Download File Normalizzato")
        col1, col2 = st.columns([2, 1])
        
        with col1:
            st.write(f"**Colonne normalizzate:** {', '.join(colonne_selezionate)}")
            st.write(f"**Formato date:** {formato_output}")
            if elabora_tutti_fogli:
                st.write(f"**Fogli elaborati:** {len(tutti_df_elaborati)} ({', '.join(tutti_df_elaborati.keys())})")
            else:
                st.write(f"**Foglio elaborato:** {foglio_selezionato}")
            if ordina_date and colonna_ordinamento:
                st.write(f"**Ordinamento:** Per data (colonna '{colonna_ordinamento}')")
        
        with col2:
            # Calcola il tasso di successo totale
            successo_totale = sum([stats['convertiti'] for stats in statistiche_conversione.values()])
            righe_totali = sum([stats['totali'] for stats in statistiche_conversione.values()])
            tasso_successo = (successo_totale / righe_totali * 100) if righe_totali > 0 else 0
            
            if tasso_successo == 100:
                st.success(f"✅ {tasso_successo:.0f}% successo")
            elif tasso_successo >= 80:
                st.warning(f"⚠️ {tasso_successo:.1f}% successo")
            else:
                st.error(f"❌ {tasso_successo:.1f}% successo")
        
        # Nome file personalizzato
        nome_file = "date_normalizzate_multifogli.xlsx" if elabora_tutti_fogli else "date_normalizzate.xlsx"
        
        st.download_button(
            label="📊 Scarica Excel con date normalizzate",
            data=output.getvalue(),
            file_name=nome_file,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        # Aggiunge una nota informativa
        if elabora_tutti_fogli:
            st.info(f"✅ Le date nelle colonne selezionate sono state normalizzate in tutti i {len(tutti_df_elaborati)} fogli elaborati e possono ora essere utilizzate per l'ordinamento cronologico.")
        else:
            st.info("✅ Le date nelle colonne selezionate sono state normalizzate e possono ora essere utilizzate per l'ordinamento cronologico.")
        
        # Aggiungiamo un footer con istruzioni
        st.markdown("---")
        st.markdown("### 📋 Suggerimenti per l'utilizzo:")
        st.markdown("1. **Selezione fogli**: Se il file Excel ha più fogli, puoi scegliere quale elaborare o elaborarli tutti.")
        st.markdown("2. **Selezione colonne**: Puoi selezionare una o più colonne contenenti date da normalizzare.")
        st.markdown("3. **Ordinamento**: I dati vengono ordinati cronologicamente in base alla colonna selezionata.")
        st.markdown("4. **Formato**: Puoi cambiare il formato di visualizzazione delle date nella barra laterale.")
        st.markdown("5. **Controllo qualità**: Controlla le statistiche e i 'valori problematici' per verificare la qualità della conversione.")
        st.markdown("6. **Export multi-foglio**: Quando elabori tutti i fogli, il file Excel scaricato conterrà tutti i fogli normalizzati.")
        st.markdown("7. **Compatibilità**: Assicurati che le colonne selezionate esistano in tutti i fogli se usi l'opzione 'Elabora tutti i fogli'.")
        
    except Exception as e:
        st.error(f"❌ Si è verificato un errore durante l'elaborazione del file: {e}")
        st.error("Dettagli dell'errore per il debug:")
        st.exception(e)