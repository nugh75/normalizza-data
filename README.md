# 📊 Normalizzazione Date in Excel

Una applicazione Streamlit per normalizzare e convertire date in file Excel, supportando file multi-foglio e selezione multipla di colonne.

## ✨ Caratteristiche Principali

### 🎯 Funzionalità Core
- **Normalizzazione intelligente delle date** con riconoscimento automatico di oltre 15 formati
- **Selezione multipla di colonne** per elaborare più colonne contemporaneamente
- **Gestione file Excel multi-foglio** con opzione di elaborazione singola o completa
- **Ordinamento cronologico** automatico basato sulla colonna selezionata
- **Formati di output personalizzabili** (gg-mm-aaaa, gg/mm/aaaa, aaaa-mm-gg)

### 🔧 Formati Date Supportati

L'applicazione riconosce automaticamente questi formati:

**Formati numerici:**
- `YYYY-MM-DD`, `DD/MM/YYYY`, `MM/DD/YYYY`
- `DD-MM-YYYY`, `MM-DD-YYYY`, `YYYY/MM/DD`
- `DD.MM.YYYY`, `MM.DD.YYYY`, `YYYY.MM.DD`
- `YYYYMMDD` (formato compatto)

**Formati con testo:**
- `DD MMM YYYY`, `DD MMMM YYYY` (es. "12 gen 2023", "12 gennaio 2023")
- `MMM DD, YYYY`, `MMMM DD, YYYY` (es. "Jan 12, 2023")
- `DD-MMM-YYYY`, `DD-MMMM-YYYY`

**Formati con giorni della settimana:**
- `DDD, DD MMM YYYY` (es. "Lun, 12 gen 2023")
- `DDDD, DD MMMM YYYY` (es. "Lunedì, 12 gennaio 2023")

**Formati speciali:**
- Timestamp Excel (numeri seriali)
- Timestamp UNIX (secondi da epoch)
- Oggetti Pandas Timestamp e Python datetime

### 📋 Lingue Supportate

- **Italiano**: gennaio, febbraio, marzo, ecc. + lunedì, martedì, ecc.
- **Inglese**: January, February, March, ecc. + Monday, Tuesday, ecc.

## 🚀 Installazione

### Prerequisiti
- Python 3.7 o superiore

### Dipendenze
Installa le dipendenze richieste:

```bash
pip install -r requirements.txt
```

Le dipendenze includono:
- `streamlit>=1.21.0` - Framework web per l'interfaccia
- `pandas>=1.5.0` - Manipolazione e analisi dati
- `numpy>=1.20.0` - Calcoli numerici
- `openpyxl>=3.0.0` - Lettura file Excel
- `xlsxwriter>=3.0.0` - Scrittura file Excel con formattazione
- `python-dateutil>=2.8.2` - Parsing avanzato delle date

## 📖 Come Utilizzare

### 1. Avviare l'Applicazione

```bash
streamlit run normalizza_date.py
```

L'applicazione sarà disponibile su `http://localhost:8501`

### 2. Interfaccia Utente

#### Barra Laterale - Opzioni Configurazione
- **Ordina per data**: Abilita/disabilita l'ordinamento cronologico
- **Formato visualizzazione**: Scegli tra `gg-mm-aaaa`, `gg/mm/aaaa`, `aaaa-mm-gg`

#### Area Principale - Workflow di Elaborazione

1. **📁 Upload File Excel**
   - Carica file `.xlsx` o `.xls`
   - Supporto per file multi-foglio

2. **📋 Selezione Foglio**
   - Se il file ha un solo foglio: selezione automatica
   - Se il file ha più fogli: selezione manuale o opzione "elabora tutti"

3. **🎯 Selezione Colonne**
   - Widget multiselect per selezionare una o più colonne
   - Anteprima delle date presenti nelle colonne selezionate

4. **⚙️ Opzioni Ordinamento** (se multiple colonne)
   - Selezione della colonna di riferimento per l'ordinamento cronologico

### 3. Elaborazione e Risultati

#### Statistiche di Conversione
Per ogni colonna e foglio elaborato vengono mostrate:
- ✅ Numero di date convertite con successo
- ⚠️ Percentuale di successo
- ❌ Valori problematici con dettagli

#### Indicatori Visivi
- **✅ Verde**: 100% successo
- **⚠️ Giallo**: 80-99% successo  
- **❌ Rosso**: <80% successo

#### Gestione Errori
- **Valori problematici**: Visualizzazione espandibile dei record non convertiti
- **Download errori**: File Excel separato con solo le righe problematiche
- **Report dettagliato**: Numero riga originale per facilitare correzioni manuali

### 4. Export e Download

#### File Normalizzato Principale
- **Foglio singolo**: Excel con colonne normalizzate e ordinate
- **Multi-foglio**: Excel con tutti i fogli elaborati mantenendo la struttura originale

#### File Date Problematiche (opzionale)
- Contiene solo le righe con date non riconosciute
- Mantiene riferimenti alle righe originali
- Disponibile sia per foglio singolo che multi-foglio

#### Formattazione Excel
- Le date vengono salvate come oggetti DateTime nativi di Excel
- Formattazione automatica `dd/mm/yyyy` per le colonne di date
- Larghezza colonne ottimizzata per la visualizzazione

## 💡 Funzionalità Avanzate

### Gestione Multi-Foglio
```python
# Logica interna per gestione multi-foglio
if elabora_tutti_fogli:
    for nome_foglio in fogli_disponibili:
        df_foglio = pd.read_excel(file, sheet_name=nome_foglio)
        df_elaborato, stats, df_temp = elabora_foglio(df_foglio, ...)
        tutti_df_elaborati[nome_foglio] = df_elaborato
```

### Normalizzazione Intelligente
```python
def normalizza_data(data, solo_formato=False):
    # Supporta oltre 15 formati diversi
    # Gestisce date italiane e inglesi
    # Fallback intelligente con dateutil.parser
    # Conversione timestamp Excel e UNIX
```

### Controllo Qualità
- **Validazione pre-elaborazione**: Verifica esistenza colonne in tutti i fogli
- **Statistiche dettagliate**: Conteggi e percentuali per ogni colonna/foglio
- **Report errori**: Identificazione precisa dei valori problematici

## 🔍 Esempi d'Uso

### Caso 1: File con Singolo Foglio
```
Input: file.xlsx (1 foglio, colonna "Data")
- Selezione automatica foglio
- Selezione colonna "Data"  
- Elaborazione e download
```

### Caso 2: File Multi-Foglio, Singola Colonna
```
Input: report.xlsx (5 fogli, colonna "Data Evento")
- Selezione "Elabora tutti i fogli"
- Selezione colonna "Data Evento"
- Elaborazione di tutti i 5 fogli
- Download con 5 fogli normalizzati
```

### Caso 3: File Multi-Colonna
```
Input: anagrafe.xlsx (colonne "Data Nascita", "Data Assunzione")
- Selezione multiple colonne
- Scelta colonna ordinamento ("Data Nascita")
- Statistiche separate per ogni colonna
```

## 🐛 Troubleshooting

### Problemi Comuni

**"Colonna non trovata nel foglio"**
- Verifica che la colonna esista in tutti i fogli selezionati
- Controlla maiuscole/minuscole e spazi nel nome colonna

**"Bassa percentuale di conversione"**
- Controlla il formato delle date nel file originale
- Verifica la presenza di celle vuote o testo non-data
- Consulta la sezione "Valori problematici" per dettagli

**"Errore durante l'elaborazione"**
- Verifica che il file Excel non sia corrotto
- Assicurati che le colonne selezionate contengano effettivamente date
- Controlla la console per messaggi di debug dettagliati

**"UserWarning: Boolean Series key will be reindexed"**
- Questo warning è stato risolto nelle versioni recenti
- Se persiste, aggiorna alle ultime dipendenze con `pip install -r requirements.txt --upgrade`

### Limitazioni

- **Fogli Excel**: Nomi limitati a 31 caratteri (limitazione Excel)
- **Dimensioni file**: Dipendente dalla memoria RAM disponibile
- **Formati date**: Solo formati gregoriani supportati

## 🔧 Sviluppo e Contributi

### Struttura del Codice

```
normalizza_date.py
├── normalizza_data()        # Funzione core normalizzazione
├── elabora_foglio()         # Elaborazione singolo foglio
├── Interfaccia Streamlit    # UI e workflow
└── Gestione Export          # Download e formattazione
```

### Estensioni Possibili

1. **Supporto CSV**: Aggiungere lettura file CSV
2. **Formati aggiuntivi**: Supporto calendari non-gregoriani
3. **API REST**: Interfaccia programmatica
4. **Batch processing**: Elaborazione multipli file
5. **Configurazione avanzata**: Template personalizzati

## 📄 Licenza

Questo progetto è disponibile sotto licenza MIT. Vedi il file LICENSE per dettagli.

## 🤝 Supporto

Per problemi, suggerimenti o richieste di funzionalità:
- Apri una issue su GitHub
- Controlla la sezione Troubleshooting
- Verifica che tu stia usando l'ultima versione delle dipendenze

---

**Versione**: 2.0  
**Ultimo aggiornamento**: 2024  
**Compatibilità**: Python 3.7+, Streamlit 1.21+
