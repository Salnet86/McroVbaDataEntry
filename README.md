# McroVbaDataEntry
         Tecnico informatico 
# Progetto di Automazione del Processo di Data Entry con VBA e Gestione Calendari

## Descrizione del Progetto

Durante il mio tirocinio, ho avuto l'opportunità di automatizzare diversi processi, tra cui il **data entry** per i corsi di formazione e la gestione dei **calendari di appuntamenti**, utilizzando **Macro VBA** in **Microsoft Excel**. Inoltre, ho integrato il sistema con un **database SQL Server** per archiviare i dati in modo strutturato e facilmente accessibile.

## Funzionalità Implementate

### 1. **Automazione del Data Entry per i Corsi di Formazione (VBA)**

- **UserForm per l'inserimento dei dati**: Ho creato una UserForm che semplifica l'inserimento di dati per gli studenti e i corsi di formazione. I dati inseriti vengono salvati automaticamente in un foglio Excel, riducendo errori manuali e aumentando l'efficienza.
  
  Dati raccolti tramite la UserForm:
  - Nome
  - Cognome
  - Data di Nascita
  - Età (calcolata automaticamente)
  - Codice Fiscale
  - Numero del Corso
  - Anno di Iscrizione
  - Tipo di Corso
  - Tipo di Istruzione

- **Salvataggio e gestione dei dati**: Ho sviluppato una logica per salvare i dati in un formato strutturato in un foglio Excel (in colonne specifiche) per facilitarne l'archiviazione e l'analisi.

- **Funzionalità di calcolo dell'età**: L'età degli studenti viene calcolata automaticamente sulla base della loro data di nascita.

### 2. **Gestione Calendari di Appuntamenti**

- **Data Entry in Excel**: Ho creato una soluzione per gestire il calendario degli appuntamenti utilizzando **Microsoft Excel**. Ogni appuntamento è stato inserito in un formato strutturato (data, ora, descrizione, ecc.), semplificando la visualizzazione e la gestione degli appuntamenti.

- **Automazione dell'Inserimento dei Dati**: Utilizzando **VBA**, ho creato macro che permettono di inserire facilmente nuovi appuntamenti, aggiornare gli esistenti e gestire le modifiche, riducendo la necessità di input manuali.

### 3. **Integrazione con SQL Server per l'Archiviazione dei Dati**

- **Connessione con SQL Server**: Ho implementato un sistema per inserire e aggiornare i dati degli appuntamenti e dei corsi di formazione in un **database SQL Server**. I dati vengono trasferiti da Excel a SQL Server tramite query SQL utilizzando VBA.
  
- **Sincronizzazione dei Dati**: Dopo aver inserito i dati in Excel, i dati vengono trasferiti al database per garantire una gestione centralizzata e sicura. La sincronizzazione è stata automatizzata tramite VBA, riducendo il rischio di errori di duplicazione o incoerenza.

- **Esempio di query SQL**:
    ```sql
    INSERT INTO Appuntamenti (Data, Ora, Descrizione)
    VALUES ('2025-01-20', '14:00', 'Appuntamento di formazione');
    ```

### 4. **Funzionalità di Salvataggio del File**

- **Salvataggio automatico**: Il file Excel può essere salvato automaticamente nella stessa posizione o in una posizione personalizzata tramite una finestra di dialogo di salvataggio, in modo che l'utente possa scegliere dove memorizzare i dati.

- **Salvataggio nel formato desiderato**: È stata implementata una funzionalità che consente di esportare i dati in un formato Excel (.xlsx) o CSV per facilitare l'integrazione con altre piattaforme e applicazioni.

## Tecnologie Utilizzate

- **Microsoft Excel**: Per la gestione dei dati e l'implementazione della macro VBA.
- **VBA (Visual Basic for Applications)**: Per automatizzare l'inserimento dei dati, la gestione dei calendari e la sincronizzazione con SQL Server.
- **SQL Server**: Per archiviare e gestire i dati in modo sicuro e centralizzato.
- **SQL**: Per le query di inserimento, aggiornamento e gestione dei dati nel database.

## Benefici del Progetto

- **Miglioramento dell'Efficienza**: L'automazione del data entry e la gestione degli appuntamenti ha ridotto significativamente il tempo speso in attività manuali, aumentando l'efficienza operativa.
  
- **Gestione Centralizzata dei Dati**: Grazie all'integrazione con SQL Server, i dati sono ora archiviati in un sistema centralizzato, consentendo accesso sicuro e strutturato alle informazioni.

- **Riduzione degli Errori Manuali**: L'automazione riduce al minimo il rischio di errori nel processo di inserimento dei dati, migliorando la precisione delle informazioni.

- **Facilità di Accesso e Recupero Dati**: Con SQL Server, è più facile accedere, recuperare e analizzare i dati relativi agli appuntamenti e ai corsi, migliorando la gestione complessiva.

## Come Usare

1. **Data Entry per i Corsi**: Compila la UserForm con le informazioni richieste e clicca sul pulsante di inserimento per salvare i dati nel foglio Excel.
   
2. **Gestione Appuntamenti**: Inserisci i nuovi appuntamenti nel calendario Excel. Puoi aggiungere, aggiornare o rimuovere appuntamenti direttamente tramite l'interfaccia di Excel.

3. **Esportazione dei Dati**: Usa la funzionalità di salvataggio per esportare i dati nel formato desiderato (Excel o CSV).

4. **Sincronizzazione con SQL Server**: La macro VBA invierà automaticamente i dati a SQL Server, garantendo che il database sia aggiornato con i dati più recenti.

## Conclusione

Questo progetto ha rappresentato una preziosa esperienza di apprendimento, dove ho migliorato le mie competenze in **VBA**, **SQL Server**, e **automazione dei processi aziendali**. La gestione centralizzata dei dati e l'automazione dei processi di inserimento e sincronizzazione hanno migliorato l'efficienza operativa e la qualità dei dati per la gestione dei corsi di formazione e degli appuntamenti.


