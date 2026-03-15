# AGENTS.md

## Principi non negoziabili

1. Il repository deve rimanere **sempre coerente tra codice, documentazione e stato del progetto**.
2. Nessuna modifica può lasciare documentazione o stato fuori sincronizzazione.
3. Se viene individuato drift tra codice e documentazione, deve essere **corretto immediatamente o tracciato come task**.
4. Ogni problema osservato che richiede lavoro futuro o influisce sulla coerenza del repository deve essere:

   * corretto immediatamente, oppure
   * registrato come task in `PROJECT_STATUS.json`.
5. L’agente non deve inventare requisiti o comportamenti non presenti nei file di progetto.

---

# Standard repository richiesto

Il repository deve contenere questi file alla root:

```
AGENTS.md
PROJECT_SPEC.md
PROJECT_STATUS.json
README.md
```

Questi file costituiscono **l'interfaccia standard tra repository e agenti automatizzati**.

---

# Inizializzazione repository

Se uno o più file standard sono **mancanti**, l’agente deve crearli.

## File da creare se mancanti

```
AGENTS.md
PROJECT_SPEC.md
PROJECT_STATUS.json
README.md
```

### PROJECT_SPEC.md minimo

Deve contenere almeno:

```
# Project Specification

## Goal
Descrizione sintetica dello scopo del progetto.

## Scope
Funzionalità incluse.

## Non Scope
Funzionalità escluse.

## Architecture
Descrizione dell'architettura principale.

## Constraints
Vincoli tecnici.
```

---

### PROJECT_STATUS.json minimo

```json
{
  "treeview": [],
  "files_big": [],
  "tasks": []
}
```

---

### README.md minimo

```
# Project

Descrizione breve del progetto.

## Setup

## Run

## Documentation

- PROJECT_SPEC.md
- PROJECT_STATUS.json
```

---

# Migrazione da altri sistemi di agent

Se il repository contiene file provenienti da altri sistemi di agent, l’agente deve **normalizzare la struttura**.

## File equivalenti possibili

Esempi comuni:

```
docs/spec.md
docs/architecture.md
docs/design.md
SPEC.md
TODO.md
TASKS.md
STATUS.json
STATE.json
AGENT.md
AGENT_RULES.md
```

---

## Regole di migrazione

Se esistono file equivalenti:

### Spec

File equivalenti possibili:

```
docs/spec.md
SPEC.md
architecture.md
design.md
```

Devono essere:

* consolidati in `PROJECT_SPEC.md`
* aggiornati o fusi se necessario

---

### Status / Todo

File equivalenti possibili:

```
TODO.md
TASKS.md
STATUS.json
STATE.json
ROADMAP.md
```

Devono essere:

* convertiti nello schema `PROJECT_STATUS.json`

---

### Agent rules

File equivalenti possibili:

```
AGENT.md
AGENT_RULES.md
AGENT_POLICY.md
```

Devono essere:

* migrati o fusi in `AGENTS.md`

---

# Normalizzazione posizione file

Se i file standard esistono ma non sono nella root:

Esempi:

```
docs/PROJECT_SPEC.md
meta/PROJECT_STATUS.json
agents/AGENTS.md
```

L’agente deve:

1. spostarli nella root del repository
2. aggiornare eventuali riferimenti nei file

---

# Gerarchia di verità

In caso di conflitto tra fonti:

1. **Repository reale / comportamento osservabile**
2. `PROJECT_SPEC.md`
3. `PROJECT_STATUS.json`
4. `README.md`
5. deduzioni non esplicitate

Il repository reale ha sempre priorità sulla documentazione.

---

# File operativi obbligatori

## PROJECT_STATUS.json

Fonte operativa del progetto.

Contiene esclusivamente:

* task pendenti
* task in progresso
* struttura repository
* file monitorati

Non contiene:

* storico task
* log narrativi
* changelog

---

## PROJECT_SPEC.md

Documento tecnico del progetto.

Contiene:

* obiettivi
* architettura
* comportamento atteso
* vincoli

Non contiene task.

---

## README.md

Documento orientato agli utenti del repository.

Contiene:

* descrizione
* setup
* run
* stack
* link alla documentazione

---

# Ordine di lettura obbligatorio

Prima di eseguire qualsiasi task, l’agente deve leggere:

1. `README.md`
2. `PROJECT_SPEC.md`
3. `PROJECT_STATUS.json`
4. repository reale

---

# Modalità di esecuzione

## Full Sync Mode

Usare quando:

* architettura cambia
* refactor
* documentazione
* struttura repository

Operazioni obbligatorie:

* ricostruzione treeview
* controllo file grandi
* verifica drift globale

---

## Light Mode

Usare per modifiche locali.

Operazioni obbligatorie:

* verifica area modificata
* controllo drift locale

---

# Struttura PROJECT_STATUS.json

Schema obbligatorio.

```json
{
  "treeview": [],
  "files_big": [],
  "tasks": []
}
```

---

## treeview

Struttura reale del repository.

Deve essere aggiornata se:

* vengono aggiunti file
* file rimossi
* file spostati

---

## files_big

File ≥800 linee.

Formato:

```
"path/file.ext:1234"
```

File ≥1500 linee devono essere valutati per split.

---

## tasks

Schema task:

```json
{
  "id": "task-name",
  "status": "pending | in_progress",
  "description": "descrizione",
  "files": ["file1"],
  "priority": 1
}
```

Regole:

* solo stati `pending` o `in_progress`
* task completati rimossi
* priorità minore = più importante

---

# Selezione task

1. continuare `in_progress`
2. altrimenti `pending` con priorità più alta
3. un solo task attivo

---

# Controllo repository reale

L’agente deve verificare:

* esistenza file dichiarati
* corrispondenza treeview
* file non documentati
* task riferiti a file inesistenti

---

# Controllo drift

Drift = incoerenza tra:

* codice
* documentazione
* stato progetto

Esempi:

* codice non documentato
* documentazione non implementata
* file non registrati
* task obsoleti

---

# Gestione drift

Quando trovato drift:

1. correggere subito se possibile
2. altrimenti creare task
3. aggiornare documentazione

Il drift non può essere ignorato.

---

# Regola file grandi

File ≥800 linee devono essere monitorati.

File ≥1500 linee devono essere valutati per refactor.

---

# Aggiornamento PROJECT_STATUS.json

Obbligatorio quando:

* cambia struttura repo
* cambiano task
* compaiono file grandi

---

# Regole chiusura task

Un task può essere rimosso solo se:

* lavoro completato
* repository coerente
* documentazione aggiornata
* nessun drift residuo

---

# Condizione chiusura prompt

Il prompt può essere chiuso solo se:

* task completato o avanzato
* repository coerente
* documentazione aggiornata
* PROJECT_STATUS.json aggiornato
* nessun drift non tracciato

---

# Formato risposta finale

Ogni risposta finale deve includere:

### Changes

modifiche effettuate

### Alignment

* file aggiornati
* drift rilevati
* drift corretti
* nuovi task creati

### Repository Status

* task attivo
* task rimanenti
* blocchi

### Notes

eventuali osservazioni

---

# Regola finale

L’agente deve mantenere il repository **sempre consistente, verificabile e documentato**, garantendo che i file standard siano presenti, corretti e aggiornati.
