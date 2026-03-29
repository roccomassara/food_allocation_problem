# Assegnamento ottimo di prodotti alimentari. 

Questo repository contiene la tesi in formato pdf della mia laurea in Informatica presso l'Università degli Studi di Firenze, il progetto python - pyomo riguardante il modello di ottimizzazione e i relativi file di output riguardante la sperimentazione tramite solver Gurobi 12.0.3.

## Descrizione contenuto package

Nella directory Food allocation problem è presente:
- Il modello di ottimizzazione in linguaggio python - pyomo (OptimizationModel.py).
- 5 file .xlsx contenenti gli input per il modello, con i dati su prodotti e strutture relative alle 5 settimane prese in esame, nel periodo che va dal 25/08/2025 al 26/09/2025.
- 5 directory suddivise in base a tali settimane e contenenti ciascuna 3 tipi di file di output:
  - assegnazione.xlsx xontiene la suddivisione in colli di merce di ciascuna partita alle relative strutture caritative.
  - output_modello.xlsx rappresenta il file excel riguardante l'output che richiede l'azienda per il corretto inserimento massivo dei dati riguardanti ciascun ordine nel database aziendale.
  - output_modello.txt contenente i risultati delle sperimentazioni, con i relativi errori sulle variabili di scarto della funzione obiettivo e mostra, in modo più approfondito, il risultato della suddivisione di prodotti per ciascuna struttura caritativa.

Al di fuori della directory sono presenti i file in formato pdf della tesi di laurea triennale e il relativo abstract. 

## Descrizione della simulazione

### Raccomandazione
ATTENZIONE: prima di avviare l'esecuzione di OptimizationModel.py assicurarsi di aver installato tutte le librerie presenti nel preambolo del programma: 
- import pyomo.environ as pyo
- import pandas as pd
- import datetime
- import os
- import math
- import openpyxl
- from collections import defaultdict
- from itertools import combinations
- from dateutil.relativedelta import relativedelta
- from datetime import date
- import time
  
Inoltre, raccomandiamo l'utilizzo di gurobi 12.0.3 come solver per l'esecuzione del programma, come da utilizzo per i test qui presenti.

### Simulazione
 
## Bibliografia

Tutto il materiale di terze parti che è stato utilizzato, è stato citato nella bibliografia presente al termine del pdf della tesi.
