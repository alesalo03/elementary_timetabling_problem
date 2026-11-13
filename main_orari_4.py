import pylightxl as xl
from ortools.sat.python import cp_model
import textwrap
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.cell.rich_text import TextBlock, CellRichText
from openpyxl.cell.text import InlineFont
import psutil
from pathlib import Path
import os

#COSTANTI####################################################
#Materie
ITA=0;STO=1;GEO=2;ART=3;MUS=4;MAT=5;SCI=6;TEC=7;ING=8;MOT=9;ALT=10;REL=11;MEN=12;POT=13

MATERIE_STR = {
    "MOTORIA": MOT, "RELIGIONE": REL, "ALTERNATIVA": ALT,
    "ITALIANO": ITA, "STORIA": STO, "GEOGRAFIA": GEO, "ARTE": ART,
    "MUSICA": MUS, "MATEMATICA": MAT, "SCIENZE": SCI, "TECNOLOGIA": TEC,
    "INGLESE": ING, "MENSA": MEN,
}

#Giorni
LUN=0;MAR=1;MER=2;GIO=3;VEN=4
#Ore
O1=0;#8:30-9:00
O2=1;#9:00-9:30
O3=2;#9:30-10:00
O4=3;#10:00-10:30
O5=4;#10:30-11:00
O6=5;#11:00-11:30
O7=6;#11:30-12:00
O8=7;#12:00-12:30
O9=8;#12:30-13:00
O10=9;#13:00-13:30
O11=10;#13:30-14:00
O12=11;#14:00-14:30
O13=12;#14:30-15:00
O14=13;#15:00-15:30
O15=14;#15:30-16:00
O16=15;#16:00-16:30
#Indici struttura dati docente
INDEX_NOME=0
INDEX_GIORNI_DISPONIBILITA=1
INDEX_PART_TIME = 2

# converte una stringa oraria 'H:MM' nel corrispondente indice di mezz'ora (0-16)
def time_to_slot(t_str):
    t_str = t_str.replace(" ", "")
    h, m = map(int, t_str.split(':'))
    return int((h - 8) * 2 + (1 if m >= 30 else 0) - 1)

# Ogni tupla è un gruppo di indici di mezz'ora che rappresentano la STESSA ora di lezione (stessa variabile). 
def blocchi_orari(c, g):
    if c < classi[3][0] and g in (LUN, VEN):
        # classi 1,2,3 al lunedì/venerdì: 2 ore + mezz'ora isolata (O5) + 2 ore
        blocchi = [(0, 1), (2, 3), (O5,), (5, 6), (7, 8)]
        # il resto (fuori orario scolastico) tanto verrà fissato a 0: accorpiamo a due a due
        fuori = list(range(9, O))
        blocchi += [tuple(fuori[i:i+2]) for i in range(0, len(fuori), 2)]
    else:
        # caso generale: sempre coppie di mezz'ore = 1 ora
        blocchi = [tuple(range(o, min(o + 2, O))) for o in range(0, O, 2)]
    return blocchi

#MAIN########################################################

db = xl.readxl(fn="test_input_adaptive.xlsx")

classi = []
counter_classi = 0

for x in range(5):
    sezioni = db.ws(ws='test_input_adaptive').index(row=x+2, col=2)
    classi.append(sezioni.split(","))
    internal = 0
    for classe in classi[x]:
        classi[x][internal] = counter_classi
        internal=internal+1
        counter_classi = counter_classi+1

NOMI_CLASSI = [] # per trovare la classe partendo dall'indice
classe_index_by_nome = {} # per trovare l'indice partendo dalla classe
SEZIONI = ["A", "B", "C"]
anno=1
for gruppo_anno in classi:
    for indice_c in gruppo_anno:
        sezione = gruppo_anno.index(indice_c)
        classe_index_by_nome[f"{anno}{SEZIONI[sezione]}"] = indice_c
        NOMI_CLASSI.append(f"{anno}{SEZIONI[sezione]}")
    anno=anno+1



counter_docenti = 0
docenti = []

# Giorni della settimana in cui gli insegnanti sono disponibili
while True:
    nome_docente = db.ws(ws='test_input_adaptive').index(row=9+counter_docenti, col=1)
    # Se la cella del nome è vuota o contiene solo spazi, interrompe il ciclo
    if not nome_docente or str(nome_docente).strip() == '':
        break

    giorni_disponibilita = [] # Conterrà 5 liste (una per giorno), ognuna da 16 valori (0 o 1)
    
    for col_giorno in range(2, 7): # Colonne da 2 a 6 (Lunedì - Venerdì)
        valore_cella = str(db.ws(ws='test_input_adaptive').index(row=9+counter_docenti, col=col_giorno)).strip().upper()
        
        disponibilita_giorno = [0] * 16 # Inizializziamo il giorno a tutto 0 (non disponibile)
        
        if valore_cella == "SI":
            # Caso 1: Disponibile tutto il giorno
            disponibilita_giorno = [1] * 16
            
        elif "-" in valore_cella:
            # Caso 3: Fascia oraria (es. "8:30-11:30")
            try:
                start_time, end_time = valore_cella.split("-")
                start_slot = time_to_slot(start_time)
                end_slot = time_to_slot(end_time)
                
                # Riempiamo di 1 solo gli slot compresi in questa fascia
                for o in range(max(0, start_slot), min(16, end_slot)):
                    disponibilita_giorno[o] = 1
            except Exception as e:
                print(f"Errore nella lettura dell'orario '{valore_cella}' per {nome_docente}. Impostato a NO.")
        
        giorni_disponibilita.append(disponibilita_giorno)

    valore_cella = str(db.ws(ws='test_input_adaptive').index(row=9+counter_docenti, col=8)).strip().upper()
    if valore_cella == "SI":
        part_time = []
        part_time.append(1)
        valore_cella = str(db.ws(ws='test_input_adaptive').index(row=9+counter_docenti, col=9)).strip().upper()
        if not valore_cella or str(valore_cella).strip() == '':
            part_time.append(0)
        else:
            part_time.append(int(valore_cella))

        valore_cella = str(db.ws(ws='test_input_adaptive').index(row=9+counter_docenti, col=10)).strip().upper()
        if not valore_cella or str(valore_cella).strip() == '':
            part_time.append(0) #no preferenza espressa per il/i giorno/i libero/i
        else:
            matrice_giorni = [0, 0, 0, 0, 0]
            
            if "LUN" in valore_cella:
                matrice_giorni[0] = 1
            if "MAR" in valore_cella:
                matrice_giorni[1] = 1
            if "MER" in valore_cella:
                matrice_giorni[2] = 1
            if "GIO" in valore_cella:
                matrice_giorni[3] = 1
            if "VEN" in valore_cella:
                matrice_giorni[4] = 1
                
            part_time.append(matrice_giorni)

    else:
        part_time = [0]*3

    docenti.append([nome_docente, giorni_disponibilita, part_time])
    counter_docenti += 1
    
    
docente_classe_materia = []
docente_classe_potenziamento = []

# per individuare per ogni classe il punto di inizio all'interno del file .xlxs di input
inizi = []
start=counter_docenti+9
counter_spazi=0
while counter_spazi<2:
    cella_letta = db.ws(ws='test_input_adaptive').index(row=start, col=1)
    if not cella_letta or str(cella_letta).strip() == '':
        counter_spazi=counter_spazi+1
    else:
        if counter_spazi==1:
            counter_spazi=0
            inizi.append(start+1)
    start=start+1


for docente in docenti:
    classe = []
    classe_pot = []
    for i in inizi:
        counter = 0
        ore_materia = []
        while True:
            nome_docente = db.ws(ws='test_input_adaptive').index(row=i+counter, col=1)
                
            # Se la cella del nome è vuota o contiene solo spazi, interrompe il ciclo
            if not nome_docente or str(nome_docente).strip() == '':
                ore_materia = [0]*13
                classe_pot.append(0)
                break

            if nome_docente == docente[0]:
                for j in range(2, 15):
                    num_ore = db.ws(ws='test_input_adaptive').index(row=i+counter, col=j)
                    if not num_ore or str(num_ore).strip() == '':
                        num_ore = 0
                    ore_materia.append(num_ore)
                num_ore_pot = db.ws(ws='test_input_adaptive').index(row=i+counter, col=15)
                if not num_ore_pot or str(num_ore_pot).strip() == '':
                    num_ore_pot = 0
                classe_pot.append(num_ore_pot)
                break
            counter=counter+1

        classe.append(ore_materia)
    docente_classe_materia.append(classe)
    docente_classe_potenziamento.append(classe_pot)

# lettura orario di accesso alla palestra per ogni classe
disponibilita_palestra = []
riga_corrente = i + counter + 5 

while True:
    cl = db.ws(ws='test_input_adaptive').index(row=riga_corrente, col=1)    
    if not cl or str(cl).strip() == '':
        break

    palestra_per_classe = []
    
    for col_giorno in range(2, 7):
        accesso_palestra = db.ws(ws='test_input_adaptive').index(row=riga_corrente, col=col_giorno)
        
        disponibilita_giorno = [0] * 16
        
        if accesso_palestra and str(accesso_palestra).strip() != '':
            fasce = str(accesso_palestra).split(",")
            for f in fasce:
                try:
                    start_time, end_time = f.split("-")
                    start_slot = time_to_slot(start_time)
                    end_slot = time_to_slot(end_time)
                    for o in range(max(0, start_slot), min(16, end_slot)):
                        disponibilita_giorno[o] = 1
                except Exception as e:
                    print(f"Errore formato ora palestra per {cl}: {f}")
        
        palestra_per_classe.append(disponibilita_giorno)
            
    disponibilita_palestra.append(palestra_per_classe)
    
    riga_corrente += 1

# lettura classi che fanno materie insieme

abbinamenti_per_materia = []  # lista di tuple (materia, docente, [classe1, classe2])
docenti_condivisi = set() # lista di docenti che hanno materie condivise tra più classi
riga_corrente += 2  # salta l'intestazione della nuova tabella
while True:
    cl1 = db.ws(ws='test_input_adaptive').index(row=riga_corrente, col=1)
    if not cl1 or str(cl1).strip() == '':
        break

    cl2 = db.ws(ws='test_input_adaptive').index(row=riga_corrente, col=2)
    materie_raw = db.ws(ws='test_input_adaptive').index(row=riga_corrente, col=3)

    c1 = classe_index_by_nome[str(cl1).strip()]
    c2 = classe_index_by_nome[str(cl2).strip()]

    materia = MATERIE_STR.get(str(materie_raw).strip().upper())
    if materia is None:
        print(f"Materia sconosciuta '{str(materie_raw).strip()}' riga {riga_corrente}, ignorata.")
        continue
    
    docente_trovato = None
    for d in range(counter_docenti):
        if docente_classe_materia[d][c1][materia]>0 and docente_classe_materia[d][c2][materia]>0:
            docente_trovato = d
            break
    
    if docente_trovato is None:
        print(f"Nessun docente comune tra {cl1} e {cl2} per '{str(materie_raw).strip()}', riga {riga_corrente} ignorata.")
        continue
    
    docenti_condivisi.add(docente_trovato)
    abbinamenti_per_materia.append((materia, docente_trovato, [c1, c2]))
    
    # se le due classi fanno insieme la materia religione allora automaticamente fanno insieme anche alternativa
    if materia == REL:
        docente_alternativa = None
        for d in range(counter_docenti):
            if docente_classe_materia[d][c1][ALT]>0 and docente_classe_materia[d][c2][ALT]>0:
                docente_alternativa = d
                break
                
        if docente_alternativa is None:
            continue
        docenti_condivisi.add(docente_alternativa)
        abbinamenti_per_materia.append((ALT, docente_alternativa, [c1, c2]))
        

    riga_corrente += 1

classi_secondarie_mot = set(
    c2 for materia, d, (c1, c2) in abbinamenti_per_materia if materia == MOT
)

# lettura di potenziamenti da assegnare a specifiche materie

potenziamento_per_materia = []  # lista di tuple (materia, docente, docente_pot, classe)
riga_corrente += 2  # salta l'intestazione della nuova tabella
while True:
    cl = db.ws(ws='test_input_adaptive').index(row=riga_corrente, col=1)
    if not cl or str(cl).strip() == '':
        break

    materia_raw = db.ws(ws='test_input_adaptive').index(row=riga_corrente, col=2)
    docente_raw = db.ws(ws='test_input_adaptive').index(row=riga_corrente, col=3)

    c = classe_index_by_nome[str(cl).strip()]

    materia = MATERIE_STR.get(str(materia_raw).strip().upper())
    if materia is None:
        print(f"Materia sconosciuta '{str(materia_raw).strip()}' riga {riga_corrente}, ignorata.")
        continue
    
    docente_trovato = None
    docente_pot = None
    for d in range(counter_docenti):
        if docente_classe_materia[d][c][materia]>0:
            docente_trovato = d
        if docenti[d][INDEX_NOME] == str(docente_raw).strip():
            docente_pot = d

    
    if docente_trovato is None:
        print(f"Nessun docente in classe {cl} che insegna '{str(materie_raw).strip()}', riga {riga_corrente} ignorata.")
        continue
    if docente_pot is None:
        print(f"Nessun docente '{str(docente_raw).strip()}', riga {riga_corrente} ignorata.")
        continue

    # assumo che per ogni ora della materia debba essere presente il potenziamento, per questo motivo se le ore di potenziamento assegnate non sono sufficienti errore
    if docente_classe_potenziamento[docente_pot][c] != docente_classe_materia[docente_trovato][c][materia]:
        print(f"'{str(docente_raw).strip()}' deve fare potenziamento durante le ore di {materia} ma ha un numero diverso di ore assegnate, riga {riga_corrente} ignorata.")
        continue
    # assumo inoltre che il docente faccia solo una materia all'interno della classe e che quindi non sia necessario distinguere quali ore sono assegnate a quali materie
    counter_materie = 0
    for m in range(13):
        if docente_classe_materia[docente_trovato][c][m]>0:
            counter_materie=counter_materie+1
    if counter_materie>1:
        print(f"'{str(docente_trovato).strip()}' deve fare solo la materia '{materia}' ma ha un numero maggiore di materie assegnate, riga {riga_corrente} ignorata.")
        continue
    
    potenziamento_per_materia.append((materia, docente_trovato, docente_pot, c))
    riga_corrente += 1



                    
##### END INPUT READING

##### VINCOLI HARD

model = cp_model.CpModel()

# creo una matrice quadri-dimensionale (docente-classe-giorno-ora) i cui elementi sono variabili booleane

D, C, G, O = counter_docenti, counter_classi, 5, 16

matrice = {}
matrice_potenziamento = {}
matrice_alternativa = {}
matrice_motoria = {}

for d in range(D):
    for c in range(C):
        for g in range(G):
            for blocco in blocchi_orari(c, g):
                o0 = blocco[0]
                v_lez = model.new_bool_var(f"matrice_{d}_{c}_{g}_{o0}")

                # creo anche una matrice quadri-dimensionale (docente-classe-giorno-ora) i cui elementi sono variabili booleane per
                # i potenziamenti
                v_pot = model.new_bool_var(f"matrice_pot_{d}_{c}_{g}_{o0}")

                # creo anche una matrice quadri-dimensionale (docente-classe-giorno-ora) i cui elementi sono variabili booleane per
                # l'alternativa
                v_alt = model.new_bool_var(f"matrice_alt_{d}_{c}_{g}_{o0}")

                # creo anche una matrice quadri-dimensionale (docente-classe-giorno-ora) i cui elementi sono variabili booleane per
                # le ore di occupazione della palestra per motoria
                v_mot = model.new_bool_var(f"matrice_mot_{d}_{c}_{g}_{o0}")

                for o in blocco:
                    matrice[d, c, g, o] = v_lez
                    matrice_potenziamento[d, c, g, o] = v_pot
                    matrice_alternativa[d, c, g, o] = v_alt
                    matrice_motoria[d, c, g, o] = v_mot



# valore 0 fisso per classi fuori dal loro orario scolastico
for d in range(D):
    for c in range(C):
        for g in range(G):
            for o in range(O):
                if c < classi[3][0] and g == MER and (o == O9 or o == O10): #le classi 1, 2 e 3 il mercoledì fanno solo 4 ore
                    model.add(matrice[d, c, g, o]==0)
                    model.add(matrice_potenziamento[d, c, g, o]==0)
                    model.add(matrice_alternativa[d, c, g, o]==0)
                elif (g == LUN or g == VEN or g == MER):
                    if (c < classi[3][0] and o > O9) or (c >= classi[3][0] and o > O10): #le classi 1, 2 e 3 escono alle 13 LUN e VEN, mentre le classi 4 e 5 alle 13:30
                        model.add(matrice[d, c, g, o]==0)
                        model.add(matrice_potenziamento[d, c, g, o]==0)
                        model.add(matrice_alternativa[d, c, g, o]==0)

# in orario scolastico ogni classe deve sempre avere almeno un docente assegnato
for c in range(C):
    for g in range(G):
        for o in range(O):
            in_orario_scolastico = False
            
            # Condizioni per le classi 1, 2, 3
            if c < classi[3][0]:
                if g in [LUN, VEN] and o <= O9:
                    in_orario_scolastico = True
                elif g in [MAR, GIO] and o <= O15:
                    in_orario_scolastico = True
                elif g == MER and o <= O7:
                    in_orario_scolastico = True
                    
            # Condizioni per le classi 4, 5
            else:
                if g in [LUN, MER, VEN] and o <= O9:     # 8:30 - 13:30
                    in_orario_scolastico = True
                elif g in [MAR, GIO] and o <= O15:       # 8:30 - 16:30
                    in_orario_scolastico = True

            if in_orario_scolastico:
                model.add(sum(matrice[d, c, g, o] for d in range(D)) == 1)


# non ha senso assegnare potenziamento per le classi 1, 2 e 3 il lunedì e il venerdì quando hanno una mezz'ora dalle 10:30 alle 11:00
for d in range(D):
    for c in range(C):
        for g in [LUN, VEN]:
            model.add(matrice_potenziamento[d, c, g, O5]==0)

# non ha senso assegnare potenziamento durante motoria o a religione salvo eccezioni specificate in input

classi_con_pot_motoria_o_religione = {c for mat, _, _, c in potenziamento_per_materia if mat == MOT or mat == REL}

for d in range(D):
    for c in range(C):
        # Escludiamo sia le classi secondarie sia quelle con potenziamento esplicitamente previsto a motoria
        if (docente_classe_materia[d][c][REL] > 0 or docente_classe_materia[d][c][MOT] > 0) and c not in classi_secondarie_mot and c not in classi_con_pot_motoria_o_religione:
            for g in range(G):
                for o in range(O):
                    model.add(sum(matrice_potenziamento[d1, c, g, o] for d1 in range(D)) == 0).only_enforce_if(matrice[d, c, g, o])

# non ha senso assegnare potenziamento durante ore di docenti che hanno SOLO materie di ARTE, RELIGIONE, MOTORIA, MENSA o MUSICA
for d in range(D):
    for c in range(C):
        # eccezione per motoria in 4A e 4B
        if c == classi[3][0] and docente_classe_materia[d][c][MOT] > 0:
            continue
        solo_materie = True
        for m in range(13):
            if docente_classe_materia[d][c][m]>0 and m not in [ART, REL, MOT, MEN, MUS]:
                solo_materie = False
        if solo_materie:
            for g in range(G):
                for o in range(O):
                    model.add(sum(matrice_potenziamento[d1, c, g, o] for d1 in range(D))==0).only_enforce_if(matrice[d, c, g, o])

# un docente non può lavorare nei giorni o nelle ore in cui non ha dato disponibilità
for d in range(D):
    for g in range(G):
        for o in range(O):
            # Se in questo giorno (g) e in questa specifica mezz'ora (o) il docente non è disponibile
            if docenti[d][INDEX_GIORNI_DISPONIBILITA][g][o] == 0:
                for c in range(C):
                    model.add(matrice[d, c, g, o] == 0)
                    model.add(matrice_potenziamento[d, c, g, o] == 0)
                    model.add(matrice_alternativa[d, c, g, o] == 0)


# ogni docente in ogni classe fa esattamente il numero di ore che deve fare da assegnazioni
for d in range(D):
    for c in range(C):
        # NOTA: moltiplico le ore delle assegnazioni per 2 perché nella matrice ragiono in mezz'ore
        # Le ore curricolari (matrice) DEVONO corrispondere a tutte le materie tranne ALT (10)
        model.add(sum(matrice[d, c, g, o] for g in range(G) for o in range(O)) == sum(docente_classe_materia[d][c][m] for m in range(13) if m != ALT) * 2)

        # Le ore di alternativa (matrice_alternativa) DEVONO corrispondere ESATTAMENTE al valore di ALT in input
        model.add(sum(matrice_alternativa[d, c, g, o] for g in range(G) for o in range(O)) == docente_classe_materia[d][c][ALT] * 2)

# ogni docente in ogni classe fa esattamente il numero di ore di POTENZIAMENTO che deve fare da assegnazioni
for d in range(D):
    for c in range(C):
        model.add(sum(matrice_potenziamento[d, c, g, o] for g in range(G) for o in range(0, O)) == docente_classe_potenziamento[d][c]*2)

# le due ore di religione non possono essere spezzate (ASSUMO CHE IL DOCENTE DI RELIGIONE FACCIA SEMPRE SOLO QUELLO)
for c in range(C):
    for d in range(D):
        if docente_classe_materia[d][c][REL]>0:
            for g in range(G):
                for o in range(O):
                    if o == O1 or o == O2:
                        model.add_implication(matrice[d, c, g, o], matrice[d, c, g, o+2])
                    # se è l'ultima ora del giorno
                    elif (((o==O15 or o==O16) and (g==MAR or g==GIO)) or ((o==O7 or o==O8) and g==MER and c<classi[3][0]) or 
                          ((o==O8 or o==O9) and (g==LUN or g==VEN) and c<classi[3][0]) or ((o==O9 or o==O10) and (g==LUN or g==MER or g==VEN) and c>=classi[3][0])):
                        model.add(matrice[d, c, g, o-2] == 1).only_enforce_if(matrice[d, c, g, o])
                    # se è un'ora intermedia
                    elif not ((c>=classi[3][0] and g in [LUN, MER,VEN] and o>O10) or (c<classi[3][0] and g in [LUN, VEN] and o>O9) or (c<classi[3][0] and g==MER and o>O8)):
                        model.add(matrice[d, c, g, o-2] + matrice[d, c, g, o+2] >= 1).only_enforce_if(matrice[d, c, g, o])


# docente di ITA e di MAT hanno blocco prime 2 ore per 2 volte in ogni classe
for d in range(D):
    for c in range(C):
        if docente_classe_materia[d][c][ITA]!=0 or docente_classe_materia[d][c][MAT]!=0:
            model.add(sum(matrice[d,c,g,o] for g in range(G) for o in range(O5))==8)
            for g in range(G):
                model.add_linear_constraint(sum(matrice[d,c,g,o] for o in range(O5)), 0, 4) #le prime due ore o sono attaccate o non ci sono

# Inglese in 3, 4 e 5 ha 3 ore settimanali: evitare che siano tutte e 3 nello stesso giorno (tipo 2 al mattino e una a pomeriggio)
for d in range(D):
    for c in range(C):
        ore_ing = docente_classe_materia[d][c][ING]
        ore_totali = sum(docente_classe_materia[d][c][m] for m in range(13) if m != ALT)
        
        # Applichiamo il limite SOLO se Inglese costituisce la maggior parte dell'orario del docente 
        # per quella classe, per non strozzare chi ha anche Matematica/Italiano (es. Ferrari in 5A)
        if ore_ing != 0 and c >= classi[2][0] and ore_totali <= ore_ing + 1:
            for g in range(G):
                model.add(sum(matrice[d, c, g, o] for o in range(O)) < ore_ing * 2)

# chi fa la mezz'ora del lunedì e del venerdì è la stesso docente per la stessa classe
for d in range(D):
    for c in range(classi[3][0]):
        model.add(matrice[d, c, LUN, O5]==matrice[d, c, VEN, O5])

# evitare di far venire qualcuno a scuola un giorno per fare solo 1 ora di lavoro, almeno 2 ore
for d in range(D):
    for g in range(G):
        giornata_lavorativa = model.new_bool_var(f"giornata_lavorativa_{d}_{g}")
        model.add((sum(matrice[d, c, g, o] for c in range(C) for o in range(O))+
                   sum(matrice_potenziamento[d, c, g, o] for c in range(C) for o in range(O)) +
                   sum(matrice_alternativa[d, c, g, o] for c in range(C) for o in range(O)))>=4).only_enforce_if(giornata_lavorativa)
        model.add((sum(matrice[d, c, g, o] for c in range(C) for o in range(O))+
                           sum(matrice_potenziamento[d, c, g, o] for c in range(C) for o in range(O)) +
                           sum(matrice_alternativa[d, c, g, o] for c in range(C) for o in range(O)))==0).only_enforce_if(~giornata_lavorativa)



# i docenti che hanno part-time devono avere il numero corretto di giorni liberi, il lunedì non può essere libero in quanto c'è programmazione
# i docenti che non hanno part-time invece non devono avere giorni liberi
for d in range(D):
    giorni_liberi_vars = []
    for g in [MAR, MER, GIO, VEN]:
        # booleana: vale 1 se il giorno g è completamente libero da lezioni/pot/alt
        e_libero = model.new_bool_var(f"libero_{d}_{g}")
            
        ore_giorno = (
            sum(matrice[d, c, g, o] for c in range(C) for o in range(O))
            + sum(matrice_potenziamento[d, c, g, o] for c in range(C) for o in range(O))
            + sum(matrice_alternativa[d, c, g, o] for c in range(C) for o in range(O))
        )
            
        model.add(ore_giorno == 0).only_enforce_if(e_libero)
        model.add(ore_giorno >= 1).only_enforce_if(~e_libero)
            
        giorni_liberi_vars.append(e_libero) # <-- Aggiungiamo la variabile alla lista
            
    ore_lunedi = (
        sum(matrice[d, c, LUN, o] for c in range(C) for o in range(O))
        + sum(matrice_potenziamento[d, c, LUN, o] for c in range(C) for o in range(O))
        + sum(matrice_alternativa[d, c, LUN, o] for c in range(C) for o in range(O))
    )
    model.add(ore_lunedi >= 1)

    if docenti[d][INDEX_PART_TIME][1]>0: # se il docente fa part-time e ha specificato un certo numero di giorni che ASSOLUTAMENTE devono essere liberi
        # La somma delle variabili deve essere uguale ai giorni richiesti
        model.add(sum(giorni_liberi_vars) == docenti[d][INDEX_PART_TIME][1]) 
    elif docenti[d][INDEX_PART_TIME][0]==0: #docente che NON fa part time
        # La somma delle variabili deve essere uguale a 0 giorni liberi
        model.add(sum(giorni_liberi_vars) == 0)
        



# ITA e MAT è bene che vedano tutti i giorni o almeno 4 giorni su 5 ogni classe in cui lavorano
for d in range(D):
    for c in range(C):
        if docente_classe_materia[d][c][ITA]>0 or docente_classe_materia[d][c][MAT]>0:
            giorni_senza_entrare = []
            for g in range(G):
                entra = model.new_bool_var(f"entra_{d}_{c}_{g}")

                ore_giorno = (sum(matrice[d, c, g, o] for o in range(O))
                              + sum(matrice_potenziamento[d, c, g, o] for o in range(O))
                              + sum(matrice_alternativa[d, c, g, o] for o in range(O)))
                
                model.add(ore_giorno > 1).only_enforce_if(entra)
                model.add(ore_giorno == 0).only_enforce_if(~entra)
                giorni_senza_entrare.append(~entra)
            model.add_at_most_one(giorni_senza_entrare)

# i docenti che hanno assegnata la mensa in una certa classe devono per forza farla alla quinta ora del martedì o del giovedì
for d in range(D):
    for c in range(C):
        model.add(sum(matrice[d, c, g, o] for g in [MAR, GIO] for o in range(O9, O11)) == docente_classe_materia[d][c][MEN]*2)

# le ore di potenziamento non ha senso che siano contemporanee alle ore di mensa
for d in range(D):
    for c in range(C):
        model.add(sum(matrice_potenziamento[d, c, g, o] for g in [MAR, GIO] for o in range(O9, O11)) == 0)

# in ogni classe le ore di alternativa vanno fatte contemporaneamente alle ore di religione
for c in range(C):
    for d1 in range(D):
        if docente_classe_materia[d1][c][REL]>0:
            for d2 in range(D):
                if docente_classe_materia[d2][c][ALT]>0:
                    for g in range(G):
                        for o in range(O):
                            model.add_implication(matrice[d1, c, g, o], matrice_alternativa[d2, c, g, o])

# In una stessa classe non ci possono essere due docenti che contemporaneamente fanno potenziamento
for c in range(C):
    for g in range(G):
        for o in range(O):
            model.add(sum(matrice_potenziamento[d,c,g,o] for d in range(D))<=1)


# Potenziamenti imposti contemporanei ad una specifica materia
for m, d, dp, c in potenziamento_per_materia:
    model.add

# vincolo per garantire la mutua esclusione del docente solo quando necessario:
# un docente che ha una certa materia m condivisa tra due classi c1 e c2 può essere solo per il numero di ore necessarie allo
# svolgimento di m in entrambe le classi c1 e c2
# indicatore dedicato "in questo slot il docente d sta facendo la materia condivisa alla classe c"
# creato solo per le combinazioni (d, materia, c) effettivamente coinvolte in un abbinamento
matrice_materia_condivisa = {}

def get_o_crea_matrice_condivisa(d, materia, c):
    chiave = (d, materia, c)
    if chiave in matrice_materia_condivisa:
        return matrice_materia_condivisa[chiave]
    # per ogni ora di ogni giorno mi creo una variabile, se questa variabile è True allora vuol dire che il docente d (input)
    # si trova nella classe c (input) e sta facendo la materia (input). A matrice_materia_condivisa data una chiave (d, materia, c)
    # è associata questa grid che contiene appunto la lista di tutte le ore in cui quella materia è fatta
    if materia == MOT:
        # riusiamo la matrice_motoria in quanto fa la stessa cosa e non voglio un duplicato parallelo
        grid = {(g, o): matrice_motoria[d, c, g, o] for g in range(G) for o in range(O)}
    elif materia == ALT:
        # riusiamo la matrice_alternativa in quanto fa la stessa cosa e non voglio un duplicato parallelo
        grid = {(g, o): matrice_alternativa[d, c, g, o] for g in range(G) for o in range(O)}
    else:
        grid = {}
        for g in range(G):
            for blocco in blocchi_orari(c, g):
                o0 = blocco[0]
                v = model.new_bool_var(f"mat_cond_{d}_{materia}_{c}_{g}_{o0}")
                model.add_implication(v, matrice[d, c, g, o0])
                for o in blocco:
                    grid[(g, o)] = v
        # il totale di questa matrice dedicata deve coincidere con le ore assegnate di quella materia
        model.add(sum(grid.values()) == docente_classe_materia[d][c][materia] * 2)
    matrice_materia_condivisa[chiave] = grid
    return grid

# sincronizzo le due classi sulla materia condivisa
for materia, d, (c1, c2) in abbinamenti_per_materia:
    grid1 = get_o_crea_matrice_condivisa(d, materia, c1)
    grid2 = get_o_crea_matrice_condivisa(d, materia, c2)
    for g in range(G):
        for o in range(O):
            model.add(grid1[(g, o)] == grid2[(g, o)]) # impongo che le due grid siano uguali

# mutua esclusione: un docente non condiviso non può mai stare in due classi insieme;
# un docente condiviso può, ma SOLO per gli slot marcati come materia condivisa
for d in range(D):
    if d in docenti_condivisi:
        for g in range(G):
            for o in range(O):
                tolleranza = [ # lista per sommare 0 e 1, quando un docente d in una certa ora si trova in più classi 
                    matrice_materia_condivisa[(d, materia, c2)][(g, o)]
                    for materia, dd, (c1, c2) in abbinamenti_per_materia
                    if dd == d and materia != ALT
                ]
                model.add(sum(matrice[d, c, g, o] for c in range(C)) <= 1 + sum(tolleranza))
    else:
        for g in range(G):
            for o in range(O):
                model.add(sum(matrice[d, c, g, o] for c in range(C)) <= 1)

# ogni docente (compresi quelli condivisi) in un certo istante può fare solo ore curricolari, ore di potenziamento o ore di alternativa

# calcolo gli abbinamenti senza duplicati, mantenendo l'ordine
abbinamenti_unici = list(dict.fromkeys(
    (materia, dd, c1, c2) for materia, dd, (c1, c2) in abbinamenti_per_materia
))

grid_curr = {d: [] for d in range(D)}  # griglie condivise dove d è il docente condiviso (qualsiasi materia)
grid_alt = {d: [] for d in range(D)}   # griglie condivise dove d è il docente condiviso solo per ALT
grid_pot = {d: [] for d in range(D)}   # griglie condivise in cui d fa potenziamento in ENTRAMBE le classi

pot_righe = {(dp, dd, mat, cl) for mat, dd, dp, cl in potenziamento_per_materia}
docenti_pot = {dp for _, _, dp, _ in potenziamento_per_materia}

for materia, dd, c1, c2 in abbinamenti_unici:
    grid = matrice_materia_condivisa[(dd, materia, c2)]
    grid_curr[dd].append(grid)
    if materia == ALT:
        grid_alt[dd].append(grid)
    for dp in docenti_pot:
        if (dp, dd, materia, c1) in pot_righe and (dp, dd, materia, c2) in pot_righe:
            grid_pot[dp].append(grid)

for d in range(D):
    for g in range(G):
        for o in range(O):
            tolleranza_curr = sum(grid[(g, o)] for grid in grid_curr[d])
            tolleranza_alt = sum(grid[(g, o)] for grid in grid_alt[d])
            tolleranza_pot = sum(grid[(g, o)] for grid in grid_pot[d])

            ore_curricolari = sum(matrice[d, c, g, o] for c in range(C)) - tolleranza_curr
            ore_potenziamento = sum(matrice_potenziamento[d, c, g, o] for c in range(C)) - tolleranza_pot
            ore_alternativa = sum(matrice_alternativa[d, c, g, o] for c in range(C)) - tolleranza_alt
            model.add(ore_curricolari + ore_potenziamento + ore_alternativa <= 1)

# un docente non può fare più di 3 ore di fila in una stessa classe
for d in range(D):
    for c in range(C):
        if sum(docente_classe_materia[d][c][m] for m in range(13)) >0:
            for g in range(G):
                blocchi = [] #lista di blocchi vicini esclusa la mezz'ora
                for blocco in blocchi_orari(c, g):
                    if len(blocco)>1:
                        blocchi.append(blocco)
                for i in range(len(blocchi)-3):
                    model.add((matrice[d,c,g,blocchi[i][0]] + matrice[d,c,g,blocchi[i+1][0]] + matrice[d,c,g,blocchi[i+2][0]] + matrice[d,c,g,blocchi[i+3][0]])<4)

# un docente non può lavorare per tutte le 8 ore delle giornate lunghe
for d in range(D):
    for g in [MAR, GIO]:
        mezzore_lavorate = []
        for o in range(O):
            tolleranza_curr = sum(grid[(g, o)] for grid in grid_curr[d])
            tolleranza_alt = sum(grid[(g, o)] for grid in grid_alt[d])
            tolleranza_pot = sum(grid[(g, o)] for grid in grid_pot[d])

            ore_curricolari = sum(matrice[d, c, g, o] for c in range(C)) - tolleranza_curr
            ore_potenziamento = sum(matrice_potenziamento[d, c, g, o] for c in range(C)) - tolleranza_pot
            ore_alternativa = sum(matrice_alternativa[d, c, g, o] for c in range(C)) - tolleranza_alt

            mezzore_lavorate.append(ore_curricolari + ore_potenziamento + ore_alternativa)

        model.add(sum(mezzore_lavorate) <= 14)

# Un docente non deve avere la maggior parte delle sue ore in una classe a pomeriggio 
# (tranne motoria in quanto dipende dagli orari di disponibilità della palestra)

for d in range(D):
    for c in range(C):
        numero_ore = sum(docente_classe_materia[d][c][m] for m in range(13) if m != MOT)
        if numero_ore > 0:
            ore_pomeriggio = sum(matrice[d, c, g, o] for g in [MAR, GIO] for o in range(O11, O))
            model.add(ore_pomeriggio < (int(numero_ore/2)+2))

# le due ore rimanenti la mattina devono essere assegnate allo stesso docente sia in A che in B

for c in range(5):
    if c != 1:
        for d in range(D):
            if (docente_classe_materia[d][classi[c][0]][MAT] + docente_classe_materia[d][classi[c][0]][ITA] + docente_classe_materia[d][classi[c][1]][MAT] + docente_classe_materia[d][classi[c][1]][ITA]) == 0:
                model.add(sum(matrice[d, classi[c][0], g, o] for g in range(G) for o in range(O5)) == 
                        sum(matrice[d, classi[c][1], g, o] for g in range(G) for o in range(O5)))

# un docente che ha solo 1 ora (escluse quelle di MEN dal conteggio) non può essere la terza ora (ora di ricreazione => persa)

for d in range(D):
    for c in range(C):
        if sum(docente_classe_materia[d][c][m] for m in range(12) if m != ALT) == 1:
            for g in range(G):
                if c < classi[3][0]:
                    if g in [LUN, VEN]:
                        model.add(matrice[d, c, g, O5]==0)
                    else:
                        model.add(sum(matrice[d, c, g, o] for o in [O5, O6])==0)
                else:
                    model.add(sum(matrice[d, c, g, o] for o in [O5, O6])==0)


# POTENZIAMENTI CONTEMPORANEI AD UNA CERTA MATERIA

for materia, docente, docente_pot, c in potenziamento_per_materia:
    for g in range(G):
        for o in range(O):
            model.add(matrice[docente, c, g, o] == matrice_potenziamento[docente_pot, c, g, o])


#VINCOLI MOTORIA

# se si occupa la palestra per motoria (matrice_motoria==1) allora per forza anche il docente è impegnato per quella classe in quella ora
for d in range(D):
    for c in range(C):
        if docente_classe_materia[d][c][MOT] > 0:
            for g in range(G):
                for o in range(O):
                    model.add_implication(matrice_motoria[d, c, g, o], matrice[d, c, g, o])

#docenti che per una certa classe non fanno motoria hanno valore 0 sempre nella matrice_motoria
for c in range(C):
    for d in range(D):
        if docente_classe_materia[d][c][MOT]==0:
            model.add(sum(matrice_motoria[d, c, g, o] for g in range(G) for o in range(O))==0)

# ogni classe fa motoria quando ha la disponibilità della palestra
for c in range(C):
    for d in range(D):
        if docente_classe_materia[d][c][MOT]>0:
            for g in range(G):
                for o in range(O):
                    model.add(matrice_motoria[d, c, g, o] <= disponibilita_palestra[c][g][o])

# il numero di ore di motoria corrisponde al numero di ore in cui la classe è in palestra
for c in range(C):
    if c in classi_secondarie_mot: continue        
    for d in range(D):
        if docente_classe_materia[d][c][MOT] > 0:
            model.add(sum(matrice_motoria[d, c, g, o] for g in range(G) for o in range(O)) == docente_classe_materia[d][c][MOT] * 2)

# la palestra è unica e condivisa tra tutte le classi => non ci possono andare più classi contemporaneamente 
# (questo vincolo non si applica alle classi che fanno motoria contemporaneamente)
for g in range(G):
    for o in range(O):
        model.add(sum(matrice_motoria[d, c, g, o] for d in range(D) for c in range(C) if c not in classi_secondarie_mot) <= 1)

# se di quarte e quinte in totale ce ne sono 4 allora una fa le due ore separate, altrimenti se sono solo 3 le fanno tutte intere
numero_classi_quarte_e_quinte = len(classi[3]) + len(classi[4])
for materia, d, (c1, c2) in abbinamenti_per_materia:
    if (materia == MOT and c1>=classi[3][0] and c2>=classi[3][0]):
        numero_classi_quarte_e_quinte= numero_classi_quarte_e_quinte-1

if (numero_classi_quarte_e_quinte) < 4:
    for c in range(classi[3][0], classi[4][len(classi[4])-1]):
        for d in range(D):
            if docente_classe_materia[d][c][MOT]>0:
                for g in range(G):
                    for o in range(O):
                        if o == O1 or o == O2:
                            model.add_implication(matrice_motoria[d, c, g, o], matrice_motoria[d, c, g, o+2])
                        # se è l'ultima ora del giorno
                        elif (((o==O9 or o==O10) and (g==LUN or g==MER or g==VEN)) or ((o==O15 or o==O16) and (g==MAR or g==GIO))):
                            model.add(matrice[d, c, g, o-2] == 1).only_enforce_if(matrice[d, c, g, o])
                        # se è un'ora intermedia
                        elif (((o>O2 and o<O9) and (g==LUN or g==MER or g==VEN)) or ((o>O2 and o<O15) and (g==MAR or g==GIO))):
                            model.add((matrice[d, c, g, o-2] + matrice[d, c, g, o+2]) >= 1).only_enforce_if(matrice[d, c, g, o])

# OBIETTIVI

# evitare di spezzettare gli orari, soprattutto al mattino
malus_spezzettamento = []
for d in range(D):
    for c in range(C):
        for g in range(G):
            for o in range(O):
                # variabile booleana che vale 1 se l'ora o è un'ora singola isolata
                ora_singola = model.new_bool_var(f"singola_{d}_{c}_{g}_{o}")
                # se è la prima ora del giorno
                if o == O1 or o == O2:
                    model.add_implication(ora_singola, matrice[d, c, g, o])
                    model.add_implication(ora_singola, ~matrice[d, c, g, O3])
                    model.add_bool_or([~matrice[d, c, g, o], matrice[d, c, g, O3], ora_singola])
                # se è l'ultima ora del giorno
                elif (((o==O15 or o==O16) and (g==MAR or g==GIO)) or ((o==O7 or o==O8) and g==MER and c<classi[3][0]) or 
                      ((o==O8 or o==O9) and (g==LUN or g==VEN) and c<classi[3][0]) or ((o==O9 or o==O10) and (g==LUN or g==MER or g==VEN) and c>=classi[3][0])):
                    model.add_implication(ora_singola, matrice[d, c, g, o])
                    model.add_implication(ora_singola, ~matrice[d, c, g, o-2])
                    model.add_bool_or([~matrice[d, c, g, o], matrice[d, c, g, o-2], ora_singola])
                # se è un'ora intermedia (devo comunque escludere le ore in cui non si hanno lezione)
                elif not ((c>=classi[3][0] and g in [LUN, MER,VEN] and o>O10) or (c<classi[3][0] and g in [LUN, VEN] and o>O9) or (c<classi[3][0] and g==MER and o>O8)):
                    model.add_implication(ora_singola, matrice[d, c, g, o])
                    model.add_implication(ora_singola, ~matrice[d, c, g, o - 2])
                    model.add_implication(ora_singola, ~matrice[d, c, g, o + 2])
                    model.add_bool_or([~matrice[d, c, g, o], matrice[d, c, g, o - 2], matrice[d, c, g, o + 2], ora_singola])

                if (o < O9 and (g == MAR or g == GIO)) or (g == LUN or g == MER or g == VEN): #le ore al mattino spezzettate le contiamo due volte per dargli maggior peso
                    malus_spezzettamento.append(ora_singola)
                malus_spezzettamento.append(ora_singola)

# ogni docente part-time ha preferenze per quanto riguarda il giorno libero
bonus_giorno_libero = []
for d in range(D):
    if docenti[d][INDEX_PART_TIME][0]:
        for g in [MAR, MER, GIO, VEN]:
            preferenze_accordate = model.new_bool_var(f"preferenze_accordate_{d}_{g}")
            ore_giorno = (
                sum(matrice[d, c, g, o] for c in range(C) for o in range(O))
                + sum(matrice_potenziamento[d, c, g, o] for c in range(C) for o in range(O))
                + sum(matrice_alternativa[d, c, g, o] for c in range(C) for o in range(O))
            )
            
            # Applica i vincoli OR-Tools solo se la preferenza Python è rispettata
            if docenti[d][INDEX_PART_TIME][2]!=0 and docenti[d][INDEX_PART_TIME][2][g]==1:
                model.add(ore_giorno == 0).only_enforce_if(preferenze_accordate)
                model.add(ore_giorno > 0).only_enforce_if(~preferenze_accordate)
            else:
                # Se non è una sua preferenza, il bonus non può scattare
                model.add(preferenze_accordate == 0)

            bonus_giorno_libero.append(preferenze_accordate)

# docenti di ITA e MAT nelle giornate corte devono dare il compito per casa quindi è utile che entrino soprattutto lun e ven
bonus_docente_ITA_MAT_entra_nelle_corte = []

for d in range(D):
    for c in range(C):
        if docente_classe_materia[d][c][ITA] or docente_classe_materia[d][c][MAT]:
            for g in [LUN, MER, VEN]:
                entra = model.new_bool_var(f"entra_{d}_{c}_{g}")
                model.add(sum(matrice[d, c, g, o] for o in range(O))>0).only_enforce_if(entra)
                model.add(sum(matrice[d, c, g, o] for o in range(O))==0).only_enforce_if(~entra)
                bonus_docente_ITA_MAT_entra_nelle_corte.append(entra)

# un docente che fa 3 ore di fila in una stessa classe è poco funzionale e deve essere solo un'eccezione

malus_docente_3_ore_di_fila = []

for d in range(D):
    for c in range(C):
        if sum(docente_classe_materia[d][c][m] for m in range(13)) >0:
            for g in range(G):
                blocchi = [] #lista di blocchi vicini esclusa la mezz'ora
                for blocco in blocchi_orari(c, g):
                    if len(blocco)>1:
                        blocchi.append(blocco)
                for i in range(len(blocchi)):
                    tre_ore_di_fila = model.new_bool_var(f"tre_ore_di_fila_{d}_{c}_{g}_{i}")
                    if i == 0:
                        model.add((matrice[d,c,g,blocchi[i][0]] + matrice[d,c,g,blocchi[i+1][0]] + matrice[d,c,g,blocchi[i+2][0]])==3).only_enforce_if(tre_ore_di_fila)
                        model.add((matrice[d,c,g,blocchi[i][0]] + matrice[d,c,g,blocchi[i+1][0]] + matrice[d,c,g,blocchi[i+2][0]])!=3).only_enforce_if(~tre_ore_di_fila)
                    elif i == len(blocchi)-1:
                        model.add((matrice[d,c,g,blocchi[i][0]] + matrice[d,c,g,blocchi[i-1][0]] + matrice[d,c,g,blocchi[i-2][0]])==3).only_enforce_if(tre_ore_di_fila)
                        model.add((matrice[d,c,g,blocchi[i][0]] + matrice[d,c,g,blocchi[i-1][0]] + matrice[d,c,g,blocchi[i-2][0]])!=3).only_enforce_if(~tre_ore_di_fila)
                    else:
                        model.add((matrice[d,c,g,blocchi[i][0]] + matrice[d,c,g,blocchi[i+1][0]] + matrice[d,c,g,blocchi[i-1][0]])==3).only_enforce_if(tre_ore_di_fila)
                        model.add((matrice[d,c,g,blocchi[i][0]] + matrice[d,c,g,blocchi[i+1][0]] + matrice[d,c,g,blocchi[i-1][0]])!=3).only_enforce_if(~tre_ore_di_fila)
                
                    malus_docente_3_ore_di_fila.append(tre_ore_di_fila)
        
# se un docente in una classe fa più di 2 ore in una giornata (nelle lunghe), meglio che siano 2 al mattino e 1 o 2 a pomeriggio

bonus_ripartizione_2_mattina_2_pomeriggio = []

for d in range(D):
    for c in range(C):
        if sum(docente_classe_materia[d][c][m] for m in range(13)) >0:
            for g in [MAR, GIO]:
                ripartizione_2_2 = model.new_bool_var(f"ripartizione_2_2_{d}_{c}_{g}")
                piu_di_2 = model.new_bool_var(f"piu_di_2_{d}_{c}_{g}")
                mattino = []
                pomeriggio = []

                model.add(sum(matrice[d, c, g, o] for o in range(0, O, 2))>2).only_enforce_if(piu_di_2)
                model.add(sum(matrice[d, c, g, o] for o in range(0, O, 2))<=2).only_enforce_if(~piu_di_2)

                for o in range(0, O, 2):
                    mattina_2 = model.new_bool_var(f"mattina_2_{d}_{c}_{g}_{o}")
                    pomeriggio_2 = model.new_bool_var(f"pomeriggio_2_{d}_{c}_{g}_{o}")
                    if o < O9: #mattina
                        if o == O1:
                            model.add_bool_and([matrice[d, c, g, o], matrice[d, c, g, o+2], ~matrice[d, c, g, o+4]]).only_enforce_if(mattina_2)
                            model.add_bool_or([~matrice[d, c, g, o], ~matrice[d, c, g, o+2], matrice[d, c, g, o+4]]).only_enforce_if(~mattina_2)
                        elif o == O7:
                            model.add_bool_and([matrice[d, c, g, o], matrice[d, c, g, o-2], ~matrice[d, c, g, o-4]]).only_enforce_if(mattina_2)
                            model.add_bool_or([~matrice[d, c, g, o], ~matrice[d, c, g, o-2], matrice[d, c, g, o-4]]).only_enforce_if(~mattina_2)
                        else:
                            cond1 = model.new_bool_var(f"m_cond1_{d}_{c}_{g}_{o}")
                            model.add_bool_and([matrice[d, c, g, o], matrice[d, c, g, o+2], ~matrice[d, c, g, o-2]]).only_enforce_if(cond1)
                            model.add_bool_or([~matrice[d, c, g, o], ~matrice[d, c, g, o+2], matrice[d, c, g, o-2]]).only_enforce_if(~cond1)

                            cond2 = model.new_bool_var(f"m_cond2_{d}_{c}_{g}_{o}")
                            model.add_bool_and([matrice[d, c, g, o], matrice[d, c, g, o-2], ~matrice[d, c, g, o+2]]).only_enforce_if(cond2)
                            model.add_bool_or([~matrice[d, c, g, o], ~matrice[d, c, g, o-2], matrice[d, c, g, o+2]]).only_enforce_if(~cond2)

                            model.add_bool_or([cond1, cond2]).only_enforce_if(mattina_2)
                            model.add_bool_and([~cond1, ~cond2]).only_enforce_if(~mattina_2)
                        mattino.append(mattina_2)
                    else: #pomeriggio
                        if o == O9:
                            model.add_bool_and([matrice[d, c, g, o], matrice[d, c, g, o+2], ~matrice[d, c, g, o+4]]).only_enforce_if(pomeriggio_2)
                            model.add_bool_or([~matrice[d, c, g, o], ~matrice[d, c, g, o+2], matrice[d, c, g, o+4]]).only_enforce_if(~pomeriggio_2)
                        elif o == O15:
                            model.add_bool_and([matrice[d, c, g, o], matrice[d, c, g, o-2], ~matrice[d, c, g, o-4]]).only_enforce_if(pomeriggio_2)
                            model.add_bool_or([~matrice[d, c, g, o], ~matrice[d, c, g, o-2], matrice[d, c, g, o-4]]).only_enforce_if(~pomeriggio_2)
                        else:
                            cond1 = model.new_bool_var(f"p_cond1_{d}_{c}_{g}_{o}")
                            model.add_bool_and([matrice[d, c, g, o], matrice[d, c, g, o+2], ~matrice[d, c, g, o-2]]).only_enforce_if(cond1)
                            model.add_bool_or([~matrice[d, c, g, o], ~matrice[d, c, g, o+2], matrice[d, c, g, o-2]]).only_enforce_if(~cond1)

                            cond2 = model.new_bool_var(f"p_cond2_{d}_{c}_{g}_{o}")
                            model.add_bool_and([matrice[d, c, g, o], matrice[d, c, g, o-2], ~matrice[d, c, g, o+2]]).only_enforce_if(cond2)
                            model.add_bool_or([~matrice[d, c, g, o], ~matrice[d, c, g, o-2], matrice[d, c, g, o+2]]).only_enforce_if(~cond2)

                            model.add_bool_or([cond1, cond2]).only_enforce_if(pomeriggio_2)
                            model.add_bool_and([~cond1, ~cond2]).only_enforce_if(~pomeriggio_2)
                        pomeriggio.append(pomeriggio_2)

                almeno_1_mattina = model.new_bool_var(f"almeno_1_mattina_{d}_{c}_{g}")
                model.add(sum(mattino) > 0).only_enforce_if(almeno_1_mattina)
                model.add(sum(mattino) == 0).only_enforce_if(~almeno_1_mattina)

                almeno_1_pomeriggio = model.new_bool_var(f"almeno_1_pomeriggio_{d}_{c}_{g}")
                model.add(sum(pomeriggio) > 0).only_enforce_if(almeno_1_pomeriggio)
                model.add(sum(pomeriggio) == 0).only_enforce_if(~almeno_1_pomeriggio)

                model.add_bool_and([almeno_1_mattina, almeno_1_pomeriggio, piu_di_2]).only_enforce_if(ripartizione_2_2)
                model.add_bool_or([~almeno_1_mattina, ~almeno_1_pomeriggio, ~piu_di_2]).only_enforce_if(~ripartizione_2_2)

                bonus_ripartizione_2_mattina_2_pomeriggio.append(ripartizione_2_2)
                
# classi 1/2/3 il lunedì e il venerdì fanno 4 ore e mezza. La mezz’ora è la terza ed è un’ora “persa” perché c’è ricreazione. 
# Si cerca di darla il più possibile a un insegnante che nella classe ha tante ore, in modo da non penalizzarlo troppo
bonus_ora_persa = []

for c in range(classi[3][0]):
    for d in range(D):
        ore_in_quella_classe = 0
        for m in range(13):
            ore_in_quella_classe=ore_in_quella_classe+docente_classe_materia[d][c][m]
        for g in [LUN, VEN]:
            ora_persa = model.new_bool_var(f"ora_persa_{d}_{c}_{g}")
            model.add_implication(ora_persa, matrice[d, c, g, O5])
            bonus_ora_persa.append(ora_persa*ore_in_quella_classe)

# i “pomeriggi” piacciono poco... si lavora peggio e meno con i bambini, quindi soprattutto le due prevalenti 
# (quella di ITA e di MAT) cercare di trattarle un po’ pari…
malus_sbilanciamento_pomeriggio_ITA_MAT = []
for c in range(C):
    ore_pomeriggio_ITA = sum(
        matrice[d, c, g, o]
        for d in range(D) if docente_classe_materia[d][c][ITA] > 0
        for g in [MAR, GIO] for o in range(O11, O)
    )
    ore_pomeriggio_MAT = sum(
        matrice[d, c, g, o]
        for d in range(D) if docente_classe_materia[d][c][MAT] > 0
        for g in [MAR, GIO] for o in range(O11, O)
    )
    # Il limite deve coprire la massima differenza possibile (12 mezz'ore: due pomeriggi con tre ore ciascuno).
    differenza = model.new_int_var(0, 12 * D, f"differenza_{c}")
    model.add_abs_equality(differenza, ore_pomeriggio_ITA - ore_pomeriggio_MAT)
    malus_sbilanciamento_pomeriggio_ITA_MAT.append(differenza)

# evitare troppi buchi nell’orario del singolo docente (ma qualche buco l’abbiamo tutti), contare anche le ore di potenziamento
# e il LUN tutti i docenti sono impegnati nella programmazione dalle 14:00 alle 16:00, quindi tra i “buchi” conta anche se il 
# lunedì un docente finisce prima delle 13:30 e deve quindi fermarsi per ore in attesa di programmare
malus_buchi = []

for d in range(D):
    occupata = {
        (g, o): model.new_bool_var(f"occupata_{d}_{g}_{o}")
        for g in range(G)
        for o in range(O)
    }

    for g in range(G):
        for o in range(O):
            lezioni_o_potenziamento = (sum(matrice[d, c, g, o] for c in range(C)) + sum(matrice_potenziamento[d, c, g, o] for c in range(C)))
            if g == LUN and O12 <= o < O16:
                model.add(occupata[LUN, o] == 1) # aggiungo la programmazione tra le 14:00 e le 16:00
            else:
                model.add(lezioni_o_potenziamento >= 1).only_enforce_if(occupata[g, o])
                model.add(lezioni_o_potenziamento == 0).only_enforce_if(~occupata[g, o])

    buchi_giorno = []
    for g in range(G):
        for o in range(O):
            prima = model.new_bool_var(f"lezione_prima_{d}_{g}_{o}")
            dopo = model.new_bool_var(f"lezione_dopo_{d}_{g}_{o}")
            buco = model.new_bool_var(f"buco_{d}_{g}_{o}")

            precedenti = [occupata[g, i] for i in range(o)]
            successive = [occupata[g, i] for i in range(o + 1, O)]
            if precedenti:
                model.add_max_equality(prima, precedenti)
            else:
                model.add(prima == 0)
            if successive:
                model.add_max_equality(dopo, successive)
            else:
                model.add(dopo == 0)

            # Un'ora è un buco solo se è libera e c'è almeno una lezione sia prima sia dopo nello stesso giorno.
            model.add(buco <= 1 - occupata[g, o]) #se l'ora è occupata allora non c'è buco
            model.add(buco <= prima) #se non ci sono lezioni prima allora non c'è buco
            model.add(buco <= dopo) #se non ci sono lezioni dopo allora non c'è buco
            model.add(buco >= (1 - occupata[g, o]) + prima + dopo - 2) #se non è occupata, ci sono lezioni prima e ce ne sono dopo, allora buco == 1
            buchi_giorno.append(buco)

    malus_buchi.append(sum(buchi_giorno))

# religione si dà solitamente di pomeriggio a una classe alta (4 o 5)
bonus_religione_pomeriggio = []

for d in range(D):
    for c in range(C):
        if docente_classe_materia[d][c][REL]>0 and c>=classi[3][0]:
            for g in [MAR, GIO]:
                for o in range(O11, O, 2):
                    religione_a_pomeriggio = model.new_bool_var(f"religione_pomeriggio_{d}_{c}_{g}_{o}")
                    model.add_implication(religione_a_pomeriggio, matrice[d, c, g, o])
                    bonus_religione_pomeriggio.append(religione_a_pomeriggio)

# motoria di solito si mandano a pomeriggio le terze (tra le 13:30 e le 16:30)
bonus_motoria_terze_pomeriggio = []

for d in range(D):
    for c in classi[2]:
        if docente_classe_materia[d][c][MOT]>0:
            for g in range(G):
                for o in range(O11, O, 2):
                    motoria_a_pomeriggio = model.new_bool_var(f"motoria_a_pomeriggio_{d}_{c}_{o}")
                    model.add_implication(motoria_a_pomeriggio, matrice[d, c, g, o])
                    bonus_motoria_terze_pomeriggio.append(motoria_a_pomeriggio)
                
# ridurre il più possibile le ore settimanali in cui non c’è né potenziamento né alternativa disponibile per le sostituzioni
malus_no_sostituzione = []

for g in range(G):
    for o in range(O):
        no_sostituzione = model.new_bool_var(f"no_sostituzione_{g}_{o}")
        
        ore_sostituzione = (
            sum(matrice_potenziamento[d, c, g, o] for d in range(D) for c in range(C)) + 
            sum(matrice_alternativa[d, c, g, o] for d in range(D) for c in range(C))
        )
        
        # Se c'è almeno un'ora di pot/alt, "no_sostituzione" è Falso (0)
        model.add(ore_sostituzione >= 1).only_enforce_if(~no_sostituzione)
        
        # Se le ore sono 0, "no_sostituzione" è Vero (1)
        model.add(ore_sostituzione == 0).only_enforce_if(no_sostituzione)
        
        malus_no_sostituzione.append(no_sostituzione)

# l’ultimo blocco di due prime ore viene dato preferibilmente a:
#   -	INGLESE (se non è lo stesso docente di ita o di mat) nelle classi 3/4/5 in cui fa 3 ore
#   -   RELIGIONE (soprattutto nelle classi 1/2/3) anche per consentire al docente di alternativa di avere le prime ore se non 
#       ne avrebbe mai
#   -	oppure a chi ha una materia di studio: STO-GEO-SCI (se non è lo stesso docente di ita e di mat)
bonus_ultimo_blocco = []

for c in range(C):
    for d in range(D):
        peso = 0
        if docente_classe_materia[d][c][MAT] == 0 and docente_classe_materia[d][c][ITA] == 0:
            if docente_classe_materia[d][c][ING] > 0 and c >= classi[2][0]:
                peso = 3
            elif docente_classe_materia[d][c][REL] > 0:
                peso = 2
            elif (docente_classe_materia[d][c][STO] > 0 or 
                  docente_classe_materia[d][c][GEO] > 0 or 
                  docente_classe_materia[d][c][SCI] > 0):
                peso = 1
        if peso == 0:
            continue

        # booleana che vale 1 se il docente ottiene il blocco delle prime due ore
        ha_blocco = model.new_bool_var(f"ha_blocco_{d}_{c}")

        model.add(sum(matrice[d, c, g, o] for g in range(G) for o in range(O5)) == 4).only_enforce_if(ha_blocco)

        bonus_ultimo_blocco.append(ha_blocco * peso)

# Un docente che entra nelle due classi parallele deve avere un numero simile di ore pomeridiane tra le due classi
malus_sbilanciamento_ore_pomeridiane_tra_A_e_B = []

for d in range(D):
    for i in range(5):
        if len(classi[i])>1:
            trovato_match = False
            for m in range(13):
                if docente_classe_materia[d][classi[i][0]][m] == docente_classe_materia[d][classi[i][1]][m] and trovato_match==False:
                    trovato_match=True
                    ore_pomeriggio_in_A = sum(matrice[d, classi[i][0], g, o] for g in [MAR, GIO] for o in range(O11, O))
                    ore_pomeriggio_in_B = sum(matrice[d, classi[i][1], g, o] for g in [MAR, GIO] for o in range(O11, O))

                    # Il limite deve coprire la massima differenza possibile (12 mezz'ore: due pomeriggi con tre ore ciascuno).
                    differenza_A_B = model.new_int_var(0, 12 * D, f"differenza_A_B_{d}_{i}")
                    model.add_abs_equality(differenza_A_B, ore_pomeriggio_in_A - ore_pomeriggio_in_B)
                    malus_sbilanciamento_ore_pomeridiane_tra_A_e_B.append(differenza_A_B)


# pesi diversi per i diversi objectives
model.maximize(
    - 5*(sum(malus_buchi))                                        # imp 7, mezz'ore
    - 3*(sum(malus_no_sostituzione))                              # imp 6, mezz'ore
    - 7*(sum(malus_spezzettamento))                               # imp 8, mezz'ore (x2 mattino già nel codice)
    + 9*(sum(bonus_docente_ITA_MAT_entra_nelle_corte))            # imp 7
    - 20*(sum(malus_docente_3_ore_di_fila))                       # imp 9
    + 4*(sum(bonus_ora_persa))                                    # imp 9, già x ore del docente
    + 13*(sum(bonus_ripartizione_2_mattina_2_pomeriggio))         # imp 8
    - 10*(sum(malus_sbilanciamento_ore_pomeridiane_tra_A_e_B))    # imp 9, mezz'ore
    + 15*(sum(bonus_ultimo_blocco))                               # imp 10, già x peso 1/2/3
    + 9*(sum(bonus_giorno_libero))                                # imp 7
    - 7*(sum(malus_sbilanciamento_pomeriggio_ITA_MAT))            # imp 8, mezz'ore
    + 3*(sum(bonus_religione_pomeriggio))                         # imp 4
    + 2*(sum(bonus_motoria_terze_pomeriggio))                     # imp 3
)


#OUTPUT

# SOLVER ####################################################

# REPORT DIAGNOSTICO ########################################

NOMI_MATERIE = [
    "ITA", "STO", "GEO", "ART", "MUS", "MAT",
    "SCI", "TEC", "ING", "MOT", "ALT", "REL", "MEN", "POT"
]

NOMI_GIORNI = ["LUNEDÌ", "MARTEDÌ", "MERCOLEDÌ", "GIOVEDÌ", "VENERDÌ"]
FASCE_ORARIE = [
    "08:30-09:00", "09:00-09:30", "09:30-10:00", "10:00-10:30",
    "10:30-11:00", "11:00-11:30", "11:30-12:00", "12:00-12:30",
    "12:30-13:00", "13:00-13:30", "13:30-14:00", "14:00-14:30",
    "14:30-15:00", "15:00-15:30", "15:30-16:00", "16:00-16:30"
]

migliore_obj = None
tutti_obj = []
solver = None
status = None
migliore_soluzione = None   # valori di TUTTE le variabili del miglior run

n_vars = len(model.proto.variables)

for seed in range(1, 11):
    model.clear_hints()

    if migliore_soluzione is not None:
        for idx in range(n_vars):
            model.add_hint(model.get_int_var_from_proto_index(idx), migliore_soluzione[idx])

    s = cp_model.CpSolver()
    s.parameters.symmetry_level = 2
    s.parameters.max_time_in_seconds = 360.0
    s.parameters.log_search_progress = True
    s.parameters.linearization_level = 2
    s.parameters.num_workers = os.cpu_count() - 1
    s.parameters.random_seed = seed

    st = s.solve(model)

    if st in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        obj = s.objective_value
        tutti_obj.append(obj)
        print(f"[seed={seed}] status={s.status_name(st)}  objective={int(obj)}")
        if migliore_obj is None or obj > migliore_obj:
            migliore_obj = obj
            solver = s
            status = st
            migliore_soluzione = list(s.response_proto.solution)
    else:
        print(f"[seed={seed}] status={s.status_name(st)} -> scartato")

print(f"\n OBJECTIVES OTTENUTI: {tutti_obj}\n")
print(f"\n>>> Miglior risultato: objective = {int(migliore_obj) if migliore_obj is not None else 'N/A'} <<<\n")

# ATTENZIONE: se nessun seed produce una soluzione, gestiscilo prima di proseguire
if solver is None:
    print("Nessun run ha prodotto una soluzione FEASIBLE/OPTIMAL su nessun seed.")
    raise SystemExit(1)

print("\n" + "=" * 125)
print(f"ESITO RISOLUZIONE: {solver.status_name(status)}")
print("=" * 125)

if status in (cp_model.OPTIMAL, cp_model.FEASIBLE):
    print(f"PUNTEGGIO TOTALE OBJECTIVE: {int(solver.objective_value)}\n")

    # 1. BREAKDOWN COMPLETO OBJECTIVE (BONUS / MALUS)
    print("=" * 125)
    print("DETTAGLIO COMPONENTI FUNZIONE OBIETTIVO")
    print("=" * 125)
    
    def valuta_componente(lista_espressioni):
        valore = sum(solver.value(e) for e in lista_espressioni)
        return valore

    componenti = [
        ("Bonus Giorno Libero (Conti/Bianchi)", sum(solver.value(x) for x in bonus_giorno_libero), "+", 0, len(bonus_giorno_libero)),
        ("Malus Spezzettamento Orario", sum(solver.value(x) for x in malus_spezzettamento), "-", 0, len(malus_spezzettamento)),
        ("Bonus ITA/MAT Entra Giorni Corti", sum(solver.value(x) for x in bonus_docente_ITA_MAT_entra_nelle_corte), "+", 0, len(bonus_docente_ITA_MAT_entra_nelle_corte)),
        ("Bonus Ora Persa (Ricreazione 1/2/3)", sum(solver.value(x) for x in bonus_ora_persa), "+", 0, "N/A"),
        ("Malus Sbilanciamento Pomeriggi ITA/MAT", sum(solver.value(x) for x in malus_sbilanciamento_pomeriggio_ITA_MAT), "-", 0, 12 * D * C),
        ("Malus Buchi nell'Orario", sum(solver.value(x) for x in malus_buchi), "-", 0, D * G * O),
        ("Bonus Religione Pomeriggio (4/5)", sum(solver.value(x) for x in bonus_religione_pomeriggio), "+", 0, len(bonus_religione_pomeriggio)),
        ("Bonus Motoria Pomeriggio (Terze)", sum(solver.value(x) for x in bonus_motoria_terze_pomeriggio), "+", 0, len(bonus_motoria_terze_pomeriggio)),
        ("Malus Ore Senza Sostituzione (No POT/ALT)", sum(solver.value(x) for x in malus_no_sostituzione), "-", 0, G * O),
        ("Bonus Scelta Ultimo Blocco Prime Ore", sum(solver.value(x) for x in bonus_ultimo_blocco), "+", 0, len(bonus_ultimo_blocco) * 3),
        ("Bonus Ripartizione 2 Mattina 2 Pomeriggio", sum(solver.value(x) for x in bonus_ripartizione_2_mattina_2_pomeriggio), "+", 0, len(bonus_ripartizione_2_mattina_2_pomeriggio) * 3),
        ("Malus Docente 3 ore di Fila", sum(solver.value(x) for x in malus_docente_3_ore_di_fila), "-", 0, len(malus_docente_3_ore_di_fila) * 3),
        ("Malus Sbilanciamento Ore Pomeridiane tra A e B", sum(solver.value(x) for x in malus_sbilanciamento_ore_pomeridiane_tra_A_e_B), "-", 0, len(malus_sbilanciamento_ore_pomeridiane_tra_A_e_B) * 3),    
    ]

    header_obj = f"| {'Componente':<45} | {'Tipo':<6} | {'Valore':<10} | {'Contributo Obj':<16} | {'Range Teorico':<15} |"
    sep_obj = "+" + "-" * 47 + "+" + "-" * 8 + "+" + "-" * 12 + "+" + "-" * 18 + "+" + "-" * 17 + "+"
    
    print(sep_obj)
    print(header_obj)
    print(sep_obj)
    
    totale_verificato = 0
    for nome_comp, val, segno, r_min, r_max in componenti:
        contributo = val if segno == "+" else -val
        totale_verificato += contributo
        range_str = f"[{r_min}, {r_max}]"
        print(f"| {nome_comp:<45} | {segno:<6} | {val:<10} | {contributo:<+16} | {range_str:<15} |")
    
    print(sep_obj)
    print(f"| {'SOMMA DEI CONTRIBUTI':<45} | {'=':<6} | {'':<10} | {totale_verificato:<+16} | {'':<15} |\n")

# 2. TABELLE ORARIO SETTIMANALE PER CLASSE
    print("=" * 125)
    print("ORARI SETTIMANALI PER CLASSE")
    print("=" * 125)

    col_w = 23
    sep_line = "+" + "-" * 15 + ("+" + "-" * (col_w + 2)) * G + "+"

    for c in range(C):
        nome_cl = NOMI_CLASSI[c]
        print(f"\n>>>>>> CLASSE {nome_cl} <<<<<<")
        print(sep_line)
        header = f"| {'FASCIA':<13} |" + "|".join([f" {NOMI_GIORNI[g]:<{col_w}} " for g in range(G)]) + "|"
        print(header)
        print(sep_line)

        for o in range(O):
            # Raccogliamo le stringhe di testo per ogni giorno in questo slot orario
            giorni_lines = []
            
            for g in range(G):
                assegnazioni_ora = []

                # Docenti e materie curricolari
                for d in range(D):
                    if solver.value(matrice[d, c, g, o]) == 1:
                        mat_doc = [NOMI_MATERIE[m] for m in range(13) if docente_classe_materia[d][c][m] > 0]
                        mat_str = "/".join(mat_doc) if mat_doc else "MAT"
                        if solver.value(matrice_motoria[d, c, g, o]) == 1:
                            mat_str = "MOT(Pal)"
                        assegnazioni_ora.append(f"{docenti[d][INDEX_NOME]} ({mat_str})")

                # Docenti di Potenziamento
                for d in range(D):
                    if solver.value(matrice_potenziamento[d, c, g, o]) == 1:
                        prefisso = "+ " if assegnazioni_ora else ""
                        assegnazioni_ora.append(f"{prefisso}{docenti[d][INDEX_NOME]} [POT]")

                # Docenti di Alternativa
                for d in range(D):
                    if solver.value(matrice_alternativa[d, c, g, o]) == 1:
                        prefisso = "+ " if assegnazioni_ora else ""
                        assegnazioni_ora.append(f"{prefisso}{docenti[d][INDEX_NOME]} [ALT]")

                if not assegnazioni_ora:
                    assegnazioni_ora = ["---"]
                
                # Usiamo textwrap per mandare a capo le stringhe più lunghe di col_w
                wrapped_lines = []
                for ass in assegnazioni_ora:
                    wrapped_lines.extend(textwrap.wrap(ass, width=col_w))
                
                giorni_lines.append(wrapped_lines)

            # Troviamo quante righe "fisiche" servono per la cella più alta di questa fascia oraria
            max_lines = max(len(lines) for lines in giorni_lines)
            if max_lines == 0:
                max_lines = 1

            # Stampiamo la cella riga per riga
            for line_idx in range(max_lines):
                # L'orario viene stampato solo sulla primissima riga
                fascia_str = FASCE_ORARIE[o] if line_idx == 0 else ""
                riga_celle = [f"| {fascia_str:<13} |"]
                
                for g in range(G):
                    # Se ci sono ancora righe di testo da stampare per questo giorno, le mettiamo
                    if line_idx < len(giorni_lines[g]):
                        cella_str = giorni_lines[g][line_idx]
                    else:
                        cella_str = "" # Altrimenti riempiamo con spazi vuoti
                        
                    riga_celle.append(f" {cella_str:<{col_w}} |")

                print("".join(riga_celle))
            print(sep_line)

    # 3. VERIFICA QUADRATURA ORE DOCENTI
    print("\n" + "=" * 125)
    print("QUADRATURA ORE DOCENTI (Mezz'ore -> Ore Effettive)")
    print("=" * 125)
    sep_doc = "+" + "-" * 22 + "+" + "-" * 15 + "+" + "-" * 15 + "+" + "-" * 15 + "+" + "-" * 15 + "+"
    print(sep_doc)
    print(f"| {'Docente':<20} | {'Curricolari':<13} | {'Alternativa':<13} | {'Potenziamento':<13} | {'Totale Ore':<13} |")
    print(sep_doc)

    for d in range(D):
        mezz_curr = sum(solver.value(matrice[d, c, g, o]) for c in range(C) for g in range(G) for o in range(O))
        mezz_alt = sum(solver.value(matrice_alternativa[d, c, g, o]) for c in range(C) for g in range(G) for o in range(O))
        mezz_pot = sum(solver.value(matrice_potenziamento[d, c, g, o]) for c in range(C) for g in range(G) for o in range(O))
        
        tot_ore = (mezz_curr + mezz_alt + mezz_pot) / 2.0
        print(f"| {docenti[d][INDEX_NOME]:<20} | {mezz_curr/2.0:<13.1f} | {mezz_alt/2.0:<13.1f} | {mezz_pot/2.0:<13.1f} | {tot_ore:<13.1f} |")
    print(sep_doc)

elif status == cp_model.INFEASIBLE:
    print("ATTENZIONE: Il modello è INFEASIBLE. Nessuna configurazione oraria soddisfa i vincoli hard impostati.")
else:
    print("Stato della ricerca sconosciuto o tempo limite raggiunto.")



####### WRITING OUTPUT TO NEW .xlsx FILE

# funzione che elimina il file se era già stato creato
# fatta solamente per comodità in modo da non dover tutte le volte chiudere il file aperto ed eliminarlo
def chiudi_ed_elimina(percorso_file: str):
    file_path = Path(percorso_file).resolve()

    if not file_path.exists():
        return

    try:
        file_path.unlink()
        print(f"File '{file_path.name}' eliminato.")
        return
    except PermissionError:
        pass

    for proc in psutil.process_iter(['name']):
        try:
            if proc.info['name'] and proc.info['name'].lower() in ('excel.exe', 'soffice.bin'):
                proc.terminate()
                proc.wait(timeout=2)
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.TimeoutExpired):
            continue

    try:
        file_path.unlink()
        print(f"File '{file_path.name}' eliminato con successo.")
    except PermissionError:
        print(f"Impossibile eliminare '{file_path.name}': permessi insufficienti o file ancora bloccato.")
    except FileNotFoundError:
        pass

chiudi_ed_elimina("Orario_Definitivo_Adaptive.xlsx")
chiudi_ed_elimina("orario_per_docente.xlsx")

print("\nGenerazione dei file Excel formattati in corso...")

from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()
wb.remove(wb.active) # Rimuoviamo il foglio vuoto di default

wb_docenti = openpyxl.Workbook()
wb_docenti.remove(wb_docenti.active)

# --- DEFINIZIONE STILI EXCEL ---
font_bold = Font(bold=True)
fill_mensa = PatternFill(start_color="FFF200", end_color="FFF200", fill_type="solid") # Giallo
fill_motoria = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid") # Verdino
align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)

thin = Side(style='thin')
border_thin = Border(left=thin, right=thin, top=thin, bottom=thin)

# Stili per il testo ricco (più colori nella stessa cella)
inline_black = InlineFont(b=False, color="000000") # Nero grassetto per docenti normali
inline_red = InlineFont(b=True, color="D13B26")   # Rosso grassetto per POT
inline_blue = InlineFont(b=True, color="0055D4")  # Blu grassetto per ALT

# --- MAPPA DINAMICA DEI BLOCCHI ORARI ---
def get_blocks_for_class_day(c, g):
    if c < classi[3][0]:
        if g in [LUN, VEN]:
            return [(0,1), (2,3), (4,4), (5,6), (7,8)]
        elif g == MER:
            return [(0,1), (2,3), (4,5), (6,7)]
        else: # MAR, GIO
            return [(0,1), (2,3), (4,5), (6,7), (8,9), (10,11), (12,13), (14,15)]
    else: # c >= classi[3][0]
        if g in [LUN, MER, VEN]:
            return [(0,1), (2,3), (4,5), (6,7), (8,9)]
        else: # MAR, GIO
            return [(0,1), (2,3), (4,5), (6,7), (8,9), (10,11), (12,13), (14,15)]


# ===============================================================
# 1. GENERAZIONE FOGLI PER CLASSE (Raggruppati per Anno)
# ===============================================================

# Raggruppo gli indici delle classi per anno (es: '1A' -> '1')
anni = {}
for c in range(C):
    anno = NOMI_CLASSI[c][0]
    if anno not in anni:
        anni[anno] = []
    anni[anno].append(c)

for anno, classi_anno in anni.items():
    ws = wb.create_sheet(title=f"Anno_{anno}")
    
    for idx, c in enumerate(classi_anno):
        classe = NOMI_CLASSI[c]
        # Offset di colonna per affiancare le sezioni nello stesso foglio (7 colonne: 1 orari + 5 giorni + 1 spazio vuoto)
        col_offset = idx * 7 
        
        # Dimensioni colonne per questa specifica classe
        ws.column_dimensions[get_column_letter(1 + col_offset)].width = 13
        for i in range(5):
            ws.column_dimensions[get_column_letter(2 + i + col_offset)].width = 24
            
        # Intestazione classe (Riga 1)
        cella_classe = ws.cell(row=1, column=1 + col_offset, value=f"Classe {classe}")
        cella_classe.font = font_bold
        cella_classe.alignment = align_center
        ws.merge_cells(start_row=1, start_column=1 + col_offset, end_row=1, end_column=6 + col_offset)
        
        # Intestazione giorni (Riga 2)
        for g in range(G):
            cell = ws.cell(row=2, column=g + 2 + col_offset, value=NOMI_GIORNI[g])
            cell.font = font_bold
            cell.alignment = align_center
            cell.border = border_thin
            
        for o in range(O):
            ws.row_dimensions[o + 3].height = 28 # Riga 3 in poi per gli orari

        # Creazione colonna Orari per questa classe
        time_blocks = get_blocks_for_class_day(c, LUN)
        for start_o, end_o in time_blocks:
            start_str = FASCE_ORARIE[start_o][:5] 
            end_str = FASCE_ORARIE[end_o][-5:]    
            label = f"{start_str}/{end_str}"
            
            c_orario = ws.cell(row=start_o + 3, column=1 + col_offset, value=label)
            c_orario.alignment = align_center
            for o in range(start_o, end_o + 1):
                ws.cell(row=o + 3, column=1 + col_offset).border = border_thin
            if start_o != end_o:
                ws.merge_cells(start_row=start_o + 3, start_column=1 + col_offset, end_row=end_o + 3, end_column=1 + col_offset)

        # Compilazione griglia materie
        for g in range(G):
            day_blocks = get_blocks_for_class_day(c, g)
            
            for start_o, end_o in day_blocks:
                is_mensa = False
                is_motoria = False
                segments = []
                
                # Docenti Curricolari
                for d in range(D):
                    if solver.value(matrice[d, c, g, start_o]) == 1:
                        mat_doc = [NOMI_MATERIE[m] for m in range(13) if docente_classe_materia[d][c][m] > 0]
                        mat_str = "/".join(mat_doc) if mat_doc else "MAT"
                        
                        if solver.value(matrice_motoria[d, c, g, start_o]) == 1:
                            mat_str = "motoria"
                            is_motoria = True
                            
                        if "MEN" in mat_doc and start_o == 8 and end_o == 9 and (g == MAR or g == GIO):
                            testo = f"MENSA\n{docenti[d][INDEX_NOME].upper()}"
                            segments.append({"text": testo, "color": inline_black})
                            is_mensa = True
                        else:
                            testo = f"{docenti[d][INDEX_NOME].upper()}\n({mat_str.lower()})"
                            segments.append({"text": testo, "color": inline_black})

                # Docenti Potenziamento (con check per condivisioni)
                for d in range(D):
                    if solver.value(matrice_potenziamento[d, c, g, start_o]) == 1:
                        altre_classi = [NOMI_CLASSI[c_other] for c_other in range(C) if c_other != c and solver.value(matrice_potenziamento[d, c_other, g, start_o]) == 1]
                        condiviso_str = f" {{ con {' e '.join(altre_classi)} }}" if altre_classi else ""
                        testo_pot = f"(+ {docenti[d][INDEX_NOME].upper()} [POT]){condiviso_str}"
                        segments.append({"text": testo_pot, "color": inline_red})

                # Docenti Alternativa (con check per condivisioni)
                for d in range(D):
                    if solver.value(matrice_alternativa[d, c, g, start_o]) == 1:
                        altre_classi = [NOMI_CLASSI[c_other] for c_other in range(C) if c_other != c and solver.value(matrice_alternativa[d, c_other, g, start_o]) == 1]
                        condiviso_str = f" {{ con {' e '.join(altre_classi)} }}" if altre_classi else ""
                        testo_alt = f"(+ {docenti[d][INDEX_NOME].upper()} [ALT]){condiviso_str}"
                        religione_in_corso = any(
                            solver.value(matrice[d_rel, c, g, start_o]) == 1 and docente_classe_materia[d_rel][c][REL] > 0 
                            for d_rel in range(D)
                        )
                        if religione_in_corso:
                            segments.append({"text": testo_alt, "color": inline_blue})
                        else:
                            segments.append({"text": testo_alt, "color": inline_black})

                # classe con cui si fa lezione condivisa (Materie Curricolari)
                for d in range(D):
                    for materia, dd, (c1, c2) in abbinamenti_per_materia:
                        if dd == d:
                            if c1 == c:
                                if solver.value(matrice_materia_condivisa[(d, materia, c)][(g, start_o)]) == 1:
                                    segments.append({"text": f" {{ con classe {NOMI_CLASSI[c2]} }}", "color": inline_black})
                            elif c2 == c:
                                if solver.value(matrice_materia_condivisa[(d, materia, c)][(g, start_o)]) == 1:
                                    segments.append({"text": f" {{ con classe {NOMI_CLASSI[c1]} }}", "color": inline_black})

                blocks_rich = []
                for i, seg in enumerate(segments):
                    txt = seg["text"]
                    if i < len(segments) - 1:
                        txt += "\n"
                    blocks_rich.append(TextBlock(seg["color"], txt))

                # Scrittura cella principale
                if not blocks_rich:
                    main_cell = ws.cell(row=start_o + 3, column=g + 2 + col_offset, value="---")
                else:
                    rich_text = CellRichText(blocks_rich)
                    main_cell = ws.cell(row=start_o + 3, column=g + 2 + col_offset, value=rich_text)
                    
                main_cell.alignment = align_center
                
                if is_mensa:
                    main_cell.fill = fill_mensa
                elif is_motoria:
                    main_cell.fill = fill_motoria

                for o in range(start_o, end_o + 1):
                    ws.cell(row=o + 3, column=g + 2 + col_offset).border = border_thin
                if start_o != end_o:
                    ws.merge_cells(start_row=start_o + 3, start_column=g + 2 + col_offset, end_row=end_o + 3, end_column=g + 2 + col_offset)

wb.save("Orario_Definitivo_Adaptive_4.xlsx")


# ===============================================================
# 2. GENERAZIONE FOGLI PER SINGOLO DOCENTE
# ===============================================================

def ottieni_attivita_docente(d, g, o):
    segments = []
    is_mensa = False
    is_motoria = False
    
    # 1. Ricerca classi per ore curricolari (incluse Mensa e Motoria)
    classi_curr = [c for c in range(C) if solver.value(matrice[d, c, g, o]) == 1]
    if classi_curr:
        # Assumiamo che per ore condivise la materia sia la stessa, prendiamo i dati dalla prima classe
        c_first = classi_curr[0]
        mat_doc = [NOMI_MATERIE[m] for m in range(13) if docente_classe_materia[d][c_first][m] > 0]
        mat_str = "/".join(mat_doc) if mat_doc else "MAT"
        
        if solver.value(matrice_motoria[d, c_first, g, o]) == 1:
            is_motoria = True
            mat_str = "MOT(Pal)"
        if "MEN" in mat_doc and o in [8, 9] and (g == MAR or g == GIO):
            is_mensa = True
            mat_str = "MENSA"
            
        nomi_c = [NOMI_CLASSI[c] for c in classi_curr]
        condivisa_str = " (condivisa)" if len(nomi_c) > 1 else ""
        testo = f"Classe {' e '.join(nomi_c)}{condivisa_str} - {mat_str}"
        segments.append({"text": testo, "color": inline_black})
        
    # 2. Ricerca classi per Potenziamento
    classi_pot = [c for c in range(C) if solver.value(matrice_potenziamento[d, c, g, o]) == 1]
    if classi_pot:
        nomi_c = [NOMI_CLASSI[c] for c in classi_pot]
        condivisa_str = " (condivisa)" if len(nomi_c) > 1 else ""
        testo = f"Classe {' e '.join(nomi_c)}{condivisa_str} - POT"
        segments.append({"text": testo, "color": inline_red})
        
    # 3. Ricerca classi per Alternativa
    classi_alt = [c for c in range(C) if solver.value(matrice_alternativa[d, c, g, o]) == 1]
    if classi_alt:
        nomi_c = [NOMI_CLASSI[c] for c in classi_alt]
        condivisa_str = " (condivisa)" if len(nomi_c) > 1 else ""
        testo = f"Classe {' e '.join(nomi_c)}{condivisa_str} - ALT"
        segments.append({"text": testo, "color": inline_blue})
        
    # Stringa univoca per capire se due mezz'ore consecutive sono identiche
    can_str = "\n".join([s["text"] for s in segments])
    
    return segments, is_mensa, is_motoria, can_str

for d in range(D):
    # Excel supporta al massimo 31 caratteri per il nome del foglio
    nome_docente = docenti[d][INDEX_NOME][:31] 
    ws_doc = wb_docenti.create_sheet(title=nome_docente)
    
    # Dimensioni
    ws_doc.column_dimensions['A'].width = 13
    for i in range(5):
        ws_doc.column_dimensions[get_column_letter(2 + i)].width = 24
        
    # Giorni della settimana
    for g in range(G):
        cell = ws_doc.cell(row=1, column=g + 2, value=NOMI_GIORNI[g])
        cell.font = font_bold
        cell.alignment = align_center
        cell.border = border_thin
        
    for o in range(O):
        ws_doc.row_dimensions[o + 2].height = 28
        c_orario = ws_doc.cell(row=o + 2, column=1, value=FASCE_ORARIE[o])
        c_orario.alignment = align_center
        c_orario.border = border_thin

    # Griglia Docente
    for g in range(G):
        # Calcoliamo prima cosa fa il docente in ogni mezz'ora di quel giorno
        schedule_giorno = [ottieni_attivita_docente(d, g, o) for o in range(O)]
        
        start_o = 0
        while start_o < O:
            segments, is_mensa, is_motoria, can_str = schedule_giorno[start_o]
            end_o = start_o
            
            # Scorriamo in avanti per raggruppare visivamente le mezz'ore consecutive uguali 
            while end_o + 1 < O and schedule_giorno[end_o + 1][3] == can_str and schedule_giorno[end_o + 1][1] == is_mensa and schedule_giorno[end_o + 1][2] == is_motoria:
                end_o += 1
            
            # Formattazione del Rich Text o cella vuota
            if not segments:
                main_cell = ws_doc.cell(row=start_o + 2, column=g + 2, value="")
            else:
                blocks_rich = []
                for i, seg in enumerate(segments):
                    txt = seg["text"]
                    # Aggiungiamo l'a capo se c'è un elemento successivo
                    if i < len(segments) - 1:
                        txt += "\n"
                    blocks_rich.append(TextBlock(seg["color"], txt))
                
                rich_text = CellRichText(blocks_rich)
                main_cell = ws_doc.cell(row=start_o + 2, column=g + 2, value=rich_text)
                
            main_cell.alignment = align_center
            
            if is_mensa:
                main_cell.fill = fill_mensa
            elif is_motoria:
                main_cell.fill = fill_motoria

            for o in range(start_o, end_o + 1):
                ws_doc.cell(row=o + 2, column=g + 2).border = border_thin
            
            if start_o != end_o:
                ws_doc.merge_cells(start_row=start_o + 2, start_column=g + 2, end_row=end_o + 2, end_column=g + 2)
            
            start_o = end_o + 1

wb_docenti.save("Orario_per_Docente_4.xlsx")

print("File 'Orario_Definitivo_Adaptive.xlsx' e 'orario_per_docente.xlsx' generati con successo!")