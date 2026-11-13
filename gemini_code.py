import pandas as pd
from ortools.sat.python import cp_model

def genera_orario():
    file_path = "test_input_adaptive.xlsx"
    df_in = pd.read_excel(file_path, sheet_name='test_input_adaptive', header=None)

    # 1. Parsing classi
    classes = []
    for i in range(1, 6):
        row = df_in.iloc[i]
        if pd.isna(row[0]): break
        grado = str(int(row[0]))
        sezioni = [s.strip() for s in str(row[1]).split(",")]
        for s in sezioni: classes.append(grado+s)

    # 2. Parsing docenti
    teachers = {}
    for i in range(8, 24):
        row = df_in.iloc[i]
        if pd.isna(row[0]): continue
        t = row[0]
        avail = [row[1], row[2], row[3], row[4], row[5]]
        teachers[t] = {
            'avail': [1 if str(a).strip().upper() == 'SI' else 0 for a in avail],
            'part_time': str(row[7]).strip().upper() == 'SI'
        }

    # 3. Parsing ore e materie
    class_assignments = {c: [] for c in classes}
    start_row = 25
    while start_row < 98:
        row = df_in.iloc[start_row]
        c_name = row[0]
        if c_name in classes:
            subs = []
            for col_idx in range(1, 15):
                val = row[col_idx]
                if pd.isna(val): break
                # Normalizzazione in maiuscolo per evitare discrepanze
                subs.append(str(val).split("-")[0].strip().upper())
            
            t_row = start_row + 1
            while t_row < 98 and df_in.iloc[t_row][0] not in classes and df_in.iloc[t_row][0] != 'Orario accesso palestra':
                t_name = df_in.iloc[t_row][0]
                for col_idx, sub in enumerate(subs):
                    hrs = df_in.iloc[t_row][col_idx + 1]
                    if not pd.isna(hrs) and float(hrs) > 0:
                        class_assignments[c_name].append({'teacher': t_name, 'subject': sub, 'hours': int(hrs)})
                t_row += 1
            start_row = t_row
        else:
            start_row += 1

    # Parsing Palestra
    gym_access = {c: {d: [] for d in range(5)} for c in classes}
    row_idx = 99
    while row_idx < 109:
        row = df_in.iloc[row_idx]
        c = row[0]
        if c in classes:
            for d_idx, col_idx in enumerate([1, 2, 3, 4, 5]):
                val = row[col_idx]
                if not pd.isna(val) and isinstance(val, str):
                    slots = val.split(",")
                    parsed = []
                    for s in slots:
                        start, end = s.split("-")
                        start_h, start_m = map(int, start.split(":"))
                        end_h, end_m = map(int, end.split(":"))
                        start_slot = (start_h - 8) * 2 + (1 if start_m == 30 else 0) - 1
                        end_slot = (end_h - 8) * 2 + (1 if end_m == 30 else 0) - 1
                        parsed.extend(list(range(start_slot, end_slot)))
                    gym_access[c][d_idx] = parsed
        row_idx += 1

    # Parsing Classi abbinate
    paired_classes = []
    row_idx = 111
    while row_idx < 114:
        row = df_in.iloc[row_idx]
        if not pd.isna(row[0]) and not pd.isna(row[1]) and not pd.isna(row[2]):
            c1, c2, sub = str(row[0]).strip().upper(), str(row[1]).strip().upper(), str(row[2]).strip().upper()
            paired_classes.append((c1, c2, sub))
            # Se la religione è congiunta, l'alternativa è forzatamente congiunta per necessità temporale
            if sub == 'RELIGIONE':
                paired_classes.append((c1, c2, 'ALTERNATIVA'))
        row_idx += 1

    # Parsing POT
    specific_pot = []
    row_idx = 115
    while row_idx < len(df_in):
        row = df_in.iloc[row_idx]
        if not pd.isna(row[0]) and not pd.isna(row[1]) and not pd.isna(row[2]):
            specific_pot.append((str(row[0]).strip().upper(), str(row[1]).strip().upper(), str(row[2]).strip()))
        row_idx += 1

    # 4. Modello CP-SAT
    model = cp_model.CpModel()
    DAYS = 5
    SLOTS = 16 
    
    active_slots = {c: {d: [] for d in range(DAYS)} for c in classes}
    blocks = {c: {d: [] for d in range(DAYS)} for c in classes}
    
    for c in classes:
        is_small = int(c[0]) <= 3
        for d in range(DAYS):
            if d in [1, 3]: # MAR, GIO 
                active_slots[c][d] = list(range(16))
                blocks[c][d] = [(i, i+1) for i in range(0, 16, 2)]
            else:
                if is_small:
                    if d in [0, 4]: 
                        active_slots[c][d] = list(range(9))
                        blocks[c][d] = [(0,1), (2,3), (5,6), (7,8)]
                    else: 
                        active_slots[c][d] = list(range(8))
                        blocks[c][d] = [(i, i+1) for i in range(0, 8, 2)]
                else:
                    active_slots[c][d] = list(range(10))
                    blocks[c][d] = [(i, i+1) for i in range(0, 10, 2)]

    # Variabili
    x = {}
    for c in classes:
        for assign in class_assignments[c]:
            t = assign['teacher']
            sub = assign['subject']
            for d in range(DAYS):
                if teachers[t]['avail'][d] == 1:
                    for s in active_slots[c][d]:
                        if sub == 'MOTORIA' and s not in gym_access[c][d]: continue
                        if sub == 'MENSA' and not (d in [1, 3] and s in [8, 9]): continue
                        x[(d, s, c, t, sub)] = model.NewBoolVar(f"x_{d}_{s}_{c}_{t}_{sub}")

    # Sincronizzazione blocchi 2 mezz'ore
    for c in classes:
        for d in range(DAYS):
            for (s1, s2) in blocks[c][d]:
                for a in class_assignments[c]:
                    k1 = (d, s1, c, a['teacher'], a['subject'])
                    k2 = (d, s2, c, a['teacher'], a['subject'])
                    if k1 in x and k2 in x: model.Add(x[k1] == x[k2])
                    elif k1 in x: model.Add(x[k1] == 0)
                    elif k2 in x: model.Add(x[k2] == 0)

    # Mensa obbligatoria
    for c in classes:
        for d in [1, 3]:
            mensa_vars = [x[(d, 8, c, a['teacher'], 'MENSA')] for a in class_assignments[c] if a['subject'] == 'MENSA' and (d, 8, c, a['teacher'], 'MENSA') in x]
            if mensa_vars: model.AddExactlyOne(mensa_vars)

    # 1 lezione curricolare per slot (No sovrapposizione materie)
    for c in classes:
        for d in range(DAYS):
            for s in active_slots[c][d]:
                lesson_vars = []
                for a in class_assignments[c]:
                    if a['subject'] not in ['POT.', 'ALTERNATIVA']:
                        k = (d, s, c, a['teacher'], a['subject'])
                        if k in x: lesson_vars.append(x[k])
                model.AddExactlyOne(lesson_vars)

    # Rispetto monte ore totale per ogni docente in ogni classe
    for c in classes:
        for a in class_assignments[c]:
            k_vars = []
            for d in range(DAYS):
                for s in active_slots[c][d]:
                    k = (d, s, c, a['teacher'], a['subject'])
                    if k in x: k_vars.append(x[k])
            model.Add(sum(k_vars) == a['hours'] * 2)

    # Alternativa sincrona a Religione
    for c in classes:
        t_rel_list = [a['teacher'] for a in class_assignments[c] if a['subject'] == 'RELIGIONE']
        t_alt_list = [a['teacher'] for a in class_assignments[c] if a['subject'] == 'ALTERNATIVA']
        if t_rel_list and t_alt_list:
            t_rel, t_alt = t_rel_list[0], t_alt_list[0]
            for d in range(DAYS):
                for s in active_slots[c][d]:
                    k_rel = (d, s, c, t_rel, 'RELIGIONE')
                    k_alt = (d, s, c, t_alt, 'ALTERNATIVA')
                    v_rel = x[k_rel] if k_rel in x else 0
                    v_alt = x[k_alt] if k_alt in x else 0
                    model.Add(v_rel == v_alt)

    # Classi abbinate
    for (c1, c2, sub) in paired_classes:
        t_list = [a['teacher'] for a in class_assignments[c1] if a['subject'] == sub]
        if t_list:
            t = t_list[0]
            for d in range(DAYS):
                for s in range(SLOTS):
                    k1 = (d, s, c1, t, sub)
                    k2 = (d, s, c2, t, sub)
                    if k1 in x and k2 in x: model.Add(x[k1] == x[k2])
                    elif k1 in x: model.Add(x[k1] == 0)
                    elif k2 in x: model.Add(x[k2] == 0)

    # Nessuna sovrapposizione oraria per il Docente (tenendo conto delle abbinate)
    for d in range(DAYS):
        for s in range(SLOTS):
            for t in teachers:
                group_vars = []
                covered_paired = set()
                for c in classes:
                    if s in active_slots[c][d]:
                        for a in class_assignments[c]:
                            if a['teacher'] == t:
                                k = (d, s, c, t, a['subject'])
                                if k in x:
                                    paired_group = next((p for p in paired_classes if a['subject'] == p[2] and c in (p[0], p[1])), None)
                                    if paired_group:
                                        if paired_group not in covered_paired:
                                            group_vars.append(x[k])
                                            covered_paired.add(paired_group)
                                    else:
                                        group_vars.append(x[k])
                if group_vars: model.Add(sum(group_vars) <= 1)

    # Potenziamento imposto in contemporanea
    for (c, sub, t_pot) in specific_pot:
        t_sub_list = [a['teacher'] for a in class_assignments[c] if a['subject'] == sub]
        if t_sub_list:
            t_sub = t_sub_list[0]
            for d in range(DAYS):
                for s in active_slots[c][d]:
                    k_sub = (d, s, c, t_sub, sub)
                    k_pot = (d, s, c, t_pot, 'POT.')
                    v_sub = x[k_sub] if k_sub in x else 0
                    v_pot = x[k_pot] if k_pot in x else 0
                    model.Add(v_pot == v_sub)

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 180.0
    status = solver.Solve(model)

    if status in [cp_model.OPTIMAL, cp_model.FEASIBLE]:
        print("Soluzione trovata! Esportazione in orario_output.xlsx...")
        with pd.ExcelWriter('orario_output.xlsx') as writer:
            for c in classes:
                df_out = pd.DataFrame(index=range(SLOTS), columns=['LUN', 'MAR', 'MER', 'GIO', 'VEN'])
                for d in range(DAYS):
                    for s in active_slots[c][d]:
                        cell_text = []
                        for a in class_assignments[c]:
                            k = (d, s, c, a['teacher'], a['subject'])
                            if k in x and solver.Value(x[k]) == 1:
                                lbl = f"{a['subject']} ({a['teacher']})"
                                if a['subject'] in ['POT.', 'ALTERNATIVA']: lbl = f"[{lbl}]"
                                cell_text.append(lbl)
                        df_out.iloc[s, d] = " \n ".join(cell_text) if cell_text else ""
                df_out.to_excel(writer, sheet_name=f"Classe_{c}")
        print("Fatto.")
    else:
        print("L'istanza risulta INFEASIBLE.")

if __name__ == '__main__':
    genera_orario()