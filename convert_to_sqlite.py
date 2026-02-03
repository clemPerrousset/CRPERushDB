import pandas as pd
import sqlite3
import csv
import io
import os

def convert_excel_to_sqlite(excel_path, db_path, table_name='quiz_card'):
    print(f"Reading {excel_path}...")
    # On lit le fichier Excel sans header car la structure semble complexe
    df = pd.read_excel(excel_path, header=None)

    # Suppression de l'ancienne DB pour repartir au propre
    if os.path.exists(db_path):
        os.remove(db_path)

    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    # --- MODIFICATION IMPORTANTE ICI ---
    # 1. id_card est maintenant TEXT
    # 2. Ajout de NOT NULL partout pour satisfaire Room
    create_table_sql = f"""
    CREATE TABLE IF NOT EXISTS {table_name} (
        id_card TEXT NOT NULL PRIMARY KEY,
        question TEXT NOT NULL,
        answer TEXT NOT NULL,
        fake_answer_one TEXT NOT NULL,
        fake_answer_two TEXT NOT NULL,
        theme TEXT NOT NULL,
        submatter TEXT NOT NULL,
        matter TEXT NOT NULL,
        explanation TEXT NOT NULL
    );
    """
    cursor.execute(create_table_sql)

    # Récupération du header pour comparaison (ligne 3, colonne 1 selon ton script original)
    try:
        header_row = df.iloc[2, 0]
        header = next(csv.reader(io.StringIO(str(header_row))))
    except Exception as e:
        print(f"Error reading header: {e}")
        return

    seen_ids = set()
    rows_to_insert = []

    print("Processing rows...")
    for idx, row in df.iterrows():
        cell_value = row[0]
        if pd.isna(cell_value):
            continue

        try:
            # Parsing du contenu CSV dans la cellule Excel
            parsed = next(csv.reader(io.StringIO(str(cell_value))))
        except Exception as e:
            continue

        # On saute la ligne si c'est le header
        if parsed == header:
            continue
        
        # Gestion des lignes mal formées (longueur incorrecte)
        if len(parsed) != 9:
            if len(parsed) > 9:
                # Si trop long, on fusionne le surplus dans l'explication
                expl = ",".join(parsed[8:])
                parsed = parsed[:8] + [expl]
            else:
                # Si trop court, on ignore (ou on pourrait padder avec des vides)
                # print(f"Skipping row {idx} (Len {len(parsed)})")
                continue

        # --- MODIFICATION DE L'ID ---
        try:
            # On nettoie l'ID : on le convertit en int d'abord pour enlever les ".0" éventuels, puis en string
            # Ex: "1.0" -> 1 -> "1"
            clean_id_int = int(float(parsed[0])) 
            row_id = str(clean_id_int) 
        except:
            # Si l'ID n'est pas un nombre valide, on saute
            continue

        if row_id in seen_ids:
            continue

        seen_ids.add(row_id)

        # Création du tuple à insérer
        # On force str() sur tout pour garantir le type TEXT
        data_tuple = (
            row_id,
            str(parsed[1]),
            str(parsed[2]),
            str(parsed[3]),
            str(parsed[4]),
            str(parsed[5]),
            str(parsed[6]),
            str(parsed[7]),
            str(parsed[8])
        )
        rows_to_insert.append(data_tuple)

    insert_sql = f"""
    INSERT INTO {table_name}
    (id_card, question, answer, fake_answer_one, fake_answer_two, theme, submatter, matter, explanation)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    """

    cursor.executemany(insert_sql, rows_to_insert)
    conn.commit()
    conn.close()

    print(f"Successfully inserted {len(rows_to_insert)} rows into '{table_name}' table.")
    print(f"Database created at: {db_path}")

if __name__ == "__main__":
    # Assure-toi que le nom du fichier Excel est correct
    convert_excel_to_sqlite('BDD CRPE RUSH 01.26.xlsx', 'quiz.db')