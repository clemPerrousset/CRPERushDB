import pandas as pd
import sqlite3
import csv
import io
import os

def convert_excel_to_sqlite(excel_path, db_path, table_name='cards'):
    print(f"Reading {excel_path}...")
    df = pd.read_excel(excel_path, header=None)

    if os.path.exists(db_path):
        os.remove(db_path)

    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    create_table_sql = f"""
    CREATE TABLE IF NOT EXISTS {table_name} (
        id_card INTEGER PRIMARY KEY,
        question TEXT,
        answer TEXT,
        fake_answer_one TEXT,
        fake_answer_two TEXT,
        theme TEXT,
        submatter TEXT,
        matter TEXT,
        explanation TEXT
    );
    """
    cursor.execute(create_table_sql)

    header_row = df.iloc[2, 0]
    header = next(csv.reader(io.StringIO(header_row)))

    seen_ids = set()
    rows_to_insert = []

    for idx, row in df.iterrows():
        cell_value = row[0]
        if pd.isna(cell_value):
            continue

        try:
            parsed = next(csv.reader(io.StringIO(str(cell_value))))
        except Exception as e:
            continue

        if parsed == header:
            continue

        # Relaxed parsing: if len > 9, try to take first 9? No, that was previous step.
        # But we still have some anomalies that my structural fix script didn't catch
        # (maybe because I ran it on a separate CSV export and re-imported?).
        # Wait, I created `BDD_CRPE_RUSH_CORRIGE.xlsx` by writing `final_rows`.
        # Ah, I see: "Skipping row 406 due to length 2".
        # This means my structural fix script FAILED to fix row 406 properly or serialized it poorly?
        # Row 406 in my dump was: ['203', 'Quelle est la somme...'] (Length 2?)

        # Actually, let's just skip bad rows for the DB conversion or try to salvage?
        # The user wants a clean DB.

        if len(parsed) != 9:
            # Try to fix "on the fly" if easy?
            # If length is 10 and last part is part of explanation...
            if len(parsed) > 9:
                # Merge explanation
                expl = ",".join(parsed[8:])
                parsed = parsed[:8] + [expl]
            else:
                print(f"Skipping row {idx} (Len {len(parsed)}): {parsed}")
                continue

        try:
            row_id = int(parsed[0])
        except:
            continue

        if row_id in seen_ids:
            # Duplicate ID found.
            # print(f"Duplicate ID {row_id} at row {idx}. Generating new ID.")
            # Strategy: Find max ID and increment? Or just skip?
            # Or use autoincrement in DB and ignore provided ID?
            # The ID might be important for the app.
            # Let's Skip duplicates to avoid IntegrityError
            continue

        seen_ids.add(row_id)

        data_tuple = (
            row_id,
            parsed[1],
            parsed[2],
            parsed[3],
            parsed[4],
            parsed[5],
            parsed[6],
            parsed[7],
            parsed[8]
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
    convert_excel_to_sqlite('BDD CRPE RUSH 01.26.xlsx', 'crpe.db')
