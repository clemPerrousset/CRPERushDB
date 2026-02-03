import pandas as pd
import csv
import io
import sys

def fix_duplicates():
    input_file = 'BDD CRPE RUSH 01.26.xlsx'
    print(f"Reading {input_file}...")

    try:
        df = pd.read_excel(input_file, header=None)
    except Exception as e:
        print(f"Error reading file: {e}")
        sys.exit(1)

    print(f"Total rows: {len(df)}")

    # We will track questions seen in the Math section
    # Map: question_text -> first_row_index
    seen_math_questions = {}

    rows_to_clear = []

    # Iterate and Identify
    for idx, row in df.iterrows():
        cell_value = row[0]
        if pd.isna(cell_value):
            continue

        try:
            # Parse CSV
            # cell_value needs to be a string
            parsed = next(csv.reader(io.StringIO(str(cell_value))))
        except Exception as e:
            # parsing error, skip
            continue

        if len(parsed) < 8:
            continue

        # Check if it is Math
        matter = parsed[7].strip().lower()
        if 'math' in matter:
            question = parsed[1].strip()

            if question in seen_math_questions:
                # Duplicate found!
                original_idx = seen_math_questions[question]
                rows_to_clear.append(idx)
            else:
                seen_math_questions[question] = idx

    print(f"Found {len(rows_to_clear)} duplicates to remove.")

    if not rows_to_clear:
        print("No duplicates found. Exiting.")
        return

    # Process removals
    for idx in rows_to_clear:
        # Get the row content before clearing for logging
        cell_value = df.iloc[idx, 0]
        parsed = next(csv.reader(io.StringIO(str(cell_value))))
        q_text = parsed[1]

        # Verify correctness of the *original* (kept) entry
        original_idx = seen_math_questions[q_text.strip()]
        orig_cell = df.iloc[original_idx, 0]
        orig_parsed = next(csv.reader(io.StringIO(str(orig_cell))))

        print(f"Removing duplicate at row {idx}.")
        print(f"   - Question: {q_text[:60]}...")
        print(f"   - Kept Entry (Row {original_idx}): Answer = {orig_parsed[2]}")

        # Clear the cell
        df.iloc[idx, 0] = "" # Set to empty string

    # Save
    print("Saving changes...")
    output_file = 'BDD CRPE RUSH 01.26.xlsx'
    # index=False, header=False to match input format (no index column, no header row)
    df.to_excel(output_file, index=False, header=False)
    print("Done.")

if __name__ == "__main__":
    fix_duplicates()
