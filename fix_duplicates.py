import pandas as pd
import csv
import io
import sys
import collections

def fix_duplicates_all_matters():
    input_file = 'BDD CRPE RUSH 01.26.xlsx'
    print(f"Reading {input_file}...")

    try:
        df = pd.read_excel(input_file, header=None)
    except Exception as e:
        print(f"Error reading file: {e}")
        sys.exit(1)

    print(f"Total rows: {len(df)}")

    # We will track questions seen, categorized by Matter.
    # Map: matter -> { question_text -> first_row_index }
    # This ensures we only consider duplicates *within the same matter*.
    # (A question about "Louis XIV" in History is distinct from a similar text in Math, though unlikely).
    seen_questions_by_matter = collections.defaultdict(dict)

    rows_to_clear = []
    duplicates_count_by_matter = collections.defaultdict(int)

    # Iterate and Identify
    for idx, row in df.iterrows():
        cell_value = row[0]
        if pd.isna(cell_value):
            continue

        # Skip empty strings (previously cleared rows)
        if str(cell_value).strip() == "":
            continue

        try:
            # Parse CSV
            parsed = next(csv.reader(io.StringIO(str(cell_value))))
        except Exception as e:
            continue

        if len(parsed) < 8:
            continue

        # Extract Matter (index 7) and Question (index 1)
        matter = parsed[7].strip().lower() # Normalize matter
        submatter = parsed[6].strip()
        question = parsed[1].strip()

        if not matter or not question:
            continue

        # Check for duplicate in this specific matter
        if question in seen_questions_by_matter[matter]:
            # Duplicate found!
            original_idx = seen_questions_by_matter[matter][question]
            rows_to_clear.append(idx)
            duplicates_count_by_matter[matter] += 1

            # Log detail
            # print(f"Duplicate in '{matter}' (Submatter: {submatter}):")
            # print(f"   - Row {idx}: {question[:60]}...")
            # print(f"   - Original at Row {original_idx}")
        else:
            seen_questions_by_matter[matter][question] = idx

    print("-" * 40)
    print("Duplicate Summary by Matter:")
    for matter, count in duplicates_count_by_matter.items():
        print(f"   - {matter}: {count} duplicates found.")
    print(f"Total duplicates to remove: {len(rows_to_clear)}")
    print("-" * 40)

    if not rows_to_clear:
        print("No duplicates found. Exiting.")
        return

    # Process removals
    print(" removing duplicates...")
    for idx in rows_to_clear:
        # Clear the cell
        df.iloc[idx, 0] = "" # Set to empty string

    # Save
    print("Saving changes...")
    output_file = 'BDD CRPE RUSH 01.26.xlsx'
    df.to_excel(output_file, index=False, header=False)
    print("Done.")

if __name__ == "__main__":
    fix_duplicates_all_matters()
