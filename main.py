from smtplib import SMTPHeloError

import pandas as pd
from collections import Counter
from itertools import combinations

INPUT_XLSX = 'assets/numbers.xlsx'
SHEET_NAME = 0
RANGE_MIN, RANGE_MAX = 1, 41
K_TOP = 20
OUTPUT_XLSX = 'assets/results.xlsx'

df = pd.read_excel(INPUT_XLSX, sheet_name=SHEET_NAME)

prefers = ["A", "B", "C", "D", "E", "F"]
if all(c in df.columns for c in prefers):
    cols = prefers
else:
    num_cols = [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]
    if len(num_cols) < 6:
        raise ValueError("Not enough numeric columns found in the Excel file.")
    cols = num_cols[:6]

data = df[cols].copy()

# === Cleaning ===
# Remove empty rows
data = data.dropna(how="all")
# Check for any NaN values in the selected columns
has_nan = data.isna().any(axis=1)

# Convert to numeric, coercing errors to NaN
for c in cols:
    data[c] = pd.to_numeric(data[c], errors="coerce")


# === Validations ===
def row_issues(row):
    vals = [v for v in row if pd.notna(v)]
    issues = []
    if len(vals) != 6:
        issues.append(f"Expected 6 numbers, got {len(vals)}")
        return issues
    # integers
    if any(float(v) != int(v) for v in vals):
        issues.append("Non-integer value found")
    vals_int = [int(v) for v in vals]
    # range check
    if any(v < RANGE_MIN or v > RANGE_MAX for v in vals_int):
        issues.append(f"Values must be between {RANGE_MIN} and {RANGE_MAX}")
    # duplicates
    if len(set(vals_int)) != 6:
        issues.append("Duplicate values found")
    return issues


issues_list = data[cols].apply(lambda r: row_issues(r.values), axis=1)
is_valid = issues_list.apply(lambda x: len(x) == 0)

valid = data.loc[is_valid, cols].copy()

# Ordering each row to detect consecutive numbers
valid_sorted = valid.apply(lambda r: sorted([int(x) for x in r.values]), axis=1, result_type="expand")
valid_sorted.columns = cols

# === Frequency Analysis ===
all_numbers = valid_sorted.to_numpy().ravel()
freq_num = Counter(all_numbers)
freq_num_df = (pd.DataFrame(freq_num.items(), columns=["Number", "Frequency"])
               .sort_values(["Frequency", "Number"], ascending=[False, True])
               .reset_index(drop=True))

# === Most common pairs and trios ===
pair_counter = Counter()
trio_counter = Counter()

for row in valid_sorted.to_numpy():
    pair_counter.update(combinations(row, 2))
    trio_counter.update(combinations(row, 3))

pairs_df = (pd.DataFrame([(a, b, c) for (a, b), c in pair_counter.items()], columns=["a", "b", "Frequency"])
            .sort_values(["Frequency", "a", "b"], ascending=[False, True, True])
            .head(K_TOP)
            .reset_index(drop=True))

trios_df = (pd.DataFrame([(a,b,c,d) for (a,b,c),d in trio_counter.items()],
                         columns=["a","b","c","frequency"])
            .sort_values(["frequency","a","b","c"], ascending=[False, True, True, True])
            .head(K_TOP)
            .reset_index(drop=True))

# === Consecutive (pairs k, k+1) ===
def runs_in_row(sorted_row):
    runs = []
    run = [sorted_row[0]]
    for x in sorted_row[1:]:
        if x == run[-1] + 1:
            run.append(x)
        else:
            if len(run) >= 2:
                runs.append(tuple(run))
            run = [x]
    if len(run) >= 2:
        runs.append(tuple(run))
    return runs

consec_pair_counter = Counter()
run_counter = Counter()

max_run_len = []
num_consec_pairs = []

for row in valid_sorted.to_numpy():
    c_pairs = 0
    for a, b in zip(row, row[1:]):
        if b == a + 1:
            consec_pair_counter[(a, b)] += 1
            c_pairs += 1
    num_consec_pairs.append(c_pairs)

    runs = runs_in_row(row)
    for r in runs:
        run_counter[r] += 1
    max_run_len.append(max([len(r) for r in runs], default=1))

consec_pairs_df = (pd.DataFrame([(a,b,c) for (a,b),c in consec_pair_counter.items()],
                                columns=["a","b","frequency"])
                   .sort_values(["frequency","a","b"], ascending=[False, True, True])
                   .head(K_TOP)
                   .reset_index(drop=True))

runs_df = (pd.DataFrame([(" - ".join(map(str, r)), len(r), c) for r,c in run_counter.items()],
                        columns=["corrida","longitud","conteo"])
           .sort_values(["conteo","longitud","corrida"], ascending=[False, False, True])
           .head(K_TOP)
           .reset_index(drop=True))

# === Metrics by row ===
row_metrics = valid_sorted.copy()
row_metrics["suma"] = valid_sorted.sum(axis=1)
row_metrics["min"] = valid_sorted.min(axis=1)
row_metrics["max"] = valid_sorted.max(axis=1)
row_metrics["rango"] = row_metrics["max"] - row_metrics["min"]
row_metrics["pares"] = valid_sorted.apply(lambda r: sum(v % 2 == 0 for v in r.values), axis=1)
row_metrics["impares"] = 6 - row_metrics["pares"]
row_metrics["num_pares_consecutivos"] = num_consec_pairs
row_metrics["max_corrida_consecutiva"] = max_run_len

# === Output to Excel ===
invalid_report = df.loc[~is_valid].copy()
invalid_report["problems"] = issues_list.loc[~is_valid].apply(lambda x: ", ".join(x))

with pd.ExcelWriter(OUTPUT_XLSX, engine="openpyxl") as writer:
    freq_num_df.to_excel(writer, index=False, sheet_name="Frecuencia_numeros")
    pairs_df.to_excel(writer, index=False, sheet_name="Top_pares")
    trios_df.to_excel(writer, index=False, sheet_name="Top_trios")
    consec_pairs_df.to_excel(writer, index=False, sheet_name="Top_consecutivos")
    runs_df.to_excel(writer, index=False, sheet_name="Top_corridas")
    row_metrics.to_excel(writer, index=False, sheet_name="Metricas_por_fila")
    invalid_report.to_excel(writer, index=False, sheet_name="Filas_invalidas")

print(f"Ready. Saved results to: {OUTPUT_XLSX}")
print(f"Valid rows: {len(valid_sorted)} | Invalid rows: {len(df) - len(valid_sorted)}")