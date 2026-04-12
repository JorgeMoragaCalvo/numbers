import pandas as pd
from collections import Counter
from itertools import combinations

INPUT_XLSX = 'assets/numbers.xlsx'
SHEET_NAME = 0
RANGE_MIN, RANGE_MAX = 1, 42
K_TOP = 20
N_RECENT = 20 # draws to consider for hot/cold analysis
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
data = data.dropna(how="all")
has_nan = data.isna().any(axis=1)

for c in cols:
    data[c] = pd.to_numeric(data[c], errors="coerce")


# === Validations ===
def row_issues(nums):
    vals = [v for v in nums if pd.notna(v)]
    issues = []
    if len(vals) != 6:
        issues.append(f"Expected 6 numbers, got {len(vals)}")
        return issues
    if any(float(v) != int(v) for v in vals):
        issues.append("Non-integer value found")
    vals_int = [int(v) for v in vals]
    if any(v < RANGE_MIN or v >= RANGE_MAX for v in vals_int):
        issues.append(f"Values must be between {RANGE_MIN} and {RANGE_MAX - 1}")
    if len(set(vals_int)) != 6:
        issues.append("Duplicate values found")
    return issues


issues_list = data[cols].apply(lambda r: row_issues(r.values), axis=1)
is_valid = issues_list.apply(lambda x: len(x) == 0)

valid = data.loc[is_valid, cols].copy()

valid_sorted = valid.apply(lambda r: sorted([int(x) for x in r.values]), axis=1, result_type="expand")
valid_sorted.columns = cols

# === Frequency Analysis ===
all_numbers = valid_sorted.to_numpy().ravel()
freq_num = Counter(all_numbers)
freq_num_df = (pd.DataFrame(freq_num.items(), columns=["Number", "Frequency"])
               .sort_values(["Frequency", "Number"], ascending=[False, True])
               .reset_index(drop=True))

# === Pairs, Trios, Quartets ===
pair_counter = Counter()
trio_counter = Counter()
quartet_counter = Counter()

for row in valid_sorted.to_numpy():
    pair_counter.update(combinations(row, 2))
    trio_counter.update(combinations(row, 3))
    quartet_counter.update(combinations(row, 4))

pairs_df = (pd.DataFrame([(a, b, c) for (a, b), c in pair_counter.items()], columns=["a", "b", "Frequency"])
            .sort_values(["Frequency", "a", "b"], ascending=[False, True, True])
            .head(K_TOP)
            .reset_index(drop=True))

trios_df = (pd.DataFrame([(a, b, c, d) for (a, b, c), d in trio_counter.items()],
                         columns=["a", "b", "c", "frequency"])
            .sort_values(["frequency", "a", "b", "c"], ascending=[False, True, True, True])
            .head(K_TOP)
            .reset_index(drop=True))

quartets_df = (pd.DataFrame([(a, b, c, d, e) for (a, b, c, d), e in quartet_counter.items()],
                             columns=["a", "b", "c", "d", "frequency"])
               .sort_values(["frequency", "a", "b", "c", "d"], ascending=[False, True, True, True, True])
               .head(K_TOP)
               .reset_index(drop=True))

# === Consecutive pairs (k, k+1) and runs ===
def runs_in_row(sorted_row):
    result = []
    run = [sorted_row[0]]
    for x in sorted_row[1:]:
        if x == run[-1] + 1:
            run.append(x)
        else:
            if len(run) >= 2:
                result.append(tuple(run))
            run = [x]
    if len(run) >= 2:
        result.append(tuple(run))
    return result


consecutive_pair_counter = Counter()
run_counter = Counter()
max_run_len = []
num_consecutive_pairs = []
delta_rows = []

for row in valid_sorted.to_numpy():
    c_pairs = 0
    for a, b in zip(row, row[1:]):
        if b == a + 1:
            consecutive_pair_counter[(a, b)] += 1
            c_pairs += 1
    num_consecutive_pairs.append(c_pairs)

    runs = runs_in_row(row)
    for run_tuple in runs:
        run_counter[run_tuple] += 1
    max_run_len.append(max([len(run_tuple) for run_tuple in runs], default=1))

    delta_rows.append([int(b) - int(a) for a, b in zip(row, row[1:])])

consecutive_pairs_df = (pd.DataFrame([(a, b, c) for (a, b), c in consecutive_pair_counter.items()],
                                columns=["a", "b", "frequency"])
                   .sort_values(["frequency", "a", "b"], ascending=[False, True, True])
                   .head(K_TOP)
                   .reset_index(drop=True))

runs_df = (pd.DataFrame([(" - ".join(map(str, run_key)), len(run_key), cnt) for run_key, cnt in run_counter.items()],
                        columns=["corrida", "longitud", "conteo"])
           .sort_values(["conteo", "longitud", "corrida"], ascending=[False, False, True])
           .head(K_TOP)
           .reset_index(drop=True))

# === Metrics by row ===
row_metrics = valid_sorted.copy()
row_metrics["suma"] = valid_sorted.sum(axis=1)
row_metrics["min"] = valid_sorted.min(axis=1)
row_metrics["max"] = valid_sorted.max(axis=1)
row_metrics["rango"] = row_metrics["max"] - row_metrics["min"]
row_metrics["pares"] = valid_sorted.apply(lambda series: sum(v % 2 == 0 for v in series.values), axis=1)
row_metrics["impares"] = 6 - row_metrics["pares"]
row_metrics["num_pares_consecutivos"] = num_consecutive_pairs
row_metrics["max_corrida_consecutiva"] = max_run_len

# Deltas between consecutive sorted numbers (5 gaps per row)
delta_df = pd.DataFrame(delta_rows,
                        columns=[f"delta_{i+1}" for i in range(5)],
                        index=valid_sorted.index)
row_metrics = pd.concat([row_metrics, delta_df], axis=1)

# Decade distribution (count of numbers in each band of 10)
def decade_counts(nums):
    return [
        sum(1  <= v <= 10 for v in nums),
        sum(11 <= v <= 20 for v in nums),
        sum(21 <= v <= 30 for v in nums),
        sum(31 <= v <= 40 for v in nums),
    ]

decade_df = valid_sorted.apply(
    lambda series: pd.Series(decade_counts(series.values),
                             index=["dec_1_10", "dec_11_20", "dec_21_30", "dec_31_40"]),
    axis=1
)
row_metrics = pd.concat([row_metrics, decade_df], axis=1)

# === Positional Frequency ===
pos_freq_df = pd.DataFrame({"Number": range(RANGE_MIN, RANGE_MAX)})
for i, col in enumerate(cols):
    counts = Counter(int(v) for v in valid_sorted[col])
    pos_freq_df[f"pos_{i+1}"] = pos_freq_df["Number"].map(counts).fillna(0).astype(int)

# === Hot / Cold Numbers (last N_RECENT draws vs. all-time) ===
recent_nums = Counter(int(v) for v in valid_sorted.tail(N_RECENT).to_numpy().ravel())
hot_cold_df = (freq_num_df.copy()
               .rename(columns={"Frequency": "freq_histórica"})
               .assign(freq_reciente=lambda d: d["Number"].map(lambda n: recent_nums.get(n, 0)))
               [["Number", "freq_histórica", "freq_reciente"]]
               .sort_values("freq_reciente", ascending=False)
               .reset_index(drop=True))

# === Number Aging (draws since last appearance) ===
last_seen = {}
for idx, row in enumerate(valid_sorted.to_numpy()):
    for num in row:
        last_seen[int(num)] = idx

total_draws = len(valid_sorted)
aging_df = pd.DataFrame([
    {"Number": n,
     "ultimo_sorteo_hace": (total_draws - 1 - last_seen[n]) if n in last_seen else total_draws}
    for n in range(RANGE_MIN, RANGE_MAX)
]).sort_values("ultimo_sorteo_hace", ascending=False).reset_index(drop=True)

# === Sum Distribution ===
vc = row_metrics["suma"].value_counts().sort_index()
sum_dist_df = pd.DataFrame({"suma": vc.index, "conteo": vc.values})

# === Affinity Matrix (co-occurrence counts for every number pair) ===
affinity = pd.DataFrame(0,
                        index=range(RANGE_MIN, RANGE_MAX),
                        columns=range(RANGE_MIN, RANGE_MAX),
                        dtype=int)
for row in valid_sorted.to_numpy():
    for a, b in combinations(row.tolist(), 2):
        affinity.at[a, b] += 1
        affinity.at[b, a] += 1
affinity.index.name = "Number"
affinity_df = affinity.reset_index()

# === Output to Excel ===
invalid_report = df.loc[~is_valid].copy()
invalid_report["problems"] = issues_list.loc[~is_valid].apply(lambda x: ", ".join(x))

with pd.ExcelWriter(OUTPUT_XLSX, engine="openpyxl") as writer:
    freq_num_df.to_excel(writer, index=False, sheet_name="Frecuencia_números")
    pairs_df.to_excel(writer, index=False, sheet_name="Top_pares")
    trios_df.to_excel(writer, index=False, sheet_name="Top_trios")
    quartets_df.to_excel(writer, index=False, sheet_name="Top_cuartetos")
    consecutive_pairs_df.to_excel(writer, index=False, sheet_name="Top_consecutivos")
    runs_df.to_excel(writer, index=False, sheet_name="Top_corridas")
    row_metrics.to_excel(writer, index=False, sheet_name="Métricas_por_fila")
    pos_freq_df.to_excel(writer, index=False, sheet_name="Frecuencia_posicional")
    hot_cold_df.to_excel(writer, index=False, sheet_name="Caliente_Frio")
    aging_df.to_excel(writer, index=False, sheet_name="Envejecimiento")
    sum_dist_df.to_excel(writer, index=False, sheet_name="Distribución_suma")
    affinity_df.to_excel(writer, index=False, sheet_name="Matriz_afinidad")
    invalid_report.to_excel(writer, index=False, sheet_name="Filas_invalidas")

print(f"Ready. Saved results to: {OUTPUT_XLSX}")
print(f"Valid rows: {len(valid_sorted)} | Invalid rows: {len(df) - len(valid_sorted)}")
