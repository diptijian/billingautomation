from pathlib import Path
import pandas as pd


E_BASE_DIR = Path("input/t3/e_base")
X_BASE_DIR = Path("input/t3/x_base")
OUTPUT_DIR = Path("output/t3")


X_COLUMNS = [
    "Serving DC Name",
    "Lower Limit",
    "Upper Limit",
    "Proposed CPK",
    "Weight",
    "Cost",
    "",
    "Serving DC Code",
    "Total input weight",
    "Total Cost",
    "Average CPK",
]


def get_excel_file(folder: Path) -> Path:
    files = (
        list(folder.glob("*.xlsx"))
        + list(folder.glob("*.xlsb"))
        + list(folder.glob("*.xls"))
    )

    if len(files) == 0:
        raise FileNotFoundError(f"No Excel file found in {folder}")

    if len(files) > 1:
        raise ValueError(f"Multiple Excel files found in {folder}. Keep only one.")

    return files[0]


def read_excel_file(path: Path, **kwargs) -> pd.DataFrame:
    if path.suffix.lower() == ".xlsb":
        return pd.read_excel(path, engine="pyxlsb", **kwargs)

    return pd.read_excel(path, **kwargs)


def clean_header(value) -> str:
    return str(value).replace("\xa0", " ").strip()


def clean_text(series: pd.Series) -> pd.Series:
    return (
        series.fillna("")
        .astype(str)
        .str.replace("\xa0", " ", regex=False)
        .str.strip()
    )


def normalize_key(series: pd.Series) -> pd.Series:
    return clean_text(series).str.upper()


def find_col(df: pd.DataFrame, possible_names: list[str]) -> str:
    normalized_columns = {
        clean_header(col).lower(): col
        for col in df.columns
    }

    for name in possible_names:
        key = clean_header(name).lower()
        if key in normalized_columns:
            return normalized_columns[key]

    raise KeyError(
        "Could not find any of these columns: "
        + ", ".join(possible_names)
        + "\nAvailable columns: "
        + ", ".join(map(str, df.columns))
    )


def standardize_x_columns(x: pd.DataFrame) -> pd.DataFrame:
    x = x.copy()

    x.columns = [clean_header(col) for col in x.columns]

    while len(x.columns) < len(X_COLUMNS):
        x[f"_blank_{len(x.columns)}"] = pd.NA

    cols = list(x.columns)

    for idx, col_name in enumerate(X_COLUMNS):
        cols[idx] = col_name

    x.columns = cols

    return x


def allocate_slab_weight(total_weight, lower_limit, upper_limit) -> float:
    total_weight = float(total_weight or 0)
    lower_limit = float(lower_limit or 0)
    upper_limit = float(upper_limit or 0)

    slab_capacity = upper_limit - lower_limit

    if slab_capacity <= 0:
        return 0.0

    allocated_weight = min(max(total_weight - lower_limit, 0), slab_capacity)

    return max(allocated_weight, 0.0)


def generate_x():
    e_base_path = get_excel_file(E_BASE_DIR)
    x_base_path = get_excel_file(X_BASE_DIR)

    e = read_excel_file(e_base_path)
    x = read_excel_file(x_base_path)

    e.columns = (
        e.columns.astype(str)
        .str.replace("\xa0", " ", regex=False)
        .str.strip()
    )

    x = standardize_x_columns(x)

    # -----------------------------
    # Step 1:
    # From E_Base:
    # Filter Hub Billing type = E2E
    # Group by Serving DC Name
    # Sum Chargeable Weight (Payable)
    # -----------------------------

    e_serving_dc_col = find_col(
        e,
        [
            "Serving DC Name",
            "Serv. DC Name",
            "Serving DC Code",
            "Serv. DC Code",
        ],
    )

    e_weight_col = find_col(
        e,
        [
            "Chargeable Weight (Payable)",
        ],
    )

    e_hub_billing_type_col = find_col(
        e,
        [
            "Hub Billing type",
            "Hub Billing Type",
        ],
    )

    e["_serving_dc_display"] = clean_text(e[e_serving_dc_col])
    e["_serving_dc_key"] = normalize_key(e[e_serving_dc_col])
    e["_hub_billing_type"] = normalize_key(e[e_hub_billing_type_col])
    e["_weight"] = pd.to_numeric(e[e_weight_col], errors="coerce").fillna(0)

    e2e = e[
        (e["_hub_billing_type"] == "E2E")
        & (e["_serving_dc_key"] != "")
        & (e["_weight"] > 0)
    ].copy()

    summary = (
        e2e.groupby("_serving_dc_key", as_index=False)
        .agg(
            serving_dc_display=("_serving_dc_display", "first"),
            total_input_weight=("_weight", "sum"),
        )
    )

    # -----------------------------
    # Prepare X columns
    # -----------------------------

    x["Weight"] = 0.0
    x["Cost"] = 0.0
    x["Serving DC Code"] = ""
    x["Total input weight"] = pd.NA
    x["Total Cost"] = pd.NA
    x["Average CPK"] = pd.NA

    x["_dc_display"] = clean_text(x["Serving DC Name"])
    x["_dc_key"] = normalize_key(x["Serving DC Name"])

    x["_lower"] = pd.to_numeric(x["Lower Limit"], errors="coerce").fillna(0)
    x["_upper"] = pd.to_numeric(x["Upper Limit"], errors="coerce").fillna(0)
    x["_cpk"] = pd.to_numeric(x["Proposed CPK"], errors="coerce").fillna(0)

    summary_map = dict(
        zip(summary["_serving_dc_key"], summary["total_input_weight"])
    )

    summary_display_map = dict(
        zip(summary["_serving_dc_key"], summary["serving_dc_display"])
    )

    # Keep right-side H/I order based on X_Base hub order.
    # Then append any E_Base DCs not present in X_Base.
    x_dc_order = x.loc[x["_dc_key"] != "", "_dc_key"].drop_duplicates().tolist()

    ordered_dc_keys = []

    for dc_key in x_dc_order:
        if dc_key in summary_map and dc_key not in ordered_dc_keys:
            ordered_dc_keys.append(dc_key)

    for dc_key in summary_map:
        if dc_key not in ordered_dc_keys:
            ordered_dc_keys.append(dc_key)

    # -----------------------------
    # Step 1 output:
    # Fill H/I:
    # Serving DC Code
    # Total input weight
    # -----------------------------

    for idx, dc_key in enumerate(ordered_dc_keys):
        while idx >= len(x):
            x.loc[len(x)] = [pd.NA] * len(x.columns)

        dc_display = summary_display_map.get(dc_key, dc_key)
        total_weight = summary_map.get(dc_key, 0)

        x.loc[idx, "Serving DC Code"] = dc_display
        x.loc[idx, "Total input weight"] = total_weight

    # -----------------------------
    # Step 2:
    # Allocate total input weight into slabs
    # Cost = allocated weight in kg * Proposed CPK
    # No *1000 because weight is already in kgs
    # -----------------------------

    total_cost_map = {}

    for dc_key in ordered_dc_keys:
        total_weight = float(summary_map.get(dc_key, 0) or 0)

        dc_rows = x[x["_dc_key"] == dc_key].sort_values(by="_lower")

        if dc_rows.empty:
            total_cost_map[dc_key] = 0.0
            continue

        dc_total_cost = 0.0

        for row_idx, row in dc_rows.iterrows():
            lower_limit = row["_lower"]
            upper_limit = row["_upper"]
            proposed_cpk = row["_cpk"]

            allocated_weight = allocate_slab_weight(
                total_weight,
                lower_limit,
                upper_limit,
            )

            cost = allocated_weight * proposed_cpk

            x.loc[row_idx, "Weight"] = allocated_weight
            x.loc[row_idx, "Cost"] = cost

            dc_total_cost += cost

        total_cost_map[dc_key] = dc_total_cost

    # -----------------------------
    # Step 3:
    # Fill J/K:
    # Total Cost
    # Average CPK = Total Cost / Total input weight
    # -----------------------------

    for idx, dc_key in enumerate(ordered_dc_keys):
        total_weight = float(summary_map.get(dc_key, 0) or 0)
        total_cost = float(total_cost_map.get(dc_key, 0) or 0)

        average_cpk = total_cost / total_weight if total_weight > 0 else 0

        x.loc[idx, "Total Cost"] = total_cost
        x.loc[idx, "Average CPK"] = average_cpk

    x.drop(
        columns=[
            "_dc_display",
            "_dc_key",
            "_lower",
            "_upper",
            "_cpk",
        ],
        inplace=True,
        errors="ignore",
    )

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    output_file = OUTPUT_DIR / "X.xlsx"
    x.to_excel(output_file, index=False)

    print("T3 complete: X generated")
    print(f"E_Base file: {e_base_path}")
    print(f"X_Base file: {x_base_path}")
    print(f"Rows in X: {len(x)}")
    print(f"Serving DCs calculated: {len(ordered_dc_keys)}")
    print(f"Output: {output_file}")


if __name__ == "__main__":
    generate_x()