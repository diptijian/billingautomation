from pathlib import Path
import pandas as pd
import re


C_DIR = Path("input/t2/c_cleaned")
D_DIR = Path("input/t2/d_cleaned")
F_DIR = Path("input/t2/f")
G_DIR = Path("input/t2/g")
H_DIR = Path("input/t2/h")
OUTPUT_DIR = Path("output/t2")
HUB_MASTER_DIR = Path("input/t2/hub_master")


def get_excel_file(folder: Path) -> Path:
    files = list(folder.glob("*.xlsx")) + list(folder.glob("*.xlsb")) + list(folder.glob("*.xls"))
    if len(files) == 0:
        raise FileNotFoundError(f"No Excel file found in {folder}")
    if len(files) > 1:
        raise ValueError(f"Multiple Excel files found in {folder}. Keep only one.")
    return files[0]


def read_excel_file(path: Path, **kwargs) -> pd.DataFrame:
    if path.suffix.lower() == ".xlsb":
        return pd.read_excel(path, engine="pyxlsb", **kwargs)
    return pd.read_excel(path, **kwargs)


def s(df, col, default=""):
    return df[col] if col in df.columns else default

def normalize_partner_name(x):
    if pd.isna(x):
        return ""

    value = str(x).replace("\xa0", " ").upper().strip()

    # Convert punctuation/dots/commas/hyphens into spaces
    # Example: P.V.T. LTD. -> P V T LTD
    value = re.sub(r"[^A-Z0-9]+", " ", value)

    raw_tokens = value.split()

    # Combine dotted/spaced abbreviations:
    # P V T -> PVT
    # L T D -> LTD
    tokens = []
    i = 0

    while i < len(raw_tokens):
        if i + 2 < len(raw_tokens) and raw_tokens[i:i+3] == ["P", "V", "T"]:
            tokens.append("PVT")
            i += 3
        elif i + 2 < len(raw_tokens) and raw_tokens[i:i+3] == ["L", "T", "D"]:
            tokens.append("LTD")
            i += 3
        else:
            tokens.append(raw_tokens[i])
            i += 1

    normalized_tokens = []

    for token in tokens:
        if token in {"PVT", "PRIVATE"}:
            normalized_tokens.append("PRIVATE")
        elif token in {"LTD", "LIMITED"}:
            normalized_tokens.append("LIMITED")
        elif token == "PVTLTD":
            normalized_tokens.extend(["PRIVATE", "LIMITED"])
        else:
            normalized_tokens.append(token)

    # Remove trailing legal suffixes
    while normalized_tokens and normalized_tokens[-1] in {"PRIVATE", "LIMITED"}:
        normalized_tokens.pop()

    return " ".join(normalized_tokens).strip()

def clean_series(series):
    return (
        series.astype("string")
        .fillna("")
        .str.replace("\xa0", " ", regex=False)
        .str.strip()
        .str.replace(r"\.0$", "", regex=True)
    )


def make_key(a, b):
    return clean_series(a) + "_" + clean_series(b)


def normalize_cd_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    # Clean column headers
    df.columns = (
        df.columns.astype(str)
        .str.replace("\xa0", " ", regex=False)
        .str.strip()
    )

    # Canonical column name used by our existing T2 code:
    # possible names in old/new C' and D' files
    aliases = {
        "CN #": ["CN #", "Consignment Number", "Reference Number"],
        "Movement Type": ["Movement Type"],
        "Status(In Trip)": ["Status(In Trip)", "Status"],
        "Created At": ["Created At", "Created at"],
        "Origin Hub": ["Origin Hub", "Origin Hub Code"],
        "(Destination Hub)": ["(Destination Hub)", "Destination Hub", "Destination Hub Code"],
        "Sender Pincode": ["Sender Pincode", "Origin Pincode"],
        "Destination Pincode": ["Destination Pincode", "Consignee Pincode"],
        "Completed Time": ["Completed Time", "Last Pickup Completed Time"],
        "First Inscan at Hub Time": ["First Inscan at Hub Time", "First Inscan At Hub"],
        "Delivered Time": ["Delivered Time"],
        "Pickup Attempts": ["Pickup Attempts", "Pickup Attempt Count"],
        "Delivery Attempts": ["Delivery Attempts", "Attempt Count"],
        "Input Weight": ["Input Weight"],
        "Number Of Pieces": ["Number Of Pieces", "Num Pieces"],
        "Verified Chargeable Weight": ["Verified Chargeable Weight", "Chargeable Weight"],
        "Customer Reference Number": ["Customer Reference Number", "Customer Reference No."],
        "Movement Classification": ["Movement Classification"],
    }

    for canonical_name, possible_names in aliases.items():
        if canonical_name not in df.columns:
            for possible_name in possible_names:
                if possible_name in df.columns:
                    df[canonical_name] = df[possible_name]
                    break
    if "Movement Classification" not in df.columns:
        df["Movement Classification"] = ""

    required_columns = [
        "CN #",
        "Movement Type",
        "Created At",
        "Origin Hub",
        "(Destination Hub)",
        "Sender Pincode",
        "Destination Pincode",
        "Completed Time",
        "First Inscan at Hub Time",
        "Delivered Time",
        "Pickup Attempts",
        "Delivery Attempts",
        "Input Weight",
        "Number Of Pieces",
        "Verified Chargeable Weight",
        "Customer Reference Number",
    ]

    missing = [col for col in required_columns if col not in df.columns]

    if missing:
        raise ValueError(
            "C'/D' input file is missing required columns after normalization: "
            + ", ".join(missing)
            + "\nAvailable columns are: "
            + ", ".join(df.columns.astype(str))
        )

    return df


def generate_e_base():
    c = read_excel_file(get_excel_file(C_DIR))
    d = read_excel_file(get_excel_file(D_DIR))

    # Adapt old/new C' and D' formats into one internal format
    c = normalize_cd_columns(c)
    d = normalize_cd_columns(d)

    f = read_excel_file(get_excel_file(F_DIR))
    g = read_excel_file(get_excel_file(G_DIR))
    h_path = get_excel_file(H_DIR)

    h = read_excel_file(
        h_path,
        sheet_name="New Rates",
        header=3
    )

    h_partner = read_excel_file(
        h_path,
        sheet_name="Hub Wise Partner Name"
    )
    hub_master = read_excel_file(get_excel_file(HUB_MASTER_DIR))

    c["Activity Type"] = "Delivery"
    d["Activity Type"] = "Pickup"

    df = pd.concat([c, d], ignore_index=True)

    e = pd.DataFrame()

    e["Concnate"] = df["CN #"].astype(str).str.strip() + "_" + df["Activity Type"]
    e["S.NO"] = range(1, len(df) + 1)
    e["Consignment Number"] = s(df, "CN #")
    e["Movement Type"] = s(df, "Movement Type")
    e["Movement Classification"] = s(df, "Movement Classification")
    e["Status(In Trip)"] = df["Activity Type"]
    e["Created At"] = s(df, "Created At")
    e["Origin Hub"] = s(df, "Origin Hub")
    e["(Destination Hub)"] = s(df, "(Destination Hub)")

    e["Billing Hub Code"] = df.apply(
        lambda r: r["Origin Hub"] if r["Activity Type"] == "Pickup" else r["(Destination Hub)"],
        axis=1,
    )

    hub_lookup = hub_master[["code", "name", "region", "state", "Serving DC Name"]].copy()
    hub_lookup["code"] = hub_lookup["code"].astype(str).str.strip()

    e["Billing Hub Code"] = e["Billing Hub Code"].astype(str).str.strip()

    e = e.merge(
        hub_lookup,
        left_on="Billing Hub Code",
        right_on="code",
        how="left",
    )

    e.rename(columns={
        "name": "Billing Hub Name",
        "region": "Hub Region",
        "state": "Hub State",
        "Serving DC Name": "Serving DC Name"
    }, inplace=True)

    e.drop(columns=["code"], inplace=True, errors="ignore")

    e["Origin Pincode"] = s(df, "Sender Pincode")
    e["Destination Pincode"] = s(df, "Destination Pincode")

    e["Billing Pincode"] = df.apply(
        lambda r: r["Sender Pincode"] if r["Activity Type"] == "Pickup" else r["Destination Pincode"],
        axis=1,
    )

    e["Completed Time"] = s(df, "Completed Time")
    e["First Inscan at Hub Time"] = s(df, "First Inscan at Hub Time")
    e["Delivered Time"] = s(df, "Delivered Time")

    e["Considered Date"] = df.apply(
        lambda r: r["Delivered Time"]
        if r["Activity Type"] == "Delivery"
        else (
            r["First Inscan at Hub Time"]
            if pd.notna(r.get("First Inscan at Hub Time"))
            else r.get("Pickup Time", "")
        ),
        axis=1,
    )

    e["Pickup Attempts"] = s(df, "Pickup Attempts")
    e["Delivery Attempts"] = s(df, "Delivery Attempts")
    e["Input Weight"] = s(df, "Input Weight")
    e["Number Of Pieces"] = s(df, "Number Of Pieces")
    e["Chargeable Weight"] = s(df, "Verified Chargeable Weight")
    e["Revised Ch. Weight"] = ""

    e["Pincode Slab Key"] = make_key(e["Billing Hub Code"], e["Billing Pincode"])

    # Prepare F lookup: Sys hub + Pincode
    f_lookup = f.copy()
    f_lookup["Pincode Slab Key"] = make_key(f_lookup["Sys hub"], f_lookup["Pincode"])

    f_cols = [
        "Pincode Slab Key",
        "HUB Classification",
        "Cluster",
        "Partner Name",
        "Final Distance slab",
    ]

    e = e.merge(
        f_lookup[f_cols],
        on="Pincode Slab Key",
        how="left",
    )

    e.rename(columns={"Final Distance slab": "Dist Slab"}, inplace=True)

        # Prepare G lookup: Hub Code -> Billing Zone
    g_lookup = g[["Hub Code", "Hub Billing Zone"]].copy()

    g_lookup["Hub Code"] = g_lookup["Hub Code"].fillna("").astype(str).str.strip()
    g_lookup["Hub Billing Zone"] = g_lookup["Hub Billing Zone"].fillna("").astype(str).str.strip()

    e["Billing Hub Code"] = e["Billing Hub Code"].fillna("").astype(str).str.strip()

    e = e.merge(
        g_lookup,
        left_on="Billing Hub Code",
        right_on="Hub Code",
        how="left",
    )

    e.rename(columns={"Hub Billing Zone": "Billing Zone"}, inplace=True)
    e.drop(columns=["Hub Code"], inplace=True, errors="ignore")

    # Ensure Partner Name exists after F merge
    if "Partner Name" not in e.columns:
        e["Partner Name"] = ""

    # -----------------------------
    # H Partner Entity Lookup
    # Logic:
    # E Partner Name = H Hub Wise Partner Name sheet Partner Name
    # Pull H HIH Partner Entity
    # -----------------------------

    h_partner = h_partner.dropna(axis=1, how="all").copy()
    h_partner.columns = (
        h_partner.columns.astype(str)
        .str.replace("\xa0", " ", regex=False)
        .str.strip()
    )

    # Expected columns in H sheet: Partner Name, HIH Partner Entity
    # Fallback: if exact names differ, use first two non-empty columns
    if "Partner Name" in h_partner.columns:
        h_partner_name_col = "Partner Name"
    elif "Partner" in h_partner.columns:
        h_partner_name_col = "Partner"
    else:
        h_partner_name_col = h_partner.columns[0]

    if "HIH Partner Entity" in h_partner.columns:
        h_partner_entity_col = "HIH Partner Entity"
    else:
        h_partner_entity_col = h_partner.columns[1]

    h_partner_lookup = h_partner[
        [h_partner_name_col, h_partner_entity_col]
    ].copy()

    h_partner_lookup.rename(
        columns={
            h_partner_name_col: "_h_partner_name",
            h_partner_entity_col: "HIH Partner Entity",
        },
        inplace=True,
    )

    h_partner_lookup["_partner_key"] = h_partner_lookup["_h_partner_name"].apply(
        normalize_partner_name
    )

    e["_partner_key"] = e["Partner Name"].apply(normalize_partner_name)

    h_partner_lookup = h_partner_lookup[
        h_partner_lookup["_partner_key"] != ""
    ]

    h_partner_lookup = h_partner_lookup.drop_duplicates(
        subset=["_partner_key"]
    )

    e = e.merge(
        h_partner_lookup[["_partner_key", "HIH Partner Entity"]],
        on="_partner_key",
        how="left",
    )

    e.drop(columns=["_partner_key"], inplace=True, errors="ignore")

    e["Hub Billing type"] = e["HUB Classification"].apply(
        lambda x: "HIH" if str(x).strip().upper() == "HIH" else "E2E"
    )
    

    weights = pd.DataFrame({
        "Input Weight": pd.to_numeric(e["Input Weight"], errors="coerce"),
        "Chargeable Weight": pd.to_numeric(e["Chargeable Weight"], errors="coerce"),
    })

    e["Chargeable Weight (Payable)"] = weights.max(axis=1).fillna(0).clip(lower=4)

    e["Floor Number"] = ""

    def clean_key_part(x):
        if pd.isna(x):
            return ""
        return " ".join(str(x).replace("\xa0", " ").strip().split()).upper()
    
    def clean_slab(x):
        value = clean_key_part(x)
        value = value.replace(" TO ", "-")
        value = value.replace(" - ", "-")
        value = value.replace(" -", "-")
        value = value.replace("- ", "-")
        return value

    # -----------------------------
    # H Rate Card Lookup
    # Logic:
    # E Partner Name  = H Partner
    # E Dist Slab     = H Last Mile Distance Slab Rev
    # E Billing Zone  = H Hub Definition REV
    # -----------------------------

    h.columns = h.columns.astype(str).str.strip()

    e["_rate_partner"] = e["Partner Name"].apply(normalize_partner_name)
    e["_rate_slab"] = e["Dist Slab"].apply(clean_slab)
    e["_rate_zone"] = e["Billing Zone"].apply(clean_key_part)

    # Keep visible/debug key in E_Base
    e["concnate"] = (
        e["_rate_partner"]
        + "_"
        + e["_rate_slab"]
        + "_"
        + e["_rate_zone"]
    )

    h_lookup = h[
        [
            "Partner",
            "Hub Definition REV",
            "Last Mile Distance Slab Rev",
            "Rate upto 4kgs_RFQ",
            ">4 kgs - Per kg rate_RFQ",
        ]
    ].copy()

    h_lookup["_rate_partner"] = h_lookup["Partner"].apply(normalize_partner_name)
    h_lookup["_rate_slab"] = h_lookup["Last Mile Distance Slab Rev"].apply(clean_slab)
    h_lookup["_rate_zone"] = h_lookup["Hub Definition REV"].apply(clean_key_part)

    h_lookup = h_lookup[
        (h_lookup["_rate_partner"] != "")
        & (h_lookup["_rate_slab"] != "")
        & (h_lookup["_rate_zone"] != "")
    ]

    h_lookup = h_lookup.drop_duplicates(
        subset=["_rate_partner", "_rate_slab", "_rate_zone"]
    )

    e = e.merge(
        h_lookup[
            [
                "_rate_partner",
                "_rate_slab",
                "_rate_zone",
                "Rate upto 4kgs_RFQ",
                ">4 kgs - Per kg rate_RFQ",
            ]
        ],
        on=["_rate_partner", "_rate_slab", "_rate_zone"],
        how="left",
    )

    is_hih = (
        e["Hub Billing type"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
        .eq("HIH")
    )

    print("========== H RATE LOOKUP DEBUG ==========")
    print("Total rows:", len(e))
    print("HIH rows:", int(is_hih.sum()))

    print("HIH rows with Partner Name:", int(
        e.loc[is_hih, "Partner Name"].fillna("").astype(str).str.strip().ne("").sum()
    ))

    print("HIH rows with Dist Slab:", int(
        e.loc[is_hih, "Dist Slab"].fillna("").astype(str).str.strip().ne("").sum()
    ))

    print("HIH rows with Billing Zone:", int(
        e.loc[is_hih, "Billing Zone"].fillna("").astype(str).str.strip().ne("").sum()
    ))

    print("HIH rows matched with H rates:", int(
        e.loc[is_hih, "Rate upto 4kgs_RFQ"].notna().sum()
    ))

    print("\nSample E lookup keys:")
    print(
        e.loc[
            is_hih,
            [
                "Partner Name",
                "Dist Slab",
                "Billing Zone",
                "_rate_partner",
                "_rate_slab",
                "_rate_zone",
            ],
        ]
        .drop_duplicates()
        .head(20)
        .to_string(index=False)
    )

    print("\nSample H lookup keys:")
    print(
        h_lookup[
            [
                "Partner",
                "Last Mile Distance Slab Rev",
                "Hub Definition REV",
                "_rate_partner",
                "_rate_slab",
                "_rate_zone",
            ]
        ]
        .drop_duplicates()
        .head(20)
        .to_string(index=False)
    )

    print("=========================================")

    e["Min. Freight (Upto 4 KGs)"] = pd.NA
    e["Freight / KG (Above 4 KGs)"] = pd.NA

    e.loc[is_hih, "Min. Freight (Upto 4 KGs)"] = e.loc[
        is_hih, "Rate upto 4kgs_RFQ"
    ]

    e.loc[is_hih, "Freight / KG (Above 4 KGs)"] = e.loc[
        is_hih, ">4 kgs - Per kg rate_RFQ"
    ]

    # Debug file to check unmatched HIH rows
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    debug_unmatched = e.loc[
        is_hih & e["Rate upto 4kgs_RFQ"].isna(),
        [
            "Billing Hub Code",
            "Partner Name",
            "Dist Slab",
            "Billing Zone",
            "_rate_partner",
            "_rate_slab",
            "_rate_zone",
            "HUB Classification",
            "Hub Billing type",
            "concnate",
        ],
    ].head(1000)

    debug_h_keys = h_lookup[
        [
            "Partner",
            "Last Mile Distance Slab Rev",
            "Hub Definition REV",
            "_rate_partner",
            "_rate_slab",
            "_rate_zone",
            "Rate upto 4kgs_RFQ",
            ">4 kgs - Per kg rate_RFQ",
        ]
    ].head(1000)

    with pd.ExcelWriter(OUTPUT_DIR / "debug_h_lookup.xlsx") as writer:
        debug_unmatched.to_excel(writer, sheet_name="Unmatched HIH E Keys", index=False)
        debug_h_keys.to_excel(writer, sheet_name="Sample H Keys", index=False)

    e.drop(
        columns=[
            "Rate upto 4kgs_RFQ",
            ">4 kgs - Per kg rate_RFQ",
            "_rate_partner",
            "_rate_slab",
            "_rate_zone",
        ],
        inplace=True,
        errors="ignore",
    )

    e["Additional Freight Rs. 1 / KG on Pickup (RVP only)"] = e.apply(
        lambda r: (
            max(r["Chargeable Weight (Payable)"] - 4, 0)
            if r["Status(In Trip)"] == "Pickup" and r["Movement Type"] == "RVP" and r["Hub Billing type"] == "HIH"
            else (
                r["Chargeable Weight (Payable)"]
                if r["Status(In Trip)"] == "Pickup" and r["Movement Type"] == "RVP" and r["Hub Billing type"] == "E2E"
                else ""
            )
        ),
        axis=1,
    )

    floor_number = pd.to_numeric(e["Floor Number"], errors="coerce")
    chargeable_weight = pd.to_numeric(e["Chargeable Weight"], errors="coerce")

    e["Floor Handling Amount"] = [
        1 if status == "Delivery" and floor >= 2 and weight >= 100
        else 0.1 if status == "Delivery" and floor >= 2 and weight >= 30
        else ""
        for status, floor, weight in zip(e["Status(In Trip)"], floor_number, chargeable_weight)
    ]
    e["Freight As per Slab"] = ""
    e["Max. Freight Charges"] = ""
    e["Promo Incentive"] = ""

    max_freight = pd.to_numeric(e["Max. Freight Charges"], errors="coerce").fillna(0)
    promo = pd.to_numeric(e["Promo Incentive"], errors="coerce").fillna(0)
    e["Total CN Cost"] = max_freight + promo

    e["Dup Check"] = ""
    e["Customer Reference No."] = s(df, "Customer Reference Number")

    prefixes = ("FY", "16", "17", "RD", "BB", "SS")

    customer_ref = e["Customer Reference No."].fillna("").astype(str).str.strip()

    e["Business Type"] = customer_ref.apply(
        lambda x: "Online" if x.startswith(prefixes) else "Offline"
    )

    e["HD/GG"] = ""
    e["Serv. DC code"] = ""
    e["Duplicate"] = ""
    e["Other Rate"] = ""
    e["Remark"] = ""

    final_columns = [
        "Concnate",
        "S.NO",
        "Consignment Number",
        "Movement Type",
        "Status(In Trip)",
        "Created At",
        "Origin Hub",
        "(Destination Hub)",
        "Billing Hub Code",
        "Billing Hub Name",
        "HUB Classification",
        "Hub Billing type",
        "Origin Pincode",
        "Destination Pincode",
        "Billing Pincode",
        "Cluster",
        "Hub Region",
        "Hub State",
        "Partner Name",
        "HIH Partner Entity",
        "Completed Time",
        "First Inscan at Hub Time",
        "Delivered Time",
        "Considered Date",
        "Pickup Attempts",
        "Delivery Attempts",
        "Input Weight",
        "Number Of Pieces",
        "Chargeable Weight",
        "Revised Ch. Weight",
        "Pincode Slab Key",
        "Dist Slab",
        "Billing Zone",
        "Chargeable Weight (Payable)",
        "Floor Number",
        "concnate",
        "Min. Freight (Upto 4 KGs)",
        "Freight / KG (Above 4 KGs)",
        "Additional Freight Rs. 1 / KG on Pickup (RVP only)",
        "Floor Handling Amount",
        "Freight As per Slab",
        "Max. Freight Charges",
        "Promo Incentive",
        "Total CN Cost",
        "Dup Check",
        "Customer Reference No.",
        "Business Type",
        "Movement Classification",
        "HD/GG",
        "Serv. DC code",
        "Serving DC Name",
        "Duplicate",
        "Other Rate",
        "Remark",
    ]

    e = e[final_columns]

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    output_file = OUTPUT_DIR / "E_Base.xlsx"
    e.to_excel(output_file, index=False)

    print("T2 complete: E Base generated")
    print(f"Rows: {len(e)}")
    print(f"Output: {output_file}")


if __name__ == "__main__":
    generate_e_base()