from pathlib import Path
import pandas as pd


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


def read_excel_file(path: Path) -> pd.DataFrame:
    if path.suffix.lower() == ".xlsb":
        return pd.read_excel(path, engine="pyxlsb")
    return pd.read_excel(path)


def s(df, col, default=""):
    return df[col] if col in df.columns else default


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
    h = pd.read_excel(
        get_excel_file(H_DIR),
        sheet_name="New Rates",
        engine="pyxlsb",
        header=3
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

    # Prepare G lookup: Hub Code -> Billing Zone + HIH Partner Entity
    g_lookup = g[["Hub Code", "Hub Billing Zone", "HIH Partner Entity"]].copy()

    g_lookup["Hub Code"] = g_lookup["Hub Code"].fillna("").astype(str).str.strip()
    g_lookup["Hub Billing Zone"] = g_lookup["Hub Billing Zone"].fillna("").astype(str).str.strip()
    g_lookup["HIH Partner Entity"] = g_lookup["HIH Partner Entity"].fillna("").astype(str).str.strip()

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

    # Prefer HIH Partner Entity from G for rate lookup.
    # F Partner Name may not match the H rate-card partner naming.
    rate_partner = e["HIH Partner Entity"].where(
        e["HIH Partner Entity"].fillna("").astype(str).str.strip() != "",
        e["Partner Name"]
    )

    # Build E-side rate lookup components
    e["_rate_partner"] = rate_partner.apply(clean_key_part)
    e["_rate_zone"] = e["Billing Zone"].apply(clean_key_part)
    e["_rate_slab"] = e["Dist Slab"].apply(clean_key_part)

    # Keep visible output/debug key
    e["concnate"] = (
        e["_rate_partner"]
        + "_"
        + e["_rate_zone"]
        + "_"
        + e["_rate_slab"]
    )

    # Clean H headers
    h.columns = h.columns.astype(str).str.strip()

    # Build H-side rate lookup from actual formula columns, NOT H["Key"]
    h_lookup = h[
        [
            "Partner",
            "Hub Definition REV",
            "Last Mile Distance Slab Rev",
            "Rate upto 4kgs_RFQ",
            ">4 kgs - Per kg rate_RFQ",
        ]
    ].copy()

    h_lookup["_rate_partner"] = h_lookup["Partner"].apply(clean_key_part)
    h_lookup["_rate_zone"] = h_lookup["Hub Definition REV"].apply(clean_key_part)
    h_lookup["_rate_slab"] = h_lookup["Last Mile Distance Slab Rev"].apply(clean_key_part)

    h_lookup = h_lookup[
        (h_lookup["_rate_partner"] != "")
        & (h_lookup["_rate_zone"] != "")
        & (h_lookup["_rate_slab"] != "")
    ]

    h_lookup = h_lookup.drop_duplicates(
        subset=["_rate_partner", "_rate_zone", "_rate_slab"]
    )

    e = e.merge(
        h_lookup[
            [
                "_rate_partner",
                "_rate_zone",
                "_rate_slab",
                "Rate upto 4kgs_RFQ",
                ">4 kgs - Per kg rate_RFQ",
            ]
        ],
        on=["_rate_partner", "_rate_zone", "_rate_slab"],
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

    e["Min. Freight (Upto 4 KGs)"] = pd.NA
    e["Freight / KG (Above 4 KGs)"] = pd.NA

    e.loc[is_hih, "Min. Freight (Upto 4 KGs)"] = e.loc[
        is_hih, "Rate upto 4kgs_RFQ"
    ]

    e.loc[is_hih, "Freight / KG (Above 4 KGs)"] = e.loc[
        is_hih, ">4 kgs - Per kg rate_RFQ"
    ]

    # Debug file to check why rates are blank
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    debug_unmatched = e.loc[
        is_hih & e["Rate upto 4kgs_RFQ"].isna(),
        [
            "Billing Hub Code",
            "Billing Zone",
            "Dist Slab",
            "Partner Name",
            "HIH Partner Entity",
            "_rate_partner",
            "_rate_zone",
            "_rate_slab",
            "HUB Classification",
            "Hub Billing type",
            "concnate",
        ],
    ].head(1000)

    debug_h_keys = h_lookup[
        [
            "Partner",
            "Hub Definition REV",
            "Last Mile Distance Slab Rev",
            "_rate_partner",
            "_rate_zone",
            "_rate_slab",
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
            "_rate_zone",
            "_rate_slab",
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