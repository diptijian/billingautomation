from pathlib import Path
import pandas as pd


C_DIR = Path("input/t2/c_cleaned")
D_DIR = Path("input/t2/d_cleaned")
F_DIR = Path("input/t2/f")
G_DIR = Path("input/t2/g")
OUTPUT_DIR = Path("output/t2")


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


def make_key(a, b):
    return a.astype(str).str.strip() + "_" + b.astype(str).str.strip()


def generate_e_base():
    c = read_excel_file(get_excel_file(C_DIR))
    d = read_excel_file(get_excel_file(D_DIR))
    f = read_excel_file(get_excel_file(F_DIR))
    g = read_excel_file(get_excel_file(G_DIR))

    c["Activity Type"] = "Delivery"
    d["Activity Type"] = "Pickup"

    df = pd.concat([c, d], ignore_index=True)

    e = pd.DataFrame()

    e["Concnate"] = df["CN #"].astype(str).str.strip() + "_" + df["Activity Type"]
    e["S.NO"] = range(1, len(df) + 1)
    e["Consignment Number"] = s(df, "CN #")
    e["Movement Type"] = s(df, "Movement Type")
    e["Status(In Trip)"] = df["Activity Type"]
    e["Created At"] = s(df, "Created At")
    e["Origin Hub"] = s(df, "Origin Hub")
    e["(Destination Hub)"] = s(df, "(Destination Hub)")

    e["Billing Hub Code"] = df.apply(
        lambda r: r["Origin Hub"] if r["Activity Type"] == "Pickup" else r["(Destination Hub)"],
        axis=1,
    )

    e["Billing Hub Name"] = e["Billing Hub Code"]

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

    # Prepare G lookup: Hub Code -> Hub Billing Zone
    g_lookup = g[["Hub Code", "Hub Billing Zone"]].copy()
    g_lookup["Hub Code"] = g_lookup["Hub Code"].astype(str).str.strip()

    e["Billing Hub Code"] = e["Billing Hub Code"].astype(str).str.strip()

    # Merge G lookup for Billing Zone
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

    e["Hub Billing type"] = ""
    e["Hub Region"] = ""
    e["Hub State"] = ""

    weights = pd.DataFrame({
        "Input Weight": pd.to_numeric(e["Input Weight"], errors="coerce"),
        "Chargeable Weight": pd.to_numeric(e["Chargeable Weight"], errors="coerce"),
    })

    e["Chargeable Weight (Payable)"] = weights.max(axis=1).fillna(0).clip(lower=4)

    e["Floor Number"] = ""
    e["concnate"] = (
        e["Billing Zone"].fillna("").astype(str)
        + "_"
        + e["Dist Slab"].fillna("").astype(str)
        + "_"
        + e["Partner Name"].fillna("").astype(str)
    )

    e["Min. Freight (Upto 4 KGs)"] = ""
    e["Freight / KG (Above 4 KGs)"] = ""
    e["Additional Freight Rs. 1 / KG on Pickup (RVP only)"] = ""
    e["Floor Handling Amount"] = ""
    e["Freight As per Slab"] = ""
    e["Max. Freight Charges"] = ""
    e["Promo Incentive"] = ""
    e["Total CN Cost"] = ""
    e["Dup Check"] = ""
    e["Customer Reference No."] = s(df, "Customer Reference Number")

    prefixes = ("FY", "16", "17", "RD", "BB", "SS")
    
    customer_ref = e["Customer Reference No."].fillna("").astype(str).str.strip()

    e["Business Type"] = customer_ref.apply(
        lambda x: "Online" if x.startswith(prefixes) else "Offline"
    )

    e["Movement Classification"] = ""
    e["HD/GG"] = ""
    e["Serv. DC code"] = ""
    e["Serv. DC Name"] = ""
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
        "Serv. DC Name",
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