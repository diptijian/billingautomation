from pathlib import Path
import pandas as pd


SHIPMENT_DIR = Path("input/t1/shipment")
TRIP_DIR = Path("input/t1/trip")
OUTPUT_DIR = Path("output/t1")


def get_excel_file(folder: Path) -> Path:
    files = (
        list(folder.glob("*.xlsx"))
        + list(folder.glob("*.xlsb"))
        + list(folder.glob("*.xls"))
    )

    if len(files) == 0:
        raise FileNotFoundError(f"No Excel file found in {folder}")

    if len(files) > 1:
        raise ValueError(
            f"Multiple Excel files found in {folder}. Keep only one."
        )

    return files[0]


def read_excel_file(path: Path) -> pd.DataFrame:
    if path.suffix.lower() == ".xlsb":
        return pd.read_excel(path, engine="pyxlsb")

    return pd.read_excel(path)


def generate_cd():

    shipment_path = get_excel_file(SHIPMENT_DIR)
    trip_path = get_excel_file(TRIP_DIR)

    shipment_df = read_excel_file(shipment_path)
    trip_df = read_excel_file(trip_path)

    # Keep only CARRIER(NP)
    valid_trips = trip_df[
        trip_df["Trip Fleet Type"].fillna("").str.strip()
        == "CARRIER(NP)"
    ]

    # Keep only matching trip numbers
    shipment_df = shipment_df.merge(
        valid_trips[["Trip Number"]],
        on="Trip Number",
        how="inner",
    )

    # Remove Non-payable vendors
    shipment_df = shipment_df[
        shipment_df["Vendor Name"]
        .fillna("")
        .str.strip()
        .str.lower()
        != "non-payable"
    ]

    # Delivered
    delivered_df = shipment_df[
        shipment_df["Status(In Trip)"]
        .fillna("")
        .str.strip()
        == "Delivered"
    ]

    # Pickup
    pickup_df = shipment_df[
        shipment_df["Status(In Trip)"]
        .fillna("")
        .str.strip()
        == "Pickup Completed"
    ]

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    delivered_df.to_excel(
        OUTPUT_DIR / "C_Delivered.xlsx",
        index=False,
    )

    pickup_df.to_excel(
        OUTPUT_DIR / "D_Pickup.xlsx",
        index=False,
    )

    print("T1 complete: C and D generated")
    print(f"Shipment file: {shipment_path}")
    print(f"Trip file: {trip_path}")
    print(f"Delivered rows: {len(delivered_df)}")
    print(f"Pickup rows: {len(pickup_df)}")