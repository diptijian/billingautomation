import pandas as pd


def process_files(shipment_path, trip_path):

    # Read .xlsb files
    shipment_df = pd.read_excel(shipment_path, engine="pyxlsb")
    trip_df = pd.read_excel(trip_path, engine="pyxlsb")

    # Keep only CARRIER(NP)
    valid_trips = trip_df[
        trip_df["Trip Fleet Type"] == "CARRIER(NP)"
    ]

    # Keep only matching trip numbers
    shipment_df = shipment_df.merge(
        valid_trips[["Trip Number"]],
        on="Trip Number",
        how="inner"
    )

    # Remove Non-payable vendors
    shipment_df = shipment_df[
    shipment_df["Vendor Name"].fillna("").str.strip().str.lower() != "non-payable"
    ]

    # Create Delivered sheet
    delivered_df = shipment_df[
        shipment_df["Status(In Trip)"] == "Delivered"
    ]

    # Create Pickup sheet
    pickup_df = shipment_df[
        shipment_df["Status(In Trip)"] == "Pickup Completed"
    ]

    return delivered_df, pickup_df