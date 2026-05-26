from processor import process_files

# Input files
shipment_file = "input/shipment.xlsb"
trip_file = "input/trip.xlsb"

# Process
delivered_df, pickup_df = process_files(
    shipment_file,
    trip_file
)

# Export
delivered_df.to_excel(
    "output/C_Delivered.xlsx",
    index=False
)

pickup_df.to_excel(
    "output/D_Pickup.xlsx",
    index=False
)

print("C and D generated successfully")