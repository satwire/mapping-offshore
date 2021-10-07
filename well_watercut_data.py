import datetime

import pandas as pd

# Name of source and destination sheet files
SOURCE_DB = "satwiko 2.xlsx"
DESTINATION_DB = "field_raw.xlsx"

# Load source and destination sheets
src_df = pd.read_excel(SOURCE_DB)
des_df = pd.read_excel(DESTINATION_DB)

# Iterate through all source rows
for src_index, src_row in src_df.iterrows():
    # Get current source row field name
    field_name = src_row["Field Name"]

    # Check if source value is in destination
    if field_name in des_df["Field"].values:
        # Find destination index that has the same field value
        des_index = (des_df.index[des_df["Field"] == field_name])[0]

        # Update destination values from source
        des_df.at[des_index, "Latest Annual Prod Year"] = src_row["Latest Annual Prod Year"]
        des_df.at[des_index, "Numb Wells"] = src_row["Numb Wells"]
        des_df.at[des_index, "Numb Producers"] = src_row["Numb Producers"]
        des_df.at[des_index, "Numb Oil Producers"] = src_row["Numb Oil Producers"]
        des_df.at[des_index, "Numb Gas Producers"] = src_row["Numb Gas Producers"]
        des_df.at[des_index, "Numb Act Producers"] = src_row["Numb Act Producers"]
        des_df.at[des_index, "Numb Act Oil Producers"] = src_row["Numb Act Oil Producers"]
        des_df.at[des_index, "Numb Act Gas Producers"] = src_row["Numb Act Gas Producers"]
        des_df.at[des_index, "Numb Act Wells Date"] = src_row["Numb Act Wells Date"]

# Save updated sheet
date = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
des_df.to_excel(f"Mapping_Offshore_{date}.xlsx")
