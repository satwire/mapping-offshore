import datetime

import pandas as pd

# Name of source and destination sheet files
SOURCE_DB = "FIELD_FULL_wincil_100000246509.xlsx"
DESTINATION_DB = "Mapping_Offshore_draft.xlsx"

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
        des_df.at[des_index, "Oil (Pp Mmbbl)"] = src_row["Oil Recoverable Pp Mmbbl"]
        des_df.at[des_index, "Gas (Pp Mmscf)"] = src_row["Gas Recoverable Pp Mmscf"]
        des_df.at[des_index, "Cond (Pp Mmbbl)"] = src_row["Cond Recoverable Pp Mmbbl"]

# Save updated sheet
date = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
des_df.to_excel(f"Mapping_Offshore_{date}.xlsx")
