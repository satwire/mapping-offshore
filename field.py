import datetime

import pandas as pd

# Name of source and destination sheet files
SOURCE1_DB = "annual_production.xlsx"
SOURCE2_DB = "monthly_production.xlsx"
DESTINATION_DB = "field.xlsx"

# Load source and destination sheets
src1_df = pd.read_excel(SOURCE1_DB)
src2_df = pd.read_excel(SOURCE2_DB)
des_df = pd.read_excel(DESTINATION_DB)

# Iterate through all source rows
for src_index, src_row in src1_df.iterrows():
    # Get current source row field name
    field_name = src_row["Field Name"]

    # Check if source value is in destination
    if field_name in des_df["Field"].values:
        # Find destination index that has the same field value
        des_index = (des_df.index[des_df["Field"] == field_name])[0]

        # Update destination values from source
        des_df.at[des_index, "Latest Annual Oil Prod Bbl"] = src_row[
            "Latest Annual Oil Prod Bbl"
        ]
        des_df.at[des_index, "Latest Annual Gas Prod Tscf"] = src_row[
            "Latest Annual Gas Prod Tscf"
        ]
        des_df.at[des_index, "Latest Annual Cond Prod Bbl"] = src_row[
            "Latest Annual Cond Prod Bbl"
        ]
        des_df.at[des_index, "Latest Annual Prod Year"] = src_row[
            "Latest Annual Prod Year"
        ]
        des_df.at[des_index, "Field Id"] = src_row["Field Id"]

# Iterate through all source rows
for src_index, src_row in src2_df.iterrows():
    # Get current source row field name and entity type
    field_name = src_row["Field Name"]
    entity_type = src_row["Entity Type"]

    # Check if source value is in destination and is correct entity type
    if field_name in des_df["Field"].values and entity_type == "Field":
        # Find destination index that has the same field value
        des_index = (des_df.index[des_df["Field"] == field_name])[0]

        # Update destination values from source
        des_df.at[des_index, "Oil Prod MMbbL"] = src_row["Oil Prod MMbbL"]
        des_df.at[des_index, "Gas Prod MMscf"] = src_row["Gas Prod MMscf"]
        des_df.at[des_index, "Cond Prod MMbbl"] = src_row["Cond Prod MMbbl"]
        des_df.at[des_index, "Entity Id"] = src_row["Entity Id"]


# Save updated sheet
date = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
des_df.to_excel(f"field_{date}.xlsx")
