
import pandas as pd
import glob

# ========= USER SETTINGS =========

TARGET_SKU = "Capstan by Pall Mall 20HL"

START_DATE = "2026-03-01"
END_DATE   = "2026-03-31"

# ================================

files = glob.glob("*.csv")
all_data = []

for file in files:
    df = pd.read_csv(file)

    # SKU filter
    df = df[df["SKU"] == TARGET_SKU]
    if df.empty:
        continue

    # Date from filename
    date = file.replace(".csv", "")
    df["Date"] = pd.to_datetime(date)

    # Total Sale
    df["Total_Sale"] = (
        df["DD_Sale_Retail_Urban"].fillna(0)
        + df["VDD_Sale_Retail_Rural"].fillna(0)
        + df["Urban_WS_Wholesale_Urban"].fillna(0)
        + df["Rural_WS_Wholesale_Rural"].fillna(0)
        + df["Mandi_WS_Wholesale_Mandi"].fillna(0)
    )

    df = df[
        [
            "Date",
            "SKU",
            "Opening_Stock",
            "DD_Sale_Retail_Urban",
            "VDD_Sale_Retail_Rural",
            "Urban_WS_Wholesale_Urban",
            "Rural_WS_Wholesale_Rural",
            "Mandi_WS_Wholesale_Mandi",
            "Total_Sale",
            "Closing_Stock",
        ]
    ]

    all_data.append(df)

# Merge
if all_data:
    final_df = pd.concat(all_data, ignore_index=True)
else:
    final_df = pd.DataFrame()

# ✅ Create full date range
date_range = pd.date_range(start=START_DATE, end=END_DATE)

full_df = pd.DataFrame({"Date": date_range})

# Merge with actual data
final_df = full_df.merge(final_df, on="Date", how="left")

# Sort
final_df = final_df.sort_values("Date")

# ✅ Mark OFF days
final_df["Day_Status"] = final_df["SKU"].apply(
    lambda x: "OFF" if pd.isna(x) else "Working"
)

# Fill missing SKU name
final_df["SKU"] = final_df["SKU"].fillna(TARGET_SKU)

# Fill numeric columns with 0
num_cols = [
    "Opening_Stock",
    "DD_Sale_Retail_Urban",
    "VDD_Sale_Retail_Rural",
    "Urban_WS_Wholesale_Urban",
    "Rural_WS_Wholesale_Rural",
    "Mandi_WS_Wholesale_Mandi",
    "Total_Sale",
    "Closing_Stock",
]

for col in num_cols:
    if col in final_df.columns:
        final_df[col] = final_df[col].fillna(0)

# Save Excel with highlight
output_file = f"{TARGET_SKU}_Stock_Register_{START_DATE}_to_{END_DATE}.xlsx"

with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
    final_df.to_excel(writer, index=False, sheet_name="Stock")

    workbook  = writer.book
    worksheet = writer.sheets["Stock"]

    # Highlight format for OFF days
    highlight_format = workbook.add_format({
        'bg_color': '#FFC7CE'
    })

    # Apply conditional formatting
    worksheet.conditional_format(
        f"A2:K{len(final_df)+1}",
        {
            'type': 'formula',
            'criteria': '=$K2="OFF"',
            'format': highlight_format
        }
    )

print("✅ Excel Created with OFF Days Highlighted:", output_file)






















# import pandas as pd
# import glob

# # ========= USER SETTINGS =========

# TARGET_SKU =("VELO Groovy Grape 17MG")
# #VELO Polar Mint 14MG
# #VELO Mango Flame 14MG
# #VELO Groovy Grape 10MG
# #VELO Groovy Grape 6MG - Nano
# #VELO Polar Mint 17MG
# #VELO Groovy Grape 17MG

# #VELO Berry Frost 6MG - Nano
# #VELO Polar Mint 6MG - Nano
# #VELO Rich Elaichi 6MG - Nano
# #VELO Berry Frost 10MG
# #VELO Polar Mint 10MG
# #VELO Frosty Lemon 10MG
# #VELO Rich Elaichi 10MG
# #VELO Strawberry Ice 10mg
# #VELO Tropical Ice 10mg
# #VELO Wintery Watermelon 10MG
# #VELO Berry Frost 14MG

# # Dunhill Lights 20HL
# # Dunhill Switch 20HL
# # Dunhill Special 20HL
# # Gold Leaf Classic 20HL
# # Gold Leaf Classic 20HL LEP
# # Benson & Hedges 20HL - New
# # CbPM Elite
# # John Player 20HL
# # Capstan Filter 20HL
# # Capstan by Pall Mall 20HL
# # CBPMO LEP 125
# # Pall Mall 20HL
# # Gold Flake by Rothmans 20HL
# # Capstan International 20HL





# START_DATE = "2026-01-11"
# END_DATE   = "2026-01-31"

# # ================================

# files = glob.glob("*.csv")
# all_data = []

# for file in files:
#     df = pd.read_csv(file)

#     # SKU filter
#     df = df[df["SKU"] == TARGET_SKU]
#     if df.empty:
#         continue

#     # Date from filename
#     date = file.replace(".csv", "")
#     df["Date"] = pd.to_datetime(date)

#     # Total Sale
#     df["Total_Sale"] = (
#         df["DD_Sale_Retail_Urban"].fillna(0)
#         + df["VDD_Sale_Retail_Rural"].fillna(0)
#         + df["Urban_WS_Wholesale_Urban"].fillna(0)
#         + df["Rural_WS_Wholesale_Rural"].fillna(0)
#         + df["Mandi_WS_Wholesale_Mandi"].fillna(0)
#     )

#     df = df[
#         [
#             "Date",
#             "SKU",
#             "Opening_Stock",
#             "DD_Sale_Retail_Urban",
#             "VDD_Sale_Retail_Rural",
#             "Urban_WS_Wholesale_Urban",
#             "Rural_WS_Wholesale_Rural",
#             "Mandi_WS_Wholesale_Mandi",
#             "Total_Sale",
#             "Closing_Stock",
#         ]
#     ]

#     all_data.append(df)

# # Merge
# final_df = pd.concat(all_data, ignore_index=True)

# # Date filter
# final_df = final_df[
#     (final_df["Date"] >= START_DATE) &
#     (final_df["Date"] <= END_DATE)
# ]

# # Sort
# final_df = final_df.sort_values("Date")

# # 🔥 IMPORTANT: Fill blanks with 0
# final_df = final_df.fillna(0)

# # Save as Excel (clean)
# output_file = f"{TARGET_SKU}_Stock_Register_{START_DATE}_to_{END_DATE}.xlsx"
# final_df.to_excel(output_file, index=False)

# print("✅ Clean Excel Stock Register Created:", output_file)










# import pandas as pd
# import glob

# # ========= USER SETTINGS =========
# TARGET_SKU = "Dunhill Lights 20HL"
# START_DATE = "2026-03-01"
# END_DATE   = "2026-03-31"
# # ================================

# files = glob.glob("*.csv")
# all_data = []

# for file in files:
#     df = pd.read_csv(file)

#     # SKU filter
#     df = df[df["SKU"] == TARGET_SKU]
#     if df.empty:
#         continue

#     # Date from filename
#     date = file.replace(".csv", "")
#     df["Date"] = pd.to_datetime(date)

#     # Fill NaN first
#     cols_fill = [
#         "Opening_Stock",
#         "Received_From_PTC",
#         "DD_Sale_Retail_Urban",
#         "VDD_Sale_Retail_Rural",
#         "Urban_WS_Wholesale_Urban",
#         "Rural_WS_Wholesale_Rural",
#         "Mandi_WS_Wholesale_Mandi",
#     ]

#     for col in cols_fill:
#         if col in df.columns:
#             df[col] = df[col].fillna(0)

#     # Total Sale
#     df["Total_Sale"] = (
#         df["DD_Sale_Retail_Urban"]
#         + df["VDD_Sale_Retail_Rural"]
#         + df["Urban_WS_Wholesale_Urban"]
#         + df["Rural_WS_Wholesale_Rural"]
#         + df["Mandi_WS_Wholesale_Mandi"]
#     )

#     # Closing Stock = Opening - Total Sale
#     df["Closing_Stock"] = df["Opening_Stock"] - df["Total_Sale"]

#     # Current Stock (same as Closing initially)
#     df["Current_Stock"] = df["Closing_Stock"]

#     df = df[
#         [
#             "Date",
#             "SKU",
#             "Opening_Stock",
#             "Received_From_PTC",
#             "DD_Sale_Retail_Urban",
#             "VDD_Sale_Retail_Rural",
#             "Urban_WS_Wholesale_Urban",
#             "Rural_WS_Wholesale_Rural",
#             "Mandi_WS_Wholesale_Mandi",
#             "Total_Sale",
#             "Closing_Stock",
#             "Current_Stock",
#         ]
#     ]

#     all_data.append(df)

# # Merge all files
# if all_data:
#     final_df = pd.concat(all_data, ignore_index=True)
# else:
#     final_df = pd.DataFrame()

# # Full date range
# date_range = pd.date_range(start=START_DATE, end=END_DATE)
# full_df = pd.DataFrame({"Date": date_range})

# # Merge
# final_df = full_df.merge(final_df, on="Date", how="left")

# # Sort
# final_df = final_df.sort_values("Date")

# # Day status
# final_df["Day_Status"] = final_df["SKU"].apply(
#     lambda x: "OFF" if pd.isna(x) else "Working"
# )

# # Fill SKU
# final_df["SKU"] = final_df["SKU"].fillna(TARGET_SKU)

# # Fill numeric columns
# num_cols = [
#     "Opening_Stock",
#     "Received_From_PTC",
#     "DD_Sale_Retail_Urban",
#     "VDD_Sale_Retail_Rural",
#     "Urban_WS_Wholesale_Urban",
#     "Rural_WS_Wholesale_Rural",
#     "Mandi_WS_Wholesale_Mandi",
#     "Total_Sale",
#     "Closing_Stock",
#     "Current_Stock",
# ]

# for col in num_cols:
#     if col in final_df.columns:
#         final_df[col] = final_df[col].fillna(0)

# # ✅ Running Total (Tsale)
# final_df["Tsale"] = final_df["Total_Sale"].cumsum()

# # Save Excel
# output_file = f"{TARGET_SKU}_Stock_Register_{START_DATE}_to_{END_DATE}.xlsx"

# with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
#     final_df.to_excel(writer, index=False, sheet_name="Stock")

#     workbook  = writer.book
#     worksheet = writer.sheets["Stock"]

#     # Highlight OFF days
#     highlight_format = workbook.add_format({'bg_color': '#FFC7CE'})

#     worksheet.conditional_format(
#         f"A2:M{len(final_df)+1}",
#         {
#             'type': 'formula',
#             'criteria': '=$M2="OFF"',
#             'format': highlight_format
#         }
#     )

# print("✅ Excel Created with New Columns & Running Sales:", output_file)