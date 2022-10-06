import pandas as pd
import xlwings as xw


# Create Dataframes
df1 = pd.read_csv("https://raw.githubusercontent.com/mwaskom/seaborn-data/master/tips.csv")
df2 = df1.groupby(by="sex").sum()[["tip"]]
df3 = df1.groupby(by="day").sum()[["tip"]]


# -----------------------------------------------------
# OPTION 1: WRITE TO NEW EXCEL FILE (will overwrite existing workbook)
# -----------------------------------------------------
with pd.ExcelWriter("output.xlsx", engine="openpyxl") as writer:
    df1.to_excel(writer, sheet_name="Sheet_1") # Default position: cell A1.
    df2.to_excel(writer, sheet_name="Sheet_1", startcol=10, startrow=1, header=True, index=False) 
    df3.to_excel(writer, sheet_name="Sheet_1", startcol=15, startrow=2, header=True, index=False)



# -----------------------------------------------------
# OPTION 2: WRITE TO EXISTING EXCEL FILE
# -----------------------------------------------------
# ADJUST THE FOLLOWING:
wb_name = "output.xlsx"
sheet_name = "xlwings_option"
df_mapping = {"A1": df1, "K1": df2, "K5": df3}

# Open Excel in background
with xw.App(visible=False) as app:
    wb = app.books.open(wb_name)
    # Add sheet if it does not exist
    current_sheets = [sheet.name for sheet in wb.sheets]
    if sheet_name not in current_sheets:
        wb.sheets.add(sheet_name)
    # Write dataframe to cell range
    for cell_target, df in df_mapping.items():
        wb.sheets(sheet_name).range(cell_target).options(pd.DataFrame, index=False).value = df
    wb.save()
