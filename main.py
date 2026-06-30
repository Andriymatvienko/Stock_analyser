import pandas as pd
import logging
import os
logging.basicConfig(level=logging.INFO ,filename="log.txt", filemode='w',format="%(asctime)s %(levelname)s %(message)s")




# Table normalizing and creating clear table
def normalizeTable():
    try:
        logging.info("Normalizing started")
        df = pd.read_csv("data.csv", encoding="windows-1251", sep=";", skiprows=2)

        df["Warehouse_Stock"] = df["Warehouse_Stock"].str.replace(' ', '')
        df["Warehouse_Stock"] = pd.to_numeric(df["Warehouse_Stock"], errors='coerce')
        df["Weeks_Cover"] = df["Weeks_Cover"].replace(' ', '')
        df["Weeks_Cover"] = pd.to_numeric(df["Weeks_Cover"], errors='coerce')
        df["NS_vs_LY"] = df["NS_vs_LY"].astype(str).str.replace('%', '').astype(float)
        df_clean = df.dropna()
        df_clean.to_excel("clean_data.xlsx", index=True)
        print("Normalizing completed")
        logging.info("Normalizing completed")
    except FileNotFoundError:
        print("File not found")
        logging.info("Something went wrong, File not found")
    # print (df_clean.columns)

def MakeIMGS(Table):
    for row in range(len(Table["Article_ID"])):
        Table.at[row, "Photo"] =f'=_xlfn.IMAGE("https://prdsdcapp44813.emea.adsint.biz/MicroStrategy/_custom/marvin/png/" &B{row + 2}& "-small.png")'
    return Table

# Table analysing
def AnalyseByStocks():
    logging.info("Analysing by Stocks started")
    # Reading Cleared Data
    df = pd.read_excel("clean_data.xlsx")
    # Getting Limits from User
    print("Enter MinimalStockLimit")
    MinStockLimit = float(input())

    # Filtering
    Order = df[(df["Store_Stock"] <= MinStockLimit) & (df["Warehouse_Stock"] > 0) & (df["NS_vs_LY"] > 0)].reset_index(drop="True")

    # Addining IMGS
    Order = MakeIMGS(Order)
    OrderDf = Order[["Photo","Article_ID", "Net_Sales_EUR" ,"NS_vs_LY", "Store_Stock", "Warehouse_Stock", "Weeks_Cover"]]


    with pd.ExcelWriter("OrderByStocks.xlsx", engine="xlsxwriter") as writer:
        OrderDf.to_excel(writer, index=False)


    # Logging
    logging.info(f"Analysing by Stocks completed, Items wanted to order {len(Order)}")



def AnalyseByWeeks():

    logging.info("Analysing by Weeks started")
    # Reading Cleaned Data
    df = pd.read_excel("clean_data.xlsx")

    # Filtering
    Order = df[(df["Weeks_Cover"] <= 4) & (df["Warehouse_Stock"] > 0)].reset_index(drop="True")
    # IMG Adding


    Order = MakeIMGS(Order)
    OrderDf = Order[["Photo", "Article_ID", "Net_Sales_EUR", "NS_vs_LY", "Store_Stock", "Warehouse_Stock", "Weeks_Cover"]]

    with pd.ExcelWriter("OrderByWeeksCorerage.xlsx", engine="xlsxwriter") as writer:
       OrderDf.to_excel(writer,index=False)

    # Logging
    logging.info(f"Analysing by Weeks completed, Items wanted to order {len(Order)}")
def Starting():
    while True:
        print ("choose function: "
               "1 - Normalize Table"
               "2 - Filter by Weeks Coverage"
               "3 - Filter by Stocks with your Limits")
        status = int(input())
        match status:
            case 1:
                print("Normalizing ur table: ///")
                normalizeTable()
            case 2:
                print("Filter by Weeks Coverage")
                AnalyseByWeeks()
                print("Filtered by Weeks Table done.")
            case 3:
                print("Filter by Stocks with ur Limit")
                AnalyseByStocks()
                print("Filtered by Stocks Table done.")
            case 4:
                break

Starting()

# & (df["Weeks_Cover"] <= 2)

# (df["Store_Stock"] < ArtLimit