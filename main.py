import pandas as pd
import numpy as np
import dataframe_image as dfi
import string

pathToExcelFile = r"YOUR_PATH_TO_EXCEL_FILE_E.G_C:\Users\YOUR_USERNAME\FILE.xlsx"
yourClass = "YOUR_CLASS_E.G_11A1"

cols = list(string.ascii_uppercase)
column = cols[(int(yourClass[0:2]) - 10) * 5 + int(yourClass[-1]) + 1]

# find neccessary informations
dfs = pd.read_excel(pathToExcelFile,
                    sheet_name = 1,
                    na_values = "",
                    usecols = f"A,{column}",
                    skiprows = [0, 1, 2, 3, 4]
)
df = dfs.replace(np.nan, "")

# convert to a easier to read dataframe
data = {}
for i in range(0, 26, 5):
  data[df.iat[i, 0]] = [df.iat[x, 1] for x in range(i, i + 5)]

cleanDf = pd.DataFrame.from_dict(data)

# export dataframe to image
dfi.export(cleanDf, "tkb.png")
