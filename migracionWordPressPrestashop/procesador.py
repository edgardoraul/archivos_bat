""" Creado por muaaaa """

import pandas as pd

df = pd.read_excel("Modelo 2.xlsx",
	sheet_name = "Hoja0",
	header = 0 )

print(df.shape)

print(df.columns)