import pandas as pd
pd.set_option("display.max_columns", None)
pd.set_option("display.width", 200)

df = pd.read_excel("input/inquiries_newyork.xlsx")

print(df)
print(df.shape)
print(df.columns)