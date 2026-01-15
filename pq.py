import pandas as pd

sheets = pd.read_excel("МЭ 25-26.xlsx", sheet_name=None)
res = pd.concat(sheets.values(), ignore_index=True)
res.to_excel("res МЭ 25-26.xlsx", index=False)
