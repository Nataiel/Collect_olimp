import pandas as pd

sheets = pd.read_excel("МЭ 25-26.xlsx", sheet_name=None)
res = pd.concat(sheets.values(), ignore_index=True)
res = res[res["Статус"].isin(['Победитель', 'Призёр'])]
res['Результат'] = res['Предмет'] + ' ('+res['Статус'] + ')'
res['Класс участника'] = res['Класс участника'].str.upper().str.replace(":", "", regex=False)
res[["Код", "ОО"]] = res["Школа"].str.split(" - ", n=1, expand=True)
res = res[["Код", "ОО", 'Класс участника', 'Участник', 'Результат']]

res.to_excel("res МЭ 25-26.xlsx", index=False)
