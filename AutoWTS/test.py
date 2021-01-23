# import xlrd
import pandas as pd
from openpyxl import load_workbook

# Import the web driver from current path
# FrireFox
# browser = webdriver.Firefox(
#     executable_path=BASE_DIR + '\\web_drivers\\geckodriver'
# )
#
# browser.get('https://web.whatsapp.com/')

# Chrome
# browser = webdriver.Chrome(
#     executable_path=BASE_DIR + '\\web_drivers\\chromedriver', options=options
# )
#
# browser.get('https://web.whatsapp.com/')

xl = pd.ExcelFile('C:\\Users\\marse\\Desktop\\pays.xlsx')
df = xl.parse('גיליון1')
row = len(df)
col = len(df.columns)

for i in range(0, row):
    costumer = (df['לקוח'][i])
    maam = (df['מע"מ'][i])
    mas = (df['מקדמות מס הכנסה'][i])
    nikuim = (df['מ"ה ניכויים'][i])
    bl = (df['ב.ל. ניכויים'][i])
    dudi = (df['דודי'][i])

    message = f"Hello dear {costumer}, you need to pay maam: {maam}, mas: {mas}\n" \
              f"The total amount of nikuim is {nikuim}, bl: {bl}, dudi: {dudi}"

    print(message)
    print()

# print(df.apply(lambda x: x.last_valid_index()))
# df.info()
# print(row)
# print(col)
