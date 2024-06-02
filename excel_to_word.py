from docxtpl import DocxTemplate
from datetime import datetime
import pandas as pd

doc = DocxTemplate("certifikat.docx")

today_date = datetime.today().strftime("%d/%m/%Y")

# name = "ime"

# potrdilo = "potrdilo"

# obrazovanje = "obrazovanje"

# ljubljana = "ljubljana"

# my_context = { 'company_name' : "World Company", 'datum' : today_date, name : "ime test", potrdilo : "potrdilo test", obrazovanje : "obrazovanje test", ljubljana : "ljubljana test"}

# df = pd.read_csv('spisak.xlsx',   encoding='ISO-8859-1', on_bad_lines='skip', engine='python')
                 #, encoding='latin1', header=None, on_bad_lines='skip')
                 # on_bad_lines='skip', encoding='ISO-8859-1')   #"ISO-8859-1") header=0, index_col=None,

df = pd.read_excel('spisak.xlsx')

for index, row in df.iterrows():
#    print(index)
 #   print(row.get)
    context = {
        'ime' : row['ime'],
        'certifikat' : row['certifikat'],
        'obrazovanje' : row['obrazovanje'],
       # '(bez taga)' : row['(bez taga)'],
        'datum' : row['datum'],
        'grad' : row['zagreb'].strftime('%d.%m.%Y.')
    }

    doc.render(context)

    doc.save(f"test/{row['ime']}.docx")