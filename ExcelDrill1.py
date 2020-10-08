
import numpy as np
import pandas as pd

# Create a DataFrame
df = pd.DataFrame(np.random.randn(3, 2), columns=['Sales', 'Expenses'],
                  index=['2018', '2019', '2020'])

writer = pd.ExcelWriter('VSCode 2020 Sales Progress Report.xlsx', engine='xlsxwriter')
df.to_excel(writer, sheet_name='Sheet1', startrow=2)

# Create book and sheet objects
book = writer.book
sheet = writer.sheets['Sheet1']

# Create a Title
bold = book.add_format({'bold': True, 'size':24})
sheet.write('A1', '2020 Sales Progress', bold)

# Negative values in red
format1 = book.add_format({'font_color': '#bd0d0d'})
sheet.conditional_format('B4:C6', {'type': 'cell', 'criteria': '<=', 'value':0, 'format': format1})

# create a chart and choose type (the book is from writer.book above)
chart = book.add_chart({'type': 'column'})
chart.add_series({'values': '=Sheet1!B4:B6',
                  'name':'sheet1!B3',
                  'categories':'Sheet1!$A$4:$A$6',
                  'border': {'color': 'black'} })
chart.add_series({'values': '=Sheet1!C4:C6', 'name': '=Sheet1!C3'})


# Insert a chart at starting cell
sheet.insert_chart('C8', chart)

# Save!
writer.save()



























































