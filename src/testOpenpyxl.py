import pandas as pd
import openpyxl
from openpyxl.chart import LineChart
from openpyxl.chart import Reference
from openpyxl.utils.dataframe import dataframe_to_rows

test_data = pd.read_csv("./csv/test.csv")


""" ---------------------------------------------------------------------------
create workbook
--------------------------------------------------------------------------- """
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "test"

""" ---------------------------------------------------------------------------
read pandas.DataFrame
--------------------------------------------------------------------------- """
for r in dataframe_to_rows(test_data, index=True, header=True):
    ws.append(r)

""" ---------------------------------------------------------------------------
chart property
--------------------------------------------------------------------------- """
x_values = Reference(ws, min_row=2, min_col=1,  max_row=test_data.shape[1]+1)
y_values = Reference(ws, min_row=1, min_col=2,  max_row=test_data.shape[1]+1, max_col=test_data.shape[1]+1)
x_title = "Name"
y_title = None

""" ---------------------------------------------------------------------------
make chart
--------------------------------------------------------------------------- """
chart = LineChart()
chart.title = "test"
chart.style = 13
chart.x_axis.title = x_title
chart.y_axis.title = y_title
chart.set_categories(x_values)
chart.add_data(y_values, titles_from_data=True)
ws.add_chart(chart, "F1")

""" ---------------------------------------------------------------------------
output
--------------------------------------------------------------------------- """
wb.save("./xlsx/test.xlsx")
