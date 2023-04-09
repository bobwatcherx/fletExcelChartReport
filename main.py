from flet import *
# INSTALL openpyxl IN YoU PC WITH PIP
import openpyxl
from openpyxl.chart import BarChart,Reference,Series


def main(page:page):

	page.theme_mode = "light"

	alldata = [
	{"name":"juki","money":231321},
	{"name":"dada","money":4354},
	{"name":"oop","money":6545},
	{"name":"kii","money":243243},
	{"name":"weqe","money":878767},
	]

	# CREATE TABLE
	mytable = DataTable(
		columns=[
			DataColumn(Text("name")),
			DataColumn(Text("money")),
		],
		rows=[]
		)

	# NOW APPEND TO TABLE FROM alldata
	for x in alldata:
		mytable.rows.append(
			DataRow(
				cells=[
					DataCell(Text(x['name'])),
					DataCell(Text(x['money'])),
				]
				)
			)

	def exporttochart(e):
		wb = openpyxl.Workbook()
		ws = wb.active
		# YOU TITLE Chart
		ws.title ="Employee chart"

		# NOW LOOP alldata 
		for row in alldata:
			ws.append([row['name'],row['money']])
		# NOW SET BAR CHART 
		chart = BarChart()
		chart.title = "Employee chart result"
		chart.x_axis_title = "Name"
		chart.x_axis_title = "money"
		x_data = Reference(ws,min_col=1,min_row=1,max_row=len(alldata))
		y_data = Reference(ws,min_col=2,min_row=1,max_row=len(alldata))

		# NOW ADD DATA TO CHART
		chart.add_data(y_data)
		chart.set_categories(x_data)

		# NOW ADD CHART TO YOU EXCEL WORKSHEET
		ws.add_chart(chart,"D1")

		# AND NOW SAVE TO EXCEL FILE AND SAVE IN YoU DIRECTORY NOW
		wb.save("you_employee_file.xlsx")

		# AND SHOW SNACKBAR IF SUCCESS
		page.snack_bar = SnackBar(
			Text("success guys",size=30),
			bgcolor="green"
			)
		page.snack_bar.open = True
		page.update()



	page.add(
		Column([
	Text("Excel export to Chart ",size=30,weight="bold"),
	mytable,
	ElevatedButton("Export to Excel Chart",
	on_click=exporttochart
		)

			])

		)
flet.app(target=main)
