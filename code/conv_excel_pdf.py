from win32com import client
import os

def excl2pdf(file_location):
	app = client.DispatchEx("Excel.Application")
	app.Interactive = False
	app.Visible = False

	workbook = app.Workbooks.open(file_location)
	output = os.path.splitext(file_location)[0]
	
	workbook.ActiveSheet.ExportAsFixedFormat(0, output)
	workbook.Close()

file_location = "C:\\Users\\sami\\OneDrive\\Desktop\\proj\\Indv. Player Profile.xlsx"
excl2pdf(file_location)