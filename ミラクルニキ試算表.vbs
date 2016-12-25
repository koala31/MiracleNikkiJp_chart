Dim filesys
Dim cwd_name
Dim items_filename
Dim chart_filename
Dim excel_app
Dim items_book
Dim chart_book

' Determine the path name of the file to be accessed
set filesys = CreateObject("Scripting.FileSystemObject")
cwd_name = filesys.getParentFolderName(WScript.ScriptFullName)
items_filename = cwd_name & "\" & "MiracleNikkiJp_items.csv"
chart_filename = cwd_name & "\" & "MiracleNikkiJp_chart.xml"

' execute Excel
set excel_app = WScript.CreateObject("Excel.Application")
excel_app.Visible = True
set items_book = excel_app.Workbooks.Open(items_filename)
set chart_book = excel_app.Workbooks.Open(chart_filename)

