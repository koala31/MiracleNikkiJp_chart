Dim filesys
Dim cwd_name
Dim items_filename
Dim balance_filename
Dim chart_filename
Dim chart_filename_xlsx
Dim excel_app
Dim items_book
Dim balance_book
Dim chart_book

On Error Resume Next

' Determine the path name of the file to be accessed
set filesys = CreateObject("Scripting.FileSystemObject")
cwd_name = filesys.getParentFolderName(WScript.ScriptFullName)
items_filename = cwd_name & "\" & "MiracleNikkiJp_items.csv"
balance_filename = cwd_name & "\" & "MiracleNikkiJp_balance.csv"
chart_filename = cwd_name & "\" & "MiracleNikkiJp_chart.xml"
chart_filename_xlsx = cwd_name & "\" & "MiracleNikkiJp_chart.xlsx"

' execute Excel
set excel_app = WScript.CreateObject("Excel.Application")
excel_app.Visible = True

' open item sheet
if filesys.FileExists(items_filename) then
    set items_book = excel_app.Workbooks.Open(items_filename)
else
    WScript.echo("not found " & items_filename)
end if
' oepn balance sheet
if filesys.FileExists(balance_filename) then
    set items_book = excel_app.Workbooks.Open(balance_filename)
else
    WScript.echo("not found " & balance_filename)
end if
' oepn chart sheet
if filesys.FileExists(chart_filename_xlsx) then
    set items_book = excel_app.Workbooks.Open(chart_filename_xlsx)
else
    if filesys.FileExists(chart_filename) then
        set items_book = excel_app.Workbooks.Open(chart_filename)
    else
        WScript.echo("not found " & chart_filename)
    end if
end if

