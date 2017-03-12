Option Explicit

' Determine the path name of the file to be accessed
Dim src_filename
src_filename = "MiracleNikkiJp_script.utf8"
Dim filesys
set filesys = CreateObject("Scripting.FileSystemObject")
If InStr(src_filename, "\") = 0 Then
    ' current directory
    Dim cwd_name
    cwd_name = filesys.getParentFolderName(WScript.ScriptFullName)
    src_filename = cwd_name & "\" & src_filename
End If
If Not filesys.FileExists(src_filename) Then
    WScript.echo("ファイルが見つかりません。 MiracleNikkiJp_script.utf8")
    WScript.Quit(-1)
End If
Dim dst_filename
dst_filename = Left(src_filename, Len(src_filename) - 4) & "vbs"
' Open source stream
Dim src_stream
Set src_stream = CreateObject("ADODB.Stream")
src_stream.Type = 2  ' text
src_stream.Charset = "utf-8"
src_stream.Open
src_stream.LoadFromFile src_filename
' Open destination stream
Dim dst_stream
Set dst_stream = CreateObject("ADODB.Stream")
dst_stream.Type = 2  ' text
dst_stream.Charset = "shift_jis"
dst_stream.Open
' copy stream
src_stream.CopyTo dst_stream
' Close source stream
src_stream.Close
' Overwrite with utf8 code with BOM
dst_stream.SaveToFile dst_filename, 2
' Close destination stream
dst_stream.Close
' build command line
Dim strCommand
strCommand = "wscript.exe " & dst_filename & " setup"
' Run vbs
Dim objShell
Set objShell = WScript.CreateObject("WScript.Shell")
objShell.Run strCommand, 1, False
