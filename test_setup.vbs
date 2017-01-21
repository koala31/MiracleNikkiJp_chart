Dim src_filename
Dim dst_filename
Dim filesys
Dim cwd_name
Dim src_stream
Dim dst_stream
Dim src_str

' Determine the path name of the file to be accessed
src_filename = "test_msg.vbs.utf8"
dst_filename = "test_msg.vbs"

set filesys = CreateObject("Scripting.FileSystemObject")
cwd_name = filesys.getParentFolderName(WScript.ScriptFullName)
if InStr(src_filename, "\") = 0 then
    ' current directory
    src_filename = cwd_name & "\" & src_filename
end if
if InStr(dst_filename, "\") = 0 then
    ' current directory
    dst_filename = cwd_name & "\" & dst_filename
end if

' message
WScript.echo(src_filename & " Çutf-8Ç≈äJÇ´Ç‹Ç∑")
' Open source stream
set src_stream = CreateObject("ADODB.Stream")
src_stream.Type = 2  ' text
src_stream.Charset = "utf-8"
src_stream.Open
src_stream.LoadFromFile src_filename

' Open destination stream
set dst_stream = CreateObject("ADODB.Stream")
dst_stream.Type = 2  ' text
dst_stream.Charset = "shift_jis"
dst_stream.Open
' read whole
src_str = src_stream.ReadText(-1)
' write to stream
dst_stream.WriteText src_str, 0
' Close source stream
src_stream.Close
' message
WScript.echo(dst_filename & " Çshijt_jisÇ≈ï€ë∂ÇµÇ‹Ç∑")
' Write
dst_stream.SaveToFile dst_filename, 2
' Close destination stream
dst_stream.Close
' message
WScript.echo dst_filename & " Çé¿çsÇµÇ‹Ç∑"

Dim wshell
set wshell = WScript.CreateObject("WSCript.shell")
wshell.Run dst_filename, 1, TRUE

' message
WScript.echo("èIóπÇµÇ‹ÇµÇΩ")
