Dim src_filename
Dim filesys
Dim cwd_name
Dim src_stream
Dim dst_stream
Dim supbook_line_str
Dim prev_line_str
Dim row_line_str
Dim line_str
Dim row_str_pos
Dim del_str_bgn
Dim del_str_end

' Determine the path name of the file to be accessed
src_filename = "MiracleNikkiJp_chart.xml"
set filesys = CreateObject("Scripting.FileSystemObject")
if InStr(src_filename, "\") = 0 then
    ' current directory
    cwd_name = filesys.getParentFolderName(WScript.ScriptFullName)
    src_filename = cwd_name & "\" & src_filename
end if

' Open source stream
set src_stream = CreateObject("ADODB.Stream")
src_stream.Type = 2  ' text
src_stream.Charset = "utf-8"
src_stream.Open
src_stream.LoadFromFile src_filename
' Open destination stream
set dst_stream = CreateObject("ADODB.Stream")
dst_stream.Type = 2  ' text
dst_stream.Charset = "utf-8"
dst_stream.Open
' message
WScript.echo("MiracleNikkiJp_chart.xmlの整形を開始します。しばらくかかる場合があります。終了したらお知らせします。")
'-----------------------------
prev_line_str = ""
row_line_str = ""
Do Until src_stream.EOS
  line_str = src_stream.ReadText(-2)    ' -1:whole, -2:line
  ' Process to be deleted
  if InStr(line_str, "<Author>") > 0 then line_str = ""
  if InStr(line_str, "<LastAuthor>") > 0 then line_str = ""
  if InStr(line_str, "<Created>") > 0 then line_str = ""
  if InStr(line_str, "<LastSaved>") > 0 then line_str = ""
  
  ' <SupBook> detect
  if InStr(line_str, "<SupBook>") > 0 then
    supbook_line_str = line_str
  end if
  ' <SupBook> mode
  if Len(supbook_line_str) > 0 then
    ' </SupBook>
    if InStr(line_str, "</SupBook>") > 0 then
      supbook_line_str = ""
    end if
    line_str = ""
  end if
  
  ' Concatenate if there is consolidation specification from the previous line
  if Len(prev_line_str) > 0 then
    line_str = prev_line_str & " " & LTrim(line_str)
    prev_line_str = ""
  end if
  
  ' Row tag detect
  row_str_pos = InStr(line_str, "<Row ") ' "<Row " position
  if row_str_pos > 0 then
    ' Error when detecting nested Row
    if Len(row_line_str) > 0 then
      WScript.echo("Error: Row nested")
      WScript.Quit
    end if
    ' Store Row tag and later in row_line_str
    row_line_str = Mid(line_str, row_str_pos)
    line_str = Left(line_str, row_str_pos - 1)
    ' If there is a valid description before the Row tag, it will be outputted as one line, otherwise it will be left behind the Row tag (if it is blank)
    if Len(Trim(line_str)) > 0 then
      dst_stream.WriteText line_str, 1
    else
      row_line_str = line_str & row_line_str
    end if
  end if
  
  ' Row-mode
  if Len(row_line_str) > 0 then
    row_line_str = row_line_str & Trim(line_str) ' Continue to the previous line
    if Right(row_line_str, 1) <> ">" then ' Preparation for connection when tag is separated into multiple lines
      row_line_str = row_line_str & " "
    end if
    ' <Row ... />
    row_str_pos = InStr(row_line_str, "/>")
    if row_str_pos > 0 then
'WScript.echo(InStr(Mid(row_line_str, InStr(row_line_str, "<Row ") + 5, row_str_pos - InStr(row_line_str, "<Row ") - 5), "<"))
'WScript.Quit
      if InStr(Mid(row_line_str, InStr(row_line_str, "<Row ") + 5, row_str_pos - InStr(row_line_str, "<Row ") - 5), "<") = 0 then
        ' Return up to /> to line_str and concatenate after /> to the next line
        line_str = Left(row_line_str, row_str_pos - 1 + 2)
        row_line_str = ""
        prev_line_str = Mid(row_line_str, row_str_pos + 2)
      end if
    end if
    ' </Row>
    row_str_pos = InStr(row_line_str, "</Row>")
    if row_str_pos > 0 then
      ' </ Row> tag is returned to line_str, and after the </ Row> tag is concatenated to the next line
      line_str = Left(row_line_str, row_str_pos - 1 + 6)
      prev_line_str = Mid(row_line_str, row_str_pos + 6)
      row_line_str = ""
    end if
  end if
  
  ' normal-mode
  if Len(row_line_str) <= 0 then
    if Right(RTrim(line_str), 1) = ">" then
      ' link update for items
      line_str = Replace(line_str, "MiracleNikkiJp_items.csv!", "'[MiracleNikkiJp_items.csv]MiracleNikkiJp_items'!")
      ' link update for balance
      line_str = Replace(line_str, "MiracleNikkiJp_balance.csv!", "'[MiracleNikkiJp_balance.csv]MiracleNikkiJp_balance'!")
      ' delete ss:Author="..."
      Do
        del_str_bgn = InStr(line_str, "ss:Author")
        if del_str_bgn > 0 then
          del_str_end = InStr(InStr(del_str_bgn + 9, line_str, """") + 1, line_str, """")
          line_str = RTrim(Left(line_str, del_str_bgn - 1)) & Mid(line_str, del_str_end + 1)
        end if
      Loop While del_str_bgn > 0
      ' delete <PhoneticText...</PhoneticText>
      Do
        del_str_bgn = InStr(line_str, "<PhoneticText")
        if del_str_bgn > 0 then
          del_str_end = InStr(del_str_bgn, line_str, "</PhoneticText>")
          line_str = RTrim(Left(line_str, del_str_bgn - 1)) & Mid(line_str, del_str_end + 15)
        end if
      Loop While del_str_bgn > 0
      ' wrie to stream
      dst_stream.WriteText line_str, 1    ' 0:string, 1:sring + CRLF
    else
      prev_line_str = RTrim(line_str)
    end if
  end if
Loop
'-----------------------------
' Close source stream
src_stream.Close
' Overwrite with utf8 code with BOM
dst_stream.SaveToFile src_filename, 2
' Close destination stream
dst_stream.Close
' message
WScript.echo("MiracleNikkiJp_chart.xmlの整形が終了しました")
