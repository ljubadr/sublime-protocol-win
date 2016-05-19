' As this was my first experience with VBScript, I apologize to anyone looking at this code

' Test this script in command prompt
'   open_file.vbs "subl://open?url=file://<path_to_file>&line=4"

' path to sublime_text.exe
Dim sublime_path

' I've tried to start sublime with `subl <path-to-file>:<line>`, but in my case it was
'   much slower than opening directly with sublime

' >>> EDIT THIS PATH TO MATCH YOUR SUBLIME INSTALL PATH / OR PATH TO SUBLIME PORTABLE <<<
sublime_path = "C:\Program Files (x86)\Sublime Text 3\sublime_text.exe"

Dim text, decoded
' get first command line argument
text         = WScript.Arguments.Item(0)

' decode URL
decoded      = URLDecode(text)

Dim get_params, params
' Split 
get_params  = Split(decoded, "?")(1)
params      = Split(get_params, "&")

Dim file_name, line
For Each field In params

  If InStr(field, "url=") > 0 Then

    file_name = Replace(field, "url=", "")

    If InStr(file_name, "file://") > 0 Then
      file_name = Replace(file_name, "file://", "")
    End If

  ElseIf InStr(field, "line=") > 0 Then
    line = Replace(field, "line=", "")
  End If

Next

' Final command
Dim run_command
run_command = """"&sublime_path&""" """&file_name&":"&line&""""

' For debugging, echo generated command
' WScript.Echo run_command
' Wscript.Quit

Dim objShell
Set objShell = WScript.CreateObject( "WScript.Shell" )
objShell.Run(run_command)
Set objShell = Nothing

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' URLDecode from http://www.motobit.com/tips/detpg_URLDecode/
',,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,

Function URLDecode(ByVal What)
'URL decode Function
'2001 Antonin Foller, PSTRUH Software, http://www.motobit.com
  Dim Pos, pPos

  'replace + To Space
  What = Replace(What, "+", " ")

  on error resume Next
  Dim Stream: Set Stream = CreateObject("ADODB.Stream")
  If err = 0 Then 'URLDecode using ADODB.Stream, If possible
    on error goto 0
    Stream.Type = 2 'String
    Stream.Open

    'replace all %XX To character
    Pos = InStr(1, What, "%")
    pPos = 1
    Do While Pos > 0
      Stream.WriteText Mid(What, pPos, Pos - pPos) + _
        Chr(CLng("&H" & Mid(What, Pos + 1, 2)))
      pPos = Pos + 3
      Pos = InStr(pPos, What, "%")
    Loop
    Stream.WriteText Mid(What, pPos)

    'Read the text stream
    Stream.Position = 0
    URLDecode = Stream.ReadText

    'Free resources
    Stream.Close
  Else 'URL decode using string concentation
    on error goto 0
    'UfUf, this is a little slow method. 
    'Do Not use it For data length over 100k
    Pos = InStr(1, What, "%")
    Do While Pos>0 
      What = Left(What, Pos-1) + _
        Chr(Clng("&H" & Mid(What, Pos+1, 2))) + _
        Mid(What, Pos+3)
      Pos = InStr(Pos+1, What, "%")
    Loop
    URLDecode = What
  End If
End Function