<div align="center">

## Easy passing parameters to program


</div>

### Description

Often you have to pass some parameters(Password,UserName,...) into

the application.This code gives you elegant and easy way to pass

as many parameters as you want.
 
### More Info
 
Add to exe line: Project1.exe /u 'UserName' /p 'Password' /d 'domain'

Where UserName,Password,domain are strings

Common formula: project1.exe {/[letter] string}

Create new project.Add code to Form1.Make Project1.exe.Now you

can pass into the application 3 parameters:Username ,Password,Domain

(Whith little change you can pass as many parameters as you wont).


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[German Bletel](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/german-bletel.md)
**Level**          |Unknown
**User Rating**    |4.0 (8 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/german-bletel-easy-passing-parameters-to-program__1-2350/archive/master.zip)





### Source Code

```
Const PARAMHEADER = "/"
Public Function getTokens(CommandLine As String) As Collection
  Dim reminder As String
  Dim col As New Collection
  Dim pos As Integer
  Dim param As String
  Dim paramValue As String
  Dim paramName As String
  reminder = CommandLine
  pos = InStr(reminder, " ")
  Do While pos > 0
    param = Trim(Left(reminder, pos - 1))
    If (Left(param, 1) = PARAMHEADER) Then
      Call AddParamCol(col, paramValue, paramName)
      paramValue = ""
      paramName = Mid(param, 2)
    Else
      paramValue = param
    End If
    reminder = Trim(Mid(reminder, pos + 1))
    pos = InStr(reminder, " ")
  Loop
  paramValue = Trim(reminder)
  Call AddParamCol(col, paramValue, paramName)
  Set getTokens = col
End Function
Private Sub AddParamCol(c As Collection, s As String, k As String)
  If k = "" Then Exit Sub
  On Error Resume Next
  Call c.Add(s, LCase(k))
End Sub
'--------------------------------------
Private Sub Form1_Load()
  Dim Args As Collection
  Set Args = getTokens(Command)
  On Error Resume Next
    User = Args("u")
    Password = Args("p")
    Domain = Args("d")
  'Add your variables and actions
End Sub
```

