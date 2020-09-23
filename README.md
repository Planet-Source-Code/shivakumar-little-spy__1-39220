<div align="center">

## Little SPY


</div>

### Description

It gets the Window Text,ClassName,HWND for any window,the mouse pointer points.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2002-05-18 23:43:32
**By**             |[shivakumar](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/shivakumar.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Little\_SPY1353249242002\.zip](https://github.com/Planet-Source-Code/shivakumar-little-spy__1-39220/archive/master.zip)

### API Declarations

```
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
```





