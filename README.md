<div align="center">

## Alphablended \(Transparent\) Windows


</div>

### Description

This code will demonstrate how to create a window with partial transparency. The color of the window and the amount of transparency (in percent, from 0 to 100% transparent) are customizable. The code involves no screen capturing, it is actually an alphablended window. The window contents update in realtime as the window is moved, sized, or manipulated in any way. This only works in Windows 2000/XP.
 
### More Info
 
This code only works in Windows 2000 or Windows XP. Expect all new versions of windows that come out to support this later on.


<span>             |<span>
---                |---
**Submitted On**   |2002-01-21 22:07:08
**By**             |[Michael Barnathan](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/michael-barnathan.md)
**Level**          |Intermediate
**User Rating**    |4.6 (46 globes from 10 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Graphics](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/graphics__1-46.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Alphablend503261212002\.zip](https://github.com/Planet-Source-Code/michael-barnathan-alphablended-transparent-windows__1-31050/archive/master.zip)

### API Declarations

```
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
```





