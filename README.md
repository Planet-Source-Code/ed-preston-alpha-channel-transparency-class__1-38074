<div align="center">

## Alpha Channel \(Transparency\) Class


</div>

### Description

Visual basic class for setting the layered window attribute and alpha channel of a form (or anything else with a window handle -> hWnd). Effectively makes the form and all its contents transparent. Implements configurable levels of transparency dictated by the caller. Pass the window handle of the form to set the alpha chanel. Turn alpha on and off using bolSetAs parameter. For the alpha value, 0 is completely transparent and 255 is opague. Be carefull when using values less than 100. Compatible with User32 DLL in Windows 2000 and XP.
 
### More Info
 
SetLayered(ByVal hwnd As Long, Byval bolSetAs As Boolean, Byval bAlpha As Byte)

ReleaseDisplay(ByVal hwnd As Long)

Compatible with User32 DLL in Windows 2000 and XP.

Be carefull when using values less than 100.


<span>             |<span>
---                |---
**Submitted On**   |2002-08-18 23:31:14
**By**             |[Ed Preston](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/ed-preston.md)
**Level**          |Advanced
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 6\.0
**Category**       |[Graphics](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/graphics__1-46.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Alpha\_Chan1195738182002\.zip](https://github.com/Planet-Source-Code/ed-preston-alpha-channel-transparency-class__1-38074/archive/master.zip)

### API Declarations

```
Private Declare Function SetWindowLong Lib "user32" _
  Alias "SetWindowLongA" _
  (ByVal hwnd As Long, _
  ByVal nIndex As Long, _
  ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" _
  Alias "GetWindowLongA" _
  (ByVal hwnd As Long, _
  ByVal nIndex As Long) As Long
Private Declare Function RedrawWindow Lib "user32" _
  (ByVal hwnd As Long, _
  lprcUpdate As RECT, _
  ByVal hrgnUpdate As Long, _
  ByVal fuRedraw As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" _
  (ByVal hwnd As Long, _
  ByVal crKey As Long, _
  ByVal bAlpha As Byte, _
  ByVal dwFlags As Long) As Long
```





