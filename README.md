<div align="center">

## Move forms together


</div>

### Description

Move forms together as if they were docked.

No Timer needed!

I had to separate forms used as a kind of toolbox. My idea was, that it would be nice if I could move these two together if they were both active. VB doesn't tell me when my form is moved. But then I realised that I read s.th. about moving forms without a titlebar. Using this trick my code will perform the following action:

1. The form detects a mousedown event

2. The mouse is released and the form moved by FormDrag

3. The other form is notified of the movement

4. The other form will follow the first

Don't forget to vote if you like it ;o)
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Andre Koester](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/andre-koester.md)
**Level**          |Beginner
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/andre-koester-move-forms-together__1-10780/archive/master.zip)

### API Declarations

```
Declare Sub ReleaseCapture Lib "user32" ()
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
```


### Source Code

```
'Put this in a global module
Public Sub FormDrag(TheForm As Form)
  ReleaseCapture
  Call SendMessage(TheForm.hwnd, &HA1, 2, 0&)
End Sub
'this code has to be in your form
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  FormDrag Me 'move form
  NameOfOtherForm.MoveMe 'notify other form
End Sub
'this is needed in the second form
Public Sub MoveMe()
  If Top > NameOfOtherForm.Top Then
    Top = NameOfOtherForm.Top + NewFrm.Height 'Place below other form
    Left = NameOfOtherForm.Left
  Else
    Top = NameOfOtherForm.Top - Height     'Place above other form
    Left = NameOfOtherForm.Left
  End If
End Sub
```

