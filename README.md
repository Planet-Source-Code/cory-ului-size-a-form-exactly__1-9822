<div align="center">

## Size a form exactly\!

<img src="PIC200071715234313.jpg">
</div>

### Description

Sizes a form to an exact size. Ever tried to size a form to say. 320 by 240 like this

Me.Width = 320 * Screen.TwipsPerPixelX

Me.Height = 240 * Screen.TwipsPerPixelY

But you want the border and title bar at the same time. And it doesn't look right at all.

Heres your answer!

(I dunno if anyone has done this before if they have. cool.) I made this years ago, but didn't know of Planet-Source-Code then.
 
### More Info
 
'Either use

'I.E.

'Call Size(Me, 320, 240)

'or

'Size Me, 320, 240

'Same result.

'Size [Form], [Width], [Height]


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Cory Ului](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/cory-ului.md)
**Level**          |Beginner
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) 
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/cory-ului-size-a-form-exactly__1-9822/archive/master.zip)





### Source Code

```
Sub Size(sForm As Form, sWidth As Integer, sHeight As Integer)
Dim t_ScaleMode As Integer, t_Width As Integer, t_Height As Integer
 t_ScaleMode = sForm.ScaleMode
 sForm.ScaleMode = 1
 t_Width = sForm.Width - sForm.ScaleWidth
 t_Height = sForm.Height - sForm.ScaleHeight
 sForm.Width = (sWidth * Screen.TwipsPerPixelX) + t_Width
 sForm.Height = (sHeight * Screen.TwipsPerPixelY) + t_Height
 sForm.ScaleMode = t_ScaleMode
End Sub
```

