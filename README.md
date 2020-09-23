<div align="center">

## Is a certain form open?


</div>

### Description

Simple API call that tells whether or not a particular form is open. Useful for managing popup forms or a series of forms.

Sample usage:

If FormIsOpen("Color Picker") Then

'   ....Do Something Here

Else

'   ...Do Something Else Here...

End If
 
### More Info
 
FormName

Goes by form title, so multiple instances of a form can be used as long as their titles are different. Also works with other applications...any window in fact, so it is not limited to your VB forms.

Boolean. True=Form is open.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Matthew Roberts](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/matthew-roberts.md)
**Level**          |Beginner
**User Rating**    |4.3 (13 globes from 3 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/matthew-roberts-is-a-certain-form-open__1-30621/archive/master.zip)

### API Declarations

```
Private Declare Function FindWindow Lib "user32" Alias _
 "FindWindowA" (ByVal lpClassName As String, ByVal _
 lpWindowName As String) As Long
```


### Source Code

```
Public Function FormIsOpen(FormCaption) As Boolean
  FormIsOpen = FindWindow(vbNullString, FormCaption)
End Function
```

