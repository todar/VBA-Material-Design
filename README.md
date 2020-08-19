# VBA Material Design

<a href="https://www.buymeacoffee.com/todar" target="_blank"><img src="https://cdn.buymeacoffee.com/buttons/default-orange.png" alt="Buy Me A Coffee" style="height: 51px !important;width: 217px !important;" ></a>

Two Class modules that are used to format a VBA Userform in a similar style as materialize css. 

![materialize](https://github.com/todar/VBA-Materialize/blob/master/materialize.jpeg "Userform Image")

Example calling the class module from a userform:

```vb

Public mz As New materialize

'====================================================================
' INTITIALIZE/ACTIVATE
'====================================================================
Private Sub UserForm_Initialize()
    mz.Form Me
    mz.TextBox Textbox1, "Email", mzEmail, "Must be email format"
    mz.TextBox TextBox2, "Pin", mzNumeric, "Must be numeric"
    mz.Button btnSubmit, mzredlighten1, mzRedAccent2, mzwhite
End Sub

Private Sub UserForm_Activate()
    
    mz.setFocus Textbox1.name
    
End Sub

'DEMONSTRATE VALIDATE
Private Sub btnSubmit_Click()

    If mz.Validate(True) = True Then
        MsgBox "Form is valid"
    End If
    
End Sub

```
