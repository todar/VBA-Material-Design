# VBA-Materialize (Currently in build status)

![Codeship](https://img.shields.io/codeship/d6c1ddd0-16a3-0132-5f85-2e35c05e22b1.svg)
[![license](https://img.shields.io/github/license/electron-userland/electron-forge.svg)](https://github.com/todar/VBA-Materialize/blob/master/LICENSE)

Two Class modules that are used to format a VBA Userform in a similar style as materialize css. 

![materialize](https://github.com/todar/VBA-Materialize/blob/master/materialize.JPEG "Userform Image")



Example calling the class module from a userform:

```vb

Public mz As New materialize

'====================================================================
' INTITIALIZE/ACTIVATE
'====================================================================
Private Sub UserForm_Initialize()
    
    'MZ.TEXTBOX TAKES IN A TEXTBOX FOR REFRENCE, A STRING FOR THE PLACEHOLDER, AND A REGULAR EXPRESSION FOR VALIDATION
    mz.TextBox division, "Division", "^\d{1,2}$"
    mz.TextBox facility, "Facility", "^\d{4}$"
    mz.TextBox userId, "User Id", "^[a-zA-Z]+\d+$"
    mz.TextBox po, "PO", "^\d+$"
    mz.TextBox period, "Period", "^\d{2,}$"
    mz.TextBox pYear, "Year", "^(\d{4}|\d{2})$"
    mz.TextBox buyerId, "Buyer Id", "\w+"
    mz.Button btnSubmit
    mz.Button btnUpload
    
End Sub

```
