# VBA-Materialize (Currently in build status)
Two Class modules that are used to format a VBA Userform in a similar style as materialize css. 

![alt text]("https://raw.githubusercontent.com/todar/VBA-Materialize/master/materialize.JPG" "")

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
