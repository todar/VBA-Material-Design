# VBA-Materialize 

Two Class modules that are used to format a VBA Userform in a similar style as materialize css. 

![materialize](https://github.com/todar/VBA-Materialize/blob/master/materialize.jpeg "Userform Image")

Example calling the class module from a userform:

```vb

Public mz As New materialize

'====================================================================
' INTITIALIZE/ACTIVATE
'====================================================================
Private Sub UserForm_Initialize()
    
    'FUNCTION TO SET FORM TO THE CENTER OF THE CURRENT EXCEL APPLICATION
    mz.CenterFormToApplication Me
    
    'MZ IS A CLASS THAT FORMATS FORMS, AND ADDS VALIDATION USING REGULAR EXPRESSIONS
    mz.TextBox division, "Division", mzTwoDigits, "Two Digits"
    mz.TextBox facility, "Facility", mzFourDigits, "Four Digits"
    mz.TextBox userId, "User Id", mzCharactersThenNumbers, "Characters then numbers"
    mz.TextBox po, "PO", mzNumeric, "Numeric"
    mz.TextBox period, "Period", mzTwoDigits, "Two Digits"
    mz.TextBox pYear, "Year", mzYear, "Year must be two or four digits"
    mz.TextBox buyerName, "Buyer Name", mzCharacters, "incorrect"
    mz.TextBox paidback, "Payback Amount (optional)", mzCurrency, "Must be Currency"
    mz.TextBox findings, "Findings (optional) ", mzCurrency, "Must be Currency"
    
    'ADD ITEMS TO THE ISSUE COMBOBOX
    issue.AddItem "Dropdown Item 1"
    issue.AddItem "Dropdown Item 2"
    
    mz.DropDown issue
    
    mz.Button btnSubmit
    mz.Button btnUpload
    
End Sub

Private Sub UserForm_Activate()
    
    'CAN ADD VALUES, AND CHECKFORMAT WILL INIT FORM
    pYear = Year(Now)
    userId = Environ("username")
    mz.checkFormat division.name
    
End Sub

```
