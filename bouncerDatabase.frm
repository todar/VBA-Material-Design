VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} bouncerDatabase 
   Caption         =   "Buyer Report Card"
   ClientHeight    =   8655
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10320
   OleObjectBlob   =   "bouncerDatabase.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "bouncerDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

'CUSTOM CLASS THAT FORMATS\VALIDATES FORMS
Public mz As New materialize


'====================================================================
' INTITIALIZE/ACTIVATE
'====================================================================
Private Sub UserForm_Initialize()
    
    'FUNCTION TO SET FORM TO THE CENTER OF THE CURRENT EXCEL APPLICATION
    mz.CenterFormToApplication Me
    
    'ADD ITEMS TO THE ISSUE COMBOBOX
    issue.AddItem ""
    issue.AddItem "VC Error"
    issue.AddItem "OI Netted"
    issue.AddItem "OI Error"
    issue.AddItem "Routed Improperly"
    issue.AddItem "Unessesary Should Be Lines"
    issue.AddItem "SA Error"
    issue.AddItem "PCA Error"
    
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


Private Sub btnSubmit_Click()
    If mz.Validate(True) = True Then
        MsgBox "Form is valid!"
    Else
        MsgBox "Please Make sure it is valid first!"
    End If
End Sub






