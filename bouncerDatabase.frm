VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} bouncerDatabase 
   Caption         =   "Buyer Report Card"
   ClientHeight    =   7845
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

Public mz As New materialize

'====================================================================
' INTITIALIZE/ACTIVATE
'====================================================================
Private Sub UserForm_Initialize()
    
    CenterForm Me
 
    mz.TextBox division, "Division", "^\d{1,2}$"
    mz.TextBox facility, "Facility", "^\d{4}$"
    mz.TextBox userId, "User Id", "^[a-zA-Z]+\d+$"
    mz.TextBox po, "PO", "^\d+$"
    mz.TextBox period, "Period", "^\d{2,}$"
    mz.TextBox pYear, "Year", "^(\d{4}|\d{2})$"
    mz.TextBox buyerId, "Buyer Id", "\w+"
    
    mz.DropDown issues, "Issues", "test"
    mz.Button btnSubmit
    mz.Button btnUpload
    
End Sub

Private Sub UserForm_Activate()
    
    pYear = Year(Now)
    userId = Environ("username")
    mz.checkFormat division.Name
    
End Sub


Private Sub btnSubmit_Click()
    If mz.Validate(True) = True Then
        MsgBox "Running..."
    Else
        MsgBox "Please Make sure it is valid first!"
    End If
End Sub

'Private Sub UserForm_Initialize()
'
'    'Dim Options As Variant
'
'    'Options = getOptions
'    CenterForm Me
'    'fc.Init Me
'    fc.TextBox division, "Division", "^\d{1,2}$", True
'    fc.TextBox facility, "Facility", "^\d{4}$"
'    fc.TextBox userId, "User Id", "^[a-zA-Z]+\d+$", False, Environ("Username")
'    fc.TextBox po, "PO", "^\d+$"
'    fc.TextBox period, "Period", "^\d{2,}$"
'    fc.TextBox pYear, "Year", "^(\d{4}|\d{2})$", , Year(Now)
'    fc.TextBox buyerId, "Buyer Id", "\w+"
''    fc.DropDown issues, "Issues", Options
''
''
''    fc.Button btnUpload
''    fc.Button btnSubmit
''    fc.Validate
'
'    userId.Enabled = False
'
'End Sub



'Private Function getOptions() As Variant
'
'    Dim arr As New cArray
'
'    arr.push "Improper routing"
'    arr.push "Unnecessary Should Be Lines"
'
'    getOptions = arr.pArray
'
'End Function




'Private Sub btnSubmit_Click()
'    If fc.Validate(True) = True Then
'        MsgBox "Running..."
'    Else
'        MsgBox "Please Make sure it is valid first!"
'    End If
'
'End Sub

'Private Sub btnUpload_Click()
'
'    Dim pacs As New cPacs
'    Dim SQL As String
'    Dim arr As Variant
'    Dim I As Integer
'
'    pacs.ActivateExtra
'
'    'GET FROM PACS
'    If pacs.CopyActivePage("Bouncer Communication") Then
'        'On Error Resume Next
'        division.value = regularExpression(StringBetween(pacs.PacsLineValue(1), "Div", "P/O"), "\d{2}")(0)
'        facility.value = regularExpression(StringBetween(pacs.PacsLineValue(2), "Fac", "Dst Cntr"), "\d{4}")(0)
'        po.value = StringBetween(pacs.PacsLineValue(1), "P/O Num ", "Foreign P/O")
'    End If
'
'    'GET FROM QUERIES
'    SQL = ReadTextFile(SQLFolderPath & "period dates.txt")
'    arr = SqlQuery(SQL)
'
'    For I = LBound(arr, 1) To UBound(arr, 1)
'        If Now >= arr(I, 1) And Now <= arr(I, 2) Then
'            pYear.value = arr(I, 4)
'            period.value = StringProperLength(CStr(arr(I, 5)), 2, "0", False)
'            Exit For
'        End If
'    Next I
'
'    fc.Validate
'
'End Sub











