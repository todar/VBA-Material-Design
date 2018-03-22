VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "materialize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

'EVENT HANDLERS
Private WithEvents pForm As MSForms.UserForm
Attribute pForm.VB_VarHelpID = -1
Private mzEvents() As New materializeEvents 'ARRAY OF CUSTOM CLASS THAT HANDLES EVENTS

'EVERY TIME A CONTROL IS ADDED, IT WILL BE STORED IN THESE GROUPS
Private mzTxtGroup As New Collection
Private mzBtnGroup As New Collection
Private mzCmbGroup As New Collection

'SETTINGS
Public mzFocusColor As Long
Public mzButtonColor As Long
Public mzButtonHoverColor As Long
Public mzBorderColor As Long
Public mzPlaceholderColor As Long
Public mzWarningColor As Long
Public mzSuccessColor As Long

Public onButton As Boolean 'BOOLEAN TO LIMIT WHEN A LOOP IS NEEDED (HELPS PERFORMANCE)
Public liveValidation As Boolean
Private Const mzFontSizeChange As Double = 4
Private Const mzPlaceholderMaxPosition As Double = 15
Private Const mzBoderStartLocation As Double = 8
Private Const mzFramesPerSec As Double = 10 '13
Private Const mzAnimationSec As Double = 0.07 '0.01
Private Const mzPaddingOffset As Integer = 8 'ADD OFFSET TO CONTROLS DUE TO TEXTBOXES HAVING A WEIRD LEFT PADDING


'NOTES:
' - THIS CLASS IS USED TO EASILY FORMAT USERFORMS, AND ADD VALIDATION using .
' - THIS CLASS HEAVILY RELIES ON ITS SIBILING CLASS, materializeEvents
' - THE LOOK AND FUNCTION OF THIS IS BASED ON THE MATERIALIZECSS FRAMEWORK THAT IS
'   USED FOR WEBSITES: http://materializecss.com/  (SEE THE SECTION ON FORMS)
' - EXAMPLE @ https://github.com/todar/VBA-Materialize

'TODO:
' - THIS IS THE TESTING STAGE. CURRENTLY NEED TO GO BACK AND CLEAN UP CODE QUITE A BIT.
'   ADD MORE ERROR HANDLING.
' - CLEAN UP THE WAY VALIDATE WORKS. CHECKFORMAT & VALIDATE ARE A BIT REPETITIVE.
' - SHOULD SEPERATE BORDER, AND PLACEHODER INTO THEIR OWN FUNCTIONS. REPETITIVE CODE.
' - CURRENTLY, REGEX ARE PASSED AS PARAMETERS. NEED TO SET UP TO
'   USE FORMVALIDATION ENUM/VALIDATIONRETURN INSTEAD

'Funtions:
' - TextBox:        formats a textbox, adds it to the ptextgroup
' - DropDown:       formats combobox, adds it to the mzCmbGroup
' - Button:         formats button, adds it to the mzBtnGroup
' - setFocus:       called from mzevents, sets focus to a form cotrol (using animations)
' - checkFormat:
' - validate:
' - resetForm:
' - toast:          @TODO - work in progress... A temporary dialog.


'FORM ENUM, USED FOR CHOOSING VALIDATION FOR TEXTBOXES
Public Enum formValidation
    mzNone
    mzEmail
    mzOneDigit
    mzTwoDigits
    mzThreeDigits
    mzFourDigits
    mzYear
    mzDate
    mzNumeric
    mzCurrency
End Enum

'FUNCTION TO RETURN THE REGEX FROM FORMVALIDATION ENUMS
Function validationReturn(Value As formValidation) As String
    
    Select Case Value
        Case mzNone: validationReturn = ""
        Case mzEmail: validationReturn = "([\w\-\.]+)@([\w\-\.]+)\.[a-zA-Z]{2,5}"
        Case mzOneDigit: validationReturn = "^\d{1}$"
        Case mzTwoDigits: validationReturn = "^\d{2}$"
        Case mzThreeDigits: validationReturn = "^\d{3}$"
        Case mzFourDigits: validationReturn = "^\d{4}$"
        Case mzYear: validationReturn = "^(\d{4}|\d{2})$"
        Case mzDate: validationReturn = "^\d{1,2}[\s-\.\/]\d{1,2}[\s-\.\/]\d{2,4}$"
        Case mzNumeric: validationReturn = "^\d+$"
        Case mzCurrency: validationReturn = "^$|\d+"
    End Select
    
End Function


'INITIALIZATION, CALLED ONCE NEW INSTANCE OF MATERIALIZE IS CREATED
Private Sub Class_Initialize()

    'SET DEFAULT COLORS TO PUBLIC CONST COLORS
    mzFocusColor = mzTealLighten1
    mzBorderColor = mzGreyLighten2
    mzPlaceholderColor = mzGreyDarken1
    mzWarningColor = mzred
    mzSuccessColor = mzGreenLighten1
    mzButtonColor = mzTealLighten1
    mzButtonHoverColor = mztealHover
    
    liveValidation = True
End Sub


'*************************************************************************
'NOTES:
' - FORMATS A TEXTBOX - ALSO ADDS: PLACEHOLDER, BOTTOM BORDER
' - ADDED TO Private Const ptextgroup - THERE IT WILL BE CALLED ON FOR
'   ANIMATIONS AND VALIDATION
'*************************************************************************
Public Sub TextBox(txt As MSForms.TextBox, placeholder As String, Optional Validation As String, Optional WarningMessage As String)
    
    Dim mz As New Collection
    Dim i As Integer
    Dim lbl As MSForms.Label
    Dim frm As MSForms.frame
    Dim warn As MSForms.Label
    
    
'TODO: DON'T ASSUME THE TEXTBOX PARENT IS THE FORM... MUST UPDATE
    Set pForm = txt.Parent
    
    'ADD NEW CONTROLS TO THE PARENT ELEMENT
    Set lbl = txt.Parent.Controls.Add("Forms.Label.1", "lbl" & txt.Name)
    Set frm = txt.Parent.Controls.Add("Forms.Frame.1", "frm" & txt.Name)
    Set warn = txt.Parent.Controls.Add("Forms.Label.1", "warn" & txt.Name)
    
    'NULL VALUE SO THAT ON FROM THE FORM IN THE EDITOR YOU CAN LABEL EACH TEXTBOX.
    txt.Value = ""
    
    '''''''''''''''''''''''''''''''''''''''''
    ' FORMAT TEXTBOX
    '''''''''''''''''''''''''''''''''''''''''
    With txt
        .BackStyle = fmBackStyleTransparent
        .borderStyle = fmBorderStyleNone
    End With
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' FORMAT BOTTOM BORDER TO TEXTBOX (FRAME)
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    With frm
        .Width = txt.Width - mzPaddingOffset
        .borderStyle = fmBorderStyleSingle
        .borderStyle = fmBorderStyleNone
        .Left = txt.Left + mzPaddingOffset
        .tabStop = False
        .Top = txt.Top + txt.Height + mzBoderStartLocation
        .BackColor = mzBorderColor
        .Height = 1
    End With
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' FORMAT A LABEL PLACEHOLDER
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    With lbl
        .Height = txt.Height
        .Width = txt.Width
        .borderStyle = fmBorderStyleNone
        .BackStyle = fmBackStyleTransparent
        .Caption = placeholder
        .MousePointer = fmMousePointerIBeam
        .Left = txt.Left + mzPaddingOffset
        .ForeColor = mzPlaceholderColor
        .Top = txt.Top
        .Font.SIZE = txt.Font.SIZE
    End With
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' FORMAT A LABEL WARNING
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    With warn
        .Height = txt.Height
        .Width = txt.Width
        .borderStyle = fmBorderStyleNone
        .BackStyle = fmBackStyleTransparent
        .Caption = ""
        .MousePointer = fmMousePointerIBeam
        .Left = txt.Left + mzPaddingOffset
        .ForeColor = mzWarningColor
        .Top = frm.Top + 3
        .Font.SIZE = txt.Font.SIZE - mzFontSizeChange
    End With
    

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ADD TEXTBOX CONTROLS TO MATERIALIZE EVENTS
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ArrayIncDim mzEvents
    
    Set mzEvents(UBound(mzEvents, 1)).mz = Me
    Set mzEvents(UBound(mzEvents, 1)).txt = txt
    Set mzEvents(UBound(mzEvents, 1)).lbl = lbl
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' MZ IS A LOCAL COLLECTION THAT ADDS ALL THE CONTROLS JUST CREATED
    ' THIS COLLECTION IS THEN ADDED TO THE MAIN COLLECTION mzTxtGroup
    '
    ' AN EXAMPLE OF GETTING DATA FROM IT: mzTxtGroup(0)("TXT")
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    'ADD GROUP CONTROLS\VARIABLES TO THE LOCAL COLLECTION
    mz.Add txt, "txt"
    mz.Add lbl, "lbl"
    mz.Add frm, "frm"
    mz.Add warn, "warn"
    mz.Add Validation, "val"
    mz.Add WarningMessage, "warnMsg"
    
    'ADD LOCAL COLLECTION TO MAIN CONTROL GROUP COLLECTION (KEY = NAME OF TEXTBOX)
    mzTxtGroup.Add mz, txt.Name
    
End Sub

Public Function DropDown(cmb As MSForms.ComboBox)
    
    Dim mz As New Collection
    Dim lbl As MSForms.Label
    Dim frm As MSForms.frame

    Set lbl = cmb.Parent.Controls.Add("Forms.Label.1", "lbl" & cmb.Name)
    Set frm = cmb.Parent.Controls.Add("Forms.Frame.1", "frm" & cmb.Name)
    
    'FORMAT CMB
    With cmb
        .borderStyle = fmBorderStyleSingle
        .borderStyle = fmBorderStyleNone
        .ShowDropButtonWhen = fmShowDropButtonWhenNever
        .Style = fmStyleDropDownList
        .ForeColor = mzFocusColor
    End With
    
    'ADD A BOTTOM BORDER TO TEXTBOX (FRAME)
    With frm
        .Width = cmb.Width
        .borderStyle = fmBorderStyleSingle
        .borderStyle = fmBorderStyleNone
        .Left = cmb.Left + mzPaddingOffset
        .tabStop = False
        .Top = cmb.Top + cmb.Height + mzBoderStartLocation
        .BackColor = mzGreyLighten2
        .Height = 1
    End With
    
    'ADD A LABEL PLACEHOLDER
    With lbl
        .Height = cmb.Height
        .Width = cmb.Width
        .borderStyle = fmBorderStyleNone
        .BackStyle = fmBackStyleTransparent
        .Caption = "Issues"
        .MousePointer = fmMousePointerIBeam
        .Left = cmb.Left + mzPaddingOffset
        .ForeColor = mzGrey
        .Top = cmb.Top - mzPlaceholderMaxPosition
        .Font.Name = cmb.Font.Name
        .Font.SIZE = cmb.Font.SIZE - mzFontSizeChange
    End With
    
    'ADD TEXTBOX TO MATERIALIZE EVENTS
    ArrayIncDim mzEvents
    
    Set mzEvents(UBound(mzEvents, 1)).mz = Me
    Set mzEvents(UBound(mzEvents, 1)).cmb = cmb
    
    'ADD GROUP CONTROLS
    mz.Add cmb, "cmb"
    mz.Add frm, "frm"
    mz.Add False, "visited"

    'ADD GROUP TO MAIN CONTROL GROUP COLLECTION (KEY = NAME OF TEXTBOX)
    mzCmbGroup.Add mz, cmb.Name
    
End Function


'==============================================================
' STYLE AND SET UP A BUTTON
'==============================================================
Public Sub Button(btn As MSForms.CommandButton)
    
    Dim mz As New Collection
    
    ArrayIncDim mzEvents
    With btn
        btn.BackColor = mzButtonColor
        btn.ForeColor = mzWhite
    End With
    
    
    Set mzEvents(UBound(mzEvents, 1)).mz = Me
    Set mzEvents(UBound(mzEvents, 1)).btn = btn
    
    mz.Add btn, "btn"
    
    mzBtnGroup.Add mz, btn.Name
    
End Sub
Public Sub BtnHover(Optional btnName As String)
    
    Dim i As Integer
    Dim mz As Collection
    
    DoEvents 'ONLY LOOP IF MOUSE OVER BUTTON
    If onButton = True Then
        For i = 1 To mzBtnGroup.Count
            Set mz = mzBtnGroup(i)
            
            'HOVER
            If mz("btn").Name = btnName Then
                mz("btn").BackColor = mzButtonHoverColor
            
            'DEFAULT
            Else
                mz("btn").BackColor = mzButtonColor
            End If
            
        Next i
    End If
    
End Sub

'============================================================================================
' SETTING FOCUS TO A TEXTBOX - CALLS FOR ANIMATIONS, AS WELL AS DOUBLE CHECKS FORMATING
'============================================================================================
Public Sub setFocus(txt As String)
    
    Dim i As Integer
    Dim mz As Collection
    
    Dim obj As New Collection
    Dim Properties As New Collection
    Dim Increment As New Collection
    
    'IN CASE TEXBOX IS NOT ENABLED
    On Error GoTo catch
    
    For i = 1 To mzTxtGroup.Count
        Set mz = mzTxtGroup(i)
        
        If mz("txt").Name = txt Then
            
            
            If mz("lbl").Top = mz("txt").Top Then
            
                'ADD TO FOCUS GROUP
                obj.Add mz("lbl")
                Properties.Add "top"
                Increment.Add (mzPlaceholderMaxPosition * -1)
                
                'ADD TO FONT SIZE GROUP
                obj.Add mz("lbl")
                Properties.Add "fontsize"
                Increment.Add (mzFontSizeChange * -1)
                
            End If
            
        ElseIf mz("txt").Value = "" And mz("lbl").Top < mz("txt").Top And mz("txt").Name <> txt Then
            
            'ADD TO LOOSE FOUCUS GROUP
            obj.Add mz("lbl")
            Properties.Add "top"
            Increment.Add mzPlaceholderMaxPosition
            
           'FONT SIZE
            obj.Add mz("lbl")
            Properties.Add "fontsize"
            Increment.Add mzFontSizeChange
               
            mz("warn") = ""
               
        End If
        
    Next i
    
    If obj.Count > 0 Then
        animation obj, Properties, Increment, mzFramesPerSec, mzAnimationSec
    End If
    
    checkFormat txt
    
catch:
    
End Sub


'==========================================================================================
' LABEL SIZE IS NOT ALWAYS ACCURATE, SO THIS WILL DOUBLE CHECK THE CORRECT SIZE.
' ALSO HIGHLIGHTS THE BOTTOM BORDER(FRAME) (DOES VALIDATION AS WELL FOR NON CURRENT TXTBOX)
'==========================================================================================
Public Function checkFormat(Optional curTxtName As String) As Boolean
    
    Dim i As Integer
    Dim txt As MSForms.TextBox
    Dim frm As MSForms.frame
    Dim lbl As MSForms.Label
    Dim Val As String
    
    
    'CHECK EACH TEXTBOX
    For i = 1 To mzTxtGroup.Count
        
        On Error GoTo Nxt
        Set txt = mzTxtGroup(i)("txt")
        Set frm = mzTxtGroup(i)("frm")
        Set lbl = mzTxtGroup(i)("lbl")
        Val = mzTxtGroup(i)("val")
        
        
        'FORMAT
        If txt.Name = curTxtName Then
            lbl.Top = txt.Top - mzPlaceholderMaxPosition
            lbl.Font.SIZE = txt.Font.SIZE - mzFontSizeChange
            lbl.ForeColor = mzFocusColor
            frm.Height = 2
            If txt.Value = "" Then frm.BackColor = mzFocusColor
            
            
        ElseIf txt.Value = "" Then
            lbl.Top = txt.Top
            lbl.Font.SIZE = txt.Font.SIZE
            lbl.ForeColor = mzPlaceholderColor
            frm.Height = 1
            frm.BackColor = mzBorderColor
            
        Else
        
            If lbl.Top = txt.Top Then
                lbl.Top = txt.Top - mzPlaceholderMaxPosition
                lbl.Font.SIZE = txt.Font.SIZE - mzFontSizeChange

            End If
            
            lbl.ForeColor = mzPlaceholderColor
            
            
            If Val <> "" And Trim(txt.Value) <> "" And liveValidation Then
            'CHECK TO SEE IF IT IS OK TO VALIDATE
                If regularExpressionExists(Trim(txt.Value), Val) = True Then
                    frm.BackColor = mzSuccessColor 'VALID
                    mzTxtGroup(i)("warn").Caption = ""
                Else
                    frm.BackColor = mzWarningColor 'INVALID, RETURN FALSE
                    mzTxtGroup(i)("warn").Caption = mzTxtGroup(i)("warnMsg")
                End If

                frm.Height = 2
            Else
                frm.Height = 1
                frm.BackColor = mzBorderColor
            End If
            
        End If
        
        'BOTTOM BORDER COLOR (FRAME)
        If Val <> "" Then
        
        Else
            If txt.Name = curTxtName Then
                frm.BackColor = mzBorderColor
            Else
                frm.BackColor = mzFocusColor
            End If
            
            mzTxtGroup(i)("warn").Caption = ""
        End If
        
Nxt:
        
    Next i
    
    
End Function



'==========================================================================================
' VALIDATES ALL THE FIEDLS THAT HAD A REG EXPRESSION PASSED IN IT
'==========================================================================================
Public Function Validate(Optional markEmpty As Boolean) As Boolean

    Dim i As Integer
    Dim txt As MSForms.TextBox
    Dim frm As MSForms.frame
    Dim cmb As MSForms.ComboBox
    Dim Val As String
    Dim visited As Boolean
    
    liveValidation = True
    Validate = True
    
    'VALIDATE TEXTBOX GROUP
    For i = 1 To mzTxtGroup.Count
    
        Set txt = mzTxtGroup(i)("txt")
        Set frm = mzTxtGroup(i)("frm")
        Val = mzTxtGroup(i)("val")
        
        If txt.Value <> "" Or markEmpty = True Then
        
        If Val <> "" Then

                'VALIDATE
                If regularExpressionExists(Trim(txt.Value), Val) = True Then
                    frm.BackColor = mzSuccessColor 'VALID
                    mzTxtGroup(i)("warn").Caption = ""
                Else
                    frm.BackColor = mzWarningColor 'INVALID, RETURN FALSE
                    Validate = False
                    mzTxtGroup(i)("warn").Caption = mzTxtGroup(i)("warnMsg")
                End If

                frm.Height = 2

        End If
        
        
        End If
        
Nxt:
    Next i
    
    'VALIDATE COMBOBOX GROUP
    For i = 1 To mzCmbGroup.Count
        
        Set cmb = mzCmbGroup(i)("cmb")
        Set frm = mzCmbGroup(i)("frm")
        visited = mzCmbGroup(i)("visited")
        
        If visited = True Or markEmpty = True Then
            If cmb.Value <> "" Then
                frm.BackColor = mzSuccessColor
            Else
                frm.BackColor = mzWarningColor
                Validate = False
            End If
            frm.Height = 2
        End If
        
    Next i
    
    
End Function


Private Function animation(mzObj As Variant, mzProperty As Variant, mzIncrement As Variant, framesPerSec As Integer, _
sec As Double)
    
    'Transition obj, arr, 13, 0.2, Inc
    Dim i As Integer
    Dim n As Integer
    Dim currentValue As Double
    Dim obj As Variant
    Dim objProperty As Variant
    Dim Increment As Variant
    
    Dim Counter As Double
    
    'CONVERT TO ARRAY. THAT WAY COLLECTIONS, ARRAYS, SINGLE PARAMETERS, EXT. CAN ALL BE PASSED IN
    obj = cArray(mzObj)
    objProperty = cArray(mzProperty)
    Increment = cArray(mzIncrement)
    
    'CALCULATE INCREMENT STEPS/TIME STEPS
    sec = (sec * 1000)
    sec = sec / framesPerSec
  
    For i = 1 To framesPerSec
        
        DoEvents
        
        'LOOP EACH CONTROL AND INCREMENT ACCORDINGLY
        For n = LBound(obj, 1) To UBound(obj, 1)
        
            'GET CURRENTVALUE
            currentValue = CallByName(obj(n), objProperty(n), VbGet)
            
            'INCREMENT PROPERTY VALUE
            CallByName obj(n), objProperty(n), VbLet, currentValue + Increment(n) / framesPerSec
            
            Counter = Counter + (Increment(n) / framesPerSec)
            
        Next n
        
        'API SLEEP FUNCTION (MILLISECONDS)
        Sleep sec
        
    Next i
    
End Function


Private Sub pForm_Click()
    checkFormat
End Sub


Private Sub pForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    BtnHover
    onButton = False
End Sub


Public Sub resetForm()
        
    Dim i As Integer
    
    For i = 1 To mzTxtGroup.Count
        mzTxtGroup(i)("txt").Value = ""
    Next i
    
    For i = 1 To mzCmbGroup.Count
        'mzCmbGroup(i)("cmb").value = Null
        mzCmbGroup(i)("cmb").ListIndex = 0
        mzCmbGroup(i)("frm").Height = 1
        mzCmbGroup(i)("frm").BackColor = mzBorderColor
    Next i
    

End Sub

Public Sub toast(S As String, mil As Long, form As MSForms.UserForm, bgColor As Long)

    Dim txt As MSForms.TextBox
    Dim L As Long
    Static running As Boolean
    
    If running = False Then
        running = True
    
        Set txt = form.Controls.Add("Forms.Textbox.1", "toast" & S)
        
        DoEvents
        With txt
            
            .AutoSize = True
            .SpecialEffect = fmSpecialEffectFlat
            .Value = S
            .MultiLine = True
            .BackColor = bgColor
            .ForeColor = mzWhite
            .Font.SIZE = 14
            .Left = (form.InsideWidth / 1.1) - .Width
            .Top = (form.InsideHeight / 6.5) + L
            
        End With
        
        animation txt, "top", -35, 7, 0.1
        
        Sleep mil
        
        animation txt, "top", -35, 7, 0.1
        form.Controls.Remove txt.Name
        
        running = False
   
    End If
    
End Sub

