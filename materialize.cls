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
Private mzEvents() As New materializeEvents 'CUSTOM EVENT HANDLER
Attribute mzEvents.VB_VarHelpID = -1

'VARIABLES
Private pTxtGroup As New Collection 'STORES COLLECTIONS, THAT ARE THE GROUPS OF CONTROLS
Private pBtnGroup As New Collection 'STORES COLLECTIONS, THAT ARE THE GROUPS OF CONTROLS

'SETTINGS
Public pFoucusColor As Long
Public pBtnColor As Long
Public pBtnHoverColor As Long
Public pBorderColor As Long
Public pLabelColor As Long
Public pWarningColor As Long
Public pSuccessColor As Long

Public onButton As Boolean 'BOOLEAN TO LIMIT WHEN A LOOP IS NEEDED (HELPS PERFORMANCE)
Private Const FontSizeReduce As Double = 4
Private Const FontPosIncrement As Double = 15
Private Const BorderLocation As Double = 7
Private Const pFramesPerSec As Double = 13
Private Const pAnimationSec As Double = 0.01

'TODO:
' - THIS IS THE TESTING STAGE. CURRENTLY NEED TO GO BACK AND CLEAN UP CODE QUITE A BIT.
' - ADD MORE ERROR HANDLING.
' - MAKE SEVERAL VARIABLES GLOBAL CONSTANTS(COLORS, FONT SIZE) TO MAKE IT EASIER TO CHANGE.


Private Sub Class_Initialize()
    pFoucusColor = mzteal
    pBorderColor = mzGreyLighten2
    pLabelColor = mzGrey
    pWarningColor = mzRedAccent1
    pSuccessColor = mzGreenLighten1
    pBtnColor = mzteal
    pBtnHoverColor = mzTealLighten1
End Sub


'==============================================================
' STYLE AND SET UP A TEXTBOX
'==============================================================
Public Sub TextBox(txt As MSForms.TextBox, placeholder As String, Optional Validation As String)
    
    Dim mz As New Collection
    Dim i As Integer
    Dim lbl As MSForms.Label
    Dim frm As MSForms.frame
    
    Set pForm = txt.Parent
    
    Set lbl = txt.Parent.Controls.Add("Forms.Label.1", "lbl" & txt.Name)
    Set frm = txt.Parent.Controls.Add("Forms.Frame.1", "frm" & txt.Name)
    
    txt.value = ""
    
    'FORMAT TEXTBOX
    With txt
        .BackStyle = fmBackStyleTransparent
        .borderStyle = fmBorderStyleNone
    End With
    
    'ADD A BOTTOM BORDER TO TEXTBOX (FRAME)
    With frm
        .Width = txt.Width
        .borderStyle = fmBorderStyleSingle
        .borderStyle = fmBorderStyleNone
        .Left = txt.Left + 8
        .tabStop = False
        .Top = txt.Top + txt.Height + BorderLocation
        .BackColor = pBorderColor
        .Height = 1
    End With
    
    'ADD A LABEL PLACEHOLDER
    With lbl
        .Height = txt.Height
        .Width = txt.Width
        .borderStyle = fmBorderStyleNone
        .BackStyle = fmBackStyleTransparent
        .Caption = placeholder
        .MousePointer = fmMousePointerIBeam
        .Left = txt.Left + 8
        .ForeColor = pLabelColor
        .Top = txt.Top
        .Font.SIZE = txt.Font.SIZE
    End With
    
    'ADD TEXTBOX TO MATERIALIZE EVENTS
    ArrayIncDim mzEvents
    
    Set mzEvents(UBound(mzEvents, 1)).mz = Me
    Set mzEvents(UBound(mzEvents, 1)).txt = txt
    Set mzEvents(UBound(mzEvents, 1)).lbl = lbl
    
    'ADD GROUP CONTROLS
    mz.Add txt, "txt"
    mz.Add lbl, "lbl"
    mz.Add frm, "frm"
    mz.Add Validation, "val"
    
    'ADD GROUP TO MAIN CONTROL GROUP COLLECTION (KEY = NAME OF TEXTBOX)
    pTxtGroup.Add mz, txt.Name
    
End Sub

'==============================================================
' STYLE AND SET UP A BUTTON
'==============================================================
Public Sub Button(btn As MSForms.CommandButton)
    
    Dim mz As New Collection
    
    ArrayIncDim mzEvents
    With btn
        btn.BackColor = pBtnColor
        btn.ForeColor = mzWhite
    End With
    
    
    Set mzEvents(UBound(mzEvents, 1)).mz = Me
    Set mzEvents(UBound(mzEvents, 1)).btn = btn
    
    mz.Add btn, "btn"
    
    pBtnGroup.Add mz, btn.Name
    
    
End Sub
Public Sub BtnHover(Optional btnName As String)
    
    Dim i As Integer
    Dim mz As Collection
    
    DoEvents 'ONLY LOOP IF MOUSE OVER BUTTON
    If onButton = True Then
        For i = 1 To pBtnGroup.Count
            Set mz = pBtnGroup(i)
            
            'HOVER
            If mz("btn").Name = btnName Then
                mz("btn").BackColor = pBtnHoverColor
            
            'DEFAULT
            Else
                mz("btn").BackColor = pBtnColor
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
    
    For i = 1 To pTxtGroup.Count
        Set mz = pTxtGroup(i)
        
        If mz("txt").Name = txt And mz("lbl").Top = mz("txt").Top Then
            
            'ADD TO FOCUS GROUP
            obj.Add mz("lbl")
            Properties.Add "top"
            Increment.Add (FontPosIncrement * -1)
            
            'FONT SIZE
            obj.Add mz("lbl")
            Properties.Add "fontsize"
            Increment.Add (FontSizeReduce * -1)
            
        ElseIf mz("txt").value = "" And mz("lbl").Top < mz("txt").Top And mz("txt").Name <> txt Then
            
            'ADD TO LOOSE FOUCUS GROUP
            obj.Add mz("lbl")
            Properties.Add "top"
            Increment.Add FontPosIncrement
            
           'FONT SIZE
            obj.Add mz("lbl")
            Properties.Add "fontsize"
            Increment.Add FontSizeReduce
            
        End If
        
    Next i
    
    If obj.Count > 0 Then
        Animation obj, Properties, Increment, pFramesPerSec, pAnimationSec
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
    For i = 1 To pTxtGroup.Count
        
        On Error GoTo Nxt
        Set txt = pTxtGroup(i)("txt")
        Set frm = pTxtGroup(i)("frm")
        Set lbl = pTxtGroup(i)("lbl")
        Val = pTxtGroup(i)("val")
        
        'FORMAT
        If txt.Name = curTxtName Then
            lbl.Top = txt.Top - FontPosIncrement
            lbl.Font.SIZE = txt.Font.SIZE - FontSizeReduce
            lbl.ForeColor = pFoucusColor
            frm.Height = 2
            If txt.value = "" Then frm.BackColor = pFoucusColor
            
        ElseIf txt.value = "" Then
            lbl.Top = txt.Top
            lbl.Font.SIZE = txt.Font.SIZE
            lbl.ForeColor = pLabelColor
            frm.Height = 1
            frm.BackColor = pBorderColor

        Else
        
            If lbl.Top = txt.Top Then
                lbl.Top = txt.Top - FontPosIncrement
                lbl.Font.SIZE = txt.Font.SIZE - FontSizeReduce
            End If
            
            lbl.ForeColor = pLabelColor
            
            If Val <> "" And Trim(txt.value) <> "" Then
            'CHECK TO SEE IF IT IS OK TO VALIDATE
                If regularExpressionExists(Trim(txt.value), Val) = True Then
                    frm.BackColor = pSuccessColor 'VALID
                Else
                    frm.BackColor = pWarningColor 'INVALID, RETURN FALSE
                End If

                frm.Height = 2
            End If
        End If
        
        'BOTTOM BORDER COLOR (FRAME)
        If Val <> "" Then
        
        Else
            If txt.Name = curTxtName Then
                frm.BackColor = pBorderColor
            Else
                frm.BackColor = pFoucusColor
            End If
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
    Dim Val As String

    Validate = True
    
    For i = 1 To pTxtGroup.Count
    
        Set txt = pTxtGroup(i)("txt")
        Set frm = pTxtGroup(i)("frm")
        Val = pTxtGroup(i)("val")
        
        
        If Val <> "" Then

            'CHECK TO SEE IF IT IS OK TO VALIDATE
            If Trim(txt.value) <> "" Then

                'VALIDATE
                If regularExpressionExists(Trim(txt.value), Val) = True Then
                    frm.BackColor = pSuccessColor 'VALID
                Else
                    frm.BackColor = pWarningColor 'INVALID, RETURN FALSE
                    Validate = False
                End If

                frm.Height = 2

            Else

                'VALUE IS "", SO THEREFORE NOT VALID
                If markEmpty = True Then frm.BackColor = pWarningColor
                Validate = False
            End If

        End If
Nxt:
    Next i


End Function


Private Function Animation(mzObj As Variant, mzProperty As Variant, mzIncrement As Variant, framesPerSec As Integer, _
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
    sec = (sec * 1000) / framesPerSec
    
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
