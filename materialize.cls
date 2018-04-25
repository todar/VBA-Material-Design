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

#If VBA7 And Win64 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

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
' - THIS CLASS IS USED TO EASILY FORMAT USERFORMS, AND ADD VALIDATION USING REGEXP .
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
    mzCharactersThenNumbers
    mzCharacters
End Enum

'COLOR ENUM (CAN MOVE THIS TO A MODULE TO USE THE COLORS FOR OTHER THINGS)
Public Enum mzColors

    mzRedLighten5 = 15657983
    mzRedLighten4 = 13815295
    mzRedLighten3 = 10132207
    mzRedLighten2 = 7566309
    mzRedLighten1 = 5264367
    mzred = 3556340
    mzRedDarken1 = 3488229
    mzRedDarken2 = 3092435
    mzRedDarken3 = 2631878
    mzRedDarken4 = 1842359
    mzRedAccent1 = 8424191
    mzRedAccent2 = 5395199
    mzRedAccent3 = 4462591
    mzRedAccent4 = 213
    mzPinkLighten5 = 15525116
    mzPinkLighten4 = 13679608
    mzPinkLighten3 = 11636724
    mzPinkLighten2 = 9593584
    mzPinkLighten1 = 8012012
    mzPink = 6495977
    mzPinkDarken1 = 6298584
    mzPinkDarken2 = 5970114
    mzPinkDarken3 = 5706925
    mzPinkDarken4 = 5181064
    mzPinkAccent1 = 11239679
    mzPinkAccent2 = 8470783
    mzPinkAccent3 = 5701877
    mzPinkAccent4 = 6427077
    mzPurpleLighten5 = 16115187
    mzPurpleLighten4 = 15187681
    mzPurpleLighten3 = 14193614
    mzPurpleLighten2 = 13134010
    mzPurpleLighten1 = 12339115
    mzPurple = 11544476
    mzPurpleDarken1 = 11150478
    mzPurpleDarken2 = 10624891
    mzPurpleDarken3 = 10099562
    mzPurpleDarken4 = 9180234
    mzPurpleAccent1 = 16548074
    mzPurpleAccent2 = 16466144
    mzPurpleAccent3 = 16318677
    mzPurpleAccent4 = 16711850
    mzDeepPurpleLighten5 = 16181229
    mzDeepPurpleLighten4 = 15320273
    mzDeepPurpleLighten3 = 14392755
    mzDeepPurpleLighten2 = 13464981
    mzDeepPurpleLighten1 = 12736382
    mzDeepPurple = 12008039
    mzDeepPurpleDarken1 = 11613534
    mzDeepPurpleDarken2 = 11021649
    mzDeepPurpleDarken3 = 10495813
    mzDeepPurpleDarken4 = 9575217
    mzDeepPurpleAccent1 = 16746675
    mzDeepPurpleAccent2 = 16731516
    mzDeepPurpleAccent3 = 16719717
    mzDeepPurpleAccent4 = 15335522
    mzIndigoLighten5 = 16181992
    mzIndigoLighten4 = 15321797
    mzIndigoLighten3 = 14330015
    mzIndigoLighten2 = 13338233
    mzIndigoLighten1 = 12610396
    mzIndigo = 11882815
    mzIndigoDarken1 = 11225401
    mzIndigoDarken2 = 10436400
    mzIndigoDarken3 = 9647400
    mzIndigoDarken4 = 8266522
    mzIndigoAccent1 = 16752268
    mzIndigoAccent2 = 16674131
    mzIndigoAccent3 = 16669245
    mzIndigoAccent4 = 16666416
    mzBlueLighten5 = 16642787
    mzBlueLighten4 = 16506555
    mzBlueLighten3 = 16370320
    mzBlueLighten2 = 16168292
    mzBlueLighten1 = 16098626
    mzblue = 15963681
    mzBlueDarken1 = 15042590
    mzBlueDarken2 = 13792793
    mzBlueDarken3 = 12608789
    mzBlueDarken4 = 10569485
    mzBlueAccent1 = 16757122
    mzBlueAccent2 = 16747076
    mzBlueAccent3 = 16742697
    mzBlueAccent4 = 16736809
    mzLightBlueLighten5 = 16709089
    mzLightBlueLighten4 = 16573875
    mzLightBlueLighten3 = 16438401
    mzLightBlueLighten2 = 16237391
    mzLightBlueLighten1 = 16168489
    mzLightBlue = 16034051
    mzLightBlueDarken1 = 15047427
    mzLightBlueDarken2 = 13731842
    mzLightBlueDarken3 = 12416770
    mzLightBlueDarken4 = 10180353
    mzLightBlueAccent1 = 16767104
    mzLightBlueAccent2 = 16761920
    mzLightBlueAccent3 = 16756736
    mzLightBlueAccent4 = 15372544
    mzCyanLighten5 = 16447456
    mzCyanLighten4 = 15920050
    mzCyanLighten3 = 15392384
    mzCyanLighten2 = 14798925
    mzCyanLighten1 = 14337574
    mzCyan = 13941760
    mzCyanDarken1 = 12692480
    mzCyanDarken2 = 10983168
    mzCyanDarken3 = 9405184
    mzCyanDarken4 = 6578176
    mzCyanAccent1 = 16777092
    mzCyanAccent2 = 16776984
    mzCyanAccent3 = 16770304
    mzCyanAccent4 = 13940736
    mzTealLighten5 = 15856352
    mzTealLighten4 = 14409650
    mzTealLighten3 = 12897152
    mzTealLighten2 = 11318861
    mzTealLighten1 = 10135078
    mzteal = 8951296
    mztealHover = 11385643
    mzTealDarken1 = 8096000
    mzTealDarken2 = 7043328
    mzTealDarken3 = 6056192
    mzTealDarken4 = 4214016
    mzTealAccent1 = 15466407
    mzTealAccent2 = 14352228
    mzTealAccent3 = 11987229
    mzTealAccent4 = 10862336
    mzGreenLighten5 = 15332840
    mzGreenLighten4 = 13231816
    mzGreenLighten3 = 10999461
    mzGreenLighten2 = 8701825
    mzGreenLighten1 = 6994790
    mzGreen = 5287756
    mzGreenDarken1 = 4694083
    mzGreenDarken2 = 3968568
    mzGreenDarken3 = 3308846
    mzGreenDarken4 = 2121243
    mzGreenAccent1 = 13301433
    mzGreenAccent2 = 11464809
    mzGreenAccent3 = 7792128
    mzGreenAccent4 = 5490688
    mzLightGreenLighten5 = 15333617
    mzLightGreenLighten4 = 13168092
    mzLightGreenLighten3 = 10871237
    mzLightGreenLighten2 = 8508846
    mzLightGreenLighten1 = 6671516
    mzLightGreen = 4899723
    mzLightGreenDarken1 = 4371324
    mzLightGreenDarken2 = 3710824
    mzLightGreenDarken3 = 3115861
    mzLightGreenDarken4 = 1993011
    mzLightGreenAccent1 = 9502668
    mzLightGreenAccent2 = 5898162
    mzLightGreenAccent3 = 262006
    mzLightGreenAccent4 = 1564004
    mzLimeLighten5 = 15203321
    mzLimeLighten4 = 12842224
    mzLimeLighten3 = 10284774
    mzLimeLighten2 = 7727068
    mzLimeLighten1 = 5759444
    mzLime = 3792077
    mzLimeDarken1 = 3394240
    mzLimeDarken2 = 2864303
    mzLimeDarken3 = 2399646
    mzLimeDarken4 = 1537922
    mzLimeAccent1 = 8519668
    mzLimeAccent2 = 4325358
    mzLimeAccent3 = 65478
    mzLimeAccent4 = 60078
    mzYellowLighten5 = 15203839
    mzYellowLighten4 = 12909055
    mzYellowLighten3 = 10352127
    mzYellowLighten2 = 7795199
    mzYellowLighten1 = 5828351
    mzYellow = 3927039
    mzYellowDarken1 = 3528957
    mzYellowDarken2 = 2998523
    mzYellowDarken3 = 2468089
    mzYellowDarken4 = 1540085
    mzYellowAccent1 = 9306111
    mzYellowAccent2 = 65535
    mzYellowAccent3 = 60159
    mzYellowAccent4 = 55039
    mzAmberLighten5 = 14809343
    mzAmberLighten4 = 11791615
    mzAmberLighten3 = 8577279
    mzAmberLighten2 = 5232127
    mzAmberLighten1 = 2673407
    mzAmber = 508415
    mzAmberDarken1 = 46079
    mzAmberDarken2 = 41215
    mzAmberDarken3 = 36863
    mzAmberDarken4 = 28671
    mzAmberAccent1 = 8381951
    mzAmberAccent2 = 4249599
    mzAmberAccent3 = 50431
    mzAmberAccent4 = 44031
    mzOrangeLighten5 = 14742527
    mzOrangeLighten4 = 11723007
    mzOrangeLighten3 = 8441087
    mzOrangeLighten2 = 5093375
    mzOrangeLighten1 = 2533375
    mzOrange = 39167
    mzOrangeDarken1 = 36091
    mzOrangeDarken2 = 31989
    mzOrangeDarken3 = 27887
    mzOrangeDarken4 = 20966
    mzOrangeAccent1 = 8442367
    mzOrangeAccent2 = 4238335
    mzOrangeAccent3 = 37375
    mzOrangeAccent4 = 28159
    mzDeepOrangeLighten5 = 15198715
    mzDeepOrangeLighten4 = 12373247
    mzDeepOrangeLighten3 = 9546751
    mzDeepOrangeLighten2 = 6654719
    mzDeepOrangeLighten1 = 4419839
    mzDeepOrange = 2250751
    mzDeepOrangeDarken1 = 1987060
    mzDeepOrangeDarken2 = 1657574
    mzDeepOrangeDarken3 = 1393624
    mzDeepOrangeDarken4 = 800447
    mzDeepOrangeAccent1 = 8429311
    mzDeepOrangeAccent2 = 4222719
    mzDeepOrangeAccent3 = 15871
    mzDeepOrangeAccent4 = 11485
    mzBrownLighten5 = 15330287
    mzBrownLighten4 = 13159639
    mzBrownLighten3 = 10791612
    mzBrownLighten2 = 8358049
    mzBrownLighten1 = 6516365
    mzBrown = 4740473
    mzBrownDarken1 = 4279405
    mzBrownDarken2 = 3620957
    mzBrownDarken3 = 3028046
    mzBrownDarken4 = 2303806
    mzGreyLighten5 = 16448250
    mzGreyLighten4 = 16119285
    mzGreyLighten3 = 15658734
    mzGreyLighten2 = 14737632
    mzGreyLighten1 = 12434877
    mzGrey = 10395294
    mzGreyDarken1 = 7697781
    mzGreyDarken2 = 6381921
    mzGreyDarken3 = 4342338
    mzGreyDarken4 = 2171169
    mzBlueGreyLighten5 = 15855596
    mzBlueGreyLighten4 = 14473423
    mzBlueGreyLighten3 = 12959408
    mzBlueGreyLighten2 = 11445392
    mzBlueGreyLighten1 = 10260600
    mzBlueGrey = 9141600
    mzBlueGreyDarken1 = 8023636
    mzBlueGreyDarken2 = 6576709
    mzBlueGreyDarken3 = 5195575
    mzBlueGreyDarken4 = 3682854
    mzWhite = 16777215
    mzblack = 0

End Enum

'FUNCTION TO RETURN THE REGEX FROM FORMVALIDATION ENUMS
Function validationReturn(value As formValidation) As String
    
    Select Case value
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
        Case mzCharactersThenNumbers: validationReturn = "^[a-zA-Z]+\d+$"
        Case mzCharacters: validationReturn = "\w+"
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


Public Sub Form(frm)

    Set pForm = frm
    
    'SET FORM TO THE CENTER OF THE APPLICATION
    With frm
        .StartUpPosition = 0
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With
    
End Sub


'*************************************************************************
'NOTES:
' - FORMATS A TEXTBOX - ALSO ADDS: PLACEHOLDER, BOTTOM BORDER
' - ADDED TO Private Const ptextgroup - THERE IT WILL BE CALLED ON FOR
'   ANIMATIONS AND VALIDATION
'*************************************************************************
Public Sub TextBox(txt As MSForms.TextBox, placeholder As String, Optional Validation As formValidation, Optional WarningMessage As String)
    
    Dim mz As New Collection
    Dim i As Integer
    Dim lbl As MSForms.Label
    Dim frm As MSForms.Frame
    Dim warn As MSForms.Label
    
    
'TODO: DON'T ASSUME THE TEXTBOX PARENT IS THE FORM... MUST UPDATE
    'Set pForm = txt.Parent
    
    'ADD NEW CONTROLS TO THE PARENT ELEMENT
    Set lbl = txt.Parent.Controls.Add("Forms.Label.1", "lbl" & txt.name)
    Set frm = txt.Parent.Controls.Add("Forms.Frame.1", "frm" & txt.name)
    Set warn = txt.Parent.Controls.Add("Forms.Label.1", "warn" & txt.name)
    
    'NULL VALUE SO THAT ON FROM THE FORM IN THE EDITOR YOU CAN LABEL EACH TEXTBOX.
    txt.value = ""
    
    '''''''''''''''''''''''''''''''''''''''''
    ' FORMAT TEXTBOX
    '''''''''''''''''''''''''''''''''''''''''
    With txt
        .BackStyle = fmBackStyleTransparent
        .borderStyle = fmBorderStyleSingle
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
    Set mzEvents(UBound(mzEvents, 1)).frm = frm
    mzEvents(UBound(mzEvents, 1)).controlType = "textbox"
    
    
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
    mz.Add validationReturn(Validation), "val"
    mz.Add WarningMessage, "warnMsg"
    
    'ADD LOCAL COLLECTION TO MAIN CONTROL GROUP COLLECTION (KEY = NAME OF TEXTBOX)
    mzTxtGroup.Add mz, txt.name
    
End Sub

Public Function DropDown(cmb As MSForms.ComboBox)
    
    Dim mz As New Collection
    Dim lbl As MSForms.Label
    Dim frm As MSForms.Frame

    Set lbl = cmb.Parent.Controls.Add("Forms.Label.1", "lbl" & cmb.name)
    Set frm = cmb.Parent.Controls.Add("Forms.Frame.1", "frm" & cmb.name)
    
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
        .Font.name = cmb.Font.name
        .Font.SIZE = cmb.Font.SIZE - mzFontSizeChange
    End With
    
    'ADD TEXTBOX TO MATERIALIZE EVENTS
    ArrayIncDim mzEvents
    
    Set mzEvents(UBound(mzEvents, 1)).mz = Me
    Set mzEvents(UBound(mzEvents, 1)).cmb = cmb
    Set mzEvents(UBound(mzEvents, 1)).lbl = lbl
    Set mzEvents(UBound(mzEvents, 1)).frm = frm
    mzEvents(UBound(mzEvents, 1)).controlType = "dropdown"
    
    'ADD GROUP CONTROLS
    mz.Add cmb, "cmb"
    mz.Add lbl, "lbl"
    mz.Add frm, "frm"
    mz.Add False, "visited"

    'ADD GROUP TO MAIN CONTROL GROUP COLLECTION (KEY = NAME OF TEXTBOX)
    mzCmbGroup.Add mz, cmb.name
    
End Function


'==============================================================
' STYLE AND SET UP A BUTTON
'==============================================================
Public Sub Button(btn As MSForms.CommandButton, Optional bgColor As Long = -1, Optional hoverColor As Long = -1, Optional fontColor As Long = -1)
    
    Dim mz As New Collection
    
    ArrayIncDim mzEvents
    
    'GET PROPER BACKGROUND COLOR
    If bgColor > -1 Then
        btn.BackColor = bgColor
        mz.Add bgColor, "bgColor"
    Else
        btn.BackColor = mzButtonColor
        mz.Add mzButtonColor, "bgColor"
    End If
    
    
    'GET PROPER HOVER COLOR
    If hoverColor > -1 Then
        mz.Add hoverColor, "hoverColor"
    Else
        mz.Add mzButtonHoverColor, "hoverColor"
    End If
    
    'GET FONT COLOR
    If fontColor > -1 Then
        btn.ForeColor = fontColor
    Else
        btn.ForeColor = mzWhite
    End If
    
    Set mzEvents(UBound(mzEvents, 1)).mz = Me
    Set mzEvents(UBound(mzEvents, 1)).btn = btn
    
    mz.Add btn, "btn"

    mzBtnGroup.Add mz, btn.name
    
End Sub
Public Sub BtnHover(Optional btnName As String)
    
    Dim i As Integer
    Dim mz As Collection
    
    
    DoEvents 'ONLY LOOP IF MOUSE OVER BUTTON
    If onButton = True Then
        For i = 1 To mzBtnGroup.Count
            Set mz = mzBtnGroup(i)
            
            'HOVER
            If mz("btn").name = btnName Then
                mz("btn").BackColor = mz("hoverColor")
            
            'DEFAULT
            Else
                mz("btn").BackColor = mz("bgColor")
            End If
            
        Next i
    End If
    
End Sub
'Public Sub Button(btn As MSForms.CommandButton, Optional mainColor As Long, Optional hoverColor As Long, Optional fgColor As Long)
'
'    Dim mz As New Collection
'
'    ArrayIncDim mzEvents
'
'    With btn
'        btn.BackColor = mzButtonColor
'        btn.ForeColor = mzWhite
'    End With
'
'
'    Set mzEvents(UBound(mzEvents, 1)).mz = Me
'    Set mzEvents(UBound(mzEvents, 1)).btn = btn
'
'    mz.Add btn, "btn"
'
'    mzBtnGroup.Add mz, btn.name
'
'End Sub
'Public Sub BtnHover(Optional btnName As String)
'
'    Dim i As Integer
'    Dim mz As Collection
'
'
'    DoEvents 'ONLY LOOP IF MOUSE OVER BUTTON
'    If onButton = True Then
'        For i = 1 To mzBtnGroup.Count
'            Set mz = mzBtnGroup(i)
'
'            'HOVER
'            If mz("btn").name = btnName Then
'                mz("btn").BackColor = mzButtonHoverColor
'
'            'DEFAULT
'            Else
'                mz("btn").BackColor = mzButtonColor
'            End If
'
'        Next i
'    End If
'
'End Sub

'============================================================================================
' SETTING FOCUS TO A TEXTBOX - CALLS FOR ANIMATIONS, AS WELL AS DOUBLE CHECKS FORMATING
'============================================================================================
Public Sub setFocus(txt As String)
    
    Dim i As Integer
    Dim mz As Collection
    
    Dim obj As New Collection
    Dim Properties As New Collection
    Dim increment As New Collection
    
    'IN CASE TEXBOX IS NOT ENABLED
    On Error GoTo catch
    
    For i = 1 To mzTxtGroup.Count
        Set mz = mzTxtGroup(i)
        
        If mz("txt").name = txt Then
            
            
            If mz("lbl").Top = mz("txt").Top Then
            
                'ADD TO FOCUS GROUP
                obj.Add mz("lbl")
                Properties.Add "top"
                increment.Add (mzPlaceholderMaxPosition * -1)
                
                'ADD TO FONT SIZE GROUP
                obj.Add mz("lbl")
                Properties.Add "fontsize"
                increment.Add (mzFontSizeChange * -1)
                
            End If
            
        ElseIf mz("txt").value = "" And mz("lbl").Top < mz("txt").Top And mz("txt").name <> txt Then
            
            'ADD TO LOOSE FOUCUS GROUP
            obj.Add mz("lbl")
            Properties.Add "top"
            increment.Add mzPlaceholderMaxPosition
            
           'FONT SIZE
            obj.Add mz("lbl")
            Properties.Add "fontsize"
            increment.Add mzFontSizeChange
               
            mz("warn") = ""
               
        End If
        
    Next i
    
    If obj.Count > 0 Then
        animation obj, Properties, increment, mzFramesPerSec, mzAnimationSec
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
    Dim frm As MSForms.Frame
    Dim lbl As MSForms.Label
    Dim val As String
    
    
    'CHECK EACH TEXTBOX
    For i = 1 To mzTxtGroup.Count
        
        On Error GoTo Nxt
        Set txt = mzTxtGroup(i)("txt")
        Set frm = mzTxtGroup(i)("frm")
        Set lbl = mzTxtGroup(i)("lbl")
        val = mzTxtGroup(i)("val")
        
        
        'FORMAT
        If txt.name = curTxtName Then
            lbl.Top = txt.Top - mzPlaceholderMaxPosition
            lbl.Font.SIZE = txt.Font.SIZE - mzFontSizeChange
            lbl.ForeColor = mzFocusColor
            frm.Height = 2
            If txt.value = "" Then frm.BackColor = mzFocusColor
            
            
        ElseIf txt.value = "" Then
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
            
            
            If val <> "" And Trim(txt.value) <> "" And liveValidation Then
            'CHECK TO SEE IF IT IS OK TO VALIDATE
                If regularExpressionExists(Trim(txt.value), val) = True Then
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
        If val <> "" Then
        
        Else
            If txt.name = curTxtName Then
                frm.BackColor = mzFocusColor
                
            Else
                frm.BackColor = mzBorderColor
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
    Dim frm As MSForms.Frame
    Dim cmb As MSForms.ComboBox
    Dim val As String
    Dim visited As Boolean
    
    liveValidation = True
    Validate = True
    
    'VALIDATE TEXTBOX GROUP
    For i = 1 To mzTxtGroup.Count
    
        Set txt = mzTxtGroup(i)("txt")
        Set frm = mzTxtGroup(i)("frm")
        val = mzTxtGroup(i)("val")
        
        If txt.value <> "" Or markEmpty = True Then
        
        If val <> "" Then

                'VALIDATE
                If regularExpressionExists(Trim(txt.value), val) = True Then
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
            If cmb.value <> "" Then
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
    Dim CurrentValue As Double
    Dim obj As Variant
    Dim objProperty As Variant
    Dim increment As Variant
    
    Dim Counter As Double
    
    'CONVERT TO ARRAY. THAT WAY COLLECTIONS, ARRAYS, SINGLE PARAMETERS, EXT. CAN ALL BE PASSED IN
    obj = cArray(mzObj)
    objProperty = cArray(mzProperty)
    increment = cArray(mzIncrement)
    
    'CALCULATE INCREMENT STEPS/TIME STEPS
    sec = (sec * 1000)
    sec = sec / framesPerSec
  
    For i = 1 To framesPerSec
        
        DoEvents
        
        'LOOP EACH CONTROL AND INCREMENT ACCORDINGLY
        For n = LBound(obj, 1) To UBound(obj, 1)
        
            'GET CURRENTVALUE
            CurrentValue = CallByName(obj(n), objProperty(n), VbGet)
            
            'INCREMENT PROPERTY VALUE
            CallByName obj(n), objProperty(n), VbLet, CurrentValue + increment(n) / framesPerSec
            
            Counter = Counter + (increment(n) / framesPerSec)
            
        Next n
        
        'API SLEEP FUNCTION (MILLISECONDS)
        Sleep sec
        
    Next i
    
End Function


Private Sub pForm_Click()
    checkFormat
End Sub


Private Sub pForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    BtnHover
    onButton = False
End Sub


Public Sub resetForm()
        
    Dim i As Integer
    
    For i = 1 To mzTxtGroup.Count
        mzTxtGroup(i)("txt").value = ""
    Next i
    
    For i = 1 To mzCmbGroup.Count
        'mzCmbGroup(i)("cmb").value = Null
        mzCmbGroup(i)("cmb").ListIndex = 0
        mzCmbGroup(i)("frm").Height = 1
        mzCmbGroup(i)("frm").BackColor = mzBorderColor
    Next i
    

End Sub


Public Sub toast(s As String, mil As Long, Form As MSForms.UserForm, bgColor As Long)

    Dim txt As MSForms.TextBox
    Dim l As Long
    Static running As Boolean
    
    If running = False Then
        running = True
    
        Set txt = Form.Controls.Add("Forms.Textbox.1", "toast" & s)
        
        DoEvents
        With txt
            
            .AutoSize = True
            .SpecialEffect = fmSpecialEffectFlat
            .value = s
            .MultiLine = True
            .BackColor = bgColor
            .ForeColor = mzWhite
            .Font.SIZE = 14
            .Left = (Form.InsideWidth / 1.1) - .Width
            .Top = (Form.InsideHeight / 6.5) + l
        End With
        
        animation txt, "top", -35, 7, 0.1
        
        Sleep mil
        
        animation txt, "top", -35, 7, 0.1
        Form.Controls.Remove txt.name
        
        running = False
   
    End If
    
End Sub

Private Function ArrayIncDim(Arr As Variant) As Variant
    
    If IsArrayEmpty(Arr) Then
        ReDim Arr(0)
    Else
        ReDim Preserve Arr(UBound(Arr, 1) + 1)
    End If
    
    ArrayIncDim = Arr
    
End Function

Private Function IsArrayEmpty(Arr As Variant) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' IsArrayEmpty
' This function tests whether the array is empty (unallocated). Returns TRUE or FALSE.
'
' The VBA IsArray function indicates whether a variable is an array, but it does not
' distinguish between allocated and unallocated arrays. It will return TRUE for both
' allocated and unallocated arrays. This function tests whether the array has actually
' been allocated.
'
' This function is really the reverse of IsArrayAllocated.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim LB As Long
    Dim UB As Long

    Err.Clear
    On Error Resume Next
    If IsArray(Arr) = False Then
        ' we weren't passed an array, return True
        IsArrayEmpty = True
    End If

    ' Attempt to get the UBound of the array. If the array is
    ' unallocated, an error will occur.
    UB = UBound(Arr, 1)
    If (Err.Number <> 0) Then
        IsArrayEmpty = True
    Else
        ''''''''''''''''''''''''''''''''''''''''''
        ' On rare occasion, under circumstances I
        ' cannot reliably replicate, Err.Number
        ' will be 0 for an unallocated, empty array.
        ' On these occasions, LBound is 0 and
        ' UBound is -1.
        ' To accommodate the weird behavior, test to
        ' see if LB > UB. If so, the array is not
        ' allocated.
        ''''''''''''''''''''''''''''''''''''''''''
        Err.Clear
        LB = LBound(Arr)
        If LB > UB Then
            IsArrayEmpty = True
        Else
            IsArrayEmpty = False
        End If
    End If

End Function

'==================================================
' CONVERT TO ARRAY
'==================================================
Private Function cArray(val As Variant) As Variant
    
    Dim i As Integer
    
    Select Case TypeName(val)
    
        Case "Collection":
            For i = 1 To val.Count
                ArrayPush cArray, val(i)
            Next i
        
        Case "Dictionary":
            cArray = val.Items()
        
        Case Else
             
            If IsArray(val) Then
                cArray = val
            Else
                ArrayPush cArray, val
            End If
            
    End Select
    
End Function

'==============================================================================================
' ADDS ELEMENT TO THE END OF A 1 DIM ARRAY - RETURNS NEW LENGHT OF ARRAY **[SINGLE DIMENSION]
'==============================================================================================
Private Function ArrayPush(ByRef Arr As Variant, value As Variant) As Long
    
    If IsArrayEmpty(Arr) Then
        ReDim Arr(0)
        
        If IsObject(value) Then
            Set Arr(0) = value
        Else
            Arr(0) = value
        End If
        
    Else
        ReDim Preserve Arr(LBound(Arr, 1) To UBound(Arr, 1) + 1)
        
        If IsObject(value) Then
            Set Arr(UBound(Arr, 1)) = value
        Else
            Arr(UBound(Arr, 1)) = value
        End If
    End If
    
    ArrayPush = UBound(Arr, 1)
    
End Function

Private Function regularExpressionExists(TestValue As String, Pattern As String) As Boolean
    
    Dim regEx As Object
    Dim match As Object
    
    'CREATE OBJECT
    Set regEx = CreateObject("vbscript.regexp")
    
    'SET PARAMETERS FOR REGULAR EXPRESSION
    With regEx
        .Global = True 'TRUE MEANS IT WILL LOOK FOR ALL MATCHES, FALSE FINDS FIRST ONLY
        .MultiLine = True
        .IgnoreCase = True
        .Pattern = Pattern 'SET THE PATTERN THAT WAS PASSED IN
    End With
    
    regularExpressionExists = regEx.test(TestValue)
    Set regEx = Nothing
    
End Function

