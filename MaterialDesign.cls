VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MaterialDesign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

'ARRAY OF EVENTS CLASS MODULE.
Private Events() As New MaterialDesignEvents

Const padding As Double = 3


'******************************************************************************************
' PUBLIC METHODS/FUNCTIONS
'******************************************************************************************
Public Sub AutoInit(form As Object)
    CenterForm form
End Sub

Public Sub StyleTextbox(txt As MSForms.TextBox, Placeholder As String)
    
    With txt
        .value = ""
        .BackStyle = fmBackStyleTransparent
        .BorderStyle = fmBorderStyleSingle
        .BorderStyle = fmBorderStyleNone
        .SelectionMargin = False
    End With
    
    'ADD CONTROLS FOR MATERIAL LOOK
    Dim lbl As MSForms.Label
    Dim Border As MSForms.Label
    Dim Notification As MSForms.Label
    Dim Background As MSForms.Label
    
    Set Border = AddBorder(txt)
    Set lbl = AddPlaceholder(txt, Placeholder)
    Set Notification = AddNotificationLabel(txt)
    Set Background = AddBackground(lbl, (Border.Top + Border.Height) - lbl.Top)
    
    Dim EventIndex As Integer
    On Error GoTo Catch
    EventIndex = UBound(Events(), 1) + 1
    
    Set Events(EventIndex).CallingClass = Me
    Set Events(EventIndex).txt = txt
    Set Events(EventIndex).lbl = lbl
    
    Exit Sub
Catch:
    If Err.Number = 9 Then
        ReDim Events(0)
        Resume Next
    End If
    Stop

End Sub

Public Sub StyleButton()
End Sub

Public Sub SetFocusOnControl()
End Sub


'******************************************************************************************
' PRIVATE METHODS/FUNCTIONS
'******************************************************************************************

Private Sub ApplyGeneralSettingsToControl(ByRef Control As MSForms.Control, Source As MSForms.Control, Optional Caption As String)

    With Control
        .Height = Source.Height
        .Width = Source.Width
        .Left = Source.Left
        .Top = Source.Top
        .BorderStyle = fmBorderStyleSingle
        .BorderStyle = fmBorderStyleNone
        .BackStyle = fmBackStyleTransparent
        .MousePointer = fmMousePointerIBeam
        .Caption = Caption
        .ForeColor = Source.ForeColor '8951296
        .Font.Size = Source.Font.Size
    End With
    
End Sub

Private Function AddBorder(Control As MSForms.Control) As MSForms.Label
    
    Dim Border As MSForms.Label
    Set Border = Control.Parent.Controls.Add("Forms.Label.1", "Border" & Control.name)
    
    ApplyGeneralSettingsToControl Border, Control
    With Border
        .Height = 2
        '.Width = 0
        .Width = Control.Width + padding + padding
        .Left = .Left - padding
        .Top = .Top + Control.Height
        .BackColor = 8951296
        .BackStyle = fmBackStyleOpaque
    End With
    
    Set AddBorder = Border
    
End Function

Private Function AddPlaceholder(Control As MSForms.Control, Caption As String) As MSForms.Label
    
    Dim Placeholder As MSForms.Label
    Set Placeholder = Control.Parent.Controls.Add("Forms.Label.1", "Placeholder" & Control.name)
    
    ApplyGeneralSettingsToControl Placeholder, Control, Caption
    With Placeholder
        .Height = .Height - 5
        .Top = .Top - .Height
        '.Font.Size = .Font.Size - 2
    End With
    
    Set AddPlaceholder = Placeholder
    
End Function

Private Function AddNotificationLabel(Control As MSForms.Control) As MSForms.Label
    
    Dim Notification As MSForms.Label
    Set Notification = Control.Parent.Controls.Add("Forms.Label.1", "Notification" & Control.name)
    
    ApplyGeneralSettingsToControl Notification, Control
    
    With Notification
        .Top = .Top + Control.Height + 3
        .Caption = "Test"
        .ForeColor = 6994790
        .MousePointer = fmMousePointerDefault
    End With

    Set AddNotificationLabel = Notification
    
End Function

Private Function AddBackground(Control As MSForms.Control, Height As Double) As MSForms.Label

    Dim bg As MSForms.Label
    Set bg = Control.Parent.Controls.Add("Forms.Label.1", "bg" & Control.name)
    
    ApplyGeneralSettingsToControl bg, Control
    With bg
        .Height = Height + padding
        .Width = .Width + padding + padding
        .Left = .Left - padding
        .Top = .Top - padding
        .BackColor = 15263976
        .BackStyle = fmBackStyleOpaque
        .ZOrder (1)
    End With
    
    Set AddBackground = bg

End Function


Private Sub CenterForm(form As Object)
    
    With form
        .StartUpPosition = 0
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With

End Sub


