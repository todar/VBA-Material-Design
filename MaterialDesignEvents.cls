VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MaterialDesignEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Public WithEvents txt As MSForms.TextBox
Attribute txt.VB_VarHelpID = -1
Public WithEvents lbl As MSForms.Label
Attribute lbl.VB_VarHelpID = -1
Public WithEvents btn As MSForms.CommandButton
Attribute btn.VB_VarHelpID = -1
Public WithEvents cmb As MSForms.ComboBox
Attribute cmb.VB_VarHelpID = -1
Public WithEvents frm As MSForms.Frame
Attribute frm.VB_VarHelpID = -1


Public CallingClass As Object


Private Sub btn_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyTab Then
    End If
End Sub

Private Sub btn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
End Sub


Private Sub txt_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    If KeyCode <> vbKeyTab And KeyCode <> vbKeyReturn Then
        'RaiseEvent TextboxEntered(txt)
    End If
    
End Sub

Private Sub txt_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyTab Or KeyCode = vbKeyReturn Then
    End If
End Sub

Private Sub txt_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'RaiseEvent TextboxEntered(txt)
End Sub
Private Sub lbl_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
End Sub



