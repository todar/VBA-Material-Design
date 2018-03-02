VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "materializeEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Public mz As materialize
Attribute mz.VB_VarHelpID = -1

Public WithEvents txt As MSForms.TextBox
Attribute txt.VB_VarHelpID = -1
Public WithEvents lbl As MSForms.Label
Attribute lbl.VB_VarHelpID = -1
Public WithEvents btn As MSForms.CommandButton
Attribute btn.VB_VarHelpID = -1


Private Sub btn_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyTab Then
        mz.setFocus ""
    End If
End Sub

Private Sub btn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    mz.onButton = True
    mz.BtnHover btn.Name
End Sub

Private Sub lbl_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
      On Error GoTo catch
    txt.setFocus
    mz.setFocus txt.Name
catch:
End Sub

Private Sub txt_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    If KeyCode <> vbKeyTab And KeyCode <> vbKeyReturn Then
        txt.setFocus
        mz.setFocus txt.Name
    End If
    
End Sub

Private Sub txt_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyTab Or KeyCode = vbKeyReturn Then
        If mz.liveValidation = True Then mz.Validate
        txt.setFocus
        mz.setFocus txt.Name
    End If
End Sub

Private Sub txt_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    mz.setFocus txt.Name
End Sub
Private Sub lbl_Click()
  
End Sub
