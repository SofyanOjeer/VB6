Attribute VB_Name = "Bia_JRN_Monitor"
Option Explicit


Public Sub mainSocExe()
frmElp_Caption = "BIA_JRN"
Set frmElp_Icon = frmJRN_SAB
blnMonitor = True

End Sub


'---------------------------------------------------------
Public Sub Msg_Monitor(Msg As String)
'---------------------------------------------------------
If Not blnMonitor Then Exit Sub

Select Case UCase$(Trim(Mid$(Msg, 1, 12)))
    Case Is = "JRN_SAB": frmJRN_SAB_Show: frmJRN_SAB.Msg_Rcv Msg:
    Case Is = "X_RESET":  main_Reset
    Case Is = "XUSRID": XUsrId_Show
    Case Is = "X_I5A7": X_I5A7_Show
End Select

End Sub
Public Sub frmJRN_SAB_Show()
Dim X As String

frmJRN_SAB.Icon = frmElp_Icon
frmJRN_SAB.Show vbModeless
frmJRN_SAB.WindowState = vbNormal
frmJRN_SAB.Visible = True
X = frmJRN_SAB.Caption
AppActivate X

End Sub



