Attribute VB_Name = "BiaGuichet_Monitor"
Option Explicit

'---------------------------------------------------------
Public Sub Msg_Monitor(Msg As String)
'---------------------------------------------------------
Select Case UCase$(Trim(mId$(Msg, 1, 12)))
   Case Is = "FRMCOMPTE", "COMPTE": frmCompte_Show: frmCompte.Msg_Rcv Msg
   Case Is = "FRMDEVCOUP", "G_COUPURES": frmDeviseCoupures_Show: frmDeviseCoupures.Msg_Rcv Msg
   Case Is = "FRMGUICHET", "GUICHET": frmGuichet_Show: frmGuichet.Msg_Rcv Msg
End Select

End Sub

Public Sub frmGuichet_Show()
Dim X As String
frmGuichet.Show vbModeless
frmGuichet.WindowState = vbNormal
frmGuichet.Visible = True
X = frmGuichet.Caption
AppActivate X

End Sub


'---------------------------------------------------------
Public Sub frmDeviseCoupures_Show()
'---------------------------------------------------------
Dim X As String

frmDeviseCoupures.Show vbModeless
frmDeviseCoupures.WindowState = vbNormal
frmDeviseCoupures.Visible = True
X = frmDeviseCoupures.Caption
AppActivate X

End Sub


'---------------------------------------------------------
Public Sub frmCompte_Show()
'---------------------------------------------------------
Dim X As String

frmCompte.Show vbModeless
frmCompte.WindowState = vbNormal
frmCompte.Visible = True
X = frmCompte.Caption
AppActivate X

End Sub



Public Sub mainSocExe()
frmElp_Caption = "BiaGuichet"
frmElp_Icon = "Misc35.ico"

AccAutId = "SRVUSRAPP "
arrCompteNbMax = 1
arrCompteNb = 0
ReDim arrCompte(1)

arrCptInfoNbMax = 1
arrCptInfoNb = 0
ReDim arrCptInfo(1)


End Sub

