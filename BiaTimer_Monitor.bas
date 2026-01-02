Attribute VB_Name = "Bia_Monitor"
Option Explicit

'---------------------------------------------------------
Public Sub frmElpChat_Show()
'---------------------------------------------------------
Dim X As String

frmElpChat.Show vbModeless
frmElpChat.WindowState = vbNormal
frmElpChat.Visible = True
X = frmElpChat.Caption
AppActivate X
End Sub


Public Sub mainSocExe()
Dim I As Integer

AccAutId = "SRVBIALR  "
frmElp_Caption = "BiaTimer"
''frmElp_Icon = "misc37.ico"

arrCompteNbMax = 1
arrCompteNb = 0
ReDim arrCompte(1)

arrCptInfoNbMax = 1
arrCptInfoNb = 0
ReDim arrCptInfo(1)

arrRacineNbMax = 1
arrRacineNb = 0
ReDim arrRacine(1)


End Sub

'---------------------------------------------------------
Public Sub frmDictio_Show()
'---------------------------------------------------------
Dim X As String

frmDictio.Show vbModeless
frmDictio.WindowState = vbNormal
frmDictio.Visible = True
X = frmDictio.Caption
AppActivate X

End Sub


'---------------------------------------------------------
Public Sub frmInformatique_Show()
'---------------------------------------------------------
Dim X As String

frmInformatique.Show vbModeless
frmInformatique.WindowState = vbNormal
frmInformatique.Visible = True
X = frmInformatique.Caption
AppActivate X

End Sub



Public Sub frmLrAttribut_Show()
Dim X As String

frmLrAttribut.Show vbModeless
frmLrAttribut.WindowState = vbNormal
frmLrAttribut.Visible = True
X = frmLrAttribut.Caption
AppActivate X

End Sub

Public Sub frmLrTiers_Show()

Dim X As String

frmLrTiers.Show vbModeless
frmLrTiers.WindowState = vbNormal
frmLrTiers.Visible = True
X = frmLrTiers.Caption
AppActivate X

End Sub

Public Sub frmLrBafi_Show()
Dim X As String

frmLrBafi.Show vbModeless
frmLrBafi.WindowState = vbNormal
frmLrBafi.Visible = True
X = frmLrBafi.Caption
AppActivate X

End Sub

Public Sub frmEchellesFusion_Show()
Dim X As String

frmEchellesFusion.Show vbModeless
frmEchellesFusion.WindowState = vbNormal
frmEchellesFusion.Visible = True
X = frmEchellesFusion.Caption
AppActivate X

End Sub


Public Sub frmCompteModif_Show()
Dim X As String

frmCompteModif.Show vbModeless
frmCompteModif.WindowState = vbNormal
frmCompteModif.Visible = True
X = frmCompteModif.Caption
AppActivate X

End Sub

Public Sub frmBiaPgmAut_Show()
Dim X As String

frmBiaPgmAut.Show vbModeless
frmBiaPgmAut.WindowState = vbNormal
frmBiaPgmAut.Visible = True
X = frmBiaPgmAut.Caption
AppActivate X

End Sub


Public Sub frmBiaPgm_Show()
Dim X As String

frmBiaPgm.Show vbModeless
frmBiaPgm.WindowState = vbNormal
frmBiaPgm.Visible = True
X = frmBiaPgm.Caption
AppActivate X

End Sub

Public Sub frmDGI_2561_Show()
Dim X As String

frmDGI_2561.Show vbModeless
frmDGI_2561.WindowState = vbNormal
frmDGI_2561.Visible = True
X = frmDGI_2561.Caption
AppActivate X

End Sub



Public Sub frmCompteExtrait_Show()
Dim X As String

frmCompteExtrait.Show vbModeless
frmCompteExtrait.WindowState = vbNormal
frmCompteExtrait.Visible = True
X = frmCompteExtrait.Caption
AppActivate X

End Sub

Public Sub frmCompteCapMoy_Show()
Dim X As String

frmCompteCapMoy.Show vbModeless
frmCompteCapMoy.WindowState = vbNormal
frmCompteCapMoy.Visible = True
X = frmCompteCapMoy.Caption
AppActivate X

End Sub


Public Sub frmLucaRisques_Show()
Dim X As String

frmLucaRisques.Show vbModeless
frmLucaRisques.WindowState = vbNormal
frmLucaRisques.Visible = True
X = frmLucaRisques.Caption
AppActivate X

End Sub

Public Sub frmPaieSage_Show()
Dim X As String

frmPaieSage.Show vbModeless
frmPaieSage.WindowState = vbNormal
frmPaieSage.Visible = True
X = frmPaieSage.Caption
AppActivate X

End Sub

Public Sub frmCptComPays_Show()
Dim X As String

frmCptComPays.Show vbModeless
frmCptComPays.WindowState = vbNormal
frmCptComPays.Visible = True
X = frmCptComPays.Caption
AppActivate X

End Sub


Public Sub frmGAdresse_Show()
Dim X As String

frmGAdresse.Show vbModeless
frmGAdresse.WindowState = vbNormal
frmGAdresse.Visible = True
X = frmGAdresse.Caption
AppActivate X

End Sub


Public Sub frmDRH_Show()
Dim X As String

frmDRH.Show vbModeless
frmDRH.WindowState = vbNormal
frmDRH.Visible = True
X = frmDRH.Caption
AppActivate X

End Sub

Public Sub frmTI_Show()
Dim X As String

frmTI.Show vbModeless
frmTI.WindowState = vbNormal
frmTI.Visible = True
X = frmTI.Caption
AppActivate X

End Sub

Public Sub frmSwift_Show()
Dim X As String

frmSwift.Show vbModeless
frmSwift.WindowState = vbNormal
frmSwift.Visible = True
X = frmSwift.Caption
AppActivate X

End Sub

Public Sub frmSwift_ShowModal(Msg As String)
Dim X As String
On Error Resume Next
Load frmSwift
frmSwift.Msg_Rcv Msg
frmSwift.Show vbModal
frmSwift.WindowState = vbMinimized
frmSwift.Visible = True
X = frmSwift.Caption
AppActivate X

End Sub


'---------------------------------------------------------
Public Sub Msg_Monitor(Msg As String)
'---------------------------------------------------------
Select Case UCase$(Trim(mId$(Msg, 1, 12)))
   Case Is = "DRH": frmDRH_Show: frmDRH.Msg_Rcv Msg
   Case Is = "FRMDICTIO", "DICTIO": frmDictio_Show: frmDictio.Msg_Rcv Msg
   Case Is = "FRMLRBAFI", "LRBAFI": frmLrBafi_Show: frmLrBafi.Msg_Rcv Msg
   Case Is = "FRMINFORMATI", "INFORMATIQ": frmInformatique_Show: frmInformatique.Msg_Rcv Msg
   Case Is = "FRMLRATTRIBU", "LRATTRIBUT": frmLrAttribut_Show: frmLrAttribut.Msg_Rcv Msg
   Case Is = "FRMLRTIERS": frmLrTiers_Show: frmLrTiers.Msg_Rcv Msg
   Case Is = "FRMECH_FUSIO", "ECH_FUSION": frmEchellesFusion_Show: frmEchellesFusion.Msg_Rcv Msg
   Case Is = "LUCARISQUE": frmLucaRisques_Show: frmLucaRisques.Msg_Rcv Msg
   Case Is = "PAIESAGE": frmPaieSage_Show: frmPaieSage.Msg_Rcv Msg
   Case Is = "COMPTE_MOD": frmCompteModif_Show: frmCompteModif.Msg_Rcv Msg
   Case Is = "BIAPGM_AUT": frmBiaPgmAut_Show: frmBiaPgmAut.Msg_Rcv Msg
   Case Is = "BIAPGM": frmBiaPgm_Show: frmBiaPgm.Msg_Rcv Msg
   Case Is = "DGI_2561": frmDGI_2561_Show: frmDGI_2561.Msg_Rcv Msg
   Case Is = "COMPTE_EXT": frmCompteExtrait_Show: frmCompteExtrait.Msg_Rcv Msg
   Case Is = "COMPTE_CAPMO": frmCompteCapMoy_Show: frmCompteCapMoy.Msg_Rcv Msg
   Case Is = "CPT_COMPAYS": frmCptComPays_Show: frmCptComPays.Msg_Rcv Msg
   Case Is = "CPT_ADRESSE": frmGAdresse_Show: frmGAdresse.Msg_Rcv Msg
   Case Is = "XUSRID_BIACP": XUsrId_Show
   Case Is = "SWIFT", "$AUTO_SWIFT": frmSwift_Show: frmSwift.Msg_Rcv Msg
   Case Is = "TI": frmTI_Show: frmTI.Msg_Rcv Msg
   Case Is = "@AUTO_SWIFT": frmSwift.Msg_Rcv Msg
   Case Is = "TIMER": ElpTimer_Init
End Select

End Sub



Public Sub frmElpTimer_Show()

End Sub
