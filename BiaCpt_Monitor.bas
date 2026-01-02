Attribute VB_Name = "BiaCpt_Monitor"
Option Explicit


'Public paramIBM_BIA_INFO_Password As String
'Public paramIBM_BIA_AUTO_Password As String
'Public paramIBM_BIA_ODBC_Password As String
'Public paramIBM_BIA_DWH_Password As String
'Public paramIBM_BO_DWH_Password As String

'---------------------------------------------------------
Public Sub frmBiaLog_Show()
'---------------------------------------------------------
Dim X As String

frmBiaLog.Show vbModeless
frmBiaLog.WindowState = vbNormal
frmBiaLog.Visible = True
X = frmBiaLog.Caption
AppActivate X

End Sub


Public Sub mainSoc_Close()

tableCDComD_Close
tableCDDossier_Close
tableCDPosting_Close
tableCDTICom_Close
tableCDTIMaster_Close
tableCptP0_Close
tableDeviseChange_Close
tableDeviseCompta_Close
tableElpBuffer_Close
tableElpTable_Close
tableLrRetris_Close
tableLrRisque_Close
tableLrSgnBnf_Close
tableLrSort_Close
tableLrTiers_Close
tableMvtP0_Close
tableSwiftHisto_Close

End Sub



Public Sub mainSocExe()
Dim I As Integer

AccAutId = "SRVBIALR  "
frmElp_Caption = "BiaCpt"
frmElp_Icon = paramFolder_Local & "\misc37.ico"

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

'frmPaieSage.Show vbModeless
'frmPaieSage.WindowState = vbNormal
'frmPaieSage.Visible = True
'X = frmPaieSage.Caption
'AppActivate X

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

Public Sub frmTI2000_Show()
Dim X As String

frmTI2000.Show vbModeless
frmTI2000.WindowState = vbNormal
frmTI2000.Visible = True
X = frmTI2000.Caption
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


Public Sub frmCptEAR_Show()
Dim X As String

frmCptEAR.Show vbModeless
frmCptEAR.WindowState = vbNormal
frmCptEAR.Visible = True
X = frmCptEAR.Caption
AppActivate X

End Sub



Public Sub frmNovaBank_Show()
Dim X As String

frmNovaBank.Show vbModeless
frmNovaBank.WindowState = vbNormal
frmNovaBank.Visible = True
X = frmNovaBank.Caption
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
   'Case Is = "PAIESAGE": frmPaieSage_Show: frmPaieSage.Msg_Rcv Msg
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
   Case Is = "TI2000": frmTI2000_Show: frmTI2000.Msg_Rcv Msg
   Case Is = "FRMBIALOG", "BIA_LOG": frmBiaLog_Show:  frmBiaLog.Msg_Rcv Msg
   Case Is = "CPT_EAR", "CPTEAR": frmCptEAR_Show: frmCptEAR.Msg_Rcv Msg
   Case Is = "NOVABANK": frmNovaBank_Show: frmNovaBank.Msg_Rcv Msg
   Case Is = "COMPTE_GAFI": frmCompteGafi_Show: frmCompteGafi.Msg_Rcv Msg
   Case Is = "SAB_CPT_REP": frmSABCPTR_Show: frmSABCptR.Msg_Rcv Msg
   Case Is = "BIA_EXPLOIT2": BIA_EXPLOIT2
''   Case Is = "BIA_EXPLOIT": frmBiaExploitation_Show: frmBiaEXploitation.Msg_Rcv Msg
   
''   Case Is = "@BIA_EXPLOIT": frmBiaEXploitation.Msg_Rcv Msg
    Case Is = "@AUTO_NOVABK": frmNovaBank.Msg_Rcv Msg
    Case Is = "@AUTO_SWIFT": frmSwift.Msg_Rcv Msg
    Case Is = "@BAFI": srvLrBafi.PeliNT_Emission Msg
    Case Is = "TIMER": ElpTimer_Init
    Case Is = "X_RESET": main_Reset

End Select

End Sub

Public Sub frmCompteGafi_Show()
Dim X As String

frmCompteGafi.Show vbModeless
frmCompteGafi.WindowState = vbNormal
frmCompteGafi.Visible = True
X = frmCompteGafi.Caption
AppActivate X

End Sub

Public Sub frmSABCPTR_Show()
Dim X As String

frmSABCptR.Show vbModeless
frmSABCptR.WindowState = vbNormal
frmSABCptR.Visible = True
X = frmSABCptR.Caption
AppActivate X

End Sub
Public Sub frmElpTimer_Show()

End Sub

Public Sub BIA_EXPLOIT2()
frmBiaLog_Show
frmBiaLog.Msg_Rcv "BIA_LOG     " & "BIA_EXPLOIT"

frmCptEAR_Show
frmCptEAR.Msg_Rcv "CPT_EAR     " & "BIA_EXPLOIT"

End Sub
