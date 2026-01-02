Attribute VB_Name = "Bia_Monitor"
Option Explicit

Public Sub frmCompteSolde_Show()
Dim X As String

frmCompteSolde.Show vbModeless
frmCompteSolde.WindowState = vbNormal
frmCompteSolde.Visible = True
X = frmCompteSolde.Caption
AppActivate X

End Sub

Public Sub frmBiaExploitation_Show()
Dim X As String

frmBiaEXploitation.Show vbModeless
frmBiaEXploitation.WindowState = vbNormal
frmBiaEXploitation.Visible = True
X = frmBiaEXploitation.Caption
AppActivate X

End Sub


Public Sub frmGarantie_Show()
Dim X As String

frmGarantie.Show vbModeless
frmGarantie.WindowState = vbNormal
frmGarantie.Visible = True
X = frmGarantie.Caption
AppActivate X

End Sub

Public Sub frmTC_Show()
Dim X As String

frmTC.Show vbModeless
frmTC.WindowState = vbNormal
frmTC.Visible = True
X = frmTC.Caption
AppActivate X

End Sub

Public Sub frmEffetCommerce_Show()
Dim X As String

frmEffetCommerce.Show vbModeless
frmEffetCommerce.WindowState = vbNormal
frmEffetCommerce.Visible = True
X = frmEffetCommerce.Caption
AppActivate X

End Sub


Public Sub frmCompteE_Show()
Dim X As String

frmCompteE.Show vbModeless
frmCompteE.WindowState = vbNormal
frmCompteE.Visible = True
X = frmCompteE.Caption
AppActivate X

End Sub


'---------------------------------------------------------
Public Sub Msg_Monitor(Msg As String)
'---------------------------------------------------------
Select Case UCase$(Trim(mId$(Msg, 1, 12)))
   Case Is = "FRMCOMPTEE", "COMPTEE", "TEST": frmCompteE_Show: frmCompteE.Msg_Rcv Msg
   Case Is = "FRMANNUAIRE", "ANNUAIRE": frmAnnuaire_Show: frmAnnuaire.Msg_Rcv Msg
   Case Is = "FRMCV", "CONTRE-VAL": frmCV_Show:  frmCV.Msg_Rcv Msg
   Case Is = "FRMCOMPTE", "COMPTE": frmCompte_Show: frmCompte.Msg_Rcv Msg
   Case Is = "FRMOPTRF", "OPTRF": frmOpTrf_Show: frmOpTrf.Msg_Rcv Msg
   Case Is = "FRMOPTRFD", "OPTRFD": frmOptrfD_show: frmOpTrfD.Msg_Rcv Msg
   Case Is = "FRMBDF", "BDF": frmBdf_Show: frmBdf.Msg_Rcv Msg
   Case Is = "FRMBIC", "BIC": frmBic_Show: frmBic.Msg_Rcv Msg
   Case Is = "FRMBICLIST": frmBicList_Show: frmBicList.Msg_Rcv Msg
   Case Is = "FRMDICTIO", "DICTIO": frmDictio_Show: frmDictio.Msg_Rcv Msg
   Case Is = "FRMDEVCOUP", "G_COUPURES": frmDeviseCoupures_Show: frmDeviseCoupures.Msg_Rcv Msg
   Case Is = "FRMDEVISECHG", "T_CHANGE", "C_CHANGE": frmDeviseChange_show: frmDeviseChange.Msg_Rcv Msg
   Case Is = "FRMGUICHET", "GUICHET": frmGuichet_Show: frmGuichet.Msg_Rcv Msg
   Case Is = "C_DIVERS", "LUCAREPORT": frmBiaCpt_Show
   Case Is = "G_CONVENT.": frmBiaGuichetConvention_Show
   Case Is = "FRMCPTRELEVÉ", "CPTRELEVÉ": frmCptRelevé_show: frmCptRelevé.Msg_Rcv Msg
   Case Is = "FRMELPDOC", "DOCUMENT..": frmElpDoc_Show:  frmElpDoc.Msg_Rcv Msg
   Case Is = "FRMELPTABLE", "TABLE": frmElpTable_Show:  frmElpTable.Msg_Rcv Msg
   Case Is = "FRMBOTC", "BOTC": frmBOTC_Show:  frmBOTC.Msg_Rcv Msg
   Case Is = "FRMDAFI", "DAFI": frmDAFI_Show:  frmDAFI.Msg_Rcv Msg
   Case Is = "FRMPRÊTPP", "PRÊTSPP": frmPrêt_Show:  frmPrêt.Msg_Rcv Msg
   Case Is = "FRMPRÊTPC", "PRÊTSPC": frmPrêt_Show:  frmPrêt.Msg_Rcv Msg
   Case Is = "COMPTE_SLD+", "COMPTE_SLD": frmCompteSolde_Show: frmCompteSolde.Msg_Rcv Msg
   Case Is = "COMPTE_SLD$": frmCompteSolde_Show: frmCompteSolde.Msg_Rcv Msg
   Case Is = "BIA_EXPLOIT": frmBiaExploitation_Show: frmBiaEXploitation.Msg_Rcv Msg
   Case Is = "@BIA_EXPLOIT": frmBiaEXploitation.Msg_Rcv Msg
   Case Is = "DAFI_GARANT": frmGarantie_Show: frmGarantie.Msg_Rcv Msg
   Case Is = "SOBF_EFFETS": frmEffetCommerce_Show: frmEffetCommerce.Msg_Rcv Msg
   Case Is = "FRMBIALOG", "BIALOG": frmBiaLog_Show:  frmBiaLog.Msg_Rcv Msg
   Case Is = "TEST", "TC", "FRMTC": frmTC_Show: frmTC.Msg_Rcv Msg
   Case Is = "XUSRID": XUsrId_Show
   Case Is = "X_RESET": main_Reset
 '2003.12.15  Case Is = "X_BIA.MDB": MDB_Copy
End Select

End Sub

Public Sub frmBicList_Show()
'---------------------------------------------------------
Dim X As String

frmBicList.Show vbModeless
frmBicList.WindowState = vbNormal
frmBicList.Visible = True
X = frmBicList.Caption
AppActivate X

End Sub



Public Sub frmCptRelevé_show()
Dim X As String
frmCptRelevé.Show vbModeless
frmCptRelevé.WindowState = vbNormal
frmCptRelevé.Visible = True
X = frmCptRelevé.Caption
AppActivate X
End Sub



Public Sub frmDeviseChange_show()
Dim X As String
frmDeviseChange.Show vbModeless
frmDeviseChange.WindowState = vbNormal
frmDeviseChange.Visible = True
X = frmDeviseChange.Caption
AppActivate X
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
Public Sub frmBdf_Show()
'---------------------------------------------------------
Dim X As String

frmBdf.Show vbModeless
frmBdf.WindowState = vbNormal
frmBdf.Visible = True
X = frmBdf.Caption
AppActivate X
End Sub



Public Sub frmOptrfD_show()
Dim X As String

frmOpTrfD.Show vbModeless
frmOpTrfD.WindowState = vbNormal
frmOpTrfD.Visible = True
X = frmOpTrfD.Caption
AppActivate X

End Sub



'---------------------------------------------------------
Public Sub frmOpTrf_Show()
'---------------------------------------------------------
Dim X As String

frmOpTrf.Show vbModeless
frmOpTrf.WindowState = vbNormal
frmOpTrf.Visible = True
X = frmOpTrf.Caption
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
Public Sub frmAnnuaire_Show()
'---------------------------------------------------------
Dim X As String

frmAnnuaire.Show vbModeless
frmAnnuaire.WindowState = vbNormal
frmAnnuaire.Visible = True
X = frmAnnuaire.Caption
AppActivate X

End Sub


'---------------------------------------------------------
Public Sub frmElpDoc_Show()
'---------------------------------------------------------
Dim X As String

frmElpDoc.Show vbModeless
frmElpDoc.WindowState = vbNormal
frmElpDoc.Visible = True
X = frmElpDoc.Caption
AppActivate X

End Sub

'---------------------------------------------------------
Public Sub frmElpTable_Show()
'---------------------------------------------------------
Dim X As String

frmElpTable.Show vbModeless
frmElpTable.WindowState = vbNormal
frmElpTable.Visible = True
X = frmElpTable.Caption
AppActivate X

End Sub


'---------------------------------------------------------
Public Sub frmBOTC_Show()
'---------------------------------------------------------
Dim X As String

frmBOTC.Show vbModeless
frmBOTC.WindowState = vbNormal
frmBOTC.Visible = True
X = frmBOTC.Caption
AppActivate X

End Sub

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


'---------------------------------------------------------
Public Sub frmDAFI_Show()
'---------------------------------------------------------
Dim X As String

frmDAFI.Show vbModeless
frmDAFI.WindowState = vbNormal
frmDAFI.Visible = True
X = frmDAFI.Caption
AppActivate X

End Sub


'---------------------------------------------------------
Public Sub frmPrêt_Show()
'---------------------------------------------------------
Dim X As String

frmPrêt.Show vbModeless
frmPrêt.WindowState = vbNormal
frmPrêt.Visible = True
X = frmPrêt.Caption
AppActivate X

End Sub

'---------------------------------------------------------
Public Sub frmCV_Show()
'---------------------------------------------------------
Dim X As String

frmCV.Show vbModeless
frmCV.WindowState = vbNormal
frmCV.Visible = True
X = frmCV.Caption
AppActivate X

End Sub



'---------------------------------------------------------
Public Sub frmBic_Show()
'---------------------------------------------------------
Dim X As String

frmBic.Show vbModeless
frmBic.WindowState = vbNormal
frmBic.Visible = True
X = frmBic.Caption
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

'usrDRH usrIdNT

AccAutId = "SRVUSRAPP "
frmElp_Icon = paramFolder_Local & "\misc34.ico"

arrCompteNbMax = 1
arrCompteNb = 0
ReDim arrCompte(1)

arrCptInfoNbMax = 1
arrCptInfoNb = 0
ReDim arrCptInfo(1)

arrBicNbMax = 1
arrBicNb = 0
ReDim arrBic(1)

arrBicIbanNbMax = 1
arrBicIbanNb = 0
ReDim arrBicIban(1)

arrBdfENbMax = 1
arrBdfENb = 0
ReDim arrBdfE(1)

arrBdfGNbMax = 1
arrBdfGNb = 0
ReDim arrBdfG(1)

arrRacineNbMax = 1
arrRacineNb = 0
ReDim arrRacine(1)

End Sub

Public Sub frmBiaCpt_Show()
Dim IdShell
IdShell = Shell(SrvDir & "BiaCpt.exe " & Command, 1)
AppActivate IdShell

End Sub
Public Sub frmBiaGuichetConvention_Show()
Dim IdShell
IdShell = Shell("c:\BiaSrv\BiaGuichetConvention.exe " & SrvDir, 1)
AppActivate IdShell

End Sub



Public Sub mainSoc_Close()

tableCptP0_Close
tableDeviseChange_Close
tableDeviseCompta_Close
tableElpBuffer_Close
tableElpDoc_Close
tableElpTable_Close
tableGuichet_Compta_Close
tableMvtP0_Close

End Sub
