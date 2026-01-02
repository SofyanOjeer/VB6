Attribute VB_Name = "BiaS820I_Monitor"
Option Explicit


Type typeBiaUsr                                  ' compatibilité   BIA.vbp
    Obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    ID                     As String * 10
    Nom                    As String * 34
    Service                As String * 3
    Coges                  As String * 2
    Groupe                 As String * 10

End Type

Public Sub lstBiaUsr_Load(lstX As ListBox, recBiaUsr As typeBiaUsr)
Dim xYbase As typeYBase

lstX.Clear
recYBase_Init xYbase
xYbase.Method = "Seek>="
xYbase.ID = constYBIATAB0
xYbase.K1 = "USER"
Do
    intReturn = tableYBase_Read(xYbase)
    If intReturn = 0 Then
        If Trim(mId$(xYbase.K1, 1, 24)) <> "USER" Then
            intReturn = -1
        Else
 
'            MsgTxt = Space$(34) & mId$(xYBase.Memo, 52, 37)
'            MsgTxtIndex = 0
'            srvYMNUUTI0_GetBuffer meYMNUUTI0

'            If meYMNUUTI0.MNUUTICGR = 0 Then
                lstX.AddItem mId$(xYbase.Text, 25, 42)
 '           End If
                xYbase.Method = "MoveNext"
        End If
    End If
    
Loop Until intReturn <> 0

'lstX.ListIndex = 0


End Sub

Public Sub mainSoc_Close()

tableElpTable_Close

tableElpKMInfo_Close
tableElpKMIndex_Close
tableElpKMLink_Close

End Sub



Public Sub mainSocExe()
frmElp_Caption = "BIAS820I"
frmElp_Icon = paramFolder_Local & "\BIA2.ico"
blnMonitor = True
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

'---------------------------------------------------------
Public Sub Msg_Monitor(Msg As String)
'---------------------------------------------------------
If Not blnMonitor Then Exit Sub
Select Case UCase$(Trim(mId$(Msg, 1, 12)))
    Case Is = "BIA_GAFI": frmBIA_GAFI_Show: frmBIA_Gafi.Msg_Rcv Msg:
    Case Is = "EDITION", "EDITION$": frmEdition_Show: frmEdition.Msg_Rcv Msg:
    Case Is = "EDITION_GEST": frmEdition_Gestion_Show: frmEdition_Gestion.Msg_Rcv Msg:
    Case Is = "CHQ_SCAN": frmCHQ_SCAN_Show: frmCHQ_SCAN.Msg_Rcv Msg:
    Case Is = "SAB_BALANCE": frmSAB_Balance_Show: frmSAB_Balance.Msg_Rcv Msg:
    Case Is = "SAB_STOCK": frmSAB_Stock_Show: frmSAB_Stock.Msg_Rcv Msg:
    Case Is = "SAB_CDO": frmSAB_CDO_Show: frmSAB_CDO.Msg_Rcv Msg:
    Case Is = "SAB_CRE": frmSAB_CRE_Show: frmSAB_CRE.Msg_Rcv Msg:
    Case Is = "SAB_TC": frmSAB_TC_Show: frmSAB_TC.Msg_Rcv Msg:
    Case Is = "SAB_TC_LIMIT": frmSAB_TC_Limites.Show: frmSAB_TC_Limites.Msg_Rcv Msg:
     Case Is = "SAB_CLI": frmSAB_CLI_Show: frmSAB_CLI.Msg_Rcv Msg:
     Case Is = "SAB_CLIENT": frmSAB_CLIENT_Show: frmSAB_Client.Msg_Rcv Msg:
  Case Is = "SAB_COMPTA": frmSAB_Compta_Show: frmSAB_Compta.Msg_Rcv Msg:
    Case Is = "SAB_DWH": frmSAB_DWH_Show: frmSAB_DWH.Msg_Rcv Msg:
    'Case Is = "BIA_FTP": frmBIA_FTP_Show: frmBIA_FTP.Msg_Rcv Msg:
    Case Is = "SAB_MNU": frmSAB_MNU_Show: frmSAB_MNU.Msg_Rcv Msg:
    Case Is = "SAB_ORDONNAN": frmSAB_Ordonnanceur_Show: frmSAB_Ordonnanceur.Msg_Rcv Msg:
    Case Is = "SAB_TAU": frmSAB_TAU_Show: frmSAB_TAU.Msg_Rcv Msg:
    Case Is = "@AUTO_TAU": frmSAB_TAU.Msg_Rcv Msg:
    Case Is = "FRMELPTABLE", "TABLE": frmElpTable_Show: frmElpTable.Msg_Rcv Msg
    Case Is = "SAA": frmSAA_Show: frmSAA.Msg_Rcv Msg:
    Case Is = "SPLFJOB": frmSPLFJOB_Show: frmSPLFJOB.Msg_Rcv Msg:
    Case Is = "TIMER": ElpTimer_Init
    Case Is = "X_DOC", "X_DOC$": frmElpKM_Show: frmElpKM.Msg_Rcv Msg
    Case Is = "BIAPGM": frmBiaPgm_Show: frmBiaPgm.Msg_Rcv Msg
    Case Is = "BIAPGM_AUT": frmBiaPgmAut_Show: frmBiaPgmAut.Msg_Rcv Msg
    Case Is = "X_DOC_SRC$": frmElpKMPgm_Show: frmElpKMPgm.Msg_Rcv Msg
    Case Is = "PAIESAGE": frmPaieSAge_Show: frmPaieSage.Msg_Rcv Msg:
    Case Is = "X_RESET":  main_Reset
    Case Is = "XUSRID": XUsrId_Show
    Case Is = "@AUTO_SPLF", "@AUTO_CLROUT": frmSPLFJOB.Msg_Rcv Msg:
    Case Is = "@PRINT_TEST", "@PRINT_PROD": frmEdition.Msg_Rcv Msg:
    Case Is = "AUTOMATE": frmAutomate_Show: frmAutomate.Msg_Rcv Msg:
    Case Is = "@AUTOMATE":  frmAutomate.Msg_Rcv Msg:
    
    Case Is = "@CHQ_DEON": mainSoc_AMJCPT_Load
                         If blnAuto_Exploitation_Ok("DATE_CPT_J", "@CHQ_DEON") Then
                            If blnAuto_Form_Show Then frmCHQ_SCAN_Show
                            frmCHQ_SCAN.Msg_Rcv Msg
                            Call blnAuto_Exploitation_Ok("Update", "@CHQ_DEON")
                        End If

    Case Is = "@AUTO_COMPTA": mainSoc_AMJCPT_Load
                         If blnAuto_Exploitation_Ok("DATE_CPT_J", "@AUTO_COMPTA") Then
                            If blnAuto_Form_Show Then frmSAB_Balance_Show
                            frmSAB_Balance.Msg_Rcv Msg
                            Call blnAuto_Exploitation_Ok("Update", "@AUTO_COMPTA")
                        End If
End Select

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


Public Sub frmElpKMPgm_Show()
Dim X As String

frmElpKMPgm.Show vbModeless
frmElpKMPgm.WindowState = vbNormal
frmElpKMPgm.Visible = True
X = frmElpKMPgm.Caption
AppActivate X

End Sub


Public Sub frmElpKM_Show()
Dim X As String

frmElpKM.Show vbModeless
frmElpKM.WindowState = vbNormal
frmElpKM.Visible = True
X = frmElpKM.Caption
AppActivate X

End Sub


Public Sub frmRTF_Show()
Dim X As String

frmRTF.Show vbModeless
frmRTF.WindowState = vbNormal
frmRTF.Visible = True
X = frmRTF.Caption
AppActivate X

End Sub

Public Sub frmSAA_Show()
Dim X As String

frmSAA.Show vbModeless
frmSAA.WindowState = vbNormal
frmSAA.Visible = True
X = frmSAA.Caption
AppActivate X

End Sub


Public Sub frmEdition_Show()
Dim X As String

frmEdition.Show vbModeless
frmEdition.WindowState = vbNormal
frmEdition.Visible = True
X = frmEdition.Caption
AppActivate X

End Sub

Public Sub frmEdition_Gestion_Show()
Dim X As String

frmEdition_Gestion.Show vbModeless
frmEdition_Gestion.WindowState = vbNormal
frmEdition_Gestion.Visible = True
X = frmEdition_Gestion.Caption
AppActivate X

End Sub


Public Sub frmSAB_MNU_Show()
Dim X As String

frmSAB_MNU.Show vbModeless
frmSAB_MNU.WindowState = vbNormal
frmSAB_MNU.Visible = True
X = frmSAB_MNU.Caption
AppActivate X

End Sub

Public Sub frmCHQ_SCAN_Show()
Dim X As String

frmCHQ_SCAN.Show vbModeless
frmCHQ_SCAN.WindowState = vbNormal
frmCHQ_SCAN.Visible = True
X = frmCHQ_SCAN.Caption
AppActivate X

End Sub


Public Sub frmSAB_CDO_Show()
Dim X As String

frmSAB_CDO.Show vbModeless
frmSAB_CDO.WindowState = vbNormal
frmSAB_CDO.Visible = True
X = frmSAB_CDO.Caption
AppActivate X

End Sub
Public Sub frmSAB_CRE_Show()
Dim X As String

frmSAB_CRE.Show vbModeless
frmSAB_CRE.WindowState = vbNormal
frmSAB_CRE.Visible = True
X = frmSAB_CRE.Caption
AppActivate X

End Sub

Public Sub frmSAB_TC_Show()
Dim X As String

frmSAB_TC.Show vbModeless
frmSAB_TC.WindowState = vbNormal
frmSAB_TC.Visible = True
X = frmSAB_TC.Caption
AppActivate X

End Sub

Public Sub frmSAB_TC_Limites_Show()
Dim X As String

frmSAB_TC_Limites.Show vbModeless
frmSAB_TC_Limites.WindowState = vbNormal
frmSAB_TC_Limites.Visible = True
X = frmSAB_TC_Limites.Caption
AppActivate X

End Sub


Public Sub frmSAB_DWH_Show()
Dim X As String

frmSAB_DWH.Show vbModeless
frmSAB_DWH.WindowState = vbNormal
frmSAB_DWH.Visible = True
X = frmSAB_DWH.Caption
AppActivate X

End Sub

Public Sub frmSAB_Balance_Show()
Dim X As String

frmSAB_Balance.Show vbModeless
frmSAB_Balance.WindowState = vbNormal
frmSAB_Balance.Visible = True
X = frmSAB_Balance.Caption
AppActivate X

End Sub

Public Sub frmSAB_Stock_Show()
Dim X As String

frmSAB_Stock.Show vbModeless
frmSAB_Stock.WindowState = vbNormal
frmSAB_Stock.Visible = True
X = frmSAB_Stock.Caption
AppActivate X

End Sub

Public Sub frmSAB_CLI_Show()
Dim X As String

frmSAB_CLI.Show vbModeless
frmSAB_CLI.WindowState = vbNormal
frmSAB_CLI.Visible = True
X = frmSAB_CLI.Caption
AppActivate X

End Sub

Public Sub frmSAB_CLIENT_Show()
Dim X As String

frmSAB_Client.Show vbModeless
frmSAB_Client.WindowState = vbNormal
frmSAB_Client.Visible = True
X = frmSAB_Client.Caption
AppActivate X

End Sub

Public Sub frmSAB_TAU_Show()
Dim X As String

frmSAB_TAU.Show vbModeless
frmSAB_TAU.WindowState = vbNormal
frmSAB_TAU.Visible = True
X = frmSAB_TAU.Caption
AppActivate X

End Sub
Public Sub frmAutomate_Show()
Dim X As String

frmAutomate.Show vbModeless
frmAutomate.WindowState = vbNormal
frmAutomate.Visible = True
X = frmAutomate.Caption
AppActivate X

End Sub

Public Sub frmBIA_GAFI_Show()
Dim X As String

frmBIA_Gafi.Show vbModeless
frmBIA_Gafi.WindowState = vbNormal
frmBIA_Gafi.Visible = True
X = frmBIA_Gafi.Caption
AppActivate X

End Sub

Public Sub frmSAB_Compta_Show()
Dim X As String
On Error Resume Next
frmSAB_Compta.Show vbModeless
frmSAB_Compta.WindowState = vbNormal
frmSAB_Compta.Visible = True
X = frmSAB_Compta.Caption
AppActivate X

End Sub

Public Sub frmSPLFJOB_Show()
Dim X As String

frmSPLFJOB.Show vbModeless
frmSPLFJOB.WindowState = vbNormal
frmSPLFJOB.Visible = True
X = frmSPLFJOB.Caption
AppActivate X

End Sub

Public Sub frmPaieSAge_Show()
Dim X As String

frmPaieSage.Show vbModeless
frmPaieSage.WindowState = vbNormal
frmPaieSage.Visible = True
X = frmPaieSage.Caption
AppActivate X

End Sub

Public Sub frmSAB_Ordonnanceur_Show()
Dim X As String

frmSAB_Ordonnanceur.Show vbModeless
frmSAB_Ordonnanceur.WindowState = vbNormal
frmSAB_Ordonnanceur.Visible = True
X = frmSAB_Ordonnanceur.Caption
AppActivate X

End Sub

Public Sub frmElpTimer_Show()

End Sub

