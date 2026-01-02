Attribute VB_Name = "BiaCD_Monitor"
Option Explicit

Public Sub mainSoc_Close()

tableCptP0_Close
tableElpBuffer_Close
tableElpTable_Close
tableMvtP0_Close

End Sub

Public Sub mainSocExe()
Dim I As Integer

AccAutId = "SRVBIALR  "
frmElp_Caption = "BiaCD"
frmElp_Icon = paramFolder_Local & "\misc36.ico"



End Sub

Public Sub frmCptComPays_Show()
Dim X As String

frmCptComPays.Show vbModeless
frmCptComPays.WindowState = vbNormal
frmCptComPays.Visible = True
X = frmCptComPays.Caption
AppActivate X

End Sub

Public Sub frmCDTauPf_Show()
Dim X As String

frmCDTauPf.Show vbModeless
frmCDTauPf.WindowState = vbNormal
frmCDTauPf.Visible = True
X = frmCDTauPf.Caption
AppActivate X

End Sub

Public Sub frmCDStat_Show()
Dim X As String

frmCDStat.Show vbModeless
frmCDStat.WindowState = vbNormal
frmCDStat.Visible = True
X = frmCDStat.Caption
AppActivate X

End Sub

Public Sub frmCDListe_Show()
Dim X As String

frmCDListe.Show vbModeless
frmCDListe.WindowState = vbNormal
frmCDListe.Visible = True
X = frmCDListe.Caption
AppActivate X

End Sub


'---------------------------------------------------------
Public Sub Msg_Monitor(Msg As String)
'---------------------------------------------------------
Select Case UCase$(Trim(mId$(Msg, 1, 12)))
    Case Is = "CD_STAT", "FRMCDSTAT": frmCDStat_Show: frmCDStat.Msg_Rcv Msg
    Case Is = "CD_LISTE": frmCDListe_Show: frmCDListe.Msg_Rcv Msg
    Case Is = "CD_COM.TAUX", "FRMCDTAUPF": frmCDTauPf_Show: frmCDTauPf.Msg_Rcv Msg
    Case Is = "CPT_COMPAYS": frmCptComPays_Show: frmCptComPays.Msg_Rcv Msg
    Case Is = "XUSRID_BIACP": XUsrId_Show
    Case Is = "TIMER": ElpTimer_Init
    Case Is = "X_RESET": main_Reset

End Select

End Sub



