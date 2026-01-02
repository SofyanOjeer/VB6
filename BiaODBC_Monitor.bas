Attribute VB_Name = "BiaODBC_Monitor"
Option Explicit

Public Sub mainSoc_Close()

tableCDComD_Close
tableCDDossier_Close
tableCDPosting_Close
tableCDTICom_Close
tableCDTIMaster_Close
tableCptP0_Close
tableElpBuffer_Close
tableElpTable_Close
tableMvtP0_Close

End Sub

Public Sub mainSocExe()
Dim I As Integer

AccAutId = "SRVBIALR  "
'frmElp_Caption = "BiaCD"   TN 30/10/01
frmElp_Caption = "BiaODBC"
frmElp_Icon = paramFolder_Local & "\misc36.ico"



End Sub

Public Sub frmTIAS400_Show()
Dim X As String

frmTIAS400.Show vbModeless
frmTIAS400.WindowState = vbNormal
frmTIAS400.Visible = True
X = frmTIAS400.Caption
AppActivate X

End Sub

'---------------------------------------------------------
Public Sub Msg_Monitor(Msg As String)
'---------------------------------------------------------
Select Case UCase$(Trim(mId$(Msg, 1, 12)))
    Case Is = "TIAS400", "@AUTO_TIAS40": frmTIAS400_Show: frmTIAS400.Msg_Rcv Msg
    Case Is = "XUSRID_BIACP": XUsrId_Show
    Case Is = "TIMER": ElpTimer_Init
    Case Is = "X_RESET": main_Reset
End Select

End Sub



