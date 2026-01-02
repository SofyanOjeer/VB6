Attribute VB_Name = "Bia_DWH_Monitor"
Option Explicit



Public Sub mainSoc_Close()



End Sub



Public Sub mainSocExe()

paramIMP_PDFCreator_Name = "PDF_BIA_DWH"
paramIMP_PDF_Path_VBP = "C:\Temp\IMP_PDF\BIA_DWH"

If Not msFileSystem.FolderExists(paramIMP_PDF_Path_VBP) Then paramIMP_PDF_Path_VBP = paramIMP_PDF_Path_Temp
paramIMP_PDF_Path = paramIMP_PDF_Path_Temp

frmElp_Caption = "BIA_DWH"
Set frmElp_Icon = frmDRENTA

blnMonitor = True

End Sub


'---------------------------------------------------------
Public Sub Msg_Monitor(Msg As String)
'---------------------------------------------------------
If Not blnMonitor Then Exit Sub
Select Case UCase$(Trim(Mid$(Msg, 1, 12)))
    Case Is = "DRENTACH": frmDRENTACH_Show: frmDRENTACH.Msg_Rcv Msg:
    Case Is = "DCOMM": frmDCOMM_Show: frmDCOMM.Msg_Rcv Msg:
    Case Is = "DCOUNIT": frmDCOUNIT_Show: frmDCOUNIT.Msg_Rcv Msg:
    Case Is = "DCRETRO": frmDCRETRO_Show: frmDCRETRO.Msg_Rcv Msg:
    Case Is = "DRENTA": frmDRENTA_Show: frmDRENTA.Msg_Rcv Msg:
    Case Is = "DAUTPIB": frmDAUTPIB_Show: frmDAUTPIB.Msg_Rcv Msg:
    Case Is = "DAUTLIB0": frmDAUTLIB0_Show: frmDAUTLIB0.Msg_Rcv Msg:
    Case Is = "DGAPPIS0": frmDGAPPIS0_Show: frmDGAPPIS0.Msg_Rcv Msg:
    Case Is = "DCREINT0": frmDCREINT0_Show: frmDCREINT0.Msg_Rcv Msg:
    Case Is = "DBIASTO0": frmDBIASTO0_Show: frmDBIASTO0.Msg_Rcv Msg:
    Case Is = "DWH_STATUT": frmDWH_Statut_Show: frmDWH_Statut.Msg_Rcv Msg:
    Case Is = "DWH_ALM": frmDWH_ALM_Show: frmDWH_ALM.Msg_Rcv Msg:
    Case Is = "X_RESET":  main_Reset
    Case Is = "XUSRID": XUsrId_Show
    Case Is = "X_I5A7": X_I5A7_Show

End Select

End Sub
Public Sub frmDRENTACH_Show()
Dim X As String
frmDRENTACH.Icon = frmElp_Icon
frmDRENTACH.Show vbModeless
frmDRENTACH.WindowState = vbNormal
frmDRENTACH.Visible = True
X = frmDRENTACH.Caption
AppActivate X

End Sub


Public Sub frmDCOMM_Show()
Dim X As String

frmDCOMM.Icon = frmElp_Icon
frmDCOMM.Show vbModeless
frmDCOMM.WindowState = vbNormal
frmDCOMM.Visible = True
X = frmDCOMM.Caption
AppActivate X

End Sub

Public Sub frmDWH_Statut_Show()
Dim X As String

frmDWH_Statut.Icon = frmElp_Icon
frmDWH_Statut.Show vbModeless
frmDWH_Statut.WindowState = vbNormal
frmDWH_Statut.Visible = True
X = frmDWH_Statut.Caption
AppActivate X

End Sub

Public Sub frmDWH_ALM_Show()
Dim X As String

frmDWH_ALM.Icon = frmElp_Icon
frmDWH_ALM.Show vbModeless
frmDWH_ALM.WindowState = vbNormal
frmDWH_ALM.Visible = True
X = frmDWH_ALM.Caption
AppActivate X

End Sub

Public Sub frmDCRETRO_Show()
Dim X As String

frmDCRETRO.Icon = frmElp_Icon
frmDCRETRO.Show vbModeless
frmDCRETRO.WindowState = vbNormal
frmDCRETRO.Visible = True
X = frmDCRETRO.Caption
AppActivate X

End Sub

Public Sub frmDCOUNIT_Show()
Dim X As String

frmDCOUNIT.Icon = frmElp_Icon
frmDCOUNIT.Show vbModeless
frmDCOUNIT.WindowState = vbNormal
frmDCOUNIT.Visible = True
X = frmDCOUNIT.Caption
AppActivate X

End Sub

Public Sub frmDRENTA_Show()
Dim X As String

frmDRENTA.Icon = frmElp_Icon
frmDRENTA.Show vbModeless
frmDRENTA.WindowState = vbNormal
frmDRENTA.Visible = True
X = frmDRENTA.Caption
AppActivate X

End Sub

Public Sub frmDAUTPIB_Show()
Dim X As String

frmDAUTPIB.Icon = frmElp_Icon
frmDAUTPIB.Show vbModeless
frmDAUTPIB.WindowState = vbNormal
frmDAUTPIB.Visible = True
X = frmDAUTPIB.Caption
AppActivate X

End Sub


Public Sub frmDAUTLIB0_Show()
Dim X As String

frmDAUTLIB0.Icon = frmElp_Icon
frmDAUTLIB0.Show vbModeless
frmDAUTLIB0.WindowState = vbNormal
frmDAUTLIB0.Visible = True
X = frmDAUTLIB0.Caption
AppActivate X

End Sub
Public Sub frmDGAPPIS0_Show()
Dim X As String

frmDGAPPIS0.Icon = frmElp_Icon
frmDGAPPIS0.Show vbModeless
frmDGAPPIS0.WindowState = vbNormal
frmDGAPPIS0.Visible = True
X = frmDGAPPIS0.Caption
AppActivate X

End Sub

Public Sub frmDCREINT0_Show()
Dim X As String

frmDCREINT0.Icon = frmElp_Icon
frmDCREINT0.Show vbModeless
frmDCREINT0.WindowState = vbNormal
frmDCREINT0.Visible = True
X = frmDCREINT0.Caption
AppActivate X

End Sub

Public Sub frmDBIASTO0_Show()
Dim X As String

frmDBIASTO0.Icon = frmElp_Icon
frmDBIASTO0.Show vbModeless
frmDBIASTO0.WindowState = vbNormal
frmDBIASTO0.Visible = True
X = frmDBIASTO0.Caption
AppActivate X

End Sub

