Attribute VB_Name = "Bia_DWH_Monitor"
Option Explicit


Type typeBiaUsr                                  ' compatibilité   BIA.vbp
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    ID                     As String * 10
    Nom                    As String * 34
    Service                As String * 3
    Coges                  As String * 2
    Groupe                 As String * 10

End Type


Public frmRTF_UsrId_Origine As String
Public frmRTF_Référence As String


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


End Sub



Public Sub mainSocExe()
frmElp_Caption = "BIA_DWH"
frmElp_Icon = paramFolder_Local & "\misc36.ico"

blnMonitor = True

End Sub


'---------------------------------------------------------
Public Sub Msg_Monitor(Msg As String)
'---------------------------------------------------------
If Not blnMonitor Then Exit Sub

Select Case UCase$(Trim(mId$(Msg, 1, 12)))
    Case Is = "DRENTACH": frmDRENTACH_Show: frmDRENTACH.Msg_Rcv Msg:
    Case Is = "DCOMM": frmDCOMM_Show: frmDCOMM.Msg_Rcv Msg:
    Case Is = "DCOUNIT": frmDCOUNIT_Show: frmDCOUNIT.Msg_Rcv Msg:
    Case Is = "DCRETRO": frmDCRETRO_Show: frmDCRETRO.Msg_Rcv Msg:
    Case Is = "DRENTA": frmDRENTA_Show: frmDRENTA.Msg_Rcv Msg:
    Case Is = "DAUTPIB": frmDAUTPIB_Show: frmDAUTPIB.Msg_Rcv Msg:
    Case Is = "X_RESET":  main_Reset
    Case Is = "XUSRID": XUsrId_Show
End Select

End Sub
Public Sub frmDRENTACH_Show()
Dim X As String

frmDRENTACH.Show vbModeless
frmDRENTACH.WindowState = vbNormal
frmDRENTACH.Visible = True
X = frmDRENTACH.Caption
AppActivate X

End Sub


Public Sub frmDCOMM_Show()
Dim X As String

frmDCOMM.Show vbModeless
frmDCOMM.WindowState = vbNormal
frmDCOMM.Visible = True
X = frmDCOMM.Caption
AppActivate X

End Sub

Public Sub frmDCRETRO_Show()
Dim X As String

frmDCRETRO.Show vbModeless
frmDCRETRO.WindowState = vbNormal
frmDCRETRO.Visible = True
X = frmDCRETRO.Caption
AppActivate X

End Sub

Public Sub frmDCOUNIT_Show()
Dim X As String

frmDCOUNIT.Show vbModeless
frmDCOUNIT.WindowState = vbNormal
frmDCOUNIT.Visible = True
X = frmDCOUNIT.Caption
AppActivate X

End Sub

Public Sub frmDRENTA_Show()
Dim X As String

frmDRENTA.Show vbModeless
frmDRENTA.WindowState = vbNormal
frmDRENTA.Visible = True
X = frmDRENTA.Caption
AppActivate X

End Sub

Public Sub frmDAUTPIB_Show()
Dim X As String

frmDAUTPIB.Show vbModeless
frmDAUTPIB.WindowState = vbNormal
frmDAUTPIB.Visible = True
X = frmDAUTPIB.Caption
AppActivate X

End Sub


Public Sub mainSoc_YBase_Load()
'pour compatibilité
End Sub
