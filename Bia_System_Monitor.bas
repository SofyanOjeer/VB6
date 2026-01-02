Attribute VB_Name = "Bia_System_Monitor"
Option Explicit

Dim meYBIAMON0 As typeYBIAMON0
Dim meElpTable  As typeElpTable


Public Sub ACTION()
Dim I As Long
Dim fic As Long
Dim XPrt As Printer
Dim fichierLog As String

    fichierLog = "\\DOCSRV2013\_GROUPS\PUBLIC\_DOSSIERS PARTAGES\INFORMATIQUE\Log_BIA2008\debugImprimantes.log"
    fic = FreeFile
    Open fichierLog For Append As #fic
    Print #fic, "BIA_SYSTEM.EXE pour " & nomDuServeur
    I = 0
    For Each XPrt In Printers
        I = I + 1
        Print #fic, CStr(I) & " - " & XPrt.Devicename
    Next
    Print #fic, CStr(Printers.Count) & " imprimantes le " & Date & " à " & Time
    Print #fic, "====================================================="

End Sub

Public Sub ALERTE_ECH()

    'L'appel ALERTE échelles a été remplacé par une tâche planifiée Windows
    'Call Shell_Exe("\\BIASRV\BIASRV.APP\ALT_ECH\Call_Alerte.cmd")

End Sub

Public Sub frmSAB_Dossier_DB_Show()
Dim X As String

frmSAB_Dossier_DB.Icon = frmElp_Icon
frmSAB_Dossier_DB.Show vbModeless
frmSAB_Dossier_DB.WindowState = vbNormal
frmSAB_Dossier_DB.Visible = True
X = frmSAB_Dossier_DB.Caption
'AppActivate X

End Sub

Public Sub frmSAB_MNU_Show()
Dim X As String
On Error Resume Next
frmSAB_MNU.Icon = frmElp_Icon
frmSAB_MNU.Show vbModeless
frmSAB_MNU.WindowState = vbNormal
frmSAB_MNU.Visible = True
X = frmSAB_MNU.Caption
AppActivate X

End Sub

Public Sub frmSysInvent_Show()
Dim X As String
On Error Resume Next
frmSysInvent.Icon = frmElp_Icon
frmSysInvent.Show vbModeless
frmSysInvent.WindowState = vbNormal
frmSysInvent.Visible = True
X = frmSysInvent.Caption
AppActivate X

End Sub

Public Sub frmSAB_CDR_Show()
Dim X As String
On Error Resume Next
frmSAB_CDR.Icon = frmElp_Icon
frmSAB_CDR.Show vbModeless
frmSAB_CDR.WindowState = vbNormal
frmSAB_CDR.Visible = True
X = frmSAB_CDR.Caption
AppActivate X

End Sub

Public Sub frmYTP7OPH0_Show()
Dim X As String
On Error Resume Next
frmYTP7OPH0.Icon = frmElp_Icon
frmYTP7OPH0.Show vbModeless
frmYTP7OPH0.WindowState = vbNormal
frmYTP7OPH0.Visible = True
X = frmYTP7OPH0.Caption
AppActivate X

End Sub


Public Sub frmBIA_VB_Habilitations_Show()
Dim X As String
On Error Resume Next
frmBIA_VB_Habilitations.Icon = frmElp_Icon
frmBIA_VB_Habilitations.Show vbModeless
frmBIA_VB_Habilitations.WindowState = vbNormal
frmBIA_VB_Habilitations.Visible = True
X = frmBIA_VB_Habilitations.Caption
AppActivate X

End Sub
Public Sub frmYPCICPT0_Show()
Dim X As String
On Error Resume Next
frmYPCICPT0.Icon = frmElp_Icon
frmYPCICPT0.Show vbModeless
frmYPCICPT0.WindowState = vbNormal
frmYPCICPT0.Visible = True
X = frmYPCICPT0.Caption
AppActivate X

End Sub

Public Sub frmYPCICPT0_Auto()
Dim meYBIAMON0 As typeYBIAMON0, V
'--------------------------------------------------------------------------------------
meYBIAMON0.MONAPP = "@BIA_SYSTEM"
meYBIAMON0.MONFLUX = "@PCI_COMPTE"
meYBIAMON0.MONSTATUS = ""
Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, "> " & meYBIAMON0.MONAPP & " " & meYBIAMON0.MONFLUX): DoEvents

V = fctExploitation_Auto_Control(meYBIAMON0)
If IsNull(V) Then
    Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, "> " & meYBIAMON0.MONAPP & " " & meYBIAMON0.MONFLUX & " en cours ...."): DoEvents
    V = cnSAB_Transaction("Commit")
    'If blnAuto_Form_Show Then
            frmYPCICPT0_Show
    frmYPCICPT0.Msg_Rcv "@PCI_COMPTE"
    V = cnSAB_Transaction("BeginTrans")
    V = fctExploitation_Auto_End(meYBIAMON0)
End If
Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, "< " & meYBIAMON0.MONAPP & " " & meYBIAMON0.MONFLUX & " Terminé"): DoEvents

End Sub

Public Sub frmYFLUTPJ0_Show()
Dim X As String
On Error Resume Next
frmYFLUTPJ0.Icon = frmElp_Icon
frmYFLUTPJ0.Show vbModeless
frmYFLUTPJ0.WindowState = vbNormal
frmYFLUTPJ0.Visible = True
X = frmYFLUTPJ0.Caption
AppActivate X

End Sub
Public Sub frmEdition_Gestion_Show()
Dim X As String
On Error Resume Next
frmEdition_Gestion.Icon = frmElp_Icon

frmEdition_Gestion.Show vbModeless
frmEdition_Gestion.WindowState = vbNormal
frmEdition_Gestion.Visible = True
X = frmEdition_Gestion.Caption
AppActivate X

End Sub



Public Sub frmBiaPgm_Show()
Dim X As String
On Error Resume Next
frmBiaPgm.Icon = frmElp_Icon

frmBiaPgm.Show vbModeless
frmBiaPgm.WindowState = vbNormal
frmBiaPgm.Visible = True
X = frmBiaPgm.Caption
AppActivate X

End Sub

Public Sub mainSocExe()

paramIMP_PDFCreator_Name = "PDF_BIA_SYSTEM"
paramIMP_PDF_Path_VBP = "C:\Temp\IMP_PDF\BIA_SYSTEM"

If Not msFileSystem.FolderExists(paramIMP_PDF_Path_VBP) Then paramIMP_PDF_Path_VBP = paramIMP_PDF_Path_Temp
paramIMP_PDF_Path = paramIMP_PDF_Path_Temp

frmElp_Caption = "BIA_SYSTEM"
Set frmElp_Icon = frmBIA_System

blnMonitor = True

If xlsManual Then
    frmElp.fra0.BackColor = &HDCF2F8
End If

End Sub


'---------------------------------------------------------
Public Sub Msg_Monitor(Msg As String)
'---------------------------------------------------------
If Not blnMonitor Then Exit Sub

Select Case UCase$(Trim(mId$(Msg, 1, 12)))
    Case Is = "@SAVNONCL": SMS_SAVNONCL
    Case Is = "@SAVNONCL_OK": SMS_SAVNONCL_OK
    Case Is = "@YBIAJOUR": SMS_YBIAJOUR
    Case Is = "@YBIAJOUR_OK": SMS_YBIAJOUR_OK
    Case Is = "@SMS_COMPTA": SMS_COMPTA: SMS_AUTO_SYSTEM
    Case Is = "@SMS_JRN": SMS_JRN
    Case Is = "@SMS_SWIFT_S": SMS_SWIFT_S
    Case Is = "@SMS_SWIFT_E": SMS_SWIFT_E
    Case Is = "@SMS_SWIFT_R": SMS_SWIFT_R
    Case Is = "@SMS_SPLF_FT": SMS_SPLF_FT
    Case Is = "@SMS_SPLF_PR": SMS_SPLF_PR
    Case Is = "@SMS_ACTIF": SMS_Actif
    Case Is = "@SMS_ACTIF_O": SMS_Actif_O
    Case Is = "@WORLDCHECK": SMS_WorldCheck_New
    Case Is = "@ADES_SERVIC":  frmBia_NET_CMD.Msg_Rcv Msg:
    Case Is = "@EUP_LAB": frmEUP_LAB.Msg_Rcv Msg:
    Case Is = "@EUP_XCOM": frmEUP_XCOM.Msg_Rcv Msg:
    Case Is = "@SMS_EUP_LAB": SMS_EUP_LAB
    Case Is = "@AUTO_SYST": AUTO_SYSTEM
    Case Is = "@PCI_COMPTE": blnAuto_Form_Show = False: frmYPCICPT0_Auto
    'L'appel ALERTE échelles a été remplacé par une tâche planifiée Windows
    'Case Is = "@ALERTE_ECH": Call ALERTE_ECH

    Case Is = "ADES": frmBIA_NET_CMD_Show: frmBia_NET_CMD.Msg_Rcv Msg:
    Case Is = "BDF_CMP": frmBDF_CMP_Show: frmBDF_CMP.Msg_Rcv Msg:
    Case Is = "BDF_CRT": frmBDF_CRT_Show: frmBDF_CRT.Msg_Rcv Msg:
    Case Is = "BIA_ACCESS": frmBIA_Access_Show: frmBIA_Access.Msg_Rcv Msg:
    Case Is = "EUP_LAB": frmEUP_LAB_Show: frmEUP_LAB.Msg_Rcv Msg:
    Case Is = "EUP_XCOM", "EUP_XCOMTEST": frmEUP_XCOM_Show: frmEUP_XCOM.Msg_Rcv Msg:
    Case Is = "BIA_EXPLOIT": frmBIA_Exploitation_Show: frmBia_Exploitation.Msg_Rcv Msg:
    Case Is = "BIA_SYSTEM": frmBIA_System_Show: frmBIA_System.Msg_Rcv Msg:
    Case Is = "BIAPGM": frmBiaPgm_Show: frmBiaPgm.Msg_Rcv Msg
    Case Is = "BIAPGM_AUT": frmBiaPgmAut_Show: frmBiaPgmAut.Msg_Rcv Msg
    Case Is = "EDITION_GEST": frmEdition_Gestion_Show: frmEdition_Gestion.Msg_Rcv Msg:
    Case Is = "LAB": frmLAB_Show: frmLAB.Msg_Rcv Msg:
    Case Is = "SAB_MNU": frmSAB_MNU_Show: frmSAB_MNU.Msg_Rcv Msg:
    Case Is = "SAB_CDR": frmSAB_CDR_Show: frmSAB_CDR.Msg_Rcv Msg:
    Case Is = "SAB_DER": frmSAB_DER_Show: frmSAB_DER.Msg_Rcv Msg:
    Case Is = "TABLE", "FRMELPTABLE": frmElpTable_Show: frmElpTable.Msg_Rcv Msg
    Case Is = "SYSINVENT": frmSysInvent_Show: frmSysInvent.Msg_Rcv Msg:
    Case Is = "FLUX_NOSTRO", "TP7_OPH": frmYTP7OPH0_Show: frmYTP7OPH0.Msg_Rcv Msg:
    Case Is = "FLUX_TREPREV", "@FLUX_TREPRE": frmYFLUTPJ0_Show: frmYFLUTPJ0.Msg_Rcv Msg:
    Case Is = "PCI_COMPTE": frmYPCICPT0_Show: frmYPCICPT0.Msg_Rcv Msg:
    Case Is = "BIA_VB_HAB": frmBIA_VB_Habilitations_Show: frmBIA_VB_Habilitations.Msg_Rcv Msg:
    Case Is = "ACP": frmACP_Show: frmACP.Msg_Rcv Msg:
    Case Is = "X_RESET":  main_Reset
    Case Is = "XUSRID": XUsrId_Show
    Case Is = "X_I5A7": X_I5A7_Show
End Select

End Sub
Public Function retourne_newDate() As String

'    retourne_newDate = CStr(Year(Now)) & "/" & mId(CStr(100 + Month(Now)), 2) & "/" & mId(CStr(100 + Day(Now)), 2) & " " & mId(CStr(100 + Hour(Now)), 2) & ":" & mId(CStr(100 + Minute(Now)), 2) & ":" & mId(CStr(100 + Second(Now)), 2)
    retourne_newDate = mId(CStr(100 + Day(Now)), 2) & "/" & mId(CStr(100 + Month(Now)), 2) & "/" & CStr(Year(Now)) & " " & mId(CStr(100 + Hour(Now)), 2) & ":" & mId(CStr(100 + Minute(Now)), 2) & ":" & mId(CStr(100 + Second(Now)), 2)
    
End Function

Public Sub SMS_EUP_LAB()
Dim V, blnSMS As Boolean
Dim wMemo As String
Dim wSUBJECT As String, wMotif As String

'A REPRENDRE

'Call lstErr_Clear(frmElp.lstErr, frmElp.cmdContext, "> SMS_EUP_LAB.....")

'blnSMS = False
'V = rsYBIATAB0_Read("SEPA", "MONITORING", "SUIVI", wMemo)
'If IsNull(V) Then
'    Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, ": SMS_EUP_LAB : SEPA [MONITORING [ SUIVI ]]]")
'    If Trim(meYBIATAB0.MONSTATUS) = "ACTIF" Then
'       If Trim(meYBIATAB0.MONFILE) <> YBIATAB0_DATE_CPT_JS1 Then blnSMS = True: wSUBJECT = "BIAJOUR : lancé, Date ? " & meYBIATAB0.MONFILE & " <> " & dateImp10(YBIATAB0_DATE_CPT_JS1)
'    Else
'       blnSMS = True: wSUBJECT = "BIAJOUR : non lancé, compta du " & dateImp10(YBIATAB0_DATE_CPT_JS1)
'    End If
'Else
'    Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, "? SMS_EUP_LAB err : " & V)
'    Call cmdSendMail_Alerte("SMS_EUP_LAB : ERREUR accès iSeries ", "SAB073SPE/YBIATAB0 [SEPA [MONITORING [ SUIVI ]]]")
'End If

'Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, "< SMS_EUP_LAB....Fin")


End Sub

Public Sub frmBiaPgmAut_Show()
Dim X As String
On Error Resume Next
frmBiaPgmAut.Icon = frmElp_Icon

frmBiaPgmAut.Show vbModeless
frmBiaPgmAut.WindowState = vbNormal
frmBiaPgmAut.Visible = True
X = frmBiaPgmAut.Caption
AppActivate X
End Sub



'---------------------------------------------------------
Public Sub frmElpTable_Show()
'---------------------------------------------------------
Dim X As String
On Error Resume Next
frmElpTable.Icon = frmElp_Icon 'LoadPicture(frmElp_Icon)

frmElpTable.Show vbModeless
frmElpTable.WindowState = vbNormal
frmElpTable.Visible = True
X = frmElpTable.Caption
AppActivate X

End Sub

Public Sub frmBIA_System_Show()
Dim X As String
On Error Resume Next
frmBIA_System.Icon = frmElp_Icon

frmBIA_System.Show vbModeless
frmBIA_System.WindowState = vbNormal
frmBIA_System.Visible = True
X = frmBIA_System.Caption
AppActivate X

End Sub

Public Sub frmBIA_Access_Show()
Dim X As String
On Error Resume Next
frmBIA_Access.Icon = frmElp_Icon
frmBIA_Access.Show vbModeless
frmBIA_Access.WindowState = vbNormal
frmBIA_Access.Visible = True
X = frmBIA_Access.Caption
AppActivate X

End Sub


Public Sub frmBIA_Exploitation_Show()
Dim X As String
On Error Resume Next
frmBia_Exploitation.Icon = frmElp_Icon
frmBia_Exploitation.Show vbModeless
frmBia_Exploitation.WindowState = vbNormal
frmBia_Exploitation.Visible = True
X = frmBia_Exploitation.Caption
AppActivate X

End Sub
Public Sub frmEUP_LAB_Show()
Dim X As String
On Error Resume Next
frmEUP_LAB.Icon = frmElp_Icon
frmEUP_LAB.Show vbModeless
frmEUP_LAB.WindowState = vbNormal
frmEUP_LAB.Visible = True
X = frmEUP_LAB.Caption
AppActivate X

End Sub

Public Sub frmEUP_XCOM_Show()
Dim X As String
On Error Resume Next
frmEUP_XCOM.Icon = frmElp_Icon
frmEUP_XCOM.Show vbModeless
frmEUP_XCOM.WindowState = vbNormal
frmEUP_XCOM.Visible = True
X = frmEUP_XCOM.Caption
AppActivate X

End Sub

Public Sub frmSAB_DER_Show()
Dim X As String
On Error Resume Next
frmSAB_DER.Icon = frmElp_Icon
frmSAB_DER.Show vbModeless
frmSAB_DER.WindowState = vbNormal
frmSAB_DER.Visible = True
X = frmSAB_DER.Caption
AppActivate X

End Sub
Public Sub frmBDF_CMP_Show()
Dim X As String
On Error Resume Next
frmBDF_CMP.Icon = frmElp_Icon
frmBDF_CMP.Show vbModeless
frmBDF_CMP.WindowState = vbNormal
frmBDF_CMP.Visible = True
X = frmBDF_CMP.Caption
AppActivate X

End Sub

Public Sub frmBDF_CRT_Show()
Dim X As String
On Error Resume Next
frmBDF_CRT.Icon = frmElp_Icon
frmBDF_CRT.Show vbModeless
frmBDF_CRT.WindowState = vbNormal
frmBDF_CRT.Visible = True
X = frmBDF_CRT.Caption
AppActivate X

End Sub
Public Sub frmACP_Show()
Dim X As String
On Error Resume Next
frmACP.Icon = frmElp_Icon
frmACP.Show vbModeless
frmACP.WindowState = vbNormal
frmACP.Visible = True
X = frmACP.Caption
AppActivate X

End Sub

Public Sub frmBIA_NET_CMD_Show()
Dim X As String
On Error Resume Next
frmBia_NET_CMD.Icon = frmElp_Icon
frmBia_NET_CMD.Show vbModeless
frmBia_NET_CMD.WindowState = vbNormal
frmBia_NET_CMD.Visible = True
X = frmBia_NET_CMD.Caption
AppActivate X

End Sub

Public Sub frmLAB_Show()
Dim X As String
On Error Resume Next
frmLAB.Icon = frmElp_Icon
frmLAB.Show vbModeless
frmLAB.WindowState = vbNormal
frmLAB.Visible = True
X = frmLAB.Caption
AppActivate X

End Sub

Public Sub SMS_SAVNONCL()
Dim V, blnSMS As Boolean
Dim wSUBJECT As String, wMotif As String
Dim paramFolder_Path As String, xFileName  As String
Dim X8 As String
Dim objFolder, objFiles
Dim fsoFile As File, fsoFile2 As File, fsoFile3 As File, TS As TextStream, X As String
Dim blnQCTL As Boolean, blnCOMPTA As Boolean, blnCONS400 As Boolean, blnCONS400_Sta As Boolean
Dim blnQPDSPAJB As Boolean, blnCOMPTA_JOB As Boolean
Dim blnQPRTSPLQ As Boolean, blnQPRTSPLQ_Sta As Boolean
Dim I As Integer

Call lstErr_Clear(frmElp.lstErr, frmElp.cmdContext, "> SMS_SAVNONCL......")
xFileName = ""
blnQPDSPAJB = False
blnSMS = False
blnQCTL = False
blnCOMPTA = False: blnCOMPTA_JOB = False
blnCONS400 = False
blnCONS400_Sta = False
blnQPRTSPLQ = False
blnQPRTSPLQ_Sta = False

paramFolder_Path = paramEditionSplf_Folder & constProduction
wMotif = paramFolder_Path

frmElp.filDoc.Path = paramFolder_Path
frmElp.filDoc.Pattern = "x.xxx"
frmElp.filDoc.Pattern = "*QPDSPAJB*.txt"
For I = 0 To frmElp.filDoc.ListCount - 1
    frmElp.filDoc.ListIndex = I
    Set fsoFile = msFileSystem.GetFile(frmElp.filDoc.Path & "\" & frmElp.filDoc.FileName)
       wMotif = wMotif & "//" & DateDiff("n", fsoFile.DateLastModified, Now) & " " & fsoFile.Name '

        If DateDiff("n", fsoFile.DateLastModified, Now) < 30 Then
            blnQPDSPAJB = True
            Set fsoFile2 = fsoFile
        End If
Next I

'Set objFolder = msFileSystem.GetFolder(paramFolder_Path)
'Set objFiles = objFolder.Files
'For Each fsoFile In objFiles
'    If Err = 0 Then
'       If InStr(fsoFile.Name, "QPDSPAJB") Then
'       wMotif = wMotif & "//" & DateDiff("n", fsoFile.DateLastModified, Now) & " " & fsoFile.Name'

'            If DateDiff("n", fsoFile.DateLastModified, Now) < 30 Then
'                blnQPDSPAJB = True
'                Set fsoFile2 = fsoFile
'            End If
'        End If
'        If InStr(fsoFile.Name, "QPRTSPLQ") Then
'            If DateDiff("n", fsoFile.DateLastModified, Now) < 30 Then
'                blnQPRTSPLQ = True
'                Set fsoFile3 = fsoFile
'            End If
'        End If
'   End If
'Next
If Not blnQPDSPAJB Then
    blnSMS = True
    wSUBJECT = "@SAVNONCL : pas de fichier QPDSPAJB récent (30 minutes)"
    
Else
    Set TS = fsoFile2.OpenAsTextStream(ForReading, TristateUseDefault)
    Do Until TS.AtEndOfStream
        X = TS.ReadLine
        
        If mId$(X, 54, 3) = "SBS" Then   '42 => 54
            blnQCTL = False: blnCOMPTA = False
            Select Case Trim(mId$(X, 6, 10))
                Case "QCTL": blnQCTL = True: wMotif = X
                Case "COMPTA": blnCOMPTA = True: wMotif = X
            End Select
        Else
            If blnCOMPTA Then
                wMotif = X
                wSUBJECT = "Il y a un job actif dans COMPTA.sbs  !!!!"
                Call SMS_MONITOR("@SAVNONCL", wSUBJECT, wMotif)
            End If
            If blnQCTL Then
                'If mId$(X, 1, 14) = "   1   CONS400" Then
                'xxxx Modif HMC du 15/10/2009
                 If mId$(X, 1, 12) = "   1   DSP01" Or mId$(X, 1, 12) = "   1   DSP02" Then
                    wMotif = X
                    blnCONS400 = True
                    If mId$(X, 120, 4) = "DLYW" Then blnCONS400_Sta = True   '115
                End If
            End If
        End If
    Loop
    TS.Close
    If Not blnCONS400 Then
        blnSMS = True
        wSUBJECT = "@SAVNONCL : non lancé, compta du " & dateImp10(YBIATAB0_DATE_CPT_JS1)
    Else
        If Not blnCONS400_Sta Then
            blnSMS = True
            wSUBJECT = "@SAVNONCL : non opérationnel, compta du  " & dateImp10(YBIATAB0_DATE_CPT_JS1)
        End If
    End If
End If

'27/09/2019 DR désactivé. A réactiver pour le 30/09/2019
'If blnSMS Then
'    Call SMS_MONITOR("@SAVNONCL", wSUBJECT, wMotif)
'End If

'_____________________________________________________________________________
blnSMS = False
wMotif = paramFolder_Path
frmElp.filDoc.Path = paramFolder_Path
frmElp.filDoc.Pattern = "x.xxx"
frmElp.filDoc.Pattern = "*QPRTSPLQ*.txt"
For I = 0 To frmElp.filDoc.ListCount - 1
    frmElp.filDoc.ListIndex = I
    Set fsoFile = msFileSystem.GetFile(frmElp.filDoc.Path & "\" & frmElp.filDoc.FileName)
       wMotif = wMotif & "//" & DateDiff("n", fsoFile.DateLastModified, Now) & " " & fsoFile.Name '

        If DateDiff("n", fsoFile.DateLastModified, Now) < 30 Then
            blnQPRTSPLQ = True
            Set fsoFile3 = fsoFile
        End If
Next I

If Not blnQPRTSPLQ Then
    blnSMS = True
    wSUBJECT = "@SAVNONCL : pas de fichier QPRTSPLQ récent (30 minutes)"
Else
    Set TS = fsoFile3.OpenAsTextStream(ForReading, TristateUseDefault)
    Do Until TS.AtEndOfStream
        X = TS.ReadLine
        
        If Trim(mId$(X, 6, 10)) = "BIAJOUR" Then
            wMotif = X
            If blnQPRTSPLQ_Sta Then
                wSUBJECT = "Il y a plusieurs jobs dans COMPTA.sbs  !!!!"
                Call SMS_MONITOR("@SAVNONCL", wSUBJECT, wMotif)
            Else
                blnQPRTSPLQ_Sta = True
                If mId$(X, 53, 3) <> "SCD" Then
                     wSUBJECT = "voir le statut du job BIAJOUR dans COMPTA.sbs  !!!!"
                    Call SMS_MONITOR("@SAVNONCL", wSUBJECT, wMotif)
                End If
            End If
        End If
   Loop
    TS.Close
    If Not blnQPRTSPLQ_Sta Then
        blnSMS = True
        wSUBJECT = "manque BIAJOUR dans la queue COMPTA " & dateImp10(YBIATAB0_DATE_CPT_JS1)
    End If
End If

If blnSMS Then
    Call SMS_MONITOR("@SAVNONCL", wSUBJECT, wMotif)
End If
'______________________________________________________________________________________
Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, "< SMS_SAVNONCL....Fin")


End Sub

Public Sub SMS_YBIAJOUR()
Dim V, blnSMS As Boolean
Dim wSUBJECT As String, wMotif As String

Call lstErr_Clear(frmElp.lstErr, frmElp.cmdContext, "> SMS_YBIAJOUR......")

blnSMS = False
meYBIAMON0.MONAPP = "SMS"
meYBIAMON0.MONFLUX = "@YBIAJOUR"
V = rsYBIAMON0_Read(meYBIAMON0)
If IsNull(V) Then
    Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, ": SMS_YBIAJOUR : " & Trim(meYBIAMON0.MONFILE) & meYBIAMON0.MONFILE)
    If Trim(meYBIAMON0.MONSTATUS) = "ACTIF" Then
       If Trim(meYBIAMON0.MONFILE) <> YBIATAB0_DATE_CPT_JS1 Then blnSMS = True: wSUBJECT = "BIAJOUR : lancé, Date ? " & meYBIAMON0.MONFILE & " <> " & dateImp10(YBIATAB0_DATE_CPT_JS1)
    Else
       blnSMS = True: wSUBJECT = "BIAJOUR : non lancé, compta du " & dateImp10(YBIATAB0_DATE_CPT_JS1)
    End If
Else
    Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, "? SMS_YBIAJOUR err : " & V)
     blnSMS = True: wSUBJECT = "SMS_YBIAJOUR : ERREUR accès iSeries "
     meYBIAMON0.MONSTATUS = "?????"
     meYBIAMON0.MONFILE = "XXXXXXXX"
End If
If blnSMS Then
    wMotif = "'SAB073SPE/YBIAMON7' : '" & meYBIAMON0.MONAPP & "' '" & meYBIAMON0.MONFLUX _
           & "?' > Statut : '" & meYBIAMON0.MONSTATUS & "' '" & meYBIAMON0.MONFILE & "'."
    Call SMS_MONITOR("@YBIAJOUR", wSUBJECT, wMotif)
End If

Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, "< SMS_YBIAJOUR....Fin")


End Sub


Public Sub SMS_SAVNONCL_OK()
Dim V, blnSMS As Boolean
Dim wSUBJECT As String, wMotif As String
Dim paramFolder_Path As String, xFileName  As String
Dim X8 As String
Dim objFolder, objFiles
Dim fsoFile As File, fsoFile2 As File, TS As TextStream, X As String
Dim blnQCTL As Boolean, blnCOMPTA As Boolean, blnCONS400 As Boolean, blnCONS400_Sta As Boolean
Dim blnQPDSPAJB As Boolean, blnCOMPTA_JOB As Boolean
Dim I As Integer

Call lstErr_Clear(frmElp.lstErr, frmElp.cmdContext, "> SMS_SAVNONCL_OK......")
xFileName = ""
blnQPDSPAJB = False
blnSMS = False
blnQCTL = False
blnCOMPTA = False: blnCOMPTA_JOB = False
blnCONS400 = False
blnCONS400_Sta = False

paramFolder_Path = paramEditionSplf_Folder & constProduction
frmElp.filDoc.Path = paramFolder_Path
frmElp.filDoc.Pattern = "x.xxx"
frmElp.filDoc.Pattern = "*QPDSPAJB*.txt"
For I = 0 To frmElp.filDoc.ListCount - 1
    frmElp.filDoc.ListIndex = I
    Set fsoFile = msFileSystem.GetFile(frmElp.filDoc.Path & "\" & frmElp.filDoc.FileName)
       wMotif = wMotif & "//" & DateDiff("n", fsoFile.DateLastModified, Now) & " " & fsoFile.Name '

        If DateDiff("n", fsoFile.DateLastModified, Now) < 30 Then
            blnQPDSPAJB = True
            Set fsoFile2 = fsoFile
        End If
Next I

'Set objFolder = msFileSystem.GetFolder(paramFolder_Path)
'Set objFiles = objFolder.Files
'For Each fsoFile In objFiles
'    If Err = 0 Then
'       If InStr(fsoFile.Name, "QPDSPAJB") Then
'        Debug.Print DateDiff("s", fsoFile.DateLastModified, Now)
'        If DateDiff("n", fsoFile.DateLastModified, Now) < 30 Then
'            blnQPDSPAJB = True
'            Set fsoFile2 = fsoFile
'        End If
'        End If
'    End If
'Next
If Not blnQPDSPAJB Then
    blnSMS = True
    wSUBJECT = "@SAVNONCL_OK : pas de fichier QPDSPAJB récent (30 minutes)"
    wMotif = paramFolder_Path
Else
    Set TS = fsoFile2.OpenAsTextStream(ForReading, TristateUseDefault)
    Do Until TS.AtEndOfStream
        X = TS.ReadLine
        
        If mId$(X, 54, 3) = "SBS" Then
            blnQCTL = False: blnCOMPTA = False
            Select Case Trim(mId$(X, 6, 10))
                Case "QCTL": blnQCTL = True: wMotif = X
                Case "COMPTA": blnCOMPTA = True: wMotif = X
            End Select
        Else
            If blnCOMPTA Then
                wMotif = X
                wSUBJECT = "Il y a un job actif dans COMPTA.sbs  !!!!"
                Call SMS_MONITOR("@SAVNONCL_OK", wSUBJECT, wMotif)
            End If
            If blnQCTL Then
                'If mId$(X, 1, 14) = "   1   CONS400" Then
                'xxxx Modif HMC du 15/10/2009
                 If mId$(X, 1, 12) = "   1   DSP01" Or mId$(X, 1, 12) = "   1   DSP02" Then
                    blnCONS400 = True
                    If mId$(X, 120, 4) <> "DSPW" Then blnCONS400_Sta = True: wMotif = X
                End If
            End If
        End If
    Loop
    TS.Close
    If blnCONS400_Sta Then
        blnSMS = True
        wSUBJECT = "@SAVNONCL_OK : non terminé, compta du " & dateImp10(YBIATAB0_DATE_CPT_JS1)
    End If
End If

If blnSMS Then
    Call SMS_MONITOR("@SAVNONCL_OK", wSUBJECT, wMotif)
End If
Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, "< SMS_SAVNONCL_ok....Fin")

End Sub

Public Sub SMS_WorldCheck()
Dim V, blnSMS As Boolean
Dim wSUBJECT As String, wMotif As String
Dim xFileName   As String
Dim X8 As String
Dim fsoFile As File, TS As TextStream, X As String
Dim blnImport_Success As Boolean, blnImport_Week As Boolean
Dim blnImport_Start As Boolean
Dim I As Integer, Nb As Integer
Dim wImport_Start As String, wImport_End As String
'Dim wImport_Day As String, wImport_Week As String
Dim wImport_Successfully As String, wImport_Load As String


xFileName = "?"
wSUBJECT = "@WorldCheck : import du " & DSys
On Error GoTo Error_Handler

Call lstErr_Clear(frmElp.lstErr, frmElp.cmdContext, "> SMS_WorldCheck......")
blnImport_Start = False
blnSMS = False
blnImport_Success = True
blnImport_Week = False
If Weekday(Now) = vbMonday Then blnImport_Week = True
wImport_Start = "$" & dateImp10_S(DSys)
wImport_End = "@" & dateImp10_S(DSys)
X = "AZ"""

'wImport_Week = "World-Check file ""D:/SWIFT/APPLI/PRODUCTION/World-Check/input/world-check-week.csv"" successfully imported."
'wImport_Day = "World-Check file ""D:/SWIFT/APPLI/PRODUCTION/World-Check/input/world-check-day.csv"" successfully imported."
wImport_Load = "[INFO] Getting lists ID's."
'Denis 28/01/2011
'wImport_Successfully = " successfully imported."
wImport_Successfully = " WC_UPDATER TERMINEE"

xFileName = paramServer("\\World-Check\Log\Log.txt")
Nb = 0

Set fsoFile = msFileSystem.GetFile(xFileName)
Set TS = fsoFile.OpenAsTextStream(ForReading, TristateUseDefault)
Do Until TS.AtEndOfStream
    X = TS.ReadLine
    
    If mId$(X, 1, 11) = wImport_Start Then
        wMotif = X
        blnImport_Start = True
        
    End If
    If blnImport_Start Then
    Debug.Print X
        If InStr(X, wImport_Load) Then blnImport_Success = False
        'Denis 28/01/2011
        'If InStr(x, wImport_Successfully) Then blnImport_Success = True: NB = NB + 1
        If InStr(UCase(X), wImport_Successfully) Then blnImport_Success = True: Nb = Nb + 1
        
        'If blnImport_Week Then
        '    If X = wImport_Week Then blnImport_Week = False: blnImport_Success = True: Nb = Nb + 1
        'Else
        '    If X = wImport_Day Or X = wImport_Week Then blnImport_Success = True: Nb = Nb + 1
        'End If
    End If
    If mId$(X, 1, 11) = wImport_End Then
        blnImport_Start = False
    End If
Loop
TS.Close

If blnImport_Start Then
    blnSMS = True
    wSUBJECT = "@WorldCheck : traitement non terminé " & dateImp10(DSys)
Else
    If Not blnImport_Success Then
        blnSMS = True
        wSUBJECT = "@WorldCheck : import SafeWatch non terminé " & dateImp10(DSys)
    'Else
    '    If blnImport_Week Then
    '        blnSMS = True
     '       wSUBJECT = "@WorldCheck : le fichier Week n'a pas été traité (lundi matin) " & dateImp10(DSys)
     '   End If
    End If
End If

If blnSMS Then
    Call SMS_MONITOR("@WorldCheck", wSUBJECT, wMotif)
End If
I = Val(mId$(Time, 1, 2))
If I < 19 Then
    If Nb < 1 Then
        wSUBJECT = "@WorldCheck : aucun traitement ce jour : " & dateImp10(DSys)
        Call SMS_MONITOR("@WorldCheck", wSUBJECT, wMotif)
    End If
Else
    If Nb < 2 Then
        wSUBJECT = "@WorldCheck : " & Nb & " traitement ce jour (matin/soir) : " & dateImp10(DSys)
        Call SMS_MONITOR("@WorldCheck", wSUBJECT, wMotif)
        
    End If
End If

Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, "< SMS_WorldCheck....Fin")
Exit Sub

Error_Handler:
    wMotif = xFileName & " : " & Error
    Call SMS_MONITOR("@WorldCheck", wSUBJECT, wMotif)


End Sub
Public Sub SMS_WorldCheck_New()
'Denis 24/04/2015
Dim V, blnSMS As Boolean
Dim wSUBJECT As String, wMotif As String
Dim xFileName   As String
Dim X8 As String
Dim fsoFile As File, TS As TextStream, X As String
Dim blnImport_Success As Boolean, blnImport_Week As Boolean
Dim blnImport_Start As Boolean
Dim I As Integer, Nb As Integer
Dim wImport_Start As String
Dim wImport_Successfully1 As String
Dim wImport_Successfully2 As String
Dim wImport_Load1 As String
Dim wImport_Load11 As String
Dim wImport_Load2 As String
Dim wImport_Complete1 As String
Dim wImport_Complete11 As String
Dim applicationMemo As String

    On Error GoTo Error_Handler
    xFileName = "?"
    wSUBJECT = "@WorldCheck : import du " & DSys
    Call lstErr_Clear(frmElp.lstErr, frmElp.cmdContext, "> SMS_WorldCheck......")
    blnImport_Start = False
    blnSMS = False
    blnImport_Success = True
    blnImport_Week = False
    If Weekday(Now) = vbMonday Then blnImport_Week = True
    wImport_Start = "@" & dateImp10_S(DSys)
    X = "AZ"""
    
    wImport_Load1 = "world-check-"
    wImport_Load11 = ".csv OK"
    wImport_Load2 = "[INFO]"
    wImport_Successfully1 = "Telechargement Termine"
    wImport_Successfully2 = "successfully imported."
    wImport_Complete1 = "WC_BATCH"
    wImport_Complete11 = "termine avec succes"
    xFileName = paramServer("\\World-Check\")
    V = rsElpTable_Read("Param", "Application", "World-Check", "log", applicationMemo)
    xFileName = xFileName & applicationMemo
    Nb = 0
    Set fsoFile = msFileSystem.GetFile(xFileName)
    Set TS = fsoFile.OpenAsTextStream(ForReading, TristateUseDefault)
    Do Until TS.AtEndOfStream
        X = TS.ReadLine
        If mId$(X, 1, 11) = wImport_Start Then
            wMotif = X
            blnImport_Start = True
        End If
        If blnImport_Start Then
            If InStr(UCase(X), UCase(wImport_Load1)) And InStr(UCase(X), UCase(wImport_Load11)) Then
                blnImport_Success = False
            End If
            If InStr(UCase(X), UCase(wImport_Load2)) Then
                blnImport_Success = False
            End If
            If InStr(UCase(X), UCase(wImport_Successfully1)) Or InStr(UCase(X), UCase(wImport_Successfully2)) Then
                blnImport_Success = True
                Nb = Nb + 1
            End If
        End If
        If InStr(UCase(X), UCase(wImport_Complete1)) And InStr(UCase(X), UCase(wImport_Complete11)) Then
            blnImport_Start = False
        End If
    Loop
    TS.Close
    
    If blnImport_Start Then
        blnSMS = True
        If wMotif = wImport_Successfully1 Then
            wMotif = wMotif & " - 1 Traitement sur 3"
        End If
        wSUBJECT = "@WorldCheck : traitement non terminé " & dateImp10(DSys)
    Else
        If Not blnImport_Success Then
            blnSMS = True
            wSUBJECT = "@WorldCheck : import SafeWatch non terminé " & dateImp10(DSys)
        End If
    End If
    
    If blnSMS Then
        Call SMS_MONITOR("@WorldCheck", wSUBJECT, wMotif)
    End If
    I = Val(mId$(Time, 1, 2))
    If I < 19 Then
        If Nb < 1 Then
            wSUBJECT = "@WorldCheck : aucun traitement ce jour : " & dateImp10(DSys)
            Call SMS_MONITOR("@WorldCheck", wSUBJECT, wMotif)
        End If
    Else
        If Nb < 2 Then
            wSUBJECT = "@WorldCheck : " & Nb & " traitement ce jour (matin/soir) : " & dateImp10(DSys)
            Call SMS_MONITOR("@WorldCheck", wSUBJECT, wMotif)
            
        End If
    End If
    
    Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, "< SMS_WorldCheck....Fin")
    Exit Sub

Error_Handler:
    wMotif = xFileName & " : " & Error
    Call SMS_MONITOR("@WorldCheck", wSUBJECT, wMotif)

End Sub

Public Sub SMS_COMPTA()
Dim V, blnSMS As Boolean
Dim wSUBJECT As String, wMotif As String
Call lstErr_Clear(frmElp.lstErr, frmElp.cmdContext, "> SMS_COMPTA......")

blnSMS = False
meYBIAMON0.MONAPP = "COMPTA"
meYBIAMON0.MONFLUX = "MAIL"
V = rsYBIAMON0_Read(meYBIAMON0)
If IsNull(V) Then
    Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, ": SMS_COMPTA : " & Trim(meYBIAMON0.MONFILE) & meYBIAMON0.MONFILE)
    If Trim(meYBIAMON0.MONSTATUS) = "" Then
       If Trim(meYBIAMON0.MONFILE) <> YBIATAB0_DATE_CPT_J Then blnSMS = True: wSUBJECT = "@COMPTA : OK Date ? " & meYBIAMON0.MONFILE & " <> " & dateImp10(YBIATAB0_DATE_CPT_JS1)
    Else
       blnSMS = True: wSUBJECT = "@COMPTA : statut " & Trim(meYBIAMON0.MONSTATUS) & " : " & meYBIAMON0.MONFILE
    End If
Else
    Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, "? SMS_COMPTA err : " & V)
     blnSMS = True: wSUBJECT = "SMS_COMPTA : ERREUR accès iSeries "
     meYBIAMON0.MONSTATUS = "?????"
     meYBIAMON0.MONFILE = "XXXXXXXX"
End If
If blnSMS Then
    wMotif = "'SAB073SPE/YBIAMON7' : '" & meYBIAMON0.MONAPP & "' '" & meYBIAMON0.MONFLUX _
           & "?' > Statut : '" & meYBIAMON0.MONSTATUS & "' '" & meYBIAMON0.MONFILE & "'."
    Call SMS_MONITOR("@SMS_COMPTA", wSUBJECT, wMotif)
End If

Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, "< SMS_COMPTA....Fin")


End Sub
Public Sub SMS_AUTO_SYSTEM()
Dim V, blnSMS As Boolean
Dim wSUBJECT As String, wMotif As String
Call lstErr_Clear(frmElp.lstErr, frmElp.cmdContext, "> SMS_AUTO_SYSTEM......")

blnSMS = False
meYBIAMON0.MONAPP = "@AUTO_SYST"
meYBIAMON0.MONFLUX = "MAIL"
V = rsYBIAMON0_Read(meYBIAMON0)
If IsNull(V) Then
    Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, ": SMS_AUTO_SYSTEM : " & Trim(meYBIAMON0.MONFILE) & meYBIAMON0.MONFILE)
    If Trim(meYBIAMON0.MONSTATUS) = "" Then
       If Trim(meYBIAMON0.MONFILE) <> YBIATAB0_DATE_CPT_J Then blnSMS = True: wSUBJECT = "@AUTO_SYSTEM : Date exe <> date compta" & meYBIAMON0.MONFILE & " <> " & dateImp10(YBIATAB0_DATE_CPT_JS1)
    Else
       blnSMS = True: wSUBJECT = "@AUTO_SYSTEM : statut " & Trim(meYBIAMON0.MONSTATUS) & " : " & meYBIAMON0.MONFILE
    End If
Else
    Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, "? SMS_AUTO_SYSTEM err : " & V)
     blnSMS = True: wSUBJECT = "SMS_AUTO_SYSTEM : ERREUR accès iSeries "
     meYBIAMON0.MONSTATUS = "?????"
     meYBIAMON0.MONFILE = "XXXXXXXX"
End If
If blnSMS Then
    wMotif = "'SAB073SPE/YBIAMON7' : '" & meYBIAMON0.MONAPP & "' '" & meYBIAMON0.MONFLUX _
           & "?' > Statut : '" & meYBIAMON0.MONSTATUS & "' '" & meYBIAMON0.MONFILE & "'."
    Call SMS_MONITOR("@SMS_COMPTA", wSUBJECT, wMotif)
End If

Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, "< SMS_AUTO_SYSTEM....Fin")


End Sub

Public Sub SMS_SWIFT_S()
Static mSWIALINUM As Long

Dim V, blnSMS As Boolean, xSql As String
Dim wSUBJECT As String, wMotif As String
Call lstErr_Clear(frmElp.lstErr, frmElp.cmdContext, "> SMS_SWIFT_S......")

blnSMS = False
Set rsSab = Nothing
If mSWIALINUM > 0 Then
    xSql = "select SWIALINUM from " & paramIBM_Library_SAB & ".ZSWIALI0 where SWIALINUM = " & mSWIALINUM
    Set rsSab = cnsab.Execute(xSql)
    If Not rsSab.EOF Then
       blnSMS = True: wSUBJECT = "@SMS_SWIFT-SAB a détecté le blocage de ZSWIALI0"
    End If
End If

xSql = "select SWIALINUM from " & paramIBM_Library_SAB & ".ZSWIALI0 order by SWIALINUM "
Set rsSab = cnsab.Execute(xSql)
If Not rsSab.EOF Then
    mSWIALINUM = rsSab("SWIALINUM")
Else
    mSWIALINUM = 0
End If


If blnSMS Then
    wMotif = "voir si l'application @TIMER_SWIFT sur \\PrintSrv est active."
    Call SMS_MONITOR("@SMS_SWIFT_S", wSUBJECT, wMotif)
End If

Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, "< SMS_SWIFT_S....Fin")


End Sub

Public Sub SMS_SWIFT_E()
Dim I As Integer, X As String

Dim vDate, blnSMS As Boolean
Dim wSUBJECT As String, wMotif As String
Call lstErr_Clear(frmElp.lstErr, frmElp.cmdContext, "> SMS_SWIFT_E......")

blnSMS = False
vDate = DateAdd("n", -10, Now)
frmElp.filDoc.Path = paramSAA_Data_from_SAB
frmElp.filDoc.Pattern = "x.xxx"
frmElp.filDoc.Pattern = "*" & paramSAA_Data_from_SAB_ExtensionP_sab
For I = 0 To frmElp.filDoc.ListCount - 1
    frmElp.filDoc.ListIndex = I
    Set msFile = msFileSystem.GetFile(frmElp.filDoc.Path & "\" & frmElp.filDoc.FileName)
    If msFile.DateLastModified < vDate Then blnSMS = True
Next I


If blnSMS Then
    wSUBJECT = "@SMS_SWIFT-Emission a détecté un blocage dans " & paramSAA_Data_from_SAB

    wMotif = "voir si l'application SAA sur \\SWIFTPROD est active."
    Call SMS_MONITOR("@SMS_SWIFT_E", wSUBJECT, wMotif)
End If

Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, "< SMS_SWIFT_E....Fin")


End Sub
Public Sub SMS_CORONA()
Dim I As Integer, X As String

Dim vDate, blnSMS As Boolean
Dim wSUBJECT As String, wMotif As String
Call lstErr_Clear(frmElp.lstErr, frmElp.cmdContext, "> SMS_CORONA......")

blnSMS = False
vDate = DateAdd("n", -10, Now)
frmElp.filDoc.Path = paramCorona_DataF_Swift_In
frmElp.filDoc.Pattern = "x.xxx"
frmElp.filDoc.Pattern = "*.*"
For I = 0 To frmElp.filDoc.ListCount - 1
    frmElp.filDoc.ListIndex = I
    Set msFile = msFileSystem.GetFile(frmElp.filDoc.Path & "\" & frmElp.filDoc.FileName)
    If msFile.DateLastModified < vDate Then blnSMS = True
Next I


If blnSMS Then
    wSUBJECT = "@SMS_SWIFT-Emission a détecté un blocage dans " & frmElp.filDoc.Path

    wMotif = "voir si l'application SAA sur \\SWIFTPROD est active."
    Call SMS_MONITOR("@SMS_CORONA", wSUBJECT, wMotif)
End If

Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, "< SMS_CORONA....Fin")


End Sub

Public Sub SMS_Actif_O()
Dim I As Integer, X As String

Dim vDate, blnSMS As Boolean
Dim wSUBJECT As String, wMotif As String
Call lstErr_Clear(frmElp.lstErr, frmElp.cmdContext, "> SMS_Actif_Ok......")



'If blnSMS Then
    wMotif = "Envoyé le " & Now
    wSUBJECT = "@SMS_Actif " & wMotif

    Call SMS_MONITOR("@SMS_Actif_O", wSUBJECT, wMotif)
'End If

Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, "< SMS_Actif_Ok....Fin")


End Sub

Public Sub SMS_SWIFT_R()
Dim I As Integer, X As String

Dim vDate, blnSMS As Boolean
Dim wSUBJECT As String, wMotif As String
Call lstErr_Clear(frmElp.lstErr, frmElp.cmdContext, "> SMS_SWIFT_R......")

blnSMS = False
vDate = DateAdd("n", -90, Now)
frmElp.filDoc.Path = paramSAA_Data_to_SAB
frmElp.filDoc.Pattern = "x.xxx"
frmElp.filDoc.Pattern = "*" & paramSAA_Data_to_SAB_ExtensionP_out
For I = 0 To frmElp.filDoc.ListCount - 1
    frmElp.filDoc.ListIndex = I
    Set msFile = msFileSystem.GetFile(frmElp.filDoc.Path & "\" & frmElp.filDoc.FileName)
    If msFile.DateLastModified < vDate Then blnSMS = True
Next I


If blnSMS Then
    wSUBJECT = "@SMS_SWIFT-Réception a détecté des messages entrants dans " & paramSAA_Data_to_SAB

    wMotif = "faire l'import des messages entrants dans SAB"
    Call SMS_MONITOR("@SMS_SWIFT_R", wSUBJECT, wMotif)
End If

Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, "< SMS_SWIFT_R....Fin")


End Sub


Public Sub SMS_Actif()
Dim I As Integer, X As String

Dim vDate, blnSMS As Boolean
Dim wSUBJECT As String, wMotif As String
Call lstErr_Clear(frmElp.lstErr, frmElp.cmdContext, "> SMS_Actif......")

blnSMS = False
vDate = DateAdd("n", -30, Now)
frmElp.filDoc.Path = paramServer("\\PeliNT\" & paramEnvironnement & "\SMS-Monitor\Send\")

frmElp.filDoc.Pattern = "x.xxx"
frmElp.filDoc.Pattern = "*.sms"
For I = 0 To frmElp.filDoc.ListCount - 1
    frmElp.filDoc.ListIndex = I
    Set msFile = msFileSystem.GetFile(frmElp.filDoc.Path & "\" & frmElp.filDoc.FileName)
    If msFile.DateLastModified < vDate Then blnSMS = True
Next I


If blnSMS Then

    X = "<body bgcolor=" & Asc34 & "MAGENTA" & Asc34 & ">" _
                & "<FONT face=" & Asc34 & prtFontName_Comic & Asc34 & ">" _
                & htmlFontColor("BLUE") & "<BR><BR>" & "Des SMS sont en attente depuis plus de 30 minutes dans le répertoire : " & frmElp.filDoc.Path _
                & "<BR><BR>" & "Vérifier l'application \\Pelisrv\SMS"
                
    Call Email_Alerte("ALERTE", "INFO", "SMS_Actif ?", X, True, "")
End If

Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, "< SMS_Actif....Fin")


End Sub

Public Sub SMS_SPLF_FT()
Dim I As Integer, X As String

Dim vDate, blnSMS As Boolean
Dim wSUBJECT As String, wMotif As String
Call lstErr_Clear(frmElp.lstErr, frmElp.cmdContext, "> SMS_SPLF_FT......")

blnSMS = False
vDate = DateAdd("n", -5, Now)
frmElp.filDoc.Path = paramFTP_SPLF
frmElp.filDoc.Pattern = "x.xxx"
frmElp.filDoc.Pattern = "*.*"
For I = 0 To frmElp.filDoc.ListCount - 1
    frmElp.filDoc.ListIndex = I
    Set msFile = msFileSystem.GetFile(frmElp.filDoc.Path & "\" & frmElp.filDoc.FileName)
    If msFile.DateLastModified < vDate Then blnSMS = True
Next I


If blnSMS Then
    wSUBJECT = "@SMS_SPLF_FT a détecté des fichiers en attente dans " & paramFTP_SPLF

    wMotif = "Vérifier sur \\PrintSrv\, si le programme @TIMER_SPLF est actif"
    Call SMS_MONITOR("@SMS_SPLF_FT", wSUBJECT, wMotif)
End If

Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, "< SMS_SPLF_FT....Fin")


End Sub
Public Sub SMS_SPLF_PR()
Dim I As Integer, X As String

Dim vDate, blnSMS As Boolean
Dim wSUBJECT As String, wMotif As String
Call lstErr_Clear(frmElp.lstErr, frmElp.cmdContext, "> SMS_SPLF_PR......")

blnSMS = False
vDate = DateAdd("n", -5, Now)
frmElp.filDoc.Path = paramEditionSplf_Folder & "\Print"
frmElp.filDoc.Pattern = "x.xxx"
frmElp.filDoc.Pattern = "*.*"
For I = 0 To frmElp.filDoc.ListCount - 1
    frmElp.filDoc.ListIndex = I
    Set msFile = msFileSystem.GetFile(frmElp.filDoc.Path & "\" & frmElp.filDoc.FileName)
    If msFile.DateLastModified < vDate Then blnSMS = True
Next I


If blnSMS Then
    wSUBJECT = "@SMS_SPLF_PR a détecté des fichiers en attente dans " & frmElp.filDoc.Path

    wMotif = "Vérifier sur \\PrintSrv\, si le programme @TIMER_SPLF est actif"
    Call SMS_MONITOR("@SMS_SPLF_PR", wSUBJECT, wMotif)
End If

Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, "< SMS_SPLF_PR....Fin")


End Sub

Public Sub SMS_JRN()
Dim V, blnSMS As Boolean
Dim wSUBJECT As String, wMotif As String

Call lstErr_Clear(frmElp.lstErr, frmElp.cmdContext, "> SMS_JRN......")
blnSMS = False
meYBIAMON0.MONAPP = "@AUTO_JRN"
meYBIAMON0.MONFLUX = "MAIL"
V = rsYBIAMON0_Read(meYBIAMON0)
If IsNull(V) Then
    Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, ": SMS_JRN : " & Trim(meYBIAMON0.MONFILE) & meYBIAMON0.MONFILE)
    If Trim(meYBIAMON0.MONSTATUS) = "" Then
       If Trim(meYBIAMON0.MONFILE) <> YBIATAB0_DATE_CPT_J Then blnSMS = True: wSUBJECT = "@AUTO_JRN : OK Date ? " & meYBIAMON0.MONFILE & " <> " & dateImp10(YBIATAB0_DATE_CPT_JS1)
    Else
       blnSMS = True: wSUBJECT = "@AUTO_JRN : statut " & Trim(meYBIAMON0.MONSTATUS) & " : " & meYBIAMON0.MONFILE
    End If
Else
    Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, "? SMS_JRN err : " & V)
     blnSMS = True: wSUBJECT = "SMS_JRN : ERREUR accès iSeries "
     meYBIAMON0.MONSTATUS = "?????"
     meYBIAMON0.MONFILE = "XXXXXXXX"
End If
If blnSMS Then
    wMotif = "'SAB073SPE/YBIAMON7' : '" & meYBIAMON0.MONAPP & "' '" & meYBIAMON0.MONFLUX _
           & "?' > Statut : '" & meYBIAMON0.MONSTATUS & "' '" & meYBIAMON0.MONFILE & "'."
    Call SMS_MONITOR("@SMS_JRN", wSUBJECT, wMotif)
End If

Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, "< SMS_JRN....Fin")

End Sub

Public Sub SMS_YBIAJOUR_OK()
Dim V, blnSMS As Boolean
Dim wSUBJECT As String, wMotif As String

Call lstErr_Clear(frmElp.lstErr, frmElp.cmdContext, "> SMS_YBIAJOUR_OK......")
blnSMS = False
mainSoc_AMJCPT_Load
meYBIAMON0.MONAPP = "SMS"
meYBIAMON0.MONFLUX = "@YBIAJOUR"
V = rsYBIAMON0_Read(meYBIAMON0)

If IsNull(V) Then
    Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, ": SMS_YBIAJOUR_OK : " & Trim(meYBIAMON0.MONFILE) & meYBIAMON0.MONFILE)
    If Trim(meYBIAMON0.MONSTATUS) = "" Then
       If Trim(meYBIAMON0.MONFILE) <> YBIATAB0_DATE_CPT_J Then blnSMS = True: wSUBJECT = "BIAJOUR : OK Date ? " & meYBIAMON0.MONFILE & " <> " & dateImp10(YBIATAB0_DATE_CPT_J)
    Else
       blnSMS = True: wSUBJECT = "BIAJOUR : statut " & Trim(meYBIAMON0.MONSTATUS) & " : " & meYBIAMON0.MONFILE
    End If
Else
    Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, "? SMS_YBIAJOUR_OK err : " & V)
     blnSMS = True: wSUBJECT = "SMS_YBIAJOUR_OK : ERREUR accès iSeries "
     meYBIAMON0.MONSTATUS = "?????"
     meYBIAMON0.MONFILE = "XXXXXXXX"
End If
If blnSMS Then
    wMotif = "'SAB073SPE/YBIAMON7' : '" & meYBIAMON0.MONAPP & "' '" & meYBIAMON0.MONFLUX _
           & "?' > Statut : '" & meYBIAMON0.MONSTATUS & "' '" & meYBIAMON0.MONFILE & "'."
    Call SMS_MONITOR("@YBIAJOUR_OK", wSUBJECT, wMotif)
End If

Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, "< SMS_YBIAJOUR_OK....Fin")


End Sub


Public Sub AUTO_SYSTEM()
Dim meYBIAMON0 As typeYBIAMON0
Dim mailYBIAMON0 As typeYBIAMON0
Dim V
On Error GoTo Mail_Exit

mainSoc_AMJCPT_Load


App_Debug = "> @AUTO_SYSTEM"
Call lstErr_Clear(frmElp.lstErr, frmElp.cmdContext, "> " & App_Debug)

'--------------------------------------------------------------------------------------
'MAIL: traitement YBIAJOUR terminé ? déjà traité ce jour ?
'--------------------------------------------------------------------------------------
meYBIAMON0.MONAPP = "SMS"
meYBIAMON0.MONFLUX = "@YBIAJOUR"
Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, "> " & App_Debug & " " & meYBIAMON0.MONAPP & " " & meYBIAMON0.MONFLUX)
V = rsYBIAMON0_Read(meYBIAMON0)

If IsNull(V) Then
    If Trim(meYBIAMON0.MONSTATUS) <> "" Then V = "BIAJOUR : statut " & Trim(meYBIAMON0.MONSTATUS) & " : " & meYBIAMON0.MONFILE
End If
If Not IsNull(V) Then
    Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, V)
    Exit Sub
End If

mailYBIAMON0.MONAPP = "@AUTO_SYST"
mailYBIAMON0.MONFLUX = "MAIL"
mailYBIAMON0.MONSTATUS = ""
Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, "> " & App_Debug & " " & meYBIAMON0.MONAPP & " " & meYBIAMON0.MONFLUX)
V = fctExploitation_Auto_Control(mailYBIAMON0)
If Not IsNull(V) Then Exit Sub
V = cnSAB_Transaction("Commit")


FLUX_TREPRE:
'--------------------------------------------------------------------------------------
meYBIAMON0.MONAPP = "@AUTO_SYST"
meYBIAMON0.MONFLUX = "@FLUX_TREP"
meYBIAMON0.MONSTATUS = ""
Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, "> " & App_Debug & " " & meYBIAMON0.MONAPP & " " & meYBIAMON0.MONFLUX)

V = fctExploitation_Auto_Control(meYBIAMON0)
If IsNull(V) Then
    V = cnSAB_Transaction("Commit")
    If blnAuto_Form_Show Then frmYFLUTPJ0_Show
    frmYFLUTPJ0.Msg_Rcv "@FLUX_TREPRE"
    V = cnSAB_Transaction("BeginTrans")
    V = fctExploitation_Auto_End(meYBIAMON0)
End If


'--------------------------------------------------------------------------------------
Mail_Exit:
On Error GoTo Error_Handler
mailYBIAMON0.MONAPP = "@AUTO_SYST"
mailYBIAMON0.MONFLUX = "MAIL"
mailYBIAMON0.MONSTATUS = "MONITOR"
Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, "> " & App_Debug & " " & meYBIAMON0.MONAPP & " " & meYBIAMON0.MONFLUX)

V = fctExploitation_Auto_Control(mailYBIAMON0)
If Not IsNull(V) Then Exit Sub

V = fctExploitation_Auto_End(mailYBIAMON0)

''''AUTO_SYSTEM_SendMail

'--------------------------------------------------------------------------------------
Error_Handler:
End Sub

Public Sub SMS_MONITOR(lApplication As String, lSubject As String, lMotif As String)
Static iSeq As Integer
Dim xFileName As String, xPath As String, X As String, X2 As String
Dim V
Dim blnOk As Boolean
Dim intFile As Integer, K As Integer


Dim applicationName As String, applicationMemo As String
Dim destinataireName As String, destinataireMemo As String, destinataireX As String

On Error GoTo Error_Handler
V = rsElpTable_Read("SMS", "Application", lApplication, applicationName, applicationMemo)

X2 = lApplication & " : " & lSubject
X = "<body bgcolor=" & Asc34 & "MAGENTA" & Asc34 & ">" _
            & "<FONT face=" & Asc34 & prtFontName_Comic & Asc34 & ">" _
            & htmlFontColor("BLUE") & "<BR><BR>" & lSubject _
            & "<BR><BR>" & "Destinataires : " & applicationMemo _
            & "<BR><BR>" & lMotif
If lApplication <> "@SMS_Actif_O" Then Call Email_Alerte("ALERTE", "INFO", lSubject, X, True, "")

 
blnOk = False
xPath = paramServer("\\PeliNT\" & paramEnvironnement & "\SMS-Monitor\Send\")
iSeq = iSeq + 1
xFileName = xPath & DSys & "_" & time_Hms & "_" & iSeq & "_" & lApplication

K = 0
If IsNull(V) Then
    Do
       destinataireX = Space_Scan(applicationMemo, K)
       If destinataireX = "" Then Exit Do
       
        Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, ": SMS_MONITOR SEND : " & destinataireX)
        
        X = xFileName & "_" & destinataireX
        If IsNull(V) Then V = File_Export_Monitor("Output", intFile, X & ".txt")
        If IsNull(V) Then V = File_Export_Monitor("Print", intFile, "$SMS-Monitor: IDENT " & Asc34 & destinataireX & Asc34)
        If IsNull(V) Then V = File_Export_Monitor("Print", intFile, "$SMS-Monitor: SUBJECT " & Asc34 & lSubject & Asc34)
        If IsNull(V) Then V = File_Export_Monitor("Close", intFile, xFileName)
        If IsNull(V) Then msFileSystem.MoveFile X & ".txt", X & ".sms"
       
    Loop
    
End If
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnElpTimer_Auto Then MsgBox V, vbCritical, "BIA_SYSTEM" & " : " & "SMS_MONITOR" & lApplication
    Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, "? SMS_MONITOR : " & V)
End Sub


