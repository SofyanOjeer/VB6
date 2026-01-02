Attribute VB_Name = "Bia_swift_Monitor"
Option Explicit


Public Const paramSAA_Queue_TRF_en_Cours = "_TRF_en_Cours"
Public Const paramSAA_Queue_Autorisation = "_MP_authorisation"
Public Const paramSAA_Queue_Modification = "_MP_mod_text"
Public Const paramSAA_Queue_SWIFT = "_SI_to_SWIFT"

Public Sub EcritLog(fonction As String, Description As String, source As String)
Dim I As Long
Dim fic As Long
Dim fichierLog As String

    fichierLog = "c:\temp\XCOM_Log.log"
    fic = FreeFile
    Open fichierLog For Append As #fic
    Print #fic, "Fonction = " & fonction & " Description = " & Description & " Source = " & source & " " & Now
    Close #fic
    
End Sub


Public Sub frmSAB_Dossier_RDE_Show()
Dim X As String

    frmSAB_Dossier_RDE.Icon = frmElp_Icon
    frmSAB_Dossier_RDE.Show vbModeless
    frmSAB_Dossier_RDE.WindowState = vbNormal
    frmSAB_Dossier_RDE.Visible = True
    X = frmSAB_Dossier_RDE.Caption

End Sub

Public Sub mainSocExe()
Dim xName As String, xMemo As String
Dim V

paramIMP_PDFCreator_Name = "PDF_BIA_SWIFT"
paramIMP_PDF_Path_VBP = "C:\Temp\IMP_PDF\BIA_SWIFT"

If Not msFileSystem.FolderExists(paramIMP_PDF_Path_VBP) Then paramIMP_PDF_Path_VBP = paramIMP_PDF_Path_Temp
paramIMP_PDF_Path = paramIMP_PDF_Path_Temp

If App_EXEName = "AIB_SWIFT" Then
    frmElp_Caption = "BIA_SWIFT"
    frmElpPrt.Hide
    Set frmElp_Icon = frmElpPrt
    frmElp.fgMain_App_X.Visible = False
Else
    frmElp_Caption = "BIA_SWIFT"
    Set frmElp_Icon = frmSwift_Messages
End If
blnMonitor = True

If xlsManual Then
    frmElp.fra0.BackColor = &HDCF2F8
End If

End Sub


'---------------------------------------------------------
Public Sub Msg_Monitor(Msg As String)
'---------------------------------------------------------
    
    If Not blnMonitor Then Exit Sub
    
    If xlsManual Then
        If appExcelPublic Is Nothing Then
            Set appExcelPublic = CreateObject("Excel.Application")
            appExcelPublic.Visible = False
            appExcelPublic.ControlCharacters = False
            appExcelPublic.Interactive = False
        End If
    End If
    
    Select Case UCase$(Trim(Mid$(Msg, 1, 12)))
        Case Is = "SAA": frmSAA_Show: frmSAA.Msg_Rcv Msg:
        Case Is = "SWI_MESSAGES": frmSwift_Messages_Show: frmSwift_Messages.Msg_Rcv Msg:
        Case Is = "SWI_OPERATIO": frmSwift_Opération_Show: frmSwift_Opération.Msg_Rcv Msg:
        Case Is = "SWI_STAT_JPL": frmYSWIDOS0_Show: frmYSWIDOS0.Msg_Rcv Msg:
        Case Is = "SWI_STAT": frmSWI_Stat_Show: frmSWI_Stat.Msg_Rcv Msg:
        Case Is = "SAB_DOSSIER": frmSAB_Dossier_Show: frmSAB_Dossier.Msg_Rcv Msg:
        Case Is = "BIA_GOS", "@YSWISAB0", "@YSWIRAM0": frmYGOSDOS0_Show: frmYGOSDOS0.Msg_Rcv Msg:
        Case Is = "BIA_GOS +": frmYGOSDOS0_Annexe_Show: frmYGOSDOS0_Annexe.Msg_Rcv Msg:
        Case Is = "SWAP_TAUX": frmYSWAMON0_Show: frmYSWAMON0.Msg_Rcv Msg:
        Case Is = "@CDO_SCAN": blnAuto_Form_Show = False: frmSAB_Dossier_Show: frmSAB_Dossier.Msg_Rcv Msg:
        Case Is = "@SAB_DOSSIER": blnAuto_Form_Show = False: frmSAB_Dossier_Auto
        Case Is = "@BIA_GOS": blnAuto_Form_Show = False: frmBIA_GOS_Auto
        Case Is = "@SAA_SORTANT"
                                If blnAuto_Form_Show Then frmSwift_Messages_Show
                                frmSwift_Messages.Msg_Rcv Msg
        Case Is = "@AUTO_SAA"
                                If blnAuto_Form_Show Then frmSwift_Messages_Show
                                frmSwift_Messages.Msg_Rcv Msg
         Case Is = "@SAA_ENTRANT"
                                If blnAuto_Form_Show Then frmSAA_Show
                                frmSAA.Msg_Rcv Msg
       Case Is = "@AUTO_SWIOPE":
                                If blnAuto_Form_Show Then frmSwift_Opération_Show
                                frmSwift_Opération.Msg_Rcv Msg
        Case Is = "@SAA_LISTES": mainSoc_AMJCPT_Load
                            If blnAuto_Exploitation_Ok("DATE_CPT_J", "@SAA_LISTES") Then
                                If blnAuto_Form_Show Then frmSwift_Messages_Show
                                frmSwift_Messages.Msg_Rcv Msg
                                Call blnAuto_Exploitation_Ok("Update", "@SAA_LISTES")
                            Else
                                Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, Msg & " déjà traité")
                            End If
    
        Case Is = "X_RESET":  main_Reset
        Case Is = "XUSRID": XUsrId_Show
        Case Is = "X_I5A7": X_I5A7_Show
    End Select
End Sub

Public Sub frmSwift_Messages_Show()
Dim X As String

frmSwift_Messages.Icon = frmElp_Icon
frmSwift_Messages.Show vbModeless
frmSwift_Messages.WindowState = vbNormal
frmSwift_Messages.Visible = True
X = frmSwift_Messages.Caption
AppActivate X

End Sub
Public Sub frmSAA_Show()
Dim X As String

frmSAA.Icon = frmElp_Icon
frmSAA.Show vbModeless
frmSAA.WindowState = vbNormal
frmSAA.Visible = True
X = frmSAA.Caption
AppActivate X

End Sub

Public Sub frmSwift_Opération_Show()
Dim X As String

frmSwift_Opération.Icon = frmElp_Icon
frmSwift_Opération.Show vbModeless
frmSwift_Opération.WindowState = vbNormal
frmSwift_Opération.Visible = True
X = frmSwift_Opération.Caption
AppActivate X

End Sub

Public Sub frmYSWIDOS0_Show()
Dim X As String

frmYSWIDOS0.Icon = frmElp_Icon
frmYSWIDOS0.Show vbModeless
frmYSWIDOS0.WindowState = vbNormal
frmYSWIDOS0.Visible = True
X = frmYSWIDOS0.Caption
AppActivate X

End Sub
Public Sub frmSWI_Stat_Show()
Dim X As String

frmSWI_Stat.Icon = frmElp_Icon
frmSWI_Stat.Show vbModeless
frmSWI_Stat.WindowState = vbNormal
frmSWI_Stat.Visible = True
X = frmSWI_Stat.Caption
'AppActivate X

End Sub

Public Sub frmYGOSDOS0_Show()
Dim X As String

frmYGOSDOS0.Icon = frmElp_Icon
frmYGOSDOS0.Show vbModeless
frmYGOSDOS0.WindowState = vbNormal
frmYGOSDOS0.Visible = True
frmYGOSDOS0.BackColor = frmElp.BackColor
X = frmYGOSDOS0.Caption
'AppActivate X

End Sub
Public Sub frmYSWAMON0_Show()
Dim X As String

frmYSWAMON0.Icon = frmElp_Icon
frmYSWAMON0.Show vbModeless
frmYSWAMON0.WindowState = vbNormal
frmYSWAMON0.Visible = True
frmYSWAMON0.BackColor = frmElp.BackColor
X = frmYSWAMON0.Caption
'AppActivate X

End Sub

Public Sub frmYGOSDOS0_Annexe_Show()
Dim X As String

frmYGOSDOS0_Annexe.Icon = frmElp_Icon
frmYGOSDOS0_Annexe.Show vbModeless
frmYGOSDOS0_Annexe.WindowState = vbNormal
frmYGOSDOS0_Annexe.Visible = True
X = frmYGOSDOS0_Annexe.Caption
'AppActivate X

End Sub
Public Sub frmSIDE_DB_Show()
Dim X As String

frmSIDE_DB.Icon = frmElp_Icon
frmSIDE_DB.Show vbModeless
frmSIDE_DB.WindowState = vbNormal
frmSIDE_DB.Visible = True
X = frmSIDE_DB.Caption
'AppActivate X

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
Public Sub frmSAB_Dossier_CDO_Show()
Dim X As String

frmSAB_Dossier_CDO.Icon = frmElp_Icon
frmSAB_Dossier_CDO.Show vbModeless
frmSAB_Dossier_CDO.WindowState = vbNormal
frmSAB_Dossier_CDO.Visible = True
X = frmSAB_Dossier_CDO.Caption
'AppActivate X

End Sub

Public Sub frmSAB_Dossier_Show()
Dim X As String
On Error Resume Next
frmSAB_Dossier.Icon = frmElp_Icon
frmSAB_Dossier.Show vbModeless
frmSAB_Dossier.WindowState = vbNormal
frmSAB_Dossier.Visible = True
X = frmSAB_Dossier.Caption
AppActivate X

End Sub
Public Sub frmSAB_Dossier_Auto()
Dim meYBIAMON0 As typeYBIAMON0, V
'--------------------------------------------------------------------------------------
V = COMPTA_YBIAJOUR_OK
If Not IsNull(V) Then
    Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, "@SAB_Dossier " & V)
    Exit Sub
End If


meYBIAMON0.MONAPP = "@BIA_SWIFT"
meYBIAMON0.MONFLUX = "@SAB_DOSSIER"
meYBIAMON0.MONSTATUS = ""
Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, "> " & App_Debug & " " & meYBIAMON0.MONAPP & " " & meYBIAMON0.MONFLUX)

V = fctExploitation_Auto_Control(meYBIAMON0)
If IsNull(V) Then
    Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, "> début du traitement ....")
    V = cnSAB_Transaction("Commit")
    If blnAuto_Form_Show Then frmSAB_Dossier_Show
    frmSAB_Dossier.Msg_Rcv "@SAB_DOSSIER"
    V = cnSAB_Transaction("BeginTrans")
    V = fctExploitation_Auto_End(meYBIAMON0)
Else
    Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, "> " & V)

End If

End Sub

Public Sub frmBIA_GOS_Auto()
Dim meYBIAMON0 As typeYBIAMON0, V
'--------------------------------------------------------------------------------------

meYBIAMON0.MONAPP = "@BIA_SWIFT"
meYBIAMON0.MONFLUX = "@BIA_GOS"
meYBIAMON0.MONSTATUS = ""
Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, "> " & App_Debug & " " & meYBIAMON0.MONAPP & " " & meYBIAMON0.MONFLUX)

V = fctExploitation_Auto_Control(meYBIAMON0)
If IsNull(V) Then
    Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, "> début du traitement ....")
    V = cnSAB_Transaction("Commit")
    If blnAuto_Form_Show Then frmYGOSDOS0_Show
    frmYGOSDOS0.Msg_Rcv "@BIA_GOS"
    V = cnSAB_Transaction("BeginTrans")
    V = fctExploitation_Auto_End(meYBIAMON0)
Else
    Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, "> " & V)

End If

End Sub

Public Sub SAA_LISTES_2008()
'Dim fic As Long
'Dim strChaine As String
'Dim oK As Boolean
'Static enCours As Boolean
'
'    If Not enCours Then
'        enCours = True
'        frmElp.Timer1.Enabled = False
'        frmElp.Timer1.Interval = 0
'        fic = FreeFile
'        Open "c:\temp\imp_pdf\Bia_Audit2008.log" For Input As #fic
'        oK = False
'        Do Until EOF(fic)
'            Line Input #fic, strChaine
'            If InStr(strChaine, "Fin") > -1 Then
'                oK = True
'            End If
'        Loop
'        Close #fic
'        If oK = False Then
'            frmElp.Timer1.Enabled = True
'            frmElp.Timer1.Interval = 1000
'            enCours = False
'        Else
'            fic = FreeFile
'            Open "c:\temp\imp_pdf\Bia_Swift2008.log" For Output As #fic
'            Print #fic, "Début SAA_LISTES_2008 --> " & CDate(Now)
'            Close #fic
'            Call frmSwift_Messages.Msg_Rcv("@SAA_LISTES")
'            appExcelPublic.Quit
'            Set appExcelPublic = Nothing
'            fic = FreeFile
'            Open "c:\temp\imp_pdf\Bia_Swift2008.log" For Append As #fic
'            Print #fic, "Fin SAA_LISTES_2008 --> " & CDate(Now)
'            Close #fic
'            End
'        End If
'    End If

End Sub


