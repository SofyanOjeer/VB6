Attribute VB_Name = "Bia_Sab_Monitor"
Option Explicit

Public Sub AUTO_COMPTA_ECRIT_LOG(sujet As String)
Dim fic As Long
Dim bureau As String

bureau = ObtenirCheminBureau
fic = FreeFile
Open bureau & "\AUTO_COMPTA.LOG" For Append As #fic
Print #fic, sujet & " --> " & Now
Close #fic

End Sub

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




Public Sub AUTO_COMPTA_2008()
Dim fic As Long
Static enCours As Boolean

    If Not enCours Then
        enCours = True
        frmElp.Timer1.Enabled = False
        frmElp.Timer1.Interval = 0
        fic = FreeFile
        Open "c:\temp\imp_pdf\Bia_Sab2008.log" For Output As #fic
        Print #fic, "Début AUTO_COMPTA_2008 --> " & CDate(Now)
        Close #fic
        Call frmBIA_Gafi.Msg_Rcv("@BIA_GAFI")
        Call frmBIA_PDC.Msg_Rcv("@BIA_PDC")
        Call frmSAB_TC_Limites.Msg_Rcv("@TC_LIMITES")
        Call frmSAB_Balance.Msg_Rcv("@BAL_6000")
        Call frmSAB_Balance.Msg_Rcv("@BAL_B/HB")
        Call frmSAB_Balance.Msg_Rcv("@BAL_PCI_DC")
        Call frmSAB_Balance.Msg_Rcv("@BAL_STOCK")
        Call frmSAB_Compta.Msg_Rcv("@SOLDEJ")
        Call frmSAB_Compta.Msg_Rcv("@JOURNAL_D")
        Call frmSAB_Compta.Msg_Rcv("@JOURNAL_S")
        Call frmSAB_Stock.Msg_Rcv("@SAB_STOCK")
        Call frmYEICGCC0_ATHIC.Msg_Rcv("@EIC_GCC")
        Call frmYICCCPT0.Msg_Rcv("@ICC_MVT")
        appExcelPublic.Quit
        Set appExcelPublic = Nothing
        fic = FreeFile
        Open "c:\temp\imp_pdf\Bia_Sab2008.log" For Append As #fic
        Print #fic, "Fin AUTO_COMPTA_2008 --> " & CDate(Now)
        Close #fic
        frmElp.Timer1.Enabled = True
        frmElp.Timer1.Interval = 1000
        enCours = False
    End If
    
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
Public Sub frmYNOTPAY0_Show()
Dim X As String

frmYNOTPAY0.Icon = frmElp_Icon
frmYNOTPAY0.Show vbModeless
frmYNOTPAY0.WindowState = vbNormal
frmYNOTPAY0.Visible = True
X = frmYNOTPAY0.Caption
AppActivate X

End Sub
Public Sub frmYCLISCO0_Show()
Dim X As String

frmYCLISCO0.Icon = frmElp_Icon
frmYCLISCO0.Show vbModeless
frmYCLISCO0.WindowState = vbNormal
frmYCLISCO0.Visible = True
X = frmYCLISCO0.Caption
'AppActivate X

End Sub
Public Sub AUTO_COMPTA_INIT_LOG()
Dim fic As Long
Dim bureau As String

bureau = ObtenirCheminBureau
fic = FreeFile
If Dir("%desktop%\AUTO_COMPTA.LOG") <> "" Then
    Kill bureau & "\AUTO_COMPTA.LOG"
End If
Open bureau & "\AUTO_COMPTA.LOG" For Output As #fic
Print #fic, "Initialisation " & Now

Close #fic
End Sub
Public Function ObtenirCheminBureau() As String
'par: Excel-Malin.com ( https://excel-malin.com )

    On Error GoTo ObtenirCheminBureauError
    Dim CheminBureau As String
    CheminBureau = ""
    Dim oWSHShell As Object
    Set oWSHShell = CreateObject("WScript.Shell")
    
    CheminBureau = oWSHShell.SpecialFolders("Desktop")
    
    If (Not (oWSHShell Is Nothing)) Then Set oWSHShell = Nothing
    ObtenirCheminBureau = CheminBureau

    Exit Function
ObtenirCheminBureauError:
    If (Not (oWSHShell Is Nothing)) Then Set oWSHShell = Nothing
    ObtenirCheminBureau = ""
End Function

Public Sub mainSocExe()

paramIMP_PDFCreator_Name = "PDF_BIA_SAB"
paramIMP_PDF_Path_VBP = "C:\Temp\IMP_PDF\BIA_SAB"

If Not msFileSystem.FolderExists(paramIMP_PDF_Path_VBP) Then paramIMP_PDF_Path_VBP = paramIMP_PDF_Path_Temp
paramIMP_PDF_Path = paramIMP_PDF_Path_Temp

If App_EXEName = "AIB_SAB" Then
        frmElp_Caption = "BIA_SAB"
        frmElpPrt.Hide
        Set frmElp_Icon = frmElpPrt
        frmElp.fgMain_App_X.Visible = False
Else
    frmElp_Caption = "BIA_SAB"
    Set frmElp_Icon = frmSAB_Balance
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
'                       '
If xlsManual Then
    If appExcelPublic Is Nothing Then
        Set appExcelPublic = CreateObject("Excel.Application")
        appExcelPublic.Visible = False
        appExcelPublic.ControlCharacters = False
        appExcelPublic.Interactive = False
    End If
End If
        
Select Case UCase$(Trim(Mid$(Msg, 1, 12)))
    Case Is = "AUTOMATE": frmAutomate_Show: frmAutomate.Msg_Rcv Msg
    Case Is = "BIA_GAFI": frmBIA_Gafi_Show: frmBIA_Gafi.Msg_Rcv Msg
    Case Is = "BIA_CLIPRO": frmBIA_CLIPRO_Show
    Case Is = "BIA_EICGCC": frmBIA_EICGCC_Show
    Case Is = "BIA_CLISTA": frmBIA_CLISTA_Show
    Case Is = "BIA_GUIMAD", "@AUTO_GUIMAD": frmBIA_GUIMAD_Show: frmBIA_GUIMAD.Msg_Rcv Msg
    Case Is = "BIA_PDC", "@BIA_PDC": frmBIA_PDC_Show: frmBIA_PDC.Msg_Rcv Msg
    Case Is = "BIA_QUID", "@RMA_CTL": frmBIA_Quid_Show: frmBIA_Quid.Msg_Rcv Msg
    Case Is = "BIA_TVAFAC", "@AUTO_TVAFAC": frmBIA_TVAFAC_Show: frmBIA_TVAFAC.Msg_Rcv Msg
    Case Is = "@BIA_IMPAYÉS": AUTO_BIA_IMPAYES
    Case Is = "BIA_IMPAYÉS": frmBIA_Impayés_Show: frmBIA_Impayés.Msg_Rcv Msg
    Case Is = "CHQ_SCAN":
            frmBIA_ATHIC_Show
            frmBIA_ATHIC.Msg_Rcv Msg
            

    Case Is = "EIC_GCC", "@EIC_GCC"
            frmYEICGCC0_ATHIC_Show
            frmYEICGCC0_ATHIC.Msg_Rcv Msg
    Case Is = "EDITION": frmEdition_Show: frmEdition.Msg_Rcv Msg:
    Case Is = "ICC_MVT", "@ICC_MVT": frmYICCCPT0_Show: frmYICCCPT0.Msg_Rcv Msg
    Case Is = "NOTATION_PAY": frmYNOTPAY0_Show: frmYNOTPAY0.Msg_Rcv Msg
    Case Is = "PAIESAGE": frmPAIESAGE_Show: frmPaieSage.Msg_Rcv Msg
    Case Is = "SAB_BALANCE": frmSAB_Balance_Show: frmSAB_Balance.Msg_Rcv Msg
    Case Is = "SAB_CDO": frmSAB_CDO_Show: frmSAB_CDO.Msg_Rcv Msg
    Case Is = "SAB_CLIENT", "@SAB_CLIENT": frmSAB_CLIENT_Show: frmSAB_Client.Msg_Rcv Msg
    Case Is = "SAB_CRE": frmSAB_CRE_Show: frmSAB_CRE.Msg_Rcv Msg
    Case Is = "SAB_DAT": frmSAB_DAT_Show: frmSAB_DAT.Msg_Rcv Msg
    Case Is = "SAB_COMPTA": frmSAB_COMPTA_Show: frmSAB_Compta.Msg_Rcv Msg
    Case Is = "SAB_CPTMVT": frmSAB_CPTMVT_Show: frmSAB_CPTMVT.Msg_Rcv Msg
    Case Is = "SAB_ECHELLES": frmSAB_Echelles_Show: frmSAB_Echelles.Msg_Rcv Msg
    Case Is = "SAB_FCI": frmSAB_FCI_Show: frmSAB_FCI.Msg_Rcv Msg
    Case Is = "SAB_STOCK": frmSAB_Stock_Show: frmSAB_Stock.Msg_Rcv Msg
    Case Is = "SAB_TAUX", "@SAB_TAUX": frmSAB_TAU_Show: frmSAB_TAU.Msg_Rcv Msg
    Case Is = "SAB_TC_LIMIT": frmSAB_TC_Limites_Show: frmSAB_TC_Limites.Msg_Rcv Msg
    Case Is = "SPLFJOB": frmSPLFJOB_Show: frmSPLFJOB.Msg_Rcv Msg
    Case Is = "SCORING_CLI": frmYCLISCO0_Show: frmYCLISCO0.Msg_Rcv Msg

    Case Is = "X_DOC": frmElpKm_Show: frmElpKM.Msg_Rcv Msg
    Case Is = "DROPI": frmDROPI_Show: frmDROPI.Msg_Rcv Msg
    Case Is = "X_RESET":  main_Reset
    Case Is = "XUSRID": XUsrId_Show
    Case Is = "@AUTO_SPLF", "@AUTO_CLROUT": frmSPLFJOB.Msg_Rcv Msg
    Case Is = "@PRINT_TEST", "@PRINT_PROD": frmEdition.Msg_Rcv Msg
    Case Is = "AUTOMATE", "*=>BDF_CB", "*=>BDF_BDP": frmAutomate_Show: frmAutomate.Msg_Rcv Msg
    Case Is = "@AUTOMATE":  frmAutomate.Msg_Rcv Msg
    Case Is = "@CHQ_DEON":
            frmBIA_ATHIC_Show
            frmBIA_ATHIC.Msg_Rcv Msg
    Case Is = "@RCOM_AUT", "@CPT_OD": frmSAB_Balance_Show: frmSAB_Balance.Msg_Rcv Msg
    Case Is = "@ENG_BEA_LFB": AUTO_ENG_BEA_LFB
    Case Is = "@AUTO_COMPTA": AUTO_COMPTA
    Case Is = "X_I5A7": X_I5A7_Show
    Case Is = "TEST":
End Select

End Sub


Public Sub frmElpKm_Show()
Dim X As String
On Error Resume Next
frmElpKM.Icon = frmElp_Icon
frmElpKM.Show vbModeless
frmElpKM.WindowState = vbNormal
frmElpKM.Visible = True
X = frmElpKM.Caption
AppActivate X

End Sub
Public Sub frmDROPI_Show()
Dim X As String
On Error Resume Next
frmDROPI.Icon = frmElp_Icon
frmDROPI.Show vbModeless
frmDROPI.WindowState = vbNormal
frmDROPI.Visible = True
X = frmDROPI.Caption
AppActivate X

End Sub


'Public Sub frmYEICGCC0_Show()
'Dim x As String
'On Error Resume Next
'frmYEICGCC0.Icon = frmElp_Icon
'frmYEICGCC0.Show vbModeless
'frmYEICGCC0.WindowState = vbNormal
'frmYEICGCC0.Visible = True
'x = frmYEICGCC0.Caption
'AppActivate x

'End Sub
Public Sub frmYEICGCC0_ATHIC_Show()
Dim X As String
On Error Resume Next
frmYEICGCC0_ATHIC.Icon = frmElp_Icon
frmYEICGCC0_ATHIC.Show vbModeless
frmYEICGCC0_ATHIC.WindowState = vbNormal
frmYEICGCC0_ATHIC.Visible = True
X = frmYEICGCC0_ATHIC.Caption
AppActivate X

End Sub


Public Sub frmYICCCPT0_Show()
Dim X As String
On Error Resume Next
frmYICCCPT0.Icon = frmElp_Icon
frmYICCCPT0.Show vbModeless
frmYICCCPT0.WindowState = vbNormal
frmYICCCPT0.Visible = True
X = frmYICCCPT0.Caption
AppActivate X

End Sub
Public Sub frmEdition_Show()
Dim X As String
On Error Resume Next
frmEdition.Icon = frmElp_Icon
frmEdition.Show vbModeless
frmEdition.WindowState = vbNormal
frmEdition.Visible = True
frmEdition.BackColor = frmElp.BackColor
X = frmEdition.Caption
AppActivate X

End Sub

Public Sub frmAutomate_Show()
Dim X As String
On Error Resume Next
frmAutomate.Icon = frmElp_Icon
frmAutomate.Show vbModeless
frmAutomate.WindowState = vbNormal
frmAutomate.Visible = True
X = frmAutomate.Caption
AppActivate X

End Sub


Public Sub frmPAIESAGE_Show()
Dim X As String
On Error Resume Next
frmPaieSage.Icon = frmElp_Icon
frmPaieSage.Show vbModeless
frmPaieSage.WindowState = vbNormal
frmPaieSage.Visible = True
X = frmPaieSage.Caption
AppActivate X

End Sub


Public Sub frmSPLFJOB_Show()
Dim X As String
On Error Resume Next
frmSPLFJOB.Icon = frmElp_Icon
frmSPLFJOB.Show vbModeless
frmSPLFJOB.WindowState = vbNormal
frmSPLFJOB.Visible = True
X = frmSPLFJOB.Caption
AppActivate X

End Sub


Public Sub frmSAB_Balance_Show()
Dim X As String
On Error Resume Next
frmSAB_Balance.Icon = frmElp_Icon
frmSAB_Balance.Show vbModeless
frmSAB_Balance.WindowState = vbNormal
frmSAB_Balance.Visible = True
X = frmSAB_Balance.Caption
AppActivate X

End Sub

Public Sub frmSAB_CDO_Show()
Dim X As String
On Error Resume Next
frmSAB_CDO.Icon = frmElp_Icon
frmSAB_CDO.Show vbModeless
frmSAB_CDO.WindowState = vbNormal
frmSAB_CDO.Visible = True
X = frmSAB_CDO.Caption
AppActivate X

End Sub
Public Sub frmSAB_FCI_Show()
Dim X As String
On Error Resume Next
frmSAB_FCI.Icon = frmElp_Icon
frmSAB_FCI.Show vbModeless
frmSAB_FCI.WindowState = vbNormal
frmSAB_FCI.Visible = True
X = frmSAB_FCI.Caption
AppActivate X

End Sub

Public Sub frmSAB_TAU_Show()
Dim X As String
On Error Resume Next
frmSAB_TAU.Icon = frmElp_Icon
frmSAB_TAU.Show vbModeless
frmSAB_TAU.WindowState = vbNormal
frmSAB_TAU.Visible = True
X = frmSAB_TAU.Caption
AppActivate X

End Sub
Public Sub frmSAB_CRE_Show()
Dim X As String
On Error Resume Next
frmSAB_CRE.Icon = frmElp_Icon
frmSAB_CRE.Show vbModeless
frmSAB_CRE.WindowState = vbNormal
frmSAB_CRE.Visible = True
X = frmSAB_CRE.Caption
AppActivate X

End Sub

Public Sub frmSAB_DAT_Show()
Dim X As String
On Error Resume Next
frmSAB_DAT.Icon = frmElp_Icon
frmSAB_DAT.Show vbModeless
frmSAB_DAT.WindowState = vbNormal
frmSAB_DAT.Visible = True
X = frmSAB_DAT.Caption
AppActivate X

End Sub

Public Sub frmSAB_CLIENT_Show()
Dim X As String
On Error Resume Next
frmSAB_Client.Icon = frmElp_Icon
frmSAB_Client.Show vbModeless
frmSAB_Client.WindowState = vbNormal
frmSAB_Client.Visible = True
X = frmSAB_Client.Caption
AppActivate X

End Sub

Public Sub frmBIA_Gafi_Show()
Dim X As String
On Error Resume Next
frmBIA_Gafi.Icon = frmElp_Icon
frmBIA_Gafi.Show vbModeless
frmBIA_Gafi.WindowState = vbNormal
frmBIA_Gafi.Visible = True
X = frmBIA_Gafi.Caption
AppActivate X

End Sub

Public Sub frmBIA_PDC_Show()
Dim X As String
On Error Resume Next

frmBIA_PDC.Icon = frmElp_Icon
frmBIA_PDC.Show vbModeless
frmBIA_PDC.WindowState = vbNormal
frmBIA_PDC.Visible = True
frmBIA_PDC.BackColor = frmElp.BackColor
X = frmBIA_PDC.Caption
AppActivate X

End Sub
Public Sub frmSAB_Echelles_Show()
Dim X As String
On Error Resume Next
frmSAB_Echelles.Icon = frmElp_Icon
frmSAB_Echelles.Show vbModeless
frmSAB_Echelles.WindowState = vbNormal
frmSAB_Echelles.Visible = True
X = frmSAB_Echelles.Caption
AppActivate X

End Sub

Public Sub frmBIA_GUIMAD_Show()
Dim X As String
On Error Resume Next
frmBIA_GUIMAD.Icon = frmElp_Icon
frmBIA_GUIMAD.Show vbModeless
frmBIA_GUIMAD.WindowState = vbNormal
frmBIA_GUIMAD.Visible = True
X = frmBIA_GUIMAD.Caption
AppActivate X

End Sub
Public Sub frmBIA_TVAFAC_Show()
Dim X As String
On Error Resume Next
frmBIA_TVAFAC.Icon = frmElp_Icon
frmBIA_TVAFAC.Show vbModeless
frmBIA_TVAFAC.WindowState = vbNormal
frmBIA_TVAFAC.Visible = True
frmBIA_TVAFAC.BackColor = frmElp.BackColor
X = frmBIA_TVAFAC.Caption
AppActivate X

End Sub

Public Sub frmBIA_Impayés_Show()
Dim X As String
On Error Resume Next
frmBIA_Impayés.Icon = frmElp_Icon
frmBIA_Impayés.Show vbModeless
frmBIA_Impayés.WindowState = vbNormal
frmBIA_Impayés.Visible = True
X = frmBIA_Impayés.Caption
AppActivate X

End Sub


'Public Sub frmCHQ_SCAN_Show()
'Dim x As String
'On Error Resume Next


'frmCHQ_SCAN.Icon = frmElp_Icon
'frmCHQ_SCAN.Show vbModeless
'frmCHQ_SCAN.WindowState = vbNormal
'frmCHQ_SCAN.Visible = True
'x = frmCHQ_SCAN.Caption

'AppActivate x

'End Sub

Public Sub frmBIA_ATHIC_Show()
Dim X As String
On Error Resume Next
frmBIA_ATHIC.Icon = frmElp_Icon
frmBIA_ATHIC.Show vbModeless
frmBIA_ATHIC.WindowState = vbNormal
frmBIA_ATHIC.Visible = True
X = frmBIA_ATHIC.Caption


AppActivate X

End Sub


Public Sub frmSAB_COMPTA_Show()
Dim X As String
On Error Resume Next
frmSAB_Compta.Icon = frmElp_Icon
frmSAB_Compta.Show vbModeless
frmSAB_Compta.WindowState = vbNormal
frmSAB_Compta.Visible = True
X = frmSAB_Compta.Caption
AppActivate X

End Sub

Public Sub frmSAB_Stock_Show()
Dim X As String
On Error Resume Next
frmSAB_Stock.Icon = frmElp_Icon
frmSAB_Stock.Show vbModeless
frmSAB_Stock.WindowState = vbNormal
frmSAB_Stock.Visible = True
X = frmSAB_Stock.Caption
AppActivate X

End Sub
Public Sub frmSAB_CPTMVT_Show()
Dim X As String
On Error Resume Next
frmSAB_CPTMVT.Icon = frmElp_Icon
frmSAB_CPTMVT.Show vbModeless
frmSAB_CPTMVT.WindowState = vbNormal
frmSAB_CPTMVT.Visible = True
X = frmSAB_CPTMVT.Caption
AppActivate X

End Sub


Public Sub frmSAB_TC_Limites_Show()
Dim X As String
On Error Resume Next
frmSAB_TC_Limites.Icon = frmElp_Icon
frmSAB_TC_Limites.Show vbModeless
frmSAB_TC_Limites.WindowState = vbNormal
frmSAB_TC_Limites.Visible = True
X = frmSAB_TC_Limites.Caption
AppActivate X

End Sub


Public Sub AUTO_COMPTA()
Dim meYBIAMON0 As typeYBIAMON0
Dim mailYBIAMON0 As typeYBIAMON0
Dim V
Dim okCOMPTA As Boolean

On Error GoTo Mail_Exit

mainSoc_AMJCPT_Load

App_Debug = "> @AUTO_COMPTA"
Call lstErr_Clear(frmElp.lstErr, frmElp.cmdContext, "> " & App_Debug)
Call AUTO_COMPTA_INIT_LOG

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

'Ecriture du flag de début de AUTO_COMPTA
If xlsManual Then
    okCOMPTA = False
    mailYBIAMON0.MONAPP = "COMPTA"
    mailYBIAMON0.MONFLUX = "MAIL"
    mailYBIAMON0.MONSTATUS = ""
    V = rsYBIAMON0_Read(mailYBIAMON0)
    If IsNull(V) Then
        If Trim(mailYBIAMON0.MONFILE) >= YBIATAB0_DATE_CPT_J Then
            okCOMPTA = True
        End If
    End If
End If
'                                       '
Call AUTO_COMPTA_ECRIT_LOG("1/27 = MAIL")
mailYBIAMON0.MONAPP = "COMPTA"
mailYBIAMON0.MONFLUX = "MAIL"
mailYBIAMON0.MONSTATUS = ""
Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, "> " & App_Debug & " " & meYBIAMON0.MONAPP & " " & meYBIAMON0.MONFLUX)
V = fctExploitation_Auto_Control(mailYBIAMON0)
If Not IsNull(V) Then Exit Sub
V = cnSAB_Transaction("Commit")

SOLDEJ:
'--------------------------------------------------------------------------------------
Call AUTO_COMPTA_ECRIT_LOG("2/27 = SOLDEJ")
meYBIAMON0.MONAPP = "COMPTA"
meYBIAMON0.MONFLUX = "SOLDEJ"
meYBIAMON0.MONSTATUS = ""
Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, "> " & App_Debug & " " & meYBIAMON0.MONAPP & " " & meYBIAMON0.MONFLUX)

V = fctExploitation_Auto_Control(meYBIAMON0)
If IsNull(V) Then
    If blnAuto_Form_Show Then frmSAB_Balance_Show
    frmSAB_Compta.Msg_Rcv "@SOLDEJ"
    V = fctExploitation_Auto_End(meYBIAMON0)
End If

BAL_B_HB:
'--------------------------------------------------------------------------------------
Call AUTO_COMPTA_ECRIT_LOG("3/27 = BAL_B/HB")
meYBIAMON0.MONAPP = "COMPTA"
meYBIAMON0.MONFLUX = "BAL_B/HB"
meYBIAMON0.MONSTATUS = ""
Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, "> " & App_Debug & " " & meYBIAMON0.MONAPP & " " & meYBIAMON0.MONFLUX)

V = fctExploitation_Auto_Control(meYBIAMON0)
If IsNull(V) Then
    If blnAuto_Form_Show Then frmSAB_Balance_Show
    frmSAB_Balance.Msg_Rcv "@BAL_B/HB"
    V = fctExploitation_Auto_End(meYBIAMON0)

End If

JOURNAL_D:
'--------------------------------------------------------------------------------------
Call AUTO_COMPTA_ECRIT_LOG("4/27 = JOURNAL_D")
meYBIAMON0.MONAPP = "COMPTA"
meYBIAMON0.MONFLUX = "JOURNAL_D"
meYBIAMON0.MONSTATUS = ""
Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, "> " & App_Debug & " " & meYBIAMON0.MONAPP & " " & meYBIAMON0.MONFLUX)

V = fctExploitation_Auto_Control(meYBIAMON0)
If IsNull(V) Then
    If blnAuto_Form_Show Then frmSAB_COMPTA_Show
    frmSAB_Compta.Msg_Rcv "@JOURNAL_D"
    V = fctExploitation_Auto_End(meYBIAMON0)
End If

JOURNAL_S:
'--------------------------------------------------------------------------------------
Call AUTO_COMPTA_ECRIT_LOG("5/27 = JOURNAL_S")
meYBIAMON0.MONAPP = "COMPTA"
meYBIAMON0.MONFLUX = "JOURNAL_S"
meYBIAMON0.MONSTATUS = ""
V = fctExploitation_Auto_Control(meYBIAMON0)
If IsNull(V) Then
Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, "> " & App_Debug & " " & meYBIAMON0.MONAPP & " " & meYBIAMON0.MONFLUX)

    If blnAuto_Form_Show Then frmSAB_COMPTA_Show
    frmSAB_Compta.Msg_Rcv "@JOURNAL_S"
    V = fctExploitation_Auto_End(meYBIAMON0)
End If

MT900:
'--------------------------------------------------------------------------------------
Call AUTO_COMPTA_ECRIT_LOG("6/27 = MT900")
meYBIAMON0.MONAPP = "COMPTA"
meYBIAMON0.MONFLUX = "MT900"
Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, "> " & App_Debug & " " & meYBIAMON0.MONAPP & " " & meYBIAMON0.MONFLUX)
V = rsYBIAMON0_Read(meYBIAMON0)
If IsNull(V) Then
    If Trim(meYBIAMON0.MONSTATUS) = "" Then
        If Trim(meYBIAMON0.MONFILE) < YBIATAB0_DATE_CPT_J Then

            If blnAuto_Form_Show Then frmSAB_COMPTA_Show
            frmSAB_Compta.Msg_Rcv "@MT900"
        End If
    End If
End If

MT950:
'--------------------------------------------------------------------------------------
Call AUTO_COMPTA_ECRIT_LOG("7/27 = MT950")
meYBIAMON0.MONAPP = "COMPTA"
meYBIAMON0.MONFLUX = "MT950"
Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, "> " & App_Debug & " " & meYBIAMON0.MONAPP & " " & meYBIAMON0.MONFLUX)
V = rsYBIAMON0_Read(meYBIAMON0)
If IsNull(V) Then
    If Trim(meYBIAMON0.MONSTATUS) = "" Then
        If Trim(meYBIAMON0.MONFILE) < YBIATAB0_DATE_CPT_J Then
            If blnAuto_Form_Show Then frmSAB_COMPTA_Show
        frmSAB_Compta.Msg_Rcv "@MT950"
        End If
    End If
End If

BIA_GAFI:
'--------------------------------------------------------------------------------------
Call AUTO_COMPTA_ECRIT_LOG("8/27 = BIA_GAFI")
meYBIAMON0.MONAPP = "COMPTA"
meYBIAMON0.MONFLUX = "BIA_GAFI"
meYBIAMON0.MONSTATUS = ""
Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, "> " & App_Debug & " " & meYBIAMON0.MONAPP & " " & meYBIAMON0.MONFLUX)

V = fctExploitation_Auto_Control(meYBIAMON0)
If IsNull(V) Then
    If blnAuto_Form_Show Then frmBIA_Gafi_Show
    frmBIA_Gafi.Msg_Rcv "@BIA_GAFI"
    V = fctExploitation_Auto_End(meYBIAMON0)
End If

CHQ_DEON:
'--------------------------------------------------------------------------------------
Call AUTO_COMPTA_ECRIT_LOG("9/27 = AUTO_FCI")
AUTO_FCI
'--------------------------------------------------------------------------------------
Call AUTO_COMPTA_ECRIT_LOG("10/27 = BAL_PCI_DC")
AUTO_BAL_PCI_DC
'--------------------------------------------------------------------------------------
Call AUTO_COMPTA_ECRIT_LOG("11/27 = AUTO_GUIMA")
AUTO_GUIMAD
'--------------------------------------------------------------------------------------
frmElp.MousePointer = vbNoDrop
Wait_SS 30
frmElp.MousePointer = vbNormal
Call AUTO_COMPTA_ECRIT_LOG("12/27 = AUTO_TVAFA")
AUTO_TVAFAC
'--------------------------------------------------------------------------------------
'$JPL 2014-10-13 AUTO_ROPDOS
'--------------------------------------------------------------------------------------
Call AUTO_COMPTA_ECRIT_LOG("13/27 = CHQ_DEON")
meYBIAMON0.MONAPP = "COMPTA"
meYBIAMON0.MONFLUX = "CHQ_DEON"
Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, "> " & App_Debug & " " & meYBIAMON0.MONAPP & " " & meYBIAMON0.MONFLUX)
V = rsYBIAMON0_Read(meYBIAMON0)
If IsNull(V) Then
    If Trim(meYBIAMON0.MONSTATUS) = "" Then
        If Trim(meYBIAMON0.MONFILE) < YBIATAB0_DATE_CPT_J Then
            If blnAuto_Form_Show Then frmBIA_ATHIC_Show 'frmCHQ_SCAN_Show
            frmBIA_ATHIC.Msg_Rcv "@CHQ_DEON"
            'frmCHQ_SCAN.Msg_Rcv "@CHQ_DEON"
        End If
    End If
End If
'--------------------------------------------------------------------------------------
frmElp.MousePointer = vbNoDrop
Wait_SS 30
frmElp.MousePointer = vbNormal
Call AUTO_COMPTA_ECRIT_LOG("14/27 = BIA_PDC")
AUTO_BIA_PDC
'--------------------------------------------------------------------------------------
frmElp.MousePointer = vbNoDrop
Wait_SS 30
frmElp.MousePointer = vbNormal
Call AUTO_COMPTA_ECRIT_LOG("16/27 = EIC_GCC")
AUTO_EIC_GCC
'--------------------------------------------------------------------------------------
frmElp.MousePointer = vbNoDrop
Wait_SS 30
frmElp.MousePointer = vbNormal
Call AUTO_COMPTA_ECRIT_LOG("17/27 = BAL_6000")
AUTO_BAL_6000 ' JPL 2011-03-22
'--------------------------------------------------------------------------------------
frmElp.MousePointer = vbNoDrop
Wait_SS 30
frmElp.MousePointer = vbNormal
Call AUTO_COMPTA_ECRIT_LOG("18/27 = RCOM_AUT")
AUTO_RCOM_AUT ' JPL 2012-06-12
'--------------------------------------------------------------------------------------
frmElp.MousePointer = vbNoDrop
Wait_SS 30
frmElp.MousePointer = vbNormal
Call AUTO_COMPTA_ECRIT_LOG("19/27 = CPT_OD")
AUTO_CPT_OD ' JPL 2012-06-12
'--------------------------------------------------------------------------------------
frmElp.MousePointer = vbNoDrop
Wait_SS 30
frmElp.MousePointer = vbNormal
Call AUTO_COMPTA_ECRIT_LOG("20/27 = BIA_IMPAYÉS")
AUTO_BIA_IMPAYES  ' jpl 2013-01-14
'--------------------------------------------------------------------------------------
frmElp.MousePointer = vbNoDrop
Wait_SS 30
frmElp.MousePointer = vbNormal
Call AUTO_COMPTA_ECRIT_LOG("21/27 = ENG_BEA_LFB")
AUTO_ENG_BEA_LFB  ' jpl 2013-05-14
'--------------------------------------------------------------------------------------
frmElp.MousePointer = vbNoDrop
Wait_SS 30
frmElp.MousePointer = vbNormal
Call AUTO_COMPTA_ECRIT_LOG("22/27 = SAB_CLIENT")
AUTO_SAB_CLIENT  ' jpl 2013-10-01
'--------------------------------------------------------------------------------------
If Mid$(YBIATAB0_DATE_CPT_J, 1, 6) <> Mid$(YBIATAB0_DATE_CPT_JS1, 1, 6) Then
    frmElp.MousePointer = vbNoDrop
    Wait_SS 30
    frmElp.MousePointer = vbNormal
    Call AUTO_COMPTA_ECRIT_LOG("24/27 = BAL_STOCK")
    AUTO_BAL_Stock
End If
'--------------------------------------------------------------------------------------
Call AUTO_COMPTA_ECRIT_LOG("23/27 = RELEVE_FOTC")
AUTO_RELEVE_FOTC ' DR 10/06/2020
'--------------------------------------------------------------------------------------
frmElp.MousePointer = vbNoDrop
Wait_SS 30
frmElp.MousePointer = vbNormal
Call AUTO_COMPTA_ECRIT_LOG("25/27 = RMA_CTL")
AUTO_RMA_CTL   'JPL 2015-01-19

Mail_Exit:
On Error GoTo Error_Handler
Call AUTO_COMPTA_ECRIT_LOG("26/27 = MAIL")
mailYBIAMON0.MONAPP = "COMPTA"
mailYBIAMON0.MONFLUX = "MAIL"
mailYBIAMON0.MONSTATUS = "MONITOR"
Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, "> " & App_Debug & " " & meYBIAMON0.MONAPP & " " & meYBIAMON0.MONFLUX)

V = fctExploitation_Auto_Control(mailYBIAMON0)
If Not IsNull(V) Then Exit Sub

V = fctExploitation_Auto_End(mailYBIAMON0)

AUTO_COMPTA_SendMail

'$JPL 2014-10-03 envoi par mail du récapitulatif des états SAB
If YBIATAB0_DATE_CPT_J <> YBIATAB0_DATE_CPT_JP0 Then
    Call frmEdition.Auto_NoPaper_Recap(paramEditionNoPaper_Folder & "PDF\Archive_" & YBIATAB0_DATE_CPT_JP0, "")
End If
Call AUTO_COMPTA_ECRIT_LOG("27/27 = NoPaper_recap")
Call frmEdition.Auto_NoPaper_Recap(paramEditionNoPaper_Folder & "PDF\Archive_" & YBIATAB0_DATE_CPT_J, "")
'--------------------------------------------------------------------------------------
Error_Handler:
End Sub
Public Sub AUTO_COMPTA_SendMail()
Dim xYBIAMON0 As typeYBIAMON0
Dim wSendMail As typeSendMail
Dim bgColor As String
Dim wPath As String
Dim xText As String
Dim XControl As String * 25
Dim xSQL As String, V


bgColor = "CYAN"
xText = ""

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIAMON7" _
       & "  where MONAPP= 'COMPTA' order by MONFLUX"
Set rsSab = cnsab.Execute(xSQL)
Do While Not rsSab.EOF
    V = rsYBIAMON0_GetBuffer(rsSab, xYBIAMON0)
    XControl = ""
    If Not IsNull(V) Then
        bgColor = "MAGENTA"
        xText = xText & V & "<BR>"
    Else
        If Trim(xYBIAMON0.MONSTATUS) <> "" Then
            bgColor = "MAGENTA"
            XControl = "? status anormal"
        End If
        If Trim(xYBIAMON0.MONFILE) <> YBIATAB0_DATE_CPT_J Then
            bgColor = "MAGENTA"
            XControl = "? Date du traitement"
        End If
        xText = xText & xYBIAMON0.MONFILE & " : " & xYBIAMON0.MONFLUX & " : " & xYBIAMON0.MONSTATUS & XControl & "<BR>"
    End If
    rsSab.MoveNext
Loop

'=====================================================================================

wSendMail.FromDisplayName = "@AUTO_COMPTA"
wSendMail.RecipientDisplayName = "INFO"

wSendMail.Subject = "Exploitation quotidienne du " & dateImp10(YBIATAB0_DATE_CPT_J)
wSendMail.Attachment = ""
wSendMail.Message = "<body bgcolor=" & Asc34 & bgColor & Asc34 & ">" _
                    & "<FONT face=" & Asc34 & prtFontName_Comic & Asc34 & ">" _
                    & htmlFontColor("BLUE") & "<BR>" & xText

wSendMail.AsHTML = True

srvSendMail.Monitor wSendMail

End Sub



Public Sub AUTO_FCI()
Dim meYBIAMON0 As typeYBIAMON0, V
AUTO_FCI:
'--------------------------------------------------------------------------------------
meYBIAMON0.MONAPP = "COMPTA"
meYBIAMON0.MONFLUX = "AUTO_FCI"
meYBIAMON0.MONSTATUS = ""
Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, "> " & App_Debug & " " & meYBIAMON0.MONAPP & " " & meYBIAMON0.MONFLUX)

V = fctExploitation_Auto_Control(meYBIAMON0)
If IsNull(V) Then
    If blnAuto_Form_Show Then frmSAB_FCI_Show
    frmSAB_FCI.Msg_Rcv "@AUTO_FCI"
    V = fctExploitation_Auto_End(meYBIAMON0)
End If

End Sub

Public Sub AUTO_GUIMAD()
Dim meYBIAMON0 As typeYBIAMON0, V
AUTO_GUIMAD:
'--------------------------------------------------------------------------------------
meYBIAMON0.MONAPP = "COMPTA"
meYBIAMON0.MONFLUX = "AUTO_GUIMA"
meYBIAMON0.MONSTATUS = ""
Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, "> " & App_Debug & " " & meYBIAMON0.MONAPP & " " & meYBIAMON0.MONFLUX)

V = fctExploitation_Auto_Control(meYBIAMON0)
If IsNull(V) Then
    If blnAuto_Form_Show Then frmBIA_GUIMAD_Show
    frmBIA_GUIMAD.Msg_Rcv "@AUTO_GUIMAD"
    V = fctExploitation_Auto_End(meYBIAMON0)
End If

End Sub

Public Sub AUTO_BIA_PDC()
Dim meYBIAMON0 As typeYBIAMON0, V
AUTO_BIA_PDC:

meYBIAMON0.MONAPP = "COMPTA"
meYBIAMON0.MONFLUX = "@BIA_PDC"
meYBIAMON0.MONSTATUS = ""
Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, "> " & App_Debug & " " & meYBIAMON0.MONAPP & " " & meYBIAMON0.MONFLUX)

V = fctExploitation_Auto_Control(meYBIAMON0)
If IsNull(V) Then
    V = cnSAB_Transaction("Commit")
    If blnAuto_Form_Show Then frmBIA_PDC_Show
    frmBIA_PDC.Msg_Rcv "@BIA_PDC"
    V = cnSAB_Transaction("BeginTrans")
    V = fctExploitation_Auto_End(meYBIAMON0)
    If blnAuto_Form_Show Then frmSAB_TC_Limites_Show
    Call AUTO_COMPTA_ECRIT_LOG("15/27 = TC_LIMITES")
    frmSAB_TC_Limites.Msg_Rcv "@TC_LIMITES"
End If

End Sub

Public Sub AUTO_EIC_GCC()
Dim meYBIAMON0 As typeYBIAMON0, V
AUTO_EIC_GCC:
'--------------------------------------------------------------------------------------
meYBIAMON0.MONAPP = "COMPTA"
meYBIAMON0.MONFLUX = "@EIC_GCC"
meYBIAMON0.MONSTATUS = ""
Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, "> " & App_Debug & " " & meYBIAMON0.MONAPP & " " & meYBIAMON0.MONFLUX)

V = fctExploitation_Auto_Control(meYBIAMON0)
If IsNull(V) Then
    V = cnSAB_Transaction("Commit")
    If blnAuto_Form_Show Then frmYEICGCC0_ATHIC_Show
    frmYEICGCC0_ATHIC.Msg_Rcv "@EIC_GCC"
    V = cnSAB_Transaction("BeginTrans")
    V = fctExploitation_Auto_End(meYBIAMON0)
End If

End Sub


Public Sub AUTO_TVAFAC()
Dim meYBIAMON0 As typeYBIAMON0, V
AUTO_TVAFAC:
'--------------------------------------------------------------------------------------
meYBIAMON0.MONAPP = "COMPTA"
meYBIAMON0.MONFLUX = "AUTO_TVAFA"
meYBIAMON0.MONSTATUS = ""
Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, "> " & App_Debug & " " & meYBIAMON0.MONAPP & " " & meYBIAMON0.MONFLUX)

V = fctExploitation_Auto_Control(meYBIAMON0)
If IsNull(V) Then
    V = cnSAB_Transaction("Commit")
    If blnAuto_Form_Show Then frmBIA_TVAFAC_Show
    frmBIA_TVAFAC.Msg_Rcv "@AUTO_TVAFAC"
    V = cnSAB_Transaction("BeginTrans")
    V = fctExploitation_Auto_End(meYBIAMON0)
End If

End Sub
Public Sub AUTO_ROPDOS()
Dim meYBIAMON0 As typeYBIAMON0, V
AUTO_ROPDOS:
'--------------------------------------------------------------------------------------
meYBIAMON0.MONAPP = "COMPTA"
meYBIAMON0.MONFLUX = "AUTO_ROPDO"
meYBIAMON0.MONSTATUS = ""
Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, "> " & App_Debug & " " & meYBIAMON0.MONAPP & " " & meYBIAMON0.MONFLUX)

V = fctExploitation_Auto_Control(meYBIAMON0)
If IsNull(V) Then
    V = cnSAB_Transaction("Commit")
    '$JPL 2014-11-03 If blnAuto_Form_Show Then frmDROPI_Show
    '$JPL 2014-11-03 frmDROPI.Msg_Rcv "@AUTO_ROPDOS"
    V = cnSAB_Transaction("BeginTrans")
    V = fctExploitation_Auto_End(meYBIAMON0)
End If

End Sub

Public Sub AUTO_BAL_PCI_DC()
Dim meYBIAMON0 As typeYBIAMON0, V

meYBIAMON0.MONAPP = "COMPTA"
meYBIAMON0.MONFLUX = "BAL_PCI_DC"
meYBIAMON0.MONSTATUS = ""
Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, "> " & App_Debug & " " & meYBIAMON0.MONAPP & " " & meYBIAMON0.MONFLUX)

V = fctExploitation_Auto_Control(meYBIAMON0)
If IsNull(V) Then
    If blnAuto_Form_Show Then frmSAB_Balance_Show
    frmSAB_Balance.Msg_Rcv "@BAL_PCI_DC"
    V = fctExploitation_Auto_End(meYBIAMON0)
End If

End Sub

Public Sub frmBIA_CLIPRO_Show()
Dim IdShell As Variant, X As String

X = Shell_VB6("BIA_CLIPRO")
X = "C:\BIASRV\YCLIPRO.exe" & " " & X
IdShell = Shell(X, 1)
AppActivate IdShell
DoEvents

End Sub
Public Sub frmBIA_EICGCC_Show()
Dim IdShell As Variant, X As String

If paramEnvironnement = constProduction Then
    X = Shell_VB6("BIA_EICGCC")
    X = "C:\BIASRV\YEICGCC.exe" & " " & X
Else
    X = Shell_VB6("BIA_EICGCC_ATHIC")
    X = "C:\BIASRV\YEICGCC_ATHIC.exe" & " " & X
End If

IdShell = Shell(X, 1)
AppActivate IdShell
DoEvents

End Sub
Public Sub frmBIA_Quid_Show()
Dim X As String
On Error Resume Next
frmBIA_Quid.Icon = frmElp_Icon
frmBIA_Quid.Show vbModeless
frmBIA_Quid.WindowState = vbNormal
frmBIA_Quid.Visible = True
X = frmBIA_Quid.Caption
AppActivate X

End Sub
Public Sub frmBIA_CLISTA_Show()
Dim IdShell As Variant, X As String

X = Shell_VB6("BIA_CLISTA")
X = "C:\BIASRV\YCLISTA.exe" & " " & X
IdShell = Shell(X, 1)
AppActivate IdShell
DoEvents

End Sub



Public Sub AUTO_BAL_6000()
BAL_6000:

Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, "> " & App_Debug & " Compta : @BAL_6000")

If blnAuto_Form_Show Then frmSAB_Balance_Show
frmSAB_Balance.Msg_Rcv "@BAL_6000"

End Sub

Public Sub AUTO_RCOM_AUT()
RCOM_AUT: ' JPL 2012-06-12
'--------------------------------------------------------------------------------------
Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, "> " & App_Debug & " Compta : @RCOM_AUT")

If blnAuto_Form_Show Then frmSAB_Balance_Show
frmSAB_Balance.Msg_Rcv "@RCOM_AUT"

End Sub

Public Sub AUTO_RMA_CTL()
RMA_CTL: ' JPL 2015-01-15
'--------------------------------------------------------------------------------------
Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, "> " & App_Debug & " Compta : @RMA_CTL")

If blnAuto_Form_Show Then frmBIA_Quid_Show
frmBIA_Quid.Msg_Rcv "@RMA_CTL"

End Sub

Public Sub AUTO_CPT_OD()
RCOM_AUT: ' JPL 2012-07-28
'--------------------------------------------------------------------------------------
Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, "> " & App_Debug & " Compta : @CPT_OD")

If blnAuto_Form_Show Then frmSAB_Balance_Show
frmSAB_Balance.Msg_Rcv "@CPT_OD"

End Sub

Public Sub AUTO_ENG_BEA_LFB()
RCOM_AUT: ' JPL 2012-07-28
'--------------------------------------------------------------------------------------
Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, "> " & App_Debug & " Compta : @CPT_OD")

If blnAuto_Form_Show Then frmSAB_Balance_Show
frmSAB_Balance.Msg_Rcv "@ENG_BEA_LFB"

End Sub

Public Sub AUTO_SAB_CLIENT()
RCOM_AUT: ' JPL 2013-10-01
'--------------------------------------------------------------------------------------
If Mid$(YBIATAB0_DATE_CPT_J, 1, 6) <> Mid$(YBIATAB0_DATE_CPT_JS1, 1, 6) Then

    Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, "> " & App_Debug & " Compta : @SAB_CLIENT")
    
    If blnAuto_Form_Show Then frmSAB_CLIENT_Show
    frmSAB_Client.Msg_Rcv "@SAB_CLIENT"
End If
End Sub

Public Sub AUTO_BIA_IMPAYES()
BIA_IMPAYES: ' JPL 2013-01-14
'--------------------------------------------------------------------------------------
Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, "> " & App_Debug & " Compta : @BIA_IMPAYES")

If blnAuto_Form_Show Then frmBIA_Impayés_Show
frmBIA_Impayés.Msg_Rcv "@BIA_IMPAYÉS"

End Sub

Public Sub Jpl_Test()
Dim IdShell, V
Dim wSendMail As typeSendMail
Dim xSQL As String
Dim xDétail_D As String, xHeader_D As String
Dim mbgColor As String
Dim X0 As String, X1 As String, curX As Currency

On Error GoTo Error_Handler

'"C:\Program Files\Axantum\AxCrypt\AxCrypt.exe" -b 2 -e -k "toto" -c -z "Bia_SAB.pdf"
IdShell = Shell("c:\temp\jplax.bat" & " > " & "c:\temp\jplax.Log", 1)
If IdShell > 0 Then
'    AppActivate IdShell
Else
    MsgBox "???", vbCritical, "Jpl_Test"
End If


wSendMail.Subject = "Test Axcrypt"



'-----------------------------------------------------------------------------------
wSendMail.FromDisplayName = "JPL"
wSendMail.RecipientDisplayName = "rrr"
'-----------------------------------------------------------------------------------
wSendMail.From = currentSSIWINMAIL

wSendMail.CcRecipient = currentSSIWINMAIL







wSendMail.Attachment = "c:\temp\Bia_SAB-pdf.txt"


wSendMail.Message = "<" & mbgColor & "><BR> Bla bla bla" _

wSendMail.AsHTML = True
srvSendMail.Monitor wSendMail
Exit Sub
'------------------------------------------
Error_Handler:
    
    V = Error
Error_MsgBox:

End Sub

Public Sub AUTO_BAL_Stock()

frmSAB_Balance.Msg_Rcv "@BAL_Stock"

frmElp.MousePointer = vbNoDrop
Wait_SS 30
frmElp.MousePointer = vbNormal

frmSAB_Stock.Msg_Rcv "@SAB_Stock"

End Sub
Public Sub AUTO_RELEVE_FOTC()

    frmSAB_Balance.Msg_Rcv "@RELEVE_FOTC"

End Sub


