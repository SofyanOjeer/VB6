Attribute VB_Name = "Bia_Audit_Monitor"
Option Explicit
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



Public Sub AUTO_JRN()
Dim meYBIAMON0 As typeYBIAMON0
Dim mailYBIAMON0 As typeYBIAMON0
Dim V
On Error GoTo Mail_Exit

If xlsManual Then
    If appExcelPublic Is Nothing Then
        Set appExcelPublic = CreateObject("Excel.Application")
        appExcelPublic.Visible = False
        appExcelPublic.ControlCharacters = False
        appExcelPublic.Interactive = False
    End If
End If

mainSoc_AMJCPT_Load
'--------------------------------------------------------------------------------------
YBIATAB0_DATE_CPT_J = YBIATAB0_DATE_CPT_JP0
'--------------------------------------------------------------------------------------


App_Debug = "> @AUTO_JRN"
'--------------------------------------------------------------------------------------
'MAIL: traitement YBIAJOUR terminé ? déjà traité ce jour ?
'--------------------------------------------------------------------------------------
Call ECRIT_LOG2008("Avant MONAPP = ""SMS"", MONFLUX = ""@YBIAJOUR""")
meYBIAMON0.MONAPP = "SMS"
meYBIAMON0.MONFLUX = "@YBIAJOUR"
V = rsYBIAMON0_Read(meYBIAMON0)
If IsNull(V) Then
    If Trim(meYBIAMON0.MONSTATUS) <> "" Then V = "BIAJOUR : statut " & Trim(meYBIAMON0.MONSTATUS) & " : " & meYBIAMON0.MONFILE
End If
If Not IsNull(V) Then
    Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, V)
    Exit Sub
End If

Call ECRIT_LOG2008("Avant MONAPP = ""@AUTO_JRN"", MONFLUX = ""MAIL"", MONSTATUS = """"")
mailYBIAMON0.MONAPP = "@AUTO_JRN"
mailYBIAMON0.MONFLUX = "MAIL"
mailYBIAMON0.MONSTATUS = ""
V = fctExploitation_Auto_Control(mailYBIAMON0)
If Not IsNull(V) Then Exit Sub
Call ECRIT_LOG2008("Avant V = cnSAB_Transaction(""Commit"") dans AUTO_JRN")
V = cnSAB_Transaction("Commit")

'DR Désactivée 11/07/2019
'JRN_CDO:
''--------------------------------------------------------------------------------------
'meYBIAMON0.MONAPP = "@AUTO_JRN"
'meYBIAMON0.MONFLUX = "@JRN_CDO"
'meYBIAMON0.MONSTATUS = ""
'
'V = fctExploitation_Auto_Control(meYBIAMON0)
'If IsNull(V) Then
'    If blnAuto_Form_Show Then frmJRN_CDO_Show
'    frmJRN_CDO.Msg_Rcv "@JRN_CDO"
'    V = fctExploitation_Auto_End(meYBIAMON0)
'End If
JRN_DAT:
'--------------------------------------------------------------------------------------
Call ECRIT_LOG2008("Avant MONAPP = ""@AUTO_JRN"", MONFLUX = ""@JRN_DAT""")
meYBIAMON0.MONAPP = "@AUTO_JRN"
meYBIAMON0.MONFLUX = "@JRN_DAT"
meYBIAMON0.MONSTATUS = ""

V = fctExploitation_Auto_Control(meYBIAMON0)
If IsNull(V) Then
    If blnAuto_Form_Show Then frmJRN_DAT_Show
    frmJRN_DAT.Msg_Rcv "@JRN_DAT"
    V = fctExploitation_Auto_End(meYBIAMON0)
End If


JRN_COMPTE:
'--------------------------------------------------------------------------------------
Call ECRIT_LOG2008("Avant MONAPP = ""@AUTO_JRN"", MONFLUX = ""@JRN_COMPT""")
meYBIAMON0.MONAPP = "@AUTO_JRN"
meYBIAMON0.MONFLUX = "@JRN_COMPT"
meYBIAMON0.MONSTATUS = ""

V = fctExploitation_Auto_Control(meYBIAMON0)
If IsNull(V) Then
    If blnAuto_Form_Show Then frmJRN_JCompte0_Show
    frmJRN_JCOMPTE0.Msg_Rcv "@JRN_COMPTE"
    V = fctExploitation_Auto_End(meYBIAMON0)
End If


JRN_CLIENT:
'--------------------------------------------------------------------------------------
Call ECRIT_LOG2008("Avant MONAPP = ""@AUTO_JRN"", MONFLUX = ""@JRN_CLIEN""")
meYBIAMON0.MONAPP = "@AUTO_JRN"
meYBIAMON0.MONFLUX = "@JRN_CLIEN"
meYBIAMON0.MONSTATUS = ""

V = fctExploitation_Auto_Control(meYBIAMON0)
If IsNull(V) Then
    If blnAuto_Form_Show Then frmJRN_JCLIENT0_Show
    frmJRN_JCLIENT0.Msg_Rcv "@JRN_CLIENT"
    V = fctExploitation_Auto_End(meYBIAMON0)
End If

'--------------------------------------------------------------------------------------

CPT_SCHEMA:
'--------------------------------------------------------------------------------------
Call ECRIT_LOG2008("Avant MONAPP = ""@AUTO_JRN"", MONFLUX = ""CPT_SCHEMA""")
meYBIAMON0.MONAPP = "@AUTO_JRN"
meYBIAMON0.MONFLUX = "CPT_SCHEMA"
meYBIAMON0.MONSTATUS = ""

V = fctExploitation_Auto_Control(meYBIAMON0)
If IsNull(V) Then
    V = cnSAB_Transaction("Commit")
    If blnAuto_Form_Show Then frmCPT_SCHEMA_Show
    frmCPT_SCHEMA.Msg_Rcv "@CPT_SCHEMA"
    V = cnSAB_Transaction("BeginTrans")
    V = fctExploitation_Auto_End(meYBIAMON0)
End If

'--------------------------------------------------------------------------------------

JRN_MNU:
'--------------------------------------------------------------------------------------
Call ECRIT_LOG2008("Avant MONAPP = ""@AUTO_JRN"", MONFLUX = ""@JRN_MNU""")
meYBIAMON0.MONAPP = "@AUTO_JRN"
meYBIAMON0.MONFLUX = "@JRN_MNU"
meYBIAMON0.MONSTATUS = ""

V = fctExploitation_Auto_Control(meYBIAMON0)
If IsNull(V) Then
    If blnAuto_Form_Show Then frmJRN_MNU_Show
    frmJRN_MNU.Msg_Rcv "@JRN_MNU"
    V = fctExploitation_Auto_End(meYBIAMON0)
End If

JRN_SWI:
'--------------------------------------------------------------------------------------
Call ECRIT_LOG2008("Avant MONAPP = ""@AUTO_JRN"", MONFLUX = ""@JRN_SWI""")
meYBIAMON0.MONAPP = "@AUTO_JRN"
meYBIAMON0.MONFLUX = "@JRN_SWI"
meYBIAMON0.MONSTATUS = ""

V = fctExploitation_Auto_Control(meYBIAMON0)
If IsNull(V) Then
    If blnAuto_Form_Show Then frmJRN_SWI_Show
    frmJRN_SWI.Msg_Rcv "@JRN_SWI"
    V = fctExploitation_Auto_End(meYBIAMON0)
End If

BIA_SSI_JRN:
'--------------------------------------------------------------------------------------
Call ECRIT_LOG2008("Avant MONAPP = ""@AUTO_JRN"", MONFLUX = ""@BIA_SSI_J""")
meYBIAMON0.MONAPP = "@AUTO_JRN"
meYBIAMON0.MONFLUX = "@BIA_SSI_J"  'RN
meYBIAMON0.MONSTATUS = ""

V = fctExploitation_Auto_Control(meYBIAMON0)
If IsNull(V) Then
    V = fctExploitation_Auto_End(meYBIAMON0)  ' shunt : incompatibilité avec màj Transaction @BIA_SSI_JRN
    If blnAuto_Form_Show Then frmJRN_SWI_Show
    frmBIA_SSI.Msg_Rcv "@BIA_SSI_JRN"
End If

CRE_ANO:
'--------------------------------------------------------------------------------------
Call ECRIT_LOG2008("Avant MONAPP = ""@AUTO_JRN"", MONFLUX = ""@CRE_ANO""")
meYBIAMON0.MONAPP = "@AUTO_JRN"
meYBIAMON0.MONFLUX = "@CRE_ANO"
meYBIAMON0.MONSTATUS = ""

V = fctExploitation_Auto_Control(meYBIAMON0)
If IsNull(V) Then
    If blnAuto_Form_Show Then frmYCREANO0_Show
    frmYCREANO0.Msg_Rcv "@CRE_ANO"
    V = fctExploitation_Auto_End(meYBIAMON0)
End If


'--------------------------------------------------------------------------------------
Mail_Exit:
On Error GoTo Error_Handler
Call ECRIT_LOG2008("Avant MONAPP = ""@AUTO_JRN"", MONFLUX = ""MAIL"", MONSTATUS = ""MONITOR""")
mailYBIAMON0.MONAPP = "@AUTO_JRN"
mailYBIAMON0.MONFLUX = "MAIL"
mailYBIAMON0.MONSTATUS = "MONITOR"

V = fctExploitation_Auto_Control(mailYBIAMON0)
If Not IsNull(V) Then Exit Sub

V = fctExploitation_Auto_End(mailYBIAMON0)

AUTO_JRN_SendMail
'--------------------------------------------------------------------------------------
Error_Handler:
End Sub

Public Sub AUTO_JRN_SendMail()
Dim xYBIAMON0 As typeYBIAMON0
Dim wSendMail As typeSendMail
Dim bgColor As String
Dim wPath As String
Dim xText As String
Dim XControl As String * 25
Dim xSql As String, V


bgColor = "CYAN"
xText = ""

xSql = "select * from " & paramIBM_Library_SABSPE & ".YBIAMON7" _
       & "  where MONAPP= '@AUTO_JRN' order by MONFLUX"
Set rsSab = cnsab.Execute(xSql)
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

wSendMail.FromDisplayName = "@AUTO_JRN"
wSendMail.RecipientDisplayName = "INFO"

wSendMail.Subject = "AUTO_JRN du " & dateImp10(YBIATAB0_DATE_CPT_J)
wSendMail.Attachment = ""
wSendMail.Message = "<body bgcolor=" & Asc34 & bgColor & Asc34 & ">" _
                    & "<FONT face=" & Asc34 & prtFontName_Comic & Asc34 & ">" _
                    & htmlFontColor("BLUE") & "<BR>" & xText

wSendMail.AsHTML = True

srvSendMail.Monitor wSendMail

End Sub




Public Sub AUTO_JRN_2008()
'Dim fic As Long
'Dim strChaine As String
'Dim ok As Boolean
'Static enCours As Boolean
'
'    If Not enCours Then
'        enCours = True
'        frmElp.Timer1.Enabled = False
'        frmElp.Timer1.Interval = 0
'        fic = FreeFile
'        Open "c:\temp\imp_pdf\Bia_SAB2008.log" For Input As #fic
'        ok = False
'        Do Until EOF(fic)
'            Line Input #fic, strChaine
'            If InStr(strChaine, "Fin") > -1 Then
'                ok = True
'            End If
'        Loop
'        Close #fic
'        If ok = False Then
'            frmElp.Timer1.Enabled = True
'            frmElp.Timer1.Interval = 1000
'            enCours = False
'        Else
'            fic = FreeFile
'            Open "c:\temp\imp_pdf\Bia_Audit2008.log" For Output As #fic
'            Print #fic, "Début AUTO_JRN_2008 --> " & CDate(Now)
'            Close #fic
'            Call frmJRN_DAT.Msg_Rcv("@JRN_DAT")
'            Call frmJRN_JCOMPTE0.Msg_Rcv("@JRN_COMPTE")
'            Call frmJRN_MNU.Msg_Rcv("@JRN_MNU")
'            Call frmJRN_SWI.Msg_Rcv("@JRN_SWI")
'            appExcelPublic.Quit
'            Set appExcelPublic = Nothing
'            fic = FreeFile
'            Open "c:\temp\imp_pdf\Bia_Audit2008.log" For Append As #fic
'            Print #fic, "Fin AUTO_JRN_2008 --> " & CDate(Now)
'            Close #fic
'            End
'        End If
'    End If

End Sub

Public Sub mainSocExe()

paramIMP_PDFCreator_Name = "PDF_BIA_AUDIT"
paramIMP_PDF_Path_VBP = "C:\Temp\IMP_PDF\BIA_AUDIT"

If Not msFileSystem.FolderExists(paramIMP_PDF_Path_VBP) Then paramIMP_PDF_Path_VBP = paramIMP_PDF_Path_Temp
paramIMP_PDF_Path = paramIMP_PDF_Path_Temp

frmElp_Caption = "BIA_AUDIT"
Set frmElp_Icon = frmJRN_CDO
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
        Case Is = "BIA_SSI", "@BIA_SSI", "@BIA_SSI_JRN": frmBIA_SSI_Show: frmBIA_SSI.Msg_Rcv Msg
        Case Is = "CRE_ANO", "@CRE_ANO": frmYCREANO0_Show: frmYCREANO0.Msg_Rcv Msg
        Case Is = "CPT_SCHEMA": frmCPT_SCHEMA_Show: frmCPT_SCHEMA.Msg_Rcv Msg
        Case Is = "CPT_FEC": frmYFECDTA0_Show: frmYFECDTA0.Msg_Rcv Msg
        Case Is = "JRN_CDO": frmJRN_CDO_Show: frmJRN_CDO.Msg_Rcv Msg
        Case Is = "JRN_COMPTE": frmJRN_JCompte0_Show: frmJRN_JCOMPTE0.Msg_Rcv Msg
        Case Is = "JRN_CLIENT": frmJRN_JCLIENT0_Show: frmJRN_JCLIENT0.Msg_Rcv Msg
        Case Is = "JRN_DAT": frmJRN_DAT_Show: frmJRN_DAT.Msg_Rcv Msg
        Case Is = "JRN_MNU", "@JRN_MNU": frmJRN_MNU_Show: frmJRN_MNU.Msg_Rcv Msg
        Case Is = "JRN_SWI": frmJRN_SWI_Show: frmJRN_SWI.Msg_Rcv Msg
        Case Is = "JRN_ZLIBEL0": frmJRN_JLIBEL0_Show: frmJRN_JLIBEL0.Msg_Rcv Msg
        Case Is = "COM_RETRO": frmYCOMRCD0_Show: frmYCOMRCD0.Msg_Rcv Msg
        Case Is = "@CPT_SCHEMA":
    
                                mainSoc_AMJCPT_Load
                                frmCPT_SCHEMA_Show
                                frmCPT_SCHEMA.Msg_Rcv Msg
        Case Is = "@JRN_CDO": mainSoc_AMJCPT_Load
                            If blnAuto_Exploitation_Ok("DATE_CPT_J", "@JRN_CDO") Then
                                If blnAuto_Form_Show Then frmJRN_CDO_Show
                                frmJRN_CDO.Msg_Rcv Msg
                                Call blnAuto_Exploitation_Ok("Update", "@JRN_CDO")
                            End If
        Case Is = "@JRN_COMPTE": mainSoc_AMJCPT_Load
                                If blnAuto_Form_Show Then frmJRN_JCompte0_Show
                                frmJRN_JCOMPTE0.Msg_Rcv Msg
         Case Is = "@JRN_CLIENT": mainSoc_AMJCPT_Load
                                If blnAuto_Form_Show Then frmJRN_JCLIENT0_Show
                                frmJRN_JCLIENT0.Msg_Rcv Msg
       Case Is = "@JRN_DAT": mainSoc_AMJCPT_Load
                                If blnAuto_Form_Show Then frmJRN_DAT_Show
                                frmJRN_DAT.Msg_Rcv Msg
      Case Is = "@JRN_SWI": mainSoc_AMJCPT_Load
                                If blnAuto_Form_Show Then frmJRN_SWI_Show
                                frmJRN_SWI.Msg_Rcv Msg
        Case Is = "@JRN_ZLIBEL0": mainSoc_AMJCPT_Load
                                If blnAuto_Form_Show Then frmJRN_JLIBEL0_Show
                                frmJRN_JLIBEL0.Msg_Rcv Msg
    
        Case Is = "X_RESET":  main_Reset
        Case Is = "XUSRID": XUsrId_Show
        Case Is = "@AUTO_JRN": AUTO_JRN
        Case Is = "X_I5A7": X_I5A7_Show
    End Select

End Sub
Public Sub frmJRN_CDO_Show()
Dim X As String

frmJRN_CDO.Icon = frmElp_Icon
frmJRN_CDO.Show vbModeless
frmJRN_CDO.WindowState = vbNormal
frmJRN_CDO.Visible = True
X = frmJRN_CDO.Caption
AppActivate X

End Sub

Public Sub frmYCREANO0_Show()
Dim X As String

frmYCREANO0.Icon = frmElp_Icon
frmYCREANO0.Show vbModeless
frmYCREANO0.WindowState = vbNormal
frmYCREANO0.Visible = True
X = frmYCREANO0.Caption
AppActivate X

End Sub

Public Sub frmYFECDTA0_Show()
Dim X As String

frmYFECDTA0.Icon = frmElp_Icon
frmYFECDTA0.Show vbModeless
frmYFECDTA0.WindowState = vbNormal
frmYFECDTA0.Visible = True
X = frmYFECDTA0.Caption
AppActivate X

End Sub


Public Sub frmYCOMRCD0_Show()
Dim X As String

frmYCOMRCD0.Icon = frmElp_Icon
frmYCOMRCD0.Show vbModeless
frmYCOMRCD0.WindowState = vbNormal
frmYCOMRCD0.Visible = True
X = frmYCOMRCD0.Caption
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
Public Sub frmJRN_MNU_Show()
Dim X As String

frmJRN_MNU.Icon = frmElp_Icon
frmJRN_MNU.Show vbModeless
frmJRN_MNU.WindowState = vbNormal
frmJRN_MNU.Visible = True
X = frmJRN_MNU.Caption
AppActivate X

End Sub

Public Sub frmCPT_SCHEMA_Show()
Dim X As String

frmCPT_SCHEMA.Icon = frmElp_Icon
frmCPT_SCHEMA.Show vbModeless
frmCPT_SCHEMA.WindowState = vbNormal
frmCPT_SCHEMA.Visible = True
X = frmCPT_SCHEMA.Caption
AppActivate X

End Sub
Public Sub frmBIA_SSI_Show()
Dim X As String

frmBIA_SSI.Icon = frmElp_Icon
frmBIA_SSI.Show vbModeless
frmBIA_SSI.WindowState = vbNormal
frmBIA_SSI.Visible = True
X = frmBIA_SSI.Caption
'AppActivate X

End Sub

Public Sub frmJRN_JCompte0_Show()
Dim X As String

frmJRN_JCOMPTE0.Icon = frmElp_Icon
frmJRN_JCOMPTE0.Show vbModeless
frmJRN_JCOMPTE0.WindowState = vbNormal
frmJRN_JCOMPTE0.Visible = True
X = frmJRN_JCOMPTE0.Caption
AppActivate X

End Sub
Public Sub frmJRN_JLIBEL0_Show()
Dim X As String

frmJRN_JLIBEL0.Icon = frmElp_Icon
frmJRN_JLIBEL0.Show vbModeless
frmJRN_JLIBEL0.WindowState = vbNormal
frmJRN_JLIBEL0.Visible = True
X = frmJRN_JLIBEL0.Caption
AppActivate X

End Sub

Public Sub frmJRN_JCLIENT0_Show()
Dim X As String

frmJRN_JCLIENT0.Icon = frmElp_Icon
frmJRN_JCLIENT0.Show vbModeless
frmJRN_JCLIENT0.WindowState = vbNormal
frmJRN_JCLIENT0.Visible = True
X = frmJRN_JCLIENT0.Caption
AppActivate X

End Sub

Public Sub frmJRN_SWI_Show()
Dim X As String

frmJRN_SWI.Icon = frmElp_Icon
frmJRN_SWI.Show vbModeless
frmJRN_SWI.WindowState = vbNormal
frmJRN_SWI.Visible = True
X = frmJRN_SWI.Caption
AppActivate X

End Sub


Public Sub frmJRN_DAT_Show()
Dim X As String

frmJRN_DAT.Icon = frmElp_Icon
frmJRN_DAT.Show vbModeless
frmJRN_DAT.WindowState = vbNormal
frmJRN_DAT.Visible = True
X = frmJRN_DAT.Caption
AppActivate X

End Sub

