Attribute VB_Name = "BIA"
Option Explicit
 
Public cnsab As New ADODB.Connection
Public rsSab As New ADODB.Recordset
Public cnSab_Update As New ADODB.Connection
Public rsSab_Update As New ADODB.Recordset

Public paramSOC_RS As String
Public paramSOC_Adresse As String
Public paramSOC_Ville As String
Public paramsoc_capital  As String
Public paramSOC_Télécom As String
Public paramSoc_Capital_Télécom As String
Public paramSOC_TVA_Intracommunautaire As String

Public AccAutId As String * 10
Public paramBiaPgm  As String * 12
Public paramBiaPgmAut As String * 12

Public paramServerSplf As String

Type typeAuthorization
    Consulter  As Boolean
    Saisir  As Boolean
    Valider As Boolean
    Comptabiliser As Boolean
    Rapprocher  As Boolean
    Swift  As Boolean
    Virement  As Boolean
    Avis  As Boolean
    Xspécial  As Boolean
    X09  As Boolean
    X10  As Boolean
    X11  As Boolean
    X12  As Boolean
    X13  As Boolean
    X14  As Boolean
    X15  As Boolean
    X16  As Boolean
    X17  As Boolean
    X18  As Boolean
End Type


Public paramElpKM_Folder As String

Public paramEdition_Print_Test As Boolean
Public paramEdition_Print_Production As Boolean
Public paramEditionSplf_Folder As String
Public oPDF As New clsPDFCreator

Public paramEditionNoPaper_Folder As String, paramEditionNoPaper_Partage As String
Public blnEditionNoPaper_Auto As Boolean, paramEditionNoPaper_Auto_Unit As String
Public paramEditionNoPaper_Auto_PgmName As String, paramEditionNoPaper_Auto_Dir As String
Public paramEditionNoPaper_Auto_Lnk As String
Public paramEditionNoPaper_Folder_MakePDF  As String
Public blnMakePDF_Actif  As Boolean, mMakePDF_Error As Long, mMakePDF_Error_Loop As Long
Public Const paramMakePDF_Name = "makePDF"
Public Const paramMakePDF2008_Name = "makePDF2008"

Public paramEditionCourrier_Folder As String
Public paramEditionFiligrane_Folder As String
Public paramEditionFtp_File As String
Public paramEditionCorbeille_Folder As String
Public paramEditionArchive_Folder As String
Public blncmdSplfMonitor As Boolean
Public blnMonitor As Boolean

Public paramIBM_AS400_FTP As String
Public paramIBM_AS400_ID As String
Public paramIBM_ODBC_SAB As String
Public paramIBM_Library_SAB As String
Public paramIBM_Library_SAB_P As String
Public paramIBM_Library_SABSPE As String
Public paramIBM_Library_SABSPE_P As String
Public paramIBM_Library_SABSPE_XXX As String
Public paramIBM_Library_SABJRN As String
Public paramIBM_Library_BIADWH As String
Public paramIBM_Library_BODWH As String
Public paramIBM_Library_File As String
Public paramIBM_Library_Src As String
Public paramIBM_Library_Obj As String
Public paramIBM_QSYSOPR As String

Public paramYBase_Data As String, paramYBase_DataF As String
Public paramYBase_Data_Extension As String, paramYBase_Data_ExtensionP As String

Public paramFTP_Out As String, paramFTP_OutF As String
Public paramFTP_Out_XCom As String, paramFTP_Out_XComF As String
Public paramFTP_In As String, paramFTP_InF As String
Public paramFTP_SPLF As String

Public paramPeliNT_Data As String, paramPeliNT_DataF As String
Public paramPeliNT_Monitor As String, paramPeliNT_MonitorF As String
Public paramPeliNT_Exe As String, paramPeliNT_ExeF As String
Public paramPeliNT_Aller_XcomF As String
Public paramPeliNT_Aller_FTPF As String
Public paramPeliNT_Retour_XcomF As String
Public paramPeliNT_Retour_FTPF As String
Public paramPeliNT_Connexion_EnCours As String

Public localPeliNT_DataF As String
Public localPeliNT_MonitorF As String
Public localPeliNT_ExeF As String

Public arrAMJCPT(30) As String * 8


Type typeUser
    Id  As String * 12
    Name  As String * 40
    Unit  As String * 4             'Service SOBF SOBI INFO .....
    ProdTest As String * 1          ' P ou T
    Edition_Hold As String * 1      ' pas d'impression automatique
    Edition_Aut As String * 1       ' Autorisation aux spoules du service 0 à 9
    QSYSOPR As String * 1           ' opérateur informatique  ==> ventilation unit
    XXXXXX As String * 1            ' libre
    Printer  As String * 40         ' nom de l'imprimante locale <> impr du service
    AliasWin As String * 12         ' Id windows pour attribution droits ACL sur fichiers spoules
End Type

Type typeUnit
    Id  As String * 12
    Name  As String * 40
    Printer  As String * 40
End Type

Public usrIdSAB As String
Public currentUnit As typeUnit
Public currentUser As typeUser, idemUser As typeUser
Public currentZMNURUT0 As typeZMNURUT0
Public currentZMNUUTI0 As typeZMNUUTI0
Public currentZMNUUTP0 As typeZMNUUTP0
Public currentZMNUHLB0 As typeZMNUHLB0
Public currentCLIENASIG As String

'$JPL 2013-10-03Public currentSSIWINUNIT As String, currentService_Lib As String
Public currentSAB_ETA As Long, currentSAB_AGE As Long, currentSAB_PLA As Long

Public currentSSIWINUIDN As Long
Public currentSSIWINUIDD As Long
Public currentSSIWINUIDX As String, currentSSIWINUIDX_U As String
Public currentSSIWINUNOM As String
Public currentSSIWINMAIL As String
Public currentSSIWINUNIT As String
Public currentSSIWINUNIT_Lib As String

Public YBIATAB0_DATE_CPT_J As String * 8, YBIATAB0_DIBM_CPT_J As String * 8
Public YBIATAB0_DATE_CPT_JP0 As String * 8
Public YBIATAB0_DATE_CPT_JP1 As String * 8, YBIATAB0_DIBM_CPT_JP1 As String * 7
Public YBIATAB0_DATE_CPT_JS1 As String * 8, YBIATAB0_DIBM_CPT_JS1 As String * 7
Public YBIATAB0_DATE_CPT_M As String * 8
Public YBIATAB0_DATE_CPT_MP1 As String * 8
Public YBIATAB0_DATE_CPT_MP2 As String * 8
Public YBIATAB0_DATE_CPT_MS1 As String * 8
Public YBIATAB0_DATE_CPT_A As String * 8
Public YBIATAB0_DATE_CPT_AP1 As String * 8
Public YBIATAB0_DATE_CPT_AS1 As String * 8

Public YBIATAB0_DATE_CAL_AP1 As String * 8
Public YBIATAB0_DATE_CAL_MP1 As String * 8

Public paramIBM_BIA_INFO_Password As String
Public paramIBM_BIA_AUTO_Password As String
Public paramIBM_BIA_ODBC_Password As String
Public paramIBM_BIA_DWH_Password As String
Public paramIBM_BO_DWH_Password As String


Public paramSAA_Data_Archive As String, paramSAA_DataF_Archive As String
Public paramSAA_DataF_Log As String
Public paramSAA_Data_from_SAB As String, paramSAA_DataF_from_SAB As String
Public paramSAA_Data_from_SAB_ExtensionP_sab As String
Public paramSAA_Data_from_SAB_ExtensionP_pcc As String
Public paramSAA_Data_from_SAB_ExtensionP_rje As String
Public paramSAA_Data_from_SAB_ExtensionP_sav As String
Public paramSAA_Data_from_SAB_YFile As String

Public paramSAA_Data_to_SAB As String, paramSAA_DataF_to_SAB As String
Public paramSAA_Data_to_SAB_ExtensionP_out As String
Public paramSAA_Data_to_SAB_ExtensionP_sav As String
Public paramSAA_Data_to_SAB_YFile As String


Public paramSAA_Data_from_MT950 As String, paramSAA_DataF_from_MT950 As String
Public paramCorona_DataF_Swift_In As String

Public paramSAA_Data_to_Corona As String, paramSAA_DataF_to_Corona As String
Public paramSAA_Data_to_Corona_ExtensionP_out As String
Public paramSAA_Data_to_Corona_ExtensionP_sav As String



Public paramSwift_BIC_YFile As String
Public paramSwift_BIC_Input As String


Public paramSwiftSAACorona_SAA_Out As String
Public paramSwiftSAACorona_Corona_Wait As String
Public paramSwiftSAACorona_Corona_In As String

Public paramSwiftLoro_SAA_In As String
Public paramSwiftLoro_SAA_Wait As String
Public paramSwiftLoro_MT950_File As String

Public paramSwiftNostro_Corona_In As String
Public paramSwiftNostro_Corona_Wait As String
Public paramSwiftNostro_MT950_File As String

Public paramSwiftHisto_Input As String

Public paramSendMail_SMTPHost  As String
Public paramSendMail_From As String
Public paramSendMail_BIA_URL As String

Public paramIBM_BIA_Auto As String
Public paramIBM_BIA_ODBC As String

Public paramODBC_DSN_SAB As String
Public paramODBC_DSN_SAB073Y As String
Public paramODBC_DSN_JRN As String

Public paramODBC_DSN_CHQ_SCAN_ARCHIVE As String
Public paramODBC_DSN_CHQ_SCAN_LOCAL As String

Public paramODBC_DSN_SQL_Server_BIA As String
Public paramODBC_DSN_SQL_Server_BIA_VM As String
Public paramODBC_SideEUPLAB0 As String

Public paramODBC_DSN_SIDE_DB As String

'-----------------------------------------------------------------------------
Type typeCV
    DeviseIso     As String * 3
    DeviseN       As String * 3
    DeviseLibellé As String * 20
    Cours         As Double
    CoursAmj      As String * 8
    maxD          As String * 1
    EuroIn        As Boolean
    CotationCertain As Boolean
    
    Montant       As Currency
    OpéAmj        As String * 8
    CoursAmjMin   As String * 8
    AchatVente    As String * 1
    Normal        As String * 1
    CoursCompta   As String * 1

End Type


Public frmRTF_blnA5 As Boolean
Public frmRTF_UsrId_Origine As String
Public frmRTF_Référence As String

Public blnSAB_Migration As Boolean

Public arrBanqueIslamique() As String, arrBanqueIslamique_Nb As Integer
Public blnBanqueIslamique_Loop As Boolean

Public paramList_Height As Integer
Public arrBiapgm(100) As String

Public blnExplorer_IFS As Boolean

Public arrMNURUTUTI() As String, arrMNURUTUTI_Nb As Integer

Public arrUSR_UTI() As String, arrUSR_Mail() As String, arrUSR_Mail_Nb As Integer, arrUSR_Mail_UCase() As String

Public arrBIA_RCOM_Code(100) As String, arrBIA_RCOM_Lib(100) As String

Public paramCDO_Dossier_Path As String, paramCDO_Dossier_Path_DROPI As String
Public paramRDE_Dossier_Path As String, paramRDE_Dossier_Path_DROPI As String
Public paramGSOP_Dossier_Path As String, paramGSOP_Dossier_Path_DROPI As String

Public paramZSCHCRO0_SPLF As String, paramYCREANO0 As String
Public paramCPT_SCHEMA_Dossier_Path As String

Public arrMail_K1() As String, arrMail_K2() As String, arrMail_Memo() As String, arrMail_Nb As Integer
Public blnBIA_SSI_Automate As Boolean

Public collection_IMP() As String




Public Function fctCLIENACLI(lCLIENACLI As String, llstCLIENACLI() As Long) As Boolean
'DR 24/09/2013
Dim nClienaCli As Long
Dim I As Long

    fctCLIENACLI = False
    nClienaCli = Val(lCLIENACLI)
    If nClienaCli > 0 Then
        For I = 1 To llstCLIENACLI(0)
            If nClienaCli = llstCLIENACLI(I) Then
                fctCLIENACLI = True
                Exit For
            End If
        Next I
    End If
    
End Function


Public Function COMPTA_YBIAJOUR_OK()
Dim meYBIAMON0 As typeYBIAMON0
Dim V
On Error GoTo Exit_sub

App_Debug = "> COMPTA_YBIAJOUR_OK"
V = "?"

meYBIAMON0.MONAPP = "SMS"
meYBIAMON0.MONFLUX = "@YBIAJOUR"
V = rsYBIAMON0_Read(meYBIAMON0)

If IsNull(V) Then
    If Trim(meYBIAMON0.MONSTATUS) <> "" Then V = "BIAJOUR : statut " & Trim(meYBIAMON0.MONSTATUS) & " : " & meYBIAMON0.MONFILE
End If

'--------------------------------------------------------------------------------------
Exit_sub:

    COMPTA_YBIAJOUR_OK = V
    
End Function

Public Function DateComptableSuivanteP(lAMJ As String) As String
Dim V, xSql As String
Dim X10 As String * 10
On Error GoTo Error_Handler

App_Debug = "> DateComptableSuivanteP"
'--------------------------------------------------------------------------------------
DateComptableSuivanteP = "00000000"

 xSql = "select * from " & paramIBM_Library_SAB & ".ZCOMHIS0" _
     & " where COMHISOLD = " & Val(lAMJ) - 19000000 & " and COMHISNUM = 390"
   
Set rsSab = cnsab.Execute(xSql)
If Not rsSab.EOF Then
        DateComptableSuivanteP = rsSab("COMHISNEW") + 19000000
End If
Exit Function

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug



End Function
Public Function DateComptableSuivanteR(lAMJ As String) As String
Dim V, xSql As String
Dim X10 As String * 10
On Error GoTo Error_Handler

App_Debug = "> DateComptableSuivanteR"
'--------------------------------------------------------------------------------------
DateComptableSuivanteR = "00000000"

 xSql = "select * from " & paramIBM_Library_SAB & ".ZCOMHIS0" _
     & " where COMHISOLD > " & Val(lAMJ) - 19000000 & " and COMHISNUM = 390 order by COMHISOLD"
   
Set rsSab = cnsab.Execute(xSql)
If Not rsSab.EOF Then
        DateComptableSuivanteR = rsSab("COMHISOLD") + 19000000
End If
Exit Function

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug



End Function
Public Function DateComptablePrecedente(lAMJ As String) As String
Dim V, xSql As String
Dim X10 As String * 10
On Error GoTo Error_Handler

App_Debug = "> DateComptableSuivante"
'--------------------------------------------------------------------------------------
DateComptablePrecedente = "00000000"

 xSql = "select * from " & paramIBM_Library_SAB & ".ZCOMHIS0" _
     & " where COMHISNEW = " & Val(lAMJ) - 19000000 & " and COMHISNUM = 390"
   
Set rsSab = cnsab.Execute(xSql)
If Not rsSab.EOF Then
        DateComptablePrecedente = rsSab("COMHISOLD") + 19000000
End If
Exit Function

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug



End Function

Public Function DateComptableJP0(lAMJ As String) As String
Dim X As String, xSql As String
DateComptableJP0 = lAMJ
X = dateFinDeMois(lAMJ)
If X <> lAMJ Then Exit Function

xSql = "select * from " & paramIBM_Library_SAB & ".ZCOMHIS0" _
     & " where COMHISOLD = " & Val(lAMJ) - 19000000 & " and COMHISNUM = 0"
   
Set rsSab = cnsab.Execute(xSql)
If rsSab.EOF Then
        DateComptableJP0 = DateComptablePrecedente(lAMJ)
End If

End Function
Public Sub CV_Calc(lK2 As String, CV1 As typeCV, CV2 As typeCV)
Dim xMemo As String, xAMJ As String
Dim dblMontant As Double

If CV1.Montant = 0 Then
    CV1.Cours = 0
    CV2.Montant = 0
    Exit Sub
End If

'$JPL 2012-12-17 calcul de la CV

Call sqlYBIATAB0_Read("PDC", CV1.DeviseIso, CV1.OpéAmj, xMemo)
If IsNumeric(Mid$(xMemo, 9, 15)) Then
    CV1.Cours = CDbl(Mid$(xMemo, 9, 15) / 1000000000)
'If IsNull(rsYBIATAB0_Read("FIXING", CV1.DeviseIso, lK2, wBIATABTXT)) Then
'    CV1.Cours = CDbl(Mid$(wBIATABTXT, 9, 15)) / 1000000000
    dblMontant = Abs(CV1.Montant) / CV1.Cours
    CV2.Montant = Fix((dblMontant + 0.00500001) * 100) / 100
    If CV1.Montant < 0 Then CV2.Montant = -CV2.Montant
Else
    CV1.Cours = 0
    CV2.Montant = 999999999999.99
    'MsgBox "Manque cours :" & CV1.DeviseIso, vbCritical, "CV_CALC"
    
End If


End Sub

'20040830 jpl $$$$$$$$$$$$$$$$$ BIAS820I   ==> BIA_MAIN.bas et BIA_SAB.bas
'---------------------------------------------------------
Public Sub prtSocInit()
'---------------------------------------------------------
prtFormType = "SOC"

frmElpPrt.prtInit

If XPrt.PaperSize = vbPRPSA5 Then prtMaxY = 7200 '7300

prtSoc

End Sub

Public Sub fctPCEC_Atribut(lPCEC As String, lDev As String, blnCptOrdinaire As Boolean, blnRIB As Boolean, blnMédiateur As Boolean, blnIban As Boolean)
Dim X5 As String
blnCptOrdinaire = False: blnRIB = False: blnMédiateur = False: blnIban = False

X5 = Mid$(Trim(lPCEC), 1, 5)
If X5 = "11120" _
Or X5 = "12120" _
Or X5 = "12121" _
Or X5 = "12122" _
Or X5 = "25112" _
Or X5 = "25113" _
Or X5 = "25114" _
Or X5 = "25115" _
Or X5 = "25116" _
Or X5 = "25117" _
                    Then
    blnCptOrdinaire = True:
End If
If X5 = "25111" Then
    blnCptOrdinaire = True: blnMédiateur = True
End If

If blnCptOrdinaire Then
    blnIban = True
    If lDev = "EUR" Then blnRIB = True
End If

End Sub
Public Function fctUser_Classe_Aut(lClasse As Long) As Boolean

fctUser_Classe_Aut = True
If lClasse > 0 And lClasse < 100 Then
    Select Case Mid$(currentZMNUUTP0.MNUUTPCLA, lClasse, 1)
        Case "1", "2"
        Case Else: fctUser_Classe_Aut = False
    End Select
End If

End Function

 



Public Sub prtAdresse(lZADRESS0 As typeZADRESS0, blnPostal As Boolean)
Dim wADRESSPAY As String, blnADRESSPAY As Boolean
Dim wADRESSRA2 As String, blnADRESSRA2 As Boolean
'-----------------------encadrement petit tirets---------
Dim wCurrentX As Integer
wCurrentX = XPrt.CurrentX
XPrt.FontBold = True
XPrt.Print lZADRESS0.ADRESSRA1;

'-----------------------------------------------------
wADRESSRA2 = Trim(lZADRESS0.ADRESSRA2)
If wADRESSRA2 = "" Then
    blnADRESSRA2 = blnPostal
Else
    blnADRESSRA2 = True
End If

XPrt.FontBold = False
If blnADRESSRA2 Then
   XPrt.CurrentY = XPrt.CurrentY + 270
    XPrt.CurrentX = wCurrentX
    XPrt.Print wADRESSRA2;
End If
'-----------------------------------3---------------
If Trim(lZADRESS0.ADRESSAD1) <> "" Then
    XPrt.CurrentY = XPrt.CurrentY + 270
    XPrt.CurrentX = wCurrentX
    XPrt.Print lZADRESS0.ADRESSAD1;
End If
'----------------------------------4-------------------
If Trim(lZADRESS0.ADRESSAD2) <> "" Then
    XPrt.CurrentY = XPrt.CurrentY + 270
    XPrt.CurrentX = wCurrentX
    XPrt.Print lZADRESS0.ADRESSAD2;
End If

'-----------------------------------5------------------
If Trim(lZADRESS0.ADRESSAD3) <> "" Then
    XPrt.CurrentY = XPrt.CurrentY + 270
    XPrt.CurrentX = wCurrentX
    XPrt.Print lZADRESS0.ADRESSAD3;
End If
'------------------------------------6------------------
blnADRESSPAY = False
wADRESSPAY = Trim(lZADRESS0.ADRESSPAY)
If blnPostal Then
    If wADRESSPAY = "" Or wADRESSPAY = "FRANCE" Then
        XPrt.CurrentY = XPrt.CurrentY + 270
    Else
        blnADRESSPAY = True
    End If
End If
If Trim(lZADRESS0.ADRESSCOP) <> "" _
Or Trim(lZADRESS0.ADRESSVIL) <> "" Then
    XPrt.CurrentY = XPrt.CurrentY + 270
    XPrt.CurrentX = wCurrentX
    If Trim(lZADRESS0.ADRESSCOP) <> "" Then XPrt.Print lZADRESS0.ADRESSCOP & "  ";
    XPrt.Print lZADRESS0.ADRESSVIL;
End If
'------------------------------------8------------------
If blnADRESSPAY Then
    XPrt.FontBold = True
    XPrt.CurrentY = XPrt.CurrentY + 270
    XPrt.CurrentX = wCurrentX
    XPrt.Print wADRESSPAY;
    XPrt.FontBold = False
End If
'------------------------------------------

End Sub

Public Sub prtAdresse_Enveloppe(lZADRESS0 As typeZADRESS0)
'-----------------------encadrement petits tirets---------

XPrt.Line (5600, 2300)-(5700, 2300), prtLineColor
XPrt.Line (5600, 2300)-(5600, 2400), prtLineColor

XPrt.Line (10900, 2300)-(11000, 2300), prtLineColor
XPrt.Line (11000, 2300)-(11000, 2400), prtLineColor

XPrt.Line (5600, 4300)-(5700, 4300), prtLineColor
XPrt.Line (5600, 4200)-(5600, 4300), prtLineColor

XPrt.Line (10900, 4300)-(11000, 4300), prtLineColor
XPrt.Line (11000, 4200)-(11000, 4300), prtLineColor
XPrt.CurrentY = 2400
XPrt.CurrentX = 5700
prtAdresse lZADRESS0, True
End Sub



'---------------------------------------------------------
Public Sub prtSAB_Compta_Mt(lMonTant As Currency, lcolDb As Integer, lcolCr As Integer)
'---------------------------------------------------------
Dim X As String

XPrt.FontBold = True
X = Format$(Abs(lMonTant), "## ### ### ### ### ##0.00")
XPrt.CurrentX = IIf(lMonTant < 0, lcolCr, lcolDb) - 100 - XPrt.TextWidth(X)
XPrt.Print X;
XPrt.FontBold = False

End Sub



Public Function cnSAB_Transaction(lFct As String)
'On Error GoTo Error_Handler

'blnOff_Line = SAB073Y > ACCESS 2000 > ADO
'sinon ODBC AS400 DB400

cnSAB_Transaction = Null
If blnOff_Line Then
    Select Case lFct
        Case "BeginTrans": cnSab_Update.Open paramODBC_DSN_SAB
                           cnSab_Update.BeginTrans
                           Call FEU_ORANGE
        Case "Commit": cnSab_Update.CommitTrans
                       cnSab_Update.Close: Set cnSab_Update = Nothing
                       cnsab.Close: cnsab.Open paramODBC_DSN_SAB
                       Call FEU_VERT
        Case "Rollback": cnSab_Update.RollbackTrans
                         cnSab_Update.Close: Set cnSab_Update = Nothing
                       cnsab.Close: cnsab.Open paramODBC_DSN_SAB
                       Call FEU_VERT
    End Select
Else
    Select Case lFct
        Case "BeginTrans":
            cnSab_Update.Open paramODBC_DSN_SAB
            Set rsSab_Update = cnSab_Update.Execute("SET TRANSACTION ISOLATION LEVEL READ COMMITTED")
            Call FEU_ORANGE
        Case "Commit", "Rollback":
                Set rsSab_Update = cnSab_Update.Execute(lFct)
            cnSab_Update.Close
            Set cnSab_Update = Nothing
            Set rsSab_Update = Nothing
            Call FEU_VERT

    End Select

End If
Exit Function

Error_Handler:

cnSAB_Transaction = Error
MsgBox Error, vbCritical, frmElp_Caption & App_Debug

End Function

Public Function Table_Ope_Unit(lK2 As String) As String
On Error GoTo Error_Handler
Dim V, wK2 As String, wUnit As String, X As String
If Mid$(lK2, 5, 1) = "*" Then
    wK2 = Mid$(lK2, 1, 6)
Else
    wK2 = lK2
End If

V = rsElpTable_Read("SAB_Param", "Ope_Unit", wK2, X, wUnit)
If Not IsNull(V) Then
    V = rsElpTable_Read("SAB_Param", "Ope_Unit", Mid$(wK2, 1, 4), X, wUnit)
    If Not IsNull(V) Then wUnit = Mid$(wK2, 1, 4)
End If

Table_Ope_Unit = wUnit

Exit Function
'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & "Table_Ope_Unit : " & lK2
End Function

Public Function Table_Ope_Unit_RDE(lMOUVEMOPE As String, lMOUVEMNUM As Long, cnAdo As ADODB.Connection, rsADO As ADODB.Recordset) As String
Dim xSql As String, X As String
' cas particulier RDE/RDI : affecter le service initiateur SOBF / (SOBI par défaut)
Table_Ope_Unit_RDE = "SOBI"
xSql = "Select ENCCARNAT from " & paramIBM_Library_SAB & ".ZENCCAR0" & " where ENCCARCOP = '" & lMOUVEMOPE & "' and  ENCCARDOS = " & lMOUVEMNUM
Set rsADO = cnAdo.Execute(xSql)
If Not rsADO.EOF Then
    X = rsADO("ENCCARNAT")
    If X = "CHQ" Or X = "EFF" Then Table_Ope_Unit_RDE = "SOBF"
End If

End Function


Public Function paramSAA_Init()
Dim V
On Error GoTo Error_Handler

paramSAA_Init = Null

paramSAA_Data_Archive = paramServer("\\SWIFT\" & paramEnvironnement & "\" & constArchive)
paramSAA_DataF_Archive = paramSAA_Data_Archive & "\"

paramSAA_DataF_Log = paramServer("\\SWIFT\" & paramEnvironnement & "\" & constLog) & "\"

paramSAA_Data_from_SAB = paramServer("\\SWIFT\" & paramEnvironnement & "\SAA_from_SAB")
paramSAA_DataF_from_SAB = paramSAA_Data_from_SAB & "\"

paramSAA_Data_from_SAB_ExtensionP_sab = ".sab"
paramSAA_Data_from_SAB_ExtensionP_pcc = ".pcc"
paramSAA_Data_from_SAB_ExtensionP_rje = ".rje"
paramSAA_Data_from_SAB_ExtensionP_sav = ".sav"
paramSAA_Data_from_SAB_YFile = "ZSWIALL0"

paramSAA_Data_to_SAB = paramServer("\\SWIFT\" & paramEnvironnement & "\SAA_to_SAB")
paramSAA_DataF_to_SAB = paramSAA_Data_to_SAB & "\"

paramSAA_Data_to_SAB_ExtensionP_out = ".out"
paramSAA_Data_to_SAB_ExtensionP_sav = ".sav"
paramSAA_Data_to_SAB_YFile = "YSWIRAL0"


paramSAA_Data_from_MT950 = paramServer("\\SWIFT\" & paramEnvironnement & "\SAA_from_MT950")
paramSAA_DataF_from_MT950 = paramSAA_Data_from_MT950 & "\"

paramCorona_DataF_Swift_In = paramServer("\\Corona\Corona_Swift_In\")

paramSAA_Data_to_Corona = paramServer("\\SWIFT\" & paramEnvironnement & "\SAA_to_Corona")
paramSAA_DataF_to_Corona = paramSAA_Data_to_Corona & "\"

paramSAA_Data_to_Corona_ExtensionP_out = ".out"
paramSAA_Data_to_Corona_ExtensionP_sav = ".sav"

'paramSwift_BIC_Input = paramServer("\\DOCSRV\Install.bia\Swift Alliance Access\BIC Directory\2004_12\FI.dat")
'paramSwift_BIC_Input = paramServer("\\BIADOCSRV\install.bia\_INFORMATIQUE\Swift\BIC DIRECTORY\2007-**\WFI.dat")
paramSwift_BIC_Input = "\\.......\install.bia\_INFORMATIQUE\Swift\BIC DIRECTORY\2007-**\WFI.dat"
'========================================================================================
'Call lstErr_Clear(frmSAA.lstErr, frmSAA.cmdContext, "BIA.mdb : table : " & recElpTable.ID & ": ok ")


Exit Function
'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V & " " & Now, vbCritical, frmElp_Caption & App_Debug
    paramSAA_Init = V
End Function


'---------------------------------------------------------
Public Sub prtSoc()
'---------------------------------------------------------
Dim X As String, I As Integer
Dim prtX As Integer
Dim mprtFontSize As Integer
Dim mprtFontName As String
Dim wX As Long

wX = prtMinX + 7800
'If XPrt.Orientation = vbPRORLandscape And XPrt.PaperSize = vbPRPSA4 Then prtSoc_Lansdcape: Exit Sub
If XPrt.Orientation = vbPRORLandscape And XPrt.PaperSize = vbPRPSA4 Then wX = prtMinX + 11000
mprtFontSize = XPrt.FontSize
mprtFontName = XPrt.FontName

I = frmElpPrt.imgSocLogo.Width * 0.17
XPrt.PaintPicture frmElpPrt.imgSocLogo.Picture _
                , wX, 10 _
                , I _
                , frmElpPrt.imgSocLogo.Height * 0.17
'$JPL 20060829 , prtMinX + (prtMaxX - prtMinX - I) / 2, 10
                

XPrt.CurrentY = prtMinY
XPrt.FontBold = True
XPrt.FontSize = 11
prtX = 2500
'frmElpPrt.prtCentré prtX, paramSOC_RS
'-----------------------------------------------------
XPrt.FontSize = 9
XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + 250
'frmElpPrt.prtCentré prtX, paramSOC_Adresse
'------------------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + 250
'frmElpPrt.prtCentré prtX, paramSOC_Ville
'--------------------------------------------------------
I = frmElpPrt.imgSocLogo_PiedPage.Width
If Printer.ColorMode = 2 Then
    If XPrt.PaperSize = vbPRPSA5 Then
        XPrt.PaintPicture frmElpPrt.imgSocLogo_PiedPage.Picture _
                        , prtMinX, prtMaxY _
                        , I _
                        , frmElpPrt.imgSocLogo_PiedPage.Height ' * 0.13
    Else
        XPrt.PaintPicture frmElpPrt.imgSocLogo_PiedPage.Picture _
                        , prtMinX, prtMaxY + 150 _
                        , I _
                        , frmElpPrt.imgSocLogo_PiedPage.Height ' * 0.13
    End If
End If

' TODO XPrt.CurrentY = prtMaxY
If XPrt.PaperSize = vbPRPSA5 Then
    XPrt.CurrentY = prtMaxY + 150
Else
    XPrt.CurrentY = prtMaxY + 150 + 150
End If
XPrt.FontSize = 7
XPrt.FontName = "Calibri"
XPrt.ForeColor = prtForeColor
X = paramSOC_RS & " - " & paramSOC_Adresse & " - " & paramSOC_Ville
XPrt.CurrentX = prtMinX + I + 100: XPrt.Print X;
'frmElpPrt.prtCentré prtMedX, X
XPrt.CurrentY = XPrt.CurrentY + 140 'TODO 250
XPrt.CurrentX = prtMinX + I + 100: XPrt.Print paramSOC_Télécom;
'frmElpPrt.prtCentré prtMedX, paramSOC_Télécom
XPrt.CurrentY = XPrt.CurrentY + 140 'TODO 250
XPrt.CurrentX = prtMinX + I + 100: XPrt.Print paramsoc_capital;
'frmElpPrt.prtCentré prtMedX, paramsoc_capital

XPrt.CurrentY = prtMinY + prtHeaderHeight + prtlineHeight
XPrt.FontSize = mprtFontSize
XPrt.FontName = mprtFontName


End Sub

'---------------------------------------------------------
Public Sub prtSoc_Lansdcape()
'---------------------------------------------------------
Dim X As String, I As Integer
Dim prtX As Integer
Dim mprtFontSize As Integer
mprtFontSize = XPrt.FontSize


I = frmElpPrt.imgSocLogo.Width * 0.15
XPrt.PaintPicture frmElpPrt.imgSocLogo.Picture _
                , 3000 - I / 2, prtMinY - 100 _
                , I _
                , frmElpPrt.imgSocLogo.Height * 0.15

                
XPrt.CurrentY = prtMinY + 300
XPrt.FontBold = True
XPrt.FontSize = 11
prtX = 5000
'frmElpPrt.prtCentré prtX, paramSOC_RS
'-----------------------------------------------------
XPrt.FontSize = 9
XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + 750
frmElpPrt.prtCentré 3000, paramSOC_Adresse & "   " & paramSOC_Ville
'------------------------------------------------------
'XPrt.CurrentY = XPrt.CurrentY + 250
'frmElpPrt.prtCentré prtX, paramSOC_Ville
'--------------------------------------------------------
XPrt.FontSize = 7
XPrt.CurrentY = XPrt.CurrentY + 300
frmElpPrt.prtCentré 3000, paramsoc_capital
XPrt.CurrentY = XPrt.CurrentY + 250
frmElpPrt.prtCentré 3000, paramSOC_Télécom
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
mprtFontSize = XPrt.FontSize


End Sub


Public Sub prtSocMini(mCurrenty As Integer, AMJ As String)
Dim mprtFontSize As Integer, I As Integer

mprtFontSize = XPrt.FontSize
I = frmElpPrt.imgSocLogo.Width * 0.17
XPrt.PaintPicture frmElpPrt.imgSocLogo.Picture _
                , prtMinX + (prtMaxX - prtMinX - I) / 2, 10 _
                , I _
                , frmElpPrt.imgSocLogo.Height * 0.17


'----------------------------------------------------
XPrt.CurrentY = mCurrenty + prtMinY
XPrt.FontBold = True
XPrt.FontSize = 8
frmElpPrt.prtCentré 2000, paramSOC_RS
'-----------------------------------------------------
XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + 250
frmElpPrt.prtCentré 2000, paramSOC_Adresse
'------------------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + 250
frmElpPrt.prtCentré 2000, paramSOC_Ville
'--------------------------------------------------------

''XPrt.CurrentY = XPrt.CurrentY + 250 * 2
''XPrt.CurrentX = 7500
''XPrt.Print "Paris, le " & dateImp_jjMoisAAAA(AMJ);
XPrt.FontSize = mprtFontSize

End Sub

'---------------------------------------------------------
Public Function BiaPgm_Init()
'---------------------------------------------------------
Dim X As String, V, I As Integer
Dim xName As String, xK2 As String, xK2x As String
Dim H As Long

On Error GoTo Error_Handler

App_Debug = "> BiaPgm_Init : applications autorisées pour " & Elp.usrId
'--------------------------------------------------------------------------------------
BiaPgm_Init = Null

usrSituationCompte_Forçage = False
usrService_DisplayAll = False
            
XListBox.Clear
XListBox.Visible = True
paramList_Height = 245
'XListBox.Height = paramList_Height
XLabel.Visible = True
XLabel.Caption = "Menu"

X = "select * from ElpTable where SNN = 0" _
    & " and id = '" & paramBiaPgmAut & "'" _
    & " and K1 = '" & Elp.usrId & "'"
    
Set rsMDB = cnMDB.Execute(X)
Do While Not rsMDB.EOF
    xK2 = rsMDB("K2")
    xK2x = Trim(xK2)
    If xK2x = "$usr_Forçage" Then usrSituationCompte_Forçage = True
    If xK2x = "$usr_Service" Then usrService_DisplayAll = True
    
    V = rsElpTable_Read(paramBiaPgm, xK2, "", xName, X)
    If xK2x = "XUsrId" Or xK2x = "X_I5A7" Then
        X = Space$(50)
        Mid$(X, 21, 20) = frmElp_Caption
    End If
    If frmElp_Caption = Trim(Mid$(X, 21, 20)) Then
        If Not blnSAB_Migration Then
            XListBox.AddItem xK2 & vbTab & vbTab & xName
        Else
            If Mid$(X, 10, 1) = "X" Then XListBox.AddItem xK2 & " I5A7"  'vbTab & xName
        End If
    End If
        
    rsMDB.MoveNext
Loop

XListBox.AddItem "X_Reset     " & vbTab & vbTab & "réplication BiaSrv"

'Elp_ResizeControl XListBox
H = paramList_Height + paramList_Height * XListBox.ListCount
'XListBox.Height = IIf(8000 < H, 8000, H)
For I = 0 To XListBox.ListCount - 1
    XListBox.ListIndex = I
    arrBiapgm(I) = XListBox.Text
Next I
XListBox.ListIndex = -1
'==========================================================================
Exit Function

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
    BiaPgm_Init = V
End Function
Public Sub X_I5A7_Show()

blnSAB_Migration = True

elpSrvXcom = ""
mainSoc
elpSrvXcom = "XXXX"

End Sub

Public Sub lstZMNURUT0_Load(lstX As ListBox)
Dim V, xSql As String
Dim X10 As String * 10
On Error GoTo Error_Handler

App_Debug = "> lstZMNURUT0_Load"
'--------------------------------------------------------------------------------------

lstX.Clear
 xSql = "select * from " & paramIBM_Library_SAB & ".ZMNURUT0" _
     & " where MNURUTLOG = 'O'"
   
Set rsSab = cnsab.Execute(xSql)
Do Until rsSab.EOF
    X10 = rsSab("MNURUTUTI")
    lstX.AddItem X10 & vbTab & rsSab("MNURUTNOM")
    rsSab.MoveNext
Loop
Exit Sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug

End Sub
Public Sub cboZMNURUT0_Load_Prod(cboX As ComboBox)
Dim V, xSql As String
Dim X10 As String * 10
On Error GoTo Error_Handler

App_Debug = "> lstZMNURUT0_Load"
'--------------------------------------------------------------------------------------

cboX.Clear
 xSql = "select MNURUTUTI,MNUUTIGR2 from " & paramIBM_Library_SABSPE & ".ZMNURUT0 , " & paramIBM_Library_SAB & ".ZMNUUTI0" _
     & " where MNURUTLOG = 'O' and MNUUTICUT = MNURUTCUT"
   
Set rsSab = cnsab.Execute(xSql)
Do Until rsSab.EOF
    If Trim(rsSab("MNUUTIGR2")) <> "G_MIN" Then
        cboX.AddItem rsSab("MNURUTUTI")
    Else
        cboX.AddItem "* " & rsSab("MNURUTUTI")
    End If
    
    rsSab.MoveNext
Loop
Exit Sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug

End Sub

Public Sub lstZMNURUT0_Load_Actif(lstX As ListBox)
Dim V, xSql As String
Dim X10 As String * 10
On Error GoTo Error_Handler

App_Debug = "> lstZMNURUT0_Load_Actif"
'--------------------------------------------------------------------------------------

lstX.Clear
 xSql = "select MNURUTUTI,MNURUTNOM,MNUUTIGR2 from " & paramIBM_Library_SAB & ".ZMNURUT0 , " & paramIBM_Library_SAB & ".ZMNUUTI0" _
     & " where MNURUTLOG = 'O' and MNUUTICUT = MNURUTCUT"
   
Set rsSab = cnsab.Execute(xSql)
Do Until rsSab.EOF
    If Trim(rsSab("MNUUTIGR2")) <> "G_MIN" Then
        X10 = rsSab("MNURUTUTI")
        lstX.AddItem X10 & " " & vbTab & rsSab("MNURUTNOM")
    End If
        rsSab.MoveNext
Loop
Exit Sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug

End Sub

Public Sub lstZMNURUT0_Load_Actif_Production(lstX As ListBox)
Dim V, xSql As String
Dim X10 As String * 10
On Error GoTo Error_Handler

'--------------------------------------------------------------------------------------
'$JPL 20130916 voir : YSSIUSR0_Actif_Load
'--------------------------------------------------------------------------------------

App_Debug = "> lstZMNURUT0_Load_Actif"
'--------------------------------------------------------------------------------------

lstX.Clear
 xSql = "select MNURUTUTI,MNURUTNOM,MNUUTIGR2 from " & paramIBM_Library_SAB_P & ".ZMNURUT0 , " & paramIBM_Library_SAB_P & ".ZMNUUTI0" _
     & " where MNURUTLOG = 'O' and MNUUTICUT = MNURUTCUT"
   
Set rsSab = cnsab.Execute(xSql)
Do Until rsSab.EOF
    If Trim(rsSab("MNUUTIGR2")) <> "G_MIN" Then
        X10 = rsSab("MNURUTUTI")
        lstX.AddItem X10 & " " & vbTab & rsSab("MNURUTNOM")
    End If
        rsSab.MoveNext
Loop
Exit Sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug

End Sub

Public Sub YSSIUSR0_Actif_Load(lstX As ListBox)
Dim V, xSql As String
Dim X10 As String '* 10
On Error GoTo Error_Handler

App_Debug = "> YSSIUSR0_Actif_Load"
'--------------------------------------------------------------------------------------

lstX.Clear
xSql = "select * from " & paramIBM_Library_SABSPE & ".YSSIDOM0 , " & paramIBM_Library_SABSPE & ".YSSIWIN0" _
     & " where SSIDOMNAT = ' ' and SSIDOMDIDX = 'WIN' and SSIDOMUNIT <> '' and SSIDOMPRFK <> 'X'" _
     & " and SSIWINNAT = ' ' and SSIWINUIDD = SSIDOMUIDD"
   
Set rsSab = cnsab.Execute(xSql)
Do Until rsSab.EOF
        X10 = rsSab("SSIDOMUIDX")
        lstX.AddItem X10 & " " & vbTab & rsSab("SSIWINUNOM")
        rsSab.MoveNext
Loop
Exit Sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug

End Sub

'---------------------------------------------------------
Public Sub mainSoc()
'---------------------------------------------------------
Dim V, xName As String, xMemo As String, X As String
Dim wIBM_Library_X2 As String
Dim xElpTable As typeElpTable
On Error GoTo Error_Handler

App_Debug = "> MainSoc : Param"
Call lstErr_Clear(frmElp.lstErr, frmElp.cmdContext, App_Debug)
'--------------------------------------------------------------------------------------
mainSocExe

Set XListBox = frmElp.lstMain
Set XLabel = frmElp.lblMain

'--------------------------------------------------------------------------------------

V = rsElpTable_Read("Param", "BiaPgm", "Programmes", xName, paramBiaPgm)
If Not IsNull(V) Then GoTo Error_MsgBox

V = rsElpTable_Read("Param", "BiaPgm", "Autorisation", xName, paramBiaPgmAut)
If Not IsNull(V) Then GoTo Error_MsgBox

V = rsElpTable_Read("Param", "XUsrId", Trim(Elp.usrId), xName, X)

If IsNull(V) Then
    usrId = UCase$(X)
    Elp.usrId = usrId
    Xcom_UsrId usrId
    usrName_UCase = usrId
End If

App_Debug = "> MainSoc : applications autorisées"
Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, App_Debug)
'--------------------------------------------------------------------------------------
XListBox.Visible = False

If elpSrvXcom = "" Then
    XListBox.Clear
    XListBox.AddItem Elp.usrId
    SocId$ = "001"
    SocAgence$ = "001"
End If

If XListBox.ListCount = 1 Then
    XListBox.ListIndex = 0
    X = Space$(100)
    X = Trim(XListBox.Text)
    Elp.usrId = Trim(X) 'Mid$(x, 1, 10)
    usrName = Elp.usrId ' mId$(X, 17, 34)
    usrName_UCase = UCase(usrName)
    usrDRH usrIdNT
Else
    XLabel.Caption = "Qui êtes-vous ?"
    XLabel.ForeColor = errUsr.ForeColor
    XLabel.Visible = True
    XListBox.Visible = True
End If

App_Debug = "> MainSoc : paramEnvironnement"
Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, App_Debug)
'--------------------------------------------------------------------------------------
If Mid$(Elp.SrvDtaqIn, 1, 2) = "PC" Then
    paramBic8 = "BIARFRPP"
    paramEnvironnement = constProduction
    paramIBM_QSYSOPR = "BIA_INFO"
    paramIBM_BIA_Auto = "BIA_AUTO"
    paramIBM_BIA_ODBC = "BIA_ODBC"
    wIBM_Library_X2 = "P_"
Else
    paramBic8 = "BIARFRP0"
    paramEnvironnement = constTest
    paramIBM_QSYSOPR = "T_BIA_INFO"
    paramIBM_BIA_Auto = "T_BIA_AUTO"
    paramIBM_BIA_ODBC = "T_BIA_ODBC"
    wIBM_Library_X2 = "T_"
    socName = Elp.SrvDTaqOut
    focusUsr.BackColor = vbYellow
    MouseMoveUsr.ForeColor = vbYellow
    MouseMoveUsr.BackColor = vbYellow
    strSocSignon = paramFolder_Local & "\BiaTEST.bmp"
    frmElp.imgSocSignon.Picture = LoadPicture(strSocSignon)
End If

'$JPL 2013-02-12 _____________________________________________________________
paramFolder_Master = paramServer("\\BiaSrv\")
DataBase_Master = UCase$(paramFolder_Master & "\BIA_SAB.mdb")
'$JPL 2013-02-12 _______________________________________________________________

paramYBase_Data = paramServer("\\YBASE\" & paramEnvironnement)
paramYBase_DataF = paramYBase_Data & "\"
paramYBase_Data_Extension = "txt"
paramYBase_Data_ExtensionP = "." & paramYBase_Data_Extension

paramFTP_Out = paramServer("\\#S820I_Out\" & paramEnvironnement)
paramFTP_OutF = paramFTP_Out & "\"
paramFTP_Out_XCom = paramServer("\\#S820I_Out\" & paramEnvironnement & "\" & constXCom)
paramFTP_Out_XComF = paramFTP_Out_XCom & "\"

paramFTP_In = paramServer("\\#S820I_In\" & paramEnvironnement)
paramFTP_InF = paramFTP_In & "\"

paramFTP_SPLF = paramServer("\\#S820I_Out\SPLF")

paramPeliNT_Data = paramServer("\\PELINT\" & paramEnvironnement)
paramPeliNT_DataF = paramPeliNT_Data & "\"

paramPeliNT_Monitor = paramPeliNT_Data & "\Monitor"
paramPeliNT_MonitorF = paramPeliNT_Monitor & "\"

paramPeliNT_Exe = paramPeliNT_Data & "\Exe"
paramPeliNT_ExeF = paramPeliNT_Exe & "\"

paramPeliNT_Aller_XcomF = paramPeliNT_Data & "\" & constAller & "\" & constXCom & "\"
paramPeliNT_Aller_FTPF = paramPeliNT_Data & "\" & constAller & "\" & constFTP & "\"
paramPeliNT_Retour_XcomF = paramPeliNT_Data & "\" & constRetour & "\" & constXCom & "\"
paramPeliNT_Retour_FTPF = paramPeliNT_Data & "\" & constRetour & "\" & constFTP & "\"
paramPeliNT_Connexion_EnCours = "Connexion_EnCours.bat"

localPeliNT_DataF = "C:\Pelint.dat\" & paramEnvironnement & "\"
If blnOff_Line Then localPeliNT_DataF = paramPeliNT_DataF

localPeliNT_MonitorF = localPeliNT_DataF & "Monitor\"
localPeliNT_ExeF = localPeliNT_DataF & "Exe\"

paramIBM_Init wIBM_Library_X2

paramSAA_Init

paramEdition_Init

'A8 pour tests sur I5A7
'-------------------------------
If blnSAB_Migration Then
    paramIBM_AS400_ID = "I5A7"
    paramIBM_ODBC_SAB = "I5A7"
    paramBic8 = "BIARFRP0"
    paramIBM_Library_SAB = "SAB073U"
    paramIBM_Library_SABSPE = "SAB073SPE"
    paramIBM_QSYSOPR = "BIA_INFO"
    paramIBM_BIA_Auto = "BIA_AUTO"
    paramIBM_BIA_ODBC = "BIA_ODBC"
    wIBM_Library_X2 = "P_"
    focusUsr.BackColor = vbGreen
    MouseMoveUsr.ForeColor = vbGreen
    MouseMoveUsr.BackColor = vbGreen
    frmElp.imgSocSignon.Picture = LoadPicture("C:\biasrv\I5A7.bmp")
    'paramEnvironnement = constTest
    ''JPL cas très particulier: paramEnvironnement = constProduction
End If

'Migration SQL2010_BIA
V = rsElpTable_Read("SIDE", "PasswordX", "SIDE_UPDATE", xName, xMemo)
paramODBC_DSN_SIDE_DB = "DSN=SIDE2010" & ";UID=SIDE_UPDATE" & "; PWD=" & xMemo
paramODBC_DSN_SQL_Server_BIA = "DSN=SQL2010_BIA" & ";UID=SIDE_UPDATE" & "; PWD=" & xMemo
paramODBC_DSN_SAB073Y = "SAB073Y"

paramODBC_DSN_SAB = "DSN=SABAT6;UID=P_QUALIOS; PWD=Bia$2031"
If paramEnvironnement = constProduction Then
    paramODBC_DSN_CHQ_SCAN_ARCHIVE = "DSN=CHQ_ARCHIVE"
    paramODBC_DSN_CHQ_SCAN_LOCAL = "DSN=CHQ_LOCAL"
Else
    paramODBC_DSN_CHQ_SCAN_ARCHIVE = "DSN=CHQ_ARCHIVE_TEST"
    paramODBC_DSN_CHQ_SCAN_LOCAL = "DSN=CHQ_LOCAL_TEST"
End If

paramODBC_SideEUPLAB0 = "SEPA"

paramEdition_Print_Production = True
V = rsElpTable_Read("Edition", "@PRINT_PROD", "Programmes", xName, xMemo)
If UCase$(Mid$(xMemo, 1, 1)) = "N" Then paramEdition_Print_Production = False
paramEdition_Print_Test = False
V = rsElpTable_Read("Edition", "@PRINT_TEST", "Programmes", xName, xMemo)
If UCase$(Mid$(xMemo, 1, 1)) = "O" Then paramEdition_Print_Test = True

App_Debug = "> MainSoc : param Test off-line"
'--------------------------------------------------------------------------------------

If blnOff_Line Then
    paramIBM_Library_SAB = "c:\biasrv\SAB073Y"
    paramIBM_Library_SAB_P = "c:\biasrv\SAB073Y"
    paramIBM_Library_SABSPE = "c:\biasrv\SAB073Y"
    paramIBM_Library_SABSPE_P = "c:\biasrv\SAB073Y"
    paramIBM_Library_BODWH = "c:\biasrv\SAB073Y"
    paramODBC_DSN_SAB = "DSN=SAB073Y" '& paramDataBase_Password
    paramODBC_DSN_SQL_Server_BIA = "DSN=SAB073Y"
    paramODBC_SideEUPLAB0 = "c:\biasrv\SAB073Y.SideEUPLAB0"
    paramSendMail_SMTPHost = "smtp.xxxxxx"
    paramSendMail_From = "inconnu@xxxxxx"
    paramSendMail_BIA_URL = "@xxxxxx"
End If

App_Debug = "> MainSoc : Chargement des tables SAB"
Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, App_Debug)
'--------------------------------------------------------------------------------------
If cnsab.State = adStateOpen Then cnsab.Close
cnsab.CommandTimeout = 0

On Error Resume Next
cnsab.Open paramODBC_DSN_SAB
If Trim(Error) <> "" Then
    If paramIBM_AS400_ID = "I5A7" Then
        paramODBC_DSN_SAB = "DSN=SAB073" & ";UID=" & paramIBM_BIA_ODBC & "; PWD=" & paramIBM_BIA_ODBC_Password
        cnsab.Open paramODBC_DSN_SAB
    Else
        GoTo Error_Handler
    End If
    
End If
On Error GoTo Error_Handler

mainSoc_AMJCPT_Load

If paramEnvironnement = constProduction Then
    If DSys <> YBIATAB0_DATE_CPT_JS1 Then
        If blnTimer_Enabled Then End
        Call MsgBox("Date comptable SAB : " & YBIATAB0_DATE_CPT_JS1 & " # Date système : " & DSys, vbCritical, "BIA.vb: initialisation")
    End If
End If
    
''20050320 mainSoc_YBase_Load  ' ! après lecture YBIATAB0

App_Debug = "> MainSoc : habilitations Utilisateur SAB"
Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, App_Debug)
'--------------------------------------------------------------------------------------
    
If blnOff_Line Then
        currentSSIWINUIDN = 1005
        currentSSIWINUIDD = 490
        currentSSIWINUIDX = "LOULERGUE"
        currentSSIWINUNOM = "LOULERGUE Jean Pierre"
        currentSSIWINUNIT = "S40"
        currentSSIWINMAIL = "loulergue.jp@bia-paris.fr"
        currentSSIWINUNIT_Lib = "Informatique"
Else
'----------------------------------------------------------------------------------
    X = "select * from " & paramIBM_Library_SABSPE & ".YSSIDOM0 ," & paramIBM_Library_SABSPE & ".YSSIWIN0" _
        & " where SSIDOMNAT = ' ' and SSIDOMDIDX = 'WIN'  and SSIDOMUIDX =  '" & usrName_UCase & "'" _
        & " and SSIWINNAT = ' ' and SSIWINUIDD = SSIDOMUIDD "
    Set rsSab = cnsab.Execute(X)
    If rsSab.EOF Then
        X = "select * from " & paramIBM_Library_SABSPE & ".YSSIDOM0 ," & paramIBM_Library_SABSPE & ".YSSIWIN0" _
            & " where SSIDOMNAT = ' ' and SSIDOMDIDX = 'WIN'  and SSIWININFO like  '%|" & usrName_UCase & "|%'" _
            & " and SSIWINNAT = ' ' and SSIWINUIDD = SSIDOMUIDD "
        Set rsSab = cnsab.Execute(X)
    End If
    
    If rsSab.EOF Then
        currentSSIWINUNIT = "S40"
        Call MsgBox("inconnu : " & X, vbExclamation, "BIA_mainSoc")
    Else
        currentSSIWINUIDN = rsSab("SSIDOMUIDN")
        currentSSIWINUIDD = rsSab("SSIDOMUIDD")
        currentSSIWINUIDX = Trim(rsSab("SSIDOMUIDX")): currentSSIWINUIDX_U = UCase(currentSSIWINUIDX)
        currentSSIWINUNOM = Trim(rsSab("SSIWINUNOM"))
        currentSSIWINUNIT = Trim(rsSab("SSIDOMUNIT"))
'JPL 2014-09-23 : conserver l'adresse mail d'origine pour les utilisateurs AIB_SWIFT (SALLE_BO, CULPIN_BO)
        'currentSSIWINMAIL = Trim(rsSab("SSIWINMAIL"))
        If Not blnBIA_VB_AIB Then currentSSIWINMAIL = Trim(rsSab("SSIWINMAIL"))
'__________________________________________________________________________________________
        X = "select *from " & paramIBM_Library_SABSPE & ".YSSIUSR0 where SSIUSRNAT = 'S' and SSIUSRUNIT = '" & currentSSIWINUNIT & "'"
        
        Set rsSab = cnsab.Execute(X)
            
        If rsSab.EOF Then
            currentSSIWINUNIT_Lib = currentSSIWINUNIT
        Else
            currentSSIWINUNIT_Lib = Trim(rsSab("SSIUSRUIDX"))
        End If
    End If
End If
'============================================================================================================
currentUser.Id = usrId
usrIdSAB = usrId
If blnOff_Line Then

Else
    X = "select * from " & paramIBM_Library_SABSPE & ".YSSIDOM0" _
        & " where SSIDOMNAT = ' ' and SSIDOMDIDX = 'SAB'  and SSIDOMUIDN = " & currentSSIWINUIDN _
        & " and SSIDOMUIDX =  '" & usrIdSAB & "'"
    
    Set rsSab = cnsab.Execute(X)
    
    If rsSab.EOF Then
        X = "select * from " & paramIBM_Library_SABSPE & ".YSSIDOM0" _
            & " where SSIDOMNAT = ' ' and SSIDOMDIDX = 'SAB'  and SSIDOMUIDN = " & currentSSIWINUIDN _
            & " and SSIDOMPRFX <> 'X'"
        Set rsSab = cnsab.Execute(X)
        If Not rsSab.EOF Then
            If Not blnBIA_VB_AIB Then usrIdSAB = Trim(rsSab("SSIDOMUIDX")): currentUser.Id = usrIdSAB
        End If
    End If
    If usrId = "SOBF_SCAN" Then currentUser.Id = "BIA_AUTO"
End If

    Call Table_User(currentUser)
    
    currentUnit.Id = currentUser.Unit: Call Table_Unit(currentUnit)
    
    currentZMNURUT0.MNURUTUTI = usrIdSAB 'currentUser.Id

If paramEnvironnement = constTest Then
    If Not blnOff_Line Then currentZMNURUT0.MNURUTUTI = "T_" & currentUser.Id
End If

If Not blnBIA_VB_AIB Then V = currentZMNU_Load
V = Null
If Not IsNull(V) Then
    MsgBox V & vbCrLf & "BYE BYE", vbCritical, frmElp_Caption & App_Debug
    End
End If
'============================================================================================================
paramIBM_Library_SABSPE_XXX = paramIBM_Library_SABSPE
App_Debug = "> MainSoc : Terminé"
Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, App_Debug)
'--------------------------------------------------------------------------------------
Set frmElp_Icon = frmElp_Icon.Icon
mainSoc_Display
mainSoc_BanqueIslamique
BIA_VB_APP '$JPL_HAB
'--------------------------------------------------------------------------------------
If blnOff_Line Then
    MsgBox "BIA MainSoc sabPays() à initialiser"
Else
    Call rsZBASTAB0_Pays(sabPays(), sabPays_NB)
    Dim K As Integer
    For K = 1 To sabPays_NB
        Select Case sabPays(K).Id
            Case "FR": sabPays_FR = K
            Case "DZ": sabPays_DZ = K
            Case "LY": sabPays_LY = K
            Case "US": sabPays_US = K
        End Select
    Next K
End If

Call arrBIA_RCOM_Load

paramCDO_Dossier_Path = paramServer("\\ROPDOS\") & "CREDOC\" & paramEnvironnement & "\"
paramCDO_Dossier_Path_DROPI = paramServer("\\ROPDOS_DROPI\" & "CREDOC\" & paramEnvironnement & "\")

paramRDE_Dossier_Path = paramServer("\\ROPDOS\") & "REMDOC\" & paramEnvironnement & "\"
paramRDE_Dossier_Path_DROPI = paramServer("\\ROPDOS_DROPI\" & "REMDOC\" & paramEnvironnement & "\")

paramGSOP_Dossier_Path = paramServer("\\ROPDOS\") & "GSOP\" & paramEnvironnement & "\"
paramGSOP_Dossier_Path_DROPI = paramServer("\\ROPDOS_DROPI\" & "GSOP\" & paramEnvironnement & "\")

paramZSCHCRO0_SPLF = paramEditionSplf_Folder & "SCHGE005P1\ZSCHCRO0_SPLF\"
paramYCREANO0 = paramEditionSplf_Folder & "SCHGE005P1\YCREANO0\"

paramCPT_SCHEMA_Dossier_Path = paramServer("\\ROPDOS\") & "CPT_SCHEMA\" & paramEnvironnement & "\"

'2013-06-10 délégation : destinataires des mails en copie
Call BIA_VB_Hab_Idem_Mail
'--------------------------------------------------------------------------------------
Exit Sub
'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
End Sub
Public Function currentZMNU_Load()

Dim xMemo As String, X As String, xSql As String
Dim V
On Error GoTo Error_Handler
currentZMNU_Load = Null

'-------------------------------------------------------
App_Debug = " currentZMNU_Load : ZMNURUT0"
'-------------------------------------------------------

xSql = "select * from " & paramIBM_Library_SAB & ".ZMNURUT0" _
     & " where MNURUTUTI ='" & currentZMNURUT0.MNURUTUTI & "'"
     
Set rsSab = cnsab.Execute(xSql)
If Not rsSab.EOF Then
    Call rsZMNURUT0_GetBuffer(rsSab, currentZMNURUT0)
Else
    V = " ? ZMNURUT0 : " & currentZMNURUT0.MNURUTUTI
    GoTo Error_MsgBox
End If

'-------------------------------------------------------
App_Debug = " currentZMNU_Load : ZMNUUTI0"
'-------------------------------------------------------

xSql = "select * from " & paramIBM_Library_SAB & ".ZMNUUTI0" _
     & " where MNUUTICUT = " & currentZMNURUT0.MNURUTCUT _
     & " and   MNUUTIETB = " & currentZMNURUT0.MNURUTETB _
     & " and   MNUUTIREF = 0"
     
Set rsSab = cnsab.Execute(xSql)
If Not rsSab.EOF Then
    Call rsZMNUUTI0_GetBuffer(rsSab, currentZMNUUTI0)
Else
    V = " ? ZMNUUTI0 : " & currentZMNURUT0.MNURUTUTI & " : " & currentZMNURUT0.MNURUTCUT
    GoTo Error_MsgBox
End If

'-------------------------------------------------------
App_Debug = " currentZMNU_Load : ZMNUHLB0"
'-------------------------------------------------------

xSql = "select * from " & paramIBM_Library_SAB & ".ZMNUHLB0" _
     & " where MNUHLBCLA = '3' and MNUHLBVAL = '1' and MNUHLBFID = 0" _
     & " and MNUHLBNOM = '" & currentZMNUUTI0.MNUUTIGR3 & "'" _
     & " and MNUHLBETB = " & currentZMNURUT0.MNURUTETB
     
Set rsSab = cnsab.Execute(xSql)
If Not rsSab.EOF Then
    Call rsZMNUHLB0_GetBuffer(rsSab, currentZMNUHLB0)
Else
    V = " ? ZMNUHLB0 : " & currentZMNURUT0.MNURUTUTI & " : " & currentZMNUUTI0.MNUUTIGR3
    GoTo Error_MsgBox
End If


'-------------------------------------------------------
App_Debug = " currentZMNU_Load : ZMNUUTP0"
'-------------------------------------------------------

xSql = "select * from " & paramIBM_Library_SAB & ".ZMNUUTP0" _
     & " where MNUUTPGRP = '" & currentZMNUUTI0.MNUUTIGR3 & "'" _
     & " and   MNUUTPREF = " & currentZMNUHLB0.MNUHLBREF _
     & " and   MNUUTPETB = " & currentZMNURUT0.MNURUTETB _
     & " and   MNUUTPAGE = " & currentZMNUUTI0.MNUUTIAGE

Set rsSab = cnsab.Execute(xSql)
If Not rsSab.EOF Then
    Call rsZMNUUTP0_GetBuffer(rsSab, currentZMNUUTP0)
Else
    V = " ? ZMNUUTP0 : " & currentZMNURUT0.MNURUTUTI & " : " & currentZMNUUTI0.MNUUTIGR3
    GoTo Error_MsgBox
End If
        
'-------------------------------------------------------
App_Debug = "currentZMNU_Load : Compte du personnel"
'-------------------------------------------------------
V = rsElpTable_Read("User", usrId, "CLIENASIG", X, xMemo)
If Not IsNull(V) Then
    currentCLIENASIG = usrId
Else
    currentCLIENASIG = xMemo
End If
        
        
        
currentSAB_ETA = currentZMNURUT0.MNURUTETB: currentSAB_PLA = 1: currentSAB_AGE = 1
Exit Function

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
    currentZMNU_Load = V
End Function

Public Sub BiaPgmAut_Init(lFct As String, recAut As typeAuthorization)
Dim xSql As String, K As Integer, X As String

recAut.Consulter = False
recAut.Saisir = False
recAut.Valider = False
recAut.Comptabiliser = False
recAut.Rapprocher = False
recAut.Swift = False
recAut.Virement = False
recAut.Avis = False
recAut.Xspécial = False

Call BIA_VB_HAB_Idem(lFct, usrName_UCase)

xSql = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
     & " where BIATABID = 'BIA_VB_HAB' and BIATABK1 = '" & UCase(Trim(lFct)) & "'and BIATABK2 = '" & idemUser.Id & "'"
Set rsSab = cnsab.Execute(xSql)

If Not rsSab.EOF Then
    X = rsSab("BIATABTXT")
    For K = 1 To 19
        If Mid$(X, K, 1) <> " " Then
            Select Case K
                Case 1: recAut.Consulter = True
                Case 2: recAut.Saisir = True
                Case 3: recAut.Valider = True
                Case 4: recAut.Comptabiliser = True
                Case 5: recAut.Rapprocher = True
                Case 6: recAut.Swift = True
                Case 7: recAut.Virement = True
                Case 8: recAut.Avis = True
                Case 9: recAut.X09 = True
                Case 10: recAut.X10 = True
                Case 11: recAut.X11 = True
                Case 12: recAut.X12 = True
                Case 13: recAut.X13 = True
                Case 14: recAut.X14 = True
                Case 15: recAut.X15 = True
                Case 16: recAut.X16 = True
                Case 17: recAut.X17 = True
                Case 18: recAut.X18 = True
               Case 19: recAut.Xspécial = True
            End Select
        End If
    Next K
End If


'Dim xName As String, xMemo As String
'Call rsElpTable_Memo(paramBiaPgmAut, Elp.usrId, X, xName, xMemo)

'recAut.Consulter = IIf(mId$(xMemo, 1, 1) = "X", True, False)
'recAut.Saisir = IIf(mId$(xMemo, 2, 1) = "X", True, False)
'recAut.Valider = IIf(mId$(xMemo, 3, 1) = "X", True, False)
'recAut.Comptabiliser = IIf(mId$(xMemo, 4, 1) = "X", True, False)
'recAut.Rapprocher = IIf(mId$(xMemo, 5, 1) = "X", True, False)
'recAut.Swift = IIf(mId$(xMemo, 6, 1) = "X", True, False)
'recAut.Virement = IIf(mId$(xMemo, 7, 1) = "X", True, False)
'recAut.Avis = IIf(mId$(xMemo, 8, 1) = "X", True, False)
'recAut.Xspécial = IIf(mId$(xMemo, 9, 1) = "X", True, False)

End Sub

Public Sub prtSocMiniFin()
XPrt.FontSize = 6
frmElpPrt.prtCentré (prtMaxX - prtMinX) / 2, paramSoc_Capital_Télécom
End Sub


Public Sub mainSoc_Environnement()

''D:\BiaSRc\Dta\BIAS820I_xx.ini
'£JPL 2013-02-01 initialisation dans mianSoc
'paramFolder_Master = "\\BIADOCSRV\.BiaSrv"
'_________________________________________________
paramFolder_Local = "C:\BiaSrv"

If Dir(paramFolder_Local & "\BiaSigno.bmp") = "" Then
    paramFolder_Local = "D:\BiaSrv"
    If Dir(paramFolder_Local & "\BiaSigno.bmp") = "" Then
        MsgBox "MANQUE :" & paramFolder_Local, vbCritical, "BIA_mainSoc_Environnement"
        ''End
    End If
End If
    

DataBase_Open = ""
DataBase_Local = UCase$(paramFolder_Local & "\BIA_SAB.mdb")

'$JPL 2013-02-12 __déplacé dans main_soc______________________________________
'DataBase_Master = UCase$(paramFolder_Master & "\BIA_SAB.mdb")
'$JPL 2013-02-12 _____________________________________________________________
DataBase_Data = UCase$(paramFolder_Local & "\SAB073Y.mdb")

Elp.SrvObj = "ELPDTAQ"
Elp.pcId = "FR"
Elp.SrvType = "AS400"
Elp.SrvId = paramIBM_AS400_ID
Elp.SrvDtaqLib = "BIADTAQ"
Elp.SrvDtaqIn = "PC000001"
Elp.SrvDTaqOut = "PC000000"
pcIdUsrIdCtl = False
strSocSignon = paramFolder_Local & "\BiaSigno.bmp"
'imgSocLogo = paramFolder_Local & "\BiaLogo.bmp"
paramSocLogo = paramFolder_Local & "\banqueBIA.bmp"
paramSocLogo_G = paramFolder_Local & "\banqueBIA_G.bmp"
paramSocLogo_PiedPage = paramFolder_Local & "\banqueBIA_PiedPage.bmp"
imgGuichet = paramFolder_Local & "\BiaGuichet.bmp"
prtFontName = prtFontName_Arial
prtFiligrane_Color = vbBlack

elpSrvXcom = "CAV4"

paramDataBase_Password = "l2206"
paramSOC_TVA_Intracommunautaire = "FR 87 302 590 070"

paramSOC_RS = "Banque BIA"
paramSOC_Adresse = "67, avenue Franklin D. Roosevelt"
paramSOC_Ville = "75008 PARIS"
paramsoc_capital = "S.A. au capital de 158 100 000 Euros - R.C.Paris B 302590070 - Code APE 6419Z - N° de TVA intracommunautaire : " & paramSOC_TVA_Intracommunautaire
paramSOC_Télécom = "Tél: 33 (0)1 53 76 62 62 - Fax: 33 (0)1 42 89 09 59 - Télex: 644 030 BIAPA - Swift: BIARFRPP "
paramSoc_Capital_Télécom = "S.A. au capital de 158 100 000 Euros - R.C.Paris B 302 590 070 - Code APE 6419Z -  Tél: 33 (0)1 53 76 62 62 - Téléfax: 33 (0)1 42 89 09 59 - Télex: 644 030 BIAPA - Swift: BIARFRPP "

socSiren = "302590070"
SocBicId = "BIARFRPP"
socName = "Banque BIA (Paris)"
strSocBdfE = "12179": strSocBdfG = "00001"
SocRibDom = "Banque BIA PARIS"
socTéléphone = "(33) 01 53 76 62 62"

paramSendMail_SMTPHost = "exchg2015"                                   ' Required the fist time, optional thereafter
'paramSendMail_SMTPHost = "exg2016a"  ' Modification du serveur EXchange  Kokou 18/10/2024
paramSendMail_BIA_URL = "@bia-paris.fr"
paramSendMail_From = "BIA_INFO" & paramSendMail_BIA_URL

End Sub

Public Sub usrDRH(lName)

End Sub



Public Sub paramIBM_Init(X2 As String)
Dim V, xName As String, xMemo As String
On Error GoTo Error_Handler

App_Debug = "> paramIBM_Init : " & X2
'--------------------------------------------------------------------------------------

'Ne pas utiliser PDFCreator si computerName=paramServerSplf=BIA2008
V = rsElpTable_Read("server", "Splf", "PDFCrea", xName, xMemo)
If Not IsNull(V) Then GoTo Error_MsgBox
paramServerSplf = paramServer(xMemo)

paramIBM_AS400_ID = "?"
paramIBM_Library_SAB = "?"
paramIBM_Library_SABSPE = "?"
paramIBM_Library_SABJRN = "?"
paramIBM_Library_File = "?"
paramIBM_Library_Src = "?"
paramIBM_Library_Obj = "?"

' 2008-10-15 JPL
'V = rsElpTable_Read("IBM", "ODBC", X2 & "SAB", xName, xMemo)
'If Not IsNull(V) Then GoTo Error_MsgBox
'paramIBM_ODBC_SAB = paramServer(xMemo)

V = rsElpTable_Read("IBM", "AS400", X2 & "ID", xName, xMemo)
If Not IsNull(V) Then GoTo Error_MsgBox
paramIBM_AS400_ID = paramServer(xMemo)

paramIBM_AS400_FTP = paramIBM_AS400_ID
paramIBM_ODBC_SAB = paramIBM_AS400_ID

V = rsElpTable_Read("IBM", "Library", X2 & "SAB", xName, xMemo)
If Not IsNull(V) Then GoTo Error_MsgBox
paramIBM_Library_SAB = paramServer(xMemo)

V = rsElpTable_Read("IBM", "Library", "P_" & "SAB", xName, xMemo)
If Not IsNull(V) Then GoTo Error_MsgBox
paramIBM_Library_SAB_P = paramServer(xMemo)
V = rsElpTable_Read("IBM", "Library", "P_" & "SABSPE", xName, xMemo)
If Not IsNull(V) Then GoTo Error_MsgBox
paramIBM_Library_SABSPE_P = paramServer(xMemo)


V = rsElpTable_Read("IBM", "Library", X2 & "SABSPE", xName, xMemo)
If Not IsNull(V) Then GoTo Error_MsgBox
paramIBM_Library_SABSPE = paramServer(xMemo)

V = rsElpTable_Read("IBM", "Library", X2 & "SABJRN", xName, xMemo)
If Not IsNull(V) Then GoTo Error_MsgBox
paramIBM_Library_SABJRN = paramServer(xMemo)

V = rsElpTable_Read("IBM", "Library", X2 & "BIADWH", xName, xMemo)
If Not IsNull(V) Then GoTo Error_MsgBox
paramIBM_Library_BIADWH = paramServer(xMemo)

V = rsElpTable_Read("IBM", "Library", X2 & "BODWH", xName, xMemo)
If Not IsNull(V) Then GoTo Error_MsgBox
paramIBM_Library_BODWH = paramServer(xMemo)

V = rsElpTable_Read("IBM", "Library", X2 & "FILE", xName, xMemo)
If Not IsNull(V) Then GoTo Error_MsgBox
paramIBM_Library_File = paramServer(xMemo)

V = rsElpTable_Read("IBM", "Library", X2 & "OBJ", xName, xMemo)
If Not IsNull(V) Then GoTo Error_MsgBox
paramIBM_Library_Obj = paramServer(xMemo)

V = rsElpTable_Read("IBM", "Library", X2 & "SRC", xName, xMemo)
If Not IsNull(V) Then GoTo Error_MsgBox
paramIBM_Library_Src = paramServer(xMemo)

V = rsElpTable_Read("IBM", "Library", X2 & "OBJ", xName, xMemo)
If Not IsNull(V) Then GoTo Error_MsgBox
paramIBM_Library_Obj = paramServer(xMemo)

V = rsElpTable_Read("IBM", "PasswordX", "BIA_INFO", xName, xMemo)
If Not IsNull(V) Then GoTo Error_MsgBox
paramIBM_BIA_INFO_Password = paramServer(xMemo)

V = rsElpTable_Read("IBM", "PasswordX", "BIA_INFO", xName, xMemo)
If Not IsNull(V) Then GoTo Error_MsgBox
paramIBM_BIA_INFO_Password = paramServer(xMemo)

V = rsElpTable_Read("IBM", "PasswordX", "BIA_AUTO", xName, xMemo)
If Not IsNull(V) Then GoTo Error_MsgBox
paramIBM_BIA_AUTO_Password = paramServer(xMemo)

V = rsElpTable_Read("IBM", "PasswordX", "BIA_ODBC", xName, xMemo)
If Not IsNull(V) Then GoTo Error_MsgBox
paramIBM_BIA_ODBC_Password = paramServer(xMemo)

V = rsElpTable_Read("IBM", "PasswordX", "BIA_DWH", xName, xMemo)
If Not IsNull(V) Then GoTo Error_MsgBox
paramIBM_BIA_DWH_Password = paramServer(xMemo)

V = rsElpTable_Read("IBM", "PasswordX", "BO_DWH", xName, xMemo)
If Not IsNull(V) Then GoTo Error_MsgBox
paramIBM_BO_DWH_Password = paramServer(xMemo)

Exit Sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
'==============================================

End Sub

Public Function Table_Unit(lUnit As typeUnit)

On Error GoTo Error_Handler

Dim V

Table_Unit = Null
V = rsElpTable_Read("Unit", lUnit.Id, "", lUnit.Name, lUnit.Printer)
If Not IsNull(V) Then GoTo Error_Handler

Exit Function
'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    'MsgBox V, vbCritical, frmElp_Caption & "Table_Printer :"
    Table_Unit = V:
End Function
Public Function Table_Printer(lPrinter As String)
On Error GoTo Error_Handler

Dim V, X As String, wPrinter_Path As String

Table_Printer = ""
V = rsElpTable_Read("Server", "Printer", lPrinter, X, wPrinter_Path)
If Not IsNull(V) Then GoTo Error_Handler

Table_Printer = wPrinter_Path

Exit Function
'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
 '   MsgBox V, vbCritical, frmElp_Caption & "Table_Printer :" & lPrinter

End Function

Public Function Table_User(lUser As typeUser)
On Error GoTo Error_Handler

Dim V
Dim lenX As Integer, K As Integer, wMemo As String
Dim K2 As Integer, X As String

Table_User = Null
V = rsElpTable_Read("User", lUser.Id, "", lUser.Name, wMemo)
If IsNull(V) Then
    lUser.Unit = Mid$(wMemo, 1, 4)
    lUser.ProdTest = Mid$(wMemo, 6, 1)
    lUser.Edition_Hold = Mid$(wMemo, 7, 1)
    lUser.Edition_Aut = Mid$(wMemo, 8, 1)        ' 9 tous les spoules Utilisateur
    lUser.QSYSOPR = Mid$(wMemo, 9, 1)
    lUser.XXXXXX = Mid$(wMemo, 10, 1)
    If Not IsNumeric(lUser.XXXXXX) Then lUser.XXXXXX = "0"
    lUser.Printer = ""
    lUser.AliasWin = ""
    K = 11
    Do
        X = Space_Scan(wMemo, K)
        If X <> "" Then
            If Mid$(X, 1, 8) = "PRINTER:" Then
                lUser.Printer = Replace(X, "PRINTER:", "")
            Else
                If Mid$(X, 1, 6) = "CACLS:" Then
                    lUser.AliasWin = Replace(X, "CACLS:", "")
                Else
                    lUser.Printer = X
                End If
            End If
        End If
    Loop Until X = ""
Else
    lUser.Unit = "?"
    lUser.ProdTest = "0"
    lUser.Edition_Hold = "1"
    lUser.Edition_Aut = "0"
    lUser.QSYSOPR = "0"
    lUser.XXXXXX = ""
    lUser.Printer = ""
    lUser.AliasWin = ""
    Table_User = V
End If

Exit Function
'------------------------------------------k
Error_Handler:
    V = Error
Error_MsgBox:
   ' MsgBox V, vbCritical, frmElp_Caption & "Table_User :"
    Table_User = V

End Function

Public Function Table_User_CACLS(lUser As typeUser)
On Error GoTo Error_Handler

Dim V
Dim wUser_Id As String, lenX As Integer

Table_User_CACLS = Null

wUser_Id = Trim(lUser.Id)
lenX = Len(wUser_Id)
If Mid$(wUser_Id, 1, 1) = "_" Or Mid$(wUser_Id, 1, 9) = "SPLFFTPW0" Then
    lUser.Id = "BIA_INFO"
    lUser.Unit = Mid$(wUser_Id, 2, lenX - 1)
Else
    If Mid$(wUser_Id, 1, 2) = "T_" Then
        lUser.Id = Mid$(wUser_Id, 3, lenX - 2)
        V = Table_User(lUser)
        If Not IsNull(V) Then
            lUser.Id = wUser_Id
            Call Table_User(lUser)
        End If
    Else
       Call Table_User(lUser)
    End If
    If Trim(lUser.AliasWin) <> "" Then lUser.Id = Trim(lUser.AliasWin)
    lUser.Unit = Trim(lUser.Unit)
End If


Exit Function
'------------------------------------------k
Error_Handler:
    V = Error
Error_MsgBox:
    Table_User_CACLS = V

End Function

Public Sub Table_User_Load(lUnit As String, cbo As ComboBox, blnUser_Test As Boolean)

Dim V, wMemo As String, X As String
Dim xUser_Test As String

X = "select * from ElpTable where ID = 'USER' "
    
Set rsMDB = cnMDB.Execute(X)

Do Until rsMDB.EOF
    wMemo = rsMDB("Memo")
    If Mid$(wMemo, 1, 4) = lUnit Then
        X = rsMDB("K1")
        cbo.AddItem X
        If blnUser_Test Then
            If Mid$(X, 1, 2) = "P_" Then
                xUser_Test = X
                Mid$(X, 1, 2) = "T_"
            Else
                xUser_Test = "T_" & Mid$(X, 1, 8)
            End If
            cbo.AddItem xUser_Test
        End If

    End If
    
    rsMDB.MoveNext
Loop

Exit Sub
'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & "Table_Printer :"

End Sub


Public Sub mainSoc_AMJCPT_Load()
''' à faire recherche des 30 dernières dates comptables SAB073 / SAB073T ZSCHTAB0 & ZCOMHIS0
Dim blnRéplication As Boolean
Dim V, X As String, xSql As String
Dim xFileName As String, intFile As Integer
Dim xName As String, xMemo As String
Dim K As Integer
Dim blnOk As Boolean
Dim xElpTable As typeElpTable
Dim xAMJ As String
Dim kWait As Integer
On Error GoTo Error_Handler

'-------------------------------------------------------
App_Debug = "mainSoc_AMJCPT_Load : Taux TVA"
'-------------------------------------------------------
tauxTVA = 0
If blnOff_Line Then
    tauxTVA = 0.2
Else
    xSql = "select * from " & paramIBM_Library_SAB & ".ZBASTAB0" _
        & " where BASTABETA = 1 " _
        & " and   BASTABNUM = 25" _
        & " and BASTABARG like 'EURTVA000%'" _
        & " order by BASTABARG"
    
    Set rsSab = cnsab.Execute(xSql)
    
    Do Until rsSab.EOF
            
        tauxTVA = CDbl(convX2P(Mid$(rsSab("BASTABDON"), 1, 8))) / 100000000000#
        
        rsSab.MoveNext
    Loop
End If

If tauxTVA = 0 Then
    tauxTVA = 0.2   '0.196
    Call MsgBox("Erreur lecture du taux de TVA dans SAB, valeur par défaut : " & tauxTVA, vbCritical, App_Debug)
End If
tauxTTC = 1 + tauxTVA
tauxTVA_Lib = Format(tauxTVA * 100, "00.00") & "%"

'-------------------------------------------------------
App_Debug = "mainSoc_AMJCPT_Load : réplication YBIATAB0"
'-------------------------------------------------------
blnRéplication = True
blnOk = False
xSql = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 where" _
     & " BIATABID = 'DATE' and BIATABK1 = 'CPT' and BIATABK2 = 'J'"
Set rsSab = cnsab.Execute(xSql)
If rsSab.EOF Then
        MsgBox "Manque SAB073*SPE/YBIATAB0", vbCritical, frmElp_Caption & App_Debug
        End
End If
xAMJ = Trim(rsSab("BIATABTXT"))
'? déjà répliquée
'-----------------
rsElpTable_Init xElpTable
xElpTable.Id = "IBM"
xElpTable.K1 = "YBIATAB0"
xElpTable.K2 = "IMPORT"
For kWait = 1 To 5
    V = rsElpTable_Read(xElpTable.Id, xElpTable.K1, xElpTable.K2, xElpTable.Name, xMemo)
    If Mid$(xMemo, 1, 11) = "OK " & xAMJ Then
        Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, App_Debug & " OK")
        blnRéplication = False
        Exit For
    End If
    If Mid$(xMemo, 4, 8) <> xAMJ Then Exit For
    If Mid$(xMemo, 1, 2) <> "OK" Then
        Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, App_Debug & " attente" & kWait & " /5" & xMemo)
        DoEvents
        Wait_SS 5 '* 10
        DoEvents
    End If
Next kWait
If blnRéplication Then
   Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, App_Debug & " Début")
   xElpTable.Memo = "XX " & xAMJ & " " & DSys & " " & Time
    V = adoElpTable_Update(rsMDB, xElpTable)
    rsYBIATAB0_Réplication
    xElpTable.Memo = "OK " & xAMJ & " " & DSys & " " & Time
    V = adoElpTable_Update(rsMDB, xElpTable)
   Call lstErr_AddItem(frmElp.lstErr, frmElp.cmdContext, App_Debug & " Terminée")
End If
 


'-------------------------------------------------------
App_Debug = "mainSoc_AMJCPT_Load : initialisation dates"
'-------------------------------------------------------

arrAMJCPT(0) = dateElp("Ouvré", -1, DSys)
For K = 1 To 29
    arrAMJCPT(K) = dateElp("Ouvré", -1, arrAMJCPT(K - 1))
Next K

V = rsYBIATAB0_Read("DATE", "CPT", "J", YBIATAB0_DATE_CPT_J)
If Not IsNull(V) Then GoTo Error_MsgBox

V = rsYBIATAB0_Read("DATE", "CPT", "JP0", YBIATAB0_DATE_CPT_JP0)
If Not IsNull(V) Then MsgBox V, vbCritical, frmElp_Caption & App_Debug & "YBIATAB0_DATE_CPT_JP0"

V = rsYBIATAB0_Read("DATE", "CPT", "JP1", YBIATAB0_DATE_CPT_JP1)
If Not IsNull(V) Then MsgBox V, vbCritical, frmElp_Caption & App_Debug & "YBIATAB0_DATE_CPT_JP1"

V = rsYBIATAB0_Read("DATE", "CPT", "JS1", YBIATAB0_DATE_CPT_JS1)
If Not IsNull(V) Then MsgBox V, vbCritical, frmElp_Caption & App_Debug & "YBIATAB0_DATE_CPT_JS1J"

V = rsYBIATAB0_Read("DATE", "CPT", "M", YBIATAB0_DATE_CPT_M)
If Not IsNull(V) Then MsgBox V, vbCritical, frmElp_Caption & App_Debug & "YBIATAB0_DATE_CPT_M"

V = rsYBIATAB0_Read("DATE", "CPT", "MP1", YBIATAB0_DATE_CPT_MP1)
If Not IsNull(V) Then MsgBox V, vbCritical, frmElp_Caption & App_Debug & "YBIATAB0_DATE_CPT_MP1"

V = rsYBIATAB0_Read("DATE", "CPT", "MS1", YBIATAB0_DATE_CPT_MS1)
If Not IsNull(V) Then MsgBox V, vbCritical, frmElp_Caption & App_Debug & "YBIATAB0_DATE_CPT_MS1"

V = rsYBIATAB0_Read("DATE", "CPT", "A", YBIATAB0_DATE_CPT_A)
If Not IsNull(V) Then MsgBox V, vbCritical, frmElp_Caption & App_Debug & "YBIATAB0_DATE_CPT_A"

V = rsYBIATAB0_Read("DATE", "CPT", "AP1", YBIATAB0_DATE_CPT_AP1)
If Not IsNull(V) Then MsgBox V, vbCritical, frmElp_Caption & App_Debug & "YBIATAB0_DATE_CPT_AP1"

V = rsYBIATAB0_Read("DATE", "CAL", "AP1", YBIATAB0_DATE_CAL_AP1)
If Not IsNull(V) Then MsgBox V, vbCritical, frmElp_Caption & App_Debug & "YBIATAB0_DATE_CAL_AP1"

V = rsYBIATAB0_Read("DATE", "CAL", "MP1", YBIATAB0_DATE_CAL_MP1)
If Not IsNull(V) Then MsgBox V, vbCritical, frmElp_Caption & App_Debug & "YBIATAB0_DATE_CAL_M"

V = rsYBIATAB0_Read("DATE", "CPT", "AS1", YBIATAB0_DATE_CPT_AS1)
If Not IsNull(V) Then MsgBox V, vbCritical, frmElp_Caption & App_Debug & "YBIATAB0_DATE_CPT_AS1"

YBIATAB0_DATE_CPT_MP2 = dateElp("FinDeMoisP", 0, YBIATAB0_DATE_CPT_MP1)
YBIATAB0_DIBM_CPT_J = dateIBM(YBIATAB0_DATE_CPT_J)
YBIATAB0_DIBM_CPT_JS1 = dateIBM(YBIATAB0_DATE_CPT_JS1)
YBIATAB0_DIBM_CPT_JP1 = dateIBM(YBIATAB0_DATE_CPT_JP1)

Exit Sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug


End Sub

Public Function fileName_AMJCPT(lName As String, lIndex As Integer) As String
Dim K As Integer, L As Integer

fileName_AMJCPT = arrAMJCPT(lIndex) & "_" & Trim(lName) & ".txt"

End Function




Public Function Printer_Set(lPrinter_Name As String) As String
Dim X As String
Dim wPrinter_Name As String, wPrinter_Name_x As String
Dim K As Integer, lenPrinter_Name_X As Integer

On Error GoTo Error_Handler
Printer_Set = ""
wPrinter_Name = UCase$(Table_Printer(lPrinter_Name))
'!!! voir BIA Table_Unit_Printer_Set

''Si l'imprimante est gérée sur ce serveur => enlever "\\....\"
K = InStr(3, wPrinter_Name, "\")
If K > 0 Then
    wPrinter_Name_x = Mid$(wPrinter_Name, K + 1, Len(wPrinter_Name) - K)
Else
    wPrinter_Name_x = wPrinter_Name
End If
lenPrinter_Name_X = Len(wPrinter_Name_x)
For Each XPrt In Printers
    X = UCase$(Trim(XPrt.Devicename))
    'Debug.Print X
    K = InStr(1, X, wPrinter_Name_x)    ' 2004.08.25 JPL supprimer l'ID (\\...\ du poste de tarvail
    If K > 0 Then
        If lenPrinter_Name_X <> Len(X) - K + 1 Then K = 0      ' attention noms semblables : IMP_INFO et IMP_INFOCOLOR
    End If

    '''Debug.Print UCase$(Trim(XPrt.Devicename))
    If wPrinter_Name = UCase$(Trim(XPrt.Devicename)) _
    Or K > 0 Then
       Set Printer = XPrt
       frmElpPrt.prtColor_Check
        Printer_Set = wPrinter_Name
        Exit Function
    End If

Next

'Imprimante non trouvée => IMP_INFO

Error_Handler:

wPrinter_Name = UCase$(Table_Printer("INFO"))

For Each XPrt In Printers
    X = UCase$(Trim(XPrt.Devicename))
    K = InStr(1, X, "INFO")
    If K > 0 Then
        If K + 4 > Len(X) Then Set Printer = XPrt: Printer_Set = "INFO": frmElpPrt.prtColor_Check: Exit Function
    End If
Next

End Function
Public Function Printer_Set_SAV(lPrinter_Name As String) As String
'Dim X As String
'Dim wPrinter_Name As String, wPrinter_Name_x As String
'Dim K As Integer, lenPrinter_Name_X As Integer
'Dim ii As Long
'
'On Error GoTo Error_Handler
'Printer_Set = ""
'wPrinter_Name = UCase$(Table_Printer(lPrinter_Name))
'''Si l'imprimante est gérée sur ce serveur => enlever "\\....\"
'K = InStr(3, wPrinter_Name, "\")
'If K > 0 Then
'    wPrinter_Name_x = Mid$(wPrinter_Name, K + 1, Len(wPrinter_Name) - K)
'Else
'    wPrinter_Name_x = wPrinter_Name
'End If
'lenPrinter_Name_X = Len(wPrinter_Name_x)
'For ii = 2 To Val(collection_IMP(1))
'    X = UCase$(collection_IMP(ii))
'    'Debug.Print X
'    K = InStr(1, X, wPrinter_Name_x)    ' 2004.08.25 JPL supprimer l'ID (\\...\ du poste de tarvail
'    If K > 0 Then
'        If lenPrinter_Name_X <> Len(X) - K + 1 Then K = 0      ' attention noms semblables : IMP_INFO et IMP_INFOCOLOR
'    End If
'    If wPrinter_Name = UCase$(collection_IMP(ii)) _
'    Or K > 0 Then
'       Set Printer = collection_IMP(ii)
'       frmElpPrt.prtColor_Check
'        Printer_Set = wPrinter_Name
'        Exit Function
'    End If
'Next ii
''Imprimante non trouvée => IMP_INFO
'Error_Handler:
'
'wPrinter_Name = UCase$(Table_Printer("INFO"))
'For ii = 2 To Val(collection_IMP(1))
'    X = UCase$(collection_IMP(ii))
'    K = InStr(1, X, "INFO")
'    If K > 0 Then
'        If K + 4 > Len(X) Then Set Printer = collection_IMP(ii)
'        Printer_Set = "INFO"
'        frmElpPrt.prtColor_Check
'        Exit Function
'    End If
'Next
End Function

Public Function Printer_Reset() As String

On Error GoTo Error_Handler

Printer_Reset = ""
If Printer_Previous_DeviceName = "" Then GoTo Error_Handler
For Each XPrt In Printers
    If XPrt.Devicename = Printer_Previous_DeviceName Then
       Set Printer = XPrt
       frmElpPrt.prtColor_Check
        Exit Function
    End If
Next

Error_Handler:
Dim X As String, K As Integer
For Each XPrt In Printers
    X = UCase$(Trim(XPrt.Devicename))
    K = InStr(1, X, "INFO")
    If K > 0 Then
        If K + 4 > Len(X) Then Set Printer = XPrt: Printer_Reset = "INFO": frmElpPrt.prtColor_Check: Exit Function
    End If
Next

End Function

Public Function Printer_PDF() As String
Dim X As String
Dim wPrinter_Name As String, wPrinter_Name_x As String
Dim K As Integer, lenPrinter_Name_X As Integer
Dim blnOk As Boolean

On Error GoTo Error_Handler
blnOk = False


If nomDuServeur <> paramServerSplf Then
    For Each XPrt In Printers
        X = UCase$(Trim(XPrt.Devicename))
        K = InStr(1, X, paramIMP_PDFCreator_Name) '"PDF_BIA_SAB") '
        If K > 0 Then
           Set Printer = XPrt
            Printer_PDF = XPrt.Devicename
            Printer.ColorMode = 2
            prtColorMode = True
            blnIMP_PDF = True
            paramIMP_PDF_Path = paramIMP_PDF_Path_VBP
            Exit Function
        End If
    Next
End If
paramIMP_PDF_Path = paramIMP_PDF_Path_Temp

If nomDuServeur <> paramServerSplf Then
    For Each XPrt In Printers
        X = UCase$(Trim(XPrt.Devicename))
        K = InStr(1, X, "PDFCREATOR")
        If K > 0 Then
           Set Printer = XPrt
            Printer_PDF = XPrt.Devicename
            Printer.ColorMode = 2
            prtColorMode = True
            blnIMP_PDF = True
            frmElpPrt.prtIMP_PDF_Monitor "Clear"
            Exit Function
        End If
    Next
    For Each XPrt In Printers
        X = UCase$(Trim(XPrt.Devicename))
        K = InStr(1, X, "ADOBE PDF")
        If K > 0 Then
           Set Printer = XPrt
            Printer_PDF = XPrt.Devicename
            Printer.ColorMode = 2
            prtColorMode = True
            blnIMP_PDF = True
            frmElpPrt.prtIMP_PDF_Monitor "Clear"
            Exit Function
        End If
    Next
End If
'Imprimante non trouvée => IMP_INFO

Error_Handler:

wPrinter_Name = UCase$(Table_Printer("INFO"))

For Each XPrt In Printers
    X = UCase$(Trim(XPrt.Devicename))
    K = InStr(1, X, "INFO")
    If K > 0 Then
        If K + 4 > Len(X) Then Set Printer = XPrt: Printer_PDF = "INFO": Exit Function
    End If
Next

End Function

Public Function Printer_Set_Unit(lUnit As String) As String
Dim xUnit As typeUnit
If Trim(lUnit) = "" Then
    xUnit.Id = "INFO"
Else
    xUnit.Id = lUnit
End If
Call Table_Unit(xUnit)
Printer_Set_Unit = Printer_Set(Trim(xUnit.Printer))
End Function



Public Function blnAuto_Exploitation_Ok(lFct As String, lApplication As String)
Dim xFileName As String, X8 As String, xIn As String
Dim fsoFile As File
Dim intFile As Integer

On Error Resume Next
blnAuto_Exploitation_Ok = True

xFileName = paramYBase_DataF & lApplication & "_Exploitation_Ok.txt"
Select Case lFct
    Case "DateLastModified"
            Set fsoFile = msFileSystem.GetFile(xFileName)
            If Err = 0 Then
                Call dateJMA6_AMJ(fsoFile.DateLastModified, X8)
                If X8 >= DSys Then blnAuto_Exploitation_Ok = False
               
            End If
    Case "Update"
        Call FEU_ROUGE
        intFile = FreeFile(0)
        Open xFileName For Output As #intFile
        Print #intFile, YBIATAB0_DATE_CPT_J, DSys, Time
        Close #intFile
        Call FEU_VERT
    Case "DATE_CPT_J"
        If Dir(xFileName) <> "" Then
            intFile = FreeFile(0)
            Open xFileName For Input As intFile
            Line Input #intFile, xIn
            If Mid$(xIn, 1, 8) >= YBIATAB0_DATE_CPT_J Then blnAuto_Exploitation_Ok = False
            Close #intFile
        End If
        
End Select
End Function

Public Function SAA_Text_Control(lX As String, lenMax As Integer) As String
Dim X As String, lenX As Integer
Dim I As Integer
X = UCase$(lX)

lenX = Len(X)
If lenX > lenMax Then X = Mid$(X, 1, lenMax): lenX = lenMax
For I = 1 To lenX
    Select Case Mid$(X, I, 1)
        Case "A" To "Z":
        Case "0" To "9"
        Case ".", "-":
        Case Chr$(200), Chr$(201), Chr$(202): Mid$(X, I, 1) = "E"
       Case Else: Mid$(X, I, 1) = " "
    End Select
Next I
SAA_Text_Control = X
End Function

Public Sub mainSoc_Display()
frmElp.lstElp_Environnement.Clear
frmElp.lstElp_Environnement.AddItem "DATE_CPT_Veille" & vbTab & " : " & dateImp(YBIATAB0_DATE_CPT_J)
frmElp.lstElp_Environnement.AddItem "Environnement" & vbTab & " : " & paramEnvironnement
frmElp.lstElp_Environnement.AddItem "BIC émission SWIFT" & vbTab & " : " & paramBic8
frmElp.lstElp_Environnement.AddItem ""
frmElp.lstElp_Environnement.AddItem "AS400" & vbTab & vbTab & " : " & paramIBM_AS400_ID
'frmElp.lstElp_Environnement.AddItem "utilisateur" & vbTab & vbTab & " : " & Trim(currentUser.Id) & " / " & currentZMNURUT0.MNURUTNOM
frmElp.lstElp_Environnement.AddItem ""
frmElp.lstElp_Environnement.AddItem "GetUserNameEx  " & vbTab & " : " & usrIdNT
frmElp.lstElp_Environnement.AddItem "utilisateur SSI WIN" & vbTab & " : " & currentSSIWINUIDX & " / " & currentSSIWINUNOM
frmElp.lstElp_Environnement.AddItem "utilisateur SAB" & vbTab & " : " & usrIdSAB
frmElp.lstElp_Environnement.AddItem "SAB groupe" & vbTab & " : " & currentZMNUUTI0.MNUUTIGR2
frmElp.lstElp_Environnement.AddItem "Service (.mdb)" & vbTab & " : " & currentUnit.Id
frmElp.lstElp_Environnement.AddItem "Imprimante" & vbTab & " : " & currentUnit.Printer
'frmElp.lstElp_Environnement.AddItem "Sigle 'Mes Comptes'" & vbTab & " : " & currentCLIENASIG
frmElp.lstElp_Environnement.AddItem "Service " & vbTab & vbTab & " : " & currentSSIWINUNIT & " " & currentSSIWINUNIT_Lib

frmElp.lstElp_Environnement.AddItem ""
frmElp.lstElp_Environnement.AddItem "ODBC_SAB" & vbTab & " : " & paramIBM_ODBC_SAB
frmElp.lstElp_Environnement.AddItem "Library_SAB" & vbTab & " : " & paramIBM_Library_SAB
frmElp.lstElp_Environnement.AddItem "Library_SAB" & vbTab & " : " & paramIBM_Library_SAB
frmElp.lstElp_Environnement.AddItem "Répertoire_Local" & vbTab & " : " & paramFolder_Local

frmElp.lstElp_Environnement.AddItem "Mail_SMTPHost" & vbTab & " : " & paramSendMail_SMTPHost
frmElp.lstElp_Environnement.AddItem "Mail_From" & vbTab & " : " & currentSSIWINMAIL ' paramSendMail_From
frmElp.lstElp_Environnement.AddItem "Mail_BIA_URL" & vbTab & " : " & paramSendMail_BIA_URL
frmElp.lstElp_Environnement.AddItem ""
frmElp.lstElp_Environnement.AddItem "socName     " & vbTab & " : " & socName

frmElp.lstElp_Environnement.AddItem "SOC_RS" & vbTab & vbTab & " : " & paramSOC_RS
frmElp.lstElp_Environnement.AddItem "SOC_Adresse" & vbTab & " : " & paramSOC_Adresse
frmElp.lstElp_Environnement.AddItem "SOC_Ville" & vbTab & " : " & paramSOC_Ville
frmElp.lstElp_Environnement.AddItem "soc_capital" & vbTab & " : " & paramsoc_capital
frmElp.lstElp_Environnement.AddItem "SOC_Télécom" & vbTab & " : " & paramSOC_Télécom
frmElp.lstElp_Environnement.AddItem ""
frmElp.lstElp_Environnement.AddItem "BIC courrier" & vbTab & " : " & SocBicId
frmElp.lstElp_Environnement.AddItem "BDF" & vbTab & vbTab & " : " & strSocBdfE & vbTab & strSocBdfG
frmElp.lstElp_Environnement.AddItem "RIB DOM" & vbTab & vbTab & " : " & SocRibDom
frmElp.lstElp_Environnement.AddItem "RIB Tél" & vbTab & vbTab & " : " & socTéléphone
frmElp.lstElp_Environnement.AddItem "width : " & Screen.Width & " height : " & Screen.Height & " pixelX : " & Screen.TwipsPerPixelX & " pixelY : " & Screen.TwipsPerPixelY
If paramEnvironnement = constProduction Then
    frmElp.SSTab1.Tab = 0
Else
    frmElp.SSTab1.Tab = 0  '1
End If
End Sub

Public Function Shell_VB6(lApplication As String) As String
Dim X As String, xName As String, xMemo As String
Dim r As Currency, X16 As String

Shell_VB6 = Space$(66)
X = DSys & time_Hms
r = Mid$(X, 1, 9) Mod 97
r = 97 - (Format$(r, "00") & Mid$(X, 10, 5) & "00") Mod 97
Mid$(Shell_VB6, 1, 16) = X & Format$(r, "00")

Call rsElpTable_Read(paramBiaPgmAut, Elp.usrId, lApplication, xName, xMemo)
Mid$(Shell_VB6, 17, 10) = Elp.usrId
Mid$(Shell_VB6, 27, 10) = xMemo
Mid$(Shell_VB6, 37, 30) = Printer.Devicename

End Function

Public Function mailAdresse_Production(lUser) As String
Dim xSql As String
Dim rsSab_Local As New ADODB.Recordset

xSql = "select SSIWINMAIL from " & paramIBM_Library_SABSPE & ".YSSIWIN0" _
     & " where SSIWINNAt = ' ' and SSIWINUIDX = '" & Trim(lUser) & "'"
     
Set rsSab_Local = cnsab.Execute(xSql)
If Not rsSab_Local.EOF Then
    mailAdresse_Production = Trim(rsSab_Local("SSIWINMAIL"))
Else
    mailAdresse_Production = ""
End If
End Function

Public Function mailAdresse_Production_Load() As String
Dim xSql As String, K As Integer, X As String
Dim rsSab As New ADODB.Recordset

xSql = "select count(*) as Tally    from " & paramIBM_Library_SABSPE & ".YSSIWIN0 " _
     & " where SSIWINNAT = ' ' and SSIWINMAIL <> ''"
Set rsSab = cnsab.Execute(xSql)
K = rsSab("Tally") + 1
ReDim arrUSR_UTI(K), arrUSR_Mail(K), arrUSR_Mail_UCase(K)

arrUSR_Mail_Nb = 0
xSql = "select SSIWINUIDX , SSIWINMAIL  from " & paramIBM_Library_SABSPE & ".YSSIWIN0 " _
     & " where SSIWINNAT = ' ' and SSIWINMAIL <> '' and SSIWINPRFX <> 'X' order by SSIWINUIDX"
     
Set rsSab = cnsab.Execute(xSql)
Do While Not rsSab.EOF
    arrUSR_Mail_Nb = arrUSR_Mail_Nb + 1
         
    arrUSR_UTI(arrUSR_Mail_Nb) = UCase$(Trim(rsSab("SSIWINUIDX")))
    arrUSR_Mail(arrUSR_Mail_Nb) = StrConv(Trim(rsSab("SSIWINMAIL")), vbProperCase)
    rsSab.MoveNext
Loop


ReDim Preserve arrUSR_UTI(arrUSR_Mail_Nb), arrUSR_Mail(arrUSR_Mail_Nb)

End Function

Public Function mailAdresse_Production_Control(lIn As String, lOUT As String)
Dim K As Integer, kLen As Integer, K1 As Integer
Dim X As String, xMail As String
Dim V
If arrUSR_Mail_Nb = 0 Then mailAdresse_Production_Load

lOUT = ""
mailAdresse_Production_Control = Null

kLen = Len(lIn)
K1 = 1
Do
    K = InStr(K1, lIn, ";")
    If K > 0 Then
        X = Trim(Mid$(lIn, K1, K - K1))
        If X <> "" Then
            V = mailAdresse_Production_Control_UTI(X, xMail)
            If IsNull(V) Then
                lOUT = lOUT & xMail & ";"
            Else
                mailAdresse_Production_Control = mailAdresse_Production_Control & vbCrLf & "? " & V
            End If
        End If
        K1 = K + 1
    Else
        If K1 < kLen Then
            X = Trim(Mid$(lIn, K1, kLen - K1 + 1))
            V = mailAdresse_Production_Control_UTI(X, xMail)
            If IsNull(V) Then
                lOUT = lOUT & xMail & ";"
            Else
                mailAdresse_Production_Control = mailAdresse_Production_Control & vbCrLf & "? " & V
            End If
        End If
        Exit Do
    End If
Loop
kLen = Len(lOUT)
If kLen > 0 Then
    If Mid$(lOUT, kLen, 1) = ";" Then Mid$(lOUT, kLen, 1) = " "
Else
    mailAdresse_Production_Control = "? vide"
End If


End Function
Public Function mailAdresse_Production_Control_UTI(lUTI As String, lUTI_Mail As String)
Dim K As Integer, X As String

mailAdresse_Production_Control_UTI = Null

If InStr(lUTI, "@") > 0 Then
    X = StrConv(Trim(lUTI), vbProperCase)
    For K = 1 To arrUSR_Mail_Nb
        If X = arrUSR_Mail(K) Then
            lUTI_Mail = arrUSR_Mail(K)
            Exit Function
        End If
    Next K
    mailAdresse_Production_Control_UTI = "- adresse mail inconnue : " & lUTI

Else
    X = UCase(lUTI)
    For K = 1 To arrUSR_Mail_Nb
        If X = arrUSR_UTI(K) Then
            lUTI_Mail = arrUSR_Mail(K)
            Exit Function
        End If
    Next K
    mailAdresse_Production_Control_UTI = "- utilisateur inconnu : " & lUTI
End If

End Function
Public Sub mainSoc_BanqueIslamique()
'Chargement des codes 'Banque Islamique'
'en dur dans frmEdition 50451
Dim X As String, Nb As Long

blnBanqueIslamique_Loop = False
arrBanqueIslamique_Nb = 0
X = "select count(*) as Tally  from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
    & " where BIATABID = 'ZLIBEL0'" _
    & " and BIATABK1 = 'INTERET'"
Set rsSab = cnsab.Execute(X)
Nb = rsSab("Tally") + 1
ReDim arrBanqueIslamique(Nb)

X = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
    & " where BIATABID = 'ZLIBEL0'" _
    & " and BIATABK1 = 'INTERET'"
    
Set rsSab = cnsab.Execute(X)
Do While Not rsSab.EOF
    X = Mid$(rsSab("BIATABK2"), 1, 5)
    If X <> "50451" Then
        blnBanqueIslamique_Loop = True
        arrBanqueIslamique_Nb = arrBanqueIslamique_Nb + 1
        arrBanqueIslamique(arrBanqueIslamique_Nb) = X
        'Call MsgBox("Nouvelle banque islamique : " & X & " revoir edition des avis", vbExclamation, "Banques Islamiques")
    End If
    rsSab.MoveNext
Loop

End Sub

Public Sub BIA_VB_HAB(lAPP As String, larrHab() As Boolean, cboSelect_SQL As ComboBox)
On Error GoTo Error_Handler
Dim V, K As Integer, X As String, Xc As String, blnOk As Boolean

App_Debug = "> BIA_VB_HAB " & lAPP
'==========================================================================
For K = 1 To 19: larrHab(K) = False: Next K

X = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
     & " where BIATABID = 'BIA_VB_HAB' and BIATABK1 = '" & lAPP & "'and BIATABK2 = '" & currentSSIWINUIDX & "'"
Set rsSab = cnsab.Execute(X)

If Not rsSab.EOF Then
    X = rsSab("BIATABTXT")
    For K = 1 To 19
        If Mid$(X, K, 1) <> " " Then larrHab(K) = True
    Next K
End If
'_____________________________________________________________________________

cboSelect_SQL.Clear

X = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
     & " where BIATABID = 'BIA_VB_MNU' and BIATABK1 = '" & lAPP & "'"
Set rsSab = cnsab.Execute(X)

Do While Not rsSab.EOF
    blnOk = True
    X = rsSab("BIATABTXT")
    For K = 1 To 19
        If Mid$(X, K, 1) <> " " And Not larrHab(K) Then blnOk = False
            
    Next K
    If blnOk Then
        Xc = Trim(rsSab("BIATABK2"))
        If Len(Xc) < 7 Then
            cboSelect_SQL.AddItem Mid$(Xc & "      ", 1, 6) & " - " & Trim(Mid$(X, 20, 79))
        Else
            cboSelect_SQL.AddItem Xc & " - " & Trim(Mid$(X, 20, 79))
        End If
        
    End If
    rsSab.MoveNext

Loop

'If cboSelect_SQL.ListCount > 0 Then cboSelect_SQL.ListIndex = 0

'cboSelect_SQL.AddItem "1 - sélection droit => Utilisateurs"


Exit Sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug

End Sub
'---------------------------------------------------------
Public Function BIA_VB_APP()
'---------------------------------------------------------
Dim X As String, V, I As Integer, arrHab17_Nb As Integer
Dim wColor As Long
Dim xName As String, xDoc As String
Dim mTop As Long
Dim xWIN_Name As String
On Error GoTo Error_Handler

App_Debug = "> BIA_VB_APP : applications autorisées pour " & usrName_UCase
'--------------------------------------------------------------------------------------
BIA_VB_APP = Null
            
'XListBox.Clear
'XListBox.Visible = True
frmElp.fgMain_App.Rows = 1: frmElp.fgMain_App_X.Rows = 1
paramList_Height = 245
XLabel.Visible = True
XLabel.Caption = "Menu"
mTop = 0
'--------------------------------------------------------------------------------------
    xWIN_Name = currentSSIWINUIDX ' usrIdNT
X = "select count(*) from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
    & " where BIATABID = 'BIA_VB_HAB' and BIATABK2 = '" & xWIN_Name & "'"
Set rsSab = cnsab.Execute(X)

'''Call MsgBox(rsSab(0) & xWIN_Name, vbInformation, "BIA_VB_APP ")

ReDim arrHab17(rsSab(0) + 1)
    X = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
        & " where BIATABID = 'BIA_VB_HAB' and BIATABK2 = '" & xWIN_Name & "'and substring(BIATABTXT,17,1) <> ' '"

If blnOff_Line Then X = Replace(X, "substring", "mid$")

Set rsSab = cnsab.Execute(X)
arrHab17_Nb = 0

Do While Not rsSab.EOF
    arrHab17_Nb = arrHab17_Nb + 1
    arrHab17(arrHab17_Nb) = Trim(rsSab("BIATABK1"))
    rsSab.MoveNext
Loop

'--------------------------------------------------------------------------------------

X = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
     & " where BIATABID = 'BIA_VB_APP' and ( substring(BIATABTXT,80,12) = '" & frmElp_Caption & "' or substring(BIATABTXT,80,12) = '' ) and BIATABK1 in " _
     & " (select BIATABK1 from " & paramIBM_Library_SABSPE & ".YBIATAB0  where BIATABID = 'BIA_VB_HAB' and BIATABK2 = '" & xWIN_Name & "') order by BIATABK1"


If blnOff_Line Then X = Replace(X, "substring", "mid$")

Set rsSab = cnsab.Execute(X)

Do While Not rsSab.EOF
    xName = rsSab("BIATABTXT")
    wColor = vbBlue 'RGB(0, 0, 64)
    
    frmElp.fgMain_App.Rows = frmElp.fgMain_App.Rows + 1
    frmElp.fgMain_App.Row = frmElp.fgMain_App.Rows - 1
    X = Trim(rsSab("BIATABK1"))
    If Mid$(X, 1, 1) = "@" Then wColor = vbMagenta: mTop = frmElp.fgMain_App.Row
    frmElp.fgMain_App.Col = 0: frmElp.fgMain_App.Text = X
    frmElp.fgMain_App.CellForeColor = wColor
    frmElp.fgMain_App.CellFontSize = 11
    frmElp.fgMain_App.CellFontBold = True
    
    xDoc = Trim(Mid$(xName, 70, 10))
    If xDoc <> "" Then
        For I = 1 To arrHab17_Nb
            If X = arrHab17(I) Then
                frmElp.fgMain_App.Col = 1: frmElp.fgMain_App.Text = xDoc
                frmElp.fgMain_App.CellForeColor = &HD0FFD0   ' vbMagenta
                frmElp.fgMain_App.CellBackColor = &HD0FFD0   ' vbMagenta
                Exit For
            End If
        Next I
    End If
    
    frmElp.fgMain_App.Col = 2: frmElp.fgMain_App.Text = Trim(Mid$(xName, 1, 69))
    frmElp.fgMain_App.CellForeColor = vbBlack 'wColor
    rsSab.MoveNext

Loop
frmElp.fgMain_App.Rows = frmElp.fgMain_App.Rows + 1
frmElp.fgMain_App.Row = frmElp.fgMain_App.Rows - 1
frmElp.fgMain_App.Col = 0: frmElp.fgMain_App.Text = "X_Reset"
frmElp.fgMain_App.CellForeColor = vbRed
frmElp.fgMain_App.CellFontBold = True
frmElp.fgMain_App.Col = 2: frmElp.fgMain_App.Text = "réplication du répertoire c:\BiaSrv"
frmElp.fgMain_App.CellForeColor = vbRed
I = frmElp.fgMain_App.Height / frmElp.fgMain_App.RowHeightMin

If frmElp.fgMain_App.Rows > I Then frmElp.fgMain_App.TopRow = frmElp.fgMain_App.Rows - I

For I = 1 To frmElp.fgMain_App.Rows - 1
    frmElp.fgMain_App.Row = I
    frmElp.fgMain_App.Col = 0
    arrBiapgm(I) = Trim(frmElp.fgMain_App.Text)
Next I
'==========================================================================
mTop = 0
X = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
     & " where BIATABID = 'BIA_VB_APP' and ( substring(BIATABTXT,80,12) <> '" & frmElp_Caption & "' and substring(BIATABTXT,80,12) <> '' ) and BIATABK1 in " _
     & " (select BIATABK1 from " & paramIBM_Library_SABSPE & ".YBIATAB0  where BIATABID = 'BIA_VB_HAB' and BIATABK2 = '" & xWIN_Name & "') order by BIATABK1"

If blnOff_Line Then X = Replace(X, "substring", "mid$")

Set rsSab = cnsab.Execute(X)

Do While Not rsSab.EOF
    xName = rsSab("BIATABTXT")
    Select Case Trim(Mid$(xName, 80, 12))
        Case frmElp_Caption: wColor = RGB(0, 64, 0)
        Case "BIA_SAB", "BIA_SWIFT": wColor = vbBlue
        Case "BIA_AUDIT", "BIA_DWH": wColor = RGB(128, 0, 255)
        Case Else: wColor = RGB(96, 96, 96)
    
    End Select
    frmElp.fgMain_App_X.Rows = frmElp.fgMain_App_X.Rows + 1
    frmElp.fgMain_App_X.Row = frmElp.fgMain_App_X.Rows - 1
    X = Trim(rsSab("BIATABK1"))
    If Mid$(X, 1, 1) = "@" Then wColor = RGB(200, 100, 0): mTop = frmElp.fgMain_App_X.Row
    frmElp.fgMain_App_X.Col = 0: frmElp.fgMain_App_X.Text = X
    frmElp.fgMain_App_X.CellForeColor = wColor
    frmElp.fgMain_App_X.CellFontBold = True
    
    xDoc = Trim(Mid$(xName, 70, 10))
    If xDoc <> "" Then
        For I = 1 To arrHab17_Nb
            If X = arrHab17(I) Then
                frmElp.fgMain_App_X.Col = 1: frmElp.fgMain_App_X.Text = xDoc
                frmElp.fgMain_App_X.CellForeColor = &HD0FFD0 ' vbMagenta
                frmElp.fgMain_App_X.CellBackColor = &HD0FFD0 ' vbMagenta

                Exit For
            End If
        Next I
    End If

    
    frmElp.fgMain_App_X.Col = 2: frmElp.fgMain_App_X.Text = Trim(Mid$(xName, 1, 69))
    frmElp.fgMain_App_X.CellForeColor = wColor
    frmElp.fgMain_App_X.Col = 3: frmElp.fgMain_App_X.Text = Trim(Mid$(xName, 80, 12))
    frmElp.fgMain_App_X.CellForeColor = wColor
    rsSab.MoveNext

Loop
'frmElp.fgMain_App_X.TopRow = mTop + 1
I = frmElp.fgMain_App_X.Height / frmElp.fgMain_App_X.RowHeightMin

If frmElp.fgMain_App_X.Rows > I Then frmElp.fgMain_App_X.TopRow = frmElp.fgMain_App_X.Rows - I

'==========================================================================


Exit Function

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
    BIA_VB_APP = V
End Function

'---------------------------------------------------------
Public Function BIA_VB_APP_XXXXX()
'---------------------------------------------------------
Dim X As String, V, I As Integer
Dim xName As String, xK2 As String, xK2x As String
Dim H As Long

On Error GoTo Error_Handler

App_Debug = "> BIA_VB_APP : applications autorisées pour " & Elp.usrId
'--------------------------------------------------------------------------------------
BIA_VB_APP = Null

            
XListBox.Clear
XListBox.Visible = True
paramList_Height = 245
XLabel.Visible = True
XLabel.Caption = "Menu"

X = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
     & " where BIATABID = 'BIA_VB_APP' and ( substring(BIATABTXT,80,12) = '" & frmElp_Caption & "' or substring(BIATABTXT,80,12) = '' ) and BIATABK1 in " _
     & " (select BIATABK1 from " & paramIBM_Library_SABSPE & ".YBIATAB0  where BIATABID = 'BIA_VB_HAB' and BIATABK2 = '" & Elp.usrId & "')"

Set rsSab = cnsab.Execute(X)

Do While Not rsSab.EOF
    XListBox.AddItem rsSab("BIATABK1") & "    " & Trim(Mid$(rsSab("BIATABTXT"), 1, 79))
    rsSab.MoveNext

Loop
XListBox.AddItem "X_Reset     " & "    " & "réplication BiaSrv"

H = paramList_Height + paramList_Height * XListBox.ListCount
For I = 0 To XListBox.ListCount - 1
    XListBox.ListIndex = I
    arrBiapgm(I) = XListBox.Text
Next I
XListBox.ListIndex = -1
'==========================================================================
Exit Function

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
    BIA_VB_APP = V
End Function




Public Sub arrBIA_RCOM_Load()
Dim K As Integer, X As String
'Initialisation RCOM ______________________________________________________________________________


For K = 0 To 99
     arrBIA_RCOM_Code(K) = "R" & Format$(K, "00") ': arrBIA_RCOM_Lib(K) = arrBIA_RCOM_Code(K)
Next K

If blnOff_Line Then Exit Sub

X = "select * from  " & paramIBM_Library_SAB & ".ZBASTAB0 " _
    & " where  BASTABETA = 1 and BASTABNUM = 6 and BASTABARG like 'CLIR%' order by BASTABARG"
Set rsSab = cnsab.Execute(X)
Do While Not rsSab.EOF
    K = Val(Mid$(rsSab("BASTABARG"), 5, 2))
    arrBIA_RCOM_Code(K) = Mid$(rsSab("BASTABARG"), 4, 3)
    arrBIA_RCOM_Lib(K) = Trim(Mid$(rsSab("BASTABDON"), 24, 10))
    '$JPL 2013-10-03 If arrBIA_RCOM_Lib(K) = "" Then arrBIA_RCOM_Lib(K) = arrBIA_RCOM_Code(K)
    '$JPL 2013-10-03 If K = 32 Then arrBIA_RCOM_Lib(K) = arrBIA_RCOM_Lib(K) & ";PERRET"
    rsSab.MoveNext
Loop

'$JPL 2013-10-03  destinataires complémentaires
X = "select * from  " & paramIBM_Library_SABSPE & ".YSSIMEL0 " _
    & " where  SSIMELNAT = '@' and SSIMELUIDX like 'RCOM.%' and SSIMELPRFK <> 'X' order by SSIMELUIDX"
Set rsSab = cnsab.Execute(X)
Do While Not rsSab.EOF
    X = Replace(rsSab("SSIMELUIDX"), "RCOM.R", "")
    K = Val(X)
    If K < 99 Then
        If arrBIA_RCOM_Lib(K) = "" Then
            arrBIA_RCOM_Lib(K) = rsSab("SSIMELINFO")
        Else
            arrBIA_RCOM_Lib(K) = arrBIA_RCOM_Lib(K) & ";" & rsSab("SSIMELINFO")
        End If
    End If
    
    rsSab.MoveNext
Loop


End Sub

Public Sub arrMNURUTUTI_Load()
Dim K As Integer, X As String
'Initialisation Utilisateur ______________________________________________________________________________
X = "select * from  " & paramIBM_Library_SAB & ".ZMNURUT0 " _
    & " order by MNURUTCUT desc"
Set rsSab = cnsab.Execute(X)

Do While Not rsSab.EOF
    K = Val(rsSab("MNURUTCUT"))
    If arrMNURUTUTI_Nb = 0 Then ReDim arrMNURUTUTI(K + 1):  arrMNURUTUTI_Nb = K
    arrMNURUTUTI(K) = rsSab("MNURUTUTI")
    rsSab.MoveNext
Loop
End Sub




Public Sub BIA_VB_HAB_Idem(lFct As String, lUsrName As String)
Dim xSql As String
xSql = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
     & " where BIATABID = 'BIA_VB_HAB=X' and BIATABK1 = '" & UCase(Trim(lFct)) & "'and BIATABK2 = '" & currentSSIWINUIDX & "'"
Set rsSab = cnsab.Execute(xSql)

If Not rsSab.EOF Then
    idemUser.Id = Trim(rsSab("BIATABTXT"))
    Call Table_User(idemUser)
    Call MsgBox("Vous avez les mêmes habilitations que l'utilisateur : " & idemUser.Id & vbCrLf & " service : " & idemUser.Unit, vbExclamation, "Délégation : " & UCase(Trim(lFct)))
Else
    idemUser.Id = currentSSIWINUIDX 'usrName_UCase
    idemUser.Unit = currentUser.Unit
End If
End Sub

Public Sub BIA_VB_Hab_Idem_Mail()
Dim xSql As String, Nb As Integer
Dim rsSab As New ADODB.Recordset

xSql = "select count(*) from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
     & " where BIATABID = 'BIA_VB_MAIL' "
Set rsSab = cnsab.Execute(xSql)

arrMail_Nb = rsSab(0)

ReDim arrMail_K1(arrMail_Nb), arrMail_K2(arrMail_Nb), arrMail_Memo(arrMail_Nb)

xSql = "select * from " & paramIBM_Library_SABSPE & ".YBIATAB0 " _
     & " where BIATABID = 'BIA_VB_MAIL' order by  BIATABK1 , BIATABK2 "
Set rsSab = cnsab.Execute(xSql)

Do While Not rsSab.EOF
    arrMail_K1(Nb) = Trim(rsSab("BIATABK1"))
    arrMail_K2(Nb) = Trim(rsSab("BIATABK2"))
    arrMail_Memo(Nb) = Trim(rsSab("BIATABTXT"))
    
    Nb = Nb + 1
    rsSab.MoveNext
Loop

End Sub

Public Function Table_Unit_SSI(lFct As String, lUnit As String) As String


If lFct = "" Then
    Select Case lUnit
        Case "ORPA", "SOBF", "GDMP", "0000", "00TR", "00MP", "00GU": Table_Unit_SSI = "S01"
        Case "00CD", "SOBI": Table_Unit_SSI = "S10"
        Case "BOTC", "DAFI", "00CR", "GDC": Table_Unit_SSI = "S32"
        Case "CSOP", "GSOP", "CCGA": Table_Unit_SSI = "S11"
        Case "INSP": Table_Unit_SSI = "S21"
        Case "CDG": Table_Unit_SSI = "S22"
        Case "DOP": Table_Unit_SSI = "S30"
        Case "ORGA": Table_Unit_SSI = "S31"
        Case "INFO": Table_Unit_SSI = "S40"
        Case "DCOM": Table_Unit_SSI = "S41"
        Case "DEON": Table_Unit_SSI = "S42"
        Case "CONF": Table_Unit_SSI = "S43"
        Case "DER": Table_Unit_SSI = "S51"
        Case "JURI": Table_Unit_SSI = "S52"
        Case "DTX": Table_Unit_SSI = "S53"
        Case "FOTC": Table_Unit_SSI = "S54"
        Case "CPT", "CPCP", "CPXX", "CPTP": Table_Unit_SSI = "S60"
        Case "DRH", "RHRH": Table_Unit_SSI = "S61"
        Case "S99": Table_Unit_SSI = "S99"
        Case Else: Table_Unit_SSI = "S00"
    End Select
Else
    Select Case lUnit
        Case "S01": Table_Unit_SSI = "GDMP"
        Case "S10": Table_Unit_SSI = "SOBI"
        Case "S11": Table_Unit_SSI = "CCGA"
        Case "S21": Table_Unit_SSI = "INSP"
        Case "S22": Table_Unit_SSI = "CDG"
        Case "S30": Table_Unit_SSI = "DOP"
        Case "S31": Table_Unit_SSI = "ORGA"
        Case "S32": Table_Unit_SSI = "GDC"
        Case "S40": Table_Unit_SSI = "INFO"
        Case "S41": Table_Unit_SSI = "DCOM"
        Case "S42": Table_Unit_SSI = "DEON"
        Case "S43": Table_Unit_SSI = "CONF"
        Case "S51": Table_Unit_SSI = "DER"
        Case "S52": Table_Unit_SSI = "JURI"
        Case "S53": Table_Unit_SSI = "DTX"
        Case "S54": Table_Unit_SSI = "FOTC"
        Case "S60": Table_Unit_SSI = "CPT"
        Case "S61": Table_Unit_SSI = "DRH"
        Case "S99": Table_Unit_SSI = "S99"
        Case Else: Table_Unit_SSI = "S00"
    End Select
End If

End Function
