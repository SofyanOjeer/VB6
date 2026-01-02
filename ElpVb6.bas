Attribute VB_Name = "ElpVb6"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
' Err = 9999 Method ?
'       9998 NoMatch
'       9997 Bof
'       9996 Eof
'       9922 Existe déjà
'       9923 N'existe pas
'-----------------------------------------------------
'Dim V
'On Error GoTo Error_Handler

'Exit Sub
'------------------------------------------
'Error_Handler:
'    V = Error
'Error_MsgBox:
'    MsgBox V, vbCritical, frmElp_Caption & App_Debug

'Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
'Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
'SetWindowRgn hWnd, CreateEllipticRgn(0, 0, 600, 500), True
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long
Public Declare Function Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) As Long

      Public Declare Function SendMessage Lib "user32" Alias _
         "SendMessageA" _
         (ByVal hwnd As Long, _
          ByVal wMsg As Long, _
          ByVal wParam As Long, _
          lParam As Any) As Long

Public Const EM_GETLINECOUNT = &HBA
'___________________________________________________________________________________________
'$JPL 2012-10-16
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long


'___________________________________________________________________________________________
'$JPL 2012-05-09
 Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" _
()
'___________________________________________________________________________________

Global Const TH32CS_SNAPPROCESS = &H2
Global Const MAX_PATH = 260

Public Type PROCESSENTRY32
dwSize As Long
cntUsage As Long
th32ProcessID As Long
th32DefaultHeapID As Long
th32ModuleID As Long
cntThreads As Long
th32ParentProcessID As Long
pcPriClassBase As Long
dwFlags As Long
szExeFile As String * MAX_PATH
End Type

Public Declare Function CreateToolhelp32Snapshot Lib "kernel32" _
(ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" _
(ByVal hObject As Long) As Long
Public Declare Function Process32First Lib "kernel32" _
(ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Boolean
Public Declare Function Process32Next Lib "kernel32" _
(ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Boolean

'_______________________________________________________________________________________
Public cnMDB As New ADODB.Connection
Public rsMDB As New ADODB.Recordset


Public msFileSystem, msFile
Public blnRéplication_Load  As Boolean
Public App_EXEName As String
Public App_Title As String
Public App_Debug As String
Public App_Error As String

Public srvIdle As Boolean
Public tauxTVA As Single, tauxTTC As Single, tauxTVA_Lib As String
Public Const Pi = 3.14159
Public Const picLineHeight = 320
Public Const lstLineHeight = 255
Public lX As Long, Vx As Variant
Public libMois(12) As String, libMonth(12) As String
Public Const constàCompta = "à Compta"
Public Const constàCompta_Valider = "à Compta_Val"
Public Const constàCompta_Annuler = "à Compta_Ann"
Public Const constCompta = "Compta"
Public Const constComptaHB = "ComptaHB"
Public Const constSwiftSnd = "SwiftSnd"
Public Const constSwiftRcv = "SwiftRcv"
Public Const constSaisie = "Saisie"
Public Const constSaisieG = "SaisieG"
Public Const constAjouter = "Ajouter"
Public Const constàValider = "à Valider"
Public Const constàModifier = "à Modifier"
Public Const constModifier = "Modifier"
Public Const constValider = "Valider"
Public Const constInvalider = "Invalider"
Public Const constAnnuler = "Annuler"
Public Const constEffacer = "Effacer"
Public Const constSeek = "Seek=       "
Public Const constAddNew = "AddNew      "
Public Const constUpdate = "Update      "
Public Const constDelete = "Delete      "
Public Const constIgnore = "Ignore      "
Public Const constEnAttente = "en Attente  "
Public Const constAnnulé = "Annulé      "
Public Const constExtourne = "Extourne    "
Public Const constcmdAbandonner = "&Abandonner"
Public Const constcmdRechercher = "&Rechercher"
Public Const constcmdEnregistrer = "&Enregistrer"
Public Const constHistorique = "Historique"
Public Const constEnCours = "En cours"
Public Const constDisplay = "Display"
Public Const constAller = "Aller"
Public Const constRetour = "Retour"

Public Const constAuto = "Auto"
Public Const constProduction = "Production"
Public Const constTest = "Test"
Public Const constArchive = "Archive"
Public Const constSystem = "System"
Public Const constCorbeille = "Corbeille"
Public Const constLog = "Log"
Public Const constXCom = "XCom"
Public Const constExe = "Exe"
Public Const constFTP = "FTP"
Public paramEnvironnement As String

Public constWinWord As String, constWinWord_D As String
Public constExcel As String, constExcel_D As String
Public constAcrord32 As String, constAcrord32_D As String
Public constWordPad As String
Public constMsPaint As String
Public constIExplorer As String

Public DataBase_Open As String, DataBase_Master As String, DataBase_Local As String
Public DataBase_Data As String
Public paramFolder_Master As String, paramFolder_Local As String
Public paramDataBase_Password As String

Public KeyAscii As Integer
Public Asc01 As String * 1
Public Asc03 As String * 1
Public Asc10 As String * 1
Public Asc13 As String * 1
Public Asc10_13 As String * 2
Public Asc34 As String * 1
Public Asc39 As String * 1
Public Asc123 As String * 1
Public Asc125 As String * 1
Public Asc232 As String * 1
Public Asc233 As String * 1

Public TSys As String * 6
Public DSys As String * 8, DSys_S As String

Public valDSys As Long, dateDSys As Date
Public DSys_VeilleC As String * 8, DSys_VeilleO As String * 8, DSys_VeilleOAP As String * 8
Public DSys_SuivantC As String * 8, DSys_SuivantO As String * 8
Public usrId As String, usrIdNT As String
Public usrName As String, usrName_UCase As String, usrName_UCase10 As String, usrName_ULCase As String
Public usrService As String
Public usrCompte As String * 11, usrRacine As String * 5
Public usrGestionnaire As String
Public pcIdUsrIdCtl As Boolean
Public usrSituationCompte_Forçage As Boolean
Public usrService_DisplayAll As Boolean
Public SrvDir As String
Public Elp As typeXcom
'-----------------------------------------------------

Public socName As String, SocRibDom As String, socTéléphone As String, socSiren As String
Public SocBdfE As Integer, strSocBdfE As String * 5
Public SocBdfG As Integer, strSocBdfG As String * 5
Public SocId$
Public SocAgence$
Public paramBic8 As String * 8, SocBicId As String * 11, SocBicIdNostro As String * 11
Public strSocSignon As String
Public paramSocLogo_G As String
Public paramSocLogo_PiedPage As String
Public paramSocLogo As String
Public imgGuichet As String
    
Public errTag
Public oldText
Public intReturn As Integer, intReturn2 As Integer, intReturn3 As Integer
   
   
Public dateAAmin As Long
Public dateAAmax As Long
Public dateSerialMin As Long
Public dateExerDeb As Long
Public dateExerFin As Long
Public xobj
Public XForm As Form
Public XPrt, XPrt_Previous, Printer_Previous_DeviceName As String
Public XListBox As ListBox
Public XLabel As Label
Public XControl As Control
Public xImage As Image
'-----------------------------------------------------
Public frmElp_Caption As String
Public frmElp_Icon As Variant
'-----------------------------------------------------
Public prtCollection_Index As Integer
Public prtSocSigle As Boolean
Public prtKillDoc As Boolean
Public prtOrientation As Integer
Public prtPaperSize As Integer
Public prtCurrentX As Integer
Public prtMaxX As Integer
Public prtMinX As Integer
Public prtMinX1 As Integer, prtMinX2 As Integer, blnMinX As Boolean, blnMinX12 As Boolean
Public prtMinMarge As Integer, prtMaxMarge As Integer
Public prtWidth As Integer, prtWidthMarge As Integer
Public prtMedX As Integer, prtMedX0 As Integer
Public prtCurrentY As Integer
Public prtMaxY As Integer, prtMaxLine As Integer
Public prtMinY As Integer
Public prtMedY As Integer
Public prtFontName As String, prtFontNameZ As String
Public prtFontSize As String
Public prtTitleText As String
Public prtTitleUsr As String
Public prtPgmName As String
Public Const prtFontName_CenturyGothic = "Century Gothic"
Public Const prtFontName_Comic = "Comic Sans MS"
Public Const prtFontName_Arial = "Arial"
Public Const prtFontName_TimesNewRoman = "Times New Roman"
Public Const prtFontName_CourierNew = "Courier New"

Public prtColorMode As Boolean
Public Const prtForeColor_Header = &HFF0000    ' vbBlack '12615680 '
Public Const prtForeColor = vbBlack

Public prtLineColor  As Long
Public Const prtLineColor_Standard = 12615680
Public Const prtLineColor_Black = vbBlack

Public prtFillColor As Long
Public Const prtFillColor_Standard = 16777210 'RGB(250, 255, 255) ' Logo RGB(0, 160, 182) ' RGB(0, 123, 141)
Public Const prtFillColor_Black = 15790320 'rgb(240,240,240)

Public prtFormType As String
Public prtLineNb As Integer
Public prtlineHeight As Integer, prtlineHeight66 As Integer
Public prtHeaderHeight As Integer
Public prtParagraphHeight As Integer
Public prtZoom As Integer
Public prtShow As Boolean
Public blnFiligrane As Boolean
Public prtFiligrane_Name As String
Public prtFiligrane_Color  As Long

Public Const paramIMP_PDF_Path_Temp = "C:\Temp\IMP_PDF"
Public paramIMP_PDF_Path   As String, paramIMP_PDF_Path_VBP  As String '= "C:\Temp\IMP_PDF\BIA_SAB"
Public paramIMP_PDFCreator_Name   As String

Public Const paramElpCypher = "AntiHacker"
Public blnIMP_PDF As Boolean
Public prtIMP_PDF_FileName As String
'---------------------------------------------------------
Type typeDate
    AA  As String * 4
    MM  As String * 2
    jj As String * 2
End Type
'---------------------------------------------------------
Type typeUsrColor
    BackColor  As Long
    ForeColor  As Long
End Type
    
Public errUsr As typeUsrColor
Public frmUsr As typeUsrColor
Public lblUsr As typeUsrColor
Public libUsr As typeUsrColor
Public lstUsr As typeUsrColor
Public picUsr As typeUsrColor
Public txtUsr As typeUsrColor
Public focusUsr As typeUsrColor
Public dbUsr As typeUsrColor
Public crUsr As typeUsrColor
Public MouseMoveUsr As typeUsrColor
Public greenColor As typeUsrColor

Public warnUsrColor As Long
Public blnBeep As Boolean
Public blnTimer_Enabled As Boolean
Public blnNetSend_Enabled As Boolean
Public frmUsr_Windowstate As Integer
Public mCommand As String

Public countTimer As Long
Public blnOff_Line As Boolean
Public paramTemp_Folder As String
Public blnAuto_Form_Show As Boolean

'---------------------------------------------------------
Public elpSrvXcom  As String
Public elpSrvTxtin As Boolean
Public elpSrvTxtOut As Boolean

Type typeXcom
   SrvObj       As String * 12
   SrvMethod    As String * 12
   SrvErr       As String * 10
   usrId       As String '* 10
   pcId        As String * 10
   SrvType     As String * 10
   SrvId       As String * 10
   SrvDtaqLib  As String * 10
   SrvDtaqIn   As String * 10
   SrvDTaqOut  As String * 10
   SrvDTaqLen  As String * 5
   jplFree     As String * 5
End Type

Public Xcom As typeXcom

Public arrX2P, arrP2X, arrX2P_D(99) As Integer, arrX2P_DF(9) As Integer

Public paramIBM_Library_JPL073 As String, paramIBM_Library_JPL073SPE As String


Public mColor_Z0 As Long, mColor_G0 As Long, mColor_W0 As Long, mColor_W1 As Long
Public mColor_Y0 As Long, mColor_Y1 As Long, mColor_Y2 As Long, mColor_Y3 As Long
Public mColor_GB As Long, mColor_G1 As Long, mColor_G2 As Long, mColor_G9 As Long
Public mColor_B0 As Long, mColor_B1 As Long, mColor_B9 As Long

Public htmlFontColor_Blue As String, htmlFontColor_Green As String, htmlFontColor_Gray As String, htmlFontColor_Red As String
Public htmlFontColor_Black As String, htmlFontColor_White As String, htmlFontColor_Magenta As String

Type typeX_Stat
   Code       As String
   Lib        As String
   Row1       As Long
   Row2       As Long
   col1       As Long
   col2       As Long
End Type

Public blnBIA_VB_AIB As Boolean

'Pour accéder au REGISTRE '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long          ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Any, lpcbData As Long) As Long
Public Const REG_SZ As Long = 1
Public Const REG_BINARY As Long = 3
Public Const REG_DWORD As Long = 4
Public Const REG_DWORD_BIG_ENDIAN = 5           ' 32-bit number
Public Const REG_DWORD_LITTLE_ENDIAN = 4        ' 32-bit number (same as REG_DWORD)

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const KEY_ALL_ACCESS = &H3F
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
'   FindExecutable()                                                                                                                                                        '
Private Const MAX_FILENAME_LEN = 260
Public Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long
Public Function FindDOCX() As String
Dim I As Long
Dim s2 As String
Dim sFile
Dim retour As String
 
    retour = "No association found !"
    'create a temp file
    sFile = App.path & "\a_garder.docx"
   
    'Create a buffer
    s2 = String(MAX_FILENAME_LEN, 32)
    'Retrieve the name and handle of the executable
    I = FindExecutable(sFile, vbNullString, s2)
    If I > 32 Then
        retour = Left(s2, InStr(s2, Chr$(0)) - 1)
    Else
    End If
    
    FindDOCX = retour
    
End Function

Public Function FindPDF() As String
Dim I As Long
Dim s2 As String
Dim sFile
Dim retour As String
 
    retour = "No association found !"
    'create a temp file
    sFile = App.path & "\a_garder.pdf"
   
    'Create a buffer
    s2 = String(MAX_FILENAME_LEN, 32)
    'Retrieve the name and handle of the executable
    I = FindExecutable(sFile, vbNullString, s2)
    If I > 32 Then
        retour = Left(s2, InStr(s2, Chr$(0)) - 1)
    Else
    End If
    
    FindPDF = retour
    
End Function

Public Sub pause_with_events(secondes As Long)
Dim dt As Date

    dt = Now
    Do While DateDiff("s", dt, Now) < secondes
        DoEvents
        Sleep 50   ' put your app to sleep in small increments
                   ' to avoid having CPU go to 100%
    Loop
    
End Sub


Public Function Retourne_Num_Client(COMPTE As String) As String
Dim xSql As String
Dim rsDenis As ADODB.Recordset

    Retourne_Num_Client = ""
    xSql = "select CLIENACLI from " & paramIBM_Library_SABSPE & ".YBIACPT0 where COMPTECOM = '" & COMPTE & "'"
    Set rsDenis = cnsab.Execute(xSql)
    If Not rsDenis.EOF Then
        Retourne_Num_Client = rsDenis("CLIENACLI")
    End If
    If rsDenis.State = adStateOpen Then
        rsDenis.Close
    End If
    Set rsDenis = Nothing

End Function

Public Function CHAINE_Rpad_0(ByVal z As String) As String
'Tronque un string avant le 1er chr(0) rencontré
Dim I As Long

    CHAINE_Rpad_0 = ""
    For I = 1 To Len(z)
        If Asc(Mid(z, I, 1)) = 0 Then
            Exit For
        Else
            CHAINE_Rpad_0 = CHAINE_Rpad_0 & Mid(z, I, 1)
        End If
    Next I
    
End Function
Public Sub REGISTRE_Ecrit_LONG(ByVal larubrique As Long, ByVal lentree As String, ByVal laclef As String, ByVal LaValeur As Long)
'Exemple de rubrique = HKEY_CURRENT_USER = &H80000001
'Exemple d'entree = "AppEvents\Schemes\Apps\.Default\.Default\.Current"
Dim lRetval As Long
Dim hKey As Long
Dim szBuffer As Long
Dim lBuffSize As Long

    szBuffer = LaValeur
    lBuffSize = 4
    lRetval = RegOpenKeyEx(larubrique, lentree, 0, KEY_ALL_ACCESS, hKey)
    lRetval = RegSetValueEx(hKey, laclef, 0, REG_DWORD, szBuffer, lBuffSize)
    RegCloseKey (hKey)

End Sub
Public Sub REGISTRE_Ecrit_BINAIRE(ByVal larubrique As Long, ByVal lentree As String, ByVal laclef As String, ByVal LaValeur As Variant)
'Exemple de rubrique = HKEY_CURRENT_USER = &H80000001
'Exemple d'entree = "AppEvents\Schemes\Apps\.Default\.Default\.Current"
Dim lRetval As Long
Dim hKey As Long
Dim szBuffer As String
Dim lBuffSize As Long

    szBuffer = LaValeur
    lBuffSize = Len(LaValeur)
    lRetval = RegOpenKeyEx(larubrique, lentree, 0, KEY_ALL_ACCESS, hKey)
    lRetval = RegSetValueEx(hKey, laclef, 0, REG_BINARY, ByVal szBuffer, lBuffSize)
    RegCloseKey (hKey)

End Sub
Public Sub REGISTRE_Ecrit_STRING(ByVal larubrique As Long, ByVal lentree As String, ByVal laclef As String, ByVal LaValeur As String)
'Exemple de rubrique = HKEY_CURRENT_USER = &H80000001
'Exemple d'entree = "AppEvents\Schemes\Apps\.Default\.Default\.Current"
Dim lRetval As Long
Dim hKey As Long
Dim szBuffer As String
Dim lBuffSize As Long

    szBuffer = LaValeur
    lBuffSize = Len(szBuffer)
    lRetval = RegOpenKeyEx(larubrique, lentree, 0, KEY_ALL_ACCESS, hKey)
    lRetval = RegSetValueEx(hKey, laclef, 0, REG_SZ, ByVal szBuffer, lBuffSize)
    RegCloseKey (hKey)

End Sub
Public Sub REGISTRE_Lit(ByVal larubrique As Long, ByVal lentree As String, ByVal valueName As String, leretour As String)
'Exemple de rubrique = HKEY_CURRENT_USER = &H80000001
'Exemple d'entree = "AppEvents\Schemes\Apps\.Default\.Default\.Current"
Dim lRetval As Long
Dim hKey As Long
Dim szBuffer As String
Dim lBuffSize As Long

    szBuffer = Space(255)
    lBuffSize = Len(szBuffer)

    lRetval = RegOpenKeyEx(larubrique, lentree, 0, KEY_ALL_ACCESS, hKey)
    lRetval = RegQueryValueEx(hKey, valueName, 0, REG_SZ, szBuffer, lBuffSize)
    RegCloseKey (hKey)
    DoEvents
    leretour = CHAINE_Rpad_0(szBuffer)
    
End Sub

Public Function get_PDFCreator_AutosaveFilename() As String
Dim larubrique As Long
Dim lentree As String
Dim leretour As String
Dim laclef As String

    get_PDFCreator_AutosaveFilename = "?"
    If nomDuServeur = paramServerSplf Then Exit Function
    larubrique = HKEY_LOCAL_MACHINE
    lentree = "SOFTWARE\PDFCreator\Program"
    laclef = "AutosaveFilename"
    leretour = ""
    Call REGISTRE_Lit(larubrique, lentree, laclef, leretour)
    If Trim(leretour) <> "" Then
        get_PDFCreator_AutosaveFilename = leretour
    Else
        MsgBox "Clef de registre introuvable !"
    End If
    
End Function
Public Function Retourne_WAIT_PDF() As Long
Dim xSql As String
Dim rs As ADODB.Recordset

    Retourne_WAIT_PDF = 3
    xSql = "select BIATABTXT from " & paramIBM_Library_SABSPE & ".YBIATAB0"
    xSql = xSql & " where BIATABID = 'PDF_WAIT' and BIATABK1 = 'SECONDS'"
    xSql = xSql & " and BIATABK2='valeur'"
    Set rs = cnsab.Execute(xSql)
    If Not rs.EOF Then
        Retourne_WAIT_PDF = CLng(rs("BIATABTXT"))
    End If
    rs.Close
    Set rs = Nothing

End Function

Public Sub set_PDFCreator_AutosaveFilename(z As String)
Dim larubrique As Long
Dim lentree As String
Dim laclef As String

    If nomDuServeur <> paramServerSplf Then
        larubrique = HKEY_LOCAL_MACHINE
        lentree = "SOFTWARE\PDFCreator\Program"
        laclef = "AutosaveFilename"
        Call REGISTRE_Ecrit_STRING(larubrique, lentree, laclef, z)
    End If

End Sub
Public Sub killProcessDotNet(z As String)

    If nomDuServeur = paramServerSplf Then
        Shell ("d:\BIASRV.APP\Production\killProcess\killProcess.exe " & z)
    Else
        Shell ("c:\BIASRV\killProcess.exe " & z)
    End If
    
End Sub

Public Function retourne_Client_CLOS(codeClient As String, nomClient As String) As Boolean
Dim xSql As String
Dim rs As ADODB.Recordset

    retourne_Client_CLOS = False
    If InStr(UCase(nomClient), "CLOS") > 0 Then
        retourne_Client_CLOS = True
        Exit Function
    End If
    If InStr(UCase(nomClient), "CLOTURE") > 0 Then
        retourne_Client_CLOS = True
        Exit Function
    End If
    xSql = "SELECT CLIRELDAF FROM " & paramIBM_Library_SAB & ".ZCLIREL0 WHERE CLIRELCLI = '" & codeClient & "'"
    Set rs = cnsab.Execute(xSql)
    If Not rs.EOF Then
        If CLng(rs("CLIRELDAF")) > 0 Then
            retourne_Client_CLOS = True
        End If
    End If
    If rs.State = adStateOpen Then
        rs.Close
    End If
    Set rs = Nothing

End Function

Public Function Retourne_Num_Document(tabId As String, tabTxt As String) As String
Dim xSql As String
Dim rs As ADODB.Recordset

    Retourne_Num_Document = ""
    xSql = "select BIATABK2 from " & paramIBM_Library_SABSPE & ".YBIATAB0"
    xSql = xSql & " where BIATABID = '" & tabId & "' and BIATABK1 = 'Courrier_Doc'"
    xSql = xSql & " and BIATABTXT = '" & tabTxt & "'"
    Set rs = cnsab.Execute(xSql)
    If Not rs.EOF Then
        Retourne_Num_Document = rs("BIATABK2")
    End If
    rs.Close
    Set rs = Nothing

End Function
Public Function dateAAAAMMJJTOJJ_MM_AAAA(lAMJ As String) As String
'en entrée 20140826 (26 août 2014)
'en sortie 26/08/2014

    dateAAAAMMJJTOJJ_MM_AAAA = Mid(lAMJ, 7, 2) & "/" & Mid(lAMJ, 5, 2) & "/" & Left(lAMJ, 4)
    
End Function

Public Sub FEU_ORANGE()
    
    frmElp.imgFEU(1).Top = frmElp.imgSocSignon.Top + 600
    frmElp.imgFEU(1).Visible = True
    frmElp.imgFEU(2).Visible = False
    frmElp.imgFEU(0).Visible = False
    DoEvents

End Sub

Public Sub FEU_ROUGE()

    frmElp.imgFEU(2).Top = frmElp.imgSocSignon.Top + 600
    frmElp.imgFEU(2).Visible = True
    frmElp.imgFEU(1).Visible = False
    frmElp.imgFEU(0).Visible = False
    DoEvents

End Sub


Public Sub FEU_VERT()

    frmElp.imgFEU(0).Top = frmElp.imgSocSignon.Top + 600
    frmElp.imgFEU(0).Visible = True
    frmElp.imgFEU(1).Visible = False
    frmElp.imgFEU(2).Visible = False
    DoEvents
    
End Sub

Public Sub MSflexGrid_Excel(lDir As String, lAPP As String, lTitle As String, fgX As MSFlexGrid, lCols As Integer)
On Error GoTo Error_Handler
Static mXls1_File As Long

Dim appExcel As Excel.Application 'Application Excel
Dim wbExcel As Excel.Workbook 'Classeur Excel
Dim wsExcel As Excel.Worksheet 'Feuille Excel

Dim wFile As String, wFilex As String
Dim mXls1_Row As Long, K As Long, K2 As Long, wForecolor As Long, wBackColor As Long
Dim X As String
Dim blnCALCS As Boolean
Dim arrNumberFormat() As String, kDec As Integer, lDec As Integer

On Error GoTo Error_Handler
ReDim arrNumberFormat(lCols + 2)
'===================================================================================
If InStr(lDir, ".xls") > 0 Then
    wFile = lDir
Else
    If lDir = "" Then lDir = "C:\Temp\"
    
    mXls1_File = mXls1_File + 1
    
    wFile = lDir & Trim(lAPP & " " & DSYS_Time & mXls1_File & ".xlsx")
    '______________________________________________
        X = InputBox("par défaut : " _
            & vbCrLf & "     =========================" & vbCrLf & wFile _
            & vbCrLf & "     =========================", "Nom du fichier d'exportation (" & fgX.Rows & " lignes)", wFile)
        If Trim(X) = "" Then mXls1_File = mXls1_File - 1: Exit Sub
        wFilex = Trim(X)
        '______________________________________________
        If wFile <> wFilex Then
            wFile = wFilex
        End If
End If
'_________________________________________


If Dir(wFile) <> "" Then msFileSystem.DeleteFile wFile

'=========================================================================================
'Call lstErr_AddItem(lstErr, cmdContext, "Fichier excel.... : "): DoEvents

Set appExcel = CreateObject("Excel.Application")
appExcel.Workbooks.Add
Set wbExcel = appExcel.ActiveWorkbook
With wbExcel
    .Title = lAPP
    .Subject = lTitle
End With

'__________________________________________________________________________________

'appExcel.Worksheets.Add

Set wsExcel = wbExcel.Sheets(1): wsExcel.Name = lAPP

Set wsExcel = wbExcel.Sheets(1)

With wsExcel.Cells
    .Borders.Weight = xlWide
    .Borders.Color = RGB(255, 255, 153)
    .Borders(xlInsideHorizontal).Weight = xlThin ' xlMedium
    .Borders(xlInsideHorizontal).Color = RGB(128, 128, 255)
    .Borders(xlInsideVertical).Weight = xlThin
    .Borders(xlInsideVertical).Color = RGB(128, 128, 255)
    .HorizontalAlignment = Excel.xlHAlignLeft
    .WrapText = False ' True
    .Font.Size = 10
    .Font.Name = "Calibri"
    .RowHeight = 17
End With

wsExcel.PageSetup.Orientation = vbPRORLandscape
wsExcel.PageSetup.HeaderMargin = 5
'wsExcel.PageSetup.Zoom = 75
wsExcel.PageSetup.Zoom = False
wsExcel.PageSetup.FitToPagesWide = 1
wsExcel.PageSetup.FitToPagesTall = False

wsExcel.PageSetup.CenterHeader = vbCr & "&B&U&14" & lTitle _
                                 & vbCr & "  (édité le " & dateImp10(DSys) & " " & Time & ")" & vbCr


wsExcel.PageSetup.CenterHorizontally = True


wsExcel.PageSetup.PrintTitleRows = "$A1:$K1"
wsExcel.PageSetup.LeftFooter = "&F - &B&A"
wsExcel.PageSetup.RightFooter = "&P / &N"

mXls1_Row = 0

fgX.Visible = False

    For K = 0 To fgX.Rows - 1
        fgX.Row = K
        

        mXls1_Row = mXls1_Row + 1
        For K2 = 0 To lCols
        
            fgX.Col = K2: X = fgX.Text
            If K = 0 Then
                wsExcel.Columns(K2 + 1).ColumnWidth = fgX.CellWidth / 100
                'If K2 > 0 Then wsExcel.Columns(K2 + 1).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
              Select Case fgX.CellAlignment
                Case Is = 0
                Case Is = 1: wsExcel.Columns(K2 + 1).HorizontalAlignment = Excel.xlHAlignRight
                Case Is = 2: wsExcel.Columns(K2 + 1).HorizontalAlignment = Excel.xlHAlignCenter
              End Select
          End If
          
          If wsExcel.Columns(K2 + 1).HorizontalAlignment = Excel.xlHAlignRight _
          And IsNumeric(X) Then
                If arrNumberFormat(K2 + 1) = "" Then
                    X = Trim(X)
                    kDec = InStr(X, ",")
                    If kDec = 0 Then kDec = InStr(X, ".")
                    If kDec > 0 Then
                        lDec = Len(X) - kDec
                        arrNumberFormat(K2 + 1) = "### ### ##0." & String(lDec, "0")
                        wsExcel.Columns(K2 + 1).NumberFormat = arrNumberFormat(K2 + 1)
                    End If
                End If
                wsExcel.Cells(mXls1_Row, K2 + 1) = num_CDec(X) 'num_String_Auto(X)  à tester
          Else
                wsExcel.Cells(mXls1_Row, K2 + 1) = X
          End If
          
            wForecolor = fgX.CellForeColor
            If wForecolor = 0 Then
                If K = 0 Then
                    wForecolor = fgX.ForeColorFixed
                Else
                    wForecolor = fgX.ForeColor
                End If
            End If
            
            wBackColor = fgX.CellBackColor
            If wBackColor = 0 Then
                If K = 0 Then
                    wBackColor = fgX.BackColorFixed
                Else
                    wBackColor = fgX.BackColor
                End If
            End If

            wsExcel.Cells(mXls1_Row, K2 + 1).Font.Color = wForecolor
            wsExcel.Cells(mXls1_Row, K2 + 1).Interior.Color = wBackColor
            If fgX.CellFontBold = True Then wsExcel.Cells(mXls1_Row, K2 + 1).Font.Bold = True
            If fgX.CellFontItalic = True Then wsExcel.Cells(mXls1_Row, K2 + 1).Font.Italic = True
            If fgX.CellFontUnderline = True Then wsExcel.Cells(mXls1_Row, K2 + 1).Font.Underline = True
            wsExcel.Cells(mXls1_Row, K2 + 1).Font.Size = fgX.CellFontSize
            wsExcel.Cells(mXls1_Row, K2 + 1).Font.Name = fgX.CellFontName
        Next K2
    Next K


'======================================================================================================

'__________________________________________________________________________________

fgX.Visible = True

'======================================================================================================

Exit_sub:
'__________________________________________________________________________________
Set rsSab = Nothing


wbExcel.SaveAs wFile

wbExcel.Close

'____________________________________________________________________________________
appExcel.Quit

Set rsSab = Nothing

Set wsExcel = Nothing
Set wbExcel = Nothing
Set appExcel = Nothing
'Call lstErr_AddItem(lstErr, cmdContext, "< Exportation terminée"): DoEvents


'_____________________________
Exit Sub

Error_Handler:
    If Not blnTimer_Enabled Then MsgBox Error, vbCritical, lAPP & " MSflexGrid_Excel"
    'Call lstErr_AddItem(lstErr, cmdContext, "< Exportation terminée"): DoEvents
    
    wbExcel.SaveAs wFile
    wbExcel.Close
    appExcel.Quit

End Sub

Public Sub MSFlexGrid_SendMail(lDest As String, lAPP As String, lSubject As String, lTitle As String, fgX As MSFlexGrid, lCols As Integer, Optional lAttachment As String)
Dim wSendMail As typeSendMail
Dim xDetail As String, mbgColor As String
Dim K As Long, htmlFontColor_K As String
Dim iRow As Integer, iCol As Integer, X As String, xTD As String, xAlign As String
Dim wForecolor As String, wBackColor As String, xColor As String
Dim arrCellAlignment() As String, wCellFontSize As Integer, arrCellWidth() As Long
Dim mWidth As Long
Dim wTag1 As String, wTag2 As String
On Error Resume Next

fgX.Visible = False
ReDim arrCellAlignment(lCols), arrCellWidth(lCols)

fgX.Row = 0
For iCol = 0 To lCols
    fgX.Col = iCol
    arrCellWidth(iCol) = fgX.CellWidth
    mWidth = mWidth + fgX.CellWidth
Next iCol
For iCol = 0 To lCols
    arrCellWidth(iCol) = arrCellWidth(iCol) * 100 / mWidth
    Debug.Print iCol; arrCellWidth(iCol)
Next iCol


xDetail = ""
For iRow = 0 To fgX.Rows - 1
    fgX.Row = iRow
    xTD = ""
    For iCol = 0 To lCols
        fgX.Col = iCol

        X = Trim(fgX.Text)
        If X = "" Then X = "&#160"
        wCellFontSize = fgX.CellFontSize
        If iRow = 0 Then
            Select Case fgX.CellAlignment
                Case Is = 0: arrCellAlignment(iCol) = " align=left"
                Case Is = 1: arrCellAlignment(iCol) = " align=right"
                Case Is = 2: arrCellAlignment(iCol) = " align=center"
             End Select
            
            wForecolor = RGB_Html_Color(fgX.ForeColorFixed)
            wBackColor = RGB_Html_Color(fgX.BackColorFixed)
            xTD = xTD _
                 & "<TD bgcolor=" & wBackColor & arrCellAlignment(iCol) & " width=" & arrCellWidth(iCol) & "%><span style='font-size:" & wCellFontSize & ".0pt;font-family:Calibri'><Font color=" & wForecolor & "><B>" _
                 & X & "</B/TD>"
        Else
            If fgX.CellForeColor <> 0 Then
                wForecolor = RGB_Html_Color(fgX.CellForeColor)
            Else
                wForecolor = RGB_Html_Color(fgX.ForeColor)
            End If
            If fgX.CellBackColor <> 0 Then
                wBackColor = RGB_Html_Color(fgX.CellBackColor)
            Else
                wBackColor = RGB_Html_Color(fgX.BackColor)
            End If
            wTag1 = "": wTag2 = ""
            If fgX.CellFontBold = True Then wTag1 = wTag1 & "<B>": wTag2 = wTag2 & "</B>"
            If fgX.CellFontItalic = True Then wTag1 = wTag1 & "<i>": wTag2 = wTag2 & "</i>"
            If fgX.CellFontUnderline = True Then wTag1 = wTag1 & "<U>": wTag2 = wTag2 & "</U>"
                xTD = xTD _
                     & "<TD bgcolor=" & wBackColor & arrCellAlignment(iCol) & " width=" & arrCellWidth(iCol) & "%><span style='font-size:" & wCellFontSize & ".0pt;font-family:Calibri'><Font color=" & wForecolor & ">" _
                     & wTag1 & X & wTag2 & "</TD>"
        End If
    Next iCol
    xDetail = xDetail & "<TR>" & xTD & "</TR>"
    

Next iRow

mbgColor = "bgcolor = #E0E0E0"

wSendMail.From = currentSSIWINMAIL
wSendMail.FromDisplayName = lAPP
wSendMail.Recipient = lDest

wSendMail.Subject = lSubject
If IsMissing(lAttachment) Then
    wSendMail.Attachment = ""
Else
    wSendMail.Attachment = lAttachment
End If

wSendMail.Message = "<body bgcolor = #FFFFFF><span style='font-size:10.0pt;font-family:Calibri'>" _
                    & htmlFontColor_Blue & lTitle & "<BR><BR>" _
                    & "<TABLE  width=100% border=0 cellspacing=3  cellpadding=0 ></B>" _
                    & xDetail _
                    & "</div></TABLE>"

wSendMail.AsHTML = True

srvSendMail.Monitor wSendMail
fgX.Visible = True

End Sub

Public Function ElpCipher_C(mMsg As String, mKey As String) As String
Dim L1 As Integer, L2 As Integer, I As Integer
Dim wMsg As String, wKey As String

L1 = Len(mMsg)
wKey = mKey
Do
    L2 = Len(wKey)
    If L2 <= L1 Then wKey = wKey & mKey
Loop Until L2 > L1
wMsg = ""
For I = 1 To L1
    wMsg = wMsg & Format$(Asc(Mid$(mMsg, I, 1)) Xor Asc(Mid$(wKey, I, 1)), "000")
Next I
ElpCipher_C = wMsg
End Function

Public Function ElpCipher_D(mMsg As String, mKey As String) As String
Dim L1 As Integer, L2 As Integer, I As Integer
Dim wMsg As String, wKey As String

L1 = Fix((Len(mMsg) - 1) / 3) + 1
wKey = mKey
Do
    L2 = Len(wKey)
    If L2 <= L1 Then wKey = wKey & mKey
Loop Until L2 > L1
wMsg = ""
L2 = 1
For I = 1 To L1
    wMsg = wMsg & Chr$(Val(Mid$(mMsg, L2, 3)) Xor Asc(Mid$(wKey, I, 1)))
    L2 = L2 + 3
Next I
ElpCipher_D = wMsg
End Function

Public Sub Shell_Exe(lFileName)
Dim IdShell As Variant, X As String
'X = "CMD/Q " & Chr$(34) & lFilename & Chr$(34)
IdShell = Shell(lFileName, 0)
If IdShell > 0 Then
'    AppActivate IdShell
Else
    MsgBox lFileName, vbCritical, "Shell_Exe"
End If

DoEvents
'AppActivate App.EXEName
'SetFocus
End Sub
Public Sub Shell_FTP(lFTP_Nt_Filename As String, lFTP_AS400_Library As String, lFTP_AS400_File As String, blnFTP_Get As Boolean, blnFTP_AS400_Binary As Boolean)
Dim wFTP_Nt_Filename As String, wFTP_Nt_Filename_Bat As String, wFTP_Nt_Filename_Dta As String
Dim K As Integer, K1 As Integer
Dim wFTP_Nt_Filename_Log As String
Dim IdShell
Dim iWait As Long, xIn As String, blnOk As Boolean, blnLog_Quit As Boolean
Dim fsoFile As File
Dim wAMJHMS As String
On Error GoTo Error_Handle
Dim X As String
Dim xTrace As String

wFTP_Nt_Filename = Trim(lFTP_Nt_Filename)
wAMJHMS = "_" & DSys & "_" & time_Hms
xTrace = " : Init "
K1 = Len(wFTP_Nt_Filename)
K = InStr(K1 - 5, wFTP_Nt_Filename, ".")
If K > 0 Then K1 = K - 1
X = Mid$(wFTP_Nt_Filename, 1, K1)
    
wFTP_Nt_Filename_Bat = X & wAMJHMS & ".bat"
wFTP_Nt_Filename_Dta = X & wAMJHMS & ".dta"
wFTP_Nt_Filename_Log = X & wAMJHMS & ".log"
Open wFTP_Nt_Filename_Dta For Output As #2
Print #2, paramIBM_BIA_Auto
Print #2, paramIBM_BIA_AUTO_Password
If blnFTP_AS400_Binary Then Print #2, "bin"
Print #2, "cd " & Trim(lFTP_AS400_Library)
If blnFTP_Get Then
    Print #2, "get " & Trim(lFTP_AS400_File) & " " & wFTP_Nt_Filename
Else
    Print #2, "put " & wFTP_Nt_Filename & " " & Trim(lFTP_AS400_File)
End If

Print #2, "quit"

Open wFTP_Nt_Filename_Bat For Output As #1
K1 = 1
Do
    K = InStr(K1, wFTP_Nt_Filename, "\")
    If K > 0 Then K1 = K + 1
Loop Until K <= 0

Print #1, "cd /d " & Mid$(wFTP_Nt_Filename, 1, K1 - 2)
Print #1, "FTP -s:" & wFTP_Nt_Filename_Dta & " " & paramIBM_AS400_FTP 'Trim("I5A7") 'paramIBM_AS400_ID
Print #1, "del " & wFTP_Nt_Filename_Dta
Print #1, "del " & wFTP_Nt_Filename_Bat

Close
xTrace = " : Shell "

blnOk = False
blnLog_Quit = False

IdShell = Shell(wFTP_Nt_Filename_Bat & " > " & wFTP_Nt_Filename_Log, 1)
DoEvents

If IdShell > 0 Then
    xTrace = " : AppActivate "
                                        On Error Resume Next            '$$$$$$$$$$$
                                        AppActivate IdShell, True       '$$$$$$$$$$$
                                        On Error GoTo Error_Handle      '$$$$$$$$$$$
    
    xTrace = " : Read Log "
    Sleep 1000
    Do
        For iWait = 1 To 30
            DoEvents
            If Dir(wFTP_Nt_Filename_Log) <> "" Then
                Open wFTP_Nt_Filename_Log For Input As #1
                Do Until EOF(1)
                    Line Input #1, xIn
                 ''
                If InStr(xIn, "File transfer completed successfully") Then blnOk = True
                If InStr(xIn, "QUIT subcommand received") Then blnLog_Quit = True
                 'If Trim(xIn) = "250 File transfer completed successfully." Then blnOk = True
                 'If Trim(xIn) = "221 QUIT subcommand received." Then blnLog_Quit = True
                Loop
                Close #1
                If blnLog_Quit Then Exit For
                Sleep 1000
            End If
           ' Debug.Print iWait, xIn
        Next iWait
        If Not blnLog_Quit Then
            X = MsgBox("Fin du processus (221 QUIT)  non détecté, Voulez-vous tester à nouveau la fin du transfert?", vbYesNo + vbQuestion + vbDefaultButton1, "Shell_FTP")
            If X = vbYes Then blnLog_Quit = True
        End If
    Loop Until blnLog_Quit
End If
Close

Sleep 1000

xTrace = " : Close Log "
If blnOk Then
    xTrace = " : Kill bat "
    If Dir(wFTP_Nt_Filename_Bat) <> "" Then msFileSystem.DeleteFile wFTP_Nt_Filename_Bat
    xTrace = " : Kill log "
    If Dir(wFTP_Nt_Filename_Log) <> "" Then msFileSystem.DeleteFile wFTP_Nt_Filename_Log

Else
    Shell_MsgBox "Shell_FTP : " & wFTP_Nt_Filename_Log & " : en erreur ", vbCritical, XForm.Caption, False
End If

DoEvents

paramIBM_AS400_FTP = paramIBM_AS400_ID

Exit Sub

Error_Handle:

Close
Shell_MsgBox wFTP_Nt_Filename & xTrace & ":" & Error, vbCritical, "Shell_FTP", False
paramIBM_AS400_FTP = paramIBM_AS400_ID

End Sub

Public Sub Shell_MsgBox(lMsg As String, lConst As Long, lTitle As String, blnNetSend As Boolean)
Dim wNt_Filename_Bat As String
Dim IdShell
Dim intFile As Integer
Dim X As String, I As Integer
On Error Resume Next  'GoTo Error_Handle

If blnTimer_Enabled Or blnNetSend Then
    wNt_Filename_Bat = paramTemp_Folder & "\" & DSys & "_" & time_Hms & "_Net_Send.bat"
    intFile = FreeFile(0)
    Open wNt_Filename_Bat For Output As #intFile
    '''x = usrId & " : " & Trim(lTitle) & " : " & lMsg
     ''X = usrId & " : " & lMsg
     X = lMsg
   For I = 1 To Len(X)
        Select Case Asc(Mid$(X, I, 1))
            
            Case Is < 32: Mid$(X, I, 1) = " "
            Case Is > 127: Mid$(X, I, 1) = " "
        End Select
    Next I
    Close intFile
    DoEvents
    
    IdShell = Shell(wNt_Filename_Bat, 1)
    
    If IdShell > 0 Then
        AppActivate IdShell, True       '$$$$$$$$$$$
    End If
    DoEvents
    Wait_SS 5
    If Dir(wNt_Filename_Bat) <> "" Then msFileSystem.DeleteFile wNt_Filename_Bat
  ''' Kill wNt_Filename_Bat

End If

DoEvents

If Not blnTimer_Enabled Then MsgBox lMsg, lConst, lTitle
Exit Sub

Error_Handle:
Close
If Not blnTimer_Enabled Then MsgBox wNt_Filename_Bat & ":" & Error, vbCritical, "Shell_MsgBox"

End Sub

Public Sub DTPicker_Amj8_tiret(C As DTPicker, lAmj8_tiret As String)
Dim X8 As String * 8

Call DTPicker_Control(C, X8)   ' X8 : format aaaammjj
lAmj8_tiret = Mid$(X8, 1, 4) & "-" & Mid$(X8, 5, 2) & "-" & Mid$(X8, 7, 2)

End Sub

Public Function Text_KeyWord(lText As String, lK As Integer, blnSelectAll As Boolean) As String
Dim Kmin As Integer, kMax As Integer, lenText As Integer, xKeyWord As String, blnOk As Boolean
Dim X1 As String, blnKeyWord As Boolean

lenText = Len(lText)
blnKeyWord = False
Do
    Kmin = lK + 1
    xKeyWord = ""
    blnOk = False
    For kMax = Kmin To lenText
        X1 = Mid$(lText, kMax, 1)
        Select Case X1
            Case ".", "-", "_":
            Case "a" To "z": xKeyWord = xKeyWord & X1: blnOk = True
            Case "0" To "9": xKeyWord = xKeyWord & X1: blnOk = True
            Case Else: If blnOk Then Exit For
        End Select
                
    Next kMax
    
    If kMax >= lenText Then blnKeyWord = True
    lK = kMax
    
    If blnSelectAll Then
        blnKeyWord = True
    Else
    
        Select Case xKeyWord
            Case "l", "le", "la", "les", "du", "de", "des", "a", "au", "et", "ou":
            Case Else:
                        If Len(xKeyWord) > 1 Then
                            blnKeyWord = True
                        Else
                            xKeyWord = ""
                        End If
                        
        End Select
    End If
Loop Until blnKeyWord

Text_KeyWord = xKeyWord
End Function
Public Function Text_LCase(lText As String) As String
Dim X As String, I As Integer

X = LCase(Trim(lText))

' Voir aussi Text_Accent

For I = 1 To Len(X)
    Select Case Mid$(X, I, 1)
        Case "à", "â", "ä": Mid$(X, I, 1) = "a"
        Case "é", "è", "ê", "ë": Mid$(X, I, 1) = "e"
        Case "î", "ï": Mid$(X, I, 1) = "i"
        Case "ô", "ö": Mid$(X, I, 1) = "o"
        Case "ù", "û", "ü": Mid$(X, I, 1) = "u"
        Case "ç": Mid$(X, I, 1) = "c"

   End Select
Next I
Text_LCase = X
End Function

Public Function Text_Apostrophe(lText As String) As String
Dim X As String, I As Integer, K As Integer
X = lText
K = 1
Do
    I = InStr(K, X, "'")
    If I > 0 Then Mid$(X, I, 1) = " ": K = I
Loop While I > 0
Text_Apostrophe = X
End Function

Public Sub Text_Accent(lText As String)
Dim X As String, I As Integer

X = LCase(Trim(lText))

' Voir aussi Text_LCase

For I = 1 To Len(X)
    Select Case Mid$(X, I, 1)
        Case "à", "â", "ä": Mid$(lText, I, 1) = "a"
        Case "é", "è", "ê", "ë": Mid$(lText, I, 1) = "e"
        Case "î", "ï": Mid$(lText, I, 1) = "i"
        Case "ô", "ö": Mid$(lText, I, 1) = "o"
        Case "ù", "û", "ü": Mid$(lText, I, 1) = "u"
        Case "ç": Mid$(lText, I, 1) = "c"

   End Select
Next I
End Sub

Public Sub TEG_Calc_VersionMensuelleàSUPxxx(lCapital As Currency, lFrais As Currency, lMensualité As Currency, lPériodeNB As Integer, lTauxAnnuel As Double, LTEG As Double)
Dim wTEG As Double, wDiff As Double
Dim wTegMin As Double, wDiffMin As Double, blnTEGMin As Boolean
Dim wTegMax As Double, wDiffMax As Double, blnTEGMax As Boolean
Dim wT As Double
Dim blnOk As Boolean

Dim lAssurance As Currency
lAssurance = 0

blnOk = False
wTegMin = 0: wDiffMin = -lCapital: blnTEGMin = True
wTegMax = 0: wDiffMax = lCapital: blnTEGMax = True

wTEG = lTauxAnnuel / 1200
Do
    wT = (1 + wTEG) ^ lPériodeNB
    wDiff = lCapital - ((lMensualité + lAssurance) * (wT - 1) / (wTEG * wT) + lFrais)
    
    If Abs(wDiff) < 0.01 Then blnOk = True
    
    If wDiff < 0 Then
        If wDiff > wDiffMin Then
            wTegMin = wTEG: wDiffMin = wDiff: blnTEGMin = True
            If wTegMax = 0 Then
                wTEG = wTEG + 0.01
            Else
                wTEG = (wTegMin + wTegMax) / 2
            End If
        Else
            blnTEGMin = False
        End If
     Else
        If wDiff < wDiffMax Then
            wTegMax = wTEG: wDiffMax = wDiff: blnTEGMax = True
            wTEG = (wTegMin + wTegMax) / 2
        Else
            blnTEGMax = False
        End If
     End If
     
    If Not blnTEGMin And Not blnTEGMax Then blnOk = True
    
Loop Until blnOk

LTEG = wTEG * 1200

End Sub

Public Sub TEG_Calc(lCapital As Currency, lFrais As Currency, lMensualité As Currency, lPériodeNB As Integer, lPériodicité As String, lTauxAnnuel As Double, LTEG As Double)
Dim wTEG As Double, wDiff As Double
Dim wTegMin As Double, wDiffMin As Double, blnTEGMin As Boolean
Dim wTegMax As Double, wDiffMax As Double, blnTEGMax As Boolean
Dim wT As Double
Dim nbIter As Integer
Dim blnOk As Boolean
Dim nbPériodeDansAnnéeCivile  As Integer
Dim nbMoisDansPériode  As Integer

Dim lAssurance As Currency
LTEG = 0

lAssurance = 0
nbIter = 0
blnOk = False
wTegMin = 0: wDiffMin = -lCapital: blnTEGMin = True
wTegMax = 0: wDiffMax = lCapital: blnTEGMax = True

Select Case lPériodicité
    Case "M": nbPériodeDansAnnéeCivile = 12: nbMoisDansPériode = 1
    Case "T": nbPériodeDansAnnéeCivile = 4: nbMoisDansPériode = 3
    Case "S": nbPériodeDansAnnéeCivile = 2: nbMoisDansPériode = 6
    Case "A": nbPériodeDansAnnéeCivile = 1: nbMoisDansPériode = 12
    Case Else: Exit Sub
End Select
wTEG = lTauxAnnuel / 1200 ''''(nbPériodeDansAnnéeCivile * 100)
Do
    wT = (1 + wTEG) ^ (lPériodeNB)  '''''* nbMoisDansPériode)
    wDiff = lCapital - ((lMensualité + lAssurance) * (wT - 1) / (wTEG * wT) + lFrais)
    
    If Abs(wDiff) < 0.01 Then blnOk = True
    
    If wDiff < 0 Then
        If wDiff > wDiffMin Then
            wTegMin = wTEG: wDiffMin = wDiff: blnTEGMin = True
            If wTegMax = 0 Then
                wTEG = wTEG + 0.01
            Else
                wTEG = (wTegMin + wTegMax) / 2
            End If
        Else
            blnTEGMin = False
        End If
     Else
        If wDiff < wDiffMax Then
            wTegMax = wTEG: wDiffMax = wDiff: blnTEGMax = True
            wTEG = (wTegMin + wTegMax) / 2
        Else
            blnTEGMax = False
        End If
     End If
     '
    If Not blnTEGMin And Not blnTEGMax Then blnOk = True
    nbIter = nbIter + 1
    If nbIter > 100 Then Exit Sub
Loop Until blnOk

LTEG = wTEG * 100 * nbPériodeDansAnnéeCivile ''/ nbMoisDansPériode

End Sub


Public Function dateJma08_Amj08(ljma08 As String, lAMJ As String)
Dim Siecle2C As String

lAMJ = "00000000"
If Trim(ljma08) = "" Then Exit Function
If Mid$(ljma08, 7, 2) >= 90 Then
   Siecle2C = "19"
Else
   Siecle2C = "20"
End If
lAMJ = Siecle2C & Mid$(ljma08, 7, 2) & Mid$(ljma08, 4, 2) & Mid$(ljma08, 1, 2)

End Function


Public Function Time_Hms_Sss(lHMS As String) As Long
Time_Hms_Sss = Mid$(lHMS, 1, 2) * 3600 + Mid$(lHMS, 3, 2) * 60 + Mid$(lHMS, 5, 2)
End Function

Public Function Time_Sys_Sss() As Long
Dim X As String
X = Time
Time_Sys_Sss = Mid$(X, 1, 2) * 3600 + Mid$(X, 4, 2) * 60 + Mid$(X, 7, 2)

End Function


Public Function Time_Sss_Hms(lSss As Long) As String
Dim L1 As Long, L2 As Long, L3 As Long, lX As Long
L1 = Fix(lSss / 3600): lX = (lSss - L1 * 3600)
L2 = Fix(lX / 60): lX = lX - L2 * 60
L3 = lX Mod 60
Time_Sss_Hms = Format(L1, "00") & Format(L2, "00") & Format(L3, "00")
End Function

Public Sub XUsrId_Show()

usrId = UCase$(InputBox("User Id", "Changement d'identité", usrId))

XUsrId_Set

End Sub

Public Sub XUsrId_Set()


usrIdNT = usrId
usrName = usrIdNT
usrName_UCase = UCase(usrName)
usrCompte = "": usrRacine = ""

Elp.usrId = usrId
Xcom_UsrId usrId
elpSrvXcom = ""
mainSoc
elpSrvXcom = "XXXX"

End Sub

Function dateCtlDsys(ByVal X As String)
'---------------------------------------------------------------------

If Mid$(X, 11, 4) = "____" And Mid$(X, 6, 2) = "__" And Mid$(X, 1, 2) = "__" Then
    dateCtlDsys = DSys
Else
    dateCtlDsys = dateCtl(X)
End If

End Function



'---------------------------------------------------------
'-----------------------------------------------------
Sub mainEnd()
'-----------------------------------------------------
If DataBase_Open <> "" Then MDB_Close
Unload frmElp
End
End Sub


'---------------------------------------------------------------------
Function dateCtl(ByVal X As String)
'---------------------------------------------------------------------
Dim jma As String

Dim AMJ As typeDate

AMJ.AA = Format$(Val(Mid$(X, 11, 4)), "0000")

AMJ.MM = Format$(Val(Mid$(X, 6, 2)), "00")
AMJ.jj = Format$(Val(Mid$(X, 1, 2)), "00")


'If Amj.aa = "____" And Amj.mm = "__" And Amj.jj = "__" Then
If AMJ.AA = "0000" And AMJ.MM = "00" And AMJ.jj = "00" Then
    dateCtl = "00000000"
Else
    If AMJ.AA = "____" Or AMJ.AA = "0000" Then
        AMJ.AA = Mid$(DSys, 1, 4)
    Else
        If Val(AMJ.AA) < 70 Then
            Mid$(AMJ.AA, 1, 2) = "20"
        Else
            If Val(AMJ.AA) < 100 Then
                Mid$(AMJ.AA, 1, 2) = "19"
            End If
        End If
    End If
    
    If AMJ.MM = "__" Or AMJ.MM = "00" Then
        AMJ.MM = Mid$(DSys, 5, 2)
    End If
    If AMJ.jj = "__" Or AMJ.jj = "00" Then
        AMJ.jj = Mid$(DSys, 7, 2)
    End If
   
 '   jma = "# " & AMJ.jj & "-" & AMJ.mm & "-" & AMJ.aa & " #"
        jma = AMJ.jj & "-" & AMJ.MM & "-" & AMJ.AA
   If Val(AMJ.MM) > 12 Then
       dateCtl = "Mois"
       Exit Function
   End If
    If Not IsDate(jma) Then
        dateCtl = "Date erronée"
    Else
        If Val(AMJ.AA) < dateAAmin Then
            dateCtl = "Année < AAmin"
            
        Else
            If Val(AMJ.AA) > dateAAmax Then
                dateCtl = "Année > AAmax"
            Else
                dateCtl = AMJ.AA & AMJ.MM & AMJ.jj
            End If
        End If
    End If
End If
End Function

'---------------------------------------------------------
Function dateImp(ByVal X As String) As String
'---------------------------------------------------------

If X = "00000000" Or X = "0" Or RTrim(X) = "" Then
    dateImp = Space$(14)
Else
    dateImp = Format$(Mid$(X, 7, 2) & Mid$(X, 5, 2) & Mid$(X, 1, 4), "@@ - @@ - @@@@")
End If

End Function

'---------------------------------------------------------
Function dateImp10(ByVal X As String) As String
'---------------------------------------------------------

If X = "00000000" Or RTrim(X) = "" Then
    dateImp10 = Space$(10)
Else
    dateImp10 = Format$(Mid$(X, 7, 2) & Mid$(X, 5, 2) & Mid$(X, 1, 4), "@@.@@.@@@@")
End If

End Function
'---------------------------------------------------------
Function dateImp10_S(ByVal X As String) As String
'---------------------------------------------------------

If X = "00000000" Or RTrim(X) = "" Then
    dateImp10_S = Space$(10)
Else
    dateImp10_S = Format$(Mid$(X, 7, 2) & Mid$(X, 5, 2) & Mid$(X, 1, 4), "@@/@@/@@@@")
End If

End Function

'---------------------------------------------------------
Function dateJma6_Imp10(ByVal X As String) As String
'---------------------------------------------------------
If Trim(X) = "" Then
    dateJma6_Imp10 = ""
Else
    dateJma6_Imp10 = Format$(Mid$(X, 1, 2) & Mid$(X, 3, 2) & Mid$(X, 5, 2), "@@.@@.@@")
End If
End Function

'---------------------------------------------------------
Function dateAMJ6_Imp10(ByVal X As String) As String
'---------------------------------------------------------
If Trim(X) = "" Then
    dateAMJ6_Imp10 = ""
Else
    dateAMJ6_Imp10 = Format$(Mid$(X, 5, 2) & Mid$(X, 3, 2) & "20" & Mid$(X, 1, 2), "@@.@@.@@@@")
End If
End Function

'---------------------------------------------------------
Function dateAMJ10(ByVal X As String) As String
'---------------------------------------------------------

If X = "00000000" Or RTrim(X) = "" Then
    dateAMJ10 = Space$(10)
Else
    dateAMJ10 = Format$(Mid$(X, 1, 4) & Mid$(X, 5, 2) & Mid$(X, 7, 2), "@@@@.@@.@@")
End If

End Function

'---------------------------------------------------------
Function dateAMJ10_S(ByVal X As String) As String
'---------------------------------------------------------

If X = "00000000" Or RTrim(X) = "" Then
    dateAMJ10_S = Space$(10)
Else
    dateAMJ10_S = Format$(Mid$(X, 1, 4) & Mid$(X, 5, 2) & Mid$(X, 7, 2), "@@@@/@@/@@")
End If

End Function


'---------------------------------------------------------
Function dateAMJ10_T(ByVal X As String) As String
'---------------------------------------------------------

If X = "00000000" Or RTrim(X) = "" Then
    dateAMJ10_T = Space$(10)
Else
    dateAMJ10_T = Format$(Mid$(X, 1, 4) & Mid$(X, 5, 2) & Mid$(X, 7, 2), "@@@@-@@-@@")
End If

End Function

'---------------------------------------------------------
Function dateIBM10(ByVal X As String, blnJMA As Boolean) As String
'---------------------------------------------------------
Dim wDate As Long
wDate = Val(X)
If wDate = 0 Then
    dateIBM10 = ""
Else
    If blnJMA Then
        dateIBM10 = dateImp10(wDate + 19000000)
    Else
        dateIBM10 = dateAMJ10(wDate + 19000000)
    End If
End If
End Function
'---------------------------------------------------------
Function dateIBM_AMJ(ByVal X As String) As String
'---------------------------------------------------------
Dim wDate As Long
wDate = Val(X)
If wDate = 0 Then
    dateIBM_AMJ = ""
Else
    dateIBM_AMJ = X + 19000000
End If
End Function

'---------------------------------------------------------
Function dateIBM(ByVal X As String) As String
'---------------------------------------------------------

dateIBM = Format$((Val(X) - 19000000), "0000000")

End Function


'---------------------------------------------------------
Function dateImp_Amj(ByVal X As String) As String
'---------------------------------------------------------

If X = "00000000" Or RTrim(X) = "" Then
    dateImp_Amj = Space$(14)
Else
    dateImp_Amj = Format$(Mid$(X, 1, 4) & Mid$(X, 5, 2) & Mid$(X, 7, 2), "@@@@-@@-@@")
End If

End Function

'---------------------------------------------------------
Function dateImpS(ByVal X As String) As String
'---------------------------------------------------------

If X = "00000000" Or RTrim(X) = "" Then
    dateImpS = Space$(8)
Else
    dateImpS = Format$(Mid$(X, 7, 2) & Mid$(X, 5, 2) & Mid$(X, 3, 2), "@@-@@-@@")
End If

End Function



'---------------------------------------------------------
Function dateImp_jjMoisAAAA(ByVal X As String) As String
'---------------------------------------------------------
Dim I As Integer
I = Val(Mid$(X, 5, 2))
If I < 0 Or I > 12 Then
    dateImp_jjMoisAAAA = ""
Else
    dateImp_jjMoisAAAA = Format$(Mid$(X, 7, 2), "@@ ") & Trim(libMois(I)) & Format$(Mid$(X, 1, 4), " @@@@")
End If

End Function

'---------------------------------------------------------
Function dateImp_ddMonthYYYY(ByVal X As String) As String
'---------------------------------------------------------
Dim I As Integer
I = Val(Mid$(X, 5, 2))
If I < 0 Or I > 12 Then
    dateImp_ddMonthYYYY = ""
Else
    dateImp_ddMonthYYYY = Format$(Mid$(X, 7, 2), "@@ ") & Trim(libMonth(I)) & Format$(Mid$(X, 1, 4), " @@@@")
End If

End Function

'---------------------------------------------------------
Function dateElp(ByVal Fct As String, ByVal Nb As Integer, ByVal X As String) As String
'---------------------------------------------------------
Dim K As Integer, K1 As Integer, K2 As Integer, K3 As Integer, K4 As Integer
Dim V, X8 As String * 8, X8B As String * 8
Dim Fct_Mod As String, Nb_Mod As Integer
Dim jj As Integer, MM As Integer, AAAA As Integer
dateElp = X
Select Case Fct
        Case "TrimestreAdd": Fct_Mod = "MoisAdd": Nb_Mod = Nb * 3
        Case "SemestreAdd": Fct_Mod = "MoisAdd": Nb_Mod = Nb * 6
        Case "AnAdd": Fct_Mod = "MoisAdd": Nb_Mod = Nb * 12
        Case Else: Fct_Mod = Fct: Nb_Mod = Nb
End Select

Select Case Fct_Mod
   Case "Decade"
        Select Case Mid$(X, 7, 2)
            Case Is < 11: dateElp = Mid$(X, 1, 6) & "10"
            Case Is < 21: dateElp = Mid$(X, 1, 6) & "20"
            Case Else: dateElp = dateFinDeMois(X)
        End Select
    Case "Ouvré"
        V = dateImp(X)
        K = 0: K1 = IIf(Nb_Mod > 0, 1, -1)
        Do Until K = Nb_Mod
            V = DateAdd("d", K1, V)
            K2 = Weekday(V)
            If K2 > 1 And K2 < 7 Then K = K + K1
        Loop
        X8 = Year(V)
        Mid$(X8, 5, 2) = Format$(Month(V), "00")
        Mid$(X8, 7, 2) = Format$(Day(V), "00")
        dateElp = X8
        
    Case "Jour"
        V = Format$(Mid$(X, 7, 2) & Mid$(X, 5, 2) & Mid$(X, 1, 4), "@@ - @@ - @@@@")
        K = 0: K1 = IIf(Nb_Mod > 0, 1, -1)
        Do Until K = Nb_Mod
            V = DateAdd("d", K1, V)
            K = K + K1
        Loop
        X8 = Year(V)
        Mid$(X8, 5, 2) = Format$(Month(V), "00")
        Mid$(X8, 7, 2) = Format$(Day(V), "00")
        dateElp = X8
    Case "FinDeMoisP"
        X8 = X
        K = Val(Mid$(X8, 5, 2))
        If K > 1 Then
            K = K - 1
            Mid$(X8, 5, 2) = Format$(K, "00")
        Else
            K = Val(Mid$(X8, 1, 4)) - 1
            Mid$(X8, 1, 4) = Format$(K, "0000")
            Mid$(X8, 5, 2) = "12"
        End If
        dateElp = dateFinDeMois(X8)
    Case "FinDAnnéeP"
        X8 = X
        K = Val(Mid$(X8, 1, 4)) - 1
        Mid$(X8, 1, 4) = Format$(K, "0000")
        Mid$(X8, 5, 4) = "1231"
        dateElp = dateFinDeMois(X8)
    
    Case "MoisAdd"
        K1 = Fix(Abs(Nb_Mod) / 12)
        K2 = Abs(Nb_Mod) Mod 12
        X8 = X
        K = Val(Mid$(X8, 5, 2))
   
        If Nb_Mod < 0 Then
            K1 = -K1
            If K > K2 Then
                K = K - K2
                Mid$(X8, 5, 2) = Format$(K, "00")
            Else
                K1 = K1 - 1
                Mid$(X8, 5, 2) = Format$(Val(Mid$(X8, 5, 2)) - K2 + 12, "00")
            End If
        Else
            If K + K2 <= 12 Then
                K = K + K2
                Mid$(X8, 5, 2) = Format$(K, "00")
            Else
                K1 = K1 + 1
                Mid$(X8, 5, 2) = Format$(Val(Mid$(X8, 5, 2)) + K2 - 12, "00")
            End If
        End If
                 
        K = Val(Mid$(X8, 1, 4)) + K1
        Mid$(X8, 1, 4) = Format$(K, "0000")
        X8B = dateFinDeMois(X8)
        If X8 > X8B Then
            dateElp = X8B
        Else
            dateElp = X8
        End If
    Case "Weekday"                          'Nb : code du jour Dimanche = 1, signe - recherche le précédent, sinon le suivant
        V = dateImp(X)
        K = Abs(Nb): K1 = IIf(Nb_Mod > 0, 1, -1)
        Do
            V = DateAdd("d", K1, V)
            K2 = Weekday(V)
        Loop Until K2 = K
        X8 = Year(V)
        Mid$(X8, 5, 2) = Format$(Month(V), "00")
        Mid$(X8, 7, 2) = Format$(Day(V), "00")
        dateElp = X8
    Case "A-FM"                          'ajout x ans => fin de mois
        X8 = X
        K = Val(Mid$(X8, 1, 4)) + Nb
        Mid$(X8, 1, 4) = Format$(K, "0000")
        dateElp = dateFinDeMois(X8)
    Case "M-FM"                          'ajout x mois => fin de mois
        'X8 = X
        'K2 = Val(Mid$(X8, 5, 2)) + Nb
        'K1 = Fix((K2 - 1) / 12)
        'K = Val(Mid$(X8, 1, 4)) + K1
        'Mid$(X8, 1, 4) = Format$(K, "0000")
        'Mid$(X8, 5, 2) = Format$(K2 - K1 * 12, "00")
        'dateElp = dateFinDeMois(X8)
'$JPL 2013-02-01 _________________________________________________________________________
        Dim xDate As String
        xDate = dateImp10_S(X)
        xDate = DateAdd("m", -3, xDate)
        dateElp = dateFinDeMois(Mid$(xDate, 7, 4) & Mid$(xDate, 4, 2) & Mid$(xDate, 1, 2))
    Case "DateMS"
        dateElp = ""
        If Mid$(X, 1, 4) > "1900" And Mid$(X, 1, 4) < "2100" Then
            If Mid$(X, 5, 2) > "00" And Mid$(X, 5, 2) < "13" Then
                If Mid$(X, 7, 2) > "00" Then
                    If Mid$(X, 7, 2) < "29" Then
                        dateElp = Format$(Mid$(X, 1, 4) & Mid$(X, 5, 2) & Mid$(X, 7, 2), "@@@@/@@/@@")
                    Else
                        X8 = dateFinDeMois(X)
                        If Mid$(X, 7, 2) <= Mid$(X8, 7, 2) Then
                            dateElp = Format$(Mid$(X, 1, 4) & Mid$(X, 5, 2) & Mid$(X, 7, 2), "@@@@/@@/@@")
                        End If
                    End If
                End If
            End If
        End If
    Case "DateJJ-MM-AAAA"
        dateElp = ""
        K2 = InStr(1, X, "-")
        If K2 > 0 Then
            jj = Val(Trim(Mid$(X, 1, K2 - 1)))
            K3 = InStr(K2 + 1, X, "-")
            If K3 > 0 Then
                MM = Val(Trim(Mid$(X, K2 + 1, K3 - K2 - 1)))
            K4 = Len(X)
            If K4 > K3 Then
                AAAA = Val(Trim(Mid$(X, K3 + 1, K4 - K3)))
                If AAAA > 1900 And AAAA < 2100 Then
                    If MM > 0 And MM < 13 Then
                        If jj > 0 Then
                            If jj < 29 Then
                                dateElp = Format$(AAAA, "0000") & Format$(MM, "00") & Format$(jj, "00")
                            Else
                                X8 = dateFinDeMois(X)
                                If jj <= Mid$(X8, 7, 2) Then
                                    dateElp = Format$(AAAA, "0000") & Format$(MM, "00") & Format$(jj, "00")
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End Select

End Function

Public Function DateAdd_AMJ(lFct As String, lNb As Long, lGECH As String) As String
Dim V As Variant, X8 As String * 8

V = dateImp(lGECH)
V = DateAdd(lFct, lNb, V)
X8 = Year(V)
Mid$(X8, 5, 2) = Format$(Month(V), "00")
Mid$(X8, 7, 2) = Format$(Day(V), "00")
DateAdd_AMJ = X8
End Function

'---------------------------------------------------------
Function DateElp_X(ByVal Msg As String, ByVal xAMJ As String) As String
'---------------------------------------------------------
Dim K As Integer, Nb As Integer
Dim Fct As String

K = InStr(Msg, " ")
Nb = Val(Mid$(Msg, 1, K - 1))
Fct = Mid$(Msg, K + 1, Len(Msg) - K + 1)
DateElp_X = dateElp(Fct, Nb, xAMJ)
End Function



'---------------------------------------------------------
Function dateTime(ByVal D As String, ByVal T As String)
'---------------------------------------------------------

dateTime = "#" & Mid$(D, 7, 2) & "-" & Mid$(D, 5, 2) & "-" & Mid$(D, 1, 4) & " " & Mid$(T, 1, 2) & ":" & Mid$(T, 3, 2) & ":" & Mid$(T, 5, 2) & "#"
End Function

'---------------------------------------------------------
Sub elpErrMsg(C As Control, E As Control, Optional Dsp)
'---------------------------------------------------------

If IsMissing(Dsp) Then
    Set E.Container = C.Container
    
    E.Top = C.Top + C.Height + 10
    E.Left = C.Left
    'E.Width = TextWidth(E.Caption)
Else
    E.Top = Dsp.Top
    E.Left = Dsp.Left
    E.Width = Dsp.Width
    Dsp.Visible = False
End If

Beep
E.BackColor = focusUsr.BackColor
E.ForeColor = errUsr.ForeColor
E.Visible = True

If TypeOf C Is TextBox Then
    C.ForeColor = errUsr.ForeColor
    
    If oldText <> C.Text Then
        C.SetFocus
        C.SelStart = 0
            errTag = C.Tag
    End If
End If
End Sub

'---------------------------------------------------------
Public Function CErr(C As Control, E As Control)
'---------------------------------------------------------
E.Visible = True
E.FontBold = True
E.FontSize = 8
E.BackColor = focusUsr.BackColor
E.ForeColor = errUsr.ForeColor
C.ForeColor = errUsr.ForeColor
CErr = C.Tag

End Function





'---------------------------------------------------------------------
Function timeCtl(ByVal X As String)
'---------------------------------------------------------------------
If Mid$(X, 1, 2) = "__" And Mid$(X, 6, 2) = "__" Then
    timeCtl = "000000"
Else
    If Mid$(X, 1, 2) = "__" Then
        Mid$(X, 1, 2) = "00"
    End If

    If Mid$(X, 6, 2) = "__" Then
        Mid$(X, 6, 2) = "00"
    End If


    If Val(Mid$(X, 1, 2)) > 24 Then
        timeCtl = "Heure > 24"
    Else
        If Val(Mid$(X, 6, 2)) > 60 Then
            timeCtl = "Minute > 60"
        Else

            timeCtl = Mid$(X, 1, 2) & Mid$(X, 6, 2) & "00"
        End If
    End If
End If

End Function

'---------------------------------------------------------
Function timeImp(ByVal X As String) As String
'---------------------------------------------------------

If X = "000000" Then
    timeImp = Space$(13)
Else
    timeImp = Format$(Mid$(X, 1, 6), "@@ : @@ : @@")
End If


End Function

'---------------------------------------------------------
Function timeImpHM(ByVal X As String) As String
'---------------------------------------------------------

If X = "000000" Then
    timeImpHM = Space$(13)
Else
    timeImpHM = Format$(Mid$(X, 1, 4), "@@ \H @@")
End If


End Function
'---------------------------------------------------------
Function timeImp8(ByVal X As String) As String
'---------------------------------------------------------

If Trim(X) = "000000" Then 'Or "" Then
    timeImp8 = Space$(8)
Else
    timeImp8 = Format$(Mid$(X, 1, 6), "@@:@@:@@")
End If


End Function
'---------------------------------------------------------
Function timeNImp8(ByVal X As Long) As String
'---------------------------------------------------------

If X = 0 Then
    timeNImp8 = Space$(8)
Else
    timeNImp8 = Format$(X, "@@:@@:@@")
End If


End Function


'---------------------------------------------------------------------
Function timeMask(C As Control, E As Control)
'---------------------------------------------------------------------
Dim X

errTag = Null
X = timeCtl(C.Text)
timeMask = X

If IsNumeric(X) Then
    C.ForeColor = txtUsr.ForeColor
    E.Visible = False
Else
    E.Caption = X
    Call elpErrMsg(C, E)
End If
C.BackColor = txtUsr.BackColor

End Function




'---------------------------------------------------------
Public Function elpNum(C As Control, KeyAscii As Integer, maxI As Integer, maxD As Integer, Optional numFormat)
'---------------------------------------------------------
Dim X, F As String
Dim K, nbD, NbE As Integer
Dim Vmax

If KeyAscii = 13 Then KeyAscii = 0: Exit Function

If KeyAscii = 8 And Len(C.Text) > 0 Then
    X = Mid$(C.Text, 1, Len(C.Text) - 1)
Else

   If (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) _
   Or KeyAscii = Asc(".") And InStr(1, C.Text, ".") = 0 Then
        X = C.Text & Chr$(KeyAscii)
    Else
         Beep
         X = C.Text
    End If
End If

KeyAscii = 0
elpNum = Val(X)

Vmax = Val(String$(maxI, "9"))
If elpNum > Vmax Then
        X = Mid$(X, 1, Len(X) - 1)
        elpNum = Val(X)
    Beep
End If
    
K = InStr(1, X, ".")
If IsMissing(numFormat) Then
    F = "### ### ### ### ###"
Else
    F = numFormat
End If

If K > 0 Then
    nbD = Len(X) - K
    If nbD > maxD Then
        nbD = maxD
        X = Mid$(X, 1, Len(X) - 1)
        elpNum = Val(X)
        Beep
    End If
    F = F & "." & String$(nbD, "0")

End If

If F = "" Then
    C.Text = X
Else
    If elpNum = 0 Then
        C.Text = ""
    Else
        X = Format$(elpNum, F)
        C.Text = RTrim(X)
    End If
End If

C.SelStart = Len(C.Text)
C.SelText = ""

End Function

'---------------------------------------------------------
Public Sub MeInit(Nb As Integer, Optional Msg)
'---------------------------------------------------------
Dim kLocked As Boolean

kLocked = True
If Not IsMissing(Msg) Then
    If Msg = "Locked" Then kLocked = False
End If

XForm.Caption = XForm.Caption & "     ( " & socName & " : " & Trim(Elp.usrId) & " )"

XForm.KeyPreview = True
Nb = 0
errTag = Null
For Each xobj In XForm.Controls
    If TypeOf xobj Is TextBox Then
        xobj.Tag = Format$(Nb, "000"): Nb = Nb + 1 'xobj.Name
        xobj.Enabled = kLocked
        xobj.ForeColor = IIf(xobj.Enabled, vbBlue, vbBlack)
    End If
Next xobj

usrColor_Set

End Sub
'---------------------------------------------------------
Public Sub usrColor_Set()
'---------------------------------------------------------
For Each xobj In XForm.Controls
    If TypeOf xobj Is TextBox Then
        xobj.BackColor = txtUsr.BackColor
        xobj.ForeColor = txtUsr.ForeColor
    Else
        If TypeOf xobj Is Label Then '_
'        Or TypeOf xobj Is SSoption Then
    
            xobj.ForeColor = IIf(Mid$(xobj.Name, 1, 3) = "lib", libUsr.ForeColor, lblUsr.ForeColor)
    Else
        If TypeOf xobj Is PictureBox Then
            xobj.BackColor = picUsr.BackColor
            xobj.ForeColor = picUsr.ForeColor
      Else
        If TypeOf xobj Is ListBox Then
            If Trim(xobj.Name) = "lstErr" Then
                xobj.BackColor = errUsr.BackColor
                xobj.ForeColor = errUsr.ForeColor
            Else
                xobj.BackColor = lstUsr.BackColor
                xobj.ForeColor = lstUsr.ForeColor
            End If
      Else
        If TypeOf xobj Is Form Then
            xobj.BackColor = frmUsr.BackColor
            xobj.ForeColor = frmUsr.ForeColor
  Else
        If TypeOf xobj Is Shape Then
            xobj.BorderColor = vbBlue
        Else
        If TypeOf xobj Is SSTab _
        Or TypeOf xobj Is Frame _
        Or TypeOf xobj Is OptionButton _
        Or TypeOf xobj Is CheckBox _
        Then
            xobj.ForeColor = lblUsr.ForeColor
        End If
            End If
            End If
            End If
        End If
    End If
End If
Next xobj

End Sub


'---------------------------------------------------------
Public Sub Main()
'---------------------------------------------------------

Dim I As Integer
Dim X As String, X2 As String, X3 As String, X4 As String, wCommand As String
Dim I1 As Integer, I2 As Integer
Dim s() As String

On Error GoTo Error_Handler
App_Debug = "> ElpVb6_Main : Initialisation"
'-------------------------------------------------------------------------
App_EXEName = UCase$(Trim(App.exeName))
App_Title = UCase$(Trim(App.Title))
 
lX = 25: X = Space(25)
Vx = GetUserName(X, lX)
usrIdNT = Mid$(X, 1, lX - 1)

usrIdNT = "BENIA"




usrName = usrIdNT
usrName_UCase = UCase(usrName)
If Len(usrName_UCase) > 10 Then
    usrName_UCase10 = Mid$(usrName_UCase, 1, 10)
Else
    usrName_UCase10 = usrName_UCase
End If

usrName_ULCase = Mid$(usrName_UCase, 1, 1) & LCase(Mid$(usrName_UCase, 2, Len(usrName_UCase) - 1))

usrCompte = "": usrRacine = ""


libMois(1) = "Janvier": libMonth(1) = "Juanary"
libMois(2) = "Février": libMonth(2) = "February"
libMois(3) = "Mars": libMonth(3) = "March"
libMois(4) = "Avril": libMonth(4) = "April"
libMois(5) = "Mai": libMonth(5) = "May"
libMois(6) = "Juin": libMonth(6) = "June"
libMois(7) = "Juillet": libMonth(7) = "July"
libMois(8) = "Août": libMonth(8) = "August"
libMois(9) = "Septembre": libMonth(9) = "September"
libMois(10) = "Octobre": libMonth(10) = "October"
libMois(11) = "Novembre": libMonth(11) = "November"
libMois(12) = "Décembre": libMonth(12) = "December"

convArray

srvIdle = True

frmUsr_Windowstate = 0

warnUsrColor = vbMagenta
errUsr.BackColor = RGB(255, 230, 230)
errUsr.ForeColor = vbRed

focusUsr.BackColor = vbHighlight
focusUsr.ForeColor = vbHighlightText

frmUsr.BackColor = vbMenuBar
frmUsr.ForeColor = vbMenuText

lblUsr.BackColor = vbMenuBar
lblUsr.ForeColor = &H4000&     '&H808000    'vbMenuText

libUsr.BackColor = vbMenuBar
libUsr.ForeColor = vbBlue 'vbActiveTitleBar

lstUsr.BackColor = vbWindowBackground
lstUsr.ForeColor = vbBlue '
picUsr.BackColor = &HE0E0E0    'vbMenuBar 'vbWindowBackground
picUsr.ForeColor = vbBlue

txtUsr.BackColor = vbWindowBackground
txtUsr.ForeColor = vbBlue

dbUsr.ForeColor = vbRed
dbUsr.BackColor = RGB(255, 230, 230)
crUsr.ForeColor = vbBlue
crUsr.BackColor = RGB(230, 255, 255)

MouseMoveUsr.ForeColor = RGB(235, 255, 255) 'vbHighlight
MouseMoveUsr.BackColor = RGB(164, 255, 164) 'RGB(235, 255, 255)  'vbHighlight

greenColor.BackColor = RGB(230, 255, 230)
greenColor.ForeColor = RGB(0, 32, 0)

DSYS_Init

Asc01 = Chr$(1)
Asc03 = Chr$(3)
Asc10 = Chr$(10)
Asc10_13 = Chr$(10) & Chr$(13)
Asc13 = Chr$(13)
Asc34 = Chr$(34)
Asc39 = Chr$(39)
Asc232 = Chr$(Asc("é")) 'Chr$(232)      'Chr$(&HE8)
Asc233 = Chr$(Asc("è"))  'Chr$(233)      'Chr$(&HE9)
Asc123 = Chr$(Asc("{")) 'Chr$(123)      'Chr$(&H7B)     {
Asc125 = Chr$(Asc("=")) 'Chr$(125)      'Chr$(&H7D)     }

mColor_GB = RGB(0, 128, 128)
mColor_Z0 = RGB(255, 255, 255)
mColor_G0 = RGB(230, 255, 230)
mColor_G1 = RGB(210, 255, 210)
mColor_G2 = RGB(164, 255, 164)
mColor_G9 = RGB(0, 196, 128)
mColor_Y0 = RGB(255, 255, 228)
mColor_Y1 = RGB(255, 255, 196)
mColor_Y2 = RGB(255, 255, 164)
mColor_Y3 = RGB(255, 255, 128)
mColor_W0 = RGB(255, 224, 255)
mColor_W1 = RGB(255, 128, 255)
mColor_B0 = RGB(190, 240, 255)
mColor_B1 = RGB(128, 190, 255)
mColor_B9 = RGB(0, 128, 255)

htmlFontColor_Blue = "<Font color = #0000FF>"
htmlFontColor_Green = "<Font color = #106020>" '"<Font color = #008080>"
htmlFontColor_Gray = "<Font color = #505050>"
htmlFontColor_Red = "<Font color =#FF0000>"
htmlFontColor_Black = "<Font color =#000000>"
htmlFontColor_White = "<Font color =#FFFFFF>"
htmlFontColor_Magenta = "<Font color =#FF00FF>"

dateAAmin = 1980
dateAAmax = 2080
dateSerialMin = DateSerial(1989, 12, 31)
blnTimer_Enabled = False
blnNetSend_Enabled = False
elpSrvTxtin = False
elpSrvTxtOut = False
elpSrvXcom = ""
pcIdUsrIdCtl = True
DataBase_Open = "": DataBase_Master = "": DataBase_Local = ""
socName = "Société ?"

prtFontName = prtFontName_Arial
prtZoom = 0
prtSocSigle = False 'True
prtShow = True
prtCollection_Index = 0

'   DENIS               '
xlsManual = False
s = Split(Trim(Command), " ")
If UBound(s) > 0 Then
    mCommand = s(0)
    If UCase(s(1)) = "TRUE" Then
        xlsManual = True
    ElseIf nomDuServeur = "BIA2008" Then
        xlsManual = True
    End If
Else
    mCommand = Trim(Command)
End If
'                      '
wCommand = mCommand
blnAuto_Form_Show = True
App_Debug = "> ElpVb6_Main : analyse command"
'-------------------------------------------------------------------------
If UCase$(Mid$(mCommand, 1, 6)) = "@TIMER" Then
    If App.PrevInstance Then End
    If UCase$(Mid$(mCommand, 7, 1)) = "_" Then blnAuto_Form_Show = False             ' Affichage de la forme en mode AUTO
    paramElpTimer_Id = Mid$(mCommand, 8, Len(mCommand) - 7)
    wCommand = ""
    blnTimer_Enabled = True
End If

If wCommand = "" Then
     If App.PrevInstance Then Call MsgBox("Il y a déjà une instance active", vbCritical, App_EXEName): End
    mainSoc_Environnement
Else
    mainSoc_Environnement
End If

If blnOff_Line Then
    'usrIdNT = "XXXXX": usrName = usrIdNT  '$$$$$$$$$$$$$$$$$
    paramTemp_Folder = "C:\Temp"
Else
    paramTemp_Folder = "C:\Temp"
End If

App_Debug = "> ElpVb6_Main : frmELP.show"
'-------------------------------------------------------------------------

prtFontNameZ = prtFontName
Elp.SrvDTaqLen = "00000"
Elp.jplFree = "00000"
Elp.usrId = UCase$(usrName)
usrId = Trim(Elp.usrId)
frmElp.Show vbModeless 'vbModal
frmElp.Enabled = False
frmElp.SSTab1.Tab = 1

frmElp.imgSocSignon.Stretch = True  'False
frmElp.imgSocSignon.Picture = LoadPicture(strSocSignon)
Call FEU_VERT
Set msFileSystem = CreateObject("Scripting.FileSystemObject")
blnRéplication_Load = True

App_Debug = "> ElpVb6_Main : open " & DataBase_Local
'-------------------------------------------------------------------------
MDB_Open DataBase_Local, paramDataBase_Password
'20040830 jpl :DTAQ à supprimer $$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
If elpSrvXcom = "CAV4" Then elpSrvXcom = ""
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
App_Debug = "> ElpVb6_Main : mainsoc"
'-------------------------------------------------------------------------

frmElp.Timer1 = False
mainSoc

If Mid$(App_EXEName, 1, 4) = "AIB_" Then
'$JPL 2015-12-07        If Not IsNull(sqlYBIATAB0_Read("BIA_VB_AIB", usrIdNT, "", X)) Then
    If Not IsNull(sqlYBIATAB0_Read("BIA_VB_AIB", usrName_UCase, "", X)) Then
        MsgBox "Vous n'êtes pas habilité à ce programme", vbCritical, frmElp_Caption & App_Debug
        End
    Else
        blnBIA_VB_AIB = True
        usrId = Trim(X)
        XUsrId_Set
    End If
End If
    
App_Debug = "> ElpVb6_Main :frmElp"
'-------------------------------------------------------------------------
    frmElp.Caption = frmElp_Caption
    frmElp.Icon = frmElp_Icon
    'If Trim(frmElp_Icon) <> "" Then frmElp.Icon = LoadPicture(frmElp_Icon)

    Load frmElpPrt: MeInit I
    frmElpPrt.imgSocLogo.Picture = LoadPicture(paramSocLogo)
    frmElpPrt.imgSocLogo_G.Picture = LoadPicture(paramSocLogo_G)
    frmElpPrt.imgSocLogo_PiedPage.Picture = LoadPicture(paramSocLogo_PiedPage)
    frmElpPrt.imgFiligrane.Picture = LoadPicture("")
    frmElpPrt.imgFiligrane.Tag = ""
    frmElp.imgSocLogo.Picture = LoadPicture(paramSocLogo)

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
elpSrvXcom = "XXXX"

frmElpPrt.WinWord_Dir

If blnTimer_Enabled Then blnElpTimer_Auto = False: ElpTimer_Init
frmElp.Msg_Rcv "ELP"

frmElp.Enabled = True
frmElp.fgMain_App.Visible = Not blnTimer_Enabled
frmElp.fgMain_App_X.Visible = Not blnTimer_Enabled
If blnBIA_VB_AIB Then
    frmElp.fgMain_App_X.Visible = False
    
    frmElp.BackColor = vbRed
    frmElp.fra0.BackColor = mColor_Y0
End If

Exit Sub

Error_Handler:
    'Call Bia_swift_Monitor.EcritLog("Main", Err.Description, Err.source)
    MsgBox "Erreur :" & Err & " : " & Error$(Err), vbCritical, frmElp_Caption & App_Debug
    End
End Sub

Public Sub Xcom_UsrId(Msg As String)
Xcom.usrId = Msg
End Sub

'---------------------------------------------------------
Public Function RibClé(E As String, G As String, C As String, IbanE As String) As Integer
'---------------------------------------------------------
Dim r As Currency, mC As String
Dim X23 As String, X1 As String * 1
Dim I As Integer
C = UCase$(Trim(C))
mC = C
I = Len(C)
Do While Not IsNumeric(C)
    If I = 0 Then Exit Do
    
    X1 = Mid$(C, I, 1)
    If Not IsNumeric(X1) Then
       Select Case X1
        Case Is = "A", "J"
            Mid$(C, I, 1) = "1"
        Case Is = "B", "K", "S"
            Mid$(C, I, 1) = "2"
        Case Is = "C", "L", "T"
            Mid$(C, I, 1) = "3"
        Case Is = "D", "M", "U"
            Mid$(C, I, 1) = "4"
        Case Is = "E", "N", "V"
            Mid$(C, I, 1) = "5"
        Case Is = "F", "O", "W"
            Mid$(C, I, 1) = "6"
        Case Is = "G", "P", "X"
            Mid$(C, I, 1) = "7"
        Case Is = "H", "Q", "Y"
            Mid$(C, I, 1) = "8"
        Case Is = "I", "R", "Z"
            Mid$(C, I, 1) = "9"
            
       End Select
       
    End If
    I = I - 1
Loop
IbanE = ""
If Not IsNumeric(C) Then
    RibClé = 99
Else
    If Len(C) > 11 Then
        RibClé = 99
        C = ""
        IbanE = ""
    Else
        X23 = Format$(Val(E), "00000") & Format$(Val(G), "00000") _
          & Format$(Val(C), "00000000000") & Format$(0, "00")
        r = Mid$(X23, 1, 9) Mod 97
        r = (Format$(r, "00") & Mid$(X23, 10, 7)) Mod 97
        RibClé = 97 - (Format$(r, "00") & Mid$(X23, 17, 7)) Mod 97
    
        'Call Iban_Calc("FR00" & Mid$(X23, 1, 21) & Format$(RibClé, "00"), IbanE)
        Dim X As String
        X = "00000000000" & mC
        Call Iban_Calc("FR00" & Mid$(X23, 1, 10) & Mid$(X, Len(X) - 10, 11) & Format$(RibClé, "00"), IbanE)
    End If
End If

End Function

'---------------------------------------------------------
Public Function Rib_Compte(X As String) As String
'---------------------------------------------------------
Dim X1 As String * 1, y As String
Dim I As Integer, K As Integer, lenX As Integer

y = "00000000000"
X = Trim(UCase$(X))
lenX = Len(X)
K = 11
For I = lenX To 1 Step -1
    X1 = Mid$(X, I, 1)
    Select Case Asc(X1)
        Case 48 To 57, 65 To 90
            If K > 0 Then Mid$(y, K, 1) = X1: K = K - 1
    End Select
Next I
Rib_Compte = y
End Function

'---------------------------------------------------------
Public Sub meEnabled(ByVal Msg As Boolean)
'---------------------------------------------------------

For Each xobj In XForm.Controls
    If TypeOf xobj Is TextBox Then
        xobj.Enabled = Msg
    End If
Next xobj

End Sub

'---------------------------------------------------------
Public Function convUCase(KeyAscii As Integer) As Integer
'---------------------------------------------------------

convUCase = Asc(UCase(Chr(KeyAscii)))

End Function

'---------------------------------------------------------
Public Function ctlNum(KeyAscii As Integer) As Integer
'---------------------------------------------------------
   
If (KeyAscii >= 48 And KeyAscii <= 57) Then
    ctlNum = KeyAscii
Else
    If KeyAscii <> 8 And KeyAscii <> 32 And KeyAscii <> 46 Then
        Beep
        ctlNum = 0
    Else
        ctlNum = KeyAscii
    End If
End If
End Function



'---------------------------------------------------------
Public Function ctlNumS(KeyAscii As Integer) As Integer
'---------------------------------------------------------
   
If (KeyAscii >= 48 And KeyAscii <= 57) Then
    ctlNumS = KeyAscii
Else
    If KeyAscii <> 8 And KeyAscii <> 32 And KeyAscii <> 46 And KeyAscii <> 45 Then
        Beep
        ctlNumS = 0
    End If
End If
End Function

'---------------------------------------------------------
Public Sub lstErr_AddItem(E As Control, C As Control, ByVal X As String)
'---------------------------------------------------------

E.Visible = True
If Not blnBeep And Mid$(X, 1, 1) = "?" Then Beep: blnBeep = True
'E.Visible = True
E.AddItem Time & " " & X
If E.ListCount < 2 Then E.Height = (E.ListCount + 1) * 200 'lstLineHeight
E.TopIndex = E.ListCount - 1
'E.BackColor = errUsr.BackColor
'E.ForeColor = errUsr.ForeColor
    
If TypeOf C Is TextBox Then
    C.ForeColor = errUsr.ForeColor
    C.BackColor = errUsr.BackColor
    E.Tag = C.Tag
End If

DoEvents

End Sub
'---------------------------------------------------------
Public Sub lstErr_ChangeLastItem(E As Control, C As Control, ByVal X As String)
'---------------------------------------------------------

If E.ListCount > 0 Then E.RemoveItem E.ListCount - 1

Call lstErr_AddItem(E, C, X)
End Sub

'---------------------------------------------------------
Public Sub lstErr_Clear(E As Control, C As Control, ByVal X As String)
'---------------------------------------------------------
blnBeep = False
E.Clear
E.Visible = True
E.Height = 200
Call lstErr_AddItem(E, C, X)
'E.BackColor = frmUsr.BackColor
'E.ForeColor = frmUsr.ForeColor
If TypeOf C Is TextBox Then
    C.BackColor = focusUsr.BackColor
    C.ForeColor = focusUsr.ForeColor
End If
End Sub

'---------------------------------------------------------
Public Sub tag_SetFocus(XTag As String)
'---------------------------------------------------------

For Each xobj In XForm.Controls
    If TypeOf xobj Is TextBox Then
        If XTag = xobj.Tag Then
            If xobj.Enabled Then xobj.SetFocus
            Exit For
        End If
    End If
Next xobj

End Sub

'---------------------------------------------------------
Public Sub usrColor_Container(lObj As Control, lColor As Long)
'---------------------------------------------------------
For Each xobj In XForm.Controls
     If TypeOf xobj Is Label Or TypeOf xobj Is CheckBox Or TypeOf xobj Is OptionButton Then
        If xobj.Container.Name = lObj.Name Then
            xobj.BackColor = lColor
        End If
    End If
Next xobj

End Sub

Public Sub usrColor_Opt(oldC As Control, newC As Control)
oldC.ForeColor = lblUsr.ForeColor
newC.ForeColor = libUsr.ForeColor
Set oldC = newC
End Sub

Public Static Function time_Hms() As String
Dim X As String
X = Time
time_Hms = Mid$(X, 1, 2) & Mid$(X, 4, 2) & Mid$(X, 7, 2)

End Function

Public Sub num_KeyAscii(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) Then
    If KeyAscii <> 8 And KeyAscii <> 32 Then KeyAscii = 0
End If

End Sub

Public Sub num_KeyAsciiS(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) Then
    If KeyAscii <> 8 And KeyAscii <> 32 And KeyAscii <> 45 Then KeyAscii = 0
End If

End Sub

Public Function num_Control(C As TextBox, valX As Variant, ByVal maxI As Integer, ByVal maxD As Integer) As String
'Public Function num_Control(C As TextBox, valX As String, ByVal maxI As Integer, ByVal maxD As Integer) As String
Dim Vdec, Vmax
Dim X As String
num_Control = ""
If Trim(C.Text) = "" Then
    Vdec = 0
Else
    Vdec = num_CDec(C.Text)
End If

If maxD = 0 Then
    Vmax = Val(String$(maxI, "9"))
    If Vdec > Vmax Then num_Control = "> " & Vmax: Exit Function
Else
    Vmax = Val(String$(maxI, "9") & "." & String$(maxD, "9"))
    If Vdec > Vmax Then num_Control = "> " & Vmax: Exit Function
End If
valX = num_String(C.Text, maxI, maxD)
C.Text = num_Display(valX, maxI, maxD, C.ForeColor, X, "#")

End Function

Public Sub num_KeyAsciiD(KeyAscii As Integer, X As String)
If Trim(X) = "" And KeyAscii = 45 Then Exit Sub
If (KeyAscii < 48 Or KeyAscii > 57) Then
    If KeyAscii <> 8 And KeyAscii <> 32 Then
        If KeyAscii = 44 Then KeyAscii = 46
        If KeyAscii <> 46 Then
        
            KeyAscii = 0: Beep
        Else
            If InStr(1, X, ".") > 0 Then KeyAscii = 0: Beep
        End If
    End If
End If
End Sub

Public Function num_String(ByVal X As String, ByVal maxI As Integer, ByVal maxD As Integer) As String
Dim Vdec
Vdec = num_CDec(X)
If maxD = 0 Then
    num_String = Format$(Vdec, String$(maxI, "0"))
Else
    num_String = Format$(Vdec, String$(maxI, "0") & "." & String$(maxD, "0"))
End If
End Function
Public Function cur_19P(ByVal lCur As Currency) As String
Dim X As String

X = Format$(Abs(lCur), "0000000000000000.00")
Mid$(X, 17, 1) = "."
If lCur <> 0 Then Mid$(X, 1, 1) = IIf(lCur < 0, "-", "+")
cur_19P = X
End Function

Public Function cur_P(ByVal lCur As Currency) As String
Dim X As String, K As Integer

X = Trim(Format$(Abs(lCur), "###############0.00"))
K = InStr(1, X, ",")
If K > 0 Then Mid$(X, K, 1) = "."
If lCur < 0 Then X = "-" & X
cur_P = X
End Function


Public Function Comma_Point(ByVal lV As Variant) As String
Dim X As String, K As Integer

X = Trim(Abs(lV))
K = InStr(1, X, ",")
If K > 0 Then Mid$(X, K, 1) = "."
If lV < 0 Then X = "-" & X

Comma_Point = X
End Function

Public Function cur_19V(ByVal lCur As Currency) As String
Dim X As String

X = Format$(Abs(lCur), "0000000000000000.00")
Mid$(X, 17, 1) = ","
If lCur <> 0 Then Mid$(X, 1, 1) = IIf(lCur < 0, "-", "+")
cur_19V = X
End Function

Public Function cur_AbsV(ByVal lCur As Currency) As String
Dim X As String, K As Integer

X = Format$(Abs(lCur), "0.00")
K = InStr(1, X, ".")

If K > 0 Then Mid$(X, K, 1) = ","
cur_AbsV = X

End Function

Public Function cur_AbsV_Dev(ByVal lCur As Currency, lDev As String) As String
Dim X As String, K As Integer

If lDev = "JPY" Then
    X = Format$(Abs(lCur), "0.")
Else
    X = Format$(Abs(lCur), "0.00")
End If
K = InStr(1, X, ".")

If K > 0 Then Mid$(X, K, 1) = ","
cur_AbsV_Dev = X
End Function

Public Function num_Display(ByVal V As Variant, ByVal maxI As Integer, ByVal maxD As Integer, ForeColor As Long, Sens As String, ByVal strD As String) As String
Dim F As String, X1 As String * 1, Vdec
X1 = IIf(strD = "0", "0", "#")
Vdec = num_CDec(V)
If Vdec = 0 Then
    num_Display = ""
    ForeColor = crUsr.ForeColor
    Sens = ""
    If maxD > 0 And X1 = "0" Then num_Display = "." & String$(maxD, X1)
    Exit Function
End If
Select Case maxI
    Case 1: F = "#"
    Case 2: F = "##"
    Case 3: F = "###"
    Case 4: F = "# ###"
    Case 5: F = "## ###"
    Case 6: F = "### ###"
    Case 7: F = "# ### ###"
    Case 8: F = "## ### ###"
    Case 9: F = "### ### ###"
    Case 10: F = "# ### ### ###"
    Case 11: F = "## ### ### ###"
    Case 12: F = "### ### ### ###"
    Case 13: F = "# ### ### ### ###"
    Case 14: F = "## ### ### ### ###"
    Case 15: F = "### ### ### ### ###"
End Select
If maxD > 0 Then F = F & "." & String$(maxD, "0") ' X1)
num_Display = Format$(Vdec, F)
If Vdec < 0 Then
    ForeColor = dbUsr.ForeColor
    Sens = "DB"
Else
    ForeColor = crUsr.ForeColor
    Sens = "CR"
End If
End Function


Public Sub txtBox_Enabled(ByVal blnValue As Boolean)
For Each xobj In XForm.Controls
    If TypeOf xobj Is TextBox Then
        xobj.Enabled = blnValue
    End If
Next xobj
End Sub

Public Function Iban_Calc(IbanX As String, IbanE As String)
Dim r As Currency
Dim X As String, X1 As String * 1, X2 As String * 2, y As String
Dim I As Integer, K As Integer, lenX As Integer

Iban_Calc = Null
IbanE = "": y = ""
X = UCase$(Trim(IbanX))
lenX = Len(X)
If lenX < 11 Then Iban_Calc = "Iban : longueur < 5 caractères": Exit Function
Mid$(X, 3, 2) = "00"

For I = 1 To lenX
    X1 = Mid$(X, I, 1)
    K = Asc(X1)
    Select Case K
        Case 48 To 57: IbanE = IbanE & X1: y = y & X1
        Case 65 To 90: X2 = Format$(K - 55, "00"): IbanE = IbanE & X1: y = y & X2
    End Select
Next I
lenX = Len(y)
X = Mid$(y, 7, lenX - 6) & Mid$(y, 1, 6)
r = Mid$(X, 1, 9) Mod 97
For I = 10 To lenX Step 7
    r = (Format$(r, "00") & Mid$(X, I, 7)) Mod 97
Next I
Mid$(IbanE, 3, 2) = Format$(98 - r, "00")
End Function

Public Function Iban_Print(IbanE As String) As String
Iban_Print = Trim(Format$(IbanE, "!@@@@ @@@@ @@@@ @@@@ @@@@ @@@@ @@@@ @@@@"))
End Function

Public Function Iban_Check(IbanX As String)
Dim IbanE As String, V
Iban_Check = Null
V = Iban_Calc(IbanX, IbanE)
If Not IsNull(V) Then
    Iban_Check = V
Else
    If Mid$(IbanX, 3, 2) <> Mid$(IbanE, 3, 2) Then Iban_Check = "Clé Iban erronée : " & Mid$(IbanE, 3, 2)
End If
End Function

Public Function num_CDec(V As Variant) As Variant
Dim I As Integer
I = InStr(1, V, ",")
If I > 0 Then Mid$(V, I, 1) = "."
num_CDec = Val(V)

End Function

Public Sub num_Montant(KeyAscii As Integer, X As TextBox)
Dim xCur As Currency, blnFormat As Boolean
Dim K As Integer, wX As String
blnFormat = False
wX = Trim(Replace(X, ",", "."))
Select Case KeyAscii
    Case 48 To 57: xCur = Val(wX & Chr$(KeyAscii)): KeyAscii = 0: blnFormat = True
    Case 44, 46:
                If InStr(1, wX, ".") = 0 Then
                    X = Format$(Val(wX), "### ### ### ### ###.")
                End If
                KeyAscii = 0
                X.SelStart = Len(X)

    Case 8
    Case 45: If Trim(wX) <> "" Then KeyAscii = 0
    Case Else: KeyAscii = 0
End Select
If blnFormat Then
    K = InStr(1, wX, ".")
    If K > 0 Then
        Select Case Len(wX) - K
            Case 0: X = Format$(xCur, "### ### ### ### ###.0")
            Case 1: X = Format$(xCur, "### ### ### ### ###.00")
            Case Else
        End Select
    Else
        X = Format$(xCur, "### ### ### ### ###")
    End If
    X.SelStart = Len(X)
End If
Exit Sub

End Sub

Public Function num_CDec_USA(V As Variant) As Variant
Dim I As Integer
Do
    I = InStr(1, V, ",")
    If I > 0 Then Mid$(V, I, 1) = " "
Loop Until I = 0
num_CDec_USA = Val(V)

End Function

Public Sub cbo_Scan(X As String, cbo As ComboBox)
Dim I As Integer, lenX As Integer
lenX = Len(X)
cbo.ListIndex = -1
For I = 0 To cbo.ListCount - 1
    cbo.ListIndex = cbo.ListIndex + 1
    If X = Mid$(cbo.List(cbo.ListIndex), 1, lenX) Then Exit Sub
Next I
cbo.ListIndex = -1
End Sub

Public Sub fileListBox_Scan(X As String, fileListBox As fileListBox)
Dim I As Integer, lenX As Integer
lenX = Len(X)
fileListBox.ListIndex = -1
For I = 0 To fileListBox.ListCount - 1
    fileListBox.ListIndex = fileListBox.ListIndex + 1
    If X = Mid$(fileListBox.List(fileListBox.ListIndex), 1, lenX) Then Exit Sub
Next I
fileListBox.ListIndex = -1

End Sub

Public Sub lst_Scan(X As String, lst As ListBox)
Dim I As Integer, lenX As Integer
lenX = Len(X)
lst.ListIndex = -1
For I = 0 To lst.ListCount - 1
    lst.ListIndex = lst.ListIndex + 1
    If X = Mid$(lst.List(lst.ListIndex), 1, lenX) Then Exit Sub
Next I
lst.ListIndex = -1
End Sub

Public Sub lst_Scan_Text(lX As String, lst As ListBox)
Dim I As Integer, I0 As Integer, wX As String, mMultiSelect As Integer
mMultiSelect = lst.MultiSelect
'lst.MultiSelect = 0

I0 = lst.ListIndex + 1
For I = I0 To (lst.ListCount - 1)
    lst.ListIndex = I
    wX = Text_LCase(lst.List(lst.ListIndex))
    If InStr(1, wX, lX) > 0 Then
        GoTo Exit_sub
    End If
'Exit Sub

Next I
lst.ListIndex = -1

Exit_sub:
 'lst.MultiSelect = mMultiSelect
End Sub

Public Sub cbo_Value(X As String, cbo As ComboBox)
X = Mid$(cbo.List(cbo.ListIndex), 1, Len(X))
End Sub

Public Function dateFinDeMois(ByVal X As String) As String
Select Case Mid$(X, 5, 2)
    Case "02":
            dateFinDeMois = Mid$(X, 1, 6) & "28"
            If Val(Mid$(X, 1, 4)) Mod 4 = 0 Then
                If Val(Mid$(X, 1, 4)) Mod 100 <> 0 Then
                    dateFinDeMois = Mid$(X, 1, 6) & "29"
                Else
                    If Val(Mid$(X, 1, 4)) Mod 400 = 0 Then
                        dateFinDeMois = Mid$(X, 1, 6) & "29"
                    End If
                End If
            End If
    Case "04", "06", "09", "11": dateFinDeMois = Mid$(X, 1, 6) & "30"
    Case Else: dateFinDeMois = Mid$(X, 1, 6) & "31"
End Select

End Function

Public Sub Elp_ResizeImg(imgX As Image)
Dim D1 As Double, D2 As Double

D1 = 1: D2 = 1
If imgX.Width <> 0 Then D1 = imgX.Container.Width / imgX.Width
If imgX.Height <> 0 Then D2 = imgX.Container.Height / imgX.Height
If D2 < D1 Then D1 = D2
imgX.Width = imgX.Width * D1
imgX.Height = imgX.Height * D1
imgX.Stretch = True
imgX.Top = imgX.Container.Height - imgX.Height
imgX.Left = imgX.Container.Width - imgX.Width

End Sub

Public Sub Elp_Form_Resize(lMe As Form, lWindowState As Integer, lHeight_0 As Integer, lWidth_0 As Integer, lHeight_2 As Integer, lWidth_2 As Integer)
Dim D1 As Double, D2 As Double, D3 As Double, wFontSize As Integer
Dim DScreen As Double
On Error Resume Next

'If Screen.Width > 19200 Then Exit Sub

DScreen = Screen.Height / Screen.Width

If lMe.WindowState = vbMaximized Then
    frmUsr_Windowstate = vbMaximized
    wFontSize = 10
    lHeight_2 = lMe.Height - 300 * DScreen: lWidth_2 = lMe.Width - 300
    If lHeight_0 <> 0 Then
        D1 = lWidth_2 / lWidth_0
        D2 = lHeight_2 / lHeight_0
        D3 = (lHeight_2 / lHeight_0) * DScreen ' 0.5
        
    Else
        Exit Sub
    End If
Else
    frmUsr_Windowstate = vbNormal
    wFontSize = 8
    lHeight_0 = lMe.Height: lWidth_0 = lMe.Width
    If lHeight_2 <> 0 Then
        D1 = lWidth_0 / lWidth_2
        D2 = lHeight_0 / lHeight_2
        D3 = (lHeight_0 / lHeight_2) / DScreen '0.5
    Else
        Exit Sub
    End If
End If

For Each xobj In lMe.Controls
'Debug.Print xobj.Name
    If TypeOf xobj.Container Is Toolbar Then
    Else
        If TypeOf xobj Is Menu _
        Or TypeOf xobj Is CoolBar _
        Or TypeOf xobj Is Toolbar _
        Or TypeOf xobj Is Timer Then
        Else
            'If TypeOf xobj Is TextBox
            'Or TypeOf xobj Is Label _
            'Or TypeOf xobj Is ComboBox _
            'Or TypeOf xobj Is fileListBox _
            'Or TypeOf xobj Is ListBox _
            'Or TypeOf xobj Is PictureBox _
            'Or TypeOf xobj Is MSFlexGrid Then
            '    xobj.FontSize = wFontSize

            'End If
            'If xobj.Left < 0 Then Debug.Print xobj.Name; xobj.Left; xobj.Container
            On Error Resume Next
            If TypeOf xobj Is Line Then
                xobj.X1 = xobj.X1 * D1: xobj.X2 = xobj.X2 * D1
                xobj.Y1 = xobj.Y1 * D2: xobj.Y2 = xobj.Y2 * D2
            Else
                If xobj.Left > 0 Then xobj.Left = xobj.Left * D1
                xobj.Top = xobj.Top * D2
                xobj.Width = xobj.Width * D1
                If TypeOf xobj Is ComboBox Then
                Else
                    If TypeOf xobj Is RichTextBox Then
                    Else
                        xobj.Font.Size = xobj.Font.Size * D1
                    End If
                    
                    If TypeOf xobj Is TextBox Then 'And xobj.Height < 500 Then
                        xobj.Height = xobj.Height * D3
                    Else
                        xobj.Height = xobj.Height * D2
                        If TypeOf xobj Is MSFlexGrid Then
                            xobj.RowHeight = xobj.RowHeight * D2
                        End If

                    End If
                    'If TypeOf xobj Is IMAGE Then
                        'Set XImage = xobj
                        'Call Elp_ResizeImg(XImage)
                    'End If
                End If
            End If
        End If
    End If
Next xobj
lWindowState = lMe.WindowState
End Sub

Public Sub Elp_MouseMove()
'commandbutton.style=1

'Public Sub MouseMoveActiveControl_Set(C As Control)
'Dim MouseMoveActiveControl_Name  As String, MouseMoveActiveControl As typeUsrColor

'MouseMoveActiveControl_Reset
'MouseMoveActiveControl_Set $$
'If MouseMoveActiveControl_Name <> C.Name Then
'    MouseMoveActiveControl_Reset
'    If Not C.Enabled Then
'        MouseMoveActiveControl_Name = ""
'    Else
'        MouseMoveActiveControl_Name = C.Name
'        If TypeOf C Is CommandButton Then
'            MouseMoveActiveControl.BackColor = C.BackColor
'            C.BackColor = MouseMoveUsr.BackColor
'        Else
'            MouseMoveActiveControl.ForeColor = C.ForeColor
'            C.ForeColor = MouseMoveUsr.ForeColor
'        End If
'    End If
'End If

'End Sub


'Public Sub MouseMoveActiveControl_Reset()
'For Each xobj In Me.Controls
'    If MouseMoveActiveControl_Name = xobj.Name Then
'        MouseMoveActiveControl_Name = ""
'        If TypeOf xobj Is CommandButton Then
'            xobj.BackColor = MouseMoveActiveControl.BackColor
'        Else
'            xobj.ForeColor = MouseMoveActiveControl.ForeColor
'        End If
'        Exit For
'    End If
'Next xobj

'End Sub

End Sub

Public Sub Elp_ResizeControl(C As Control)
Dim H As Integer, Hmax As Integer

If TypeOf C Is ListBox Then
    Hmax = C.Container.Height - C.Top - 100
    H = C.ListCount * 200 + 200
    C.Height = IIf(Hmax < H, Hmax, H)
End If

End Sub

Public Sub filDoc_Pattern(filDoc As fileListBox, X As String)
Dim X1 As String, X2 As String
Dim I As Integer, L As Integer

L = Len(X)
For I = L To 1 Step -1
    If Mid$(X, I, 1) = "\" Then Exit For
Next I

X1 = "": X2 = ""
If I > 1 Then X1 = Mid$(X, 1, I)
If I < L Then X2 = Mid$(X, I + 1, L - I)
I = InStr(X2, ".")
If I > 0 Then
    X2 = X2 & "*"
Else
    X2 = X2 & "*.*"
End If

filDoc.path = X1
filDoc.Pattern = X2 & "*"

End Sub

Public Sub fileName_Split(lFileName As String, lFolder As String, lName As String, lExtension As String)
Dim K As Integer, L As Integer

lFolder = "": lName = "": lExtension = ""

L = Len(lFileName)
For K = L To 1 Step -1
    If Mid$(lFileName, K, 1) = "\" Then Exit For
Next K


If K > 1 Then lFolder = Mid$(lFileName, 1, K)
If K < L Then lName = Mid$(lFileName, K + 1, L - K)
K = InStr(lName, ".")
If K > 0 Then
    lExtension = Mid$(lName, K + 1, Len(lName) - K)
    lName = Mid$(lName, 1, K - 1)
End If


End Sub
Public Function fileName_Extension(lFileName As String) As String
Dim K As Integer, L As Integer

L = Len(lFileName)
For K = L To 1 Step -1
    If Mid$(lFileName, K, 1) = "." Then Exit For
Next K

If K > 0 Then
    fileName_Extension = Mid$(lFileName, K + 1, Len(lFileName) - K)
Else
    fileName_Extension = ""
End If


End Function

Public Function fileName_Change(lFileName As String, lOld As String, lNew As String) As String
Dim K As Integer, L As Integer, lX As Integer
Dim xOld As String, xNew As String
Dim X1 As String, X2 As String

fileName_Change = ""
xOld = Trim(lOld)
xNew = Trim(lNew)

K = InStr(lFileName, xOld)
If K > 0 Then
    L = Len(lFileName)
    lX = Len(xOld)
    X1 = "": X2 = ""
    If K > 1 Then X1 = Mid$(lFileName, 1, K - 1)
    K = K + lX
    If K < L Then X2 = Mid$(lFileName, K, L - K + 1)
    fileName_Change = X1 & xNew & X2
End If

End Function


Public Sub DTPicker_LostFocus(C As DTPicker)

C.CalendarBackColor = txtUsr.BackColor
C.CalendarForeColor = txtUsr.ForeColor

End Sub
Public Sub DTPicker_GotFocus(C As DTPicker)

C.CalendarBackColor = focusUsr.BackColor
C.CalendarForeColor = focusUsr.ForeColor

End Sub


Public Sub DTPicker_Set(C As DTPicker, AMJ As String)
If IsNumeric(AMJ) Then
    If AMJ >= 19000101 Then
        C.Day = "01"
        C.Year = Mid$(AMJ, 1, 4)
        C.Month = Mid$(AMJ, 5, 2)
        C.Day = Mid$(AMJ, 7, 2)
    End If
End If
End Sub
'-------------------------------------------------'
Public Function DTPicker_Control(C As DTPicker, xAMJ As String)
'-------------------------------------------------'

Dim X As String
DTPicker_Control = Null
X = Format$(C.Year, "0000") & Format$(C.Month, "00") & Format$(C.Day, "00")
If Not IsNumeric(X) Then
    xAMJ = "00000000"
    DTPicker_Control = "? erreur date"
    DTPicker_Now C
Else
    xAMJ = Mid$(X, 1, 8)
End If

End Function

Public Sub DTPicker_Now(C As DTPicker)
C.Year = Year(Now)
C.Day = 1
C.Month = Month(Now)
C.Day = Day(Now)

End Sub



Public Function Périodicité_Nbj(lPériodicité As String) As Integer
Select Case lPériodicité
    Case "M": Périodicité_Nbj = 30
    Case "T": Périodicité_Nbj = 90
    Case "S": Périodicité_Nbj = 180
    Case "A": Périodicité_Nbj = 360
    Case Else: Périodicité_Nbj = 0
End Select

End Function

Public Sub meEnabled_Container(ByVal lContainer As String, ByVal lTrueFalse As Boolean)
On Error Resume Next
For Each xobj In XForm.Controls

    If TypeOf xobj Is TextBox _
    Or TypeOf xobj Is DTPicker _
    Or TypeOf xobj Is ComboBox _
    Or TypeOf xobj Is ListBox _
    Or TypeOf xobj Is CheckBox _
    Or TypeOf xobj Is CheckBox _
    Or TypeOf xobj Is Frame _
    Or TypeOf xobj Is CommandButton _
    Or TypeOf xobj Is OptionButton Then
        If xobj.Container.Name = lContainer Then
            xobj.Enabled = lTrueFalse
        End If
    End If
Next xobj
End Sub

Public Sub lbl_Style(C As Label, lTF As Boolean)
If lTF Then
    C.ForeColor = warnUsrColor
    C.BackColor = focusUsr.BackColor
    C.BorderStyle = 1
Else
    C.ForeColor = lblUsr.ForeColor
    C.BackColor = lblUsr.BackColor
    C.BorderStyle = 0
End If

End Sub
Public Sub chk_Style(C As CheckBox, lTF As Boolean)
If lTF Then
    C.ForeColor = warnUsrColor
    C.BackColor = focusUsr.BackColor
'    C.Style = 1
Else
    C.ForeColor = lblUsr.ForeColor
    C.BackColor = lblUsr.BackColor
'    C.Style = 0
End If

End Sub


Public Sub cbo_Load(lId As String, lK1 As String, cbo As ComboBox, lK2_len As Integer)
Dim X As String
cbo.Clear
X = "select * from ElpTable where SNN = 0" _
    & " and id = '" & lId & "'" _
    & " and K1 = '" & lK1 & "'"
    
Set rsMDB = cnMDB.Execute(X)
Do While Not rsMDB.EOF
    cbo.AddItem Mid$(rsMDB("K2"), 1, lK2_len) & " " & Trim(rsMDB("Name"))
    rsMDB.MoveNext
Loop
End Sub

Public Sub cbo_LoadK2(lId As String, lK1 As String, cbo As ComboBox)
Dim X As String
cbo.Clear
X = "select * from ElpTable where SNN = 0" _
    & " and id = '" & lId & "'" _
    & " and K1 = '" & lK1 & "'"
    
Set rsMDB = cnMDB.Execute(X)
Do While Not rsMDB.EOF
    cbo.AddItem rsMDB("K2")
    rsMDB.MoveNext
Loop

End Sub
Public Sub cbo_LoadName(lId As String, lK1 As String, cbo As ComboBox)
Dim X As String
cbo.Clear
X = "select * from ElpTable where SNN = 0" _
    & " and id = '" & lId & "'" _
    & " and K1 = '" & lK1 & "'"
    
Set rsMDB = cnMDB.Execute(X)
Do While Not rsMDB.EOF
    cbo.AddItem rsMDB("Name")
    rsMDB.MoveNext
Loop

End Sub

Public Sub lst_LoadK2(lId As String, lK1 As String, lst As ListBox, blnDisplay_K2 As Boolean)
Dim X As String
X = "select * from ElpTable where SNN = 0" _
    & " and id = '" & lId & "'" _
    & " and K1 = '" & lK1 & "'"
    
Set rsMDB = cnMDB.Execute(X)
Do While Not rsMDB.EOF
    If blnDisplay_K2 Then
        X = Trim(rsMDB("K2")) & " " & Trim(rsMDB("Name"))
    Else
         X = Trim(rsMDB("Name"))
   End If
    
    lst.AddItem X
    rsMDB.MoveNext
Loop

End Sub

Public Sub lst_LoadK1(lId As String, lst As ListBox)
Dim X As String
X = "select * from ElpTable where SNN = 0" _
    & " and id = '" & lId & "'"
    
Set rsMDB = cnMDB.Execute(X)
Do While Not rsMDB.EOF
    lst.AddItem rsMDB("K1") & vbTab & Trim(rsMDB("Name"))
    rsMDB.MoveNext
Loop

End Sub

Public Sub cbo_LoadId(lId As String, cbo As ComboBox)
Dim X As String
cbo.Clear
X = "select * from ElpTable where SNN = 0" _
    & " and id = '" & lId & "'"
    
Set rsMDB = cnMDB.Execute(X)
Do While Not rsMDB.EOF
    cbo.AddItem rsMDB("K1")
    rsMDB.MoveNext
Loop

End Sub

Public Sub cbo_LoadId_K2(lId As String, lK2 As String, cbo As ComboBox)
Dim X As String
cbo.Clear
X = "select * from ElpTable where SNN = 0" _
    & " and id = '" & lId & "' and K2 = '" & lK2 & "'"
    
Set rsMDB = cnMDB.Execute(X)
Do While Not rsMDB.EOF
    cbo.AddItem rsMDB("K1")
    rsMDB.MoveNext
Loop

End Sub
Public Sub cbo_Load_Unit(cbo As ComboBox) ' 2007-01-29 bricolage changement nom de service
Dim X As String
cbo.Clear
X = "select * from ElpTable where SNN = 0" _
    & " and id = 'Unit' and K2 = '' and Name not like '=%'"
    
Set rsMDB = cnMDB.Execute(X)
Do While Not rsMDB.EOF
    cbo.AddItem rsMDB("K1")
    rsMDB.MoveNext
Loop

End Sub

Public Function dateJMA_AMJ(lJma As Variant, lAMJ As String)
lAMJ = ""
lAMJ = Mid$(lJma, 7, 4) & Mid$(lJma, 4, 2) & Mid$(lJma, 1, 2)

End Function

Public Function dateJMA6_AMJ(lJma As String, lAMJ As String)
lAMJ = ""
If Len(lJma) < 8 Then
    lAMJ = "20" & Mid$(lJma, 7, 2) & Mid$(lJma, 4, 2) & Mid$(lJma, 1, 2)
Else
    lAMJ = Mid$(lJma, 7, 4) & Mid$(lJma, 4, 2) & Mid$(lJma, 1, 2)
End If

End Function
Public Function dateX8_N8(lJma As String) As Long

dateX8_N8 = 20000000 + Val(Mid$(lJma, 7, 2)) * 10000 + Val(Mid$(lJma, 4, 2)) * 100 + Val(Mid$(lJma, 1, 2))

End Function

Public Function dateX6_N8(lJma As String) As Long

dateX6_N8 = 20000000 + Val(Mid$(lJma, 5, 2)) * 10000 + Val(Mid$(lJma, 3, 2)) * 100 + Val(Mid$(lJma, 1, 2))

End Function

Public Function dateAMJ8_JMA6(lAMJ As String, lJma As String)
lJma = Mid$(lAMJ, 7, 2) & Mid$(lAMJ, 5, 2) & Mid$(lAMJ, 3, 2)

End Function

Public Function dateAMJ_JMA(lAMJ As String, lJma As String)
lJma = ""
lJma = Mid$(lAMJ, 7, 2) & Mid$(lAMJ, 5, 2) & Mid$(lAMJ, 1, 4)

End Function

Public Function dateJma10_Amj(ljma10 As String, lAMJ As String)
lAMJ = "00000000"
lAMJ = Mid$(ljma10, 7, 4) & Mid$(ljma10, 4, 2) & Mid$(ljma10, 1, 2)

End Function

Public Function curMaxD(lcurX As Currency, lMaxD As String) As Currency

Select Case lMaxD
    Case 0: curMaxD = Fix(lcurX + 0.5000001)
    Case Else: curMaxD = Fix((lcurX + 0.00500001)) / 100
End Select

End Function

Public Function paramServer(lMsg) As String
Dim X As String, I1 As Integer, I2 As Integer
Dim V, xName As String, xMemo As String
On Error GoTo Error_Handler

App_Debug = "> paramServer : " & lMsg
'--------------------------------------------------------------------------------------


X = Trim(lMsg)
paramServer = X
If Mid$(lMsg, 1, 2) = "\\" Then
    I1 = InStr(3, X, "\")
    If I1 > 0 Then
        ''recparamServer.K2 = mId$(X, 3, I1 - 3)
        ''recparamServer.ID = "Server"
        ''recparamServer.K1 = "Application"

        V = rsElpTable_Read("Server", "Application", Mid$(X, 3, I1 - 3), xName, xMemo)
        ''If Not IsNull(V) Then GoTo Error_MsgBox

        If IsNull(V) Then
            If Not IsNull(xMemo) Then
                If blnOff_Line Then xMemo = paramTemp_Folder: I1 = 2
                I2 = Len(lMsg)
                If Mid$(xMemo, 2, 1) = ":" Then
                    paramServer = Trim(xMemo) & Mid$(X, I1, I2 - I1 + 1)
               Else
                    paramServer = "\\" & Trim(xMemo) & Mid$(X, I1, I2 - I1 + 1)
                End If
            End If
        End If
    End If
End If
Exit Function

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    If Not blnTimer_Enabled Then MsgBox V & " " & Now, vbCritical, frmElp_Caption & App_Debug
End Function


Public Sub MDB_Open(lDataBase_Name As String, lDataBase_PassWord As String)
If lDataBase_Name <> "" Then
    If DataBase_Open <> "" Then MDB_Close

    DataBase_Open = lDataBase_Name
'    Set MDB = OpenDatabase(paramDataBase_Password, False, False, lDataBase_PassWord)
'    tableElpTable_Open

'    MDB.Execute "delete * from ElpBuffer": tableElpBuffer_Open:


    Set cnMDB = New ADODB.Connection
    cnMDB.Provider = "Microsoft.Jet.OLEDB.4.0"
    cnMDB.Properties("JET OLEDB:Database Password") = lDataBase_PassWord
    cnMDB.Mode = adModeReadWrite
    
    cnMDB.Open lDataBase_Name
    
    If UCase$(lDataBase_Name) <> UCase$(cnMDB.Properties("Data Source Name")) Then
        MsgBox lDataBase_Name, vbCritical, " non conforme "
        cnAdo_Info cnMDB
        End
    End If
'cnAdo_Info cnMDB
End If

End Sub

Public Sub MDB_Master()
Dim X As String
If DataBase_Local <> DataBase_Master Then
    X = MsgBox("Choisissez-vous la base principale [" & DataBase_Master & "]?", vbYesNo + vbQuestion + vbDefaultButton1, "Choix de la base à mettre à jour: LOCALE ou PRINCIPALE")
    If X = vbYes Then
        MDB_Open DataBase_Master, paramDataBase_Password
    Else
        If DataBase_Open <> DataBase_Local Then MDB_Open DataBase_Local, paramDataBase_Password
    End If
End If
End Sub
Public Sub MDB_Local()
Dim X As String
If DataBase_Local <> DataBase_Master Then
    If DataBase_Open = DataBase_Master Then
        Call MsgBox("Réplication de la base principale vers la base locale", vbInformation, "Elp : MDB_Local")
        MDB_Close
        msFileSystem.CopyFile DataBase_Master, DataBase_Local
        MDB_Open DataBase_Local, paramDataBase_Password
    End If
End If

End Sub

Public Sub MDB_Replication()
Dim X As String

If DataBase_Local <> DataBase_Master Then
    X = MsgBox("Réplication de la base principale vers la base locale", vbInformation + vbYesNo + vbDefaultButton2, "Elp : MDB_Local")
    If X = vbYes Then
        X = DataBase_Open
        If X <> "" Then MDB_Close
        DataBase_Open = ""
        msFileSystem.CopyFile DataBase_Master, DataBase_Local
        If X <> "" Then MDB_Open X, paramDataBase_Password
    End If
Else
    Call MsgBox("Base principale = base locale", vbInformation, "Elp : MDB_Replication")
End If

End Sub

Public Sub MDB_ReplaceMaster()
On Error GoTo Error_Handler
Dim X As String

If DataBase_Local <> DataBase_Master Then
    X = MsgBox("REMPLACEMENT de la base principale PAR la base locale", vbInformation + vbYesNo + vbDefaultButton2, "Elp : MDB_Local")
    If X = vbYes Then
        X = DataBase_Open
        DataBase_Open = ""
        If X <> "" Then MDB_Close
        msFileSystem.CopyFile DataBase_Local, DataBase_Master
        If X <> "" Then MDB_Open X, paramDataBase_Password
    End If
Else
    Call MsgBox("Base principale = base locale", vbInformation, "Elp : MDB_ReplaceMaster")
End If
Exit Sub
Error_Handler:
Shell_MsgBox "ELPVB_MDB_ReplaceMaster : " & Error, vbCritical, frmElp_Caption, False
End Sub


Public Sub MDB_Close()
cnMDB.Close
Set cnMDB = Nothing

'mainSoc_Close
'MDB.Close
DataBase_Open = ""
End Sub

Public Sub main_Reset()
Dim X As String, IdShell
On Error GoTo Error_Exit
X = MsgBox("Voulez-vous réellement mettre à jour la base BIASRV ?", vbQuestion & vbYesNo, "XReset")
If X = vbNo Then Exit Sub

If Not msFileSystem.FolderExists(paramFolder_Master) Then
    Call MsgBox("paramFolder_Master inconnu : " & paramFolder_Master, vbCritical, "ElpVb4 : main_Reset")
    Exit Sub
End If

If Not msFileSystem.FolderExists(paramFolder_Local) Then
    paramFolder_Local = ""
    Call MsgBox("paramFolder_Local inconnu : " & paramFolder_Local, vbCritical, "ElpVb4 : main_Reset")
    Exit Sub
End If

X = DataBase_Open
If X <> "" Then MDB_Close
'msFileSystem.CopyFolder paramFolder_Master, paramFolder_Local, True
'msFileSystem.CopyFile paramFolder_Master & "\*.*", paramFolder_Local, True
X = UCase$(frmElp_Caption & ".exe ") & mCommand
IdShell = Shell(paramFolder_Master & "\BIASrv_Copy.cmd " & X, 1)
'IdShell = Shell(paramFolder_Local & "\" & frmElp_Caption & ".exe", 1)
'AppActivate IdShell
'DoEvents

mainEnd

Exit Sub

Error_Exit:
    Call MsgBox("Erreur : " & Error, vbCritical, "ElpVb4 : main_Reset : " & paramFolder_Master & "\*.*")
    Exit Sub
    
End Sub

Public Function CSV_Scan(lText As String, lK As Integer) As String
Dim Kmin As Integer, kMax As Integer, lenText As Integer
Kmin = lK + 1
lenText = Len(lText)
CSV_Scan = ""
For kMax = Kmin To lenText
  If Mid$(lText, kMax, 1) = ";" Then Exit For
Next kMax

If kMax > Kmin Then CSV_Scan = Mid$(lText, Kmin, kMax - Kmin)
lK = kMax

End Function
Public Function DateJMA_Scan(lText As String, lK As Integer) As String
Dim K1 As Integer, K2 As Integer, lenText As Integer
Dim jj As Long, MM As Long, AA As Long, jma As String

K1 = lK + 1
lenText = Len(lText)
DateJMA_Scan = ""
K2 = InStr(K1, lText, "/")
If K2 > 0 Then
    jj = Val(Mid$(lText, K1, K2 - K1))
    K1 = K2 + 1
    K2 = InStr(K1, lText, "/")
    If K2 > 0 Then
        MM = Val(Mid$(lText, K1, K2 - K1))
        K1 = K2 + 1
        K2 = InStr(K1, lText, " ")
        If K2 > 0 Then
            AA = Val(Mid$(lText, K1, K2 - K1))
            If AA < 100 Then AA = 2000 + AA
        End If
    End If
End If
jma = Format(jj, "00") & "-" & Format(MM, "00") & "-" & Format(AA, "0000")
If IsDate(jma) Then
    DateJMA_Scan = Mid$(jma, 7, 4) & Mid$(jma, 4, 2) & Mid$(jma, 1, 2)
End If
lK = K2

End Function

Public Function TimeHMS_Scan(lText As String, lK As Integer) As String
Dim K1 As Integer, K2 As Integer, KScan As Integer, lenText As Integer
Dim HH As Long, MM As Long, ss As Long
K1 = lK + 1
KScan = K1
lenText = Len(lText)
TimeHMS_Scan = ""
K2 = InStr(K1, lText, ":")
If K2 > 0 Then
    KScan = K2
    HH = Val(Mid$(lText, K1, K2 - K1))
    K1 = K2 + 1
    K2 = InStr(K1, lText, ":")
    If K2 > 0 Then
        KScan = K2
        MM = Val(Mid$(lText, K1, K2 - K1))
        K1 = K2 + 1
        K2 = InStr(K1, lText, " ")
        If K2 > 0 Then
            KScan = K2
            ss = Val(Mid$(lText, K1, K2 - K1))
        End If
    End If
End If
TimeHMS_Scan = Format(HH, "00") & Format(MM, "00") & Format(ss, "00")
lK = KScan

End Function



Public Function Space_Scan(lText As String, lK As Integer) As String
Dim Kmin As Integer, kMax As Integer, lenText As Integer
Dim blnOk As Boolean

Kmin = lK + 1
lenText = Len(lText)
Space_Scan = ""
blnOk = False
For kMax = Kmin To lenText
    If Mid$(lText, kMax, 1) = " " Then
        If blnOk Then Exit For
    Else
        blnOk = True
    End If
    
Next kMax

If kMax > Kmin Then Space_Scan = Trim(Mid$(lText, Kmin, kMax - Kmin))
lK = kMax

End Function

Public Sub strMoveR(lX As Variant, lDest As String, lPos As Integer, lLen As Integer)
Dim lenX As Integer
lenX = Len(lX)
If lenX > lLen Then
    Mid$(lDest, lPos, lLen) = Mid$(lX, lenX - lLen + 1, lLen)
Else
    Mid$(lDest, lPos, lLen) = Space$(lLen - lenX) & lX
End If
End Sub

Public Sub DTPicker_Amj7(C As DTPicker, lAmj7 As Long)
Dim X8 As String * 8
Call DTPicker_Control(C, X8)
lAmj7 = CLng(X8) - 19000000

End Sub

Public Function file_Archive(lFile As String, lArchive_Folder As String) As String
Dim K As Integer, lenX As Integer
Dim X As String
On Error GoTo Error_Handler

K = InStr(1, lFile, lArchive_Folder)
If K > 0 Then
    file_Archive = "Fichier déjà archivé " & lFile
    Exit Function
End If
lenX = Len(lFile)
For K = lenX To 1 Step -1
    If Mid$(lFile, K, 1) = "\" Then Exit For
Next K

X = lArchive_Folder & DSys & "_" & time_Hms & "_" & Mid$(lFile, K + 1, lenX - K)
file_Archive = X
msFileSystem.MoveFile lFile, X

Exit Function

Error_Handler:
file_Archive = Error
Shell_MsgBox "#file_Archive# " & lFile & " / " & ":" & Error, vbCritical, "Archivage : " & lArchive_Folder, False
End Function


Public Function File_Export_Monitor(lFct As String, lId As Integer, lText As String)
Static mFile As String
Dim X As String
On Error GoTo Error_Handle
File_Export_Monitor = Null
Select Case lFct
    Case "Output":   mFile = lText
                    X = Dir(lText)
                    If X <> "" Then
                        X = MsgBox("Voulez-vous effacer le fichier existant ?", vbYesNo + vbQuestion + vbDefaultButton2, lText)
                        If X <> vbYes Then
                            File_Export_Monitor = "Exit"
                            Exit Function
                        End If
                    End If
                    lId = FreeFile
                    Open lText For Output As #lId
    Case "Print": Print #lId, lText
    Case "Input":   mFile = lText
                    X = Dir(lText)
                    If X = "" Then
                        X = MsgBox("Le fichier n'existe pas . ", vbCritical, lText)
                        File_Export_Monitor = "Exit"
                        Exit Function
                    End If
                    lId = FreeFile
                    Open lText For Input As #lId
    Case "Close": Close lId
    Case Else: Error 9999
End Select
Exit Function

Error_Handle:
File_Export_Monitor = Err & Error
MsgBox mFile & " : " & Error, vbCritical, "File_Export_Monitor"
Close
End Function


Public Sub Wait_SS(lSS As Long)
Dim SSS As Long
SSS = Time_Sys_Sss + lSS

Do
    DoEvents
Loop Until Time_Sys_Sss > SSS
End Sub

Public Sub DSYS_Init()
DSys = Year(Now)
Mid$(DSys, 5, 2) = Format$(Month(Now), "00")
Mid$(DSys, 7, 2) = Format$(Day(Now), "00")
valDSys = Val(DSys)
DSys_VeilleC = dateElp("Jour", -1, DSys)
DSys_VeilleO = dateElp("Ouvré", -1, DSys)
DSys_VeilleOAP = dateElp("Ouvré", -3, DSys) 'passé de 2 à 3 jours, le 17/10/2018
DSys_SuivantC = dateElp("Jour", 1, DSys)
DSys_SuivantO = dateElp("Ouvré", 1, DSys)
DSys_S = dateImp10_S(DSys)
dateDSys = dateImp10_S(DSys)
End Sub

Public Sub num_XPrt_Long(lNum As Long, lCurrentX As Long)
Dim X As String

If lNum = 0 Then
    X = "-"
Else
    X = Format(lNum, "### ### ##0")
End If
XPrt.CurrentX = lCurrentX - XPrt.TextWidth(X)
XPrt.Print X;

End Sub
Public Sub num_XPrt_Currency(lNum As Currency, lCurrentX As Long)
Dim X As String

If lNum = 0 Then
    X = "-"
Else
    X = Format(lNum, "### ### ### ### ##0.00")
End If
XPrt.CurrentX = lCurrentX - XPrt.TextWidth(X)
XPrt.Print X;

End Sub


Public Function time_N6(lTxt As String) As Long
Dim wTime As Long
wTime = Val(lTxt)
If wTime > 0 Then
    If wTime < 100 Then wTime = wTime * 100
    If wTime < 10000 Then wTime = wTime * 100
End If
time_N6 = wTime
End Function

Public Function SQL_Date_Time(lAMJ As String, lHMS As Long) As String
SQL_Date_Time = "{ts '" & Format$(lAMJ, "@@@@-@@-@@ ") & Format$(lHMS, "00:00:00") & ".000'}"
End Function

Public Function DSYS_Time() As String
DSYS_Time = DSys & "_" & time_Hms & "_"
End Function

Public Sub SAA_X32(lX As String, lX32V As Long, lX32D As String, lX32A As Currency)
Dim X8 As String
'format = :32A:vvvvvvdddaaaaa.....,aa
'==============================
lX32V = 0
lX32D = ""
lX32A = 0
If Len(lX) >= 14 Then
    Call dateJMA6_AMJ(Mid$(lX, 6, 6), X8)
    lX32V = CLng(X8)
    lX32D = Mid$(lX, 12, 3)
    lX32A = CCur(Mid$(lX, 15, Len(lX) - 14))
End If

End Sub


Public Function htmlFontColor(lX1 As String) As String
htmlFontColor = "<font color=" & Asc34 & lX1 & Asc34 & ">"
End Function

Public Function htmlbgColor(lX1 As String) As String
htmlbgColor = "<bgcolor=" & Asc34 & lX1 & Asc34 & ">"
End Function

Public Function colorHex_RGB(lColor As Long) As Long
Dim xColor As String, X As String
Dim lRed As Integer, lGreen As Integer, LBlue As Integer
lRed = lColor Mod 256
lGreen = Int(lColor / 256) Mod 256
LBlue = Int(lColor / 65536) Mod 256
colorHex_RGB = RGB(lRed, lGreen, LBlue)
End Function

Public Sub cnAdo_Info(cnAdo As ADODB.Connection)
On Error Resume Next
Dim X As String
Dim I As Integer
X = ""
For I = 0 To cnAdo.Properties.Count - 1
    X = X & vbCr & cnAdo.Properties(I).Name & " : " & cnAdo.Properties(I).value
Next I


MsgBox X

End Sub

Public Function EBCDIC_ASCII(lText As String) As String
Dim X As String, I As Integer
Dim lenX As Integer
X = ""
lenX = Len(lText)
For I = 1 To lenX Step 2
    Select Case Mid$(lText, I, 2)
        Case "40": X = X + " "
        
        Case "C1": X = X + "A"
        Case "C2": X = X + "B"
        Case "C3": X = X + "C"
        Case "C4": X = X + "D"
        Case "C5": X = X + "E"
        Case "C6": X = X + "F"
        Case "C7": X = X + "G"
        Case "C8": X = X + "H"
        Case "C9": X = X + "I"
        Case "D1": X = X + "J"
        Case "D2": X = X + "K"
        Case "D3": X = X + "L"
        Case "D4": X = X + "M"
        Case "D5": X = X + "N"
        Case "D6": X = X + "O"
        Case "D7": X = X + "P"
        Case "D8": X = X + "Q"
        Case "D9": X = X + "R"
        Case "E2": X = X + "S"
        Case "E3": X = X + "T"
        Case "E4": X = X + "U"
        Case "E5": X = X + "V"
        Case "E6": X = X + "W"
        Case "E7": X = X + "X"
        Case "E8": X = X + "Y"
        Case "E9": X = X + "Z"
         
        Case "F0": X = X + "0"
        Case "F1": X = X + "1"
        Case "F2": X = X + "2"
        Case "F3": X = X + "3"
        Case "F4": X = X + "4"
        Case "F5": X = X + "5"
        Case "F6": X = X + "6"
        Case "F7": X = X + "7"
        Case "F8": X = X + "8"
        Case "F9": X = X + "9"
 
        Case "81": X = X + "a"
        Case "82": X = X + "b"
        Case "83": X = X + "c"
        Case "84": X = X + "d"
        Case "85": X = X + "e"
        Case "86": X = X + "f"
        Case "87": X = X + "g"
        Case "88": X = X + "h"
        Case "89": X = X + "i"
        Case "91": X = X + "j"
        Case "92": X = X + "k"
        Case "93": X = X + "l"
        Case "94": X = X + "m"
        Case "95": X = X + "n"
        Case "96": X = X + "o"
        Case "97": X = X + "p"
        Case "98": X = X + "q"
        Case "99": X = X + "r"
        Case "A2": X = X + "s"
        Case "A3": X = X + "t"
        Case "A4": X = X + "u"
        Case "A5": X = X + "v"
        Case "A6": X = X + "w"
        Case "A7": X = X + "x"
        Case "A8": X = X + "y"
        Case "A9": X = X + "z"
       
        Case "5B": X = X + "$"
        Case "5C": X = X + "*"
        Case "6D": X = X + "_"
        Case Else: MsgBox Mid$(lText, I, 2), vbExclamation, "EBCDIC_ASCII"
    End Select
Next I
EBCDIC_ASCII = X
End Function

Public Function convX2P(lX As String)
Dim wX As String, lenX As Integer
Dim K As Integer
lenX = Len(lX)
wX = ""
For K = 1 To lenX
    wX = wX & arrX2P(Asc(Mid$(lX, K, 1)))
Next K
Select Case Asc(Mid$(lX, lenX, 1))
    Case 13, 29, 5, 21, 40, 41, 95, 39, 253, 184: convX2P = wX & "-"
    Case Else: convX2P = wX
End Select

End Function

Public Sub convX2P_IBMAMJ(lIBMAMJ As String, lA1 As Integer, lA2 As Integer, lA3 As Integer, lA4 As Integer)

lA1 = arrX2P_D(Val(Mid$(lIBMAMJ, 1, 2)))
lA2 = arrX2P_D(Val(Mid$(lIBMAMJ, 3, 2)))
lA3 = arrX2P_D(Val(Mid$(lIBMAMJ, 5, 2)))
lA4 = arrX2P_DF(Val(Mid$(lIBMAMJ, 7, 1)))

End Sub



Public Function convP2X(lV As Variant, lenX As Integer)
Dim wX As String, X As String
Dim K As Integer, kSign As Integer

If lV < 0 Then
    kSign = 120
Else
    kSign = 100
End If
X = Format$(lV, String(lenX, "0"))
wX = ""
For K = 1 To lenX - 1 Step 2
    'Debug.Print K, Asc(Mid$(X, K, 2))
    wX = wX & arrP2X(Val(Mid$(X, K, 2)))
    'Debug.Print K, Mid$(X, K, 2), Asc(arrP2X(Val(Mid$(X, K, 2))))
Next K
If lenX Mod 2 = 0 Then
    convP2X = wX
Else
    convP2X = wX & arrP2X(kSign + Val(Mid$(X, lenX, 1)))
End If
End Function

Public Sub convArray_JPL()
'wAMJ = CLng(convX2P(Mid$(xZBASTAB0.BASTABARG, 4, 4)))
'wCours = CDbl(convX2P(Mid$(xZBASTAB0.BASTABDON, 1, 8)))
'X2 = convP2X(wAMJ, 7)
'X2 = convP2X(wCours, 15)

'  0  0
'  1  1
'  2  2
'  3  3
'  4  156
'  5  9
'  6  134
'  7  127
'  8  151
'  9  141
'  10  16
'  11  17
'  12  18
'  13  19
'  14  157
'  15  133
'  16  8
'  17  135
'  18  24
'  19  25
'  20  128
'  21  129
'  22  130
'  23  131
'  24  132
'  25  10
'  26  23
'  27  27
'  28  136
'  29  137
'  30  144
'  31  145
'  32  22
'  33  147
'  34  148
'  35  149
'  36  150
'  37  4
'  38  152
'  39  153
'  40  32
'  41  160
'  42  226
'  43  228
'  44  64
'  45  225
'  46  227
'  47  229
'  48  92
'  49  241
'  50  38
'  51  123
'  52  234
'  53  235
'  54  125
'  55  237
'  56  238
'  57  239
'  58  236
'  59  223
'  60  45
'  61  47
'  62  194
'  63  196
'  64  192
'  65  193
'  66  195
'  67  197
'  68  199
'  69  209
'  70  248
'  71  201
'  72  202
'  73  203
'  74  200
'  75  205
'  76  206
'  77  207
'  78  204
'  79  181
'  80  216
'  81  97
'  82  98
'  83  99
'  84  100
'  85  101
'  86  102
'  87  103
'  88  104
'  89  105
'  90  91
'  91  106
'  92  107
'  93  108
'  94  109
'  95  110
'  96  111
'  97  112
'  98  113
'  99  114
'  100  15
'  101  31
'  102  7
'  103  26
'  104  33
'  105  94
'  106  63
'  107  34
'  108  177
'  109  164
'  110  12
'  111  28
'  112  140
'  113  20
'  114  60
'  115  42
'  116  37
'  117  224
'  118  240
'  119  230
'  120  13
'  121  29
'  122  5
'  123  21
'  124  40
'  125  41
'  126  95
'  127  39
'  128  253
'  129  184

arrX2P = Array("00", "01", "02", "03", "37", "2 ", "  ", "2 ", "16", "05", _
"25", "  ", "0 ", "0 ", "  ", "0 ", "10", "11", "12", "13", "3 ", "3 ", "32", "26", "18", "19", "3 ", "27", "1 ", "1 ", _
"  ", "1 ", "40", "4 ", "7 ", "  ", "  ", "6 ", "50", "7 ", "4 ", "5 ", "5 ", "  ", "  ", "60", "  ", "61", "  ", "  ", _
"  ", "  ", "  ", "  ", "  ", "  ", "  ", "  ", "  ", "  ", "4 ", "  ", "  ", "6 ", "44", "  ", "  ", "  ", "  ", "  ", _
"  ", "  ", "  ", "  ", "  ", "  ", "  ", "  ", "  ", "  ", "  ", "  ", "  ", "  ", "  ", "  ", "  ", "  ", "  ", "  ", _
"  ", "90", "48", "  ", "5 ", "6 ", "  ", "81", "82", "83", "84", "85", "86", "87", "88", "89", "91", "92", "93", "94", _
"95", "96", "97", "98", "99", "  ", "  ", "  ", "  ", "  ", "  ", "  ", "  ", "51", "  ", "54", "  ", "07", "20", "21", _
"22", "23", "24", "15", "06", "17", "28", "29", "  ", "  ", "2 ", "09", "  ", "  ", "30", "31", "  ", "33", "34", "35", _
"36", "08", "38", "39", "  ", "  ", "04", "14", "  ", "  ", "41", "  ", "  ", "  ", "9 ", "  ", "  ", "  ", "  ", "  ", _
"  ", "  ", "  ", "  ", "  ", "  ", "  ", "8 ", "  ", "  ", "  ", "79", "  ", "  ", "9 ", "  ", "  ", "  ", "  ", "  ", _
"  ", "  ", "64", "65", "62", "66", "63", "67", "  ", "68", "74", "71", "72", "73", "78", "75", "76", "77", "  ", "69", _
"  ", "  ", "  ", "  ", "  ", "  ", "80", "  ", "  ", "  ", "  ", "  ", "  ", "59", "7 ", "45", "42", "46", "43", "47", _
"9 ", "  ", "  ", "  ", "52", "53", "58", "55", "56", "57", "8 ", "49", "  ", "  ", "  ", "  ", "  ", "  ", "70", "  ", _
"  ", "  ", "  ", "8 ", "  ", "  ")


arrP2X = Array(Chr$(0), Chr$(1), Chr$(2), Chr$(3), Chr$(156), Chr$(9), Chr$(134), Chr$(127), Chr$(151), Chr$(141), _
Chr$(16), Chr$(17), Chr$(18), Chr$(19), Chr$(157), Chr$(133), Chr$(8), Chr$(135), Chr$(24), Chr$(25), Chr$(128), Chr$(129), Chr$(130), Chr$(131), Chr$(132), Chr$(10), Chr$(23), Chr$(27), Chr$(136), Chr$(137), _
Chr$(144), Chr$(145), Chr$(22), Chr$(147), Chr$(148), Chr$(149), Chr$(150), Chr$(4), Chr$(152), Chr$(153), Chr$(32), Chr$(160), Chr$(226), Chr$(228), Chr$(64), Chr$(225), Chr$(227), Chr$(229), Chr$(92), Chr$(241), _
Chr$(38), Chr$(123), Chr$(234), Chr$(235), Chr$(125), Chr$(237), Chr$(238), Chr$(239), Chr$(236), Chr$(223), Chr$(45), Chr$(47), Chr$(194), Chr$(196), Chr$(192), Chr$(193), Chr$(195), Chr$(197), Chr$(199), Chr$(209), _
Chr$(248), Chr$(201), Chr$(202), Chr$(203), Chr$(200), Chr$(205), Chr$(206), Chr$(207), Chr$(204), Chr$(181), Chr$(216), Chr$(97), Chr$(98), Chr$(99), Chr$(100), Chr$(101), Chr$(102), Chr$(103), Chr$(104), Chr$(105), _
Chr$(91), Chr$(106), Chr$(107), Chr$(108), Chr$(109), Chr$(110), Chr$(111), Chr$(112), Chr$(113), Chr$(114), Chr$(15), Chr$(31), Chr$(7), Chr$(26), Chr$(33), Chr$(94), Chr$(63), Chr$(34), Chr$(177), Chr$(164), _
Chr$(12), Chr$(28), Chr$(140), Chr$(20), Chr$(60), Chr$(42), Chr$(37), Chr$(224), Chr$(240), Chr$(230), Chr$(13), Chr$(29), Chr$(5), Chr$(21), Chr$(40), Chr$(41), Chr$(95), Chr$(39) & Chr$(39), Chr$(253), Chr$(184), _
Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), _
Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), _
Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), _
Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), _
Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), _
Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), _
Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32))

arrX2P_D(1) = 1
arrX2P_D(2) = 2
arrX2P_D(3) = 3
arrX2P_D(4) = 156
arrX2P_D(7) = 127
arrX2P_D(8) = 151
arrX2P_D(10) = 16
arrX2P_D(11) = 17
arrX2P_D(12) = 18
arrX2P_D(13) = 19
arrX2P_D(14) = 157
arrX2P_D(17) = 135
arrX2P_D(18) = 24
arrX2P_D(20) = 128
arrX2P_D(21) = 129
arrX2P_D(22) = 130
arrX2P_D(23) = 131
arrX2P_D(24) = 132
arrX2P_D(27) = 27
arrX2P_D(28) = 136
arrX2P_D(29) = 137
arrX2P_D(30) = 144
arrX2P_D(31) = 145
arrX2P_D(32) = 22
arrX2P_D(33) = 147
arrX2P_D(34) = 148
arrX2P_D(35) = 149
arrX2P_D(36) = 147
arrX2P_D(37) = 4
arrX2P_D(40) = 32
arrX2P_D(41) = 160
arrX2P_D(42) = 226
arrX2P_D(43) = 228
arrX2P_D(47) = 229
arrX2P_D(48) = 228
arrX2P_D(50) = 38
arrX2P_D(51) = 123
arrX2P_D(52) = 234
arrX2P_D(53) = 235
arrX2P_D(56) = 23
arrX2P_D(57) = 239
arrX2P_D(58) = 236
arrX2P_D(60) = 45
arrX2P_D(61) = 47
arrX2P_D(62) = 194
arrX2P_D(63) = 196
arrX2P_D(64) = 192
arrX2P_D(68) = 199
arrX2P_D(70) = 248
arrX2P_D(71) = 201
arrX2P_D(72) = 202
arrX2P_D(73) = 203
arrX2P_D(74) = 204
arrX2P_D(75) = 205
arrX2P_D(80) = 216
arrX2P_D(81) = 97
arrX2P_D(82) = 98
arrX2P_D(83) = 99
arrX2P_D(86) = 102
arrX2P_D(87) = 103
arrX2P_D(88) = 104
arrX2P_D(89) = 105
arrX2P_D(90) = 91
arrX2P_D(91) = 106
arrX2P_D(92) = 107
arrX2P_D(93) = 108
arrX2P_D(94) = 109
arrX2P_D(95) = 110
arrX2P_D(96) = 111
arrX2P_D(97) = 112
arrX2P_D(98) = 113
arrX2P_D(99) = 114


arrX2P_DF(0) = 15
arrX2P_DF(1) = 31
arrX2P_DF(2) = 7
arrX2P_DF(3) = 26
arrX2P_DF(4) = 33
arrX2P_DF(5) = 94
arrX2P_DF(6) = 63
arrX2P_DF(7) = 34
arrX2P_DF(8) = 26
arrX2P_DF(9) = 26
End Sub

Public Sub convArray()
'wAMJ = CLng(convX2P(Mid$(xZBASTAB0.BASTABARG, 4, 4)))
'wCours = CDbl(convX2P(Mid$(xZBASTAB0.BASTABDON, 1, 8)))
'X2 = convP2X(wAMJ, 7)
'X2 = convP2X(wCours, 15)

'  0  0
'  1  1
'  2  2
'  3  3
'  4  156
'  5  9
'  6  134
'  7  127
'  8  151
'  9  141
'  10  16
'  11  17
'  12  18
'  13  19
'  14  157
'  15  133
'  16  8
'  17  135
'  18  24
'  19  25
'  20  128
'  21  129
'  22  130
'  23  131
'  24  132
'  25  10
'  26  23
'  27  27
'  28  136
'  29  137
'  30  144
'  31  145
'  32  22
'  33  147
'  34  148
'  35  149
'  36  150
'  37  4
'  38  152
'  39  153
'  40  32
'  41  160
'  42  226
'  43  228
'  44  64
'  45  225
'  46  227
'  47  229
'  48  92
'  49  241
'  50  38
'  51  123
'  52  234
'  53  235
'  54  125
'  55  237
'  56  238
'  57  239
'  58  236
'  59  223
'  60  45
'  61  47
'  62  194
'  63  196
'  64  192
'  65  193
'  66  195
'  67  197
'  68  199
'  69  209
'  70  248
'  71  201
'  72  202
'  73  203
'  74  200
'  75  205
'  76  206
'  77  207
'  78  204
'  79  181
'  80  216
'  81  97
'  82  98
'  83  99
'  84  100
'  85  101
'  86  102
'  87  103
'  88  104
'  89  105
'  90  91
'  91  106
'  92  107
'  93  108
'  94  109
'  95  110
'  96  111
'  97  112
'  98  113
'  99  114
'  100  15
'  101  31
'  102  7
'  103  26
'  104  33
'  105  94
'  106  63
'  107  34
'  108  177
'  109  164
'  110  12
'  111  28
'  112  140
'  113  20
'  114  60
'  115  42
'  116  37
'  117  224
'  118  240
'  119  230
'  120  13
'  121  29
'  122  5
'  123  21
'  124  40
'  125  41
'  126  95
'  127  39
'  128  253
'  129  184

arrX2P = Array("00", "01", "02", "03", "37", "2 ", "  ", "2 ", "16", "05", _
"25", "  ", "0 ", "0 ", "  ", "0 ", "10", "11", "12", "13", "3 ", "3 ", "32", "26", "18", "19", "3 ", "27", "1 ", "1 ", _
"  ", "1 ", "40", "4 ", "7 ", "  ", "  ", "6 ", "50", "7 ", "4 ", "5 ", "5 ", "  ", "  ", "60", "  ", "61", "  ", "  ", _
"  ", "  ", "  ", "  ", "  ", "  ", "  ", "  ", "  ", "  ", "4 ", "  ", "  ", "6 ", "44", "  ", "  ", "  ", "  ", "  ", _
"  ", "  ", "  ", "  ", "  ", "  ", "  ", "  ", "  ", "  ", "  ", "  ", "  ", "  ", "  ", "  ", "  ", "  ", "  ", "  ", _
"  ", "90", "48", "  ", "5 ", "6 ", "  ", "81", "82", "83", "84", "85", "86", "87", "88", "89", "91", "92", "93", "94", _
"95", "96", "97", "98", "99", "  ", "  ", "  ", "  ", "  ", "  ", "  ", "  ", "51", "  ", "54", "  ", "07", "20", "21", _
"22", "23", "24", "15", "06", "17", "28", "29", "  ", "  ", "2 ", "09", "  ", "  ", "30", "31", "  ", "33", "34", "35", _
"36", "08", "38", "39", "  ", "  ", "04", "14", "  ", "  ", "41", "  ", "  ", "  ", "9 ", "  ", "  ", "  ", "  ", "  ", _
"  ", "  ", "  ", "  ", "  ", "  ", "  ", "8 ", "  ", "  ", "  ", "79", "  ", "  ", "9 ", "  ", "  ", "  ", "  ", "  ", _
"  ", "  ", "64", "65", "62", "66", "63", "67", "  ", "68", "74", "71", "72", "73", "78", "75", "76", "77", "  ", "69", _
"  ", "  ", "  ", "  ", "  ", "  ", "80", "  ", "  ", "  ", "  ", "  ", "  ", "59", "7 ", "45", "42", "46", "43", "47", _
"9 ", "  ", "  ", "  ", "52", "53", "58", "55", "56", "57", "8 ", "49", "  ", "  ", "  ", "  ", "  ", "  ", "70", "  ", _
"  ", "  ", "  ", "8 ", "  ", "  ")


arrP2X = Array(Chr$(0), Chr$(1), Chr$(2), Chr$(3), Chr$(156), Chr$(9), Chr$(134), Chr$(127), Chr$(151), Chr$(141), _
Chr$(16), Chr$(17), Chr$(18), Chr$(19), Chr$(157), Chr$(133), Chr$(8), Chr$(135), Chr$(24), Chr$(25), Chr$(128), Chr$(129), Chr$(130), Chr$(131), Chr$(132), Chr$(10), Chr$(23), Chr$(27), Chr$(136), Chr$(137), _
Chr$(144), Chr$(145), Chr$(22), Chr$(147), Chr$(148), Chr$(149), Chr$(150), Chr$(4), Chr$(152), Chr$(153), Chr$(32), Chr$(160), Chr$(226), Chr$(228), Chr$(64), Chr$(225), Chr$(227), Chr$(229), Chr$(92), Chr$(241), _
Chr$(38), Chr$(123), Chr$(234), Chr$(235), Chr$(125), Chr$(237), Chr$(238), Chr$(239), Chr$(236), Chr$(223), Chr$(45), Chr$(47), Chr$(194), Chr$(196), Chr$(192), Chr$(193), Chr$(195), Chr$(197), Chr$(199), Chr$(209), _
Chr$(248), Chr$(201), Chr$(202), Chr$(203), Chr$(200), Chr$(205), Chr$(206), Chr$(207), Chr$(204), Chr$(181), Chr$(216), Chr$(97), Chr$(98), Chr$(99), Chr$(100), Chr$(101), Chr$(102), Chr$(103), Chr$(104), Chr$(105), _
Chr$(91), Chr$(106), Chr$(107), Chr$(108), Chr$(109), Chr$(110), Chr$(111), Chr$(112), Chr$(113), Chr$(114), Chr$(15), Chr$(31), Chr$(7), Chr$(26), Chr$(33), Chr$(94), Chr$(63), Chr$(34), Chr$(177), Chr$(164), _
Chr$(12), Chr$(28), Chr$(140), Chr$(20), Chr$(60), Chr$(42), Chr$(37), Chr$(224), Chr$(240), Chr$(230), Chr$(13), Chr$(29), Chr$(5), Chr$(21), Chr$(40), Chr$(41), Chr$(95), Chr$(39) & Chr$(39), Chr$(253), Chr$(184), _
Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), _
Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), _
Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), _
Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), _
Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), _
Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), _
Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32), Chr$(32))

arrX2P_D(1) = 1
arrX2P_D(2) = 2
arrX2P_D(3) = 3
arrX2P_D(10) = 16
arrX2P_D(11) = 17
arrX2P_D(12) = 18
arrX2P_D(13) = 19
arrX2P_D(20) = 26
arrX2P_D(21) = 26
arrX2P_D(22) = 26
arrX2P_D(23) = 26
arrX2P_D(30) = 26
arrX2P_D(31) = 26
arrX2P_D(32) = 22
arrX2P_D(33) = 26
arrX2P_D(40) = 32
arrX2P_D(41) = 26
arrX2P_D(42) = 26
arrX2P_D(43) = 26
arrX2P_D(50) = 38
arrX2P_D(51) = 123
arrX2P_D(52) = 26
arrX2P_D(53) = 26
arrX2P_D(60) = 45
arrX2P_D(61) = 47
arrX2P_D(62) = 26
arrX2P_D(63) = 26
arrX2P_D(70) = 26
arrX2P_D(71) = 26
arrX2P_D(72) = 26
arrX2P_D(73) = 26
arrX2P_D(80) = 26
arrX2P_D(81) = 97
arrX2P_D(82) = 98
arrX2P_D(83) = 99
arrX2P_D(90) = 91
arrX2P_D(91) = 106
arrX2P_D(92) = 107
arrX2P_D(93) = 108

arrX2P_DF(0) = 15
arrX2P_DF(1) = 31
arrX2P_DF(2) = 7
arrX2P_DF(3) = 26
arrX2P_DF(4) = 33
arrX2P_DF(5) = 94
arrX2P_DF(6) = 63
arrX2P_DF(7) = 34
arrX2P_DF(8) = 26
arrX2P_DF(9) = 26
End Sub


Public Function File_CACLS(lFileName As String, lUser As String, lUnit As String)
Dim wCacls As String, IdShell, X As String
Dim K As Integer, wText As String
Dim xName As String, xMemo As String
Dim V
Dim wUser As String

File_CACLS = Null
'DR 15/12/2020
'If InStr(1, UCase(lFileName), "CPT096P1") > 0 Then
'    Call ECRIT_LOG_CPT096P1("CACLS " & lFileName & " \ " & lUser & " \ " & lUnit)
'End If
'               '
wUser = Trim(lUser)
X = ""
If Trim(lUnit) <> "" Then
    'DR 16/12/2020
    'If InStr(1, UCase(lFileName), "CPT096P1") > 0 And lUnit = "CPT" Then
    '    V = rsElpTable_Read("Unit", "CPTP", "CACLS", xName, xMemo)
    'Else
        V = rsElpTable_Read("Unit", lUnit, "CACLS", xName, xMemo)
    'End If
    '               '
    If Trim(xMemo) <> "" Then
        K = 0
        Do
            wText = CSV_Scan(xMemo, K)
            If wText <> "" And wText <> wUser Then X = X & " bia-paris\" & wText & ":F"
        Loop Until wText = ""
    End If
End If

wCacls = "cacls " & lFileName & " /E  /r ""bia-paris\Utilisa. du domaine""   /g bia-paris\" & wUser & ":F" & X
IdShell = Shell(wCacls, 0)
DoEvents
If IdShell > 0 Then
    On Error Resume Next            '$$$$$$$$$$$
    AppActivate IdShell, True       '$$$$$$$$$$$
End If
DoEvents

End Function
Public Sub ECRIT_LOG_CPT096P1(mes As String)
Dim fic As Long

    fic = FreeFile
    Open "C:\Temp\CPT096P1.log" For Append As #fic
    Print #fic, mes & Format(Now, "dd/MM/yyyy  HH:nn:ss")
    Close #fic
    
End Sub
Public Function File_ICACLS(lFileName As String, lUser As String, lUnit As String)
'_____________________________________________
'$JPL 2014-09-22 nouvelle version ICACLS
'_____________________________________________
Dim wCacls As String, IdShell, wUnit As String
Dim K As Integer, wText As String
Dim xName As String, xMemo As String
Dim V
Dim wUser As String

File_ICACLS = Null
'DR 15/12/2020
'If InStr(1, UCase(lFileName), "CPT096P1") > 0 Then
'    Call ECRIT_LOG_CPT096P1("ICACLS " & lFileName & " \ " & lUser & " \ " & lUnit)
'End If
'               '

If Trim(lUser) <> "" Then wUser = "BIA-PARIS\" & Trim(lUser) & ":R"

wUnit = ""
If Trim(lUnit) <> "" Then
    'DR 16/12/2020
    'If InStr(1, UCase(lFileName), "CPT096P1") > 0 And lUnit = "CPT" Then
    '    V = rsElpTable_Read("Unit", "CPTP", "CACLS", xName, xMemo)
    'Else
        V = rsElpTable_Read("Unit", lUnit, "CACLS", xName, xMemo)
    'End If
    '               '
    If Trim(xMemo) <> "" Then
        K = 0
        Do
            wText = CSV_Scan(xMemo, K)
            If wText <> "" And wText <> wUser Then wUnit = wUnit & " bia-paris\" & wText & ":F"
        Loop Until wText = ""
    End If
End If

If wUser = "" And wUnit = "" Then
    wCacls = "icacls " & Asc34 & lFileName & Asc34 & " /grant ""BIA-PARIS\Utilisa. du domaine:R""  "
Else

    wCacls = "icacls " & Asc34 & lFileName & Asc34 & " /remove:g ""BIA-PARIS\Utilisa. du domaine""  /grant " & wUser & wUnit
    
    If paramEnvironnement = constTest Then wCacls = wCacls & " bia-paris\LOULERGUE:F"
End If

    IdShell = Shell(wCacls, 0)
    DoEvents
    If IdShell > 0 Then
        On Error Resume Next            '$$$$$$$$$$$
        AppActivate IdShell, True       '$$$$$$$$$$$
    End If
    DoEvents
'End If


End Function

Public Function Date_VB(lAMJ As Long, lHMS As Long) As Date
If lHMS = 0 Then
    Date_VB = CDate(Format$(Mid$(lAMJ, 7, 2) & Mid$(lAMJ, 5, 2) & Mid$(lAMJ, 1, 4), "@@/@@/@@@@"))
Else
    Date_VB = CDate(Format$(Mid$(lAMJ, 7, 2) & Mid$(lAMJ, 5, 2) & Mid$(lAMJ, 1, 4), "@@/@@/@@@@") & " " & Format$(Mid$(lHMS, 1, 6), " @@:@@:@@"))
End If

End Function

Public Function num_Taux_Display(lTaux As Double)
Dim X As String, K As Integer, K1 As Integer

X = Format$(Abs(lTaux), "##0.00000")
K = InStr(X, ",")
If K > 0 Then
    For K1 = Len(X) To K + 2 Step -1
        If Mid$(X, K1, 1) = "0" Then
            Mid$(X, K1, 1) = " "
        Else
            Exit For
        End If
    Next K1
    
End If

Select Case lTaux
    Case Is > 0: num_Taux_Display = " + " & X
    Case Is < 0: num_Taux_Display = " - " & X
    Case Else: num_Taux_Display = ""
End Select
End Function

Public Function Windows_Processus_Actif(lProcess) As Integer
    Dim hThreadSnap As Long
    Dim bRet As Boolean
    Dim hProcessSnap As Long
    Dim pe32 As PROCESSENTRY32
    Dim Nb As Integer
    
    hProcessSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    If (hProcessSnap = -1) Then Exit Function
    pe32.dwSize = Len(pe32)
    bRet = Process32First(hProcessSnap, pe32)
    If (bRet = 1) Then
        While (Process32Next(hProcessSnap, pe32))
        'Debug.Print pe32.szExeFile
        If InStr(pe32.szExeFile, lProcess) > 0 Then Nb = Nb + 1
            
        Wend
    End If
Windows_Processus_Actif = Nb
End Function

Public Function num_String_Auto(ByVal lX As String) As Variant
Dim maxI As Integer, maxD As Integer
Dim X As String, K As Integer, blnD As Boolean
X = Trim(lX)
For K = 1 To Len(X)
    If blnD Then
        maxD = maxD + 1
    Else
        If Mid$(X, K, 1) = "." Or Mid$(X, K, 1) = "," Then
            blnD = True
        Else
            maxI = maxI + 1
        End If
    End If
Next K
num_String_Auto = num_String(X, maxI, maxD)
End Function

Public Function KillProcess(ByVal ProcessName As String) As Boolean
    Dim svc As Object
    Dim sQuery As String
    Dim oproc
    Set svc = GetObject("winmgmts:root\cimv2")
    sQuery = "select * from win32_process where name='" & ProcessName & "'"
    For Each oproc In svc.execquery(sQuery)
        oproc.Terminate
    Next
    Set svc = Nothing

End Function

Public Function Windows_Service_Actif(lProcess) As Integer
    Dim svc As Object
    Dim sQuery As String
    Dim oserv
    On Error Resume Next
 
    Set svc = GetObject("winmgmts:root\cimv2")
    sQuery = "select * from win32_service"
    For Each oserv In svc.execquery(sQuery)
        'Debug.Print oserv.Name & " : " & oserv.Pathname & " : " & oserv.State
        If InStr(oserv.Name, lProcess) > 0 Then
            If oserv.State = "Running" Then
                Windows_Service_Actif = 1: Exit Function
            Else
                Windows_Service_Actif = -1: Exit Function
            End If
        End If
    Next
    Set svc = Nothing


End Function
