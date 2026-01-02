Attribute VB_Name = "ElpVb4"
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
'Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
'Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
'SetWindowRgn hWnd, CreateEllipticRgn(0, 0, 600, 500), True
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long
Public Declare Function Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) As Long


Public paramIBM_BIA_Auto As String
Public paramIBM_BIA_ODBC As String

Public paramODBC_DSN_SAB As String
Public paramODBC_DSN_BIADWH As String
Public paramODBC_DSN_JRN As String
Public paramODBC_DSN_SAB073Y As String

Public paramODBC_DSN_CHQ_SCAN_ARCHIVE As String
Public paramODBC_DSN_CHQ_SCAN_LOCAL As String


Public MDB As Database
Public msFileSystem, msFile
Public blnRéplication_Load  As Boolean
Public App_EXEName As String
Public App_Title As String

Public srvIdle As Boolean
Public tauxTVA As Single, tauxTTC As Single
Public Const Pi = 3.14159
Public Const picLineHeight = 320
Public Const lstLineHeight = 255
Public lX As Long, Vx As Variant
Public libMois(12) As String, libMonth(12) As String
Public Const constDateZ = "__ - __ - ____"
Public Const Format_Date = "## - ## - ####"
Public Const Format_Time = "## : ##"
Public Const Format_N5 = "#####"
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
Public paramEnvironement As String

Public constWinWord As String, constWinWord_D As String
Public constExcel As String, constExcel_D As String
Public Const constWordPad = "c:\Program Files\Windows NT\Accessoires\WordPad.exe"
Public Const constMsPaint = "c:\WinNT\System32\MsPaint.exe"
Public DataBase_Open As String, DataBase_Master As String, DataBase_Local As String
Public DataBase_Data As String
Public Const paramDataBase_Password = ";pwd=l2206"
Public paramFolder_Master As String, paramFolder_Local As String

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
Public DSys As String * 8
Public valDSys As Long
Public DSys_VeilleC As String * 8, DSys_VeilleO As String * 8, DSys_VeilleOAP As String * 8
Public DSys_SuivantC As String * 8, DSys_SuivantO As String * 8
Public usrId As String, usrIdNT As String
Public usrName As String, usrName_UCase As String
Public usrService As String
Public usrCompte As String * 11, usrRacine As String * 5
Public usrGestionnaire As String
Public pcIdUsrIdCtl As Boolean
Public usrSituationCompte_Forçage As Boolean
Public usrService_DisplayAll As Boolean
Public SrvDir As String
Public Elp As typeXcom
'-----------------------------------------------------
Public socName As String, SocRibDom As String, socTéléphone As String
Public SocBdfE As Integer, strSocBdfE As String * 5
Public SocBdfG As Integer, strSocBdfG As String * 5
Public SocId$
Public SocAgence$
Public paramBic8 As String * 8, SocBicId As String * 11, SocBicIdNostro As String * 11
Public strSocSignon As String
Public imgSocSigle As String
Public imgSocLogo As String
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
Public XPrt
Public XListBox As ListBox
Public XLabel As Label
Public XControl As Control
Public XImage As IMAGE
'-----------------------------------------------------
Public frmElp_Caption As String
Public frmElp_Icon As String
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
Public Const prtForeColor_Header = vbGreen
Public Const prtForeColor = vbBlack

Public prtFormType As String
Public prtLineNb As Integer
Public prtlineHeight As Integer, prtlineHeight66 As Integer
Public prtHeaderHeight As Integer
Public prtParagraphHeight As Integer
Public prtZoom As Integer
Public prtShow As Boolean
Public blnFiligrane As Boolean
Public prtFiligrane_Name As String
Public Const paramElpCypher = "AntiHacker"

'---------------------------------------------------------
Type typeDate
    AA  As String * 4
    MM  As String * 2
    JJ As String * 2
End Type
'---------------------------------------------------------
Type typeUsrColor
    BackColor  As Long
    ForeColor  As Long
End Type
'---------------------------------------------------------
Type typeParamSnap
    AmjMin  As String * 8
    HMSMin  As String * 6
    AmjMax  As String * 8
    HMSMax  As String * 6
    selK1   As String * 20
    selK2   As String * 20
    selK3   As String * 20
    sortK1  As String * 20
    sortK2  As String * 20
    sortK3  As String * 20
    prtDétail  As Boolean
    prtK1  As Boolean
    prtK2  As Boolean
    prtK3  As Boolean
    Nb  As Long
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
Public recparamServer As typeElpTable
Public mCommand As String

Public countTimer As Long
Public blnJPL As Boolean
Public paramTemp_Folder As String
Public blnAuto_Form_Show As Boolean





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
    wMsg = wMsg & Format$(Asc(mId$(mMsg, I, 1)) Xor Asc(mId$(wKey, I, 1)), "000")
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
    wMsg = wMsg & Chr$(Val(mId$(mMsg, L2, 3)) Xor Asc(mId$(wKey, I, 1)))
    L2 = L2 + 3
Next I
ElpCipher_D = wMsg
End Function

Public Sub DTPicker_Amj8_tiret(C As DTPicker, lAmj8_tiret As String)
Dim X8 As String * 8

Call DTPicker_Control(C, X8)   ' X8 : format aaaammjj
lAmj8_tiret = mId$(X8, 1, 4) & "-" & mId$(X8, 5, 2) & "-" & mId$(X8, 7, 2)

End Sub

Public Function Text_KeyWord(lText As String, lK As Integer, blnSelectAll As Boolean) As String
Dim Kmin As Integer, Kmax As Integer, lenText As Integer, xKeyWord As String, blnOk As Boolean
Dim X1 As String, blnKeyWord As Boolean

lenText = Len(lText)
blnKeyWord = False
Do
    Kmin = lK + 1
    xKeyWord = ""
    blnOk = False
    For Kmax = Kmin To lenText
        X1 = mId$(lText, Kmax, 1)
        Select Case X1
            Case ".", "-", "_":
            Case "a" To "z": xKeyWord = xKeyWord & X1: blnOk = True
            Case "0" To "9": xKeyWord = xKeyWord & X1: blnOk = True
            Case Else: If blnOk Then Exit For
        End Select
                
    Next Kmax
    
    If Kmax >= lenText Then blnKeyWord = True
    lK = Kmax
    
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
    Select Case mId$(X, I, 1)
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

Public Sub Text_Accent(lText As String)
Dim X As String, I As Integer

X = LCase(Trim(lText))

' Voir aussi Text_LCase

For I = 1 To Len(X)
    Select Case mId$(X, I, 1)
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


Public Function dateJma08_Amj08(ljma08 As String, lAmj As String)
Dim Siecle2C As String

lAmj = "00000000"
If Trim(ljma08) = "" Then Exit Function
If mId$(ljma08, 7, 2) >= 90 Then
   Siecle2C = "19"
Else
   Siecle2C = "20"
End If
lAmj = Siecle2C & mId$(ljma08, 7, 2) & mId$(ljma08, 4, 2) & mId$(ljma08, 1, 2)

End Function


Public Function Time_Hms_Sss(lHms As String) As Long
Time_Hms_Sss = mId$(lHms, 1, 2) * 3600 + mId$(lHms, 3, 2) * 60 + mId$(lHms, 5, 2)
End Function

Public Function Time_Sys_Sss() As Long
Dim X As String
X = Time
Time_Sys_Sss = mId$(X, 1, 2) * 3600 + mId$(X, 4, 2) * 60 + mId$(X, 7, 2)
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

usrIdNT = usrId
usrName = usrIdNT
usrCompte = "": usrRacine = ""

Elp.usrId = usrId
Xcom_UsrId usrId
elpSrvXcom = ""
mainSoc
elpSrvXcom = "XXXX"

End Sub

Function dateCtlDsys(ByVal X As String)
'---------------------------------------------------------------------

If mId$(X, 11, 4) = "____" And mId$(X, 6, 2) = "__" And mId$(X, 1, 2) = "__" Then
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

AMJ.AA = Format$(Val(mId$(X, 11, 4)), "0000")

AMJ.MM = Format$(Val(mId$(X, 6, 2)), "00")
AMJ.JJ = Format$(Val(mId$(X, 1, 2)), "00")


'If Amj.aa = "____" And Amj.mm = "__" And Amj.jj = "__" Then
If AMJ.AA = "0000" And AMJ.MM = "00" And AMJ.JJ = "00" Then
    dateCtl = "00000000"
Else
    If AMJ.AA = "____" Or AMJ.AA = "0000" Then
        AMJ.AA = mId$(DSys, 1, 4)
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
        AMJ.MM = mId$(DSys, 5, 2)
    End If
    If AMJ.JJ = "__" Or AMJ.JJ = "00" Then
        AMJ.JJ = mId$(DSys, 7, 2)
    End If
   
 '   jma = "# " & AMJ.jj & "-" & AMJ.mm & "-" & AMJ.aa & " #"
        jma = AMJ.JJ & "-" & AMJ.MM & "-" & AMJ.AA
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
                dateCtl = AMJ.AA & AMJ.MM & AMJ.JJ
            End If
        End If
    End If
End If
End Function

'---------------------------------------------------------------------
Sub dateDsp(ByVal X As String, C As Control)
'---------------------------------------------------------------------
If TypeOf C Is MaskEdBox Then
    C.Mask = Format$("")
End If

C.Text = dateImp(X)

If TypeOf C Is MaskEdBox Then
    C.Mask = Format_Date
End If

End Sub
'---------------------------------------------------------
Function dateImp(ByVal X As String) As String
'---------------------------------------------------------

If X = "00000000" Or RTrim(X) = "" Then
    dateImp = Space$(14)
Else
    dateImp = Format$(mId$(X, 7, 2) & mId$(X, 5, 2) & mId$(X, 1, 4), "@@ - @@ - @@@@")
End If

End Function

'---------------------------------------------------------
Function dateImp10(ByVal X As String) As String
'---------------------------------------------------------

If X = "00000000" Or RTrim(X) = "" Then
    dateImp10 = Space$(10)
Else
    dateImp10 = Format$(mId$(X, 7, 2) & mId$(X, 5, 2) & mId$(X, 1, 4), "@@.@@.@@@@")
End If

End Function
'---------------------------------------------------------
Function dateJma6_Imp10(ByVal X As String) As String
'---------------------------------------------------------
If Trim(X) = "" Then
    dateJma6_Imp10 = ""
Else
    dateJma6_Imp10 = Format$(mId$(X, 1, 2) & mId$(X, 3, 2) & mId$(X, 5, 2), "@@.@@.@@")
End If
End Function

'---------------------------------------------------------
Function dateAMJ6_Imp10(ByVal X As String) As String
'---------------------------------------------------------
If Trim(X) = "" Then
    dateAMJ6_Imp10 = ""
Else
    dateAMJ6_Imp10 = Format$(mId$(X, 5, 2) & mId$(X, 3, 2) & "20" & mId$(X, 1, 2), "@@.@@.@@@@")
End If
End Function

'---------------------------------------------------------
Function dateAMJ10(ByVal X As String) As String
'---------------------------------------------------------

If X = "00000000" Or RTrim(X) = "" Then
    dateAMJ10 = Space$(10)
Else
    dateAMJ10 = Format$(mId$(X, 1, 4) & mId$(X, 5, 2) & mId$(X, 7, 2), "@@@@.@@.@@")
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
    dateImp_Amj = Format$(mId$(X, 1, 4) & mId$(X, 5, 2) & mId$(X, 7, 2), "@@@@-@@-@@")
End If

End Function

'---------------------------------------------------------
Function dateImpS(ByVal X As String) As String
'---------------------------------------------------------

If X = "00000000" Or RTrim(X) = "" Then
    dateImpS = Space$(8)
Else
    dateImpS = Format$(mId$(X, 7, 2) & mId$(X, 5, 2) & mId$(X, 3, 2), "@@-@@-@@")
End If

End Function



'---------------------------------------------------------
Function dateImp_jjMoisAAAA(ByVal X As String) As String
'---------------------------------------------------------
Dim I As Integer
I = Val(mId$(X, 5, 2))
If I < 0 Or I > 12 Then
    dateImp_jjMoisAAAA = ""
Else
    dateImp_jjMoisAAAA = Format$(mId$(X, 7, 2), "@@ ") & Trim(libMois(I)) & Format$(mId$(X, 1, 4), " @@@@")
End If

End Function

'---------------------------------------------------------
Function dateImp_ddMonthYYYY(ByVal X As String) As String
'---------------------------------------------------------
Dim I As Integer
I = Val(mId$(X, 5, 2))
If I < 0 Or I > 12 Then
    dateImp_ddMonthYYYY = ""
Else
    dateImp_ddMonthYYYY = Format$(mId$(X, 7, 2), "@@ ") & Trim(libMonth(I)) & Format$(mId$(X, 1, 4), " @@@@")
End If

End Function

'---------------------------------------------------------
Function dateElp(ByVal Fct As String, ByVal Nb As Integer, ByVal X As String) As String
'---------------------------------------------------------

Dim K As Integer, K1 As Integer, K2 As Integer
Dim V, X8 As String * 8, X8B As String * 8
Dim Fct_Mod As String, Nb_Mod As Integer

dateElp = X
Select Case Fct
        Case "TrimestreAdd": Fct_Mod = "MoisAdd": Nb_Mod = Nb * 3
        Case "SemestreAdd": Fct_Mod = "MoisAdd": Nb_Mod = Nb * 6
        Case "AnAdd": Fct_Mod = "MoisAdd": Nb_Mod = Nb * 12
        Case Else: Fct_Mod = Fct: Nb_Mod = Nb
End Select

Select Case Fct_Mod
   Case "Decade"
        Select Case mId$(X, 7, 2)
            Case Is < 11: dateElp = mId$(X, 1, 6) & "10"
            Case Is < 21: dateElp = mId$(X, 1, 6) & "20"
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
        V = Format$(mId$(X, 7, 2) & mId$(X, 5, 2) & mId$(X, 1, 4), "@@ - @@ - @@@@")
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
        K = Val(mId$(X8, 5, 2))
        If K > 1 Then
            K = K - 1
            Mid$(X8, 5, 2) = Format$(K, "00")
        Else
            K = Val(mId$(X8, 1, 4)) - 1
            Mid$(X8, 1, 4) = Format$(K, "0000")
            Mid$(X8, 5, 2) = "12"
        End If
        dateElp = dateFinDeMois(X8)
    Case "FinDAnnéeP"
        X8 = X
        K = Val(mId$(X8, 1, 4)) - 1
        Mid$(X8, 1, 4) = Format$(K, "0000")
        Mid$(X8, 5, 4) = "1231"
        dateElp = dateFinDeMois(X8)
    
    Case "MoisAdd"
        K1 = Fix(Abs(Nb_Mod) / 12)
        K2 = Abs(Nb_Mod) Mod 12
        X8 = X
        K = Val(mId$(X8, 5, 2))
   
        If Nb_Mod < 0 Then
            K1 = -K1
            If K > K2 Then
                K = K - K2
                Mid$(X8, 5, 2) = Format$(K, "00")
            Else
                K1 = K1 - 1
                Mid$(X8, 5, 2) = Format$(Val(mId$(X8, 5, 2)) - K2 + 12, "00")
            End If
        Else
            If K + K2 <= 12 Then
                K = K + K2
                Mid$(X8, 5, 2) = Format$(K, "00")
            Else
                K1 = K1 + 1
                Mid$(X8, 5, 2) = Format$(Val(mId$(X8, 5, 2)) + K2 - 12, "00")
            End If
        End If
                 
        K = Val(mId$(X8, 1, 4)) + K1
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

End Select

End Function

'---------------------------------------------------------
Function DateElp_X(ByVal Msg As String, ByVal xAMJ As String) As String
'---------------------------------------------------------
Dim K As Integer, Nb As Integer
Dim Fct As String

K = InStr(Msg, " ")
Nb = Val(mId$(Msg, 1, K - 1))
Fct = mId$(Msg, K + 1, Len(Msg) - K + 1)
DateElp_X = dateElp(Fct, Nb, xAMJ)
End Function


'---------------------------------------------------------------------
Function dateMask(C As Control, E As Control)
'---------------------------------------------------------------------
Dim X

errTag = Null
X = dateCtl(C.Text)
dateMask = X

If IsNumeric(X) Then
    Call dateDsp(X, C)
    C.ForeColor = txtUsr.ForeColor
    E.Visible = False
Else
    E.Caption = X
    Call elpErrMsg(C, E)
     errTag = C.Tag
End If
C.BackColor = txtUsr.BackColor

End Function

'---------------------------------------------------------
Function dateTime(ByVal D As String, ByVal T As String)
'---------------------------------------------------------

dateTime = "#" & mId$(D, 7, 2) & "-" & mId$(D, 5, 2) & "-" & mId$(D, 1, 4) & " " & mId$(T, 1, 2) & ":" & mId$(T, 3, 2) & ":" & mId$(T, 5, 2) & "#"
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

If TypeOf C Is TextBox _
Or TypeOf C Is MaskEdBox Then
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




'-------------------------------------------------------
Sub txtGetFocus(C As Control)
'-------------------------------------------------------
'''''Do While Screen.MousePointer = vbHourglass
'''''    DoEvents
'''''Loop

'If IsNull(errTag) Or errTag = C.Tag Then
    C.ForeColor = txtUsr.ForeColor
    C.BackColor = focusUsr.BackColor
'    oldText = C.Text
'End If
If TypeOf C Is MaskEdBox Then
    If Val(C.Text) = 0 Then
        C.SelStart = 0
        C.SelLength = 0
    End If
End If

End Sub

'---------------------------------------------------------------------
Function timeCtl(ByVal X As String)
'---------------------------------------------------------------------
If mId$(X, 1, 2) = "__" And mId$(X, 6, 2) = "__" Then
    timeCtl = "000000"
Else
    If mId$(X, 1, 2) = "__" Then
        Mid$(X, 1, 2) = "00"
    End If

    If mId$(X, 6, 2) = "__" Then
        Mid$(X, 6, 2) = "00"
    End If


    If Val(mId$(X, 1, 2)) > 24 Then
        timeCtl = "Heure > 24"
    Else
        If Val(mId$(X, 6, 2)) > 60 Then
            timeCtl = "Minute > 60"
        Else

            timeCtl = mId$(X, 1, 2) & mId$(X, 6, 2) & "00"
        End If
    End If
End If

End Function

'---------------------------------------------------------------------
Sub timeDsp(ByVal X As String, C As Control)
'---------------------------------------------------------------------
C.Mask = Format$("")

If X = "000000" Then
    C.Text = Space$(7)
Else
    C.Text = Format$(mId$(X, 1, 4), "@@ : @@")
End If

C.Mask = Format_Time

End Sub

'---------------------------------------------------------
Function timeImp(ByVal X As String) As String
'---------------------------------------------------------

If X = "000000" Then
    timeImp = Space$(13)
Else
    timeImp = Format$(mId$(X, 1, 6), "@@ : @@ : @@")
End If


End Function

'---------------------------------------------------------
Function timeImpHM(ByVal X As String) As String
'---------------------------------------------------------

If X = "000000" Then
    timeImpHM = Space$(13)
Else
    timeImpHM = Format$(mId$(X, 1, 4), "@@ \H @@")
End If


End Function
'---------------------------------------------------------
Function timeImp8(ByVal X As String) As String
'---------------------------------------------------------

If X = "000000" Or X = "" Then
    timeImp8 = Space$(8)
Else
    timeImp8 = Format$(mId$(X, 1, 6), "@@:@@:@@")
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
Dim K, NbD, NbE As Integer
Dim Vmax

If KeyAscii = 13 Then KeyAscii = 0: Exit Function

If KeyAscii = 8 And Len(C.Text) > 0 Then
    X = mId$(C.Text, 1, Len(C.Text) - 1)
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
        X = mId$(X, 1, Len(X) - 1)
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
    NbD = Len(X) - K
    If NbD > maxD Then
        NbD = maxD
        X = mId$(X, 1, Len(X) - 1)
        elpNum = Val(X)
        Beep
    End If
    F = F & "." & String$(NbD, "0")

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
    If TypeOf xobj Is TextBox _
    Or TypeOf xobj Is MaskEdBox Then
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
    If TypeOf xobj Is TextBox _
    Or TypeOf xobj Is MaskEdBox Then
        xobj.BackColor = txtUsr.BackColor
        xobj.ForeColor = txtUsr.ForeColor
    Else
        If TypeOf xobj Is Label Then '_
'        Or TypeOf xobj Is SSoption Then
    
            xobj.ForeColor = IIf(mId$(xobj.Name, 1, 3) = "lib", libUsr.ForeColor, lblUsr.ForeColor)
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


On Error GoTo ErrorX

App_EXEName = UCase$(Trim(App.EXEName))
App_Title = UCase$(Trim(App.Title))

lX = 25: X = Space(25)
Vx = GetUserName(X, lX)
usrIdNT = mId$(X, 1, lX - 1)

usrName = usrIdNT
usrName_UCase = UCase(usrName)

usrCompte = "": usrRacine = ""

tauxTVA = 0.196: tauxTTC = 1 + tauxTVA

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
MouseMoveUsr.BackColor = RGB(235, 255, 255)  'vbHighlight

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

dateAAmin = 1980
dateAAmax = 2080
dateSerialMin = DateSerial(1989, 12, 31)
blnTimer_Enabled = False
blnNetSend_Enabled = False
'''If Command = "" Then
    SrvDir = ""
'''Else
'    If Mid$(SrvDir, I, 1) = "\" Then
'       Mid$(SrvDir, I, 1) = ""
'    End If
'''End If

elpSrvTxtin = False
elpSrvTxtOut = False
elpSrvXcom = ""
pcIdUsrIdCtl = True
DataBase_Open = "": DataBase_Master = "": DataBase_Local = ""
socName = "Société ?"
socName = "Banque Intercontinentale Arabe (Paris)"

prtFontName = prtFontName_Arial
prtZoom = 0
prtSocSigle = True
prtShow = True
prtCollection_Index = 0

''SrvDir = "D:\BiaSrv\"
mCommand = Trim(Command)
wCommand = mCommand
blnAuto_Form_Show = True

If UCase$(mId$(mCommand, 1, 6)) = "@TIMER" Then
    If UCase$(mId$(mCommand, 7, 1)) = "_" Then blnAuto_Form_Show = False             ' Affichage de la forme en mode AUTO

    paramElpTimer_Id = mId$(mCommand, 8, Len(mCommand) - 7)
    wCommand = ""
    blnTimer_Enabled = True
End If

If wCommand = "" Then
    mainSoc_Environment
Else
   ' I = Len(X)
   ' For I1 = I To 1 Step -1
   '     If mId$(wCommand, I1, 1) = "\" Then SrvDir = mId$(X, 1, I1): Exit For
   ' Next I1
   
    mainSoc_Environment

    Open wCommand For Input As #1
    
    Do While Not EOF(1)
        
        Line Input #1, X
        I1 = InStr(1, X, Chr$(34))
        If I1 > 0 Then
            I2 = InStr(I1 + 1, X, Chr$(34))
            X2 = Trim(UCase$(mId$(X, I1 + 1, I2 - I1 - 1)))
            Select Case UCase$(mId$(X, 1, I1 - 1))
                Case "SRVOBJ="
                        Elp.SrvObj = X2
                Case "PCID="
                        Elp.pcId = X2
 '''''''''      Case "USRID=": Elp.usrId = X2
                Case "SRVID="
                        Elp.SrvId = X2
                Case "SRVTYPE="
                        Elp.SrvType = X2
                Case "SRVDTAQLIB="
                        Elp.SrvDtaqLib = X2
                Case "SRVDTAQIN="
                        Elp.SrvDtaqIn = X2
                Case "SRVDTAQOUT="
                        Elp.SrvDTaqOut = X2
                Case "PCIDUSRIDCTL="
                        If X2 = "NON" Then pcIdUsrIdCtl = False
                Case "IMGSOCSIGNON="
                        strSocSignon = X2
                Case "IMGSOCLOGO="
                        imgSocLogo = X2
                Case "IMGSOCSIGLE="
                        imgSocSigle = X2
                Case "IMGGUICHET="
                        imgGuichet = X2
                Case "PRTFONTNAME="
                        prtFontName = X2
                Case "PRTZOOM="
                        prtZoom = Val(X2)
                Case "DATABASENAME=", "DATABASE_LOCAL="
                        DataBase_Local = X2
                        'I1 = InStr(3, X2, "\")                  ' \\FR....\ ou D:\
                        'I2 = InStr(I1 + 1, X2, "\")
                        'paramFolder_Local = mId$(X2, 1, I2 - 1)
                        
                        Call fileName_Split(X2, paramFolder_Local, X3, X4)
                 Case "DATABASE_MASTER="
                        DataBase_Master = X2
               Case "SRVTXTOUT="
                        If X2 = "OUI" Then elpSrvTxtOut = True
                Case "SRVXCOM="
                        elpSrvXcom = X2
                Case "TIMER="
                        If X2 = "OUI" Then blnTimer_Enabled = True
                Case "IMGSOCSICON="
                        frmElp_Icon = X2
                Case "BLNJPL="
                    blnJPL = X2 'True: MsgBox "ElpVb4.main : blnJPL" 'False

            End Select
        End If
    Loop
    
    Close #1
End If

If blnJPL Then
    usrIdNT = "LOULERGUE": usrName = usrIdNT  '$$$$$$$$$$$$$$$$$
    paramTemp_Folder = "C:\Temp"
Else
    paramTemp_Folder = "C:\Temp"
End If

prtFontNameZ = prtFontName
Elp.SrvDTaqLen = "00000"
Elp.jplFree = "00000"
Elp.usrId = UCase$(usrName)
usrId = Elp.usrId
frmElp_Caption = "Bia"
'frmElp_Icon = "Misc34.ico"

frmElp.Show vbModeless 'vbModal
frmElp.Enabled = False

frmElp.imgSocSignon.Stretch = False
frmElp.imgSocSignon.Picture = LoadPicture(strSocSignon)
Elp_ResizeImg frmElp.imgSocSignon
Set msFileSystem = CreateObject("Scripting.FileSystemObject")
blnRéplication_Load = True

MDB_Open DataBase_Local, paramDataBase_Password

paramServer_Init

' Call Sleep(2000)

'20040830 jpl :DTAQ à supprimer $$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
If elpSrvXcom = "CAV4" Then elpSrvXcom = ""
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
If Not IsNull(SndRcv_Init) Then
    End
Else
    frmElp.Timer1 = False
    mainSoc
    frmElp.Caption = frmElp_Caption
    If Trim(frmElp_Icon) <> "" Then frmElp.Icon = LoadPicture(frmElp_Icon)

    Load frmElpPrt: MeInit I
    frmElpPrt.imgSocLogo.Picture = LoadPicture(imgSocLogo)
    frmElpPrt.imgSocSigle.Picture = LoadPicture(imgSocSigle)
    frmElpPrt.imgFiligrane.Picture = LoadPicture("")
    frmElpPrt.imgFiligrane.Tag = ""
End If

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
elpSrvXcom = "XXXX"
'20040830 jpl :DTAQ à supprimer :maintenu pour frmSAB_TAU, frmAUTOMATE, frmSAA
' frmSAB_Compta_cmdRelevéA4W_Click
'======================================================================
'If elpSrvXcom = "XXXX" Then
'    elpSrvXcom = "CAV4"
'    If Not IsNull(SndRcv_Init) Then MsgBox "ElpVb4_Main_SndRcv_Init", vbCritical: End
'End If
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

'2004.02.24 DATQ pour BIA.exe
Select Case App_EXEName
    Case "BIA", "BIACPT", "BIAODBC", "BIACD", "BIA_C", "BIACPT_C":
            elpSrvXcom = "CAV4"
            If Not IsNull(SndRcv_Init) Then MsgBox "ElpVb4_Main_SndRcv_Init", vbCritical: End
    Case Else:
            elpSrvXcom = "XXXX"

End Select

frmElpPrt.WinWord_Dir


If blnTimer_Enabled Then blnElpTimer_Auto = False: ElpTimer_Init
frmElp.Msg_Rcv "ELP"
frmElp.Enabled = True
'jpl.2000.01.26 test countTimer = 0: frmElp.Timer1.Enabled = True: frmElp.Timer1.Interval = 1
Exit Sub

ErrorX:
    MsgBox "Erreur :" & Err & " : " & Error$(Err), vbCritical, "Elp.Main : " & X
    End
End Sub

'---------------------------------------------------------
Public Function RibClé(E As String, G As String, C As String, IbanE As String) As Integer
'---------------------------------------------------------
Dim r As Currency
Dim X23 As String, X1 As String * 1
Dim I As Integer
C = UCase$(C)
I = Len(C)
Do While Not IsNumeric(C)
    If I = 0 Then Exit Do
    
    X1 = mId$(C, I, 1)
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
        r = mId$(X23, 1, 9) Mod 97
        r = (Format$(r, "00") & mId$(X23, 10, 7)) Mod 97
        RibClé = 97 - (Format$(r, "00") & mId$(X23, 17, 7)) Mod 97
    
        Call Iban_Calc("FR00" & mId$(X23, 1, 21) & Format$(RibClé, "00"), IbanE)
    End If
End If

End Function

'---------------------------------------------------------
Public Function Rib_Compte(X As String) As String
'---------------------------------------------------------
Dim X1 As String * 1, Y As String
Dim I As Integer, K As Integer, lenX As Integer

Y = "00000000000"
X = Trim(UCase$(X))
lenX = Len(X)
K = 11
For I = lenX To 1 Step -1
    X1 = mId$(X, I, 1)
    Select Case Asc(X1)
        Case 48 To 57, 65 To 90
            If K > 0 Then Mid$(Y, K, 1) = X1: K = K - 1
    End Select
Next I
Rib_Compte = Y
End Function

'---------------------------------------------------------
Public Sub meEnabled(ByVal Msg As Boolean)
'---------------------------------------------------------

For Each xobj In XForm.Controls
    If TypeOf xobj Is TextBox _
    Or TypeOf xobj Is MaskEdBox Then
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
    End If
End If
End Function



'---------------------------------------------------------
Public Sub lstErr_AddItem(E As Control, C As Control, ByVal X As String)
'---------------------------------------------------------

E.Visible = True
If Not blnBeep And mId$(X, 1, 1) = "?" Then Beep: blnBeep = True
'E.Visible = True
E.AddItem Time & " " & X
If E.ListCount < 6 Then E.Height = (E.ListCount + 1) * 200 'lstLineHeight
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
    If TypeOf xobj Is TextBox _
    Or TypeOf xobj Is MaskEdBox Then
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
time_Hms = mId$(X, 1, 2) & mId$(X, 4, 2) & mId$(X, 7, 2)

End Function

Public Sub num_KeyAscii(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) Then
    If KeyAscii <> 8 And KeyAscii <> 32 Then KeyAscii = 0
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
    If TypeOf xobj Is TextBox _
    Or TypeOf xobj Is MaskEdBox Then
        xobj.Enabled = blnValue
    End If
Next xobj
End Sub

Public Function Iban_Calc(IbanX As String, IbanE As String)
Dim r As Currency
Dim X As String, X1 As String * 1, X2 As String * 2, Y As String
Dim I As Integer, K As Integer, lenX As Integer

Iban_Calc = Null
IbanE = "": Y = ""
X = UCase$(IbanX)
lenX = Len(X)
If lenX < 5 Then Iban_Calc = "Iban : longueur < 5 caractères": Exit Function
Mid$(X, 3, 2) = "00"

For I = 1 To lenX
    X1 = mId$(X, I, 1)
    K = Asc(X1)
    Select Case K
        Case 48 To 57: IbanE = IbanE & X1: Y = Y & X1
        Case 65 To 90: X2 = Format$(K - 55, "00"): IbanE = IbanE & X1: Y = Y & X2
    End Select
Next I
lenX = Len(Y)
X = mId$(Y, 7, lenX - 6) & mId$(Y, 1, 6)
r = mId$(X, 1, 9) Mod 97
For I = 10 To lenX Step 7
    r = (Format$(r, "00") & mId$(X, I, 7)) Mod 97
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
    If mId$(IbanX, 3, 2) <> mId$(IbanE, 3, 2) Then Iban_Check = "Clé Iban erronée : " & mId$(IbanE, 3, 2)
End If
End Function

Public Function num_CDec(V As Variant) As Variant
Dim I As Integer
I = InStr(1, V, ",")
If I > 0 Then Mid$(V, I, 1) = "."
num_CDec = Val(V)

End Function

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
    If X = mId$(cbo.List(cbo.ListIndex), 1, lenX) Then Exit Sub
Next I
cbo.ListIndex = -1
End Sub

Public Sub fileListBox_Scan(X As String, fileListBox As fileListBox)
Dim I As Integer, lenX As Integer
lenX = Len(X)
fileListBox.ListIndex = -1
For I = 0 To fileListBox.ListCount - 1
    fileListBox.ListIndex = fileListBox.ListIndex + 1
    If X = mId$(fileListBox.List(fileListBox.ListIndex), 1, lenX) Then Exit Sub
Next I
fileListBox.ListIndex = -1

End Sub

Public Sub lst_Scan(X As String, lst As ListBox)
Dim I As Integer, lenX As Integer
lenX = Len(X)
lst.ListIndex = -1
For I = 0 To lst.ListCount - 1
    lst.ListIndex = lst.ListIndex + 1
    If X = mId$(lst.List(lst.ListIndex), 1, lenX) Then Exit Sub
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
        GoTo Exit_Sub
    End If
'Exit Sub

Next I
lst.ListIndex = -1

Exit_Sub:
 'lst.MultiSelect = mMultiSelect
End Sub

Public Sub cbo_Value(X As String, cbo As ComboBox)
X = mId$(cbo.List(cbo.ListIndex), 1, Len(X))
End Sub

Public Function dateFinDeMois(ByVal X As String) As String
Select Case mId$(X, 5, 2)
    Case "02":
            dateFinDeMois = mId$(X, 1, 6) & "28"
            If Val(mId$(X, 1, 4)) Mod 4 = 0 Then
                If Val(mId$(X, 1, 4)) Mod 100 <> 0 Then
                    dateFinDeMois = mId$(X, 1, 6) & "29"
                Else
                    If Val(mId$(X, 1, 4)) Mod 400 = 0 Then
                        dateFinDeMois = mId$(X, 1, 6) & "29"
                    End If
                End If
            End If
    Case "04", "06", "09", "11": dateFinDeMois = mId$(X, 1, 6) & "30"
    Case Else: dateFinDeMois = mId$(X, 1, 6) & "31"
End Select

End Function

Public Sub Elp_ResizeImg(imgX As IMAGE)
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
On Error Resume Next
If lMe.WindowState = vbMaximized Then
    frmUsr_Windowstate = vbMaximized
    wFontSize = 12
    lHeight_2 = lMe.Height: lWidth_2 = lMe.Width
    If lHeight_0 <> 0 Then
        D1 = lWidth_2 / lWidth_0
        D2 = lHeight_2 / lHeight_0
        D3 = (lHeight_2 / lHeight_0) * 0.75
        
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
        D3 = (lHeight_0 / lHeight_2) / 0.75
    Else
        Exit Sub
    End If
End If

For Each xobj In lMe.Controls
    If TypeOf xobj.Container Is Toolbar Then
    Else
        If TypeOf xobj Is Menu _
        Or TypeOf xobj Is CoolBar _
        Or TypeOf xobj Is Toolbar _
        Or TypeOf xobj Is Timer Then
        Else
            If TypeOf xobj Is TextBox _
            Or TypeOf xobj Is Label _
            Or TypeOf xobj Is ComboBox _
            Or TypeOf xobj Is fileListBox _
            Or TypeOf xobj Is ListBox _
            Or TypeOf xobj Is PictureBox _
            Or TypeOf xobj Is MSFlexGrid Then
                xobj.FontSize = wFontSize
            End If
            
            On Error Resume Next
            If xobj.Left > 0 Then xobj.Left = xobj.Left * D1
            xobj.Top = xobj.Top * D2
            xobj.Width = xobj.Width * D1
            If TypeOf xobj Is ComboBox Then
            Else
                If TypeOf xobj Is TextBox Then
                        xobj.Height = xobj.Height * D3
                Else
                        xobj.Height = xobj.Height * D2
                End If
                If TypeOf xobj Is IMAGE Then
                    Set XImage = xobj
                    Call Elp_ResizeImg(XImage)
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
    If mId$(X, I, 1) = "\" Then Exit For
Next I

X1 = "": X2 = ""
If I > 1 Then X1 = mId$(X, 1, I)
If I < L Then X2 = mId$(X, I + 1, L - I)
I = InStr(X2, ".")
If I > 0 Then
    X2 = X2 & "*"
Else
    X2 = X2 & "*.*"
End If

filDoc.PATH = X1
filDoc.Pattern = X2 & "*"

End Sub

Public Sub fileName_Split(lFileName As String, lFolder As String, lName As String, lExtension As String)
Dim K As Integer, L As Integer

lFolder = "": lName = "": lExtension = ""

L = Len(lFileName)
For K = L To 1 Step -1
    If mId$(lFileName, K, 1) = "\" Then Exit For
Next K


If K > 1 Then lFolder = mId$(lFileName, 1, K)
If K < L Then lName = mId$(lFileName, K + 1, L - K)
K = InStr(lName, ".")
If K > 0 Then
    lExtension = mId$(lName, K + 1, Len(lName) - K)
    lName = mId$(lName, 1, K - 1)
End If


End Sub
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
    If K > 1 Then X1 = mId$(lFileName, 1, K - 1)
    K = K + lX
    If K < L Then X2 = mId$(lFileName, K, L - K + 1)
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
    If AMJ > 19000101 Then
        C.Day = "01"
        C.Year = mId$(AMJ, 1, 4)
        C.Month = mId$(AMJ, 5, 2)
        C.Day = mId$(AMJ, 7, 2)
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
    xAMJ = mId$(X, 1, 8)
End If

End Function

Public Sub DTPicker_Now(C As DTPicker)
C.Year = Year(Now)
C.Day = 1
C.Month = Month(Now)
C.Day = Day(Now)

End Sub



Public Sub X_JPL()
'Private Sub txtCompte_Change()
'fgSelect.Clear
'End Sub

'Private Sub txtXXX_GotFocus()
'txt_GotFocus txtXXX
'End Sub
'
'Private Sub txtXXX_KeyPress(KeyAscii As Integer)
'KeyAscii = convUCase(KeyAscii)
'KeyAscii =ctlNum (KeyAscii)
'End Sub

'Private Sub txtXXX_LostFocus()
'txt_LostFocus txtXXX
'If blnControl Then cmdControl

'End Sub


'RESIZE $$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'======
'>Declarations
'Dim mHeight_0 As Integer, mWidth_0 As Integer, mHeight_2 As Integer, mWidth_2 As Integer, mWindowState As Integer

'>Msg_Rcv
'mWindowState = Me.WindowState
'If Me.WindowState <> frmUsr_Windowstate Then Me.WindowState = frmUsr_Windowstate


'>Form_Load
'mHeight_0 = Me.Height: mWidth_0 = Me.Width: mHeight_2 = 0: mWidth_2 = 0: mWindowState = Me.WindowState

'>Private Sub Form_Resize()

'If mWindowState <> Me.WindowState Then
'    If Me.WindowState = 0 Or Me.WindowState = 2 Then
'        Elp_Form_Resize Me, mWindowState, mHeight_0, mWidth_0, mHeight_2, mWidth_2
'    End If
'End If
'RESIZE $$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'======

'XListBox.AddItem "X_Reset     " & Chr$(9) & Chr$(9) & "réplication BiaSrv"
'Nb = Nb + 1
'If Nb >= UBound(arrBiaPgm_Name) Then ReDim arrBiaPgm_Name(Nb + 10)
'arrBiaPgm_Name(Nb) = "X_Reset     " & " : " & "réplication BiaSrv"

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


Public Sub cbo_Load(lElpTable As typeElpTable, cbo As ComboBox, lK2_len As Integer)

cbo.Clear

Dim mK1 As String * 12

mK1 = lElpTable.K1

lElpTable.Method = "Seek>="
lElpTable.Err = 0

Do
    lElpTable.Err = tableElpTable_Read(lElpTable)
    If lElpTable.Err = 0 Then
        If lElpTable.K1 <> mK1 Then
            lElpTable.Err = 9996
        Else
            cbo.AddItem mId$(lElpTable.K2, 1, lK2_len) & " " & Trim(lElpTable.Name)

            lElpTable.Method = "Seek>"
       End If
    End If
Loop While lElpTable.Err = 0
End Sub

Public Sub cbo_LoadK2(lElpTable As typeElpTable, cbo As ComboBox)
Dim mK1 As String * 12

mK1 = lElpTable.K1

lElpTable.Method = "Seek>="
lElpTable.Err = 0

Do
    lElpTable.Err = tableElpTable_Read(lElpTable)
    If lElpTable.Err = 0 Then
        If lElpTable.K1 <> mK1 Then
            lElpTable.Err = 9996
        Else
            cbo.AddItem lElpTable.K2

            lElpTable.Method = "Seek>"
       End If
    End If
Loop While lElpTable.Err = 0

End Sub
Public Sub lst_LoadK2(lElpTable As typeElpTable, lst As ListBox)
Dim mK1 As String * 12

mK1 = lElpTable.K1

lElpTable.Method = "Seek>="
lElpTable.Err = 0

Do
    lElpTable.Err = tableElpTable_Read(lElpTable)
    If lElpTable.Err = 0 Then
        If lElpTable.K1 <> mK1 Then
            lElpTable.Err = 9996
        Else
            lst.AddItem Trim(lElpTable.K2) & " " & Trim(lElpTable.Name)

            lElpTable.Method = "Seek>"
       End If
    End If
Loop While lElpTable.Err = 0

End Sub

Public Sub cbo_LoadId(lElpTable As typeElpTable, cbo As ComboBox)
Dim mId As String * 12

mId = lElpTable.ID

lElpTable.Method = "Seek>="
lElpTable.Err = 0

Do
    lElpTable.Err = tableElpTable_Read(lElpTable)
    If lElpTable.Err = 0 Then
        If lElpTable.ID <> mId Then
            lElpTable.Err = 9996
        Else
            cbo.AddItem Trim(lElpTable.K1)

            lElpTable.Method = "Seek>"
       End If
    End If
Loop While lElpTable.Err = 0

End Sub

Public Function dateJMA_AMJ(lJma As String, lAmj As String)
lAmj = ""
lAmj = mId$(lJma, 7, 4) & mId$(lJma, 4, 2) & mId$(lJma, 1, 2)

End Function

Public Function dateJMA6_AMJ(lJma As String, lAmj As String)
lAmj = ""
If Len(lJma) < 8 Then
    lAmj = "20" & mId$(lJma, 7, 2) & mId$(lJma, 4, 2) & mId$(lJma, 1, 2)
Else
    lAmj = mId$(lJma, 7, 4) & mId$(lJma, 4, 2) & mId$(lJma, 1, 2)
End If

End Function
Public Function dateAMJ8_JMA6(lAmj As String, lJma As String)
lJma = mId$(lAmj, 7, 2) & mId$(lAmj, 5, 2) & mId$(lAmj, 3, 2)

End Function

Public Function dateAMJ_JMA(lAmj As String, lJma As String)
lJma = ""
lJma = mId$(lAmj, 7, 2) & mId$(lAmj, 5, 2) & mId$(lAmj, 1, 4)

End Function

Public Function dateJma10_Amj(ljma10 As String, lAmj As String)
lAmj = "00000000"
lAmj = mId$(ljma10, 7, 4) & mId$(ljma10, 4, 2) & mId$(ljma10, 1, 2)

End Function

Public Function curMaxD(lcurX As Currency, lMaxD As String) As Currency

Select Case lMaxD
    Case 0: curMaxD = Fix(lcurX + 0.5000001)
    Case Else: curMaxD = Fix((lcurX + 0.00500001)) / 100
End Select

End Function

Public Function paramServer(lMsg) As String
Dim X As String, I1 As Integer, I2 As Integer, V

X = Trim(lMsg)
paramServer = X
If mId$(lMsg, 1, 2) = "\\" Then
    I1 = InStr(3, X, "\")
    If I1 > 0 Then
        recparamServer.K2 = mId$(X, 3, I1 - 3)
        V = dbElpTable_ReadE(recparamServer)
        If IsNull(V) Then
            If Not IsNull(recparamServer.Memo) Then
                If blnJPL Then recparamServer.Memo = paramTemp_Folder: I1 = 2
                I2 = Len(lMsg)
                If mId$(recparamServer.Memo, 2, 1) = ":" Then
                    paramServer = Trim(recparamServer.Memo) & mId$(X, I1, I2 - I1 + 1)
               Else
                    paramServer = "\\" & Trim(recparamServer.Memo) & mId$(X, I1, I2 - I1 + 1)
                End If
            End If
        End If
    End If
End If
End Function
Public Sub paramServer_Init()
recElpTable_Init recparamServer
recparamServer.Method = "Seek="
recparamServer.ID = "Server"
recparamServer.K1 = "Application"

End Sub


Public Sub MDB_Open(lDataBase_Name As String, lDataBase_PassWord As String)
If lDataBase_Name <> "" Then
    If DataBase_Open <> "" Then MDB_Close

    DataBase_Open = lDataBase_Name
    Set MDB = OpenDatabase(lDataBase_Name, False, False, lDataBase_PassWord)
    tableElpTable_Open

    MDB.Execute "delete * from ElpBuffer": tableElpBuffer_Open:
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


Public Sub MDB_CompactDataBase()
On Error GoTo Error_Handler
Dim X As String, xNew As String, xOld As String
Dim wFolder As String, wName As String, wExtension As String
    X = MsgBox("CompactDataBase : " & DataBase_Local, vbInformation + vbYesNo + vbDefaultButton2, "Elp : MDB_Local")
    If X = vbYes Then
        X = DataBase_Open
        If X <> "" Then MDB_Close
        DataBase_Open = ""
        
        Call fileName_Split(DataBase_Local, wFolder, wName, wExtension)
        xNew = wFolder & wName & "_New." & wExtension
        xOld = wFolder & wName & "_Old." & wExtension
        If Dir(xNew) <> "" Then Kill xNew
        
        DBEngine.CompactDatabase DataBase_Local, xNew, , , paramDataBase_Password
        
        If Dir(xOld) <> "" Then Kill xOld
        Name DataBase_Local As xOld
        Name xNew As DataBase_Local
        Kill xOld
        
         If X <> "" Then MDB_Open X, paramDataBase_Password
    End If
    
Exit Sub
Error_Handler:
Shell_MsgBox "ELPVB_MDB_CompactDataBase : " & Error, vbCritical, frmElp_Caption, False
End Sub

Public Sub MDB_Close()

mainSoc_Close
MDB.Close
DataBase_Open = ""
End Sub

Public Sub main_Reset()
Dim X As String, IdShell
On Error GoTo Error_Exit

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
Dim Kmin As Integer, Kmax As Integer, lenText As Integer
Kmin = lK + 1
lenText = Len(lText)
CSV_Scan = ""
For Kmax = Kmin To lenText
  If mId$(lText, Kmax, 1) = ";" Then Exit For
Next Kmax

If Kmax > Kmin Then CSV_Scan = mId$(lText, Kmin, Kmax - Kmin)
lK = Kmax

End Function
Public Function DateJMA_Scan(lText As String, lK As Integer) As String
Dim K1 As Integer, K2 As Integer, lenText As Integer
Dim JJ As Long, MM As Long, AA As Long, jma As String

K1 = lK + 1
lenText = Len(lText)
DateJMA_Scan = ""
K2 = InStr(K1, lText, "/")
If K2 > 0 Then
    JJ = Val(mId$(lText, K1, K2 - K1))
    K1 = K2 + 1
    K2 = InStr(K1, lText, "/")
    If K2 > 0 Then
        MM = Val(mId$(lText, K1, K2 - K1))
        K1 = K2 + 1
        K2 = InStr(K1, lText, " ")
        If K2 > 0 Then
            AA = Val(mId$(lText, K1, K2 - K1))
            If AA < 100 Then AA = 2000 + AA
        End If
    End If
End If
jma = Format(JJ, "00") & "-" & Format(MM, "00") & "-" & Format(AA, "0000")
If IsDate(jma) Then
    DateJMA_Scan = mId$(jma, 7, 4) & mId$(jma, 4, 2) & mId$(jma, 1, 2)
End If
lK = K2

End Function

Public Function TimeHMS_Scan(lText As String, lK As Integer) As String
Dim K1 As Integer, K2 As Integer, KScan As Integer, lenText As Integer
Dim HH As Long, MM As Long, SS As Long
K1 = lK + 1
KScan = K1
lenText = Len(lText)
TimeHMS_Scan = ""
K2 = InStr(K1, lText, ":")
If K2 > 0 Then
    KScan = K2
    HH = Val(mId$(lText, K1, K2 - K1))
    K1 = K2 + 1
    K2 = InStr(K1, lText, ":")
    If K2 > 0 Then
        KScan = K2
        MM = Val(mId$(lText, K1, K2 - K1))
        K1 = K2 + 1
        K2 = InStr(K1, lText, " ")
        If K2 > 0 Then
            KScan = K2
            SS = Val(mId$(lText, K1, K2 - K1))
        End If
    End If
End If
TimeHMS_Scan = Format(HH, "00") & Format(MM, "00") & Format(SS, "00")
lK = KScan

End Function



Public Function Space_Scan(lText As String, lK As Integer) As String
Dim Kmin As Integer, Kmax As Integer, lenText As Integer
Dim blnOk As Boolean

Kmin = lK + 1
lenText = Len(lText)
Space_Scan = ""
blnOk = False
For Kmax = Kmin To lenText
    If mId$(lText, Kmax, 1) = " " Then
        If blnOk Then Exit For
    Else
        blnOk = True
    End If
    
Next Kmax

If Kmax > Kmin Then Space_Scan = Trim(mId$(lText, Kmin, Kmax - Kmin))
lK = Kmax

End Function

Public Sub strMoveR(lX As Variant, lDest As String, lPos As Integer, lLen As Integer)
Dim lenX As Integer
lenX = Len(lX)
If lenX > lLen Then
    Mid$(lDest, lPos, lLen) = mId$(lX, lenX - lLen + 1, lLen)
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
    If mId$(lFile, K, 1) = "\" Then Exit For
Next K

X = lArchive_Folder & DSys & "_" & time_Hms & "_" & mId$(lFile, K + 1, lenX - K)
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
DSys_VeilleOAP = dateElp("Ouvré", -2, DSys)
DSys_SuivantC = dateElp("Jour", 1, DSys)
DSys_SuivantO = dateElp("Ouvré", 1, DSys)

End Sub

Public Sub num_Xprt_Long(lNum As Long, lCurrentX As Long)
Dim X As String

If lNum = 0 Then
    X = "-"
Else
    X = Format(lNum, "### ### ##0")
End If
XPrt.CurrentX = lCurrentX - XPrt.TextWidth(X)
XPrt.Print X;

End Sub
Public Sub num_Xprt_Currency(lNum As Currency, lCurrentX As Long)
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

Public Function SQL_Date_Time(lAmj As String, lHms As Long) As String
SQL_Date_Time = "{ts '" & Format$(lAmj, "@@@@-@@-@@ ") & Format$(lHms, "00:00:00") & ".000'}"
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
    Call dateJMA6_AMJ(mId$(lX, 6, 6), X8)
    lX32V = CLng(X8)
    lX32D = mId$(lX, 12, 3)
    lX32A = CCur(mId$(lX, 15, Len(lX) - 14))
End If

End Sub


Public Function htmlFontColor(lX1 As String) As String
htmlFontColor = "<font color=" & Asc34 & lX1 & Asc34 & ">"
End Function

Public Sub cnAdo_Info(cnADO As ADODB.Connection)
Dim X As String

X = cnADO.DefaultDatabase & vbCr & _
"ADO Version: " & cnADO.Version & vbCr & _
"DBMS Name: " & cnADO.Properties("DBMS Name") & vbCr & _
"DBMS Version: " & cnADO.Properties("DBMS Version") & vbCr & _
"OLE DB Version: " & cnADO.Properties("OLE DB Version") & vbCr & _
"Provider Name: " & cnADO.Properties("Provider Name") & vbCr & _
"Provider Version: " & cnADO.Properties("Provider Version") & vbCr & _
"Driver Name: " & cnADO.Properties("Driver Name") & vbCr & _
"Driver Version: " & cnADO.Properties("Driver Version") & vbCr & _
"Driver ODBC Version: " & cnADO.Properties("Driver ODBC Version")

MsgBox X

End Sub
