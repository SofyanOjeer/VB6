Attribute VB_Name = "prtCptMvt"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim I As Integer, solde As Currency, mCurrenty As Integer, Height8_6 As Integer
Dim Line1 As Integer, Line2 As Integer, Line3 As Integer, Line4 As Integer, Line5 As Integer
Dim col1 As Integer, col2 As Integer, col3 As Integer
Dim Col4 As Integer, Col5 As Integer, Col6 As Integer, Col7 As Integer, Col8 As Integer
Dim Col As Integer
Dim K As Integer, K1 As Integer, K2 As Integer, K3 As Integer
Dim X As String
Dim nbLigne As Integer, NbPage As Integer
Dim NbLigneMax As Integer, NbPageMax As Integer
Dim NbImprimé As Integer

Private recCptInfo As typeCptInfo
Private recCptMvt As typeCptMvt
Dim prtListe As Boolean
Dim prtPréimprimé As Boolean
Dim valAmjMin As String, valAmjMax As String

Dim curCumulDébit As Currency, curCumulCrédit As Currency
Dim optCV_Euro As Boolean
Dim mCV As typeCV
Dim blnA4_Form As Boolean
Dim blnMsgInfo As Boolean, mMsgInfo As String, mExtraitNuméro As String

Public Sub XXX_prtCptMvtX(Msg As String)
'---------------------------------------------------------
On Error GoTo prtError


optCV_Euro = True

Set XPrt = Printer
K1 = Val(mId$(Msg, 1, 6))
K2 = Val(mId$(Msg, 7, 6))

If mId$(Msg, 13, 1) = "-" Then
    K3 = K1: K1 = K2: K2 = K3: K3 = -1
Else
    K3 = 1
End If

prtListe = IIf(mId$(Msg, 14, 1) = "L", True, False)
prtPréimprimé = IIf(mId$(Msg, 14, 1) = "P", True, False)
optCV_Euro = IIf(mId$(Msg, 15, 1) = "E", True, False)

valAmjMin = mId$(Msg, 16, 8)
valAmjMax = mId$(Msg, 24, 8)
If valAmjMin = "00000000" Then valAmjMin = arrCptMvt(K1).AmjTraitement
If valAmjMax = "00000000" Then valAmjMax = arrCptMvt(K2).AmjTraitement

If Not IsNull(srvCompteCptInfo) Then
    MsgBox "erreur lecture CptInfo", vbCritical, "module : prtCptMvtX"
    Exit Sub
End If

frmElpPrt.Show vbModeless
recCptInfo = arrCptInfo(0)

CV_Init mCV
mCV.DeviseN = recCptInfo.Devise
CV_AttributN mCV

prtTitleText = "Extrait de Compte"
prtPgmName = "prtCptMvt"
prtTitleUsr = usrName

prtLineNb = 1
prtlineHeight = 250


solde = arrCptMvt(K1).SoldeVeille
nbLigne = 0: NbPage = 1
curCumulDébit = 0: curCumulCrédit = 0

If prtListe Then
    prtHeaderHeight = 900
    prtOrientation = vbPRORLandscape
    frmElpPrt.prtStdInit
    NbLigneMax = 36
    prtListeX
Else
    prtHeaderHeight = 300
    prtOrientation = vbPRORPortrait
    
    If prtPréimprimé Then
        prtFormType = ""
        frmElpPrt.prtInit
        NbLigneMax = 30
        prtExtrait
    Else
        prtSocInit
        NbLigneMax = 35
        prtA4
    End If
End If
Term:
frmElpPrt.prtEndDoc
frmElpPrt.Hide
Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub
Public Sub prtCptMvt_Monitor(Msg As String)
'---------------------------------------------------------
prtCptMvt_Open Msg

If Not IsNull(srvCompteCptInfo) Then
    MsgBox "erreur lecture CptInfo", vbCritical, "module : prtCptMvtX"
    Exit Sub
End If

prtCptMvt_Extrait Msg, arrCptInfo(0), "", ""

Term:
prtCptMvt_Close

Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub

Public Sub prtCptMvt_Open(Msg As String)
'---------------------------------------------------------
On Error GoTo prtError


optCV_Euro = True

Set XPrt = Printer
frmElpPrt.Show vbModeless
Height8_6 = frmElpPrt.prtHeightDelta(8, 6)

CV_Init mCV

prtTitleText = "Extrait de Compte"
prtPgmName = "prtCptMvt"
prtTitleUsr = usrName

prtLineNb = 1
prtlineHeight = 250

prtListe = IIf(mId$(Msg, 14, 1) = "L", True, False)
prtPréimprimé = IIf(mId$(Msg, 14, 1) = "P", True, False)
optCV_Euro = IIf(mId$(Msg, 15, 1) = "E", True, False)
If prtListe Then
    prtHeaderHeight = 900
    prtOrientation = vbPRORLandscape
    frmElpPrt.prtStdInit
    NbLigneMax = 36
Else
    prtHeaderHeight = 300
    prtOrientation = vbPRORPortrait
    
    If prtPréimprimé Then
        prtFormType = ""
        frmElpPrt.prtInit
        NbLigneMax = 30
    Else
        prtSocInit
        NbLigneMax = 35
    End If
End If

blnA4_Form = False
blnMsgInfo = False

Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub

Public Sub prtCptMvt_Close()
'---------------------------------------------------------
On Error GoTo prtError

frmElpPrt.prtEndDoc
frmElpPrt.Hide
Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub



Public Sub prtCptMvt_X2(Msg As String)

arrCompteIndex = 1: arrCompteNb = 1
arrCompte(1).Société = arrCptMvt(1).Société
arrCompte(1).Agence = arrCptMvt(1).Agence
arrCompte(1).Devise = arrCptMvt(1).Devise
arrCompte(1).Numéro = arrCptMvt(1).Compte
arrCompte(1).BiaTyp = "000"
arrCompte(1).BiaNum = "00"
arrCompte(1).NuméroAncien = "00000"

prtCptMvt_Monitor Msg
End Sub
'---------------------------------------------------------
Public Sub prtA4Form(Msg As String)
'---------------------------------------------------------
Dim X As String
Dim mCurrenty

prtA4Rib
If optCV_Euro Then Call frmElpPrt.prtTrame(Col6, Line3, Col7, Line4, " ", 250)

Call frmElpPrt.prtTrame(Col4, Line3, Col5, Line4, " ", 250)
Call frmElpPrt.prtTrame(col1, Line2, Col8, Line3, " ", 240)
XPrt.CurrentY = prtMinY + prtlineHeight * 4

If optCV_Euro Then prtA4Form_Euro

XPrt.DrawWidth = 3
XPrt.Line (Col4 + 200, Line1)-(Col6 - 200, Line1)
XPrt.DrawWidth = 2
XPrt.Line (col1 + 200, Line2)-(Col8, Line2)
XPrt.Line (col1, Line3)-(Col8, Line3)
XPrt.Line (col1 + 200, Line4)-(Col8, Line4)
XPrt.DrawWidth = 3
XPrt.Line (Col4 + 200, Line5)-(Col6 - 200, Line5)
XPrt.DrawWidth = 2
XPrt.Line (col1, Line2 + 200)-(col1, Line4 - 200)
XPrt.DrawWidth = 1
XPrt.Line (col2, Line2)-(col2, Line4)
XPrt.DrawWidth = 1
XPrt.Line (col3, Line2)-(col3, Line4)
XPrt.DrawWidth = 3
XPrt.Line (Col4, Line1 + 200)-(Col4, Line5 - 200)
XPrt.DrawWidth = 1
XPrt.Line (Col5, Line1)-(Col5, Line5)
XPrt.DrawWidth = 3
XPrt.Line (Col6, Line1 + 200)-(Col6, Line5 - 200)

XPrt.CurrentY = Line2 + 50
XPrt.FontBold = True

XPrt.FontSize = prtFontSize
frmElpPrt.prtCentré (col1 + col2) / 2, "Date"
frmElpPrt.prtCentré (col2 + col3) / 2, "Libellé"
frmElpPrt.prtCentré (col3 + Col4) / 2, "Date Valeur"
frmElpPrt.prtCentré (Col4 + Col5) / 2, "Débit"
frmElpPrt.prtCentré (Col5 + Col6) / 2, "Crédit"
If optCV_Euro Then
    frmElpPrt.prtCentré (Col6 + Col7) / 2, "Débit"
    frmElpPrt.prtCentré (Col7 + Col8) / 2, "Crédit"
End If
'XPrt.CurrentX = 11150
'XPrt.Print "---";

'------------------------
XPrt.DrawWidth = 2

XPrt.CurrentY = 0
XPrt.CurrentX = 0
XPrt.Circle Step(col1 + 200, Line2 + 200), 200, 0, 0.5 * Pi, Pi
XPrt.DrawWidth = 3

XPrt.CurrentY = 0
XPrt.CurrentX = 0
XPrt.Circle Step(Col6 - 200, Line1 + 200), 200, 0, 0, 0.5 * Pi

XPrt.CurrentY = 0
XPrt.CurrentX = 0
XPrt.Circle Step(Col4 + 200, Line1 + 200), 200, 0, 0.5 * Pi, Pi

XPrt.DrawWidth = 2
XPrt.CurrentY = 0
XPrt.CurrentX = 0
XPrt.Circle Step(col1 + 200, Line4 - 200), 200, 0, Pi, 1.5 * Pi

XPrt.DrawWidth = 3
XPrt.CurrentY = 0
XPrt.CurrentX = 0
XPrt.Circle Step(Col4 + 200, Line5 - 200), 200, 0, Pi, 1.5 * Pi



XPrt.CurrentY = 0
XPrt.CurrentX = 0
XPrt.Circle Step(Col6 - 200, Line5 - 200), 200, 0, 1.5 * Pi, 2 * Pi

'----------------------------------------ligne 1-----------------
XPrt.FontSize = 10
XPrt.CurrentY = prtMinY + prtlineHeight * 10 - XPrt.TextHeight("test")
'----------------------------------1------------
XPrt.FontBold = True

XPrt.CurrentX = 5800
XPrt.Print recCptInfo.Intitulé;
XPrt.FontBold = False
'-----------------------------------2-------------
If Trim(recCptInfo.Adresse2) <> "" Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = 5800
    XPrt.Print recCptInfo.Adresse2;
End If
'------------------------------------3---------------
If Trim(recCptInfo.Adresse3) <> "" Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = 5800
    XPrt.Print recCptInfo.Adresse3;
End If
'----------------------------------4-------------------
If Trim(recCptInfo.Adresse4) <> "" Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = 5800
    XPrt.Print recCptInfo.Adresse4;
End If

'-----------------------------------5------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 5800
XPrt.Print recCptInfo.Adresse5;
'------------------------------------6------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 5800
If Trim(recCptInfo.AdresseCP) <> "" Then XPrt.Print recCptInfo.AdresseCP & "  ";
XPrt.Print recCptInfo.AdresseBD;
'------------------------------------8------------------
 XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 5800
XPrt.Print recCptInfo.AdressePays;

XPrt.FontSize = 6
XPrt.CurrentX = Col8 - 350
XPrt.Print "  G " & recCptInfo.Gestionnaire & "-" & recCptInfo.Courrier;

XPrt.FontSize = 8

XPrt.CurrentY = Line1 - prtlineHeight * 3 + 50
XPrt.FontSize = 10
XPrt.FontBold = True

X = "RELEVE DE COMPTE" & "   " & mExtraitNuméro
Col = Col4 + (Col8 - Col4 - XPrt.TextWidth(X)) / 2
Call frmElpPrt.prtTrame(Col, XPrt.CurrentY, Col + XPrt.TextWidth(X) + 100, XPrt.CurrentY + prtlineHeight, " ", 240)
XPrt.CurrentX = Col + 50
XPrt.Print X;
'''frmElpPrt.prtCentré (Col8 + Col4) / 2, X
XPrt.CurrentY = XPrt.CurrentY + Height8_6
XPrt.FontBold = False
XPrt.FontSize = 6
XPrt.Print "  " & Format$(NbPage, "###") & " / " & Format$(NbPageMax, "###");
XPrt.CurrentY = XPrt.CurrentY - Height8_6
XPrt.FontSize = 8
'----------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
'XPrt.CurrentX = 800
'XPrt.Print "Numéro ";
XPrt.CurrentX = col1 + 50
'XPrt.Print ": ";
XPrt.FontBold = True
'XPrt.Print Format$(recCptInfo.Numéro, "@@@@@.@@@.@@.@") ;
XPrt.Print recCptInfo.Intitulé2;
'-------------------------------------------------------
XPrt.FontBold = False
Call DevX(recCptInfo.Devise)
XPrt.FontSize = 5
frmElpPrt.prtCentré (Col4 + Col6) / 2, Trim(XDevise.DevLib)
If optCV_Euro Then frmElpPrt.prtCentré (Col6 + Col8) / 2, "Contre-valeur en Euros"

XPrt.FontSize = 8
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'XPrt.CurrentX = 800
'XPrt.Print "Type";
XPrt.CurrentX = col1 + 50
'XPrt.Print ": ";
XPrt.FontBold = True
XPrt.Print Trim(DicLib(13, recCptInfo.BiaTyp)) & "-" & Trim(XDevise.DevLib);

'---------------------------------------
'XPrt.FontBold = False

'XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'XPrt.CurrentX = 400
'XPrt.Print "Devise";
'XPrt.CurrentX = Col2
'XPrt.Print ": ";
'XPrt.FontBold = True
'XPrt.Print Format$(recCptInfo.Devise, "000") & "-" & XDevise.DevLib;

'------------------------------------9--------------
'---------------------------------------
XPrt.FontBold = False


XPrt.FontSize = prtFontSize

XPrt.CurrentY = Line1 + 50
XPrt.CurrentX = Col4 - 100 - XPrt.TextWidth(Msg)
XPrt.Print Msg;
prtCptMvtMt (solde)
If optCV_Euro Then prtCptMvtMt_Euro (solde)

XPrt.CurrentY = Line3 - prtlineHeight + 50

End Sub

'---------------------------------------------------------
Public Sub prtA4Line()
'---------------------------------------------------------
Dim X As String, I As Integer, libCV As String, blnCV As Boolean

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
 
XPrt.FontSize = prtFontSize
XPrt.FontBold = False


XPrt.CurrentX = col1 + 50
XPrt.Print dateImp(recCptMvt.AmjTraitement);

libCV = "": blnCV = False

Select Case recCptMvt.Devise
    Case 978: If mId$(recCptMvt.Libellé, 47, 4) = " FRF" Then blnCV = True: libCV = Trim(mId$(recCptMvt.Libellé, 30, 21)): Mid$(recCptMvt.Libellé, 30, 21) = Space$(23)
    Case 1: If mId$(recCptMvt.Libellé, 47, 4) = " EUR" Then blnCV = True: libCV = Trim(mId$(recCptMvt.Libellé, 30, 21)): Mid$(recCptMvt.Libellé, 30, 21) = Space$(23)
End Select

XPrt.CurrentX = col2 + 50
For I = prtFontSize To 4 Step -1
    XPrt.FontSize = I
    If XPrt.TextWidth(Trim(recCptMvt.Libellé)) <= (col3 - col2 - 100) Then Exit For
Next I

XPrt.Print recCptMvt.Libellé;
If blnCV Then
    XPrt.FontSize = 6: XPrt.FontItalic = True
    
    XPrt.CurrentX = col3 - 50 - XPrt.TextWidth(libCV)
    XPrt.Print libCV;
    XPrt.FontItalic = False
     
End If

XPrt.FontSize = prtFontSize
'XPrt.Print StrConv(recCptMvt.Libellé, vbProperCase);
XPrt.CurrentX = col3 + 50
XPrt.Print dateImp(recCptMvt.AmjValeur);
If recCptMvt.CptComplémentaire = "3" Or recCptMvt.CptComplémentaire = "4" Then XPrt.Print " ***";
prtCptMvtMt (recCptMvt.MT)
If optCV_Euro Then prtCptMvtMt_Euro (recCptMvt.MT)
End Sub


'---------------------------------------------------------
Public Sub prtCptMvtMt(MT As Currency)
'---------------------------------------------------------
Dim X As String

XPrt.FontBold = True
X = Format$(Abs(MT), "## ### ### ### ### ##0.00")
XPrt.CurrentX = IIf(MT < 0, Col5, Col6) - 100 - XPrt.TextWidth(X)
XPrt.Print X;

End Sub

'---------------------------------------------------------
Public Sub prtCptMvtMt_Euro(MT As Currency)
'---------------------------------------------------------
Dim X As String

CV_X1.Montant = Abs(MT)
Call CV_Calc(CV_X1, CV_X2, CV_X3)
XPrt.FontBold = True
X = Format$(Abs(CV_X2.Montant), "## ### ### ### ### ##0.00")
XPrt.CurrentX = IIf(MT < 0, Col7, Col8) - 100 - XPrt.TextWidth(X)
XPrt.Print X;

End Sub

'---------------------------------------------------------
Public Sub prtCptMvtMt_cvFRF(MT As Currency)
'---------------------------------------------------------
Dim X As String
CV_X1.Montant = Abs(MT)
CV_X2.DeviseIso = "FRF": CV_X2.DeviseN = 0
Call CV_Transitoire(CV_X1, CV_X2, CV_X3, X)

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
Call frmElpPrt.prtTrame(col1, XPrt.CurrentY, Col8, XPrt.CurrentY + prtlineHeight - 10, " ", 225)

XPrt.FontBold = False
XPrt.CurrentX = col2 + 50
XPrt.Print "Cours de conversion : 1 " & Trim(CV_X3.DeviseLibellé) & " = " & Format$(CV_X2.Cours, "##.##### ") & Trim(CV_X2.DeviseLibellé);

XPrt.FontBold = True
X = Format$(Abs(CV_X2.Montant), "## ### ### ### ### ##0.00")
XPrt.CurrentX = IIf(MT < 0, Col5, Col6) - 100 - XPrt.TextWidth(X)
XPrt.Print X & " " & CV_X2.DeviseIso;

End Sub

'---------------------------------------------------------
Public Sub prtA4_Médiateur()
'---------------------------------------------------------
Dim X As String
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight + 50
Call frmElpPrt.prtTrame(col1, XPrt.CurrentY, Col8, XPrt.CurrentY + prtlineHeight * 2.3, " ", 225)

XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + 100
XPrt.CurrentX = col1 + 200
XPrt.Print "Nous vous informons qu'un médiateur est à votre disposition à l'adresse suivante : ";
XPrt.FontBold = True
XPrt.Print "  M. le MEDIATEUR   -   B.P. 151   -   75422 PARIS CEDEX 09";

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = col2 + 50
XPrt.FontBold = False
frmElpPrt.prtCentré prtMedX, "pour tout problème que vous n'avez pu résoudre préalablement avec la banque."



End Sub

'---------------------------------------------------------
Public Sub prtExtrait()
'---------------------------------------------------------
Dim CTLAMJ As String * 8

NbPageMax = Fix((Abs(K2 - K1)) / NbLigneMax) + 1
prtExtraitForm "Solde au : " & dateImp(dateElp("Jour", -1, valAmjMin))
CTLAMJ = arrCptMvt(K1).AmjTraitement

For K = K1 To K2 Step K3
If nbLigne = NbLigneMax Then
    XPrt.CurrentY = Line4 + 50
    XPrt.CurrentX = 4600
'    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 32
    XPrt.Print "                    Report";
    prtCptMvtMt (solde)
    nbLigne = 0: NbPage = NbPage + 1
    frmElpPrt.prtNewPage
    prtExtraitForm "                    Report"
End If

nbLigne = nbLigne + 1
recCptMvt = arrCptMvt(K)
prtExtraitLine

    If CTLAMJ <> recCptMvt.AmjTraitement Then
        If recCptMvt.AmjTraitement <> "00000000" _
        And recCptMvt.AmjTraitement <> DSys Then
            If solde <> recCptMvt.SoldeVeille Then
                XPrt.CurrentX = col2
                MsgBox "erreur Solde .........", vbCritical, "prtCptMvt"
                XPrt.FontSize = 14
                XPrt.Print "ERREUR SOLDE ............."
                Exit For
            End If
        End If
        CTLAMJ = recCptMvt.AmjTraitement
    End If

solde = solde + arrCptMvt(K).MT

DoEvents: If prtKillDoc Then frmElpPrt.Hide: Exit Sub

Next K
XPrt.CurrentY = Line4 + 50
XPrt.CurrentX = 5000
XPrt.Print "Solde au : " & dateImp(valAmjMax);
prtCptMvtMt (solde)

End Sub

'---------------------------------------------------------
Public Sub prtListeX()
'---------------------------------------------------------
Dim CTLAMJ As String * 8

Col4 = 7000: Col5 = 8700: Col6 = 10300
prtListeForm "Solde au : " & dateImp(dateElp("Jour", -1, valAmjMin))
CTLAMJ = arrCptMvt(K1).AmjTraitement

For K = K1 To K2 Step K3
If nbLigne = NbLigneMax Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight + 50
    XPrt.DrawWidth = 1
    Call frmElpPrt.prtTrame(Col4, XPrt.CurrentY - 40, Col6, XPrt.CurrentY + prtlineHeight, "B")
    Call prtCptMvtMt(solde)
    nbLigne = 0
    XPrt.DrawWidth = 1
    XPrt.Line (Col4, prtMinY)-(Col4, prtMaxY)
    XPrt.Line (prtMinX + 10100, prtMinY)-(prtMinX + 10100, prtMaxY)
    
    frmElpPrt.prtNewPage
    prtListeForm "report"
End If

nbLigne = nbLigne + 1
recCptMvt = arrCptMvt(K)
NbImprimé = NbImprimé + 1
If NbImprimé > 3 Then
''    Call frmElpPrt.prtTrame(prtMinX + 20, XPrt.CurrentY - prtlineHeight * 2 - 50, prtMaxX - 20, XPrt.CurrentY + prtlineHeight - 50, " ")
    Call frmElpPrt.prtTrame(prtMinX + 20, XPrt.CurrentY + prtlineHeight - 50, prtMaxX - 20, XPrt.CurrentY + prtlineHeight * 2 - 30, " ", 250)
    If NbImprimé = 6 Then NbImprimé = 0
End If
prtListeLine

    If CTLAMJ <> recCptMvt.AmjTraitement Then
        If recCptMvt.AmjTraitement <> "00000000" _
        And recCptMvt.AmjTraitement <> DSys Then
            If solde <> recCptMvt.SoldeVeille Then
                XPrt.CurrentX = col2
                MsgBox "erreur Solde .........", vbCritical, "prtCptMvt"
                XPrt.FontSize = 14
                XPrt.Print "ERREUR SOLDE ............."
                Exit For
            End If
        End If
        CTLAMJ = recCptMvt.AmjTraitement
    End If

solde = solde + arrCptMvt(K).MT
X = Format$(solde, "## ### ### ### ### ##0.00")
XPrt.CurrentX = 12500 - XPrt.TextWidth(X)
XPrt.Print X;

If arrCptMvt(K).MT < 0 Then
    curCumulDébit = curCumulDébit + arrCptMvt(K).MT
Else
    curCumulCrédit = curCumulCrédit + arrCptMvt(K).MT
End If


DoEvents: If prtKillDoc Then frmElpPrt.Hide: Exit Sub

Next K

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight + 50
XPrt.DrawWidth = 1
Call frmElpPrt.prtTrame(Col4, XPrt.CurrentY - 40, Col6, XPrt.CurrentY + prtlineHeight, "B")
If curCumulDébit <> 0 Then Call prtCptMvtMt(curCumulDébit)
If curCumulCrédit <> 0 Then Call prtCptMvtMt(curCumulCrédit)


XPrt.CurrentY = XPrt.CurrentY + prtlineHeight + 50
XPrt.DrawWidth = 1
Call frmElpPrt.prtTrame(Col4, XPrt.CurrentY - 40, Col6, XPrt.CurrentY + prtlineHeight, "B")
Call prtCptMvtMt(solde)
X = "Solde au : " & dateImp(valAmjMax)
XPrt.CurrentX = Col4 - 50 - XPrt.TextWidth(X)
XPrt.Print X;

XPrt.DrawWidth = 1
XPrt.Line (Col4, prtMinY)-(Col4, prtMaxY)
XPrt.Line (prtMinX + 10100, prtMinY)-(prtMinX + 10100, prtMaxY)
XPrt.Line (12600, prtMinY + prtHeaderHeight)-(12600, prtMaxY)

End Sub

'---------------------------------------------------------
Public Sub prtListeLine()
'---------------------------------------------------------
Dim X As String

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight

XPrt.FontSize = 8
prtCptMvtMt (recCptMvt.MT)

XPrt.FontBold = False
 
XPrt.CurrentX = prtMinX + 100
XPrt.Print dateImp(recCptMvt.AmjTraitement);

If recCptMvt.AmjValeur <> recCptMvt.AmjTraitement Then
    XPrt.CurrentX = prtMinX + 1325 + 100
    XPrt.Print dateImp(recCptMvt.AmjValeur);
End If


XPrt.CurrentX = prtMinX + 2750
XPrt.Print StrConv(recCptMvt.Libellé, vbProperCase);


XPrt.CurrentX = 12700 '10400
XPrt.Print DicLib(4, recCptMvt.Service);

'XPrt.CurrentX = 12700
'XPrt.Print DicLib(27, recCptMvt.CodeOpération);

XPrt.CurrentX = 14800
XPrt.Print recCptMvt.Pièce;

If recCptMvt.Automatique = "*" Then
   'XPrt.CurrentX = 15200
    XPrt.Print "*";
End If

If recCptMvt.Exonéré = "1" Then
   'XPrt.CurrentX = 15400
    XPrt.Print "E";
End If


If recCptMvt.EditionAvis = "1" Then
   'XPrt.CurrentX = 15600
  XPrt.Print "A";
End If

End Sub

'---------------------------------------------------------
Public Sub prtListeForm(Msg As String)
'---------------------------------------------------------
 
 NbImprimé = 0
 XPrt.DrawWidth = 1
Call frmElpPrt.prtTrame(Col4, prtMinY + 10, Col6, prtMinY + 10 + prtHeaderHeight, "B")

XPrt.FontSize = 8
XPrt.FontBold = True
XPrt.DrawWidth = 3

Call frmElpPrt.prtTrame(prtMinX, prtMinY + prtHeaderHeight, prtMaxX, prtMinY + prtHeaderHeight + prtlineHeight, "B")

'XPrt.Line (prtMinX, prtMinY)-(prtMinX, prtMaxY)
'XPrt.Line (prtMaxX, prtMinY)-(prtMaxX, prtMaxY)

'---------------------------------------------------------

XPrt.CurrentY = prtMinY + 50 + (prtlineHeight - XPrt.TextHeight("toto")) / 2

XPrt.CurrentX = 300
XPrt.FontBold = True
XPrt.Print Format$(recCptInfo.Devise, "000") & ".";
XPrt.Print Compte_Imp(recCptInfo.Numéro);

XPrt.CurrentX = 2000
XPrt.Print Trim(DicLib(13, recCptInfo.BiaTyp)) & "  " & mCV.DeviseLibellé;

XPrt.FontBold = False
XPrt.CurrentX = 13500
XPrt.Print "Automatique";
XPrt.CurrentX = 15000
XPrt.Print ": *";
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'--------------------------------------------------
XPrt.FontBold = True

XPrt.CurrentX = 300
XPrt.Print recCptInfo.Intitulé;

XPrt.FontBold = False
XPrt.CurrentX = 13500
XPrt.Print "Exonéré";
XPrt.CurrentX = 15000
XPrt.Print ": E";
'--------------------------------------------------
XPrt.FontBold = True
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight

XPrt.CurrentX = Col4 - 50 - XPrt.TextWidth(Msg)
XPrt.Print Msg;

XPrt.CurrentX = IIf(solde < 0, Col5, Col5 - 300)
XPrt.Print mCV.DeviseIso;
Call prtCptMvtMt(solde)

XPrt.CurrentX = 300
XPrt.Print recCptInfo.Intitulé2;

XPrt.FontBold = False
XPrt.CurrentX = 13500
XPrt.Print "Avis";
XPrt.CurrentX = 15000
XPrt.Print ": A";

XPrt.FontBold = True
XPrt.CurrentY = prtMinY + prtHeaderHeight + (prtlineHeight - XPrt.TextHeight(X)) / 2
XPrt.CurrentX = prtMinX + 300
X = "Date Traitement"
XPrt.CurrentX = 300
XPrt.Print X;

X = "Date de Valeur"
XPrt.CurrentX = 1600
XPrt.Print X;

X = "Libellé"
XPrt.CurrentX = 3000
XPrt.Print X;

X = "Débit"
XPrt.CurrentX = 8200
XPrt.Print X;

X = "Crédit"
XPrt.CurrentX = 9600
XPrt.Print X;

XPrt.CurrentX = 12000
XPrt.Print "Solde";

''X = "Opération"
X = "Service"
XPrt.CurrentX = 12700
XPrt.Print X;

X = "Pièce"
XPrt.CurrentX = 14700
XPrt.Print X;

XPrt.CurrentY = XPrt.CurrentY + 50
End Sub

Public Sub prtListeBox()

mCurrenty = XPrt.CurrentY

XPrt.DrawWidth = 3
XPrt.FillStyle = 0
XPrt.ForeColor = RGB(0, 0, 0)
XPrt.FillColor = RGB(250, 250, 250)
XPrt.Line (Col4, XPrt.CurrentY - 20)-(Col6, XPrt.CurrentY + prtlineHeight), , B

XPrt.FillStyle = 1
XPrt.CurrentY = mCurrenty

End Sub

Public Sub prtExtraitLine()
'---------------------------------------------------------
Dim X As String

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
 
XPrt.FontSize = 9
XPrt.FontBold = False

XPrt.CurrentX = 0

X = recCptMvt.AmjTraitement
XPrt.Print Format$(mId$(X, 7, 2) & mId$(X, 5, 2) & mId$(X, 1, 4), "@@-@@-@@@@");

XPrt.CurrentX = prtMinX + 1000
XPrt.Print mId$(recCptMvt.Libellé, 1, 25);
XPrt.CurrentX = prtMinX + 4600
X = recCptMvt.AmjValeur
XPrt.Print Format$(mId$(X, 7, 2) & mId$(X, 5, 2) & mId$(X, 1, 4), "@@-@@-@@@@");


prtCptMvtMt (recCptMvt.MT)

End Sub

Public Sub prtExtraitForm(Msg As String)
'---------------------------------------------------------
Dim X As String
Dim mCurrenty
'---------------------------------------------------------
Dim W As Integer, H As Integer
Dim SW As Single, SH As Single, SX As Single
Dim L As Integer

SH = 300 / frmElpPrt.imgSocSigle.Height

W = SH * frmElpPrt.imgSocSigle.Width
H = SH * frmElpPrt.imgSocSigle.Height

XPrt.PaintPicture frmElpPrt.imgSocSigle.Picture _
                , prtMinX + (prtMaxX - prtMinX - W) / 2 _
                , prtMaxY - 500 _
                , W, H

    

Line1 = prtlineHeight * 25 + 100

Line2 = Line1 + prtlineHeight * 2 - 50
Line3 = Line2 + prtlineHeight + 50
Line4 = Line3 + prtlineHeight * NbLigneMax - 50
Line5 = Line4 - prtlineHeight + 50
col1 = prtMinX
col2 = col1 + 1325
col3 = col1 + 5725
Col4 = col1 + 6950
Col5 = col1 + 8400
Col6 = col1 + 11000

'----------------------------------1------------
XPrt.FontSize = 9
XPrt.FontBold = True

XPrt.CurrentY = prtlineHeight * 11 + 100
XPrt.CurrentX = 5200
XPrt.Print recCptInfo.Intitulé;

'-----------------------------------2-------------
If Trim(recCptInfo.Adresse2) <> "" Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = 5200
    XPrt.Print recCptInfo.Adresse2;
End If
'------------------------------------3---------------
If Trim(recCptInfo.Adresse3) <> "" Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = 5200
   XPrt.Print recCptInfo.Adresse3;
End If
'----------------------------------4-------------------
If Trim(recCptInfo.Adresse4) <> "" Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = 5200
    XPrt.Print recCptInfo.Adresse4;
End If

'-----------------------------------5------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 5200
XPrt.Print recCptInfo.Adresse5;
'------------------------------------6------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 5200
If Trim(recCptInfo.AdresseCP) <> "" Then XPrt.Print recCptInfo.AdresseCP & "  ";
XPrt.Print recCptInfo.AdresseBD;
'------------------------------------8------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 5200
XPrt.Print recCptInfo.AdressePays;

XPrt.FontBold = True
XPrt.FontSize = 7

XPrt.CurrentY = prtlineHeight * 11 + 100
XPrt.CurrentX = 600
XPrt.Print recCptInfo.Intitulé;

XPrt.FontSize = 9
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2 + 100

XPrt.CurrentX = 0
XPrt.Print strSocBdfE;
XPrt.CurrentX = 1000
'-----------------------------------------
XPrt.Print strSocBdfG;
'----------------------------------------------
XPrt.CurrentX = 2000
XPrt.Print Format$(recCptInfo.Numéro, "@@@@@.@@@.@@.@");
XPrt.CurrentX = 4100
XPrt.Print recCptInfo.CléRib;

XPrt.CurrentX = 800
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2 + 100
XPrt.Print "Siège";
XPrt.CurrentX = 800
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "01 53 76 62 62";
'-------------------------------------------------------
XPrt.CurrentX = 2450
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 7
XPrt.Print recCptInfo.LibTyp;

'---------------------------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 2450
XPrt.Print Format$(recCptInfo.Devise, "000") & "-" & XDevise.DevLib;

'------------------------------------9--------------
 XPrt.CurrentY = XPrt.CurrentY + prtlineHeight

XPrt.CurrentX = 1950
XPrt.Print recCptInfo.Intitulé2;

XPrt.CurrentY = Line1 + 50
XPrt.CurrentX = 5000
XPrt.Print Msg;
prtCptMvtMt (solde)

XPrt.CurrentY = Line3 - prtlineHeight

End Sub

Public Sub prtA4()
'---------------------------------------------------------
Dim CTLAMJ As String * 8

CV_X1.DeviseIso = ""
CV_X1.DeviseN = Format$(recCptInfo.Devise, "000")
CV_X2.DeviseIso = "EUR"
CV_X1.Montant = 10000
Call CV_Transitoire(CV_X1, CV_X2, CV_X3, X)

If optCV_Euro Then
    prtFontSize = 7
    col1 = prtMinX
    col2 = col1 + 1060
    col3 = col1 + 4820
    Col4 = col1 + 5800
    Col5 = col1 + 7100
    Col6 = col1 + 8400
    Col7 = col1 + 9650
    Col8 = col1 + 10900
   
Else
    prtFontSize = 8
    col1 = prtMinX
    col2 = col1 + 1100 '1325
    col3 = col1 + 6100 '6025
    Col4 = col1 + 7250 '6950
    Col5 = col1 + 9075 '8925
    Col6 = col1 + 10900
    Col7 = col1 + 10900
    Col8 = col1 + 10900
End If

Line1 = prtlineHeight * 21

Line2 = Line1 + prtlineHeight + 50
Line3 = Line2 + prtlineHeight + 50
Line4 = Line3 + prtlineHeight * NbLigneMax + 50
Line5 = Line4 + prtlineHeight + 50

NbPageMax = Fix((Abs(K2 - K1)) / NbLigneMax) + 1

If blnA4_Form Then frmElpPrt.prtNewPage
blnA4_Form = True
prtA4Form "Solde au : " & dateImp(dateElp("Jour", -1, valAmjMin))
CTLAMJ = arrCptMvt(K1).AmjTraitement

For K = K1 To K2 Step K3
    If nbLigne = NbLigneMax Then
        XPrt.CurrentY = Line4 + 50
        prtCptMvtMt (solde)
        nbLigne = 0: NbPage = NbPage + 1
        frmElpPrt.prtNewPage
        prtA4Form "Report"
    End If
    nbLigne = nbLigne + 1
    recCptMvt = arrCptMvt(K)
    prtA4Line
    If CTLAMJ <> recCptMvt.AmjTraitement Then
        If recCptMvt.AmjTraitement <> "00000000" _
        And recCptMvt.AmjTraitement <> DSys _
        And recCptMvt.CptComplémentaire <> "3" _
        And recCptMvt.CptComplémentaire <> "4" Then
            If solde <> recCptMvt.SoldeVeille Then
                XPrt.CurrentX = col2
                MsgBox "erreur Solde .........", vbCritical, "prtCptMvt"
                XPrt.FontSize = 14
                XPrt.Print "ERREUR SOLDE ............."
                Exit For
            End If
        End If
        CTLAMJ = recCptMvt.AmjTraitement
    End If
    
    solde = solde + arrCptMvt(K).MT
    
    DoEvents: If prtKillDoc Then frmElpPrt.Hide: Exit Sub

Next K
XPrt.CurrentY = Line4 + 50
X = "Solde au : " & dateImp(valAmjMax)
XPrt.CurrentX = Col4 - XPrt.TextWidth(X) - 200
XPrt.Print X;
XPrt.CurrentX = 5000
prtCptMvtMt (solde)
If optCV_Euro Then
    prtCptMvtMt_Euro (solde)
    
    XPrt.FontSize = 5: XPrt.FontBold = False
    XPrt.CurrentX = IIf(solde < 0, Col4, Col5) + 20
    XPrt.Print CV_X1.DeviseIso;
    XPrt.CurrentX = IIf(solde < 0, Col6, Col7) + 20
    XPrt.Print CV_X2.DeviseIso;
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
    XPrt.CurrentX = Col6
    XPrt.Print "Cours de conversion : 1 " & Trim(CV_X3.DeviseLibellé) & " = " & Format$(CV_X1.Cours, "##.##### ") & Trim(CV_X1.DeviseLibellé);
Else
'$JPL 2002.12.26    If CV_X1.DeviseIso = "EUR" Then prtCptMvtMt_cvFRF (solde)
End If

'$JPL 2002.12.26 médiateur
If recCptInfo.BiaTyp = "001" And recCptInfo.NatureTitulaire = "01" And recCptInfo.Numéro > "30000000000" Then
    prtA4_Médiateur
Else
    If blnMsgInfo Then
        XPrt.FontBold = True: XPrt.FontSize = 10
        XPrt.CurrentY = XPrt.CurrentY + prtlineHeight ''* 2
        Call frmElpPrt.prtTrame(col1, XPrt.CurrentY, Col8, XPrt.CurrentY + prtlineHeight - 10, " ", 245)
        frmElpPrt.prtCentré 5500, mMsgInfo
    End If
End If

End Sub


Public Sub prtA4Rib()
Dim iY As Integer
Dim blnRib As Boolean

'--------------------------TRAME---------------------------
Dim X As String, IbanE As String
blnRib = False

If recCptInfo.Numéro > "10000000000" And mId$(recCptInfo.Numéro, 6, 3) = "001" Then blnRib = True
 
XPrt.DrawWidth = 1
iY = 1500
Call frmElpPrt.prtTrame(200, iY, 4750, iY + 250, "", 240)

Call frmElpPrt.prtTrame(200, iY + 1450, 4750, iY + 1700, "B", 240)

'------------------------verticaux avec arrondi
XPrt.Line (200, iY + 200)-(200, iY + 2800)
XPrt.Line (1100, iY + 1450)-(1100, iY + 2100)
XPrt.Line (2000, iY + 1450)-(2000, iY + 2100)
XPrt.Line (4200, iY + 1450)-(4200, iY + 2100)
XPrt.Line (4750, iY + 200)-(4750, iY + 2800)
'------------------------horizontaux
XPrt.Line (400, iY)-(4550, iY)
XPrt.Line (200, iY + 250)-(4750, iY + 250)

XPrt.Line (200, iY + 2100)-(4750, iY + 2100)
XPrt.Line (400, iY + 3000)-(4550, iY + 3000)
'------------------------
XPrt.DrawWidth = 1

XPrt.CurrentY = 0
XPrt.CurrentX = 0
XPrt.Circle Step(200 + 200, iY + 200), 200, 0, 0.5 * Pi, Pi

XPrt.CurrentY = 0
XPrt.CurrentX = 0
XPrt.Circle Step(4750 - 200, iY + 200), 200, 0, 0, 0.5 * Pi

XPrt.CurrentY = 0
XPrt.CurrentX = 0
XPrt.Circle Step(200 + 200, iY + 3000 - 200), 200, 0, Pi, 1.5 * Pi

XPrt.CurrentY = 0
XPrt.CurrentX = 0
XPrt.Circle Step(4750 - 200, iY + 3000 - 200), 200, 0, 1.5 * Pi, 2 * Pi

XPrt.CurrentY = iY + prtlineHeight - 200
XPrt.FontSize = 8
XPrt.FontBold = True
If blnRib Then frmElpPrt.prtCentré 2500, "RELEVE D'IDENTITE BANCAIRE"
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight - 50
XPrt.FontBold = False
XPrt.FontSize = 6
If blnRib Then frmElpPrt.prtCentré 2500, "Cadre réservé au destinataire du R.I.B"
'------------------------------------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 5
XPrt.FontBold = False
XPrt.FontSize = 6
XPrt.CurrentX = 250
XPrt.Print "Code Banque";
XPrt.CurrentX = 1200
XPrt.Print "Code Guichet";
XPrt.CurrentX = 2600
XPrt.Print "Numéro de compte";
XPrt.CurrentX = 4250
XPrt.Print "clé R.I.B";
'----------------------------------------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight + 100
'XPrt.FontBold = True
XPrt.FontSize = 10
XPrt.CurrentX = 450
XPrt.Print strSocBdfE;
XPrt.CurrentX = 1350
XPrt.Print strSocBdfG;
XPrt.CurrentX = 2600
XPrt.Print Format$(recCptInfo.Numéro, "@@@@@@@@@@@");
XPrt.CurrentX = 4400
XPrt.Print Format$(recCptInfo.CléRib, "@@");
'------------------------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight + 50
XPrt.FontSize = 8
XPrt.CurrentX = 300
XPrt.Print SocRibDom;
'--------------------------------------------------------------
'XPrt.FontBold = False
XPrt.CurrentX = 3300
XPrt.Print socTéléphone;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight + 50
XPrt.CurrentX = 300
XPrt.Print "Titulaire";
XPrt.CurrentX = 900
XPrt.Print ":";
XPrt.CurrentX = 1100
XPrt.Print recCptInfo.Intitulé;
'------------------------------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight + 50
XPrt.CurrentX = 300
XPrt.Print "IBAN";
XPrt.CurrentX = 900
XPrt.Print ":";
XPrt.CurrentX = 1100
'$$$$$X = "FR00" & strSocBdfE & strSocBdfG & Format$(recCptInfo.Numéro, "@@@@@@@@@@@") & Format$(recCptInfo.CléRib, "@@")
X = "FR00" & strSocBdfE & strSocBdfG & Format$(recCptInfo.Numéro, "00000000000") & Format$(recCptInfo.CléRib, "00")
Call Iban_Calc(X, IbanE)
XPrt.Print Iban_Print(IbanE);

End Sub

Public Sub prtA4Form_Euro()

XPrt.DrawWidth = 3
XPrt.Line (Col6 + 200, Line1)-(Col8 - 200, Line1)
XPrt.DrawWidth = 3
XPrt.Line (Col6 + 200, Line5)-(Col8 - 200, Line5)
XPrt.DrawWidth = 1
XPrt.Line (Col7, Line1)-(Col7, Line5)
XPrt.DrawWidth = 3
XPrt.Line (Col8, Line1 + 200)-(Col8, Line5 - 200)

XPrt.DrawWidth = 2

XPrt.CurrentY = 0
XPrt.CurrentX = 0
XPrt.Circle Step(Col6 + 200, Line1 + 200), 200, 0, 0.5 * Pi, Pi
XPrt.DrawWidth = 3

XPrt.CurrentY = 0
XPrt.CurrentX = 0
XPrt.Circle Step(Col8 - 200, Line1 + 200), 200, 0, 0, 0.5 * Pi


XPrt.DrawWidth = 3
XPrt.CurrentY = 0
XPrt.CurrentX = 0
XPrt.Circle Step(Col6 + 200, Line5 - 200), 200, 0, Pi, 1.5 * Pi



XPrt.CurrentY = 0
XPrt.CurrentX = 0
XPrt.Circle Step(Col8 - 200, Line5 - 200), 200, 0, 1.5 * Pi, 2 * Pi

End Sub


Public Sub prtCptMvt_Extrait(Msg As String, mCptInfo As typeCptInfo, xMsgInfo As String, xExtraitNuméro As String)


mExtraitNuméro = xExtraitNuméro

If Trim(xMsgInfo) = "" Then
    blnMsgInfo = False
Else
    blnMsgInfo = True: mMsgInfo = xMsgInfo
End If

K1 = Val(mId$(Msg, 1, 6))
K2 = Val(mId$(Msg, 7, 6))

If mId$(Msg, 13, 1) = "-" Then
    K3 = K1: K1 = K2: K2 = K3: K3 = -1
Else
    K3 = 1
End If

prtListe = IIf(mId$(Msg, 14, 1) = "L", True, False)
prtPréimprimé = IIf(mId$(Msg, 14, 1) = "P", True, False)
optCV_Euro = IIf(mId$(Msg, 15, 1) = "E", True, False)

valAmjMin = mId$(Msg, 16, 8)
valAmjMax = mId$(Msg, 24, 8)
If valAmjMin = "00000000" Then valAmjMin = arrCptMvt(K1).AmjTraitement
If valAmjMax = "00000000" Then valAmjMax = arrCptMvt(K2).AmjTraitement

recCptInfo = mCptInfo

'$JPL 2002.06.21 : vérifier reccompte = reccptmvt
For K = K1 To K2 Step K3

    If recCptInfo.Numéro <> arrCptMvt(K).Compte Or recCptInfo.Devise <> arrCptMvt(K).Devise Then
        XPrt.CurrentX = 2000
        XPrt.CurrentY = (prtMinY + prtMaxY) / 2

        MsgBox "erreur COMPTE / CPTMVT .........", vbCritical, "prtCptMvt"
        XPrt.FontSize = 14
        XPrt.Print "erreur COMPTE / CPTMVT ............."
        Exit Sub
    End If
Next K
'$JPL 2002.06.21 : vérifier reccompte = reccptmvt

mCV.DeviseN = recCptInfo.Devise
CV_AttributN mCV

If optCV_Euro Then optCV_Euro = mCV.EuroIn

solde = arrCptMvt(K1).SoldeVeille
nbLigne = 0: NbPage = 1
curCumulDébit = 0: curCumulCrédit = 0

If prtListe Then
    prtListeX
Else
    If prtPréimprimé Then
        prtExtrait
    Else
        prtA4
    End If
End If

End Sub
