Attribute VB_Name = "prtCptBalance"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------

Private recCompte As typeCompte
Dim I As Integer, Height8_6 As Integer

Dim mDevCode As Integer

Private recCptInfo As typeCptInfo
Private Mt As Currency
Private totalCV As Currency, totalDev As Currency

Dim blnBalance As Boolean

Dim mAMJCours As String * 8
Public arrCptBalance() As typeCompte
'---------------------------------------------------------
 Public Sub prtCptBalance_Monitor(Msg As String)
'---------------------------------------------------------
Dim K As Integer, K1 As Integer, K2 As Integer
Dim X As String
On Error GoTo prtError

Set XPrt = Printer
Height8_6 = frmElpPrt.prtHeightDelta(8, 6)
K1 = 1   'Val(mId$(Msg, 1, 6))
K2 = UBound(arrCptBalance) - 1 ''Val(mId$(Msg, 7, 6))

CV_X1 = CV_Euro
CV_X1.CoursCompta = "C"

CV_X2 = CV_X1
CV_X3 = CV_X1

mDevCode = 0
mAMJCours = DSys
CV_X1.OpéAmj = mAMJCours
CV_X2.OpéAmj = mAMJCours
CV_X3.OpéAmj = mAMJCours

    prtTitleText = Msg
    prtLineNb = 1
    blnBalance = True
mDevCode = 0: totalDev = 0: totalCV = 0

frmElpPrt.Show vbModeless

recCompteInit recCompte

prtOrientation = vbPRORLandscape
prtPgmName = "prtCptBalance"
prtTitleUsr = usrName

prtlineHeight = 250
prtHeaderHeight = 300

frmElpPrt.prtStdInit
prtCptBalance_Form

For K = K1 To K2
recCompte = arrCptBalance(K)
    CV_X1.DeviseN = Format$(recCompte.Devise, "000")
    CV_X1.DeviseIso = ""
    CV_X1.Montant = recCompte.SoldeInstantané
    Call CV_Transitoire(CV_X1, CV_X2, CV_X3, X)

prtCptBalance_Line

DoEvents: If prtKillDoc Then frmElpPrt.Hide: Exit Sub

Next K
XPrt.DrawWidth = 5
frmElpPrt.prtLineY

K = 0
XPrt.CurrentY = XPrt.CurrentY + 50
XPrt.FontBold = True

If mDevCode > 0 Then
    K = 1
    X = Format$(totalDev, "#### ### ### ### ##0.00")
    If totalDev >= 0 Then
        XPrt.CurrentX = prtMinX + 9000 - XPrt.TextWidth(X)
    Else
        XPrt.CurrentX = prtMinX + 10750 - XPrt.TextWidth(X)
    End If

    XPrt.Print X & " ";
End If

K = IIf(totalCV < 0, 12500, 13950)
X = Format$(totalCV, "#### ### ### ### ##0.00")
XPrt.CurrentX = prtMinX + K - XPrt.TextWidth(X)
XPrt.Print X;

prtCurrentY = XPrt.CurrentY + prtlineHeight + 50
If K > 0 Then XPrt.Line (prtMinX + 7500, prtCurrentY)-(prtMinX + 14000, prtCurrentY)

frmElpPrt.prtEndDoc
frmElpPrt.Hide

Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide
End Sub
'---------------------------------------------------------
 Public Sub prtCptBalance_CptMvt(valAmjMin As String, valAmjMax As String)
'---------------------------------------------------------
Dim K As Integer, I As Integer, Msg As String * 50

For I = 1 To UBound(arrCptBalance) - 1
    Call arrCptMvt_Load(arrCptBalance(I), valAmjMin, valAmjMax)
    If arrCptMvtNb > 0 Then
        Msg = Format$(1, "000000") & Format$(arrCptMvtNb, "000000")
        Mid$(Msg, 14, 1) = "L"
        Mid$(Msg, 15, 1) = " "
        Mid$(Msg, 16, 8) = valAmjMin
        Mid$(Msg, 24, 8) = valAmjMax

        prtCptMvt_X2 Msg
    End If
Next I

End Sub

'---------------------------------------------------------
Public Sub prtCptBalance_Form()
'---------------------------------------------------------
Dim X As String

XPrt.FontSize = 8
XPrt.FontBold = True
XPrt.DrawWidth = 3


Call frmElpPrt.prtTrame(prtMinX, prtMinY, prtMaxX, prtMinY + prtHeaderHeight, "B")
Call frmElpPrt.prtTrame(prtMinX + 7500, prtMinY + prtHeaderHeight + 10, prtMinX + 9350, prtMaxY - 10, " ", 250)
Call frmElpPrt.prtTrame(prtMinX + 11100, prtMinY + prtHeaderHeight + 10, prtMinX + 12550, prtMaxY - 10, " ", 250)

XPrt.DrawWidth = 2

XPrt.Line (prtMinX + 7500, prtMinY)-(prtMinX + 7500, prtMaxY)
XPrt.Line (prtMinX + 11100, prtMinY)-(prtMinX + 11100, prtMaxY)
XPrt.Line (prtMinX + 14000, prtMinY)-(prtMinX + 14000, prtMaxY)
XPrt.Line (prtMinX + 15000, prtMinY)-(prtMinX + 15000, prtMaxY)

XPrt.Line (prtMinX + 1600, prtMinY)-(prtMinX + 1600, prtMaxY)
'---------------------------------------------------------

X = "N°de Compte"
XPrt.CurrentY = prtMinY + (prtHeaderHeight - XPrt.TextHeight(X)) / 2
XPrt.CurrentX = prtMinX + 300
XPrt.Print X;

X = "Intitulé"
XPrt.CurrentX = prtMinX + 1750
XPrt.Print X;

XPrt.CurrentX = 15300
XPrt.Print "Mvt le";


X = "Débit"
XPrt.CurrentX = 8400
XPrt.Print X;

X = "Crédit"
XPrt.CurrentX = 10250
XPrt.Print X;

XPrt.CurrentX = 12500
XPrt.Print "EUROS";
prtCurrentY = prtMinY + prtHeaderHeight

End Sub

'---------------------------------------------------------
Public Sub prtCptBalance_Line()
'---------------------------------------------------------
Dim X As String, K As Integer, wsdCurrentX As Integer, wsdCurrentX2 As Integer
Dim Situation As String

If prtCurrentY + prtParagraphHeight > prtMaxY Then
    frmElpPrt.prtNewPage
    prtCptBalance_Form
End If

XPrt.FontBold = False
'------------------------------------------ligne 1--------------

XPrt.CurrentY = prtCurrentY + prtlineHeight - XPrt.TextHeight("test")


XPrt.FontSize = 6
XPrt.CurrentY = XPrt.CurrentY + Height8_6
XPrt.CurrentX = prtMinX + 1750
XPrt.Print recCompte.Intitulé;

XPrt.CurrentX = prtMinX + 5500
If recCompte.TypeGA = "A" Then
    XPrt.Print Trim(DicLib(13, recCompte.BiaTyp));
End If
Select Case recCompte.Situation
    Case " ": Situation = ""
    Case "A": Situation = " **Annulé**"
    Case "B": Situation = " **Bloqué**"
    Case Else: Situation = " ?? " & recCompte.Situation
End Select

XPrt.Print Situation;

XPrt.CurrentX = prtMinX + 50
XPrt.Print CV_X1.DeviseN;

XPrt.FontSize = 8
XPrt.CurrentY = XPrt.CurrentY - Height8_6
XPrt.FontBold = True

XPrt.Print "." & Compte_Imp(recCompte.Numéro);
       
X = Format$(recCompte.SoldeInstantané, "#### ### ### ### ##0.00")
If recCompte.SoldeInstantané < 0 Then
    wsdCurrentX = prtMinX + 9000: wsdCurrentX2 = wsdCurrentX + 3500
Else
    wsdCurrentX = prtMinX + 10750: wsdCurrentX2 = wsdCurrentX + 3000
End If

XPrt.CurrentX = wsdCurrentX - XPrt.TextWidth(X)

K = Val(recCompte.Devise)
If K <> mDevCode Then mDevCode = IIf(mDevCode = 0, K, -1)

totalDev = totalDev + recCompte.SoldeInstantané
XPrt.Print X & " ";
XPrt.FontBold = False
XPrt.FontSize = 6
XPrt.CurrentY = XPrt.CurrentY + Height8_6
XPrt.Print CV_X1.DeviseIso;
X = dateImp(recCompte.MvtAmj)
XPrt.CurrentX = 16000 - XPrt.TextWidth(X)
XPrt.Print X;

X = Format$(CV_X1.Cours, "####0.00000")
XPrt.CurrentX = 15000 - XPrt.TextWidth(X)
XPrt.Print X;

XPrt.FontSize = 8
XPrt.CurrentY = XPrt.CurrentY - Height8_6
XPrt.FontBold = True
X = Format$(CV_X2.Montant, "#### ### ### ### ##0.00")
XPrt.CurrentX = wsdCurrentX2 - XPrt.TextWidth(X)
XPrt.Print X;
totalCV = totalCV + CV_X2.Montant
    
    
XPrt.FontBold = False

prtCurrentY = prtCurrentY + prtParagraphHeight

End Sub




