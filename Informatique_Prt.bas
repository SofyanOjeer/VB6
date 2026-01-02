Attribute VB_Name = "prtInformatique"
Option Explicit
Dim I As Integer, mCurrenty As Integer
Dim colP As Integer, colA As Integer, colV As Integer, colX As Integer

Dim V, Height8_6 As Integer
'---------------------------------------------------------
 Public Sub prtInformatique_Open()
'---------------------------------------------------------
Dim K As Integer, K1 As Integer, K2 As Integer
Dim X As String

Set XPrt = Printer
Height8_6 = frmElpPrt.prtHeightDelta(8, 6)


frmElpPrt.Show vbModeless

prtOrientation = vbPRORPortrait
prtTitleText = "Liste des comptes en devise IN "
prtPgmName = "prtInformatique"
prtTitleUsr = usrName

prtLineNb = 1
prtlineHeight = 250
prtHeaderHeight = 300

frmElpPrt.prtStdInit
colP = prtMinX
colA = prtMinX + 2000
colV = prtMinX + 9000
colX = prtMaxX
prtInformatique_Form

End Sub
'---------------------------------------------------------
 Public Sub prtInformatique_Close()
'---------------------------------------------------------
prtInformatique_Form_End
frmElpPrt.prtEndDoc
frmElpPrt.Hide

End Sub

'---------------------------------------------------------
Public Sub prtInformatique_Form()
'---------------------------------------------------------
Dim X As String, K As Integer

XPrt.FontSize = 8
XPrt.FontBold = False

XPrt.DrawWidth = 3
Call frmElpPrt.prtTrame(prtMinX, prtMinY, prtMaxX, prtMinY + prtHeaderHeight, "B")

XPrt.Line (prtMinX, prtMinY)-(prtMinX, prtMaxY)
XPrt.Line (prtMaxX, prtMinY)-(prtMaxX, prtMaxY)
K = prtMinY + prtHeaderHeight + 10

'---------------------------------------------------------
XPrt.FontBold = True
XPrt.CurrentY = prtMinY + 50
XPrt.CurrentX = prtMinX + 50
XPrt.Print "Euro                Situation";

XPrt.CurrentX = colA + 100

XPrt.Print "Devise Compte           Intitulé";
XPrt.CurrentX = colV + 1000
XPrt.Print "Solde";
XPrt.FontBold = False

XPrt.CurrentY = prtMinY + prtHeaderHeight + prtlineHeight - XPrt.TextHeight("test")

End Sub

'---------------------------------------------------------
Public Sub prtInformatique_Line(lMsg As String, lCptinfo As typeCptInfo, lDeviseIso As String)
'---------------------------------------------------------
Dim X As String, X1 As String, X2 As String

If XPrt.CurrentY + prtlineHeight > prtMaxY Then
    prtInformatique_Form_End
    frmElpPrt.prtNewPage
    prtInformatique_Form
'Else
    'frmElpPrt.prtLineY
End If

XPrt.FontBold = False
'------------------------------------------ligne 1--------------

XPrt.FontSize = 8
XPrt.CurrentX = colP + 100: XPrt.Print lMsg;
XPrt.CurrentX = colA - 200: XPrt.Print lCptinfo.Situation;
XPrt.CurrentX = colA + 100: XPrt.Print lCptinfo.Devise & " " & Compte_Imp(lCptinfo.Numéro);
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 6

X2 = Trim(lCptinfo.Intitulé2)
X1 = Trim(lCptinfo.Intitulé)
If XPrt.TextWidth(X1 & X2) < 6000 Then
    X1 = X1 & " " & X2
    X2 = ""
End If
XPrt.CurrentX = colA + 1700: XPrt.Print X1;

XPrt.CurrentX = colV - 700: XPrt.Print dateImp10(lCptinfo.AmjDernierMouvement);
XPrt.CurrentX = colX + 50: If lCptinfo.SoldeInstantané < 0 Then XPrt.Print "db";
If lCptinfo.SoldeInstantané <> 0 Then
    X = Format$(Abs(lCptinfo.SoldeInstantané), "### ### ### ### ##0.00")
    XPrt.CurrentX = colX - 100 - XPrt.TextWidth(X)
    XPrt.Print X;
End If

If X2 <> "" Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = colA + 1700: XPrt.Print X2;
End If

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight - Height8_6

End Sub









Public Sub prtInformatique_Form_End()
XPrt.DrawWidth = 2
XPrt.Line (colA, prtMinY)-(colA, prtMaxY)
XPrt.Line (colV, prtMinY)-(colV, prtMaxY)

End Sub


