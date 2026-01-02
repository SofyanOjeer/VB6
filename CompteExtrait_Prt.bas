Attribute VB_Name = "prtCompteExtrait"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------

Dim recLine As String

'---------------------------------------------------------
Public Sub prtCompteExtrait_Monitor(Msg As String)
'---------------------------------------------------------
Dim K As Integer, K1 As Integer, K2 As Integer, mDicrub As Integer
Dim X
Set XPrt = Printer


'frmElpPrt.Show vbModal 'vbModeless

prtOrientation = vbPRORPortrait
prtTitleText = "Liste des extraits de compte imprimés"
prtPgmName = "prtCompteExtrait"
prtTitleUsr = usrName
'prtFontName = "Courier"
prtLineNb = 1
prtlineHeight = 250
prtHeaderHeight = 300

frmElpPrt.prtStdInit
prtMaxLine = prtMaxY - 50 - prtlineHeight

prtCompteExtrait_Form

Open Msg For Input As #1
Do Until EOF(1)
    Input #1, recLine
    prtCompteExtrait_Line
Loop
Close #1
prtFontName = prtFontNameZ
frmElpPrt.prtEndDoc

frmElpPrt.Hide

End Sub



'---------------------------------------------------------
Public Sub prtCompteExtrait_Form()
'---------------------------------------------------------
Dim X As String

XPrt.FontSize = 7
XPrt.FontBold = True

XPrt.DrawWidth = 3
XPrt.ForeColor = RGB(0, 0, 0)

XPrt.Line (prtMinX, prtMinY)-(prtMaxX, prtMaxY), , B
XPrt.Line (prtMinX, prtMinY + prtHeaderHeight)-(prtMaxX, prtMinY + prtHeaderHeight)


XPrt.DrawWidth = 1
XPrt.Line (prtMinX + 3000, prtMinY)-(prtMinX + 3000, prtMaxY)
XPrt.Line (prtMinX + 6100, prtMinY)-(prtMinX + 6100, prtMaxY)

'----------------------------------------ligne 1-----------------

XPrt.CurrentY = prtMinY + (prtHeaderHeight - XPrt.TextHeight(X)) / 2
XPrt.CurrentX = prtMinX + 50: XPrt.Print "Message";
XPrt.CurrentX = prtMinX + 1100: XPrt.Print "Devise";
XPrt.CurrentX = prtMinX + 2000: XPrt.Print "Compte";
XPrt.CurrentX = prtMinX + 3100: XPrt.Print "Solde débiteur";
XPrt.CurrentX = prtMinX + 5000: XPrt.Print "Solde créditeur";
XPrt.CurrentX = prtMinX + 6200: XPrt.Print "N° extrait";
XPrt.CurrentX = prtMinX + 7200: XPrt.Print "Date";

XPrt.CurrentY = prtMinY + prtHeaderHeight - XPrt.TextHeight("test")
XPrt.FontBold = False


End Sub


'---------------------------------------------------------
Public Sub prtCompteExtrait_Line()
'---------------------------------------------------------
Dim curX As Currency, X As String

If XPrt.CurrentY > prtMaxLine Then
    frmElpPrt.prtNewPage
    prtCompteExtrait_Form
End If
'------------------------------------------ligne 1--------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'----------------------------------ligne2

XPrt.CurrentX = prtMinX
'XPrt.Print mId$(recLine, 1, 12) = "SRVCPTUPD"
'XPrt.Print mId$(recLine, 13, 12) = "ComptaExt"
XPrt.Print mId$(recLine, 25, 10);
'XPrt.Print mId$(recLine, 34 + 1, 3) = recCptInfo.Société
'XPrt.Print mId$(recLine, 34 + 4, 3) = recCptInfo.Agence
XPrt.CurrentX = prtMinX + 1200: XPrt.Print mId$(recLine, 34 + 7, 3);
XPrt.CurrentX = prtMinX + 1800: XPrt.Print Compte_Imp(mId$(recLine, 34 + 10, 11));
X = Format$(Abs(CCur(Val(mId$(recLine, 34 + 21, 15)) / 100)), "## ### ### ### ### ##0.00")
XPrt.CurrentX = prtMinX + 4000 - XPrt.TextWidth(X)
If mId$(recLine, 34 + 36, 1) = "C" Then XPrt.CurrentX = XPrt.CurrentX + 2000
XPrt.Print X;
XPrt.CurrentX = prtMinX + 6300: XPrt.Print mId$(recLine, 34 + 37, 3);
XPrt.CurrentX = prtMinX + 7000: XPrt.Print dateImp(mId$(recLine, 34 + 40, 8));

End Sub




