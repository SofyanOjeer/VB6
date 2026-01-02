Attribute VB_Name = "prtEmploi"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------

Private recEmploi As typeEmploi
Dim I As Integer, Height8_6 As Integer, nbLigne As Integer
Dim mAmjEchéance As String
'---------------------------------------------------------
 Public Sub prtEmploi_Monitor(Msg As String)
'---------------------------------------------------------
Dim K As Integer, K1 As Integer, K2 As Integer
Dim X As String
On Error GoTo prtError

Set XPrt = Printer
Height8_6 = frmElpPrt.prtHeightDelta(8, 6)
K1 = 1
K2 = arrEmploiNb

CV_X1 = CV_Euro
prtTitleText = Msg
prtLineNb = 1

frmElpPrt.Show vbModeless


prtOrientation = vbPRORLandscape
prtPgmName = "prtEmploi"
prtTitleUsr = usrName

prtlineHeight = 250
prtHeaderHeight = 300

frmElpPrt.prtStdInit
recEmploi = arrEmploi(K)
prtEmploi_Form

For K = K1 To K2
recEmploi = arrEmploi(K)
    CV_X1.DeviseN = Format$(recEmploi.Devise, "000")
    CV_AttributN CV_X1

prtEmploi_Line

DoEvents: If prtKillDoc Then frmElpPrt.Hide: Exit Sub

Next K
XPrt.DrawWidth = 5: prtCurrentY = XPrt.CurrentY
frmElpPrt.prtLineY

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
Public Sub prtEmploi_Form()
'---------------------------------------------------------
Dim X As String

XPrt.FontSize = 8
XPrt.FontBold = True
XPrt.DrawWidth = 3


Call frmElpPrt.prtTrame(prtMinX, prtMinY, prtMaxX, prtMinY + prtHeaderHeight, "B")

XPrt.DrawWidth = 2

XPrt.Line (prtMinX + 1250, prtMinY)-(prtMinX + 1250, prtMaxY)
'XPrt.Line (prtMinX + 7800, prtMinY)-(prtMinX + 7800, prtMaxY)
XPrt.Line (prtMinX + 10100, prtMinY)-(prtMinX + 10100, prtMaxY)
XPrt.Line (prtMinX + 13800, prtMinY)-(prtMinX + 13800, prtMaxY)

'---------------------------------------------------------

XPrt.CurrentY = prtMinY + (prtHeaderHeight - XPrt.TextHeight("X")) / 2
XPrt.CurrentX = prtMinX + 100
XPrt.Print "Echéance";

XPrt.CurrentX = prtMinX + 1600: XPrt.Print "Compte";
XPrt.CurrentX = prtMinX + 2950: XPrt.Print "Intitulé";
XPrt.CurrentX = prtMinX + 9300: XPrt.Print "Capital";
XPrt.CurrentX = prtMinX + 11000: XPrt.Print "Taux";
XPrt.CurrentX = prtMinX + 12200: XPrt.Print "Intérêts";
XPrt.CurrentX = prtMinX + 13300: XPrt.Print "Base";
XPrt.CurrentX = prtMinX + 14000: XPrt.Print "Départ";
XPrt.CurrentX = prtMinX + 14900: XPrt.Print "Nbj courus";

XPrt.CurrentY = prtMinY + prtHeaderHeight + 50 '+ prtlineHeight
nbLigne = 0
mAmjEchéance = recEmploi.AmjEchéance

End Sub

'---------------------------------------------------------
Public Sub prtEmploi_Line()
'---------------------------------------------------------
Dim X As String, K As Integer, wsdCurrentX As Integer
Dim Situation As String

If XPrt.CurrentY + prtParagraphHeight > prtMaxY Then
    frmElpPrt.prtNewPage
    prtEmploi_Form
End If

XPrt.FontBold = False
'------------------------------------------ligne 1--------------
If mAmjEchéance = recEmploi.AmjEchéance Then
    nbLigne = nbLigne + 1
Else
    nbLigne = 4: mAmjEchéance = recEmploi.AmjEchéance
     XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
End If

'If nbLigne = 4 Then nbLigne = 1: XPrt.CurrentY = XPrt.CurrentY + prtlineHeight


XPrt.FontSize = 6
XPrt.CurrentY = XPrt.CurrentY + Height8_6
XPrt.CurrentX = prtMinX + 2950
XPrt.Print Trim(recEmploi.Intitulé);
If Trim(recEmploi.Intitulé2) <> "" Then XPrt.Print " / " & Trim(recEmploi.Intitulé2);
XPrt.CurrentX = prtMinX + 9800
XPrt.Print CV_X1.DeviseIso;

XPrt.CurrentX = prtMinX + 1350
XPrt.Print CV_X1.DeviseN;

XPrt.FontSize = 8
XPrt.CurrentY = XPrt.CurrentY - Height8_6
XPrt.FontBold = True

XPrt.Print "." & Compte_Imp(recEmploi.Compte);

       
X = Format$(recEmploi.Capital, "#### ### ### ### ##0.00")
XPrt.CurrentX = prtMinX + 9750 - XPrt.TextWidth(X)
XPrt.Print X;

XPrt.FontBold = False
XPrt.CurrentX = prtMinX + 100
XPrt.Print dateImp(recEmploi.AmjEchéance) & recEmploi.TagEchéance;
       
If recEmploi.Taux <> 0 Then
    X = Format$(recEmploi.Taux, "###0.000000")
    XPrt.CurrentX = prtMinX + 11500 - XPrt.TextWidth(X)
    XPrt.Print X;
End If

If recEmploi.Intérêts <> 0 Then
    X = Format$(recEmploi.Intérêts, "#### ### ### ### ##0.00")
    XPrt.CurrentX = prtMinX + 13000 - XPrt.TextWidth(X)
    XPrt.Print X;
End If
       
XPrt.CurrentX = prtMinX + 13500
XPrt.Print recEmploi.NbjBase;
       
XPrt.CurrentX = prtMinX + 14000
XPrt.Print dateImp(recEmploi.AmjDépart);
       
X = Format$(recEmploi.NbjCouru, "### ### ##0")
XPrt.CurrentX = prtMaxX - XPrt.TextWidth(X)
XPrt.Print X;

    
XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtParagraphHeight

End Sub





