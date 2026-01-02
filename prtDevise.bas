Attribute VB_Name = "prtDevise"
Option Explicit
Dim recDevise1 As typeDevise, recDevise2 As typeDevise
Private recDeviseCours As typeDeviseCours
Dim I As Integer, mCurrenty As Integer
Dim colP As Integer, colA As Integer, colV As Integer, colX As Integer
Dim blnList As Boolean
'---------------------------------------------------------
 Public Sub prtDeviseCours(Msg As String, Text As String)
'---------------------------------------------------------
Dim K As Integer, K1 As Integer, K2 As Integer
Dim X As String

Set XPrt = Printer
K1 = Val(Mid$(Msg, 1, 6))
K2 = Val(Mid$(Msg, 7, 6))


frmElpPrt.Show vbModeless

recDeviseCours_Init recDeviseCours
If Trim(Text) = "" Then
    blnList = True
    prtOrientation = vbPRORPortrait
Else
    blnList = False
    prtOrientation = vbPRORLandscape
End If
prtTitleText = "Cours de Change" & Text
prtPgmName = "prtDevise"
prtTitleUsr = usrName

prtLineNb = 1
prtlineHeight = 250
prtHeaderHeight = 600

frmElpPrt.prtStdinit
colP = prtMinX
colA = prtMinX + 4000
colV = prtMinX + 7200
colX = prtMinX + 10400
recDevise1.DevX = arrDeviseCours(1).Id1
prtDeviseCoursForm

For K = K1 To K2
recDeviseCours = arrDeviseCours(K)
If Trim(recDeviseCours.Method) <> constIgnore Then
    If recDevise1.DevX <> recDeviseCours.Id1 Then
        mCurrenty = XPrt.CurrentY
        XPrt.Line (prtMinX, mCurrenty - 50)-(prtMaxX, mCurrenty - 100)
        XPrt.CurrentY = mCurrenty
    End If
    DevCode recDeviseCours.Id1: recDevise1 = XDevise
    DevCode recDeviseCours.Id2: recDevise2 = XDevise
    prtDeviseCoursLine
End If
DoEvents: If prtKillDoc Then frmElpPrt.Hide: Exit Sub

Next K
XPrt.Line (prtMinX, XPrt.CurrentY - 50)-(prtMaxX, XPrt.CurrentY - 100)
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.FontSize = 12
XPrt.FontBold = True
frmElpPrt.prtCentré prtMaxX / 2, "Pour tout montant supérieur ou égal à  EUR 10 000 ou contrevaleur : AVISEZ la TC"
frmElpPrt.prtEndDoc
frmElpPrt.Hide

End Sub
'---------------------------------------------------------
Public Sub prtDeviseCoursForm()
'---------------------------------------------------------
Dim X As String, K As Integer

XPrt.FontSize = 8
XPrt.FontBold = False

'XPrt.FillStyle = 0
XPrt.DrawWidth = 3
'XPrt.ForeColor = RGB(0, 0, 0)
'XPrt.FillStyle = 1
Call frmElpPrt.prtTrame(prtMinX, prtMinY, prtMaxX, prtMinY + prtHeaderHeight, "B")

XPrt.Line (prtMinX, prtMinY)-(prtMinX, prtMaxY)
XPrt.Line (prtMaxX, prtMinY)-(prtMaxX, prtMaxY)
K = prtMinY + prtHeaderHeight + 10
Call frmElpPrt.prtTrame(colP + 1600, K, colP + 3000, prtMaxY - 10, " ")
Call frmElpPrt.prtTrame(colA + 1100, K, colA + 2100, prtMaxY - 10, " ")
Call frmElpPrt.prtTrame(colV + 1100, K, colV + 2100, prtMaxY - 10, " ")

XPrt.DrawWidth = 2
XPrt.Line (colA, prtMinY)-(colA, prtMaxY)
XPrt.Line (colV, prtMinY)-(colV, prtMaxY)
XPrt.Line (colX, prtMinY)-(colX, prtMaxY)
XPrt.DrawWidth = 1
XPrt.Line (colA, prtMinY + 300)-(colX, prtMinY + 300)
'---------------------------------------------------------
XPrt.FontBold = True

XPrt.CurrentY = prtMinY + (prtHeaderHeight - XPrt.TextHeight(X)) / 2
XPrt.CurrentX = colP + 3000
XPrt.Print "cours pivot";

XPrt.CurrentY = prtMinY + 50
X = "Achat"
XPrt.CurrentX = (colA + colV - XPrt.TextWidth(X)) / 2: XPrt.Print X;
X = "Vente"
XPrt.CurrentX = (colV + colX - XPrt.TextWidth(X)) / 2: XPrt.Print X;
XPrt.FontBold = False

XPrt.CurrentY = prtMinY + 350
XPrt.CurrentX = colA + 200: XPrt.Print "en Compte";
XPrt.FontBold = True
XPrt.CurrentX = colA + 1400: XPrt.Print "Billets";
XPrt.FontBold = False
XPrt.CurrentX = colA + 2300: XPrt.Print "Privilégié";
XPrt.CurrentX = colV + 200: XPrt.Print "en Compte";
XPrt.FontBold = True
XPrt.CurrentX = colV + 1400: XPrt.Print "Billets";
XPrt.FontBold = False
XPrt.CurrentX = colV + 2300: XPrt.Print "Privilégié";
If Not blnList Then
    XPrt.CurrentX = colX + 50: XPrt.Print "Saisie";
    XPrt.CurrentX = colX + 2650: XPrt.Print "Validation";
End If
XPrt.CurrentY = prtMinY + prtHeaderHeight + prtlineHeight - XPrt.TextHeight("test")

End Sub

'---------------------------------------------------------
Public Sub prtDeviseCoursLine()
'---------------------------------------------------------
Dim X As String, K As Integer
Dim Situation As String

If prtCurrentY + prtParagraphHeight > prtMaxY Then
    frmElpPrt.prtNewPage
    prtDeviseCoursForm
'Else
    'frmElpPrt.prtLineY
End If

XPrt.FontBold = False
'------------------------------------------ligne 1--------------

XPrt.FontSize = 8

XPrt.CurrentX = colP + 50
XPrt.Print Trim(recDevise1.DevLib) & " / " & Trim(recDevise2.DevX);
If recDeviseCours.QD1 <> 1 Then
    X = Format$(recDeviseCours.QD1, "### ##0")
    XPrt.CurrentX = colP + 2100 - XPrt.TextWidth(X)
    XPrt.Print X;
End If

XPrt.CurrentX = colP + 2200
XPrt.Print Trim(recDevise1.DevX) & " / " & Trim(recDevise2.DevX);

X = Format$(recDeviseCours.QD2CoursPivot, "### ### ##0.00 000")
XPrt.CurrentX = colP + 3900 - XPrt.TextWidth(X)
XPrt.Print X;

X = Format$(recDeviseCours.QD2AchatEnCompte, "### ### ##0.00 000")
XPrt.CurrentX = colA + 950 - XPrt.TextWidth(X)
XPrt.Print X;

X = Format$(recDeviseCours.QD2AchatNormal, "### ### ##0.00 000")
XPrt.CurrentX = colA + 2000 - XPrt.TextWidth(X)
XPrt.Print X;

X = Format$(recDeviseCours.QD2AchatPrivilégié, "### ### ##0.00 000")
XPrt.CurrentX = colA + 3050 - XPrt.TextWidth(X)
XPrt.Print X;

X = Format$(recDeviseCours.QD2VenteEnCompte, "### ### ##0.00 000")
XPrt.CurrentX = colV + 950 - XPrt.TextWidth(X)
XPrt.Print X;

X = Format$(recDeviseCours.QD2VenteNormal, "### ### ##0.00 000")
XPrt.CurrentX = colV + 2000 - XPrt.TextWidth(X)
XPrt.Print X;

X = Format$(recDeviseCours.QD2VentePrivilégié, "### ### ##0.00 000")
XPrt.CurrentX = colV + 3050 - XPrt.TextWidth(X)
XPrt.Print X;
If blnList Then
    If Trim(recDeviseCours.ValidationUsr) = "" Then XPrt.CurrentX = colX + 50: XPrt.Print "?";

Else
    XPrt.FontSize = 6
    XPrt.CurrentX = colX + 50
    XPrt.Print recDeviseCours.SaisieUsr;
    XPrt.CurrentX = colX + 900
    XPrt.Print dateImp(recDeviseCours.SaisieAMJ);
    XPrt.CurrentX = colX + 2000
    XPrt.Print Mid$(timeImp(recDeviseCours.SaisieHMS), 1, 7);
    
    XPrt.CurrentX = colX + 2650
    XPrt.Print recDeviseCours.ValidationUsr;
    XPrt.CurrentX = colX + 3600
    XPrt.Print dateImp(recDeviseCours.ValidationAMJ);
    XPrt.CurrentX = colX + 4700
    XPrt.Print Mid$(timeImp(recDeviseCours.ValidationHMS), 1, 7);
    prtDeviseCoursMarge
End If

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5

End Sub








Public Sub prtDeviseCoursMarge()
Dim dblX As Double, X As String

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight

dblX = (recDeviseCours.QD2AchatEnCompte - recDeviseCours.QD2CoursPivot) / recDeviseCours.QD2CoursPivot
    dblX = Fix(dblX * 10000 - 0.5) / 100
X = Format$(dblX, "###0.00")
XPrt.CurrentX = colA + 950 - XPrt.TextWidth(X)
XPrt.Print X & " %";

dblX = (recDeviseCours.QD2AchatNormal - recDeviseCours.QD2CoursPivot) / recDeviseCours.QD2CoursPivot
    dblX = Fix(dblX * 10000 - 0.5) / 100
X = Format$(dblX, "###0.00")
XPrt.CurrentX = colA + 2000 - XPrt.TextWidth(X)
XPrt.Print X & " %";

dblX = (recDeviseCours.QD2AchatPrivilégié - recDeviseCours.QD2CoursPivot) / recDeviseCours.QD2CoursPivot
    dblX = Fix(dblX * 10000 - 0.5) / 100
X = Format$(dblX, "###0.00")
XPrt.CurrentX = colA + 3050 - XPrt.TextWidth(X)
XPrt.Print X & " %";

dblX = (recDeviseCours.QD2VenteEnCompte - recDeviseCours.QD2CoursPivot) / recDeviseCours.QD2CoursPivot
    dblX = Fix(dblX * 10000 + 0.5) / 100
X = Format$(dblX, "###0.00")
XPrt.CurrentX = colV + 950 - XPrt.TextWidth(X)
XPrt.Print X & " %";

dblX = (recDeviseCours.QD2VenteNormal - recDeviseCours.QD2CoursPivot) / recDeviseCours.QD2CoursPivot
    dblX = Fix(dblX * 10000 + 0.5) / 100
X = Format$(dblX, "###0.00")
XPrt.CurrentX = colV + 2000 - XPrt.TextWidth(X)
XPrt.Print X & " %";

dblX = (recDeviseCours.QD2VentePrivilégié - recDeviseCours.QD2CoursPivot) / recDeviseCours.QD2CoursPivot
    dblX = Fix(dblX * 10000 + 0.5) / 100
X = Format$(dblX, "###0.00")
XPrt.CurrentX = colV + 3050 - XPrt.TextWidth(X)
XPrt.Print X & " %";

End Sub
