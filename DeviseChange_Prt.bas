Attribute VB_Name = "prtDeviseChange"
Option Explicit
Dim recDevise1 As typeDevise, recDevise2 As typeDevise
Private recDeviseChange As typeDeviseChange
Dim I As Integer, mCurrenty As Integer
Dim colP As Integer, colA As Integer, colV As Integer, colX As Integer
Dim blnList As Boolean
Dim mOrigine As String
'---------------------------------------------------------
 Public Sub prtDeviseChangeX(Msg As String, Text As String)
'---------------------------------------------------------
Dim K As Integer, K1 As Integer, K2 As Integer
Dim X As String
On Error GoTo prtError

Set XPrt = Printer
K1 = Val(mId$(Msg, 1, 6))
K2 = Val(mId$(Msg, 7, 6))
mOrigine = mId$(Msg, 13, 1)


frmElpPrt.Show vbModeless

recDeviseChange_Init recDeviseChange
If Trim(Text) = "" Then
    blnList = True
    prtOrientation = vbPRORPortrait
Else
    blnList = False
    prtOrientation = vbPRORLandscape
End If
If mOrigine = "C" Then
    prtTitleText = "Comptabilité : Cours de Change" & Text
Else
    prtTitleText = "Trésorerie : Cours de Change" & Text
End If
prtTitleText = "Cours de Change" & Text
prtPgmName = "prtDeviseChange"
prtTitleUsr = usrName

prtLineNb = 1
prtlineHeight = 250
prtHeaderHeight = 600

frmElpPrt.prtStdInit
colP = prtMinX
colA = prtMinX + 4000
colV = prtMinX + 7200
colX = prtMinX + 10400
recDevise1.DevX = arrDeviseChange(1).Id1
prtDeviseChangeForm

For K = K1 To K2
recDeviseChange = arrDeviseChange(K)
If Trim(recDeviseChange.Method) <> constIgnore _
Or Trim(recDeviseChange.Method) <> constDelete Then
     If recDevise1.DevX <> recDeviseChange.Id1 Then
        mCurrenty = XPrt.CurrentY
        XPrt.Line (prtMinX, mCurrenty - 50)-(prtMaxX, mCurrenty - 100)
        XPrt.CurrentY = mCurrenty
    End If
    DevCode recDeviseChange.Id1: recDevise1 = XDevise
    DevCode recDeviseChange.Id2: recDevise2 = XDevise
    prtDeviseChangeLine
End If
DoEvents: If prtKillDoc Then frmElpPrt.Hide: Exit Sub

Next K
XPrt.Line (prtMinX, XPrt.CurrentY - 50)-(prtMaxX, XPrt.CurrentY - 100)
    
If mOrigine = "T" Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
    XPrt.FontSize = 12
    XPrt.FontBold = True
    frmElpPrt.prtCentré prtMaxX / 2, "Pour tout montant supérieur ou égal à  EUR 10 000 ou contrevaleur : AVISEZ la TC"
End If

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
Public Sub prtDeviseChangeForm()
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
Public Sub prtDeviseChangeLine()
'---------------------------------------------------------
Dim X As String, K As Integer
Dim Situation As String

If prtCurrentY + prtParagraphHeight > prtMaxY Then
    frmElpPrt.prtNewPage
    prtDeviseChangeForm
'Else
    'frmElpPrt.prtLineY
End If

XPrt.FontBold = False
'------------------------------------------ligne 1--------------

XPrt.FontSize = 6

XPrt.CurrentX = colP + 50
XPrt.Print Trim(recDevise1.DevLib) & " / " & Trim(recDevise2.DevLib);
XPrt.FontSize = 8
If recDeviseChange.QD1 <> 1 Then
    X = Format$(recDeviseChange.QD1, "### ##0")
    XPrt.CurrentX = colP + 2100 - XPrt.TextWidth(X)
    XPrt.Print X;
End If

XPrt.CurrentX = colP + 2200
XPrt.Print Trim(recDevise1.DevX) & " / " & Trim(recDevise2.DevX);

X = Format$(recDeviseChange.QD2CoursPivot, "### ### ##0.00 000")
XPrt.CurrentX = colP + 3900 - XPrt.TextWidth(X)
XPrt.Print X;

X = Format$(recDeviseChange.QD2AchatEnCompte, "### ### ##0.00 000")
XPrt.CurrentX = colA + 950 - XPrt.TextWidth(X)
XPrt.Print X;

X = Format$(recDeviseChange.QD2AchatNormal, "### ### ##0.00 000")
XPrt.CurrentX = colA + 2000 - XPrt.TextWidth(X)
XPrt.Print X;

X = Format$(recDeviseChange.QD2AchatPrivilégié, "### ### ##0.00 000")
XPrt.CurrentX = colA + 3050 - XPrt.TextWidth(X)
XPrt.Print X;

X = Format$(recDeviseChange.QD2VenteEnCompte, "### ### ##0.00 000")
XPrt.CurrentX = colV + 950 - XPrt.TextWidth(X)
XPrt.Print X;

X = Format$(recDeviseChange.QD2VenteNormal, "### ### ##0.00 000")
XPrt.CurrentX = colV + 2000 - XPrt.TextWidth(X)
XPrt.Print X;

X = Format$(recDeviseChange.QD2VentePrivilégié, "### ### ##0.00 000")
XPrt.CurrentX = colV + 3050 - XPrt.TextWidth(X)
XPrt.Print X;
If blnList Then
    If Trim(recDeviseChange.ValidationUsr) = "" Then XPrt.CurrentX = colX + 50: XPrt.Print "?";

Else
    XPrt.FontSize = 6
    XPrt.CurrentX = colX + 50
    XPrt.Print recDeviseChange.SaisieUsr;
    XPrt.CurrentX = colX + 900
    XPrt.Print dateImp(recDeviseChange.SaisieAMJ);
    XPrt.CurrentX = colX + 2000
    XPrt.Print mId$(timeImp(recDeviseChange.SaisieHMS), 1, 7);
    
    XPrt.CurrentX = colX + 2650
    XPrt.Print recDeviseChange.ValidationUsr;
    XPrt.CurrentX = colX + 3600
    XPrt.Print dateImp(recDeviseChange.ValidationAMJ);
    XPrt.CurrentX = colX + 4700
    XPrt.Print mId$(timeImp(recDeviseChange.ValidationHMS), 1, 7);
    If recDeviseChange.Origine = "T" Then prtDeviseChangeMarge
End If

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5

End Sub








Public Sub prtDeviseChangeMarge()
Dim dblX As Double, X As String

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight

dblX = (recDeviseChange.QD2AchatEnCompte - recDeviseChange.QD2CoursPivot) / recDeviseChange.QD2CoursPivot
    dblX = Fix(dblX * 10000 - 0.5) / 100
X = Format$(dblX, "###0.00")
XPrt.CurrentX = colA + 950 - XPrt.TextWidth(X)
XPrt.Print X & " %";

dblX = (recDeviseChange.QD2AchatNormal - recDeviseChange.QD2CoursPivot) / recDeviseChange.QD2CoursPivot
    dblX = Fix(dblX * 10000 - 0.5) / 100
X = Format$(dblX, "###0.00")
XPrt.CurrentX = colA + 2000 - XPrt.TextWidth(X)
XPrt.Print X & " %";

dblX = (recDeviseChange.QD2AchatPrivilégié - recDeviseChange.QD2CoursPivot) / recDeviseChange.QD2CoursPivot
    dblX = Fix(dblX * 10000 - 0.5) / 100
X = Format$(dblX, "###0.00")
XPrt.CurrentX = colA + 3050 - XPrt.TextWidth(X)
XPrt.Print X & " %";

dblX = (recDeviseChange.QD2VenteEnCompte - recDeviseChange.QD2CoursPivot) / recDeviseChange.QD2CoursPivot
    dblX = Fix(dblX * 10000 + 0.5) / 100
X = Format$(dblX, "###0.00")
XPrt.CurrentX = colV + 950 - XPrt.TextWidth(X)
XPrt.Print X & " %";

dblX = (recDeviseChange.QD2VenteNormal - recDeviseChange.QD2CoursPivot) / recDeviseChange.QD2CoursPivot
    dblX = Fix(dblX * 10000 + 0.5) / 100
X = Format$(dblX, "###0.00")
XPrt.CurrentX = colV + 2000 - XPrt.TextWidth(X)
XPrt.Print X & " %";

dblX = (recDeviseChange.QD2VentePrivilégié - recDeviseChange.QD2CoursPivot) / recDeviseChange.QD2CoursPivot
    dblX = Fix(dblX * 10000 + 0.5) / 100
X = Format$(dblX, "###0.00")
XPrt.CurrentX = colV + 3050 - XPrt.TextWidth(X)
XPrt.Print X & " %";

End Sub
