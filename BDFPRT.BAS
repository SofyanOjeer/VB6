Attribute VB_Name = "prtBDF"
Option Explicit
'---------------------------------------------------------
Private recBdfE As typeBdfE
Private recBdfG As typeBdfG
Dim I As Integer

'---------------------------------------------------------
Private Sub prtBdfEForm()
'---------------------------------------------------------
Dim X As String

XPrt.FontSize = 8
XPrt.FontBold = True

XPrt.FillStyle = 0
XPrt.DrawWidth = 3
XPrt.ForeColor = RGB(0, 0, 0)
XPrt.FillStyle = 1

XPrt.Line (prtMinX, prtMinY)-(prtMaxX, prtMaxY), , B
XPrt.Line (prtMinX, prtMinY + prtHeaderHeight)-(prtMaxX, prtMinY + prtHeaderHeight)

XPrt.DrawWidth = 1

XPrt.Line (5000, prtMinY)-(5000, prtMaxY)

XPrt.Line (800, prtMinY)-(800, prtMaxY)

'---------------------------------------------------------

X = "Code"
XPrt.CurrentY = prtMinY + (prtHeaderHeight - XPrt.TextHeight(X)) / 2
XPrt.CurrentX = prtMinX + 50
XPrt.Print X;

X = "Dénomination / adresse"
XPrt.CurrentX = 850
XPrt.Print X;
prtCurrentY = prtMinY + prtHeaderHeight

End Sub


'---------------------------------------------------------
Private Sub prtBdfELine()
'---------------------------------------------------------
Dim X As String

If prtCurrentY + prtParagraphHeight > prtMaxY Then
    frmElpPrt.prtNewPage
    prtBdfEForm
Else
    frmElpPrt.prtLineY
End If

XPrt.FontBold = False

XPrt.CurrentY = prtCurrentY + prtlineHeight - XPrt.TextHeight("test")

XPrt.FontSize = 7
XPrt.CurrentX = 5050
XPrt.Print "Code situation";
XPrt.CurrentX = 7500
XPrt.Print ":";
XPrt.CurrentX = 11500
XPrt.Print "Date début de validité";
XPrt.CurrentX = 13500
XPrt.Print ":";

XPrt.FontSize = 8

XPrt.CurrentX = prtMinX + 50
XPrt.Print recBdfE.BdfE2;
'---------------------------------------------------------

XPrt.CurrentX = 850
XPrt.Print recBdfE.BdfE4;
XPrt.CurrentX = 8000 - XPrt.TextWidth(recBdfE.BdfE3)
XPrt.Print recBdfE.BdfE3;

X = dateImp(recBdfE.BdfE7)
XPrt.CurrentX = 15000 - XPrt.TextWidth(X)
XPrt.Print X;

'---------------------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight

XPrt.FontSize = 7
XPrt.CurrentX = 5050
XPrt.Print "Code nature activité";
XPrt.CurrentX = 7500
XPrt.Print ":";

XPrt.CurrentX = 11500
XPrt.Print "Date fin de validité";
XPrt.CurrentX = 13500
XPrt.Print ":";


XPrt.FontSize = 8
XPrt.CurrentX = 850
XPrt.Print recBdfE.BdfE10A;
XPrt.CurrentX = 8000 - XPrt.TextWidth(recBdfE.BdfE5)
XPrt.Print recBdfE.BdfE5;

X = dateImp(recBdfE.BdfE8)
XPrt.CurrentX = 15000 - XPrt.TextWidth(X)
XPrt.Print X;

'---------------------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight

XPrt.FontSize = 7
XPrt.CurrentX = 5050
XPrt.Print "Code établissement nouveau";
XPrt.CurrentX = 7500
XPrt.Print ":";


XPrt.CurrentX = 11500
XPrt.Print "Date de disparition";
XPrt.CurrentX = 13500
XPrt.Print ":";


XPrt.FontSize = 8
XPrt.CurrentX = 850
XPrt.Print recBdfE.BdfE10B;
XPrt.CurrentX = 8000 - XPrt.TextWidth(recBdfE.BdfE12)
XPrt.Print recBdfE.BdfE12;

X = dateImp(recBdfE.BdfE9)
XPrt.CurrentX = 15000 - XPrt.TextWidth(X)
XPrt.Print X;

'---------------------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight

XPrt.FontSize = 7
XPrt.CurrentX = 5050
XPrt.Print "Organe représentatif";
XPrt.CurrentX = 7500
XPrt.Print ":";


XPrt.CurrentX = 11500
XPrt.Print "Date prise effet identifiant";
XPrt.CurrentX = 13500
XPrt.Print ":";


XPrt.FontSize = 8
XPrt.CurrentX = 850
XPrt.Print recBdfE.BdfE10C;
XPrt.CurrentX = 8000 - XPrt.TextWidth(recBdfE.BdfE6)
XPrt.Print recBdfE.BdfE6;
X = dateImp(recBdfE.BdfE14)
XPrt.CurrentX = 15000 - XPrt.TextWidth(X)
XPrt.Print X;


'---------------------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight

XPrt.FontSize = 7

XPrt.CurrentX = 5050
XPrt.Print "option identifiant invariant";
XPrt.CurrentX = 7500
XPrt.Print ":";


XPrt.FontSize = 8

XPrt.CurrentX = 850
XPrt.Print recBdfE.BdfE10CP;
XPrt.CurrentX = 8000 - XPrt.TextWidth(recBdfE.BdfE13)
XPrt.Print recBdfE.BdfE13;
XPrt.CurrentX = 1300
XPrt.Print recBdfE.BdfE10D;

'---------------------------------------------------------

prtCurrentY = prtCurrentY + prtParagraphHeight
End Sub

'---------------------------------------------------------
Public Sub prtBdfE(Msg As String)
'---------------------------------------------------------
Dim K As Integer, K1 As Integer, K2 As Integer
On Error GoTo prtError


Set XPrt = Printer
K1 = Val(Mid$(Msg, 1, 6))
K2 = Val(Mid$(Msg, 7, 6))

frmElpPrt.Show vbModeless


prtOrientation = vbPRORLandscape
prtTitleText = "Banque de France : liste des Etablissements "
prtPgmName = "prtBdfE"
prtTitleUsr = usrName

prtLineNb = 5
prtlineHeight = 250
prtHeaderHeight = 300

frmElpPrt.prtStdInit
prtBdfEForm

For K = K1 To K2

recBdfE = arrBdfE(K)
prtBdfELine

DoEvents: If prtKillDoc Then frmElpPrt.Hide: Exit Sub

Next K
    
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
Public Sub prtBdfG(Msg As String)
'---------------------------------------------------------
Dim K As Integer, K1 As Integer, K2 As Integer
On Error GoTo prtError

Set XPrt = Printer
K1 = Val(Mid$(Msg, 1, 6))
K2 = Val(Mid$(Msg, 7, 6))

frmElpPrt.Show vbModeless

recBdfEInit recBdfE

prtOrientation = vbPRORLandscape
prtTitleText = "Banque de France : liste des Guichets"
prtPgmName = "prtBdfG"
prtTitleUsr = usrName

prtLineNb = 8
prtlineHeight = 250
prtHeaderHeight = 300

frmElpPrt.prtStdInit
prtBdfGForm

For K = K1 To K2

recBdfG = arrBdfG(K)
prtBdfGLine

DoEvents: If prtKillDoc Then frmElpPrt.Hide: Exit Sub

Next K

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
Public Sub prtBdfGForm()
'---------------------------------------------------------
Dim X As String

XPrt.FontSize = 8
XPrt.FontBold = True

XPrt.FillStyle = 0
XPrt.DrawWidth = 3
XPrt.ForeColor = RGB(0, 0, 0)
XPrt.FillStyle = 1

XPrt.Line (prtMinX, prtMinY)-(prtMaxX, prtMaxY), , B
XPrt.Line (prtMinX, prtMinY + prtHeaderHeight)-(prtMaxX, prtMinY + prtHeaderHeight)

XPrt.DrawWidth = 1

XPrt.Line (5000, prtMinY)-(5000, prtMaxY)

XPrt.Line (800, prtMinY)-(800, prtMaxY)

'---------------------------------------------------------

X = "Code"
XPrt.CurrentY = prtMinY + (prtHeaderHeight - XPrt.TextHeight(X)) / 2
XPrt.CurrentX = prtMinX + 50
XPrt.Print X;

X = "Dénomination / adresse"
XPrt.CurrentX = 850
XPrt.Print X;

prtCurrentY = prtMinY + prtHeaderHeight

End Sub

'---------------------------------------------------------
'           Document Réalisé en partie par J/F
'---------------------------------------------------------
Public Sub prtBdfGLine()
'---------------------------------------------------------
Dim X As String

If prtCurrentY + prtParagraphHeight > prtMaxY Then
    frmElpPrt.prtNewPage
    prtBdfGForm
Else
    frmElpPrt.prtLineY
End If

XPrt.FontBold = False
'------------------------------------------ligne 1--------------

XPrt.CurrentY = prtCurrentY + prtlineHeight - XPrt.TextHeight("test")
XPrt.FontSize = 8

recBdfE.BdfE2 = recBdfG.BdfG2
If Not IsNull(srvBdfEFind(recBdfE)) Then recBdfE.BdfE4 = "? Etablissement ?"
XPrt.CurrentX = prtMinX + 50
XPrt.Print recBdfG.BdfG2;
XPrt.CurrentX = 850
XPrt.Print recBdfE.BdfE4;
'------------------------------------------ligne 2--------------

XPrt.CurrentY = prtCurrentY + prtlineHeight - XPrt.TextHeight("test")

XPrt.FontSize = 7
XPrt.CurrentX = 5050
XPrt.Print "Nature du guichet";
XPrt.CurrentX = 6300
XPrt.Print ":";
XPrt.CurrentX = 8700
XPrt.Print "comptoir B.D.F";
XPrt.CurrentX = 10700
XPrt.Print ":";
XPrt.CurrentX = 12500
XPrt.Print "Code Situation";
XPrt.CurrentX = 14500
XPrt.Print ":";

XPrt.FontSize = 8

XPrt.CurrentX = prtMinX + 50
XPrt.Print recBdfG.BdfG2;
XPrt.CurrentX = 850
XPrt.Print recBdfG.BdfG5;
XPrt.CurrentX = 6500
XPrt.Print recBdfG.BdfG21;
XPrt.CurrentX = 10900
XPrt.Print recBdfG.BdfG13;
XPrt.CurrentX = 14700
XPrt.Print recBdfG.BdfG4;
XPrt.Print X;

'---------------------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'------------------------------------------ligne 3---------------


XPrt.FontSize = 7

XPrt.CurrentX = 5050
XPrt.Print "Nom du guichet";
XPrt.CurrentX = 6300
XPrt.Print ":";

XPrt.CurrentX = 8700
XPrt.Print "Nom comptoir B.D.F";
XPrt.CurrentX = 10700
XPrt.Print ":";
XPrt.CurrentX = 12500
XPrt.Print "Code établissement nouveau";
XPrt.CurrentX = 14500
XPrt.Print ":";

XPrt.FontSize = 8
XPrt.CurrentX = prtMinX + 50
XPrt.Print recBdfG.BdfG3;
XPrt.CurrentX = 850
XPrt.Print recBdfG.BdfG14A;
XPrt.CurrentX = 6500
XPrt.Print recBdfG.BdfG6;
XPrt.CurrentX = 10900
XPrt.Print recBdfG.BdfG23;
XPrt.CurrentX = 14700
XPrt.Print recBdfG.BdfG7;

XPrt.Print X;

'-----------------------------------------ligne 4----------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight

XPrt.FontSize = 7
XPrt.CurrentX = 5050
XPrt.Print "Domiciliation RIB";
XPrt.CurrentX = 6300
XPrt.Print ":";

XPrt.CurrentX = 8700
XPrt.Print "Code localité de compens 1";
XPrt.CurrentX = 10700
XPrt.Print ":";
XPrt.CurrentX = 12500
XPrt.Print "Code guichet nouveau";
XPrt.CurrentX = 14500
XPrt.Print ":";

XPrt.FontSize = 8

XPrt.CurrentX = 850
XPrt.Print recBdfG.BdfG14B;
XPrt.CurrentX = 6500
XPrt.Print recBdfG.BdfG20;
XPrt.CurrentX = 10900
XPrt.Print recBdfG.BdfG10;
XPrt.CurrentX = 14700
XPrt.Print recBdfG.BdfG8;

'----------------------------------------ligne 5-----------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight

XPrt.FontSize = 7
XPrt.CurrentX = 5050
XPrt.Print "Adresse Swift";
XPrt.CurrentX = 6300
XPrt.Print ":";

XPrt.CurrentX = 8700
XPrt.Print "Code localité de compens 2";
XPrt.CurrentX = 10700
XPrt.Print ":";
XPrt.CurrentX = 12500
XPrt.Print "Date début de validité";
XPrt.CurrentX = 14500
XPrt.Print ":";

XPrt.FontSize = 8

XPrt.CurrentX = 850
XPrt.Print recBdfG.BdfG14C;
XPrt.CurrentX = 6500
XPrt.Print recBdfG.BdfG15;
XPrt.CurrentX = 10900
XPrt.Print recBdfG.BdfG11;
XPrt.CurrentX = 14700
XPrt.Print dateImp(recBdfG.BdfG16);


'---------------------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'----------------------------------------ligne 6-----------------

XPrt.FontSize = 7

XPrt.CurrentX = 8700
XPrt.Print "Code recouvreur";
XPrt.CurrentX = 10700
XPrt.Print ":";
XPrt.CurrentX = 12500
XPrt.Print "Date fin de validité";
XPrt.CurrentX = 14500
XPrt.Print ":";

XPrt.FontSize = 8

XPrt.CurrentX = 850
XPrt.Print recBdfG.BdfG14CP;
XPrt.CurrentX = 1300
XPrt.Print recBdfG.BdfG14D;
XPrt.CurrentX = 10900
XPrt.Print recBdfG.BdfG12;
XPrt.CurrentX = 14700
XPrt.Print dateImp(recBdfG.BdfG17);

'---------------------------------------ligne 7------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight

XPrt.FontSize = 7


XPrt.CurrentX = 8700
XPrt.Print "Code routage guichet domicile";
XPrt.CurrentX = 10700
XPrt.Print ":";
XPrt.CurrentX = 12500
XPrt.Print "Date disparition du FGD";
XPrt.CurrentX = 14500
XPrt.Print ":";

XPrt.FontSize = 8

XPrt.CurrentX = 10900
XPrt.Print recBdfG.BdfG19;
XPrt.CurrentX = 14700
XPrt.Print dateImp(recBdfG.BdfG18);


'----------------------------------------ligne 8------------

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight

XPrt.FontSize = 7
XPrt.CurrentX = 8700
XPrt.Print "Code règlement crédits";
XPrt.CurrentX = 10700
XPrt.Print ":";

XPrt.FontSize = 8

XPrt.CurrentX = 10900
XPrt.Print recBdfG.BdfG22;


prtCurrentY = prtCurrentY + prtParagraphHeight
End Sub


