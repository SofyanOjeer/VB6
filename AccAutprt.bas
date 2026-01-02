Attribute VB_Name = "prtAccAut"
Option Explicit

Dim recAccAut As typeAccAut
'---------------------------------------------------------
Public Sub prtAccAutX(Msg As String)
'---------------------------------------------------------
Dim K As Integer, K1 As Integer, K2 As Integer
On Error GoTo prtError


Set XPrt = Printer
If prtShow Then frmElpPrt.Show vbModeless

prtOrientation = vbPRORPortrait
prtTitleText = "liste des Autorisations "
prtPgmName = "prtAccAut"
prtTitleUsr = usrName

prtLineNb = 1
prtlineHeight = 250
prtHeaderHeight = 300

frmElpPrt.prtStdInit
prtAccAutForm

K1 = Val(Mid$(Msg, 1, 6))
K2 = Val(Mid$(Msg, 7, 6))
For K = K1 To K2
    recAccAut = arrAccAut(K)
    prtAccAutLine
    DoEvents: If prtKillDoc Then frmElpPrt.Hide: Exit Sub
Next K

'frmElpPrt.prtLineY
frmElpPrt.prtEndDoc
If prtShow Then frmElpPrt.Hide
Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub

'---------------------------------------------------------
Private Sub prtAccAutForm()
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
XPrt.Line (3800, prtMinY)-(3800, prtMaxY)
XPrt.Line (6800, prtMinY)-(6800, prtMaxY)
'XPrt.Line (13600, prtMinY)-(13600, prtMaxY)
'---------------------------------------------------------
X = "Nature"
XPrt.CurrentY = prtMinY + (prtHeaderHeight - XPrt.TextHeight(X)) / 2
XPrt.CurrentX = 300
XPrt.Print X;
XPrt.CurrentX = 5200: XPrt.Print "Droits";
XPrt.CurrentX = 7000: XPrt.Print "Validité du";
XPrt.CurrentX = 9300: XPrt.Print "Au";
'XPrt.CurrentX = 14400: XPrt.Print "Validité";
'XPrt.CurrentX = 10100: XPrt.Print "Banque";
'XPrt.CurrentX = 10900: XPrt.Print "Guichet";
'XPrt.CurrentX = 12400: XPrt.Print "Avis";
XPrt.CurrentY = prtMinY + prtHeaderHeight + 50

End Sub
'---------------------------------------------------------
Private Sub prtAccAutLine()
'---------------------------------------------------------
Dim X As String

If XPrt.CurrentY + prtlineHeight > prtMaxY Then
    frmElpPrt.prtNewPage
    prtAccAutForm
'Else
 '   frmElpPrt.prtLineY
End If

XPrt.FontBold = False


'------------------------------------ligne 1
XPrt.FontSize = 8

XPrt.CurrentX = 300
XPrt.Print recAccAut.AccAutId;

XPrt.CurrentX = 1400
XPrt.Print recAccAut.AccAutK1;

XPrt.CurrentX = 2700
XPrt.Print recAccAut.AccAutK2;

XPrt.CurrentX = 4500
XPrt.Print recAccAut.AccAutTxt;

XPrt.CurrentX = 7000
XPrt.Print dateImp(recAccAut.AccAutDD);

XPrt.CurrentX = 8200
XPrt.Print timeImp(recAccAut.AccAutHD);

XPrt.CurrentX = 9000
XPrt.Print dateImp(recAccAut.AccAutDF);

XPrt.CurrentX = 10200
XPrt.Print timeImp(recAccAut.AccAutHF);


XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
End Sub




