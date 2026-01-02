Attribute VB_Name = "prtPaieSage"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------

Dim recLine As String
'---------------------------------------------------------
Public Sub prtPaieSageX(Msg As String)
'---------------------------------------------------------
Dim K As Integer, K1 As Integer, K2 As Integer, mDicrub As Integer
Dim X
Set XPrt = Printer


'frmElpPrt.Show vbModal 'vbModeless

prtOrientation = vbPRORPortrait
prtTitleText = "Paie SAGE : Virements"
prtPgmName = "prtPaieSage"
prtTitleUsr = usrName
prtFontName = "Courier"
prtLineNb = 1
prtlineHeight = 250
prtHeaderHeight = 300

frmElpPrt.prtStdInit
prtPaieSageForm

Open Msg For Input As #1
Do Until EOF(1)
    Input #1, recLine
    prtPaieSageLine
Loop
Close #1
prtFontName = prtFontNameZ
Call frmElpPrt.prtEndDoc(1000)

frmElpPrt.Hide

End Sub



'---------------------------------------------------------
Public Sub prtPaieSageForm()
'---------------------------------------------------------
Dim X As String

XPrt.FontSize = 7
XPrt.FontBold = True

XPrt.DrawWidth = 3
XPrt.ForeColor = RGB(0, 0, 0)

XPrt.Line (prtMinX, prtMinY)-(prtMaxX, prtMaxY), prtLineColor, B
XPrt.Line (prtMinX, prtMinY + prtHeaderHeight)-(prtMaxX, prtMinY + prtHeaderHeight), prtLineColor


XPrt.DrawWidth = 1
'----------------------------------------ligne 1-----------------

XPrt.CurrentY = prtMinY + (prtHeaderHeight - XPrt.TextHeight(X)) / 2
XPrt.CurrentX = prtMinX
XPrt.Print "";

XPrt.CurrentY = prtMinY + prtHeaderHeight - XPrt.TextHeight("test")
XPrt.FontBold = False


End Sub


'---------------------------------------------------------
Public Sub prtPaieSageLine()
'---------------------------------------------------------
Dim X As String, K As Integer, mCurrenty As Integer

If XPrt.CurrentY + prtlineHeight * 2 > prtMaxY Then
    frmElpPrt.prtNewPage
    prtPaieSageForm
End If

'------------------------------------------ligne 1--------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'----------------------------------ligne2

XPrt.CurrentX = prtMinX

XPrt.Print recLine;
End Sub



