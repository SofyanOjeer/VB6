Attribute VB_Name = "prtLrBafiMsg"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------

Dim recLine As String
'---------------------------------------------------------
Public Sub prtLrBafiMsgX(Msg As String)
'---------------------------------------------------------
Dim K As Integer, K1 As Integer, K2 As Integer, mDicrub As Integer
Dim X
Set XPrt = Printer


'frmElpPrt.Show vbModal 'vbModeless

prtOrientation = vbPRORPortrait
prtTitleText = "LrBafi : Compte-rendu d'extraction AS400"
prtPgmName = "prtLrBafiMsg"
prtTitleUsr = usrName
prtFontName = "Courier"
prtLineNb = 1
prtlineHeight = 250
prtHeaderHeight = 300

frmElpPrt.prtStdInit
prtLrBafiMsgForm

Open Msg For Input As #1
Do Until EOF(1)
    Input #1, recLine
    prtLrBafiMsgLine
Loop
Close #1
prtFontName = prtFontNameZ
frmElpPrt.prtEndDoc

frmElpPrt.Hide

End Sub



'---------------------------------------------------------
Public Sub prtLrBafiMsgForm()
'---------------------------------------------------------
Dim X As String

XPrt.FontSize = 7
XPrt.FontBold = True

XPrt.DrawWidth = 3
XPrt.ForeColor = RGB(0, 0, 0)

XPrt.Line (prtMinX, prtMinY)-(prtMaxX, prtMaxY), , B
XPrt.Line (prtMinX, prtMinY + prtHeaderHeight)-(prtMaxX, prtMinY + prtHeaderHeight)


XPrt.DrawWidth = 1
'----------------------------------------ligne 1-----------------

XPrt.CurrentY = prtMinY + (prtHeaderHeight - XPrt.TextHeight(X)) / 2
XPrt.CurrentX = prtMinX
XPrt.Print "Origine";
XPrt.CurrentX = prtMinX + 3500
XPrt.Print "Devise";
XPrt.CurrentX = XPrt.CurrentX + 500
XPrt.Print "Compte";
XPrt.CurrentX = XPrt.CurrentX + 900
XPrt.Print "Solde Compte";
XPrt.CurrentX = XPrt.CurrentX + 900
XPrt.Print "Cumul emploi";
XPrt.CurrentX = XPrt.CurrentX + 900
XPrt.Print "Différence";

XPrt.CurrentY = prtMinY + prtHeaderHeight - XPrt.TextHeight("test")
XPrt.FontBold = False


End Sub


'---------------------------------------------------------
Public Sub prtLrBafiMsgLine()
'---------------------------------------------------------
Dim X As String, K As Integer, mCurrenty As Integer

If XPrt.CurrentY + prtlineHeight * 2 > prtMaxY Then
    frmElpPrt.prtNewPage
    prtLrBafiMsgForm
End If

'------------------------------------------ligne 1--------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'----------------------------------ligne2

XPrt.CurrentX = prtMinX

XPrt.Print recLine;
End Sub



