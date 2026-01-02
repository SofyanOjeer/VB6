Attribute VB_Name = "prtSAB_CPTMVT"
Option Explicit

Public Sub prtSAB_CPTMVT_Open(lText As String)
On Error GoTo prtError

Set XPrt = Printer
frmElpPrt.Show vbModeless

'Height8_6 = frmElpPrt.prtHeightDelta(8, 6)

blnFiligrane = False
prtOrientation = vbPRORPortrait '
prtPgmName = "prtSAB_CPTMVT"
prtTitleUsr = usrName
prtTitleText = lText
prtLineNb = 1
prtlineHeight = 300
prtHeaderHeight = 300


prtFormType = ""
frmElpPrt.prtStdInit
prtSAB_CPTMVT_Form
Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub

'---------------------------------------------------------
Public Sub prtSAB_CPTMVT_Form()
'---------------------------------------------------------
Dim X As String

XPrt.DrawWidth = 2
XPrt.FontSize = 7: XPrt.FontBold = True

Call frmElpPrt.prtTrame(prtMinX, prtMinY, prtMaxX, prtMinY + prtHeaderHeight, "B", 240)
XPrt.Line (prtMinX + 6000 - 50, prtMinY)-(prtMinX + 6000 - 50, prtMaxY), prtLineColor
XPrt.Line (prtMinX + 8000 - 50, prtMinY)-(prtMinX + 8000 - 50, prtMaxY), prtLineColor
XPrt.Line (prtMinX + 10000 - 50, prtMinY)-(prtMinX + 10000 - 50, prtMaxY), prtLineColor
'---------------------------------------------------------
XPrt.CurrentY = prtMinY + 50

XPrt.CurrentX = prtMinX + 50: XPrt.Print "Compte";
XPrt.CurrentX = prtMinX + 2000: XPrt.Print "Intitulé";

XPrt.CurrentX = prtMinX + 5000 + 300: XPrt.Print "Dossier";
XPrt.CurrentX = prtMinX + 6000 + 1200: XPrt.Print "Solde Db";
XPrt.CurrentX = prtMinX + 8000 + 1200: XPrt.Print "Solde CR";
XPrt.CurrentX = prtMinX + 10000 + 100: XPrt.Print "Date Rbt";

XPrt.CurrentY = prtMinY + 50 + prtHeaderHeight - prtlineHeight

XPrt.FontBold = False

End Sub
Public Sub prtSAB_CPTMVT_NewLine()
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
If XPrt.CurrentY + 300 > prtMaxY Then
    frmElpPrt.prtNewPage
    prtSAB_CPTMVT_Form
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight

End If

End Sub

Public Sub prtSAB_CPTMVT_Close()
Dim X As String
On Error GoTo prtError
XPrt.DrawWidth = 5
prtSAB_CPTMVT_NewLine
XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY), prtLineColor

Call frmElpPrt.prtEndDoc(1000)
frmElpPrt.Hide
Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub


