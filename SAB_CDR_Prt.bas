Attribute VB_Name = "prtSAB_CDR"
Option Explicit

Public Sub prtSAB_CDR_Open(lText As String)
On Error GoTo prtError

Set XPrt = Printer
frmElpPrt.Show vbModeless

'Height8_6 = frmElpPrt.prtHeightDelta(8, 6)

blnFiligrane = False
prtOrientation = vbPRORPortrait '
prtPgmName = "prtSAB_CDR"
prtTitleUsr = usrName
prtTitleText = lText
prtLineNb = 1
prtlineHeight = 300
prtHeaderHeight = 300


prtFormType = ""
frmElpPrt.prtStdInit
prtSAB_CDR_Form
Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub

'---------------------------------------------------------
Public Sub prtSAB_CDR_Form()
'---------------------------------------------------------
Dim X As String

XPrt.DrawWidth = 2
XPrt.FontSize = 7: XPrt.FontBold = True

Call frmElpPrt.prtTrame(prtMinX, prtMinY, prtMaxX, prtMinY + prtHeaderHeight, "B", 240)
XPrt.Line (prtMinX + 500 - 50, prtMinY)-(prtMinX + 500 - 50, prtMaxY), prtLineColor
XPrt.Line (prtMinX + 2000 - 50, prtMinY)-(prtMinX + 2000 - 50, prtMaxY), prtLineColor
XPrt.Line (prtMinX + 6000 - 50, prtMinY)-(prtMinX + 6000 - 50, prtMaxY), prtLineColor
'---------------------------------------------------------
XPrt.CurrentY = prtMinY + 50

XPrt.CurrentX = prtMinX + 50: XPrt.Print "Table";
XPrt.CurrentX = prtMinX + 500: XPrt.Print "Identification";

XPrt.CurrentX = prtMinX + 2000: XPrt.Print "Paramètres";
XPrt.CurrentX = prtMinX + 6000: XPrt.Print "Intitulé";
XPrt.CurrentY = prtMinY + 50 + prtHeaderHeight

XPrt.FontBold = False

End Sub
Public Sub prtSAB_CDR_NewLine()
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
If XPrt.CurrentY + 300 > prtMaxY Then
    frmElpPrt.prtNewPage
    prtSAB_CDR_Form
End If

End Sub

Public Sub prtSAB_CDR_Close()
Dim X As String
On Error GoTo prtError

frmElpPrt.prtEndDoc 1000
frmElpPrt.Hide
Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub


