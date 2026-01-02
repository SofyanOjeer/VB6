Attribute VB_Name = "prtCptComPays"
Option Explicit

Dim mCol As Integer, mCol1 As Integer, mCol2 As Integer, mCol3 As Integer
'---------------------------------------------------------
 Public Sub prtCptComPays_Open(Msg As String)
'---------------------------------------------------------
Dim X As String

On Error GoTo prtError

Set XPrt = Printer
'Height8_6 = frmElpPrt.prtHeightDelta(8, 6)
prtTitleText = "Commission : code pays " 'lEnTête

prtLineNb = 1

frmElpPrt.Show vbModeless
prtlineHeight = 350
prtHeaderHeight = 350

prtOrientation = vbPRORPortrait

prtPgmName = "prtCptComPays"
prtTitleUsr = usrName
frmElpPrt.prtStdInit

mCol1 = prtMinX + 50
mCol2 = prtMinX + 3550
mCol3 = prtMinX + 7050

prtCptComPays_Form
Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide
End Sub

'---------------------------------------------------------
 Public Sub prtCptComPays_Close()
'---------------------------------------------------------

On Error GoTo prtError

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
Public Sub prtCptComPays_Line(lX1 As String, lX2 As String)
'---------------------------------------------------------

If XPrt.CurrentY + prtlineHeight * 1.9 > prtMaxY Then
    XPrt.CurrentY = prtMinY + prtHeaderHeight + 50
    Select Case mCol
        Case mCol1: mCol = mCol2
        Case mCol2: mCol = mCol3
        Case Else: frmElpPrt.prtNewPage: prtCptComPays_Form
    End Select
End If


XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = mCol: XPrt.Print lX1 & "      " & lX2;

End Sub

'---------------------------------------------------------
Public Sub prtCptComPays_Form()
'---------------------------------------------------------
Dim X As String, K As Integer
XPrt.FontSize = 8

XPrt.FontBold = True
XPrt.DrawWidth = 3

Call frmElpPrt.prtTrame(prtMinX, prtMinY, prtMaxX, prtMinY + prtHeaderHeight, " ", 230)
XPrt.Line (prtMinX, prtMinY)-(prtMaxX, prtMinY)

XPrt.CurrentY = prtMinY + 50
XPrt.CurrentX = prtMinX + 400: XPrt.Print "Dossier";
mCol = mCol1
XPrt.CurrentY = prtMinY + prtHeaderHeight + 50
End Sub


