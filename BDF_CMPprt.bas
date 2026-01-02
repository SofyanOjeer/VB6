Attribute VB_Name = "prtBDF_CMP"
Option Explicit


Public Sub prtBDF_CMP_Open(lK As Integer, lText As String)
On Error GoTo prtError

Set XPrt = Printer
frmElpPrt.Show vbModeless

'Height8_6 = frmElpPrt.prtHeightDelta(8, 6)

blnFiligrane = False
If lK = 3 Or lK = 2 Then
    prtOrientation = vbPRORPortrait '
Else
    prtOrientation = vbPRORLandscape '
End If
prtPgmName = "prtBDF_CMP"
prtTitleUsr = usrName
prtTitleText = lText
prtLineNb = 1
prtlineHeight = 300
prtHeaderHeight = 300


prtFormType = ""
frmElpPrt.prtStdInit
    Select Case lK
        Case 3: prtBDF_CMP_Form_3
        Case 4: prtBDF_CMP_Form_4
    End Select
Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub

'---------------------------------------------------------
Public Sub prtBDF_CMP_Form_3()
'---------------------------------------------------------
Dim X As String

XPrt.DrawWidth = 2
XPrt.FontSize = 7: XPrt.FontBold = True

Call frmElpPrt.prtTrame(prtMinX, prtMinY, prtMaxX, prtMinY + prtHeaderHeight, "B", 240)
'---------------------------------------------------------
XPrt.CurrentY = prtMinY + 50

XPrt.CurrentX = prtMinX + 50: XPrt.Print "type";
XPrt.CurrentY = prtMinY + 50 + prtHeaderHeight - prtlineHeight

XPrt.FontBold = False

End Sub
'---------------------------------------------------------
'---------------------------------------------------------
Public Sub prtBDF_CMP_Form_4()
'---------------------------------------------------------
Dim X As String

XPrt.DrawWidth = 2
XPrt.FontSize = 9: XPrt.FontBold = True

Call frmElpPrt.prtTrame(prtMinX, prtMinY, prtMaxX, prtMinY + prtHeaderHeight, "B", 240)
'---------------------------------------------------------
XPrt.CurrentY = prtMinY + 50

XPrt.CurrentX = prtMinX + 50: XPrt.Print "type";
XPrt.CurrentX = prtMinX + 3000: XPrt.Print "virements nationaux";
XPrt.CurrentX = prtMinX + 7500: XPrt.Print "virements union européenne";
XPrt.CurrentX = prtMinX + 13000: XPrt.Print "virements internationaux";

XPrt.CurrentY = prtMinY + 50 + prtHeaderHeight - prtlineHeight

XPrt.FontBold = False

End Sub
'---------------------------------------------------------


'---------------------------------------------------------
Public Sub prtBDF_CMP_Form_2()
'---------------------------------------------------------
Dim X As String

XPrt.DrawWidth = 2
XPrt.FontSize = 7: XPrt.FontBold = True

Call frmElpPrt.prtTrame(prtMinX, prtMinY, prtMaxX, prtMinY + prtHeaderHeight, "B", 240)
'---------------------------------------------------------
XPrt.CurrentY = prtMinY + 50

XPrt.CurrentX = prtMinX + 50: XPrt.Print "type";
XPrt.CurrentY = prtMinY + 50 + prtHeaderHeight - prtlineHeight

XPrt.FontBold = False

End Sub
'---------------------------------------------------------


'---------------------------------------------------------
Public Sub prtBDF_CMP_Form_3_Col(lMaxY)
'---------------------------------------------------------
Dim X As String, K As Integer, K2 As Integer

XPrt.DrawWidth = 2

XPrt.Line (prtMinX + 2100, prtMinY)-(prtMinX + 2100, lMaxY), prtLineColor
XPrt.Line (prtMinX + 5100, prtMinY)-(prtMinX + 5100, lMaxY), prtLineColor
XPrt.Line (prtMinX + 8100, prtMinY)-(prtMinX + 8100, lMaxY), prtLineColor
XPrt.Line (prtMaxX, prtMinY)-(prtMaxX, lMaxY), prtLineColor


End Sub
'---------------------------------------------------------
Public Sub prtBDF_CMP_Form_4_Col(lMaxY)
'---------------------------------------------------------
Dim X As String, K As Integer, K2 As Integer

XPrt.DrawWidth = 2

XPrt.Line (prtMinX + 1000, prtMinY)-(prtMinX + 1000, lMaxY), prtLineColor
XPrt.Line (prtMinX + 6100, prtMinY)-(prtMinX + 6100, lMaxY), prtLineColor
XPrt.Line (prtMinX + 11100, prtMinY)-(prtMinX + 11100, lMaxY), prtLineColor
XPrt.Line (prtMaxX, prtMinY)-(prtMaxX, lMaxY), prtLineColor


End Sub

'---------------------------------------------------------
Public Sub prtBDF_CMP_Form_2_Col(lMaxY)
'---------------------------------------------------------
Dim X As String, K As Integer, K2 As Integer

XPrt.DrawWidth = 2

XPrt.Line (prtMinX + 4100, prtMinY)-(prtMinX + 4100, lMaxY), prtLineColor
XPrt.Line (prtMinX + 7100, prtMinY)-(prtMinX + 7100, lMaxY), prtLineColor
XPrt.Line (prtMinX + 9100, prtMinY)-(prtMinX + 9100, lMaxY), prtLineColor
XPrt.Line (prtMaxX, prtMinY)-(prtMaxX, lMaxY), prtLineColor


End Sub

Public Sub prtBDF_CMP_NewLine(lK As Integer)
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
If XPrt.CurrentY + 300 > prtMaxY Then
    Select Case lK
        Case 2: prtBDF_CMP_Form_2_Col (prtMaxY)
        Case 3: prtBDF_CMP_Form_3_Col (prtMaxY)
        Case 4: prtBDF_CMP_Form_4_Col (prtMaxY)
    End Select
    frmElpPrt.prtNewPage
    Select Case lK
        Case 2: prtBDF_CMP_Form_2
        Case 3: prtBDF_CMP_Form_3
        Case 4: prtBDF_CMP_Form_4
        
    End Select
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight

End If

End Sub

Public Sub prtBDF_CMP_Close(lK As Integer)
Dim X As String
On Error GoTo prtError
XPrt.DrawWidth = 5
prtBDF_CMP_NewLine lK
    Select Case lK
        Case 3: prtBDF_CMP_Form_3_Col (XPrt.CurrentY)
        Case 2: prtBDF_CMP_Form_2_Col (XPrt.CurrentY)
        Case 4: prtBDF_CMP_Form_4_Col (XPrt.CurrentY)
    End Select

XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY), prtLineColor

frmElpPrt.prtEndDoc 1000
frmElpPrt.Hide
Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub




