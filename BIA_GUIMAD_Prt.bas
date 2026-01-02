Attribute VB_Name = "prtBIA_GUIMAD"
Option Explicit

Dim wMM As Integer, wAAAA As Integer

Public Sub prtBIA_GUIMAD_Open(lK As Integer, lText As String, lMM As Integer, lAAAA As Integer)
On Error GoTo prtError

wMM = lMM
wAAAA = lAAAA
Set XPrt = Printer
frmElpPrt.Show vbModeless

'Height8_6 = frmElpPrt.prtHeightDelta(8, 6)

blnFiligrane = False
If lK = 7 Then
    prtOrientation = vbPRORPortrait '
Else
    prtOrientation = vbPRORLandscape '
End If
prtPgmName = "prtBIA_GUIMAD"
prtTitleUsr = usrName
prtTitleText = lText
prtLineNb = 1
prtlineHeight = 300
prtHeaderHeight = 300


prtFormType = ""
frmElpPrt.prtStdInit
    Select Case lK
        Case 1: prtBIA_GUIMAD_Form_1
        Case 6: prtBIA_GUIMAD_Form_6
        Case 7: prtBIA_GUIMAD_Form_7
    End Select
Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub

'---------------------------------------------------------
Public Sub prtBIA_GUIMAD_Form_1()
'---------------------------------------------------------
Dim X As String

XPrt.DrawWidth = 2
XPrt.FontSize = 7: XPrt.FontBold = True

Call frmElpPrt.prtTrame(prtMinX, prtMinY, prtMaxX, prtMinY + prtHeaderHeight, "B", 240)
'---------------------------------------------------------
XPrt.CurrentY = prtMinY + 50

XPrt.CurrentX = prtMinX + 50: XPrt.Print "Date";
XPrt.CurrentX = prtMinX + 1050: XPrt.Print "Opération";

XPrt.CurrentX = prtMinX + 3300: XPrt.Print "Montant";
XPrt.CurrentX = prtMinX + 4200: XPrt.Print "Client";
XPrt.CurrentX = prtMinX + 8000: XPrt.Print "Bénéficiaire / pour compte de";
XPrt.CurrentX = prtMinX + 11000 + 100: XPrt.Print "Motif";
XPrt.CurrentY = prtMinY + 50 + prtHeaderHeight - prtlineHeight

XPrt.FontBold = False

End Sub
'---------------------------------------------------------
Public Sub prtBIA_GUIMAD_Form_7()
'---------------------------------------------------------
Dim X As String

XPrt.DrawWidth = 2
XPrt.FontSize = 7: XPrt.FontBold = True

Call frmElpPrt.prtTrame(prtMinX, prtMinY, prtMaxX, prtMinY + prtHeaderHeight, "B", 240)
'---------------------------------------------------------
XPrt.CurrentY = prtMinY + 50

XPrt.CurrentX = prtMinX + 50: XPrt.Print "Bénéficiaire";
XPrt.CurrentX = prtMinX + 5000: XPrt.Print "Nb d'opérations";

XPrt.CurrentX = prtMinX + 8000: XPrt.Print "Total EUR";
XPrt.CurrentY = prtMinY + 50 + prtHeaderHeight - prtlineHeight

XPrt.FontBold = False

End Sub

'---------------------------------------------------------
Public Sub prtBIA_GUIMAD_Form_6()
'---------------------------------------------------------
Dim X As String, K As Integer, K2 As Integer
Dim xMM As Integer, xAAAA As Integer

XPrt.DrawWidth = 2
XPrt.FontSize = 7: XPrt.FontBold = True

Call frmElpPrt.prtTrame(prtMinX, prtMinY, prtMaxX, prtMinY + prtHeaderHeight, "B", 240)

'---------------------------------------------------------
XPrt.CurrentY = prtMinY + 50

XPrt.CurrentX = prtMinX + 350: XPrt.Print "Client";
XPrt.CurrentX = prtMinX + 50: XPrt.Print "Dev";
XPrt.CurrentX = prtMinX + 15000: XPrt.Print "Total";
xMM = wMM
xAAAA = wAAAA
For K = 1 To 12
    XPrt.CurrentX = prtMaxX - (14 - K) * 900 + 200: XPrt.Print Format$(xMM, "00") & "." & xAAAA;
    xMM = xMM + 1
    If xMM = 13 Then xMM = 1: xAAAA = xAAAA + 1
Next K



XPrt.CurrentY = prtMinY + 50 + prtHeaderHeight - prtlineHeight

XPrt.FontBold = False

End Sub

'---------------------------------------------------------
Public Sub prtBIA_GUIMAD_Form_6_Col(lMaxY)
'---------------------------------------------------------
Dim X As String, K As Integer, K2 As Integer

XPrt.DrawWidth = 2
For K = 1 To 14
    K2 = prtMaxX - K * 900 + 900
    XPrt.Line (K2, prtMinY)-(K2, lMaxY), prtLineColor
Next K



End Sub

'---------------------------------------------------------
Public Sub prtBIA_GUIMAD_Form_1_Col(lMaxY)
'---------------------------------------------------------
Dim X As String, K As Integer, K2 As Integer

XPrt.DrawWidth = 2

XPrt.Line (prtMinX + 2100, prtMinY)-(prtMinX + 2100, lMaxY), prtLineColor
XPrt.Line (prtMinX + 4100, prtMinY)-(prtMinX + 4100, lMaxY), prtLineColor
'XPrt.Line (prtMinX + 13000, prtMinY)-(prtMinX + 13000, lMaxY), prtLineColor


End Sub

Public Sub prtBIA_GUIMAD_NewLine(lK As Integer)
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
If XPrt.CurrentY + 300 > prtMaxY Then
    Select Case lK
        Case 1: prtBIA_GUIMAD_Form_1_Col (prtMaxY)
        Case 6: prtBIA_GUIMAD_Form_6_Col (prtMaxY)
    End Select
    frmElpPrt.prtNewPage
    Select Case lK
        Case 1: prtBIA_GUIMAD_Form_1
        Case 6:  prtBIA_GUIMAD_Form_6
    End Select
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight

End If

End Sub

Public Sub prtBIA_GUIMAD_Close(lK As Integer)
Dim X As String
On Error GoTo prtError
XPrt.DrawWidth = 5
prtBIA_GUIMAD_NewLine lK
    Select Case lK
        Case 1: prtBIA_GUIMAD_Form_1_Col (XPrt.CurrentY)
        Case 6: prtBIA_GUIMAD_Form_6_Col (XPrt.CurrentY)
    End Select

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



