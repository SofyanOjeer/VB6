Attribute VB_Name = "prtSAB_Dossier"
Option Explicit
Dim mForm As String
Dim X As String, I As Integer, Height8_6 As Integer
Dim curX As Currency, curX1 As Currency, curX2 As Currency
Dim prtMaxY_YDOSSLD0 As Long, prtMinY_YDOSSLD0 As Long

Dim mAMJ_7Past As String, mAMJ_7Ante As String

Public Sub prtYDOSSLD0_Open()
On Error GoTo prtError

Set XPrt = Printer
frmElpPrt.Show vbModeless

Height8_6 = frmElpPrt.prtHeightDelta(8, 6)

blnFiligrane = False
 prtOrientation = vbPRORLandscape '
prtPgmName = "prtYDOSSLD0"
prtTitleUsr = usrName

prtLineNb = 1
prtlineHeight = 250
prtHeaderHeight = 300


prtFormType = ""
frmElpPrt.prtStdInit

prtFontName = prtFontName_Arial
Select Case mForm
    Case "YDOSSLD0": prtYDOSSLD0_Form
End Select
XPrt.CurrentY = XPrt.CurrentY - prtlineHeight
Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub

Public Sub prtYDOSSLD0_Close(blnEnd As Boolean)
On Error GoTo prtError

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight

    prtMaxY_YDOSSLD0 = XPrt.CurrentY
    Select Case mForm
        Case "YDOSSLD0"
                prtYDOSSLD0_Col
    End Select
    XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY), prtLineColor
If blnEnd Then
    frmElpPrt.prtEndDoc 1000
    frmElpPrt.Hide
End If
Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub


Public Sub prtYDOSSLD0_NewLine()
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
If XPrt.CurrentY + 300 > prtMaxY Then
    prtMaxY_YDOSSLD0 = XPrt.CurrentY
    Select Case mForm
        Case "YDOSSLD0":
                prtYDOSSLD0_Col
                frmElpPrt.prtNewPage
                prtYDOSSLD0_Form
    End Select
End If

End Sub




'---------------------------------------------------------
Public Sub prtYDOSSLD0_Col()
'---------------------------------------------------------

XPrt.DrawWidth = 1
prtLineColor = vbCyan
XPrt.Line (prtMinX + 1950, prtMinY_YDOSSLD0)-(prtMinX + 1950, prtMaxY_YDOSSLD0), prtLineColor
XPrt.Line (prtMinX + 7550, prtMinY_YDOSSLD0)-(prtMinX + 7550, prtMaxY_YDOSSLD0), prtLineColor
XPrt.Line (prtMinX + 13450, prtMinY_YDOSSLD0)-(prtMinX + 13450, prtMaxY_YDOSSLD0), prtLineColor
End Sub

Public Sub prtYDOSSLD0_Form()
Dim wId As String
Dim X As String

XPrt.FontSize = 7
XPrt.FontBold = True
XPrt.DrawWidth = 2

XPrt.FontSize = 8
prtFillColor = vbCyan
XPrt.ForeColor = vbWhite
Call frmElpPrt.prtTrame_Color(prtMinX, prtMinY, prtMaxX, prtMinY + prtHeaderHeight, " ")
'---------------------------------------------------------
prtFillColor = RGB(200, 255, 255)

XPrt.CurrentY = prtMinY + 50

XPrt.CurrentX = prtMinX: XPrt.Print "D. compta";
XPrt.CurrentX = prtMinX + 1000: XPrt.Print "Service   ";
XPrt.CurrentX = prtMinX + 2000: XPrt.Print "Compte débité";
XPrt.CurrentX = prtMinX + 3500: XPrt.Print "Intitulé";
XPrt.CurrentX = prtMinX + 6800: XPrt.Print "Montant";
XPrt.CurrentX = prtMinX + 7600: XPrt.Print "n° chèque";
XPrt.CurrentX = prtMinX + 8600: XPrt.Print "Bénéficiaire";
'XPrt.CurrentX = prtMinX + 10800: XPrt.Print "Bq créditée";
XPrt.CurrentX = prtMinX + 11500: XPrt.Print "Archivage interne";
XPrt.CurrentX = prtMinX + 13500: XPrt.Print "Numérisation: date jpg";
'XPrt.CurrentX = prtMinX + 14400: XPrt.Print "référence";
'XPrt.CurrentX = prtMinX + 15200: XPrt.Print "MSEL";
XPrt.CurrentX = prtMinX + 15600: XPrt.Print "Id";

'XPrt.CurrentY = XPrt.CurrentY + prtlineHeight

'XPrt.CurrentX = prtMinX + 1000: XPrt.Print "Opération   ";
'XPrt.CurrentX = prtMinX + 2000: XPrt.Print "Client";
'XPrt.CurrentX = prtMinX + 3500: XPrt.Print "motif économique";
'XPrt.CurrentX = prtMinX + 7600: XPrt.Print "d. émission";
'XPrt.CurrentX = prtMinX + 8600: XPrt.Print "Bq créditée / Compte";
'XPrt.CurrentX = prtMinX + 11500: XPrt.Print "Archivage externe";
'XPrt.CurrentX = prtMinX + 13500: XPrt.Print "réf remise";
''XPrt.CurrentX = prtMinX + 15600: XPrt.Print "sta";

XPrt.FontSize = 8
XPrt.FontBold = False
XPrt.ForeColor = vbBlack
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY), prtLineColor
XPrt.CurrentY = XPrt.CurrentY + 50
prtMinY_YDOSSLD0 = prtMinY + prtHeaderHeight '* 2


End Sub
Public Sub prtYDOSSLD0_Line(lYDOSSLD0 As typeYDOSSLD0)
Dim X As String, xSql As String
Dim wCLIENARA1 As String
Dim wColor As Long
prtYDOSSLD0_NewLine



End Sub




