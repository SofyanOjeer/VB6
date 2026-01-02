Attribute VB_Name = "prtSAB_TC_Limites"
Option Explicit
Dim mFct1 As String

Dim X As String, I As Integer, Height8_6 As Integer
Dim curX As Currency, curX1 As Currency, curX2 As Currency

Dim blnPage As Boolean
Dim xZAUTSYC0  As typeZAUTSYC0

Dim wAmj1_8C As String, wAmjMax_8C As String
Public Sub prtSAB_TC_Lmites_NewLine_xlsManual(ByRef currentRow As Long, ByRef wsExcel As Excel.Worksheet, ByRef comptageRows As Long, maxRows As Long, maxRowsPlus As Long)

    If currentRow >= maxRows + maxRowsPlus Then
        If comptageRows >= maxRows Then
            Call insere_entete_page(wsExcel, "1:3", 3, currentRow)
            comptageRows = 3
            currentRow = currentRow + 3
        End If
    End If
    comptageRows = comptageRows + 1
    currentRow = currentRow + 1

End Sub

Public Sub prtSAB_TC_Lmites_Open(lText As String, lAmj1_8C As String, lAmjMax_8C As String)
On Error GoTo prtError

wAmj1_8C = lAmj1_8C
wAmjMax_8C = lAmjMax_8C

Set XPrt = Printer
frmElpPrt.Show vbModeless

Height8_6 = frmElpPrt.prtHeightDelta(8, 6)

blnFiligrane = False
 prtOrientation = vbPRORLandscape '
prtPgmName = "prtSAB_TC_Lmites"
prtTitleUsr = usrName
prtTitleText = "Trésorerie : Etat des limites " & lText

prtLineNb = 1
prtlineHeight = 250
prtHeaderHeight = 300


prtFormType = ""
frmElpPrt.prtStdInit

prtFontName = prtFontName_Arial
prtSAB_TC_Lmites_Form
XPrt.CurrentY = XPrt.CurrentY - prtlineHeight
Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub

Public Sub prtSAB_TC_Lmites_Close()
On Error GoTo prtError


Call frmElpPrt.prtEndDoc(1000)
frmElpPrt.Hide
Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub


Public Sub prtSAB_TC_Lmites_NewLine()
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
If XPrt.CurrentY + 300 > prtMaxY Then
    frmElpPrt.prtNewPage
    prtSAB_TC_Lmites_Form
End If

End Sub



Public Sub prtSAB_TC_Lmites_Form()
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

XPrt.CurrentX = prtMinX
XPrt.Print "Abrégé";
XPrt.CurrentX = prtMinX + 1300: XPrt.Print "Racine";
XPrt.CurrentX = prtMinX + 2000: XPrt.Print "Intitulé  /  Dossier";
'XPrt.CurrentX = prtMinX + 4000: XPrt.Print "Date nég";
XPrt.CurrentX = prtMinX + 4000: XPrt.Print "Date MAD";
XPrt.CurrentX = prtMinX + 5000: XPrt.Print "Echéance";
X = "Autorisation / Opé"
XPrt.CurrentX = prtMinX + 7300 - XPrt.TextWidth(X): XPrt.Print X;
XPrt.CurrentX = prtMinX + 7500: XPrt.Print "Aut";
XPrt.CurrentX = prtMinX + 8000: XPrt.Print ">>>> MAD";
XPrt.CurrentX = prtMinX + 9000: XPrt.Print "%";
XPrt.CurrentX = prtMinX + 9400: XPrt.Print "Autorisation € %";
XPrt.CurrentX = prtMinX + 10800: XPrt.Print "Encours €";
'XPrt.CurrentX = prtMinX + 11850: XPrt.Print "Dépassement € %";
'XPrt.CurrentX = prtMinX + 11900: XPrt.Print "Disponible (%)";
XPrt.CurrentX = prtMinX + 12400: XPrt.Print "Today";
XPrt.CurrentX = prtMinX + 13600: XPrt.Print "Tom next";
XPrt.CurrentX = prtMinX + 15200: XPrt.Print "Spot";


XPrt.FontSize = 8
XPrt.FontBold = False
XPrt.ForeColor = vbBlack
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY), prtLineColor
XPrt.CurrentY = XPrt.CurrentY + 50


End Sub






Public Sub prtSAB_TC_Lmites_Open_xlsManual(lText As String, lAmj1_8C As String, lAmjMax_8C As String, wsExcel As Excel.Worksheet)

    wAmj1_8C = lAmj1_8C
    wAmjMax_8C = lAmjMax_8C
    prtTitleText = "Trésorerie : Etat des limites " & lText
    wsExcel.Cells(1, 4) = prtTitleText

End Sub


