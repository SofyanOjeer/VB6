Attribute VB_Name = "prtYEICGCC0"
Option Explicit
Dim mForm As String
Dim X As String, I As Integer, Height8_6 As Integer
Dim curX As Currency, curX1 As Currency, curX2 As Currency
Dim prtMaxY_YEICGCC0 As Long, prtMinY_YEICGCC0 As Long

Dim mAMJ_7Past As String, mAMJ_7Ante As String

Public Sub prtYEICGCC0_Line_Echéancier_xlsManual(lYEICGCC0 As typeYEICGCC0, lYEICGCCLOG As typeYEICGCCLOG, ByRef currentRow As Long, wsExcel As Excel.Worksheet, ByRef comptageRows As Long, maxRows As Long, maxRowsPlus As Long)
Dim X As String, xSQL As String
Dim wCLIENARA1 As String
Dim wColor As Long

If currentRow >= maxRows + maxRowsPlus Then
    If comptageRows >= maxRows Then
        Call insere_entete_page(wsExcel, "1:3", 3, currentRow)
        comptageRows = 3
        currentRow = currentRow + 3
    End If
End If
comptageRows = comptageRows + 1
currentRow = currentRow + 1
wsExcel.Activate
Range("4:4").Select
Selection.Copy
Range("A" & CStr(currentRow)).Select
ActiveSheet.Paste
wColor = vbMagenta
wsExcel.Cells(currentRow, 1) = "'" & dateImp10(lYEICGCC0.EICGCCAMJ)
wsExcel.Cells(currentRow, 1).Font.Color = wColor
wsExcel.Cells(currentRow, 2) = lYEICGCC0.EICGCCOPE & "  " & lYEICGCC0.EICGCCDOS
wsExcel.Cells(currentRow, 2).Font.Color = wColor
wColor = RGB(0, 0, 128)
If IsNumeric(lYEICGCC0.EICGCCECPT) Then
    wsExcel.Cells(currentRow, 3) = Format(lYEICGCC0.EICGCCECPT, "@@@@@ @@@ @@@ @@@@@@@")
Else
    wsExcel.Cells(currentRow, 3) = lYEICGCC0.EICGCCECPT
End If
wsExcel.Cells(currentRow, 3).Font.Color = wColor
xSQL = "select CLIENARA1, COMPTEINT from " & paramIBM_Library_SABSPE & ".YBIACPT0 " _
     & " Where COMPTECOM = '" & lYEICGCC0.EICGCCECPT & "'"
Set rsSab = cnsab.Execute(xSQL)
If Not rsSab.EOF Then
    wCLIENARA1 = rsSab("CLIENARA1")
    wsExcel.Cells(currentRow, 4) = rsSab("COMPTEINT")
    wsExcel.Cells(currentRow, 4).Font.Color = wColor
Else
    wCLIENARA1 = "?"
End If
'______________________________________________________________________________________
If lYEICGCC0.EICGCCEIND <> " " Then
    wsExcel.Cells(currentRow, 7) = "/" & lYEICGCC0.EICGCCEIND
    wsExcel.Cells(currentRow, 7).Font.Color = wColor
End If

X = Format$(lYEICGCC0.EICGCCID, "### ### ### ##0")
wsExcel.Cells(currentRow, 10) = X
wsExcel.Cells(currentRow, 10).Font.Color = wColor
'___________________________________________________________________________________

X = Format$(lYEICGCC0.EICGCCEMT, "### ### ### ##0.00")
wsExcel.Cells(currentRow, 5) = X
wsExcel.Cells(currentRow, 5).Font.Color = wColor
wsExcel.Cells(currentRow, 6) = lYEICGCC0.EICGCCECHQ
wsExcel.Cells(currentRow, 6).Font.Color = wColor
wsExcel.Cells(currentRow, 8) = lYEICGCCLOG.EICGCCLOGK
wsExcel.Cells(currentRow, 8).Font.Color = wColor
wsExcel.Cells(currentRow, 9) = lYEICGCCLOG.EICGCCLOGX
wsExcel.Cells(currentRow, 9).Font.Color = wColor

End Sub

Public Sub prtYEICGCC0_Line_xlsManual(lYEICGCC0 As typeYEICGCC0, ByRef currentRow As Long, wsExcel As Excel.Worksheet, ByRef comptageRows As Long, maxRows As Long, maxRowsPlus As Long)
Dim X As String, xSQL As String
Dim wCLIENARA1 As String
Dim wColor As Long

If currentRow >= maxRows + maxRowsPlus Then
    If comptageRows >= maxRows Then
        Call insere_entete_page(wsExcel, "1:3", 3, currentRow)
        comptageRows = 3
        currentRow = currentRow + 3
    End If
End If
comptageRows = comptageRows + 1
currentRow = currentRow + 1
wsExcel.Activate
Range("4:4").Select
Selection.Copy
Range("A" & CStr(currentRow)).Select
ActiveSheet.Paste

wColor = vbBlack
Select Case lYEICGCC0.EICGCCSTA
    Case "I", "A", "R": wColor = RGB(128, 128, 128)
    Case "V", "@": wColor = RGB(64, 128, 64)
    Case Else
        If lYEICGCC0.EICGCCDOS = 0 Then
            wColor = vbRed
        Else
            If lYEICGCC0.EICGCCVJPG = 0 Then
                If lYEICGCC0.EICGCCAMJ < mAMJ_7Ante Then
                    wColor = vbRed
                Else
                    wColor = vbMagenta
                End If
            Else
                Select Case lYEICGCC0.EICGCCSTAK
                    Case "X": wColor = vbMagenta
                    Case Else: wColor = vbBlue
                End Select
            End If
        End If
End Select
wsExcel.Cells(currentRow, 1) = dateImp10(lYEICGCC0.EICGCCAMJ)
wsExcel.Cells(currentRow, 1).Font.Color = wColor
wsExcel.Cells(currentRow, 2) = lYEICGCC0.EICGCCOPE & "  " & lYEICGCC0.EICGCCDOS
wsExcel.Cells(currentRow, 2).Font.Color = wColor
wColor = RGB(0, 0, 128)
If IsNumeric(lYEICGCC0.EICGCCECPT) Then
    wsExcel.Cells(currentRow, 3) = Format(lYEICGCC0.EICGCCECPT, "@@@@@ @@@ @@@ @@@@@@@")
Else
    wsExcel.Cells(currentRow, 3) = lYEICGCC0.EICGCCECPT
End If
wsExcel.Cells(currentRow, 3).Font.Color = wColor
xSQL = "select CLIENARA1, COMPTEINT from " & paramIBM_Library_SABSPE & ".YBIACPT0 " _
     & " Where COMPTECOM = '" & lYEICGCC0.EICGCCECPT & "'"
Set rsSab = cnsab.Execute(xSQL)
If Not rsSab.EOF Then
    wCLIENARA1 = rsSab("CLIENARA1")
    wsExcel.Cells(currentRow, 4) = rsSab("COMPTEINT")
    wsExcel.Cells(currentRow, 4).Font.Color = wColor
Else
    wCLIENARA1 = "?"
End If
X = Format$(lYEICGCC0.EICGCCEMT, "### ### ### ##0.00")
wsExcel.Cells(currentRow, 5) = X
wsExcel.Cells(currentRow, 5).Font.Color = wColor
'______________________________________________________________________________________
wsExcel.Cells(currentRow, 6) = lYEICGCC0.EICGCCECHQ
wsExcel.Cells(currentRow, 6).Font.Color = wColor
If lYEICGCC0.EICGCCEIND <> " " Then
    wsExcel.Cells(currentRow, 7) = "/" & lYEICGCC0.EICGCCEIND
    wsExcel.Cells(currentRow, 7).Font.Color = wColor
End If
wsExcel.Cells(currentRow, 8) = lYEICGCC0.EICGCCXNOM
wsExcel.Cells(currentRow, 8).Font.Color = wColor
If lYEICGCC0.EICGCCVAMJ > 0 Then
    wsExcel.Cells(currentRow, 10) = dateImp10(lYEICGCC0.EICGCCVAMJ) & " " & lYEICGCC0.EICGCCVJPG
    wsExcel.Cells(currentRow, 10).Font.Color = wColor
End If
X = Format$(lYEICGCC0.EICGCCID, "### ### ### ##0")
wsExcel.Cells(currentRow, 11) = X
wsExcel.Cells(currentRow, 11).Font.Color = wColor
'___________________________________________________________________________________

End Sub

Public Sub prtYEICGCC0_Open()
On Error GoTo prtError

Set XPrt = Printer
frmElpPrt.Show vbModeless

Height8_6 = frmElpPrt.prtHeightDelta(8, 6)

blnFiligrane = False
 prtOrientation = vbPRORLandscape '
prtPgmName = "prtYEICGCC0"
prtTitleUsr = usrName

prtLineNb = 1
prtlineHeight = 250
prtHeaderHeight = 300


prtFormType = ""
frmElpPrt.prtStdInit

prtFontName = prtFontName_Arial
Select Case mForm
    Case "YEICGCC0": prtYEICGCC0_Form
    Case "YEICGCCLOG": prtYEICGCCLOG_Form
    Case "Echéancier": prtYEICGCC0_Form_Echéancier
    Case "Statistiques": prtYEICGCC0_Form_Statistiques
End Select
XPrt.CurrentY = XPrt.CurrentY - prtlineHeight
Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub
Public Sub prtYEICGCC0_Init(lForm As String, lText As String)

mForm = lForm
prtTitleText = "Gestion des chèques circulants : " & lText
mAMJ_7Past = dateElp("Jour", 7, YBIATAB0_DATE_CPT_JS1)
mAMJ_7Ante = dateElp("Jour", -7, YBIATAB0_DATE_CPT_JS1)

End Sub

Public Sub prtYEICGCC0_Close(blnEnd As Boolean)
On Error GoTo prtError

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight

    prtMaxY_YEICGCC0 = XPrt.CurrentY
    Select Case mForm
        Case "YEICGCC0"
                prtYEICGCC0_Col
        Case "YEICGCCLOG"
                prtYEICGCCLOG_Col
                
        Case "Echéancier"
                prtYEICGCC0_Col_echéancier
        Case "Statistiques"
                prtYEICGCC0_Col_Statistiques
    End Select
    XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY), prtLineColor
If blnEnd Then
    Call frmElpPrt.prtEndDoc(1000)
    frmElpPrt.Hide
End If
Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub


Public Sub prtYEICGCC0_NewLine()
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
If XPrt.CurrentY + 300 > prtMaxY Then
    prtMaxY_YEICGCC0 = XPrt.CurrentY
    Select Case mForm
        Case "YEICGCC0":
                prtYEICGCC0_Col
                frmElpPrt.prtNewPage
                prtYEICGCC0_Form
        Case "YEICGCCLOG":
                prtYEICGCCLOG_Col
                frmElpPrt.prtNewPage
                prtYEICGCCLOG_Form
        Case "Echéancier":
                prtYEICGCC0_Col_echéancier
                frmElpPrt.prtNewPage
                prtYEICGCC0_Form_Echéancier
End Select
End If

End Sub




'---------------------------------------------------------
Public Sub prtYEICGCC0_Col_echéancier()
'---------------------------------------------------------

XPrt.DrawWidth = 1
prtLineColor = vbCyan
XPrt.Line (prtMinX + 1950, prtMinY_YEICGCC0)-(prtMinX + 1950, prtMaxY_YEICGCC0), prtLineColor
XPrt.Line (prtMinX + 7550, prtMinY_YEICGCC0)-(prtMinX + 7550, prtMaxY_YEICGCC0), prtLineColor
End Sub

'---------------------------------------------------------
Public Sub prtYEICGCC0_Col_Statistiques()
'---------------------------------------------------------

XPrt.DrawWidth = 1
prtLineColor = vbCyan
XPrt.Line (prtMinX + 3500 + 100, prtMinY_YEICGCC0)-(prtMinX + 3500 + 100, prtMaxY_YEICGCC0), prtLineColor
XPrt.Line (prtMinX + 5000 + 100, prtMinY_YEICGCC0)-(prtMinX + 5000 + 100, prtMaxY_YEICGCC0), prtLineColor
XPrt.Line (prtMinX + 6500 + 100, prtMinY_YEICGCC0)-(prtMinX + 6500 + 100, prtMaxY_YEICGCC0), prtLineColor
XPrt.Line (prtMinX + 8000 + 100, prtMinY_YEICGCC0)-(prtMinX + 8000 + 100, prtMaxY_YEICGCC0), prtLineColor
XPrt.Line (prtMinX + 9500 + 100, prtMinY_YEICGCC0)-(prtMinX + 9500 + 100, prtMaxY_YEICGCC0), prtLineColor
XPrt.Line (prtMinX + 11000 + 100, prtMinY_YEICGCC0)-(prtMinX + 11000 + 100, prtMaxY_YEICGCC0), prtLineColor
XPrt.Line (prtMinX + 12500 + 100, prtMinY_YEICGCC0)-(prtMinX + 12500 + 100, prtMaxY_YEICGCC0), prtLineColor
XPrt.Line (prtMinX + 14000 + 100, prtMinY_YEICGCC0)-(prtMinX + 14000 + 100, prtMaxY_YEICGCC0), prtLineColor

End Sub

'---------------------------------------------------------
Public Sub prtYEICGCC0_Col()
'---------------------------------------------------------

XPrt.DrawWidth = 1
prtLineColor = vbCyan
XPrt.Line (prtMinX + 1950, prtMinY_YEICGCC0)-(prtMinX + 1950, prtMaxY_YEICGCC0), prtLineColor
XPrt.Line (prtMinX + 7550, prtMinY_YEICGCC0)-(prtMinX + 7550, prtMaxY_YEICGCC0), prtLineColor
XPrt.Line (prtMinX + 13450, prtMinY_YEICGCC0)-(prtMinX + 13450, prtMaxY_YEICGCC0), prtLineColor
End Sub


'---------------------------------------------------------
Public Sub prtYEICGCCLOG_Col()
'---------------------------------------------------------

XPrt.DrawWidth = 1
prtLineColor = vbCyan
XPrt.Line (prtMinX + 4400, prtMinY_YEICGCC0)-(prtMinX + 4400, prtMaxY_YEICGCC0), prtLineColor
XPrt.Line (prtMinX + 7800, prtMinY_YEICGCC0)-(prtMinX + 7800, prtMaxY_YEICGCC0), prtLineColor
End Sub


Public Sub prtYEICGCC0_Form()
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
prtMinY_YEICGCC0 = prtMinY + prtHeaderHeight '* 2


End Sub
Public Sub prtYEICGCC0_Form_Echéancier()
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
XPrt.CurrentX = prtMinX + 8700: XPrt.Print "Action";
XPrt.CurrentX = prtMinX + 10000: XPrt.Print "Commentaire";
XPrt.CurrentX = prtMinX + 15600: XPrt.Print "Id";

XPrt.FontSize = 8
XPrt.FontBold = False
XPrt.ForeColor = vbBlack
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentY = XPrt.CurrentY + 50
prtMinY_YEICGCC0 = prtMinY + prtHeaderHeight


End Sub

Public Sub prtYEICGCC0_Form_Statistiques()
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
X = "Nature": XPrt.CurrentX = prtMinX + 100: XPrt.Print X;
X = "Total": XPrt.CurrentX = prtMinX + 5000 - XPrt.TextWidth(X): XPrt.Print X;
X = "En cours": XPrt.CurrentX = prtMinX + 6500 - XPrt.TextWidth(X): XPrt.Print X;
X = "Vérifiés": XPrt.CurrentX = prtMinX + 8000 - XPrt.TextWidth(X): XPrt.Print X;
X = "Annulés": XPrt.CurrentX = prtMinX + 9500 - XPrt.TextWidth(X): XPrt.Print X;
X = "à ignorer": XPrt.CurrentX = prtMinX + 11000 - XPrt.TextWidth(X): XPrt.Print X;
X = "non circulants": XPrt.CurrentX = prtMinX + 12500 - XPrt.TextWidth(X): XPrt.Print X;
X = "Rejetés": XPrt.CurrentX = prtMinX + 14000 - XPrt.TextWidth(X): XPrt.Print X;

XPrt.FontSize = 8
XPrt.FontBold = False
XPrt.ForeColor = vbBlack
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentY = XPrt.CurrentY + 50
prtMinY_YEICGCC0 = prtMinY + prtHeaderHeight


End Sub

Public Sub prtYEICGCCLOG_Form()
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

XPrt.CurrentX = prtMinX + 200: XPrt.Print "Date";
XPrt.CurrentX = prtMinX + 1300: XPrt.Print "Heure";
XPrt.CurrentX = prtMinX + 2800: XPrt.Print "Utilisateur";
XPrt.CurrentX = prtMinX + 5000: XPrt.Print "dossier";
XPrt.CurrentX = prtMinX + 6000: XPrt.Print "échéance";
XPrt.CurrentX = prtMinX + 7900: XPrt.Print "Action";
XPrt.CurrentX = prtMinX + 9400: XPrt.Print "Commentaire";


XPrt.FontSize = 8
XPrt.FontBold = False
XPrt.ForeColor = vbBlack
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentY = XPrt.CurrentY + 50
prtMinY_YEICGCC0 = prtMinY + prtHeaderHeight


End Sub

Public Sub prtYEICGCC0_Line(lYEICGCC0 As typeYEICGCC0)
Dim X As String, xSQL As String
Dim wCLIENARA1 As String
Dim wColor As Long
prtYEICGCC0_NewLine


XPrt.ForeColor = vbBlack 'prtForeColor
Select Case lYEICGCC0.EICGCCSTA
    Case "I", "A", "R": XPrt.ForeColor = RGB(128, 128, 128)
    Case "V", "@": XPrt.ForeColor = RGB(64, 128, 64)
    Case Else
        If lYEICGCC0.EICGCCDOS = 0 Then
            XPrt.ForeColor = vbRed
        Else
            If lYEICGCC0.EICGCCVJPG = 0 Then
                If lYEICGCC0.EICGCCAMJ < mAMJ_7Ante Then
                    XPrt.ForeColor = vbRed
                Else
                    XPrt.ForeColor = vbMagenta 'RGB(255, 96, 32)
                End If
            Else
                Select Case lYEICGCC0.EICGCCSTAK
                    Case "X": XPrt.ForeColor = vbMagenta
                    Case Else: XPrt.ForeColor = vbBlue
                End Select
            End If
        End If
        
End Select
wColor = XPrt.ForeColor
XPrt.FontSize = 8

XPrt.CurrentX = prtMinX
XPrt.Print dateImp10(lYEICGCC0.EICGCCAMJ);
XPrt.CurrentX = prtMinX + 1000: XPrt.Print lYEICGCC0.EICGCCOPE & "  " & lYEICGCC0.EICGCCDOS;

XPrt.ForeColor = RGB(0, 0, 128)

XPrt.CurrentX = prtMinX + 2000
If IsNumeric(lYEICGCC0.EICGCCECPT) Then
    XPrt.Print Format(lYEICGCC0.EICGCCECPT, "@@@@@ @@@ @@@ @@@@@@@");
Else
    XPrt.Print lYEICGCC0.EICGCCECPT;
End If

xSQL = "select CLIENARA1, COMPTEINT from " & paramIBM_Library_SABSPE & ".YBIACPT0 " _
     & " Where COMPTECOM = '" & lYEICGCC0.EICGCCECPT & "'"
Set rsSab = cnsab.Execute(xSQL)
If Not rsSab.EOF Then
    wCLIENARA1 = rsSab("CLIENARA1")
    XPrt.CurrentX = prtMinX + 3500: XPrt.Print rsSab("COMPTEINT");
Else
    wCLIENARA1 = "?"
End If
'______________________________________________________________________________________
XPrt.CurrentY = XPrt.CurrentY + Height8_6
XPrt.FontSize = 6
XPrt.CurrentX = prtMinX + 8250: If lYEICGCC0.EICGCCEIND <> " " Then XPrt.Print "/" & lYEICGCC0.EICGCCEIND;

XPrt.CurrentX = prtMinX + 8600: XPrt.Print lYEICGCC0.EICGCCXNOM;
X = Format$(lYEICGCC0.EICGCCID, "### ### ### ##0")
XPrt.CurrentX = prtMaxX - XPrt.TextWidth(X): XPrt.Print X;

XPrt.FontSize = 8
XPrt.CurrentY = XPrt.CurrentY - Height8_6
'___________________________________________________________________________________

X = Format$(lYEICGCC0.EICGCCEMT, "### ### ### ##0.00")

XPrt.CurrentX = prtMinX + 7500 - XPrt.TextWidth(X): XPrt.Print X;
XPrt.CurrentX = prtMinX + 7600: XPrt.Print lYEICGCC0.EICGCCECHQ;
'If lYEICGCC0.EICGCCXID > 0 Then XPrt.Print "-" & lYEICGCC0.EICGCCXID;
XPrt.CurrentX = prtMinX + 11500: XPrt.Print lYEICGCC0.EICGCCVINT;
XPrt.CurrentX = prtMinX + 10800

If lYEICGCC0.EICGCCVAMJ > 0 Then
    XPrt.CurrentX = prtMinX + 13500: XPrt.Print dateImp10(lYEICGCC0.EICGCCVAMJ);
    XPrt.CurrentX = prtMinX + 14400: XPrt.Print lYEICGCC0.EICGCCVJPG;
End If


Exit Sub


prtYEICGCC0_NewLine
XPrt.ForeColor = wColor

XPrt.CurrentX = prtMinX + 1000: XPrt.Print lYEICGCC0.EICGCCSER & "  " & lYEICGCC0.EICGCCSSE;
XPrt.CurrentX = prtMinX: XPrt.Print lYEICGCC0.EICGCCSTA;

XPrt.ForeColor = RGB(0, 0, 128)

XPrt.CurrentX = prtMinX + 2000: XPrt.Print lYEICGCC0.EICGCCECLI;
'XPrt.CurrentX = prtMinX + 3500: XPrt.Print wCLIENARA1;
If lYEICGCC0.EICGCCEAMJ > 0 Then XPrt.CurrentX = prtMinX + 7600: XPrt.Print dateImp10(lYEICGCC0.EICGCCEAMJ);
XPrt.CurrentX = prtMinX + 3500: XPrt.Print lYEICGCC0.EICGCCXECO;

XPrt.CurrentX = prtMinX + 8600
If lYEICGCC0.EICGCCXBQ <> strSocBdfE Then
    XPrt.Print lYEICGCC0.EICGCCXBQ;
Else
    If IsNumeric(lYEICGCC0.EICGCCXCPT) Then
        XPrt.Print Format(lYEICGCC0.EICGCCXCPT, "@@@@@ @@@ @@@ @@@@@@@");
    Else
        XPrt.Print lYEICGCC0.EICGCCXCPT;
    End If
End If

XPrt.CurrentX = prtMinX + 11500: XPrt.Print lYEICGCC0.EICGCCVEXT;
If lYEICGCC0.EICGCCVREM > 0 Then XPrt.CurrentX = prtMinX + 13500: XPrt.Print lYEICGCC0.EICGCCVREM;

'X = "----"
'If lYEICGCC0.EICGCCKMT <> " " Then Mid$(X, 1, 1) = lYEICGCC0.EICGCCKMT
'If lYEICGCC0.EICGCCKSIG <> " " Then Mid$(X, 2, 1) = lYEICGCC0.EICGCCKSIG
'If lYEICGCC0.EICGCCKEND <> " " Then Mid$(X, 3, 1) = lYEICGCC0.EICGCCKEND
'If lYEICGCC0.EICGCCKLAB <> " " Then Mid$(X, 4, 1) = lYEICGCC0.EICGCCKLAB

'XPrt.CurrentX = prtMinX + 15200: XPrt.Print X;

'XPrt.FontItalic = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight / 2

End Sub

Public Sub prtYEICGCC0_Line_Echéancier(lYEICGCC0 As typeYEICGCC0, lYEICGCCLOG As typeYEICGCCLOG)
Dim X As String, xSQL As String
Dim wCLIENARA1 As String
Dim wColor As Long
prtYEICGCC0_NewLine


XPrt.ForeColor = vbMagenta
XPrt.FontSize = 8

XPrt.CurrentX = prtMinX
XPrt.Print dateImp10(lYEICGCC0.EICGCCAMJ);
XPrt.CurrentX = prtMinX + 1000: XPrt.Print lYEICGCC0.EICGCCOPE & "  " & lYEICGCC0.EICGCCDOS;

XPrt.ForeColor = RGB(0, 0, 128)

XPrt.CurrentX = prtMinX + 2000
If IsNumeric(lYEICGCC0.EICGCCECPT) Then
    XPrt.Print Format(lYEICGCC0.EICGCCECPT, "@@@@@ @@@ @@@ @@@@@@@");
Else
    XPrt.Print lYEICGCC0.EICGCCECPT;
End If

xSQL = "select CLIENARA1, COMPTEINT from " & paramIBM_Library_SABSPE & ".YBIACPT0 " _
     & " Where COMPTECOM = '" & lYEICGCC0.EICGCCECPT & "'"
Set rsSab = cnsab.Execute(xSQL)
If Not rsSab.EOF Then
    wCLIENARA1 = rsSab("CLIENARA1")
    XPrt.CurrentX = prtMinX + 3500: XPrt.Print rsSab("COMPTEINT");
Else
    wCLIENARA1 = "?"
End If
'______________________________________________________________________________________
XPrt.CurrentY = XPrt.CurrentY + Height8_6
XPrt.FontSize = 6
XPrt.CurrentX = prtMinX + 8250: If lYEICGCC0.EICGCCEIND <> " " Then XPrt.Print "/" & lYEICGCC0.EICGCCEIND;

'   XPrt.CurrentX = prtMinX + 8600: XPrt.Print lYEICGCC0.EICGCCXNOM;
X = Format$(lYEICGCC0.EICGCCID, "### ### ### ##0")
XPrt.CurrentX = prtMaxX - XPrt.TextWidth(X): XPrt.Print X;

XPrt.FontSize = 8
XPrt.CurrentY = XPrt.CurrentY - Height8_6
'___________________________________________________________________________________

X = Format$(lYEICGCC0.EICGCCEMT, "### ### ### ##0.00")

XPrt.CurrentX = prtMinX + 7500 - XPrt.TextWidth(X): XPrt.Print X;
XPrt.CurrentX = prtMinX + 7600: XPrt.Print lYEICGCC0.EICGCCECHQ;
XPrt.CurrentX = prtMinX + 8700: XPrt.Print lYEICGCCLOG.EICGCCLOGK;
XPrt.CurrentX = prtMinX + 10000: XPrt.Print lYEICGCCLOG.EICGCCLOGX;


XPrt.CurrentY = XPrt.CurrentY + prtlineHeight / 2

End Sub

Public Sub prtYEICGCC0_Line_Statistiques(fgX As MSFlexGrid)
Dim X As String, xSQL As String
Dim wCLIENARA1 As String
Dim K As Long, Nb As Long

fgX.Row = 0: fgX.Col = 0


XPrt.FontSize = 8: XPrt.FontBold = False
For K = 1 To fgX.Rows - 1
    fgX.Row = K

    fgX.Col = 0
    If fgX.CellBackColor <> 0 Then 'InStr(fgX.Text, "Ventilation") > 0 Then
        prtFillColor = fgX.CellBackColor
        'XPrt.ForeColor = vbWhite
        Call frmElpPrt.prtTrame_Color(prtMinX, XPrt.CurrentY, prtMaxX, XPrt.CurrentY + prtlineHeight, " ")
        '---------------------------------------------------------
        prtFillColor = RGB(200, 255, 255)
        XPrt.CurrentX = prtMinX + 100: XPrt.Print fgX.Text;
        fgX.Col = 1: X = Trim(fgX.Text)
        XPrt.CurrentX = prtMinX + 5000 - XPrt.TextWidth(X): XPrt.Print X;
        fgX.Col = 2: X = Trim(fgX.Text)
        XPrt.CurrentX = prtMinX + 6500 - XPrt.TextWidth(X): XPrt.Print X;
        fgX.Col = 3: X = Trim(fgX.Text)
        XPrt.CurrentX = prtMinX + 8000 - XPrt.TextWidth(X): XPrt.Print X;
        fgX.Col = 4: X = Trim(fgX.Text)
        XPrt.CurrentX = prtMinX + 9500 - XPrt.TextWidth(X): XPrt.Print X;
        fgX.Col = 5: X = Trim(fgX.Text)
        XPrt.CurrentX = prtMinX + 11000 - XPrt.TextWidth(X): XPrt.Print X;
        fgX.Col = 6: X = Trim(fgX.Text)
        XPrt.CurrentX = prtMinX + 12500 - XPrt.TextWidth(X): XPrt.Print X;
        fgX.Col = 7: X = Trim(fgX.Text)
        XPrt.CurrentX = prtMinX + 14000 - XPrt.TextWidth(X): XPrt.Print X;
    Else
        XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    
        XPrt.CurrentX = prtMinX + 100: XPrt.Print fgX.Text;
        fgX.Col = 1: X = Format(Val(fgX.Text), "### ### ###")
        XPrt.CurrentX = prtMinX + 5000 - XPrt.TextWidth(X): XPrt.Print X;
        fgX.Col = 2: X = Format(Val(fgX.Text), "### ### ###")
        XPrt.CurrentX = prtMinX + 6500 - XPrt.TextWidth(X): XPrt.Print X;
        fgX.Col = 3: X = Format(Val(fgX.Text), "### ### ###")
        XPrt.CurrentX = prtMinX + 8000 - XPrt.TextWidth(X): XPrt.Print X;
        fgX.Col = 4: X = Format(Val(fgX.Text), "### ### ###")
        XPrt.CurrentX = prtMinX + 9500 - XPrt.TextWidth(X): XPrt.Print X;
        fgX.Col = 5: X = Format(Val(fgX.Text), "### ### ###")
        XPrt.CurrentX = prtMinX + 11000 - XPrt.TextWidth(X): XPrt.Print X;
        fgX.Col = 6: X = Format(Val(fgX.Text), "### ### ###")
        XPrt.CurrentX = prtMinX + 12500 - XPrt.TextWidth(X): XPrt.Print X;
        fgX.Col = 7: X = Format(Val(fgX.Text), "### ### ###")
        XPrt.CurrentX = prtMinX + 14000 - XPrt.TextWidth(X): XPrt.Print X;
    End If
    
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY), prtLineColor
    XPrt.CurrentY = XPrt.CurrentY - prtlineHeight
Next K


End Sub


Public Sub prtYEICGCCLOG_Line(lYEICGCCLOG As typeYEICGCCLOG)
Dim X As String, xSQL As String
Dim wCLIENARA1 As String
Dim wColor As Long
prtYEICGCC0_NewLine


XPrt.ForeColor = vbBlack 'prtForeColor
Select Case lYEICGCCLOG.EICGCCLOGA
    Case "A", "I": XPrt.ForeColor = RGB(128, 128, 128)
    Case "V": XPrt.ForeColor = RGB(32, 96, 32)
    Case Else
        If lYEICGCCLOG.EICGCCLOGE > 0 Then
            XPrt.ForeColor = vbMagenta
        Else
            XPrt.ForeColor = vbBlue
        End If
        
End Select
If lYEICGCCLOG.EICGCCLOGI > 0 Then
    X = Format$(lYEICGCCLOG.EICGCCLOGI, "### ### ### ##0")
    XPrt.CurrentX = prtMinX + 5400 - XPrt.TextWidth(X): XPrt.Print X;
End If

If lYEICGCCLOG.EICGCCLOGE > 0 And lYEICGCCLOG.EICGCCLOGA = " " Then
    XPrt.CurrentX = prtMinX + 6000: XPrt.Print dateImp10(lYEICGCCLOG.EICGCCLOGE);
End If
XPrt.CurrentX = prtMinX + 7500: XPrt.Print lYEICGCCLOG.EICGCCLOGA;

'wColor = XPrt.ForeColor
XPrt.ForeColor = vbBlack
XPrt.FontSize = 8

XPrt.CurrentX = prtMinX + 100
XPrt.Print dateImp10(lYEICGCCLOG.EICGCCLOGD);
XPrt.CurrentX = prtMinX + 1300: XPrt.Print timeImp8(lYEICGCCLOG.EICGCCLOGH) & " - " & lYEICGCCLOG.EICGCCLOGS;

XPrt.CurrentX = prtMinX + 2800: XPrt.Print lYEICGCCLOG.EICGCCLOGU;
XPrt.CurrentX = prtMinX + 8000: XPrt.Print lYEICGCCLOG.EICGCCLOGK;
XPrt.CurrentX = prtMinX + 9400: XPrt.Print lYEICGCCLOG.EICGCCLOGX;

End Sub

Public Sub prtYEICGCCLOG_Line_xlsManual(lYEICGCCLOG As typeYEICGCCLOG, ByRef currentRow As Long, wsExcel As Excel.Worksheet, ByRef comptageRows As Long, maxRows As Long, maxRowsPlus As Long)
Dim X As String, xSQL As String
Dim wCLIENARA1 As String
Dim wColor As Long

If currentRow >= maxRows + maxRowsPlus Then
    If comptageRows >= maxRows Then
        Call insere_entete_page(wsExcel, "1:3", 3, currentRow)
        comptageRows = 3
        currentRow = currentRow + 3
    End If
End If
comptageRows = comptageRows + 1
currentRow = currentRow + 1
Range("4:4").Select
Selection.Copy
Range("A" & CStr(currentRow)).Select
ActiveSheet.Paste
wColor = vbBlack
Select Case lYEICGCCLOG.EICGCCLOGA
    Case "A", "I": wColor = RGB(128, 128, 128)
    Case "V": wColor = RGB(32, 96, 32)
    Case Else
        If lYEICGCCLOG.EICGCCLOGE > 0 Then
            wColor = vbMagenta
        Else
            wColor = vbBlue
        End If
End Select
If lYEICGCCLOG.EICGCCLOGI > 0 Then
    X = Format$(lYEICGCCLOG.EICGCCLOGI, "### ### ### ##0")
    wsExcel.Cells(currentRow, 5) = X
End If

If lYEICGCCLOG.EICGCCLOGE > 0 And lYEICGCCLOG.EICGCCLOGA = " " Then
    wsExcel.Cells(currentRow, 4) = "'" & dateImp10(lYEICGCCLOG.EICGCCLOGE)
End If
wsExcel.Cells(currentRow, 5) = lYEICGCCLOG.EICGCCLOGA

wColor = vbBlack
wsExcel.Cells(currentRow, 1) = "'" & dateImp10(lYEICGCCLOG.EICGCCLOGD)
wsExcel.Cells(currentRow, 1).Font.Color = wColor
wsExcel.Cells(currentRow, 2) = timeImp8(lYEICGCCLOG.EICGCCLOGH) & " - " & lYEICGCCLOG.EICGCCLOGS
wsExcel.Cells(currentRow, 2).Font.Color = wColor
wsExcel.Cells(currentRow, 3) = lYEICGCCLOG.EICGCCLOGU
wsExcel.Cells(currentRow, 3).Font.Color = wColor
wsExcel.Cells(currentRow, 6) = lYEICGCCLOG.EICGCCLOGK
wsExcel.Cells(currentRow, 6).Font.Color = wColor
wsExcel.Cells(currentRow, 7) = lYEICGCCLOG.EICGCCLOGX
wsExcel.Cells(currentRow, 7).Font.Color = wColor
End Sub


