Attribute VB_Name = "prtYICCCPT0"
Option Explicit
Dim mForm As String
Dim X As String, I As Integer, Height8_6 As Integer
Dim curX As Currency, curX1 As Currency, curX2 As Currency
Dim prtMaxY_YICCCPT0 As Long, prtMinY_YICCCPT0 As Long

Dim blnLine_Trame As Boolean


Public Sub prtYICCCPT0_Close_xlsManual(blnEnd As Boolean, ByRef currentRow As Long, wsExcel As Excel.Worksheet, ByRef comptageRows As Long, maxRows As Long, maxRowsPlus As Long)

If currentRow >= maxRows + maxRowsPlus Then
    If comptageRows >= maxRows Then
        Call insere_entete_page(wsExcel, "1:2", 2, currentRow)
        wsExcel.Cells(1, 4) = prtTitleText
        comptageRows = 2
        currentRow = currentRow + 2
    End If
End If
comptageRows = comptageRows + 1
currentRow = currentRow + 1
wsExcel.Activate
Range("6:6").Select
Selection.Copy
Range("A" & CStr(currentRow)).Select
ActiveSheet.Paste

If blnEnd Then
    'ne rien faire
    'wsExcel.Cells(1, 4) = prtTitleText
Else
    wsExcel.Cells(1, 4) = prtTitleText 'substitution du premier titre en haut de la feuille
    Call insere_entete_page(wsExcel, "1:2", 2, currentRow)
    comptageRows = 2
    currentRow = currentRow + 2
End If

End Sub

Public Sub prtYICCCPT0_Line_Detail(lYICCMVT0 As typeYICCMVT0)
Dim X As String
prtYICCCPT0_NewLine

XPrt.ForeColor = vbBlue
XPrt.FontSize = 8

XPrt.CurrentX = prtMinX
XPrt.CurrentX = prtMinX + 2500: XPrt.Print lYICCMVT0.ICCMVTSER & " " & lYICCMVT0.ICCMVTSSE & " " _
              & lYICCMVT0.ICCMVTOPE & " " & lYICCMVT0.ICCMVTNAT & " " & lYICCMVT0.ICCMVTDOS & " " & lYICCMVT0.ICCMVTEVE;

XPrt.CurrentX = prtMinX + 5700: XPrt.Print dateImp10(lYICCMVT0.ICCMVTAMJ);

'______________________________________________________
Select Case lYICCMVT0.ICCMVTOPE

    Case "EMP", "EM1":
        If lYICCMVT0.ICCMVTRBT <> lYICCMVT0.ICCMVTTDB Then
            prtFillColor = RGB(255, 160, 255)
            Call frmElpPrt.prtTrame_Color(prtMinX + 6700, XPrt.CurrentY - 50, prtMinX + 8150, XPrt.CurrentY + prtlineHeight - 50, " ")
        End If
        If lYICCMVT0.ICCMVTPRO <> lYICCMVT0.ICCMVTTCR Then
            prtFillColor = RGB(255, 160, 255)
            Call frmElpPrt.prtTrame_Color(prtMinX + 14200, XPrt.CurrentY - 50, prtMinX + 15650, XPrt.CurrentY + prtlineHeight - 50, " ")
        End If
    Case Else:
        If Mid$(lYICCMVT0.ICCMVTOPE, 1, 1) <> "*" And Mid$(lYICCMVT0.ICCMVTOPE, 1, 1) <> " " Then
            If lYICCMVT0.ICCMVTRBT <> lYICCMVT0.ICCMVTTCR Then
                prtFillColor = RGB(255, 160, 255)
                Call frmElpPrt.prtTrame_Color(prtMinX + 6700, XPrt.CurrentY - 50, prtMinX + 8150, XPrt.CurrentY + prtlineHeight - 50, " ")
            End If
            If lYICCMVT0.ICCMVTPRO <> lYICCMVT0.ICCMVTTDB Then
                prtFillColor = RGB(255, 160, 255)
                Call frmElpPrt.prtTrame_Color(prtMinX + 14200, XPrt.CurrentY - 50, prtMinX + 15650, XPrt.CurrentY + prtlineHeight - 50, " ")
            End If
        End If
End Select

'___________________________________________________________________________________

If lYICCMVT0.ICCMVTRBT <> 0 Then
    XPrt.ForeColor = vbMagenta 'IIf(lYICCMVT0.ICCMVTRBT < 0, vbMagenta, RGB(32, 96, 32))
    X = Format$(lYICCMVT0.ICCMVTRBT, "### ### ### ##0.00")
    XPrt.CurrentX = prtMinX + 8100 - XPrt.TextWidth(X): XPrt.Print X;
End If
If lYICCMVT0.ICCMVTTDB <> 0 Then
    XPrt.ForeColor = IIf(lYICCMVT0.ICCMVTTDB < 0, vbRed, vbBlue)
    X = Format$(lYICCMVT0.ICCMVTTDB, "### ### ### ##0.00")
    XPrt.CurrentX = prtMinX + 11100 - XPrt.TextWidth(X): XPrt.Print X;
End If
If lYICCMVT0.ICCMVTTCR <> 0 Then
    XPrt.ForeColor = IIf(lYICCMVT0.ICCMVTTCR < 0, vbRed, vbBlue)
    X = Format$(lYICCMVT0.ICCMVTTCR, "### ### ### ##0.00")
    XPrt.CurrentX = prtMinX + 12600 - XPrt.TextWidth(X): XPrt.Print X;
End If

If lYICCMVT0.ICCMVTPRO <> 0 Then
    XPrt.ForeColor = RGB(32, 96, 32) 'IIf(lYICCMVT0.ICCMVTPRO < 0, vbMagenta, RGB(32, 96, 32))
    X = Format$(lYICCMVT0.ICCMVTPRO, "### ### ### ##0.00")
    XPrt.CurrentX = prtMinX + 15600 - XPrt.TextWidth(X): XPrt.Print X;
End If

End Sub

Public Sub prtYICCCPT0_Line_xlsManual(lYICCCPT0 As typeYICCCPT0, Total As typeYICCMVT0, lsoldeD As typeYICCMVT0, lsoldeF As typeYICCMVT0, blnDetail As Boolean, lYICCMVT0_Nb As Long, blnAvance As Boolean, ByRef currentRow As Long, wsExcel As Excel.Worksheet, ByRef comptageRows As Long, maxRows As Long, maxRowsPlus As Long)
Dim X As String, xSQL As String, wSoldeD As Currency, wSoldeF As Currency
Dim wColor_Cr As Long
Dim wColor As Long

Call prtYICCCPT0_NewLine_xlsManual(currentRow, wsExcel, comptageRows, maxRows, maxRowsPlus)
If blnDetail Then
    wColor_Cr = vbBlack
Else
    wColor_Cr = vbBlue
End If
If lsoldeD.ICCMVTRBT <> 0 Or lsoldeD.ICCMVTPRO <> 0 Then
    prtFillColor = RGB(255, 230, 240): blnLine_Trame = True
Else
    If lYICCMVT0_Nb = 0 Then
        prtFillColor = RGB(250, 250, 250)
    Else
        prtFillColor = RGB(245, 255, 245): blnLine_Trame = True
    End If
    If blnDetail Then blnLine_Trame = True
End If
If blnLine_Trame Then
    blnLine_Trame = False
    wsExcel.Rows(currentRow).Interior.Color = prtFillColor
Else
    blnLine_Trame = True
End If
wColor = vbBlack
wsExcel.Cells(currentRow, 1) = lYICCCPT0.ICCCPTDEV
wsExcel.Cells(currentRow, 2) = lYICCCPT0.ICCCPTCOM
xSQL = "select COMPTEINT from " & paramIBM_Library_SABSPE & ".YBIACPT0 " _
     & " Where COMPTECOM = '" & lYICCCPT0.ICCCPTCOM & "'"
Set rsSab = cnsab.Execute(xSQL)
If Not rsSab.EOF Then
    wsExcel.Cells(currentRow, 3) = rsSab("COMPTEINT")
End If
'___________________________________________________________________________________
wSoldeD = lsoldeD.ICCMVTTDB + lsoldeD.ICCMVTTCR
If Not blnAvance Then
    If Total.ICCMVTRBT + lsoldeD.ICCMVTTDB + lsoldeD.ICCMVTTCR <> 0 Then
        prtFillColor = RGB(255, 226, 200)
        wsExcel.Cells(currentRow, 6).Interior.Color = prtFillColor
        prtFillColor = RGB(255, 180, 255)
        wsExcel.Cells(currentRow, 5).Interior.Color = prtFillColor
    End If
End If
If Total.ICCMVTRBT <> 0 Then
    wColor = vbMagenta
    X = Format$(Total.ICCMVTRBT, "### ### ### ##0.00")
    wsExcel.Cells(currentRow, 5) = X
    wsExcel.Cells(currentRow, 5).Font.Color = wColor
End If
If wSoldeD <> 0 Then
    wColor = IIf(wSoldeD < 0, vbRed, wColor_Cr)
    X = Format$(wSoldeD, "### ### ### ##0.00")
    wsExcel.Cells(currentRow, 6) = X
    wsExcel.Cells(currentRow, 6).Font.Color = wColor
End If
If Total.ICCMVTTDB <> 0 Then
    wColor = IIf(Total.ICCMVTTDB < 0, vbRed, wColor_Cr)
    X = Format$(Total.ICCMVTTDB, "### ### ### ##0.00")
    wsExcel.Cells(currentRow, 7) = X
    wsExcel.Cells(currentRow, 7).Font.Color = wColor
End If
If Total.ICCMVTTCR <> 0 Then
    wColor = IIf(Total.ICCMVTTCR < 0, vbRed, wColor_Cr)
    X = Format$(Total.ICCMVTTCR, "### ### ### ##0.00")
    wsExcel.Cells(currentRow, 8) = X
    wsExcel.Cells(currentRow, 8).Font.Color = wColor
End If
If Not blnAvance Then
    If Total.ICCMVTPRO - lsoldeF.ICCMVTTDB - lsoldeF.ICCMVTTCR <> 0 Then
        prtFillColor = RGB(255, 226, 200)
        wsExcel.Cells(currentRow, 9).Interior.Color = prtFillColor
        prtFillColor = RGB(255, 180, 255)
        wsExcel.Cells(currentRow, 10).Interior.Color = prtFillColor
    End If
End If
If Total.ICCMVTPRO <> 0 Then
    wColor = RGB(32, 96, 32)
    X = Format$(Total.ICCMVTPRO, "### ### ### ##0.00")
    wsExcel.Cells(currentRow, 10) = X
    wsExcel.Cells(currentRow, 10).Font.Color = wColor
End If

wSoldeF = lsoldeF.ICCMVTTDB + lsoldeF.ICCMVTTCR
If wSoldeF <> wSoldeD + Total.ICCMVTTDB + Total.ICCMVTTCR Then
    prtFillColor = RGB(255, 160, 255)
    wsExcel.Cells(currentRow, 9).Interior.Color = prtFillColor
End If
If wSoldeF <> 0 Then
    wColor = IIf(wSoldeF < 0, vbRed, wColor_Cr)
    X = Format$(wSoldeF, "### ### ### ##0.00")
    wsExcel.Cells(currentRow, 9) = X
    wsExcel.Cells(currentRow, 9).Font.Color = wColor
End If
End Sub

Public Sub prtYICCCPT0_NewLine_xlsManual(ByRef currentRow As Long, wsExcel As Excel.Worksheet, ByRef comptageRows As Long, maxRows As Long, maxRowsPlus As Long)

    If currentRow >= maxRows + maxRowsPlus Then
        If comptageRows >= maxRows Then
            Call insere_entete_page(wsExcel, "1:2", 2, currentRow)
            wsExcel.Cells(1, 4) = prtTitleText
            comptageRows = 2
            currentRow = currentRow + 2
        End If
    End If
    comptageRows = comptageRows + 1
    currentRow = currentRow + 1
    wsExcel.Activate
    Range("4:4").Select
    Selection.Copy
    Range("A" & CStr(currentRow)).Select
    ActiveSheet.Paste

End Sub

Public Sub prtYICCCPT0_Open()
On Error GoTo prtError

Set XPrt = Printer
frmElpPrt.Show vbModeless

Height8_6 = frmElpPrt.prtHeightDelta(8, 6)

blnFiligrane = False
 prtOrientation = vbPRORLandscape '
prtPgmName = "prtYICCCPT0"
prtTitleUsr = usrName

prtLineNb = 1
prtlineHeight = 300
prtHeaderHeight = 300


prtFormType = ""
frmElpPrt.prtStdInit

prtFontName = prtFontName_Arial
prtYICCCPT0_Form
XPrt.CurrentY = XPrt.CurrentY - prtlineHeight
Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub

Public Sub prtYICCCPT0_Close(blnEnd As Boolean)
On Error GoTo prtError

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight

    prtMaxY_YICCCPT0 = XPrt.CurrentY
                prtYICCCPT0_Col
    XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY), prtLineColor
If blnEnd Then
    Call frmElpPrt.prtEndDoc(1000)
    frmElpPrt.Hide
Else
    frmElpPrt.prtNewPage
    prtYICCCPT0_Form
End If
Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub

Public Sub prtYICCCPT0_NewLine()
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
If XPrt.CurrentY + 300 > prtMaxY Then
    prtMaxY_YICCCPT0 = XPrt.CurrentY
                prtYICCCPT0_Col
                frmElpPrt.prtNewPage
                prtYICCCPT0_Form
End If

End Sub



'---------------------------------------------------------
Public Sub prtYICCCPT0_Col()
'---------------------------------------------------------

XPrt.DrawWidth = 1
prtLineColor = RGB(0, 123, 141)

XPrt.Line (prtMinX + 450, prtMinY_YICCCPT0)-(prtMinX + 450, prtMaxY_YICCCPT0), prtLineColor
XPrt.Line (prtMinX + 2450, prtMinY_YICCCPT0)-(prtMinX + 2450, prtMaxY_YICCCPT0), prtLineColor
XPrt.Line (prtMinX + 6650, prtMinY_YICCCPT0)-(prtMinX + 6650, prtMaxY_YICCCPT0), prtLineColor
XPrt.Line (prtMaxX, prtMinY_YICCCPT0)-(prtMaxX, prtMaxY_YICCCPT0), prtLineColor

XPrt.DrawWidth = 3
XPrt.Line (prtMinX + 6650, prtMinY_YICCCPT0)-(prtMinX + 6650, prtMaxY_YICCCPT0), prtLineColor
XPrt.Line (prtMinX + 12650, prtMinY_YICCCPT0)-(prtMinX + 12650, prtMaxY_YICCCPT0), prtLineColor
XPrt.Line (prtMinX + 9650, prtMinY_YICCCPT0)-(prtMinX + 9650, prtMaxY_YICCCPT0), prtLineColor

XPrt.DrawWidth = 12
XPrt.Line (prtMinX + 8150, prtMinY_YICCCPT0)-(prtMinX + 8150, prtMaxY_YICCCPT0), prtLineColor
XPrt.Line (prtMinX + 14150, prtMinY_YICCCPT0)-(prtMinX + 14150, prtMaxY_YICCCPT0), prtLineColor
End Sub

Public Sub prtYICCCPT0_Form()
Dim wId As String
Dim X As String

blnLine_Trame = False
XPrt.FontSize = 7
XPrt.FontBold = True
XPrt.DrawWidth = 2

XPrt.FontSize = 8
prtFillColor = RGB(0, 123, 141)
XPrt.ForeColor = vbWhite
Call frmElpPrt.prtTrame_Color(prtMinX, prtMinY, prtMaxX, prtMinY + prtHeaderHeight, " ")
'---------------------------------------------------------
prtFillColor = RGB(200, 255, 255)

XPrt.CurrentY = prtMinY + 50

XPrt.CurrentX = prtMinX: XPrt.Print "Devise";
XPrt.CurrentX = prtMinX + 500: XPrt.Print "Compte";
XPrt.CurrentX = prtMinX + 2500: XPrt.Print "Intitulé";
X = "Provision M-1"
XPrt.CurrentX = prtMinX + 8100 - XPrt.TextWidth(X): XPrt.Print X;
X = "Solde M-1"
XPrt.CurrentX = prtMinX + 9600 - XPrt.TextWidth(X): XPrt.Print X;
X = "Mvt Débit"
XPrt.CurrentX = prtMinX + 11100 - XPrt.TextWidth(X): XPrt.Print X;
X = "Mvt Crédit"
XPrt.CurrentX = prtMinX + 12600 - XPrt.TextWidth(X): XPrt.Print X;
X = "Solde M"
XPrt.CurrentX = prtMinX + 14100 - XPrt.TextWidth(X): XPrt.Print X;
X = "Provision M"
XPrt.CurrentX = prtMinX + 15600 - XPrt.TextWidth(X): XPrt.Print X;


XPrt.FontSize = 8
XPrt.FontBold = False
XPrt.ForeColor = vbBlack
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentY = XPrt.CurrentY + 50
prtMinY_YICCCPT0 = prtMinY + prtHeaderHeight '* 2


End Sub
Public Sub prtYICCCPT0_Line(lYICCCPT0 As typeYICCCPT0, Total As typeYICCMVT0, lsoldeD As typeYICCMVT0, lsoldeF As typeYICCMVT0, blnDetail As Boolean, lYICCMVT0_Nb As Long, blnAvance As Boolean)
Dim X As String, xSQL As String, wSoldeD As Currency, wSoldeF As Currency
Dim wColor_Cr As Long

prtYICCCPT0_NewLine
If blnDetail Then
    wColor_Cr = vbBlack
Else
    wColor_Cr = vbBlue
End If

If lsoldeD.ICCMVTRBT <> 0 Or lsoldeD.ICCMVTPRO <> 0 Then
    prtFillColor = RGB(255, 230, 240): blnLine_Trame = True
Else
    If lYICCMVT0_Nb = 0 Then
        prtFillColor = RGB(250, 250, 250)
    Else
        prtFillColor = RGB(245, 255, 245): blnLine_Trame = True
    End If
    If blnDetail Then blnLine_Trame = True
End If

If blnDetail Then
    If lYICCMVT0_Nb = 0 Then
        XPrt.CurrentY = XPrt.CurrentY + 50
    Else
        XPrt.CurrentY = XPrt.CurrentY + 100
    End If
End If


If blnLine_Trame Then
    blnLine_Trame = False
    
    Call frmElpPrt.prtTrame_Color(prtMinX, XPrt.CurrentY - 50, prtMaxX, XPrt.CurrentY - 50 + prtlineHeight, " ")
Else
    blnLine_Trame = True
End If
'If blnDetail Then
'    prtLineColor = vbBlue
'    XPrt.Line (prtMinX, XPrt.CurrentY - 50)-(prtMaxX, XPrt.CurrentY - 50), prtLineColor
'End If


XPrt.ForeColor = vbBlack 'prtForeColor

'wColor = XPrt.ForeColor
XPrt.FontSize = 8

XPrt.CurrentX = prtMinX
XPrt.CurrentX = prtMinX + 50: XPrt.Print lYICCCPT0.ICCCPTDEV;
XPrt.CurrentX = prtMinX + 500: XPrt.Print lYICCCPT0.ICCCPTCOM;
xSQL = "select COMPTEINT from " & paramIBM_Library_SABSPE & ".YBIACPT0 " _
     & " Where COMPTECOM = '" & lYICCCPT0.ICCCPTCOM & "'"
Set rsSab = cnsab.Execute(xSQL)
If Not rsSab.EOF Then
    XPrt.CurrentX = prtMinX + 2500: XPrt.Print rsSab("COMPTEINT");
End If

'___________________________________________________________________________________
wSoldeD = lsoldeD.ICCMVTTDB + lsoldeD.ICCMVTTCR
'$jpl 2010-09-06  If lsoldeD.ICCMVTRBT <> 0 Then
'$jpl 2010-09-06    prtFillColor = RGB(255, 180, 255)
'$jpl 2010-09-06    Call frmElpPrt.prtTrame_Color(prtMinX + 6700, XPrt.CurrentY - 50, prtMinX + 8150, XPrt.CurrentY + prtlineHeight - 50, " ")
'$jpl 2010-09-06End If

If Not blnAvance Then
    If Total.ICCMVTRBT + lsoldeD.ICCMVTTDB + lsoldeD.ICCMVTTCR <> 0 Then
        prtFillColor = RGB(255, 226, 200)
        Call frmElpPrt.prtTrame_Color(prtMinX + 8200, XPrt.CurrentY - 50, prtMinX + 9650, XPrt.CurrentY + prtlineHeight - 50, " ")
        prtFillColor = RGB(255, 180, 255)
        Call frmElpPrt.prtTrame_Color(prtMinX + 6700, XPrt.CurrentY - 50, prtMinX + 8150, XPrt.CurrentY + prtlineHeight - 50, " ")
    End If
End If
If Total.ICCMVTRBT <> 0 Then
    XPrt.ForeColor = vbMagenta 'IIf(Total.ICCMVTRBT < 0, vbMagenta, RGB(32, 96, 32))
    X = Format$(Total.ICCMVTRBT, "### ### ### ##0.00")
    XPrt.CurrentX = prtMinX + 8100 - XPrt.TextWidth(X): XPrt.Print X;
End If

If wSoldeD <> 0 Then
    XPrt.ForeColor = IIf(wSoldeD < 0, vbRed, wColor_Cr)
    X = Format$(wSoldeD, "### ### ### ##0.00")
    XPrt.CurrentX = prtMinX + 9600 - XPrt.TextWidth(X): XPrt.Print X;
End If

If Total.ICCMVTTDB <> 0 Then
    XPrt.ForeColor = IIf(Total.ICCMVTTDB < 0, vbRed, wColor_Cr)
    X = Format$(Total.ICCMVTTDB, "### ### ### ##0.00")
    XPrt.CurrentX = prtMinX + 11100 - XPrt.TextWidth(X): XPrt.Print X;
End If
If Total.ICCMVTTCR <> 0 Then
    XPrt.ForeColor = IIf(Total.ICCMVTTCR < 0, vbRed, wColor_Cr)
    X = Format$(Total.ICCMVTTCR, "### ### ### ##0.00")
    XPrt.CurrentX = prtMinX + 12600 - XPrt.TextWidth(X): XPrt.Print X;
End If

'$jpl 2010-09-06If lsoldeD.ICCMVTPRO <> 0 Then
'$jpl 2010-09-06    prtFillColor = RGB(255, 180, 255)
'$jpl 2010-09-06    Call frmElpPrt.prtTrame_Color(prtMinX + 14200, XPrt.CurrentY - 50, prtMinX + 15650, XPrt.CurrentY + prtlineHeight - 50, " ")
'$jpl 2010-09-06End If

If Not blnAvance Then
    If Total.ICCMVTPRO - lsoldeF.ICCMVTTDB - lsoldeF.ICCMVTTCR <> 0 Then
        prtFillColor = RGB(255, 226, 200)
        Call frmElpPrt.prtTrame_Color(prtMinX + 12700, XPrt.CurrentY - 50, prtMinX + 14150, XPrt.CurrentY + prtlineHeight - 50, " ")
        prtFillColor = RGB(255, 180, 255)
        Call frmElpPrt.prtTrame_Color(prtMinX + 14200, XPrt.CurrentY - 50, prtMinX + 15650, XPrt.CurrentY + prtlineHeight - 50, " ")
    End If
End If

If Total.ICCMVTPRO <> 0 Then
    XPrt.ForeColor = RGB(32, 96, 32) 'IIf(Total.ICCMVTPRO < 0, vbMagenta, RGB(32, 96, 32))
    X = Format$(Total.ICCMVTPRO, "### ### ### ##0.00")
    XPrt.CurrentX = prtMinX + 15600 - XPrt.TextWidth(X): XPrt.Print X;
End If

wSoldeF = lsoldeF.ICCMVTTDB + lsoldeF.ICCMVTTCR
If wSoldeF <> wSoldeD + Total.ICCMVTTDB + Total.ICCMVTTCR Then
    prtFillColor = RGB(255, 160, 255)
    Call frmElpPrt.prtTrame_Color(prtMinX + 12700, XPrt.CurrentY - 50, prtMinX + 14150, XPrt.CurrentY + prtlineHeight - 50, " ")
End If
If wSoldeF <> 0 Then
    XPrt.ForeColor = IIf(wSoldeF < 0, vbRed, wColor_Cr)
    X = Format$(wSoldeF, "### ### ### ##0.00")
    XPrt.CurrentX = prtMinX + 14100 - XPrt.TextWidth(X):     XPrt.Print X;
End If
End Sub
Public Sub prtYICCCPT0_Line_Detail_xlsManual(lYICCMVT0 As typeYICCMVT0, ByRef currentRow As Long, wsExcel As Excel.Worksheet, ByRef comptageRows As Long, maxRows As Long, maxRowsPlus As Long)
Dim X As String
Dim wColor As Long

Call prtYICCCPT0_NewLine_xlsManual(currentRow, wsExcel, comptageRows, maxRows, maxRowsPlus)
wColor = vbBlue
wsExcel.Cells(currentRow, 3) = lYICCMVT0.ICCMVTSER & " " & lYICCMVT0.ICCMVTSSE & " " _
              & lYICCMVT0.ICCMVTOPE & " " & lYICCMVT0.ICCMVTNAT & " " & lYICCMVT0.ICCMVTDOS & " " & lYICCMVT0.ICCMVTEVE
wsExcel.Cells(currentRow, 3).Font.Color = wColor
wsExcel.Cells(currentRow, 4) = dateImp10(lYICCMVT0.ICCMVTAMJ)
wsExcel.Cells(currentRow, 4).Font.Color = wColor
'______________________________________________________
Select Case lYICCMVT0.ICCMVTOPE
    Case "EMP", "EM1":
        If lYICCMVT0.ICCMVTRBT <> lYICCMVT0.ICCMVTTDB Then
            prtFillColor = RGB(255, 160, 255)
            wsExcel.Cells(currentRow, 6).Interior.Color = prtFillColor
        End If
        If lYICCMVT0.ICCMVTPRO <> lYICCMVT0.ICCMVTTCR Then
            prtFillColor = RGB(255, 160, 255)
            wsExcel.Cells(currentRow, 10).Interior.Color = prtFillColor
        End If
    Case Else:
        If Mid$(lYICCMVT0.ICCMVTOPE, 1, 1) <> "*" And Mid$(lYICCMVT0.ICCMVTOPE, 1, 1) <> " " Then
            If lYICCMVT0.ICCMVTRBT <> lYICCMVT0.ICCMVTTCR Then
                prtFillColor = RGB(255, 160, 255)
            wsExcel.Cells(currentRow, 6).Interior.Color = prtFillColor
            End If
            If lYICCMVT0.ICCMVTPRO <> lYICCMVT0.ICCMVTTDB Then
                prtFillColor = RGB(255, 160, 255)
            wsExcel.Cells(currentRow, 10).Interior.Color = prtFillColor
            End If
        End If
End Select
'___________________________________________________________________________________
If lYICCMVT0.ICCMVTRBT <> 0 Then
    wColor = vbMagenta
    X = Format$(lYICCMVT0.ICCMVTRBT, "### ### ### ##0.00")
    wsExcel.Cells(currentRow, 5) = X
    wsExcel.Cells(currentRow, 5).Font.Color = wColor
End If
If lYICCMVT0.ICCMVTTDB <> 0 Then
    wColor = IIf(lYICCMVT0.ICCMVTTDB < 0, vbRed, vbBlue)
    X = Format$(lYICCMVT0.ICCMVTTDB, "### ### ### ##0.00")
    wsExcel.Cells(currentRow, 7) = X
    wsExcel.Cells(currentRow, 7).Font.Color = wColor
End If
If lYICCMVT0.ICCMVTTCR <> 0 Then
    wColor = IIf(lYICCMVT0.ICCMVTTCR < 0, vbRed, vbBlue)
    X = Format$(lYICCMVT0.ICCMVTTCR, "### ### ### ##0.00")
    wsExcel.Cells(currentRow, 8) = X
    wsExcel.Cells(currentRow, 8).Font.Color = wColor
End If
If lYICCMVT0.ICCMVTPRO <> 0 Then
    wColor = RGB(32, 96, 32)
    X = Format$(lYICCMVT0.ICCMVTPRO, "### ### ### ##0.00")
    wsExcel.Cells(currentRow, 10) = X
    wsExcel.Cells(currentRow, 10).Font.Color = wColor
End If

End Sub


