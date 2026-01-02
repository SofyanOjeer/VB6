Attribute VB_Name = "prtBIA_PDC"
Option Explicit

Dim curX As Currency
Dim xSQL As String, V

Dim mEtat_PDCPOS As String, mEtat_Exclure_HB As String, mEtat_Exclure_PDCMVTKCUT As String

Dim Height8_6 As Integer

Dim prtMinY_PDCMVT As Integer, prtMaxY_PDCMVT As Integer, blnPDCMVT As Boolean

Dim prtMinY_PDCPOS As Integer, prtMaxY_PDCPOS As Integer

Dim prtMinY_PDCLOG As Integer, prtMaxY_PDCLOG As Integer, blnPDCLOG As Boolean

Dim blnPDCOPE As Boolean, blnPDCPOS As Boolean


Public Sub prtBIA_PDC_MT_xlsManual(lcurX As Currency, laColonne As Long, ByRef currentRow As Long, wsExcel As Excel.Worksheet)
Dim X As String
Dim wColor As Long

If lcurX <> 0 Then
    If lcurX < 0 Then
        wColor = vbRed
    Else
        wColor = vbBlue
    End If
    X = Format$(lcurX, "### ### ### ##0.00")
    wsExcel.Cells(currentRow, laColonne) = X
    wsExcel.Cells(currentRow, laColonne).Font.Color = wColor
End If

End Sub


Public Sub prtBIA_PDC_NewLine_xlsManual(ByRef currentRow As Long, wsExcel As Excel.Worksheet, butoir As String, ByRef comptageRows As Long, maxRows As Long, maxRowsPlus As Long)

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
    Range("A4" & ":" & butoir & "4").Select
    Selection.Copy
    Range("A" & CStr(currentRow)).Select
    ActiveSheet.Paste
    wsExcel.Rows(currentRow).RowHeight = 6
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
    Range("A5" & ":" & butoir & "5").Select
    Selection.Copy
    Range("A" & CStr(currentRow)).Select
    ActiveSheet.Paste

End Sub

Public Sub prtBIA_PDCLOG_Line_xlsManual(lYPDCLOG0 As typeYPDCLOG0, ByRef currentRow As Long, wsExcel As Excel.Worksheet, ByRef comptageRowsZero As Long, maxRowsZero As Long, maxRowsPlusZero As Long)
Dim K As Long, X As String
Dim wColor As Long

    Call prtBIA_PDC_NewLine_xlsManual(currentRow, wsExcel, "F", comptageRowsZero, maxRowsZero, maxRowsPlusZero)
    wColor = vbBlack
    If Mid$(lYPDCLOG0.PDCLOGNAT, 3, 1) <> " " Then wColor = vbRed
    If Mid$(lYPDCLOG0.PDCLOGNAT, 1, 1) = "7" Then wColor = vbMagenta
    If Mid$(lYPDCLOG0.PDCLOGNAT, 1, 2) = "5=" Then wColor = vbMagenta
    wsExcel.Cells(currentRow, 1) = "'" & dateImp(lYPDCLOG0.PDCLOGDTR)
    wsExcel.Cells(currentRow, 1).Font.Color = wColor
    wsExcel.Cells(currentRow, 2) = "'" & dateImp(lYPDCLOG0.PDCLOGUAMJ) & "  " & timeImp(lYPDCLOG0.PDCLOGUHMS) & " - " & lYPDCLOG0.PDCLOGUSEQ
    wsExcel.Cells(currentRow, 2).Font.Color = wColor
    wsExcel.Cells(currentRow, 3) = "'" & lYPDCLOG0.PDCLOGSTA & " " & lYPDCLOG0.PDCLOGNAT
    wsExcel.Cells(currentRow, 3).Font.Color = wColor
    wsExcel.Cells(currentRow, 4) = lYPDCLOG0.PDCLOGTXT
    wsExcel.Cells(currentRow, 4).Font.Color = wColor
    If lYPDCLOG0.PDCLOGPIE <> 0 Then
        wsExcel.Cells(currentRow, 5) = lYPDCLOG0.PDCLOGPIE & " - " & lYPDCLOG0.PDCLOGECR
        wsExcel.Cells(currentRow, 5).Font.Color = wColor
    End If
    wsExcel.Cells(currentRow, 6) = lYPDCLOG0.PDCLOGUUSR
    wsExcel.Cells(currentRow, 6).Font.Color = wColor

End Sub

Public Sub prtBIA_PDCMVT_Line_xlsManual(lYPDCMVT0 As typeYPDCMVT0, lKCUT As String, ByRef currentRow As Long, wsExcel As Excel.Worksheet, ByRef comptageRows As Long, maxRows As Long, maxRowsPlus As Long)
Dim K As Long, X As String
Dim devFixing As Currency, devPP As Currency, eurPP As Currency
Dim wPDCMVTMTE As Currency, wPDCMVTMTD As Currency

Call prtBIA_PDC_NewLine_xlsManual(currentRow, wsExcel, "I", comptageRows, maxRows, maxRowsPlus)
wPDCMVTMTE = -lYPDCMVT0.PDCMVTMTE
wPDCMVTMTD = -lYPDCMVT0.PDCMVTMTD
wsExcel.Cells(currentRow, 1) = "'" & dateImp(lYPDCMVT0.PDCMVTDTR) & "   " & lYPDCMVT0.PDCMVTDEV
Call prtBIA_PDC_MT_xlsManual(wPDCMVTMTE, 2, currentRow, wsExcel)
Call prtBIA_PDC_MT_xlsManual(wPDCMVTMTD, 3, currentRow, wsExcel)
X = Format$(lYPDCMVT0.PDCMVTTAUX, "### ##0.000000")
wsExcel.Cells(currentRow, 4) = X
X = "'" & dateImp(lYPDCMVT0.PDCMVTDVA)
wsExcel.Cells(currentRow, 5) = X
X = lYPDCMVT0.PDCMVTSTA2
X = X & " " & lYPDCMVT0.PDCMVTSER
X = X & " " & lYPDCMVT0.PDCMVTSSE
X = X & " " & lYPDCMVT0.PDCMVTOPEC
X = X & Format$(lYPDCMVT0.PDCMVTOPEN, "### ##0")
wsExcel.Cells(currentRow, 6) = X
If lYPDCMVT0.PDCMVTKCUT = " " Then
    wsExcel.Cells(currentRow, 7) = lYPDCMVT0.PDCMVTCLI
Else
    If lKCUT = "" Then
        wsExcel.Cells(currentRow, 7) = lYPDCMVT0.PDCMVTCLI & " cut"
    Else
        wsExcel.Cells(currentRow, 7) = lKCUT
    End If
    wsExcel.Cells(currentRow, 7).Font.Color = vbMagenta
End If
wsExcel.Cells(currentRow, 8) = lYPDCMVT0.PDCMVTSTA & " " & lYPDCMVT0.PDCMVTPIE & " - " & lYPDCMVT0.PDCMVTECR
wsExcel.Cells(currentRow, 9) = lYPDCMVT0.PDCMVTCPT
End Sub

Public Sub prtBIA_PDCMVT_POS_xlsManual(lYPDCPOS0 As typeYPDCPOS0, lFct As String, ByRef currentRow As Long, wsExcel As Excel.Worksheet, ByRef comptageRows As Long, maxRows As Long, maxRowsPlus As Long)
Dim X As String
Dim wY As Integer

    If currentRow >= maxRows + maxRowsPlus Then
        If comptageRows >= maxRows Then
            Call insere_entete_page(wsExcel, "1:3", 3, currentRow)
            comptageRows = 3
            currentRow = currentRow + 3
        End If
    End If
    comptageRows = comptageRows + 1
    currentRow = currentRow + 1
    Range("A4:I4").Select
    Selection.Copy
    Range("A" & CStr(currentRow)).Select
    ActiveSheet.Paste
    wsExcel.Rows(currentRow).RowHeight = 6
    If currentRow >= maxRows + maxRowsPlus Then
        If comptageRows >= maxRows Then
            Call insere_entete_page(wsExcel, "1:3", 3, currentRow)
            comptageRows = 3
            currentRow = currentRow + 3
        End If
    End If
    comptageRows = comptageRows + 1
    currentRow = currentRow + 1
    Range("A7:I7").Select
    Selection.Copy
    Range("A" & CStr(currentRow)).Select
    ActiveSheet.Paste
    wsExcel.Cells(currentRow, 1) = "'" & dateImp(lYPDCPOS0.PDCPOSDTR) & "  " & lYPDCPOS0.PDCPOSDEV
    Call prtBIA_PDC_MT_xlsManual(lYPDCPOS0.PDCPOSPOSE, 2, currentRow, wsExcel)
    Call prtBIA_PDC_MT_xlsManual(lYPDCPOS0.PDCPOSPOSD, 3, currentRow, wsExcel)
    X = Format$(lYPDCPOS0.PDCPOSPRIX, "### ##0.000000")
    wsExcel.Cells(currentRow, 4) = X

End Sub

Public Sub prtBIA_PDCOPE_YPDCOPE0_xlsManual(lYPDCOPE0 As typeYPDCOPE0, ByRef currentRow As Long, wsExcel As Excel.Worksheet, ByRef comptageRows As Long, maxRows As Long, maxRowsPlus As Long)
Dim K As Long, X As String
Dim wPDCMVTMTE As Currency, wPDCMVTMTD As Currency

    Call prtBIA_PDC_NewLine_xlsManual(currentRow, wsExcel, "J", comptageRows, maxRows, maxRowsPlus)
    If lYPDCOPE0.PDCOPESENS = "A" Then
        wPDCMVTMTE = -lYPDCOPE0.PDCOPEMTD1
        wPDCMVTMTD = -lYPDCOPE0.PDCOPEMTD2
    Else
        wPDCMVTMTE = -lYPDCOPE0.PDCOPEMTD1
        wPDCMVTMTD = -lYPDCOPE0.PDCOPEMTD2
    End If
    wsExcel.Cells(currentRow, 1) = "'" & dateImp(lYPDCOPE0.PDCOPEIAMJ) & "   " & lYPDCOPE0.PDCOPESENS
    Call prtBIA_PDC_MT_xlsManual(wPDCMVTMTE, 2, currentRow, wsExcel)
    wsExcel.Cells(currentRow, 3) = lYPDCOPE0.PDCOPEDEV1
    Call prtBIA_PDC_MT_xlsManual(wPDCMVTMTD, 4, currentRow, wsExcel)
    wsExcel.Cells(currentRow, 5) = lYPDCOPE0.PDCOPEDEV2
    X = Format$(lYPDCOPE0.PDCOPETAUX, "### ##0.000000")
    wsExcel.Cells(currentRow, 6) = X
    X = "'" & dateImp(lYPDCOPE0.PDCOPEDVA + 19000000)
    wsExcel.Cells(currentRow, 7) = X
    X = lYPDCOPE0.PDCOPESER
    X = X & " " & lYPDCOPE0.PDCOPESSE
    X = X & " " & lYPDCOPE0.PDCOPEOPEC & "-" & lYPDCOPE0.PDCOPEID
    X = X & " " & Format$(lYPDCOPE0.PDCOPEOPEN, "### ##0")
    wsExcel.Cells(currentRow, 8) = X
    wsExcel.Cells(currentRow, 9) = lYPDCOPE0.PDCOPECLI
    X = lYPDCOPE0.PDCOPESTA & lYPDCOPE0.PDCOPESTA2 & lYPDCOPE0.PDCOPESTA3
    If lYPDCOPE0.PDCOPEREF <> 0 Then
        X = X & " " & lYPDCOPE0.PDCOPEREF
    End If
    If lYPDCOPE0.PDCOPEIAMJ <> lYPDCOPE0.PDCOPEDTR Then
        X = X & " report"
    End If
    wsExcel.Cells(currentRow, 10) = X

End Sub

Public Sub prtBIA_PDCOPE_ZCHGOPE0_xlsManual(lZCHGOPE0 As typeZCHGOPE0, lTxt As String, ByRef currentRow As Long, wsExcel As Excel.Worksheet, ByRef comptageRows As Long, maxRows As Long, maxRowsPlus As Long)
Dim K As Long, X As String
Dim wPDCMVTMTE As Currency, wPDCMVTMTD As Currency
    
    Call prtBIA_PDC_NewLine_xlsManual(currentRow, wsExcel, "J", comptageRows, maxRows, maxRowsPlus)
    If lZCHGOPE0.CHGOPESEN = "A" Then
        wPDCMVTMTE = lZCHGOPE0.CHGOPEMO1
        wPDCMVTMTD = -lZCHGOPE0.CHGOPEMO2
    Else
        wPDCMVTMTE = -lZCHGOPE0.CHGOPEMO1
        wPDCMVTMTD = lZCHGOPE0.CHGOPEMO2
    End If
    wsExcel.Cells(currentRow, 1) = "'" & dateImp(lZCHGOPE0.CHGOPECRE + 19000000) & "   " & lZCHGOPE0.CHGOPESEN
    Call prtBIA_PDC_MT_xlsManual(wPDCMVTMTE, 2, currentRow, wsExcel)
    wsExcel.Cells(currentRow, 3) = lZCHGOPE0.CHGOPEDE1
    Call prtBIA_PDC_MT_xlsManual(wPDCMVTMTD, 4, currentRow, wsExcel)
    wsExcel.Cells(currentRow, 5) = lZCHGOPE0.CHGOPEDE2
    If lZCHGOPE0.CHGOPECO3 <> 0 Then
        X = Format$(lZCHGOPE0.CHGOPECO3, "### ##0.000000")
    Else
        X = Format$(lZCHGOPE0.CHGOPECO1, "### ##0.000000")
    End If
    wsExcel.Cells(currentRow, 6) = X
    X = "'" & dateImp(lZCHGOPE0.CHGOPEDT1 + 19000000)
    wsExcel.Cells(currentRow, 7) = X
    X = lZCHGOPE0.CHGOPESER
    X = X & " " & lZCHGOPE0.CHGOPESSE
    X = X & " " & lZCHGOPE0.CHGOPEOPE
    X = X & " " & Format$(lZCHGOPE0.CHGOPEDOS, "### ##0")
    wsExcel.Cells(currentRow, 8) = X
    wsExcel.Cells(currentRow, 9) = lZCHGOPE0.CHGOPECON
    wsExcel.Cells(currentRow, 10) = lTxt

End Sub

'---------------------------------------------------------
Public Sub prtBIA_PDCPOS_Form()
'---------------------------------------------------------
Dim X As String
blnPDCPOS = True

XPrt.DrawWidth = 3
XPrt.FontSize = 14: XPrt.FontBold = True
'XPrt.CurrentY = prtMinY + prtlineHeight * 3
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.FontUnderline = True
XPrt.ForeColor = RGB(0, 123, 141)

XPrt.CurrentX = prtMinX + 8000 - XPrt.TextWidth(mEtat_PDCPOS) / 2
XPrt.Print mEtat_PDCPOS;
XPrt.FontUnderline = False
XPrt.FontSize = 10
XPrt.ForeColor = vbMagenta

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
If mEtat_Exclure_HB <> "" Then
    XPrt.CurrentX = prtMinX + 8000 - XPrt.TextWidth(mEtat_Exclure_HB) / 2
    XPrt.Print mEtat_Exclure_HB;
End If
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
If mEtat_Exclure_PDCMVTKCUT <> "" Then
    XPrt.CurrentX = prtMinX + 8000 - XPrt.TextWidth(mEtat_Exclure_PDCMVTKCUT) / 2
    XPrt.Print mEtat_Exclure_PDCMVTKCUT;
End If

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.FontBold = True
prtCurrentY = XPrt.CurrentY
prtMinY_PDCPOS = prtCurrentY
prtFillColor = RGB(0, 123, 141)
XPrt.ForeColor = vbWhite

Call frmElpPrt.prtTrame_Color(prtMinX, prtCurrentY, prtMaxX, prtCurrentY + prtHeaderHeight, " ")
'---------------------------------------------------------
XPrt.CurrentY = prtCurrentY + 50

XPrt.CurrentX = prtMinX + 100: XPrt.Print "Devise";
X = "Position EUR"
XPrt.CurrentX = prtMinX + 3000 - 100 - XPrt.TextWidth(X): XPrt.Print X;
X = "Position DEV"
XPrt.CurrentX = prtMinX + 4400 - 100 - XPrt.TextWidth(X): XPrt.Print X;
X = "Prix Position"
XPrt.CurrentX = prtMinX + 5800 - 100 - XPrt.TextWidth(X): XPrt.Print X;
X = "Fixing"
XPrt.CurrentX = prtMinX + 7200 - 100 - XPrt.TextWidth(X): XPrt.Print X;
X = "PP Devises"
XPrt.CurrentX = prtMinX + 8600 - 100 - XPrt.TextWidth(X): XPrt.Print X;
X = "Reval Jour"
XPrt.CurrentX = prtMinX + 10000 - 100 - XPrt.TextWidth(X): XPrt.Print X;
X = "PP Jour"
XPrt.CurrentX = prtMinX + 11400 - 100 - XPrt.TextWidth(X): XPrt.Print X;
X = "PP J-1"
XPrt.CurrentX = prtMinX + 12800 - 100 - XPrt.TextWidth(X): XPrt.Print X;
X = "Fixing J-1"
XPrt.CurrentX = prtMinX + 14400 - 100 - XPrt.TextWidth(X): XPrt.Print X;
X = "RPC"
XPrt.CurrentX = prtMaxX - 100 - XPrt.TextWidth(X): XPrt.Print X;
XPrt.ForeColor = vbBlack
prtFillColor = prtFillColor_Standard

XPrt.FontSize = 8: XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtHeaderHeight - prtlineHeight

End Sub
'---------------------------------------------------------
Public Sub prtBIA_PDCMVT_Form()
'---------------------------------------------------------


Dim X As String
blnPDCMVT = True

prtBIA_PDC_Etat

prtMinY_PDCMVT = XPrt.CurrentY
prtMaxY_PDCMVT = XPrt.CurrentY + prtHeaderHeight
XPrt.DrawWidth = 3
XPrt.FontSize = 8: XPrt.FontBold = True
prtFillColor = RGB(0, 123, 141)
XPrt.ForeColor = vbWhite

Call frmElpPrt.prtTrame_Color(prtMinX, XPrt.CurrentY, prtMaxX, XPrt.CurrentY + prtHeaderHeight, " ")
'---------------------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + 50

XPrt.CurrentX = prtMinX + 50: XPrt.Print "Date TRT";
X = "Montant EUR"
XPrt.CurrentX = prtMinX + 3200 - 100 - XPrt.TextWidth(X): XPrt.Print X;
X = "Montant DEV"
XPrt.CurrentX = prtMinX + 4800 - 100 - XPrt.TextWidth(X): XPrt.Print X;
X = "Taux"
XPrt.CurrentX = prtMinX + 6400 - 100 - XPrt.TextWidth(X): XPrt.Print X;
X = "Date valeur"
XPrt.CurrentX = prtMinX + 8000 - 100 - XPrt.TextWidth(X): XPrt.Print X;
X = "Opération"
XPrt.CurrentX = prtMinX + 9600 - 100 - XPrt.TextWidth(X): XPrt.Print X;
X = "Client"
XPrt.CurrentX = prtMinX + 11200 + 100: XPrt.Print X;
X = "Pièce comptable"
XPrt.CurrentX = prtMinX + 12800 + 100: XPrt.Print X;
X = "Compte"
XPrt.CurrentX = prtMaxX - 100 - XPrt.TextWidth(X): XPrt.Print X;
XPrt.ForeColor = vbBlack
prtFillColor = prtFillColor_Standard

XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtHeaderHeight - prtlineHeight
End Sub
'---------------------------------------------------------
Public Sub prtBIA_PDCOPE_Form()
'---------------------------------------------------------


Dim X As String
blnPDCOPE = True

prtBIA_PDC_Etat

prtMinY_PDCMVT = XPrt.CurrentY
prtMaxY_PDCMVT = XPrt.CurrentY + prtHeaderHeight
XPrt.DrawWidth = 3
XPrt.FontSize = 8: XPrt.FontBold = True
prtFillColor = RGB(0, 123, 141)
XPrt.ForeColor = vbWhite

Call frmElpPrt.prtTrame_Color(prtMinX, XPrt.CurrentY, prtMaxX, XPrt.CurrentY + prtHeaderHeight, " ")
'---------------------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + 50

XPrt.CurrentX = prtMinX + 50: XPrt.Print "Date CRE";
X = "Montant DEV"
XPrt.CurrentX = prtMinX + 3700 - 100 - XPrt.TextWidth(X): XPrt.Print X;
X = "Montant DEV"
XPrt.CurrentX = prtMinX + 5800 - 100 - XPrt.TextWidth(X): XPrt.Print X;
X = "Taux"
XPrt.CurrentX = prtMinX + 7400 - 100 - XPrt.TextWidth(X): XPrt.Print X;
X = "Date valeur"
XPrt.CurrentX = prtMinX + 9000 - 100 - XPrt.TextWidth(X): XPrt.Print X;
X = "Opération"
XPrt.CurrentX = prtMinX + 10600 - 100 - XPrt.TextWidth(X): XPrt.Print X;
X = "Client"
XPrt.CurrentX = prtMinX + 12200 + 100: XPrt.Print X;
X = "Commentaire"
XPrt.CurrentX = prtMinX + 13800 + 100: XPrt.Print X;
XPrt.ForeColor = vbBlack
prtFillColor = prtFillColor_Standard

XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtHeaderHeight - prtlineHeight
End Sub

Public Sub prtBIA_PDCPOS_Form_xlsManual(ByRef currentRow As Long, wsExcel As Excel.Worksheet)
blnPDCPOS = True

wsExcel.Cells(currentRow, 4) = mEtat_PDCPOS
If mEtat_Exclure_HB <> "" Then
    wsExcel.Cells(currentRow, 4) = mEtat_Exclure_HB
End If
If mEtat_Exclure_PDCMVTKCUT <> "" Then
    wsExcel.Cells(currentRow, 4) = mEtat_Exclure_PDCMVTKCUT
End If

End Sub

Public Sub prtBIA_PDCPOS_Line_xlsManual(fgX As MSFlexGrid, ByRef currentRow As Long, wsExcel As Excel.Worksheet)
Dim K As Long, X As String
Dim wColor As Long
Dim comptageRows As Long
Dim maxRows As Long
Dim maxRowsPlus As Long

comptageRows = currentRow
maxRows = 45
maxRowsPlus = 3
For K = 1 To fgX.Rows - 1
    fgX.Row = K
    If K = fgX.Rows - 1 Then
        If currentRow >= maxRows + maxRowsPlus Then
            If comptageRows >= maxRows Then
                Call insere_entete_page(wsExcel, "1:3", 3, currentRow)
                comptageRows = 3
                currentRow = currentRow + 3
            End If
        End If
        comptageRows = comptageRows + 1
        currentRow = currentRow + 1
        Range("A4:K4").Select
        Selection.Copy
        Range("A" & CStr(currentRow)).Select
        ActiveSheet.Paste
        If currentRow >= maxRows + maxRowsPlus Then
            If comptageRows >= maxRows Then
                Call insere_entete_page(wsExcel, "1:3", 3, currentRow)
                comptageRows = 3
                currentRow = currentRow + 3
            End If
        End If
        comptageRows = comptageRows + 1
        currentRow = currentRow + 1
        Range("A6:K6").Select
        Selection.Copy
        Range("A" & CStr(currentRow)).Select
        ActiveSheet.Paste
    Else
        Call prtBIA_PDC_NewLine_xlsManual(currentRow, wsExcel, "K", comptageRows, maxRows, maxRowsPlus)
    End If
    fgX.Col = 0: wColor = fgX.CellForeColor
    wsExcel.Cells(currentRow, 1) = fgX.Text
    fgX.Col = 1: X = Trim(fgX.Text): wColor = fgX.CellForeColor
    wsExcel.Cells(currentRow, 2) = X
    wsExcel.Cells(currentRow, 2).Font.Color = wColor
    fgX.Col = 2: X = Trim(fgX.Text): wColor = fgX.CellForeColor
    wsExcel.Cells(currentRow, 3) = X
    wsExcel.Cells(currentRow, 3).Font.Color = wColor
    fgX.Col = 3: X = Trim(fgX.Text): wColor = fgX.CellForeColor
    wsExcel.Cells(currentRow, 4) = X
    wsExcel.Cells(currentRow, 4).Font.Color = wColor
    fgX.Col = 4: X = Trim(fgX.Text): wColor = fgX.CellForeColor
    wsExcel.Cells(currentRow, 5) = X
    wsExcel.Cells(currentRow, 5).Font.Color = wColor
    fgX.Col = 5: X = Trim(fgX.Text): wColor = fgX.CellForeColor
    wsExcel.Cells(currentRow, 6) = X
    wsExcel.Cells(currentRow, 6).Font.Color = wColor
    fgX.Col = 6: X = Trim(fgX.Text): wColor = fgX.CellForeColor
    wsExcel.Cells(currentRow, 7) = X
    wsExcel.Cells(currentRow, 7).Font.Color = wColor
    fgX.Col = 7: X = Trim(fgX.Text): wColor = fgX.CellForeColor
    wsExcel.Cells(currentRow, 8) = X
    wsExcel.Cells(currentRow, 7).Font.Color = wColor
    fgX.Col = 8: X = Trim(fgX.Text): wColor = fgX.CellForeColor
    wsExcel.Cells(currentRow, 9) = X
    wsExcel.Cells(currentRow, 9).Font.Color = wColor
    fgX.Col = 9: X = Trim(fgX.Text): wColor = fgX.CellForeColor
    wsExcel.Cells(currentRow, 10) = X
    wsExcel.Cells(currentRow, 10).Font.Color = wColor
    fgX.Col = 10: X = Trim(fgX.Text): wColor = fgX.CellForeColor
    wsExcel.Cells(currentRow, 11) = X
    wsExcel.Cells(currentRow, 11).Font.Color = wColor
Next K

End Sub

'---------------------------------------------------------
Public Sub prtBIA_PDCTER_Form()
'---------------------------------------------------------


Dim X As String
XPrt.DrawWidth = 3
XPrt.FontSize = 8: XPrt.FontBold = True
prtFillColor = RGB(0, 123, 141)
XPrt.ForeColor = vbWhite

Call frmElpPrt.prtTrame_Color(prtMinX, XPrt.CurrentY, prtMaxX, XPrt.CurrentY + prtHeaderHeight, " ")
'---------------------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + 50

XPrt.CurrentX = prtMinX + 100: XPrt.Print "Devise";
X = "Opération"
XPrt.CurrentX = prtMinX + 1500 + 100: XPrt.Print X;
X = "D.engagement"
XPrt.CurrentX = prtMinX + 2500 + 100: XPrt.Print X;
X = "D. échéance"
XPrt.CurrentX = prtMinX + 4000 + 100: XPrt.Print X;
X = "Taux"
XPrt.CurrentX = prtMinX + 6500 - 100 - XPrt.TextWidth(X): XPrt.Print X;
X = "Débit €  "
XPrt.CurrentX = prtMinX + 8000 - 100 - XPrt.TextWidth(X): XPrt.Print X;
X = "Crédit € "
XPrt.CurrentX = prtMinX + 10000 - 100 - XPrt.TextWidth(X): XPrt.Print X;
X = "Dev "
XPrt.CurrentX = prtMinX + 10300 + 100: XPrt.Print X;
X = "Débit devise"
XPrt.CurrentX = prtMinX + 12200 - 100 - XPrt.TextWidth(X): XPrt.Print X;
X = "Crédit devise"
XPrt.CurrentX = prtMinX + 14000 - 100 - XPrt.TextWidth(X): XPrt.Print X;
X = "Réeval jour"
XPrt.CurrentX = prtMaxX - 100 - XPrt.TextWidth(X): XPrt.Print X;
XPrt.ForeColor = vbBlack
prtFillColor = prtFillColor_Standard

XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtHeaderHeight - prtlineHeight
End Sub


'---------------------------------------------------------
Public Sub prtBIA_PDCLOG_Form()
'---------------------------------------------------------


Dim X As String
blnPDCLOG = True

prtBIA_PDC_Etat

prtMinY_PDCLOG = XPrt.CurrentY
prtMaxY_PDCLOG = XPrt.CurrentY + prtHeaderHeight
XPrt.DrawWidth = 3
XPrt.FontSize = 8: XPrt.FontBold = True
prtFillColor = RGB(0, 123, 141)
XPrt.ForeColor = vbWhite

Call frmElpPrt.prtTrame_Color(prtMinX, XPrt.CurrentY, prtMaxX, XPrt.CurrentY + prtHeaderHeight, " ")
'---------------------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + 50

XPrt.CurrentX = prtMinX + 50: XPrt.Print "Date CPT";
XPrt.CurrentX = prtMinX + 1600 + 100: XPrt.Print "mise à jour le";
XPrt.CurrentX = prtMinX + 4800 + 100: XPrt.Print "Nature";
XPrt.CurrentX = prtMinX + 6400 + 100: XPrt.Print "Libellé";
XPrt.CurrentX = prtMinX + 12800 + 100: XPrt.Print "Pièce comptable";
X = "màj par"
XPrt.CurrentX = prtMaxX - 100 - XPrt.TextWidth(X): XPrt.Print X;
XPrt.ForeColor = vbBlack
prtFillColor = prtFillColor_Standard

XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtHeaderHeight - prtlineHeight
End Sub


'---------------------------------------------------------
Public Sub prtBIA_PDCPOS_Cumul(lYPDCPOS0 As typeYPDCPOS0, blnDeviseU As Boolean)
'---------------------------------------------------------
Dim X As String

prtBIA_PDC_NewLine

XPrt.DrawWidth = 3
XPrt.FontSize = 10: XPrt.FontBold = True
prtMaxY_PDCPOS = XPrt.CurrentY + prtHeaderHeight
prtFillColor = RGB(180, 255, 255)

Call frmElpPrt.prtTrame_Color(prtMinX, XPrt.CurrentY, prtMaxX, XPrt.CurrentY + prtHeaderHeight, "B")
'---------------------------------------------------------
If Not blnDeviseU Then
    XPrt.CurrentY = XPrt.CurrentY + 50
    XPrt.CurrentX = prtMinX + 50: XPrt.Print "EUR";
    
    Call prtBIA_PDC_MT(prtMinX + 3200 - 100, lYPDCPOS0.PDCPOSPOSE)
    
    Call prtBIA_PDC_MT(prtMinX + 12800 - 100, lYPDCPOS0.PDCPOSPOSD)
    
    Call prtBIA_PDC_MT(prtMinX + 14400 - 100, lYPDCPOS0.PDCPOSPNL)
    
    Call prtBIA_PDC_MT(prtMaxX - 100, lYPDCPOS0.PDCPOSRPC)
End If
XPrt.FontSize = 8: XPrt.FontBold = False

prtFillColor = prtFillColor_Standard
prtBIA_PDCPOS_Col



End Sub

'---------------------------------------------------------
Public Sub prtBIA_PDCMVT_POS(lYPDCPOS0 As typeYPDCPOS0, lFct As String)
'---------------------------------------------------------
Dim X As String
Dim wY As Integer

prtBIA_PDC_NewLine
If lFct <> "B" Then prtBIA_PDC_NewLine

XPrt.DrawWidth = 3
XPrt.FontSize = 8: XPrt.FontBold = True
prtMaxY_PDCMVT = XPrt.CurrentY + prtHeaderHeight
If lFct = " " Then
    prtFillColor = RGB(210, 255, 255)
    Call frmElpPrt.prtTrame_Color(prtMinX, XPrt.CurrentY, prtMaxX, XPrt.CurrentY + prtHeaderHeight, " ")
Else
    prtFillColor = RGB(180, 255, 255)
    Call frmElpPrt.prtTrame_Color(prtMinX, XPrt.CurrentY, prtMaxX, XPrt.CurrentY + prtHeaderHeight, "B")
End If

XPrt.CurrentY = XPrt.CurrentY + 50
XPrt.CurrentX = prtMinX + 50: XPrt.Print dateImp(lYPDCPOS0.PDCPOSDTR) & "   " & lYPDCPOS0.PDCPOSDEV;

Call prtBIA_PDC_MT(prtMinX + 3200 - 100, lYPDCPOS0.PDCPOSPOSE)

Call prtBIA_PDC_MT(prtMinX + 4800 - 100, lYPDCPOS0.PDCPOSPOSD)

X = Format$(lYPDCPOS0.PDCPOSPRIX, "### ##0.000000")
XPrt.CurrentX = prtMinX + 6400 - 100 - XPrt.TextWidth(X): XPrt.Print X;
prtFillColor = prtFillColor_Standard
'---------------------------------------------------------

End Sub
'---------------------------------------------------------
Public Sub prtBIA_PDCMVT_Col()
'---------------------------------------------------------

XPrt.DrawWidth = 1
XPrt.Line (prtMinX, prtMinY_PDCMVT)-(prtMinX, prtMaxY_PDCMVT), prtLineColor
XPrt.Line (prtMinX + 1600, prtMinY_PDCMVT)-(prtMinX + 1600, prtMaxY_PDCMVT), prtLineColor
XPrt.Line (prtMinX + 3200, prtMinY_PDCMVT)-(prtMinX + 3200, prtMaxY_PDCMVT), prtLineColor
XPrt.Line (prtMinX + 4800, prtMinY_PDCMVT)-(prtMinX + 4800, prtMaxY_PDCMVT), prtLineColor
XPrt.Line (prtMinX + 6400, prtMinY_PDCMVT)-(prtMinX + 6400, prtMaxY_PDCMVT), prtLineColor
XPrt.Line (prtMinX + 8000, prtMinY_PDCMVT)-(prtMinX + 8000, prtMaxY_PDCMVT), prtLineColor
'XPrt.Line (prtMinX + 9600, prtMinY_pdcmvt)-(prtMinX + 9600,xprt.currenty), prtLineColor
XPrt.Line (prtMinX + 11200, prtMinY_PDCMVT)-(prtMinX + 11200, prtMaxY_PDCMVT), prtLineColor
XPrt.Line (prtMinX + 12800, prtMinY_PDCMVT)-(prtMinX + 12800, prtMaxY_PDCMVT), prtLineColor
XPrt.Line (prtMinX + 14400, prtMinY_PDCMVT)-(prtMinX + 14400, prtMaxY_PDCMVT), prtLineColor
XPrt.Line (prtMaxX, prtMinY_PDCMVT)-(prtMaxX, prtMaxY_PDCMVT), prtLineColor
End Sub

'---------------------------------------------------------
Public Sub prtBIA_PDCOPE_Col()
'---------------------------------------------------------

XPrt.DrawWidth = 1
XPrt.Line (prtMinX, prtMinY_PDCMVT)-(prtMinX, prtMaxY_PDCMVT), prtLineColor
XPrt.Line (prtMinX + 1600, prtMinY_PDCMVT)-(prtMinX + 1600, prtMaxY_PDCMVT), prtLineColor
XPrt.Line (prtMinX + 3700, prtMinY_PDCMVT)-(prtMinX + 3700, prtMaxY_PDCMVT), prtLineColor
XPrt.Line (prtMinX + 5800, prtMinY_PDCMVT)-(prtMinX + 5800, prtMaxY_PDCMVT), prtLineColor
XPrt.Line (prtMinX + 7400, prtMinY_PDCMVT)-(prtMinX + 7400, prtMaxY_PDCMVT), prtLineColor
XPrt.Line (prtMinX + 9000, prtMinY_PDCMVT)-(prtMinX + 9000, prtMaxY_PDCMVT), prtLineColor
XPrt.Line (prtMinX + 12200, prtMinY_PDCMVT)-(prtMinX + 12200, prtMaxY_PDCMVT), prtLineColor
XPrt.Line (prtMinX + 13800, prtMinY_PDCMVT)-(prtMinX + 13800, prtMaxY_PDCMVT), prtLineColor
XPrt.Line (prtMaxX, prtMinY_PDCMVT)-(prtMaxX, prtMaxY_PDCMVT), prtLineColor
End Sub


'---------------------------------------------------------
Public Sub prtBIA_PDCLOG_Col()
'---------------------------------------------------------

XPrt.DrawWidth = 1
XPrt.Line (prtMinX, prtMinY_PDCLOG)-(prtMinX, prtMaxY_PDCLOG), prtLineColor
XPrt.Line (prtMinX + 1600, prtMinY_PDCLOG)-(prtMinX + 1600, prtMaxY_PDCLOG), prtLineColor
XPrt.Line (prtMinX + 4800, prtMinY_PDCLOG)-(prtMinX + 4800, prtMaxY_PDCLOG), prtLineColor
XPrt.Line (prtMinX + 6400, prtMinY_PDCLOG)-(prtMinX + 6400, prtMaxY_PDCLOG), prtLineColor
XPrt.Line (prtMinX + 12800, prtMinY_PDCLOG)-(prtMinX + 12800, prtMaxY_PDCLOG), prtLineColor
XPrt.Line (prtMinX + 14400, prtMinY_PDCLOG)-(prtMinX + 14400, prtMaxY_PDCLOG), prtLineColor
XPrt.Line (prtMaxX, prtMinY_PDCLOG)-(prtMaxX, prtMaxY_PDCLOG), prtLineColor
End Sub

Public Sub prtBIA_PDC_Open()
On Error GoTo prtError

Set XPrt = Printer
frmElpPrt.Show vbModeless

Height8_6 = frmElpPrt.prtHeightDelta(8, 6)

blnFiligrane = False
prtOrientation = vbPRORLandscape '
prtPgmName = "prtBIA_PDC"
prtTitleUsr = usrName
prtTitleText = "Position de change"
prtFontName = "Arial Unicode MS"
prtLineNb = 1
prtlineHeight = 350
prtHeaderHeight = 300


prtFormType = ""
frmElpPrt.prtStdInit
XPrt.CurrentY = prtMinY
Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub

Public Sub prtBIA_PDC_Etat_Init(lEtat_PDCPOS As String)

mEtat_PDCPOS = lEtat_PDCPOS
blnPDCMVT = False
blnPDCLOG = False
blnPDCPOS = False
blnPDCOPE = False
End Sub

Public Sub prtBIA_PDC_Init(lEtat_PDCPOS As String, lEtat_Exclure_HB As String, lEtat_Exclure_PDCMVTKCUT As String)

mEtat_PDCPOS = lEtat_PDCPOS
mEtat_Exclure_HB = lEtat_Exclure_HB
mEtat_Exclure_PDCMVTKCUT = lEtat_Exclure_PDCMVTKCUT

blnPDCMVT = False
blnPDCLOG = False
blnPDCPOS = False
blnPDCOPE = False

End Sub

Public Sub prtBIA_PDC_Close()
On Error GoTo prtError

If blnPDCMVT Then
    prtBIA_PDCMVT_Col
    XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY), prtLineColor
End If
If blnPDCLOG Then
    prtMaxY_PDCLOG = XPrt.CurrentY + prtlineHeight

    prtBIA_PDCLOG_Col
    XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY), prtLineColor
End If

Call frmElpPrt.prtEndDoc(1000)
frmElpPrt.Hide
Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub

Public Sub prtBIA_PDCMVT_End()
On Error GoTo prtError

If blnPDCMVT Then
    prtMaxY_PDCMVT = XPrt.CurrentY + prtHeaderHeight

    prtBIA_PDCMVT_Col
    XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY), prtLineColor
    blnPDCMVT = False
End If
Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub

Public Sub prtBIA_PDCLOG_End()
On Error GoTo prtError

If blnPDCLOG Then
    prtMaxY_PDCLOG = XPrt.CurrentY + prtHeaderHeight

    prtBIA_PDCLOG_Col
    XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY), prtLineColor
    blnPDCLOG = False
End If
Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub

Public Sub prtBIA_PDCOPE_End()
On Error GoTo prtError

If blnPDCOPE Then
    prtMaxY_PDCMVT = XPrt.CurrentY + prtHeaderHeight

    prtBIA_PDCOPE_Col
    XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY), prtLineColor
    blnPDCOPE = False
End If
Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub


Public Sub prtBIA_PDC_NewLine()
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
If XPrt.CurrentY + 300 > prtMaxY Then
    If blnPDCMVT Then
        prtMaxY_PDCMVT = prtMaxY
        prtBIA_PDCMVT_Col
        frmElpPrt.prtNewPage
        prtBIA_PDCMVT_Form

    Else
         If blnPDCLOG Then
            prtMaxY_PDCLOG = prtMaxY
            prtBIA_PDCLOG_Col
            frmElpPrt.prtNewPage
            prtBIA_PDCLOG_Form
    
        Else
            If blnPDCLOG Then
                prtMaxY_PDCPOS = prtMaxY
                prtBIA_PDCPOS_Col
                frmElpPrt.prtNewPage
                prtBIA_PDCPOS_Form
            End If
        End If
    End If
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    
End If

End Sub



Public Sub prtBIA_PDCPOS_Linexxx(lYPDCPOS0 As typeYPDCPOS0, blnTerme As Boolean)
Dim K As Long, X As String
Dim devFixing As Currency, devPP As Currency, eurPP As Currency
Dim wPDCPOSPOSE As Currency, wPDCPOSPOSD As Currency
Dim wCours_format As String

prtBIA_PDC_NewLine

If Not blnTerme Then
    wPDCPOSPOSE = -lYPDCPOS0.PDCPOSPOSE
    wPDCPOSPOSD = -lYPDCPOS0.PDCPOSPOSD
Else
    wPDCPOSPOSE = -lYPDCPOS0.PDCPOSPOSE - lYPDCPOS0.PDCPOSTERE - lYPDCPOS0.PDCPOSSWPE
    wPDCPOSPOSD = -lYPDCPOS0.PDCPOSPOSD - lYPDCPOS0.PDCPOSTERD - lYPDCPOS0.PDCPOSSWPD
End If

prtFillColor = RGB(180, 255, 255)

Call frmElpPrt.prtTrame_Color(prtMinX, XPrt.CurrentY, prtMinX + 1600, XPrt.CurrentY + prtlineHeight, " ")

XPrt.CurrentX = prtMinX + 50: XPrt.Print "EUR / " & lYPDCPOS0.PDCPOSDEV;
XPrt.CurrentY = XPrt.CurrentY + Height8_6
XPrt.FontSize = 6
XPrt.CurrentX = prtMinX + 1100: XPrt.Print Format$(Mid$(lYPDCPOS0.PDCPOSDTR, 7, 2) & Mid$(lYPDCPOS0.PDCPOSDTR, 5, 2) & Mid$(lYPDCPOS0.PDCPOSDTR, 3, 2), "@@.@@.@@");
XPrt.CurrentY = XPrt.CurrentY - Height8_6
XPrt.FontSize = 8

Call prtBIA_PDC_MT(prtMinX + 3200 - 100, wPDCPOSPOSE)

Call prtBIA_PDC_MT(prtMinX + 4800 - 100, wPDCPOSPOSD)

If lYPDCPOS0.PDCPOSDEV = "JPY" Then
    wCours_format = "#### ##0.00"
Else
    wCours_format = "### ##0.000000"
End If

X = Format$(lYPDCPOS0.PDCPOSPRIX, wCours_format)
XPrt.CurrentX = prtMinX + 6400 - 100 - XPrt.TextWidth(X): XPrt.Print X;

X = Format$(lYPDCPOS0.PDCPOSFIXT, wCours_format)
XPrt.CurrentX = prtMinX + 8000 - 100 - XPrt.TextWidth(X): XPrt.Print X;

devFixing = -Round(wPDCPOSPOSE * lYPDCPOS0.PDCPOSFIXT, 2)
Call prtBIA_PDC_MT(prtMinX + 9600 - 100, devFixing)

devPP = wPDCPOSPOSD - devFixing
Call prtBIA_PDC_MT(prtMinX + 11200 - 100, devPP)

eurPP = Round(devPP / lYPDCPOS0.PDCPOSFIXT, 2)
Call prtBIA_PDC_MT(prtMinX + 12800 - 100, eurPP)

Call prtBIA_PDC_MT(prtMinX + 14400 - 100, lYPDCPOS0.PDCPOSPNL)

Call prtBIA_PDC_MT(prtMaxX - 100, lYPDCPOS0.PDCPOSRPC)

End Sub



Public Sub prtBIA_PDCMVT_Line(lYPDCMVT0 As typeYPDCMVT0, lKCUT As String)
Dim K As Long, X As String
Dim devFixing As Currency, devPP As Currency, eurPP As Currency
Dim wPDCMVTMTE As Currency, wPDCMVTMTD As Currency


prtBIA_PDC_NewLine
XPrt.FontSize = 8: XPrt.FontBold = False

If lYPDCMVT0.PDCMVTOPEC = "RPC" Then
    prtFillColor = RGB(255, 200, 268)
    Call frmElpPrt.prtTrame_Color(prtMinX, XPrt.CurrentY - 50, prtMaxX, XPrt.CurrentY + prtlineHeight, " ")
End If
wPDCMVTMTE = -lYPDCMVT0.PDCMVTMTE
wPDCMVTMTD = -lYPDCMVT0.PDCMVTMTD

XPrt.CurrentX = prtMinX + 50: XPrt.Print dateImp(lYPDCMVT0.PDCMVTDTR) & "   " & lYPDCMVT0.PDCMVTDEV;

Call prtBIA_PDC_MT(prtMinX + 3200 - 100, wPDCMVTMTE)

Call prtBIA_PDC_MT(prtMinX + 4800 - 100, wPDCMVTMTD)

X = Format$(lYPDCMVT0.PDCMVTTAUX, "### ##0.000000")
XPrt.CurrentX = prtMinX + 6400 - 100 - XPrt.TextWidth(X): XPrt.Print X;

X = dateImp(lYPDCMVT0.PDCMVTDVA)
XPrt.CurrentX = prtMinX + 8000 - 100 - XPrt.TextWidth(X): XPrt.Print X;

XPrt.CurrentX = prtMinX + 8000 + 100: XPrt.Print lYPDCMVT0.PDCMVTSTA2;
XPrt.CurrentX = prtMinX + 8000 + 300: XPrt.Print lYPDCMVT0.PDCMVTSER;
XPrt.CurrentX = prtMinX + 8000 + 600: XPrt.Print lYPDCMVT0.PDCMVTSSE;
XPrt.CurrentX = prtMinX + 8000 + 1000: XPrt.Print lYPDCMVT0.PDCMVTOPEC;
X = Format$(lYPDCMVT0.PDCMVTOPEN, "### ##0")
XPrt.CurrentX = prtMinX + 11000 - 100 - XPrt.TextWidth(X): XPrt.Print X;

XPrt.CurrentX = prtMinX + 11200 + 100
If lYPDCMVT0.PDCMVTKCUT = " " Then
    XPrt.Print lYPDCMVT0.PDCMVTCLI;
Else
    XPrt.ForeColor = vbMagenta
    If lKCUT = "" Then
        XPrt.Print lYPDCMVT0.PDCMVTCLI & " cut";
    Else
        XPrt.Print lKCUT;
    End If
    XPrt.ForeColor = vbBlack
End If
XPrt.CurrentX = prtMinX + 12800 + 100: XPrt.Print lYPDCMVT0.PDCMVTSTA & " " & lYPDCMVT0.PDCMVTPIE & " - " & lYPDCMVT0.PDCMVTECR;
XPrt.CurrentX = prtMinX + 14400 + 100: XPrt.Print lYPDCMVT0.PDCMVTCPT;
prtFillColor = prtFillColor_Standard

End Sub
Public Sub prtBIA_PDCTER_Line(fgX As MSFlexGrid)
Dim K As Long, X As String
Dim devFixing As Currency, devPP As Currency, eurPP As Currency
Dim wPDCTERMTE As Currency, wPDCTERMTD As Currency

prtBIA_PDC_Etat_Init "Echéancier Terme"
prtBIA_PDC_Etat
prtBIA_PDCTER_Form

XPrt.FontSize = 8: XPrt.FontBold = False
For K = 1 To fgX.Rows - 1
    fgX.Row = K
    prtBIA_PDC_NewLine
    fgX.Col = 6: prtFillColor = fgX.CellBackColor

    Call frmElpPrt.prtTrame_Color(prtMinX, XPrt.CurrentY, prtMaxX, XPrt.CurrentY + prtHeaderHeight, " ")

    fgX.Col = 0: XPrt.ForeColor = fgX.CellForeColor
    XPrt.CurrentX = prtMinX + 100: XPrt.Print fgX.Text;
    fgX.Col = 1: XPrt.ForeColor = fgX.CellForeColor
    XPrt.CurrentX = prtMinX + 1500 + 100: XPrt.Print fgX.Text;
    fgX.Col = 2: XPrt.ForeColor = fgX.CellForeColor
    XPrt.CurrentX = prtMinX + 2500 + 100: XPrt.Print fgX.Text;
    fgX.Col = 3: XPrt.ForeColor = fgX.CellForeColor
    XPrt.CurrentX = prtMinX + 4000 + 100: XPrt.Print fgX.Text;
    fgX.Col = 4:    X = Trim(fgX.Text): XPrt.ForeColor = fgX.CellForeColor
    XPrt.CurrentX = prtMinX + 6500 - 100 - XPrt.TextWidth(X): XPrt.Print X;
    
    fgX.Col = 6: X = Trim(fgX.Text): XPrt.ForeColor = fgX.CellForeColor
    XPrt.CurrentX = prtMinX + 8000 - 100 - XPrt.TextWidth(X): XPrt.Print X;
    fgX.Col = 7: X = Trim(fgX.Text): XPrt.ForeColor = fgX.CellForeColor
    XPrt.CurrentX = prtMinX + 10000 - 100 - XPrt.TextWidth(X): XPrt.Print X;
    
    fgX.Col = 8: XPrt.ForeColor = fgX.CellForeColor
    XPrt.CurrentX = prtMinX + 10300 - 100: XPrt.Print fgX.Text;
    
    fgX.Col = 9: X = Trim(fgX.Text): XPrt.ForeColor = fgX.CellForeColor
    XPrt.CurrentX = prtMinX + 12200 - 100 - XPrt.TextWidth(X): XPrt.Print X;
    fgX.Col = 10: X = Trim(fgX.Text): XPrt.ForeColor = fgX.CellForeColor
    XPrt.CurrentX = prtMinX + 14000 - 100 - XPrt.TextWidth(X): XPrt.Print X;
    fgX.Col = 11: X = Trim(fgX.Text): XPrt.ForeColor = fgX.CellForeColor
    XPrt.CurrentX = prtMaxX - 100 - XPrt.TextWidth(X): XPrt.Print X;

Next K
prtFillColor = prtFillColor_Standard

End Sub
Public Sub prtBIA_PDCPOS_Line(fgX As MSFlexGrid)
Dim K As Long, X As String
Dim devFixing As Currency, devPP As Currency, eurPP As Currency
Dim wPDCTERMTE As Currency, wPDCTERMTD As Currency


XPrt.FontSize = 8: XPrt.FontBold = False
For K = 1 To fgX.Rows - 1
    fgX.Row = K
    prtBIA_PDC_NewLine
    fgX.Col = 1: prtFillColor = fgX.CellBackColor
    Call frmElpPrt.prtTrame_Color(prtMinX, XPrt.CurrentY, prtMaxX, XPrt.CurrentY + prtHeaderHeight, " ")

    fgX.Col = 0: XPrt.ForeColor = fgX.CellForeColor
    XPrt.CurrentX = prtMinX + 100: XPrt.Print fgX.Text;
    fgX.Col = 1: X = Trim(fgX.Text): XPrt.ForeColor = fgX.CellForeColor
    XPrt.CurrentX = prtMinX + 3000 - 100 - XPrt.TextWidth(X): XPrt.Print X;
    fgX.Col = 2: X = Trim(fgX.Text): XPrt.ForeColor = fgX.CellForeColor
    XPrt.CurrentX = prtMinX + 4400 - 100 - XPrt.TextWidth(X): XPrt.Print X;
    fgX.Col = 3: X = Trim(fgX.Text): XPrt.ForeColor = fgX.CellForeColor
    XPrt.CurrentX = prtMinX + 5800 - 100 - XPrt.TextWidth(X): XPrt.Print X;
    fgX.Col = 4: X = Trim(fgX.Text): XPrt.ForeColor = fgX.CellForeColor
    XPrt.CurrentX = prtMinX + 7200 - 100 - XPrt.TextWidth(X): XPrt.Print X;
    fgX.Col = 5: X = Trim(fgX.Text): XPrt.ForeColor = fgX.CellForeColor
    XPrt.CurrentX = prtMinX + 8600 - 100 - XPrt.TextWidth(X): XPrt.Print X;
    
    fgX.Col = 6: X = Trim(fgX.Text): XPrt.ForeColor = fgX.CellForeColor
    XPrt.CurrentX = prtMinX + 10000 - 100 - XPrt.TextWidth(X): XPrt.Print X;
    fgX.Col = 7: X = Trim(fgX.Text): XPrt.ForeColor = fgX.CellForeColor
    XPrt.CurrentX = prtMinX + 11400 - 100 - XPrt.TextWidth(X): XPrt.Print X;
    
    fgX.Col = 8: X = Trim(fgX.Text): XPrt.ForeColor = fgX.CellForeColor
    XPrt.CurrentX = prtMinX + 12800 - 100 - XPrt.TextWidth(X): XPrt.Print X;
    fgX.Col = 9: X = Trim(fgX.Text): XPrt.ForeColor = fgX.CellForeColor
    XPrt.CurrentX = prtMinX + 14400 - 100 - XPrt.TextWidth(X): XPrt.Print X;
    fgX.Col = 10: X = Trim(fgX.Text): XPrt.ForeColor = fgX.CellForeColor
    XPrt.CurrentX = prtMaxX - 100 - XPrt.TextWidth(X): XPrt.Print X;

Next K
prtFillColor = prtFillColor_Standard

End Sub
Public Sub prtBIA_PDCOPE_ZCHGOPE0(lZCHGOPE0 As typeZCHGOPE0, lTxt As String)
Dim K As Long, X As String
Dim wPDCMVTMTE As Currency, wPDCMVTMTD As Currency


prtBIA_PDC_NewLine
XPrt.FontSize = 8: XPrt.FontBold = False
If lZCHGOPE0.CHGOPESEN = "A" Then
    wPDCMVTMTE = lZCHGOPE0.CHGOPEMO1
    wPDCMVTMTD = -lZCHGOPE0.CHGOPEMO2
Else
    wPDCMVTMTE = -lZCHGOPE0.CHGOPEMO1
    wPDCMVTMTD = lZCHGOPE0.CHGOPEMO2
End If

XPrt.CurrentX = prtMinX + 50: XPrt.Print dateImp(lZCHGOPE0.CHGOPECRE + 19000000) & "   " & lZCHGOPE0.CHGOPESEN;

Call prtBIA_PDC_MT(prtMinX + 3300 - 100, wPDCMVTMTE)
XPrt.CurrentX = prtMinX + 3300: XPrt.Print lZCHGOPE0.CHGOPEDE1;

Call prtBIA_PDC_MT(prtMinX + 5400 - 100, wPDCMVTMTD)
XPrt.CurrentX = prtMinX + 5400: XPrt.Print lZCHGOPE0.CHGOPEDE2;

If lZCHGOPE0.CHGOPECO3 <> 0 Then
    X = Format$(lZCHGOPE0.CHGOPECO3, "### ##0.000000")
Else
    X = Format$(lZCHGOPE0.CHGOPECO1, "### ##0.000000")
End If

XPrt.CurrentX = prtMinX + 7400 - 100 - XPrt.TextWidth(X): XPrt.Print X;

X = dateImp(lZCHGOPE0.CHGOPEDT1 + 19000000)
XPrt.CurrentX = prtMinX + 9000 - 100 - XPrt.TextWidth(X): XPrt.Print X;

XPrt.CurrentX = prtMinX + 9000 + 300: XPrt.Print lZCHGOPE0.CHGOPESER;
XPrt.CurrentX = prtMinX + 9000 + 600: XPrt.Print lZCHGOPE0.CHGOPESSE;
XPrt.CurrentX = prtMinX + 9000 + 1000: XPrt.Print lZCHGOPE0.CHGOPEOPE;
X = Format$(lZCHGOPE0.CHGOPEDOS, "### ##0")
XPrt.CurrentX = prtMinX + 12000 - 100 - XPrt.TextWidth(X): XPrt.Print X;

XPrt.CurrentX = prtMinX + 12200 + 100
XPrt.Print lZCHGOPE0.CHGOPECON;

XPrt.CurrentX = prtMinX + 13800 + 100: XPrt.Print lTxt;

prtFillColor = prtFillColor_Standard

End Sub
Public Sub prtBIA_PDCOPE_YPDCOPE0(lYPDCOPE0 As typeYPDCOPE0)
Dim K As Long, X As String
Dim wPDCMVTMTE As Currency, wPDCMVTMTD As Currency


prtBIA_PDC_NewLine
XPrt.FontSize = 8: XPrt.FontBold = False
If lYPDCOPE0.PDCOPESENS = "A" Then
    wPDCMVTMTE = -lYPDCOPE0.PDCOPEMTD1
    wPDCMVTMTD = -lYPDCOPE0.PDCOPEMTD2
Else
    wPDCMVTMTE = -lYPDCOPE0.PDCOPEMTD1
    wPDCMVTMTD = -lYPDCOPE0.PDCOPEMTD2
End If

XPrt.CurrentX = prtMinX + 50: XPrt.Print dateImp(lYPDCOPE0.PDCOPEIAMJ) & "   " & lYPDCOPE0.PDCOPESENS;

Call prtBIA_PDC_MT(prtMinX + 3300 - 100, wPDCMVTMTE)
XPrt.CurrentX = prtMinX + 3300: XPrt.Print lYPDCOPE0.PDCOPEDEV1;

Call prtBIA_PDC_MT(prtMinX + 5400 - 100, wPDCMVTMTD)
XPrt.CurrentX = prtMinX + 5400: XPrt.Print lYPDCOPE0.PDCOPEDEV2;

X = Format$(lYPDCOPE0.PDCOPETAUX, "### ##0.000000")
XPrt.CurrentX = prtMinX + 7400 - 100 - XPrt.TextWidth(X): XPrt.Print X;

X = dateImp(lYPDCOPE0.PDCOPEDVA + 19000000)
XPrt.CurrentX = prtMinX + 9000 - 100 - XPrt.TextWidth(X): XPrt.Print X;

XPrt.CurrentX = prtMinX + 9000 + 300: XPrt.Print lYPDCOPE0.PDCOPESER;
XPrt.CurrentX = prtMinX + 9000 + 600: XPrt.Print lYPDCOPE0.PDCOPESSE;
XPrt.CurrentX = prtMinX + 9000 + 1000: XPrt.Print lYPDCOPE0.PDCOPEOPEC & "-" & lYPDCOPE0.PDCOPEID;
X = Format$(lYPDCOPE0.PDCOPEOPEN, "### ##0")
XPrt.CurrentX = prtMinX + 12000 - 100 - XPrt.TextWidth(X): XPrt.Print X;

XPrt.CurrentX = prtMinX + 12200 + 100
XPrt.Print lYPDCOPE0.PDCOPECLI;

XPrt.CurrentX = prtMinX + 13800 + 100: XPrt.Print lYPDCOPE0.PDCOPESTA & lYPDCOPE0.PDCOPESTA2 & lYPDCOPE0.PDCOPESTA3;
If lYPDCOPE0.PDCOPEREF <> 0 Then XPrt.Print " " & lYPDCOPE0.PDCOPEREF;
If lYPDCOPE0.PDCOPEIAMJ <> lYPDCOPE0.PDCOPEDTR Then XPrt.Print " report";
prtFillColor = prtFillColor_Standard

End Sub

Public Sub prtBIA_PDCLOG_Line(lYPDCLOG0 As typeYPDCLOG0)
Dim K As Long, X As String


prtBIA_PDC_NewLine
XPrt.ForeColor = vbBlack
If Mid$(lYPDCLOG0.PDCLOGNAT, 3, 1) <> " " Then XPrt.ForeColor = vbRed
If Mid$(lYPDCLOG0.PDCLOGNAT, 1, 1) = "7" Then XPrt.ForeColor = vbMagenta
If Mid$(lYPDCLOG0.PDCLOGNAT, 1, 2) = "5=" Then XPrt.ForeColor = vbMagenta

XPrt.CurrentX = prtMinX + 50: XPrt.Print dateImp(lYPDCLOG0.PDCLOGDTR);

XPrt.CurrentX = prtMinX + 1600 + 100: XPrt.Print dateImp(lYPDCLOG0.PDCLOGUAMJ) & "  " & timeImp(lYPDCLOG0.PDCLOGUHMS) & " - "; lYPDCLOG0.PDCLOGUSEQ;
XPrt.CurrentX = prtMinX + 4800 + 100: XPrt.Print lYPDCLOG0.PDCLOGSTA & " " & lYPDCLOG0.PDCLOGNAT;
XPrt.CurrentX = prtMinX + 6400 + 100: XPrt.Print lYPDCLOG0.PDCLOGTXT;
If lYPDCLOG0.PDCLOGPIE <> 0 Then XPrt.CurrentX = prtMinX + 12800 + 100: XPrt.Print lYPDCLOG0.PDCLOGPIE & " - " & lYPDCLOG0.PDCLOGECR;
XPrt.CurrentX = prtMinX + 14400 + 100: XPrt.Print lYPDCLOG0.PDCLOGUUSR;
'prtFillColor = prtFillColor_Standard

End Sub

Public Sub prtBIA_PDC_MT(lCurrentY As Integer, lcurX As Currency)
Dim X As String
If lcurX <> 0 Then
    If lcurX < 0 Then
        XPrt.ForeColor = vbRed
    Else
        XPrt.ForeColor = vbBlue
    End If
    X = Format$(lcurX, "### ### ### ##0.00")
    XPrt.CurrentX = lCurrentY - XPrt.TextWidth(X): XPrt.Print X;
    XPrt.ForeColor = prtForeColor
End If
End Sub


Public Sub prtBIA_PDC_Etat()

prtlineHeight = 300

If XPrt.CurrentY > prtMaxY - prtlineHeight * 10 Then frmElpPrt.prtNewPage

XPrt.FontSize = 14: XPrt.FontBold = True
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.FontUnderline = True
XPrt.ForeColor = RGB(0, 123, 141)

XPrt.CurrentX = prtMinX + 8000 - XPrt.TextWidth(mEtat_PDCPOS) / 2
XPrt.Print mEtat_PDCPOS;
XPrt.FontUnderline = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.FontSize = 8: XPrt.FontBold = True

End Sub

Public Sub prtBIA_PDCPOS_Col()

XPrt.DrawWidth = 1
XPrt.Line (prtMinX, prtMinY_PDCPOS)-(prtMinX, prtMaxY_PDCPOS), prtLineColor
XPrt.Line (prtMinX + 1600, prtMinY_PDCPOS)-(prtMinX + 1600, prtMaxY_PDCPOS), prtLineColor
XPrt.Line (prtMinX + 3200, prtMinY_PDCPOS)-(prtMinX + 3200, prtMaxY_PDCPOS), prtLineColor
XPrt.Line (prtMinX + 4800, prtMinY_PDCPOS)-(prtMinX + 4800, prtMaxY_PDCPOS), prtLineColor
XPrt.Line (prtMinX + 6400, prtMinY_PDCPOS)-(prtMinX + 6400, prtMaxY_PDCPOS), prtLineColor
XPrt.Line (prtMinX + 8000, prtMinY_PDCPOS)-(prtMinX + 8000, prtMaxY_PDCPOS), prtLineColor
XPrt.Line (prtMinX + 9600, prtMinY_PDCPOS)-(prtMinX + 9600, prtMaxY_PDCPOS), prtLineColor
XPrt.Line (prtMinX + 11200, prtMinY_PDCPOS)-(prtMinX + 11200, prtMaxY_PDCPOS), prtLineColor
XPrt.Line (prtMinX + 12800, prtMinY_PDCPOS)-(prtMinX + 12800, prtMaxY_PDCPOS), prtLineColor
XPrt.Line (prtMinX + 14400, prtMinY_PDCPOS)-(prtMinX + 14400, prtMaxY_PDCPOS), prtLineColor
XPrt.Line (prtMaxX, prtMinY_PDCPOS)-(prtMaxX, prtMaxY_PDCPOS), prtLineColor

End Sub
Public Sub prtBIA_PDCTER_Line_xlsManual(fgX As MSFlexGrid, ByRef currentRow As Long, wsExcel As Excel.Worksheet)
Dim K As Long
Dim X As String
Dim wColor As Long
Dim comptageRows As Long
Dim maxRows As Long
Dim maxRowsPlus As Long

prtBIA_PDC_Etat_Init "Echéancier Terme"
wsExcel.Cells(currentRow, 5) = "Echéancier Terme"
currentRow = 6
comptageRows = currentRow
maxRows = 45
maxRowsPlus = 3
For K = 1 To fgX.Rows - 1
    fgX.Row = K
    Call prtBIA_PDC_NewLine_xlsManual(currentRow, wsExcel, "K", comptageRows, maxRows, maxRowsPlus)
    fgX.Col = 0: wColor = fgX.CellForeColor
    wsExcel.Cells(currentRow, 1) = fgX.Text
    wsExcel.Cells(currentRow, 1).Font.Color = wColor
    fgX.Col = 1: wColor = fgX.CellForeColor
    wsExcel.Cells(currentRow, 2) = fgX.Text
    wsExcel.Cells(currentRow, 2).Font.Color = wColor
    fgX.Col = 2: wColor = fgX.CellForeColor
    wsExcel.Cells(currentRow, 3) = fgX.Text
    wsExcel.Cells(currentRow, 3).Font.Color = wColor
    fgX.Col = 3: wColor = fgX.CellForeColor
    wsExcel.Cells(currentRow, 4) = fgX.Text
    wsExcel.Cells(currentRow, 4).Font.Color = wColor
    fgX.Col = 4:    X = Trim(fgX.Text): wColor = fgX.CellForeColor
    wsExcel.Cells(currentRow, 5) = fgX.Text
    wsExcel.Cells(currentRow, 5).Font.Color = wColor
    fgX.Col = 6: X = Trim(fgX.Text): wColor = fgX.CellForeColor
    wsExcel.Cells(currentRow, 7) = X
    wsExcel.Cells(currentRow, 7).Font.Color = wColor
    fgX.Col = 7: X = Trim(fgX.Text): wColor = fgX.CellForeColor
    wsExcel.Cells(currentRow, 8) = X
    wsExcel.Cells(currentRow, 8).Font.Color = wColor
    fgX.Col = 8: wColor = fgX.CellForeColor
    wsExcel.Cells(currentRow, 9) = fgX.Text
    wsExcel.Cells(currentRow, 9).Font.Color = wColor
    fgX.Col = 9: X = Trim(fgX.Text): wColor = fgX.CellForeColor
    wsExcel.Cells(currentRow, 10) = X
    wsExcel.Cells(currentRow, 10).Font.Color = wColor
    fgX.Col = 10: X = Trim(fgX.Text): wColor = fgX.CellForeColor
    wsExcel.Cells(currentRow, 11) = X
    wsExcel.Cells(currentRow, 11).Font.Color = wColor
Next K
If currentRow >= maxRows + maxRowsPlus Then
    If comptageRows >= maxRows Then
        Call insere_entete_page(wsExcel, "1:3", 3, currentRow)
        comptageRows = 3
        currentRow = currentRow + 3
    End If
End If
comptageRows = comptageRows + 1
currentRow = currentRow + 1
Range("A4:K4").Select
Selection.Copy
Range("A" & CStr(currentRow)).Select
ActiveSheet.Paste
If currentRow >= maxRows + maxRowsPlus Then
    If comptageRows >= maxRows Then
        Call insere_entete_page(wsExcel, "1:3", 3, currentRow)
        comptageRows = 3
        currentRow = currentRow + 3
    End If
End If
comptageRows = comptageRows + 1
currentRow = currentRow + 1
Range("A6:K6").Select
Selection.Copy
Range("A" & CStr(currentRow)).Select
ActiveSheet.Paste

End Sub


