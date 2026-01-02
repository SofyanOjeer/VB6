Attribute VB_Name = "prtSAB_Stock"
Option Explicit
Dim mFct1 As String

Dim X As String, I As Integer, Height8_6 As Integer
Dim curX As Currency, curX1 As Currency, curX2 As Currency

Dim blnPage As Boolean

Dim meYBIACPT0 As typeYBIACPT0, prevYBIACPT0 As typeYBIACPT0
Dim meYBIASTO0 As typeYBIASTO0, xYBIASTO0 As typeYBIASTO0

Dim cnAdo As New ADODB.Connection
Dim rsAdo As New ADODB.Recordset




Private Sub prtSAB_Stock_Détail_xlsManual(wsExcel As Excel.Worksheet, ByRef currentRow As Long)
Dim Nb As Long
Dim xSQL As String
Dim V
Set rsAdo = Nothing
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIASTO0 where " _
     & "YSTOPCI = '" & meYBIASTO0.YSTOPCI & "'" _
     & "AND YSTODEV = '" & meYBIASTO0.YSTODEV & "'" _
     & "AND YSTOCLI = " & meYBIASTO0.YSTOCLI

Set rsAdo = cnAdo.Execute(xSQL)
Nb = 0
Do While Not rsAdo.EOF
    V = rsYBIASTO0_GetBuffer(rsAdo, xYBIASTO0)
    If Not IsNull(V) Then
        MsgBox V, vbCritical, "prtSAB_Stock_SQL"
        Exit Sub
    Else
        Nb = Nb + 1
        If Nb > 1 Then
            currentRow = currentRow + 1
        End If
        Call prtSAB_Stock_Line_Détail_xlsManual(xYBIASTO0, wsExcel, currentRow)
    End If
    rsAdo.MoveNext
Loop

End Sub

Public Sub prtSAB_Stock_Form_xlsManual(mFct1 As String, wsExcel As Excel.Worksheet)
Dim wId As String
Dim X As String
Dim nbCol As Long

wsExcel.Cells(5, 1) = "Client / Compte"
wsExcel.Cells(5, 2) = "Intitulé"
wsExcel.Cells(5, 3) = "Solde comptable"
wsExcel.Cells(5, 4) = "Devise"
wsExcel.Cells(5, 5) = "Ecart"
wsExcel.Cells(5, 6) = "Encours"
nbCol = 6
If Trim(mFct1) = "D" Then
    nbCol = 9
    wsExcel.Cells(5, 7) = "Contrat"
    wsExcel.Cells(5, 8) = "du"
    wsExcel.Cells(5, 9) = "au"
End If
For I = 1 To nbCol
    wsExcel.Cells(5, I).Font.Name = "Arial"
    wsExcel.Cells(5, I).Font.Size = 8
    wsExcel.Cells(5, I).Font.Bold = True
    wsExcel.Cells(5, I).Font.Color = xlsBlue
    wsExcel.Cells(5, I).Interior.Color = xlsBackEntete
    wsExcel.Cells(5, I).Borders(xlEdgeLeft).Weight = xlThin
    wsExcel.Cells(5, I).Borders(xlEdgeLeft).Color = xlsBlue
    wsExcel.Cells(5, I).Borders(xlEdgeRight).Weight = xlThin
    wsExcel.Cells(5, I).Borders(xlEdgeRight).Color = xlsBlue
    wsExcel.Cells(5, I).Borders(xlEdgeTop).Weight = xlThin
    wsExcel.Cells(5, I).Borders(xlEdgeTop).Color = xlsBlue
    wsExcel.Cells(5, I).Borders(xlEdgeBottom).Weight = xlThin
    wsExcel.Cells(5, I).Borders(xlEdgeBottom).Color = xlsBlue
    wsExcel.Cells(5, I).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
Next I
wsExcel.Rows(5).RowHeight = 17
wsExcel.Rows(5).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter

wsExcel.Columns(1).ColumnWidth = 15
wsExcel.Columns(2).ColumnWidth = 28
wsExcel.Columns(3).ColumnWidth = 15
wsExcel.Columns(4).ColumnWidth = 6
wsExcel.Columns(5).ColumnWidth = 15
wsExcel.Columns(6).ColumnWidth = 15
If Trim(mFct1) = "D" Then
    wsExcel.Columns(7).ColumnWidth = 20
    wsExcel.Columns(8).ColumnWidth = 12
    wsExcel.Columns(9).ColumnWidth = 12
End If

End Sub

Public Sub prtSAB_Stock_Line_Détail_xlsManual(lYBIASTO0 As typeYBIASTO0, wsExcel As Excel.Worksheet, ByRef currentRow As Long)
Dim X As String
Dim xlsForeColor As Long

xlsForeColor = xlsBlue
xlsFontSize = 6
wsExcel.Rows(currentRow).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
'                                   '
xlsCol = 6
If lYBIASTO0.YSTOMON < 0 Then
    xlsForeColor = xlsRed
    X = Format$(lYBIASTO0.YSTOMON, "### ### ### ### ##0.00")
Else
    xlsForeColor = xlsBlue
    X = Format$(Abs(lYBIASTO0.YSTOMON), "### ### ### ### ##0.00")
End If
wsExcel.Cells(currentRow, xlsCol) = X
wsExcel.Cells(currentRow, xlsCol).Font.Size = xlsFontSize
wsExcel.Cells(currentRow, xlsCol).Font.Color = xlsForeColor
wsExcel.Cells(currentRow, xlsCol).Font.Italic = True
'                                   '
xlsCol = 7
X = Trim(lYBIASTO0.YSTOAPP) & " " & Trim(lYBIASTO0.YSTOOPE) & " " & lYBIASTO0.YSTONAT & "     " & Format$(lYBIASTO0.YSTONUM, "### ### ### ###")
wsExcel.Cells(currentRow, xlsCol) = X
wsExcel.Cells(currentRow, xlsCol).Font.Size = xlsFontSize
wsExcel.Cells(currentRow, xlsCol).Font.Color = xlsForeColor
wsExcel.Cells(currentRow, xlsCol).Font.Italic = True
'                                   '
xlsCol = 8
wsExcel.Cells(currentRow, xlsCol) = "'" & dateImp10(lYBIASTO0.YSTODEB)
wsExcel.Cells(currentRow, xlsCol).Font.Size = xlsFontSize
wsExcel.Cells(currentRow, xlsCol).Font.Color = xlsForeColor
wsExcel.Cells(currentRow, xlsCol).Font.Italic = True
'                                   '
xlsCol = 9
wsExcel.Cells(currentRow, xlsCol) = "'" & dateImp10(lYBIASTO0.YSTOFIN)
wsExcel.Cells(currentRow, xlsCol).Font.Size = xlsFontSize
wsExcel.Cells(currentRow, xlsCol).Font.Color = xlsForeColor
wsExcel.Cells(currentRow, xlsCol).Font.Italic = True

End Sub

Private Sub prtSAB_Stock_Line_xlsManual(lFct As String, wsExcel As Excel.Worksheet, ByRef currentRow As Long)
Dim I As Long
Dim X As String
Dim rng As Excel.Range

xlsFontSize = 8

currentRow = currentRow + 1
If Trim(lFct) <> "D" Then
    If currentRow Mod 2 = 0 Then
        xlsBackColor = xlsGray
    Else
        xlsBackColor = xlsWhite
    End If
Else
    xlsBackColor = xlsGray
End If
xlsForeColor = xlsBlue
curX1 = Abs(meYBIACPT0.SOLDECEN)
curX2 = Abs(meYBIASTO0.YSTOMON)
curX = Abs(curX1 - curX2)
'                                   '
xlsCol = 1
wsExcel.Cells(currentRow, xlsCol) = meYBIASTO0.YSTOCLI
wsExcel.Cells(currentRow, xlsCol).Font.Size = xlsFontSize
wsExcel.Cells(currentRow, xlsCol).Font.Color = xlsForeColor
wsExcel.Cells(currentRow, xlsCol).Interior.Color = xlsBackColor
wsExcel.Cells(currentRow, xlsCol).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
'                                   '
xlsCol = 2
wsExcel.Cells(currentRow, xlsCol) = meYBIACPT0.CLIENARA1
wsExcel.Cells(currentRow, xlsCol).Font.Size = xlsFontSize
wsExcel.Cells(currentRow, xlsCol).Font.Color = xlsForeColor
wsExcel.Cells(currentRow, xlsCol).Interior.Color = xlsBackColor
wsExcel.Cells(currentRow, xlsCol).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
'                                   '
xlsCol = 3
wsExcel.Cells(currentRow, xlsCol).Font.Size = xlsFontSize
wsExcel.Cells(currentRow, xlsCol).Interior.Color = xlsBackColor
If curX <> 0 Then
    wsExcel.Cells(currentRow, xlsCol) = Format$(curX, "### ### ### ###.00")
    wsExcel.Cells(currentRow, xlsCol).Font.Color = xlsRed
Else
    wsExcel.Cells(currentRow, xlsCol) = ""
End If
wsExcel.Cells(currentRow, xlsCol).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
'                                   '
xlsCol = 4
wsExcel.Cells(currentRow, xlsCol) = meYBIASTO0.YSTODEV
wsExcel.Cells(currentRow, xlsCol).Font.Size = xlsFontSize
wsExcel.Cells(currentRow, xlsCol).Font.Color = xlsForeColor
wsExcel.Cells(currentRow, xlsCol).Interior.Color = xlsBackColor
wsExcel.Cells(currentRow, xlsCol).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
'                                   '
xlsCol = 5
wsExcel.Cells(currentRow, xlsCol).Font.Size = xlsFontSize
wsExcel.Cells(currentRow, xlsCol).Interior.Color = xlsBackColor
If curX2 <> 0 Then
    wsExcel.Cells(currentRow, xlsCol) = Format$(curX2, "### ### ### ###.00")
    wsExcel.Cells(currentRow, xlsCol).Font.Color = xlsRed
Else
    wsExcel.Cells(currentRow, xlsCol) = ""
End If
wsExcel.Cells(currentRow, xlsCol).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
'                                   '
xlsCol = 6
wsExcel.Cells(currentRow, xlsCol).Font.Size = xlsFontSize
wsExcel.Cells(currentRow, xlsCol).Interior.Color = xlsBackColor
If curX1 <> 0 Then
    wsExcel.Cells(currentRow, xlsCol) = Format$(curX2, "### ### ### ###.00")
    wsExcel.Cells(currentRow, xlsCol).Font.Color = xlsRed
Else
    wsExcel.Cells(currentRow, xlsCol) = ""
End If
wsExcel.Cells(currentRow, xlsCol).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
wsExcel.Rows(currentRow).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
If Trim(lFct) = "D" Then
    For I = 7 To 9
        xlsCol = I
        wsExcel.Cells(currentRow, xlsCol) = ""
        wsExcel.Cells(currentRow, xlsCol).Font.Size = xlsFontSize
        wsExcel.Cells(currentRow, xlsCol).Font.Color = xlsForeColor
        wsExcel.Cells(currentRow, xlsCol).Interior.Color = xlsBackColor
    Next I
End If
'////////////////////////////////////////////////////////////////////////////////////////////////
currentRow = currentRow + 1
If Trim(lFct) <> "D" Then
    If currentRow Mod 2 = 0 Then
        xlsBackColor = xlsGray
    Else
        xlsBackColor = xlsWhite
    End If
Else
    xlsBackColor = xlsWhite
End If
xlsForeColor = xlsBlue
xlsFontSize = 6
wsExcel.Rows(currentRow).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
'                                   '
xlsCol = 1
wsExcel.Cells(currentRow, xlsCol) = meYBIACPT0.COMPTECOM
wsExcel.Cells(currentRow, xlsCol).Font.Size = xlsFontSize
wsExcel.Cells(currentRow, xlsCol).Font.Color = xlsForeColor
wsExcel.Cells(currentRow, xlsCol).Interior.Color = xlsBackColor
'                                   '
xlsCol = 2
wsExcel.Cells(currentRow, xlsCol) = meYBIACPT0.COMPTEINT
wsExcel.Cells(currentRow, xlsCol).Font.Size = xlsFontSize
wsExcel.Cells(currentRow, xlsCol).Font.Color = xlsForeColor
wsExcel.Cells(currentRow, xlsCol).Interior.Color = xlsBackColor
If Trim(lFct) <> "D" Then
    For I = 3 To 6
        xlsCol = I
        wsExcel.Cells(currentRow, xlsCol) = ""
        wsExcel.Cells(currentRow, xlsCol).Font.Size = xlsFontSize
        wsExcel.Cells(currentRow, xlsCol).Font.Color = xlsForeColor
        wsExcel.Cells(currentRow, xlsCol).Interior.Color = xlsBackColor
    Next I
End If

End Sub

Public Function prtSAB_Stock_Monitor_xlsManual(lFct As String, fgW As MSFlexGrid, larrYBIASTO0() As typeYBIASTO0, larrYBIACPT0() As typeYBIACPT0, larrYBIACPT0_Nb As Long, ByRef wsExcel As Excel.Worksheet) As Long
' B : balance 1,1 : "B"
'             2,1 : "D" (Rupture Devise/ PCEC)              blnBalance_B_COMPTEDEV
Dim wIndex As Long, I As Integer
Dim xlsRow As Long
Dim xlsCol As Long
Dim currentRow As Long

currentRow = 0
prtSAB_Stock_Monitor_xlsManual = 0
If Trim(lFct) = "D" Then cnAdo.Open paramODBC_DSN_SAB 'JRN

xlsBlue = RGB(13, 87, 155)
xlsBlack = RGB(0, 0, 0)
xlsWhite = RGB(255, 255, 255)
xlsRed = RGB(192, 0, 0)

wsExcel.Cells.Font.Name = prtFontName_Arial
wsExcel.Cells(1, 1) = "Banque BIA (Paris)"
wsExcel.Cells(1, 1).Font.Size = 6
wsExcel.Cells(1, 1).Font.Color = xlsBlue
wsExcel.Cells(1, 1).Interior.Color = xlsWhite
wsExcel.Cells(1, 2) = "Rapprochement Opérations / Soldes comptables - au " & dateImp(YBIATAB0_DATE_CPT_J)
wsExcel.Cells(1, 2).Font.Size = 10
wsExcel.Cells(1, 2).Font.Color = xlsBlue
wsExcel.Cells(1, 2).Interior.Color = xlsWhite
wsExcel.Cells(1, 2).Font.Bold = True

If Trim(lFct) = "D" Then
    wsExcel.Cells(1, 9) = "BIA_INFO"
    wsExcel.Cells(1, 9).Font.Size = 6
    wsExcel.Cells(1, 9).Font.Color = xlsBlue
    wsExcel.Cells(1, 9).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
Else
    wsExcel.Cells(1, 6) = "BIA_INFO"
    wsExcel.Cells(1, 6).Font.Size = 6
    wsExcel.Cells(1, 6).Font.Color = xlsBlue
    wsExcel.Cells(1, 6).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
End If
Call prtSAB_Stock_Form_xlsManual(lFct, wsExcel)

currentRow = 5
For I = 1 To fgW.Rows - 1
    
    fgW.Row = I
    fgW.Col = fgW.Cols - 1: wIndex = Val(fgW.Text)
    meYBIACPT0 = larrYBIACPT0(wIndex)
    meYBIASTO0 = larrYBIASTO0(wIndex)

    Select Case Trim(lFct)
        Case "L": Call prtSAB_Stock_Line_xlsManual(lFct, wsExcel, currentRow)
        Case "D": Call prtSAB_Stock_Line_xlsManual(lFct, wsExcel, currentRow)
                  Call prtSAB_Stock_Détail_xlsManual(wsExcel, currentRow)
     End Select
     
     prevYBIACPT0 = meYBIACPT0
Next I

  
If Trim(lFct) = "D" Then cnAdo.Close: Set cnAdo = Nothing

prtSAB_Stock_Monitor_xlsManual = currentRow

End Function


Public Sub prtSAB_Stock_Open(lFct As String)
On Error GoTo prtError

Set XPrt = Printer
frmElpPrt.Show vbModeless

Height8_6 = frmElpPrt.prtHeightDelta(8, 6)

blnFiligrane = False
Select Case lFct
    Case "L": prtOrientation = vbPRORPortrait
    Case Else: prtOrientation = vbPRORLandscape '
End Select
prtPgmName = "prtSAB_Stock"
prtTitleUsr = usrName

prtLineNb = 1
prtlineHeight = 250
prtHeaderHeight = 50 ' 100


prtFormType = ""
frmElpPrt.prtStdInit

Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub


Public Sub prtSAB_Stock_Close()
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


Public Sub prtSAB_Stock_NewLine()
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
If XPrt.CurrentY + 300 > prtMaxY Then
    frmElpPrt.prtNewPage
    prtSAB_Stock_Form
End If

End Sub




Public Sub prtSAB_Stock_Monitor(lFct As String, fgW As MSFlexGrid, larrYBIASTO0() As typeYBIASTO0, larrYBIACPT0() As typeYBIACPT0, larrYBIACPT0_Nb As Long)

' B : balance 1,1 : "B"
'             2,1 : "D" (Rupture Devise/ PCEC)              blnBalance_B_COMPTEDEV
Dim wIndex As Long, I As Integer

mFct1 = Mid$(lFct, 1, 1)

If mFct1 = "D" Then cnAdo.Open paramODBC_DSN_SAB 'JRN

prtTitleText = "Rapprochement Opérations / Soldes comptables - au " & dateImp(YBIATAB0_DATE_CPT_J)

prtFontName = prtFontName_Arial
prtSAB_Stock_Open mFct1
prtHeaderHeight = 300
prtSAB_Stock_Form
XPrt.CurrentY = XPrt.CurrentY - prtlineHeight

'XPrt.FontSize = 8
For I = 1 To fgW.Rows - 1
    
    fgW.Row = I
    fgW.Col = fgW.Cols - 1: wIndex = Val(fgW.Text)
    meYBIACPT0 = larrYBIACPT0(wIndex)
    meYBIASTO0 = larrYBIASTO0(wIndex)

    Select Case mFct1
        Case "L": prtSAB_Stock_Line
        Case "D": prtSAB_Stock_Line
                  prtSAB_Stock_Détail '!!!!! meYBIASTO0
     End Select
     
     prevYBIACPT0 = meYBIACPT0
Next I


prtSAB_Stock_Close
    
If mFct1 = "D" Then cnAdo.Close: Set cnAdo = Nothing

End Sub

Public Sub prtSAB_Stock_Form()
Dim wId As String
Dim X As String

XPrt.FontSize = 7
XPrt.FontBold = True
XPrt.DrawWidth = 2

XPrt.CurrentY = prtMinY + 50

XPrt.CurrentX = prtMinX
XPrt.Print "Client / Compte";

XPrt.CurrentX = prtMinX + 1800
XPrt.Print "Intitulé";
XPrt.CurrentX = prtMinX + 5200
XPrt.Print "Solde comptable";
XPrt.CurrentX = prtMinX + 6500
XPrt.Print "Devise";
XPrt.CurrentX = prtMinX + 8500
XPrt.Print "Ecart";
XPrt.CurrentX = prtMinX + 10300
XPrt.Print "Encours";

If mFct1 = "D" Then
    XPrt.CurrentX = prtMinX + 11500
    XPrt.Print "Contrat";
    XPrt.CurrentX = prtMinX + 14200
    XPrt.Print "du";
    XPrt.CurrentX = prtMinX + 15500
    XPrt.Print "au";
End If

'XPrt.FontSize = 8
XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY), prtLineColor
XPrt.CurrentY = XPrt.CurrentY + 50


End Sub


Public Sub prtSAB_Stock_Line()
Dim X As String

prtSAB_Stock_NewLine

curX1 = Abs(meYBIACPT0.SOLDECEN)
curX2 = Abs(meYBIASTO0.YSTOMON)
curX = Abs(curX1 - curX2)
'$JPL 20101129
'curX1 = meYBIACPT0.SOLDECEN
'curX2 = meYBIASTO0.YSTOMON
'curX = curX1 - curX2

Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY - 20, prtMaxX, XPrt.CurrentY + prtlineHeight - 50, " ", 240)

XPrt.ForeColor = prtForeColor
XPrt.FontSize = 7

XPrt.CurrentX = prtMinX
XPrt.Print meYBIASTO0.YSTOCLI;
XPrt.CurrentX = prtMinX + 1800
XPrt.Print meYBIACPT0.CLIENARA1;
XPrt.CurrentX = prtMinX + 6500
XPrt.Print meYBIASTO0.YSTODEV;
If curX <> 0 Then
    XPrt.ForeColor = vbRed
    X = Format$(curX, "### ### ### ###.00")
    XPrt.CurrentX = prtMinX + 8900 - XPrt.TextWidth(X)
    XPrt.Print X;
    XPrt.ForeColor = prtForeColor
End If

X = Format$(curX2, "### ### ### ### ##0.00")
XPrt.CurrentX = prtMinX + 10900 - XPrt.TextWidth(X)
XPrt.Print X;

If curX1 <> 0 Then
    XPrt.FontBold = False
    X = Format$(curX1, "### ### ### ### ##0.00")
    XPrt.CurrentX = prtMinX + 6400 - XPrt.TextWidth(X)
    XPrt.Print X;
End If

prtSAB_Stock_NewLine
XPrt.ForeColor = vbBlue
XPrt.FontSize = 6

XPrt.CurrentX = prtMinX
XPrt.Print meYBIACPT0.COMPTECOM;

XPrt.CurrentX = prtMinX + 1800
XPrt.Print meYBIACPT0.COMPTEINT;

End Sub
Public Sub prtSAB_Stock_Line_Détail(lYBIASTO0 As typeYBIASTO0)

XPrt.ForeColor = vbBlue
XPrt.FontItalic = True
'xxxx Modification montant négatif 21/12/2009 Denis R.
'Imprime les montants négatifs et les lignes négatives en rouge
If lYBIASTO0.YSTOMON < 0 Then
    XPrt.ForeColor = vbRed
    X = Format$(lYBIASTO0.YSTOMON, "### ### ### ### ##0.00")
Else
    X = Format$(Abs(lYBIASTO0.YSTOMON), "### ### ### ### ##0.00")
End If
XPrt.CurrentX = prtMinX + 10900 - XPrt.TextWidth(X)
XPrt.Print X;

XPrt.CurrentX = prtMinX + 11000
XPrt.Print lYBIASTO0.YSTOAPP;
XPrt.CurrentX = prtMinX + 11500
XPrt.Print lYBIASTO0.YSTOOPE;
XPrt.CurrentX = prtMinX + 12000
XPrt.Print lYBIASTO0.YSTONAT;
X = Format$(lYBIASTO0.YSTONUM, "### ### ### ###")
XPrt.CurrentX = prtMinX + 13500 - XPrt.TextWidth(X)
XPrt.Print X;
XPrt.CurrentX = prtMinX + 13700
XPrt.Print dateImp10(lYBIASTO0.YSTODEB);
XPrt.CurrentX = prtMinX + 15000
XPrt.Print dateImp10(lYBIASTO0.YSTOFIN);
XPrt.FontItalic = False

End Sub

Public Sub prtSAB_Stock_Détail()
Dim Nb As Long
Dim xSQL As String
Dim V
Set rsAdo = Nothing
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIASTO0 where " _
     & "YSTOPCI = '" & meYBIASTO0.YSTOPCI & "'" _
     & "AND YSTODEV = '" & meYBIASTO0.YSTODEV & "'" _
     & "AND YSTOCLI = " & meYBIASTO0.YSTOCLI

Set rsAdo = cnAdo.Execute(xSQL)
Nb = 0
Do While Not rsAdo.EOF
    V = rsYBIASTO0_GetBuffer(rsAdo, xYBIASTO0)
    If Not IsNull(V) Then
        MsgBox V, vbCritical, "prtSAB_Stock_SQL"
        Exit Sub
    Else
        Nb = Nb + 1
        If Nb > 1 Then prtSAB_Stock_NewLine
        prtSAB_Stock_Line_Détail xYBIASTO0
    End If
    rsAdo.MoveNext
Loop

End Sub
