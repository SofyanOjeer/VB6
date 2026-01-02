Attribute VB_Name = "prtBIA_Compta_Control"
Option Explicit
Type typeYBIACPT_C

    COMPTECOM       As String * 20                    ' NUMERO COMPTE
    COMPTEOBL       As String * 10                    ' COMPTE OBLIGATOIRE
    COMPTEINT       As String * 32                    ' INTITULE
    COMPTEDEV       As String * 3                     ' TABLES BASE 013
    SOLDEDMO        As Long                           ' DATE DERNIER MVT
    SOLDECEN        As Currency                       ' SOLDE ENCOURS
    
    SOLDEJ_2        As Currency                       ' SOLDE J-2
    MOUVEMMON_DB    As Currency                       ' MONTANT
    MOUVEMMON_CR    As Currency                       ' MONTANT
    
    SOLDEJ_2_DB     As Currency                       ' SOLDE J-2
    SOLDEJ_2_CR     As Currency                       ' SOLDE J-2
    SOLDEJ_1_DB     As Currency                       ' SOLDE J-2
    SOLDEJ_1_CR     As Currency                       ' SOLDE J-2
End Type

Public arrYBIACPT_C() As typeYBIACPT_C, arrYBIACPT_C_Nb As Long, arrYBIACPT_C_NbMax As Long
Public arrDev_B() As typeYBIACPT_C, arrDev_HB() As typeYBIACPT_C, arrDev_Nb As Long
Dim curX As Currency
Dim xSQL As String, V


Dim Height8_6 As Integer




Public Sub prtBIA_Compta_Control_Anomalie_xlsManual(lText As String, ByRef currentRow As Long, wsExcel As Excel.Worksheet)

    Call prtBIA_Compta_Control_NewLine_xlsManual(currentRow, wsExcel)
    wsExcel.Cells(currentRow, 1) = lText

End Sub

Public Sub prtBIA_Compta_Control_Cumul_xlsManual(ByRef currentRow As Long, wsExcel As Excel.Worksheet)
Dim K As Long, X As String
Dim kDev As Integer
kDev = 1

For K = 1 To arrYBIACPT_C_Nb
    If arrYBIACPT_C(K).COMPTEDEV <> arrDev_B(kDev).COMPTEDEV Then
        For kDev = 1 To arrDev_Nb
            If arrYBIACPT_C(K).COMPTEDEV = arrDev_B(kDev).COMPTEDEV Then Exit For
        Next kDev
    End If
    If Mid$(arrYBIACPT_C(K).COMPTEOBL, 1, 1) <> "9" Then
        arrDev_B(kDev).SOLDEDMO = arrDev_B(kDev).SOLDEDMO + 1
        arrDev_B(kDev).MOUVEMMON_DB = arrDev_B(kDev).MOUVEMMON_DB + arrYBIACPT_C(K).MOUVEMMON_DB
        arrDev_B(kDev).MOUVEMMON_CR = arrDev_B(kDev).MOUVEMMON_CR + arrYBIACPT_C(K).MOUVEMMON_CR
        If arrYBIACPT_C(K).SOLDEJ_2 < 0 Then
            arrDev_B(kDev).SOLDEJ_2_CR = arrDev_B(kDev).SOLDEJ_2_CR + arrYBIACPT_C(K).SOLDEJ_2
        Else
            arrDev_B(kDev).SOLDEJ_2_DB = arrDev_B(kDev).SOLDEJ_2_DB + arrYBIACPT_C(K).SOLDEJ_2
        End If
         If arrYBIACPT_C(K).SOLDECEN < 0 Then
            arrDev_B(kDev).SOLDEJ_1_CR = arrDev_B(kDev).SOLDEJ_1_CR + arrYBIACPT_C(K).SOLDECEN
        Else
            arrDev_B(kDev).SOLDEJ_1_DB = arrDev_B(kDev).SOLDEJ_1_DB + arrYBIACPT_C(K).SOLDECEN
        End If
    Else
        arrDev_HB(kDev).SOLDEDMO = arrDev_HB(kDev).SOLDEDMO + 1
        arrDev_HB(kDev).MOUVEMMON_DB = arrDev_HB(kDev).MOUVEMMON_DB + arrYBIACPT_C(K).MOUVEMMON_DB
        arrDev_HB(kDev).MOUVEMMON_CR = arrDev_HB(kDev).MOUVEMMON_CR + arrYBIACPT_C(K).MOUVEMMON_CR
        If arrYBIACPT_C(K).SOLDEJ_2 < 0 Then
            arrDev_HB(kDev).SOLDEJ_2_CR = arrDev_HB(kDev).SOLDEJ_2_CR + arrYBIACPT_C(K).SOLDEJ_2
        Else
            arrDev_HB(kDev).SOLDEJ_2_DB = arrDev_HB(kDev).SOLDEJ_2_DB + arrYBIACPT_C(K).SOLDEJ_2
        End If
         If arrYBIACPT_C(K).SOLDECEN < 0 Then
            arrDev_HB(kDev).SOLDEJ_1_CR = arrDev_HB(kDev).SOLDEJ_1_CR + arrYBIACPT_C(K).SOLDECEN
        Else
            arrDev_HB(kDev).SOLDEJ_1_DB = arrDev_HB(kDev).SOLDEJ_1_DB + arrYBIACPT_C(K).SOLDECEN
        End If
    End If
Next K
For kDev = 1 To arrDev_Nb
    If arrDev_B(kDev).SOLDEDMO > 1 Then
        Call prtBIA_Compta_Control_NewLine3_xlsManual(currentRow, wsExcel)
        If arrDev_B(kDev).SOLDEJ_2_DB + arrDev_B(kDev).SOLDEJ_2_CR <> 0 _
        Or arrDev_B(kDev).MOUVEMMON_DB + arrDev_B(kDev).MOUVEMMON_CR <> 0 _
        Or arrDev_B(kDev).SOLDEJ_1_DB + arrDev_B(kDev).SOLDEJ_1_CR <> 0 Then
            wsExcel.Cells(currentRow, 9) = "ERREUR B / HB"
            wsExcel.Cells(currentRow, 9).Font.Color = vbMagenta
        End If
        wsExcel.Cells(currentRow, 1) = arrDev_B(kDev).COMPTEDEV
        wsExcel.Cells(currentRow, 2) = "B"
        X = Format$(arrDev_B(kDev).SOLDEJ_2_DB, "### ### ### ### ##0.00")
        wsExcel.Cells(currentRow, 3) = X
        X = Format$(arrDev_B(kDev).SOLDEJ_2_CR, "### ### ### ### ##0.00")
        wsExcel.Cells(currentRow, 4) = X
        X = Format$(arrDev_B(kDev).MOUVEMMON_DB, "### ### ### ### ##0.00")
        wsExcel.Cells(currentRow, 5) = X
        X = Format$(arrDev_B(kDev).MOUVEMMON_CR, "### ### ### ### ##0.00")
        wsExcel.Cells(currentRow, 6) = X
        X = Format$(arrDev_B(kDev).SOLDEJ_1_DB, "### ### ### ### ##0.00")
        wsExcel.Cells(currentRow, 7) = X
        X = Format$(arrDev_B(kDev).SOLDEJ_1_CR, "### ### ### ### ##0.00")
        wsExcel.Cells(currentRow, 8) = X
    End If
    If arrDev_HB(kDev).SOLDEDMO > 1 Then
        Call prtBIA_Compta_Control_NewLine3_xlsManual(currentRow, wsExcel)
       If arrDev_HB(kDev).SOLDEJ_2_DB + arrDev_HB(kDev).SOLDEJ_2_CR <> 0 _
        Or arrDev_HB(kDev).MOUVEMMON_DB + arrDev_HB(kDev).MOUVEMMON_CR <> 0 _
        Or arrDev_HB(kDev).SOLDEJ_1_DB + arrDev_HB(kDev).SOLDEJ_1_CR <> 0 Then
            wsExcel.Cells(currentRow, 9) = "ERREUR B / HB"
            wsExcel.Cells(currentRow, 9).Font.Color = vbMagenta
        End If
        wsExcel.Cells(currentRow, 1) = arrDev_HB(kDev).COMPTEDEV
        wsExcel.Cells(currentRow, 2) = "HB"
        X = Format$(arrDev_HB(kDev).SOLDEJ_2_DB, "### ### ### ### ##0.00")
        wsExcel.Cells(currentRow, 3) = X
        X = Format$(arrDev_HB(kDev).SOLDEJ_2_CR, "### ### ### ### ##0.00")
        wsExcel.Cells(currentRow, 4) = X
        X = Format$(arrDev_HB(kDev).MOUVEMMON_DB, "### ### ### ### ##0.00")
        wsExcel.Cells(currentRow, 5) = X
        X = Format$(arrDev_HB(kDev).MOUVEMMON_CR, "### ### ### ### ##0.00")
        wsExcel.Cells(currentRow, 6) = X
        X = Format$(arrDev_HB(kDev).SOLDEJ_1_DB, "### ### ### ### ##0.00")
        wsExcel.Cells(currentRow, 7) = X
        X = Format$(arrDev_HB(kDev).SOLDEJ_1_CR, "### ### ### ### ##0.00")
        wsExcel.Cells(currentRow, 8) = X
    End If
Next kDev
currentRow = currentRow + 1
Range("A5:I5").Select
Selection.Copy
Range("A" & CStr(currentRow)).Select
ActiveSheet.Paste
'On supprime les lignes modèles de la feuille SOLDES
Rows("2:5").Select
Selection.Delete
currentRow = currentRow - 4
End Sub

'---------------------------------------------------------
Public Sub prtBIA_Compta_Control_Form()
'---------------------------------------------------------
Dim X As String

XPrt.DrawWidth = 1
XPrt.FontSize = 7: XPrt.FontBold = True
prtCurrentY = XPrt.CurrentY
Call frmElpPrt.prtTrame(prtMinX, prtCurrentY, prtMaxX, prtCurrentY + prtHeaderHeight, "B", 240)
XPrt.Line (prtMinX + 1100, prtCurrentY)-(prtMinX + 1100, prtMaxY), prtLineColor
XPrt.Line (prtMinX + 5100, prtCurrentY)-(prtMinX + 5100, prtMaxY), prtLineColor
XPrt.Line (prtMinX + 9100, prtCurrentY)-(prtMinX + 9100, prtMaxY), prtLineColor
XPrt.Line (prtMinX + 13100, prtCurrentY)-(prtMinX + 13100, prtMaxY), prtLineColor
'---------------------------------------------------------
XPrt.CurrentY = prtCurrentY + 50

XPrt.CurrentX = prtMinX + 50: XPrt.Print "Devise";
'XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 6
XPrt.CurrentX = prtMinX + 2500: XPrt.Print "Soldes au " & dateImp(YBIATAB0_DATE_CPT_JP1);
XPrt.CurrentX = prtMinX + 6500: XPrt.Print "Mouvements du jour";
XPrt.CurrentX = prtMinX + 10500: XPrt.Print "Soldes au " & dateImp(YBIATAB0_DATE_CPT_J);
'XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 8

XPrt.FontBold = False
XPrt.FontSize = 10
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight + 50

End Sub

Public Sub prtBIA_Compta_Control_NewLine_xlsManual(ByRef currentRow As Long, wsExcel As Excel.Worksheet)

    currentRow = currentRow + 1
    Range("A2:H2").Select
    Selection.Copy
    Range("A" & CStr(currentRow)).Select
    ActiveSheet.Paste

End Sub

Public Sub prtBIA_Compta_Control_NewLine3_xlsManual(ByRef currentRow As Long, wsExcel As Excel.Worksheet)

    currentRow = currentRow + 1
    Range("A3:H3").Select
    Selection.Copy
    Range("A" & CStr(currentRow)).Select
    ActiveSheet.Paste

End Sub

Public Sub prtBIA_Compta_Control_Open()
On Error GoTo prtError

Set XPrt = Printer
frmElpPrt.Show vbModeless

Height8_6 = frmElpPrt.prtHeightDelta(8, 6)

blnFiligrane = False
prtOrientation = vbPRORLandscape '
prtPgmName = "prtBIA_Compta_Control"
prtTitleUsr = usrName
prtLineNb = 1
prtlineHeight = 250
prtHeaderHeight = 300


prtFormType = ""
frmElpPrt.prtStdInit
    XPrt.CurrentY = prtMinX
Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub


Public Sub prtBIA_Compta_Control_Close()
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




Public Sub prtBIA_Compta_Control_NewLine()
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
If XPrt.CurrentY + 300 > prtMaxY Then
    frmElpPrt.prtNewPage
    XPrt.CurrentY = prtMinX
    prtBIA_Compta_Control_Form
End If

End Sub



Public Sub prtBIA_Compta_Control_Anomalie(lText As String, blnCentré As Boolean)
prtBIA_Compta_Control_NewLine
XPrt.CurrentX = prtMinX
XPrt.FontBold = True
If blnCentré Then
    frmElpPrt.prtCentré prtMedX, lText
Else
    XPrt.Print lText;
End If

XPrt.FontBold = False

End Sub

Public Sub prtBIA_Compta_Control_Cumul()
Dim K As Long, X As String
Dim kDev As Integer
kDev = 1

For K = 1 To arrYBIACPT_C_Nb
    If arrYBIACPT_C(K).COMPTEDEV <> arrDev_B(kDev).COMPTEDEV Then
        For kDev = 1 To arrDev_Nb
            If arrYBIACPT_C(K).COMPTEDEV = arrDev_B(kDev).COMPTEDEV Then Exit For
        Next kDev
    End If
    If Mid$(arrYBIACPT_C(K).COMPTEOBL, 1, 1) <> "9" Then
        arrDev_B(kDev).SOLDEDMO = arrDev_B(kDev).SOLDEDMO + 1
        arrDev_B(kDev).MOUVEMMON_DB = arrDev_B(kDev).MOUVEMMON_DB + arrYBIACPT_C(K).MOUVEMMON_DB
        arrDev_B(kDev).MOUVEMMON_CR = arrDev_B(kDev).MOUVEMMON_CR + arrYBIACPT_C(K).MOUVEMMON_CR
        If arrYBIACPT_C(K).SOLDEJ_2 < 0 Then
            arrDev_B(kDev).SOLDEJ_2_CR = arrDev_B(kDev).SOLDEJ_2_CR + arrYBIACPT_C(K).SOLDEJ_2
        Else
            arrDev_B(kDev).SOLDEJ_2_DB = arrDev_B(kDev).SOLDEJ_2_DB + arrYBIACPT_C(K).SOLDEJ_2
        End If
         If arrYBIACPT_C(K).SOLDECEN < 0 Then
            arrDev_B(kDev).SOLDEJ_1_CR = arrDev_B(kDev).SOLDEJ_1_CR + arrYBIACPT_C(K).SOLDECEN
        Else
            arrDev_B(kDev).SOLDEJ_1_DB = arrDev_B(kDev).SOLDEJ_1_DB + arrYBIACPT_C(K).SOLDECEN
        End If
    Else
        arrDev_HB(kDev).SOLDEDMO = arrDev_HB(kDev).SOLDEDMO + 1
        arrDev_HB(kDev).MOUVEMMON_DB = arrDev_HB(kDev).MOUVEMMON_DB + arrYBIACPT_C(K).MOUVEMMON_DB
        arrDev_HB(kDev).MOUVEMMON_CR = arrDev_HB(kDev).MOUVEMMON_CR + arrYBIACPT_C(K).MOUVEMMON_CR
        If arrYBIACPT_C(K).SOLDEJ_2 < 0 Then
            arrDev_HB(kDev).SOLDEJ_2_CR = arrDev_HB(kDev).SOLDEJ_2_CR + arrYBIACPT_C(K).SOLDEJ_2
        Else
            arrDev_HB(kDev).SOLDEJ_2_DB = arrDev_HB(kDev).SOLDEJ_2_DB + arrYBIACPT_C(K).SOLDEJ_2
        End If
         If arrYBIACPT_C(K).SOLDECEN < 0 Then
            arrDev_HB(kDev).SOLDEJ_1_CR = arrDev_HB(kDev).SOLDEJ_1_CR + arrYBIACPT_C(K).SOLDECEN
        Else
            arrDev_HB(kDev).SOLDEJ_1_DB = arrDev_HB(kDev).SOLDEJ_1_DB + arrYBIACPT_C(K).SOLDECEN
        End If
    End If
          
Next K

prtBIA_Compta_Control_Form
For kDev = 1 To arrDev_Nb
    If arrDev_B(kDev).SOLDEDMO > 1 Then
        prtBIA_Compta_Control_NewLine
        If arrDev_B(kDev).SOLDEJ_2_DB + arrDev_B(kDev).SOLDEJ_2_CR <> 0 _
        Or arrDev_B(kDev).MOUVEMMON_DB + arrDev_B(kDev).MOUVEMMON_CR <> 0 _
        Or arrDev_B(kDev).SOLDEJ_1_DB + arrDev_B(kDev).SOLDEJ_1_CR <> 0 Then
            XPrt.CurrentX = prtMinX + 13300
            XPrt.FontBold = True
            XPrt.ForeColor = vbMagenta
            XPrt.Print "ERREUR B / HB";
        Else
            XPrt.FontBold = False
            XPrt.ForeColor = vbBlack
        End If

        XPrt.CurrentX = prtMinX + 50
        XPrt.Print arrDev_B(kDev).COMPTEDEV;
        XPrt.CurrentX = prtMinX + 650
        XPrt.Print "B";
        
        X = Format$(arrDev_B(kDev).SOLDEJ_2_DB, "### ### ### ### ##0.00")
        XPrt.CurrentX = prtMinX + 3000 - XPrt.TextWidth(X)
        XPrt.Print X;
        X = Format$(arrDev_B(kDev).SOLDEJ_2_CR, "### ### ### ### ##0.00")
        XPrt.CurrentX = prtMinX + 5000 - XPrt.TextWidth(X)
        XPrt.Print X;

        X = Format$(arrDev_B(kDev).MOUVEMMON_DB, "### ### ### ### ##0.00")
        XPrt.CurrentX = prtMinX + 7000 - XPrt.TextWidth(X)
        XPrt.Print X;
        X = Format$(arrDev_B(kDev).MOUVEMMON_CR, "### ### ### ### ##0.00")
        XPrt.CurrentX = prtMinX + 9000 - XPrt.TextWidth(X)
        XPrt.Print X;
        
        X = Format$(arrDev_B(kDev).SOLDEJ_1_DB, "### ### ### ### ##0.00")
        XPrt.CurrentX = prtMinX + 11000 - XPrt.TextWidth(X)
        XPrt.Print X;
        X = Format$(arrDev_B(kDev).SOLDEJ_1_CR, "### ### ### ### ##0.00")
        XPrt.CurrentX = prtMinX + 13000 - XPrt.TextWidth(X)
        XPrt.Print X;
    End If
    If arrDev_HB(kDev).SOLDEDMO > 1 Then
        prtBIA_Compta_Control_NewLine
       If arrDev_HB(kDev).SOLDEJ_2_DB + arrDev_HB(kDev).SOLDEJ_2_CR <> 0 _
        Or arrDev_HB(kDev).MOUVEMMON_DB + arrDev_HB(kDev).MOUVEMMON_CR <> 0 _
        Or arrDev_HB(kDev).SOLDEJ_1_DB + arrDev_HB(kDev).SOLDEJ_1_CR <> 0 Then
            XPrt.CurrentX = prtMinX + 13300
            XPrt.FontBold = True
            XPrt.ForeColor = vbMagenta
            XPrt.Print "ERREUR B / HB";
        Else
            XPrt.FontBold = False
            XPrt.ForeColor = vbBlack
        End If
        XPrt.CurrentX = prtMinX + 50
        XPrt.Print arrDev_HB(kDev).COMPTEDEV;
        XPrt.CurrentX = prtMinX + 650
        XPrt.Print "HB";
        
        X = Format$(arrDev_HB(kDev).SOLDEJ_2_DB, "### ### ### ### ##0.00")
        XPrt.CurrentX = prtMinX + 3000 - XPrt.TextWidth(X)
        XPrt.Print X;
        X = Format$(arrDev_HB(kDev).SOLDEJ_2_CR, "### ### ### ### ##0.00")
        XPrt.CurrentX = prtMinX + 5000 - XPrt.TextWidth(X)
        XPrt.Print X;

        X = Format$(arrDev_HB(kDev).MOUVEMMON_DB, "### ### ### ### ##0.00")
        XPrt.CurrentX = prtMinX + 7000 - XPrt.TextWidth(X)
        XPrt.Print X;
        X = Format$(arrDev_HB(kDev).MOUVEMMON_CR, "### ### ### ### ##0.00")
        XPrt.CurrentX = prtMinX + 9000 - XPrt.TextWidth(X)
        XPrt.Print X;
        
        X = Format$(arrDev_HB(kDev).SOLDEJ_1_DB, "### ### ### ### ##0.00")
        XPrt.CurrentX = prtMinX + 11000 - XPrt.TextWidth(X)
        XPrt.Print X;
        X = Format$(arrDev_HB(kDev).SOLDEJ_1_CR, "### ### ### ### ##0.00")
        XPrt.CurrentX = prtMinX + 13000 - XPrt.TextWidth(X)
        XPrt.Print X;
    End If

Next kDev
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY), prtLineColor
End Sub
