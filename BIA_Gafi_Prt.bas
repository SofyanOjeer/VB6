Attribute VB_Name = "prtBIA_Gafi"


'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim X As String, I As Integer, Height8_6 As Integer
Dim V
Dim Nb1 As Integer, Nb2 As Integer

Dim meZCOMPTE0 As typeZCOMPTE0
Dim meYBIAMVT0 As typeYBIAMVT0

Dim meCV1 As typeCV, meCV2 As typeCV
Dim X1 As String, X2 As String
Dim mLog_Compte      As String * 11
Dim Col4 As Integer, Col5 As Integer, Col6 As Integer, Col7 As Integer, Col8 As Integer

Dim Conversion As String
Dim curEur As Currency, mDevIso As String, curT As Currency
Dim nbDB As Long, nbCR As Long, curDB As Currency, curCR As Currency
Dim paramCompteGafi_Seuil As Currency, paramCompteGafi_curMin As Currency
Dim paramCompteGafi_Etat As String
Dim mCLIENARES As String, xCLIENARES As String


Dim blnGAFI_Open As Boolean, blnPrintCompte As Boolean
Public arrYBIAMVT0() As typeYBIAMVT0, arrYBIAMVT0_Nb As Long, arrYBIAMVT0_Max As Long
Dim arrYBIAMVT0_K1 As Long, arrYBIAMVT0_K2 As Long
Dim meYBIACPT0 As typeYBIACPT0
Public Sub prtBIA_Gafi_Line9(arrYBIAMVT9 As typeYBIAMVT9, ByRef blnPrintCpt As Boolean)
Dim lX As Long, lMax As Long
Dim wUnit As String, X As String

    If XPrt.CurrentY + prtlineHeight * 3 > prtMaxY Then
        frmElpPrt.prtNewPage
        prtBIA_Gafi_Form
    End If
    Nb1 = Nb1 + 1
    XPrt.FontSize = 8
    XPrt.FontBold = False
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinX + 50
    XPrt.FontSize = 6
    XPrt.CurrentY = XPrt.CurrentY + Height8_6
    If blnPrintCpt Then
        blnPrintCpt = False
        Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY, Col4 - 20, XPrt.CurrentY + prtlineHeight, " ", 235)
        XPrt.CurrentY = XPrt.CurrentY + 50
        XPrt.FontBold = True
        XPrt.CurrentX = prtMinX + 50
        XPrt.Print arrYBIAMVT9.COMPTEDEV & "  ";
        XPrt.Print arrYBIAMVT9.MOUVEMCOM;
        If XPrt.CurrentX < 1500 Then
            XPrt.CurrentX = prtMinX + 1400
        Else
            XPrt.CurrentX = prtMinX + 1800
        End If
        XPrt.Print arrYBIAMVT9.CLIENARSD & " - " & arrYBIAMVT9.COMPTEINT;
        XPrt.FontBold = False
        XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    End If
    XPrt.FontBold = True
    X = Format$(Abs(arrYBIAMVT9.MOUVEMMON), "## ### ### ### ### ##0.00")
    XPrt.CurrentX = IIf(arrYBIAMVT9.MOUVEMMON > 0, Col5, Col6) - 50 - XPrt.TextWidth(X)
    XPrt.Print X;
    XPrt.FontBold = False
    If Trim(arrYBIAMVT9.COMPTEDEV) <> "EUR" Then
        curEur = prtBIA_Gafi_CV9(arrYBIAMVT9.COMPTEDEV, arrYBIAMVT9.MOUVEMDTR, arrYBIAMVT9.MOUVEMMON)
        XPrt.FontItalic = True
        X = Format$(curEur, "## ### ### ### ### ##0.00")
        XPrt.CurrentX = Col6 + 1200 - XPrt.TextWidth(X)
        XPrt.Print X;
        XPrt.FontItalic = False
        XPrt.CurrentX = Col4 - 400
        XPrt.Print arrYBIAMVT9.COMPTEDEV;
    Else
        curEur = 0
        XPrt.FontItalic = True
        X = "                         "
        XPrt.CurrentX = Col6 + 1200 - XPrt.TextWidth(X)
        XPrt.Print X;
        XPrt.FontItalic = False
        XPrt.CurrentX = Col4 - 400
    End If
    If arrYBIAMVT9.MOUVEMDTR <> arrYBIAMVT9.MOUVEMDOP Then
        XPrt.CurrentX = prtMaxX - 3000
        Call frmElpPrt.prtTrame(XPrt.CurrentX - 50, XPrt.CurrentY - 50, XPrt.CurrentX + 800, XPrt.CurrentY + prtlineHeight - 50, " ", 235)
        XPrt.FontBold = True
    End If
    XPrt.CurrentX = prtMaxX - 3000
    XPrt.Print dateIBM10(arrYBIAMVT9.MOUVEMDOP, False);
    XPrt.FontBold = False
    XPrt.CurrentX = prtMaxX - 4000
    If arrYBIAMVT9.MOUVEMDVA <> arrYBIAMVT9.MOUVEMDOP Then XPrt.Print dateIBM10(arrYBIAMVT9.MOUVEMDVA, False);
    XPrt.CurrentX = prtMinX + 1400
    XPrt.Print Trim(arrYBIAMVT9.LIBELLIB1) & " " & Trim(arrYBIAMVT9.LIBELLIB2) & Trim(arrYBIAMVT9.LIBELLIB3);
    wUnit = Table_Ope_Unit(arrYBIAMVT9.MOUVEMSER & arrYBIAMVT9.MOUVEMSSE & arrYBIAMVT9.MOUVEMOPE)
    XPrt.CurrentX = prtMaxX - 2000: XPrt.Print wUnit;
    XPrt.CurrentX = prtMaxX - 1600: XPrt.Print arrYBIAMVT9.MOUVEMOPE;
    XPrt.CurrentX = prtMaxX - 1300: XPrt.Print arrYBIAMVT9.MOUVEMNUM;
    XPrt.CurrentX = prtMaxX - 500: XPrt.Print arrYBIAMVT9.MOUVEMEVE;
    XPrt.CurrentY = XPrt.CurrentY - Height8_6
    
End Sub
Public Function prtBIA_Gafi_CV9(lDevise As String, lDtr As Long, lMonTant As Currency) As Currency
If lDevise <> "EUR" Then
    meCV1.DeviseIso = lDevise
    meCV1.DeviseN = 0
    meCV1.Montant = lMonTant
    meCV1.OpéAmj = lDtr + 19000000
    meCV2.OpéAmj = meCV1.OpéAmj
       
    Call CV_Calc("J  ", meCV1, meCV2)
    meCV2.Montant = meCV2.Montant
Else
    meCV2.Montant = lMonTant
End If
prtBIA_Gafi_CV9 = meCV2.Montant

End Function
Public Sub prtBIA_Gafi_Close9(aPrinter As String)
Dim xXPrt As Printer
                        
On Error GoTo prtError
        
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY), prtLineColor
    XPrt.CurrentY = XPrt.CurrentY + 50
    XPrt.FontBold = True
    XPrt.CurrentX = prtMinX: XPrt.Print Nb1 & " mouvements";
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY), prtLineColor
    DoEvents: If prtKillDoc Then frmElpPrt.Hide: Exit Sub
    Call frmElpPrt.prtEndDoc(1000)
    frmElpPrt.Hide

    For Each xXPrt In Printers
        If InStr(1, UCase$(Trim(xXPrt.Devicename)), UCase(aPrinter)) > 0 Then
           Set Printer = xXPrt
           Exit For
        End If
    Next

Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "prtBIA_Gafi_Close9")
frmElpPrt.Hide
End Sub

Public Sub prtBIA_Gafi_Line9_xlsManual(arrYBIAMVT9 As typeYBIAMVT9, ByRef blnPrintCpt As Boolean, ligneorigine As Long, ByRef currentRow As Long, wsExcel As Excel.Worksheet)
Dim lX As Long, lMax As Long
Dim wUnit As String, X As String
Dim laColonne As Long

    If blnPrintCpt Then
        currentRow = currentRow + 1
        Range("A4:J5").Select
        Selection.Copy
        Range("A" & CStr(currentRow)).Select
        ActiveSheet.Paste
        currentRow = currentRow + 1
        wsExcel.Cells(currentRow, 1) = arrYBIAMVT9.COMPTEDEV & "  " & arrYBIAMVT9.MOUVEMCOM
        wsExcel.Cells(currentRow, 2) = arrYBIAMVT9.CLIENARSD & " - " & arrYBIAMVT9.COMPTEINT
    End If
    currentRow = currentRow + 1
    ' DR En dur car pas le temps de développer saut de page + entete +... 07/02/2020 ----------------------------------
    If currentRow = 48 Then
        Call insere_entete_page(wsExcel, "1:3", 3, currentRow)
        currentRow = currentRow + 3
    End If
    '----------------------------------
    Range("A" & CStr(ligneorigine) & ":J" & CStr(ligneorigine)).Select
    Selection.Copy
    Range("A" & CStr(currentRow)).Select
    ActiveSheet.Paste
    X = Format$(Abs(arrYBIAMVT9.MOUVEMMON), "## ### ### ### ### ##0.00")
    laColonne = IIf(arrYBIAMVT9.MOUVEMMON > 0, 4, 5)
    wsExcel.Cells(currentRow, laColonne) = X
    If Trim(arrYBIAMVT9.COMPTEDEV) <> "EUR" Then
        curEur = prtBIA_Gafi_CV9(arrYBIAMVT9.COMPTEDEV, arrYBIAMVT9.MOUVEMDTR, arrYBIAMVT9.MOUVEMMON)
        X = Format$(curEur, "## ### ### ### ### ##0.00")
        wsExcel.Cells(currentRow, 6) = X
        wsExcel.Cells(currentRow, 3) = arrYBIAMVT9.COMPTEDEV
    Else
        wsExcel.Cells(currentRow, 6) = ""
        wsExcel.Cells(currentRow, 3) = ""
    End If
    wsExcel.Cells(currentRow, 8) = dateIBM10(arrYBIAMVT9.MOUVEMDOP, False)
    If arrYBIAMVT9.MOUVEMDVA <> arrYBIAMVT9.MOUVEMDOP Then
        wsExcel.Cells(currentRow, 7) = dateIBM10(arrYBIAMVT9.MOUVEMDVA, False)
    End If
    wsExcel.Cells(currentRow, 2) = Trim(arrYBIAMVT9.LIBELLIB1) & " " & Trim(arrYBIAMVT9.LIBELLIB2) & Trim(arrYBIAMVT9.LIBELLIB3)
    wUnit = Table_Ope_Unit(arrYBIAMVT9.MOUVEMSER & arrYBIAMVT9.MOUVEMSSE & arrYBIAMVT9.MOUVEMOPE)
    wsExcel.Cells(currentRow, 9) = wUnit & " " & arrYBIAMVT9.MOUVEMOPE
    wsExcel.Cells(currentRow, 10) = arrYBIAMVT9.MOUVEMNUM & " " & arrYBIAMVT9.MOUVEMEVE

End Sub

Public Sub prtBIA_Gafi_Open9()
Dim xXPrt As Printer
Dim K As Long

On Error GoTo prtError

If nomDuServeur <> paramServerSplf Then
    For Each xXPrt In Printers
        K = InStr(1, UCase$(Trim(xXPrt.Devicename)), "PDFCREATOR")
        If K > 0 Then
           Set Printer = xXPrt
           Shell "C:\Program Files\PDFCreator\PDFCreator.exe", vbMinimizedNoFocus
           Exit For
        End If
    Next
End If
Set XPrt = Printer

Height8_6 = frmElpPrt.prtHeightDelta(8, 6)

prtLineNb = 1

frmElpPrt.Show vbModeless

prtOrientation = vbPRORLandscape
prtPgmName = "prtBIA_Gafi"
prtTitleUsr = usrName

prtlineHeight = 250
prtHeaderHeight = 300

frmElpPrt.prtStdInit

'recCompteInit meZCOMPTE0
'meCV1 = CV_Euro
meCV1.CoursCompta = "C"
meCV1.OpéAmj = DSys
meCV1.Normal = "P"
meCV1.AchatVente = " "
meCV2 = meCV1:
Col4 = 7000: Col5 = 8700: Col6 = 10300
prtBIA_Gafi_Form

blnGAFI_Open = True

Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "prtBIA_Gafi_Open9")
frmElpPrt.Hide

End Sub


'---------------------------------------------------------
 Public Sub prtBIA_Gafi_Open()
'---------------------------------------------------------

On Error GoTo prtError

Set XPrt = Printer
Height8_6 = frmElpPrt.prtHeightDelta(8, 6)


prtLineNb = 1

frmElpPrt.Show vbModeless


prtOrientation = vbPRORLandscape
prtPgmName = "prtBIA_Gafi"
prtTitleUsr = usrName

prtlineHeight = 250
prtHeaderHeight = 300

frmElpPrt.prtStdInit

'recCompteInit meZCOMPTE0
'meCV1 = CV_Euro
meCV1.CoursCompta = "C"
meCV1.OpéAmj = DSys
meCV1.Normal = "P"
meCV1.AchatVente = " "
meCV2 = meCV1:
Col4 = 7000: Col5 = 8700: Col6 = 10300
prtBIA_Gafi_Form

blnGAFI_Open = True

Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "prtBIA_Gafi_Open")
frmElpPrt.Hide
End Sub
'---------------------------------------------------------
 Public Sub prtBIA_Gafi_Close()
'---------------------------------------------------------
                        
On Error GoTo prtError
        
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY), prtLineColor
XPrt.CurrentY = XPrt.CurrentY + 50
XPrt.FontBold = True
XPrt.CurrentX = prtMinX: XPrt.Print Nb1 & " mouvements";
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY), prtLineColor

DoEvents: If prtKillDoc Then frmElpPrt.Hide: Exit Sub

Call frmElpPrt.prtEndDoc(1000)
frmElpPrt.Hide

Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "prtBIA_Gafi_Close")
frmElpPrt.Hide
End Sub

'---------------------------------------------------------
Public Sub prtBIA_Gafi_Form()
'---------------------------------------------------------
Dim X As String
XPrt.DrawWidth = 3
XPrt.FontSize = 8
Call frmElpPrt.prtTrame(prtMinX, prtMinY, prtMaxX, prtMinY + prtHeaderHeight, "B", 235)

'---------------------------------------------------------


XPrt.DrawWidth = 1
XPrt.Line (Col4, prtMinY)-(Col4, prtMaxY), prtLineColor
XPrt.Line (Col6, prtMinY)-(Col6, prtMaxY), prtLineColor
XPrt.CurrentY = prtMinY + 50

XPrt.FontBold = True
XPrt.FontBold = True
XPrt.CurrentX = 400
XPrt.Print "Compte";

XPrt.CurrentX = 1600
XPrt.Print "Intitulé";

XPrt.CurrentX = 8200
XPrt.Print "Débit";

XPrt.CurrentX = 9600
XPrt.Print "Crédit";

XPrt.FontSize = 6

XPrt.CurrentX = prtMaxX - 3000: XPrt.Print "Date Opé";

XPrt.CurrentX = prtMaxX - 4000: XPrt.Print "Date Valeur";

XPrt.CurrentX = 11000
XPrt.FontItalic = True
XPrt.Print "cv/EUR";
XPrt.FontItalic = False

XPrt.CurrentX = prtMaxX - 2000: XPrt.Print "Service";
XPrt.CurrentX = prtMaxX - 1500: XPrt.Print "Référence";
''XPrt.CurrentX = prtMaxX - 500: XPrt.Print "Evé";

'---------------------------------------------------------

XPrt.CurrentY = prtMinY + prtHeaderHeight - XPrt.TextHeight("X")

XPrt.FontSize = 8

End Sub

Public Sub prtBIA_Gafi_01_MVTP0()
Dim K As Integer, xSQL As String
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
For K = 1 To arrYBIAMVT0_Nb
    meYBIAMVT0 = arrYBIAMVT0(K)
    
    If meYBIACPT0.COMPTECOM <> meYBIAMVT0.MOUVEMCOM Then
        blnPrintCompte = True
        xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIACPT0 " _
             & " where COMPTECOM = '" & meYBIAMVT0.MOUVEMCOM & "'"
             
        Set rsSab = cnsab.Execute(xSQL)
        
        V = rsYBIACPT0_GetBuffer(rsSab, meYBIACPT0)
        If Not IsNull(V) Then
            MsgBox Error, vbCritical, "COMPTECOM = '" & meYBIAMVT0.MOUVEMCOM & "'"
            Exit Sub
        End If
    End If
    
    curEur = prtBIA_Gafi_CV
    mDevIso = meYBIAMVT0.COMPTEDEV

    prtBIA_Gafi_Line
Next K


        
End Sub

Public Sub prtBIA_Gafi_02_MVTP0()
Dim K As Long, blnNéant As Boolean

prtBIA_Gafi_02_Z
arrYBIAMVT0(0) = arrYBIAMVT0(1)
arrYBIAMVT0_K1 = 1
For K = 1 To arrYBIAMVT0_Nb

        If arrYBIAMVT0(0).MOUVEMCOM <> arrYBIAMVT0(K).MOUVEMCOM Then
            arrYBIAMVT0_K2 = K - 1
            If curT > paramCompteGafi_Seuil Then prtBIA_Gafi_02_Compte
            
            prtBIA_Gafi_02_Z
            arrYBIAMVT0(0) = arrYBIAMVT0(K)
            arrYBIAMVT0_K1 = K
        End If

        meYBIAMVT0 = arrYBIAMVT0(K)
        curEur = prtBIA_Gafi_CV
        curT = curT + Abs(curEur)
        If Abs(curEur) < paramCompteGafi_curMin Then
            If curEur > 0 Then
                nbDB = nbDB + 1: curDB = curDB + curEur
            Else
                 nbCR = nbCR + 1: curCR = curCR + curEur
           End If
        End If
Next K
arrYBIAMVT0_K2 = arrYBIAMVT0_Nb
If curT > paramCompteGafi_Seuil Then prtBIA_Gafi_02_Compte

End Sub

Public Sub prtBIA_Gafi_02_Compte()
Dim xSQL As String
Dim K As Long

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIACPT0 " _
     & " where COMPTECOM = '" & arrYBIAMVT0(0).MOUVEMCOM & "'"
     
Set rsSab = cnsab.Execute(xSQL)

V = rsYBIACPT0_GetBuffer(rsSab, meYBIACPT0)
If Not IsNull(V) Then
    MsgBox Error, vbCritical, "COMPTECOM = '" & arrYBIAMVT0(0).MOUVEMCOM & "'"
    Exit Sub
End If


If Not blnGAFI_Open Then prtBIA_Gafi_Open
blnPrintCompte = True
For K = arrYBIAMVT0_K1 To arrYBIAMVT0_K2

    meYBIAMVT0 = arrYBIAMVT0(K)
    curEur = prtBIA_Gafi_CV
    mDevIso = meYBIAMVT0.COMPTEDEV
            
    If Abs(curEur) >= paramCompteGafi_curMin Then prtBIA_Gafi_Line
Next K
End Sub

'---------------------------------------------------------
Public Sub prtBIA_Gafi_Line()
'---------------------------------------------------------
Dim lX As Long, lMax As Long
Dim wUnit As String, X As String


If Not blnGAFI_Open Then prtBIA_Gafi_Open

If XPrt.CurrentY + prtlineHeight * 3 > prtMaxY Then
    frmElpPrt.prtNewPage
    prtBIA_Gafi_Form
End If

Nb1 = Nb1 + 1

XPrt.FontSize = 8
XPrt.FontBold = False
'_______________________________________________________________ligne 1-

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinX + 50
XPrt.FontSize = 6
XPrt.CurrentY = XPrt.CurrentY + Height8_6

If blnPrintCompte Then
    blnPrintCompte = False
    Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY, Col4 - 20, XPrt.CurrentY + prtlineHeight, " ", 235)
    XPrt.CurrentY = XPrt.CurrentY + 50
    XPrt.FontBold = True
    XPrt.CurrentX = prtMinX + 50
    XPrt.Print meYBIAMVT0.COMPTEDEV & "  ";
    XPrt.Print meYBIAMVT0.MOUVEMCOM;
    If XPrt.CurrentX < 1500 Then
        XPrt.CurrentX = prtMinX + 1400
    Else
        XPrt.CurrentX = prtMinX + 1800
    End If
    
    XPrt.Print meYBIACPT0.CLIENARSD & " - " & meYBIACPT0.COMPTEINT;
  '20050502jpl  XPrt.CurrentX = Col4 - 400
  '20050502jpl  XPrt.Print meYBIACPT0.CLIENARSD;

    XPrt.FontBold = False
    
    If paramCompteGafi_Etat = "02" Then
        XPrt.FontItalic = True
        X = Format$(curT, "## ### ### ### ### ##0.00")
        XPrt.CurrentX = Col4 - 50 - XPrt.TextWidth(X)
        XPrt.Print X;
        If curDB <> 0 Then
            X = Format$(Abs(curDB), "## ### ### ### ### ##0.00")
            XPrt.CurrentX = Col5 - 50 - XPrt.TextWidth(X)
            XPrt.Print X;
        End If
        If curCR <> 0 Then
            X = Format$(Abs(curCR), "## ### ### ### ### ##0.00")
            XPrt.CurrentX = Col6 - 50 - XPrt.TextWidth(X)
            XPrt.Print X;
        End If
        XPrt.CurrentX = Col6 + 100
        If (nbDB + nbCR) > 0 Then XPrt.Print "Cumul (EUR) de ";
        If nbDB > 0 Then XPrt.Print nbDB & " mvts au débit , ";
        If nbCR > 0 Then XPrt.Print nbCR & " mvts au crédit ";
       
        XPrt.FontItalic = False
    End If

    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    
End If



'XPrt.FontSize = 8
'XPrt.CurrentY = XPrt.CurrentY - Height8_6

XPrt.FontBold = True

X = Format$(Abs(meYBIAMVT0.MOUVEMMON), "## ### ### ### ### ##0.00")
XPrt.CurrentX = IIf(meYBIAMVT0.MOUVEMMON > 0, Col5, Col6) - 50 - XPrt.TextWidth(X)
XPrt.Print X;

XPrt.FontBold = False

If meYBIAMVT0.COMPTEDEV <> "EUR" Then
'    meCV1.DeviseIso = ""
'    meCV1.DeviseN = meYBIAMVT0.Devise
'    meCV1.Montant = meYBIAMVT0.Mt
'    meCV1.OpéAmj = meYBIAMVT0.AmjOpération
'    meCV2.OpéAmj = meYBIAMVT0.AmjOpération
'    Call CV_Transitoire(meCV1, meCV2, meCV3, Conversion)
'    X = Format$(meCV2.Montant, "## ### ### ### ### ##0.00")
    XPrt.FontItalic = True
    X = Format$(curEur, "## ### ### ### ### ##0.00")
    XPrt.CurrentX = Col6 + 1200 - XPrt.TextWidth(X)
    XPrt.Print X;
    XPrt.FontItalic = False
    XPrt.CurrentX = Col4 - 400
    XPrt.Print meYBIAMVT0.COMPTEDEV;

End If

If meYBIAMVT0.MOUVEMDTR <> meYBIAMVT0.MOUVEMDOP Then
    XPrt.CurrentX = prtMaxX - 3000
    Call frmElpPrt.prtTrame(XPrt.CurrentX - 50, XPrt.CurrentY - 50, XPrt.CurrentX + 800, XPrt.CurrentY + prtlineHeight - 50, " ", 235)
    XPrt.FontBold = True
End If
XPrt.CurrentX = prtMaxX - 3000
XPrt.Print dateIBM10(meYBIAMVT0.MOUVEMDOP, False);
XPrt.FontBold = False

XPrt.CurrentX = prtMaxX - 4000
If meYBIAMVT0.MOUVEMDVA <> meYBIAMVT0.MOUVEMDOP Then XPrt.Print dateIBM10(meYBIAMVT0.MOUVEMDVA, False);


XPrt.CurrentX = prtMinX + 1400
XPrt.Print Trim(meYBIAMVT0.LIBELLIB1) & " " & Trim(meYBIAMVT0.LIBELLIB2) & Trim(meYBIAMVT0.LIBELLIB3); ' & " " & Trim(meYBIAMVT0.LIBELLIB4);
wUnit = Table_Ope_Unit(meYBIAMVT0.MOUVEMSER & meYBIAMVT0.MOUVEMSSE & meYBIAMVT0.MOUVEMOPE)
XPrt.CurrentX = prtMaxX - 2000: XPrt.Print wUnit;
XPrt.CurrentX = prtMaxX - 1600: XPrt.Print meYBIAMVT0.MOUVEMOPE;
XPrt.CurrentX = prtMaxX - 1300: XPrt.Print meYBIAMVT0.MOUVEMNUM;
XPrt.CurrentX = prtMaxX - 500: XPrt.Print meYBIAMVT0.MOUVEMEVE;

'If Trim(meYBIAMVT0.LIBELLIB3) <> "" Then
'    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'    XPrt.CurrentX = prtMinX + 1400
'    XPrt.Print Trim(meYBIAMVT0.LIBELLIB3) & " " & Trim(meYBIAMVT0.LIBELLIB4);
'End If
XPrt.CurrentY = XPrt.CurrentY - Height8_6

End Sub



Public Sub prtBIA_Gafi_Monitor(lEtat As String, lcurDB As Currency, lcurCR As Currency, lAMJMin As String, lAMJMax As String, lCLIENARES As String)
Dim X As String
paramCompteGafi_Seuil = lcurCR
paramCompteGafi_curMin = Abs(lcurDB)
paramCompteGafi_Etat = lEtat

mCLIENARES = lCLIENARES
xCLIENARES = "Responsable : "
X = "select * from YBIATAB0 where" _
    & " BIATABID = 'RESPONSABLE'" _
    & " and BIATABK1 ='" & mCLIENARES & "'"
    
Set rsMDB = cnMDB.Execute(X)
If Not rsMDB.EOF Then xCLIENARES = Trim(Mid$(rsMDB("BIATABTXT"), 34, 12)) & " : "

Nb1 = 0
blnGAFI_Open = False
rsYBIACPT0_Init meYBIACPT0
X = Format$(paramCompteGafi_Seuil, "### ### ### ##0.00") & " Eur -  (du : " & dateImp10(lAMJMin) & "  au : " & dateImp10(lAMJMax) & ")"

Select Case paramCompteGafi_Etat
    Case "01":
        prtTitleText = xCLIENARES & lCLIENARES & "  Dispositif de lutte contre le blanchiment : Mvt > " & X
        prtBIA_Gafi_01_MVTP0
    Case "02":
        prtTitleText = xCLIENARES & lCLIENARES & "  dispositif de lutte contre le blanchiment : cumul  > " & X
        prtBIA_Gafi_02_MVTP0
    Case "03":
        prtTitleText = "Dispositif de lutte contre le blanchiment : Surveillance des comptes 'T'" & "  -  (du : " & dateImp10(lAMJMin) & "  au : " & dateImp10(lAMJMax) & ")"
        prtBIA_Gafi_01_MVTP0
    Case "04":
        prtTitleText = "Dispositif de lutte contre le blanchiment : Surveillance des comptes 'PARADIS FISCAUX' : cumul  > " & X
        prtBIA_Gafi_02_MVTP0
    Case "05":
        prtTitleText = xCLIENARES & lCLIENARES & "  Dispositif de lutte contre le blanchiment : Surveillance des comptes 'PTNC' " & "  -  (du : " & dateImp10(lAMJMin) & "  au : " & dateImp10(lAMJMax) & ")"
        prtBIA_Gafi_02_MVTP0
    Case "06":
        prtTitleText = xCLIENARES & lCLIENARES & "  Dispositif de lutte contre le blanchiment : Surveillance des comptes 'X' " & "  -  (du : " & dateImp10(lAMJMin) & "  au : " & dateImp10(lAMJMax) & ")"
        prtBIA_Gafi_02_MVTP0
    Case "07":
        prtTitleText = xCLIENARES & lCLIENARES & "  Dispositif de lutte contre le blanchiment : Surveillance des comptes OBNL " & "  -  (du : " & dateImp10(lAMJMin) & "  au : " & dateImp10(lAMJMax) & ")"
        prtBIA_Gafi_02_MVTP0
    Case "08":
        prtTitleText = xCLIENARES & lCLIENARES & "  Dispositif de lutte contre le blanchiment : Surveillance des PPE " & "  -  (du : " & dateImp10(lAMJMin) & "  au : " & dateImp10(lAMJMax) & ")"
        prtBIA_Gafi_02_MVTP0
End Select
If blnGAFI_Open Then prtBIA_Gafi_Close

End Sub

Public Sub prtBIA_Gafi_02_Z()
curT = 0
nbDB = 0: curDB = 0
nbCR = 0: curCR = 0
arrYBIAMVT0_K1 = 0: arrYBIAMVT0_K2 = 0

End Sub

Public Function prtBIA_Gafi_CV() As Currency
If meYBIAMVT0.COMPTEDEV <> "EUR" Then
    meCV1.DeviseIso = meYBIAMVT0.COMPTEDEV
    meCV1.DeviseN = 0
    meCV1.Montant = meYBIAMVT0.MOUVEMMON
    meCV1.OpéAmj = meYBIAMVT0.MOUVEMDTR + 19000000
    meCV2.OpéAmj = meCV1.OpéAmj
       
    Call CV_Calc("J  ", meCV1, meCV2)
    meCV2.Montant = meCV2.Montant
Else
    meCV2.Montant = meYBIAMVT0.MOUVEMMON
End If
prtBIA_Gafi_CV = meCV2.Montant
End Function
