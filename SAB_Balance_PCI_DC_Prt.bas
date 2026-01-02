Attribute VB_Name = "prtSAB_Balance_PCI_DC"
Option Explicit
Dim mFct1 As String

Dim X As String, I As Integer, Height8_6 As Integer
Dim curX As Currency, curX1 As Currency, curX2 As Currency

Dim blnPage As Boolean
Dim xZAUTSYC0  As typeZAUTSYC0

Public Sub prtBalance_PCI_DC_Close_xlsManual(lNb As Long, lErr_PCI As Long, lErr_Sens As Long, currentrow As Long, wsexcel As Excel.Worksheet, ByRef comptageRows As Long, maxRows As Long)
On Error GoTo prtError

        If comptageRows >= maxRows Then
            Call insere_entete_page(wsexcel, "1:3", 3, currentrow)
            comptageRows = 3
            currentrow = currentrow + 3
        End If
        comptageRows = comptageRows + 1
        currentrow = currentrow + 1
        Range("A7:I7").Select
        Selection.Copy
        Range("A" & CStr(currentrow)).Select
        ActiveSheet.Paste
        
        If comptageRows >= maxRows Then
            Call insere_entete_page(wsexcel, "1:3", 3, currentrow)
            comptageRows = 3
            currentrow = currentrow + 3
        End If
        comptageRows = comptageRows + 1
        currentrow = currentrow + 1
        Range("A8:I8").Select
        Selection.Copy
        Range("A" & CStr(currentrow)).Select
        ActiveSheet.Paste
        wsexcel.Cells(currentrow, 4) = "Nombre de comptes traités : " & lNb

        If comptageRows >= maxRows Then
            Call insere_entete_page(wsexcel, "1:3", 3, currentrow)
            comptageRows = 3
            currentrow = currentrow + 3
        End If
        comptageRows = comptageRows + 1
        currentrow = currentrow + 1
        Range("A8:I8").Select
        Selection.Copy
        Range("A" & CStr(currentrow)).Select
        ActiveSheet.Paste
        wsexcel.Cells(currentrow, 4) = "Nombre PCI inconnu : " & lErr_PCI

        If comptageRows >= maxRows Then
            Call insere_entete_page(wsexcel, "1:3", 3, currentrow)
            comptageRows = 3
            currentrow = currentrow + 3
        End If
        comptageRows = comptageRows + 1
        currentrow = currentrow + 1
        Range("A8:I8").Select
        Selection.Copy
        Range("A" & CStr(currentrow)).Select
        ActiveSheet.Paste
        wsexcel.Cells(currentrow, 4) = "Nombre anomalies Db / Cr : " & lErr_Sens

Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")

End Sub

Public Sub prtBalance_PCI_DC_Line_xlsManual(lYBIACPT0 As typeYBIACPT0, lSens As String, currentrow As Long, wsexcel As Excel.Worksheet, ByRef comptageRows As Long, maxRows As Long)
Dim X As String

If comptageRows >= maxRows Then
    Call insere_entete_page(wsexcel, "1:3", 3, currentrow)
    comptageRows = 3
    currentrow = currentrow + 3
End If
currentrow = currentrow + 1
comptageRows = comptageRows + 1
Range("A5:I5").Select
Selection.Copy
Range("A" & CStr(currentrow)).Select
ActiveSheet.Paste
wsexcel.Cells(currentrow, 1) = lYBIACPT0.PLANCOPRO
wsexcel.Cells(currentrow, 2) = lSens
wsexcel.Cells(currentrow, 3) = Trim(lYBIACPT0.COMPTEOBL)
wsexcel.Cells(currentrow, 4) = Trim(lYBIACPT0.COMPTECOM)
wsexcel.Cells(currentrow, 5) = Trim(lYBIACPT0.COMPTEINT)
wsexcel.Cells(currentrow, 6) = dateIBM10(lYBIACPT0.SOLDEDMO, True)

X = Format$(Abs(lYBIACPT0.SOLDECEN), "### ### ### ### ##0.00")
If lYBIACPT0.SOLDECEN > 0 Then
    wsexcel.Cells(currentrow, 7) = X
Else
    wsexcel.Cells(currentrow, 8) = X
End If
wsexcel.Cells(currentrow, 9) = lYBIACPT0.COMPTEDEV

End Sub

Public Sub prtBalance_PCI_DC_Open(lText As String)
On Error GoTo prtError

Set XPrt = Printer
frmElpPrt.Show vbModeless

Height8_6 = frmElpPrt.prtHeightDelta(8, 6)

blnFiligrane = False
 prtOrientation = vbPRORLandscape '
prtPgmName = "prtBalance_PCI_DC"
prtTitleUsr = usrName
prtTitleText = "Comptabilité : Etat des anomalies de sens des comptes / PCI " & lText

prtLineNb = 1
prtlineHeight = 250
prtHeaderHeight = 300


prtFormType = ""
frmElpPrt.prtStdInit

prtFontName = prtFontName_Arial
prtBalance_PCI_DC_Form
XPrt.CurrentY = XPrt.CurrentY - prtlineHeight

Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub


Public Sub prtBalance_PCI_DC_Close(lNb As Long, lErr_PCI As Long, lErr_Sens As Long)
On Error GoTo prtError
XPrt.DrawWidth = 4
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY), prtLineColor

prtBalance_PCI_DC_NewLine
XPrt.CurrentX = prtMinX + 2000
XPrt.Print "Nombre de comptes traités : " & lNb;

prtBalance_PCI_DC_NewLine
XPrt.CurrentX = prtMinX + 2000
XPrt.Print "Nombre PCI inconnu : " & lErr_PCI;


prtBalance_PCI_DC_NewLine
XPrt.CurrentX = prtMinX + 2000
XPrt.Print "Nombre anomalies Db / Cr : " & lErr_Sens;

Call frmElpPrt.prtEndDoc(1000)
frmElpPrt.Hide
Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub

Public Sub prtBalance_PCI_DC_NewLine()
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
If XPrt.CurrentY + 300 > prtMaxY Then
    frmElpPrt.prtNewPage
    prtBalance_PCI_DC_Form
End If

End Sub




Public Sub prtBalance_PCI_DC_Form()
Dim mCurrenty As Long
Dim X As String

XPrt.FontSize = 7
XPrt.FontBold = True
XPrt.DrawWidth = 2

XPrt.CurrentY = prtMinY + 50

XPrt.CurrentX = prtMinX + 50
XPrt.Print "produit";
XPrt.CurrentX = prtMinX + 1300
XPrt.Print "PCI";
XPrt.CurrentX = prtMinX + 800
XPrt.Print "D/C";
XPrt.CurrentX = prtMinX + 2000
XPrt.Print "Compte";
XPrt.CurrentX = prtMinX + 4000
XPrt.Print "Intitulé";
XPrt.CurrentX = prtMinX + 9300
XPrt.Print "Date der mvt";

XPrt.CurrentX = prtMinX + 11800
XPrt.Print "Débit";
XPrt.CurrentX = prtMinX + 13800

XPrt.Print "Crédit";

XPrt.CurrentX = prtMinX + 14900
XPrt.Print "Devise";


'XPrt.FontSize = 8
XPrt.FontBold = False

mCurrenty = XPrt.CurrentY + prtlineHeight
XPrt.Line (prtMinX, mCurrenty)-(prtMaxX, mCurrenty), prtLineColor
XPrt.Line (prtMinX + 1900, prtMinY)-(prtMinX + 1900, prtMaxY), prtLineColor
XPrt.Line (prtMinX + 10400, prtMinY)-(prtMinX + 10400, prtMaxY), prtLineColor
XPrt.Line (prtMinX + 14600, prtMinY)-(prtMinX + 14600, prtMaxY), prtLineColor

XPrt.CurrentY = mCurrenty + 50


End Sub


Public Sub prtBalance_PCI_DC_Line(lYBIACPT0 As typeYBIACPT0, lSens As String)
Dim wId As String
Dim X As String
prtBalance_PCI_DC_NewLine

XPrt.CurrentX = prtMinX + 50
XPrt.Print lYBIACPT0.PLANCOPRO;
XPrt.CurrentX = prtMinX + 1300
XPrt.Print Trim(lYBIACPT0.COMPTEOBL);
XPrt.CurrentX = prtMinX + 800
XPrt.Print lSens;
XPrt.CurrentX = prtMinX + 2000
XPrt.Print lYBIACPT0.COMPTECOM;
XPrt.CurrentX = prtMinX + 4000
XPrt.Print lYBIACPT0.COMPTEINT;
XPrt.CurrentX = prtMinX + 9500
XPrt.Print dateIBM10(lYBIACPT0.SOLDEDMO, True);

X = Format$(Abs(lYBIACPT0.SOLDECEN), "### ### ### ### ##0.00")
If lYBIACPT0.SOLDECEN > 0 Then
    XPrt.CurrentX = prtMinX + 12400 - XPrt.TextWidth(X)
Else
    XPrt.CurrentX = prtMinX + 14400 - XPrt.TextWidth(X)
End If
XPrt.Print X;

XPrt.CurrentX = prtMinX + 15000
XPrt.Print lYBIACPT0.COMPTEDEV;

End Sub




