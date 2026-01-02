Attribute VB_Name = "prtSAB_Balance_Cumul"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------

Type typeSAB_Balance_Cumul
    ID                 As String
    Dev                As String
    Bilan_Nb           As Long
    Bilan_DB           As Currency
    Bilan_CR           As Currency
    HorsBilan_Nb           As Long
    HorsBilan_DB           As Currency
    HorsBilan_CR           As Currency
End Type


Dim xSAB_Balance_Cumul As typeSAB_Balance_Cumul
Public Sub prtSAB_Balance_Cumul_Close()
On Error GoTo prtError
prtSAB_Balance_Cumul_Colonne
frmElpPrt.prtEndDoc
frmElpPrt.Hide
Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub


Public Sub prtSAB_Balance_Cumul_Open()
On Error GoTo prtError

Set XPrt = Printer
frmElpPrt.Show vbModeless

'Height8_6 = frmElpPrt.prtHeightDelta(8, 6)

blnFiligrane = False
prtOrientation = vbPRORPortrait '
prtPgmName = "prtSAB_Balance_Cumul"
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


Public Sub prtSAB_Balance_Cumul_Monitor(larrService_Balance_Cumul() As typeSAB_Balance_Cumul, larrService_Nb As Long, larrDevise_Nb As Long)
Dim iService   As Integer, iDevise As Integer
Dim X As String, blnErreur As Boolean


prtTitleText = "Balance : Service / Devise - au " & dateIBM10(YBIATAB0_DATE_CPT_J, True)

prtFontName = prtFontName_Arial
prtSAB_Balance_Cumul_Open
prtHeaderHeight = 300
prtSAB_Balance_Cumul_Form
XPrt.CurrentY = XPrt.CurrentY - prtlineHeight
'==========================================================================================

For iService = 1 To larrService_Nb
    If XPrt.CurrentY + 600 > prtMaxY Then XPrt.CurrentY = prtMaxY

    prtSAB_Balance_Cumul_NewLine
'    If iService > 1 Then
'        XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY)
'        XPrt.CurrentY = XPrt.CurrentY + 20
'    End If
    Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY, prtMaxX - 20, XPrt.CurrentY + prtlineHeight, " ", 240)
    XPrt.CurrentY = XPrt.CurrentY + 50
    XPrt.CurrentX = prtMinX + 20
    XPrt.FontBold = True
    XPrt.Print larrService_Balance_Cumul(iService, 0).ID;
    XPrt.FontBold = False
    For iDevise = 1 To larrDevise_Nb
        xSAB_Balance_Cumul = larrService_Balance_Cumul(iService, iDevise)
        If xSAB_Balance_Cumul.Bilan_Nb > 0 Or xSAB_Balance_Cumul.HorsBilan_Nb > 0 Then
            prtSAB_Balance_Cumul_Line
            larrService_Balance_Cumul(0, iDevise).Bilan_Nb = larrService_Balance_Cumul(0, iDevise).Bilan_Nb + xSAB_Balance_Cumul.Bilan_Nb
            larrService_Balance_Cumul(0, iDevise).Bilan_DB = larrService_Balance_Cumul(0, iDevise).Bilan_DB + xSAB_Balance_Cumul.Bilan_DB
            larrService_Balance_Cumul(0, iDevise).Bilan_CR = larrService_Balance_Cumul(0, iDevise).Bilan_CR + xSAB_Balance_Cumul.Bilan_CR
            larrService_Balance_Cumul(0, iDevise).HorsBilan_Nb = larrService_Balance_Cumul(0, iDevise).HorsBilan_Nb + xSAB_Balance_Cumul.HorsBilan_Nb
            larrService_Balance_Cumul(0, iDevise).HorsBilan_DB = larrService_Balance_Cumul(0, iDevise).HorsBilan_DB + xSAB_Balance_Cumul.HorsBilan_DB
            larrService_Balance_Cumul(0, iDevise).HorsBilan_CR = larrService_Balance_Cumul(0, iDevise).HorsBilan_CR + xSAB_Balance_Cumul.HorsBilan_CR
        End If
    Next iDevise

Next iService
'==========================================================================================
XPrt.DrawWidth = 10
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight + 20
XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY)
    Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY, prtMaxX - 20, XPrt.CurrentY + prtlineHeight, " ", 240)
    XPrt.CurrentY = XPrt.CurrentY + 50
    XPrt.CurrentX = prtMinX + 20
    XPrt.FontBold = True
    XPrt.Print "Total";
blnErreur = False
For iDevise = 1 To larrDevise_Nb
    xSAB_Balance_Cumul = larrService_Balance_Cumul(0, iDevise)
    If xSAB_Balance_Cumul.Bilan_Nb > 0 Or xSAB_Balance_Cumul.HorsBilan_Nb > 0 Then
        prtSAB_Balance_Cumul_Line
        If larrService_Balance_Cumul(0, iDevise).Bilan_DB <> larrService_Balance_Cumul(0, iDevise).Bilan_CR Then
            blnErreur = True
            XPrt.CurrentX = prtMinX + 1050
            XPrt.Print "???";
        End If
        If larrService_Balance_Cumul(0, iDevise).HorsBilan_DB <> larrService_Balance_Cumul(0, iDevise).HorsBilan_CR Then
            blnErreur = True
            XPrt.CurrentX = prtMinX + 6150
            XPrt.Print "???";
        End If
      
    End If
Next iDevise
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight + 20
XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY)
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight + 20
Call frmElpPrt.prtTrame(prtMinX + 1200 + 20, XPrt.CurrentY, prtMinX + 6200 - 20, XPrt.CurrentY + prtlineHeight, " ", 240)
If blnErreur Then
    X = "??? ERREUR ???"
Else
    X = "Balance équilibrée"
End If
XPrt.FontSize = 12
frmElpPrt.prtCentré prtMinX + 3500, X
'==========================================================================================

prtSAB_Balance_Cumul_Close
    
End Sub

Public Sub prtSAB_Balance_Cumul_Form()
Dim wId As String
Dim X As String

XPrt.FontSize = 7
XPrt.FontBold = True
XPrt.DrawWidth = 2
XPrt.Line (prtMaxX, prtMinY)-(prtMaxX, prtMaxY)

XPrt.CurrentY = prtMinY + 50

XPrt.CurrentX = prtMinX
XPrt.Print "Devise";
XPrt.CurrentX = prtMinX + 1500
XPrt.Print "Bilan Nb";

XPrt.CurrentX = prtMinX + 3200
XPrt.Print "Bilan Débit";
XPrt.CurrentX = prtMinX + 5200
XPrt.Print "Bilan Crédit";
XPrt.CurrentX = prtMinX + 6600
XPrt.Print "HB Nb";

XPrt.CurrentX = prtMinX + 8300
XPrt.Print "HB Débit";
XPrt.CurrentX = prtMinX + 10200
XPrt.Print "HB crédit";

'XPrt.FontSize = 8
XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY)
XPrt.CurrentY = XPrt.CurrentY + 50


End Sub

Public Sub prtSAB_Balance_Cumul_Colonne()
Dim wId As String
Dim X As String

XPrt.DrawWidth = 2
XPrt.Line (prtMinX + 6200, prtMinY)-(prtMinX + 6200, prtMaxY)
XPrt.Line (prtMinX + 1200, prtMinY)-(prtMinX + 1200, prtMaxY)


End Sub


Public Sub prtSAB_Balance_Cumul_Line()
Dim X As String

prtSAB_Balance_Cumul_NewLine


XPrt.CurrentX = prtMinX
XPrt.Print xSAB_Balance_Cumul.Dev;

If xSAB_Balance_Cumul.Bilan_Nb <> 0 Then
    X = Format$(xSAB_Balance_Cumul.Bilan_Nb, "### ### ### ###")
    XPrt.CurrentX = prtMinX + 2000 - XPrt.TextWidth(X)
    XPrt.Print X;
End If

If xSAB_Balance_Cumul.Bilan_DB <> 0 Then
    X = Format$(xSAB_Balance_Cumul.Bilan_DB, "### ### ### ###.00")
    XPrt.CurrentX = prtMinX + 4000 - XPrt.TextWidth(X)
    XPrt.Print X;
End If

If xSAB_Balance_Cumul.Bilan_CR <> 0 Then
    X = Format$(xSAB_Balance_Cumul.Bilan_CR, "### ### ### ###.00")
    XPrt.CurrentX = prtMinX + 6000 - XPrt.TextWidth(X)
    XPrt.Print X;
End If

If xSAB_Balance_Cumul.HorsBilan_Nb <> 0 Then
    X = Format$(xSAB_Balance_Cumul.HorsBilan_Nb, "### ### ### ###")
    XPrt.CurrentX = prtMinX + 7000 - XPrt.TextWidth(X)
    XPrt.Print X;
End If


If xSAB_Balance_Cumul.HorsBilan_DB <> 0 Then
    X = Format$(xSAB_Balance_Cumul.HorsBilan_DB, "### ### ### ###.00")
    XPrt.CurrentX = prtMaxX - 2000 - XPrt.TextWidth(X)
    XPrt.Print X;
End If

If xSAB_Balance_Cumul.HorsBilan_CR <> 0 Then
    X = Format$(xSAB_Balance_Cumul.HorsBilan_CR, "### ### ### ###.00")
    XPrt.CurrentX = prtMaxX - 50 - XPrt.TextWidth(X)
    XPrt.Print X;
End If

End Sub





Public Sub prtSAB_Balance_Cumul_NewLine()
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
If XPrt.CurrentY + 300 > prtMaxY Then
    prtSAB_Balance_Cumul_Colonne
    frmElpPrt.prtNewPage
    prtSAB_Balance_Cumul_Form
End If

End Sub

