Attribute VB_Name = "prtSWI_Opération"
Option Explicit
Dim mFct1 As String

Dim X As String, I As Integer, Height8_6 As Integer
'Dim curX As Currency, curX1 As Currency, curX2 As Currency
Dim rsADO_Local As New ADODB.Recordset

Dim blnDossier_Line As Boolean

Dim zYSWIOPE0 As typeYSWIOPE0, mYSWIOPE0 As typeYSWIOPE0, xYSWIOPE0 As typeYSWIOPE0
Dim meZCDODOS0 As typeZCDODOS0
Public Sub prtSWI_Opération_Line(lYSWIOPE0 As typeYSWIOPE0, cnAdo As ADODB.Connection)
Dim X As String
Dim wFontBold_1 As Boolean, wFontBold_2 As Boolean
Dim blnOk  As Boolean
Dim xSql As String
Dim V, wBIC As String
Dim blnDossier As Boolean

On Error GoTo Error_Handler
Set rsADO_Local = Nothing
'Lecture dossier
'-----------------
If lYSWIOPE0.SWISABDOS = 0 Then
    blnDossier = False
Else
    blnDossier = True
    xSql = "select * from " & paramIBM_Library_SAB & ".ZCDODOS0 " _
         & " where CDODOSDOS = " & lYSWIOPE0.SWISABDOS _
         & " and CDODOSCOP = '" & lYSWIOPE0.SWISABCOP & "'"
         
    Set rsADO_Local = cnAdo.Execute(xSql)
    
    If Not rsADO_Local.EOF Then
        V = rsZCDODOS0_GetBuffer(rsADO_Local, meZCDODOS0)
    Else
        MsgBox xSql, vbCritical
        GoTo Error_Handler
    End If
End If

XPrt.FontSize = 7
xYSWIOPE0 = lYSWIOPE0
wBIC = Mid$(xYSWIOPE0.SWIOPEXBIC, 1, 8)
If wBIC <> Mid$(mYSWIOPE0.SWIOPEXBIC, 1, 8) Then
    prtSWI_Opération_NewLine
    Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY - 20, prtMaxX, XPrt.CurrentY + prtlineHeight, " ", 245)
    XPrt.FontBold = True
    XPrt.CurrentX = prtMinX + 50: XPrt.Print wBIC;
    'Lecture Correspondant
    '-----------------
    
    If blnDossier Then
        XPrt.CurrentX = prtMinX + 1000: XPrt.Print prtCDO_RA1(meZCDODOS0.CDODOSCOT, meZCDODOS0.CDODOSCOR);
    End If
    XPrt.FontBold = False
End If

prtSWI_Opération_NewLine
XPrt.CurrentX = prtMinX + 2000


Select Case xYSWIOPE0.SWIOPESTA
    Case "E200": XPrt.FontBold = True: XPrt.Print "Non intégré dans SAB";
                 XPrt.CurrentX = prtMinX + 6000: XPrt.Print xYSWIOPE0.SWIOPEXTRN;
    Case "E300": XPrt.FontBold = True: XPrt.Print "Non comptabilisé";
End Select

XPrt.CurrentX = prtMinX + 4000: XPrt.Print dateImp10(xYSWIOPE0.SWIOPEFLUD);

If blnDossier Then
    XPrt.CurrentX = prtMinX + 5000: XPrt.Print xYSWIOPE0.SWISABCOP & " " & xYSWIOPE0.SWISABDOS;
    
    XPrt.FontBold = True
    XPrt.CurrentX = prtMinX + 12200
    If Trim(meZCDODOS0.CDODOSBAB) = "" Then
        XPrt.Print "D";
    Else
        XPrt.Print "Bq";
    End If
    XPrt.FontBold = False

    'Lecture Donneur d'ordre
    '-----------------
    X = prtCDO_RA1(meZCDODOS0.CDODOSDOR, meZCDODOS0.CDODOSDON)
    XPrt.CurrentX = prtMinX + 6000: XPrt.Print X;
    'Lecture Bénéficiaire
    '-----------------
    X = prtCDO_RA1(meZCDODOS0.CDODOSBER, meZCDODOS0.CDODOSBEN)
    XPrt.CurrentX = prtMinX + 9000: XPrt.Print X;
    
    Select Case meZCDODOS0.CDODOSCON
        Case "C": X = "Conf."
        Case "N": X = "Not."
        Case "P": X = "Part."
        Case Else: X = meZCDODOS0.CDODOSCON
    End Select
    XPrt.CurrentX = prtMinX + 14300: XPrt.Print X;

End If
If xYSWIOPE0.SWIOPEX32A <> 0 Then
    X = Format$(xYSWIOPE0.SWIOPEX32A, "### ### ### ##0.00")
    XPrt.CurrentX = prtMinX + 13500 - XPrt.TextWidth(X)
    XPrt.Print X;
    XPrt.CurrentX = prtMinX + 13700: XPrt.Print xYSWIOPE0.SWIOPEX32D;
End If


If xYSWIOPE0.SWISABCPTD > 0 Then XPrt.CurrentX = prtMinX + 15000: XPrt.Print dateImp10(xYSWIOPE0.SWISABCPTD);

XPrt.FontBold = False
mYSWIOPE0 = lYSWIOPE0
'=======================
Exit Sub

Error_Handler:
XPrt.FontSize = 7
XPrt.FontBold = True
prtSWI_Opération_NewLine
prtSWI_Opération_NewLine
XPrt.CurrentX = prtMinX + 50: XPrt.Print "ERREUR $$$$$$$$$$$$$$$$$$$$$$$$$$";
XPrt.CurrentX = prtMinX + 5000: XPrt.Print lYSWIOPE0.SWISABCOP & " " & lYSWIOPE0.SWISABDOS & ": " & Error;
XPrt.FontBold = False
prtSWI_Opération_NewLine

End Sub

Public Sub prtSWI_Opération_Open(lMsg As String)
On Error GoTo prtError

Set XPrt = Printer
frmElpPrt.Show vbModeless

Height8_6 = frmElpPrt.prtHeightDelta(8, 6)

blnFiligrane = False
 prtOrientation = vbPRORLandscape '
prtPgmName = "prtSWI_Opération"
prtTitleUsr = usrName
prtTitleText = lMsg

prtLineNb = 1
prtlineHeight = 250
prtHeaderHeight = 300


prtFormType = ""
frmElpPrt.prtStdInit

prtFontName = prtFontName_Arial
prtSWI_Opération_Form
XPrt.CurrentY = XPrt.CurrentY - prtlineHeight

srvYSWIOPE0_Init zYSWIOPE0

'mYSWIOPE0.SWIOPEDOS = -1

Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub


Public Sub prtSWI_Opération_Close()
On Error GoTo prtError

XPrt.DrawWidth = 5

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY)

frmElpPrt.prtEndDoc 1000
frmElpPrt.Hide
Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub


Public Sub prtSWI_Opération_NewLine()
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
If XPrt.CurrentY + 300 > prtMaxY Then
    frmElpPrt.prtNewPage
    prtSWI_Opération_Form
End If

End Sub




Public Sub prtSWI_Opération_Form()
Dim wId As String
Dim X As String

XPrt.FontSize = 7
XPrt.FontBold = True
XPrt.DrawWidth = 5

XPrt.CurrentY = prtMinY + 50
Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY - 20, prtMaxX, XPrt.CurrentY + prtlineHeight - 40, " ", 235)

XPrt.CurrentX = prtMinX + 50: XPrt.Print "Banque émettrice";
XPrt.CurrentX = prtMinX + 3900: XPrt.Print "Date réception";
XPrt.CurrentX = prtMinX + 5000: XPrt.Print "N/Référence";
XPrt.CurrentX = prtMinX + 6000: XPrt.Print "Donneur d'ordre ";
XPrt.CurrentX = prtMinX + 9000: XPrt.Print "Bénéficiaire";
XPrt.CurrentX = prtMinX + 11700: XPrt.Print "Bq / Direct";
XPrt.CurrentX = prtMinX + 12800: XPrt.Print "Mt MT700";
XPrt.CurrentX = prtMinX + 13700: XPrt.Print "Devise";
XPrt.CurrentX = prtMinX + 14900: XPrt.Print "Date compta";

XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight

XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY), prtLineColor
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight + 50

blnDossier_Line = False
XPrt.DrawWidth = 1

End Sub









Private Function prtCDO_RA1(lCode As String, lId As String) As String
Dim X As String, xSql As String

If lCode = "T" Then
    xSql = "select * from " & paramIBM_Library_SAB & ".ZCDOTIE0 " _
     & " where CDOTIETIE = '" & lId & "'"
     
        Set rsADO_Local = cnsab.Execute(xSql)
        
        If Not rsADO_Local.EOF Then
            X = rsADO_Local("CDOTIERA1")
        Else
           ' MsgBox xSql, vbCritical, "prtCDO_RA1"
            X = "???" & lId
        End If
Else
    xSql = "select * from " & paramIBM_Library_SAB & ".ZCLIENA0 " _
     & " where CLIENACLI = '" & lId & "'"
     
        Set rsADO_Local = cnsab.Execute(xSql)
        
        If Not rsADO_Local.EOF Then
            X = rsADO_Local("CLIENARA1")
        Else
           ' MsgBox xSql, vbCritical, "prtCDO_RA1"
            X = "???" & lId
        End If
End If
prtCDO_RA1 = X
End Function
