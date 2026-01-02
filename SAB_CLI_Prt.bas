Attribute VB_Name = "prtSAB_CLI"

'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim X As String, I As Integer, Height8_6 As Integer
Dim V
Dim blnPage As Boolean

Dim meYCLIENA0 As typeYCLIENA0, xYCLIENA0 As typeYCLIENA0
Dim meYADRESS0 As typeYADRESS0, xYADRESS0 As typeYADRESS0
Dim meYSWIBIC0 As typeYSWIBIC0
Dim meMVTP0 As typeMvtP0
Dim xYCLIREF0 As typeYCLIREF0

Dim rsAdo As New ADODB.Recordset

Public Sub prtSAB_CLI_BIC(fgW As MSFlexGrid, cnAdo As ADODB.Connection)
Dim X As String
Dim xSql As String

recYSWIBIC0_Init meYSWIBIC0
Set rsAdo = Nothing


prtTitleText = "SAB : Liste des banques / SWIFT BIC File"
prtFontName = prtFontName_Arial
prtOrientation = vbPRORLandscape 'Portrait '
prtSAB_CLI_Open
prtHeaderHeight = 300
prtSAB_CLI_BIC_Form
XPrt.CurrentY = XPrt.CurrentY - prtlineHeight

XPrt.FontSize = 8
For I = 1 To fgW.Rows - 1
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight + 100
    If XPrt.CurrentY + 500 > prtMaxY Then
        frmElpPrt.prtNewPage
        prtSAB_CLI_BIC_Form
    End If
    
    fgW.Row = I
    fgW.Col = 0
    X = Trim(fgW.Text)
    xSql = "select * from ZCLIENA0 WHERE CLIENACLI = '" & X & "'"
    Set rsAdo = cnAdo.Execute(xSql)
    If rsAdo.EOF Then
        MsgBox xSql, vbCritical, "prtSAB_CLI_BIC:"
    Else

        Call srvYCLIENA0_GetBuffer_ODBC(rsAdo, meYCLIENA0)
   
        xSql = "select * from ZADRESS0 WHERE ADRESSTYP = '4' AND ADRESSNUM = ' " & X & "'"
        Set rsAdo = cnAdo.Execute(xSql)
        If rsAdo.EOF Then
            'MsgBox xSQL, vbCritical, "prtSAB_CLI_BIC:"
            recYSWIBIC0_Init meYSWIBIC0
            meYSWIBIC0.SWIBICBIC = "??????"
        Else
            X = mId$(rsAdo("ADRESSRA1"), 11, 11)
            xSql = "select * from ZSWIBIC0 WHERE SWIBICBIC = '" & X & "'"
            Set rsAdo = cnAdo.Execute(xSql)
            If rsAdo.EOF Then
                'MsgBox xSQL, vbCritical, "prtSAB_CLI_BIC:"
                recYSWIBIC0_Init meYSWIBIC0
                meYSWIBIC0.SWIBICBIC = "??????"
            Else
                Call srvYSWIBIC0_GetBuffer_ODBC(rsAdo, meYSWIBIC0)
            End If
         End If
         
         XPrt.FontSize = 7: XPrt.FontBold = False
         XPrt.CurrentX = prtMinX: XPrt.Print meYCLIENA0.CLIENACLI;
         XPrt.FontBold = True
         If Trim(meYCLIENA0.CLIENASIG) <> mId$(meYSWIBIC0.SWIBICBIC, 1, 8) Then
             XPrt.CurrentX = prtMinX + 800: XPrt.Print "!!";
         End If
         XPrt.CurrentX = prtMinX + 1000: XPrt.Print meYCLIENA0.CLIENASIG;
         XPrt.FontSize = 6: XPrt.FontBold = False
         XPrt.CurrentX = prtMinX + 2500: XPrt.Print meYCLIENA0.CLIENARA1;
         
         XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
         XPrt.FontSize = 5: XPrt.FontBold = False
        ' XPrt.CurrentX = prtMinX: XPrt.Print Trim(xYCLIREF0.CLIREFREF);
         
         XPrt.FontSize = 7: XPrt.FontBold = True
         XPrt.CurrentX = prtMinX + 1000: XPrt.Print meYSWIBIC0.SWIBICBIC;
         XPrt.FontSize = 6: XPrt.FontBold = False
         XPrt.CurrentX = prtMinX + 2500: XPrt.Print meYSWIBIC0.SWIBICINT;
         XPrt.CurrentX = prtMinX + 10000: XPrt.Print meYSWIBIC0.SWIBICVIL;
         XPrt.CurrentX = prtMinX + 13000: XPrt.Print meYSWIBIC0.SWIBICCOM;

    End If
Next I

prtSAB_CLI_Close

End Sub

Public Sub prtSAB_CLI_Adresse(fgW As MSFlexGrid, cnAdo As ADODB.Connection)
Dim xSql As String
Set rsAdo = Nothing

recYSWIBIC0_Init meYSWIBIC0
meYSWIBIC0.Method = "SeekP0"
recMvtP0_Init meMVTP0
meMVTP0.Method = "Seek="


prtTitleText = "SAB : Liste des banques / Adresse"
prtFontName = prtFontName_Arial
prtOrientation = vbPRORPortrait '
prtSAB_CLI_Open
prtHeaderHeight = 300
prtSAB_CLI_BIC_Form
XPrt.CurrentY = XPrt.CurrentY - prtlineHeight

XPrt.FontSize = 8
For I = 1 To fgW.Rows - 1
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight + 100
    If XPrt.CurrentY + 500 > prtMaxY Then
        frmElpPrt.prtNewPage
        prtSAB_CLI_Adresse_Form
    End If
    
    fgW.Row = I
    fgW.Col = 0
    
    xSql = "select * from ZCLIENA0 where CLIENACLI = '" & Trim(fgW.Text) & "'"
    Set rsAdo = cnAdo.Execute(xSql)
    If Not rsAdo.EOF Then
        V = srvYCLIENA0_GetBuffer_ODBC(rsAdo, xYCLIENA0)
        If Not IsNull(V) Then
            MsgBox V, vbCritical, "prtSAB_CLI_Adresse : Lecture ZCLIENT0 : "
            Exit For
        End If
' recherche BIC
'----------------
        xSql = "select * from ZADRESS0 WHERE ADRESSTYP = '4' AND ADRESSNUM = ' " & xYCLIENA0.CLIENACLI & "'"
        Set rsAdo = cnAdo.Execute(xSql)
        If rsAdo.EOF Then
            'MsgBox xSQL, vbCritical, "prtSAB_CLI_BIC:"
            recYSWIBIC0_Init meYSWIBIC0
            meYSWIBIC0.SWIBICBIC = "??????"
        Else
            meYSWIBIC0.SWIBICBIC = mId$(rsAdo("ADRESSRA1"), 11, 11)

        End If
' recherche ADRESSE
'-------------------
        xYADRESS0.ADRESSNUM = xYCLIENA0.CLIENACLI
        Call srvYADRESS0_Client(xYADRESS0, cnAdo)
        
        XPrt.FontSize = 7: XPrt.FontBold = False
        XPrt.CurrentX = prtMinX: XPrt.Print xYCLIENA0.CLIENACLI;
        XPrt.FontBold = True
        XPrt.CurrentX = prtMinX + 1000: XPrt.Print xYCLIENA0.CLIENASIG;
        XPrt.FontSize = 6: XPrt.FontBold = False
        XPrt.CurrentX = prtMinX + 2500: XPrt.Print xYCLIENA0.CLIENARA1;
        XPrt.CurrentX = prtMinX + 6000: XPrt.Print Trim(xYADRESS0.ADRESSRA1) & " " & Trim(xYADRESS0.ADRESSRA2);
        
        XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
       '''' XPrt.CurrentX = prtMinX: XPrt.Print Trim(xYCLIREF0.CLIREFREF);
        XPrt.CurrentX = prtMinX + 1000: XPrt.Print meYSWIBIC0.SWIBICBIC;
        
        XPrt.CurrentX = prtMinX + 6000: XPrt.Print Trim(xYADRESS0.ADRESSAD1);
        If Trim(xYADRESS0.ADRESSAD2) <> "" Then
            XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
            XPrt.CurrentX = prtMinX + 6000: XPrt.Print Trim(xYADRESS0.ADRESSAD2);
        End If
        If Trim(xYADRESS0.ADRESSAD3) <> "" Then
            XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
            XPrt.CurrentX = prtMinX + 6000: XPrt.Print Trim(xYADRESS0.ADRESSAD2);
        End If
        XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
        XPrt.CurrentX = prtMinX + 6000: XPrt.Print Trim(xYADRESS0.ADRESSCOP) & " " & Trim(xYADRESS0.ADRESSVIL);
        XPrt.CurrentX = prtMinX + 8500: XPrt.Print Trim(xYADRESS0.ADRESSPAY);
    End If
Next I

prtSAB_CLI_Close

End Sub


'---------------------------------------------------------
Public Sub prtSAB_CLI_Adresse_Form()
'---------------------------------------------------------
Dim X As String

XPrt.DrawWidth = 1
XPrt.FontSize = 7: XPrt.FontBold = True

Call frmElpPrt.prtTrame(prtMinX, prtMinY, prtMaxX, prtMinY + prtHeaderHeight, "B", 240)
'---------------------------------------------------------
XPrt.CurrentY = prtMinY + 50

XPrt.CurrentX = prtMinX + 50: XPrt.Print "Racine";
XPrt.CurrentX = prtMinX + 1000: XPrt.Print "nom usuel / BIC";
XPrt.CurrentX = prtMinX + 2500: XPrt.Print "intitulé interne ";
XPrt.CurrentX = prtMinX + 6000: XPrt.Print "Adresse";
XPrt.CurrentY = prtMinY + 50 + prtHeaderHeight
XPrt.FontBold = False

End Sub


'---------------------------------------------------------"
Public Sub prtSAB_CLI_BIC_Form()
'---------------------------------------------------------
Dim X As String

XPrt.DrawWidth = 1
XPrt.FontSize = 7: XPrt.FontBold = True

Call frmElpPrt.prtTrame(prtMinX, prtMinY, prtMaxX, prtMinY + prtHeaderHeight, "B", 240)
'---------------------------------------------------------
XPrt.CurrentY = prtMinY + 50

XPrt.CurrentX = prtMinX + 50: XPrt.Print "Racine";
XPrt.CurrentX = prtMinX + 1000: XPrt.Print "nom usuel / BIC";
XPrt.CurrentX = prtMinX + 2500: XPrt.Print "intitulé interne / BIC";
XPrt.CurrentX = prtMinX + 10000: XPrt.Print "ville";
XPrt.CurrentY = prtMinY + 50 + prtHeaderHeight
XPrt.FontBold = False

End Sub





Public Sub prtSAB_CLI_Close()
On Error GoTo prtError


frmElpPrt.prtEndDoc
frmElpPrt.Hide
Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub



Public Sub prtSAB_CLI_Open()
On Error GoTo prtError

Set XPrt = Printer
frmElpPrt.Show vbModeless

Height8_6 = frmElpPrt.prtHeightDelta(8, 6)

blnFiligrane = False
prtPgmName = "prtSAB_CLI"
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


