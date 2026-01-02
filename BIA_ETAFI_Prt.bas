Attribute VB_Name = "prtBIA_ETAFI"
Option Explicit
'!!!!!!!! pour export vers CEGID _ ETAFI : DEBIT signé + , CREDIT signé -

Dim meYETAFI0 As typeYETAFI0, zYETAFI0 As typeYETAFI0
Dim devYETAFI0(100) As typeYETAFI0, devYETAFI0_Nb
Dim cnAdo As New ADODB.Connection
Dim rsADO_ETAFI0 As New ADODB.Recordset

Dim curX As Currency
Dim xSQL As String, V

    Dim arrSD1_Nb  As Integer
    Dim arrSD1_C1A5(100) As typeYETAFI0
    Dim arrSD1_C6A8(100) As typeYETAFI0
    Dim arrSD1_C9(100) As typeYETAFI0

    Dim arrSD0_Nb  As Integer
    Dim arrSD0_C1A5(100) As typeYETAFI0
    Dim arrSD0_C6A8(100) As typeYETAFI0
    Dim arrSD0_C9(100) As typeYETAFI0

Dim Height8_6 As Integer

Public Sub prtBIA_ETAFI_Export()
Dim X250 As String * 250
Dim curX As Currency

 curX = meYETAFI0.ETAFISD1X - (meYETAFI0.ETAFISD0X + meYETAFI0.ETAFIDBX + meYETAFI0.ETAFICRX)
 If curX > 0 Then
     meYETAFI0.ETAFIDBX = meYETAFI0.ETAFIDBX + curX
 Else
     meYETAFI0.ETAFICRX = meYETAFI0.ETAFICRX + curX
End If

If meYETAFI0.ETAFISD1X <> 0 Or meYETAFI0.ETAFISD0X <> 0 _
Or meYETAFI0.ETAFIDBX <> 0 Or meYETAFI0.ETAFICRX <> 0 Then
    X250 = ""
    Mid$(X250, 1, 20) = Format(meYETAFI0.ETAFICOM, "00000000000")
    Mid$(X250, 21, 1) = ";"
    Mid$(X250, 22, 10) = meYETAFI0.ETAFIOBL
    Mid$(X250, 32, 1) = ";"
    Mid$(X250, 33, 32) = meYETAFI0.ETAFIINT
    Mid$(X250, 65, 1) = ";"
           
    Mid$(X250, 66, 19) = cur_19V(meYETAFI0.ETAFISD0X)
    Mid$(X250, 85, 1) = ";"
    Mid$(X250, 86, 19) = cur_19V(meYETAFI0.ETAFIDBX)
    Mid$(X250, 105, 1) = ";"
    Mid$(X250, 106, 19) = cur_19V(meYETAFI0.ETAFICRX)
    Mid$(X250, 125, 1) = ";"
    Mid$(X250, 126, 19) = cur_19V(meYETAFI0.ETAFISD1X)
    Print #2, X250

End If

End Sub


Public Sub prtBIA_ETAFI_Monitor(lFileName As String)
Dim xIn As String, X As String
Dim seq As Long
Dim blnOk As Boolean
On Error GoTo Error_Handler

prtBIA_ETAFI_Init
Call FEU_ROUGE
Open "C:\TEMP\ETAFI.txt" For Output As #2

cnAdo.Open paramODBC_DSN_SAB
Set rsADO_ETAFI0 = Nothing
Set rsADO_ETAFI0 = Nothing
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YETAFI0"
Set rsADO_ETAFI0 = cnAdo.Execute(xSQL)

Do While Not rsADO_ETAFI0.EOF
    V = rsYETAFI0_GetBuffer(rsADO_ETAFI0, meYETAFI0)
    
    If Not IsNull(V) Then
        MsgBox V, vbCritical, "prtBIA_ETAFI_Monitor : rsADO_ETAFI0"
    Else
        prtBIA_ETAFI_Dev_Classe_Cumul
        prtBIA_ETAFI_Export
    End If
    rsADO_ETAFI0.MoveNext
Loop

prtTitleText = "ETAFI : récapitulatif Bilan / Hors Bilan par devise DEBUT d'EXERCICE"
prtBIA_ETAFI_Open
prtBIA_ETAFI_Form
prtBIA_ETAFI_Dev_Classe_Prt_SD0

prtTitleText = "ETAFI :récapitulatif Bilan / Hors Bilan par devise FIN d'EXERCICE"
frmElpPrt.prtNewPage
prtBIA_ETAFI_Form
prtBIA_ETAFI_Dev_Classe_Prt_SD1  'vite fait par duplication (utilisé 1 fois par an !) jpl

prtBIA_ETAFI_Close
Call FEU_VERT

GoTo Exit_sub

Error_Handler:

Shell_MsgBox "me.cmdYBIAMVT0_Import#  & error ", vbCritical, "prtBIA_ETAFI_Export", False
Exit_sub:
Close

cnAdo.Close
Set cnAdo = Nothing


End Sub

'---------------------------------------------------------
Public Sub prtBIA_ETAFI_Form()
'---------------------------------------------------------
Dim X As String

XPrt.DrawWidth = 1
XPrt.FontSize = 7: XPrt.FontBold = True

Call frmElpPrt.prtTrame(prtMinX, prtMinY, prtMaxX, prtMinY + prtHeaderHeight, "B", 240)
XPrt.Line (prtMinX + 12000, prtMinY)-(prtMinX + 12000, prtMaxY), prtLineColor
XPrt.Line (prtMinX + 7500, prtMinY)-(prtMinX + 7500, prtMaxY), prtLineColor
XPrt.Line (prtMaxX, prtMinY)-(prtMaxX, prtMaxY), prtLineColor
'---------------------------------------------------------
XPrt.CurrentY = prtMinY + 50

XPrt.CurrentX = prtMinX + 100: XPrt.Print " ";
'XPrt.CurrentX = prtMinX + 400: XPrt.Print "Compte ";
'XPrt.CurrentX = prtMinX + 2000: XPrt.Print "Intitulé";
'XPrt.CurrentX = prtMinX + 10500: XPrt.Print "Devise";
XPrt.CurrentX = prtMinX + 8900: XPrt.Print "Débit";
XPrt.CurrentX = prtMinX + 10900: XPrt.Print "Crédit";
XPrt.CurrentX = prtMinX + 13100: XPrt.Print "Débit";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 6
XPrt.FontBold = True
XPrt.CurrentX = prtMinX + 14100: XPrt.Print "/ EUR /";
XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 8

XPrt.CurrentX = prtMinX + 15100: XPrt.Print "Crédit";

XPrt.CurrentY = prtMinY + prtHeaderHeight + 100
XPrt.FontBold = False

End Sub

Public Sub prtBIA_ETAFI_Open()
On Error GoTo prtError

Set XPrt = Printer
frmElpPrt.Show vbModeless

Height8_6 = frmElpPrt.prtHeightDelta(8, 6)

blnFiligrane = False
prtOrientation = vbPRORLandscape '
prtPgmName = "prtBIA_ETAFI"
prtTitleUsr = usrName

prtLineNb = 1
prtlineHeight = 250
prtHeaderHeight = 300


prtFormType = ""
frmElpPrt.prtStdInit

Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub


Public Sub prtBIA_ETAFI_Close()
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




Public Sub prtBIA_ETAFI_Init()
Dim xSQL As String

zYETAFI0.ETAFICOM = ""
zYETAFI0.ETAFIOBL = ""
zYETAFI0.ETAFIINT = ""
zYETAFI0.ETAFISD0X = 0
zYETAFI0.ETAFIDBX = 0
zYETAFI0.ETAFICRX = 0
zYETAFI0.ETAFISD1X = 0
zYETAFI0.ETAFISD0 = 0
zYETAFI0.ETAFIDB = 0
zYETAFI0.ETAFICR = 0
zYETAFI0.ETAFISD1 = 0
zYETAFI0.ETAFIDBNB = 0
zYETAFI0.ETAFICRNB = 0
zYETAFI0.ETAFISTA = ""
zYETAFI0.ETAFIDEV = ""

arrSD1_Nb = 0
xSQL = "select * from YBIATAB0 " _
    & " where BIATABID = 'DEVISE'" _
    & " and BIATABK1 = 'ISO'"
    
Set rsMDB = cnMDB.Execute(xSQL)
Do While Not rsMDB.EOF

        arrSD1_Nb = arrSD1_Nb + 1
        arrSD1_C1A5(arrSD1_Nb) = zYETAFI0
        arrSD1_C1A5(arrSD1_Nb).ETAFIDEV = rsMDB("BIATABK2")
        arrSD1_C6A8(arrSD1_Nb) = zYETAFI0
        arrSD1_C6A8(arrSD1_Nb).ETAFIDEV = arrSD1_C1A5(arrSD1_Nb).ETAFIDEV
        arrSD1_C9(arrSD1_Nb) = zYETAFI0
        arrSD1_C9(arrSD1_Nb).ETAFIDEV = arrSD1_C1A5(arrSD1_Nb).ETAFIDEV
        
        arrSD0_Nb = arrSD0_Nb + 1
        arrSD0_C1A5(arrSD0_Nb) = arrSD1_C1A5(arrSD1_Nb)
        arrSD0_C6A8(arrSD0_Nb) = arrSD1_C6A8(arrSD1_Nb)
        arrSD0_C9(arrSD0_Nb) = arrSD1_C9(arrSD1_Nb)
    rsMDB.MoveNext
Loop

' Code devise non trouvé
arrSD1_C1A5(arrSD1_Nb + 1) = zYETAFI0
arrSD1_C6A8(arrSD1_Nb + 1) = zYETAFI0
arrSD1_C9(arrSD1_Nb + 1) = zYETAFI0

End Sub
Public Sub prtBIA_ETAFI_Dev_Classe_Cumul()
Dim K As Integer

For K = 1 To arrSD1_Nb
    If meYETAFI0.ETAFIDEV = arrSD1_C1A5(K).ETAFIDEV Then Exit For
Next K
    
Select Case Mid$(meYETAFI0.ETAFIOBL, 1, 1)
    Case Is <= 5
        If meYETAFI0.ETAFISD1 < 0 Then
             arrSD1_C1A5(K).ETAFICR = arrSD1_C1A5(K).ETAFICR + meYETAFI0.ETAFISD1
             arrSD1_C1A5(K).ETAFICRX = arrSD1_C1A5(K).ETAFICRX + meYETAFI0.ETAFISD1X
        Else
             arrSD1_C1A5(K).ETAFIDB = arrSD1_C1A5(K).ETAFIDB + meYETAFI0.ETAFISD1
             arrSD1_C1A5(K).ETAFIDBX = arrSD1_C1A5(K).ETAFIDBX + meYETAFI0.ETAFISD1X
        End If
        
         If meYETAFI0.ETAFISD0 < 0 Then
             arrSD0_C1A5(K).ETAFICR = arrSD0_C1A5(K).ETAFICR + meYETAFI0.ETAFISD0
             arrSD0_C1A5(K).ETAFICRX = arrSD0_C1A5(K).ETAFICRX + meYETAFI0.ETAFISD0X
        Else
             arrSD0_C1A5(K).ETAFIDB = arrSD0_C1A5(K).ETAFIDB + meYETAFI0.ETAFISD0
             arrSD0_C1A5(K).ETAFIDBX = arrSD0_C1A5(K).ETAFIDBX + meYETAFI0.ETAFISD0X
        End If
       
    Case Is <= 8
        If meYETAFI0.ETAFISD1 < 0 Then
             arrSD1_C6A8(K).ETAFICR = arrSD1_C6A8(K).ETAFICR + meYETAFI0.ETAFISD1
             arrSD1_C6A8(K).ETAFICRX = arrSD1_C6A8(K).ETAFICRX + meYETAFI0.ETAFISD1X
        Else
             arrSD1_C6A8(K).ETAFIDB = arrSD1_C6A8(K).ETAFIDB + meYETAFI0.ETAFISD1
             arrSD1_C6A8(K).ETAFIDBX = arrSD1_C6A8(K).ETAFIDBX + meYETAFI0.ETAFISD1X
        End If
        
          If meYETAFI0.ETAFISD0 < 0 Then
             arrSD0_C6A8(K).ETAFICR = arrSD0_C6A8(K).ETAFICR + meYETAFI0.ETAFISD0
             arrSD0_C6A8(K).ETAFICRX = arrSD0_C6A8(K).ETAFICRX + meYETAFI0.ETAFISD0X
        Else
             arrSD0_C6A8(K).ETAFIDB = arrSD0_C6A8(K).ETAFIDB + meYETAFI0.ETAFISD0
             arrSD0_C6A8(K).ETAFIDBX = arrSD0_C6A8(K).ETAFIDBX + meYETAFI0.ETAFISD0X
        End If
  Case Else
         If meYETAFI0.ETAFISD1 < 0 Then
             arrSD1_C9(K).ETAFICR = arrSD1_C9(K).ETAFICR + meYETAFI0.ETAFISD1
              arrSD1_C9(K).ETAFICRX = arrSD1_C9(K).ETAFICRX + meYETAFI0.ETAFISD1X
       Else
             arrSD1_C9(K).ETAFIDB = arrSD1_C9(K).ETAFIDB + meYETAFI0.ETAFISD1
             arrSD1_C9(K).ETAFIDBX = arrSD1_C9(K).ETAFIDBX + meYETAFI0.ETAFISD1X
        End If
        
        If meYETAFI0.ETAFISD0 < 0 Then
             arrSD0_C9(K).ETAFICR = arrSD0_C9(K).ETAFICR + meYETAFI0.ETAFISD0
              arrSD0_C9(K).ETAFICRX = arrSD0_C9(K).ETAFICRX + meYETAFI0.ETAFISD0X
       Else
             arrSD0_C9(K).ETAFIDB = arrSD0_C9(K).ETAFIDB + meYETAFI0.ETAFISD0
             arrSD0_C9(K).ETAFIDBX = arrSD0_C9(K).ETAFIDBX + meYETAFI0.ETAFISD0X
        End If

End Select
End Sub

Public Sub prtBIA_ETAFI_Dev_Classe_Prt_SD1()
Dim curX As Currency, X As String
Dim blnSaut As Boolean
Dim wNb As Integer, K As Integer

wNb = arrSD1_Nb + 2
arrSD1_C1A5(wNb) = zYETAFI0
arrSD1_C6A8(wNb) = zYETAFI0
arrSD1_C9(wNb) = zYETAFI0

arrSD1_C1A5(arrSD1_Nb + 1).ETAFIDEV = "???"
arrSD1_C6A8(arrSD1_Nb + 1).ETAFIDEV = "???"
arrSD1_C9(arrSD1_Nb + 1).ETAFIDEV = "???"
arrSD1_C1A5(arrSD1_Nb + 2).ETAFIDEV = "***"
arrSD1_C6A8(arrSD1_Nb + 2).ETAFIDEV = "***"
arrSD1_C9(arrSD1_Nb + 2).ETAFIDEV = "***"
XPrt.FontSize = 6

For K = 1 To wNb
    blnSaut = False
    If arrSD1_C1A5(K).ETAFIDBX <> 0 Or arrSD1_C1A5(K).ETAFICRX <> 0 Then
        XPrt.FontBold = True
        blnSaut = True
        prtBIA_ETAFI_NewLine
        XPrt.CurrentX = prtMinX + 5000: XPrt.Print arrSD1_C1A5(K).ETAFIDEV & " - Total classes 1 à 5 ";
        XPrt.CurrentX = prtMinX + 11600: XPrt.Print arrSD1_C1A5(K).ETAFIDEV;
    
        curX = Abs(arrSD1_C1A5(K).ETAFIDB)
        If curX <> 0 Then
            X = Format$(curX, "### ### ### ### ##0.00")
            XPrt.CurrentX = prtMinX + 9400 - XPrt.TextWidth(X)
            XPrt.Print X;
        End If
        
        curX = Abs(arrSD1_C1A5(K).ETAFICR)
        If curX <> 0 Then
            X = Format$(curX, "### ### ### ### ##0.00")
            XPrt.CurrentX = prtMinX + 11400 - XPrt.TextWidth(X)
            XPrt.Print X;
        End If
        
        curX = Abs(arrSD1_C1A5(K).ETAFIDBX)
        If curX <> 0 Then
            X = Format$(curX, "### ### ### ### ##0.00")
            XPrt.CurrentX = prtMaxX - 2100 - XPrt.TextWidth(X)
            XPrt.Print X;
        End If
        
        curX = Abs(arrSD1_C1A5(K).ETAFICRX)
        If curX <> 0 Then
            X = Format$(curX, "### ### ### ### ##0.00")
            XPrt.CurrentX = prtMaxX - 100 - XPrt.TextWidth(X)
            XPrt.Print X;
        End If
    End If

    If arrSD1_C6A8(K).ETAFIDBX <> 0 Or arrSD1_C6A8(K).ETAFICRX <> 0 Then
        XPrt.FontBold = True
        blnSaut = True
        prtBIA_ETAFI_NewLine
        XPrt.CurrentX = prtMinX + 5000: XPrt.Print arrSD1_C6A8(K).ETAFIDEV & " - Total classes 6 à 8 ";
        XPrt.CurrentX = prtMinX + 11600: XPrt.Print arrSD1_C6A8(K).ETAFIDEV;
    
        curX = Abs(arrSD1_C6A8(K).ETAFIDB)
        If curX <> 0 Then
            X = Format$(curX, "### ### ### ### ##0.00")
            XPrt.CurrentX = prtMinX + 9400 - XPrt.TextWidth(X)
            XPrt.Print X;
        End If
        
        curX = Abs(arrSD1_C6A8(K).ETAFICR)
        If curX <> 0 Then
            X = Format$(curX, "### ### ### ### ##0.00")
            XPrt.CurrentX = prtMinX + 11400 - XPrt.TextWidth(X)
            XPrt.Print X;
        End If
        
        curX = Abs(arrSD1_C6A8(K).ETAFIDBX)
        If curX <> 0 Then
            X = Format$(curX, "### ### ### ### ##0.00")
            XPrt.CurrentX = prtMaxX - 2100 - XPrt.TextWidth(X)
            XPrt.Print X;
        End If
        
        curX = Abs(arrSD1_C6A8(K).ETAFICRX)
        If curX <> 0 Then
            X = Format$(curX, "### ### ### ### ##0.00")
            XPrt.CurrentX = prtMaxX - 100 - XPrt.TextWidth(X)
            XPrt.Print X;
        End If
    End If
    If arrSD1_C9(K).ETAFIDBX <> 0 Or arrSD1_C9(K).ETAFICRX <> 0 Then
        XPrt.FontBold = False
        blnSaut = True
        prtBIA_ETAFI_NewLine
        XPrt.CurrentX = prtMinX + 5000: XPrt.Print arrSD1_C9(K).ETAFIDEV & " - Total classe 9 ";
        XPrt.CurrentX = prtMinX + 11600: XPrt.Print arrSD1_C9(K).ETAFIDEV;
    
        curX = Abs(arrSD1_C9(K).ETAFIDB)
        If curX <> 0 Then
            X = Format$(curX, "### ### ### ### ##0.00")
            XPrt.CurrentX = prtMinX + 9400 - XPrt.TextWidth(X)
            XPrt.Print X;
        End If
        
        curX = Abs(arrSD1_C9(K).ETAFICR)
        If curX <> 0 Then
            X = Format$(curX, "### ### ### ### ##0.00")
            XPrt.CurrentX = prtMinX + 11400 - XPrt.TextWidth(X)
            XPrt.Print X;
        End If
        
        curX = Abs(arrSD1_C9(K).ETAFIDBX)
        If curX <> 0 Then
            X = Format$(curX, "### ### ### ### ##0.00")
            XPrt.CurrentX = prtMaxX - 2100 - XPrt.TextWidth(X)
            XPrt.Print X;
        End If
        
        curX = Abs(arrSD1_C9(K).ETAFICRX)
        If curX <> 0 Then
            X = Format$(curX, "### ### ### ### ##0.00")
            XPrt.CurrentX = prtMaxX - 100 - XPrt.TextWidth(X)
            XPrt.Print X;
        End If
    End If
    
    curX = arrSD1_C1A5(K).ETAFIDB + arrSD1_C1A5(K).ETAFICR _
            + arrSD1_C6A8(K).ETAFIDB + arrSD1_C6A8(K).ETAFICR
    If curX <> 0 Then
        prtBIA_ETAFI_NewLine
        XPrt.CurrentX = prtMinX + 2000: XPrt.Print arrSD1_C1A5(K).ETAFIDEV & "   ?????????? ERREUR BILAN ";
        X = Format$(curX, "### ### ### ### ##0.00")
        XPrt.CurrentX = prtMinX + 1900 - XPrt.TextWidth(X)
        XPrt.Print X;

    End If
    
    
    curX = arrSD1_C9(K).ETAFIDB + arrSD1_C9(K).ETAFICR
    If curX <> 0 Then
        prtBIA_ETAFI_NewLine
        XPrt.CurrentX = prtMinX + 2000: XPrt.Print arrSD1_C9(K).ETAFIDEV & "   ?????????? ERREUR HORS-BILAN ";
        X = Format$(curX, "### ### ### ### ##0.00")
        XPrt.CurrentX = prtMinX + 1900 - XPrt.TextWidth(X)
        XPrt.Print X;
    End If
   
   arrSD1_C1A5(wNb).ETAFIDBX = arrSD1_C1A5(wNb).ETAFIDBX + arrSD1_C1A5(K).ETAFIDBX
   arrSD1_C1A5(wNb).ETAFICRX = arrSD1_C1A5(wNb).ETAFICRX + arrSD1_C1A5(K).ETAFICRX
   arrSD1_C6A8(wNb).ETAFIDBX = arrSD1_C6A8(wNb).ETAFIDBX + arrSD1_C6A8(K).ETAFIDBX
   arrSD1_C6A8(wNb).ETAFICRX = arrSD1_C6A8(wNb).ETAFICRX + arrSD1_C6A8(K).ETAFICRX
   arrSD1_C9(wNb).ETAFIDBX = arrSD1_C9(wNb).ETAFIDBX + arrSD1_C9(K).ETAFIDBX
   arrSD1_C9(wNb).ETAFICRX = arrSD1_C9(wNb).ETAFICRX + arrSD1_C9(K).ETAFICRX
   
   If blnSaut Then prtBIA_ETAFI_NewLine

Next K


End Sub

Public Sub prtBIA_ETAFI_Dev_Classe_Prt_SD0()
Dim curX As Currency, X As String
Dim blnSaut As Boolean
Dim wNb As Integer, K As Integer

wNb = arrSD0_Nb + 2
arrSD0_C1A5(wNb) = zYETAFI0
arrSD0_C6A8(wNb) = zYETAFI0
arrSD0_C9(wNb) = zYETAFI0

arrSD0_C1A5(arrSD0_Nb + 1).ETAFIDEV = "???"
arrSD0_C6A8(arrSD0_Nb + 1).ETAFIDEV = "???"
arrSD0_C9(arrSD0_Nb + 1).ETAFIDEV = "???"
arrSD0_C1A5(arrSD0_Nb + 2).ETAFIDEV = "***"
arrSD0_C6A8(arrSD0_Nb + 2).ETAFIDEV = "***"
arrSD0_C9(arrSD0_Nb + 2).ETAFIDEV = "***"
XPrt.FontSize = 6

For K = 1 To wNb
    blnSaut = False
    If arrSD0_C1A5(K).ETAFIDBX <> 0 Or arrSD0_C1A5(K).ETAFICRX <> 0 Then
        XPrt.FontBold = True
        blnSaut = True
        prtBIA_ETAFI_NewLine
        XPrt.CurrentX = prtMinX + 5000: XPrt.Print arrSD0_C1A5(K).ETAFIDEV & " - Total classes 1 à 5 ";
        XPrt.CurrentX = prtMinX + 11600: XPrt.Print arrSD0_C1A5(K).ETAFIDEV;
    
        curX = Abs(arrSD0_C1A5(K).ETAFIDB)
        If curX <> 0 Then
            X = Format$(curX, "### ### ### ### ##0.00")
            XPrt.CurrentX = prtMinX + 9400 - XPrt.TextWidth(X)
            XPrt.Print X;
        End If
        
        curX = Abs(arrSD0_C1A5(K).ETAFICR)
        If curX <> 0 Then
            X = Format$(curX, "### ### ### ### ##0.00")
            XPrt.CurrentX = prtMinX + 11400 - XPrt.TextWidth(X)
            XPrt.Print X;
        End If
        
        curX = Abs(arrSD0_C1A5(K).ETAFIDBX)
        If curX <> 0 Then
            X = Format$(curX, "### ### ### ### ##0.00")
            XPrt.CurrentX = prtMaxX - 2100 - XPrt.TextWidth(X)
            XPrt.Print X;
        End If
        
        curX = Abs(arrSD0_C1A5(K).ETAFICRX)
        If curX <> 0 Then
            X = Format$(curX, "### ### ### ### ##0.00")
            XPrt.CurrentX = prtMaxX - 100 - XPrt.TextWidth(X)
            XPrt.Print X;
        End If
    End If

    If arrSD0_C6A8(K).ETAFIDBX <> 0 Or arrSD0_C6A8(K).ETAFICRX <> 0 Then
        XPrt.FontBold = True
        blnSaut = True
        prtBIA_ETAFI_NewLine
        XPrt.CurrentX = prtMinX + 5000: XPrt.Print arrSD0_C6A8(K).ETAFIDEV & " - Total classes 6 à 8 ";
        XPrt.CurrentX = prtMinX + 11600: XPrt.Print arrSD0_C6A8(K).ETAFIDEV;
    
        curX = Abs(arrSD0_C6A8(K).ETAFIDB)
        If curX <> 0 Then
            X = Format$(curX, "### ### ### ### ##0.00")
            XPrt.CurrentX = prtMinX + 9400 - XPrt.TextWidth(X)
            XPrt.Print X;
        End If
        
        curX = Abs(arrSD0_C6A8(K).ETAFICR)
        If curX <> 0 Then
            X = Format$(curX, "### ### ### ### ##0.00")
            XPrt.CurrentX = prtMinX + 11400 - XPrt.TextWidth(X)
            XPrt.Print X;
        End If
        
        curX = Abs(arrSD0_C6A8(K).ETAFIDBX)
        If curX <> 0 Then
            X = Format$(curX, "### ### ### ### ##0.00")
            XPrt.CurrentX = prtMaxX - 2100 - XPrt.TextWidth(X)
            XPrt.Print X;
        End If
        
        curX = Abs(arrSD0_C6A8(K).ETAFICRX)
        If curX <> 0 Then
            X = Format$(curX, "### ### ### ### ##0.00")
            XPrt.CurrentX = prtMaxX - 100 - XPrt.TextWidth(X)
            XPrt.Print X;
        End If
    End If
    If arrSD0_C9(K).ETAFIDBX <> 0 Or arrSD0_C9(K).ETAFICRX <> 0 Then
        XPrt.FontBold = False
        blnSaut = True
        prtBIA_ETAFI_NewLine
        XPrt.CurrentX = prtMinX + 5000: XPrt.Print arrSD0_C9(K).ETAFIDEV & " - Total classe 9 ";
        XPrt.CurrentX = prtMinX + 11600: XPrt.Print arrSD0_C9(K).ETAFIDEV;
    
        curX = Abs(arrSD0_C9(K).ETAFIDB)
        If curX <> 0 Then
            X = Format$(curX, "### ### ### ### ##0.00")
            XPrt.CurrentX = prtMinX + 9400 - XPrt.TextWidth(X)
            XPrt.Print X;
        End If
        
        curX = Abs(arrSD0_C9(K).ETAFICR)
        If curX <> 0 Then
            X = Format$(curX, "### ### ### ### ##0.00")
            XPrt.CurrentX = prtMinX + 11400 - XPrt.TextWidth(X)
            XPrt.Print X;
        End If
        
        curX = Abs(arrSD0_C9(K).ETAFIDBX)
        If curX <> 0 Then
            X = Format$(curX, "### ### ### ### ##0.00")
            XPrt.CurrentX = prtMaxX - 2100 - XPrt.TextWidth(X)
            XPrt.Print X;
        End If
        
        curX = Abs(arrSD0_C9(K).ETAFICRX)
        If curX <> 0 Then
            X = Format$(curX, "### ### ### ### ##0.00")
            XPrt.CurrentX = prtMaxX - 100 - XPrt.TextWidth(X)
            XPrt.Print X;
        End If
    End If
    
    curX = arrSD0_C1A5(K).ETAFIDB + arrSD0_C1A5(K).ETAFICR _
            + arrSD0_C6A8(K).ETAFIDB + arrSD0_C6A8(K).ETAFICR
    If curX <> 0 Then
        prtBIA_ETAFI_NewLine
        XPrt.CurrentX = prtMinX + 2000: XPrt.Print arrSD0_C1A5(K).ETAFIDEV & "   ?????????? ERREUR BILAN ";
        X = Format$(curX, "### ### ### ### ##0.00")
        XPrt.CurrentX = prtMinX + 1900 - XPrt.TextWidth(X)
        XPrt.Print X;

    End If
    
    
    curX = arrSD0_C9(K).ETAFIDB + arrSD0_C9(K).ETAFICR
    If curX <> 0 Then
        prtBIA_ETAFI_NewLine
        XPrt.CurrentX = prtMinX + 2000: XPrt.Print arrSD0_C9(K).ETAFIDEV & "   ?????????? ERREUR HORS-BILAN ";
        X = Format$(curX, "### ### ### ### ##0.00")
        XPrt.CurrentX = prtMinX + 1900 - XPrt.TextWidth(X)
        XPrt.Print X;
    End If
   
   arrSD0_C1A5(wNb).ETAFIDBX = arrSD0_C1A5(wNb).ETAFIDBX + arrSD0_C1A5(K).ETAFIDBX
   arrSD0_C1A5(wNb).ETAFICRX = arrSD0_C1A5(wNb).ETAFICRX + arrSD0_C1A5(K).ETAFICRX
   arrSD0_C6A8(wNb).ETAFIDBX = arrSD0_C6A8(wNb).ETAFIDBX + arrSD0_C6A8(K).ETAFIDBX
   arrSD0_C6A8(wNb).ETAFICRX = arrSD0_C6A8(wNb).ETAFICRX + arrSD0_C6A8(K).ETAFICRX
   arrSD0_C9(wNb).ETAFIDBX = arrSD0_C9(wNb).ETAFIDBX + arrSD0_C9(K).ETAFIDBX
   arrSD0_C9(wNb).ETAFICRX = arrSD0_C9(wNb).ETAFICRX + arrSD0_C9(K).ETAFICRX
   
   If blnSaut Then prtBIA_ETAFI_NewLine

Next K


End Sub

Public Sub prtBIA_ETAFI_NewLine()
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
If XPrt.CurrentY + 300 > prtMaxY Then
    frmElpPrt.prtNewPage
    prtBIA_ETAFI_Form
End If

End Sub


