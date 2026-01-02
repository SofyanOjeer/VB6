Attribute VB_Name = "prtDGI_2561"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim X As String, I As Integer, Height8_6 As Integer

Public prtrecDGI_2561 As typeDGI_2561
Dim bln2561Bis As Boolean
'---------------------------------------------------------
 Public Sub prtDGI_2561_Monitor(Msg As String)
'---------------------------------------------------------
Dim X As String

On Error GoTo prtError

Set XPrt = Printer
Height8_6 = frmElpPrt.prtHeightDelta(8, 6)

prtTitleText = "DGI_2561"

prtLineNb = 1

frmElpPrt.Show vbModeless


prtOrientation = vbPRORPortrait
prtPgmName = "prtDGI_2561"
prtTitleUsr = usrName

prtlineHeight = 300
prtHeaderHeight = 300

bln2561Bis = False
frmElpPrt.prtStdBlankInit
'frmElpPrt.prtFiligrane paramTemp_Folder & "\IFU\2561_2003.bmp"   'prtFiligrane_Name
prtMinX = prtMinX + 400
prtMinY = prtMinY + 1000

prtDGI_2561_Form
prtDGI_2561_Line
frmElpPrt.prtEndDoc

'bln2561Bis = True
'frmElpPrt.prtStdBlankInit
'frmElpPrt.prtFiligrane paramTemp_Folder & "\IFU\2561_2003_bis.bmp"   'prtFiligrane_Name
'prtMinX = prtMinX + 400
'prtMinY = prtMinY + 1400

'prtDGI_2561_Bis_Form
'prtDGI_2561_Bis_Line

'frmElpPrt.prtEndDoc


frmElpPrt.Hide
prtOrientation = vbPRORPortrait
prtPgmName = "prtDGI_2561"
prtTitleUsr = usrName


Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide
End Sub
'---------------------------------------------------------
Public Sub prtDGI_2561_Form()
'---------------------------------------------------------
Dim X As String
''XPrt.PaintPicture frmDGI_2561.img.Picture, 0, 0
XPrt.FontSize = 9
XPrt.FontBold = True
XPrt.CurrentX = prtMinX + 2300
XPrt.CurrentY = prtMinY + 1430
XPrt.Print "BANQUE INTERCONTINENTALE ARABE";
XPrt.CurrentX = prtMinX + 8800
XPrt.Print "C";
XPrt.CurrentY = prtMinY + 1700
XPrt.CurrentX = prtMinX + 8800
XPrt.Print "B";


XPrt.CurrentX = prtMinX + 2300
XPrt.CurrentY = prtMinY + 1950
XPrt.Print "67";
XPrt.CurrentX = prtMinX + 8800
XPrt.CurrentY = prtMinY + 2000
''XPrt.Print "2003";

XPrt.CurrentY = prtMinY + 2320
XPrt.CurrentX = prtMinX + 8800
XPrt.Print "12179";



XPrt.CurrentX = prtMinX + 2300
XPrt.CurrentY = prtMinY + 2250
XPrt.Print "AVENUE FRANKLIN ROOSEVELT";

XPrt.CurrentY = prtMinY + 2600
XPrt.CurrentX = prtMinX + 8800
XPrt.Print "00001";

XPrt.CurrentX = prtMinX + 2300
XPrt.CurrentY = prtMinY + 2620
XPrt.Print "PARIS";

XPrt.CurrentX = prtMinX + 2300
XPrt.CurrentY = prtMinY + 2880
XPrt.Print "75008";

XPrt.CurrentX = prtMinX + 2300
XPrt.CurrentY = prtMinY + 3600
XPrt.Print "302590070 00017";


End Sub

'---------------------------------------------------------
Public Sub prtDGI_2561_Bis_Form()
'---------------------------------------------------------
Dim X As String
''XPrt.PaintPicture frmDGI_2561.img.Picture, 0, 0
XPrt.FontSize = 9
XPrt.FontBold = True
XPrt.CurrentX = prtMinX + 2300
XPrt.CurrentY = prtMinY + 1700
XPrt.Print "BANQUE INTERCONTINENTALE ARABE";
XPrt.CurrentY = prtMinY + 2050
XPrt.CurrentX = prtMinX + 8800
XPrt.Print "C";
XPrt.CurrentY = prtMinY + 1850
XPrt.CurrentX = prtMinX + 8800
XPrt.Print "B";


XPrt.CurrentX = prtMinX + 2100
XPrt.CurrentY = prtMinY + 2500
XPrt.Print "67";
XPrt.CurrentX = prtMinX + 8800
XPrt.CurrentY = prtMinY + 2480
''XPrt.Print "2003";

XPrt.CurrentY = prtMinY + 2520
XPrt.CurrentX = prtMinX + 8800
XPrt.Print "12179";



XPrt.CurrentX = prtMinX + 2300
XPrt.CurrentY = prtMinY + 2400
XPrt.Print "AVENUE FRANKLIN ROOSEVELT";

XPrt.CurrentY = prtMinY + 2750
XPrt.CurrentX = prtMinX + 8800
XPrt.Print "0001";

XPrt.CurrentX = prtMinX + 2300
XPrt.CurrentY = prtMinY + 2790
XPrt.Print "PARIS";

XPrt.CurrentX = prtMinX + 2300
XPrt.CurrentY = prtMinY + 3100
XPrt.Print "75008";

XPrt.CurrentX = prtMinX + 2300
XPrt.CurrentY = prtMinY + 3850
XPrt.Print "302590070 00017";


End Sub


'---------------------------------------------------------
Public Sub prtDGI_2561_Line()
'---------------------------------------ZONE 1
Dim X As String

XPrt.FontSize = 9
XPrt.CurrentY = prtMinY + 2900
XPrt.CurrentX = prtMinX + 8800
XPrt.Print (prtrecDGI_2561.AI);

XPrt.CurrentY = prtMinY + 3400
XPrt.CurrentX = prtMinX + 8800
XPrt.Print (prtrecDGI_2561.AH);

XPrt.CurrentY = prtMinY + 3700
XPrt.CurrentX = prtMinX + 8800
XPrt.Print (prtrecDGI_2561.BR);

XPrt.CurrentY = prtMinY + 4200
XPrt.CurrentX = prtMinX + 8800
XPrt.Print (prtrecDGI_2561.AC);

XPrt.CurrentY = prtMinY + 4920
XPrt.CurrentX = prtMinX + 8800
XPrt.Print (prtrecDGI_2561.AE);

XPrt.CurrentY = prtMinY + 5170
XPrt.CurrentX = prtMinX + 8800
XPrt.Print (prtrecDGI_2561.AF);

XPrt.CurrentY = prtMinY + 5400
If Trim(prtrecDGI_2561.AO) = "1" Then
    XPrt.CurrentX = prtMinX + 9100
Else
    XPrt.CurrentX = prtMinX + 9900
End If

XPrt.Print (prtrecDGI_2561.AO);

XPrt.CurrentY = prtMinY + 5700
XPrt.CurrentX = prtMinX + 8800
XPrt.Print (prtrecDGI_2561.CT);


'------------------------------------ZONE 2
XPrt.CurrentY = prtMinY + 4220
XPrt.CurrentX = prtMinX + 2300
XPrt.Print (prtrecDGI_2561.ZC);

XPrt.CurrentY = prtMinY + 4450
XPrt.CurrentX = prtMinX + 2300
XPrt.Print (prtrecDGI_2561.ZD);

XPrt.CurrentY = prtMinY + 5170
XPrt.CurrentX = prtMinX + 2300
XPrt.Print (prtrecDGI_2561.ZG);

XPrt.CurrentY = prtMinY + 5550
XPrt.CurrentX = prtMinX + 2300
XPrt.Print (prtrecDGI_2561.ZH);

XPrt.CurrentY = prtMinY + 5840
XPrt.CurrentX = prtMinX + 2300
XPrt.Print (prtrecDGI_2561.ZI);

XPrt.CurrentY = prtMinY + 6080
XPrt.CurrentX = prtMinX + 2300
XPrt.Print (prtrecDGI_2561.ZJ);

If Not bln2561Bis Then
    '--------------------------------------ZONE 3
    XPrt.CurrentY = prtMinY + 7980
    X = Format$(prtrecDGI_2561.AR, "#### ### ### ### ###")
    XPrt.CurrentX = prtMinX + 4600 - XPrt.TextWidth(X)
    XPrt.Print X;
    
    '--------------------------------------ZONE 4
    XPrt.CurrentY = prtMinY + 10820
    X = Format$(prtrecDGI_2561.BN, "#### ### ### ### ###")
    XPrt.CurrentX = prtMinX + 10200 - XPrt.TextWidth(X)
    XPrt.Print X;
        
    
    XPrt.CurrentY = prtMinY + 11050
    X = Format$(prtrecDGI_2561.BP, "#### ### ### ### ###")
    XPrt.CurrentX = prtMinX + 10200 - XPrt.TextWidth(X)
    XPrt.Print X;

End If

End Sub
'---------------------------------------------------------
Public Sub prtDGI_2561_Bis_Line()
'---------------------------------------ZONE 1
Dim X As String

XPrt.FontSize = 9
XPrt.CurrentY = prtMinY + 3550
XPrt.CurrentX = prtMinX + 8800
XPrt.Print (prtrecDGI_2561.AI);

XPrt.CurrentY = prtMinY + 3930
XPrt.CurrentX = prtMinX + 8800
XPrt.Print (prtrecDGI_2561.AH);

XPrt.CurrentY = prtMinY + 4170
XPrt.CurrentX = prtMinX + 8800
XPrt.Print (prtrecDGI_2561.BR);

XPrt.CurrentY = prtMinY + 4600
XPrt.CurrentX = prtMinX + 8800
XPrt.Print (prtrecDGI_2561.AC);

XPrt.CurrentY = prtMinY + 5270
XPrt.CurrentX = prtMinX + 8800
XPrt.Print (prtrecDGI_2561.AE);

XPrt.CurrentY = prtMinY + 5480
XPrt.CurrentX = prtMinX + 8800
XPrt.Print (prtrecDGI_2561.AF);

XPrt.CurrentY = prtMinY + 5700
If Trim(prtrecDGI_2561.AO) = "1" Then
    XPrt.CurrentX = prtMinX + 8800
Else
    XPrt.CurrentX = prtMinX + 9200
End If

XPrt.Print (prtrecDGI_2561.AO);

XPrt.CurrentY = prtMinY + 5900
XPrt.CurrentX = prtMinX + 8800
XPrt.Print (prtrecDGI_2561.CT);


'------------------------------------ZONE 2
XPrt.CurrentY = prtMinY + 4630
XPrt.CurrentX = prtMinX + 2300
XPrt.Print (prtrecDGI_2561.ZC);

XPrt.CurrentY = prtMinY + 4850
XPrt.CurrentX = prtMinX + 2300
XPrt.Print (prtrecDGI_2561.ZD);

XPrt.CurrentY = prtMinY + 5470
XPrt.CurrentX = prtMinX + 2300
XPrt.Print (prtrecDGI_2561.ZG);

XPrt.CurrentY = prtMinY + 5750
XPrt.CurrentX = prtMinX + 2300
XPrt.Print (prtrecDGI_2561.ZH);

XPrt.CurrentY = prtMinY + 6100
XPrt.CurrentX = prtMinX + 2300
XPrt.Print (prtrecDGI_2561.ZI);

XPrt.CurrentY = prtMinY + 6330
XPrt.CurrentX = prtMinX + 2300
XPrt.Print (prtrecDGI_2561.ZJ);


End Sub














