Attribute VB_Name = "prtROPDOS"
Option Explicit

Dim X As String, I As Integer, Height8_6 As Integer

Dim mTextColor As Long
Const Col2_D As Integer = 3700
Const Col2_P As Integer = 6000
Dim Col2_P_Txt As Integer
Const Col2_A As Integer = 6000 '11000
Dim blnTopOfPage  As Boolean
Dim blnDossier_New  As Boolean

Type typeYROPPRT0
    ROPPRTDEST  As String
    ROPPRTID    As Long
    ROPPRTIDP   As Long
    ROPPRTIDT   As Long
    ROPPRTarrI  As Long
    ROPPRTarrD  As Long
    ROPPRTGECH  As String
    ROPPRTGUSR  As String
    
End Type
Public Sub prtROPDOS_Open(lK As Integer, lText As String)
On Error GoTo prtError

Set XPrt = Printer
frmElpPrt.Show vbModeless

Height8_6 = frmElpPrt.prtHeightDelta(8, 6)

blnFiligrane = False
If lK = 1 Then
    prtOrientation = vbPRORLandscape 'vbPRORPortrait '
    prtFontName = prtFontName_Arial
Else
    prtOrientation = vbPRORLandscape '
End If
prtPgmName = "prtROPDOS"
prtTitleUsr = usrName
prtTitleText = lText
prtLineNb = 1
prtlineHeight = 320
prtHeaderHeight = 300

prtFormType = ""
Select Case lK
    Case 1:
        frmElpPrt.prtStdInit: prtMaxX = prtMaxX + 200
        prtROPDOS_Form_1
    Case 2:
        frmElpPrt.prtStdInit ': prtMaxX = prtMaxX + 200
        Col2_P_Txt = prtMinX + Col2_P + 100
        prtROPDOS_Form_2
End Select
Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub
'---------------------------------------------------------
Public Sub prtROPDOS_Form_1()

'---------------------------------------------------------
Dim X As String
Dim curX As Currency

XPrt.DrawWidth = 1

XPrt.CurrentY = prtMinY + prtlineHeight
XPrt.FontSize = 8
prtFillColor = vbCyan ' vbBlue ' RGB(128, 255, 255)
XPrt.ForeColor = vbWhite
Call frmElpPrt.prtTrame_Color(prtMinX, prtMinY, prtMaxX, prtMinY + prtHeaderHeight, "B")
'---------------------------------------------------------
XPrt.CurrentY = prtMinY + 50
XPrt.ForeColor = vbWhite


XPrt.CurrentX = prtMinX + 50: XPrt.Print "Dossier";
XPrt.CurrentX = prtMinX + 800: XPrt.Print "Gestionnaire";
XPrt.CurrentX = prtMinX + 2300: XPrt.Print "Echéance";
XPrt.CurrentY = XPrt.CurrentY + 100
prtFillColor = prtFillColor_Standard

XPrt.FontBold = False
XPrt.FontSize = 8
XPrt.ForeColor = prtForeColor
blnTopOfPage = True
End Sub


'---------------------------------------------------------
Public Sub prtROPDOS_Form_2()

'---------------------------------------------------------
Dim X As String
Dim curX As Currency

XPrt.DrawWidth = 1

XPrt.CurrentY = prtMinY + prtlineHeight
XPrt.FontSize = 8
prtFillColor = vbCyan ' vbBlue ' RGB(128, 255, 255)
XPrt.ForeColor = vbWhite
Call frmElpPrt.prtTrame_Color(prtMinX, prtMinY, prtMaxX, prtMinY + prtHeaderHeight, "B")
'---------------------------------------------------------
XPrt.CurrentY = prtMinY + 50
XPrt.ForeColor = vbWhite


XPrt.CurrentX = prtMinX + 50: XPrt.Print "Dossier";
XPrt.CurrentX = prtMinX + 800: XPrt.Print "Gestionnaire";
XPrt.CurrentX = prtMinX + 2300: XPrt.Print "Echéance";
XPrt.CurrentY = XPrt.CurrentY + 100
prtFillColor = prtFillColor_Standard

XPrt.FontBold = False
XPrt.FontSize = 8
XPrt.ForeColor = prtForeColor
blnTopOfPage = True
End Sub



'---------------------------------------------------------
Public Sub prtROPDOS_Form_1_Col()
'---------------------------------------------------------
Dim X As String, K As Integer, K2 As Integer

XPrt.DrawWidth = 2
XPrt.Line (prtMinX + 3500, prtMinY)-(prtMinX + 3500, prtMaxY), &H808080 'prtLineColor

End Sub

'---------------------------------------------------------
Public Sub prtROPDOS_Form_2_Col()
'---------------------------------------------------------
Dim X As String, K As Integer, K2 As Integer

XPrt.DrawWidth = 2
XPrt.Line (prtMinX + 3500, prtMinY)-(prtMinX + 3500, prtMaxY), &H808080 'prtLineColor

End Sub

Public Sub prtROPDOS_Close(lK As Integer)
Dim X As String
On Error GoTo prtError
    Select Case lK
        Case 1: prtROPDOS_Form_1_Col
        Case 2: prtROPDOS_Form_2_Col
    End Select
frmElpPrt.prtEndDoc
frmElpPrt.Hide
Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub

Public Sub prtROPDOS_NewLine(lK As Integer)
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
If XPrt.CurrentY > prtMaxY - 100 Then
    Select Case lK
        Case 1: prtROPDOS_Form_1_Col
        Case 2: prtROPDOS_Form_2_Col
    End Select
    frmElpPrt.prtNewPage
    Select Case lK
        Case 1: prtROPDOS_Form_1
        Case 2: prtROPDOS_Form_2
    End Select
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight

End If

End Sub


Public Sub prtROPDOS_Dossier(lYROPDOS0 As typeYROPDOS0)
Dim X As String
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight

blnDossier_New = True
If XPrt.CurrentY + 3 * prtlineHeight > prtMaxY Then
    XPrt.CurrentY = prtMaxY
    prtROPDOS_NewLine 1
End If

XPrt.ForeColor = vbBlue
Call prtROPDOS_STAK("D", lYROPDOS0.ROPDOSSTAK, 0, 3500)

XPrt.FontBold = True
XPrt.CurrentX = prtMinX + 3400: XPrt.Print lYROPDOS0.ROPDOSSTA;

XPrt.CurrentX = prtMinX + 50: XPrt.Print Trim(lYROPDOS0.ROPDOSID);
XPrt.FontBold = False
XPrt.CurrentX = prtMinX + 800: XPrt.Print frmROPDOS.tvwSelect_Display_USR(lYROPDOS0.ROPDOSGUSR);

XPrt.ForeColor = vbBlue
XPrt.CurrentX = prtMinX + 2300: XPrt.Print dateImp10(lYROPDOS0.ROPDOSGECH);
XPrt.CurrentX = prtMinX + 3550: XPrt.Print lYROPDOS0.ROPDOSXDOM;
XPrt.CurrentX = prtMinX + 5000: XPrt.Print lYROPDOS0.ROPDOSXAPP;

If Trim(lYROPDOS0.ROPDOSXID) <> "" Then
    XPrt.CurrentX = prtMinX + 6500: XPrt.Print "L/Réf : ";
    XPrt.ForeColor = vbRed
    XPrt.Print lYROPDOS0.ROPDOSXID;
End If
XPrt.CurrentY = XPrt.CurrentY + Height8_6
XPrt.FontBold = False
XPrt.FontSize = 6
XPrt.ForeColor = &H606060

If Trim(lYROPDOS0.ROPDOSIREF) <> "" Then
    XPrt.CurrentX = prtMinX + 8500: XPrt.Print "N/Réf : " & lYROPDOS0.ROPDOSIREF & " ";
End If

XPrt.CurrentX = prtMinX + 10000
XPrt.Print lYROPDOS0.ROPDOSGPRV & ": "; Trim(frmROPDOS.tvwSelect_Display_USR(lYROPDOS0.ROPDOSISRV)) & "-" & frmROPDOS.tvwSelect_Display_USR(lYROPDOS0.ROPDOSIUSR);

XPrt.CurrentX = prtMinX + 12300: XPrt.Print dateImp10(lYROPDOS0.ROPDOSIAMJ);

XPrt.CurrentX = prtMinX + 13400: XPrt.Print "Nature : " & lYROPDOS0.ROPDOSGNAT;

XPrt.CurrentX = prtMinX + 14400: XPrt.Print "Gravité : " & lYROPDOS0.ROPDOSGGRA;
XPrt.CurrentX = prtMinX + 15400: XPrt.Print "Priorité : " & lYROPDOS0.ROPDOSGPRI;
XPrt.FontSize = 8
XPrt.CurrentY = XPrt.CurrentY - Height8_6


End Sub
Public Sub prtROPDOS_Détail_1(lYROPINF0 As typeYROPINF0)
Dim X As String, blnEnd As Boolean, blnROPINFIDT_Print As Boolean
Dim K As Integer, K1 As Integer, lenX As Integer

Select Case lYROPINF0.ROPINFGNAT
    Case "P": XPrt.DrawWidth = 8: K1 = 0: mTextColor = vbBlue
    Case "A", "F": XPrt.DrawWidth = 1: K1 = 0: mTextColor = &H606060
    Case Else: XPrt.DrawWidth = 1: K1 = 3500: mTextColor = &H808080
End Select

If blnDossier_New Then
    blnDossier_New = False
Else
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.Line (prtMinX + K1, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY), mTextColor
    XPrt.CurrentY = XPrt.CurrentY - prtlineHeight
End If
prtROPDOS_NewLine 1
XPrt.CurrentY = XPrt.CurrentY + 50

XPrt.ForeColor = prtForeColor
If lYROPINF0.ROPINFSTA = "A" Then
    XPrt.FontItalic = True
    XPrt.FontSize = 6
    XPrt.CurrentY = XPrt.CurrentY + Height8_6
End If
blnROPINFIDT_Print = False
Select Case lYROPINF0.ROPINFGNAT
    Case "P":
        XPrt.ForeColor = vbBlue
        If lYROPINF0.ROPINFSTAK <> " " Then Call prtROPDOS_STAK("P", lYROPINF0.ROPINFSTAK, 0, 100)
        XPrt.CurrentX = prtMinX: XPrt.Print "§ " & lYROPINF0.ROPINFIDP;
        XPrt.CurrentX = prtMinX + 800: XPrt.Print frmROPDOS.tvwSelect_Display_USR(lYROPINF0.ROPINFGUSR);
        XPrt.CurrentX = prtMinX + 2300: XPrt.Print dateImp10(lYROPINF0.ROPINFGECH);
        mTextColor = vbBlue
    Case "A", "F":
        If lYROPINF0.ROPINFSTAK <> " " Then Call prtROPDOS_STAK("A", lYROPINF0.ROPINFSTAK, 0, 100)
        XPrt.CurrentX = prtMinX:     XPrt.Print "+ " & lYROPINF0.ROPINFIDT;
        XPrt.CurrentX = prtMinX + 800: XPrt.Print frmROPDOS.tvwSelect_Display_USR(lYROPINF0.ROPINFGUSR);
        XPrt.CurrentX = prtMinX + 2300: XPrt.Print dateImp10(lYROPINF0.ROPINFGECH);
        blnROPINFIDT_Print = True
    Case "J":
        mTextColor = RGB(0, 0, 160)
     Case Else:
        mTextColor = vbBlack
End Select

XPrt.FontSize = 6
XPrt.CurrentY = XPrt.CurrentY + Height8_6
XPrt.ForeColor = mTextColor
If lYROPINF0.ROPINFSTA = "A" Then
    XPrt.ForeColor = vbRed
    XPrt.CurrentX = prtMinX + 3400: XPrt.Print lYROPINF0.ROPINFSTA;
    XPrt.ForeColor = &H606060
End If

If lYROPINF0.ROPINFGUO <> 0 Then
    X = Format$(lYROPINF0.ROPINFGUO / 100, "##0.00")
    XPrt.CurrentX = prtMinX + 3450 - XPrt.TextWidth(X)
    XPrt.Print X;
End If

XPrt.CurrentY = XPrt.CurrentY - Height8_6
XPrt.FontSize = 8

blnEnd = False
K = 1
lenX = Len(lYROPINF0.ROPINFGTXT)
Do
    K1 = InStr(K, lYROPINF0.ROPINFGTXT, vbCrLf)
    If K1 > 0 Then
        Call prtROPDOS_Détail_1_Txt(Mid$(lYROPINF0.ROPINFGTXT, K, K1 - K))
        K = K1 + 2
        prtROPDOS_NewLine 1
    Else
        blnEnd = True
        Call prtROPDOS_Détail_1_Txt(Mid$(lYROPINF0.ROPINFGTXT, K, lenX - K + 1))
    End If
Loop Until blnEnd
XPrt.FontItalic = False
End Sub

Public Sub prtROPDOS_Détail_2(lYROPINF0 As typeYROPINF0)
Dim X As String, blnEnd As Boolean, blnROPINFIDT_Print As Boolean
Dim K As Integer, K1 As Integer, lenX As Integer

Select Case lYROPINF0.ROPINFGNAT
    Case "P": XPrt.DrawWidth = 8: K1 = 0: mTextColor = vbBlue
    Case "A", "F": XPrt.DrawWidth = 1: K1 = 0: mTextColor = &H606060
    Case Else: XPrt.DrawWidth = 1: K1 = 3500: mTextColor = &H808080
End Select

If blnDossier_New Then
    blnDossier_New = False
Else
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.Line (prtMinX + K1, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY), mTextColor
    XPrt.CurrentY = XPrt.CurrentY - prtlineHeight
End If
prtROPDOS_NewLine 1
XPrt.CurrentY = XPrt.CurrentY + 50

XPrt.ForeColor = prtForeColor
If lYROPINF0.ROPINFSTA = "A" Then
    XPrt.FontItalic = True
    XPrt.FontSize = 6
    XPrt.CurrentY = XPrt.CurrentY + Height8_6
End If
blnROPINFIDT_Print = False
Select Case lYROPINF0.ROPINFGNAT
    Case "P":
        XPrt.ForeColor = vbBlue
        If lYROPINF0.ROPINFSTAK <> " " Then Call prtROPDOS_STAK("P", lYROPINF0.ROPINFSTAK, 0, 100)
        XPrt.CurrentX = prtMinX: XPrt.Print "§ " & lYROPINF0.ROPINFIDP;
        XPrt.CurrentX = prtMinX + 800: XPrt.Print frmROPDOS.tvwSelect_Display_USR(lYROPINF0.ROPINFGUSR);
        XPrt.CurrentX = prtMinX + 2300: XPrt.Print dateImp10(lYROPINF0.ROPINFGECH);
        mTextColor = vbBlue
    Case "A", "F":
        If lYROPINF0.ROPINFSTAK <> " " Then Call prtROPDOS_STAK("A", lYROPINF0.ROPINFSTAK, 0, 100)
        XPrt.CurrentX = prtMinX:     XPrt.Print "+ " & lYROPINF0.ROPINFIDT;
        XPrt.CurrentX = prtMinX + 800: XPrt.Print frmROPDOS.tvwSelect_Display_USR(lYROPINF0.ROPINFGUSR);
        XPrt.CurrentX = prtMinX + 2300: XPrt.Print dateImp10(lYROPINF0.ROPINFGECH);
        blnROPINFIDT_Print = True
    Case "J":
        mTextColor = RGB(0, 0, 160)
     Case Else:
        mTextColor = vbBlack
End Select

XPrt.FontSize = 6
XPrt.CurrentY = XPrt.CurrentY + Height8_6
XPrt.ForeColor = mTextColor
If lYROPINF0.ROPINFSTA = "A" Then
    XPrt.ForeColor = vbRed
    XPrt.CurrentX = prtMinX + 3400: XPrt.Print lYROPINF0.ROPINFSTA;
    XPrt.ForeColor = &H606060
End If

If lYROPINF0.ROPINFGUO <> 0 Then
    X = Format$(lYROPINF0.ROPINFGUO / 100, "##0.00")
    XPrt.CurrentX = prtMinX + 3450 - XPrt.TextWidth(X)
    XPrt.Print X;
End If

XPrt.CurrentY = XPrt.CurrentY - Height8_6
XPrt.FontSize = 8

blnEnd = False
K = 1
lenX = Len(lYROPINF0.ROPINFGTXT)
Do
    K1 = InStr(K, lYROPINF0.ROPINFGTXT, vbCrLf)
    If K1 > 0 Then
        Call prtROPDOS_Détail_1_Txt(Mid$(lYROPINF0.ROPINFGTXT, K, K1 - K))
        K = K1 + 2
        prtROPDOS_NewLine 1
    Else
        blnEnd = True
        Call prtROPDOS_Détail_1_Txt(Mid$(lYROPINF0.ROPINFGTXT, K, lenX - K + 1))
    End If
Loop Until blnEnd
XPrt.FontItalic = False
End Sub

Public Sub prtROPDOS_Détail_2P(lYROPDOS0 As typeYROPDOS0, lYROPINF0 As typeYROPINF0)
Dim X As String, blnEnd As Boolean
Dim K As Integer, K1 As Integer, lenX As Integer

XPrt.DrawWidth = 5
If Not blnTopOfPage Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY), vbBlue
    XPrt.CurrentY = XPrt.CurrentY - prtlineHeight + 100
End If
blnTopOfPage = False
prtROPDOS_NewLine 2

XPrt.ForeColor = prtForeColor

XPrt.CurrentX = prtMinX + 1100: XPrt.Print lYROPDOS0.ROPDOSXDOM;
XPrt.CurrentX = prtMinX + 2300: XPrt.Print lYROPDOS0.ROPDOSXAPP;

XPrt.CurrentX = prtMinX + 50: XPrt.Print lYROPDOS0.ROPDOSID & "-" & Format$(lYROPINF0.ROPINFIDP, "00");

Call prtROPDOS_Détail_2_ROPINFGTXT(lYROPINF0)

XPrt.FontItalic = False
End Sub

Public Sub prtROPDOS_Détail_2A(lYROPINF0 As typeYROPINF0, blnLineX As Boolean)

XPrt.DrawWidth = 1
If blnLineX Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY), &H808080
Else
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.Line (prtMinX + Col2_D, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY), &H808080
End If
XPrt.CurrentY = XPrt.CurrentY - prtlineHeight + 100

prtROPDOS_NewLine 2
blnTopOfPage = False

XPrt.ForeColor = vbBlue 'prtForeColor
If lYROPINF0.ROPINFSTA = "A" Then
    XPrt.FontItalic = True
    XPrt.FontSize = 6
    XPrt.CurrentY = XPrt.CurrentY + Height8_6
End If

XPrt.CurrentX = prtMinX + Col2_D + 50: XPrt.Print frmROPDOS.tvwSelect_Display_USR(lYROPINF0.ROPINFGUSR);
XPrt.CurrentX = prtMinX + Col2_D + 1400: XPrt.Print dateImp10(lYROPINF0.ROPINFGECH);

XPrt.CurrentX = prtMinX + Col2_D - 200: XPrt.Print Format$(lYROPINF0.ROPINFIDT, "00");

Call prtROPDOS_Détail_2_ROPINFGTXT(lYROPINF0)

XPrt.FontItalic = False
End Sub

Public Sub prtROPDOS_Détail_2_ROPINFGTXT(lYROPINF0 As typeYROPINF0)
Dim X As String, blnEnd As Boolean
Dim K As Integer, K1 As Integer, lenX As Integer

XPrt.FontSize = 8

blnEnd = False
K = 1
lenX = Len(lYROPINF0.ROPINFGTXT)
Do
    K1 = InStr(K, lYROPINF0.ROPINFGTXT, vbCrLf)
    If K1 > 0 Then
        Call prtROPDOS_Détail_2_Txt(Mid$(lYROPINF0.ROPINFGTXT, K, K1 - K), Col2_P_Txt)
        K = K1 + 2
        prtROPDOS_NewLine 2
    Else
        blnEnd = True
        Call prtROPDOS_Détail_2_Txt(Mid$(lYROPINF0.ROPINFGTXT, K, lenX - K + 1), Col2_P_Txt)
    End If
Loop Until blnEnd

End Sub

Public Sub prtROPDOS_STAK(lOrigine As String, lSTAK As String, lMin As Integer, lMax As Integer)
Dim blnOk As Boolean
Dim mCurrenty As Integer
mTextColor = prtForeColor
Select Case lSTAK
    Case "V": prtFillColor = vbGreen:  blnOk = True
    Case "R": prtFillColor = vbRed:  blnOk = True
    Case "B": prtFillColor = vbCyan:  blnOk = True
    Case "O": prtFillColor = RGB(255, 255, 190): blnOk = True '
    Case "!": prtFillColor = vbMagenta: blnOk = True
    Case Else:: blnOk = False
End Select
If blnOk Then
    If lOrigine = "D" Then
        Call frmElpPrt.prtTrame_Color(prtMinX + lMin, XPrt.CurrentY, prtMinX + lMax, XPrt.CurrentY + prtlineHeight - 120, " ")
        prtFillColor = RGB(192, 255, 255)
        Call frmElpPrt.prtTrame_Color(prtMinX + lMax, XPrt.CurrentY, prtMaxX, XPrt.CurrentY + prtlineHeight - 120, " ")
        mCurrenty = XPrt.CurrentY
        XPrt.Line (prtMinX, XPrt.CurrentY + prtlineHeight - 120)-(prtMaxX, XPrt.CurrentY + prtlineHeight - 120), vbBlue
        XPrt.CurrentY = mCurrenty
    Else
        'XPrt.Line (prtMinX + lMin, XPrt.CurrentY - 100)-(prtMaxX, XPrt.CurrentY - 100), prtFillColor
        Call frmElpPrt.prtTrame_Color(prtMinX + lMin, XPrt.CurrentY, prtMinX + lMax, XPrt.CurrentY + prtlineHeight - 120, " ")
    End If
Else
    If lOrigine = "D" Then
        prtFillColor = RGB(192, 255, 255)
        Call frmElpPrt.prtTrame_Color(prtMinX, XPrt.CurrentY, prtMaxX, XPrt.CurrentY + prtlineHeight - 120, " ")
        mCurrenty = XPrt.CurrentY
        XPrt.Line (prtMinX, XPrt.CurrentY + prtlineHeight - 120)-(prtMaxX, XPrt.CurrentY + prtlineHeight - 120), vbBlue
        XPrt.CurrentY = mCurrenty
    End If
End If
End Sub

Public Sub prtROPDOS_Détail_1_Txt(lTxt As String)
Dim wLen As Long, I As Long, iW As Long, Ic As Long
 
XPrt.CurrentX = prtMinX + 3550
If XPrt.TextWidth(lTxt) < prtMaxX - XPrt.CurrentX Then
    XPrt.Print lTxt;
Else
    For I = 1 To Len(lTxt)
        If XPrt.CurrentX > prtMaxX Then
            prtROPDOS_NewLine 1
            XPrt.CurrentX = prtMinX + 3550
        End If
        XPrt.Print Mid$(lTxt, I, 1);
    Next I

End If
End Sub
Public Sub prtROPDOS_Détail_2_Txt(lTxt As String, lCol As Integer)
Dim wLen As Long, I As Long, iW As Long, Ic As Long
 
XPrt.CurrentX = prtMinX + 3550
If XPrt.TextWidth(lTxt) < prtMaxX - XPrt.CurrentX Then
    XPrt.Print lTxt;
Else
    For I = 1 To Len(lTxt)
        If XPrt.CurrentX > prtMaxX Then
            prtROPDOS_NewLine 1
            XPrt.CurrentX = prtMinX + 3550
        End If
        XPrt.Print Mid$(lTxt, I, 1);
    Next I

End If


End Sub

