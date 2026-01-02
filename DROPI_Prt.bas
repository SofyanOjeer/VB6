Attribute VB_Name = "prtDROPI"
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

Type typeDROPI_rtf
    K        As Long
    Length   As Long
    Color    As Long
    
End Type

Public Sub prtDROPI_Open(lK As Integer, lText As String)
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
prtPgmName = "prtDROPI"
prtTitleUsr = usrName
prtTitleText = lText
prtLineNb = 1
prtlineHeight = 320
prtHeaderHeight = 300

prtFormType = ""
Select Case lK
    Case 1:
        frmElpPrt.prtStdInit: prtMaxX = prtMaxX + 200
        prtDROPI_Form_1
    Case 2:
        frmElpPrt.prtStdInit ': prtMaxX = prtMaxX + 200
        Col2_P_Txt = prtMinX + Col2_P + 100
        prtDROPI_Form_2
End Select
Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub
'---------------------------------------------------------
Public Sub prtDROPI_Form_1()

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
Public Sub prtDROPI_Form_2()

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
Public Sub prtDROPI_Form_1_Col()
'---------------------------------------------------------
Dim X As String, K As Integer, K2 As Integer

XPrt.DrawWidth = 2
XPrt.Line (prtMinX + 3500, prtMinY)-(prtMinX + 3500, prtMaxY), &H808080 'prtLineColor

End Sub

'---------------------------------------------------------
Public Sub prtDROPI_Form_2_Col()
'---------------------------------------------------------
Dim X As String, K As Integer, K2 As Integer

XPrt.DrawWidth = 2
XPrt.Line (prtMinX + 3500, prtMinY)-(prtMinX + 3500, prtMaxY), &H808080 'prtLineColor

End Sub

Public Sub prtDROPI_Close(lK As Integer)
Dim X As String
On Error GoTo prtError
    Select Case lK
        Case 1: prtDROPI_Form_1_Col
        Case 2: prtDROPI_Form_2_Col
    End Select
Call frmElpPrt.prtEndDoc(1000)
frmElpPrt.Hide
Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub

Public Sub prtDROPI_NewLine(lK As Integer)
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
If XPrt.CurrentY > prtMaxY - 100 Then
    mTextColor = XPrt.ForeColor
    Select Case lK
        Case 1: prtDROPI_Form_1_Col
        Case 2: prtDROPI_Form_2_Col
    End Select
    frmElpPrt.prtNewPage
    Select Case lK
        Case 1: prtDROPI_Form_1
        Case 2: prtDROPI_Form_2
    End Select
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.ForeColor = mTextColor
End If

End Sub


Public Sub prtDROPI_Dossier(lYROPDOS0 As typeYROPDOS0)
Dim X As String
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight

blnDossier_New = True
If XPrt.CurrentY + 3 * prtlineHeight > prtMaxY Then
    XPrt.CurrentY = prtMaxY
    prtDROPI_NewLine 1
End If

XPrt.ForeColor = vbBlue
Call prtDROPI_STAK("D", lYROPDOS0.ROPDOSSTAK, 0, 3500)

XPrt.FontBold = True
XPrt.CurrentX = prtMinX + 3400: XPrt.Print lYROPDOS0.ROPDOSSTA;

XPrt.CurrentX = prtMinX + 50: XPrt.Print Trim(lYROPDOS0.ROPDOSID);
XPrt.FontBold = False
XPrt.CurrentX = prtMinX + 800: XPrt.Print frmDROPI.fraDossier_Display_USR(lYROPDOS0.ROPDOSGUSR);

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
XPrt.Print lYROPDOS0.ROPDOSGPRV & ": "; Trim(frmDROPI.fraDossier_Display_USR(lYROPDOS0.ROPDOSISRV)) & "-" & frmDROPI.fraDossier_Display_USR(lYROPDOS0.ROPDOSIUSR);

XPrt.CurrentX = prtMinX + 12300: XPrt.Print dateImp10(lYROPDOS0.ROPDOSIAMJ);

XPrt.CurrentX = prtMinX + 13400: XPrt.Print "Nature : " & lYROPDOS0.ROPDOSGNAT;

XPrt.CurrentX = prtMinX + 14400: XPrt.Print "Gravité : " & lYROPDOS0.ROPDOSGGRA;
XPrt.CurrentX = prtMinX + 15400: XPrt.Print "Priorité : " & lYROPDOS0.ROPDOSGPRI;
XPrt.FontSize = 8
XPrt.CurrentY = XPrt.CurrentY - Height8_6


End Sub
Public Sub prtDROPI_Détail_1(lYROPINF0 As typeYROPINF0)
Dim X As String, blnEnd As Boolean, blnROPINFIDT_Print As Boolean
Dim K As Integer, K1 As Integer, lenX As Integer

Select Case lYROPINF0.ROPINFGNAT
    Case "P": XPrt.DrawWidth = 8: K1 = 0: mTextColor = RGB(0, 0, 200)
    Case "A", "F": XPrt.DrawWidth = 1: K1 = 0: mTextColor = RGB(0, 128, 196)
    Case Else: XPrt.DrawWidth = 1: K1 = 3500: mTextColor = RGB(164, 164, 164)
End Select

If blnDossier_New Then
    blnDossier_New = False
Else
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.Line (prtMinX + K1, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY), mTextColor
    XPrt.CurrentY = XPrt.CurrentY - prtlineHeight
End If
prtDROPI_NewLine 1
XPrt.CurrentY = XPrt.CurrentY + 50

XPrt.ForeColor = mTextColor
If lYROPINF0.ROPINFSTA = "A" Then
    XPrt.FontItalic = True
    XPrt.FontSize = 6
    XPrt.CurrentY = XPrt.CurrentY + Height8_6
End If
blnROPINFIDT_Print = False
Select Case lYROPINF0.ROPINFGNAT
    Case "P":
        XPrt.ForeColor = vbBlue
        If lYROPINF0.ROPINFSTAK <> " " Then Call prtDROPI_STAK("P", lYROPINF0.ROPINFSTAK, 0, 100)
        XPrt.CurrentX = prtMinX: XPrt.Print lYROPINF0.ROPINFIDP;
        XPrt.CurrentX = prtMinX + 800: XPrt.Print frmDROPI.fraDossier_Display_USR(lYROPINF0.ROPINFGUSR);
        XPrt.CurrentX = prtMinX + 2300: XPrt.Print dateImp10(lYROPINF0.ROPINFGECH);
        mTextColor = RGB(0, 0, 96)
    Case "A", "F":
        If lYROPINF0.ROPINFSTAK <> " " Then Call prtDROPI_STAK("A", lYROPINF0.ROPINFSTAK, 0, 100)
        XPrt.CurrentX = prtMinX:     XPrt.Print lYROPINF0.ROPINFIDP & "-" & lYROPINF0.ROPINFIDT;
        XPrt.CurrentX = prtMinX + 800: XPrt.Print frmDROPI.fraDossier_Display_USR(lYROPINF0.ROPINFGUSR);
        XPrt.CurrentX = prtMinX + 2300: XPrt.Print dateImp10(lYROPINF0.ROPINFGECH);
        If lYROPINF0.ROPINFSTA = " " Then
            mTextColor = RGB(0, 128, 196)
        Else
            mTextColor = RGB(128, 128, 128)
        End If
        blnROPINFIDT_Print = True
    Case "J", "N":
        mTextColor = RGB(164, 164, 200)
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
        Call prtDROPI_Détail_1_Txt(Mid$(lYROPINF0.ROPINFGTXT, K, K1 - K))
        K = K1 + 2
        prtDROPI_NewLine 1
    Else
        blnEnd = True
        Call prtDROPI_Détail_1_Txt(Mid$(lYROPINF0.ROPINFGTXT, K, lenX - K + 1))
    End If
Loop Until blnEnd
XPrt.FontItalic = False
End Sub

Public Sub prtDROPI_Détail_2(lYROPINF0 As typeYROPINF0)
Dim X As String, blnEnd As Boolean, blnROPINFIDT_Print As Boolean
Dim K As Integer, K1 As Integer, lenX As Integer

Select Case lYROPINF0.ROPINFGNAT
    Case "P": XPrt.DrawWidth = 8: K1 = 0: mTextColor = RGB(0, 0, 96)
    Case "A", "F": XPrt.DrawWidth = 1: K1 = 0: mTextColor = RGB(0, 128, 196)
    Case Else: XPrt.DrawWidth = 1: K1 = 3500: mTextColor = RGB(164, 164, 164)
End Select

If blnDossier_New Then
    blnDossier_New = False
Else
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.Line (prtMinX + K1, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY), mTextColor
    XPrt.CurrentY = XPrt.CurrentY - prtlineHeight
End If
prtDROPI_NewLine 1
XPrt.CurrentY = XPrt.CurrentY + 50

XPrt.ForeColor = mTextColor
If lYROPINF0.ROPINFSTA = "A" Then
    XPrt.FontItalic = True
    XPrt.FontSize = 6
    XPrt.CurrentY = XPrt.CurrentY + Height8_6
End If
blnROPINFIDT_Print = False
Select Case lYROPINF0.ROPINFGNAT
    Case "P":
        XPrt.ForeColor = vbBlue
        If lYROPINF0.ROPINFSTAK <> " " Then Call prtDROPI_STAK("P", lYROPINF0.ROPINFSTAK, 0, 100)
        XPrt.CurrentX = prtMinX: XPrt.Print lYROPINF0.ROPINFIDP;
        XPrt.CurrentX = prtMinX + 800: XPrt.Print frmDROPI.fraDossier_Display_USR(lYROPINF0.ROPINFGUSR);
        XPrt.CurrentX = prtMinX + 2300: XPrt.Print dateImp10(lYROPINF0.ROPINFGECH);
        mTextColor = RGB(0, 0, 96)
    Case "A", "F":
        If lYROPINF0.ROPINFSTAK <> " " Then Call prtDROPI_STAK("A", lYROPINF0.ROPINFSTAK, 0, 100)
        XPrt.CurrentX = prtMinX:     XPrt.Print lYROPINF0.ROPINFIDP & "-" & lYROPINF0.ROPINFIDT;
        XPrt.CurrentX = prtMinX + 800: XPrt.Print frmDROPI.fraDossier_Display_USR(lYROPINF0.ROPINFGUSR);
        XPrt.CurrentX = prtMinX + 2300: XPrt.Print dateImp10(lYROPINF0.ROPINFGECH);
        If lYROPINF0.ROPINFSTA = " " Then
            mTextColor = RGB(0, 128, 196)
        Else
            mTextColor = RGB(128, 128, 128)
        End If
        blnROPINFIDT_Print = True
    Case "J", "N":
        mTextColor = RGB(164, 164, 200)
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
        Call prtDROPI_Détail_1_Txt(Mid$(lYROPINF0.ROPINFGTXT, K, K1 - K))
        K = K1 + 2
        prtDROPI_NewLine 1
    Else
        blnEnd = True
        Call prtDROPI_Détail_1_Txt(Mid$(lYROPINF0.ROPINFGTXT, K, lenX - K + 1))
    End If
Loop Until blnEnd
XPrt.FontItalic = False
End Sub

Public Sub prtDROPI_Détail_2_ROPINFGTXT(lYROPINF0 As typeYROPINF0)
Dim X As String, blnEnd As Boolean
Dim K As Integer, K1 As Integer, lenX As Integer

XPrt.FontSize = 8

blnEnd = False
K = 1
lenX = Len(lYROPINF0.ROPINFGTXT)
Do
    K1 = InStr(K, lYROPINF0.ROPINFGTXT, vbCrLf)
    If K1 > 0 Then
        Call prtDROPI_Détail_2_Txt(Mid$(lYROPINF0.ROPINFGTXT, K, K1 - K), Col2_P_Txt)
        K = K1 + 2
        prtDROPI_NewLine 2
    Else
        blnEnd = True
        Call prtDROPI_Détail_2_Txt(Mid$(lYROPINF0.ROPINFGTXT, K, lenX - K + 1), Col2_P_Txt)
    End If
Loop Until blnEnd

End Sub

Public Sub prtDROPI_STAK(lOrigine As String, lSTAK As String, lMin As Integer, lMax As Integer)
Dim blnOk As Boolean
Dim mCurrenty As Integer
mTextColor = prtForeColor
Select Case lSTAK
    Case "V": prtFillColor = RGB(200, 255, 200): blnOk = True: mTextColor = RGB(0, 164, 0)
    Case "R": prtFillColor = RGB(255, 200, 200): blnOk = True: mTextColor = RGB(200, 0, 0)
    Case "B": prtFillColor = RGB(192, 255, 255): blnOk = True: mTextColor = RGB(0, 128, 196)
    Case "O": prtFillColor = RGB(255, 255, 190): blnOk = True: mTextColor = RGB(255, 128, 0)
    Case "!": prtFillColor = RGB(255, 200, 255): blnOk = True: mTextColor = vbMagenta
    Case Else:: blnOk = False: mTextColor = RGB(64, 64, 64)
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
        'Call frmElpPrt.prtTrame_Color(prtMinX + lMin, XPrt.CurrentY, prtMinX + lMax, XPrt.CurrentY + prtlineHeight - 120, " ")
        XPrt.ForeColor = mTextColor
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

Public Sub prtDROPI_Détail_1_Txt(lTxt As String)
Dim wLen As Long, I As Long, iW As Long, Ic As Long
 
XPrt.CurrentX = prtMinX + 3550
'If XPrt.TextWidth(lTxt) < prtMaxX - XPrt.CurrentX Then
If Len(lTxt) < 150 Then
    XPrt.Print lTxt;
Else
    For I = 1 To Len(lTxt)
        If XPrt.CurrentX > prtMaxX Then
            prtDROPI_NewLine 1
            XPrt.CurrentX = prtMinX + 3550
        End If
        XPrt.Print Mid$(lTxt, I, 1);
    Next I

End If
End Sub
Public Sub prtDROPI_Détail_2_Txt(lTxt As String, lCol As Integer)
Dim wLen As Long, I As Long, iW As Long, Ic As Long
 
XPrt.CurrentX = prtMinX + 3550
If XPrt.TextWidth(lTxt) < prtMaxX - XPrt.CurrentX Then
    XPrt.Print lTxt;
Else
    For I = 1 To Len(lTxt)
        If XPrt.CurrentX > prtMaxX Then
            prtDROPI_NewLine 1
            XPrt.CurrentX = prtMinX + 3550
        End If
        XPrt.Print Mid$(lTxt, I, 1);
    Next I

End If


End Sub

