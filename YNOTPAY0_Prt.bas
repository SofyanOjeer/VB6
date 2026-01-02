Attribute VB_Name = "prtYNOTPAY0"
Option Explicit

Dim Height8_6 As Integer
Dim prtMinY_YNOTPAY0 As Integer, prtMaxY_YNOTPAY0 As Integer
Dim mForm As String
Public Sub prtYNOTPAY0_Line(fgX As MSFlexGrid)
Dim K As Long, X As String, xSQL As String
Dim wFillColor As Long
Dim kCoface_New As Integer, kOCDE_New As Integer, kSp_New As Integer, kBIAN_New As Integer
Dim wJORCV As Long, wJOSEQN As Long
Dim mCurrenty As Long


'prtYNOTPAY0_Etat_Init "Echéancier Terme"
'prtYNOTPAY0_Etat

fgX.Row = 0: fgX.Col = 0: wFillColor = fgX.CellBackColor


prtYNOTPAY0_Form

XPrt.FontSize = 8: XPrt.FontBold = False
For K = 1 To fgX.Rows - 1
    fgX.Row = K
    prtYNOTPAY0_NewLine
    If K Mod 2 = 0 Then
        prtFillColor = wFillColor
    Else
        prtFillColor = prtFillColor_Standard
    End If
    kCoface_New = 0: kOCDE_New = 0: kSp_New = 0: kBIAN_New = 0

    Call frmElpPrt.prtTrame_Color(prtMinX, XPrt.CurrentY, prtMaxX, XPrt.CurrentY + prtHeaderHeight, " ")

    fgX.Col = 0: XPrt.ForeColor = fgX.CellForeColor
    XPrt.CurrentX = prtMinX + 100: XPrt.Print fgX.Text;
XPrt.FontBold = True

    fgX.Col = 3: X = Trim(fgX.Text): XPrt.ForeColor = fgX.CellForeColor
    kCoface_New = InStr(fgX.Text, ">")
    If kCoface_New > 0 Then
        prtFillColor = fgX.CellBackColor
        Call frmElpPrt.prtTrame_Color(prtMinX + 4900 + 20, XPrt.CurrentY, prtMinX + 7000 - 20, XPrt.CurrentY + prtHeaderHeight, " ")
    End If
    XPrt.CurrentX = prtMinX + 4600 + 100: XPrt.Print X;
    
    fgX.Col = 6:    X = Trim(fgX.Text): XPrt.ForeColor = fgX.CellForeColor
    kOCDE_New = InStr(fgX.Text, ">")
    If kOCDE_New > 0 Then
         prtFillColor = fgX.CellBackColor
        Call frmElpPrt.prtTrame_Color(prtMinX + 6100 + 20, XPrt.CurrentY, prtMinX + 7300 - 20, XPrt.CurrentY + prtHeaderHeight, " ")
    End If
    XPrt.CurrentX = prtMinX + 6100 + 100: XPrt.Print X;
    
    fgX.Col = 8: X = Trim(fgX.Text): XPrt.ForeColor = fgX.CellForeColor
    kSp_New = InStr(fgX.Text, ">")
    If kSp_New > 0 Then
        prtFillColor = fgX.CellBackColor
        Call frmElpPrt.prtTrame_Color(prtMinX + 7300 + 20, XPrt.CurrentY, prtMinX + 8700 - 20, XPrt.CurrentY + prtHeaderHeight, " ")
    End If
    XPrt.CurrentX = prtMinX + 7300 + 100: XPrt.Print X;
    
    fgX.Col = 11: X = Trim(fgX.Text): XPrt.ForeColor = fgX.CellForeColor
    kBIAN_New = InStr(fgX.Text, ">")
    If kBIAN_New > 0 Then
        prtFillColor = fgX.CellBackColor
        Call frmElpPrt.prtTrame_Color(prtMinX + 9100 + 20, XPrt.CurrentY, prtMinX + 10500 - 20, XPrt.CurrentY + prtHeaderHeight, " ")
        XPrt.CurrentX = prtMinX + 9100 + 100: XPrt.Print X;
    Else
        XPrt.CurrentX = prtMinX + 9400 - 100 - XPrt.TextWidth(X): XPrt.Print X;
    End If
    
    fgX.Col = 13: X = Trim(fgX.Text): XPrt.ForeColor = fgX.CellForeColor
    XPrt.CurrentX = prtMinX + 11000 - 100 - XPrt.TextWidth(X): XPrt.Print X;
XPrt.FontBold = False

    fgX.Col = 14: X = Trim(fgX.Text): XPrt.ForeColor = fgX.CellForeColor
    XPrt.CurrentX = prtMinX + 11500 - 100 - XPrt.TextWidth(X): XPrt.Print X;
    
    fgX.Col = 1: XPrt.ForeColor = fgX.CellForeColor
    X = Trim(fgX.Text)
    If X <> "" Then
        prtFillColor = fgX.CellBackColor
        Call frmElpPrt.prtTrame_Color(prtMinX + 3800 + 20, XPrt.CurrentY, prtMinX + 4600 - 20, XPrt.CurrentY + prtHeaderHeight, " ")
    End If
    XPrt.CurrentX = prtMinX + 3800 + 100: XPrt.Print X;
    
    
XPrt.CurrentY = XPrt.CurrentY + Height8_6
XPrt.FontSize = 6
    fgX.Col = 2: XPrt.ForeColor = fgX.CellForeColor
    XPrt.CurrentX = prtMinX + 3600 + 100: XPrt.Print fgX.Text;
    
    
    fgX.Col = 4: X = Trim(fgX.Text): XPrt.ForeColor = fgX.CellForeColor
    XPrt.CurrentX = prtMinX + 5100 + 100: XPrt.Print X;
    
    fgX.Col = 5: XPrt.ForeColor = fgX.CellForeColor
    If kCoface_New > 0 Then
        'XPrt.CurrentX = prtMinX + 6100 - 300: XPrt.Print Mid$(fgX.Text, 1, 1);
    Else
        XPrt.CurrentX = prtMinX + 5300 + 100: XPrt.Print fgX.Text;
    End If
    
    fgX.Col = 7: XPrt.ForeColor = fgX.CellForeColor
    If kOCDE_New > 0 Then
        'XPrt.CurrentX = prtMinX + 7300 - 300: XPrt.Print Mid$(fgX.Text, 1, 1);
    Else
        XPrt.CurrentX = prtMinX + 6500 + 100: XPrt.Print fgX.Text;
    End If
    
    fgX.Col = 9: X = Trim(fgX.Text): XPrt.ForeColor = fgX.CellForeColor
    If kSp_New > 0 Then
        'XPrt.CurrentX = prtMinX + 8700 - 300: XPrt.Print Mid$(fgX.Text, 1, 1);
    Else
        XPrt.CurrentX = prtMinX + 7900 + 100: XPrt.Print X;
    End If
    
    
    fgX.Col = 12: X = Trim(fgX.Text): XPrt.ForeColor = fgX.CellForeColor
    If kBIAN_New > 0 Then
        'XPrt.CurrentX = prtMinX + 10200 - 300: XPrt.Print Mid$(fgX.Text, 1, 1);
    Else
        XPrt.CurrentX = prtMinX + 9400 + 100: XPrt.Print fgX.Text;
    End If
    
    fgX.Col = 15: X = Trim(fgX.Text): XPrt.ForeColor = fgX.CellForeColor
    XPrt.CurrentX = prtMaxX - 2100: XPrt.Print X;
    fgX.Col = 16: X = Trim(fgX.Text): XPrt.ForeColor = fgX.CellForeColor
    XPrt.CurrentX = prtMinX + 11500 + 100: XPrt.Print X;
XPrt.CurrentY = XPrt.CurrentY - Height8_6

XPrt.FontSize = 8
    
fgX.Col = 10
If Trim(fgX.Text) <> "" Then
    prtFillColor = fgX.CellBackColor
    Call frmElpPrt.prtTrame_Color(prtMinX + 8700, XPrt.CurrentY, prtMinX + 9100, XPrt.CurrentY + prtHeaderHeight, " ")
    XPrt.CurrentX = prtMinX + 8700 + 100: XPrt.Print fgX.Text;
End If
    
'_____________________________________________________________________________________________
fgX.Col = 17
wJORCV = Val(fgX.Text)

If wJORCV > 0 Then
    fgX.Col = 18
    wJOSEQN = Val(fgX.Text)
    xSQL = "select JOENTT from " & paramIBM_Library_SABJRN & ".JRNENT0 " _
         & " where jorcv = " & wJORCV _
         & " and joSEQN = " & wJOSEQN
         
    Set rsSab = cnsab.Execute(xSQL)
    
    If Not rsSab.EOF Then
        X = rsSab("JOENTT")
        X = JOENTT_Lib(X)
        XPrt.FontSize = 6
        XPrt.CurrentY = XPrt.CurrentY + 200
        mCurrenty = XPrt.CurrentY
        XPrt.CurrentX = prtMaxX - 2100: XPrt.Print X;
        XPrt.Line (prtMaxX - 2100, XPrt.CurrentY + 180)-(prtMaxX, XPrt.CurrentY + 180), prtLineColor
        XPrt.CurrentY = mCurrenty - 100
    End If
End If
XPrt.FontSize = 8
'_____________________________________________________________________________________________


Next K
prtFillColor = prtFillColor_Standard

End Sub

Public Sub prtYNOTPAY0_Pays(fgX As MSFlexGrid)
Dim K As Long, X As String
Dim wFillColor As Long

fgX.Row = 0: fgX.Col = 0: wFillColor = RGB(230, 230, 230)


prtYNOTPAY0_Pays_Form

XPrt.FontSize = 8: XPrt.FontBold = False
For K = 1 To fgX.Rows - 1
    fgX.Row = K
    prtYNOTPAY0_NewLine
    If K Mod 2 = 0 Then
        prtFillColor = wFillColor
    Else
        prtFillColor = prtFillColor_Standard
    End If

    Call frmElpPrt.prtTrame_Color(prtMinX, XPrt.CurrentY, prtMaxX, XPrt.CurrentY + prtHeaderHeight, " ")

    fgX.Col = 0
    XPrt.CurrentX = prtMinX + 100: XPrt.Print fgX.Text;
    fgX.Col = 1:    X = Trim(fgX.Text)
    XPrt.CurrentX = prtMinX + 1000 + 100: XPrt.Print X;
    fgX.Col = 2:    X = Trim(fgX.Text)
    XPrt.CurrentX = prtMinX + 5000 + 100: XPrt.Print X;
    fgX.Col = 3: X = Trim(fgX.Text)
    XPrt.CurrentX = prtMinX + 8000 + 100: XPrt.Print X;
    fgX.Col = 4: X = Trim(fgX.Text)
    XPrt.CurrentX = prtMinX + 9000 + 100: XPrt.Print X;
    fgX.Col = 5: X = Trim(fgX.Text)
    XPrt.CurrentX = prtMinX + 12000 + 100: XPrt.Print X;
    

Next K
prtFillColor = prtFillColor_Standard

End Sub

Public Sub prtYNOTPAY0_Log(fgX As MSFlexGrid)
Dim K As Long, X As String
Dim wFillColor As Long

fgX.Row = 0: fgX.Col = 0: wFillColor = RGB(230, 230, 230)


prtYNOTPAY0_Log_Form

XPrt.FontSize = 8: XPrt.FontBold = False
For K = 1 To fgX.Rows - 1
    fgX.Row = K
    prtYNOTPAY0_NewLine
    If K Mod 2 = 0 Then
        prtFillColor = wFillColor
    Else
        prtFillColor = prtFillColor_Standard
    End If

    Call frmElpPrt.prtTrame_Color(prtMinX, XPrt.CurrentY, prtMaxX, XPrt.CurrentY + prtHeaderHeight, " ")

    fgX.Col = 0
    XPrt.CurrentX = prtMinX + 100: XPrt.Print fgX.Text;
    fgX.Col = 1:    X = Trim(fgX.Text)
    XPrt.CurrentX = prtMinX + 1200 + 100: XPrt.Print X;
    fgX.Col = 2:    X = Trim(fgX.Text)
    XPrt.CurrentX = prtMinX + 2200 + 100: XPrt.Print X;
    fgX.Col = 3: X = Trim(fgX.Text)
    XPrt.CurrentX = prtMinX + 3300 + 100: XPrt.Print X;
    fgX.Col = 4: X = Trim(fgX.Text)
    XPrt.CurrentX = prtMinX + 3700 + 100: XPrt.Print X;
    

Next K
prtFillColor = prtFillColor_Standard

End Sub

Public Sub prtYNOTPAY0_lstX(lstX As ListBox)
Dim K As Long, X As String
Dim wFillColor As Long

XPrt.FontSize = 8: XPrt.FontBold = False
For K = 1 To lstX.ListCount - 1
    prtYNOTPAY0_NewLine
    lstX.ListIndex = K
    XPrt.CurrentX = prtMinX + 100: XPrt.Print RTrim(lstX.Text);
    

Next K
prtFillColor = prtFillColor_Standard

End Sub
Public Sub prtYNOTPAY0_Close()
On Error GoTo prtError
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight

    prtMaxY_YNOTPAY0 = XPrt.CurrentY
    Select Case mForm
        Case "Form"
            prtYNOTPAY0_Col
        Case "Pays_Form"
            prtYNOTPAY0_Pays_Col
    End Select
    XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY), prtLineColor


Call frmElpPrt.prtEndDoc(1000)
frmElpPrt.Hide
Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub
'---------------------------------------------------------
Public Sub prtYNOTPAY0_Col()
'---------------------------------------------------------

XPrt.DrawWidth = 1
prtLineColor = RGB(220, 110, 0)
XPrt.Line (prtMinX, prtMinY_YNOTPAY0)-(prtMinX, prtMaxY_YNOTPAY0), prtLineColor
XPrt.Line (prtMinX + 4600, prtMinY_YNOTPAY0)-(prtMinX + 4600, prtMaxY_YNOTPAY0), prtLineColor
XPrt.Line (prtMinX + 6100, prtMinY_YNOTPAY0)-(prtMinX + 6100, prtMaxY_YNOTPAY0), prtLineColor
XPrt.Line (prtMinX + 7300, prtMinY_YNOTPAY0)-(prtMinX + 7300, prtMaxY_YNOTPAY0), prtLineColor
XPrt.Line (prtMinX + 8700, prtMinY_YNOTPAY0)-(prtMinX + 8700, prtMaxY_YNOTPAY0), prtLineColor
XPrt.Line (prtMinX + 10200, prtMinY_YNOTPAY0)-(prtMinX + 10200, prtMaxY_YNOTPAY0), prtLineColor
XPrt.Line (prtMinX + 11000, prtMinY_YNOTPAY0)-(prtMinX + 11000, prtMaxY_YNOTPAY0), prtLineColor
XPrt.Line (prtMinX + 11500, prtMinY_YNOTPAY0)-(prtMinX + 11500, prtMaxY_YNOTPAY0), prtLineColor
XPrt.Line (prtMinX + 13600, prtMinY_YNOTPAY0)-(prtMinX + 13600, prtMaxY_YNOTPAY0), prtLineColor
XPrt.Line (prtMaxX, prtMinY_YNOTPAY0)-(prtMaxX, prtMaxY_YNOTPAY0), prtLineColor
End Sub

'---------------------------------------------------------
Public Sub prtYNOTPAY0_Pays_Col()
'---------------------------------------------------------
Exit Sub
XPrt.DrawWidth = 1
prtLineColor = RGB(0, 0, 255)
XPrt.Line (prtMinX, prtMinY_YNOTPAY0)-(prtMinX, prtMaxY_YNOTPAY0), prtLineColor
XPrt.Line (prtMinX + 1000, prtMinY_YNOTPAY0)-(prtMinX + 1000, prtMaxY_YNOTPAY0), prtLineColor
XPrt.Line (prtMinX + 5000, prtMinY_YNOTPAY0)-(prtMinX + 5000, prtMaxY_YNOTPAY0), prtLineColor
XPrt.Line (prtMinX + 8000, prtMinY_YNOTPAY0)-(prtMinX + 8000, prtMaxY_YNOTPAY0), prtLineColor
XPrt.Line (prtMinX + 9000, prtMinY_YNOTPAY0)-(prtMinX + 9000, prtMaxY_YNOTPAY0), prtLineColor
XPrt.Line (prtMinX + 12000, prtMinY_YNOTPAY0)-(prtMinX + 12000, prtMaxY_YNOTPAY0), prtLineColor
XPrt.Line (prtMaxX, prtMinY_YNOTPAY0)-(prtMaxX, prtMaxY_YNOTPAY0), prtLineColor
End Sub


Public Sub prtYNOTPAY0_Open(lForm As String, lText As String)
On Error GoTo prtError

mForm = lForm
Set XPrt = Printer
frmElpPrt.Show vbModeless

Height8_6 = frmElpPrt.prtHeightDelta(8, 6)

blnFiligrane = False
If mForm = "lstX" Or mForm = "Log" Then
    prtOrientation = vbPRORPortrait
Else
    prtOrientation = vbPRORLandscape '
End If
prtPgmName = "prtYNOTPAY0"
prtTitleUsr = usrName
prtTitleText = lText
prtFontName = "Arial Unicode MS"
prtLineNb = 1
prtHeaderHeight = 300
Select Case mForm
    Case "Form": prtlineHeight = 350
    Case "Log": prtlineHeight = 200: prtFontName = prtFontName_CourierNew
    Case "lstX": prtlineHeight = 180: prtFontName = prtFontName_CourierNew
    Case "Notation pays : Journalisation": prtlineHeight = 200
    Case Else:    prtlineHeight = 280
    
End Select

prtFormType = ""
frmElpPrt.prtStdInit
XPrt.CurrentY = prtMinY
Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub

Public Sub prtYNOTPAY0_NewLine()
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
If XPrt.CurrentY + 300 > prtMaxY Then
    prtMaxY_YNOTPAY0 = prtMaxY
    Select Case mForm
        Case "Form"
            prtYNOTPAY0_Col
            frmElpPrt.prtNewPage
            prtYNOTPAY0_Form
        Case "Pays_Form"
            prtYNOTPAY0_Pays_Col
            frmElpPrt.prtNewPage
            prtYNOTPAY0_Pays_Form
         Case "Log"
            frmElpPrt.prtNewPage
            prtYNOTPAY0_Log_Form
       Case Else
            frmElpPrt.prtNewPage
    End Select
    
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    
End If

End Sub
'---------------------------------------------------------
Public Sub prtYNOTPAY0_Pays_Form()
'---------------------------------------------------------


Dim X As String
XPrt.DrawWidth = 3
XPrt.FontSize = 8: XPrt.FontBold = True
prtFillColor = RGB(0, 123, 141)
'prtFillColor = RGB(220, 110, 0) '&H80FF
XPrt.ForeColor = vbWhite

Call frmElpPrt.prtTrame_Color(prtMinX, XPrt.CurrentY, prtMaxX, XPrt.CurrentY + prtHeaderHeight, " ")
'---------------------------------------------------------
prtMinY_YNOTPAY0 = XPrt.CurrentY '+ prtHeaderHeight
XPrt.CurrentY = XPrt.CurrentY + 50

XPrt.CurrentX = prtMinX + 100: XPrt.Print "Pays";

X = "SAB"
XPrt.CurrentX = prtMinX + 1000 + 100: XPrt.Print X;
X = "Coface"
XPrt.CurrentX = prtMinX + 5000 + 100: XPrt.Print X;
X = "OCDE"
XPrt.CurrentX = prtMinX + 8000 + 100: XPrt.Print X;
X = "OCDE libellé"
XPrt.CurrentX = prtMinX + 9000 + 100: XPrt.Print X;
X = "S & P"
XPrt.CurrentX = prtMinX + 12000 + 100: XPrt.Print X;
XPrt.ForeColor = vbBlack
prtFillColor = prtFillColor_Standard

XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtHeaderHeight - prtlineHeight
End Sub

'---------------------------------------------------------
Public Sub prtYNOTPAY0_Log_Form()
'---------------------------------------------------------


Dim X As String
XPrt.DrawWidth = 3
XPrt.FontSize = 8: XPrt.FontBold = True
prtFillColor = RGB(0, 123, 141)
'prtFillColor = RGB(220, 110, 0) '&H80FF
XPrt.ForeColor = vbWhite

Call frmElpPrt.prtTrame_Color(prtMinX, XPrt.CurrentY, prtMaxX, XPrt.CurrentY + prtHeaderHeight, " ")
'---------------------------------------------------------
prtMinY_YNOTPAY0 = XPrt.CurrentY '+ prtHeaderHeight
XPrt.CurrentY = XPrt.CurrentY + 50

XPrt.CurrentX = prtMinX + 100: XPrt.Print "Date";

X = "Heure"
XPrt.CurrentX = prtMinX + 1200 + 100: XPrt.Print X;
X = "Utilisateur"
XPrt.CurrentX = prtMinX + 2200 + 100: XPrt.Print X;
X = "Seq"
XPrt.CurrentX = prtMinX + 3300 + 100: XPrt.Print X;
X = "Libellé"
XPrt.CurrentX = prtMinX + 3700 + 100: XPrt.Print X;
XPrt.ForeColor = vbBlack
prtFillColor = prtFillColor_Standard

XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtHeaderHeight - prtlineHeight
End Sub


'---------------------------------------------------------
Public Sub prtYNOTPAY0_Form()
'---------------------------------------------------------


Dim X As String
XPrt.DrawWidth = 3
XPrt.FontSize = 8: XPrt.FontBold = True
'prtFillColor = RGB(0, 123, 141)
prtFillColor = RGB(220, 110, 0) '&H80FF
XPrt.ForeColor = vbWhite

Call frmElpPrt.prtTrame_Color(prtMinX, XPrt.CurrentY, prtMaxX, XPrt.CurrentY + prtHeaderHeight, " ")
'---------------------------------------------------------
prtMinY_YNOTPAY0 = XPrt.CurrentY '+ prtHeaderHeight
XPrt.CurrentY = XPrt.CurrentY + 50

XPrt.CurrentX = prtMinX + 100: XPrt.Print "Pays";
X = "Prov / D. arrêté"
XPrt.CurrentX = prtMinX + 3300 + 100: XPrt.Print X;
X = "Coface (+affaires)"
XPrt.CurrentX = prtMinX + 4600 + 100: XPrt.Print X;
X = "OCDE"
XPrt.CurrentX = prtMinX + 6100 + 100: XPrt.Print X;
X = "S & P"
XPrt.CurrentX = prtMinX + 7300 + 100: XPrt.Print X;
X = "       BIA"
XPrt.CurrentX = prtMinX + 8700 + 100: XPrt.Print X;
X = "Taux"
XPrt.CurrentX = prtMinX + 10400 + 100: XPrt.Print X;
X = "Fisc"
XPrt.CurrentX = prtMinX + 11000 + 100: XPrt.Print X;
X = "Commentaire"
XPrt.CurrentX = prtMinX + 11500 + 100: XPrt.Print X;
XPrt.ForeColor = vbBlack
prtFillColor = prtFillColor_Standard

XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtHeaderHeight - prtlineHeight
End Sub


