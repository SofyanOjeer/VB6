Attribute VB_Name = "prtBIA_TVAFAC"
Option Explicit
Dim Page_No As Integer, blnNewPage As Boolean
Dim mCurrenty_Top As Long, mCurrentY_Page As Long
Dim me_prtMaxY As Long
Dim mYTVAFAC0 As typeYTVAFAC0, mZADRESS0 As typeZADRESS0, mZCLIENA0 As typeZCLIENA0
Dim mTVANIFCLIT_Pays As Boolean
Public Sub prtBIA_TVAFAC_Close(lK As Integer)
Dim X As String
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

Public Sub prtBIA_TVAFAC_Open(lK As Integer, lText As String)
On Error GoTo prtError

Set XPrt = Printer
frmElpPrt.Msg_Rcv 'Show vbModeless

'Height8_6 = frmElpPrt.prtHeightDelta(8, 6)

blnFiligrane = False
If lK = 1 Or lK = 4 Then
    prtOrientation = vbPRORPortrait '
Else
    prtOrientation = vbPRORLandscape '
End If
prtPgmName = "prtBIA_TVAFAC"
prtTitleUsr = usrName
prtTitleText = lText
prtLineNb = 1
prtlineHeight = 300
prtHeaderHeight = 300

prtFormType = ""
Select Case lK
    Case 1:

        prtSocInit
        me_prtMaxY = prtMaxY - 400
    Case 2:
        frmElpPrt.prtStdInit
        me_prtMaxY = prtMaxY - 400
        prtBIA_TVACOM_Form_2
    Case 3:
        frmElpPrt.prtStdInit
        me_prtMaxY = prtMaxY - 400
        prtBIA_TVAFAC_Form_3
    Case 4:
        frmElpPrt.prtStdInit
        me_prtMaxY = prtMaxY - 400
        prtBIA_TVAFAC_Form_4
    Case 9:
        frmElpPrt.prtStdInit
        me_prtMaxY = prtMaxY - 400
        prtBIA_TVAFAC_Form_9
    Case 10:
        frmElpPrt.prtStdInit
        me_prtMaxY = prtMaxY - 400
        prtBIA_TVAFAC_Form_10
End Select
blnNewPage = False
Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub
'---------------------------------------------------------
Public Sub prtBIA_TVAFAC_Form_1_Col(lFct As String)
'---------------------------------------------------------
Dim X As String, K As Integer, K2 As Integer


XPrt.DrawWidth = 2
If lFct = "C" Then
    XPrt.Line (prtMinX, mCurrenty_Top)-(prtMinX, XPrt.CurrentY), RGB(0, 123, 141)
    XPrt.Line (prtMinX + 1000, mCurrenty_Top)-(prtMinX + 1000, XPrt.CurrentY), RGB(0, 123, 141)
    XPrt.Line (prtMinX + 2400, mCurrenty_Top)-(prtMinX + 2400, XPrt.CurrentY), RGB(0, 123, 141)
    XPrt.Line (prtMinX + 6000, mCurrenty_Top)-(prtMinX + 6000, XPrt.CurrentY), RGB(0, 123, 141)
    XPrt.Line (prtMinX + 6400, mCurrenty_Top)-(prtMinX + 6400, XPrt.CurrentY), RGB(0, 123, 141)
End If
XPrt.Line (prtMinX + 8250, mCurrenty_Top)-(prtMinX + 8250, XPrt.CurrentY), RGB(0, 123, 141)
XPrt.Line (prtMinX + 8550, mCurrenty_Top)-(prtMinX + 8550, XPrt.CurrentY), RGB(0, 123, 141)
XPrt.Line (prtMinX + 9950, mCurrenty_Top)-(prtMinX + 9950, XPrt.CurrentY), RGB(0, 123, 141)
XPrt.Line (prtMaxX, mCurrenty_Top)-(prtMaxX, XPrt.CurrentY), RGB(0, 123, 141)

mCurrenty_Top = XPrt.CurrentY

End Sub


'---------------------------------------------------------
Public Sub prtBIA_TVACOM_Form_2_Col()
'---------------------------------------------------------
Dim X As String, K As Integer, K2 As Integer

XPrt.DrawWidth = 2
XPrt.Line (prtMinX, mCurrenty_Top)-(prtMinX, prtMaxY), prtLineColor
XPrt.Line (prtMinX + 6100, mCurrenty_Top)-(prtMinX + 6100, prtMaxY), prtLineColor
'XPrt.Line (prtMinX + 8450, mCurrentY_Top)-(prtMinX + 8450,prtmaxy), prtLineColor
XPrt.Line (prtMinX + 9400, mCurrenty_Top)-(prtMinX + 9400, prtMaxY), prtLineColor
XPrt.Line (prtMinX + 11900, mCurrenty_Top)-(prtMinX + 11900, prtMaxY), prtLineColor
XPrt.Line (prtMinX + 12900, mCurrenty_Top)-(prtMinX + 12900, prtMaxY), prtLineColor
XPrt.Line (prtMinX + 13700, mCurrenty_Top)-(prtMinX + 13700, prtMaxY), prtLineColor
XPrt.Line (prtMaxX, mCurrenty_Top)-(prtMaxX, prtMaxY), prtLineColor

End Sub


'---------------------------------------------------------
Public Sub prtBIA_TVAFAC_Form_10_Col()
'---------------------------------------------------------
Dim X As String, K As Integer, K2 As Integer

XPrt.DrawWidth = 2
XPrt.Line (prtMinX, mCurrenty_Top)-(prtMinX, prtMaxY), prtLineColor
XPrt.Line (prtMinX + 1750, mCurrenty_Top)-(prtMinX + 1750, prtMaxY), prtLineColor
XPrt.Line (prtMinX + 4050, mCurrenty_Top)-(prtMinX + 4050, prtMaxY), prtLineColor
XPrt.Line (prtMinX + 6350, mCurrenty_Top)-(prtMinX + 6350, prtMaxY), prtLineColor
XPrt.Line (prtMinX + 8650, mCurrenty_Top)-(prtMinX + 8650, prtMaxY), prtLineColor
XPrt.Line (prtMinX + 10950, mCurrenty_Top)-(prtMinX + 10950, prtMaxY), prtLineColor
XPrt.Line (prtMinX + 13250, mCurrenty_Top)-(prtMinX + 13250, prtMaxY), prtLineColor
XPrt.Line (prtMaxX, mCurrenty_Top)-(prtMaxX, prtMaxY), prtLineColor

End Sub



'---------------------------------------------------------
Public Sub prtBIA_TVAFAC_Form_3_Col()
'---------------------------------------------------------
Dim X As String, K As Integer, K2 As Integer

XPrt.DrawWidth = 2
XPrt.Line (prtMinX, mCurrenty_Top)-(prtMinX, prtMaxY), prtLineColor
XPrt.Line (prtMinX + 7450, mCurrenty_Top)-(prtMinX + 7450, prtMaxY), prtLineColor
XPrt.Line (prtMinX + 8950, mCurrenty_Top)-(prtMinX + 8950, prtMaxY), prtLineColor
XPrt.Line (prtMinX + 10450, mCurrenty_Top)-(prtMinX + 10450, prtMaxY), prtLineColor
XPrt.Line (prtMinX + 11550, mCurrenty_Top)-(prtMinX + 11550, prtMaxY), prtLineColor
XPrt.Line (prtMinX + 14050, mCurrenty_Top)-(prtMinX + 14050, prtMaxY), prtLineColor
XPrt.Line (prtMaxX, mCurrenty_Top)-(prtMaxX, prtMaxY), prtLineColor
XPrt.Line (prtMinX + 15450, mCurrenty_Top)-(prtMinX + 15450, prtMaxY), prtLineColor

End Sub

'---------------------------------------------------------
Public Sub prtBIA_TVAFAC_Form_9_Col()
'---------------------------------------------------------
Dim X As String, K As Integer, K2 As Integer

XPrt.DrawWidth = 2
XPrt.Line (prtMinX, mCurrenty_Top)-(prtMinX, prtMaxY), prtLineColor
XPrt.Line (prtMinX + 7350, mCurrenty_Top)-(prtMinX + 7350, prtMaxY), prtLineColor
XPrt.Line (prtMinX + 8550, mCurrenty_Top)-(prtMinX + 8550, prtMaxY), prtLineColor
XPrt.Line (prtMinX + 9750, mCurrenty_Top)-(prtMinX + 9750, prtMaxY), prtLineColor
XPrt.Line (prtMinX + 10200, mCurrenty_Top)-(prtMinX + 10200, prtMaxY), prtLineColor
XPrt.Line (prtMinX + 12000, mCurrenty_Top)-(prtMinX + 12000, prtMaxY), prtLineColor
XPrt.Line (prtMinX + 13850, mCurrenty_Top)-(prtMinX + 13850, prtMaxY), prtLineColor
XPrt.Line (prtMinX + 15200, mCurrenty_Top)-(prtMinX + 15200, prtMaxY), prtLineColor
XPrt.Line (prtMaxX, mCurrenty_Top)-(prtMaxX, prtMaxY), prtLineColor

End Sub


'---------------------------------------------------------
Public Sub prtBIA_TVAFAC_Form_4_Col()
'---------------------------------------------------------
Dim X As String, K As Integer, K2 As Integer

XPrt.DrawWidth = 2
XPrt.Line (prtMinX, mCurrenty_Top)-(prtMinX, prtMaxY), prtLineColor
XPrt.Line (prtMinX + 6950, mCurrenty_Top)-(prtMinX + 6950, prtMaxY), prtLineColor
XPrt.Line (prtMinX + 7450, mCurrenty_Top)-(prtMinX + 7450, prtMaxY), prtLineColor
XPrt.Line (prtMaxX, mCurrenty_Top)-(prtMaxX, prtMaxY), prtLineColor

End Sub

Public Sub prtBIA_TVAFAC_NewLine(lK As Integer)
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
If XPrt.CurrentY > me_prtMaxY Then
    Select Case lK
        Case 1: prtBIA_TVAFAC_Form_1_Col ("C")
                'XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY), prtLineColor
                'XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
                XPrt.CurrentX = prtMinX + 10300: XPrt.Print "---/---";
                If Page_No = 1 Then
                    XPrt.CurrentY = mCurrentY_Page
                    XPrt.FontSize = 8
                    XPrt.CurrentX = prtMinX + 10000: XPrt.Print "Page : " & Page_No;
                End If

         Case 2:  prtBIA_TVACOM_Form_2_Col
         Case 3:  prtBIA_TVAFAC_Form_3_Col
         Case 4:  prtBIA_TVAFAC_Form_4_Col
         Case 9:  prtBIA_TVAFAC_Form_9_Col
         Case 10:  prtBIA_TVAFAC_Form_10_Col
    End Select
    frmElpPrt.prtNewPage
    Select Case lK
        Case 1: prtBIA_TVAFAC_Form_1
        Case 2: prtBIA_TVACOM_Form_2
        Case 3: prtBIA_TVAFAC_Form_3
        Case 4: prtBIA_TVAFAC_Form_4
        Case 10: prtBIA_TVAFAC_Form_10
    End Select
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight

End If

End Sub
'---------------------------------------------------------
Public Sub prtBIA_TVAFAC_Form_1()
'---------------------------------------------------------
Dim X As String
Dim curX As Currency
blnNewPage = True

XPrt.DrawWidth = 1

XPrt.CurrentY = prtMinY + prtlineHeight
XPrt.FontSize = 10

'XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'XPrt.CurrentX = prtMinX + 50: XPrt.FontBold = False
'XPrt.Print "identifiant TVA";
'XPrt.CurrentX = prtMinX + 1500: XPrt.FontBold = True
'XPrt.Print ": " & paramSOC_TVA_Intracommunautaire;

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinX + 50: XPrt.FontBold = False
XPrt.ForeColor = RGB(0, 123, 141)
XPrt.Print "Facture N° ";
XPrt.CurrentX = prtMinX + 1500: XPrt.FontBold = True
XPrt.Print ": ";
XPrt.ForeColor = prtForeColor
XPrt.Print mYTVAFAC0.TVAFACFACN;

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinX + 50: XPrt.FontBold = False
XPrt.ForeColor = RGB(0, 123, 141)
XPrt.Print "émise le ";
XPrt.CurrentX = prtMinX + 1500: XPrt.FontBold = True
If mYTVAFAC0.TVAFACDTR > 0 Then
    XPrt.Print ": ";
    XPrt.ForeColor = prtForeColor
    XPrt.Print dateIBM10(mYTVAFAC0.TVAFACDTR, True);
End If

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinX + 50: XPrt.FontBold = False
XPrt.ForeColor = RGB(0, 123, 141)
XPrt.Print "N/Réf";
XPrt.CurrentX = prtMinX + 1500: XPrt.FontBold = True
XPrt.Print ": ";
XPrt.ForeColor = prtForeColor
If Trim(mZCLIENA0.CLIENARES) <> "" Then XPrt.Print mZCLIENA0.CLIENARES & " - ";
XPrt.Print mYTVAFAC0.TVAFACCLIC & " " & mYTVAFAC0.TVAFACCLI & " - " & mYTVAFAC0.TVAFACCLIP;


If mYTVAFAC0.TVAFACSTA <> "F" Then
    XPrt.FontSize = 10: XPrt.FontBold = True
    XPrt.CurrentX = 5700
    XPrt.ForeColor = vbRed
    XPrt.Print "DOCUMENT INTERNE"
    'XPrt.Print "Paris, le  " & dateImp10(DSys);
    XPrt.ForeColor = prtForeColor
End If

'If (Trim(mYTVAFAC0.TVAFACCLIT) = "" And mYTVAFAC0.TVAFACSTA = "V") _
'Or mYTVAFAC0.TVAFACSTA = "2" Then
If mTVANIFCLIT_Pays Then
    XPrt.CurrentY = 2400 ''XPrt.CurrentY + prtlineHeight * 3
    XPrt.CurrentX = prtMinX + 50: XPrt.FontBold = False
    XPrt.ForeColor = RGB(0, 123, 141)
    XPrt.Print "identifiant TVA";
    XPrt.CurrentX = prtMinX + 1500: XPrt.FontBold = True
    XPrt.Print ": ";
    XPrt.ForeColor = prtForeColor
    XPrt.Print TVANIFCLIT_Format(mYTVAFAC0.TVAFACCLIT);
End If

'XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 6



Call prtAdresse_Enveloppe(mZADRESS0)

XPrt.ForeColor = RGB(0, 123, 141)

XPrt.FontSize = 13: XPrt.FontBold = True: XPrt.FontUnderline = True
' Titre de l'édition
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 6
frmElpPrt.prtCentré prtMedX, "Justificatif de prestations de services fournies"
XPrt.FontBold = False:: XPrt.FontUnderline = False
XPrt.FontSize = 7
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
'frmElpPrt.prtCentré prtMedX, "En application de la directive européenne 2001/115/CE du 20-12-2001 relative aux règles de facturation en matière de TVA"
'modification DR du 19/03/2019
'frmElpPrt.prtCentré prtMedX, "En application des directives européennes 2006/11/CE, 2008/8/CE et 2008/9/CE relatives à la TVA"
frmElpPrt.prtCentré prtMedX, "En application des directives européennes 2006/112/CE, 2008/8/CE et 2008/9/CE relatives à la TVA"
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'frmElpPrt.prtCentré prtMedX, "(articles 289 et 289 bis modifiés du CGI)"
frmElpPrt.prtCentré prtMedX, "(articles 259 à 259D du CGI)"


' Entête de colonne
XPrt.ForeColor = prtForeColor
XPrt.DrawWidth = 1
XPrt.FontSize = 8: XPrt.FontBold = True
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight

Page_No = Page_No + 1
If Page_No = 1 Then
    mCurrentY_Page = XPrt.CurrentY
Else
    XPrt.FontSize = 8
    XPrt.CurrentX = prtMinX + 10000: XPrt.Print "Page : " & Page_No;
End If
'---------------------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
prtFillColor = RGB(0, 123, 141)
XPrt.ForeColor = vbWhite
Call frmElpPrt.prtTrame_Color(prtMinX, XPrt.CurrentY, prtMaxX, XPrt.CurrentY + prtlineHeight + 50, " ")
'---------------------------------------------------------

'---------------------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + 50
mCurrenty_Top = XPrt.CurrentY

XPrt.CurrentX = prtMinX + 80: XPrt.Print "Date";
XPrt.CurrentX = prtMinX + 1100: XPrt.Print "Opération";
XPrt.CurrentX = prtMinX + 2450: XPrt.Print "Prestation";
XPrt.CurrentX = prtMinX + 6200: XPrt.Print "Q.";
XPrt.CurrentX = prtMinX + 6700: XPrt.Print "Prix unitaire";
XPrt.CurrentX = prtMinX + 7800: XPrt.Print "Dev";
XPrt.CurrentX = prtMinX + 8300: XPrt.Print "Tx";
XPrt.CurrentX = prtMinX + 8850: XPrt.Print "Montant HT €";
XPrt.CurrentX = prtMinX + 10350: XPrt.Print "TVA €";
XPrt.CurrentY = XPrt.CurrentY + 100
XPrt.ForeColor = vbBlack
prtFillColor = prtFillColor_Standard

XPrt.FontBold = False
XPrt.FontSize = 8

End Sub

'---------------------------------------------------------
Public Sub prtBIA_TVACOM_Form_2()
'---------------------------------------------------------
Dim X As String
Dim curX As Currency
blnNewPage = True

XPrt.DrawWidth = 1

XPrt.CurrentY = prtMinY + prtlineHeight
XPrt.FontSize = 8
mCurrenty_Top = prtMinY
Call frmElpPrt.prtTrame(prtMinX, prtMinY, prtMaxX, prtMinY + prtHeaderHeight, "B", 240)
'---------------------------------------------------------
XPrt.CurrentY = prtMinY + 50

XPrt.ForeColor = prtForeColor_Header

XPrt.CurrentX = prtMinX + 80: XPrt.Print "Tiers";
XPrt.CurrentX = prtMinX + 7100: XPrt.Print "Opération";
XPrt.CurrentX = prtMinX + 8600: XPrt.Print "D.TRT";
XPrt.CurrentX = prtMinX + 9500: XPrt.Print "Commission";
XPrt.CurrentX = prtMinX + 11200: XPrt.Print "TVA";
XPrt.CurrentX = prtMinX + 12000: XPrt.Print "Commission";
XPrt.CurrentX = prtMinX + 13000: XPrt.Print "Rés";
XPrt.CurrentX = prtMinX + 13400: XPrt.Print "Tax";
XPrt.CurrentX = prtMinX + 14100: XPrt.Print "N° facture";
XPrt.CurrentY = XPrt.CurrentY + 100
prtFillColor = prtFillColor_Standard

XPrt.FontBold = False
XPrt.FontSize = 8

End Sub
'---------------------------------------------------------
Public Sub prtBIA_TVAFAC_Form_10()
'---------------------------------------------------------
Dim X As String
Dim curX As Currency
blnNewPage = True

XPrt.DrawWidth = 1

XPrt.CurrentY = prtMinY + prtlineHeight
XPrt.FontSize = 8
mCurrenty_Top = prtMinY
Call frmElpPrt.prtTrame(prtMinX, prtMinY, prtMaxX, prtMinY + prtHeaderHeight, "B", 240)
'---------------------------------------------------------
XPrt.CurrentY = prtMinY + 50

XPrt.ForeColor = prtForeColor_Header

XPrt.CurrentX = prtMinX + 50: XPrt.Print "Srv";
XPrt.CurrentX = prtMinX + 550: XPrt.Print "Opé";
XPrt.CurrentX = prtMinX + 1050: XPrt.Print "Nature";
XPrt.CurrentX = prtMinX + 2500: XPrt.Print "BIA : Débit (nb,mt€)";
XPrt.CurrentX = prtMinX + 4700: XPrt.Print "BIA : Crédit (nb,mt€)";
XPrt.CurrentX = prtMinX + 7100: XPrt.Print "CDO : Débit (nb,mt€)";
XPrt.CurrentX = prtMinX + 9300: XPrt.Print "CDO : Crédit (nb,mt€)";
XPrt.CurrentX = prtMinX + 11700: XPrt.Print "TRF : Débit (nb,mt€)";
XPrt.CurrentX = prtMinX + 14000: XPrt.Print "TRF : Crédit (nb,mt€)";
XPrt.CurrentY = XPrt.CurrentY + 100
prtFillColor = prtFillColor_Standard

XPrt.FontBold = False
XPrt.FontSize = 8

End Sub

'---------------------------------------------------------
Public Sub prtBIA_TVAFAC_Form_3()
'---------------------------------------------------------
Dim X As String
Dim curX As Currency
blnNewPage = True

XPrt.DrawWidth = 1

XPrt.CurrentY = prtMinY + prtlineHeight
XPrt.FontSize = 8
mCurrenty_Top = prtMinY
Call frmElpPrt.prtTrame(prtMinX, prtMinY, prtMaxX, prtMinY + prtHeaderHeight, "B", 240)
'---------------------------------------------------------
XPrt.CurrentY = prtMinY + 50
XPrt.ForeColor = prtForeColor_Header


XPrt.CurrentX = prtMinX + 80: XPrt.Print "Tiers";
XPrt.CurrentX = prtMinX + 7500: XPrt.Print "Pays";
XPrt.CurrentX = prtMinX + 8200: XPrt.Print "id TVA";
XPrt.CurrentX = prtMinX + 9200: XPrt.Print "D.TRT";
XPrt.CurrentX = prtMinX + 9850: XPrt.Print "N° fac";
XPrt.CurrentX = prtMinX + 10700: XPrt.Print "Mt exonéré";
XPrt.CurrentX = prtMinX + 12200: XPrt.Print "Mt taxable";
XPrt.CurrentX = prtMinX + 13700: XPrt.Print "TVA";
XPrt.CurrentX = prtMinX + 15000: XPrt.Print "Total";
XPrt.CurrentY = XPrt.CurrentY + 100
prtFillColor = prtFillColor_Standard

XPrt.FontBold = False
XPrt.FontSize = 8

End Sub
'---------------------------------------------------------
Public Sub prtBIA_TVAFAC_Form_9()
'---------------------------------------------------------
Dim X As String
Dim curX As Currency
blnNewPage = True

XPrt.DrawWidth = 1

XPrt.CurrentY = prtMinY + prtlineHeight
XPrt.FontSize = 8
mCurrenty_Top = prtMinY
Call frmElpPrt.prtTrame(prtMinX, prtMinY, prtMaxX, prtMinY + prtHeaderHeight, "B", 240)
'---------------------------------------------------------
XPrt.CurrentY = prtMinY + 50
XPrt.ForeColor = prtForeColor_Header


XPrt.CurrentX = prtMinX + 80: XPrt.Print "Tiers";
XPrt.CurrentX = prtMinX + 7450: XPrt.Print "Date";
XPrt.CurrentX = prtMinX + 9000: XPrt.Print "Facture";
XPrt.CurrentX = prtMinX + 9800: XPrt.Print "Pays";
XPrt.CurrentX = prtMinX + 10300: XPrt.Print "Identification TVA";
XPrt.CurrentX = prtMinX + 12700: XPrt.Print "montant taxable";
XPrt.CurrentX = prtMinX + 14200: XPrt.Print "montant TVA";
XPrt.CurrentX = prtMinX + 15300: XPrt.Print "n° xml";
XPrt.CurrentY = XPrt.CurrentY + 100
prtFillColor = prtFillColor_Standard

XPrt.FontBold = False
XPrt.FontSize = 8

End Sub

'---------------------------------------------------------
Public Sub prtBIA_TVAFAC_Form_4()
'---------------------------------------------------------
Dim X As String
Dim curX As Currency
blnNewPage = True

XPrt.DrawWidth = 1

XPrt.CurrentY = prtMinY + prtlineHeight
XPrt.FontSize = 8
mCurrenty_Top = prtMinY
Call frmElpPrt.prtTrame(prtMinX, prtMinY, prtMaxX, prtMinY + prtHeaderHeight, "B", 240)
'---------------------------------------------------------
XPrt.CurrentY = prtMinY + 50
XPrt.ForeColor = prtForeColor_Header


XPrt.CurrentX = prtMinX + 80: XPrt.Print "Tiers";
XPrt.CurrentX = prtMinX + 7000: XPrt.Print "Pays";
XPrt.CurrentX = prtMinX + 7500: XPrt.Print "numéro TVA intracommunautaire";
XPrt.CurrentY = XPrt.CurrentY + 100
prtFillColor = prtFillColor_Standard

XPrt.FontBold = False
XPrt.FontSize = 8

End Sub

Public Sub prtBIA_TVAFAC_Init_1(lYTVAFAC0 As typeYTVAFAC0, lZADRESS0 As typeZADRESS0, lZCLIENA0 As typeZCLIENA0, lTVANIFCLIT_Pays As Boolean)
'---------------------------------------------------------
mYTVAFAC0 = lYTVAFAC0
mZADRESS0 = lZADRESS0
mZCLIENA0 = lZCLIENA0
mTVANIFCLIT_Pays = lTVANIFCLIT_Pays
Page_No = 0
End Sub





