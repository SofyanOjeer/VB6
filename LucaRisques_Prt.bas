Attribute VB_Name = "prtLucaRisques"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim Colonne As Integer
Dim Height6_5 As Integer
Dim NbLg As Integer

Dim mRFBENF As String, mCDCPCO As String, maxRFBENF As String
Dim arrAM(12) As String, strAM(12) As String
Dim col1 As Integer, col2 As Integer, ColAM As Integer
Dim Line1 As Integer, Line2 As Integer, Line3 As Integer, Line4 As Integer
Dim Line5 As Integer, Line6 As Integer, Line7 As Integer, Line8 As Integer
 
Public DTCENTenCours As String
Dim blnNewPage As Boolean
Dim col1Title As String
Dim Fct As String
Dim totalLrRisque As typeLrRisque


Public LrCdr_CV1 As typeCV, LrCdr_CV2 As typeCV, LrCdr_CV3 As typeCV

Dim arrExport_Rubriques(20) As String * 2
Dim arrExport_CDR(20, 12) As typeExport_CDR
Type typeExport_CDR
    mtBIA    As Currency
    mtBDF   As Currency
End Type
    

'---------------------------------------------------------
Public Sub prtLucaRisques_Form()
'---------------------------------------------------------
Dim K As Integer

XPrt.DrawWidth = 3
Call frmElpPrt.prtTrame(prtMinX, prtMinY, prtMaxX, prtMinY + prtHeaderHeight, "B", 235)

'-----------------------------------------------------
XPrt.FontBold = False
XPrt.FontSize = 8
XPrt.CurrentY = prtMinY + (prtHeaderHeight - XPrt.TextHeight("X")) / 2
XPrt.FontBold = True
XPrt.CurrentX = prtMinX
XPrt.Print col1Title;

For K = 1 To 12
    frmElpPrt.prtCentré prtMinX + col1 + col2 * (K - 0.5), strAM(K)
Next K


XPrt.FontBold = False

'---------------------------------------------------------
XPrt.FontSize = 6
XPrt.CurrentY = prtMinY + prtHeaderHeight + prtlineHeight - XPrt.TextHeight("X")

NbLg = 0

End Sub


'---------------------------------------------------------
Public Sub prtLucaRisques_FormRubrique()
'---------------------------------------------------------
prtLucaRisques_Form

Line1 = XPrt.CurrentY
XPrt.CurrentX = prtMinX: XPrt.FontBold = True
XPrt.Print "Crédits à court terme";
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY - 60, prtMaxX, XPrt.CurrentY + prtlineHeight - 50, " ")
XPrt.CurrentX = prtMinX: XPrt.FontBold = False
XPrt.Print "CO - Comptes ordinaires débiteurs";
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY - 60, prtMaxX, XPrt.CurrentY + prtlineHeight - 50, " ")
XPrt.CurrentX = prtMinX: XPrt.FontBold = False
XPrt.Print "CA - Autres crédits";
XPrt.CurrentX = prtMinX: XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.Print "CC - Dont crédits liés à des";
XPrt.CurrentX = prtMinX: XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "     créances commerciales";
XPrt.CurrentX = prtMinX: XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.Print "CD - Dont crédits en devises";

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
Line2 = XPrt.CurrentY
XPrt.CurrentX = prtMinX: XPrt.FontBold = True
XPrt.Print "Crédits à moyen et long terme";
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY - 60, prtMaxX, XPrt.CurrentY + prtlineHeight - 50, " ")
XPrt.CurrentX = prtMinX: XPrt.FontBold = False
XPrt.Print "TE - Crédits à l'exportation";
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY - 60, prtMaxX, XPrt.CurrentY + prtlineHeight - 50, " ")
XPrt.CurrentX = prtMinX: XPrt.FontBold = False
XPrt.Print "TA - Autres crédits";
XPrt.CurrentX = prtMinX: XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.Print "TD - Dont crédits en devises";

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
Line3 = XPrt.CurrentY
XPrt.CurrentX = prtMinX: XPrt.FontBold = True
XPrt.Print "Titrisation";
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY - 60, prtMaxX, XPrt.CurrentY + prtlineHeight - 50, " ")
XPrt.CurrentX = prtMinX: XPrt.FontBold = False
XPrt.Print "IT - Crédits titrisés";

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
Line4 = XPrt.CurrentY
XPrt.CurrentX = prtMinX: XPrt.FontBold = True
XPrt.Print "Avals et cautions";
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY - 60, prtMaxX, XPrt.CurrentY + prtlineHeight - 50, " ")
XPrt.CurrentX = prtMinX: XPrt.FontBold = False
XPrt.Print "AC - Avals et cautions";

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
Line5 = XPrt.CurrentY
XPrt.CurrentX = prtMinX: XPrt.FontBold = True
XPrt.Print "Ouverture de crédits";
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY - 60, prtMaxX, XPrt.CurrentY + prtlineHeight - 50, " ")
XPrt.CurrentX = prtMinX: XPrt.FontBold = False
XPrt.Print "OC - Autorisations disponibles sur";
XPrt.CurrentX = prtMinX: XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "     crédits confirmés à terme";
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY - 60, prtMaxX, XPrt.CurrentY + prtlineHeight - 50, " ")
XPrt.CurrentX = prtMinX: XPrt.FontBold = False
XPrt.Print "OD - Fraction disponible sur";
XPrt.CurrentX = prtMinX: XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "  ouverture de crédits documentaires";

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
Line6 = XPrt.CurrentY
XPrt.CurrentX = prtMinX: XPrt.FontBold = True
XPrt.Print "Crédit et location avec option d'achat";
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY - 60, prtMaxX, XPrt.CurrentY + prtlineHeight - 50, " ")
XPrt.CurrentX = prtMinX: XPrt.FontBold = False
XPrt.Print "BM - Oprérations mobilières";
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY - 60, prtMaxX, XPrt.CurrentY + prtlineHeight - 50, " ")
XPrt.CurrentX = prtMinX: XPrt.FontBold = False
XPrt.Print "BI - Opérations immobilières ";

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
Line7 = XPrt.CurrentY
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY - 60, prtMaxX, XPrt.CurrentY + prtlineHeight - 50, " ", 240)
XPrt.CurrentX = prtMinX: XPrt.FontBold = True
XPrt.Print "Total";

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
Line8 = XPrt.CurrentY
Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY - 60, prtMaxX, XPrt.CurrentY + prtlineHeight - 50, " ")
XPrt.CurrentX = prtMinX: XPrt.FontBold = True
XPrt.Print "Cotation Banque de France";

XPrt.FontBold = False

End Sub



'---------------------------------------------------------
 Public Sub prtLucaRisquesX(Msg As String)
'---------------------------------------------------------
Dim K As Integer, K1 As Integer, K2 As Integer
Dim X As String
NbLg = 0
Set XPrt = Printer
frmElpPrt.Show vbModeless

prtOrientation = vbPRORLandscape
prtTitleText = ""
prtPgmName = "prtLucaRisques"
prtTitleUsr = "Centralisation des risques bancaires ( en Milliers )" ' usrName
recLrSgnBnf.RFBENF = "TOTAL"
recLrSgnBnf.Method = "Seek="
dbLrSgnBnf_ReadZ recLrSgnBnf
prtTitleUsr = recLrSgnBnf.NOMBNF

prtLineNb = 1
prtlineHeight = 200
prtHeaderHeight = 300
Colonne = 0
XPrt.FontSize = 6
Height6_5 = XPrt.TextHeight("X")
XPrt.FontSize = 5
Height6_5 = Height6_5 - XPrt.TextHeight("X")
prtFormType = "STD"
frmElpPrt.prtInit
blnNewPage = False
mCDCPCO = ""
prtLucaRisque_Init
Fct = mId$(Msg, 1, 6)
Select Case Fct
    Case "RFBENF": mRFBENF = mId$(Msg, 7, 16): maxRFBENF = mRFBENF: prtLucaRisques_Rubriques
    Case "RFBALL": mRFBENF = "": maxRFBENF = "999999999999999": prtLucaRisques_Rubriques
    Case "TOTAL ": mRFBENF = "TOTAL": maxRFBENF = mRFBENF: prtLucaRisques_Rubriques
    Case "COTBDF": mRFBENF = "TOTAL0   ": maxRFBENF = "TOTALZZZZ": prtLucaRisques_Rubriques
    Case "TOTBDF": mRFBENF = "TOTAL0   ": maxRFBENF = "TOTALZZZZ": prtLucaRisques_TOTBDF
    Case "LSTBNF": mRFBENF = mId$(Msg, 7, 16): maxRFBENF = mRFBENF: prtLucaRisques_LSTBNF
    Case "LSTBPM": mRFBENF = mId$(Msg, 7, 16): maxRFBENF = mRFBENF: prtLucaRisques_LSTBPM
    Case "EXPORT": mRFBENF = mId$(Msg, 7, 16): maxRFBENF = mRFBENF: prtLucaRisques_ExportMonitor
End Select


If blnNewPage Then frmElpPrt.prtEndDoc
frmElpPrt.Hide

End Sub
'---------------------------------------------------------
Public Sub prtLucaRisques_Line()
'---------------------------------------------------------
Dim X As String, K As Integer

For K = 1 To 12
    If recLrRisque.DTCENT1 = arrAM(K) Then Exit For
Next K

ColAM = col1 + col2 * K - 150

XPrt.CurrentY = Line1 + prtlineHeight
''prtLucaRisques_Mt  recLrRisque.MT01
prtLucaRisques_Ratio recLrRisque.MT01, recLrRetris.MT01

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
'prtLucaRisques_Mt  recLrRisque.MT02
prtLucaRisques_Ratio recLrRisque.MT02, recLrRetris.MT02

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
'prtLucaRisques_Mt  recLrRisque.MT03
prtLucaRisques_Ratio recLrRisque.MT03, recLrRetris.MT03

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
'prtLucaRisques_Mt  recLrRisque.MT04
prtLucaRisques_Ratio recLrRisque.MT04, recLrRetris.MT04


XPrt.CurrentY = Line2 + prtlineHeight
'prtLucaRisques_Mt  recLrRisque.MT05
prtLucaRisques_Ratio recLrRisque.MT05, recLrRetris.MT05

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
'prtLucaRisques_Mt  recLrRisque.MT06
prtLucaRisques_Ratio recLrRisque.MT06, recLrRetris.MT06

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
'prtLucaRisques_Mt  recLrRisque.MT07
prtLucaRisques_Ratio recLrRisque.MT07, recLrRetris.MT07


XPrt.CurrentY = Line3 + prtlineHeight
'prtLucaRisques_Mt  recLrRisque.MT13
prtLucaRisques_Ratio recLrRisque.MT13, recLrRetris.MT13

XPrt.CurrentY = Line4 + prtlineHeight
'prtLucaRisques_Mt  recLrRisque.MT10
prtLucaRisques_Ratio recLrRisque.MT10, recLrRetris.MT10


XPrt.CurrentY = Line5 + prtlineHeight
'prtLucaRisques_Mt  recLrRisque.MT08
prtLucaRisques_Ratio recLrRisque.MT08, recLrRetris.MT08

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
'prtLucaRisques_Mt  recLrRisque.MT09
prtLucaRisques_Ratio recLrRisque.MT09, recLrRetris.MT09


XPrt.CurrentY = Line6 + prtlineHeight
'prtLucaRisques_Mt  recLrRisque.MT11
prtLucaRisques_Ratio recLrRisque.MT11, recLrRetris.MT11

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
'prtLucaRisques_Mt  recLrRisque.MT12
prtLucaRisques_Ratio recLrRisque.MT12, recLrRetris.MT12

XPrt.CurrentY = Line7 + prtlineHeight
'prtLucaRisques_Mt  recLrRisque.MTTOTAL
prtLucaRisques_Ratio recLrRisque.MTTOTAL, recLrRetris.MTTOTAL
'-------------------------------------------------------------
XPrt.CurrentY = Line8
XPrt.CurrentX = ColAM - XPrt.TextWidth(recLrRetris.COTBDF)
XPrt.Print recLrRetris.COTBDF;


End Sub
'---------------------------------------------------------
Public Sub prtLucaRisques_ExportLine()
'---------------------------------------------------------
Dim X As String, K As Integer
For K = 1 To 12
    If recLrRisque.DTCENT1 = arrAM(K) Then Exit For
Next K
arrExport_CDR(0, K).mtBIA = recLrRisque.MTTOTAL
arrExport_CDR(0, K).mtBDF = recLrRetris.MTTOTAL
arrExport_CDR(1, K).mtBIA = recLrRisque.MT01
arrExport_CDR(1, K).mtBDF = recLrRetris.MT01
arrExport_CDR(2, K).mtBIA = recLrRisque.MT02
arrExport_CDR(2, K).mtBDF = recLrRetris.MT02
arrExport_CDR(3, K).mtBIA = recLrRisque.MT03
arrExport_CDR(3, K).mtBDF = recLrRetris.MT03
arrExport_CDR(4, K).mtBIA = recLrRisque.MT04
arrExport_CDR(4, K).mtBDF = recLrRetris.MT04
arrExport_CDR(5, K).mtBIA = recLrRisque.MT05
arrExport_CDR(5, K).mtBDF = recLrRetris.MT05
arrExport_CDR(6, K).mtBIA = recLrRisque.MT06
arrExport_CDR(6, K).mtBDF = recLrRetris.MT06
arrExport_CDR(7, K).mtBIA = recLrRisque.MT07
arrExport_CDR(7, K).mtBDF = recLrRetris.MT07
arrExport_CDR(8, K).mtBIA = recLrRisque.MT08
arrExport_CDR(8, K).mtBDF = recLrRetris.MT08
arrExport_CDR(9, K).mtBIA = recLrRisque.MT09
arrExport_CDR(9, K).mtBDF = recLrRetris.MT09
arrExport_CDR(10, K).mtBIA = recLrRisque.MT10
arrExport_CDR(10, K).mtBDF = recLrRetris.MT10
arrExport_CDR(11, K).mtBIA = recLrRisque.MT11
arrExport_CDR(11, K).mtBDF = recLrRetris.MT11
arrExport_CDR(12, K).mtBIA = recLrRisque.MT12
arrExport_CDR(12, K).mtBDF = recLrRetris.MT12
arrExport_CDR(13, K).mtBIA = recLrRisque.MT13
arrExport_CDR(13, K).mtBDF = recLrRetris.MT13

End Sub

'---------------------------------------------------------
Public Sub prtLucaRisques_MTTOTAL()
'---------------------------------------------------------
Dim X As String, K As Integer

For K = 1 To 12
    If recLrRisque.DTCENT1 = arrAM(K) Then Exit For
Next K

ColAM = col1 + col2 * K - 150

prtLucaRisques_Ratio recLrRisque.MTTOTAL, recLrRetris.MTTOTAL

End Sub

Public Sub prtLucaRisques_Trait()
Dim K As Integer
XPrt.Line (prtMinX + col1, prtMinY)-(prtMinX + col1, prtMaxY)
'XPrt.Line (prtMinX + Col1 + Col2 * 3, prtMinY)-(prtMinX + Col1 + Col2 * 3, prtMaxY)
'XPrt.Line (prtMinX + Col1 + Col2 * 6, prtMinY)-(prtMinX + Col1 + Col2 * 6, prtMaxY)
'XPrt.Line (prtMinX + Col1 + Col2 * 9, prtMinY)-(prtMinX + Col1 + Col2 * 9, prtMaxY)
'XPrt.Line (prtMinX + Col1 + Col2 * 12, prtMinY)-(prtMinX + Col1 + Col2 * 12, prtMaxY)
For K = 1 To 11
    Call frmElpPrt.prtTrame(prtMinX + col1 + col2 * K, prtMinY + 20, prtMinX + col1 + col2 * K + 10, prtMaxY - 20, " ", 255)
Next K
End Sub

Public Sub prtLucaRisque_Init()
Dim K As Integer, X8 As String

col1 = 2500
col2 = 1100
X8 = DTCENTenCours & "01"
For K = 1 To 12
    arrAM(K) = mId$(dateElp("MoisAdd", -K + 1, X8), 1, 6)
    strAM(K) = mId$(arrAM(K), 5, 2) & "-" & mId$(arrAM(K), 1, 4)
Next K


End Sub

Public Sub prtLucaRisques_Mt(MT As Currency)
Dim X As String

If MT <> 0 Then
    X = Format$(MT / 100000, "###### ##0.00")

    XPrt.CurrentX = ColAM - XPrt.TextWidth(X)
    XPrt.Print X;
End If

End Sub

Public Sub prtLucaRisques_Ratio(mtBIA As Currency, mtBDF As Currency)
Dim intV As Integer, X As String
If mtBIA <> 0 Or mtBDF <> 0 Then
    XPrt.FontBold = True
    X = Format$(Fix(mtBIA / 100000 + 0.5), "###### ##0")
    XPrt.CurrentX = ColAM - XPrt.TextWidth(X)
    XPrt.Print X;
    XPrt.FontBold = False
    
    If mtBDF = 0 Then
        intV = 0
    Else
        If (mtBIA / mtBDF) > 1.01 Then
            intV = 0
        Else
            intV = Fix(mtBIA / mtBDF * 100 + 0.5)
        End If
    End If
    If intV <= 100 And intV > 0 Then
        XPrt.FontSize = 5
        X = Format(intV, "###") & "%"
        XPrt.CurrentX = ColAM + 350 - XPrt.TextWidth(X)
        XPrt.CurrentY = XPrt.CurrentY + Height6_5
        XPrt.Print X;
        XPrt.CurrentY = XPrt.CurrentY - Height6_5
        XPrt.FontSize = 6
    End If
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    X = Format$(Fix(mtBDF / 100000 + 0.5), "###### ##0")
    XPrt.CurrentX = ColAM - XPrt.TextWidth(X)
    XPrt.Print X;
    XPrt.CurrentY = XPrt.CurrentY - prtlineHeight
End If
End Sub

Public Sub prtLucaRisques_Bénéficiaire()
Dim libCDCCPO As String, intReturn As Integer

If mCDCPCO = "2" Then
    libCDCCPO = "-collectif"
Else
    libCDCCPO = ""
End If

recLrSgnBnf.RFBENF = recLrRisque.RFBENF
recLrSgnBnf.Method = "Seek="
dbLrSgnBnf_ReadZ recLrSgnBnf
prtTitleText = mId$(recLrRisque.RFBENF, 1, 5) & libCDCCPO & " : " & Trim(recLrSgnBnf.NOMBNF)
If Trim(recLrSgnBnf.NSIREN2) <> "" Then prtTitleText = prtTitleText & " ( " & Trim(recLrSgnBnf.NPREFI2) & "_" & Trim(recLrSgnBnf.NSIREN2) & "_ " & Trim(recLrSgnBnf.NSUFFI2) & " )"

If blnNewPage Then
   blnNewPage = False
   frmElpPrt.prtNewPage
Else
    frmElpPrt.prtStd
End If
prtLucaRisques_FormRubrique

recLrRisque.RFBENF = mRFBENF
recLrRisque.CDCPCO = mCDCPCO
recLrRisque.DTCENT1 = arrAM(12)
recLrRisque.Method = "Seek>="
dbLrRisque_Read recLrRisque

Do
    If recLrRisque.DTCENT1 >= arrAM(12) And recLrRisque.DTCENT1 <= arrAM(1) Then
        
        recLrRetris.RFBENF = recLrRisque.RFBENF
        recLrRetris.DTCENT1 = recLrRisque.DTCENT1
        recLrRetris.CDCPCO = recLrRisque.CDCPCO
        recLrRetris.Method = "Seek="
        dbLrRetris_ReadZ recLrRetris
        
        prtLucaRisques_Line
    End If
    
    recLrRisque.Method = "MoveNext"
    intReturn = tableLrRisque_Read(recLrRisque)
  
Loop While intReturn = 0 And mRFBENF = recLrRisque.RFBENF And mCDCPCO = recLrRisque.CDCPCO

prtLucaRisques_Trait
blnNewPage = True
End Sub

Public Sub prtLucaRisques_ExportBénéficiaire()
Dim libCDCCPO As String, intReturn As Integer
Dim K1 As Integer, K2 As Integer, K3 As Integer
Dim wmtBIA As Currency, wmtBDF As Currency

'If mCDCPCO = "2" Then
'    libCDCCPO = "-collectif"
'Else
'    libCDCCPO = ""
'End If

recLrSgnBnf.RFBENF = recLrRisque.RFBENF
recLrSgnBnf.Method = "Seek="
dbLrSgnBnf_ReadZ recLrSgnBnf
prtTitleText = mId$(recLrRisque.RFBENF, 1, 5) & libCDCCPO & " : " & Trim(recLrSgnBnf.NOMBNF)
If Trim(recLrSgnBnf.NSIREN2) <> "" Then prtTitleText = prtTitleText & " ( " & Trim(recLrSgnBnf.NPREFI2) & "_" & Trim(recLrSgnBnf.NSIREN2) & "_ " & Trim(recLrSgnBnf.NSUFFI2) & " )"

'If blnNewPage Then
'   blnNewPage = False
'   frmElpPrt.prtNewPage
'Else
'    frmElpPrt.prtStd
'End If
'prtLucaRisques_FormRubrique
For K1 = 0 To 20
    For K2 = 1 To 12
        arrExport_CDR(K1, K2).mtBIA = 0
        arrExport_CDR(K1, K2).mtBDF = 0
    Next K2
Next K1

recLrRisque.RFBENF = mRFBENF
recLrRisque.CDCPCO = mCDCPCO
recLrRisque.DTCENT1 = arrAM(12)
recLrRisque.Method = "Seek>="
dbLrRisque_Read recLrRisque

Do
    If recLrRisque.DTCENT1 >= arrAM(12) And recLrRisque.DTCENT1 <= arrAM(1) Then
        
        recLrRetris.RFBENF = recLrRisque.RFBENF
        recLrRetris.DTCENT1 = recLrRisque.DTCENT1
        recLrRetris.CDCPCO = recLrRisque.CDCPCO
        recLrRetris.Method = "Seek="
        dbLrRetris_ReadZ recLrRetris
        
        prtLucaRisques_ExportLine
    End If
    
    recLrRisque.Method = "MoveNext"
    intReturn = tableLrRisque_Read(recLrRisque)
  
Loop While intReturn = 0 And mRFBENF = recLrRisque.RFBENF And mCDCPCO = recLrRisque.CDCPCO

'' prtLucaRisques_ExportBénéficiaireDétail

Dim X100 As String * 100
X100 = Space$(100)
K3 = 1
wmtBIA = Round(arrExport_CDR(0, 1).mtBIA, 0)
wmtBDF = Round(arrExport_CDR(0, 1).mtBDF, 0)

X100 = mId$(recLrSgnBnf.RFBENF, 1, 5) _
        & "," & Format$(wmtBIA, "00000000000000") & "," & Format$(wmtBDF, "00000000000000")
        
Print #1, Trim(X100)

'prtLucaRisques_Trait
'blnNewPage = True
End Sub


Public Sub prtLucaRisques_MTTOTALLine()
Dim libCDCCPO As String, intReturn As Integer, blnCOTBDF As Boolean
Dim mCOTBDF As String, mCOTBDF_AM As String, X As String

If XPrt.CurrentY + prtlineHeight * 3 > prtMaxY Then
    prtLucaRisques_Trait
    frmElpPrt.prtNewPage
    prtLucaRisques_Form
End If

'If NbLg = 2 Then
'    NbLg = -1
'    Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY - 60, prtMaxX, XPrt.CurrentY + prtlineHeight * 2 - 50, " ")
'End If

Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY - 60, prtMaxX, XPrt.CurrentY + prtlineHeight - 20, " ")

If mCDCPCO = "2" Then
    libCDCCPO = "_C : "
Else
    libCDCCPO = " : "
End If

recLrSgnBnf.RFBENF = mRFBENF
recLrSgnBnf.Method = "Seek="
dbLrSgnBnf_ReadZ recLrSgnBnf
XPrt.FontSize = 5
XPrt.CurrentY = XPrt.CurrentY + Height6_5
XPrt.CurrentX = prtMinX
XPrt.Print mId$(Trim(recLrSgnBnf.NOMBNF), 1, 40);
XPrt.CurrentY = XPrt.CurrentY - Height6_5
XPrt.FontSize = 6
XPrt.CurrentX = prtMinX
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print mId$(recLrSgnBnf.RFBENF, 1, 5) & libCDCCPO;
XPrt.Print Trim(recLrSgnBnf.NSIREN2);

XPrt.CurrentY = XPrt.CurrentY - prtlineHeight

blnCOTBDF = False: mCOTBDF = "": mCOTBDF_AM = "000000"

recLrRisque.RFBENF = mRFBENF
recLrRisque.CDCPCO = mCDCPCO
recLrRisque.DTCENT1 = arrAM(12)
recLrRisque.Method = "Seek>="
dbLrRisque_Read recLrRisque

Do
    If recLrRisque.DTCENT1 >= arrAM(12) And recLrRisque.DTCENT1 <= arrAM(1) Then
        If Fct = "TOTBDF" Then
            totalLrRisque.RFBENF = "TOTAL"
            totalLrRisque.DTCENT1 = recLrRisque.DTCENT1
            totalLrRisque.CDCPCO = recLrRisque.CDCPCO
            totalLrRisque.Method = "Seek="
            dbLrRisque_ReadZ totalLrRisque
            recLrRetris.MTTOTAL = totalLrRisque.MTTOTAL
            prtLucaRisques_MTTOTAL
       Else
            recLrRetris.RFBENF = recLrRisque.RFBENF
            recLrRetris.DTCENT1 = recLrRisque.DTCENT1
            recLrRetris.CDCPCO = recLrRisque.CDCPCO
            recLrRetris.Method = "Seek="
            dbLrRetris_ReadZ recLrRetris
            prtLucaRisques_MTTOTAL
                                  
            XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
            X = recLrRetris.COTBDF
            XPrt.CurrentX = ColAM + 350 - XPrt.TextWidth(X)
            XPrt.Print X;
            XPrt.CurrentY = XPrt.CurrentY - prtlineHeight * 2
            blnCOTBDF = True

            If recLrRetris.DTCENT1 > mCOTBDF_AM Then
'20000925                blnCOTBDF = True
                mCOTBDF_AM = recLrRetris.DTCENT1
                mCOTBDF = recLrRetris.COTBDF
            End If
        End If
    End If
    
    recLrRisque.Method = "Seek>"
    intReturn = tableLrRisque_Read(recLrRisque)
  
Loop While intReturn = 0 And mRFBENF = recLrRisque.RFBENF And mCDCPCO = recLrRisque.CDCPCO

If blnCOTBDF Then
    blnCOTBDF = True
'    XPrt.CurrentX = Col1 - 150
'    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'    XPrt.Print mCOTBDF;
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
    X = recLrSgnBnf.CDACCO
    XPrt.CurrentX = prtMinX
    XPrt.Print X & "-" & Trim(DicLib(5, X));
    XPrt.CurrentY = XPrt.CurrentY - prtlineHeight
End If

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
NbLg = NbLg + 1
End Sub


Public Sub prtLucaRisques_Rubriques()
Dim intReturn As Integer

col1Title = "Rubrique"

recLrRisque.RFBENF = mRFBENF
recLrRisque.CDCPCO = mCDCPCO
recLrRisque.DTCENT1 = "999999"
recLrRisque.Method = "Seek>="  '"MoveFirst"

Do
    intReturn = tableLrRisque_Read(recLrRisque)
    If intReturn = 0 Then
        If Trim(recLrRisque.RFBENF) > maxRFBENF Then
            intReturn = 1
        Else
            mRFBENF = recLrRisque.RFBENF
            mCDCPCO = recLrRisque.CDCPCO
            prtLucaRisques_Bénéficiaire
            recLrRisque.RFBENF = mRFBENF
            recLrRisque.CDCPCO = mCDCPCO
            recLrRisque.DTCENT1 = "999999"
            recLrRisque.Method = "Seek>="
        End If
     End If
Loop While intReturn = 0

End Sub
Public Sub prtLucaRisques_ExportMonitor()
Dim intReturn As Integer

'col1Title = "Rubrique"

recLrRisque.RFBENF = mRFBENF
recLrRisque.CDCPCO = mCDCPCO
recLrRisque.DTCENT1 = "999999"
recLrRisque.Method = "Seek>="  '"MoveFirst"

arrExport_Rubriques(1) = "CO"
arrExport_Rubriques(2) = "CA"
arrExport_Rubriques(3) = "CC"
arrExport_Rubriques(4) = "CD"
arrExport_Rubriques(5) = "TE"
arrExport_Rubriques(6) = "TA"
arrExport_Rubriques(7) = "TD"
arrExport_Rubriques(8) = "OC"
arrExport_Rubriques(9) = "OD"
arrExport_Rubriques(10) = "AC"
arrExport_Rubriques(11) = "BM"
arrExport_Rubriques(12) = "BI"
arrExport_Rubriques(13) = "IT"

Open "C:\BiaSrv\Export_Cdr.txt" For Output As #1

If Trim(maxRFBENF) = "" Then maxRFBENF = "999999999999999"

Do
    intReturn = tableLrRisque_Read(recLrRisque)
    If intReturn = 0 Then
        If Trim(recLrRisque.RFBENF) > maxRFBENF Then
            intReturn = 1
        Else
            mRFBENF = recLrRisque.RFBENF
            mCDCPCO = recLrRisque.CDCPCO
            prtLucaRisques_ExportBénéficiaire
            recLrRisque.RFBENF = mRFBENF
            recLrRisque.CDCPCO = mCDCPCO
            recLrRisque.DTCENT1 = "999999"
            recLrRisque.Method = "Seek>="
        End If
     End If
Loop While intReturn = 0
Close #1

End Sub

Public Sub prtLucaRisques_LSTBNF()
Dim intReturnLrSort As Integer

col1Title = "Bénéficiaire"
prtTitleText = "Liste des bénéficiaires (par montant total déclaré décroissant)"
frmElpPrt.prtStd
prtLucaRisques_Form
NbLg = 0

recLrSort.Method = "MoveFirst"

Do
    intReturnLrSort = tableLrSort_Read(recLrSort)
    If intReturnLrSort = 0 Then
        mRFBENF = recLrSort.RFBENF: maxRFBENF = mRFBENF
        mCDCPCO = recLrSort.CDCPCO
        prtLucaRisques_MTTOTALLine
        recLrSort.Method = "MoveNext"
    End If
Loop While intReturnLrSort = 0
prtLucaRisques_Trait
blnNewPage = True

End Sub

Public Sub prtLucaRisques_LSTBPM()
Dim intReturnLrSort As Integer
Dim LSTBPM_NB As Integer

col1Title = "Bénéficiaire"
prtTitleText = "Liste de 30 bénéficiaires 'SOCIETE' (par montant total déclaré décroissant)"
frmElpPrt.prtStd
prtLucaRisques_Form
NbLg = 0
LSTBPM_NB = 0
recLrSort.Method = "MoveFirst"

Do
    intReturnLrSort = tableLrSort_Read(recLrSort)
    If intReturnLrSort = 0 Then
        If mId$(recLrSort.RFBENF, 1, 5) >= "30000" And mId$(recLrSort.RFBENF, 1, 5) < "60000" Then
            mRFBENF = recLrSort.RFBENF: maxRFBENF = mRFBENF
            mCDCPCO = recLrSort.CDCPCO
            prtLucaRisques_MTTOTALLine
            LSTBPM_NB = LSTBPM_NB + 1
            If LSTBPM_NB >= 30 Then intReturnLrSort = 1
        End If
        recLrSort.Method = "MoveNext"
    End If
Loop While intReturnLrSort = 0
prtLucaRisques_Trait
blnNewPage = True

End Sub


Public Sub prtLucaRisques_TOTBDF()
Dim intReturn As Integer

col1Title = "Cotation"
prtTitleText = "Répartition par cotation BDF des montants déclarés"
frmElpPrt.prtStd
prtLucaRisques_Form

recLrRisque.RFBENF = mRFBENF
recLrRisque.CDCPCO = mCDCPCO
recLrRisque.DTCENT1 = "999999"
recLrRisque.Method = "Seek>="

Do
    intReturn = tableLrRisque_Read(recLrRisque)
    If intReturn = 0 Then
        If Trim(recLrRisque.RFBENF) > maxRFBENF Then
            intReturn = 1
        Else
            mRFBENF = recLrRisque.RFBENF
            mCDCPCO = recLrRisque.CDCPCO
            prtLucaRisques_MTTOTALLine
            recLrRisque.RFBENF = mRFBENF
            recLrRisque.CDCPCO = mCDCPCO
            recLrRisque.DTCENT1 = "999999"
            recLrRisque.Method = "Seek>="
        End If
     End If
Loop While intReturn = 0

prtLucaRisques_Trait
blnNewPage = True

End Sub


Public Sub prtLucaRisques_ExportBénéficiaireDétail()
Dim K1 As Integer, K2 As Integer, K3 As Integer

Dim X700 As String * 700
X700 = Space$(700)
For K1 = 1 To 13 '20

    For K3 = 1 To 12
        arrExport_CDR(K1, K3).mtBIA = Fix(arrExport_CDR(K1, K3).mtBIA / 100000 + 0.5)
        arrExport_CDR(K1, K3).mtBDF = Fix(arrExport_CDR(K1, K3).mtBDF / 100000 + 0.5)
    Next K3

    X700 = mId$(recLrSgnBnf.RFBENF, 1, 5) & "," & Trim(recLrSgnBnf.NOMBNF) & "," & arrExport_Rubriques(K1) _
            & "," & Format$(arrExport_CDR(K1, 1).mtBIA, "00000000000") & "," & Format$(arrExport_CDR(K1, 1).mtBDF, "00000000000") _
            & "," & Format$(arrExport_CDR(K1, 2).mtBIA, "00000000000") & "," & Format$(arrExport_CDR(K1, 2).mtBDF, "00000000000") _
            & "," & Format$(arrExport_CDR(K1, 3).mtBIA, "00000000000") & "," & Format$(arrExport_CDR(K1, 3).mtBDF, "00000000000") _
            & "," & Format$(arrExport_CDR(K1, 4).mtBIA, "00000000000") & "," & Format$(arrExport_CDR(K1, 4).mtBDF, "00000000000") _
            & "," & Format$(arrExport_CDR(K1, 5).mtBIA, "00000000000") & "," & Format$(arrExport_CDR(K1, 5).mtBDF, "00000000000") _
            & "," & Format$(arrExport_CDR(K1, 6).mtBIA, "00000000000") & "," & Format$(arrExport_CDR(K1, 6).mtBDF, "00000000000") _
            & "," & Format$(arrExport_CDR(K1, 7).mtBIA, "00000000000") & "," & Format$(arrExport_CDR(K1, 7).mtBDF, "00000000000") _
            & "," & Format$(arrExport_CDR(K1, 8).mtBIA, "00000000000") & "," & Format$(arrExport_CDR(K1, 8).mtBDF, "00000000000") _
            & "," & Format$(arrExport_CDR(K1, 9).mtBIA, "00000000000") & "," & Format$(arrExport_CDR(K1, 9).mtBDF, "00000000000") _
            & "," & Format$(arrExport_CDR(K1, 10).mtBIA, "00000000000") & "," & Format$(arrExport_CDR(K1, 10).mtBDF, "00000000000") _
            & "," & Format$(arrExport_CDR(K1, 11).mtBIA, "00000000000") & "," & Format$(arrExport_CDR(K1, 11).mtBDF, "00000000000") _
            & "," & Format$(arrExport_CDR(K1, 12).mtBIA, "00000000000") & "," & Format$(arrExport_CDR(K1, 12).mtBDF, "00000000000")
            
    Print #1, Trim(X700)
Next K1

End Sub
