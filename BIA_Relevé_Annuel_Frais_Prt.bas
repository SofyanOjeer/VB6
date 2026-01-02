Attribute VB_Name = "prtBIA_Relevé_Annuel_Frais"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim I As Integer, solde As Currency, mCurrenty As Integer, Height8_6 As Integer
Dim Line1 As Integer, Line2 As Integer, Line3 As Integer, Line4 As Integer, Line5 As Integer
Dim col1 As Integer, col2 As Integer, col3 As Integer
Dim Col4 As Integer, Col5 As Integer, Col6 As Integer, Col7 As Integer, Col8 As Integer
Dim Col As Integer
Dim K As Integer, K1 As Integer, K2 As Integer, K3 As Integer
Dim X As String
Dim nbLigne As Integer, NbPage As Integer
Dim NbLigneMax As Integer, NbPageMax As Integer
Dim NbImprimé As Integer

Dim valAmjMin As String, valAmjMax As String
Dim IbmAmjMin As String, IbmAmjMax As String

Dim curCumulDébit As Currency, curCumulCrédit As Currency
Dim blnA4_Form As Boolean
Dim blnMsgInfo As Boolean, mMsgInfo As String, mExtraitNuméro As String

Dim xYBIAMVT0 As typeYBIAMVT0, mYBIAMVT0 As typeYBIAMVT0
Dim meYBIACPT0 As typeYBIACPT0

Dim blnCptOrdinaire As Boolean, blnRIB As Boolean, blnMédiateur As Boolean, blnIban As Boolean
Dim blnConvention_Print As Boolean
Dim mRib_Compte As String, mRib_Clé As String, mRib_IbanE As String
Dim mResponsable As String
Dim zZADRESS0 As typeZADRESS0, xZADRESS0 As typeZADRESS0, fiscalZADRESS0 As typeZADRESS0
Dim xZRELEVE0 As typeZRELEVE0
Dim intFile As Integer
Dim blnFRS_Info As Boolean

Dim mYEAR As String
Dim MOUVEMOPE_OPE As String, MOUVEMOPE_Nb As Long, MOUVEMOPE_Mt As Currency

Public imprimanteParDefaut As String

Public Sub prtBIA_Relevé_Annuel_Frais_OpenX()
'---------------------------------------------------------
On Error GoTo prtError

'$20060605_JPL Set XPrt = Printer
'$20060605_JPL frmElpPrt.Show vbModeless
prtBIA_Relevé_Annuel_Frais_OpenX_Reset



Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub

Public Sub prtBIA_Relevé_Annuel_Frais_Close()
'---------------------------------------------------------
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



'---------------------------------------------------------
Public Sub prtBIA_Relevé_Annuel_Frais_Form(Msg As String, lRéférence As String)
'---------------------------------------------------------
Dim X As String
Dim mCurrenty As Integer

prtBIA_Relevé_Annuel_Frais_RIB

' blnConvention_Print = False
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
If Not blnConvention_Print Then
    prtBIA_Relevé_Annuel_Frais_OpenX_Reset_Line " "
Else
    prtBIA_Relevé_Annuel_Frais_OpenX_Reset_Line "M"
    blnConvention_Print = False
    nbLigne = 5
    mCurrenty = XPrt.CurrentY + prtlineHeight * 2.5

    XPrt.ForeColor = RGB(0, 0, 160) 'RGB(0, 123, 141)
    XPrt.FontItalic = True
    XPrt.FontBold = True
    XPrt.FontSize = 9
    XPrt.CurrentY = mCurrenty
    XPrt.CurrentX = col1 + 50
    XPrt.Print "Vous trouverez, ci-dessous, le récapitulatif des frais réglés en " & mYEAR & " relatifs aux produits et services liés à la gestion de votre ";
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = col1 + 50
    XPrt.Print "compte de dépôt, édité conformément aux dispositions de l'article 24 de la loi du 3 janvier 2008 pour le développement de la";
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = col1 + 50
    XPrt.Print "concurrence au service des consommateurs (Loi CHATEL).";
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = col1 + 50
    XPrt.Print "Attention, ce document d'information n'est ni un relevé ni une facture: vous avez déjà réglé les frais mentionnés ci-dessous.";
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = col1 + 50
    XPrt.Print "Vous recevrez désormais ce récapitulatif annuel des frais au mois de janvier de chaque année.";

'------------------------------
MSG_FIN:

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.FontBold = False
    XPrt.FontSize = prtFontSize
    XPrt.ForeColor = vbBlack
    XPrt.FontItalic = False


End If


Call frmElpPrt.prtTrame(Col4, Line3, Col5, Line4, " ", 200)
XPrt.CurrentY = prtMinY + prtlineHeight * 4


'_________________________________________________________________
prtFillColor = RGB(240, 255, 255) 'RGB(0, 123, 141)
XPrt.ForeColor = vbBlack 'vbWhite
Call frmElpPrt.prtTrame_Color(col1, Line2, Col8, Line3, "B")
'---------------------------------------------------------
XPrt.DrawWidth = 3
XPrt.Line (Col4 + 200, Line1)-(Col6 - 200, Line1), prtLineColor
XPrt.DrawWidth = 2
XPrt.Line (col1 + 200, Line2)-(Col8, Line2), prtLineColor

XPrt.Line (col1, Line3)-(Col8, Line3), prtLineColor
XPrt.Line (col1 + 200, Line4)-(Col8, Line4), prtLineColor
XPrt.DrawWidth = 3
XPrt.Line (Col4 + 200, Line5)-(Col6 - 200, Line5), prtLineColor
XPrt.DrawWidth = 2
XPrt.Line (col1, Line2 + 200)-(col1, Line4 - 200), prtLineColor
XPrt.DrawWidth = 1
XPrt.Line (col2, Line2)-(col2, Line4), prtLineColor
XPrt.DrawWidth = 1
XPrt.Line (col3, Line2)-(col3, Line4), prtLineColor
XPrt.DrawWidth = 3
XPrt.Line (Col4, Line1 + 200)-(Col4, Line5 - 200), prtLineColor
XPrt.DrawWidth = 1
XPrt.Line (Col5, Line1)-(Col5, Line5), prtLineColor
XPrt.DrawWidth = 3
XPrt.Line (Col6, Line1 + 200)-(Col6, Line5 - 200), prtLineColor

XPrt.CurrentY = Line2 + 50
XPrt.FontBold = True

XPrt.FontSize = prtFontSize
frmElpPrt.prtCentré (col1 + col2) / 2, "Date"
frmElpPrt.prtCentré (col2 + col3) / 2, "Libellé"
frmElpPrt.prtCentré (col3 + Col4) / 2, "Date Valeur"
frmElpPrt.prtCentré (Col4 + Col5) / 2, "Débit"
frmElpPrt.prtCentré (Col5 + Col6) / 2, "Crédit"

'------------------------
prtFillColor = RGB(250, 255, 255)
XPrt.ForeColor = vbBlack
XPrt.DrawWidth = 2

XPrt.CurrentY = 0
XPrt.CurrentX = 0
XPrt.Circle Step(col1 + 200, Line2 + 200), 200, prtLineColor, 0.5 * Pi, Pi
XPrt.DrawWidth = 3

XPrt.CurrentY = 0
XPrt.CurrentX = 0
XPrt.Circle Step(Col6 - 200, Line1 + 200), 200, prtLineColor, 0, 0.5 * Pi

XPrt.CurrentY = 0
XPrt.CurrentX = 0
XPrt.Circle Step(Col4 + 200, Line1 + 200), 200, prtLineColor, 0.5 * Pi, Pi

XPrt.DrawWidth = 2
XPrt.CurrentY = 0
XPrt.CurrentX = 0
XPrt.Circle Step(col1 + 200, Line4 - 200), 200, prtLineColor, Pi, 1.5 * Pi

XPrt.DrawWidth = 3
XPrt.CurrentY = 0
XPrt.CurrentX = 0
XPrt.Circle Step(Col4 + 200, Line5 - 200), 200, prtLineColor, Pi, 1.5 * Pi



XPrt.CurrentY = 0
XPrt.CurrentX = 0
XPrt.Circle Step(Col6 - 200, Line5 - 200), 200, prtLineColor, 1.5 * Pi, 2 * Pi

'----------------------------------------ligne 1-----------------
XPrt.FontSize = 10
XPrt.CurrentY = prtMinY + prtlineHeight * 10 - XPrt.TextHeight("test")
'----------------------------------1------------
XPrt.FontBold = True

XPrt.CurrentX = 5800
XPrt.Print xZADRESS0.ADRESSRA1;
XPrt.FontBold = False
'-----------------------------------2-------------
''If Trim(xZADRESS0.ADRESSRA2) <> "" Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = 5800
    XPrt.Print xZADRESS0.ADRESSRA2;
''End If
'------------------------------------3---------------
If Trim(xZADRESS0.ADRESSAD1) <> "" Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = 5800
    XPrt.Print xZADRESS0.ADRESSAD1;
End If
'----------------------------------4-------------------
If Trim(xZADRESS0.ADRESSAD2) <> "" Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = 5800
    XPrt.Print xZADRESS0.ADRESSAD2;
End If

'-----------------------------------5------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 5800
XPrt.Print xZADRESS0.ADRESSAD3;
'------------------------------------6------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 5800
If Trim(xZADRESS0.ADRESSCOP) <> "" Then XPrt.Print xZADRESS0.ADRESSCOP & "  ";
XPrt.Print xZADRESS0.ADRESSVIL;
'------------------------------------8------------------
If Trim(xZADRESS0.ADRESSPAY) <> "FRANCE" Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = 5800
    XPrt.Print xZADRESS0.ADRESSPAY;
End If

XPrt.FontSize = 8

XPrt.CurrentY = Line1 - prtlineHeight * 1.5
XPrt.FontBold = True


'Call frmElpPrt.prtTrame(col3, XPrt.CurrentY - 100, Col6, XPrt.CurrentY + prtlineHeight, " ", 240)
prtFillColor = RGB(240, 255, 255)
Call frmElpPrt.prtTrame_Color(col3 - 300, XPrt.CurrentY - 100, Col6 + 50, XPrt.CurrentY + prtlineHeight, " ")
XPrt.CurrentX = col3 - 200
XPrt.ForeColor = vbBlack 'RGB(0, 123, 141)
XPrt.CurrentY = XPrt.CurrentY - Height8_6
XPrt.FontSize = 10
XPrt.Print meYBIACPT0.COMPTEDEV;
XPrt.CurrentY = XPrt.CurrentY + Height8_6
XPrt.FontSize = 8

XPrt.Print "  -  RECAPITULATIF DES FRAIS DE L'ANNEE " & mYEAR;
XPrt.FontBold = True
XPrt.CurrentY = XPrt.CurrentY - Height8_6
XPrt.FontSize = 10
XPrt.CurrentY = XPrt.CurrentY + Height8_6
XPrt.FontBold = False
XPrt.FontSize = 8
XPrt.Print "     page : " & Format$(NbPage, "###");

XPrt.CurrentY = XPrt.CurrentY - Height8_6
XPrt.FontSize = 8
XPrt.ForeColor = vbBlack
'----------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.CurrentX = col1 + 50
XPrt.FontBold = True
'-------------------------------------------------------
XPrt.FontBold = False
XPrt.FontSize = 8

XPrt.FontSize = 8
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = col1 + 50
XPrt.FontBold = True
'---------------------------------------
'---------------------------------------
XPrt.FontBold = False


XPrt.FontSize = prtFontSize

XPrt.CurrentY = Line1 + 50
XPrt.CurrentX = Col4 - 100 - XPrt.TextWidth(Msg)
XPrt.Print Msg;
prtBIA_Relevé_Annuel_Frais_Montant (solde)

nbLigne = 0
blnA4_Form = True

XPrt.CurrentY = Line3 - prtlineHeight + 50

'   denis   '
XPrt.DrawWidth = 4
XPrt.Line (col1, Line2)-(col1, Line2 + 150), 16777215
XPrt.Line (col1, Line2)-(col1 + 150, Line2), 16777215

XPrt.CurrentY = Line3 - prtlineHeight + 50

End Sub

'---------------------------------------------------------
Public Sub prtBIA_Relevé_Annuel_Frais_Line()
'---------------------------------------------------------
Dim X As String, I As Integer, libCV As String, blnCV As Boolean
Dim blnLine2 As Boolean, xLine1 As String, xLine2 As String
Dim kJust As Integer, kMax As Integer
Dim widthCOL3_5 As Integer

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
 
XPrt.FontSize = prtFontSize
XPrt.FontBold = False


XPrt.CurrentX = col1 + 50
XPrt.Print dateImp(Val(xYBIAMVT0.MOUVEMDTR) + 19000000);

XPrt.CurrentX = col3 + 50
XPrt.Print dateImp(Val(xYBIAMVT0.MOUVEMDVA) + 19000000);
prtBIA_Relevé_Annuel_Frais_Montant (xYBIAMVT0.MOUVEMMON)

If Mid$(xYBIAMVT0.MOUVEMANA, 1, 3) = "FRS" And Trim(meYBIACPT0.CLIENASRN) = "" _
And xYBIAMVT0.MOUVEMMON > 0 Then blnFRS_Info = True: XPrt.CurrentX = XPrt.CurrentX + 150: XPrt.Print "#";

XPrt.CurrentX = col2 + 50
If xYBIAMVT0.MOUVEMOPE = "-RM" Then Mid$(xYBIAMVT0.LIBELLIB2, 13, 18) = Space$(18)
xLine1 = Trim(xYBIAMVT0.LIBELLIB1) & " " & Trim(xYBIAMVT0.LIBELLIB2)
xLine2 = Trim(xYBIAMVT0.LIBELLIB3) & " " & Trim(xYBIAMVT0.LIBELLIB4)
X = Replace(xLine1 & " " & xLine2, "  ", " ")
X = Replace(X, "  ", " ")
widthCOL3_5 = (col3 - XPrt.CurrentX)
XPrt.FontSize = 7
If XPrt.TextWidth(X) <= widthCOL3_5 Then
    XPrt.Print X;
Else
    For kMax = Len(X) To 1 Step -1
        xLine1 = Mid$(X, 1, kMax)
        If XPrt.TextWidth(xLine1) <= widthCOL3_5 Then Exit For
    Next kMax
  
    kJust = kMax
    For I = kMax To kMax - 10 Step -1
        If Mid$(X, I, 1) = " " Then kJust = I: Exit For
    Next I
    xLine1 = Mid$(X, 1, kJust)
    xLine2 = Mid$(X, kJust + 1, Len(X) - kJust)
        
    XPrt.Print xLine1; ' & " -";
    If nbLigne = NbLigneMax Then prtBIA_Relevé_Annuel_Frais_Report
    nbLigne = nbLigne + 1
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = col2 + 50
    XPrt.Print xLine2;
End If

XPrt.FontSize = prtFontSize

End Sub


'---------------------------------------------------------
Public Sub prtBIA_Relevé_Annuel_Frais_Line_Total()
'---------------------------------------------------------
Dim X As String, I As Integer, libCV As String, blnCV As Boolean
Dim blnLine2 As Boolean, xLine1 As String, xLine2 As String
Dim kJust As Integer, kMax As Integer
Dim widthCOL3_5 As Integer
Dim xLib As String, xLib_Ope As String

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Line (col1, XPrt.CurrentY)-(Col8, XPrt.CurrentY), prtLineColor
 
XPrt.CurrentY = XPrt.CurrentY + 50
XPrt.FontSize = prtFontSize
XPrt.FontBold = True
XPrt.ForeColor = RGB(0, 123, 141)

nbLigne = nbLigne + 2
XPrt.CurrentX = col2 + 50
If MOUVEMOPE_Nb > 1 Then

    Select Case MOUVEMOPE_OPE
        Case "*C6": xLib_Ope = " écritures liées à des opérations diverses"
        Case "AL1": xLib_Ope = " écritures liées à des rejets de LCR"
        Case "AP1": xLib_Ope = " écritures liées à des rejets de prélévement"
        Case "AT1": xLib_Ope = " écritures liées à des rejets de TIP"
        Case "AV0": xLib_Ope = " écritures liées à des opérations de virement"
        Case "CPT": xLib_Ope = " écritures liées à des opérations de change au comptant"
        Case "ECH": xLib_Ope = " écritures d'agios"
        Case "ENG": xLib_Ope = " écritures liées à des opérations d'engagement"
        Case "FCI": xLib_Ope = " écritures liées à des incidents sur chèque"
        Case "FRS": xLib_Ope = " écritures liées à des services divers"
        Case "PTF": xLib_Ope = " écritures liées à des opérations sur portefeuille"
        Case "REM": xLib_Ope = " écritures liées à des opérations de remise documentaire"
        Case "TRF": xLib_Ope = " écritures liées à des opérations de transfert"
        Case Else: xLib_Ope = xLib & MOUVEMOPE_OPE
    End Select
Else
    Select Case MOUVEMOPE_OPE
        Case "*C6": xLib_Ope = " écriture liée à une opération diverse"
        Case "AL1": xLib_Ope = " écriture liée à un rejet de LCR"
        Case "AP1": xLib_Ope = " écriture liée à un rejet de prélévement"
        Case "AT1": xLib_Ope = " écriture liée à un rejet de TIP"
        Case "AV0": xLib_Ope = " écriture liée à une opération de virement"
        Case "CPT": xLib_Ope = " écriture liée à une opération de change au comptant"
        Case "ECH": xLib_Ope = " écriture d'agios"
        Case "ENG": xLib_Ope = " écriture liée à une opération d'engagement"
        Case "FCI": xLib_Ope = " écriture liée à une incident sur chèque"
        Case "FRS": xLib_Ope = " écriture liée à un service divers"
        Case "PTF": xLib_Ope = " écriture liée à une opération sur portefeuille"
        Case "REM": xLib_Ope = " écriture liée à une opération de remise documentaire"
        Case "TRF": xLib_Ope = " écriture liée à une opération de transfert"
        Case Else: xLib_Ope = xLib & MOUVEMOPE_OPE
    End Select
End If

XPrt.Print " " & MOUVEMOPE_Nb & xLib_Ope;

XPrt.CurrentX = col3 + 50
prtBIA_Relevé_Annuel_Frais_Montant (MOUVEMOPE_Mt)
XPrt.Line (col1, XPrt.CurrentY + prtlineHeight)-(Col8, XPrt.CurrentY + prtlineHeight), prtLineColor

XPrt.FontSize = prtFontSize
XPrt.ForeColor = vbBlack
End Sub



'---------------------------------------------------------
Public Sub prtBIA_Relevé_Annuel_Frais_Montant(MT As Currency)
'---------------------------------------------------------
Dim X As String

XPrt.FontBold = True
X = Format$(Abs(MT), "## ### ### ### ### ##0.00")
XPrt.CurrentX = IIf(MT < 0, Col6, Col5) - 100 - XPrt.TextWidth(X)
XPrt.Print X;
XPrt.FontBold = False

End Sub


Public Sub prtBIA_Relevé_Annuel_Frais_Extrait(lMOUVEMCOM As String, lAMJMin As String, lAMJMax As String, blnCSV As Boolean, lstErr As ListBox, lRELEVEREL As String, lRéférence As String, blnNewPage As Boolean)
'---------------------------------------------------------
Dim rsLocal As ADODB.Recordset, rsW As ADODB.Recordset
Dim xSQL As String
Dim Nb As Integer
Dim CTLAMJ As String
Dim V
Dim wAMJ_Solde As String

mYEAR = Mid$(lAMJMax, 1, 4)
valAmjMin = lAMJMin
valAmjMax = lAMJMax
IbmAmjMin = dateIBM(lAMJMin)
IbmAmjMax = dateIBM(lAMJMax)
wAMJ_Solde = dateElp("Jour", -1, valAmjMin)
rsZADRESS0_Init zZADRESS0
blnNewPage = False
blnFRS_Info = False
MOUVEMOPE_OPE = "": MOUVEMOPE_Nb = 0: MOUVEMOPE_Mt = 0

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIACPT0" _
     & " where COMPTECOM = '" & lMOUVEMCOM & "'"
     
Set rsLocal = Nothing
Set rsLocal = cnsab.Execute(xSQL)
V = rsYBIACPT0_GetBuffer(rsLocal, meYBIACPT0)
If Not IsNull(V) Then
    MsgBox "prtBIA_Relevé_Annuel_Frais_Extrait " & V
    Exit Sub
End If


prtBIA_Relevé_Annuel_Frais_Compte lRELEVEREL


xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIAMVTHF" _
     & " where MOUVEMCOM = '" & lMOUVEMCOM & "'" _
     & " and MOUVEMDTR >= " & IbmAmjMin _
     & " and MOUVEMDTR <= " & IbmAmjMax _
     & " order by MOUVEMOPE,MOUVEMDTR"
     
Set rsLocal = Nothing
Set rsLocal = cnsab.Execute(xSQL)
If rsLocal.EOF Then
' pas de mouvement dans la période
    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIAMVTHF" _
         & " where MOUVEMCOM = '" & lMOUVEMCOM & "'" _
         & " and MOUVEMDTR >= " & IbmAmjMax _
         & " order by MOUVEMOPE,MOUVEMDTR"
    Set rsW = cnsab.Execute(xSQL)
    If rsW.EOF Then
        CTLAMJ = meYBIACPT0.SOLDEDMO
        solde = meYBIACPT0.SOLDECEN
    Else
        CTLAMJ = rsW("MOUVEMDTR")
        solde = rsW("BIAMVTSD0")
   End If
    
Else
    V = rsYBIAMVT0_GetBuffer(rsLocal, mYBIAMVT0)
    If Not IsNull(V) Then
        MsgBox "prtBIA_Relevé_Annuel_Frais_Extrait " & V
        Exit Sub
    End If
    
    xYBIAMVT0 = mYBIAMVT0
    CTLAMJ = mYBIAMVT0.MOUVEMDTR
    solde = 0
End If

prtBIA_Relevé_Annuel_Frais_OpenX_Reset

blnA4_Form = False
mExtraitNuméro = mYEAR

prtFontSize = 8

NbPageMax = 0
NbPage = 1

If blnCSV Then
    Call FEU_ROUGE
    intFile = FreeFile(0)
    Open "C:\TEMP\Extrait_" & lMOUVEMCOM & ".csv" For Output As #intFile
    xYBIAMVT0.MOUVEMDTR = wAMJ_Solde - 19000000
    xYBIAMVT0.MOUVEMDVA = xYBIAMVT0.MOUVEMDTR
    xYBIAMVT0.MOUVEMMON = solde
    xYBIAMVT0.LIBELLIB1 = "Solde initial "
    xYBIAMVT0.LIBELLIB2 = "": xYBIAMVT0.LIBELLIB3 = "": xYBIAMVT0.LIBELLIB4 = ""
    prtYBIAMVT0_CSV
    
    xYBIAMVT0 = mYBIAMVT0

End If
Do Until rsLocal.EOF
        
           If xYBIAMVT0.MOUVEMDTR > IbmAmjMax Then Exit Do
               
               
               
                If xYBIAMVT0.MOUVEMDTR >= IbmAmjMin Then
           
                    If Not blnA4_Form Then
                        prtBIA_Relevé_Annuel_Frais_Form "", lRéférence
                        blnNewPage = True
                    End If
                    If MOUVEMOPE_OPE <> xYBIAMVT0.MOUVEMOPE Then
                        If MOUVEMOPE_Nb <> 0 Then
                            If nbLigne > NbLigneMax - 1 Then prtBIA_Relevé_Annuel_Frais_Report
                            nbLigne = nbLigne + 1
                            prtBIA_Relevé_Annuel_Frais_Line_Total
                            
                        End If
                        MOUVEMOPE_OPE = xYBIAMVT0.MOUVEMOPE: MOUVEMOPE_Nb = 0: MOUVEMOPE_Mt = 0
                    
                    End If
                    
                    If nbLigne = NbLigneMax Then
                        prtBIA_Relevé_Annuel_Frais_Report
                        lstErr.RemoveItem lstErr.ListCount - 1
                        lstErr.AddItem xYBIAMVT0.MOUVEMCOM & " page : " & NbPage
                    End If
                    nbLigne = nbLigne + 1
                    MOUVEMOPE_Nb = MOUVEMOPE_Nb + 1
                    MOUVEMOPE_Mt = MOUVEMOPE_Mt + xYBIAMVT0.MOUVEMMON
                    
                    prtBIA_Relevé_Annuel_Frais_Line
                    
                    If blnCSV Then prtYBIAMVT0_CSV

                    DoEvents: If prtKillDoc Then frmElpPrt.Hide: Exit Sub
                End If
            
            solde = solde + xYBIAMVT0.MOUVEMMON
            
           rsLocal.MoveNext
           Call rsYBIAMVT0_GetBuffer(rsLocal, xYBIAMVT0)

Loop


'Pas de mouvement dans la période : imprimer un extrait (sauf mensuel)
If Not blnA4_Form Then
    'If lRELEVEREL <> "M" Then prtBIA_Relevé_Annuel_Frais_Form "Solde au : " & dateImp(dateElp("Jour", -1, valAmjMin)), lRéférence
    'prtBIA_Relevé_Annuel_Frais_Form "Solde au : " & dateImp(meYBIACPT0.SOLDEDMO + 19000000), lRéférence
    prtBIA_Relevé_Annuel_Frais_Form "Total au : " & dateImp(dateElp("Jour", -1, lAMJMin)), lRéférence
    
End If

If blnA4_Form Then
    If MOUVEMOPE_Nb <> 0 Then prtBIA_Relevé_Annuel_Frais_Line_Total

    XPrt.CurrentY = Line4 + 50
    If blnFRS_Info Then XPrt.CurrentX = col2:       XPrt.Print "# : frais";

    X = "Total au : " & dateImp(valAmjMax)
    XPrt.CurrentX = Col4 - XPrt.TextWidth(X) - 200
    XPrt.Print X;
    XPrt.CurrentX = 5000
    XPrt.ForeColor = RGB(0, 123, 141)
    prtBIA_Relevé_Annuel_Frais_Montant (solde)
    XPrt.ForeColor = vbBlack

    '$JPL 2002.12.26 médiateur
    blnMédiateur = retourne_mediateur(meYBIACPT0.CLIENACLI)
    If blnMédiateur Then
        Call prtBIA_Relevé_Annuel_Frais_Médiateur
    Else
        If blnMsgInfo Then
            XPrt.FontBold = True: XPrt.FontSize = 10
            XPrt.CurrentY = XPrt.CurrentY + prtlineHeight ''* 2
            Call frmElpPrt.prtTrame(col1, XPrt.CurrentY, Col8, XPrt.CurrentY + prtlineHeight - 10, " ", 245)
            frmElpPrt.prtCentré 5500, mMsgInfo
        End If
    End If
End If
lstErr.AddItem xYBIAMVT0.MOUVEMCOM & " FIN : " & NbPage

If InStr(XPrt.Devicename, "PDF") < 1 Then
    prtBIA_Relevé_Annuel_Frais_Close
End If
If blnCSV Then
    xYBIAMVT0.MOUVEMDTR = valAmjMax - 19000000
    xYBIAMVT0.MOUVEMDVA = xYBIAMVT0.MOUVEMDTR
    xYBIAMVT0.MOUVEMMON = solde
    xYBIAMVT0.LIBELLIB1 = "Solde final "
    xYBIAMVT0.LIBELLIB2 = "": xYBIAMVT0.LIBELLIB3 = "": xYBIAMVT0.LIBELLIB4 = ""
    prtYBIAMVT0_CSV

    Close intFile
    Call FEU_VERT

End If
Set rsLocal = Nothing

End Sub


Public Sub prtBIA_Relevé_Annuel_Frais_OpenX_ResetPDF()
    Set XPrt = Printer
    frmElpPrt.Show vbModeless
    Height8_6 = frmElpPrt.prtHeightDelta(8, 6)
    prtTitleText = "Extrait de Compte"
    prtPgmName = "prtBIA_Relevé_Annuel_Frais"
    prtTitleUsr = usrName
    prtFontName = prtFontName_Arial
    prtLineNb = 1
    prtlineHeight = 250
    prtHeaderHeight = 300
    prtOrientation = vbPRORPortrait
    prtFormType = ""
    prtSocInit
    col1 = prtMinX
    col2 = col1 + 1100 '1325
    col3 = col1 + 6100 '6025
    Col4 = col1 + 7250 '6950
    Col5 = col1 + 9075 '8925
    Col6 = col1 + 10900
    Col7 = col1 + 10900
    Col8 = col1 + 10900
    prtBIA_Relevé_Annuel_Frais_OpenX_Reset_Line " "

End Sub


Public Sub prtBIA_Relevé_Annuel_Frais_RIB()
Dim iY As Integer
Dim wCurrentY As Integer
'--------------------------TRAME---------------------------
Dim X As String

XPrt.DrawWidth = 1
iY = 1500
prtFillColor = RGB(240, 255, 255)
XPrt.Line (200, iY + 1450)-(4750, iY + 1700), prtLineColor, B
Call frmElpPrt.prtTrame_Color(203, iY + 1454, 4753, iY + 1696, " ")

'------------------------verticaux avec arrondi
XPrt.Line (200, iY + 200)-(200, iY + 3200), prtLineColor
XPrt.Line (1100, iY + 1450)-(1100, iY + 2100), prtLineColor
XPrt.Line (2000, iY + 1450)-(2000, iY + 2100), prtLineColor
XPrt.Line (4200, iY + 1450)-(4200, iY + 2100), prtLineColor
XPrt.Line (4750, iY + 200)-(4750, iY + 3200), prtLineColor

'------------------------horizontaux
XPrt.Line (400, iY)-(4550, iY), prtLineColor
XPrt.Line (200, iY + 250)-(4750, iY + 250), prtLineColor

XPrt.Line (200, iY + 2100)-(4750, iY + 2100), prtLineColor
XPrt.Line (400, iY + 3400)-(4550, iY + 3400), prtLineColor

Call frmElpPrt.prtTrame_Color(200, iY + 4, 4750, iY + 243, " ")
XPrt.Line (200, iY + 200)-(200, iY + 243), prtLineColor
XPrt.Line (4750, iY + 200)-(4750, iY + 243), prtLineColor

'------------------------

XPrt.DrawWidth = 1

XPrt.CurrentY = 0
XPrt.CurrentX = 0
XPrt.Circle Step(200 + 200, iY + 200), 200, prtLineColor, 0.5 * Pi, Pi

XPrt.CurrentY = 0
XPrt.CurrentX = 0
XPrt.Circle Step(4750 - 200, iY + 200), 200, prtLineColor, 0, 0.5 * Pi

XPrt.CurrentY = 0
XPrt.CurrentX = 0
XPrt.Circle Step(200 + 200, iY + 3400 - 200), 200, prtLineColor, Pi, 1.5 * Pi

XPrt.CurrentY = 0
XPrt.CurrentX = 0
XPrt.Circle Step(4750 - 200, iY + 3400 - 200), 200, prtLineColor, 1.5 * Pi, 2 * Pi

XPrt.CurrentY = iY + prtlineHeight - 200
XPrt.FontSize = 8
XPrt.FontBold = True
If blnRIB Then
    frmElpPrt.prtCentré 2500, "RELEVE D'IDENTITE BANCAIRE"
Else
    If blnIban Then frmElpPrt.prtCentré 2500, "IBAN International Bank Account Number"
End If

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight - 50
XPrt.FontBold = False
XPrt.FontSize = 6
If blnRIB Then frmElpPrt.prtCentré 2500, "Cadre réservé au destinataire du R.I.B"

'------------------------------------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 5
wCurrentY = XPrt.CurrentY
'prtFillColor = RGB(240, 255, 255) 'RGB(0, 160, 180)
'XPrt.ForeColor = vbBlack 'vbWhite
'Call frmElpPrt.prtTrame_Color(200, XPrt.CurrentY - 50, 4750, XPrt.CurrentY + prtlineHeight, " ")
XPrt.CurrentY = wCurrentY
'---------------------------------------------------------
XPrt.FontBold = False
XPrt.FontSize = 6
If blnRIB Then
    XPrt.CurrentX = 250
    XPrt.Print "Code Banque";
    XPrt.CurrentX = 1200
    XPrt.Print "Code Guichet";
    XPrt.CurrentX = 4250
    XPrt.Print "clé R.I.B";
End If
XPrt.CurrentX = 2600
XPrt.Print "Numéro de compte";
'------------------------
prtFillColor = RGB(250, 255, 255)
XPrt.ForeColor = vbBlack

'----------------------------------------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight '+ 100
XPrt.FontBold = True
XPrt.FontSize = 9
If blnRIB Then
    XPrt.CurrentX = 400
    XPrt.Print strSocBdfE;
    XPrt.CurrentX = 1300
    XPrt.Print strSocBdfG;
    XPrt.CurrentX = 2400
    XPrt.Print Format$(mRib_Compte, "@@@  @@@  @@@  @@@");
    XPrt.CurrentX = 4400
    XPrt.Print Format$(mRib_Clé, "@@");
Else
    frmElpPrt.prtCentré 3100, Trim(mRib_Compte)
End If
XPrt.FontBold = False
XPrt.FontSize = 8
'------------------------------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight + 150
If blnIban Then
    XPrt.CurrentX = 300
    XPrt.Print "BIC";
    XPrt.CurrentX = 1050
    XPrt.Print ":";
    XPrt.CurrentX = 1200
    XPrt.FontBold = True
    XPrt.Print paramBic8;
    XPrt.FontBold = False
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight '+ 50
    XPrt.CurrentX = 300
    XPrt.Print "IBAN";
    XPrt.CurrentX = 1050
    XPrt.Print ":";
    XPrt.CurrentX = 1200
    XPrt.FontBold = True
    XPrt.Print Iban_Print(mRib_IbanE);
    XPrt.FontBold = False
End If
'--------------------------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight '+ 150
XPrt.FontBold = False
XPrt.CurrentX = 300
XPrt.Print "Titulaire";
XPrt.CurrentX = 1050
XPrt.Print ":";
XPrt.CurrentX = 1200
XPrt.Print meYBIACPT0.COMPTEINT;
XPrt.FontBold = False
'------------------------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 300
XPrt.Print "Dom";
XPrt.CurrentX = 1050
XPrt.Print ":";
XPrt.FontBold = True
XPrt.CurrentX = 1200
XPrt.Print SocRibDom;
XPrt.FontBold = False
'------------------------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 300
XPrt.Print "Tél";
XPrt.CurrentX = 1050
XPrt.Print ":";
XPrt.CurrentX = 1200
XPrt.Print socTéléphone;
XPrt.FontBold = False
XPrt.CurrentX = 4350
XPrt.FontBold = True
XPrt.Print mResponsable;
XPrt.FontBold = False

End Sub


Public Sub prtBIA_Relevé_Annuel_Frais_Compte(lRELEVEREL As String)
Dim X20 As String * 20, X As String
Dim xId As String
Dim wCompte As String

xZADRESS0 = zZADRESS0
fiscalZADRESS0 = zZADRESS0

Call fctPCEC_Atribut(meYBIACPT0.COMPTEOBL, meYBIACPT0.COMPTEDEV, blnCptOrdinaire, blnRIB, blnMédiateur, blnIban)

mRib_Compte = Trim(meYBIACPT0.COMPTECOM)
wCompte = mRib_Compte
mRib_Clé = Format$(RibClé(strSocBdfE, strSocBdfG, wCompte, mRib_IbanE), "00")

If mRib_Clé = 99 Then blnRIB = False: blnMédiateur = False: blnIban = False

blnConvention_Print = blnCptOrdinaire
'_________________________________________________________________________________________

mResponsable = meYBIACPT0.CLIENARES

xZRELEVE0.RELEVECOM = meYBIACPT0.COMPTECOM
xZRELEVE0.RELEVEREL = lRELEVEREL

Call rsZRELEVE0_Read(xZRELEVE0)

xZADRESS0.ADRESSNUM = xZRELEVE0.RELEVENUM
xZADRESS0.ADRESSTYP = xZRELEVE0.RELEVETYP
xZADRESS0.ADRESSCOA = xZRELEVE0.RELEVEADR
If xZADRESS0.ADRESSTYP = "1" Then
    Call rsZADRESS0_Client(xZADRESS0)
Else
    Call rsZADRESS0_Compte(xZADRESS0)
End If
If Trim(xZADRESS0.ADRESSRA1) = "" Then xZADRESS0.ADRESSRA1 = meYBIACPT0.COMPTEINT
End Sub

Public Sub prtBIA_Relevé_Annuel_Frais_Report()
XPrt.CurrentY = Line4 + 50
prtBIA_Relevé_Annuel_Frais_Montant (solde)
NbPage = NbPage + 1
frmElpPrt.prtNewPage
prtBIA_Relevé_Annuel_Frais_Form "Report", ""

End Sub

Public Sub prtBIA_Relevé_Annuel_Frais_OpenX_Reset()

Set XPrt = Printer
frmElpPrt.Show vbModeless

Height8_6 = frmElpPrt.prtHeightDelta(8, 6)


prtTitleText = "Extrait de Compte"
prtPgmName = "prtBIA_Relevé_Annuel_Frais"
prtTitleUsr = usrName
prtFontName = prtFontName_Arial

prtLineNb = 1
prtlineHeight = 250

prtHeaderHeight = 300
prtOrientation = vbPRORPortrait

prtFormType = ""
prtSocInit
'prtInit
col1 = prtMinX
col2 = col1 + 1100 '1325
col3 = col1 + 6100 '6025
Col4 = col1 + 7250 '6950
Col5 = col1 + 9075 '8925
Col6 = col1 + 10900
Col7 = col1 + 10900
Col8 = col1 + 10900
prtBIA_Relevé_Annuel_Frais_OpenX_Reset_Line " "

End Sub


Public Sub prtYBIAMVT0_CSV()
Dim X As String
If xYBIAMVT0.MOUVEMMON < 0 Then
    X = ";" & cur_AbsV(xYBIAMVT0.MOUVEMMON)
Else
    X = cur_AbsV(xYBIAMVT0.MOUVEMMON) & ";"
End If

Print #intFile, xYBIAMVT0.MOUVEMCOM _
         ; ";"; xYBIAMVT0.COMPTEDEV _
         ; ";"; xYBIAMVT0.MOUVEMDTR + 19000000 _
         ; ";"; xYBIAMVT0.MOUVEMDVA + 19000000 _
         ; ";"; X _
         ; ";"; Trim(xYBIAMVT0.LIBELLIB1) & Trim(xYBIAMVT0.LIBELLIB2) & Trim(xYBIAMVT0.LIBELLIB3) & Trim(xYBIAMVT0.LIBELLIB4)
End Sub

Public Sub prtBIA_Relevé_Annuel_Frais_OpenX_Reset_Line(lFct As String)
If lFct = "M" Then
    NbLigneMax = 24 '20 '27
    Line1 = prtlineHeight * 29
Else
    NbLigneMax = 32 '28 ' 35
    Line1 = prtlineHeight * 21
End If

Line2 = Line1 + prtlineHeight + 50
Line3 = Line2 + prtlineHeight + 50
Line4 = Line3 + (prtlineHeight * NbLigneMax) + 50
Line5 = Line4 + prtlineHeight + 50

End Sub
Private Sub prtBIA_Relevé_Annuel_Frais_Médiateur()
Dim X As String
Dim aFontName As String

    aFontName = XPrt.FontName
    XPrt.FontName = "Calibri"
    XPrt.FontSize = 7
    XPrt.CurrentY = Line4 + 400
    XPrt.ForeColor = RGB(0, 0, 160)
    XPrt.FontBold = False
    XPrt.CurrentX = col1 + 50
    XPrt.Print "Pour toute insatisfaction ou désaccord, vous pouvez contacter:";
    XPrt.CurrentY = XPrt.CurrentY + 210
    XPrt.CurrentX = col1 + 100
    XPrt.Print "1. Le chargé de clientèle: votre premier interlocuteur ;";
    XPrt.CurrentY = XPrt.CurrentY + 190
    XPrt.CurrentX = col1 + 100
    XPrt.Print "2. Le Service Réclamations à l'adresse suivante Service Réclamations Banque BIA 67, avenue Franklin Roosevelt 75008 PARIS ou à l'adresse E-mail suivante : contact@bia-paris.fr ;";
    XPrt.CurrentY = XPrt.CurrentY + 190
    XPrt.CurrentX = col1 + 100
    XPrt.Print "3. Le médiateur, en dernier recours, une fois les deux recours exercés successivement ou sans réponse de la Banque BIA à l'issue d'un délai de 60 jours à compter de la réception de votre demande";
    XPrt.CurrentY = XPrt.CurrentY + 180
    XPrt.CurrentX = col1 + 100
    XPrt.Print "par le  Service  Réclamations, vous disposez de la  possibilité de saisir  gratuitement un  Médiateur indépendant de la Fédération Bancaire Française, en adressant un courrier  à  l'adresse suivante :";
    XPrt.CurrentY = XPrt.CurrentY + 180
    XPrt.CurrentX = col1 + 100
    XPrt.FontBold = True
    XPrt.Print "Le Médiateur auprès de la FBF, CS 151 75422 Paris Cedex 7";
    XPrt.FontBold = False
    XPrt.Print ", ou par voie électronique sur le site internet du Médiateur : ";
    XPrt.FontBold = True
    XPrt.Print "www.lemediateur.fbf.fr";
    XPrt.FontBold = False
    XPrt.Print ".";
    XPrt.CurrentY = XPrt.CurrentY + 180
    XPrt.CurrentX = col1 + 100
    XPrt.Print "Le Médiateur répondra dans un délai maximum de 2 mois à réception du dossier complet.";
    XPrt.ForeColor = vbBlack
    XPrt.FontSize = 8
    XPrt.FontName = aFontName
    
End Sub

