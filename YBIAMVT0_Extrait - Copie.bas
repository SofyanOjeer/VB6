Attribute VB_Name = "prtYBIAMVT0_extrait"
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

Dim xZBAGFAC0 As typeZBAGFAC0, mZBAGFAC0 As typeZBAGFAC0
Dim meZBAGFAC0 As typeZBAGFAC0

Dim blnCptOrdinaire As Boolean, blnRIB As Boolean, blnMédiateur As Boolean, blnIban As Boolean
Dim blnConvention_Print As Boolean
Dim mRib_Compte As String, mRib_Clé As String, mRib_IbanE As String
Dim mResponsable As String
Dim zZADRESS0 As typeZADRESS0, xZADRESS0 As typeZADRESS0, fiscalZADRESS0 As typeZADRESS0
Dim xZRELEVE0 As typeZRELEVE0
Dim intFile As Integer
Dim blnFRS_Info As Boolean


'XXX DSP - DEBUT - 30102009
Public g_blnFlagDSP As Boolean
'XXX DSP - FIN - 30102009

Dim blnPauget_Constans As Boolean
Dim curFrs_Total As Currency, dblTAEG As Double

Dim mZAUTSYC0 As typeZAUTSYC0, blnAUTSYCAUT_Dec As Boolean, blnAUTSYCAUT_Compte As Boolean


Type typeECHTAB
    Code       As String
    AMJ7       As Long
    Taux       As Double
End Type
Public arrECHTAB(100) As typeECHTAB, arrECHTAB_K As Integer, arrECHTAB_Nb As Integer
Public arrPays() As typePays, arrPays_NB As Integer

Public Sub prtYBIAMVT0_A4_InfoFGDR()

    If retourne_Eligibilite(meYBIACPT0.CLIENACLI, meYBIACPT0.COMPTECOM) Then
        XPrt.FontSize = 7
        XPrt.CurrentY = 14676: XPrt.Print "";
        XPrt.ForeColor = RGB(0, 0, 160)
        XPrt.CurrentX = col1 + 200
        XPrt.Print "Toutes les sommes déposées sur ce compte sont couvertes par la garantie des dépôts,";
        XPrt.Print " dès lors qu'elles sont libellées en euro, en franc CFP ou dans une devise de l'Espace Économique";
        XPrt.CurrentY = XPrt.CurrentY + 180
        XPrt.CurrentX = col1 + 200
        XPrt.Print "Européen (EEE). Les déposants sont indemnisés à hauteur de 100 000 € par personne et par établissement de crédit.";
        XPrt.ForeColor = vbBlack
        XPrt.FontSize = 8
    Else
        XPrt.FontSize = 7
        XPrt.CurrentY = 14676: XPrt.Print "";
        XPrt.ForeColor = RGB(0, 0, 160)
        XPrt.CurrentX = col1 + 200
        XPrt.Print "";
        XPrt.Print "";
        XPrt.CurrentY = XPrt.CurrentY + 180
        XPrt.CurrentX = col1 + 200
        XPrt.Print "";
        XPrt.ForeColor = vbBlack
        XPrt.FontSize = 8
    End If

End Sub

Public Function retourne_Eligibilite(lClient As String, lCompte As String) As Boolean
Dim V, xName As String, xMemo As String
Dim xSQL As String
Dim rs As ADODB.Recordset
Dim myCon As New ADODB.Connection

    retourne_Eligibilite = True
    'Voir si existe dans la table sql server VUC_Clients
    V = rsElpTable_Read("SIDE", "PasswordX", "SIDE_READ", xName, xMemo)
    myCon.Open "W4SRV", "READ_ONLY", xMemo
    Set rs = myCon.Execute("select * from [VUC_V3].[dbo].VUC_Clients where titulaire='" & lClient & "' and compte ='" & lCompte & "'")
    If rs.BOF Or rs.EOF Then
        retourne_Eligibilite = False
        If rs.State = adStateOpen Then
            rs.Close
        End If
        Set rs = Nothing
        Exit Function
    End If
    'Si oui on applique les requêtes nécessaires aux balises xml
    xSQL = "SELECT '??' FROM " & paramIBM_Library_SAB & ".ZCLIENA0, " & paramIBM_Library_SAB & ".ZCOMPTE0 WHERE CLIENACLI = 'comptecli' AND CLIENACAT IN('PAR','PRR','PER','GAR','PAV') AND CLIENAETA IN('MR','MME','MLLE','EIMR') AND COMPTECOM = 'comptecompte' AND (LEFT(COMPTEOBL, 3) = '255' OR LEFT(COMPTEOBL, 4) = '2511' OR LEFT(COMPTEOBL, 3) = '262' OR LEFT(COMPTEOBL, 3) = '253') AND COMPTEDEV NOT IN('EUR','GBP','BGN','HRK','DKK','HUF','ISK','LVL','LTL','NOK','PLN','CZK','RON','SEK')"
    xSQL = Replace(xSQL, "comptecli", Trim(lClient))
    xSQL = Replace(xSQL, "comptecompte", Trim(lCompte))
    Set rs = cnsab.Execute(xSQL)
    If Not rs.EOF Then
        If rs(0) = "??" Then
            retourne_Eligibilite = False
        End If
    End If
    If rs.State = adStateOpen Then
        rs.Close
    End If
    Set rs = Nothing
    
    If Not retourne_Eligibilite Then Exit Function
    
    xSQL = "SELECT '??' FROM " & paramIBM_Library_SAB & ".ZCLIENA0, " & paramIBM_Library_SAB & ".ZCOMPTE0 WHERE CLIENACLI = 'comptecli' AND CLIENACAT IN('PAR','PRR','PER','GAR','PAV') AND CLIENAETA in('MRS','M/ME','MMME') AND COMPTECOM = 'comptecompte' AND (LEFT(COMPTEOBL, 3) = '255' OR LEFT(COMPTEOBL, 4) = '2511' OR LEFT(COMPTEOBL, 3) = '262' OR LEFT(COMPTEOBL, 3) = '253') AND COMPTEDEV NOT IN ('EUR','GBP','BGN','HRK','DKK','HUF','ISK','LVL','LTL','NOK','PLN','CZK','RON','SEK')"
    xSQL = Replace(xSQL, "comptecli", Trim(lClient))
    xSQL = Replace(xSQL, "comptecompte", Trim(lCompte))
    Set rs = cnsab.Execute(xSQL)
    If Not rs.EOF Then
        If rs(0) = "??" Then
            retourne_Eligibilite = False
        End If
    End If
    If rs.State = adStateOpen Then
        rs.Close
    End If
    Set rs = Nothing

    If Not retourne_Eligibilite Then Exit Function
    
    xSQL = "SELECT '??' FROM " & paramIBM_Library_SAB & ".ZCLIENA0, " & paramIBM_Library_SAB & ".ZCOMPTE0 WHERE CLIENACLI = 'comptecli' AND CLIENACAT IN('BQG','BQE','STE','EI','ASR','ASS','AMB','ADM') AND CLIENAETA IN('MR','MME','MLLE','EIMR','BANQ','SARL','Sa','SA','SETR','SAS','SNC','SCV','AMBA','SEP','SCI','ASSO','SCS','EURL','SCA','PMAD','SA D') AND COMPTECOM = 'comptecompte' AND (LEFT(COMPTEOBL, 3) = '255' OR LEFT(COMPTEOBL, 4) = '2511' OR LEFT(COMPTEOBL, 3) = '262' OR LEFT(COMPTEOBL, 3) = '253') AND COMPTEDEV NOT IN ('EUR','GBP','BGN','HRK','DKK','HUF','ISK','LVL','LTL','NOK','PLN','CZK','RON','SEK')"
    xSQL = Replace(xSQL, "comptecli", Trim(lClient))
    xSQL = Replace(xSQL, "comptecompte", Trim(lCompte))
    Set rs = cnsab.Execute(xSQL)
    If Not rs.EOF Then
        If rs(0) = "??" Then
            retourne_Eligibilite = False
        End If
    End If
    If rs.State = adStateOpen Then
        rs.Close
    End If
    Set rs = Nothing

    If Not retourne_Eligibilite Then Exit Function
    
    xSQL = "SELECT '??' FROM " & paramIBM_Library_SAB & ".ZCLIENA0, " & paramIBM_Library_SAB & ".ZCOMPTE0 WHERE CLIENACLI = 'comptecli' AND CLIENACAT IN('BQG','BQE','STE','EI','ASR','ASS','AMB','ADM') AND CLIENAETA IN('MRS','M/ME','MMME') AND COMPTECOM = 'comptecompte' AND (LEFT(COMPTEOBL, 3) = '255' OR LEFT(COMPTEOBL, 4) = '2511' OR LEFT(COMPTEOBL, 3) = '262' OR LEFT(COMPTEOBL, 3) = '253') AND COMPTEDEV NOT IN ('EUR','GBP','BGN','HRK','DKK','HUF','ISK','LVL','LTL','NOK','PLN','CZK','RON','SEK')"
    xSQL = Replace(xSQL, "comptecli", Trim(lClient))
    xSQL = Replace(xSQL, "comptecompte", Trim(lCompte))
    Set rs = cnsab.Execute(xSQL)
    If Not rs.EOF Then
        If rs(0) = "??" Then
            retourne_Eligibilite = False
        End If
    End If
    If rs.State = adStateOpen Then
        rs.Close
    End If
    Set rs = Nothing

End Function

Public Sub prtZBAGFAC0_A4_Compte(lRELEVEREL As String)
Dim X20 As String * 20, X As String
Dim xId As String
Dim wCompte As String

    xZADRESS0 = zZADRESS0
    fiscalZADRESS0 = zZADRESS0
    Call fctPCEC_Atribut(meYBIACPT0.COMPTEOBL, meYBIACPT0.COMPTEDEV, blnCptOrdinaire, blnRIB, blnMédiateur, blnIban)
    mRib_Compte = Trim(meYBIACPT0.COMPTECOM)
    wCompte = mRib_Compte
    mRib_Clé = Format$(RibClé(strSocBdfE, strSocBdfG, wCompte, mRib_IbanE), "00")
    If mRib_Clé = 99 Then
        blnRIB = False: blnMédiateur = False: blnIban = False
    End If
    blnConvention_Print = blnCptOrdinaire
    If Mid$(meYBIACPT0.COMPTEOBL, 1, 4) = "2511" Then
        blnPauget_Constans = True
    Else
        blnPauget_Constans = False
    End If
    If lRELEVEREL = "M" And meYBIACPT0.COMPTEDEV = "EUR" Then
        'blnPauget_Constans = blnMédiateur
        If blnPauget_Constans Then
            prtYBIAMVT0_A4_Pauget_Constans_Init
        End If
    Else
        If valAmjMax = YBIATAB0_DATE_CPT_J Then
            'blnPauget_Constans = blnMédiateur
            If blnPauget_Constans Then
                prtYBIAMVT0_A4_Pauget_Constans_Init
            End If
        End If
    End If
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
    If Trim(xZADRESS0.ADRESSRA1) = "" Then
        xZADRESS0.ADRESSRA1 = meYBIACPT0.COMPTEINT
    End If
    
End Sub
Public Sub prtZBAGFAC0_A4_Extrait_Frais(lMOUVEMCOM As String, lAMJMin As String, lAMJMax As String, blnCSV As Boolean, lstErr As ListBox, lRELEVEREL As String, lRéférence As String, blnNewPage As Boolean)
Dim rsLocal As ADODB.Recordset, rsW As ADODB.Recordset
Dim xSQL As String
Dim Nb As Integer
Dim CTLAMJ As String
Dim V
Dim wAMJ_Solde As String

    valAmjMin = lAMJMin
    valAmjMax = lAMJMax
    IbmAmjMin = dateIBM(lAMJMin)
    IbmAmjMax = dateIBM(lAMJMax)
    wAMJ_Solde = dateElp("Jour", -1, valAmjMin)
    rsZADRESS0_Init zZADRESS0
    blnNewPage = False
    blnFRS_Info = False
    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIACPT0" _
     & " where COMPTECOM = '" & lMOUVEMCOM & "'"
    Set rsLocal = Nothing
    Set rsLocal = cnsab.Execute(xSQL)
    If Not rsLocal.EOF Then
        V = rsYBIACPT0_GetBuffer(rsLocal, meYBIACPT0)
        If Not IsNull(V) Then
            MsgBox "prtZBAGFAC0_A4_Extrait " & V
            Exit Sub
        End If
    End If
    xSQL = "select BAGFACDOP, BAGFACDVA, BAGFACMCO, BAGFACNAT, BAGFACLI1, BAGFACLI2, BAGFACCPT"
    xSQL = xSQL & " from " & paramIBM_Library_SAB & ".ZBAGFAC0"
    xSQL = xSQL & " where BAGFACCPT = '" & lMOUVEMCOM & "'"
    xSQL = xSQL & " and BAGFACCOE = 1"
    xSQL = xSQL & " and (BAGFACOPE = 'FRS' and BAGFACNAT IN('OCB','DEB','CDD','CHD','RPR','ATD','JUR','ADM','INT','NOT','DCB'))"
    'xSQL = xSQL & " or BAGFACOPE = 'A0V') and (BAGFACDOP >= " & IbmAmjMin & " and BAGFACDOP <= " & IbmAmjMax & ")"
    xSQL = xSQL & " and BAGFACDOP >= " & dateIBM(DSys)
    Set rsLocal = Nothing
    Set rsLocal = cnsab.Execute(xSQL)
    If Not rsLocal.EOF Then
        solde = 0
        V = rsZBAGFAC0_GetBuffer(rsLocal, xZBAGFAC0)
        If Not IsNull(V) Then
            MsgBox "prtZBAGFAC0_A4_Frais " & V
            Exit Sub
        End If
        Call prtZBAGFAC0_A4_Compte(lRELEVEREL)
        prtZBAGFAC0_A4_OpenX_Reset
        blnA4_Form = False
        mExtraitNuméro = libMois(Mid$(lAMJMax, 5, 2))
        prtFontSize = 8
        NbPageMax = 0
        NbPage = 1
        Do Until rsLocal.EOF
                If xZBAGFAC0.BAGFACDOP >= IbmAmjMin Then
                    If Not blnA4_Form Then
                        Call prtZBAGFAC0_A4_Form("Édité le " & dateImp(DSys), "")
                        blnNewPage = True
                    End If
                    If nbLigne = NbLigneMax Then
                        prtZBAGFAC0_A4_Report
                        lstErr.RemoveItem lstErr.ListCount - 1
                        lstErr.AddItem xZBAGFAC0.BAGFACDOP & " page : " & NbPage
                    End If
                    nbLigne = nbLigne + 1
                    prtZBAGFAC0_A4_Line
                    DoEvents: If prtKillDoc Then frmElpPrt.Hide: Exit Sub
                End If
           rsLocal.MoveNext
           Call rsZBAGFAC0_GetBuffer(rsLocal, xZBAGFAC0)
        Loop
        XPrt.CurrentY = Line4 + 50
        X = "           Total"
        XPrt.CurrentX = Col5 - XPrt.TextWidth(X) - 200
        XPrt.Print X;
        XPrt.CurrentX = 5000
        Call prtZBAGFAC0_A4_Montant(solde)
        Call prtZBAGFAC0_A4_Info_Decret
        If blnMédiateur Then
            prtYBIAMVT0_A4_Médiateur
        Else
            If blnMsgInfo Then
                XPrt.FontBold = True: XPrt.FontSize = 10
                XPrt.CurrentY = XPrt.CurrentY + prtlineHeight ''* 2
                Call frmElpPrt.prtTrame(col1, XPrt.CurrentY, Col8, XPrt.CurrentY + prtlineHeight - 10, " ", 245)
                frmElpPrt.prtCentré 5500, mMsgInfo
            End If
        End If
    End If
    lstErr.AddItem xZBAGFAC0.BAGFACMCO & " FIN : " & NbPage
    prtYBIAMVT0_A4_Close
    Set rsLocal = Nothing
    
End Sub
Public Sub prtZBAGFAC0_A4_Form(Msg As String, lRéférence As String)
Dim X As String
Dim mCurrenty
Dim aFont As String

    prtYBIAMVT0_A4_RIB
    prtZBAGFAC0_A4_OpenX_Reset_Line " "
    Call frmElpPrt.prtTrame(Col4, Line3, Col5, Line4, " ", 250)
    prtFillColor = RGB(240, 255, 255)
    Call frmElpPrt.prtTrame_Color(col1, Line2, Col8, Line3, " ")
    prtFillColor = prtFillColor_Standard
    XPrt.CurrentY = prtMinY + prtlineHeight * 4
    XPrt.DrawWidth = 1
    XPrt.Line (col1 + 200, Line2)-(Col8 - 200, Line2), prtLineColor
    XPrt.Line (col1, Line3)-(Col8, Line3), prtLineColor
    XPrt.Line (col1 + 200, Line4)-(Col8, Line4), prtLineColor
    XPrt.DrawWidth = 1
    XPrt.Line (Col5 + 200, Line5)-(Col6 - 200, Line5), prtLineColor
    XPrt.DrawWidth = 1
    XPrt.Line (col1, Line2 + 200)-(col1, Line4 - 200), prtLineColor
    XPrt.DrawWidth = 1
    XPrt.Line (col2, Line2)-(col2, Line4), prtLineColor
    XPrt.DrawWidth = 1
    XPrt.Line (Col4, Line2)-(Col4, Line4), prtLineColor
    XPrt.DrawWidth = 1
    XPrt.Line (Col5, Line2)-(Col5, Line5 - 200), prtLineColor
    XPrt.DrawWidth = 1
    XPrt.Line (Col6, Line2 + 200)-(Col6, Line5 - 200), prtLineColor
    XPrt.CurrentY = Line2 + 50
    XPrt.FontBold = True
    XPrt.FontSize = prtFontSize
    frmElpPrt.prtCentré (col1 + col2) / 2, "Date"
    frmElpPrt.prtCentré (col2 + col3) / 2, "Libellé"
    frmElpPrt.prtCentré (Col4 + Col5) / 2, "Date de Valeur"
    frmElpPrt.prtCentré (Col5 + Col6) / 2, "Montant"
    XPrt.DrawWidth = 2
    XPrt.CurrentY = 0
    XPrt.CurrentX = 0
    XPrt.Circle Step(col1 + 200, Line2 + 200), 200, prtLineColor, 0.5 * Pi, Pi
    XPrt.DrawWidth = 3
    XPrt.CurrentY = 0
    XPrt.CurrentX = 0
    XPrt.Circle Step(Col6 - 200, Line2 + 200), 200, prtLineColor, 0, 0.5 * Pi
    XPrt.CurrentY = 0
    XPrt.CurrentX = 0
    XPrt.DrawWidth = 2
    XPrt.CurrentY = 0
    XPrt.CurrentX = 0
    XPrt.Circle Step(col1 + 200, Line4 - 200), 200, prtLineColor, Pi, 1.5 * Pi
    XPrt.DrawWidth = 3
    XPrt.CurrentY = 0
    XPrt.CurrentX = 0
    XPrt.Circle Step(Col5 + 200, Line5 - 200), 200, prtLineColor, Pi, 1.5 * Pi
    XPrt.CurrentY = 0
    XPrt.CurrentX = 0
    XPrt.Circle Step(Col6 - 200, Line5 - 200), 200, prtLineColor, 1.5 * Pi, 2 * Pi
    XPrt.FontSize = 10
    XPrt.CurrentY = prtMinY + prtlineHeight * 9 - XPrt.TextHeight("test")
    XPrt.FontBold = True
    XPrt.CurrentX = 6200
    XPrt.Print xZADRESS0.ADRESSRA1;
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = 6200
    XPrt.Print xZADRESS0.ADRESSRA2;
    XPrt.FontBold = False
    If Trim(xZADRESS0.ADRESSAD1) <> "" Then
        XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
        XPrt.CurrentX = 6200
        XPrt.Print xZADRESS0.ADRESSAD1;
    End If
    If Trim(xZADRESS0.ADRESSAD2) <> "" Then
        XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
        XPrt.CurrentX = 6200
        XPrt.Print xZADRESS0.ADRESSAD2;
    End If
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = 6200
    XPrt.Print xZADRESS0.ADRESSAD3;
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = 6200
    If Trim(xZADRESS0.ADRESSCOP) <> "" Then XPrt.Print xZADRESS0.ADRESSCOP & "  ";
    XPrt.Print xZADRESS0.ADRESSVIL;
    If Trim(xZADRESS0.ADRESSPAY) <> "FRANCE" Then
        XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
        XPrt.CurrentX = 6200
        XPrt.Print xZADRESS0.ADRESSPAY;
    End If
    XPrt.FontSize = 8
    XPrt.CurrentY = Line1 - prtlineHeight * 3 + 150
    XPrt.FontBold = True
    prtFillColor = RGB(240, 255, 255)
    Call frmElpPrt.prtTrame_Color(col3, XPrt.CurrentY - 100, Col6, XPrt.CurrentY + prtlineHeight, " ")
    prtFillColor = prtFillColor_Standard
    XPrt.CurrentX = col3 + 200
    XPrt.CurrentY = XPrt.CurrentY - Height8_6
    XPrt.FontSize = 10
    XPrt.Print meYBIACPT0.COMPTEDEV;
    XPrt.CurrentY = XPrt.CurrentY + Height8_6
    XPrt.FontBold = False
    XPrt.FontSize = 8
    XPrt.Print "  -  ";
    XPrt.FontBold = True
    XPrt.CurrentY = XPrt.CurrentY - Height8_6
    XPrt.FontSize = 10
    XPrt.Print mExtraitNuméro;
    XPrt.CurrentY = XPrt.CurrentY + Height8_6
    XPrt.FontBold = False
    XPrt.FontSize = 8
    XPrt.Print "     / " & Format$(NbPage, "###");
    XPrt.CurrentX = Col6 + 20
    XPrt.Print lRéférence;
    
    XPrt.CurrentY = XPrt.CurrentY + (prtlineHeight * 1.8)
    XPrt.FontBold = True
    XPrt.FontSize = 10
    XPrt.CurrentX = col2 + 1300
    aFont = XPrt.FontName
    XPrt.FontName = "Calibri"
    XPrt.Print "INFORMATION PREALABLE EN MATIERE DE FRAIS BANCAIRES"
    XPrt.FontName = aFont
    
    XPrt.FontSize = 8
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
    XPrt.CurrentX = col1 + 50
    XPrt.FontBold = True
    XPrt.FontBold = False
    XPrt.FontSize = 8
    XPrt.FontSize = 8
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = col1 + 50
    XPrt.FontBold = True
    XPrt.FontBold = False
    XPrt.FontSize = prtFontSize
    XPrt.CurrentY = Line1 + 50
    XPrt.CurrentX = Col5 + 100
    XPrt.Print Msg;
    nbLigne = 0
    blnA4_Form = True
    XPrt.CurrentY = Line3 - prtlineHeight + 50

End Sub

Public Sub prtZBAGFAC0_A4_Info_Decret()

    XPrt.FontSize = 7
    XPrt.CurrentY = 14676: XPrt.Print "";
    XPrt.ForeColor = RGB(0, 0, 160)
    XPrt.CurrentX = col1 + 200
    XPrt.FontBold = True
    XPrt.Print "Conformément au décrêt n° 2014-739 du 30 juin2014 relatif à l'information préalable du consommateur";
    XPrt.Print " en matière de frais bancaires, veuillez trouver ci-dessus le récapitulatif des frais";
    XPrt.CurrentY = XPrt.CurrentY + 180
    XPrt.CurrentX = col1 + 200
    XPrt.Print "débités sur votre compte le mois prochain.";
    XPrt.ForeColor = vbBlack
    XPrt.FontSize = 8
    XPrt.FontBold = False

End Sub

Public Sub prtZBAGFAC0_A4_Line()
Dim X As String, I As Integer, libCV As String, blnCV As Boolean
Dim blnLine2 As Boolean, xLine1 As String, xLine2 As String
Dim kJust As Integer, kMax As Integer
Dim widthCOL3_5 As Integer
    
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.FontSize = prtFontSize
    XPrt.FontBold = False
    XPrt.ForeColor = vbBlack
    XPrt.CurrentX = col1 + 50
    XPrt.Print dateImp(Val(xZBAGFAC0.BAGFACDOP) + 19000000);
    XPrt.CurrentX = Col4 + 400
    XPrt.Print dateImp(Val(xZBAGFAC0.BAGFACDVA) + 19000000);
    Call prtZBAGFAC0_A4_Montant(xZBAGFAC0.BAGFACMCO)
    XPrt.CurrentX = col2 + 50
    xLine1 = Trim(xZBAGFAC0.BAGFACLI1) & " " & Trim(xZBAGFAC0.BAGFACLI2)
    xLine2 = ""
    X = Replace(xLine1 & " " & xLine2, "  ", " ")
    X = Replace(X, "  ", " ")
    widthCOL3_5 = (col3 - XPrt.CurrentX)
    XPrt.FontSize = 8
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
        If nbLigne = NbLigneMax Then
            Call prtZBAGFAC0_A4_Report
        End If
        nbLigne = nbLigne + 1
        XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
        XPrt.CurrentX = col2 + 50
        XPrt.Print xLine2;
    End If
    XPrt.ForeColor = vbBlack
    XPrt.FontSize = prtFontSize

End Sub
Public Sub prtZBAGFAC0_A4_Montant(MT As Currency)
Dim X As String

    XPrt.FontBold = True
    X = Format$(Abs(MT), "## ### ### ### ### ##0.00")
    XPrt.CurrentX = Col6 - 100 - XPrt.TextWidth(X)
    XPrt.Print X;
    XPrt.FontBold = False
    solde = solde + MT
    
End Sub

Public Sub prtZBAGFAC0_A4_OpenX()
    
    On Error GoTo prtError
    Call prtZBAGFAC0_A4_OpenX_Reset

Exit Sub

prtError:

    Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
    frmElpPrt.Hide

End Sub

Public Sub prtZBAGFAC0_A4_OpenX_Reset()

    Set XPrt = Printer
    frmElpPrt.Show vbModeless
    Height8_6 = frmElpPrt.prtHeightDelta(8, 6)
    prtTitleText = "Frais bancaires"
    prtPgmName = "prtZBAGFAC0_A4"
    prtTitleUsr = usrName
    prtFontName = "Calibri"
    prtLineNb = 1
    prtlineHeight = 250
    prtHeaderHeight = 300
    prtOrientation = vbPRORPortrait
    prtFormType = ""
    prtSocInit
    col1 = prtMinX
    col2 = col1 + 1100
    col3 = col1 + 6100
    Col4 = col1 + 7250
    Col5 = col1 + 9075
    Col6 = col1 + 10900
    Col7 = col1 + 10900
    Col8 = col1 + 10900
    prtZBAGFAC0_A4_OpenX_Reset_Line " "
    
End Sub

Public Sub prtZBAGFAC0_A4_OpenX_Reset_Line(lFct As String)

    If lFct = "M" Then
        NbLigneMax = 27
        Line1 = prtlineHeight * 29
    Else
        NbLigneMax = 35
        Line1 = prtlineHeight * 21
    End If
    If blnPauget_Constans Then
        NbLigneMax = NbLigneMax - 3
    End If
    Line2 = Line1 + prtlineHeight + 50
    Line3 = Line2 + prtlineHeight + 50
    Line4 = Line3 + prtlineHeight * NbLigneMax + 50
    Line5 = Line4 + prtlineHeight + 50

End Sub

Public Sub prtZBAGFAC0_A4_Report()
    
    XPrt.CurrentY = Line4 + 50
    NbPage = NbPage + 1
    frmElpPrt.prtNewPage
    prtZBAGFAC0_A4_Form "Report", ""

End Sub



Public Sub prtYBIAMVT0_A4_OpenX()
'---------------------------------------------------------
On Error GoTo prtError

'$20060605_JPL Set XPrt = Printer
'$20060605_JPL frmElpPrt.Show vbModeless

prtYBIAMVT0_A4_OpenX_Reset



Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub

Public Sub prtYBIAMVT0_A4_Close()
'---------------------------------------------------------
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



'---------------------------------------------------------
Public Sub prtYBIAMVT0_A4_Form(Msg As String, lRéférence As String)
'---------------------------------------------------------
Dim X As String
Dim mCurrenty

prtYBIAMVT0_A4_RIB

 blnConvention_Print = False
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
If Not blnConvention_Print Then
    prtYBIAMVT0_A4_OpenX_Reset_Line " "
Else
    prtYBIAMVT0_A4_OpenX_Reset_Line "M"
    blnConvention_Print = False
    nbLigne = 5
    XPrt.FontBold = True
'JPL : Modifié le 28.09.2007
'------------------------------
'GoTo MSG_200709 'fin
'JPL : Modifié le 28.03.2006
'------------------------------
GoTo MSG_200611 'fin
    XPrt.FontSize = 9
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 4
    XPrt.CurrentX = col1 + 50
    XPrt.Print "La BIA a mis à jour les conditions générales appliquées à la clientèle le 1er juillet 2006, ";
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = col1 + 50
    XPrt.Print "date depuis laquelle la nouvelle grille tarifaire est à votre disposition.";
'JPL : Supprimé le 28.03.2006
'------------------------------
MSG_200611:
'XXX DSP - DEBUT - 30102009
If g_blnFlagDSP = True Then GoTo MSG_DSP
'XXX DSP - FIN - 30102009
'XXX DSP - COMMENTAIRES - 26112009
'    XPrt.FontSize = 7
'    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
'    XPrt.CurrentX = col1 + 50
'    XPrt.Print "Nous vous rappelons que nous mettons à votre disposition une convention de compte pour contractualiser tous les aspects de votre relation bancaire.";
'    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'    XPrt.CurrentX = col1 + 50
'    XPrt.Print "Cette convention reprend les conditions générales de fonctionnement de votre compte ainsi que des moyens de paiement qui y sont attachés.";
'    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'    XPrt.CurrentX = col1 + 50
'    XPrt.Print "Dans l'hypothèse où vous n'auriez pas signé cette convention, votre conseiller peut vous en adresser gratuitement 2 exemplaires sur simple demande.";
'    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'    XPrt.CurrentX = col1 + 50
'    XPrt.Print "Il vous appartiendra alors de nous en retourner un exemplaire signé.";
'    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'    XPrt.CurrentX = col1 + 50
'    XPrt.Print "Nous vous rappelons enfin qu'il est également possible de signer cette convention auprès de votre conseiller qui se tient à votre disposition pour tout";
'    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'    XPrt.CurrentX = col1 + 50
'    XPrt.Print "renseignement complémentaire.";
GoTo MSG_FIN
'XXX DSP - DEBUT - 30102009
MSG_DSP:
'XXX DSP - DEBUT - 29122009 Suppression du message DSP
'    XPrt.FontSize = 7
'    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
'    XPrt.CurrentX = col1 + 50
'    XPrt.Print "Nous vous informons qu'une nouvelle réglementation européenne dite " & Chr(171) & " Directive sur les Services de Paiement " & Chr(187) & " est entrée en vigueur en France";
'    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'    XPrt.CurrentX = col1 + 50
'    XPrt.Print "depuis le 1er novembre 2009. L'avenant à votre convention de compte est disponible à nos guichets. Une lettre d'information vous a été adressée";
'    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'    XPrt.CurrentX = col1 + 50
'    XPrt.Print "et sera consultable sur notre site internet www.bia-paris.fr.";
'    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'    XPrt.CurrentX = col1 + 50
'    XPrt.Print "We inform you that a new European legislation known as " & Chr(171) & " Payment Services Directive " & Chr(187) & " is in force in France since 1st of November 2009. An";
'    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'    XPrt.CurrentX = col1 + 50
'    XPrt.Print "amendment  to your account  agreement is available at our counters. A letter of information  has been sent to you and will be available on our";
'    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'    XPrt.CurrentX = col1 + 50
'    XPrt.Print "web site www.bia-paris.fr.";
'XXX DSP - FIN - 29122009
GoTo MSG_FIN
'XXX DSP - FIN - 30102009
MSG_200709:
    XPrt.ForeColor = vbBlue
    XPrt.FontSize = 7
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
    XPrt.CurrentX = col2 + 50
    XPrt.Print "Nous informons notre aimable clientèle que de nouvelles conditions générales seront applicables à compter du 1er janvier 2008,";
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = col2 + 50
    XPrt.Print "et que toutes les échelles d'intérêts seront désormais calculées mensuellement.";
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = col2 + 50
    XPrt.Print "La nouvelle tarification est consultable sur notre site internet : ";
    XPrt.FontBold = True: XPrt.FontUnderline = True
    XPrt.Print "www.bia-paris.com";
    XPrt.FontBold = True:: XPrt.FontUnderline = False
    XPrt.Print ", et est disponible sur simple demande au siège";
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = col2 + 50
    XPrt.Print "de notre établissement.";
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.ForeColor = vbBlack
GoTo MSG_FIN

'------------------------------
MSG_FIN:

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.FontBold = False
    XPrt.FontSize = prtFontSize


End If


Call frmElpPrt.prtTrame(Col4, Line3, Col5, Line4, " ", 250)
'Call frmElpPrt.prtTrame(col1, Line2, Col8, Line3, " ", 240)
prtFillColor = RGB(240, 255, 255)
'Call frmElpPrt.prtTrame_Color(Col4, Line3, Col5, Line4, " ")
Call frmElpPrt.prtTrame_Color(col1, Line2, Col8, Line3, " ")
prtFillColor = prtFillColor_Standard

XPrt.CurrentY = prtMinY + prtlineHeight * 4


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
XPrt.CurrentY = prtMinY + prtlineHeight * 9 - XPrt.TextHeight("test")
'----------------------------------1------------
XPrt.FontBold = True

XPrt.CurrentX = 6200
XPrt.Print xZADRESS0.ADRESSRA1;
'-----------------------------------2-------------
''If Trim(xZADRESS0.ADRESSRA2) <> "" Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = 6200
    XPrt.Print xZADRESS0.ADRESSRA2;
''End If
XPrt.FontBold = False
'------------------------------------3---------------
If Trim(xZADRESS0.ADRESSAD1) <> "" Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = 6200
    XPrt.Print xZADRESS0.ADRESSAD1;
End If
'----------------------------------4-------------------
If Trim(xZADRESS0.ADRESSAD2) <> "" Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = 6200
    XPrt.Print xZADRESS0.ADRESSAD2;
End If

'-----------------------------------5------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 6200
XPrt.Print xZADRESS0.ADRESSAD3;
'------------------------------------6------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 6200
If Trim(xZADRESS0.ADRESSCOP) <> "" Then XPrt.Print xZADRESS0.ADRESSCOP & "  ";
XPrt.Print xZADRESS0.ADRESSVIL;
'------------------------------------8------------------
If Trim(xZADRESS0.ADRESSPAY) <> "FRANCE" Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = 6200
    XPrt.Print xZADRESS0.ADRESSPAY;
End If
'XPrt.FontSize = 6
'XPrt.CurrentX = Col8 - 350
'XPrt.Print "  G " '& recCptInfo.Gestionnaire & "-" & recCptInfo.Courrier;

XPrt.FontSize = 8

XPrt.CurrentY = Line1 - prtlineHeight * 3 + 150
XPrt.FontBold = True

'$$ jpl X = "RELEVE DE COMPTE   " & mYBIAMVT0.COMPTEDEV & "   " & mExtraitNuméro
'$$ jpl Col = Col4 + (Col8 - Col4 - XPrt.TextWidth(X)) / 2
'$$ jpl Call frmElpPrt.prtTrame(Col, XPrt.CurrentY, Col + XPrt.TextWidth(X) + 100, XPrt.CurrentY + prtlineHeight, " ", 240)
'$$ jpl XPrt.CurrentX = Col + 50
'$$ jpl XPrt.Print X;
'$$ jpl XPrt.CurrentY = XPrt.CurrentY + Height8_6
'$$ jpl XPrt.FontBold = False
'$$ jpl XPrt.FontSize = 6

'Call frmElpPrt.prtTrame(col3, XPrt.CurrentY - 100, Col6, XPrt.CurrentY + prtlineHeight, " ", 240)

prtFillColor = RGB(240, 255, 255)
Call frmElpPrt.prtTrame_Color(col3, XPrt.CurrentY - 100, Col6, XPrt.CurrentY + prtlineHeight, " ")
prtFillColor = prtFillColor_Standard

XPrt.CurrentX = col3 + 200

XPrt.CurrentY = XPrt.CurrentY - Height8_6
XPrt.FontSize = 10
XPrt.Print meYBIACPT0.COMPTEDEV;
XPrt.CurrentY = XPrt.CurrentY + Height8_6
XPrt.FontBold = False
XPrt.FontSize = 8

XPrt.Print "  -  RELEVE DE COMPTE : ";
XPrt.FontBold = True
XPrt.CurrentY = XPrt.CurrentY - Height8_6
XPrt.FontSize = 10
XPrt.Print mExtraitNuméro;
XPrt.CurrentY = XPrt.CurrentY + Height8_6
XPrt.FontBold = False
XPrt.FontSize = 8
XPrt.Print "     / " & Format$(NbPage, "###");
XPrt.CurrentX = Col6 + 20
XPrt.Print lRéférence;

'XPrt.Print " -" & Format$(NbPage, "###");  '$$ jpl 2003.03.31$   & " / " & Format$(NbPageMax, "###");
XPrt.CurrentY = XPrt.CurrentY - Height8_6
XPrt.FontSize = 8
'----------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
'XPrt.CurrentX = 800
'XPrt.Print "Numéro ";
XPrt.CurrentX = col1 + 50
'XPrt.Print ": ";
XPrt.FontBold = True
'XPrt.Print Format$(recCptInfo.Numéro, "@@@@@.@@@.@@.@") ;
''XPrt.Print "recCptInfo.Intitulé2";
'-------------------------------------------------------
XPrt.FontBold = False
'Call DevX("recCptInfo.Devise")
XPrt.FontSize = 8
''frmElpPrt.prtCentré (Col4 + Col6) / 2, Trim(mYBIAMVT0.COMPTEDEV)

XPrt.FontSize = 8
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'XPrt.CurrentX = 800
'XPrt.Print "Type";
XPrt.CurrentX = col1 + 50
'XPrt.Print ": ";
XPrt.FontBold = True
'XPrt.Print Trim(DicLib(13, recCptInfo.BiaTyp)) & "-" & Trim(xDevise.DevLib);

'---------------------------------------
'XPrt.FontBold = False

'XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'XPrt.CurrentX = 400
'XPrt.Print "Devise";
'XPrt.CurrentX = Col2
'XPrt.Print ": ";
'XPrt.FontBold = True
'XPrt.Print Format$(recCptInfo.Devise, "000") & "-" & XDevise.DevLib;

'------------------------------------9--------------
'---------------------------------------
XPrt.FontBold = False


XPrt.FontSize = prtFontSize

XPrt.CurrentY = Line1 + 50
XPrt.CurrentX = Col4 - 100 - XPrt.TextWidth(Msg)
XPrt.Print Msg;
prtYBIAMVT0_A4_Montant (solde)

nbLigne = 0
blnA4_Form = True

XPrt.CurrentY = Line3 - prtlineHeight + 50

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
' 20050330 Mettre l'instruction suivante en commentaire : le message date du 31.03.2005
'                                                         A garder jusqu'à 30.09.2005 inclus
' blnConvention_Print = False
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'If blnConvention_Print Then
'    nbLigne = 5
'    blnConvention_Print = False
'    XPrt.FontBold = True
'    XPrt.FontSize = 7
'    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
'    XPrt.CurrentX = col2 + 50
'    XPrt.Print "Nous vous rappelons que nous mettons à votre dispostion une convention de compte";
'    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'    XPrt.CurrentX = col2 + 50
'    XPrt.Print "pour contractualiser tous les aspects de votre relation bancaire.";
'    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'    XPrt.CurrentX = col2 + 50
'    XPrt.Print "Cette convention reprend les conditions générales de fonctionnenment de votre compte ainsi que";
'    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'    XPrt.FontBold = False
'    XPrt.FontSize = prtFontSize
'
'    XPrt.Print "La BIA a mis à jour les conditions générales appliquées";
'    XPrt.Print "à la clientèle le 1er juillet 2005. La nouvelle grille tarifaire ";
'    XPrt.Print "est à votre disposition à nos guichets.";


'End If



End Sub

'---------------------------------------------------------
Public Sub prtYBIAMVT0_A4_Line()
'---------------------------------------------------------
Dim X As String, I As Integer, libCV As String, blnCV As Boolean
Dim blnLine2 As Boolean, xLine1 As String, xLine2 As String
Dim kJust As Integer, kMax As Integer
Dim widthCOL3_5 As Integer
Dim blnFRS_Color As Boolean
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
 
XPrt.FontSize = prtFontSize
XPrt.FontBold = False
XPrt.ForeColor = vbBlack
'If Mid$(xYBIAMVT0.MOUVEMANA, 1, 3) = "FRS" And Trim(meYBIACPT0.CLIENASRN) = ""
If Mid$(xYBIAMVT0.MOUVEMANA, 1, 3) = "FRS" _
And xYBIAMVT0.MOUVEMMON > 0 Then
    XPrt.ForeColor = RGB(0, 0, 160)
    curFrs_Total = curFrs_Total + xYBIAMVT0.MOUVEMMON
    blnFRS_Color = True
Else
    blnFRS_Color = False
End If

XPrt.CurrentX = col1 + 50
XPrt.Print dateImp(Val(xYBIAMVT0.MOUVEMDTR) + 19000000);

XPrt.CurrentX = col3 + 50
XPrt.Print dateImp(Val(xYBIAMVT0.MOUVEMDVA) + 19000000);
prtYBIAMVT0_A4_Montant (xYBIAMVT0.MOUVEMMON)

If blnFRS_Color Then blnFRS_Info = True: XPrt.CurrentX = XPrt.CurrentX + 150: XPrt.Print "#";

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
    If nbLigne = NbLigneMax Then prtYBIAMVT0_A4_Report
    nbLigne = nbLigne + 1
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = col2 + 50
    XPrt.Print xLine2;
End If

XPrt.ForeColor = vbBlack

'XPrt.FontUnderline = False: XPrt.FontItalic = False

'blnLine2 = True

'For I = prtFontSize To 6 Step -1  ' 6
'    XPrt.FontSize = I
'    If XPrt.TextWidth(X) <= (col3 - XPrt.CurrentX - 100) Then blnLine2 = False: Exit For
'Next I

'If Not blnLine2 Then
'    XPrt.Print X;
'Else
    'XPrt.FontSize = prtFontSize
   ' If XPrt.TextWidth(xLine1) > (col3 - XPrt.CurrentX - 100) Then XPrt.FontSize = prtFontSize - 1
'    XPrt.Print xLine1;
'    If nbLigne = NbLigneMax Then prtYBIAMVT0_A4_Report
'    nbLigne = nbLigne + 1
'    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'    XPrt.CurrentX = col2 + 50
    'If XPrt.TextWidth(xLine2) > (col3 - XPrt.CurrentX - 100) Then XPrt.FontSize = prtFontSize - 1
'    XPrt.Print xLine2;
'End If

XPrt.FontSize = prtFontSize

End Sub


'---------------------------------------------------------
Public Sub prtYBIAMVT0_A4_Montant(MT As Currency)
'---------------------------------------------------------
Dim X As String

XPrt.FontBold = True
X = Format$(Abs(MT), "## ### ### ### ### ##0.00")
XPrt.CurrentX = IIf(MT < 0, Col6, Col5) - 100 - XPrt.TextWidth(X)
XPrt.Print X;
XPrt.FontBold = False

End Sub

'---------------------------------------------------------
Public Sub prtYBIAMVT0_A4_Médiateur()
'---------------------------------------------------------
Dim X As String
XPrt.FontSize = 7
XPrt.CurrentY = 15231: XPrt.Print "";
'XPrt.CurrentY = XPrt.CurrentY + (prtlineHeight / 2)     ' TODO + 50

'prtFillColor = RGB(240, 255, 255)
'Call frmElpPrt.prtTrame_Color(col1, XPrt.CurrentY, Col8, XPrt.CurrentY + prtlineHeight * 2.3, " ")
'prtFillColor = prtFillColor_Standard

XPrt.ForeColor = RGB(0, 0, 160)

XPrt.FontBold = False
'XPrt.CurrentY = XPrt.CurrentY + 20 ' TODO 100
XPrt.CurrentX = col1 + 200
XPrt.Print "Nous vous informons qu'un médiateur est à votre disposition à l'adresse suivante : ";
XPrt.FontBold = True
XPrt.Print "  M. le MEDIATEUR   -   CS 151   -   75422 PARIS CEDEX 09";

XPrt.CurrentY = XPrt.CurrentY + 180 ' TODO prtlineHeight
XPrt.CurrentX = col2 + 50
XPrt.FontBold = False
frmElpPrt.prtCentré prtMedX, "pour tout problème que vous n'avez pu résoudre préalablement avec la banque."

XPrt.ForeColor = vbBlack
XPrt.FontSize = 8


End Sub


'---------------------------------------------------------
Public Sub prtYBIAMVT0_A4_Pauget_Constans(lRELEVEREL As String)
'---------------------------------------------------------
Dim X As String

XPrt.FontSize = 7

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
prtFillColor = RGB(240, 255, 255)
Call frmElpPrt.prtTrame_Color(col1, XPrt.CurrentY, Col8, XPrt.CurrentY + prtlineHeight * 1.3, " ")
prtFillColor = prtFillColor_Standard
XPrt.ForeColor = RGB(0, 0, 160)

If blnPauget_Constans Then 'blnMédiateur Then
    XPrt.FontBold = False
    XPrt.CurrentY = XPrt.CurrentY + 100
    XPrt.CurrentX = col1
    If lRELEVEREL = "M" Then
        XPrt.Print "# Total mensuel des frais bancaires : ";
    Else
        XPrt.Print "# Total des frais bancaires : ";
    End If
    
    XPrt.FontBold = True
    If curFrs_Total <> 0 Then
        XPrt.Print Format(curFrs_Total, "### ### ##0.00") & " " & meYBIACPT0.COMPTEDEV;
    Else
        XPrt.Print "néant";
    End If
End If

XPrt.CurrentX = col1 + 3800
XPrt.FontBold = False
XPrt.Print "Montant de l'autorisation de découvert : ";
XPrt.FontBold = True
If mZAUTSYC0.AUTSYCMON <> 0 Then
    XPrt.Print Format(mZAUTSYC0.AUTSYCMON, "### ### ##0.00") & " " & mZAUTSYC0.AUTSYCDEV;
    XPrt.FontBold = False
    XPrt.Print ""; '"  jusqu'au ";
    XPrt.FontBold = True
    XPrt.Print ""; 'dateImp10(mZAUTSYC0.AUTSYCFIN + 19000000);
    
    XPrt.CurrentX = Col5 + 200
    
    prtYBIAMVT0_A4_Pauget_Constans_TAEG
    XPrt.FontBold = False
    If blnMédiateur Then
        XPrt.Print "  TAEG : ";
    Else
        XPrt.Print "  TEG : ";
    End If
    
    XPrt.FontBold = True
    If dblTAEG > 0 And dblTAEG < 30 Then
        XPrt.Print Format(dblTAEG, "#0.00000") & " %";
    Else
        XPrt.Print "    %";
    End If
    
    If Not blnAUTSYCAUT_Compte Then
        'XPrt.FontSize = 6
        XPrt.CurrentY = XPrt.CurrentY + prtlineHeight - 100
        XPrt.CurrentX = col1 + 3800
        XPrt.FontBold = False
        XPrt.Print "(pour l'ensemble de vos comptes à vue)";
        XPrt.CurrentY = XPrt.CurrentY - prtlineHeight + 100
       ' XPrt.FontSize = 7
    End If
    
Else
    XPrt.Print "néant";
End If

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight / 2
XPrt.FontBold = False

XPrt.ForeColor = vbBlack
XPrt.FontSize = 8

End Sub


Public Sub prtYBIAMVT0_A4_Extrait(lMOUVEMCOM As String, lAMJMin As String, lAMJMax As String, blnCSV As Boolean, lstErr As ListBox, lRELEVEREL As String, lRéférence As String, blnNewPage As Boolean)
'---------------------------------------------------------
Dim rsLocal As ADODB.Recordset, rsW As ADODB.Recordset
Dim xSQL As String
Dim Nb As Integer
Dim CTLAMJ As String
Dim V
Dim wAMJ_Solde As String
Dim blnInfoFGDR As Boolean

valAmjMin = lAMJMin
valAmjMax = lAMJMax
IbmAmjMin = dateIBM(lAMJMin)
IbmAmjMax = dateIBM(lAMJMax)
wAMJ_Solde = dateElp("Jour", -1, valAmjMin)
rsZADRESS0_Init zZADRESS0
blnNewPage = False
blnFRS_Info = False


'lMOUVEMCOM = "50155978001"

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIACPT0" _
     & " where COMPTECOM = '" & lMOUVEMCOM & "'"
     
Set rsLocal = Nothing
Set rsLocal = cnsab.Execute(xSQL)
V = rsYBIACPT0_GetBuffer(rsLocal, meYBIACPT0)
If Not IsNull(V) Then
    MsgBox "prtYBIAMVT0_A4_Extrait " & V
    Exit Sub
End If

blnInfoFGDR = False
If retourne_Eligibilite(meYBIACPT0.CLIENACLI, meYBIACPT0.COMPTECOM) Then
    blnInfoFGDR = True
End If
prtYBIAMVT0_A4_Compte lRELEVEREL


xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIAMVTH" _
     & " where MOUVEMCOM = '" & lMOUVEMCOM & "'" _
     & " and MOUVEMDTR >= " & IbmAmjMin _
     & " and MOUVEMDTR <= " & IbmAmjMax _
     & " order by MOUVEMDTR, MOUVEMPIE, MOUVEMECR"
     
Set rsLocal = Nothing
Set rsLocal = cnsab.Execute(xSQL)
If rsLocal.EOF Then
' pas de mouvement dans la période
    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIAMVTH" _
         & " where MOUVEMCOM = '" & lMOUVEMCOM & "'" _
         & " and MOUVEMDTR >= " & IbmAmjMax _
         & " order by MOUVEMDTR, MOUVEMPIE, MOUVEMECR"
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
        MsgBox "prtYBIAMVT0_A4_Extrait " & V
        Exit Sub
    End If
    
    xYBIAMVT0 = mYBIAMVT0
    CTLAMJ = mYBIAMVT0.MOUVEMDTR
    solde = mYBIAMVT0.BIAMVTSD0
End If

prtYBIAMVT0_A4_OpenX_Reset

blnA4_Form = False
If lRELEVEREL = "M" Then
    mExtraitNuméro = libMois(Mid$(lAMJMax, 5, 2))
Else
    mExtraitNuméro = "__________"
End If

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
               
               If CTLAMJ <> xYBIAMVT0.MOUVEMDTR Then
                       If solde <> xYBIAMVT0.BIAMVTSD0 Then
                           XPrt.CurrentX = col2
                           MsgBox "erreur Solde .........", vbCritical, "prtCptMvt"
                           XPrt.FontSize = 14
                           XPrt.Print "ERREUR SOLDE ............."
                           Exit Do
                       End If
                   CTLAMJ = xYBIAMVT0.MOUVEMDTR
               End If
               
               
                If xYBIAMVT0.MOUVEMDTR >= IbmAmjMin Then
           
                    If Not blnA4_Form Then
                        prtYBIAMVT0_A4_Form "Solde au : " & dateImp(wAMJ_Solde), lRéférence
                        blnNewPage = True
                    End If
                    If nbLigne = NbLigneMax Then
                        prtYBIAMVT0_A4_Report
                        lstErr.RemoveItem lstErr.ListCount - 1
                        lstErr.AddItem xYBIAMVT0.MOUVEMCOM & " page : " & NbPage
                    End If
                    nbLigne = nbLigne + 1
                    
                    prtYBIAMVT0_A4_Line
                    
                    If blnCSV Then prtYBIAMVT0_CSV

                    DoEvents: If prtKillDoc Then frmElpPrt.Hide: Exit Sub
                End If
            
            solde = solde + xYBIAMVT0.MOUVEMMON
            
           rsLocal.MoveNext
           Call rsYBIAMVT0_GetBuffer(rsLocal, xYBIAMVT0)

Loop


'Pas de mouvement dans la période : imprimer un extrait (sauf mensuel)
If Not blnA4_Form Then
    'If lRELEVEREL <> "M" Then prtYBIAMVT0_A4_Form "Solde au : " & dateImp(dateElp("Jour", -1, valAmjMin)), lRéférence
    'prtYBIAMVT0_A4_Form "Solde au : " & dateImp(meYBIACPT0.SOLDEDMO + 19000000), lRéférence
    prtYBIAMVT0_A4_Form "Solde au : " & dateImp(dateElp("Jour", -1, lAMJMin)), lRéférence
    
End If

If blnA4_Form Then
    'XPrt.CurrentY = Line4 + 50
    XPrt.CurrentY = 13950
    X = "Solde au : " & dateImp(valAmjMax)
    XPrt.CurrentX = Col4 - XPrt.TextWidth(X) - 200
    XPrt.Print X;
    XPrt.CurrentX = 5000
    prtYBIAMVT0_A4_Montant (solde)
    
    '$JPL 2011.05.25 Pauget_Constans
    If blnPauget_Constans Then
        prtYBIAMVT0_A4_Pauget_Constans lRELEVEREL
    Else
        If blnFRS_Info Then
            XPrt.ForeColor = RGB(0, 0, 160)
            XPrt.CurrentX = col2
            XPrt.Print "# : frais bancaires";
            XPrt.ForeColor = vbBlack
        End If
    End If
    
   
    '$JPL 2002.12.26 médiateur
    If blnInfoFGDR Then
        Call prtYBIAMVT0_A4_InfoFGDR
        Call prtYBIAMVT0_A4_Médiateur
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

prtYBIAMVT0_A4_Close
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

Public Sub prtYBIAMVT0_A4_RIB()
Dim iY As Integer
'--------------------------TRAME---------------------------
Dim X As String

XPrt.DrawWidth = 1
iY = 1500
'Call frmElpPrt.prtTrame(200, iY, 4750, iY + 250, "", 240)

'Call frmElpPrt.prtTrame(200, iY + 1450, 4750, iY + 1700, "B", 240)

prtFillColor = RGB(240, 255, 255)
Call frmElpPrt.prtTrame_Color(200, iY, 4750, iY + 250, "")
Call frmElpPrt.prtTrame_Color(200, iY + 1450, 4750, iY + 1700, "B")
prtFillColor = prtFillColor_Standard

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


Public Sub prtYBIAMVT0_A4_Compte(lRELEVEREL As String)
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

If Mid$(meYBIACPT0.COMPTEOBL, 1, 4) = "2511" Then
    blnPauget_Constans = True
Else
    blnPauget_Constans = False
End If

'$JPL TEST 20110530Call MsgBox("lRELEVEREL = M: blnMédiateur = True", vbExclamation, "TEST JPL")
'$JPL TEST 20110530lRELEVEREL = "M": blnMédiateur = True

If lRELEVEREL = "M" And meYBIACPT0.COMPTEDEV = "EUR" Then
        'blnPauget_Constans = blnMédiateur
        If blnPauget_Constans Then prtYBIAMVT0_A4_Pauget_Constans_Init
Else
    If valAmjMax = YBIATAB0_DATE_CPT_J Then
        'blnPauget_Constans = blnMédiateur
        If blnPauget_Constans Then prtYBIAMVT0_A4_Pauget_Constans_Init
    End If
End If

'__________________________________________________________________________________________
'blnConvention_Print = blnMédiateur
' 20050330 filtre des PCI pour impression d'un message sur les extraits de compte
'If Mid$(meYBIACPT0.COMPTEOBL, 1, 6) = "121209" Then     ' Ajout le 30.06.2005
'    blnConvention_Print = True
'Else
'    Select Case Mid$(meYBIACPT0.COMPTEOBL, 1, 5)
'        Case "25111", "25113", "25112", "25114", "25115", "25117": blnConvention_Print = True
'    End Select
'End If
' 20051123 filtre des PCI pour impression d'un message sur les extraits de compte
'JPL Supprimé le 23.03.2006
'--------------------------------
'blnConvention_Print = False
'If Mid$(meYBIACPT0.COMPTEOBL, 1, 5) = "25111" Then blnConvention_Print = True
'If Mid$(meYBIACPT0.COMPTEOBL, 1, 5) = "25117" Then blnConvention_Print = True
'__________________________________________________________________________________________

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

Public Sub prtYBIAMVT0_A4_Report()
XPrt.CurrentY = Line4 + 50
prtYBIAMVT0_A4_Montant (solde)
NbPage = NbPage + 1
frmElpPrt.prtNewPage
prtYBIAMVT0_A4_Form "Report", ""

End Sub

Public Sub prtYBIAMVT0_A4_OpenX_Reset()

Set XPrt = Printer
frmElpPrt.Show vbModeless

Height8_6 = frmElpPrt.prtHeightDelta(8, 6)


prtTitleText = "Extrait de Compte"
prtPgmName = "prtYBIAMVT0_A4"
prtTitleUsr = usrName
prtFontName = "Calibri" 'prtFontName_Arial

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
prtYBIAMVT0_A4_OpenX_Reset_Line " "

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

Public Sub prtYBIAMVT0_A4_OpenX_Reset_Line(lFct As String)
If lFct = "M" Then
    NbLigneMax = 27
    Line1 = prtlineHeight * 29
Else
    NbLigneMax = 32 '35
    Line1 = prtlineHeight * 21
End If
'If blnPauget_Constans Then
'    NbLigneMax = NbLigneMax - 3
'End If
Line2 = Line1 + prtlineHeight + 50
Line3 = Line2 + prtlineHeight + 50
Line4 = Line3 + prtlineHeight * NbLigneMax + 50
Line5 = Line4 + prtlineHeight + 50

End Sub
Public Sub prtYBIAMVT0_A4_Pauget_Constans_Init()
Dim rsW As ADODB.Recordset, xSQL As String
Dim xCOMPTECOM As String
blnAUTSYCAUT_Dec = False: blnAUTSYCAUT_Compte = False
curFrs_Total = 0: mZAUTSYC0.AUTSYCMON = 0: dblTAEG = 0
xCOMPTECOM = Trim(meYBIACPT0.COMPTECOM)

'xSQL = "select * from " & paramIBM_Library_SAB & ".ZAUTSYC0" _
     & " where AUTSYCCLI = '" & meYBIACPT0.CLIENACLI & "' and AUTSYCAUT in ('DEC' , '" & xCOMPTECOM & "')"
xSQL = "select * from " & paramIBM_Library_SAB & ".ZAUTSYC0" _
     & " where AUTSYCCLI = '" & meYBIACPT0.CLIENACLI & "' and AUTSYCAUT in ('" & xCOMPTECOM & "')"
     
Set rsW = cnsab.Execute(xSQL)

Do While Not rsW.EOF

    If rsW("AUTSYCMON") <> 0 Then
        Select Case Trim(rsW("AUTSYCAUT"))
            Case "DEC":
                        If Not blnAUTSYCAUT_Compte Then
                            blnAUTSYCAUT_Dec = True
                            Call rsZAUTSYC0_GetBuffer(rsW, mZAUTSYC0)
                        End If
            Case xCOMPTECOM:
                        
                            blnAUTSYCAUT_Compte = True
                            Call rsZAUTSYC0_GetBuffer(rsW, mZAUTSYC0)
        End Select
    End If
    
    rsW.MoveNext
Loop

If blnAUTSYCAUT_Compte Or blnAUTSYCAUT_Dec Then
    If mZAUTSYC0.AUTSYCFIN < IbmAmjMax And mZAUTSYC0.AUTSYCFIN <> 0 Then
        mZAUTSYC0.AUTSYCMON = 0
    End If
End If
'==============================================================================

'____________________________________________________________________

End Sub

Public Sub prtYBIAMVT0_A4_Pauget_Constans_TAEG()
Dim rsW As ADODB.Recordset, xSQL As String
Dim wAmj As Long, wCours As Double, wMarge As Double
Dim X As String, xCode As String, xFiscal As String
Dim blnOk As Boolean, blnZECHTAB0 As Boolean, xECHTABDON As String

'======================================================================
dblTAEG = 0: wMarge = 0
blnOk = False
blnZECHTAB0 = False
xSQL = "select * from " & paramIBM_Library_SAB & ".ZECHTAB0" _
     & " where ECHTABETA = " & currentZMNURUT0.MNURUTETB _
     & " and   ECHTABNUM = 9" _
     & " and   ECHTABARG like '%" & meYBIACPT0.COMPTECOM & "   IDE%'"
Set rsW = cnsab.Execute(xSQL)


Do While Not rsW.EOF
    xECHTABDON = rsW("ECHTABDON")
    If Mid$(xECHTABDON, 219, 1) = " " Then blnZECHTAB0 = True: Exit Do
    rsW.MoveNext
Loop

'______________________________________________________________________
If Not blnZECHTAB0 Then
    xFiscal = prtYBIAMVT0_A4_Pauget_Constans_Fiscal(Trim(meYBIACPT0.CLIENARSD))
    X = Space(66)
    Mid$(X, 1, 3) = "CAV"
    Mid$(X, 21, 3) = "EUR"
    Mid$(X, 31, 3) = meYBIACPT0.CLIENACAT
    Mid$(X, 41, 1) = xFiscal
    Mid$(X, 64, 3) = "IDE"
    
    xSQL = "select * from " & paramIBM_Library_SAB & ".ZECHTAB0" _
         & " where ECHTABETA = " & currentZMNURUT0.MNURUTETB _
         & " and   ECHTABNUM = 8" _
         & " and   ECHTABARG like '" & X & "%'"
    Set rsW = cnsab.Execute(xSQL)
    Do While Not rsW.EOF
        xECHTABDON = rsW("ECHTABDON")
        If Mid$(xECHTABDON, 219, 1) = " " Then blnZECHTAB0 = True: Exit Do
        rsW.MoveNext
    Loop
End If
'______________________________________________________________________
If Not blnZECHTAB0 Then
    Mid$(X, 41, 1) = " "
    
    xSQL = "select * from " & paramIBM_Library_SAB & ".ZECHTAB0" _
         & " where ECHTABETA = " & currentZMNURUT0.MNURUTETB _
         & " and   ECHTABNUM = 8" _
         & " and   ECHTABARG like '" & X & "%'"
    Set rsW = cnsab.Execute(xSQL)
        
    Do While Not rsW.EOF
        xECHTABDON = rsW("ECHTABDON")
        If Mid$(xECHTABDON, 219, 1) = " " Then blnZECHTAB0 = True: Exit Do
        rsW.MoveNext
    Loop

End If

'______________________________________________________________________
If blnZECHTAB0 Then
    xCode = Mid$(xECHTABDON, 25, 6)
    wMarge = CDbl(convX2P(Mid$(xECHTABDON, 35, 5))) / 1000000
End If

'======================================================================

For arrECHTAB_K = 1 To arrECHTAB_Nb
    If xCode = arrECHTAB(arrECHTAB_K).Code Then blnOk = True: Exit For
Next arrECHTAB_K


If Not blnOk Then
    xSQL = "select * from " & paramIBM_Library_SAB & ".ZBASTAB0" _
         & " where BASTABETA = " & currentZMNURUT0.MNURUTETB _
         & " and   BASTABNUM = 25" _
         & " and   BASTABARG like 'EUR" & xCode & "%'" _
         & " order by BASTABARG"
    Set rsW = cnsab.Execute(xSQL)
    If Not rsW.EOF Then
        arrECHTAB_Nb = arrECHTAB_Nb + 1
        arrECHTAB_K = arrECHTAB_Nb
        arrECHTAB(arrECHTAB_Nb).Code = xCode
        arrECHTAB(arrECHTAB_Nb).Taux = 0
        arrECHTAB(arrECHTAB_Nb).AMJ7 = 0
        Do While Not rsW.EOF
            wAmj = 19000000 + convX2P(Mid$(rsW("BASTABARG"), 10, 4))
            If wAmj > arrECHTAB(arrECHTAB_Nb).AMJ7 And wAmj <= valAmjMax Then
                blnOk = True
                arrECHTAB(arrECHTAB_Nb).AMJ7 = wAmj
                arrECHTAB(arrECHTAB_Nb).Taux = CDbl(convX2P(Mid$(rsW("BASTABDON"), 1, 8))) / 1000000000
            End If
            rsW.MoveNext
        Loop
    End If
End If

If blnOk Then
    dblTAEG = ((1 + (wMarge + arrECHTAB(arrECHTAB_K).Taux) / 400) ^ 4 - 1) * 100

End If
End Sub

Public Function prtYBIAMVT0_A4_Pauget_Constans_Fiscal(lPays As String) As String
Static K As Integer
If arrPays_NB = 0 Then
    Call rsZBASTAB0_Pays(arrPays(), arrPays_NB)
    K = 0: arrPays(0).Id = "?"
End If
'___________________________________________________________________________
If lPays = arrPays(K).Id Then
    prtYBIAMVT0_A4_Pauget_Constans_Fiscal = arrPays(K).Fiscal
Else
    prtYBIAMVT0_A4_Pauget_Constans_Fiscal = ""
    For K = 1 To arrPays_NB
        If lPays = arrPays(K).Id Then prtYBIAMVT0_A4_Pauget_Constans_Fiscal = arrPays(K).Fiscal: Exit For
    Next K
    If K > arrPays_NB Then K = 1
End If
End Function







