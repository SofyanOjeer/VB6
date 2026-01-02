Attribute VB_Name = "prtCompta"
Option Explicit
Public CV1 As typeCV, CV2 As typeCV, CV3 As typeCV
Public arrCV030(10) As typeCpj030W0
Dim recCV030 As typeCpj030W0
Dim recCptMvt As typeCptMvt

Dim I As Integer, solde As Currency, mCurrenty As Integer
Dim Line1 As Integer, Line2 As Integer, Line3 As Integer, Line4 As Integer, Line5 As Integer
Dim Col1 As Integer, Col2 As Integer, Col3 As Integer
Dim Col4 As Integer, Col5 As Integer, Col6 As Integer, Col7 As Integer, Col8 As Integer
Dim Col As Integer, Height8_6 As Integer
Dim K As Integer, K1 As Integer, K2 As Integer, K3 As Integer
Dim X As String
Dim nbLigne As Integer, NbPage As Integer
Dim NbLigneMax As Integer, NbPageMax As Integer
Dim NbImprimé As Integer

Dim recCompte As typeCompte
Dim curCumulDébit As Currency, curCumulCrédit As Currency
Dim mDevise As String
Dim blnCVàImprimer As Boolean, blnForm As Boolean
Dim mCptMvtPièce As Long
Dim iReturn As Integer
Dim xLong As Long, prtMaxY_4 As Integer

Dim blnValidation As Boolean
Dim blnRéférence As Boolean, mRéférence As String

Dim mCurrentY_Opération As Integer
Dim mComptaUsr As String

Dim blnGuichet As Boolean, blnJournal As Boolean, blnCptMvtPièce_Rupture As Boolean
Dim blnTotalSolde As Boolean, blnEAR As Boolean, blnEAR_Imp As Boolean

Public mJournal As String
Dim blnSoldeProvisoire_Print As Boolean
'---------------------------------------------------------
Public Sub prtCompta_Monitor(Msg As String, Text As String, strTitle As String)
'---------------------------------------------------------
On Error GoTo prtError


Set XPrt = Printer
K1 = Val(mId$(Msg, 1, 6))
K2 = Val(mId$(Msg, 7, 6))
Height8_6 = frmElpPrt.prtHeightDelta(8, 6)


frmElpPrt.Show vbModeless

prtOrientation = vbPRORLandscape
prtPgmName = "prtCompta"
prtTitleUsr = usrName
prtTitleText = strTitle & " : " & Trim(Text)
 
'Select Case strTitle
'    Case conststrGuichet_Compta
'        prtTitleText = strTitle & " : " & " Opérations en instance de validation"
'    Case Else: prtTitleText = strTitle & " : " & Trim(Text)
'End Select

prtLineNb = 1
prtlineHeight = 280
prtHeaderHeight = 900
nbLigne = 0: NbPage = 1
curCumulDébit = 0: curCumulCrédit = 0: solde = 0
Col4 = 7000: Col5 = 8700: Col6 = 10300

If strTitle = "Service_Pièce" Or strTitle = "Service_Compte" Or strTitle = "Liste_Compte" Then
   
    prtTitleText = mId$(Msg, 13, 3) & "_" & Trim(DicLib(4, mId$(Msg, 13, 3))) & " :  " & Trim(Text)
    prtHeaderHeight = 300
End If

frmElpPrt.prtStdInit
blnEAR = False: blnEAR_Imp = False
blnRéférence = False
blnCVàImprimer = False
blnJournal = False
blnTotalSolde = True
blnValidation = IIf(Trim(Text) = constDemandeDeValidation, True, False)
recCpj030W0_Init recCV030
prtMaxY_4 = prtMaxY - prtlineHeight * 4

mComptaUsr = ""
blnCptMvtPièce_Rupture = True
blnSoldeProvisoire_Print = False

Select Case strTitle

    Case conststrGuichet_Compta, conststrTFlux_Compta
                    blnRéférence = True: blnGuichet = True
                ''    mComptaUsr = K2
                    prtGuichet_Compta
     Case conststrGuichet_Comptabilisé, conststrTFlux_Comptabilisé
                    blnRéférence = True: blnGuichet = True
                    prtGuichet_Comptabilisé Msg
     Case "Service_Pièce"
                    blnGuichet = False: blnJournal = True
                    prtMVTP0 Msg
                    If blnEAR Then
                        blnEAR_Imp = True
                        prtTitleText = mId$(Msg, 13, 3) & "_" & Trim(DicLib(4, mId$(Msg, 13, 3))) & " : E A R"
                        frmElpPrt.prtNewPage
                        prtMVTP0 Msg
                    End If
    Case "Service_Compte"
                    blnGuichet = False: blnJournal = True
                    blnCptMvtPièce_Rupture = False
                    prtMVTP0 Msg
                    If blnEAR Then
                        blnEAR_Imp = True
                        prtTitleText = mId$(Msg, 13, 3) & "_" & Trim(DicLib(4, mId$(Msg, 13, 3))) & " : E A R"
                        frmElpPrt.prtNewPage
                        prtMVTP0 Msg
                    End If
                   
    Case "Liste_Compte"
                    blnGuichet = False: blnJournal = True
                    blnCptMvtPièce_Rupture = False
                    blnTotalSolde = False
                    prtMVTP0 Msg
Case Else: blnCVàImprimer = True
                    prtCompta_Lot

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
'---------------------------------------------------------
Public Sub prtCompta_Lot()
'---------------------------------------------------------
recCV030 = arrCV030(1)
mDevise = recCV030.Devise
prtCompta_Form
prtCompta_Devise

For K = K1 To K2
    If arrCV030(K).MONDEV <> 0 Then
        recCV030 = arrCV030(K)
        prtCompta_Line
    End If
Next K

prtCompta_Total
prtCompta_Trait
If blnValidation Then prtCompta_Validation "VALIDATION"

End Sub


'---------------------------------------------------------
Public Sub prtGuichet_Compta()
'---------------------------------------------------------
mCptMvtPièce = -1
blnForm = True
recGuichet_Compta.Method = "MoveFirst" ' "Seek=" '
curCumulDébit = 0
curCumulCrédit = 0
G_CV2.Montant = 0
recCV030.NUMLOT = ""

Do
    iReturn = tableGuichet_Compta_Read(recGuichet_Compta)
    If iReturn = 0 Then
        recCV030.COSOC = recGuichet_Compta.Société
        recCV030.Agence = recGuichet_Compta.Agence
        recCV030.Devise = recGuichet_Compta.Devise
 ''''       recCV030.BIACOP = recGuichet_Compta.CodeOpération
 '======== Demande de compta Guichet : imprimer solde provisoire
       If Trim(recGuichet_Compta.CodeOpération) = "G008" And recGuichet_Compta.Compte > "10000000000" Then
            blnSoldeProvisoire_Print = True
        Else
             blnSoldeProvisoire_Print = False
       End If
        
        If IsNumeric(recGuichet_Compta.ComptaUsr) Then recCV030.NUMLOT = recGuichet_Compta.ComptaUsr
        recCV030.NUMPIE = recGuichet_Compta.CptMvtPièce Mod 10000
        recCV030.NOLIGN = recGuichet_Compta.CptMvtLigne Mod 1000
        recCV030.NOMOP = recGuichet_Compta.SaisieUsr
        recCV030.SERVIC = recGuichet_Compta.Service
        recCV030.Compte = recGuichet_Compta.Compte
        recCV030.MONDEV = recGuichet_Compta.Montant
        recCV030.SENECR = recGuichet_Compta.Sens
        recCV030.AMJOPE = recGuichet_Compta.AmjOpération
        recCV030.AMJVAL = recGuichet_Compta.AmjValeur
        recCV030.LIBELE = recGuichet_Compta.Libellé
        recCV030.AMJSAI = recGuichet_Compta.SaisieAmj
        recCV030.CODFOR = recGuichet_Compta.chkCompte
        recCV030.FOROPO = recGuichet_Compta.chkSolde
        recCV030.OPOCHQ = recGuichet_Compta.chkChèque
        recCV030.FORVAL = recGuichet_Compta.chkAmjValeur
        If mCptMvtPièce <> recGuichet_Compta.CptMvtPièce Then
    '''        If Not blnForm Then
     ''           prtCompta_Total
     ''           prtCompta_Trait
    '''            If blnValidation Then prtCompta_Validation
    ''''           frmElpPrt.prtNewPage
    '        End If
    ''        blnForm = True
            XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
            XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY)
            XPrt.CurrentY = XPrt.CurrentY + 20 - prtlineHeight
            mCptMvtPièce = recGuichet_Compta.CptMvtPièce
        End If
        
        If blnForm Then
            blnForm = False
            prtGuichet_Compta_Form
            mDevise = recCV030.Devise
            prtCompta_Devise
        End If
        mRéférence = recGuichet_Compta.Référence
        prtCompta_Line
    End If
    recGuichet_Compta.Method = "MoveNext"
Loop Until iReturn <> 0

prtCompta_Total
prtCompta_Trait
If blnValidation Then prtCompta_Validation "VALIDATION du lot N° : " & mComptaUsr

End Sub
'---------------------------------------------------------
Public Sub prtGuichet_Opération()
'---------------------------------------------------------
Dim xxxDevise As String

'mCptMvtPièce = -1
blnForm = True
recGuichet_Compta.Method = "MoveFirst"
curCumulDébit = 0
curCumulCrédit = 0

Do
    iReturn = tableGuichet_Compta_Read(recGuichet_Compta)
    If iReturn = 0 Then
        If recGuichet_Compta.Société <> "000" Then
            iReturn = 1
        Else
           recCV030.Devise = recGuichet_Compta.Devise
           recCV030.BIACOP = recGuichet_Compta.CodeOpération
           recCV030.NUMPIE = recGuichet_Compta.CptMvtPièce
           recCV030.NOLIGN = recGuichet_Compta.CptMvtLigne Mod 10000
           recCV030.NOMOP = recGuichet_Compta.SaisieUsr
           recCV030.SERVIC = recGuichet_Compta.Service
           recCV030.Compte = recGuichet_Compta.Compte
           recCV030.MONDEV = recGuichet_Compta.Montant
           recCV030.SENECR = recGuichet_Compta.Sens
           recCV030.AMJOPE = recGuichet_Compta.AmjOpération
           recCV030.AMJVAL = recGuichet_Compta.AmjValeur
           recCV030.LIBELE = recGuichet_Compta.Libellé
           recCV030.AMJSAI = recGuichet_Compta.SaisieAmj
           recCV030.CODFOR = recGuichet_Compta.chkCompte
           recCV030.FOROPO = recGuichet_Compta.chkSolde
           recCV030.OPOCHQ = recGuichet_Compta.chkChèque
           recCV030.FORVAL = recGuichet_Compta.chkAmjValeur
           
           G_CV2.DeviseIso = recGuichet_Compta.Devise2
           G_CV2.Montant = recGuichet_Compta.Montant2
           G_CV3.Montant = recGuichet_Compta.Montant3
        
           If blnForm Then
               blnForm = False
               prtCompta_Form
               xxxDevise = recGuichet_Compta.Agence
               mDevise = xxxDevise
               prtCompta_Devise
           End If
            
            If xxxDevise <> recGuichet_Compta.Agence Then
                prtCompta_Total
                xxxDevise = recGuichet_Compta.Agence
                mDevise = xxxDevise
                prtCompta_Devise
            End If
 
            mRéférence = recGuichet_Compta.Référence
           mDevise = recCV030.Devise
           recCV030.COSOC = SocId$
           recCV030.Agence = SocAgence$
           prtCompta_Line
        End If
    End If
    recGuichet_Compta.Method = "MoveNext"
Loop Until iReturn <> 0

prtCompta_Total
prtCompta_Trait
If blnValidation Then prtCompta_Validation "VALIDATION"
frmElpPrt.prtNewPage

End Sub


'---------------------------------------------------------
Public Sub prtCompta_Form()
'---------------------------------------------------------
nbLigne = 0: NbImprimé = 0
XPrt.DrawWidth = 1
''Call frmElpPrt.prtTrame(Col4, prtMinY + 10, Col6, prtMinY + 10 + prtHeaderHeight, "B")

XPrt.DrawWidth = 3

Call frmElpPrt.prtTrame(prtMinX, prtMinY + prtHeaderHeight, prtMaxX, prtMinY + prtHeaderHeight + prtlineHeight, "B", 235)

'---------------------------------------------------------

XPrt.CurrentY = prtMinY - prtlineHeight + 100
If Not blnJournal Then
    prtCompta_CV CV1
    
    XPrt.CurrentX = 10400
    XPrt.Print "Service";
    XPrt.CurrentX = 11000
    XPrt.Print ": " & Trim(recCV030.SERVIC) & "_" & Trim(DicLib(4, recCV030.SERVIC));
    
    XPrt.Print " / " & recCV030.NOMOP;

    prtCompta_CV CV2
    mCurrentY_Opération = XPrt.CurrentY

    
    XPrt.CurrentX = 10400
    XPrt.Print "Opération";
    XPrt.CurrentX = 11000
    X = Trim(recCV030.BIACOP)
    If X <> "" Then XPrt.Print ": " & X & "_" & DicLib(27, recCV030.BIACOP);
    
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = 10400
    XPrt.Print "Lot";
    XPrt.CurrentX = 11000
    XPrt.FontBold = True
    XPrt.Print ": " & recCV030.NUMLOT;
    XPrt.FontBold = False
End If


XPrt.CurrentY = prtMinY + prtHeaderHeight + (prtlineHeight - XPrt.TextHeight(X)) / 2
XPrt.CurrentX = prtMinX + 300

XPrt.FontSize = 8
XPrt.FontBold = True
XPrt.CurrentX = 400
XPrt.Print "Compte";

XPrt.CurrentX = 1600
XPrt.Print "Intitulé";

XPrt.CurrentX = 8000
XPrt.Print "Débit";

XPrt.CurrentX = 9600
XPrt.Print "Crédit";

XPrt.CurrentX = Col4 - 1800: XPrt.Print "Date Opé";

XPrt.CurrentX = Col4 - 950: XPrt.Print "Date Valeur";

XPrt.CurrentX = 10400
XPrt.Print "Libellé";

XPrt.CurrentX = prtMaxX - 2000: XPrt.Print "Service";
XPrt.CurrentX = prtMaxX - 1300: XPrt.Print "Référence";
XPrt.CurrentX = prtMaxX - 500: XPrt.Print "Pièce";

XPrt.CurrentX = 1100

XPrt.CurrentY = XPrt.CurrentY - 10


End Sub

'---------------------------------------------------------
Public Sub prtMt(Mt As Currency)
'---------------------------------------------------------
Dim X As String

X = Format$(Abs(Mt), "## ### ### ### ### ##0.00")
XPrt.CurrentX = IIf(Mt < 0, Col5, Col6) - 100 - XPrt.TextWidth(X)
XPrt.Print X;

End Sub


'---------------------------------------------------------
Public Sub prtCompta_CV(CV As typeCV)
'---------------------------------------------------------
Dim strAv As String

XPrt.CurrentX = prtMinX
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.FontBold = False
XPrt.FontSize = 6

If Not blnCVàImprimer Then Exit Sub

XPrt.Print Format$(CV.DeviseN, "000") & "   " & CV.DeviseLibellé;

XPrt.CurrentX = 1800
strAv = ""
If CV.Normal <> " " Then
    Select Case CV.AchatVente
        Case "A": XPrt.Print "(Achat) ";: strAv = "  (Vente)"
        Case "V": XPrt.Print "(Vente) ";: strAv = "  (Achat)"
    End Select
End If

XPrt.FontBold = False
XPrt.CurrentX = 2400
If CV.CotationCertain Then
    XPrt.Print CV3.DeviseIso & "  /  " & CV.DeviseIso;
Else
    XPrt.Print CV.DeviseIso & "  /  " & CV3.DeviseIso;
End If
XPrt.FontBold = True
XPrt.CurrentX = 3100
XPrt.Print strAv;
XPrt.FontBold = False

X = Format$(CV.Cours, "## ##0.00 000 00")
XPrt.CurrentX = 4800 - XPrt.TextWidth(X)
XPrt.Print X;
       
XPrt.FontBold = True
XPrt.CurrentX = 4900

Select Case CV.Normal
    Case "N": XPrt.Print "cours Normal ";
    Case "P": XPrt.Print "cours Privilégié ";
    Case "C": XPrt.Print "cours en Compte ";
    Case Else: XPrt.Print "Cours Pivot ";
End Select

XPrt.CurrentX = Col4 - 1000
XPrt.FontBold = False
XPrt.Print dateImp(CV.CoursAmj);

End Sub



'---------------------------------------------------------
Public Sub prtCV030_Line()
'---------------------------------------------------------

Dim X As String

XPrt.FontSize = 6
XPrt.CurrentY = XPrt.CurrentY + Height8_6

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinX

XPrt.FontBold = True
XPrt.Print Format$(recCompte.Devise, "000") & "  ";

XPrt.FontSize = 8
XPrt.CurrentY = XPrt.CurrentY - Height8_6
XPrt.Print Compte_Imp(recCompte.Numéro);

prtMt (recCV030.MONDEV)
XPrt.FontBold = False

XPrt.FontSize = 6
XPrt.CurrentY = XPrt.CurrentY + Height8_6

If recCV030.AMJSAI <> recCV030.AMJOPE Then XPrt.CurrentX = Col4 - 1800: XPrt.Print dateImp(recCV030.AMJOPE);

XPrt.CurrentX = Col4 - 800
XPrt.Print dateImp(recCV030.AMJVAL);


XPrt.CurrentX = 1700
XPrt.Print Trim(recCompte.Intitulé);
 '======== Demande de compta Guichet : imprimer solde provisoire
If blnSoldeProvisoire_Print Then
    XPrt.Print "   ";
    XPrt.FontItalic = True
    XPrt.FontUnderline = True
    If recCompte.SoldeInstantané < 0 Then
        XPrt.Print "DB ";
    Else
        XPrt.Print "CR ";
    End If
   XPrt.Print Trim(Format$(recCompte.SoldeInstantané, "##### ### ### ### ##0.00"));
    XPrt.FontItalic = False
    XPrt.FontUnderline = False

End If

XPrt.CurrentX = Col6 + 100
XPrt.Print recCV030.LIBELE;

If blnRéférence Then
    If G_CV2.Montant <> 0 Then
        X = Format$(Abs(G_CV2.Montant), "##### ### ### ### ##0.00")
        XPrt.CurrentX = 14800 - XPrt.TextWidth(X)
        XPrt.Print X & " " & G_CV2.DeviseIso;
    End If
    X = Format$(mRéférence, "#### ### ###")
    XPrt.CurrentX = prtMaxX - XPrt.TextWidth(X): XPrt.Print X;
'    If G_CV3.Montant <> 0 Then
'        X = Format$(Abs(G_CV3.Montant), "##### ### ### ### ##0.00")
'        XPrt.CurrentX = 16000 - XPrt.TextWidth(X)
'        XPrt.Print X;
'    End If

'µJPL_20000905 If blnJournal Then
Else
    XPrt.CurrentX = prtMaxX - 1800: XPrt.Print recCV030.SERVIC;
    XPrt.CurrentX = prtMaxX - 1300: XPrt.Print recCV030.NOMOP;
    XPrt.CurrentX = prtMaxX - 500: XPrt.Print Format$(recCV030.NUMPIE, "0000-") & Format$(recCV030.NOLIGN, "000");
'End If
End If

XPrt.CurrentY = XPrt.CurrentY - Height8_6


If Trim(recCompte.Intitulé2) <> "" Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight + Height8_6
    XPrt.CurrentX = 2700
    XPrt.Print recCompte.Intitulé2;
    XPrt.CurrentY = XPrt.CurrentY - Height8_6
End If

If recCV030.CODFOR <> "0" And recCV030.CODFOR <> " " Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = 1100
    XPrt.Print "Compte bloqué";
End If

If recCV030.FOROPO <> "0" And recCV030.FOROPO <> " " Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = 1100
    XPrt.Print "Solde débiteur";
End If

If recCV030.FORVAL <> "0" And recCV030.FORVAL <> " " Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = 1100
    XPrt.Print "Date valeur";
End If

End Sub








Public Sub prtCompta_Total()
Dim X As String

XPrt.FontBold = False
XPrt.FontSize = 6

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight + 50
XPrt.DrawWidth = 1
Call frmElpPrt.prtTrame(Col4, XPrt.CurrentY - 40, Col6, XPrt.CurrentY + prtlineHeight, " ", 250)
'Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY - 40, prtMaxX, XPrt.CurrentY + prtlineHeight, "B")
XPrt.CurrentY = XPrt.CurrentY + Height8_6
If curCumulDébit <> 0 Then Call prtMt(curCumulDébit)
Call prtMt(curCumulCrédit)

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.DrawWidth = 1
If blnTotalSolde And solde <> 0 Then
    Call frmElpPrt.prtTrame(Col4, XPrt.CurrentY, Col6, XPrt.CurrentY + prtlineHeight, "B", 250)
    Call prtMt(solde)
    XPrt.CurrentX = Col6 + 100
    XPrt.Print "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!";
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
End If
XPrt.CurrentY = XPrt.CurrentY - Height8_6
XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY)
XPrt.CurrentY = XPrt.CurrentY + 20
XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY)
XPrt.CurrentY = XPrt.CurrentY - prtlineHeight

curCumulDébit = 0: curCumulCrédit = 0: solde = 0
nbLigne = 0: NbImprimé = 0
End Sub

Public Sub prtCompta_Trait()
XPrt.DrawWidth = 1
XPrt.Line (Col4, prtMinY)-(Col4, prtMaxY)
XPrt.Line (Col6, prtMinY)-(Col6, prtMaxY)

End Sub

Public Sub prtCompta_Devise()
XPrt.FontSize = 9
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.DrawWidth = 1
Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY - 50, prtMaxX, XPrt.CurrentY + prtlineHeight - 50, " ")
DevX mDevise
XPrt.FontBold = True
XPrt.CurrentX = prtMinX
XPrt.Print Trim(XDevise.DevLib);

''frmElpPrt.prtCentré (Col6 + Col4) / 2, Trim(XDevise.DevLib)

End Sub

Public Sub prtCompta_Line()
If XPrt.CurrentY > prtMaxY_4 Then
     prtCompta_Trait
     frmElpPrt.prtNewPage
     
     If blnJournal Then prtCompta_Form
 End If
 
 nbLigne = nbLigne + 1
 
 If mDevise <> recCV030.Devise Then
     prtCompta_Total
     mDevise = recCV030.Devise
     prtCompta_Devise
 End If
 
 If recCV030.SENECR = "D" Then recCV030.MONDEV = -recCV030.MONDEV
 
 recCompteInit recCompte
 recCompte.Method = "SeekL1"
 recCompte.Société = recCV030.COSOC
 recCompte.Agence = recCV030.Agence
 recCompte.Devise = Format$(Val(recCV030.Devise), "000")
 recCompte.Numéro = recCV030.Compte

If blnSoldeProvisoire_Print Then
    recCompte.Method = "SeekL1      "
    recCompte.BiaTyp = "000"
    recCompte.BiaNum = "00000"
    recCompte.NuméroAncien = "00000000000"
    If Not IsNull(srvCompteMon(recCompte)) Then recCompte.Intitulé = "??????"
Else
    mdbCptP0_Find recCompte
End If

 prtCV030_Line
 solde = solde + recCV030.MONDEV
 If recCV030.MONDEV < 0 Then
     curCumulDébit = curCumulDébit + recCV030.MONDEV
 Else
   curCumulCrédit = curCumulCrédit + recCV030.MONDEV
 End If
 
' NbImprimé = NbImprimé + 1
' If NbImprimé = 4 Then
'     Call frmElpPrt.prtTrame(prtMinX + 20, XPrt.CurrentY - prtlineHeight - 50, prtMaxX - 20, XPrt.CurrentY + prtlineHeight - 50, " ")
'     NbImprimé = 0
' End If
 
 DoEvents: If prtKillDoc Then frmElpPrt.Hide: Exit Sub

End Sub

Public Sub prtCompta_Validation(Msg As String)
Call frmElpPrt.prtTrame(prtMaxX - 5000, prtMaxY - prtlineHeight * 5, prtMaxX - 10, prtMaxY - 10, " ", "230")
XPrt.CurrentY = prtMaxY - prtlineHeight * 4
XPrt.FontSize = 10
XPrt.FontUnderline = True
frmElpPrt.prtCentré prtMaxX - 2500, Trim(Msg)
XPrt.FontUnderline = False
XPrt.CurrentY = prtMaxY - prtlineHeight * 4
End Sub


Public Sub prtGuichet_Compta_Form()
Dim mCurrenty As Integer

If mComptaUsr = "" Then
    mComptaUsr = Format$(recGuichet_Compta.ComptaUsr, "### ### ###")
    recCV030.NUMLOT = mId$(recGuichet_Compta.ComptaUsr, 7, 4)
End If

prtCompta_Form

mCurrenty = XPrt.CurrentY
XPrt.CurrentY = mCurrentY_Opération
XPrt.FontBold = False
XPrt.FontSize = 6
XPrt.CurrentX = 12000: XPrt.Print "saisi par :";
XPrt.FontBold = True
XPrt.CurrentX = 11000: XPrt.Print mJournal;
XPrt.CurrentX = 12700: XPrt.Print recGuichet_Compta.SaisieUsr;
XPrt.FontBold = False
XPrt.CurrentX = 14000: XPrt.Print dateImp(recGuichet_Compta.ValidationAMJ);
XPrt.CurrentX = 15100: XPrt.Print timeImp(recGuichet_Compta.ValidationHMS);


XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.FontBold = True
''XPrt.CurrentX = 11000: XPrt.Print mComptaUsr;
XPrt.CurrentX = 12700: XPrt.Print recGuichet_Compta.ValidationUsr;
XPrt.FontBold = False
XPrt.CurrentX = 12000: XPrt.Print "validé par :";
XPrt.CurrentX = 14000: XPrt.Print dateImp(recGuichet_Compta.ComptaAMJ);
XPrt.CurrentX = 15100: XPrt.Print timeImp(recGuichet_Compta.ComptaHMS);

XPrt.FontSize = 8
XPrt.CurrentY = mCurrenty
End Sub

Public Sub prtGuichet_Comptabilisé(Msg As String)
Dim iReturn As Integer
Dim G_ElpBuffer As typeElpBuffer

mCptMvtPièce = -1
blnForm = True
curCumulDébit = 0
curCumulCrédit = 0
G_CV2.Montant = 0
recCpj030W0_Init recCV030

G_ElpBuffer.Method = "Seek=" '
G_ElpBuffer.Id = Msg
G_ElpBuffer.Seq = 1

Do
    iReturn = tableElpBuffer_Read(G_ElpBuffer)
    If iReturn = 0 Then
        MsgTxt = G_ElpBuffer.Data
        prtCV030_CptMvt
    End If
    G_ElpBuffer.Seq = G_ElpBuffer.Seq + 1
Loop Until iReturn <> 0

prtCompta_Total
prtCompta_Trait

End Sub
Public Sub prtMVTP0(Msg As String)
Dim iReturn As Integer, blnNéant As Boolean
Dim recMvtp0 As typeMvtP0
mdbMvtP0.tableMvtP0_Open

blnNéant = True
mCptMvtPièce = -1
blnForm = True
curCumulDébit = 0
curCumulCrédit = 0
G_CV2.Montant = 0
recCpj030W0_Init recCV030

recMvtp0.Method = "MoveFirst" '

Do
    iReturn = tableMvtP0_Read(recMvtp0)
    If iReturn = 0 Then
        MsgTxt = Space$(recCptMvtLen)
        Mid$(MsgTxt, 35, memoCptMvtLen) = mId$(recMvtp0.Text, 1, memoCptMvtLen)
        
        If blnEAR_Imp Then
            If mId$(recMvtp0.Text, 10, 8) = "00038890" Then
                prtCV030_CptMvt
                blnNéant = False
            End If
        Else
            If mId$(recMvtp0.Text, 10, 8) = "00038890" Then
                blnEAR = True
            End If
            prtCV030_CptMvt
            blnNéant = False
        End If
    End If
    recMvtp0.Method = "MoveNext"
Loop Until iReturn <> 0

If blnNéant Then
    XPrt.FontBold = False
    XPrt.FontSize = 32
    XPrt.CurrentY = (prtMinY + prtMaxY) / 2
    frmElpPrt.prtCentré (prtMinX + prtMaxX) / 2, "NEANT"
Else
    prtCompta_Total
    prtCompta_Trait
End If

If blnEAR_Imp Then
    XPrt.FontBold = False
    XPrt.FontSize = 32
    XPrt.CurrentY = prtMaxY - 3 * prtlineHeight
    XPrt.CurrentX = prtMaxX - 3000: XPrt.Print "E A R";
End If

mdbMvtP0.tableMvtP0_Close

End Sub


Public Sub prtCV030_CptMvt()
MsgTxtIndex = 0
srvCptMvtGetBuffer recCptMvt

recCV030.COSOC = recCptMvt.Société
recCV030.Agence = recCptMvt.Agence
recCV030.Devise = recCptMvt.Devise
recCV030.NUMLOT = recCptMvt.Lot
recCV030.NUMPIE = recCptMvt.Pièce Mod 10000
recCV030.NOLIGN = recCptMvt.Ligne Mod 1000
recCV030.NOMOP = recCptMvt.OpérateurSaisie
recCV030.SERVIC = recCptMvt.Service
recCV030.Compte = recCptMvt.Compte
recCV030.MONDEV = recCptMvt.Mt
recCV030.AMJOPE = recCptMvt.AmjOpération
recCV030.AMJVAL = recCptMvt.AmjValeur
recCV030.LIBELE = recCptMvt.Libellé
recCV030.AMJSAI = recCptMvt.AmjTraitement 'AmjSaisie
'recCV030.CODFOR = recCptMvt.chkCompte
'recCV030.FOROPO = recCptMvt.chkSolde
'recCV030.OPOCHQ = recCptMvt.chkChèque
'recCV030.FORVAL = recCptMvt.chkAmjValeur
If blnCptMvtPièce_Rupture Then
    If mCptMvtPièce <> recCptMvt.Pièce Then
        XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
        XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY)
        XPrt.CurrentY = XPrt.CurrentY + 20 - prtlineHeight
        mCptMvtPièce = recCptMvt.Pièce
    End If
End If

If blnForm Then
    blnForm = False
    prtCompta_Form
    mDevise = recCV030.Devise
    prtCompta_Devise
End If
'    mRéférence = recCptMvt.Référence
prtCompta_Line

End Sub
