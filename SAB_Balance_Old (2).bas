Attribute VB_Name = "prtSAB_Balance"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Type typeSAB_BALANCE
    Obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    CLIENARSD       As String * 3                     ' Pays de résidence
    COMPTEDEV       As String * 3                     ' TABLES BASE 013
    COMPTEOBL       As String * 10                    ' COMPTE OBLIGATOIRE
    COMPTECOM       As String * 20                    ' NUMERO COMPTE
    COMPTEINT       As String * 32                    ' INTITULE
    DB              As Currency                       '
    CR              As Currency                       '
    DB_EUR          As Currency                       '
    CR_EUR          As Currency
    blnPrint        As Boolean          '
    blnPrint_Fontbold        As Boolean          '
    blnPrint_Line        As Boolean          '
    iPrint_Trame        As Integer          '
End Type

Dim arrSAB_BALANCE() As typeSAB_BALANCE
Dim arrSAB_BALANCE_Nb As Long, arrSAB_BALANCE_NbMax As Long
Dim zSAB_BALANCE As typeSAB_BALANCE
Dim totalSAB_BALANCE(10) As typeSAB_BALANCE
Dim prevSAB_BALANCE As typeSAB_BALANCE
Dim détailSAB_BALANCE As typeSAB_BALANCE

Dim arrDEV_C1A5(100) As typeSAB_BALANCE
Dim arrDEV_C6A8(100) As typeSAB_BALANCE
Dim arrDEV_C9(100) As typeSAB_BALANCE
Dim arrDev_Nb As Long, arrDev_K As Integer


Dim X As String, I As Integer, Height8_6 As Integer

Dim blnPage As Boolean

Dim xYbase As typeYBase
Dim curX As Currency, curCumul_Db As Currency, curCumul_Cr As Currency
Dim curClient_Db As Currency, curClient_Cr As Currency, nbClient_Line As Long
Dim curListe_Db As Currency, curListe_Cr As Currency, nbListe_Line As Long
Dim curW_Db As Currency, curW_Cr As Currency
Dim IbmAmjMin As String, IbmAmjMax As String
Dim meYBIACPT0 As typeYBIACPT0, prevYBIACPT0 As typeYBIACPT0

Dim blnCompte As Boolean
Dim prtY As Integer
Dim mBalance_CV As String '* 3
Dim meCV1 As typeCV, meCV2 As typeCV

Dim blnMOUVEMDCO As Boolean, blnRésidence As Boolean
Dim mRésidence As String
Dim blnSoldeZ As Boolean
Dim blnClient_Line As Boolean

Dim blnBalance_B_COMPTEDEV As Boolean, blnBalance_B_Détail As Boolean, blnBalance_B_Récap As Boolean
Dim curCOMPTEDEV_Db As Currency, curCOMPTEDEV_Cr As Currency, nbCOMPTEDEV_Line As Long
Dim curCOMPTEDEV_Db_EUR As Currency, curCOMPTEDEV_Cr_EUR As Currency
Dim curCOMPTEOBL_Db As Currency, curCOMPTEOBL_Cr As Currency, nbCOMPTEOBL_Line As Long
Dim curCOMPTEOBL_Db_EUR As Currency, curCOMPTEOBL_Cr_EUR As Currency
Dim mBalance_YSOLDE0 As Integer
Dim blnBalance_Compte_Soldé As Boolean
Dim blnBalance_Pays As Boolean, xPays As String
Dim blnBalance_Récap_Bilan As Boolean

Dim meYPLAN0 As typeYPLAN0
Dim xYSOLDE0 As typeYSOLDE0

Dim idFile_CSV As Integer, blnFile_CSV As Boolean

Type typeSAB_Client_Stat
    Nb_Client           As Long
    Nb_Compte           As Long
    Nb_Compte_Annulé    As Long
    Solde_DB           As Currency
    Solde_CR           As Currency
End Type

Dim SAB_Client_Stat As typeSAB_Client_Stat
Dim SAB_Client_Stat_Actif As typeSAB_Client_Stat
Dim SAB_Client_Stat_Annulé As typeSAB_Client_Stat
Dim SAB_Client_Stat_Produit As typeSAB_Client_Stat
Dim SAB_Client_Stat_Produit_Lib As String

Dim meYBIATAB0 As typeYBIATAB0
Dim xTitleText As String
Dim mMsg As String
Dim blnChk_BalanceEquilibrée As Boolean, blnChk_BalanceStock As Boolean
Dim blnChk_COMPTEOUV As Boolean

Dim meYSTOMON As Currency, meDORCPTDMV As Long

Dim wAMJ_6M_00 As Long, wAMJ_6M_99 As Long
Public Sub prtSAB_Balance_Monitor(lFct As String, lAmjMin As String, lAmjMax As String, fgW As MSFlexGrid, larrYBIACPT0() As typeYBIACPT0, larrYBIACPT0_Nb As Long, lMsg As String, lYSTOMON() As Currency, lDORCPTDMV() As Long)

' B : balance 1,1 : "B"
'             2,1 : "D" (Rupture Devise/ PCEC)              blnBalance_B_COMPTEDEV
'                 : "P" (Rupture PCEC/Devise)
'             3,1 : "D" (imprimer lignes détail des compes  blnBalance_Détail
'             4,1 : "V" (solde Veille)
'                 : "M" (solde fin de mois)
'                 : "2" (solde fin de mois -2)
'                 : "A" (solde fin d'année)
'             5,1 : "1" Recap
'             6,1 : "1" CSV
'             7,3 : N0 Fichier
'            10,1 : "1" Ignorer les comptes soldés
'            11,1 : "1" Balance par pays
'            12,1 : "1" Recap Bilan /Hors-Bilan

'            16,6 : "1"Recap niveau 0 (DEV),Gras,Souliné,Trame
'            22,6 : "1" Recap niveau 1
'            28,6 : "1" Recap niveau 2
'            34,6 : "1" Recap niveau 3
'            40,6 : "1" Recap niveau 4
'            46,6 : "1" Recap niveau 5
'            52,6 : "1" Recap niveau 6
'            58,6 : "1"  total de la balance Détail niveau 7

Dim wIndex As Long, I As Integer
Dim mFct1 As String

blnChk_BalanceEquilibrée = True
blnChk_BalanceStock = False
blnChk_COMPTEOUV = False

mFct1 = mId$(lFct, 1, 1)
If mFct1 = "S" Then
    mFct1 = "B"
    blnChk_BalanceEquilibrée = False
    blnChk_BalanceStock = True
    
End If

mMsg = lMsg
IbmAmjMin = dateIBM(lAmjMin)
IbmAmjMax = dateIBM(lAmjMax)

wAMJ_6M_00 = Fix((dateElp("MoisAdd", -6, YBIATAB0_DATE_CPT_J) - 19000000) / 100) * 100
wAMJ_6M_99 = wAMJ_6M_00 + 99

meCV1.DeviseN = 0
meCV1.Montant = 0

mBalance_CV = "J  "
meCV1.OpéAmj = YBIATAB0_DATE_CPT_J
meCV2.OpéAmj = YBIATAB0_DATE_CPT_J
If lAmjMax = YBIATAB0_DATE_CPT_MP1 Then
        mBalance_CV = "MP1"
        meCV1.OpéAmj = YBIATAB0_DATE_CPT_MP1
        meCV2.OpéAmj = YBIATAB0_DATE_CPT_MP1
End If
If lAmjMax = YBIATAB0_DATE_CPT_AP1 Then
        mBalance_CV = "AP1"
        meCV1.OpéAmj = YBIATAB0_DATE_CPT_AP1
        meCV2.OpéAmj = YBIATAB0_DATE_CPT_AP1
End If



blnMOUVEMDCO = False
blnRésidence = False: mRésidence = "-"
blnCompte = False:    curCumul_Db = 0: curCumul_Cr = 0
blnClient_Line = False
nbClient_Line = 0: nbListe_Line = 0
curClient_Db = 0: curClient_Cr = 0
curListe_Db = 0: curListe_Cr = 0
curCOMPTEDEV_Db = 0: curCOMPTEDEV_Cr = 0
curCOMPTEDEV_Db_EUR = 0: curCOMPTEDEV_Cr_EUR = 0
curCOMPTEOBL_Db = 0: curCOMPTEOBL_Cr = 0
curCOMPTEOBL_Db_EUR = 0: curCOMPTEOBL_Cr_EUR = 0
curW_Db = 0: curW_Cr = 0
recYBIACPT0_Init prevYBIACPT0

prtSAB_Balance_Total_Init
blnBalance_B_Détail = False
blnBalance_B_Récap = False
blnBalance_B_COMPTEDEV = False
blnBalance_Compte_Soldé = False
blnBalance_Pays = False
blnBalance_Récap_Bilan = False
nbCOMPTEOBL_Line = 0: nbCOMPTEDEV_Line = 0

mBalance_YSOLDE0 = 0
idFile_CSV = 0
blnFile_CSV = False

Select Case mFct1
    Case "L": prtTitleText = "Liste au " & dateImp10(lAmjMax)
                If mId$(lFct, 2, 1) = "T" Then blnClient_Line = True
                
    Case "B":
                If mId$(lFct, 2, 1) = "D" Then blnBalance_B_COMPTEDEV = True
                If mId$(lFct, 3, 1) = "1" Then blnBalance_B_Détail = True
                If mId$(lFct, 5, 1) = "1" Then blnBalance_B_Récap = True
                If mId$(lFct, 10, 1) = "1" Then blnBalance_Compte_Soldé = True
                If mId$(lFct, 12, 1) = "1" Then blnBalance_Récap_Bilan = True
                
                Select Case mId$(lFct, 4, 1)
                    Case "M":   mBalance_YSOLDE0 = 1
                                mBalance_CV = "MP1"
                                meCV1.OpéAmj = YBIATAB0_DATE_CPT_MP1
                                meCV2.OpéAmj = YBIATAB0_DATE_CPT_MP1
                    Case "2":   mBalance_YSOLDE0 = 2
                                mBalance_CV = YBIATAB0_DATE_CPT_MP2
                                meCV1.OpéAmj = YBIATAB0_DATE_CPT_MP2
                                meCV2.OpéAmj = YBIATAB0_DATE_CPT_MP2
                    Case "A":   mBalance_YSOLDE0 = 1 + Val(mId$(YBIATAB0_DATE_CPT_MP1, 5, 2))
                                mBalance_CV = "AP1"
                                meCV1.OpéAmj = YBIATAB0_DATE_CPT_AP1
                                meCV2.OpéAmj = YBIATAB0_DATE_CPT_AP1
                End Select
                
                If mId$(lFct, 6, 1) = "1" Then blnFile_CSV = True
                idFile_CSV = Val(mId$(lFct, 7, 3))

                prtSAB_Balance_B_Total_Init mId$(lFct, 16, 6), totalSAB_BALANCE(0)
                prtSAB_Balance_B_Total_Init mId$(lFct, 22, 6), totalSAB_BALANCE(1)
                prtSAB_Balance_B_Total_Init mId$(lFct, 28, 6), totalSAB_BALANCE(2)
                prtSAB_Balance_B_Total_Init mId$(lFct, 34, 6), totalSAB_BALANCE(3)
                prtSAB_Balance_B_Total_Init mId$(lFct, 40, 6), totalSAB_BALANCE(4)
                prtSAB_Balance_B_Total_Init mId$(lFct, 46, 6), totalSAB_BALANCE(5)
                prtSAB_Balance_B_Total_Init mId$(lFct, 52, 6), totalSAB_BALANCE(6)
                prtSAB_Balance_B_Total_Init mId$(lFct, 58, 6), détailSAB_BALANCE
                                
                If blnChk_BalanceStock Then
                    prtTitleText = mMsg & " - Balance / Stock au " & dateImp10(meCV1.OpéAmj)
                Else
                     prtTitleText = "Balance au " & dateImp10(meCV1.OpéAmj)
               End If
                
                If mId$(lFct, 11, 1) = "1" Then blnBalance_Pays = True: prtTitleText = mMsg & " Balance par PAYS de résidence au " & dateImp10(meCV1.OpéAmj)
                xTitleText = prtTitleText
    Case "C": prtTitleText = "Cumul de mouvements du " & dateImp10(lAmjMin) & " au " & dateImp10(lAmjMax)
               If mId$(lFct, 2, 3) = "DCO" Then blnMOUVEMDCO = True
               If mId$(lFct, 5, 1) = "R" Then blnRésidence = True
    Case "R": prtTitleText = "Relevé du " & dateImp10(lAmjMin) & " au " & dateImp10(lAmjMax)

End Select

If mId$(lFct, 6, 1) = "Z" Then
    blnSoldeZ = True
Else
    blnSoldeZ = False
End If

prtFontName = prtFontName_Arial
prtSAB_Balance_Open
prtHeaderHeight = 300
prtSAB_Balance_Form
XPrt.CurrentY = XPrt.CurrentY - prtlineHeight

XPrt.FontSize = 8
For I = 1 To fgW.Rows - 1
    
    fgW.Row = I
    fgW.Col = fgW.Cols - 1: wIndex = Val(fgW.Text)
    meYBIACPT0 = larrYBIACPT0(wIndex)
    meCV1.DeviseIso = meYBIACPT0.COMPTEDEV
    If blnChk_BalanceStock Then meYSTOMON = lYSTOMON(wIndex): meDORCPTDMV = lDORCPTDMV(wIndex)
    Select Case mFct1
        Case "B": prtSAB_Balance_B_Line
        Case "R": prtSAB_Balance_R
        Case "C": prtSAB_Balance_C
        Case Else: meDORCPTDMV = lDORCPTDMV(wIndex): prtSAB_Balance_L_Line
     End Select
     
     prevYBIACPT0 = meYBIACPT0
Next I

If mFct1 = "C" And blnRésidence Then prtSAB_Balance_C_Cumul


If mFct1 = "B" Then
    prtSAB_Balance_B_Fin

Else
    If mFct1 = "L" Then
        
        prtSAB_Balance_L_Rupture
        If curListe_Db <> curW_Db Or curListe_Cr <> curW_Cr Then
            prtSAB_Balance_NewLine
            XPrt.FontSize = 12: XPrt.FontBold = True
            frmElpPrt.prtCentré prtMedX, "ERREUR TOTALISATION"
            XPrt.FontSize = 8: XPrt.FontBold = False
        End If
        XPrt.DrawWidth = 10
        XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
        XPrt.Line (prtMinX + 12000, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY)
        XPrt.CurrentY = XPrt.CurrentY - prtlineHeight
        prevYBIACPT0.CLIENACLI = ""
        prevYBIACPT0.CLIENASIG = ""
        nbClient_Line = nbListe_Line
        curClient_Db = curListe_Db
        curClient_Cr = curListe_Cr
        prtSAB_Balance_L_Rupture
    Else
    
        XPrt.DrawWidth = 10
        XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
        XPrt.Line (prtMinX + 7500, XPrt.CurrentY)-(prtMinX + 12000, XPrt.CurrentY)
    End If
End If

prtSAB_Balance_Close

End Sub
Public Sub prtSAB_Client_Stat(lFct As String, lAmjMin As String, lAmjMax As String, fgW As MSFlexGrid, larrYBIACPT0() As typeYBIACPT0, larrYBIACPT0_Nb As Long)
Dim wIndex As Long, I As Integer
Dim blnCumul As Boolean

prtTitleText = "Répartition de la clientèle par catégorie"
meCV1.DeviseN = 0
meCV1.Montant = 0

mBalance_CV = "J  "
meCV1.OpéAmj = YBIATAB0_DATE_CPT_J
meCV2.OpéAmj = YBIATAB0_DATE_CPT_J

prtFontName = prtFontName_Arial
prtSAB_Balance_Open
prtHeaderHeight = 300
XPrt.CurrentY = XPrt.CurrentY - prtlineHeight
prtSAB_Client_Stat_Form

prtSAB_Client_Stat_Z SAB_Client_Stat
prtSAB_Client_Stat_Z SAB_Client_Stat_Actif
prtSAB_Client_Stat_Z SAB_Client_Stat_Annulé
prtSAB_Client_Stat_Z SAB_Client_Stat_Produit

recYBIACPT0_Init prevYBIACPT0
blnCumul = False

fgW.Row = 1
fgW.Col = fgW.Cols - 1: wIndex = Val(fgW.Text)
prevYBIACPT0 = larrYBIACPT0(wIndex)
prtSAB_Client_Stat_PLANCOPRO prevYBIACPT0.PLANCOPRO

XPrt.FontSize = 8
For I = 1 To fgW.Rows - 1
    
    fgW.Row = I
    fgW.Col = fgW.Cols - 1: wIndex = Val(fgW.Text)
    meYBIACPT0 = larrYBIACPT0(wIndex)
    
    If meYBIACPT0.PLANCOPRO <> prevYBIACPT0.PLANCOPRO Then
        prtSAB_Client_Stat_Cumul
        prtSAB_Client_Stat_Line
        prtSAB_Client_Stat_Produit
        prtSAB_Client_Stat_PLANCOPRO meYBIACPT0.PLANCOPRO

    Else
         If meYBIACPT0.CLIENACAT <> prevYBIACPT0.CLIENACAT Then
             prtSAB_Client_Stat_Cumul
             prtSAB_Client_Stat_Line
         Else
             If meYBIACPT0.CLIENACLI <> prevYBIACPT0.CLIENACLI Then
                 prtSAB_Client_Stat_Cumul
             End If
        End If
    End If
    
    If meYBIACPT0.COMPTEFON = "4" Then
        SAB_Client_Stat.Nb_Compte_Annulé = SAB_Client_Stat.Nb_Compte_Annulé + 1
    Else
        SAB_Client_Stat.Nb_Compte = SAB_Client_Stat.Nb_Compte + 1
    End If
    meCV1.DeviseIso = meYBIACPT0.COMPTEDEV
    prtSAB_Balance_CV meYBIACPT0.SOLDECEN
    If meCV2.Montant > 0 Then
        SAB_Client_Stat.Solde_DB = SAB_Client_Stat.Solde_DB + meCV2.Montant
    Else
        SAB_Client_Stat.Solde_CR = SAB_Client_Stat.Solde_CR - meCV2.Montant
    End If
    prevYBIACPT0 = meYBIACPT0
Next I

prtSAB_Client_Stat_Cumul
prtSAB_Client_Stat_Line
prtSAB_Client_Stat_Produit

prtSAB_Client_Stat_Form_End

prtSAB_Balance_Close

End Sub


'---------------------------------------------------------
Public Sub prtSAB_Balance_Form()
'---------------------------------------------------------
Dim X As String

XPrt.DrawWidth = 1
XPrt.FontSize = 7: XPrt.FontBold = True

Call frmElpPrt.prtTrame(prtMinX, prtMinY, prtMaxX, prtMinY + prtHeaderHeight, "B", 240)
XPrt.Line (prtMinX + 12000, prtMinY)-(prtMinX + 12000, prtMaxY)
XPrt.Line (prtMinX + 7500, prtMinY)-(prtMinX + 7500, prtMaxY)
XPrt.Line (prtMaxX, prtMinY)-(prtMaxX, prtMaxY)
'---------------------------------------------------------
XPrt.CurrentY = prtMinY + 50

XPrt.CurrentX = prtMinX + 100: XPrt.Print " ";
XPrt.CurrentX = prtMinX + 400: XPrt.Print "Compte ";
XPrt.CurrentX = prtMinX + 2000: XPrt.Print "Intitulé";
If blnChk_BalanceStock Then
    XPrt.CurrentX = prtMinX + 5700: XPrt.Print "Stock opé";
    XPrt.CurrentX = prtMinX + 6500: XPrt.Print "Der Mvt";
    XPrt.CurrentX = prtMinX + 12100: XPrt.Print "Contrôle";
End If

'XPrt.CurrentX = prtMinX + 10500: XPrt.Print "Devise";
XPrt.CurrentX = prtMinX + 9000: XPrt.Print "Débit";
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


Public Sub prtSAB_Balance_Close()
On Error GoTo prtError

If blnChk_BalanceStock Then
    XPrt.FontBold = True
    XPrt.FontSize = 6
    
    prtSAB_Balance_NewLine
    XPrt.CurrentX = prtMinX + 12100: XPrt.Print "=";
    XPrt.CurrentX = prtMinX + 12400: XPrt.Print "Solde balance = cumul des contrats";
    prtSAB_Balance_NewLine
    XPrt.CurrentX = prtMinX + 12100: XPrt.Print "##";
    XPrt.CurrentX = prtMinX + 12400: XPrt.Print "Solde balance <> cumul des contrats ";
    prtSAB_Balance_NewLine
    XPrt.CurrentX = prtMinX + 12100: XPrt.Print "??";
    XPrt.CurrentX = prtMinX + 12400: XPrt.Print "Solde balance, aucun contrat rattaché??";
    prtSAB_Balance_NewLine
    XPrt.CurrentX = prtMinX + 2000: XPrt.Print "CAV et LOR : date de dernier mouvement hors échelles, facturation ...(+) indique qu'il y a des mouvements postérieurs initiés par la banque";
    XPrt.CurrentX = prtMinX + 12100: XPrt.Print "6M";
    XPrt.CurrentX = prtMinX + 12400: XPrt.Print "Compte ouvert depuis 6 mois ";
    XPrt.DrawWidth = 1: XPrt.Line (prtMinX + 12300, prtMinY + prtHeaderHeight)-(prtMinX + 12300, prtMaxY)
End If

frmElpPrt.prtEndDoc
frmElpPrt.Hide
Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub



Public Sub prtSAB_Balance_Open()
On Error GoTo prtError

Set XPrt = Printer
frmElpPrt.Show vbModeless

Height8_6 = frmElpPrt.prtHeightDelta(8, 6)

blnFiligrane = False
prtOrientation = vbPRORLandscape '
prtPgmName = "prtSAB_Balance"
prtTitleUsr = usrName

prtLineNb = 1
prtlineHeight = 250
prtHeaderHeight = 50 ' 100


prtFormType = ""
frmElpPrt.prtStdInit

Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub





Public Sub prtSAB_Balance_R()
Dim xYBIAMVT0 As typeYBIAMVT0, wSolde As Currency

wSolde = meYBIACPT0.SOLDECEN
blnCompte = False
recYBIAMVT0_Init xYBIAMVT0

xYbase.ID = constYBIAMVT0
xYbase.K1 = meYBIACPT0.COMPTECOM & IbmAmjMin
xYbase.Method = "Seek>"

Do
    intReturn = tableYBase_Read(xYbase)
    If Trim(xYbase.ID) <> constYBIAMVT0 Then intReturn = -1
    If intReturn = 0 Then
        MsgTxt = Space$(34) & xYbase.Text
        MsgTxtIndex = 0
        srvYBIAMVT0_GetBuffer xYBIAMVT0
        If xYBIAMVT0.MOUVEMCOM = meYBIACPT0.COMPTECOM And xYBIAMVT0.MOUVEMDTR <= IbmAmjMax Then
            If Not blnCompte Then
                blnCompte = True
                wSolde = xYBIAMVT0.BIAMVTSD0
                XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 0.5
                prtSAB_Balance_L "R0", wSolde
            End If
            
            prtSAB_Balance_Mvt xYBIAMVT0
            wSolde = wSolde + xYBIAMVT0.MOUVEMMON
        Else
            intReturn = -1
        End If
        
    End If

Loop Until intReturn <> 0




If Not blnCompte Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 0.5
    prtSAB_Balance_L "R1", wSolde
Else
    prtSAB_Balance_NewLine
    prtY = XPrt.CurrentY + prtlineHeight
    Call frmElpPrt.prtTrame(prtMinX + 7520, XPrt.CurrentY, prtMinX + 11980, prtY - 50, " ", 240)
   ' Call frmElpPrt.prtTrame(prtMinX + 12020, XPrt.CurrentY, prtMaxX - 20, prtY - 50, " ", 240)
     XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 6
   XPrt.FontBold = True
    prtSAB_Balance_Montant wSolde
    XPrt.FontBold = False
    XPrt.CurrentX = prtMinX + 11600: XPrt.Print meYBIACPT0.COMPTEDEV;
    XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 8
End If
'XPrt.Line (prtMinX + 7500, prtY)-(prtMinX + 12000, prtY)
'XPrt.CurrentY = XPrt.CurrentY - prtlineHeight

End Sub
Public Sub prtSAB_Balance_C()
Dim xYBIAMVT0 As typeYBIAMVT0, wSolde As Currency
Dim curDB As Currency, curCR As Currency
Dim xDate As String

If blnRésidence Then
    If mId$(meYBIACPT0.COMPTECOM, 10, 1) <> mRésidence Then
        prtSAB_Balance_C_Cumul
        mRésidence = mId$(meYBIACPT0.COMPTECOM, 10, 1)
        blnCompte = False:    curCumul_Db = 0: curCumul_Cr = 0
    End If
End If

curDB = 0: curCR = 0
wSolde = meYBIACPT0.SOLDECEN
recYBIAMVT0_Init xYBIAMVT0

xYbase.ID = constYBIAMVT0
xYbase.K1 = meYBIACPT0.COMPTECOM & IbmAmjMin
xYbase.Method = "Seek>"

Do
    intReturn = tableYBase_Read(xYbase)
    If Trim(xYbase.ID) <> constYBIAMVT0 Then intReturn = -1
    If intReturn = 0 Then
        MsgTxt = Space$(34) & xYbase.Text
        MsgTxtIndex = 0
        srvYBIAMVT0_GetBuffer xYBIAMVT0
        If blnMOUVEMDCO Then
            xDate = xYBIAMVT0.MOUVEMDCO
        Else
            xDate = xYBIAMVT0.MOUVEMDTR
        End If
        
        If xYBIAMVT0.MOUVEMCOM = meYBIACPT0.COMPTECOM Then
            If xDate <= IbmAmjMax And xDate >= IbmAmjMin Then
            
                If xYBIAMVT0.MOUVEMMON < 0 Then
                    curCR = curCR + xYBIAMVT0.MOUVEMMON
                Else
                    curDB = curDB + xYBIAMVT0.MOUVEMMON
                End If
            End If
        Else
            intReturn = -1
        End If
        
    End If

Loop Until intReturn <> 0

If blnSoldeZ Or curDB <> 0 Or curCR <> 0 Then
'''If curDB = 0 And curCR = 0 Then
    blnCompte = True
    curCumul_Db = curCumul_Db + curDB
    curCumul_Cr = curCumul_Cr + curCR
    
    prtSAB_Balance_L "C1", curCR
    If curDB <> 0 Then prtSAB_Balance_Montant curDB
End If

End Sub


Public Sub prtSAB_Balance_L(lFct As String, lcurX As Currency)

prtSAB_Balance_NewLine
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 6

Select Case lFct
    Case "R0"
     '   Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY, prtMinX + 7480, XPrt.CurrentY + prtlineHeight - 50, " ", 240)
    '    Call frmElpPrt.prtTrame(prtMinX + 7520, XPrt.CurrentY, prtMinX + 11980, XPrt.CurrentY- 50, " ", 240)
    '    Call frmElpPrt.prtTrame(prtMinX + 12020, XPrt.CurrentY , prtMaxX - 20, XPrt.CurrentY- 50, " ", 240)
        XPrt.FontBold = True: XPrt.ForeColor = prtForeColor_Header

        prtSAB_Balance_Montant lcurX
        XPrt.FontBold = False: XPrt.ForeColor = prtForeColor
    Case "R1"
        prtY = XPrt.CurrentY + prtlineHeight
       ' Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY, prtMinX + 7480, prtY - 50, " ", 240)
        Call frmElpPrt.prtTrame(prtMinX + 7520, XPrt.CurrentY, prtMinX + 11980, prtY - 50, " ", 240)
       ' Call frmElpPrt.prtTrame(prtMinX + 12020, XPrt.CurrentY, prtMaxX - 20, prtY - 50, " ", 240)
        XPrt.FontBold = True
        prtSAB_Balance_Montant lcurX
        XPrt.FontBold = False
    Case "C1"
        prtSAB_Balance_Montant lcurX
    Case Else
        prtSAB_Balance_Montant lcurX
End Select

XPrt.CurrentX = prtMinX + 100: XPrt.Print meYBIACPT0.PLANCOPRO;
If lFct = "B" Then
    XPrt.FontBold = False
Else
    XPrt.FontBold = True
End If

''Dim mRib_IbanE As String, mRib_Clé As String
''mRib_Clé = Format$(RibClé(strSocBdfE, strSocBdfG, Trim(meYBIACPT0.COMPTECOM), mRib_IbanE), "00")

XPrt.CurrentX = prtMinX + 400: XPrt.Print meYBIACPT0.COMPTECOM;  '''"& "    " & mRib_Clé;

XPrt.CurrentX = prtMinX + 2000: XPrt.Print meYBIACPT0.COMPTEINT;
XPrt.FontBold = False
If lFct = "L" Then
'_____________________________
           If meDORCPTDMV > 0 Then
                XPrt.CurrentX = prtMinX + 6800: XPrt.Print dateIBM10(meDORCPTDMV, True);
                If meDORCPTDMV <> meYBIACPT0.SOLDEDMO Then XPrt.Print " +";
            Else
                XPrt.CurrentX = prtMinX + 6800: XPrt.Print dateIBM10(meYBIACPT0.SOLDEDMO, True);
            End If
'____________________

'    XPrt.CurrentX = prtMinX + 6800: XPrt.Print dateIBM10(meYBIACPT0.SOLDEDMO, True);
End If
XPrt.CurrentX = prtMinX + 11600: XPrt.Print meYBIACPT0.COMPTEDEV;

XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 8

End Sub

Public Sub prtSAB_Balance_B_Total_Prt(lK As Long, lSAB_BALANCE As typeSAB_BALANCE)
Dim curX As Currency, X As String

prtSAB_Balance_NewLine
If lSAB_BALANCE.iPrint_Trame <> 255 Then
    XPrt.CurrentY = XPrt.CurrentY - 20
    Call frmElpPrt.prtTrame(prtMinX + 6800, XPrt.CurrentY, prtMinX + 7500 - 20, XPrt.CurrentY + prtlineHeight, " ", lSAB_BALANCE.iPrint_Trame)
    Call frmElpPrt.prtTrame(prtMinX + 7520, XPrt.CurrentY, prtMinX + 12000 - 20, XPrt.CurrentY + prtlineHeight, " ", lSAB_BALANCE.iPrint_Trame)
    Call frmElpPrt.prtTrame(prtMinX + 12020, XPrt.CurrentY, prtMaxX - 20, XPrt.CurrentY + prtlineHeight, " ", lSAB_BALANCE.iPrint_Trame)
    XPrt.CurrentY = XPrt.CurrentY + 20
End If

If lSAB_BALANCE.blnPrint_Line Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight - 30
    XPrt.Line (prtMinX + 6800, XPrt.CurrentY)-(prtMaxX - 20, XPrt.CurrentY)
    XPrt.CurrentY = XPrt.CurrentY - prtlineHeight + 30
End If

XPrt.FontSize = 6
XPrt.FontBold = lSAB_BALANCE.blnPrint_Fontbold

If lK = 6 Then
    XPrt.CurrentX = prtMinX + 400: XPrt.Print lSAB_BALANCE.COMPTEOBL;
    If blnBalance_Pays Then
        XPrt.FontSize = 6: XPrt.FontBold = False
        XPrt.CurrentX = prtMinX - 180: XPrt.Print lSAB_BALANCE.CLIENARSD;
    End If

Else
    XPrt.CurrentX = prtMinX + 7000: XPrt.Print lSAB_BALANCE.COMPTEOBL;
End If



XPrt.CurrentX = prtMinX + 2000: XPrt.Print lSAB_BALANCE.COMPTEINT;
XPrt.CurrentX = prtMinX + 11600: XPrt.Print lSAB_BALANCE.COMPTEDEV;

curX = Abs(lSAB_BALANCE.DB)
If curX <> 0 Then
    X = Format$(curX, "### ### ### ### ##0.00")
    XPrt.CurrentX = prtMinX + 9400 - XPrt.TextWidth(X)
    XPrt.Print X;
End If

curX = Abs(lSAB_BALANCE.CR)
If curX <> 0 Then
    X = Format$(curX, "### ### ### ### ##0.00")
    XPrt.CurrentX = prtMinX + 11400 - XPrt.TextWidth(X)
    XPrt.Print X;
End If

curX = Abs(lSAB_BALANCE.DB_EUR)
If curX <> 0 Then
    X = Format$(curX, "### ### ### ### ##0.00")
    XPrt.CurrentX = prtMaxX - 2100 - XPrt.TextWidth(X)
    XPrt.Print X;
End If

curX = Abs(lSAB_BALANCE.CR_EUR)
If curX <> 0 Then
    X = Format$(curX, "### ### ### ### ##0.00")
    XPrt.CurrentX = prtMaxX - 100 - XPrt.TextWidth(X)
    XPrt.Print X;
End If



If XPrt.FontSize = 6 Then XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 8

End Sub

Public Sub prtSAB_Balance_B_Dev_Classe_Prt()
Dim curX As Currency, X As String
Dim curDB As Currency, curCR As Currency
Dim blnSaut As Boolean, blnErreur As Boolean
Dim wNb As Integer

blnErreur = False
wNb = arrDev_Nb + 2
arrDEV_C1A5(wNb) = zSAB_BALANCE
arrDEV_C6A8(wNb) = zSAB_BALANCE
arrDEV_C9(wNb) = zSAB_BALANCE

arrDEV_C1A5(arrDev_Nb + 1).COMPTEDEV = "???"
arrDEV_C6A8(arrDev_Nb + 1).COMPTEDEV = "???"
arrDEV_C9(arrDev_Nb + 1).COMPTEDEV = "???"
arrDEV_C1A5(arrDev_Nb + 2).COMPTEDEV = "***"
arrDEV_C6A8(arrDev_Nb + 2).COMPTEDEV = "***"
arrDEV_C9(arrDev_Nb + 2).COMPTEDEV = "***"

prtTitleText = "Récapitulatif Bilan / Hors Bilan " & xTitleText
frmElpPrt.prtNewPage
prtSAB_Balance_Form
XPrt.FontSize = 6

For arrDev_K = 1 To wNb
    blnSaut = False
    If arrDEV_C1A5(arrDev_K).DB_EUR <> 0 Or arrDEV_C1A5(arrDev_K).CR_EUR <> 0 Then
        XPrt.FontBold = True
        blnSaut = True
        prtSAB_Balance_NewLine
        XPrt.CurrentX = prtMinX + 5000: XPrt.Print arrDEV_C1A5(arrDev_K).COMPTEDEV & " - Total classes 1 à 5 ";
        XPrt.CurrentX = prtMinX + 11600: XPrt.Print arrDEV_C1A5(arrDev_K).COMPTEDEV;
    
        curX = Abs(arrDEV_C1A5(arrDev_K).DB)
        If curX <> 0 Then
            X = Format$(curX, "### ### ### ### ##0.00")
            XPrt.CurrentX = prtMinX + 9400 - XPrt.TextWidth(X)
            XPrt.Print X;
        End If
        
        curX = Abs(arrDEV_C1A5(arrDev_K).CR)
        If curX <> 0 Then
            X = Format$(curX, "### ### ### ### ##0.00")
            XPrt.CurrentX = prtMinX + 11400 - XPrt.TextWidth(X)
            XPrt.Print X;
        End If
        
        curX = Abs(arrDEV_C1A5(arrDev_K).DB_EUR)
        If curX <> 0 Then
            X = Format$(curX, "### ### ### ### ##0.00")
            XPrt.CurrentX = prtMaxX - 2100 - XPrt.TextWidth(X)
            XPrt.Print X;
        End If
        
        curX = Abs(arrDEV_C1A5(arrDev_K).CR_EUR)
        If curX <> 0 Then
            X = Format$(curX, "### ### ### ### ##0.00")
            XPrt.CurrentX = prtMaxX - 100 - XPrt.TextWidth(X)
            XPrt.Print X;
        End If
    End If

    If arrDEV_C6A8(arrDev_K).DB_EUR <> 0 Or arrDEV_C6A8(arrDev_K).CR_EUR <> 0 Then
        XPrt.FontBold = True
        blnSaut = True
        prtSAB_Balance_NewLine
        XPrt.CurrentX = prtMinX + 5000: XPrt.Print arrDEV_C6A8(arrDev_K).COMPTEDEV & " - Total classes 6 à 8 ";
        XPrt.CurrentX = prtMinX + 11600: XPrt.Print arrDEV_C6A8(arrDev_K).COMPTEDEV;
    
        curX = Abs(arrDEV_C6A8(arrDev_K).DB)
        If curX <> 0 Then
            X = Format$(curX, "### ### ### ### ##0.00")
            XPrt.CurrentX = prtMinX + 9400 - XPrt.TextWidth(X)
            XPrt.Print X;
        End If
        
        curX = Abs(arrDEV_C6A8(arrDev_K).CR)
        If curX <> 0 Then
            X = Format$(curX, "### ### ### ### ##0.00")
            XPrt.CurrentX = prtMinX + 11400 - XPrt.TextWidth(X)
            XPrt.Print X;
        End If
        
        curX = Abs(arrDEV_C6A8(arrDev_K).DB_EUR)
        If curX <> 0 Then
            X = Format$(curX, "### ### ### ### ##0.00")
            XPrt.CurrentX = prtMaxX - 2100 - XPrt.TextWidth(X)
            XPrt.Print X;
        End If
        
        curX = Abs(arrDEV_C6A8(arrDev_K).CR_EUR)
        If curX <> 0 Then
            X = Format$(curX, "### ### ### ### ##0.00")
            XPrt.CurrentX = prtMaxX - 100 - XPrt.TextWidth(X)
            XPrt.Print X;
        End If
    End If
    If arrDEV_C9(arrDev_K).DB_EUR <> 0 Or arrDEV_C9(arrDev_K).CR_EUR <> 0 Then
        XPrt.FontBold = False
        blnSaut = True
        prtSAB_Balance_NewLine
        XPrt.CurrentX = prtMinX + 5000: XPrt.Print arrDEV_C9(arrDev_K).COMPTEDEV & " - Total classe 9 ";
        XPrt.CurrentX = prtMinX + 11600: XPrt.Print arrDEV_C9(arrDev_K).COMPTEDEV;
    
        curX = Abs(arrDEV_C9(arrDev_K).DB)
        If curX <> 0 Then
            X = Format$(curX, "### ### ### ### ##0.00")
            XPrt.CurrentX = prtMinX + 9400 - XPrt.TextWidth(X)
            XPrt.Print X;
        End If
        
        curX = Abs(arrDEV_C9(arrDev_K).CR)
        If curX <> 0 Then
            X = Format$(curX, "### ### ### ### ##0.00")
            XPrt.CurrentX = prtMinX + 11400 - XPrt.TextWidth(X)
            XPrt.Print X;
        End If
        
        curX = Abs(arrDEV_C9(arrDev_K).DB_EUR)
        If curX <> 0 Then
            X = Format$(curX, "### ### ### ### ##0.00")
            XPrt.CurrentX = prtMaxX - 2100 - XPrt.TextWidth(X)
            XPrt.Print X;
        End If
        
        curX = Abs(arrDEV_C9(arrDev_K).CR_EUR)
        If curX <> 0 Then
            X = Format$(curX, "### ### ### ### ##0.00")
            XPrt.CurrentX = prtMaxX - 100 - XPrt.TextWidth(X)
            XPrt.Print X;
        End If
    End If
    
    curX = arrDEV_C1A5(arrDev_K).DB + arrDEV_C1A5(arrDev_K).CR _
            + arrDEV_C6A8(arrDev_K).DB + arrDEV_C6A8(arrDev_K).CR
    If curX <> 0 And blnChk_BalanceEquilibrée Then
        blnErreur = True
        prtSAB_Balance_NewLine
        XPrt.CurrentX = prtMinX + 2000: XPrt.Print arrDEV_C1A5(arrDev_K).COMPTEDEV & "   ?????????? ERREUR BILAN ";
        X = Format$(curX, "### ### ### ### ##0.00")
        XPrt.CurrentX = prtMinX + 1900 - XPrt.TextWidth(X)
        XPrt.Print X;

    End If
    
    
    curX = arrDEV_C9(arrDev_K).DB + arrDEV_C9(arrDev_K).CR
    If curX <> 0 And blnChk_BalanceEquilibrée Then
        blnErreur = True
        prtSAB_Balance_NewLine
        XPrt.CurrentX = prtMinX + 2000: XPrt.Print arrDEV_C9(arrDev_K).COMPTEDEV & "   ?????????? ERREUR HORS-BILAN ";
        X = Format$(curX, "### ### ### ### ##0.00")
        XPrt.CurrentX = prtMinX + 1900 - XPrt.TextWidth(X)
        XPrt.Print X;
    End If
   If arrDev_K < wNb Then
        arrDEV_C1A5(wNb).DB_EUR = arrDEV_C1A5(wNb).DB_EUR + arrDEV_C1A5(arrDev_K).DB_EUR
        arrDEV_C1A5(wNb).CR_EUR = arrDEV_C1A5(wNb).CR_EUR + arrDEV_C1A5(arrDev_K).CR_EUR
        arrDEV_C6A8(wNb).DB_EUR = arrDEV_C6A8(wNb).DB_EUR + arrDEV_C6A8(arrDev_K).DB_EUR
        arrDEV_C6A8(wNb).CR_EUR = arrDEV_C6A8(wNb).CR_EUR + arrDEV_C6A8(arrDev_K).CR_EUR
        arrDEV_C9(wNb).DB_EUR = arrDEV_C9(wNb).DB_EUR + arrDEV_C9(arrDev_K).DB_EUR
        arrDEV_C9(wNb).CR_EUR = arrDEV_C9(wNb).CR_EUR + arrDEV_C9(arrDev_K).CR_EUR
   End If
   
   If blnSaut Then prtSAB_Balance_NewLine

Next arrDev_K

XPrt.FontBold = True
blnSaut = True
prtSAB_Balance_NewLine
XPrt.CurrentX = prtMinX + 5000: XPrt.Print arrDEV_C9(wNb).COMPTEDEV & " - Total classes 1-8 ";

curDB = Abs(arrDEV_C1A5(wNb).DB_EUR + arrDEV_C6A8(wNb).DB_EUR)
If curDB <> 0 Then
    X = Format$(curDB, "### ### ### ### ##0.00")
    XPrt.CurrentX = prtMaxX - 2100 - XPrt.TextWidth(X)
    XPrt.Print X;
End If

curCR = Abs(arrDEV_C1A5(wNb).CR_EUR + arrDEV_C6A8(wNb).CR_EUR)
If curCR <> 0 Then
    X = Format$(curCR, "### ### ### ### ##0.00")
    XPrt.CurrentX = prtMaxX - 100 - XPrt.TextWidth(X)
    XPrt.Print X;
End If

If blnChk_BalanceEquilibrée Then
    Call frmElpPrt.prtTrame(prtMinX + 7500 + 20, XPrt.CurrentY, prtMinX + 12000 - 20, XPrt.CurrentY + prtlineHeight, " ", 240)
    If blnErreur Then
        X = "??? ERREUR ???"
    Else
        X = "Balance équilibrée"
    End If
    XPrt.FontSize = 12
    frmElpPrt.prtCentré prtMinX + 9750, X
End If

End Sub


Public Sub prtSAB_Balance_Mvt(lYBIAMVT0 As typeYBIAMVT0)
prtSAB_Balance_NewLine

XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 6
XPrt.FontItalic = True
 prtSAB_Balance_Montant lYBIAMVT0.MOUVEMMON
XPrt.CurrentX = prtMinX + 400: XPrt.Print lYBIAMVT0.MOUVEMOPE & " " & lYBIAMVT0.MOUVEMNUM&; " " & lYBIAMVT0.MOUVEMEVE;
XPrt.CurrentX = prtMinX + 2000: XPrt.Print Trim(lYBIAMVT0.LIBELLIB1) & " " & Trim(lYBIAMVT0.LIBELLIB2) & " " & Trim(lYBIAMVT0.LIBELLIB3);
XPrt.CurrentX = prtMinX + 6800: XPrt.Print dateIBM10(lYBIAMVT0.MOUVEMDTR, True);
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 8
XPrt.FontItalic = False
End Sub

Public Sub prtSAB_Balance_Montant(lcurX As Currency)
Dim X As String

prtSAB_Balance_CV lcurX

X = Format$(Abs(lcurX), "### ### ### ### ##0.00")
If lcurX > 0 Then
    XPrt.CurrentX = prtMinX + 9400 - XPrt.TextWidth(X)
Else
    XPrt.CurrentX = prtMinX + 11400 - XPrt.TextWidth(X)
End If
XPrt.Print X;


X = Format$(Abs(meCV2.Montant), "### ### ### ### ##0.00")
If meCV2.Montant > 0 Then
    XPrt.CurrentX = prtMaxX - 2100 - XPrt.TextWidth(X)
Else
    XPrt.CurrentX = prtMaxX - 100 - XPrt.TextWidth(X)
End If
XPrt.Print X;

End Sub

Public Sub prtSAB_Balance_CV(lcurX As Currency)

meCV1.Montant = lcurX
If meCV1.DeviseIso <> "EUR" Then
    Call CV_Calc(mBalance_CV, meCV1, meCV2)
Else
    meCV2.Montant = lcurX
End If

End Sub

Public Sub prtSAB_Balance_Montant_Cumul(lcurDB As Currency, lcurCR As Currency)
Dim X As String
X = Format$(Abs(lcurDB), "### ### ### ### ##0.00")
XPrt.CurrentX = prtMaxX - 2100 - XPrt.TextWidth(X)
XPrt.Print X;
X = Format$(Abs(lcurCR), "### ### ### ### ##0.00")
XPrt.CurrentX = prtMaxX - 100 - XPrt.TextWidth(X)
XPrt.Print X;

End Sub


Public Sub prtSAB_Balance_NewLine()
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
If XPrt.CurrentY + 300 > prtMaxY Then
    If blnChk_BalanceStock Then XPrt.DrawWidth = 1: XPrt.Line (prtMinX + 12300, prtMinY + prtHeaderHeight)-(prtMinX + 12300, prtMaxY)

    frmElpPrt.prtNewPage
    prtSAB_Balance_Form
End If

End Sub

Public Sub prtSAb_Client_Stat_NewLine()
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
If XPrt.CurrentY + 300 > prtMaxY Then
    prtSAB_Client_Stat_Form_End
    frmElpPrt.prtNewPage
    prtSAB_Client_Stat_Form
End If

End Sub

Public Sub prtSAB_Balance_C_Cumul()
If blnCompte Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    Call frmElpPrt.prtTrame(prtMinX + 400, XPrt.CurrentY - 50, 2000, XPrt.CurrentY + prtlineHeight, " ", 240)
    Call frmElpPrt.prtTrame(prtMinX + 12020, XPrt.CurrentY - 50, prtMaxX - 20, XPrt.CurrentY + prtlineHeight, " ", 240)
    XPrt.FontBold = True
    XPrt.CurrentX = prtMinX + 400
    XPrt.Print "Code résidence : " & mRésidence;
    prtSAB_Balance_Montant_Cumul curCumul_Db, curCumul_Cr
    XPrt.FontBold = False
End If

End Sub

Public Sub prtSAB_Balance_L_Rupture()

nbListe_Line = nbListe_Line + 1

If blnClient_Line Then
    prtSAB_Balance_NewLine
    XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 6: XPrt.FontBold = True

    XPrt.CurrentX = prtMinX + 6000: XPrt.Print prevYBIACPT0.CLIENACLI & " " & prevYBIACPT0.CLIENASIG;

    Call frmElpPrt.prtTrame(prtMinX + 12020, XPrt.CurrentY, prtMaxX - 20, XPrt.CurrentY + prtlineHeight - 50, " ", 240)
    If curClient_Db <> 0 Then
        X = Format$(Abs(curClient_Db), "### ### ### ### ##0.00")
        XPrt.CurrentX = prtMaxX - 2100 - XPrt.TextWidth(X)
        XPrt.Print X;
    End If
    If curClient_Cr <> 0 Then
        X = Format$(Abs(curClient_Cr), "### ### ### ### ##0.00")
        XPrt.CurrentX = prtMaxX - 100 - XPrt.TextWidth(X)
        XPrt.Print X;
    End If
    XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 8:: XPrt.FontBold = False

End If

curW_Db = curW_Db + curClient_Db
curW_Cr = curW_Cr + curClient_Cr

nbClient_Line = 0
curClient_Db = 0: curClient_Cr = 0

End Sub
Public Sub prtSAB_Balance_B_COMPTEDEV()

prtSAB_Balance_B_COMPTEOBL
If blnBalance_Pays Then curCOMPTEDEV_Db = 0: curCOMPTEDEV_Cr = 0

If blnBalance_B_Détail And nbCOMPTEDEV_Line > 0 Then

    prtSAB_Balance_NewLine
    XPrt.FontSize = 8: XPrt.FontBold = True
    Call frmElpPrt.prtTrame(prtMinX + 7520, XPrt.CurrentY, prtMinX + 12000 - 20, XPrt.CurrentY + prtlineHeight, " ", 240)
    Call frmElpPrt.prtTrame(prtMinX + 12020, XPrt.CurrentY, prtMaxX - 20, XPrt.CurrentY + prtlineHeight, " ", 240)
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.Line (prtMinX + 7500, XPrt.CurrentY)-(prtMaxX - 20, XPrt.CurrentY)
    XPrt.CurrentY = XPrt.CurrentY - prtlineHeight

    
    If blnBalance_B_COMPTEDEV Then
        XPrt.CurrentX = prtMinX + 11600: XPrt.Print prevYBIACPT0.COMPTEDEV;
        If curCOMPTEDEV_Db <> 0 Then
            X = Format$(Abs(curCOMPTEDEV_Db), "### ### ### ### ##0.00")
            XPrt.CurrentX = prtMinX + 9400 - XPrt.TextWidth(X)
            XPrt.Print X;
        End If
        If curCOMPTEDEV_Cr <> 0 Then
             X = Format$(Abs(curCOMPTEDEV_Cr), "### ### ### ### ##0.00")
            XPrt.CurrentX = prtMinX + 11400 - XPrt.TextWidth(X)
            XPrt.Print X;
        End If
    End If

    If curCOMPTEDEV_Db_EUR <> 0 Then
        X = Format$(Abs(curCOMPTEDEV_Db_EUR), "### ### ### ### ##0.00")
        XPrt.CurrentX = prtMaxX - 2100 - XPrt.TextWidth(X)
        XPrt.Print X;
    End If
    If curCOMPTEDEV_Cr_EUR <> 0 Then
        X = Format$(Abs(curCOMPTEDEV_Cr_EUR), "### ### ### ### ##0.00")
        XPrt.CurrentX = prtMaxX - 100 - XPrt.TextWidth(X)
        XPrt.Print X;
    End If

     XPrt.FontBold = False

End If

curW_Db = curW_Db + curCOMPTEDEV_Db
curW_Cr = curW_Cr + curCOMPTEDEV_Cr

curCOMPTEDEV_Db = 0: curCOMPTEDEV_Cr = 0
curCOMPTEDEV_Db_EUR = 0: curCOMPTEDEV_Cr_EUR = 0
nbCOMPTEDEV_Line = 0
End Sub
Public Sub prtSAB_Balance_B_COMPTEOBL()


Call srvYPLAN0_Import_Read(prevYBIACPT0.COMPTEOBL, meYPLAN0)

If blnBalance_B_Détail And détailSAB_BALANCE.blnPrint And nbCOMPTEOBL_Line > 0 Then
    
    prtSAB_Balance_NewLine
    XPrt.FontBold = détailSAB_BALANCE.blnPrint_Fontbold
    
    If détailSAB_BALANCE.iPrint_Trame <> 255 Then
        Call frmElpPrt.prtTrame(prtMinX + 2000, XPrt.CurrentY, prtMinX + 7500 - 20, XPrt.CurrentY + prtlineHeight, " ", détailSAB_BALANCE.iPrint_Trame)
        Call frmElpPrt.prtTrame(prtMinX + 7520, XPrt.CurrentY, prtMinX + 12000 - 20, XPrt.CurrentY + prtlineHeight, " ", détailSAB_BALANCE.iPrint_Trame)
        Call frmElpPrt.prtTrame(prtMinX + 12020, XPrt.CurrentY, prtMaxX - 20, XPrt.CurrentY + prtlineHeight, " ", détailSAB_BALANCE.iPrint_Trame)
    End If
    If détailSAB_BALANCE.blnPrint_Line Then
        XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
        XPrt.Line (prtMinX + 2000, XPrt.CurrentY)-(prtMaxX - 20, XPrt.CurrentY)
        XPrt.CurrentY = XPrt.CurrentY - prtlineHeight
    End If
    
    XPrt.FontSize = 6: XPrt.CurrentY = XPrt.CurrentY + Height8_6

    XPrt.CurrentX = prtMinX + 2000: XPrt.Print meYPLAN0.PLANINTIT;
    XPrt.CurrentX = prtMinX + 6500: XPrt.Print prevYBIACPT0.COMPTEDEV & " " & prevYBIACPT0.COMPTEOBL;
    

    If blnBalance_B_COMPTEDEV Then
        XPrt.CurrentX = prtMinX + 11600: XPrt.Print prevYBIACPT0.COMPTEDEV;
        If curCOMPTEOBL_Db <> 0 Then
            X = Format$(Abs(curCOMPTEOBL_Db), "### ### ### ### ##0.00")
            XPrt.CurrentX = prtMinX + 9400 - XPrt.TextWidth(X)
            XPrt.Print X;
        End If
        If curCOMPTEOBL_Cr <> 0 Then
             X = Format$(Abs(curCOMPTEOBL_Cr), "### ### ### ### ##0.00")
            XPrt.CurrentX = prtMinX + 11400 - XPrt.TextWidth(X)
            XPrt.Print X;
        End If
    End If

    If curCOMPTEOBL_Db_EUR <> 0 Then
        X = Format$(Abs(curCOMPTEOBL_Db_EUR), "### ### ### ### ##0.00")
        XPrt.CurrentX = prtMaxX - 2100 - XPrt.TextWidth(X)
        XPrt.Print X;
    End If
    If curCOMPTEOBL_Cr_EUR <> 0 Then
        X = Format$(Abs(curCOMPTEOBL_Cr_EUR), "### ### ### ### ##0.00")
        XPrt.CurrentX = prtMaxX - 100 - XPrt.TextWidth(X)
        XPrt.Print X;
    End If

     XPrt.FontBold = False
    XPrt.CurrentY = XPrt.CurrentY - Height8_6
End If

curW_Db = curW_Db + curCOMPTEOBL_Db
curW_Cr = curW_Cr + curCOMPTEOBL_Cr

If arrSAB_BALANCE_Nb = arrSAB_BALANCE_NbMax Then
    arrSAB_BALANCE_NbMax = arrSAB_BALANCE_NbMax + 100
    ReDim Preserve arrSAB_BALANCE(arrSAB_BALANCE_NbMax)
End If
arrSAB_BALANCE_Nb = arrSAB_BALANCE_Nb + 1
arrSAB_BALANCE(arrSAB_BALANCE_Nb) = totalSAB_BALANCE(6)
arrSAB_BALANCE(arrSAB_BALANCE_Nb).CLIENARSD = prevYBIACPT0.CLIENARSD
arrSAB_BALANCE(arrSAB_BALANCE_Nb).COMPTEDEV = prevYBIACPT0.COMPTEDEV
arrSAB_BALANCE(arrSAB_BALANCE_Nb).COMPTEOBL = prevYBIACPT0.COMPTEOBL
arrSAB_BALANCE(arrSAB_BALANCE_Nb).COMPTECOM = ""
arrSAB_BALANCE(arrSAB_BALANCE_Nb).COMPTEINT = meYPLAN0.PLANINTIT
arrSAB_BALANCE(arrSAB_BALANCE_Nb).DB = arrSAB_BALANCE(arrSAB_BALANCE_Nb).DB + curCOMPTEOBL_Db
arrSAB_BALANCE(arrSAB_BALANCE_Nb).CR = arrSAB_BALANCE(arrSAB_BALANCE_Nb).CR + curCOMPTEOBL_Cr
arrSAB_BALANCE(arrSAB_BALANCE_Nb).DB_EUR = arrSAB_BALANCE(arrSAB_BALANCE_Nb).DB_EUR + curCOMPTEOBL_Db_EUR
arrSAB_BALANCE(arrSAB_BALANCE_Nb).CR_EUR = arrSAB_BALANCE(arrSAB_BALANCE_Nb).CR_EUR + curCOMPTEOBL_Cr_EUR


curCOMPTEOBL_Db = 0: curCOMPTEOBL_Cr = 0
curCOMPTEOBL_Db_EUR = 0: curCOMPTEOBL_Cr_EUR = 0
nbCOMPTEOBL_Line = 0
End Sub


Public Sub prtSAB_Balance_L_Line()

If prevYBIACPT0.CLIENACLI <> meYBIACPT0.CLIENACLI Then
    If Trim(prevYBIACPT0.COMPTECOM) <> "" Then
        prtSAB_Balance_L_Rupture
    End If
End If

prtSAB_Balance_L "L", meYBIACPT0.SOLDECEN

nbClient_Line = nbClient_Line + 1

If meCV2.Montant > 0 Then
    curClient_Db = curClient_Db + meCV2.Montant
    curListe_Db = curListe_Db + meCV2.Montant
Else
    curClient_Cr = curClient_Cr + meCV2.Montant
    curListe_Cr = curListe_Cr + meCV2.Montant
End If

End Sub

Public Sub prtSAB_Balance_B_Line()
Dim curX As Currency, curX1 As Currency, curX2 As Currency
Dim X As String, XS As String

If blnBalance_Pays Then                                  ' Tri Pays / PCi / Compte / Dev
    If prevYBIACPT0.CLIENARSD <> meYBIACPT0.CLIENARSD Then
        prtSAB_Balance_B_COMPTEOBL
        xPays = srvYBIATAB0_Pays(meYBIACPT0.CLIENARSD)
        prtTitleText = xPays & " - " & xTitleText
        If blnBalance_B_Détail Then
            frmElpPrt.prtNewPage
            prtSAB_Balance_Form
        End If
    End If
    If prevYBIACPT0.COMPTEOBL <> meYBIACPT0.COMPTEOBL Or prevYBIACPT0.COMPTEDEV <> meYBIACPT0.COMPTEDEV Then
        If Trim(prevYBIACPT0.COMPTEOBL) <> "" Then
            prtSAB_Balance_B_COMPTEOBL
        End If
    End If
Else
    If prevYBIACPT0.COMPTEDEV <> meYBIACPT0.COMPTEDEV Then
        If Trim(prevYBIACPT0.COMPTEDEV) <> "" Then
            prtSAB_Balance_B_COMPTEDEV
        End If
    End If
    If prevYBIACPT0.COMPTEOBL <> meYBIACPT0.COMPTEOBL Then
        If Trim(prevYBIACPT0.COMPTEOBL) <> "" Then
            prtSAB_Balance_B_COMPTEOBL
        End If
    End If
End If
    If prevYBIACPT0.COMPTEOBL <> meYBIACPT0.COMPTEOBL Then
        If Trim(prevYBIACPT0.COMPTEOBL) <> "" Then
            prtSAB_Balance_B_COMPTEOBL
        End If
    End If

If mBalance_YSOLDE0 = 0 Then
    curX = meYBIACPT0.SOLDECEN
Else
    srvYSOLDE0_Import_Read meYBIACPT0.COMPTECOM, xYSOLDE0
    Select Case mBalance_YSOLDE0
            Case 1: curX = xYSOLDE0.SOLDEC01
            Case 2: curX = xYSOLDE0.SOLDEC02
            Case 3: curX = xYSOLDE0.SOLDEC03
            Case 4: curX = xYSOLDE0.SOLDEC04
            Case 5: curX = xYSOLDE0.SOLDEC05
            Case 6: curX = xYSOLDE0.SOLDEC06
            Case 7: curX = xYSOLDE0.SOLDEC07
            Case 8: curX = xYSOLDE0.SOLDEC08
            Case 9: curX = xYSOLDE0.SOLDEC09
            Case 10: curX = xYSOLDE0.SOLDEC10
            Case 11: curX = xYSOLDE0.SOLDEC11
            Case 12: curX = xYSOLDE0.SOLDEC12
    End Select
End If

If Not blnBalance_Compte_Soldé Or curX <> 0 Then
    nbCOMPTEOBL_Line = nbCOMPTEOBL_Line + 1
    nbCOMPTEDEV_Line = nbCOMPTEDEV_Line + 1
    If blnBalance_B_Détail Then
    
        prtSAB_Balance_L "B", curX
        
        If blnChk_BalanceStock Then
            XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 6
           If meDORCPTDMV > 0 Then
                XPrt.CurrentX = prtMinX + 6500: XPrt.Print dateIBM10(meDORCPTDMV, True);
                If meDORCPTDMV <> meYBIACPT0.SOLDEDMO Then XPrt.Print " +";
            Else
                XPrt.CurrentX = prtMinX + 6500: XPrt.Print dateIBM10(meYBIACPT0.SOLDEDMO, True);
            End If
            XPrt.FontBold = True
            If meYSTOMON <> -2 Then
                XPrt.FontBold = True
                If meYSTOMON = -1 Then
                    X = "??"
                Else
                    curX1 = Abs(meYBIACPT0.SOLDECEN)
                    curX2 = Abs(meYSTOMON)
                    curX = Abs(curX1 - curX2)
                    If curX = 0 Then
                        X = "="
                    Else
                        X = Format$(Abs(curX2), "### ### ### ### ##0.00")
                        XPrt.CurrentX = prtMinX + 6400 - XPrt.TextWidth(X)
                        XPrt.Print X;
    
                        X = "##"
                    End If
                End If
                XPrt.CurrentX = prtMinX + 12100: XPrt.Print X;
                XPrt.FontBold = False
    
            End If
            
            XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 8
       End If
        
        If meYBIACPT0.PLANCOPRO = "CAV" Or meYBIACPT0.PLANCOPRO = "LOR" Then
            If meYBIACPT0.COMPTEOUV > wAMJ_6M_00 And meYBIACPT0.COMPTEOUV < wAMJ_6M_99 Then
                XPrt.FontBold = True
                XPrt.CurrentX = prtMinX + 12050: XPrt.Print "6M";
                XPrt.FontBold = False
            End If
        End If
        
        If blnBalance_Pays Then
            XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 6
            XPrt.CurrentX = prtMinX - 180: XPrt.Print meYBIACPT0.CLIENARSD;
            XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 8
        End If
    Else
        prtSAB_Balance_CV curX
    End If
    
    nbClient_Line = nbClient_Line + 1
    
    prtSAB_Balance_B_Dev_Classe_Cumul
    
    If meCV2.Montant > 0 Then
        curCOMPTEOBL_Db = curCOMPTEOBL_Db + meCV1.Montant
        curCOMPTEDEV_Db = curCOMPTEDEV_Db + meCV1.Montant
        curCOMPTEOBL_Db_EUR = curCOMPTEOBL_Db_EUR + meCV2.Montant
        curCOMPTEDEV_Db_EUR = curCOMPTEDEV_Db_EUR + meCV2.Montant
        curListe_Db = curListe_Db + meCV2.Montant
    Else
        curCOMPTEOBL_Cr = curCOMPTEOBL_Cr + meCV1.Montant
        curCOMPTEDEV_Cr = curCOMPTEDEV_Cr + meCV1.Montant
        curCOMPTEOBL_Cr_EUR = curCOMPTEOBL_Cr_EUR + meCV2.Montant
        curCOMPTEDEV_Cr_EUR = curCOMPTEDEV_Cr_EUR + meCV2.Montant
        curListe_Cr = curListe_Cr + meCV2.Montant
    End If
    If blnFile_CSV Then
    
        X = meYBIACPT0.COMPTEDEV & ";" & meYBIACPT0.COMPTEOBL & ";" & meYBIACPT0.COMPTECOM & ";" & meYBIACPT0.COMPTEINT & ";"
        If meCV2.Montant > 0 Then
            XS = cur_AbsV(meCV1.Montant) & "; ;" & cur_AbsV(meCV2.Montant) & "; "
        Else
            XS = " ;" & cur_AbsV(meCV1.Montant) & "; ;" & cur_AbsV(meCV2.Montant)
        End If
        Call File_Export_Monitor("Print", idFile_CSV, X & XS)
    End If
End If

End Sub



Public Sub prtSAB_Balance_B_Fin()
Dim I As Long, K As Long
Dim blnOk As Boolean

prtSAB_Balance_B_COMPTEOBL
If blnBalance_B_Détail Then
    XPrt.DrawWidth = 10
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight + 20
    XPrt.Line (prtMinX + 7500, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY)
    XPrt.DrawWidth = 3
End If

If blnBalance_Pays Then                                  ' Tri Pays / PCi / Compte / Dev
    prtSAB_Balance_B_COMPTEOBL
Else
    prtSAB_Balance_B_COMPTEDEV
    prtSAB_Balance_B_COMPTEOBL
End If

If blnBalance_B_Récap Then
    prtTitleText = "Récapitulatif de la " & xTitleText
    
    If blnBalance_B_Détail Then
            frmElpPrt.prtNewPage
            prtSAB_Balance_Form
    End If
    
    prevSAB_BALANCE = arrSAB_BALANCE(1)
    
    For I = 1 To arrSAB_BALANCE_Nb
        blnOk = True
        
        If Not blnBalance_Pays Then
            If prevSAB_BALANCE.COMPTEDEV <> arrSAB_BALANCE(I).COMPTEDEV Then
            
                blnOk = False
                prtSAB_Balance_B_Total_Cumul 0
            End If
        End If
        
        If blnOk Then
            For K = 1 To 5
                If mId$(prevSAB_BALANCE.COMPTEOBL, 1, K) <> mId$(arrSAB_BALANCE(I).COMPTEOBL, 1, K) Then
                    prtSAB_Balance_B_Total_Cumul K
                    Exit For
                End If
            Next K
        End If
        
        If totalSAB_BALANCE(6).blnPrint Then
            If arrSAB_BALANCE(I).DB <> 0 Or arrSAB_BALANCE(I).CR <> 0 Then
                prtSAB_Balance_B_Total_Prt 6, arrSAB_BALANCE(I)
            End If
        End If
        
        prevSAB_BALANCE = arrSAB_BALANCE(I)
        totalSAB_BALANCE(5).DB = totalSAB_BALANCE(5).DB + prevSAB_BALANCE.DB
        totalSAB_BALANCE(5).CR = totalSAB_BALANCE(5).CR + prevSAB_BALANCE.CR
        totalSAB_BALANCE(5).DB_EUR = totalSAB_BALANCE(5).DB_EUR + prevSAB_BALANCE.DB_EUR
        totalSAB_BALANCE(5).CR_EUR = totalSAB_BALANCE(5).CR_EUR + prevSAB_BALANCE.CR_EUR
    
    Next I
    If blnBalance_Pays Then
        prevSAB_BALANCE.COMPTEDEV = "***"
        For K = 0 To 6
                totalSAB_BALANCE(K).COMPTEDEV = ""
                totalSAB_BALANCE(K).DB = 0: totalSAB_BALANCE(K).CR = 0
        Next K
    End If
    prtSAB_Balance_B_Total_Cumul 0

End If

If blnBalance_Récap_Bilan Then
    prtSAB_Balance_B_Dev_Classe_Prt
End If
End Sub

Public Sub prtSAB_Balance_Total_Init()
Dim I As Integer

ReDim arrSAB_BALANCE(101) As typeSAB_BALANCE
arrSAB_BALANCE_NbMax = 100: arrSAB_BALANCE_Nb = 0
zSAB_BALANCE.COMPTEDEV = ""
zSAB_BALANCE.COMPTEOBL = ""
zSAB_BALANCE.COMPTECOM = ""
zSAB_BALANCE.COMPTEINT = ""
zSAB_BALANCE.DB = 0
zSAB_BALANCE.CR = 0
zSAB_BALANCE.DB_EUR = 0
zSAB_BALANCE.CR_EUR = 0
zSAB_BALANCE.blnPrint = False
zSAB_BALANCE.blnPrint_Fontbold = False
zSAB_BALANCE.blnPrint_Line = False
zSAB_BALANCE.iPrint_Trame = 0

prevSAB_BALANCE = zSAB_BALANCE
détailSAB_BALANCE = zSAB_BALANCE
For I = 0 To 10
    totalSAB_BALANCE(I) = zSAB_BALANCE
Next I

arrDev_Nb = 0
xYbase.ID = constYBIATAB0
xYbase.K1 = "DEVISE"
xYbase.Method = "Seek>"
Do
    intReturn = tableYBase_Read(xYbase)
    If Trim(xYbase.ID) <> constYBIATAB0 Then intReturn = -1
    If mId$(xYbase.K1, 1, 6) <> "DEVISE" Then intReturn = -1
    If intReturn = 0 Then
        arrDev_Nb = arrDev_Nb + 1
        arrDEV_C1A5(arrDev_Nb) = zSAB_BALANCE
        arrDEV_C1A5(arrDev_Nb).COMPTEDEV = mId$(xYbase.Text, 25, 3)
        arrDEV_C6A8(arrDev_Nb) = zSAB_BALANCE
        arrDEV_C6A8(arrDev_Nb).COMPTEDEV = mId$(xYbase.Text, 25, 3)
        arrDEV_C9(arrDev_Nb) = zSAB_BALANCE
        arrDEV_C9(arrDev_Nb).COMPTEDEV = mId$(xYbase.Text, 25, 3)
    End If
        
Loop Until intReturn <> 0

' Code devise non trouvé
arrDEV_C1A5(arrDev_Nb + 1) = zSAB_BALANCE
arrDEV_C6A8(arrDev_Nb + 1) = zSAB_BALANCE
arrDEV_C9(arrDev_Nb + 1) = zSAB_BALANCE


End Sub

Public Sub prtSAB_Balance_B_Total_Cumul(lK As Long)
Dim K As Long

For K = 5 To lK Step -1
    totalSAB_BALANCE(K).COMPTEDEV = prevSAB_BALANCE.COMPTEDEV
    totalSAB_BALANCE(K).COMPTEOBL = mId$(prevSAB_BALANCE.COMPTEOBL, 1, K)
    
    If totalSAB_BALANCE(K).blnPrint Then prtSAB_Balance_B_Total_Prt K, totalSAB_BALANCE(K)
    
    If K > 0 Then
        totalSAB_BALANCE(K - 1).DB = totalSAB_BALANCE(K - 1).DB + totalSAB_BALANCE(K).DB
        totalSAB_BALANCE(K - 1).CR = totalSAB_BALANCE(K - 1).CR + totalSAB_BALANCE(K).CR
        totalSAB_BALANCE(K - 1).DB_EUR = totalSAB_BALANCE(K - 1).DB_EUR + totalSAB_BALANCE(K).DB_EUR
        totalSAB_BALANCE(K - 1).CR_EUR = totalSAB_BALANCE(K - 1).CR_EUR + totalSAB_BALANCE(K).CR_EUR
    End If
    totalSAB_BALANCE(K).DB = 0
    totalSAB_BALANCE(K).CR = 0
    totalSAB_BALANCE(K).DB_EUR = 0
    totalSAB_BALANCE(K).CR_EUR = 0
Next K
End Sub

Public Sub prtSAB_Balance_B_Total_Init(lX As String, lSAB_BALANCE As typeSAB_BALANCE)

If mId$(lX, 1, 1) = "1" Then lSAB_BALANCE.blnPrint = True  'Imprimer ce niveau
If mId$(lX, 2, 1) = "1" Then lSAB_BALANCE.blnPrint_Fontbold = True  'gras
If mId$(lX, 3, 1) = "1" Then lSAB_BALANCE.blnPrint_Line = True  'ligne séparation
lSAB_BALANCE.iPrint_Trame = 255 - Val(mId$(lX, 4, 3)) 'trame

End Sub

Public Sub prtSAB_Client_Stat_Z(lSAB_Client_Stat As typeSAB_Client_Stat)
lSAB_Client_Stat.Nb_Client = 0
lSAB_Client_Stat.Nb_Compte = 0
lSAB_Client_Stat.Nb_Compte_Annulé = 0
lSAB_Client_Stat.Solde_DB = 0
lSAB_Client_Stat.Solde_CR = 0
End Sub

Public Sub prtSAB_Client_Stat_Cumul()
If SAB_Client_Stat.Nb_Compte <> 0 Then
    SAB_Client_Stat_Actif.Nb_Client = SAB_Client_Stat_Actif.Nb_Client + 1
    SAB_Client_Stat_Actif.Nb_Compte = SAB_Client_Stat_Actif.Nb_Compte + SAB_Client_Stat.Nb_Compte
    SAB_Client_Stat_Actif.Nb_Compte_Annulé = SAB_Client_Stat_Actif.Nb_Compte_Annulé + SAB_Client_Stat.Nb_Compte_Annulé
    SAB_Client_Stat_Actif.Solde_DB = SAB_Client_Stat_Actif.Solde_DB + SAB_Client_Stat.Solde_DB
    SAB_Client_Stat_Actif.Solde_CR = SAB_Client_Stat_Actif.Solde_CR + SAB_Client_Stat.Solde_CR

    SAB_Client_Stat_Produit.Nb_Client = SAB_Client_Stat_Produit.Nb_Client + 1
    SAB_Client_Stat_Produit.Nb_Compte = SAB_Client_Stat_Produit.Nb_Compte + SAB_Client_Stat.Nb_Compte
    SAB_Client_Stat_Produit.Nb_Compte_Annulé = SAB_Client_Stat_Produit.Nb_Compte_Annulé + SAB_Client_Stat.Nb_Compte_Annulé
    SAB_Client_Stat_Produit.Solde_DB = SAB_Client_Stat_Produit.Solde_DB + SAB_Client_Stat.Solde_DB
    SAB_Client_Stat_Produit.Solde_CR = SAB_Client_Stat_Produit.Solde_CR + SAB_Client_Stat.Solde_CR

Else
    If SAB_Client_Stat.Nb_Compte_Annulé <> 0 Then
        SAB_Client_Stat_Annulé.Nb_Client = SAB_Client_Stat_Annulé.Nb_Client + 1
        SAB_Client_Stat_Annulé.Nb_Compte = SAB_Client_Stat_Annulé.Nb_Compte + SAB_Client_Stat.Nb_Compte
        SAB_Client_Stat_Annulé.Nb_Compte_Annulé = SAB_Client_Stat_Annulé.Nb_Compte_Annulé + SAB_Client_Stat.Nb_Compte_Annulé

    End If
End If
prtSAB_Client_Stat_Z SAB_Client_Stat

End Sub

Public Sub prtSAB_Client_Stat_Line()
Dim wId As String
Dim X As String

'prevYBIACPT0.PLANCOPRO,
wId = "SAB         CLIENACAT   CLI" & prevYBIACPT0.CLIENACAT
Call srvYBIATAB0_Import_Read(wId, meYBIATAB0)

prtSAb_Client_Stat_NewLine

XPrt.CurrentX = prtMinX
XPrt.Print prevYBIACPT0.CLIENACAT;
XPrt.CurrentX = prtMinX + 400
XPrt.Print "- " & mId$(meYBIATAB0.BIATABTEXT, 36 + 13, 30);

num_Xprt_Long SAB_Client_Stat_Actif.Nb_Client, prtMinX + 5000
num_Xprt_Currency SAB_Client_Stat_Actif.Solde_DB, prtMinX + 7000
num_Xprt_Currency SAB_Client_Stat_Actif.Solde_CR, prtMinX + 9000

num_Xprt_Long SAB_Client_Stat_Actif.Nb_Compte, prtMinX + 11000
num_Xprt_Long SAB_Client_Stat_Actif.Nb_Compte_Annulé, prtMinX + 12000

num_Xprt_Long SAB_Client_Stat_Annulé.Nb_Client, prtMinX + 14500
num_Xprt_Long SAB_Client_Stat_Annulé.Nb_Compte_Annulé, prtMinX + 15500


prtSAB_Client_Stat_Z SAB_Client_Stat_Actif
prtSAB_Client_Stat_Z SAB_Client_Stat_Annulé

End Sub
Public Sub prtSAB_Client_Stat_Form()
Dim wId As String
Dim X As String

XPrt.FontSize = 7
XPrt.FontBold = True
XPrt.DrawWidth = 2

XPrt.CurrentY = prtMinY + 50

XPrt.CurrentX = prtMinX
XPrt.Print "Produit / catégorie clientèle";

X = "Nb clients"
XPrt.CurrentX = prtMinX + 5000 - XPrt.TextWidth(X)
XPrt.Print X;

X = "Soldes débiteurs"
XPrt.CurrentX = prtMinX + 7000 - XPrt.TextWidth(X)
XPrt.Print X;

X = "Soldes créditeurs"
XPrt.CurrentX = prtMinX + 9000 - XPrt.TextWidth(X)
XPrt.Print X;

X = "Comptes .... Actifs"
XPrt.CurrentX = prtMinX + 11000 - XPrt.TextWidth(X)
XPrt.Print X;

X = "...Annulés"
XPrt.CurrentX = prtMinX + 12000 - XPrt.TextWidth(X)
XPrt.Print X;

X = "Clients non actifs"
XPrt.CurrentX = prtMinX + 14500 - XPrt.TextWidth(X)
XPrt.Print X;

X = "Cpt annulés"
XPrt.CurrentX = prtMinX + 15500 - XPrt.TextWidth(X)
XPrt.Print X;

XPrt.FontSize = 8
XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY)
XPrt.CurrentY = XPrt.CurrentY + 50


End Sub

Public Sub prtSAB_Client_Stat_Produit()
Dim wId As String
Dim X As String

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.25
XPrt.FontSize = 7
XPrt.FontBold = True

Call frmElpPrt.prtTrame(prtMinX + 390, XPrt.CurrentY - 50, prtMaxX, XPrt.CurrentY + prtlineHeight, " ", 240)

XPrt.CurrentX = prtMinX + 400
XPrt.Print "  " & SAB_Client_Stat_Produit_Lib;

num_Xprt_Long SAB_Client_Stat_Produit.Nb_Client, prtMinX + 5000
num_Xprt_Currency SAB_Client_Stat_Produit.Solde_DB, prtMinX + 7000
num_Xprt_Currency SAB_Client_Stat_Produit.Solde_CR, prtMinX + 9000

num_Xprt_Long SAB_Client_Stat_Produit.Nb_Compte, prtMinX + 11000
num_Xprt_Long SAB_Client_Stat_Produit.Nb_Compte_Annulé, prtMinX + 12000

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 0.5
XPrt.FontSize = 8
XPrt.FontBold = False

prtSAB_Client_Stat_Z SAB_Client_Stat_Produit

End Sub

Public Sub prtSAB_Client_Stat_PLANCOPRO(lX As String)
Dim wId As String

wId = "SAB         PLANCOPRO   " & lX
Call srvYBIATAB0_Import_Read(wId, meYBIATAB0)
XPrt.FontBold = True

prtSAb_Client_Stat_NewLine

Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY - 50, prtMaxX, XPrt.CurrentY + prtlineHeight, " ", 240)

XPrt.CurrentX = prtMinX
XPrt.Print lX;
XPrt.CurrentX = prtMinX + 400
SAB_Client_Stat_Produit_Lib = mId$(meYBIATAB0.BIATABTEXT, 36 + 13, 30)
XPrt.Print "- " & SAB_Client_Stat_Produit_Lib;
XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 0.5


End Sub



Public Sub prtSAB_Client_Stat_Form_End()
Dim mCurrenty As Long

mCurrenty = XPrt.CurrentY + prtlineHeight
XPrt.Line (prtMinX + 5100, prtMinY)-(prtMinX + 5100, mCurrenty)
XPrt.Line (prtMinX + 9100, prtMinY)-(prtMinX + 9100, mCurrenty)
XPrt.DrawWidth = 5

XPrt.Line (prtMinX + 12100, prtMinY)-(prtMinX + 12100, mCurrenty)

End Sub

Public Sub prtSAB_Balance_B_Dev_Classe_Cumul()
For arrDev_K = 1 To arrDev_Nb
    If meYBIACPT0.COMPTEDEV = arrDEV_C1A5(arrDev_K).COMPTEDEV Then Exit For
Next arrDev_K
Select Case mId$(meYBIACPT0.COMPTEOBL, 1, 1)
    Case Is <= 5
        
        If meCV2.Montant > 0 Then
            arrDEV_C1A5(arrDev_K).DB = arrDEV_C1A5(arrDev_K).DB + meCV1.Montant
            arrDEV_C1A5(arrDev_K).DB_EUR = arrDEV_C1A5(arrDev_K).DB_EUR + meCV2.Montant
        Else
            arrDEV_C1A5(arrDev_K).CR = arrDEV_C1A5(arrDev_K).CR + meCV1.Montant
            arrDEV_C1A5(arrDev_K).CR_EUR = arrDEV_C1A5(arrDev_K).CR_EUR + meCV2.Montant
        End If
    Case Is <= 8
        If meCV2.Montant > 0 Then
            arrDEV_C6A8(arrDev_K).DB = arrDEV_C6A8(arrDev_K).DB + meCV1.Montant
            arrDEV_C6A8(arrDev_K).DB_EUR = arrDEV_C6A8(arrDev_K).DB_EUR + meCV2.Montant
        Else
            arrDEV_C6A8(arrDev_K).CR = arrDEV_C6A8(arrDev_K).CR + meCV1.Montant
            arrDEV_C6A8(arrDev_K).CR_EUR = arrDEV_C6A8(arrDev_K).CR_EUR + meCV2.Montant
        End If
    Case Else
        If meCV2.Montant > 0 Then
            arrDEV_C9(arrDev_K).DB = arrDEV_C9(arrDev_K).DB + meCV1.Montant
            arrDEV_C9(arrDev_K).DB_EUR = arrDEV_C9(arrDev_K).DB_EUR + meCV2.Montant
        Else
            arrDEV_C9(arrDev_K).CR = arrDEV_C9(arrDev_K).CR + meCV1.Montant
            arrDEV_C9(arrDev_K).CR_EUR = arrDEV_C9(arrDev_K).CR_EUR + meCV2.Montant
        End If
End Select
End Sub
