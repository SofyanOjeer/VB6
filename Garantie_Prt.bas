Attribute VB_Name = "prtGarantie"

'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Private recCompte As typeCompte
Dim X As String, I As Integer, Height8_6 As Integer

Private recCptInfo As typeCptInfo
Public arrTFlux() As typeTFlux, recTFlux As typeTFlux

Public recTope As typeTOpe
Public CV1 As typeCV

Dim CapitalRestantDû As Currency, totalCommission As Currency
Dim blnTableauAmortissement As Boolean

Public P_arrTOpe() As typeTOpe, P_arrTFlux() As typeTFlux
Dim prtLineNb As Integer


Public Sub prtGarantie_Avis(lTOpe As typeTOpe, lTflux As typeTFlux)
On Error GoTo prtError

Set XPrt = Printer
frmElpPrt.Show vbModeless

prtOrientation = vbPRORPortrait
prtTitleText = "Avis"
prtPgmName = "prtGarantie_Avis"
prtTitleUsr = usrName

prtLineNb = 1
prtlineHeight = 300
prtHeaderHeight = 300

prtFormType = ""
frmElpPrt.prtInit

recTope = lTOpe
recTFlux = lTflux

prtGarantie_AvisA4

frmElpPrt.prtEndDoc
frmElpPrt.Hide
Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide
'Public Sub prtAvis(lTOpe As typeTOpe, lTflux As typeTFlux)
'rectflux_Init rectflux
'rectflux.Référence = lTOpe.RéférenceInterne
'rectflux.CodeOpération = lTflux.CodeOpération
'rectflux.Société = SocId$
'rectflux.Agence = SocAgence$
'Call CV_AttributS(lTOpe.Devise, CV1)
'rectflux.Devise = CV1.DeviseN
'rectflux.Compte = lTOpe.EngagementCompte
'rectflux.Brut = lTflux.Intérêts
'Select Case lTflux.CodeOpération
'    Case "GA51", "Ga52", "Ga53": rectflux.Sens = "D"
 '   Case Else: rectflux.Sens = "C"
'End Select

'rectflux.AmjOpération = lTflux.AmjOpération
'rectflux.AmjValeur = lTflux.AmjValeur
'rectflux.Libellé = TFlux_Compta.Param_CodeOpération(lTflux.CodeOpération)
'rectflux.optAvis = "1"
'prtAvisX rectflux
'End Sub

End Sub
Public Sub prtGarantie_AvisA4()
Dim wSens As String * 1
Dim libTitre As String, libInfo1 As String, libInfo2 As String

prtSocMini XPrt.CurrentY, recTFlux.AmjOpération

XPrt.FontBold = False
XPrt.FontSize = 9
XPrt.CurrentX = prtMinX + 100
XPrt.CurrentY = prtMinY + 2100
XPrt.Print "N/Référence :";
XPrt.FontBold = True
XPrt.CurrentX = prtMinX + 1500
XPrt.Print Trim(recTope.RéférenceInterne);
XPrt.FontBold = False

If Trim(recTope.RéférenceExterne) <> "" Then
    XPrt.CurrentX = prtMinX + 100
    XPrt.CurrentY = prtMinY + 2400
    XPrt.Print "V/Référence :";
    XPrt.FontBold = True
    XPrt.CurrentX = prtMinX + 1500
    XPrt.Print Trim(recTope.RéférenceExterne);
    XPrt.FontBold = False
End If
Call CV_AttributS(recTope.Devise, CV1)

recCptInfoInit recCptInfo
recCptInfo.Method = "JoinL1"
recCptInfo.Société = SocId$
recCptInfo.Agence = SocId$
recCptInfo.Devise = CV1.DeviseN
recCptInfo.Numéro = recTope.EngagementCompte
recCptInfo.BiaTyp = "000"
recCptInfo.BiaNum = "00000"
recCptInfo.NuméroAncien = "00000000000"
If Not IsNull(srvCptInfoFind(recCptInfo)) Then
    Call MsgBox("prtGarantie_AvisA4 : compte d'engagement inconnu", vbCritical, "Impression")
    Exit Sub
End If

XPrt.CurrentY = 0

prtAdresse XPrt.CurrentY, recCptInfo

Call frmElpPrt.prtTrame(prtMinX, prtMinY + 5600, prtMaxX, prtMinY + 7700, " ", 245)

Select Case recTFlux.CodeOpération
    Case "GA99": wSens = "C"
                libTitre = "AVIS DE CREDIT / CREDIT ADVICE"
                libInfo1 = "Nous avons l'honneur de vous informer que nous créditons votre compte numéro :"
                libInfo2 = "We beg to inform you that we are crediting your account number :"

    Case Else: wSens = "D"
                libTitre = "AVIS DE DEBIT / DEBIT ADVICE"
                libInfo1 = "Nous avons l'honneur de vous informer que nous débitons votre compte numéro :"
                libInfo2 = "We beg to inform you that we are debiting your account number :"
End Select

XPrt.FontBold = True
XPrt.FontSize = 14
XPrt.CurrentX = prtMinX + 100
XPrt.CurrentY = prtMinY + 5000
XPrt.Print libTitre;

XPrt.FontBold = False
XPrt.FontSize = 9
XPrt.CurrentX = prtMinX + 100
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.Print libInfo1;
      
XPrt.FontBold = True
XPrt.CurrentX = prtMaxX - 4000
XPrt.Print Compte_Imp(recTope.EchéanceCompte);
XPrt.FontBold = False

XPrt.CurrentX = prtMinX + 100
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print libInfo2;

XPrt.FontBold = True
XPrt.CurrentX = prtMinX + 100
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "montant net de la commission / commission net amount : ";
X = CV1.DeviseIso & num_Display(recTFlux.Intérêts, 15, CV1.maxD, lX, X, "0")
XPrt.CurrentX = 10950 - XPrt.TextWidth(X)
XPrt.Print X;
    
XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
X = "(" & MontantEnLettres(recTFlux.Intérêts, CV1.DeviseLibellé) & ")"
XPrt.FontSize = 7
XPrt.CurrentX = 10950 - XPrt.TextWidth(X)
XPrt.Print X;

XPrt.FontSize = 8
XPrt.CurrentX = prtMinX + 100
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "date valeur / value date : " & dateImp(recTFlux.AmjValeur);

XPrt.CurrentY = prtMinY + 8500
XPrt.CurrentX = prtMinX + 100
XPrt.Print "Nature de l'opération / Object";
XPrt.CurrentX = prtMinX + 2500: XPrt.Print ":";
TFlux_Compta.param_Nature recTope.Nature
XPrt.CurrentX = prtMinX + 2600: XPrt.Print paramTFlux_Nature;
    
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight: XPrt.CurrentX = prtMinX + 100: XPrt.Print "Date d'émission / Date of issue";
XPrt.CurrentX = prtMinX + 2500: XPrt.Print ":";
XPrt.CurrentX = prtMinX + 2600: XPrt.Print dateImp(recTope.AmjDébut);

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight: XPrt.CurrentX = prtMinX + 100: XPrt.Print "Montant / Amount";
XPrt.CurrentX = prtMinX + 2500: XPrt.Print ":";
XPrt.FontBold = True
X = CV1.DeviseIso & num_Display(recTope.Capital, 15, CV1.maxD, lX, X, "0")
XPrt.CurrentX = prtMinX + 2600: XPrt.Print X;
XPrt.FontBold = False
    
If recTope.TauxMarge <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight: XPrt.CurrentX = prtMinX + 100: XPrt.Print "Commission taux / Rate";
    XPrt.CurrentX = prtMinX + 2500: XPrt.Print ":";
    XPrt.CurrentX = prtMinX + 2600: XPrt.Print Trim(Format(recTope.TauxMarge, "#0.#####")) & " % l'an";
End If

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight: XPrt.CurrentX = prtMinX + 100: XPrt.Print "Période / Period";
XPrt.CurrentX = prtMinX + 2500: XPrt.Print ":";
XPrt.CurrentX = prtMinX + 2600: XPrt.Print dateImp(recTFlux.AmjDébut) & " au " & dateImp(recTFlux.AmjFin);
    
XPrt.FontBold = True
XPrt.CurrentY = 12000: XPrt.CurrentX = prtMinX + 100
XPrt.Print "Veuillez agréer nos salutations distingués / Yours faithfully"
'XPrt.CurrentX = 10000
'XPrt.Print "Visa BIA";
    
XPrt.CurrentY = prtMaxY
XPrt.FontBold = False
prtSocMiniFin

End Sub


'---------------------------------------------------------
 Public Sub prtGarantieList_Monitor(Msg As String)
'---------------------------------------------------------
Dim K As Integer, K1 As Integer, K2 As Integer, Kmin As Integer, Kmax As Integer
Dim X As String

On Error GoTo prtError

Set XPrt = Printer
Height8_6 = frmElpPrt.prtHeightDelta(8, 6)
K1 = Val(mId$(Msg, 1, 6))
K2 = Val(mId$(Msg, 7, 6))

recElpTable_Init recElpTable
xElpTable = recElpTable
recElpTable.Method = "Seek="
recElpTable.Id = "BiaPgm"
prtTitleText = "Liste des Garanties"

prtLineNb = 1

frmElpPrt.Show vbModeless


prtOrientation = vbPRORLandscape
prtPgmName = "prtGarantieList"
prtTitleUsr = usrName

prtlineHeight = 300
prtHeaderHeight = 300

frmElpPrt.prtStdInit

recCompteInit recCompte
recCompte.Société = SocId$
recCompte.Agence = SocAgence$
recCompte.Devise = "001"
recCompte.BiaTyp = "000"
recCompte.BiaNum = "00"
recCompte.Method = "SeekL1"

prtGarantieList_Form
For K = K1 To K2
    recTope = P_arrTOpe(K)
    recElpTable.K1 = recTope.Nature
    tableElpTable_Read recElpTable
    
    prtGarantieList_Line
    
    DoEvents: If prtKillDoc Then frmElpPrt.Hide: Exit Sub

Next K
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
 Public Sub prtGarantieListEchéancier_Monitor(Msg As String)
'---------------------------------------------------------
Dim K As Integer, K1 As Integer, K2 As Integer, Kmin As Integer, Kmax As Integer
Dim X As String

On Error GoTo prtError

Set XPrt = Printer
Height8_6 = frmElpPrt.prtHeightDelta(8, 6)
K1 = Val(mId$(Msg, 1, 6))
K2 = Val(mId$(Msg, 7, 6))

recElpTable_Init recElpTable
xElpTable = recElpTable
recElpTable.Method = "Seek="
recElpTable.Id = "BiaPgm"
prtTitleText = "Echéancier des Garanties"

recTOpe_Init recTope

prtLineNb = 1

frmElpPrt.Show vbModeless


prtOrientation = vbPRORLandscape
prtPgmName = "prtGarantieListEchéancier"
prtTitleUsr = usrName

prtlineHeight = 300
prtHeaderHeight = 300

frmElpPrt.prtStdInit

recCompteInit recCompte
recCompte.Société = SocId$
recCompte.Agence = SocAgence$
recCompte.Devise = "001"
recCompte.BiaTyp = "000"
recCompte.BiaNum = "00"
recCompte.Method = "SeekL1"

prtGarantieListEchéancier_Form
For K = K1 To K2
    recTFlux = P_arrTFlux(K)
    recTope.IdRéférence = recTFlux.IdRéférence
    recTope.Method = "SeekP0"
    recTope.Application = paramTFlux_Service
    If IsNull(srvTOpe_Monitor(recTope)) Then
        recElpTable.K1 = recTope.Nature
        tableElpTable_Read recElpTable
        
        prtGarantieListechéancier_Line
    End If
    DoEvents: If prtKillDoc Then frmElpPrt.Hide: Exit Sub

Next K
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
Public Sub prtGarantieList_Form()
'---------------------------------------------------------
Dim X As String
prtCurrentY = XPrt.CurrentY
XPrt.FontSize = 8

XPrt.FontBold = True
XPrt.DrawWidth = 3

Call frmElpPrt.prtTrame(prtMinX, prtCurrentY, prtMaxX, prtCurrentY + prtlineHeight, "B", 250)
Call frmElpPrt.prtTrame(1300, prtMinY + prtHeaderHeight + 10, 2600, prtMaxY - 10, "", 250)
Call frmElpPrt.prtTrame(7400, prtMinY + prtHeaderHeight + 10, 9100, prtMaxY - 10, "", 250)

XPrt.DrawWidth = 1


'XPrt.Line (prtMinX, prtMaxY)-(prtMaxX, prtMaxY)

'---------------------------------------------------------

XPrt.CurrentY = prtCurrentY + 50
XPrt.CurrentX = prtMinX + 1300: XPrt.Print "Compte";

X = "Intitulé": XPrt.CurrentX = prtMinX + 3500 - XPrt.TextWidth(X)
XPrt.Print X;

X = "Montant": XPrt.CurrentX = prtMinX + 8500 - XPrt.TextWidth(X)
XPrt.Print X;

X = "%": XPrt.CurrentX = prtMinX + 9700 - XPrt.TextWidth(X)
XPrt.Print X;

X = "Période": XPrt.CurrentX = prtMinX + 11100 - XPrt.TextWidth(X)
XPrt.Print X;

X = "Réfèrences": XPrt.CurrentX = prtMinX + 15200 - XPrt.TextWidth(X)
XPrt.Print X;

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight

End Sub
'---------------------------------------------------------
Public Sub prtGarantieListEchéancier_Form()
'---------------------------------------------------------
Dim X As String
prtCurrentY = XPrt.CurrentY
XPrt.FontSize = 8

XPrt.FontBold = True
XPrt.DrawWidth = 3

Call frmElpPrt.prtTrame(prtMinX, prtCurrentY, prtMaxX, prtCurrentY + prtlineHeight, "B", 250)
Call frmElpPrt.prtTrame(1300, prtMinY + prtHeaderHeight + 10, 2600, prtMaxY - 10, "", 250)
Call frmElpPrt.prtTrame(7400, prtMinY + prtHeaderHeight + 10, 9100, prtMaxY - 10, "", 250)

XPrt.DrawWidth = 1


'XPrt.Line (prtMinX, prtMaxY)-(prtMaxX, prtMaxY)

'---------------------------------------------------------

XPrt.CurrentY = prtCurrentY + 50
XPrt.CurrentX = prtMinX + 1300: XPrt.Print "Compte";

X = "Intitulé": XPrt.CurrentX = prtMinX + 3500 - XPrt.TextWidth(X)
XPrt.Print X;

X = "Montant": XPrt.CurrentX = prtMinX + 8500 - XPrt.TextWidth(X)
XPrt.Print X;

X = "%": XPrt.CurrentX = prtMinX + 9700 - XPrt.TextWidth(X)
XPrt.Print X;

X = "Période": XPrt.CurrentX = prtMinX + 11100 - XPrt.TextWidth(X)
XPrt.Print X;

X = "Réfèrences": XPrt.CurrentX = prtMinX + 15200 - XPrt.TextWidth(X)
XPrt.Print X;

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight

End Sub


Public Sub prtGarantieList_Line()
'------------------------------------------------------ligne 1---
Dim iReturn As Integer

If XPrt.CurrentY + prtlineHeight * 5 > prtMaxY Then
    frmElpPrt.prtNewPage
    prtGarantieList_Form
End If

XPrt.FontBold = True
'XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY, prtMaxX, XPrt.CurrentY + prtlineHeight, " ", 245)


XPrt.CurrentX = prtMinX + 150: XPrt.Print param_Statut(recTope.Statut & recTope.StatutPlus);

TFlux_Compta.param_Nature recTope.Nature
XPrt.FontBold = False
XPrt.FontSize = 8
XPrt.CurrentX = prtMinX + 2450: XPrt.Print paramTFlux_Nature;

XPrt.CurrentX = prtMinX + 14300: XPrt.Print (recTope.RéférenceInterne);
'_______________________________________________________________ligne 2
XPrt.FontSize = 6
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentY = XPrt.CurrentY + Height8_6
XPrt.CurrentX = prtMinX + 150: XPrt.Print "Donneur d'ordre :";
XPrt.CurrentY = XPrt.CurrentY - Height8_6
XPrt.FontSize = 8
XPrt.CurrentX = prtMinX + 1200: XPrt.Print Compte_Imp(recTope.EngagementCompte);
Call CV_AttributS(recTope.Devise, CV1)
recCompte.Devise = CV1.DeviseN
recCompte.Intitulé = "": recCompte.Numéro = recTope.EngagementCompte:
mdbCptP0_Find recCompte

XPrt.FontSize = 6
XPrt.CurrentY = XPrt.CurrentY + Height8_6
XPrt.CurrentX = prtMinX + 2450: XPrt.Print Trim(recCompte.Intitulé);
XPrt.CurrentX = prtMinX + 8450: XPrt.Print recTope.Devise;
XPrt.CurrentX = prtMinX + 10500: XPrt.Print "du  : " & dateImp(recTope.AmjDébut);
XPrt.CurrentX = prtMinX + 11500: XPrt.Print "  au : " & dateImp(recTope.AmjFin);
XPrt.CurrentX = prtMinX + 13000
If recTope.PréavisNbj = 999 Then
    XPrt.Print "attente de la main levée";
Else
    XPrt.Print "délai courrier :" & recTope.PréavisNbj;
End If

XPrt.CurrentY = XPrt.CurrentY - Height8_6
XPrt.FontSize = 8
XPrt.FontBold = True
X = Format$(recTope.Capital, "## ### ### ### ### ##0.00")
XPrt.CurrentX = prtMinX + 8400 - XPrt.TextWidth(X): XPrt.Print X;
XPrt.FontBold = False

XPrt.CurrentX = prtMinX + 14300: XPrt.Print (recTope.RéférenceExterne);


'-------------------------------------------------ligne 3

XPrt.FontSize = 6
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentY = XPrt.CurrentY + Height8_6
XPrt.CurrentX = prtMinX + 150: XPrt.Print "Commission :";
XPrt.CurrentY = XPrt.CurrentY - Height8_6
XPrt.FontSize = 8
XPrt.CurrentX = prtMinX + 1200: XPrt.Print Compte_Imp(recTope.EchéanceCompte);
recCompte.Intitulé = "": recCompte.Numéro = recTope.EngagementCompte:
mdbCptP0_Find recCompte
XPrt.FontSize = 6
XPrt.CurrentY = XPrt.CurrentY + Height8_6
XPrt.CurrentX = prtMinX + 2450: XPrt.Print Trim(recCompte.Intitulé);
If recTope.TauxMarge <> 0 Then
    XPrt.CurrentX = prtMinX + 9500: XPrt.Print Trim(Format(recTope.TauxMarge, "##.#####")) & " %";
    XPrt.CurrentX = prtMinX + 10500
    XPrt.Print "com : " & dateImp(recTope.AmjEchéance1) & ".... " & Trim(recPériodicité_Libellé(recTope.Périodicité)) & " / " & param_AmjEchéanceS(recTope.AmjEchéanceS);
End If
'--------------------------------------------------ligne4
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinX + 150: XPrt.Print "Garantie :";
XPrt.CurrentY = XPrt.CurrentY - Height8_6
XPrt.FontSize = 8
XPrt.CurrentX = prtMinX + 1200: XPrt.Print Compte_Imp(recTope.EngagementCorrCompte);
recCompte.Intitulé = "": recCompte.Numéro = recTope.EngagementCorrCompte:
mdbCptP0_Find recCompte
XPrt.FontSize = 6
XPrt.CurrentY = XPrt.CurrentY + Height8_6
XPrt.CurrentX = prtMinX + 2450: XPrt.Print Trim(recCompte.Intitulé);

If recTope.Frais <> 0 Then
    X = Trim(Format$(recTope.Frais, "## ### ### ### ### ##0.00"))
    XPrt.CurrentX = prtMinX + 10500: XPrt.Print "Commission flat :  " & X;
End If
XPrt.CurrentX = prtMinX + 14300: XPrt.Print "(" & Trim(recTope.IdRéférence) & ")";
XPrt.CurrentY = XPrt.CurrentY - Height8_6
XPrt.FontSize = 8

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY)
XPrt.CurrentY = XPrt.CurrentY + 50
End Sub

Public Sub prtGarantieListechéancier_Line()
'------------------------------------------------------ligne 1---
Dim iReturn As Integer

If XPrt.CurrentY + prtlineHeight * 5 > prtMaxY Then
    frmElpPrt.prtNewPage
    prtGarantieList_Form
End If

XPrt.FontBold = True
'XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY, prtMaxX, XPrt.CurrentY + prtlineHeight, " ", 245)


XPrt.CurrentX = prtMinX + 1200: XPrt.Print param_Statut(recTope.Statut & recTope.StatutPlus);

TFlux_Compta.param_Nature recTope.Nature
XPrt.FontBold = False
XPrt.FontSize = 8
XPrt.CurrentX = prtMinX + 2450: XPrt.Print paramTFlux_Nature;
XPrt.CurrentX = prtMinX + 10500: XPrt.Print TFlux_Compta.Param_CodeOpération(recTFlux.CodeOpération);

XPrt.CurrentX = prtMinX + 14300: XPrt.Print (recTope.RéférenceInterne);
'_______________________________________________________________ligne 2
XPrt.FontSize = 6
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentY = XPrt.CurrentY + Height8_6
XPrt.CurrentX = prtMinX + 150: XPrt.Print "Donneur d'ordre :";
XPrt.CurrentY = XPrt.CurrentY - Height8_6
XPrt.FontSize = 8
XPrt.CurrentX = prtMinX + 1200: XPrt.Print Compte_Imp(recTope.EngagementCompte);
Call CV_AttributS(recTope.Devise, CV1)
recCompte.Devise = CV1.DeviseN
recCompte.Intitulé = "": recCompte.Numéro = recTope.EngagementCompte:
mdbCptP0_Find recCompte

XPrt.FontSize = 6
XPrt.CurrentY = XPrt.CurrentY + Height8_6
XPrt.CurrentX = prtMinX + 2450: XPrt.Print Trim(recCompte.Intitulé);
XPrt.CurrentX = prtMinX + 8450: XPrt.Print recTope.Devise;
'XPrt.FontSize = 6
'XPrt.CurrentY = XPrt.CurrentY + Height8_6
XPrt.CurrentX = prtMinX + 6800: XPrt.Print recTope.Devise;
XPrt.FontBold = True
X = Format$(recTope.Capital, "## ### ### ### ### ##0.00")
XPrt.CurrentX = prtMinX + 6700 - XPrt.TextWidth(X): XPrt.Print X;
XPrt.FontBold = False
'XPrt.CurrentY = XPrt.CurrentY - Height8_6
'XPrt.FontSize = 8
XPrt.CurrentY = XPrt.CurrentY - Height8_6
XPrt.FontSize = 8
XPrt.FontBold = True

X = Format$(recTFlux.Capital + recTFlux.Intérêts, "#### ### ### ### ##0.00")
XPrt.CurrentX = prtMinX + 8400 - XPrt.TextWidth(X): XPrt.Print X;
XPrt.FontBold = False

XPrt.CurrentX = prtMinX + 9500: XPrt.Print Trim(Format(recTFlux.Taux, "##.#####")) & " %";
Call frmElpPrt.prtTrame(prtMinX + 10500, XPrt.CurrentY - 50, prtMinX + 12200, XPrt.CurrentY + prtlineHeight - 50, "", 230)
XPrt.CurrentX = prtMinX + 10500: XPrt.Print "Valeur : " & dateImp(recTFlux.AmjValeur);

XPrt.CurrentX = prtMinX + 14300: XPrt.Print (recTope.RéférenceExterne);


'-------------------------------------------------ligne 3

XPrt.FontSize = 6
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentY = XPrt.CurrentY + Height8_6
XPrt.CurrentX = prtMinX + 150: XPrt.Print "Commission :";
XPrt.CurrentY = XPrt.CurrentY - Height8_6
XPrt.FontSize = 8
XPrt.CurrentX = prtMinX + 1200: XPrt.Print Compte_Imp(recTope.EchéanceCompte);
recCompte.Intitulé = "": recCompte.Numéro = recTope.EchéanceCompte:
mdbCptP0_Find recCompte
XPrt.FontSize = 6
XPrt.CurrentY = XPrt.CurrentY + Height8_6
XPrt.CurrentX = prtMinX + 2450: XPrt.Print Trim(recCompte.Intitulé);
XPrt.CurrentY = XPrt.CurrentY - Height8_6
XPrt.CurrentX = prtMinX + 10500: XPrt.Print "du : " & dateImp(recTFlux.AmjDébut);
XPrt.CurrentX = prtMinX + 11500: XPrt.Print "  au : " & dateImp(recTFlux.AmjFin);

XPrt.FontSize = 6
XPrt.CurrentY = XPrt.CurrentY + Height8_6
'--------------------------------------------------ligne4
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinX + 150: XPrt.Print "Garantie :";
XPrt.CurrentY = XPrt.CurrentY - Height8_6
XPrt.FontSize = 8
XPrt.CurrentX = prtMinX + 1200: XPrt.Print Compte_Imp(recTope.EngagementCorrCompte);
recCompte.Intitulé = "": recCompte.Numéro = recTope.EngagementCorrCompte:
mdbCptP0_Find recCompte
XPrt.FontSize = 6
XPrt.CurrentY = XPrt.CurrentY + Height8_6
XPrt.CurrentX = prtMinX + 2450: XPrt.Print Trim(recCompte.Intitulé);
'XPrt.CurrentY = XPrt.CurrentY - Height8_6

'XPrt.CurrentY = XPrt.CurrentY + Height8_6
XPrt.CurrentX = prtMinX + 14300: XPrt.Print "(" & Trim(recTFlux.IdRéférence) & "." & Trim(recTFlux.IdSéquence); ")";
XPrt.CurrentY = XPrt.CurrentY - Height8_6
XPrt.FontSize = 8

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY)
XPrt.CurrentY = XPrt.CurrentY + 50
End Sub


'---------------------------------------------------------
 Public Sub prtGarantie_Monitor(Msg As String)
'---------------------------------------------------------
Dim K As Integer, K1 As Integer, K2 As Integer, Kmin As Integer, Kmax As Integer
Dim X As String
Dim blnSaut As Boolean, blnValidation As Boolean

On Error GoTo prtError

Set XPrt = Printer
Height8_6 = frmElpPrt.prtHeightDelta(8, 6)
K1 = Val(mId$(Msg, 1, 6))
K2 = Val(mId$(Msg, 7, 6))

blnTableauAmortissement = False: blnSaut = True
blnValidation = False
blnSaut = False

Select Case mId$(Msg, 13, 1)
    Case "à":  prtTitleText = "Garantie à Valider : ": blnValidation = True
                Kmin = 10: Kmax = K2 - 10: blnSaut = True
    Case Else:   prtTitleText = "Garantie : "
                blnTableauAmortissement = True
                Kmin = K2 + 1: Kmax = 0
End Select
If Kmax <= 0 Then Kmax = K2

recElpTable_Init recElpTable
xElpTable = recElpTable
recElpTable.Method = "Seek="
recElpTable.Id = "BiaPgm"
recElpTable.K1 = recTope.Nature
tableElpTable_Read recElpTable
prtTitleText = prtTitleText & Trim(recElpTable.Name) & recTope.RéférenceInterne

prtLineNb = 1

frmElpPrt.Show vbModeless


prtOrientation = vbPRORLandscape 'PRORPortrait
prtPgmName = "prtGarantie"
prtTitleUsr = usrName

prtlineHeight = 300
prtHeaderHeight = 300
    
frmElpPrt.prtStdInit

recCompteInit recCompte
recCompte.Société = SocId$
recCompte.Agence = SocAgence$
recCompte.Devise = CV1.DeviseN
recCompte.BiaTyp = "000"
recCompte.BiaNum = "00"
recCompte.Method = "SeekL1"

prtGarantie_Header

prtGarantie_Form
CapitalRestantDû = 0
totalCommission = 0
For K = K1 To K2
    recTFlux = arrTFlux(K)
    If recTFlux.Statut <> "A" Then
        Select Case recTFlux.CodeOpération
            Case "GA01", "GA11": CapitalRestantDû = CapitalRestantDû - recTFlux.Capital
            Case Else: CapitalRestantDû = CapitalRestantDû + recTFlux.Capital
        End Select
        totalCommission = totalCommission + recTFlux.Intérêts
    End If
    
        If K <= Kmin Or K >= Kmax Then
            prtGarantie_Line
        Else
            If blnSaut = True Then
                blnSaut = False
                XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
            End If
        End If
            
    DoEvents: If prtKillDoc Then frmElpPrt.Hide: Exit Sub

Next K
XPrt.DrawWidth = 5
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight

Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY - 50, prtMaxX, XPrt.CurrentY + prtlineHeight - 50, " ", 240)

X = Format$(totalCommission, "#### ### ### ### ##0.00")
XPrt.CurrentX = prtMinX + 6500 - XPrt.TextWidth(X)
XPrt.Print X;

X = Format$(CapitalRestantDû, "#### ### ### ### ##0.00")
XPrt.CurrentX = prtMinX + 5000 - XPrt.TextWidth(X)
XPrt.Print X;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY)


If CapitalRestantDû <> 0 Then
    XPrt.FontSize = 14: XPrt.CurrentX = prtMinX: XPrt.Print "CapitalRestantDû <> 0"
End If

If blnValidation Then prtCompta_Validation "VALIDATION de la Garantie : " & recTope.RéférenceInterne

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
Public Sub prtGarantie_Form()
'---------------------------------------------------------
Dim X As String
prtCurrentY = XPrt.CurrentY
XPrt.FontSize = 10

XPrt.FontBold = True
XPrt.DrawWidth = 3

Call frmElpPrt.prtTrame(prtMinX, prtCurrentY - 50, prtMaxX, prtCurrentY + prtlineHeight - 50, " ", 240)
XPrt.DrawWidth = 1

'XPrt.Line (prtMinX, prtCurrentY)-(prtMinX, prtMaxY)
'XPrt.Line (prtMinX + 1700, prtCurrentY)-(prtMinX + 1700, prtMaxY)
'XPrt.Line (prtMinX + 4000, prtCurrentY)-(prtMinX + 4000, prtMaxY)
'XPrt.Line (prtMinX + 6300, prtCurrentY)-(prtMinX + 6300, prtMaxY)
'XPrt.Line (prtMinX + 8600, prtCurrentY)-(prtMinX + 8600, prtMaxY)
'XPrt.Line (prtMaxX, prtCurrentY)-(prtMaxX, prtMaxY)
'XPrt.Line (prtMinX, prtMaxY)-(prtMaxX, prtMaxY)

'---------------------------------------------------------
XPrt.FontSize = 6

XPrt.CurrentY = prtCurrentY + 50
XPrt.CurrentX = prtMinX + 250: XPrt.Print "Echéance";

XPrt.CurrentX = prtMinX + 1300: XPrt.Print "Opération";

X = "Montant": XPrt.CurrentX = prtMinX + 5000 - XPrt.TextWidth(X)
XPrt.Print X;

X = "Commission": XPrt.CurrentX = prtMinX + 6500 - XPrt.TextWidth(X)
XPrt.Print X;

XPrt.CurrentX = prtMinX + 6700: XPrt.Print "Statut";
XPrt.CurrentX = prtMinX + 9200: XPrt.Print "Période";
XPrt.CurrentX = prtMinX + 11000: XPrt.Print "Date valeur";
XPrt.CurrentX = prtMinX + 12000: XPrt.Print "Pièce";
XPrt.CurrentX = prtMinX + 13000: XPrt.Print "Mise à jour";
XPrt.FontSize = 8
prtLineNb = 0
End Sub

'---------------------------------------------------------
Public Sub prtGarantie_Line()
'---------------------------------------------------------
Dim blnHeight8_6 As Boolean

If XPrt.CurrentY + prtlineHeight * 1.5 > prtMaxY Then
    frmElpPrt.prtNewPage
    XPrt.CurrentY = prtMinX + prtlineHeight * 3
    prtGarantie_Form
End If

XPrt.FontBold = False
'------------------------------------------ligne 1--------------
prtLineNb = prtLineNb + 1
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
If prtLineNb > 2 Then
    
    Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY - 50, prtMaxX, XPrt.CurrentY + prtlineHeight - 50, " ", 250)
    If prtLineNb = 4 Then prtLineNb = 0
End If


If recTFlux.Statut = "A" Then
    blnHeight8_6 = True
    XPrt.FontSize = 6
    XPrt.CurrentY = XPrt.CurrentY + Height8_6
Else
    blnHeight8_6 = False
End If

XPrt.CurrentX = prtMinX + 50: XPrt.Print dateImp(recTFlux.AmjEchéanceTrt);
XPrt.CurrentX = prtMinX + 1300: XPrt.Print TFlux_Compta.Param_CodeOpération(recTFlux.CodeOpération);
    
If recTFlux.Statut <> "A" Then

    If recTFlux.Capital <> 0 Then
        X = Format$(recTFlux.Capital, "#### ### ### ### ##0.00")
        XPrt.CurrentX = prtMinX + 5000 - XPrt.TextWidth(X)
        XPrt.Print X;
    End If
    If recTFlux.Intérêts <> 0 Then
        X = Format$(recTFlux.Intérêts, "#### ### ### ### ##0.00")
        XPrt.CurrentX = prtMinX + 6500 - XPrt.TextWidth(X)
        XPrt.Print X;
    End If
End If

If Not blnHeight8_6 Then
    XPrt.FontSize = 6
    XPrt.CurrentY = XPrt.CurrentY + Height8_6
End If

XPrt.CurrentX = prtMinX + 6700: XPrt.Print param_Statut(recTFlux.Statut & recTFlux.StatutPlus);

XPrt.CurrentX = prtMinX + 8500: XPrt.Print dateImp(recTFlux.AmjDébut);
XPrt.CurrentX = prtMinX + 9500: XPrt.Print dateImp(recTFlux.AmjFin);
XPrt.CurrentX = prtMinX + 11000: XPrt.Print dateImp(recTFlux.AmjValeur);
If recTFlux.CptMvtPièce <> 0 Then XPrt.CurrentX = prtMinX + 12000: XPrt.Print Format$(recTFlux.CptMvtPièce, "#####") & "." & Format$(recTFlux.CptMvtLigne, "####");

XPrt.CurrentX = prtMinX + 13000: XPrt.Print dateImp(recTFlux.CptMvtAMJ);
XPrt.CurrentX = prtMinX + 14000: XPrt.Print timeImp(recTFlux.CptMvtHMS);
XPrt.CurrentX = prtMinX + 14800: XPrt.Print recTFlux.CptMvtUsr;

XPrt.CurrentY = XPrt.CurrentY - Height8_6
XPrt.FontSize = 8

End Sub






Public Sub prtGarantie_Header()
Dim strdev As String
Call CV_AttributS(recTope.Devise, CV1)
strdev = " " & CV1.DeviseLibellé
XPrt.FontSize = 6
XPrt.CurrentY = prtMinY + prtlineHeight: XPrt.CurrentX = prtMinX: XPrt.Print "Donneur d'ordre ";
XPrt.CurrentX = prtMinX + 1300: XPrt.Print ":";
'XPrt.FontSize = 8
XPrt.CurrentX = prtMinX + 1400: XPrt.Print Compte_Imp(recTope.EngagementCompte);
XPrt.FontSize = 6
recCompte.Intitulé = "": recCompte.Numéro = recTope.EngagementCompte:
mdbCptP0_Find recCompte
prtCurrentX = XPrt.CurrentX
XPrt.Print " " & Trim(recCompte.Intitulé);

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight - 100
If recCompte.TypeGA = "A" Then
    XPrt.CurrentX = prtCurrentX: XPrt.Print Trim(DicLib(13, recCompte.BiaTyp));
End If

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight - 100: XPrt.CurrentX = prtMinX: XPrt.Print "Montant du Garantie ";
XPrt.CurrentX = prtMinX + 1300: XPrt.Print ":";
'XPrt.FontSize = 8
XPrt.FontBold = True
XPrt.CurrentX = prtMinX + 1400: XPrt.Print Trim(Format(recTope.Capital, "### ### ### ###.00")) & " " & recTope.Devise; strdev;
XPrt.FontBold = False
XPrt.FontSize = 6
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight: XPrt.CurrentX = prtMinX: XPrt.Print "Date d'emission ";
XPrt.CurrentX = prtMinX + 1300: XPrt.Print ":";
'XPrt.FontSize = 8
XPrt.CurrentX = prtMinX + 1400: XPrt.Print dateImp(recTope.AmjDébut);
XPrt.FontSize = 6
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight: XPrt.CurrentX = prtMinX: XPrt.Print "Date de validité";
XPrt.CurrentX = prtMinX + 1300: XPrt.Print ":";
'XPrt.FontSize = 8
XPrt.CurrentX = prtMinX + 1400: XPrt.Print dateImp(recTope.AmjFin);
XPrt.FontSize = 6
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight: XPrt.CurrentX = prtMinX: XPrt.Print "délai courrier (jours)";
XPrt.CurrentX = prtMinX + 1300: XPrt.Print ":";
'XPrt.FontSize = 8
XPrt.CurrentX = prtMinX + 1400:
If recTope.PréavisNbj = 999 Then
    XPrt.Print "attente de la main levée";
Else
    XPrt.Print recTope.PréavisNbj;
End If

XPrt.CurrentY = prtMinY + prtlineHeight

XPrt.FontSize = 6
 XPrt.CurrentX = prtMinX + 5500: XPrt.Print "Nature";
XPrt.CurrentX = prtMinX + 6900: XPrt.Print ":";
'XPrt.FontSize = 8
TFlux_Compta.param_Nature recTope.Nature
XPrt.FontBold = True
XPrt.CurrentX = prtMinX + 7000: XPrt.Print paramTFlux_Nature;
XPrt.FontBold = False


XPrt.FontSize = 6
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight: XPrt.CurrentX = prtMinX + 5500: XPrt.Print "Commission à prélever ";
XPrt.CurrentX = prtMinX + 6900: XPrt.Print ":";
'XPrt.FontSize = 8
XPrt.CurrentX = prtMinX + 7000: XPrt.Print Compte_Imp(recTope.EchéanceCompte);
XPrt.FontSize = 6
recCompte.Intitulé = "": recCompte.Numéro = recTope.EchéanceCompte:
mdbCptP0_Find recCompte
prtCurrentX = XPrt.CurrentX
XPrt.Print " " & Trim(recCompte.Intitulé);

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight - 100
If recCompte.TypeGA = "A" Then
    XPrt.CurrentX = prtCurrentX: XPrt.Print Trim(DicLib(13, recCompte.BiaTyp));
End If

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight - 100
XPrt.CurrentX = prtMinX + 6900: XPrt.Print ":";

XPrt.CurrentX = prtMinX + 5500
If Trim(recTope.TauxRéférence) = "Montant" Then
    XPrt.Print "commission fixe";
    XPrt.FontBold = True
    XPrt.CurrentX = prtMinX + 7000: XPrt.Print Trim(Format(recTope.TauxMarge, "### ###.00"));
Else
    XPrt.Print "Taux de commission";
    XPrt.FontBold = True
    XPrt.CurrentX = prtMinX + 7000: XPrt.Print Trim(Format(recTope.TauxMarge, "##.#####")) & " %";
End If

XPrt.FontBold = False
XPrt.FontSize = 6
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight: XPrt.CurrentX = prtMinX + 5500: XPrt.Print "Première perception ";
XPrt.CurrentX = prtMinX + 6900: XPrt.Print ":";
'XPrt.FontSize = 8
XPrt.CurrentX = prtMinX + 7000: XPrt.Print dateImp(recTope.AmjEchéance1) & ".... " & Trim(recPériodicité_Libellé(recTope.Périodicité)) & " / " & param_AmjEchéanceS(recTope.AmjEchéanceS);
XPrt.FontSize = 6
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight: XPrt.CurrentX = prtMinX + 5500: XPrt.Print "Commission flat";
XPrt.CurrentX = prtMinX + 6900: XPrt.Print ":";
'XPrt.FontSize = 8
If recTope.Frais <> 0 Then XPrt.CurrentX = prtMinX + 7000: XPrt.Print Trim(Format(recTope.Frais, "### ### ### ###.00")) & strdev;

XPrt.CurrentY = prtMinY + prtlineHeight
XPrt.FontSize = 6
XPrt.CurrentX = prtMinX + 11000: XPrt.Print "Contrepartie ";
XPrt.CurrentX = prtMinX + 12400: XPrt.Print ":";
'XPrt.FontSize = 8
XPrt.CurrentX = prtMinX + 12500: XPrt.Print Compte_Imp(recTope.EngagementCorrCompte);
'XPrt.FontSize = 6
recCompte.Intitulé = "": recCompte.Numéro = recTope.EngagementCorrCompte:
mdbCptP0_Find recCompte
XPrt.Print " " & Trim(recCompte.Intitulé);

XPrt.FontSize = 6
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight: XPrt.CurrentX = prtMinX + 11000: XPrt.Print "Référence BIA";
XPrt.CurrentX = prtMinX + 12400: XPrt.Print ":";
'XPrt.FontSize = 8
XPrt.CurrentX = prtMinX + 12500: XPrt.Print recTope.RéférenceInterne;
XPrt.FontSize = 6
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight: XPrt.CurrentX = prtMinX + 11000: XPrt.Print "Référence du contrat";
XPrt.CurrentX = prtMinX + 12400: XPrt.Print ":";
'XPrt.FontSize = 8
XPrt.CurrentX = prtMinX + 12500: XPrt.Print recTope.RéférenceExterne;
XPrt.FontSize = 6
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight: XPrt.CurrentX = prtMinX + 11000: XPrt.Print "Référence informatique";
XPrt.CurrentX = prtMinX + 12400: XPrt.Print ":";
'XPrt.FontSize = 8
XPrt.CurrentX = prtMinX + 12500: XPrt.Print recTope.IdRéférence;
XPrt.FontSize = 6
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight: XPrt.CurrentX = prtMinX + 11000: XPrt.Print "Saisie";
XPrt.CurrentX = prtMinX + 12400: XPrt.Print ":";
'XPrt.FontSize = 8
XPrt.CurrentX = prtMinX + 12500: XPrt.Print dateImp(recTope.MajAMJ) & "  " & timeImp(recTope.MajHMS) & "   " & recTope.MajUsr;
XPrt.FontSize = 6
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight: XPrt.CurrentX = prtMinX + 11000: XPrt.Print "Validation";
XPrt.CurrentX = prtMinX + 12400: XPrt.Print ":";
'XPrt.FontSize = 8
XPrt.CurrentX = prtMinX + 12500: XPrt.Print dateImp(recTope.valAMJ) & "  " & timeImp(recTope.ValHMS) & "   " & recTope.ValUsr;

XPrt.CurrentY = prtMinY + prtlineHeight * 7
End Sub


