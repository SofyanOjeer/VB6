Attribute VB_Name = "prtPrêt"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Private recCompte As typeCompte
Dim x As String, I As Integer, Height8_6 As Integer

Private recCptInfo As typeCptInfo
Public P_arrTFlux() As typeTFlux, recTFlux As typeTFlux

Public recTope As typeTOpe
Public CV1 As typeCV

Dim CapitalRestantDû As Currency, totalMensualité As Currency, totalIntérêts As Currency, totalAmortissement As Currency
Dim curFrais As Currency, curIntérêtsIntermédiaires As Currency
Dim blnTableauAmortissement As Boolean

Public P_arrTOpe() As typeTOpe
Dim prtLineNb As Integer
'---------------------------------------------------------
 Public Sub prtPrêtListEchéancier_Monitor(Msg As String)
'---------------------------------------------------------
Dim K As Integer, K1 As Integer, K2 As Integer, Kmin As Integer, Kmax As Integer
Dim x As String

On Error GoTo prtError

Set XPrt = Printer
Height8_6 = frmElpPrt.prtHeightDelta(8, 6)
K1 = Val(mId$(Msg, 1, 6))
K2 = Val(mId$(Msg, 7, 6))

recElpTable_Init recElpTable
xElpTable = recElpTable
recElpTable.Method = "Seek="
recElpTable.Id = "BiaPgm"
prtTitleText = "Echéancier des Prêts"

recTOpe_Init recTope

prtLineNb = 1

frmElpPrt.Show vbModeless


prtOrientation = vbPRORLandscape
prtPgmName = "prtPrêtListEchéancier"
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

prtPrêtListEchéancier_Form
For K = K1 To K2
    recTFlux = P_arrTFlux(K)
    recTope.IdRéférence = recTFlux.IdRéférence
    recTope.Method = "SeekP0"
    recTope.Application = paramTFlux_Service
    If IsNull(srvTOpe_Monitor(recTope)) Then
        recElpTable.K1 = recTope.Nature
        tableElpTable_Read recElpTable
        
        prtPrêtListechéancier_Line
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

Public Sub prtPrêtListechéancier_Line()
'------------------------------------------------------ligne 1---
Dim iReturn As Integer

If XPrt.CurrentY + prtlineHeight * 5 > prtMaxY Then
    frmElpPrt.prtNewPage
    prtPrêtList_Form
End If

XPrt.FontBold = True
'XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY, prtMaxX, XPrt.CurrentY + prtlineHeight, " ", 245)


XPrt.CurrentX = prtMinX + 1200: XPrt.Print param_Statut(recTope.Statut & recTope.StatutPlus);

TFlux_Compta.Param_Nature recTope.Nature
XPrt.FontBold = False
XPrt.FontSize = 8
XPrt.CurrentX = prtMinX + 2450: XPrt.Print paramTFlux_Nature;
XPrt.FontSize = 6
XPrt.CurrentY = XPrt.CurrentY + Height8_6
XPrt.CurrentX = prtMinX + 6800: XPrt.Print recTope.Devise;
XPrt.FontBold = True
x = Format$(recTope.Capital, "## ### ### ### ### ##0.00")
XPrt.CurrentX = prtMinX + 6700 - XPrt.TextWidth(x): XPrt.Print x;
XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY - Height8_6
XPrt.FontSize = 8
XPrt.CurrentX = prtMinX + 10500: XPrt.Print TFlux_Compta.Param_CodeOpération(recTFlux.CodeOpération);

XPrt.CurrentX = prtMinX + 14300: XPrt.Print (recTope.RéférenceInterne);
'_______________________________________________________________ligne 2
XPrt.FontSize = 6
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentY = XPrt.CurrentY + Height8_6
XPrt.CurrentX = prtMinX + 150: XPrt.Print "Prêt :";
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
XPrt.CurrentY = XPrt.CurrentY - Height8_6
XPrt.FontSize = 8
XPrt.FontBold = True

x = Format$(recTFlux.Capital + recTFlux.Intérêts, "#### ### ### ### ##0.00")
XPrt.CurrentX = prtMinX + 8400 - XPrt.TextWidth(x): XPrt.Print x;
XPrt.FontBold = False

XPrt.CurrentX = prtMinX + 9500: XPrt.Print Trim(Format(recTFlux.Taux, "##.#####")) & " %";
Call frmElpPrt.prtTrame(prtMinX + 10500, XPrt.CurrentY - 50, prtMinX + 12200, XPrt.CurrentY + prtlineHeight - 50, "", 230)

XPrt.CurrentX = prtMinX + 10500: XPrt.Print "Valeur : " & dateImp(recTFlux.AmjValeur);

XPrt.CurrentX = prtMinX + 14300: XPrt.Print (recTope.RéférenceExterne);


'-------------------------------------------------ligne 3

XPrt.FontSize = 6
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentY = XPrt.CurrentY + Height8_6
XPrt.CurrentX = prtMinX + 150: XPrt.Print "Prélèvement :";
XPrt.CurrentY = XPrt.CurrentY - Height8_6
XPrt.FontSize = 8
XPrt.CurrentX = prtMinX + 1200: XPrt.Print Compte_Imp(recTope.EchéanceCompte);
recCompte.Intitulé = "": recCompte.Numéro = recTope.EngagementCompte:
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
XPrt.CurrentX = prtMinX + 150: XPrt.Print "Déblocage :";
XPrt.CurrentY = XPrt.CurrentY - Height8_6
XPrt.FontSize = 8
XPrt.CurrentX = prtMinX + 1200: XPrt.Print Compte_Imp(recTope.EngagementCorrCompte);
recCompte.Intitulé = "": recCompte.Numéro = recTope.EngagementCorrCompte:
mdbCptP0_Find recCompte
XPrt.FontSize = 6
XPrt.CurrentY = XPrt.CurrentY + Height8_6
XPrt.CurrentX = prtMinX + 2450: XPrt.Print Trim(recCompte.Intitulé);
XPrt.CurrentY = XPrt.CurrentY - Height8_6

XPrt.CurrentY = XPrt.CurrentY + Height8_6
XPrt.CurrentX = prtMinX + 14300: XPrt.Print "(" & Trim(recTope.IdRéférence) & ")";
XPrt.CurrentY = XPrt.CurrentY - Height8_6
XPrt.FontSize = 8

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY)
XPrt.CurrentY = XPrt.CurrentY + 50
End Sub


'---------------------------------------------------------
Public Sub prtPrêtListEchéancier_Form()
'---------------------------------------------------------
Dim x As String
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

x = "Intitulé": XPrt.CurrentX = prtMinX + 3500 - XPrt.TextWidth(x)
XPrt.Print x;

x = "Montant": XPrt.CurrentX = prtMinX + 8500 - XPrt.TextWidth(x)
XPrt.Print x;

x = "%": XPrt.CurrentX = prtMinX + 9700 - XPrt.TextWidth(x)
XPrt.Print x;

x = "Période": XPrt.CurrentX = prtMinX + 11100 - XPrt.TextWidth(x)
XPrt.Print x;

x = "Réfèrences": XPrt.CurrentX = prtMinX + 15200 - XPrt.TextWidth(x)
XPrt.Print x;

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight

End Sub



'---------------------------------------------------------
 Public Sub prtPrêtList_Monitor(Msg As String)
'---------------------------------------------------------
Dim K As Integer, K1 As Integer, K2 As Integer, Kmin As Integer, Kmax As Integer
Dim x As String

On Error GoTo prtError

Set XPrt = Printer
Height8_6 = frmElpPrt.prtHeightDelta(8, 6)
K1 = Val(mId$(Msg, 1, 6))
K2 = Val(mId$(Msg, 7, 6))

recElpTable_Init recElpTable
xElpTable = recElpTable
recElpTable.Method = "Seek="
recElpTable.Id = "BiaPgm"
prtTitleText = "Liste des prêts"

prtLineNb = 1

frmElpPrt.Show vbModeless


prtOrientation = vbPRORLandscape
prtPgmName = "prtPrêtList"
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

prtPrêtList_Form
For K = K1 To K2
    recTope = P_arrTOpe(K)
    recElpTable.K1 = recTope.Nature
    tableElpTable_Read recElpTable
    
    prtPrêtList_Line
    
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
Public Sub prtPrêtList_Form()
'---------------------------------------------------------
Dim x As String
prtCurrentY = XPrt.CurrentY
'!!!! Xprt.currenty à définir avant appel proc
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

x = "Intitulé": XPrt.CurrentX = prtMinX + 3500 - XPrt.TextWidth(x)
XPrt.Print x;

x = "Capital / mensualité": XPrt.CurrentX = prtMinX + 8900 - XPrt.TextWidth(x)
XPrt.Print x;

x = "%": XPrt.CurrentX = prtMinX + 9700 - XPrt.TextWidth(x)
XPrt.Print x;

x = "Période": XPrt.CurrentX = prtMinX + 11100 - XPrt.TextWidth(x)
XPrt.Print x;

x = "Réfèrences": XPrt.CurrentX = prtMinX + 15200 - XPrt.TextWidth(x)
XPrt.Print x;

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight

End Sub

Public Sub prtPrêtList_Line()
'------------------------------------------------------ligne 1---
Dim iReturn As Integer

If XPrt.CurrentY + prtlineHeight * 5 > prtMaxY Then
    frmElpPrt.prtNewPage
    prtPrêtList_Form
End If

XPrt.FontBold = True
'XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY, prtMaxX, XPrt.CurrentY + prtlineHeight, " ", 245)

xElpTable.Id = "Param"
xElpTable.K1 = "Statut"
xElpTable.K2 = recTope.Statut & recTope.StatutPlus
xElpTable.Method = "Seek="
iReturn = tableElpTable_Read(xElpTable)
If iReturn <> 0 Then xElpTable.Name = xElpTable.K2

XPrt.CurrentX = prtMinX + 1200: XPrt.Print xElpTable.Name;

TFlux_Compta.Param_Nature recTope.Nature
XPrt.FontBold = False
XPrt.FontSize = 8
XPrt.CurrentX = prtMinX + 2450: XPrt.Print paramTFlux_Nature;


XPrt.CurrentX = prtMinX + 14300: XPrt.Print (recTope.RéférenceInterne);

'_______________________________________________________________ligne 2


XPrt.FontSize = 6
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentY = XPrt.CurrentY + Height8_6
XPrt.CurrentX = prtMinX + 150: XPrt.Print "Prêt :";
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
XPrt.CurrentY = XPrt.CurrentY - Height8_6
XPrt.FontSize = 8
XPrt.FontBold = True
x = Format$(recTope.Capital, "## ### ### ### ### ##0.00")
XPrt.CurrentX = prtMinX + 8400 - XPrt.TextWidth(x): XPrt.Print x;
XPrt.FontBold = False
XPrt.CurrentX = prtMinX + 9500: XPrt.Print Trim(Format(recTope.TauxMarge, "##.#####")) & " %";
XPrt.CurrentX = prtMinX + 10450: XPrt.Print (recTope.PériodeNb);

xElpTable.Id = "Param"
xElpTable.K1 = "PériodicitéX"
xElpTable.K2 = recTope.Périodicité
xElpTable.Method = "Seek="
iReturn = tableElpTable_Read(xElpTable)
If iReturn <> 0 Then xElpTable.Name = xElpTable.K2
XPrt.CurrentX = prtMinX + 10800: XPrt.Print xElpTable.Name;

XPrt.CurrentX = prtMinX + 14300: XPrt.Print (recTope.RéférenceExterne);


'-------------------------------------------------ligne 3

XPrt.FontSize = 6
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentY = XPrt.CurrentY + Height8_6
XPrt.CurrentX = prtMinX + 150: XPrt.Print "Prélèvement :";
XPrt.CurrentY = XPrt.CurrentY - Height8_6
XPrt.FontSize = 8
XPrt.CurrentX = prtMinX + 1200: XPrt.Print Compte_Imp(recTope.EchéanceCompte);
recCompte.Intitulé = "": recCompte.Numéro = recTope.EngagementCompte:
mdbCptP0_Find recCompte
XPrt.FontSize = 6
XPrt.CurrentY = XPrt.CurrentY + Height8_6
XPrt.CurrentX = prtMinX + 2450: XPrt.Print Trim(recCompte.Intitulé);
XPrt.CurrentY = XPrt.CurrentY - Height8_6
XPrt.FontSize = 8
x = Format$(recTope.Mensualité, "## ### ### ### ### ##0.00")
XPrt.CurrentX = prtMinX + 8400 - XPrt.TextWidth(x): XPrt.Print x;

XPrt.CurrentY = XPrt.CurrentY + Height8_6
XPrt.FontSize = 6
XPrt.CurrentX = prtMinX + 9000: XPrt.Print "TEG.";
XPrt.CurrentX = prtMinX + 9500: XPrt.Print Trim(Format(recTope.TEG, "##.#####")) & " %";
XPrt.CurrentY = XPrt.CurrentY - Height8_6

XPrt.FontSize = 8
XPrt.CurrentX = prtMinX + 10500: XPrt.Print dateImp(recTope.AmjEchéance1);

XPrt.FontSize = 6
XPrt.CurrentY = XPrt.CurrentY + Height8_6
Select Case recTope.AmjEchéanceS
    Case "A"
        XPrt.CurrentX = prtMinX + 11600: XPrt.Print "Anniversaire";
    Case Else
        XPrt.CurrentX = prtMinX + 11600: XPrt.Print "Fin de Mois";
 End Select
        XPrt.CurrentY = XPrt.CurrentY - Height8_6
XPrt.FontSize = 8
XPrt.CurrentX = prtMinX + 12500: XPrt.Print dateImp(recTope.AmjFin);
'--------------------------------------------------ligne4
XPrt.FontSize = 6
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentY = XPrt.CurrentY + Height8_6
XPrt.CurrentX = prtMinX + 150: XPrt.Print "Déblocage :";
XPrt.CurrentY = XPrt.CurrentY - Height8_6
XPrt.FontSize = 8
XPrt.CurrentX = prtMinX + 1200: XPrt.Print Compte_Imp(recTope.EngagementCorrCompte);
recCompte.Intitulé = "": recCompte.Numéro = recTope.EngagementCorrCompte:
mdbCptP0_Find recCompte
XPrt.FontSize = 6
XPrt.CurrentY = XPrt.CurrentY + Height8_6
XPrt.CurrentX = prtMinX + 2450: XPrt.Print Trim(recCompte.Intitulé);
XPrt.CurrentY = XPrt.CurrentY - Height8_6

XPrt.CurrentY = XPrt.CurrentY + Height8_6
XPrt.CurrentX = prtMinX + 9000: XPrt.Print "T act";
XPrt.CurrentX = prtMinX + 9500: XPrt.Print Trim(Format(recTope.TauxActuariel, "##.#####")) & " %";
XPrt.CurrentX = prtMinX + 14300: XPrt.Print "(" & Trim(recTope.IdRéférence) & ")";
XPrt.CurrentY = XPrt.CurrentY - Height8_6
XPrt.FontSize = 8
XPrt.CurrentX = prtMinX + 10500: XPrt.Print dateImp(recTope.AmjDébut);

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY)
XPrt.CurrentY = XPrt.CurrentY + 50

End Sub

'---------------------------------------------------------
 Public Sub prtPrêt_Monitor(Msg As String)
'---------------------------------------------------------
Dim K As Integer, K1 As Integer, K2 As Integer, Kmin As Integer, Kmax As Integer
Dim x As String
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
    Case "T":   prtTitleText = "Tableau d'amortissement : "
                blnTableauAmortissement = True
                Kmin = K2 + 1: Kmax = 0
    Case "à":  prtTitleText = "à Valider : ": blnValidation = True
                Kmin = 10: Kmax = K2 - 10: blnSaut = True
    Case Else:  prtTitleText = "Caractéristiques "
                Kmin = 10: Kmax = K2 - 10: blnSaut = True
End Select

recElpTable_Init recElpTable
xElpTable = recElpTable
recElpTable.Method = "Seek="
recElpTable.Id = "BiaPgm"
recElpTable.K1 = recTope.Nature
tableElpTable_Read recElpTable
prtTitleText = prtTitleText & Trim(recElpTable.Name) & " - Réf : " & recTope.RéférenceInterne

prtLineNb = 1

frmElpPrt.Show vbModeless


prtOrientation = vbPRORPortrait
prtPgmName = "prtPrêt"
prtTitleUsr = usrName

prtlineHeight = 300
prtHeaderHeight = 300

If Not blnTableauAmortissement Then
    frmElpPrt.prtStdInit
Else
    prtFormType = "   "
    
    frmElpPrt.prtInit
    frmElpPrt.prtStdTop
End If

recCompteInit recCompte
recCompte.Société = SocId$
recCompte.Agence = SocAgence$
recCompte.Devise = CV1.DeviseN
recCompte.BiaTyp = "000"
recCompte.BiaNum = "00"
recCompte.Method = "SeekL1"

curFrais = 0: curIntérêtsIntermédiaires = 0
For K = K1 To K2
    If P_arrTFlux(K).CodeOpération = "PR02" Then Exit For
    If P_arrTFlux(K).CodeOpération = "PR03" Then curFrais = P_arrTFlux(K).Capital + P_arrTFlux(K).Intérêts
    If P_arrTFlux(K).CodeOpération = "PR04" Then curIntérêtsIntermédiaires = P_arrTFlux(K).Capital + P_arrTFlux(K).Intérêts
Next K

prtPrêt_Header

prtPrêt_Form
CapitalRestantDû = recTope.Capital
totalMensualité = 0: totalIntérêts = 0: totalAmortissement = 0

For K = K1 To K2
    recTFlux = P_arrTFlux(K)
    If recTFlux.CodeOpération = "PR02" Then
        CapitalRestantDû = CapitalRestantDû - recTFlux.Capital
        totalMensualité = totalMensualité + recTope.Mensualité
        totalIntérêts = totalIntérêts + recTFlux.Intérêts
        totalAmortissement = totalAmortissement + recTFlux.Capital
        If K <= Kmin Or K >= Kmax Then
            prtPrêt_Line
        Else
            If blnSaut = True Then
                blnSaut = False
                XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
            End If
        End If
        
    End If
    
    DoEvents: If prtKillDoc Then frmElpPrt.Hide: Exit Sub

Next K
XPrt.DrawWidth = 5
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight

'XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY)
'XPrt.CurrentY = XPrt.CurrentY + 100
Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY - 50, prtMaxX, XPrt.CurrentY + prtlineHeight - 50, " ", 240)

x = Format$(totalMensualité, "#### ### ### ### ##0.00")
XPrt.CurrentX = prtMinX + 3900 - XPrt.TextWidth(x)
XPrt.Print x;
x = Format$(totalIntérêts, "#### ### ### ### ##0.00")
XPrt.CurrentX = prtMinX + 6200 - XPrt.TextWidth(x)
XPrt.Print x;
x = Format$(totalAmortissement, "#### ### ### ### ##0.00")
XPrt.CurrentX = prtMinX + 8500 - XPrt.TextWidth(x)
XPrt.Print x;

x = Format$(CapitalRestantDû, "#### ### ### ### ##0.00")
XPrt.CurrentX = prtMaxX - 100 - XPrt.TextWidth(x)
XPrt.Print x;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY)

If totalMensualité <> totalAmortissement + totalIntérêts Then
    XPrt.FontSize = 14: XPrt.CurrentX = prtMinX: XPrt.Print "totalMensualité <> totalAmortissement + totalIntérêts"
End If

If recTope.Capital <> totalAmortissement Then
    XPrt.FontSize = 14: XPrt.CurrentX = prtMinX: XPrt.Print "opération.Capital <> totalAmortissement"
End If


If CapitalRestantDû <> 0 Then
    XPrt.FontSize = 14: XPrt.CurrentX = prtMinX: XPrt.Print "CapitalRestantDû <> 0"
End If

If blnValidation Then prtCompta_Validation "VALIDATION du prêt N° : " & recTope.RéférenceInterne

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
Public Sub prtPrêt_Form()
'---------------------------------------------------------
Dim x As String
prtCurrentY = XPrt.CurrentY
XPrt.FontSize = 10

XPrt.FontBold = True
XPrt.DrawWidth = 3

Call frmElpPrt.prtTrame(prtMinX, prtCurrentY, prtMaxX, prtCurrentY + prtlineHeight, " ", 240)
XPrt.DrawWidth = 1

'XPrt.Line (prtMinX, prtCurrentY)-(prtMinX, prtMaxY)
'XPrt.Line (prtMinX + 1700, prtCurrentY)-(prtMinX + 1700, prtMaxY)
'XPrt.Line (prtMinX + 4000, prtCurrentY)-(prtMinX + 4000, prtMaxY)
'XPrt.Line (prtMinX + 6300, prtCurrentY)-(prtMinX + 6300, prtMaxY)
'XPrt.Line (prtMinX + 8600, prtCurrentY)-(prtMinX + 8600, prtMaxY)
'XPrt.Line (prtMaxX, prtCurrentY)-(prtMaxX, prtMaxY)
'XPrt.Line (prtMinX, prtMaxY)-(prtMaxX, prtMaxY)

'---------------------------------------------------------

XPrt.CurrentY = prtCurrentY + 50
XPrt.CurrentX = prtMinX + 300: XPrt.Print "Echéance";

x = "Mensualité": XPrt.CurrentX = prtMinX + 3900 - XPrt.TextWidth(x)
XPrt.Print x;

x = "Intérêts": XPrt.CurrentX = prtMinX + 6200 - XPrt.TextWidth(x)
XPrt.Print x;

x = "Amortissement": XPrt.CurrentX = prtMinX + 8500 - XPrt.TextWidth(x)
XPrt.Print x;

x = "Capital restant dû": XPrt.CurrentX = prtMaxX - 100 - XPrt.TextWidth(x)
XPrt.Print x;
XPrt.FontSize = 8
prtLineNb = 0
End Sub

'---------------------------------------------------------
Public Sub prtPrêt_Line()
'---------------------------------------------------------

If XPrt.CurrentY + prtlineHeight * 1.5 > prtMaxY Then
    frmElpPrt.prtNewPage
    XPrt.CurrentY = prtMinX + prtlineHeight * 3
    prtPrêt_Form
End If

XPrt.FontBold = False
'------------------------------------------ligne 1--------------
prtLineNb = prtLineNb + 1
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
If prtLineNb > 2 Then
    
    Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY - 50, prtMaxX, XPrt.CurrentY + prtlineHeight - 50, " ", 250)
    If prtLineNb = 4 Then prtLineNb = 0
End If


XPrt.CurrentX = prtMinX + 250

XPrt.Print dateImp(recTFlux.AmjEchéanceTrt);
x = Format$(recTope.Mensualité, "#### ### ### ### ##0.00")
XPrt.CurrentX = prtMinX + 3900 - XPrt.TextWidth(x)
XPrt.Print x;
x = Format$(recTFlux.Intérêts, "#### ### ### ### ##0.00")
XPrt.CurrentX = prtMinX + 6200 - XPrt.TextWidth(x)
XPrt.Print x;
x = Format$(recTFlux.Capital, "#### ### ### ### ##0.00")
XPrt.CurrentX = prtMinX + 8500 - XPrt.TextWidth(x)
XPrt.Print x;

x = Format$(CapitalRestantDû, "#### ### ### ### ##0.00")
XPrt.CurrentX = prtMaxX - 100 - XPrt.TextWidth(x)
XPrt.Print x;

End Sub






Public Sub prtPrêt_Header()
Dim strdev As String
strdev = " " & CV1.DeviseLibellé
XPrt.FontSize = 6
XPrt.CurrentY = prtMinY + prtlineHeight: XPrt.CurrentX = prtMinX: XPrt.Print "Compte de prêt ";
XPrt.CurrentX = prtMinX + 1500: XPrt.Print ":";
'XPrt.FontSize = 8
XPrt.CurrentX = prtMinX + 1600: XPrt.Print Compte_Imp(recTope.EngagementCompte);
XPrt.FontSize = 6
recCompte.Intitulé = "": recCompte.Numéro = recTope.EngagementCompte:
mdbCptP0_Find recCompte
prtCurrentX = XPrt.CurrentX
XPrt.Print " " & Trim(recCompte.Intitulé);

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight - 100
If recCompte.TypeGA = "A" Then
    XPrt.CurrentX = prtCurrentX: XPrt.Print Trim(DicLib(13, recCompte.BiaTyp));
End If

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight - 100: XPrt.CurrentX = prtMinX: XPrt.Print "Montant du prêt ";
XPrt.CurrentX = prtMinX + 1500: XPrt.Print ":";
'XPrt.FontSize = 8
XPrt.CurrentX = prtMinX + 1600: XPrt.Print Trim(Format(recTope.Capital, "### ### ### ###.00")) & strdev;
XPrt.FontSize = 6
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight: XPrt.CurrentX = prtMinX: XPrt.Print "Date du prêt ";
XPrt.CurrentX = prtMinX + 1500: XPrt.Print ":";
'XPrt.FontSize = 8
XPrt.CurrentX = prtMinX + 1600: XPrt.Print dateImp(recTope.AmjDébut);
XPrt.FontSize = 6
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight: XPrt.CurrentX = prtMinX: XPrt.Print "Nombre de période";
XPrt.CurrentX = prtMinX + 1500: XPrt.Print ":";
'XPrt.FontSize = 8
XPrt.CurrentX = prtMinX + 1600: XPrt.Print Trim(Format(recTope.PériodeNb, "### ##0"));
XPrt.FontSize = 6
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight: XPrt.CurrentX = prtMinX: XPrt.Print "Périodicité";
XPrt.CurrentX = prtMinX + 1500: XPrt.Print ":";
'XPrt.FontSize = 8
XPrt.CurrentX = prtMinX + 1600: XPrt.Print recPériodicité_Libellé(recTope.Périodicité);
XPrt.FontSize = 6
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight: XPrt.CurrentX = prtMinX: XPrt.Print "Taux nominal";
XPrt.CurrentX = prtMinX + 1500: XPrt.Print ":";
'XPrt.FontSize = 8
XPrt.CurrentX = prtMinX + 1600: XPrt.Print Trim(Format(recTope.TauxMarge, "##.#####")) & " %";

XPrt.FontSize = 6
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight: XPrt.CurrentX = prtMinX: XPrt.Print "Déblocage des fonds ";
XPrt.CurrentX = prtMinX + 1500: XPrt.Print ":";
'XPrt.FontSize = 8
XPrt.CurrentX = prtMinX + 1600: XPrt.Print Compte_Imp(recTope.EngagementCorrCompte);
XPrt.FontSize = 6
recCompte.Intitulé = "": recCompte.Numéro = recTope.EngagementCorrCompte:
mdbCptP0_Find recCompte
XPrt.Print " " & Trim(recCompte.Intitulé);

XPrt.FontSize = 6
XPrt.CurrentY = prtMinY + prtlineHeight: XPrt.CurrentX = prtMinX + 5500: XPrt.Print "Compte à prélever ";
XPrt.CurrentX = prtMinX + 7100: XPrt.Print ":";
'XPrt.FontSize = 8
XPrt.CurrentX = prtMinX + 7200: XPrt.Print Compte_Imp(recTope.EchéanceCompte);
XPrt.FontSize = 6
recCompte.Intitulé = "": recCompte.Numéro = recTope.EchéanceCompte:
mdbCptP0_Find recCompte
prtCurrentX = XPrt.CurrentX
XPrt.Print " " & Trim(recCompte.Intitulé);

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight - 100
If recCompte.TypeGA = "A" Then
    XPrt.CurrentX = prtCurrentX: XPrt.Print Trim(DicLib(13, recCompte.BiaTyp));
End If

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight - 100: XPrt.CurrentX = prtMinX + 5500: XPrt.Print "Mensualité ";
XPrt.CurrentX = prtMinX + 7100: XPrt.Print ":";
'XPrt.FontSize = 8
XPrt.CurrentX = prtMinX + 7200: XPrt.Print Trim(Format(recTope.Mensualité, "### ### ### ###.00")) & strdev;
XPrt.FontSize = 6
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight: XPrt.CurrentX = prtMinX + 5500: XPrt.Print "Dernière échéance ";
XPrt.CurrentX = prtMinX + 7100: XPrt.Print ":";
'XPrt.FontSize = 8
XPrt.CurrentX = prtMinX + 7200: XPrt.Print dateImp(recTope.AmjFin);
XPrt.FontSize = 6
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight: XPrt.CurrentX = prtMinX + 5500: XPrt.Print "Frais";
XPrt.CurrentX = prtMinX + 7100: XPrt.Print ":";
'XPrt.FontSize = 8
If recTope.Frais <> 0 Then XPrt.CurrentX = prtMinX + 7200: XPrt.Print Trim(Format(curFrais, "### ### ### ###.00")) & strdev;
XPrt.FontSize = 6
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight: XPrt.CurrentX = prtMinX + 5500: XPrt.Print "Intérêts intermédiaires";
XPrt.CurrentX = prtMinX + 7100: XPrt.Print ":";
'XPrt.FontSize = 8
XPrt.CurrentX = prtMinX + 7200: XPrt.Print Trim(Format(curIntérêtsIntermédiaires, "### ### ### ###.00")) & strdev;
XPrt.FontSize = 6
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight: XPrt.CurrentX = prtMinX + 5500: XPrt.Print "TEG";
XPrt.CurrentX = prtMinX + 7100: XPrt.Print ":";
'XPrt.FontSize = 8
XPrt.CurrentX = prtMinX + 7200: XPrt.Print Trim(Format(recTope.TEG, "##.#####")) & " %";
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight: XPrt.CurrentX = prtMinX + 5500: XPrt.Print "Taux actuariel";
XPrt.CurrentX = prtMinX + 7100: XPrt.Print ":";
'XPrt.FontSize = 8
XPrt.CurrentX = prtMinX + 7200: XPrt.Print Trim(Format(recTope.TauxActuariel, "##.#####")) & " %";

XPrt.CurrentY = prtMinY + prtlineHeight * 9
End Sub
