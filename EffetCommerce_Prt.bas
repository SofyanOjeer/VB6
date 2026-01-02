Attribute VB_Name = "prtEffetCommerce"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Private meCompte As typeCompte
Dim X As String, I As Integer, Height8_6 As Integer

Private meCptInfo As typeCptInfo
Private meCptMvt As typeCptMvt

Private meCV1 As typeCV

Dim prtLineNb As Integer

Dim meElpTable As typeElpTable
Dim meGMemo As typegMemo
Dim blnPage As Boolean

Dim col1 As Integer, col2 As Integer, col3 As Integer, Col4 As Integer, Col5 As Integer
Dim Line1 As Integer, Line2 As Integer, Line3 As Integer

Dim curTotal As Currency
'---------------------------------------------------------
Public Sub prtEffetCommerce_Dossier_Form()
'---------------------------------------------------------
Dim X As String
prtCurrentY = XPrt.CurrentY
XPrt.FontSize = 10

XPrt.FontBold = True
XPrt.DrawWidth = 3

Call frmElpPrt.prtTrame(prtMinX, Line3 - 50, prtMaxX, Line3 + prtlineHeight - 50, " ", 220)
XPrt.DrawWidth = 1

'---------------------------------------------------------
XPrt.FontSize = 6

XPrt.CurrentY = Line3 + 50
XPrt.CurrentX = prtMinX: XPrt.Print "Echéancier";
col1 = prtMinX + 1300
col2 = prtMinX + 3500
col3 = prtMinX + 5500
Col4 = prtMinX + 12000
Col5 = prtMinX + 13000
XPrt.CurrentX = col1: XPrt.Print "Statut";
XPrt.CurrentX = col2: XPrt.Print "Opération";

XPrt.CurrentX = col3: XPrt.Print "Compte";
XPrt.CurrentX = Col4 - 1100: XPrt.Print "Montant";
XPrt.CurrentX = Col4 + 100: XPrt.Print "Date valeur";
XPrt.CurrentX = Col5: XPrt.Print "Libellé";
XPrt.FontSize = 8
prtLineNb = 0
End Sub

'---------------------------------------------------------
Public Sub prtEffetCommerce_Dossier_GEch(lparam As typeGParam, lGope As typeGOpe, lGech As typeGEch, lGmemo_Nb As Integer, lGmemo() As typegMemo)
'---------------------------------------------------------
Dim wGMemo_Index As Integer
Dim blnHeight8_6 As Boolean, blnLine As Boolean

If XPrt.CurrentY + prtlineHeight * 1.5 > prtMaxY Then
    frmElpPrt.prtNewPage
    XPrt.CurrentY = prtMinX + prtlineHeight * 3
    prtEffetCommerce_Dossier_Form
End If

XPrt.FontBold = False
'------------------------------------------ligne 1--------------
prtLineNb = prtLineNb + 1
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight


'If lGech.Statut = "A" Then
'    blnHeight8_6 = True
    XPrt.FontSize = 6
'    XPrt.CurrentY = XPrt.CurrentY + Height8_6
'Else
'    blnHeight8_6 = False
'End If
Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY - 100, prtMaxX, XPrt.CurrentY + prtlineHeight - 100, " ", 240)

XPrt.CurrentX = prtMinX: XPrt.Print dateImp(lGech.EchAMJ);
XPrt.CurrentX = col1: XPrt.Print lGech.Statut & lGech.StatutPlus & " " & dateImp10(lGech.ActionAmj) & timeImp(lGech.ActionHms);
XPrt.CurrentX = col2: XPrt.Print lGech.EchFct;
curTotal = 0
blnLine = False
For wGMemo_Index = 1 To lGmemo_Nb
        XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
        prtEffetCommerce_Dossier_GMemo lGmemo(wGMemo_Index)
Next wGMemo_Index

If curTotal <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
    XPrt.fpntbold = True: XPrt.FontSize = 14
    XPrt.CurrentX = 10000: XPrt.Print "pièce non équilibrée";
End If

'If Not blnHeight8_6 Then
'    XPrt.FontSize = 6
'    XPrt.CurrentY = XPrt.CurrentY + Height8_6
'End If



'XPrt.CurrentY = XPrt.CurrentY - Height8_6
'XPrt.FontSize = 8

End Sub
'---------------------------------------------------------
Public Sub prtEffetCommerce_Dossier_GFlux(lparam As typeGParam, lGope As typeGOpe, lGFlux As typeGFlux)
'---------------------------------------------------------
Dim blnHeight8_6 As Boolean

If XPrt.CurrentY + prtlineHeight * 1.5 > prtMaxY Then
    frmElpPrt.prtNewPage
    XPrt.CurrentY = prtMinX + prtlineHeight * 3
    prtEffetCommerce_Dossier_Form
End If

XPrt.FontBold = False
'------------------------------------------ligne 1--------------
prtLineNb = prtLineNb + 1
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight


If lGFlux.Statut = "A" Then
    blnHeight8_6 = True
    XPrt.FontSize = 6
    XPrt.CurrentY = XPrt.CurrentY + Height8_6
Else
    blnHeight8_6 = False
End If

XPrt.CurrentX = prtMinX + 50: XPrt.Print dateImp(lGFlux.AmjEchéanceTrt);
XPrt.CurrentX = prtMinX + 1300: XPrt.Print TFlux_Compta.Param_CodeOpération(lGFlux.OpérationCode);
    
If lGFlux.Statut <> "A" Then

    If lGFlux.Montant1 <> 0 Then
        X = Format$(lGFlux.Montant1, "#### ### ### ### ##0.00")
        XPrt.CurrentX = prtMinX + 5000 - XPrt.TextWidth(X)
        XPrt.Print X;
    End If
    If lGFlux.Montant2 <> 0 Then
        X = Format$(lGFlux.Montant2, "#### ### ### ### ##0.00")
        XPrt.CurrentX = prtMinX + 6500 - XPrt.TextWidth(X)
        XPrt.Print X;
    End If
End If

If Not blnHeight8_6 Then
    XPrt.FontSize = 6
    XPrt.CurrentY = XPrt.CurrentY + Height8_6
End If

XPrt.CurrentX = prtMinX + 6700: XPrt.Print param_Statut(lGFlux.Statut & lGFlux.StatutPlus);

XPrt.CurrentX = prtMinX + 8500: XPrt.Print dateImp(lGFlux.AmjDébut);
XPrt.CurrentX = prtMinX + 9500: XPrt.Print dateImp(lGFlux.AmjFin);
XPrt.CurrentX = prtMinX + 11000: XPrt.Print dateImp(lGFlux.AmjValeur);

XPrt.CurrentY = XPrt.CurrentY - Height8_6
XPrt.FontSize = 8

End Sub

Public Sub prtEffetCommerce_Dossier_GOpe(lparam As typeGParam, lGope As typeGOpe, lGech As typeGEch, lGmemo As typegMemo)
''Dim strdev As String
''Call CV_AttributS(lGope.Devise1, meCV1)
''strdev = "_" & meCV1.DeviseLibellé
Line1 = prtMinY + prtlineHeight / 2
Line2 = Line1 + prtlineHeight * 5
Line3 = Line2 + prtlineHeight * 4

'XPrt.Line (prtMinX, line2 - 100)-(prtMaxX, line2 - 100)
col1 = prtMinX
col2 = prtMinX + 1300
col3 = prtMinX + 1400
XPrt.FontSize = 6

XPrt.CurrentY = Line1
XPrt.CurrentX = col1: XPrt.Print "Compte à créditer ";
XPrt.CurrentX = col2: XPrt.Print ":";
XPrt.FontBold = True
XPrt.CurrentX = col3: XPrt.Print lGope.Devise1 & " " & Compte_Imp(lGope.EchéanceCompte);
XPrt.FontBold = False
meCompte.Intitulé = "": meCompte.Numéro = lGope.EchéanceCompte:
mdbCptP0_Find meCompte
If meCompte.TypeGA = "A" Then XPrt.Print "   " & Trim(DicLib(13, meCompte.BiaTyp));

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = col3: XPrt.Print Trim(meCompte.Intitulé);

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = col1: XPrt.Print "Portefeuille";
XPrt.CurrentX = col2: XPrt.Print ":";
XPrt.FontBold = True
XPrt.CurrentX = col3: XPrt.Print lGope.Devise1 & " " & Compte_Imp(lGope.EngagementCompte);
XPrt.FontBold = False
meCompte.Intitulé = "": meCompte.Numéro = lGope.EngagementCompte:
mdbCptP0_Find meCompte
If meCompte.TypeGA = "A" Then XPrt.Print " " & Trim(DicLib(13, meCompte.BiaTyp));

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = col3: XPrt.Print Trim(meCompte.Intitulé);

XPrt.CurrentY = Line2

XPrt.CurrentX = col1: XPrt.Print "Date de remise ";
XPrt.CurrentX = col2: XPrt.Print ":";
XPrt.CurrentX = col3: XPrt.Print dateImp(lGope.AmjDébut);


XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
If lGope.TauxMarge2 <> 0 Then
    XPrt.CurrentX = col1: XPrt.Print "Taux d'escompte";
    XPrt.CurrentX = col2: XPrt.Print ":";
    XPrt.CurrentX = col3: XPrt.Print lGope.TauxMarge2 & " %";
    
    If mId$(lGope.TauxRéférence2, 3, 1) = "1" Then XPrt.Print " + TauxMajoré (> 90jours)";
    
    If mId$(lGope.TauxRéférence2, 4, 1) = "1" Then XPrt.Print " + taux effet non accepté";
    
    If lGope.TauxMarge1 <> lGope.TauxMarge2 Then XPrt.Print " => " & lGope.TauxMarge1 & " %";
End If
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
If lGope.PériodeNb <> 0 Then
    XPrt.CurrentX = col1: XPrt.Print "Période";
    XPrt.CurrentX = col2: XPrt.Print ":";
    XPrt.CurrentX = col3: XPrt.Print lGope.PériodeNb & " jours ( " & dateImp10(lGope.AmjDébut) & " - " & dateImp10(lGope.AmjFin) & " )";
End If


XPrt.CurrentY = Line1
col1 = prtMinX + 5500
col2 = prtMinX + 6900
col3 = prtMinX + 7000

XPrt.CurrentX = col1: XPrt.Print "Nature";
XPrt.CurrentX = col2: XPrt.Print ":";
''srvGSub.Param_Nature lParam
XPrt.FontBold = True
XPrt.CurrentX = col3: XPrt.Print lparam.NatureCode & " " & lparam.NatureLib;
XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = col1: XPrt.Print "Date d'échéance";
XPrt.CurrentX = col2: XPrt.Print ":";
XPrt.FontBold = True
XPrt.CurrentX = col3: XPrt.Print dateImp(lGope.AmjFin);
XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = col1: XPrt.Print "Montant nominal ";
XPrt.CurrentX = col2: XPrt.Print ":";
XPrt.FontBold = True
X = Trim(Format(lGope.Montant1, "### ### ### ###.00"))
XPrt.CurrentX = col3 + 100: prtCurrentX = XPrt.CurrentX + XPrt.TextWidth(X)
XPrt.Print X & "  " & lGope.Devise1;
XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = col1: XPrt.Print "Montant net ";
XPrt.CurrentX = col2: XPrt.Print ":";
XPrt.FontBold = True
X = Trim(Format(lGope.Montant1 - lGope.Montant2, "### ### ### ###.00"))
XPrt.CurrentX = prtCurrentX - XPrt.TextWidth(X)
XPrt.Print X & "  " & lGope.Devise1; '' strdev;
XPrt.FontBold = False

XPrt.CurrentY = Line2
XPrt.CurrentX = col1: XPrt.Print "Agios";
XPrt.CurrentX = col2: XPrt.Print ":";
XPrt.FontBold = True
X = Trim(Format(lGope.Montant2, "### ### ### ###.00"))
XPrt.CurrentX = prtCurrentX - XPrt.TextWidth(X)
XPrt.Print X & "  " & lGope.Devise2;
XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = col1: XPrt.Print "Intérêts ";
XPrt.CurrentX = col2: XPrt.Print ":";
If lGope.Mensualité <> 0 Then
    X = Trim(Format(lGope.Mensualité, "### ### ### ###.00"))
    XPrt.CurrentX = prtCurrentX - XPrt.TextWidth(X)
    XPrt.Print X;
End If

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = col1: XPrt.Print "Commission d'endos";
XPrt.CurrentX = col2: XPrt.Print ":";
If lGope.Frais1 <> 0 Then
    X = Trim(Format(lGope.Frais1, "### ### ### ###.00"))
    XPrt.CurrentX = prtCurrentX - XPrt.TextWidth(X)
    XPrt.Print X;
End If

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = col1: XPrt.Print "Com manipulation & TVA";
XPrt.CurrentX = col2: XPrt.Print ":";
If lGope.Frais2 <> 0 Then
    X = Trim(Format(lGope.Frais2, "### ### ### ###.00"))
    XPrt.CurrentX = prtCurrentX - XPrt.TextWidth(X)
    XPrt.Print X;
End If
If lGope.Frais3 <> 0 Then XPrt.Print "  + " & Trim(Format(lGope.Frais3, "### ### ##0.00"));

XPrt.CurrentY = Line1
col1 = prtMinX + 11000
col2 = prtMinX + 12400
col3 = prtMinX + 12500

XPrt.CurrentX = col1: XPrt.Print "Référence interne";
XPrt.CurrentX = col2: XPrt.Print ":";
XPrt.FontBold = True
XPrt.CurrentX = col3: XPrt.Print lGope.RéférenceInterne;
XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = col1: XPrt.Print "Réf bordereau de remise";
XPrt.CurrentX = col2: XPrt.Print ":";
XPrt.CurrentX = col3: XPrt.Print lGope.RéférenceExterne;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight: XPrt.CurrentX = col1: XPrt.Print "Référence informatique";
XPrt.CurrentX = col2: XPrt.Print ":";
XPrt.CurrentX = col3: XPrt.Print Trim(lGope.Application) & "_ " & lGope.IdRéférence;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight: XPrt.CurrentX = col1: XPrt.Print "Saisie";
XPrt.CurrentX = col2: XPrt.Print ":";
XPrt.CurrentX = col3: XPrt.Print dateImp(lGech.EchAMJ) & "    " & timeImp(lGech.EchHMS) & "   " & lGech.EchUsr;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight: XPrt.CurrentX = col1: XPrt.Print "Validation";
XPrt.CurrentX = col2: XPrt.Print ":";
XPrt.CurrentX = col3: XPrt.Print dateImp(lGech.ActionAmj) & "    " & timeImp(lGech.ActionHms) & "   " & lGech.ActionUsr;

XPrt.CurrentY = Line2 + prtlineHeight
XPrt.CurrentX = col1: XPrt.Print "Tiré";
XPrt.CurrentX = col2: XPrt.Print ":";
XPrt.CurrentX = col3: XPrt.Print mId$(lGmemo.MemoText, 1, 50);

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = col1: XPrt.Print "Domiciliation";
XPrt.CurrentX = col2: XPrt.Print ":";
XPrt.CurrentX = col3: XPrt.Print mId$(lGmemo.MemoText, 51, 50);

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = col1: XPrt.Print "Référence du tiré";
XPrt.CurrentX = col2: XPrt.Print ":";
XPrt.CurrentX = col3: XPrt.Print mId$(lGmemo.MemoText, 101, 50);
End Sub



Public Sub prtEffetCommerce_Avis_Open()
On Error GoTo prtError

Set XPrt = Printer
frmElpPrt.Show vbModeless

prtOrientation = vbPRORPortrait
prtTitleText = "Avis"
prtPgmName = "prtEffetCommerce_Avis"
prtTitleUsr = usrName

prtLineNb = 1
prtlineHeight = 300
prtHeaderHeight = 300

prtFormType = ""
frmElpPrt.prtInit

meCV1 = CV_Euro
recElpTable_Init meElpTable
recGMemo_Init meGMemo
blnPage = False

Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub
Public Sub prtEffetCommerce_Dossier_Open()
On Error GoTo prtError

Set XPrt = Printer
frmElpPrt.Show vbModeless

prtOrientation = vbPRORLandscape
prtTitleText = "Effet de commerce"
prtPgmName = "prtEffetCommerce_Dossier"
prtTitleUsr = usrName

prtLineNb = 1
prtlineHeight = 300
prtHeaderHeight = 300

prtFormType = ""
'frmElpPrt.prtInit

meCV1 = CV_Euro
frmElpPrt.prtStdInit
recCompteInit meCompte

recElpTable_Init meElpTable
recGMemo_Init meGMemo
blnPage = False

Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub

Public Sub prtEffetCommerce_Close()
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

Public Sub prtEffetCommerce_Avis_Line(lGope As typeGOpe)
Dim V
lGope = lGope

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight

Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY, prtMaxX, XPrt.CurrentY + prtlineHeight, " ", 230)
XPrt.FontBold = True
XPrt.FontSize = 9
XPrt.CurrentY = XPrt.CurrentY + 50
XPrt.CurrentX = prtMinX + 100: XPrt.Print lGope.RéférenceInterne;
XPrt.CurrentX = prtMinX + 1500: XPrt.Print dateImp10(lGope.AmjFin);
    
X = Format$(lGope.Montant1, "### ### ### ##0.00")
XPrt.CurrentX = prtMinX + 4500 - XPrt.TextWidth(X): XPrt.Print X;
    
If lGope.Frais2 <> 0 Then
    X = Format$(lGope.Frais2, "### ### ### ##0.00")
    XPrt.CurrentX = prtMinX + 6000 - XPrt.TextWidth(X): XPrt.Print X;
End If

If lGope.Frais3 <> 0 Then
    X = Format$(lGope.Frais3, "### ### ### ##0.00")
    XPrt.CurrentX = prtMinX + 7000 - XPrt.TextWidth(X): XPrt.Print X;
End If

If lGope.Frais1 <> 0 Then
    X = Format$(lGope.Frais1, "### ### ### ##0.00")
    XPrt.CurrentX = prtMinX + 8000 - XPrt.TextWidth(X): XPrt.Print X;
End If

If lGope.Mensualité <> 0 Then
    X = Format$(lGope.Mensualité, "### ### ### ##0.00")
    XPrt.CurrentX = prtMinX + 9500 - XPrt.TextWidth(X): XPrt.Print X;
End If

X = Format$(lGope.Montant1 - lGope.Montant2, "### ### ### ##0.00")
XPrt.CurrentX = prtMaxX - XPrt.TextWidth(X): XPrt.Print X;


prtCurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentY = prtCurrentY
XPrt.FontBold = False
XPrt.FontSize = 8

XPrt.CurrentX = prtMinX + 5500
XPrt.Print "Taux d'escompte " & lGope.TauxMarge2 & " %";

If mId$(lGope.TauxRéférence2, 3, 1) = "1" Then
'    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'    XPrt.CurrentX = prtMinX + 100
    XPrt.Print " + TauxMajoré (> 90jours)";
End If

If mId$(lGope.TauxRéférence2, 4, 1) = "1" Then
'    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'    XPrt.CurrentX = prtMinX + 100
    XPrt.Print " + taux effet non accepté";
End If

If lGope.TauxMarge1 <> lGope.TauxMarge2 Then XPrt.Print " : " & lGope.TauxMarge1 & " %";

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinX + 5500
XPrt.Print "du " & dateImp10(lGope.AmjDébut); "au " & dateImp10(lGope.AmjFin) & "  :  " & lGope.PériodeNb & " jours";


XPrt.CurrentY = prtCurrentY
XPrt.CurrentX = prtMinX + 1500

XPrt.CurrentY = prtCurrentY
XPrt.CurrentX = prtMinX + 1500
XPrt.Print Trim(mId$(meGMemo.MemoText, 1, 50));

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinX + 1500
XPrt.Print Trim(mId$(meGMemo.MemoText, 51, 50));

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinX + 1500
XPrt.Print Trim(mId$(meGMemo.MemoText, 101, 50));


End Sub
Public Sub prtEffetCommerce_Avis_Form(lGope As typeGOpe)
Dim V

If blnPage Then: frmElpPrt.prtNewPage
blnPage = True
XPrt.CurrentY = prtMaxY
XPrt.FontBold = False
prtSocMiniFin

XPrt.CurrentY = prtMinY
prtSocMini XPrt.CurrentY, lGope.AmjEngagement

XPrt.FontBold = False
XPrt.FontSize = 9
XPrt.CurrentX = prtMinX + 100
XPrt.CurrentY = prtMinY + 2100
XPrt.Print "N/Référence :";
XPrt.FontBold = True
XPrt.CurrentX = prtMinX + 1500
XPrt.Print Trim(lGope.RéférenceExterne);
XPrt.FontBold = False

Call CV_AttributS(lGope.Devise1, meCV1)


recCptInfoInit meCptInfo
meCptInfo.Method = "JoinL1"
meCptInfo.Société = SocId$
meCptInfo.Agence = SocId$
meCptInfo.Devise = meCV1.DeviseN
meCptInfo.Numéro = lGope.EngagementCompte
meCptInfo.BiaTyp = "000"
meCptInfo.BiaNum = "00000"
meCptInfo.NuméroAncien = "00000000000"
If Not IsNull(mdbCptInfoP0_Find(meCptInfo)) Then
    Call MsgBox("prtEffetCommerce_Avis_Form : compte d'engagement inconnu", vbCritical, "Impression")
    Exit Sub
End If

XPrt.CurrentY = 0

prtAdresse XPrt.CurrentY, meCptInfo

meElpTable.Method = "Seek="
meElpTable.Id = "GFlux_EC"
meElpTable.K1 = "Nature"
meElpTable.K2 = lGope.Nature
tableElpTable_Read meElpTable

meGMemo.Method = "SeekP0"
meGMemo.IdRéférence = lGope.IdRéférence
meGMemo.MemoSéquence = 1
V = srvGMemo.srvGMemo_Monitor(meGMemo)
If Not IsNull(V) Then recGMemo_Init meGMemo

'Call frmElpPrt.prtTrame(prtMinX, prtMinY + 4000, prtMaxX, prtMinY + 4500, " ", 245)

XPrt.FontBold = True
XPrt.FontSize = 14
XPrt.CurrentX = prtMinX + 100
XPrt.CurrentY = prtMinY + 4200
frmElpPrt.prtCentré 5500, "Bordereau de remise d'effets en " & lGope.Devise1
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
frmElpPrt.prtCentré 5500, Trim(meElpTable.Name)


Call frmElpPrt.prtTrame(prtMinX, prtMinY + 5500, prtMaxX, prtMinY + 5500 + prtlineHeight, " ", 210)
XPrt.CurrentY = prtMinY + 5550

XPrt.FontBold = True
XPrt.FontSize = 7
XPrt.CurrentX = prtMinX + 100: XPrt.Print "Référence";
XPrt.CurrentX = prtMinX + 1500: XPrt.Print "Echéance";
X = "Nominal"
XPrt.CurrentX = prtMinX + 4500 - XPrt.TextWidth(X): XPrt.Print X;
X = "com manipulation"
XPrt.CurrentX = prtMinX + 6000 - XPrt.TextWidth(X): XPrt.Print X;
X = "TVA"
XPrt.CurrentX = prtMinX + 7000 - XPrt.TextWidth(X): XPrt.Print X;
X = "com d'endos"
XPrt.CurrentX = prtMinX + 8000 - XPrt.TextWidth(X): XPrt.Print X;
X = "Escompte"
XPrt.CurrentX = prtMinX + 9500 - XPrt.TextWidth(X): XPrt.Print X;
X = "Montant Net"
XPrt.CurrentX = prtMaxX - XPrt.TextWidth(X): XPrt.Print X;


XPrt.FontBold = False

End Sub


Public Sub prtEffetCommerce_Dossier_GMemo(lGmemo As typegMemo)
Dim X As String
X = lGmemo.Statut & lGmemo.StatutPlus
If X <> "àC " Then XPrt.CurrentX = col1: XPrt.Print X;
XPrt.CurrentX = col2: XPrt.Print lGmemo.MemoNature
If lGmemo.MemoLien1 > 0 Then XPrt.Print "  " & lGmemo.MemoLien1 & "  " & lGmemo.MemoLien2;

If Trim(lGmemo.MemoNature) <> constCompta Then
    XPrt.CurrentX = col3: XPrt.Print lGmemo.MemoText
Else

    Call srvCptMvt_GetX(meCptMvt, lGmemo.MemoText)
    If meCptMvt.Mt <> 0 Then
        curTotal = curTotal + meCptMvt.Mt
        
        XPrt.CurrentX = col3: XPrt.Print meCptMvt.Devise & "." & Compte_Imp(meCptMvt.Compte);
        
        XPrt.CurrentX = col3 + 1200
        meCompte.Devise = meCptMvt.Devise
        meCompte.Numéro = meCptMvt.Compte
        If IsNull(mdbCptP0_Find(meCompte)) Then
                XPrt.Print Trim(meCompte.Intitulé);
        Else
            XPrt.Print "???????";
        End If
       
        
        XPrt.FontBold = True
        If meCptMvt.Mt < 0 Then
            XPrt.CurrentX = Col4 - 1100
        Else
            XPrt.CurrentX = Col4 - 100
        End If

        X = Format$(Abs(meCptMvt.Mt), "### ### ### ### ##0.00")
        XPrt.CurrentX = XPrt.CurrentX - XPrt.TextWidth(X)
        XPrt.Print X;
        XPrt.FontBold = False
        
        XPrt.CurrentX = Col4 + 100
        XPrt.Print dateImp(meCptMvt.AmjValeur);
        XPrt.CurrentX = Col5
        XPrt.Print meCptMvt.Libellé;

    End If
End If

End Sub
