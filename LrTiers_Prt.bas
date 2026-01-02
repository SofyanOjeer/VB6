Attribute VB_Name = "prtLrTiers"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------

Private recLrTiers As typeLrTiers
Dim I As Integer
Dim NbImprimé As Integer
Dim kPage As Integer

Dim K As Integer, K1 As Integer, K2 As Integer
Dim X As String
Dim blnList As Boolean

'---------------------------------------------------------
Public Sub prtLrTiers_Form()
'---------------------------------------------------------
Dim X As String
NbImprimé = 0
XPrt.FontSize = 8
XPrt.FontBold = True

XPrt.DrawWidth = 2
XPrt.Line (prtMinX, prtMinY + prtHeaderHeight)-(prtMaxX, prtMinY + prtHeaderHeight)

XPrt.DrawWidth = 1


XPrt.Line (prtMinX, prtMinY)-(prtMinX, prtMaxY)
XPrt.Line (prtMaxX, prtMinY)-(prtMaxX, prtMaxY)

'---------------------------------------------------------

X = "Attribut"
XPrt.CurrentY = prtMinY + (prtHeaderHeight - XPrt.TextHeight(X)) / 2
XPrt.CurrentX = prtMinX + 20
XPrt.Print X;

XPrt.CurrentX = 2400: XPrt.Print "Codification";

XPrt.CurrentY = prtMinY + prtHeaderHeight - XPrt.TextHeight("X")

End Sub

'---------------------------------------------------------
Public Sub prtLrTiersList_Form()
'---------------------------------------------------------
Dim X As String
NbImprimé = 0
XPrt.FontSize = 8
XPrt.FontBold = True

XPrt.DrawWidth = 2
XPrt.Line (prtMinX, prtMinY + prtHeaderHeight)-(prtMaxX, prtMinY + prtHeaderHeight)

XPrt.DrawWidth = 1


XPrt.Line (prtMinX, prtMinY)-(prtMinX, prtMaxY)
XPrt.Line (prtMaxX, prtMinY)-(prtMaxX, prtMaxY)



'---------------------------------------------------------

X = "Racine / Siren / Bdf"
XPrt.CurrentY = prtMinY + (prtHeaderHeight - XPrt.TextHeight(X) + 100) / 2
XPrt.CurrentX = prtMinX + 20
XPrt.Print X;

XPrt.CurrentX = 3200: XPrt.Print "Intitulé.";
XPrt.CurrentX = 6100: XPrt.Print "Rés.";
XPrt.CurrentX = 7200: XPrt.Print "APE.";
XPrt.CurrentX = 8100: XPrt.Print "JUR.";
XPrt.CurrentX = 9100: XPrt.Print "AGECO.";
XPrt.CurrentX = 11200: XPrt.Print "Prénoms / Date Naissance / Commune / Pays&Département.";
XPrt.CurrentX = 15200: XPrt.Print "Sexe.";


XPrt.CurrentY = prtMinY + prtHeaderHeight - XPrt.TextHeight("X")

End Sub


'---------------------------------------------------------
 Public Sub prtLrTiersX(Msg As String)
'---------------------------------------------------------

K1 = Val(Mid$(Msg, 1, 6))
K2 = Val(Mid$(Msg, 7, 6))
blnList = IIf(Mid$(Msg, 13, 1) = "L", True, False)

Set XPrt = Printer
frmElpPrt.Show vbModeless
If blnList Then

    prtOrientation = vbPRORLandscape
    prtTitleText = "Liste des bénéficiaires (Luca Risques)"
    prtPgmName = "prtLrTiers"
    prtTitleUsr = usrName
    
    prtLineNb = 1
    prtlineHeight = 250
    prtHeaderHeight = 300
    
    frmElpPrt.prtStdinit
    
    prtLrTiersList_Form
    
    For K = K1 To K2
        
 '''       recLrTiers = arrLrTiers(K)
        prtLrTiersList_Line
        XPrt.Line (10000, prtMinY)-(10000, XPrt.CurrentY)
         DoEvents: If prtKillDoc Then frmElpPrt.Hide: Exit Sub
    Next K
Else

    prtOrientation = vbPRORPortrait
    prtTitleText = "Fiche bénéficiaire (Luca Risques)"
    prtPgmName = "prtLrTiers"
    prtTitleUsr = usrName
    
    prtLineNb = 1
    prtlineHeight = 250
    prtHeaderHeight = 300
    
    frmElpPrt.prtStdinit
    
    prtLrTiers_Form
    
    For K = K1 To K2
        
''''''''''''''''''        recLrTiers = arrLrTiers(K)
        prtLrTiers_Line
        
        DoEvents: If prtKillDoc Then frmElpPrt.Hide: Exit Sub
    Next K
End If

frmElpPrt.prtEndDoc

frmElpPrt.Hide
End Sub




'---------------------------------------------------------
Public Sub prtLrTiers_Line()
'---------------------------------------------------------
Dim X As String, K As Integer
Dim Situation As String

If XPrt.CurrentY + prtlineHeight * 17 > prtMaxY Then
   frmElpPrt.prtNewPage
   prtLrTiers_Form
End If
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.FontBold = False
Call frmElpPrt.prtTrame(prtMinX + 10, XPrt.CurrentY, prtMaxX, XPrt.CurrentY + prtlineHeight * 2, " ")

XPrt.CurrentX = prtMinX + 10
XPrt.Print "Référence";
XPrt.CurrentX = prtMinX + 1800
XPrt.Print ":";
XPrt.CurrentX = prtMinX + 2000
XPrt.FontBold = True
XPrt.Print recLrTiers.RFBENF;
XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinX + 10
XPrt.Print "Nom / Conjoint";
XPrt.CurrentX = prtMinX + 1800
XPrt.Print ":";
XPrt.CurrentX = prtMinX + 2000
XPrt.Print recLrTiers.NOMBNF;
XPrt.CurrentX = prtMinX + 3000
XPrt.Print recLrTiers.NOMCJT;
'--------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinX + 10
XPrt.Print "Prénoms";
XPrt.CurrentX = prtMinX + 1800
XPrt.Print ":";
XPrt.CurrentX = prtMinX + 2000
XPrt.Print recLrTiers.PRENOM;
'---------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinX + 10
XPrt.Print "Adresse";
XPrt.CurrentX = prtMinX + 1800
XPrt.Print ":";
XPrt.CurrentX = prtMinX + 2000
XPrt.Print recLrTiers.NOVOIE;
'-----------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinX + 10
XPrt.Print "Code Postal";
XPrt.CurrentX = prtMinX + 1800
XPrt.Print ":";
XPrt.CurrentX = prtMinX + 2000
XPrt.Print recLrTiers.CDPOST;
XPrt.CurrentX = prtMinX + 3000
XPrt.Print recLrTiers.LBCOMM2;

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinX + 10
XPrt.Print "Pays /Département";
XPrt.CurrentX = prtMinX + 1800
XPrt.Print ":";
XPrt.CurrentX = prtMinX + 2000
XPrt.Print recLrTiers.CDPAYS2;
XPrt.CurrentX = prtMinX + 2400
XPrt.Print recLrTiers.CDDEPT2;
XPrt.CurrentX = prtMinX + 10000
XPrt.Print recLrTiers.CDPAYN2;

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.DrawWidth = 1
XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY)
XPrt.CurrentY = XPrt.CurrentY + 100


XPrt.CurrentX = prtMinX + 10
XPrt.Print "Numéro SIREN";
XPrt.CurrentX = prtMinX + 1800
XPrt.Print ":";
XPrt.CurrentX = prtMinX + 2000
XPrt.Print recLrTiers.NSIREN;

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinX + 10
XPrt.Print "Code BDF";
XPrt.CurrentX = prtMinX + 1800
XPrt.Print ":";
XPrt.CurrentX = prtMinX + 2000
XPrt.Print Format$(recLrTiers.NBDF1, "@@ @@@@@@@@@ @@  ");
XPrt.Print "(" & dateImp(recLrTiers.AMJ1) & ")";
XPrt.CurrentX = prtMinX + 6000
XPrt.Print Format$(recLrTiers.NBDF2, "@@ @@@@@@@@@ @@  ");
XPrt.Print "(" & dateImp(recLrTiers.AMJ2) & ")";

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinX + 10
XPrt.Print "Code Résident";
XPrt.CurrentX = prtMinX + 1800
XPrt.Print ":";
XPrt.CurrentX = prtMinX + 2000
XPrt.Print recLrTiers.CDRESI;
XPrt.CurrentX = prtMinX + 6000
XPrt.Print recLrTiers.CDRESI1;
XPrt.CurrentX = prtMinX + 10000
XPrt.Print recLrTiers.CDRESI2;

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.DrawWidth = 1
XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY)
XPrt.CurrentY = XPrt.CurrentY + 100

XPrt.CurrentX = prtMinX + 10
XPrt.Print "Activité Economique";
XPrt.CurrentX = prtMinX + 1800
XPrt.Print ":";
XPrt.CurrentX = prtMinX + 2000
XPrt.Print recLrTiers.CDACCO;
XPrt.CurrentX = prtMinX + 6000
XPrt.Print recLrTiers.CDACCO2;
XPrt.CurrentX = prtMinX + 10000
XPrt.Print recLrTiers.CDACEN1;

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinX + 10
XPrt.Print "Catégorie Juridique";
XPrt.CurrentX = prtMinX + 1800
XPrt.Print ":";
XPrt.CurrentX = prtMinX + 2000
XPrt.Print recLrTiers.CTJURI;
XPrt.CurrentX = prtMinX + 6000
XPrt.Print recLrTiers.CTJURN1;

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinX + 10
XPrt.Print "Agent Economique";
XPrt.CurrentX = prtMinX + 1800
XPrt.Print ":";
XPrt.CurrentX = prtMinX + 2000
XPrt.Print recLrTiers.CDAGCO;

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinX + 10
XPrt.Print "Sexe";
XPrt.CurrentX = prtMinX + 1800
XPrt.Print ":";
XPrt.CurrentX = prtMinX + 2000
XPrt.Print recLrTiers.CDSEXE;
XPrt.CurrentX = prtMinX + 2200
XPrt.Print recLrTiers.CDSEXE1;

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinX + 10
XPrt.Print "Date de Naissance";
XPrt.CurrentX = prtMinX + 1800
XPrt.Print ":";
XPrt.CurrentX = prtMinX + 2000
XPrt.Print recLrTiers.JMA3;


XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinX + 10
XPrt.Print "Commune";
XPrt.CurrentX = prtMinX + 1800
XPrt.Print ":";
XPrt.CurrentX = prtMinX + 2000
XPrt.Print recLrTiers.CDCOMM1;
XPrt.CurrentX = prtMinX + 2400
XPrt.Print recLrTiers.LBCOMM1;

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinX + 10
XPrt.Print "Pays / Département";
XPrt.CurrentX = prtMinX + 1800
XPrt.Print ":";
XPrt.CurrentX = prtMinX + 2000
XPrt.Print recLrTiers.CDPAYS1;
XPrt.CurrentX = prtMinX + 2400
XPrt.Print recLrTiers.CDDEPT1;

XPrt.CurrentX = prtMinX + 12000
XPrt.Print recLrTiers.CDPAYN1;

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.DrawWidth = 5
XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY)

End Sub

'---------------------------------------------------------
Public Sub prtLrTiersList_Line()
'---------------------------------------------------------
Dim X As String, K As Integer
Dim Situation As String

If XPrt.CurrentY + prtlineHeight * 4 > prtMaxY Then
    XPrt.Line (10300, prtMinY)-(10300, prtMaxY)
    frmElpPrt.prtNewPage
    prtLrTiersList_Form
End If
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.FontBold = False
'Call frmElpPrt.prtTrame(prtMinX + 10, XPrt.CurrentY, prtMaxX, XPrt.CurrentY + prtlineHeight * 2, " ")
'----------------------------------------ligne 1
XPrt.CurrentX = prtMinX + 10
XPrt.FontBold = True
XPrt.Print recLrTiers.RFBENF;
XPrt.CurrentX = prtMinX + 2000
XPrt.FontBold = False
XPrt.Print recLrTiers.NOMBNF;
XPrt.CurrentX = prtMinX + 5400
XPrt.Print recLrTiers.NOMCJT;
XPrt.CurrentX = prtMinX + 6000
XPrt.Print recLrTiers.CDRESI;
XPrt.CurrentX = prtMinX + 7000
XPrt.Print recLrTiers.CDACCO;
XPrt.CurrentX = prtMinX + 8000
XPrt.Print recLrTiers.CTJURI;
XPrt.CurrentX = prtMinX + 9000
XPrt.Print recLrTiers.CDAGCO;
XPrt.CurrentX = prtMinX + 10000
XPrt.Print recLrTiers.JMA3;
XPrt.CurrentX = prtMinX + 11000
XPrt.Print recLrTiers.PRENOM;
XPrt.CurrentX = prtMinX + 15100
XPrt.Print recLrTiers.CDSEXE;

'-----------------------------------ligne 2-
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinX + 10
XPrt.Print recLrTiers.NSIREN;
XPrt.CurrentX = prtMinX + 2000
XPrt.Print recLrTiers.NOVOIE;
XPrt.CurrentX = prtMinX + 6000
XPrt.Print recLrTiers.CDRESI1;
XPrt.CurrentX = prtMinX + 7000
XPrt.Print recLrTiers.CDACEN1;
XPrt.CurrentX = prtMinX + 8000
XPrt.Print recLrTiers.CTJURN1;
XPrt.CurrentX = prtMinX + 9000
XPrt.Print recLrTiers.CDCOMM1;
XPrt.CurrentX = prtMinX + 9200
XPrt.Print recLrTiers.LBCOMM1;
XPrt.CurrentX = prtMinX + 10000
XPrt.Print recLrTiers.CDPAYS1;
XPrt.CurrentX = prtMinX + 15100
XPrt.Print recLrTiers.CDSEXE1;

'------------------------------------ligne 3-
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinX + 10
XPrt.Print Format$(recLrTiers.NBDF1, "@@ @@@@@@@@@ @@  ");
'XPrt.Print "(" & dateImp(recLrTiers.AMJ1) & ")"
XPrt.CurrentX = prtMinX + 2800
XPrt.Print recLrTiers.CDPOST;
XPrt.CurrentX = prtMinX + 2000
XPrt.Print recLrTiers.LBCOMM2;
XPrt.CurrentX = prtMinX + 6000
XPrt.Print recLrTiers.CDRESI2;
XPrt.CurrentX = prtMinX + 7000
XPrt.Print recLrTiers.CDACCO2;
XPrt.CurrentX = prtMinX + 10000
XPrt.Print recLrTiers.CDPAYN1; '

'------------------------------------ligne 4-
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinX + 10
XPrt.Print Format$(recLrTiers.NBDF2, "@@ @@@@@@@@@ @@  ");
'XPrt.Print "(" & dateImp(recLrTiers.AMJ2) & ")";
XPrt.CurrentX = prtMinX + 2000
XPrt.Print recLrTiers.CDPAYS2;
XPrt.CurrentX = prtMinX + 2300
XPrt.Print recLrTiers.CDDEPT2;
XPrt.CurrentX = prtMinX + 11000
XPrt.Print recLrTiers.CDPAYN2;




XPrt.CurrentY = XPrt.CurrentY + prtlineHeight + 100
XPrt.DrawWidth = 1
XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY - 200) '

End Sub


