Attribute VB_Name = "prtLrSgnBnf"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------

Private recLrSgnBnf As typeLrSgnBnf
Dim I As Integer
Dim NbImprimé As Integer
Dim kPage As Integer

Dim K As Integer, K1 As Integer, K2 As Integer
Dim X As String
Dim blnList As Boolean

'---------------------------------------------------------
Public Sub prtLrSgnBnf_Form()
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
Public Sub prtLrSgnBnfList_Form()
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
 Public Sub prtLrSgnBnfX(Msg As String)
'---------------------------------------------------------

K1 = Val(Mid$(Msg, 1, 6))
K2 = Val(Mid$(Msg, 7, 6))
blnList = IIf(Mid$(Msg, 13, 1) = "L", True, False)

Set XPrt = Printer
frmElpPrt.Show vbModeless
If blnList Then

    prtOrientation = vbPRORLandscape
    prtTitleText = "Liste des bénéficiaires (Luca Risques)"
    prtPgmName = "prtLrSgnBnf"
    prtTitleUsr = usrName
    
    prtLineNb = 1
    prtlineHeight = 250
    prtHeaderHeight = 300
    
    frmElpPrt.prtStdinit
    
    prtLrSgnBnfList_Form
    
    For K = K1 To K2
        
        recLrSgnBnf = arrLrSgnBnf(K)
        prtLrSgnBnfList_Line
        XPrt.Line (10000, prtMinY)-(10000, XPrt.CurrentY)
         DoEvents: If prtKillDoc Then frmElpPrt.Hide: Exit Sub
    Next K
Else

    prtOrientation = vbPRORPortrait
    prtTitleText = "Fiche bénéficiaire (Luca Risques)"
    prtPgmName = "prtLrSgnBnf"
    prtTitleUsr = usrName
    
    prtLineNb = 1
    prtlineHeight = 250
    prtHeaderHeight = 300
    
    frmElpPrt.prtStdinit
    
    prtLrSgnBnf_Form
    
    For K = K1 To K2
        
        recLrSgnBnf = arrLrSgnBnf(K)
        prtLrSgnBnf_Line
        
        DoEvents: If prtKillDoc Then frmElpPrt.Hide: Exit Sub
    Next K
End If

frmElpPrt.prtEndDoc

frmElpPrt.Hide
End Sub




'---------------------------------------------------------
Public Sub prtLrSgnBnf_Line()
'---------------------------------------------------------
Dim X As String, K As Integer
Dim Situation As String

If XPrt.CurrentY + prtlineHeight * 17 > prtMaxY Then
   frmElpPrt.prtNewPage
   prtLrSgnBnf_Form
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
XPrt.Print recLrSgnBnf.RFBENF;
XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinX + 10
XPrt.Print "Nom / Conjoint";
XPrt.CurrentX = prtMinX + 1800
XPrt.Print ":";
XPrt.CurrentX = prtMinX + 2000
XPrt.Print recLrSgnBnf.NOMBNF;
XPrt.CurrentX = prtMinX + 3000
XPrt.Print recLrSgnBnf.NOMCJT;
'--------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinX + 10
XPrt.Print "Prénoms";
XPrt.CurrentX = prtMinX + 1800
XPrt.Print ":";
XPrt.CurrentX = prtMinX + 2000
XPrt.Print recLrSgnBnf.PRENOM;
'---------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinX + 10
XPrt.Print "Adresse";
XPrt.CurrentX = prtMinX + 1800
XPrt.Print ":";
XPrt.CurrentX = prtMinX + 2000
XPrt.Print recLrSgnBnf.NOVOIE;
'-----------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinX + 10
XPrt.Print "Code Postal";
XPrt.CurrentX = prtMinX + 1800
XPrt.Print ":";
XPrt.CurrentX = prtMinX + 2000
XPrt.Print recLrSgnBnf.CDPOST;
XPrt.CurrentX = prtMinX + 3000
XPrt.Print recLrSgnBnf.LBCOMM2;

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinX + 10
XPrt.Print "Pays /Département";
XPrt.CurrentX = prtMinX + 1800
XPrt.Print ":";
XPrt.CurrentX = prtMinX + 2000
XPrt.Print recLrSgnBnf.CDPAYS2;
XPrt.CurrentX = prtMinX + 2400
XPrt.Print recLrSgnBnf.CDDEPT2;
XPrt.CurrentX = prtMinX + 10000
XPrt.Print recLrSgnBnf.CDPAYN2;

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.DrawWidth = 1
XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY)
XPrt.CurrentY = XPrt.CurrentY + 100


XPrt.CurrentX = prtMinX + 10
XPrt.Print "Numéro SIREN";
XPrt.CurrentX = prtMinX + 1800
XPrt.Print ":";
XPrt.CurrentX = prtMinX + 2000
XPrt.Print recLrSgnBnf.NSIREN;

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinX + 10
XPrt.Print "Code BDF";
XPrt.CurrentX = prtMinX + 1800
XPrt.Print ":";
XPrt.CurrentX = prtMinX + 2000
XPrt.Print Format$(recLrSgnBnf.NBDF1, "@@ @@@@@@@@@ @@  ");
XPrt.Print "(" & dateImp(recLrSgnBnf.AMJ1) & ")";
XPrt.CurrentX = prtMinX + 6000
XPrt.Print Format$(recLrSgnBnf.NBDF2, "@@ @@@@@@@@@ @@  ");
XPrt.Print "(" & dateImp(recLrSgnBnf.AMJ2) & ")";

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinX + 10
XPrt.Print "Code Résident";
XPrt.CurrentX = prtMinX + 1800
XPrt.Print ":";
XPrt.CurrentX = prtMinX + 2000
XPrt.Print recLrSgnBnf.CDRESI;
XPrt.CurrentX = prtMinX + 6000
XPrt.Print recLrSgnBnf.CDRESI1;
XPrt.CurrentX = prtMinX + 10000
XPrt.Print recLrSgnBnf.CDRESI2;

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.DrawWidth = 1
XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY)
XPrt.CurrentY = XPrt.CurrentY + 100

XPrt.CurrentX = prtMinX + 10
XPrt.Print "Activité Economique";
XPrt.CurrentX = prtMinX + 1800
XPrt.Print ":";
XPrt.CurrentX = prtMinX + 2000
XPrt.Print recLrSgnBnf.CDACCO;
XPrt.CurrentX = prtMinX + 6000
XPrt.Print recLrSgnBnf.CDACCO2;
XPrt.CurrentX = prtMinX + 10000
XPrt.Print recLrSgnBnf.CDACEN1;

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinX + 10
XPrt.Print "Catégorie Juridique";
XPrt.CurrentX = prtMinX + 1800
XPrt.Print ":";
XPrt.CurrentX = prtMinX + 2000
XPrt.Print recLrSgnBnf.CTJURI;
XPrt.CurrentX = prtMinX + 6000
XPrt.Print recLrSgnBnf.CTJURN1;

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinX + 10
XPrt.Print "Agent Economique";
XPrt.CurrentX = prtMinX + 1800
XPrt.Print ":";
XPrt.CurrentX = prtMinX + 2000
XPrt.Print recLrSgnBnf.CDAGCO;

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinX + 10
XPrt.Print "Sexe";
XPrt.CurrentX = prtMinX + 1800
XPrt.Print ":";
XPrt.CurrentX = prtMinX + 2000
XPrt.Print recLrSgnBnf.CDSEXE;
XPrt.CurrentX = prtMinX + 2200
XPrt.Print recLrSgnBnf.CDSEXE1;

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinX + 10
XPrt.Print "Date de Naissance";
XPrt.CurrentX = prtMinX + 1800
XPrt.Print ":";
XPrt.CurrentX = prtMinX + 2000
XPrt.Print recLrSgnBnf.JMA3;


XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinX + 10
XPrt.Print "Commune";
XPrt.CurrentX = prtMinX + 1800
XPrt.Print ":";
XPrt.CurrentX = prtMinX + 2000
XPrt.Print recLrSgnBnf.CDCOMM1;
XPrt.CurrentX = prtMinX + 2400
XPrt.Print recLrSgnBnf.LBCOMM1;

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinX + 10
XPrt.Print "Pays / Département";
XPrt.CurrentX = prtMinX + 1800
XPrt.Print ":";
XPrt.CurrentX = prtMinX + 2000
XPrt.Print recLrSgnBnf.CDPAYS1;
XPrt.CurrentX = prtMinX + 2400
XPrt.Print recLrSgnBnf.CDDEPT1;

XPrt.CurrentX = prtMinX + 12000
XPrt.Print recLrSgnBnf.CDPAYN1;

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinX + 10
XPrt.Print "Divers";
XPrt.CurrentX = prtMinX + 1800
XPrt.Print ":";
XPrt.CurrentX = prtMinX + 2000
XPrt.Print recLrSgnBnf.CDTRI1 & " : ";
XPrt.Print recLrSgnBnf.CDTRI2 & " : ";
XPrt.Print recLrSgnBnf.CDHABI & " : ";
XPrt.Print recLrSgnBnf.AMJ4 & " : ";
XPrt.Print recLrSgnBnf.HMSC & " : ";
XPrt.Print recLrSgnBnf.CDPHMO & " : ";
XPrt.Print recLrSgnBnf.CDCRMD & " : ";
XPrt.Print recLrSgnBnf.INDSIR & " : ";
XPrt.Print recLrSgnBnf.FILL01 & " : ";
XPrt.Print recLrSgnBnf.FILL02;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.DrawWidth = 5
XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY)

End Sub

'---------------------------------------------------------
Public Sub prtLrSgnBnfList_Line()
'---------------------------------------------------------
Dim X As String, K As Integer
Dim Situation As String

If XPrt.CurrentY + prtlineHeight * 4 > prtMaxY Then
    XPrt.Line (10300, prtMinY)-(10300, prtMaxY)
    frmElpPrt.prtNewPage
    prtLrSgnBnfList_Form
End If
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.FontBold = False
'Call frmElpPrt.prtTrame(prtMinX + 10, XPrt.CurrentY, prtMaxX, XPrt.CurrentY + prtlineHeight * 2, " ")
'----------------------------------------ligne 1
XPrt.CurrentX = prtMinX + 10
XPrt.FontBold = True
XPrt.Print recLrSgnBnf.RFBENF;
XPrt.CurrentX = prtMinX + 2000
XPrt.FontBold = False
XPrt.Print recLrSgnBnf.NOMBNF;
XPrt.CurrentX = prtMinX + 5400
XPrt.Print recLrSgnBnf.NOMCJT;
XPrt.CurrentX = prtMinX + 6000
XPrt.Print recLrSgnBnf.CDRESI;
XPrt.CurrentX = prtMinX + 7000
XPrt.Print recLrSgnBnf.CDACCO;
XPrt.CurrentX = prtMinX + 8000
XPrt.Print recLrSgnBnf.CTJURI;
XPrt.CurrentX = prtMinX + 9000
XPrt.Print recLrSgnBnf.CDAGCO;
XPrt.CurrentX = prtMinX + 10000
XPrt.Print recLrSgnBnf.JMA3;
XPrt.CurrentX = prtMinX + 11000
XPrt.Print recLrSgnBnf.PRENOM;
XPrt.CurrentX = prtMinX + 15100
XPrt.Print recLrSgnBnf.CDSEXE;

'-----------------------------------ligne 2-
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinX + 10
XPrt.Print recLrSgnBnf.NSIREN;
XPrt.CurrentX = prtMinX + 2000
XPrt.Print recLrSgnBnf.NOVOIE;
XPrt.CurrentX = prtMinX + 6000
XPrt.Print recLrSgnBnf.CDRESI1;
XPrt.CurrentX = prtMinX + 7000
XPrt.Print recLrSgnBnf.CDACEN1;
XPrt.CurrentX = prtMinX + 8000
XPrt.Print recLrSgnBnf.CTJURN1;
XPrt.CurrentX = prtMinX + 9000
XPrt.Print recLrSgnBnf.CDCOMM1;
XPrt.CurrentX = prtMinX + 9200
XPrt.Print recLrSgnBnf.LBCOMM1;
XPrt.CurrentX = prtMinX + 10000
XPrt.Print recLrSgnBnf.CDPAYS1;
XPrt.CurrentX = prtMinX + 15100
XPrt.Print recLrSgnBnf.CDSEXE1;

'------------------------------------ligne 3-
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinX + 10
XPrt.Print Format$(recLrSgnBnf.NBDF1, "@@ @@@@@@@@@ @@  ");
'XPrt.Print "(" & dateImp(recLrSgnBnf.AMJ1) & ")"
XPrt.CurrentX = prtMinX + 2800
XPrt.Print recLrSgnBnf.CDPOST;
XPrt.CurrentX = prtMinX + 2000
XPrt.Print recLrSgnBnf.LBCOMM2;
XPrt.CurrentX = prtMinX + 6000
XPrt.Print recLrSgnBnf.CDRESI2;
XPrt.CurrentX = prtMinX + 7000
XPrt.Print recLrSgnBnf.CDACCO2;
XPrt.CurrentX = prtMinX + 10000
XPrt.Print recLrSgnBnf.CDPAYN1; '

'------------------------------------ligne 4-
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinX + 10
XPrt.Print Format$(recLrSgnBnf.NBDF2, "@@ @@@@@@@@@ @@  ");
'XPrt.Print "(" & dateImp(recLrSgnBnf.AMJ2) & ")";
XPrt.CurrentX = prtMinX + 2000
XPrt.Print recLrSgnBnf.CDPAYS2;
XPrt.CurrentX = prtMinX + 2300
XPrt.Print recLrSgnBnf.CDDEPT2;
XPrt.CurrentX = prtMinX + 11000
XPrt.Print recLrSgnBnf.CDPAYN2;




XPrt.CurrentY = XPrt.CurrentY + prtlineHeight + 100
XPrt.DrawWidth = 1
XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY - 200) '

End Sub


