Attribute VB_Name = "prtRacine"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------

Private recRacine As typeRacine
Private recTitulaire As typeTitulaire


'---------------------------------------------------------
Public Sub prtRacineX(Msg As String)
'---------------------------------------------------------
Dim K As Integer, K1 As Integer, K2 As Integer
Dim X As String
On Error GoTo prtError

Set XPrt = Printer
K1 = Val(mId$(Msg, 1, 6))
K2 = Val(mId$(Msg, 7, 6))


frmElpPrt.Show vbModeless

prtOrientation = vbPRORPortrait
prtTitleText = "Fiche Racine"
prtPgmName = "prtRacine"
prtTitleUsr = usrName

prtLineNb = 1
prtlineHeight = 250
prtHeaderHeight = 300

frmElpPrt.prtStdInit

For K = K1 To K2
    prtRacineForm
    recRacine = arrRacine(K)
    prtRacineLine
    If K < K2 Then
        frmElpPrt.prtNewPage
    Else
        frmElpPrt.prtEndDoc
    End If
    DoEvents: If prtKillDoc Then frmElpPrt.Hide: Exit Sub
Next K

frmElpPrt.prtEndDoc

prtTitulaireX
Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub




'---------------------------------------------------------
Public Sub prtRacineForm()
'---------------------------------------------------------
Dim X As String

XPrt.FontSize = 8
XPrt.FontBold = True

XPrt.FillStyle = 0
XPrt.DrawWidth = 3
XPrt.ForeColor = RGB(0, 0, 0)
XPrt.FillStyle = 1

XPrt.Line (prtMinX, prtMinY)-(prtMaxX, prtMaxY), , B
XPrt.Line (prtMinX, prtMinY + prtHeaderHeight)-(prtMaxX, prtMinY + prtHeaderHeight)

XPrt.DrawWidth = 2


'----------------------------------------ligne 1-----------------

XPrt.CurrentY = prtMinY + (prtHeaderHeight - XPrt.TextHeight(X)) / 2

XPrt.CurrentX = 9000
XPrt.Print "Agence";
XPrt.CurrentX = 9650
XPrt.Print ":";
XPrt.CurrentX = 9800
XPrt.Print strSocBdfE;
XPrt.Print strSocBdfG;




prtCurrentY = prtMinY + prtHeaderHeight


End Sub

'---------------------------------------------------------
Public Sub prtRacineLine()
'---------------------------------------------------------
Dim X As String, K As Integer, mCurrenty As Integer

'If prtCurrentY + prtParagraphHeight > prtMaxY Then
 '   frmElpPrt.prtNewPage
'    prtCompteForm
'Else
    'frmElpPrt.prtLineY
'End If
XPrt.DrawWidth = 1
XPrt.FontBold = False

'------------------------------------------ligne 1--------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight + 200 - XPrt.TextHeight("test")
XPrt.FontSize = 8
XPrt.CurrentX = 400
XPrt.Print "Racine";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.CurrentX = 2450
XPrt.FontBold = True
XPrt.Print Format$(recRacine.Numéro, "00000");
'----------------------------------ligne2
XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Intitulé";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.CurrentX = 2450
XPrt.FontBold = True
XPrt.Print recRacine.Intitulé;
'---------------------------------ligne3------------
XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Recherche Alpha";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.CurrentX = 2450
XPrt.FontBold = True
XPrt.Print recRacine.Alpha;
'----------------------------------4-----------
XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Type";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.CurrentX = 2450
XPrt.FontBold = True
XPrt.Print recRacine.TypeBanqueClient;
XPrt.CurrentX = 3200
XPrt.FontBold = False
Select Case recRacine.TypeBanqueClient
    Case "C"
    XPrt.Print "Client";
    Case "B"
    XPrt.Print "Banque";
End Select
'-----------------------------------5--------
XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Actionnaire";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.CurrentX = 2450
XPrt.FontBold = True
XPrt.Print recRacine.Actionnaire;
XPrt.CurrentX = 3200
XPrt.FontBold = False
XPrt.Print DicLib(7, recRacine.Actionnaire);

'-----------------------------------6----------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Nature Titulaire";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.CurrentX = 2450
XPrt.FontBold = True
XPrt.Print recRacine.NatureTitulaire;
XPrt.CurrentX = 3200
XPrt.FontBold = False
If recRacine.TypeBanqueClient = "B" Then
    XPrt.Print DicLib(18, recRacine.NatureTitulaire);
Else
    XPrt.Print DicLib(62, recRacine.NatureTitulaire);
End If
'------------------------------------7---------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Nombre de Titulaire";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.CurrentX = 2400
XPrt.FontBold = True
XPrt.Print recRacine.NombreTitulaire;


'-------------------------------------8--------
XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Membre du Personnel";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.CurrentX = 2450
XPrt.FontBold = True
XPrt.Print recRacine.MembreduPersonnel;
XPrt.CurrentX = 3200
XPrt.FontBold = False
XPrt.Print DicLib(7, recRacine.MembreduPersonnel);

'------------------------------------9---------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Nomination";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.CurrentX = 2450
XPrt.FontBold = True
XPrt.Print recRacine.Nomination;
XPrt.CurrentX = 3200
XPrt.FontBold = False
XPrt.Print DicLib(11, recRacine.Nomination);

'-----------------------------------10----------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Opposition";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.CurrentX = 2450
XPrt.FontBold = True
XPrt.Print recRacine.Opposition;
XPrt.CurrentX = 3200
XPrt.FontBold = False
XPrt.Print DicLib(16, recRacine.Opposition);

'------------------------------------11-
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY)
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Adresse1";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.CurrentX = 2450
XPrt.FontBold = True
XPrt.Print recRacine.Adresse1;
'-------------------------------------12--------
XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Adresse02";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.CurrentX = 2450
XPrt.FontBold = True
XPrt.Print recRacine.Adresse2;
'--------------------------------------13-------
XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Adresse03";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.CurrentX = 2450
XPrt.FontBold = True
XPrt.Print recRacine.Adresse3;
'--------------------------------------14-------
XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Code Postal";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.CurrentX = 2450
XPrt.FontBold = True
XPrt.Print recRacine.AdresseCP;
'---------------------------------------15-------
XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY)
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "N° de Téléphone 1";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.CurrentX = 2450
XPrt.FontBold = True
XPrt.Print recRacine.Téléphone1;
'---------------------------------------16-------
XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "N° de Téléphone 2";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.CurrentX = 2450
XPrt.FontBold = True
XPrt.Print recRacine.Téléphone2;
'--------------------------------------17-----
XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Numéro de Télex";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.CurrentX = 2450
XPrt.FontBold = True
XPrt.Print recRacine.Télex;
'--------------------------------------18-
XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Numéro de Fax";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.CurrentX = 2450
XPrt.FontBold = True
XPrt.Print recRacine.Fax;
'-------------------------------------19----
XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Numéro Swift";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.CurrentX = 2450
XPrt.FontBold = True
XPrt.Print recRacine.Swift;
'--------------------------------------20--------

XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY)
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Type de Compte";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.CurrentX = 2450
XPrt.FontBold = True
XPrt.Print recRacine.TypeOuverture;
XPrt.CurrentX = 3200
XPrt.FontBold = False
XPrt.Print DicLib(63, recRacine.TypeOuverture);



'-----------------------------------------21--------
XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Résident Bdf";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.CurrentX = 2450
XPrt.FontBold = True
XPrt.Print recRacine.Résident;
XPrt.CurrentX = 3200
XPrt.FontBold = False
XPrt.Print DicLib(23, recRacine.Résident);

'----------------------------------------22---------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Résident Fiscal";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.CurrentX = 2450
XPrt.FontBold = True
XPrt.Print recRacine.RésidentFiscal;
XPrt.CurrentX = 3200
XPrt.FontBold = False
XPrt.Print DicLib(7, recRacine.RésidentFiscal);

'-----------------------------------------23--------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Pays de Résidence";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.CurrentX = 2450
XPrt.FontBold = True
XPrt.Print recRacine.RésidentPays;
XPrt.CurrentX = 3200
XPrt.FontBold = False
XPrt.Print DicLib(19, recRacine.RésidentPays);

'-----------------------------------------24--------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Succession";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.CurrentX = 2450
XPrt.FontBold = True
XPrt.Print recRacine.Succession;
XPrt.CurrentX = 3200
XPrt.FontBold = False
XPrt.Print DicLib(7, recRacine.Succession);

'-------------------------------------25----
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Date de Naissance";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.CurrentX = 2450
XPrt.FontBold = True
XPrt.Print dateImp(recRacine.NaissanceAmj);

'--------------------------------------26-----
XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "optionTva";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.CurrentX = 2450
XPrt.FontBold = True
XPrt.Print recRacine.optionTva;
XPrt.CurrentX = 3200
XPrt.FontBold = False
XPrt.Print DicLib(7, recRacine.optionTva);

'--------------------------------------27----
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Nationalité";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.CurrentX = 2450
XPrt.FontBold = True
XPrt.Print recRacine.Nationalité;
XPrt.CurrentX = 3200
XPrt.FontBold = False
XPrt.Print DicLib(19, recRacine.Nationalité);

'---------------------------------------28--
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Identification Bdf";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.CurrentX = 2450
XPrt.FontBold = True
XPrt.Print recRacine.BdfIdentification;
'---------------------------------------29---
XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Régime Matrimonial";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.CurrentX = 2450
XPrt.FontBold = True
XPrt.Print recRacine.RégimeMatrimonial;
XPrt.CurrentX = 3200
XPrt.FontBold = False
XPrt.Print DicLib(70, recRacine.RégimeMatrimonial);

'-----------------------------------------30-------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY)
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Regroupement Client";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.CurrentX = 2450
XPrt.FontBold = True
XPrt.Print recRacine.GroupeIdentification;
XPrt.CurrentX = 3200
XPrt.FontBold = False
XPrt.Print DicLib(65, recRacine.GroupeIdentification);

'----------------------------------------31---------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Apporteur";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.CurrentX = 2450
XPrt.FontBold = True
XPrt.Print recRacine.Apporteur;
XPrt.CurrentX = 3200
XPrt.FontBold = False
XPrt.Print DicLib(60, recRacine.Apporteur);
'----------------------------------------31---------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Groupe Economique";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.CurrentX = 2450
XPrt.FontBold = True
XPrt.Print recRacine.GroupeEconomique;
XPrt.CurrentX = 3200
XPrt.FontBold = False
XPrt.Print DicLib(64, recRacine.GroupeEconomique);


'-----------------------------------------32--------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Gestionnaire";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.CurrentX = 2450
XPrt.FontBold = True
XPrt.Print recRacine.Gestionnaire;
XPrt.CurrentX = 3200
XPrt.FontBold = False
XPrt.Print DicLib(60, recRacine.Gestionnaire);
'-----------------------------------33
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Courrier";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.CurrentX = 2450
XPrt.FontBold = True
XPrt.Print recRacine.Courrier;
XPrt.CurrentX = 3200
XPrt.FontBold = False
XPrt.Print DicLib(61, recRacine.Courrier);

'-----------------------------------34
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Commentaire";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.CurrentX = 2450
XPrt.FontBold = True
XPrt.Print recRacine.Commentaire;
'-----------------------------------35
XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "N° Siren";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.CurrentX = 2450
XPrt.FontBold = True
XPrt.Print recRacine.SIREN;
'-------------------------------------36----
XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Racine Siège";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.CurrentX = 2400
XPrt.FontBold = True
XPrt.Print recRacine.SiègeRacine;
'-------------------------------------37----
XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Interdiction de Chq";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.CurrentX = 2450
XPrt.FontBold = True
XPrt.Print recRacine.ChèquierInterdit;
XPrt.CurrentX = 3200
XPrt.FontBold = False
XPrt.Print DicLib(7, recRacine.ChèquierInterdit);

'--------------------------------------38---
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Date Interdiction";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.CurrentX = 2450
XPrt.FontBold = True
XPrt.Print dateImp(recRacine.ChèquierInterditAmj);
'---------------------------------------39--
XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Code Ape";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.CurrentX = 2450
XPrt.FontBold = True
XPrt.Print recRacine.APE;
XPrt.CurrentX = 3200
XPrt.FontBold = False
XPrt.Print DicLib(5, recRacine.APE);


'--------------------------------------40----
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY)
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Date de Création";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.CurrentX = 2450
XPrt.FontBold = True
XPrt.Print dateImp(recRacine.AmjCréation);

'--------------------------------------41-
XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Date de Modification";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.CurrentX = 2450
XPrt.FontBold = True
XPrt.Print dateImp(recRacine.AmjModification);
'-------------------------------------42--
XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Date Annulation";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.CurrentX = 2450
XPrt.FontBold = True
XPrt.Print dateImp(recRacine.AmjAnnulation);
'--------------------------------------43-
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.FontBold = False
XPrt.CurrentX = 400
XPrt.Print "Date de Réactivation";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.CurrentX = 2450
XPrt.FontBold = True
XPrt.Print dateImp(recRacine.AmjRéactivation);
'-------------------------------------44--
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.FontBold = False

XPrt.CurrentX = 400
XPrt.Print "Opérateur";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.CurrentX = 2450
XPrt.FontBold = True
XPrt.Print recRacine.Opérateur;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY)
XPrt.FontBold = False


prtCurrentY = prtCurrentY + prtParagraphHeight
        

End Sub


'---------------------------------------------------------
Public Sub prtTitulaireForm()
'---------------------------------------------------------
Dim X As String

XPrt.FontSize = 8
XPrt.FontBold = True

XPrt.FillStyle = 0
XPrt.DrawWidth = 3
XPrt.ForeColor = RGB(0, 0, 0)
XPrt.FillStyle = 1

XPrt.Line (prtMinX, prtMinY)-(prtMaxX, prtMaxY), , B
XPrt.Line (prtMinX, prtMinY + prtHeaderHeight)-(prtMaxX, prtMinY + prtHeaderHeight)

XPrt.DrawWidth = 2


'----------------------------------------ligne 1-----------------

XPrt.CurrentY = prtMinY + (prtHeaderHeight - XPrt.TextHeight(X)) / 2

XPrt.CurrentX = 9000
XPrt.Print "Agence";
XPrt.CurrentX = 9650
XPrt.Print ":";
XPrt.CurrentX = 9800
XPrt.Print strSocBdfE;
XPrt.Print strSocBdfG;

prtCurrentY = prtMinY + prtHeaderHeight


End Sub

Public Sub prtTitulaireLine()
'---------------------------------------------------------
Dim X As String, K As Integer, mCurrenty As Integer

If prtCurrentY + prtParagraphHeight > prtMaxY Then
    frmElpPrt.prtNewPage
    prtTitulaireForm
'Else
    'frmElpPrt.prtLineY
End If
XPrt.DrawWidth = 1
XPrt.FontBold = False

'------------------------------------------ligne 1--------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight + 250 - XPrt.TextHeight("test")
XPrt.FontSize = 8
XPrt.CurrentX = 400
XPrt.Print "Racine";
XPrt.CurrentX = 2350
XPrt.Print ":";
XPrt.CurrentX = 2600
XPrt.FontBold = True
XPrt.Print Format$(recTitulaire.Racine, "00000");
'-----------------------------------------ligne2
XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Nationalité";
XPrt.CurrentX = 2350
XPrt.Print ":";
XPrt.CurrentX = 2600
XPrt.FontBold = True
XPrt.Print recTitulaire.Nationalité;
XPrt.FontBold = False
XPrt.CurrentX = 3300
XPrt.Print DicLib(19, recTitulaire.Nationalité);
'------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Pays de Résidence";
XPrt.CurrentX = 2350
XPrt.Print ":";
XPrt.CurrentX = 2600
XPrt.FontBold = True
XPrt.Print recTitulaire.RésidencePays;
XPrt.FontBold = False
XPrt.CurrentX = 3300
XPrt.Print DicLib(19, recTitulaire.RésidencePays);
'------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Zone Géographique Fiscale";
XPrt.CurrentX = 2350
XPrt.Print ":";
XPrt.CurrentX = 2600
XPrt.FontBold = True
XPrt.Print recTitulaire.RésidenceZone;
XPrt.FontBold = False
XPrt.CurrentX = 3300
XPrt.Print DicLib(71, recTitulaire.RésidenceZone);

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY)

'------------------------------------------
XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Nom Patronymique";
XPrt.CurrentX = 2350
XPrt.Print ":";
XPrt.CurrentX = 2600
XPrt.FontBold = True
XPrt.Print recTitulaire.NomPatronymique;
'------------------------------------------
XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Prénom";
XPrt.CurrentX = 2350
XPrt.Print ":";
XPrt.CurrentX = 2600
XPrt.FontBold = True
XPrt.Print recTitulaire.Prénom;
'------------------------------------------
XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Nom du Mari";
XPrt.CurrentX = 2350
XPrt.Print ":";
XPrt.CurrentX = 2600
XPrt.FontBold = True
XPrt.Print recTitulaire.MariNom;
'------------------------------------------
XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Prénom du Mari";
XPrt.CurrentX = 2350
XPrt.Print ":";
XPrt.CurrentX = 2600
XPrt.FontBold = True
XPrt.Print recTitulaire.MariPrénom;
'------------------------------------------
XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Adresse Fiscale Client 1";
XPrt.CurrentX = 2350
XPrt.Print ":";
XPrt.CurrentX = 2600
XPrt.FontBold = True
XPrt.Print recTitulaire.Adresse1;
'------------------------------------------
XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Adresse Fiscale Client 2";
XPrt.CurrentX = 2350
XPrt.Print ":";
XPrt.CurrentX = 2600
XPrt.FontBold = True
XPrt.Print recTitulaire.Adresse2;
'------------------------------------------
XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Adresse Fiscale Client 3";
XPrt.CurrentX = 2350
XPrt.Print ":";
XPrt.CurrentX = 2600
XPrt.FontBold = True
XPrt.Print recTitulaire.Adresse3;
'------------------------------------------
XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Adresse Fiscale Client 4";
XPrt.CurrentX = 2350
XPrt.Print ":";
XPrt.CurrentX = 2600
XPrt.FontBold = True
XPrt.Print recTitulaire.Adresse4;
'------------------------------------------
XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Cd Postal Fiscale Client 5";
XPrt.CurrentX = 2350
XPrt.Print ":";
XPrt.CurrentX = 2600
XPrt.FontBold = True
XPrt.Print recTitulaire.AdresseCP & "  " & recTitulaire.AdresseBD;
'-------------------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY)
'------------------------------------------
XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Date de Naissance";
XPrt.CurrentX = 2350
XPrt.Print ":";
XPrt.CurrentX = 2600
XPrt.FontBold = True
XPrt.Print dateImp(recTitulaire.NaissanceAmj);
'------------------------------------------
XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Département";
XPrt.CurrentX = 2350
XPrt.Print ":";
XPrt.CurrentX = 2600
XPrt.FontBold = True
XPrt.Print recTitulaire.NaissanceDépN;
XPrt.FontBold = False
XPrt.CurrentX = 3300
XPrt.Print DicLib(51, recTitulaire.NaissanceDépN);
'------------------------------------------
XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Département ou Pays";
XPrt.CurrentX = 2350
XPrt.Print ":";
XPrt.CurrentX = 2600
XPrt.FontBold = True
XPrt.Print recTitulaire.NaissanceDépX;
'------------------------------------------
XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Commune";
XPrt.CurrentX = 2350
XPrt.Print ":";
XPrt.CurrentX = 2600
XPrt.FontBold = True
XPrt.Print recTitulaire.NaissanceCommune;
'------------------------------------------
XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Pays";
XPrt.CurrentX = 2350
XPrt.Print ":";
XPrt.CurrentX = 2600
XPrt.FontBold = True
XPrt.Print recTitulaire.NaissancePays;
XPrt.FontBold = False
XPrt.CurrentX = 3300
XPrt.Print DicLib(19, recTitulaire.NaissancePays);
'------------------------------------------
XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Zone Géographique";
XPrt.CurrentX = 2350
XPrt.Print ":";
XPrt.CurrentX = 2600
XPrt.FontBold = True
XPrt.Print recTitulaire.NaissanceZone;
XPrt.FontBold = False
XPrt.CurrentX = 3300
XPrt.Print DicLib(71, recTitulaire.NaissanceZone);

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY)
'------------------------------------------
XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Nat.Document Identité";
XPrt.CurrentX = 2350
XPrt.Print ":";
XPrt.CurrentX = 2600
XPrt.FontBold = True
XPrt.Print recTitulaire.DocIdentitéType;
XPrt.FontBold = False
XPrt.CurrentX = 3300
XPrt.Print DicLib(73, recTitulaire.DocIdentitéType);


XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Délivré Par";
XPrt.CurrentX = 2350
XPrt.Print ":";
XPrt.CurrentX = 2600
XPrt.FontBold = True
XPrt.Print recTitulaire.DocIdentitéDélivréPar;
'------------------------------------------
XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Délivré Le";
XPrt.CurrentX = 2350
XPrt.Print ":";
XPrt.CurrentX = 2600
XPrt.FontBold = True
XPrt.Print dateImp(recTitulaire.DocIdentitéAmj);
'------------------------------------------
XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Numéro";
XPrt.CurrentX = 2350
XPrt.Print ":";
XPrt.CurrentX = 2600
XPrt.FontBold = True
XPrt.Print recTitulaire.DocIdentitéNo;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY)
'------------------------------------------
XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Forme Juridique";
XPrt.CurrentX = 2350
XPrt.Print ":";
XPrt.CurrentX = 2600
XPrt.FontBold = False
XPrt.CurrentX = 3300
XPrt.Print DicLib(3, recTitulaire.PmFormeJuridique);
'------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "N° Siren";
XPrt.CurrentX = 2350
XPrt.Print ":";
XPrt.CurrentX = 2600
XPrt.FontBold = True
XPrt.Print recTitulaire.PmSiren;
'------------------------------------------
XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Code Ape";
XPrt.CurrentX = 2350
XPrt.Print ":";
XPrt.CurrentX = 2600
XPrt.FontBold = True
XPrt.Print recTitulaire.PmApe;
XPrt.FontBold = False
XPrt.CurrentX = 3300
XPrt.Print DicLib(5, recTitulaire.PmApe);
'------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Identification Bénéficiaire";
XPrt.CurrentX = 2350
XPrt.Print ":";
XPrt.CurrentX = 2600
XPrt.FontBold = True
XPrt.Print recTitulaire.PmBdfId;
'------------------------------------------
XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Société Mère";
XPrt.CurrentX = 2350
XPrt.Print ":";
XPrt.CurrentX = 2600
XPrt.FontBold = True
XPrt.Print recTitulaire.PmNationalitéMère;
XPrt.FontBold = False
XPrt.CurrentX = 3300
XPrt.Print DicLib(19, recTitulaire.PmNationalitéMère);
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY)
'------------------------------------------

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Qualité Titulaire";
XPrt.CurrentX = 2350
XPrt.Print ":";
XPrt.CurrentX = 2600
XPrt.FontBold = True
XPrt.Print recTitulaire.PpQualité;
XPrt.FontBold = False
XPrt.CurrentX = 3300
XPrt.Print DicLib(69, recTitulaire.PpQualité);
'------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Régime Matrimonial";
XPrt.CurrentX = 2350
XPrt.Print ":";
XPrt.CurrentX = 2600
XPrt.FontBold = True
XPrt.Print recTitulaire.PpRégimeMatrimonial;
XPrt.FontBold = False
XPrt.CurrentX = 3300
XPrt.Print DicLib(70, recTitulaire.PpRégimeMatrimonial);
'------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Nombre Enfant";
XPrt.CurrentX = 2350
XPrt.Print ":";
XPrt.CurrentX = 2600
XPrt.FontBold = True
XPrt.Print recTitulaire.PpNombreEnfant;
'------------------------------------------
XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Domiciliation Clé Bdf";
XPrt.CurrentX = 2350
XPrt.Print ":";
XPrt.CurrentX = 2600
XPrt.FontBold = True
XPrt.Print recTitulaire.PpBdfId;
'------------------------------------------
XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Profession";
XPrt.CurrentX = 2350
XPrt.Print ":";
XPrt.CurrentX = 2600
XPrt.FontBold = True
XPrt.Print recTitulaire.PpProfession;
'------------------------------------------
XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Catégorie Socio Profes";
XPrt.CurrentX = 2350
XPrt.Print ":";
XPrt.CurrentX = 2600
XPrt.FontBold = True
XPrt.Print recTitulaire.PpCsp;
XPrt.FontBold = False
XPrt.CurrentX = 3300
XPrt.Print DicLib(72, recTitulaire.PpCsp);
'------------------------------------------
XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Date Entrée en France";
XPrt.CurrentX = 2350
XPrt.Print ":";
XPrt.CurrentX = 2600
XPrt.FontBold = True
XPrt.Print dateImp(recTitulaire.PpAmjEntrée);
'------------------------------------------
XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Date Sortie de France";
XPrt.CurrentX = 2350
XPrt.Print ":";
XPrt.CurrentX = 2600
XPrt.FontBold = True
XPrt.Print dateImp(recTitulaire.PpAmjSortie);
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY)
'------------------------------------------
XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Date de Création";
XPrt.CurrentX = 2350
XPrt.Print ":";
XPrt.CurrentX = 2600
XPrt.FontBold = True
XPrt.Print dateImp(recTitulaire.AmjCre);
'------------------------------------------
XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Date de Modification";
XPrt.CurrentX = 2350
XPrt.Print ":";
XPrt.CurrentX = 2600
XPrt.FontBold = True
XPrt.Print dateImp(recTitulaire.AmjMod);
'------------------------------------------
XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Date Annulation";
XPrt.CurrentX = 2350
XPrt.Print ":";
XPrt.CurrentX = 2600
XPrt.FontBold = True
XPrt.Print dateImp(recTitulaire.AmjAnn);
'------------------------------------------
XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Date de Réactivation";
XPrt.CurrentX = 2350
XPrt.Print ":";
XPrt.CurrentX = 2600
XPrt.FontBold = True
XPrt.Print dateImp(recTitulaire.AmjRea);
'------------------------------------------
XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Nom Opérateur";
XPrt.CurrentX = 2350
XPrt.Print ":";
XPrt.CurrentX = 2600
XPrt.FontBold = True
XPrt.Print recTitulaire.NOMOP;







End Sub

Public Sub prtTitulaireX()
Dim K As Integer
On Error GoTo prtError

Set XPrt = Printer


frmElpPrt.Show vbModeless

prtOrientation = vbPRORPortrait
prtTitleText = "Fiche Titulaire"
prtPgmName = "prtRacine"
prtTitleUsr = usrName

prtLineNb = 1
prtlineHeight = 250
prtHeaderHeight = 300

frmElpPrt.prtStdInit
prtTitulaireForm

For K = 1 To arrTitulaireNb
recTitulaire = arrTitulaire(K)
prtTitulaireLine

DoEvents: If prtKillDoc Then frmElpPrt.Hide: Exit Sub
prtCurrentY = prtMaxY

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
