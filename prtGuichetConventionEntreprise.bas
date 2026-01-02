Attribute VB_Name = "prtGuichetConventionEntreprise"
Option Explicit
'Dim I As Integer
'---------------------------------------------------------
Public Sub prtGuichetConventionEntrepriseForm()
'---------------------------------------------------------
Dim X As String
XPrt.FontSize = 8
XPrt.FontBold = False

XPrt.FillStyle = 0
XPrt.DrawWidth = 3
XPrt.ForeColor = RGB(0, 0, 0)
XPrt.FillStyle = 1

XPrt.Line (prtMinX, prtMinY + prtHeaderHeight)-(prtMaxX, prtMinY + prtHeaderHeight)
XPrt.DrawWidth = 3

Call frmElpPrt.prtTrame(prtMinX, prtMinY, prtMaxX, prtMinY + prtHeaderHeight, "B")

Call frmElpPrt.prtTrame(200, 8800, 11100, 9100, "B")
Call frmElpPrt.prtTrame(200, 10700, 11100, 11000, "B")
Call frmElpPrt.prtTrame(200, 13000, 11100, 13300, "B")
'Call frmElpPrt.prtTrame(200, 1700, 11100, 3500, "")
'Call frmElpPrt.prtTrame(200, 5900, 11100, 7700, "")

XPrt.DrawWidth = 1

'---------------------------------------------------------
XPrt.FontSize = 8
XPrt.FontBold = True
XPrt.CurrentY = prtMinY + (prtHeaderHeight - XPrt.TextHeight(X)) / 2
frmElpPrt.prtCentré prtMinX + 5400, "CONVENTION DE COMPTE ENTREPRISE"
XPrt.FontBold = False



prtCurrentY = prtMinY + prtHeaderHeight

End Sub
Public Sub prtGuichetConventionEntrepriseLine()
Dim X As String, K As Integer
Dim Situation As String
'Dim Titulaire1Civilité As String, Titulaire2Civilité As String
'Titulaire1Civilité = ""
'Titulaire2Civilité = ""

XPrt.FontBold = False
XPrt.FontSize = 8
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight - XPrt.TextHeight("test")
'------------------------------------------1-----------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight - 100
XPrt.CurrentX = 300
XPrt.Print "Raison Social";
XPrt.FontSize = 9
XPrt.FontBold = True
XPrt.CurrentX = 2400
XPrt.Print ":";
XPrt.FontSize = 9
XPrt.FontBold = True
XPrt.CurrentX = 2500
XPrt.Print recGuichetConventionEntreprise.RaisonSociale;
'-----------------------------------------3-----
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.FontSize = 8
XPrt.FontBold = False
XPrt.CurrentX = 300
XPrt.Print "Dénomination Commerciale";
XPrt.CurrentX = 2400
XPrt.Print ":";
XPrt.FontSize = 9
XPrt.FontBold = True
XPrt.CurrentX = 2500
XPrt.Print recGuichetConventionEntreprise.DénominationCommerciale;
'-----------------------------------------------4------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.FontSize = 8
XPrt.FontBold = False
XPrt.CurrentX = 300
XPrt.Print "Objet Social";
XPrt.CurrentX = 2400
XPrt.Print ":";
XPrt.CurrentX = 2500
XPrt.FontSize = 9
XPrt.FontBold = True
XPrt.Print recGuichetConventionEntreprise.objetSocial;

'----------------------------------------------5-------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.FontSize = 8
XPrt.FontBold = False
XPrt.CurrentX = 300
XPrt.Print "Date de Création";
XPrt.CurrentX = 2400
XPrt.Print ":";
XPrt.FontSize = 9
XPrt.FontBold = True
XPrt.CurrentX = 2500
XPrt.Print dateImp(recGuichetConventionEntreprise.CréationDate);

'----------------------------------------------5-------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.FontSize = 8
XPrt.FontBold = False
XPrt.CurrentX = 300
XPrt.Print "Forme Juridique";
XPrt.CurrentX = 2400
XPrt.Print ":";
XPrt.CurrentX = 2500
XPrt.FontSize = 9
XPrt.FontBold = True
XPrt.Print recGuichetConventionEntreprise.FormeJuridique;
'----------------------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.FontSize = 8
XPrt.FontBold = False
XPrt.CurrentX = 300
XPrt.Print "Capital Social";
XPrt.CurrentX = 2400
XPrt.Print ":";
XPrt.CurrentX = 2500
XPrt.FontSize = 9
XPrt.FontBold = True
XPrt.Print recGuichetConventionEntreprise.CapitalSocial;

'----------------------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.FontSize = 8
XPrt.FontBold = False
XPrt.CurrentX = 300
XPrt.Print "Numéro de Siren";
XPrt.CurrentX = 2400
XPrt.Print ":";
XPrt.CurrentX = 2500
XPrt.FontSize = 9
XPrt.FontBold = True
XPrt.Print recGuichetConventionEntreprise.NuméroSiren;

'----------------------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.FontSize = 8
XPrt.FontBold = False
XPrt.CurrentX = 300
XPrt.Print "Code APE";
XPrt.CurrentX = 2400
XPrt.Print ":";
XPrt.CurrentX = 2500
XPrt.FontSize = 9
XPrt.FontBold = True
XPrt.Print recGuichetConventionEntreprise.CodeApe;

'----------------------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.FontSize = 8
XPrt.FontBold = False
XPrt.CurrentX = 300
XPrt.Print "Pays d'origine";
XPrt.CurrentX = 2400
XPrt.Print ":";
XPrt.CurrentX = 2500
XPrt.FontSize = 9
XPrt.FontBold = True
XPrt.Print recGuichetConventionEntreprise.OriginePays;
'----------------------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.FontSize = 8
XPrt.FontBold = False
XPrt.CurrentX = 300
XPrt.Print "Pays de résidence";
XPrt.CurrentX = 2400
XPrt.Print ":";
XPrt.CurrentX = 2500
XPrt.FontSize = 9
XPrt.FontBold = True
XPrt.Print recGuichetConventionEntreprise.Résidencepays;
'-----------------------------------------------------------16
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.FontBold = False
XPrt.FontSize = 8
XPrt.CurrentX = 300
XPrt.Print "Adresse Courrier";
XPrt.CurrentX = 2400
XPrt.Print ":";
XPrt.FontSize = 9
XPrt.FontBold = True
XPrt.CurrentX = 2500
XPrt.Print recGuichetConventionEntreprise.AdresseCourrier1;
'-----------------------------------------------------------17
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 2500
XPrt.Print recGuichetConventionEntreprise.AdresseCourrier2;
'------------------------------------------------------------18
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 2500
XPrt.Print recGuichetConventionEntreprise.AdresseCourrier3;
'------------------------------------------------------------18
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 2500
XPrt.Print recGuichetConventionEntreprise.AdresseCourrier4;





XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
End Sub
Public Sub prtGuichetConventionEntreprisetrait()


'XPrt.Line (prtMinX + 5400, prtMinY + 300)-(prtMinX + 5400, prtMaxY - 7025)


End Sub



'---------------------------------------------------------
 Public Sub prtGuichetConventionEntrepriseX()
'---------------------------------------------------------
Dim K As Integer, K1 As Integer, K2 As Integer
Dim X As String

Set XPrt = Printer
frmElpPrt.Show vbModeless

prtFormType = "+"
prtOrientation = vbPRORPortrait
prtTitleText = ""
prtPgmName = "prtGuichetConventionEntreprise"
prtTitleUsr = usrName

prtLineNb = 1
prtlineHeight = 300
prtHeaderHeight = 300
prtFormType = "SOC"

frmElpPrt.prtInit
prtsoc
prtMinY = 1000
prtGuichetConventionEntrepriseForm
prtGuichetConventionEntrepriseLine
prtGuichetConventionEntreprisetrait

frmElpPrt.prtEndDoc
frmElpPrt.Hide

End Sub





