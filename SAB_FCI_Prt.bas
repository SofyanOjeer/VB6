Attribute VB_Name = "prtSAB_FCI"
'-----------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim X As String, I As Integer, Height8_6 As Integer

Dim blnNewPage As Boolean, blnOpen As Boolean

 
Dim prtRéférenceY As Integer, prtCorpsY As Integer
Dim xDocRéférence As String

Dim meZADRESS0 As typeZADRESS0
Dim wREF As String
Dim xAMJ_Print As String

Dim meZFCIGCO0 As typeZFCIGCO0
Public Sub prtSAB_FCI_Close()
On Error GoTo prtError

blnOpen = False
Call frmElpPrt.prtEndDoc(1000)
frmElpPrt.Hide
Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub

'---------------------------------------------------------
Public Sub prtSAB_FCI_form()
'---------------------------------------------------------
Dim X As String

If Not blnOpen Then prtSAB_FCI_Open
If blnNewPage Then frmElpPrt.prtNewPage   'XPrt.NewPage
blnNewPage = True

XPrt.DrawWidth = 1
XPrt.FontSize = 10: XPrt.FontBold = False

XPrt.CurrentX = prtMinX + 6800
XPrt.CurrentY = prtMinY + prtlineHeight * 4

XPrt.Print "Paris, le  " & meZFCIGCO0.FCIGCODAJ;

Call prtAdresse_Enveloppe(meZADRESS0)

XPrt.FontSize = 8
XPrt.FontBold = False
XPrt.CurrentY = prtRéférenceY
XPrt.CurrentX = prtMinMarge: XPrt.Print "N/Référence " & meZFCIGCO0.FCIGCOCPT & " / " & meZFCIGCO0.FCIGCONUC

XPrt.CurrentX = prtMinMarge + 1250: XPrt.Print ":";
XPrt.FontBold = True

XPrt.FontSize = 10: XPrt.FontBold = False

End Sub
Public Sub prtSAB_FCI_Open()
On Error GoTo prtError
blnOpen = True
blnNewPage = False
Set XPrt = Printer
frmElpPrt.Show vbModeless
XPrt.FontItalic = False

Height8_6 = frmElpPrt.prtHeightDelta(8, 6)

blnFiligrane = False
prtPgmName = "prtSAB_FCI"
prtTitleUsr = usrName
prtOrientation = vbPRORPortrait
prtTitleText = "FCI_Courrier"
prtFontName = prtFontName_Arial   'TimesNewRoman '


prtLineNb = 1
prtlineHeight = 250
prtHeaderHeight = 50 ' 100

prtFormType = ""
prtSocInit

prtRéférenceY = prtMinY + prtlineHeight * 9
prtCorpsY = prtMinY + prtlineHeight * 18


Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub

Public Sub prtSAB_FCI_Monitor(lZFCIGCO0 As typeZFCIGCO0)
Dim I As Integer
Dim X As String
blnOpen = False

meZFCIGCO0 = lZFCIGCO0
rsZADRESS0_Init meZADRESS0
If Trim(lZFCIGCO0.FCIGCOCPT) <> "" Then
    meZADRESS0.ADRESSNUM = lZFCIGCO0.FCIGCOCPT     ' String * 20                    ' ou numéro de client
Else
    meZADRESS0.ADRESSNUM = lZFCIGCO0.FCIGCOCLI     ' String * 20                    ' ou numéro de client
End If
If Trim(lZFCIGCO0.FCIGCONRD) <> "" Then
    meZADRESS0.ADRESSRA1 = lZFCIGCO0.FCIGCONRD     ' String * 32                    ' ou raison sociale 1
    meZADRESS0.ADRESSRA2 = lZFCIGCO0.FCIGCOPRD     ' String * 32                    ' ou raison sociale 2
    meZADRESS0.ADRESSAD1 = lZFCIGCO0.FCIGCOA1D     ' String * 32                    ' Adresse 1
    meZADRESS0.ADRESSAD2 = lZFCIGCO0.FCIGCOA2D      ' String * 32                    ' Adresse 2
    meZADRESS0.ADRESSAD3 = lZFCIGCO0.FCIGCOA3D    ' String * 32                    ' Adresse 3
    meZADRESS0.ADRESSCOP = lZFCIGCO0.FCIGCOCPD     ' String * 6                     ' Code postal
    meZADRESS0.ADRESSVIL = lZFCIGCO0.FCIGCOVID    ' String * 25                    ' Ville
    meZADRESS0.ADRESSPAY = lZFCIGCO0.FCIGCOLPD      ' String * 25                    ' Pays
Else
    meZADRESS0.ADRESSRA1 = lZFCIGCO0.FCIGCONOT     ' String * 32                    ' ou raison sociale 1
    meZADRESS0.ADRESSRA2 = lZFCIGCO0.FCIGCOPRT     ' String * 32                    ' ou raison sociale 2
    meZADRESS0.ADRESSAD1 = lZFCIGCO0.FCIGCO1DT     ' String * 32                    ' Adresse 1
    meZADRESS0.ADRESSAD2 = lZFCIGCO0.FCIGCO2DT      ' String * 32                    ' Adresse 2
    meZADRESS0.ADRESSAD3 = lZFCIGCO0.FCIGCO3DT    ' String * 32                    ' Adresse 3
    meZADRESS0.ADRESSCOP = lZFCIGCO0.FCIGCOPOT     ' String * 6                     ' Code postal
    meZADRESS0.ADRESSVIL = lZFCIGCO0.FCIGCOVIT    ' String * 25                    ' Ville
    meZADRESS0.ADRESSPAY = lZFCIGCO0.FCIGCOLPT      ' String * 25                    ' Pays
End If

prtSAB_FCI_form
prtSAB_FCI_Détail

If blnOpen Then prtSAB_FCI_Close

End Sub




Public Sub prtSAB_FCI_Détail()
Dim wCol1 As Long
prtFontName = prtFontName_CourierNew
XPrt.FontBold = False
XPrt.FontSize = 7
wCol1 = prtMinX
XPrt.CurrentY = prtCorpsY
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOETA    4S";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "ETABLISSEMENT";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOETA;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 2
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOCLI    7A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "RESPONSABLE";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOCLI;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 3
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOPLA    3S";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "NUMERO PLAN";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOPLA;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 4
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOCPT   20A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "NUMERO COMPTE";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOCPT;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 5
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCONUC    7S";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "NUMERO CHEQUE";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCONUC;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 6
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOCAR   16A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "NUMERO CARTE";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOCAR;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 7
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOSES    5S";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "NUM¢ SEQUENCE STATUT";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOSES;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 8
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOSEA    3S";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "NUM¢ SEQUENCE ACTION";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOSEA;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 9
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCODLI   10A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "DATE LIMITE RETENT¢";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCODLI;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 10
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCODAJ   10A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "DATE JOUR";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCODAJ;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 12
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOCOU    6A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "CODE COURRIER TRANSM";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOCOU;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 13
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOLIB   30A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "LIBELLE COURRIER";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOLIB;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 14
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOTYC    1A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "TYPE COURRIER TRANSM";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOTYC;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 15
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOLTY   30A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "LIBELLE TYPE COURR.";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOLTY;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 16
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOENV    1A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "ENVOI RECOMANDE";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOENV;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 17
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOREC   30A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "LIBELLE RECOMMANDE";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOREC
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 18
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCODCP   10A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "DATE COURRIER PRECED";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCODCP;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 19;
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOEDI   10A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "DATE EDITION";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOEDI;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 20;
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCORED    1A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "REEDITION (O/N)";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCORED;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 21;
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCONDE    7A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "NUMERO CLIENT DESTIN";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCONDE;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 22;
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOLTD   30A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "LIBELLE ETAT DESTINA";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOLTD;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 23;
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCONRD   32A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "NOM/RAISON DESTINATA";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCONRD;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 24;
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOPRD   32A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "PRENOM/RAISON DESTIN";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOPRD;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 25;
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOA1D   32A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "ADRESSE 1 DESTINATAT";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOA1D;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 26;
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOA2D   32A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "ADRESSE 2 DESTINATAT";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOA2D;

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 27;
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOA3D   32A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "ADRESSE 3 DESTINATAT";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOA3D;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 28;
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOCPD    6A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "CODE POSTAL DESTINAT";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOCPD;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 29;
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOVID   25A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "VILLE DESTINATAIRE";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOVID;

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 30;
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOLPD   25A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "LIBELL.PAYS DESTINAT";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOLPD;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 31;
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOCLD    7A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "NUMERO CLIENT";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOCLD;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 32
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOLTC   30A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "LIBELLE ETAT CLIENT";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOLTC;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 33;
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCONRC   32A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "NOM/RAISON CLIENT";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCONRC
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 34;
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOPRC   32A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "PRENOM/RAISON CLIENT";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOPRC;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 35;
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOAD1   32A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "ADRESSE 1 CLIENT";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOAD1;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 36;
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOAD2   32A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "ADRESSE 2 CLIENT";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOAD2;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 37;
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOAD3   32A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "ADRESSE 3 CLIENT";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOAD3;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 38;
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOCPC    6A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "CODE POSTAL CLIENT";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOCPC;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 39;
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOVIC   25A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "VILLE CLIENT";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOVIC;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 40
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOLPC   25A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "LIBELL.PAYS CLIENT";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOLPC;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 41
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOCLB    7A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "NUMERO CLIENT BENEFI";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOCLB;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 42
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOLTB   30A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "LIBEL ETAT CLI BENEF";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOLTB;
'________________________________________________________________________________________________
XPrt.CurrentY = prtCorpsY
wCol1 = prtMinX + 5500



XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 43
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOBNR   32A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "NOM/RAIS CLI BENENEF";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOBNR;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 44
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOBPR   32A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "PRENOM/RAIS CLI BENE";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOBPR;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 45;;
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOA1B   32A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "ADRESSE 1 BENEFICIAC";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOA1B;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 46;
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOA2B   32A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "ADRESSE 2 BENEFICIAC";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOA2B;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 47;
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOA3B   32A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "ADRESSE 3 BENEFICIAC";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOA3B;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 48
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOCPB    6A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "CODE POSTAL BENEFICI";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOCPB;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 49
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOBVI   25A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "VILLE BENEFICIAIRE";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOBVI;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 50
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOLPB   25A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "LIBELL.PAYS BENEFIC";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOLPB;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 51
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOCLP    7A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "N¢ CLIENT PORTEUR";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOCLP;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 52
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOLTP   30A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "LIBELLE ETAT PORTEUR";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOLTP;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 53
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCONRP   32A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "NOM/RAISON PORTEUR";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCONRP;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 54
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOPRP   32A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "PRENOM/RAISON PORTEU";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOPRP;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 55
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCO1DP   32A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "ADRESSE 1 PORTEUR";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCO1DP;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 56
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCO2DP   32A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "ADRESSE 2 PORTEUR";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCO2DP;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 57
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCO3DP   32A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "ADRESSE 3 PORTEUR";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCO3DP;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 58
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOCPP    6A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "CODE POSTAL PORTEUR";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOCPP;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 59;
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOVIP   25A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "VILLE PORTEUR";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOVIP;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 60;
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOLPP   25A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "LIBELL.PAYS PORTEUR";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOLPP;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 61
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOCLT    7A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "N¢ CLIENT TITULAIRE";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOCLT;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 62
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOLIT   30A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "LIBELLE ETAT TITULAI";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOLIT;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 63
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCONOT   32A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "NOM/RAISON TITULALAI";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCONOT;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 64
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOPRT   32A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "PRENOM/RAISON TITULA";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOPRT;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 65;
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCO1DT   32A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "ADRESSE 1 TITULAIRE";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCO1DT;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 66
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCO2DT   32A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "ADRESSE 2 TITULAIRE";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCO2DT;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 67
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCO3DT   32A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "ADRESSE 3 TITULAIRE";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCO3DT;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 68
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOPOT    6A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "CODE POSTAL TITULAIR";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOPOT;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 69
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOVIT   25A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "VILLE TITULAIRE";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOVIT;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 70
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOLPT   25A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "LIBELL.PAYS TITULAIR";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOLPT;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 71
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOCLC    7A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "N¢ CLIENT COTITULAIR";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOCLC;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 72
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOLIC   30A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "LIBELLE ETAT COTITUL";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOLIC;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 73
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCONOC   32A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "NOM/RAISON COTITULAL";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCONOC;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 74
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOPCO   32A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "PRENOM/RAISON COTITU";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOPCO;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 75
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCO1DC   32A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "ADRESSE 1 COTITULAIR";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCO1DC;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 76
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCO2DC   32A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "ADRESSE 2 COTITULAIR";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCO2DC;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 77
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCO3DC   32A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "ADRESSE 3 COTITULAIR";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCO3DC;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 78
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOPOC    6A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "CODE POSTAL COTITULA";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOPOC;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 79
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOVLC   25A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "VILLE COTITULAIRE";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOVLC;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 80
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOPAC   25A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "LIBELL.PAYS COTITULA";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOPAC;

'________________________________________________________________________________________________
frmElpPrt.prtNewPage
XPrt.CurrentY = 1500
XPrt.CurrentX = prtMinMarge: XPrt.Print "N/Référence " & meZFCIGCO0.FCIGCOCPT & " / " & meZFCIGCO0.FCIGCONUC

XPrt.CurrentY = 2000
wCol1 = prtMinX
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 81
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOCLM    7A";

XPrt.CurrentX = wCol1 + 1300: XPrt.Print "N¢ CLIENT MANDATAIRE";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOCLM;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 82
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOLIM   30A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "LIBELLE ETAT MANDATA";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOLIM;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 83
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCONOM   32A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "NOM/RAISON MANDATAIR";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCONOM;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 84
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOPRM   32A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "PRENOM/RAISON MANDAT";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOPRM;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 85
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCO1DM   32A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "ADRESSE 1 MANDATAIRE";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCO1DM;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 86
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCO2DM   32A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "ADRESSE 2 MANDATAIRE";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCO2DM;

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 87
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCO3DM   32A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "ADRESSE 3 MANDATAIRE";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCO3DM;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 88
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOCPM    6A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "CODE POSTAL MANDATAI";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOCPM;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 89
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOVIM   25A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "VILLE MANDATAIRE";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOVIM;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 90
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOLPM   25A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "LIBELL.PAYS MANDATAI";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOLPM;

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 91
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOCLG    7A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "N¢ CLIENT GREFFE TRI";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOCLG;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 92
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOLIG   30A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "LIBELLE ETAT GREFFE";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOLIG;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 93
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCONOG   32A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "NOM/RAISON GREFFE";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCONOG;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 94
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOPRG   32A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "PRENOM/RAISON GREFFE";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOPRG;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 95
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCO1DG   32A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "ADRESSE 1 GREFFE";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCO1DG;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 96
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCO2DG   32A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "ADRESSE 2 GREFFE";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCO2DG;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 97
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCO3DG   32A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "ADRESSE 3 GREFFE";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCO3DG;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 98
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOCPG    6A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "CODE POSTAL GREFFE";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOCPG;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 99
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOVIG   25A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "VILLE GREFFE";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOVIG;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 100
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOLPG   25A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "LIBELL.PAYS GREFFE";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOLPG;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 101
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOLED   30A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "LIEU EDITION";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOLED;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 102
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOGES   32A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "NOM GESTIONNAIRE";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOGES;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 103
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOREL   32A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "REFERENCE LIBRE";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOREL;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 104
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOTEL   20A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "TELEPHONE GESTIONNAI";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOTEL;

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 105
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOREJ    6A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "CODE REJET";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOREJ;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 106;
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOLIR   30A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "LIBELLE REJET";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOLIR;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 107
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOMCH   20A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "REJETE";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOMCH;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 108
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCODEV    3A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "DEVISE MONTANT";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCODEV;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 109
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOAT1    1A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "CAS N¢1 ATTESTION";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOAT1;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 110
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOAT2    1A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "CAS N¢2 ATTESTION";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOAT2;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 111
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOAT3    1A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "CAS N¢3 ATTESTION";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOAT3;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 112
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOAT4    1A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "NON UTILISE PREVISIONN";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOAT4;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 113
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOIJ1    1A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "CAS N¢1 INJONCTION";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOIJ1;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 114
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOIJ2    1A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "CAS N¢2 INJONCTION";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOIJ2;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 115
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOIJ3    1A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "CAS N¢3 INJONCTION";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOIJ3;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 116
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOIJ4    1A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "NON UTILISE PREVISIONN";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOIJ4;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 117
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOMSD   20A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "MONTANT SOLDE DISPON";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOMSD;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 118
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCODES    3A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "DEVISE SOLDE DISPONI";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCODES;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 119
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCODEB   12A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "MENTION DEBITEUR";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCODEB;

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 120
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOIC1    1A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "CAS N¢1 PAS PAYE";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOIC1;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 121
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOIC2    1A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "CAS N¢2 PAYE PARTIEL";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOIC2;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 122
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOMPP   20A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "MONTANT PAIEMT PARTI";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOMPP;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 123
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCODPP    3A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "DEVISE MT PAIT PARTI";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCODPP;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 124
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOCH1    7A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "NUMERO CHEQUE 1";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOCH1;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 125
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOMC1   20A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "MONTANT CHEQUE 1";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOMC1;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 126
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCODE1    3A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "DEVISE MONTANT 1";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCODE1;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 127
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOCH2    7A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "NUMERO CHEQUE 2";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOCH2;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 128
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOMC2   20A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "MONTANT CHEQUE 2";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOMC2;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 129
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCODE2    3A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "DEVISE MONTANT 2";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCODE2;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 130
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOCH3    7A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "NUMERO CHEQUE 3";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOCH3;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 131
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOMC3   20A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "MONTANT CHEQUE 3";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOMC3;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 132
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCODE3    3A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "DEVISE MONTANT 3";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCODE3;

'________________________________________________________________________________________________
XPrt.CurrentY = 2000
wCol1 = prtMinX + 5500

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 133
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOCH4    7A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "NUMERO CHEQUE 4";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOCH4;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 134
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOMC4   20A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "MONTANT CHEQUE 4";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOMC4;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 135
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCODE4    3A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "DEVISE MONTANT 4";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCODE4;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 136
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOCH5    7A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "NUMERO CHEQUE 5";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOCH5;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 137
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOMC5   20A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "MONTANT CHEQUE 5";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOMC5;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 138
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCODE5    3A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "DEVISE MONTANT 5";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCODE5;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 139
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOCH6    7A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "NON UTILISE PREVISIO";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOCH6;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 140
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOMC6   20A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "NON UTILISE PREVISIO";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOMC6;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 141
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCODE6    3A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "NON UTILISE PREVISIO";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCODE6;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 142
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCODRJ   10A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "DATE REJET DES CHQ";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCODRJ;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 143
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCODEI   10A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "DATE DEPART INTERDIT";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCODEI;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 144
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCODLR   10A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "DATE LIMITE REGULARI";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCODLR;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 145
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOMPN   20A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "MONTANT PENALITE";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOMPN;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 146
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCODPN    3A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "DEVISE MT PENALITE";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCODPN;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 147
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOJ21    1A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "CAS N1 INJ 2 EV PREC";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOJ21;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 148
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOJ22    1A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "CAS N2 INJ 2 EV PREC";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOJ22;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 149
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOINT   30A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "INTITULE COMPTE";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOINT;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 150
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCONAG   32A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "NOM AGENCE";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCONAG;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 151
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCODAP   10A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "DU CHEQUE";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCODAP;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 152
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOMIM   20A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "MONTANT IMPAYE";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOMIM;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 153
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCODIM    3A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "DEVISE MT IMPAYE";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCODIM;

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 154
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOCHB    7A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "FRAIS PUBLICITE";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOCHB;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 155
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOMCB   20A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "MONTANT CHQ BANQUE";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOMCB;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 156
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCODCH    3A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "DEVISE MT CHQ BANQUE";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCODCH;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 157
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCONBQ   32A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "NOM BANQUE";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCONBQ;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 158
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOMTA   20A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "MONTANT ABUSIF";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOMTA;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 159
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCODEA    3A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "DEVISE MONT. ABUSIF";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCODEA;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 160
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOLCA   32A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "LIBELLE TYPE CARTE";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOLCA;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 161
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOLNA   32A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "LIBELLE NATURE CARTE";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOLNA;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 162
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOVAL   10A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "DATE VALIDITE";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOVAL;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 163
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCODUR    2A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "DUREE VALIDITE";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCODUR;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 164
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOLI1   32A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "CHAMP LIBRE 1";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOLI1;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 165
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOLI2   32A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "CHAMP LIBRE 2";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOLI2;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 166
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOPUP   10A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "DATE PURGE POSSIBLE";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOPUP;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 167
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOTYI   30A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "TYPE INTERDIT";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOTYI;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 168
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCODDI   10A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "TION INTERNE";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCODDI;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 169
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCODFI   10A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "ON INTERNE";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCODFI;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 170
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCODDB   10A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "TION BANCAIRE";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCODDB;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 171
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCODFB   10A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "DATE FIN INTERDICTI-";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCODFB;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 172
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOACP   20A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "COMPTE AV CONVERSION";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOACP;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight  '.... 173
XPrt.CurrentX = wCol1: XPrt.Print "FCIGCOIBA   20A";
XPrt.CurrentX = wCol1 + 1300: XPrt.Print "NUMERO IBAN";
XPrt.CurrentX = wCol1 + 3200: XPrt.Print meZFCIGCO0.FCIGCOIBA;

End Sub
