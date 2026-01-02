Attribute VB_Name = "prtVir"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim recCptInfo As typeCptInfo
Dim recOpCpt As typeOpCpt
Dim recBic As typeBic
Dim recBdfG As typeBdfG

Dim I As Integer
Dim blnPrinter As Boolean



'----------------------------------
Public Sub prtVirx(Msg As String)
'---------------------------------------------------------
blnPrinter = IIf(Trim(Msg) = "picVirement", False, True)
recOpCpt = XopCpt(2)
recCptInfoInit recCptInfo
recCptInfo.Method = "JoinL1"
recCptInfo.Société = recOpCpt.Société
recCptInfo.Agence = recOpCpt.Agence
recCptInfo.Devise = recOpCpt.Devise
recCptInfo.Numéro = XopCpt(1).Compte
recCptInfo.BiaTyp = "000"
recCptInfo.BiaNum = "00000"
recCptInfo.NuméroAncien = "00000000000"
If IsNull(srvCptInfoFind(recCptInfo)) Then
    recBicInit recBic
    recBic.Method = "SeekP0"
    recBic.BicId = "ISBKTRIS"
    If IsNull(srvBicFind(recBic)) Then
        recBdfGInit recBdfG
        recBdfG.Method = "SeekP0"
        recBdfG.BdfG2 = Mid$(recOpCpt.Libellé, 28, 5)
        recBdfG.BdfG3 = Mid$(recOpCpt.Libellé, 33, 5)
        srvBdfGFind recBdfG
    
        prtOrientation = vbPRORPortrait
        prtTitleText = "Vir"
        prtPgmName = "prtVir"
        prtTitleUsr = usrName
        
        prtLineNb = 1: prtlineHeight = 250: prtHeaderHeight = 300
        prtFormType = ""
        If blnPrinter Then
            Set XPrt = Printer
            frmElpPrt.Show vbModeless
        
            frmElpPrt.prtInit
            Call prtVirForm(0, "VIREMENT")
            prtVirFin2 0
            prtVirLine 0
            XPrt.CurrentY = 5150
            frmElpPrt.prtTiret
            Call prtVirForm(5350, "AVIS DE CREDIT")
            prtVirFin3 5350
            prtVirLine 5350
            XPrt.CurrentY = 10650
            frmElpPrt.prtTiret
            Call prtVirForm(10900, "VIREMENT INTERBANCAIRE")
            prtVirFin4 10900
            prtVirLine 10900
            frmElpPrt.prtEndDoc
            frmElpPrt.Hide
        Else
            Call prtVirForm(0, "VIREMENT")
            prtVirFin2 0
            prtVirLine 0
        End If
    End If
End If
End Sub

'----------------------------------
Public Sub prtVirForm(H As Integer, titre As String)
'----------------------------------
Dim X As String, I As Integer
'---------------------------BOXES--------------------
XPrt.FillStyle = 1
XPrt.DrawWidth = 1
Call frmElpPrt.prtTrame(0, H + 400, 4000, H + 800, "B", 250)
Call frmElpPrt.prtTrame(4200, H + 400, 8900, H + 800, "B", 250)
XPrt.DrawWidth = 8
Call frmElpPrt.prtTrame(9000, H + 400, 11000, 1000 + H, "B", 250)

Call frmElpPrt.prtTrame(0, 1300 + H, 4000, 2500 + H, "B", 250)

Call frmElpPrt.prtTrame(0, 2800 + H, 4000, 3300 + H, "B", 250)
XPrt.DrawWidth = 1
XPrt.Line (4800, 2800 + H)-(10000, 4300 + H), , B

  

                                                            
'----------------TRAITS VERTICAUX---------------------
XPrt.Line (1000, 400 + H)-(1000, 800 + H)
XPrt.Line (0, 400 + H)-(0, 1300 + H)
XPrt.Line (1000, 400 + H)-(1000, 1300 + H)
XPrt.Line (2100, 400 + H)-(2100, 800 + H)
XPrt.Line (3500, 400 + H)-(3500, 800 + H)
XPrt.Line (5400, 400 + H)-(5400, 800 + H)
XPrt.Line (6400, 400 + H)-(6400, 800 + H)
XPrt.Line (7300, 400 + H)-(7300, 800 + H)
XPrt.Line (8400, 400 + H)-(8400, 800 + H)
XPrt.Line (850, 2800 + H)-(850, 3300 + H)
XPrt.Line (1900, 2800 + H)-(1900, 3300 + H)
XPrt.Line (3500, 2800 + H)-(3500, 3300 + H)

'----------------------petits traits verticaux----------
XPrt.Line (4800, 1250 + H)-(4800, 1500 + H)
XPrt.Line (10000, 1250 + H)-(10000, 1500 + H)
XPrt.Line (4800, 2400 + H)-(4800, 2600 + H)
XPrt.Line (10000, 2400 + H)-(10000, 2600 + H)
XPrt.Line (10400, 2900 + H)-(10400, 3300 + H)
XPrt.Line (11100, 2900 + H)-(11100, 3300 + H)
'---------------------petit trait horizontaux------------
XPrt.Line (4800, 1250 + H)-(5000, 1250 + H)
XPrt.Line (9800, 1250 + H)-(10000, 1250 + H)
XPrt.Line (4800, 2600 + H)-(5000, 2600 + H)
XPrt.Line (9800, 2600 + H)-(10000, 2600 + H)


XPrt.Line (10400, 3300 + H)-(11100, 3300 + H)


XPrt.DrawWidth = 2
XPrt.FontSize = 7
XPrt.FontBold = True
XPrt.CurrentX = 1300
XPrt.Print " ";


XPrt.FontSize = 7
XPrt.FontBold = True
XPrt.CurrentX = 1300: XPrt.CurrentY = H

XPrt.Print "BANQUE INTERCONTINENTALE ARABE";
XPrt.FontSize = 10
XPrt.FontBold = True
XPrt.CurrentX = 6200
XPrt.Print titre;

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight / 2 + 50
XPrt.FontBold = False
XPrt.FontSize = 6
XPrt.CurrentX = 50
XPrt.Print "ORDONNATEUR :";
XPrt.FontSize = 6
XPrt.CurrentX = 1200
XPrt.Print "R.C. Paris B 302590070 - L.B.E. 116 - 75008 PARIS";


'------------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.FontSize = 6
XPrt.FontBold = False
XPrt.CurrentX = 100
XPrt.Print "CODE ETAB.";
XPrt.CurrentX = 1100
XPrt.Print "CODE GUICHET";
XPrt.CurrentX = 2400
XPrt.Print "N° DE COMPTE";
XPrt.CurrentX = 3550
XPrt.Print "Clé RIB";
XPrt.CurrentX = 4300
XPrt.Print "N° EMETTEUR";
XPrt.CurrentX = 5700
XPrt.Print "REF. 1";
XPrt.CurrentX = 6700
XPrt.Print "REF. 2";
XPrt.CurrentX = 7600
XPrt.Print "DATE";
XPrt.CurrentX = 8500
XPrt.Print "CAT .";
XPrt.FontSize = 10
XPrt.FontBold = True
XPrt.CurrentX = 9100
XPrt.Print "FRF";
XPrt.FontSize = 7
XPrt.FontBold = False

XPrt.CurrentX = 10000
XPrt.Print "MONTANT";
'---------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight / 2
XPrt.FontSize = 10
XPrt.FontBold = True
XPrt.CurrentX = 150
 XPrt.Print strSocBdfE;
XPrt.CurrentX = 1300
XPrt.Print strSocBdfG;
'---------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight + 50
XPrt.FontSize = 5
XPrt.FontBold = False
XPrt.CurrentX = 50
XPrt.Print "PRESENTATEUR";
XPrt.FontSize = 6
XPrt.FontBold = False

XPrt.CurrentX = 1400
XPrt.Print "MOTIF DE L'OPERATION";
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight / 2
XPrt.FontSize = 10
XPrt.CurrentX = 150
XPrt.Print "30003";

'---------------------------------------------
XPrt.FontSize = 6
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.CurrentX = 850
XPrt.Print "RAISON SOCIALE - LOCALITE - GUICHET";
XPrt.FontBold = True
XPrt.CurrentX = 10500
XPrt.Print "DONNEUR";
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 10500
XPrt.Print "D'ORDRE";
XPrt.FontBold = True
XPrt.CurrentX = 10100
XPrt.Print "<==";
'-----------------------------------------------
XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2 + 100
XPrt.CurrentX = 800
XPrt.Print "Etablissement tenant le compte du bénéficiaire";
'-----------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.FontSize = 7
XPrt.CurrentX = 1000
XPrt.Print "SOCIETE GENERALE N° 27";
'------------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.FontSize = 6
XPrt.FontBold = False
XPrt.CurrentX = 50
XPrt.Print "CODE ETAB.";
XPrt.CurrentX = 900
XPrt.Print "CODE GUICHET";
XPrt.CurrentX = 2200
XPrt.Print "N° DE COMPTE";
XPrt.CurrentX = 3550
XPrt.Print "Clé RIB";
XPrt.FontSize = 6
XPrt.FontBold = False

XPrt.CurrentX = 10500
XPrt.Print "NUMERO";
'-----------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2 + 25
XPrt.CurrentX = 1200
XPrt.Print "RIB DU BENEFICIAIRE";
'-------------------------------------------------


'-------------------------------------------------
XPrt.FontBold = True
XPrt.CurrentX = 10300
XPrt.Print "BENEFICIAIRE";
XPrt.FontBold = True
XPrt.CurrentX = 10050
XPrt.Print "<==";


'-------------------------------------------------




    
    
End Sub




Public Sub prtVirLine(H As Integer)
Dim Mt As Currency
XPrt.DrawWidth = 1
XPrt.FontBold = True
XPrt.FontSize = 8
'------------------------------------------ligne 1--------------
XPrt.CurrentY = H + prtlineHeight * 2 + 50
XPrt.FontSize = 10
XPrt.CurrentX = 2250
XPrt.Print recCptInfo.Numéro;
XPrt.CurrentX = 3650
XPrt.Print recCptInfo.CléRib;
XPrt.FontSize = 8
XPrt.CurrentY = XPrt.CurrentY + 50
XPrt.CurrentX = 4300
XPrt.Print recOpCpt.CodeOpération;
XPrt.CurrentX = 5700
XPrt.Print recOpCpt.Référence;
XPrt.CurrentX = 7400
XPrt.Print dateImp(recOpCpt.AmjValeur);
XPrt.FontSize = 10
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight - 100
XPrt.CurrentX = 9200
Mt = recOpCpt.Brut - recOpCpt.ComMontantDevise

XPrt.Print Format$(Mt, "### ### ### ###.##");
'------------------------------------

XPrt.FontSize = 8
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2 + 50
XPrt.CurrentX = 5000
XPrt.Print recCptInfo.Intitulé;
XPrt.FontBold = False

If Trim(recCptInfo.Adresse2) <> "" Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = 5000
    XPrt.Print recCptInfo.Adresse2;
End If
'------------------------------------3---------------
If Trim(recCptInfo.Adresse3) <> "" Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = 5000
    XPrt.Print recCptInfo.Adresse3;
End If
'----------------------------------4-------------------
If Trim(recCptInfo.Adresse4) <> "" Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = 5000
    XPrt.Print recCptInfo.Adresse4;
End If

'-----------------------------------5------------------
If Trim(recCptInfo.Adresse5) <> "" Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = 5000
    XPrt.Print recCptInfo.Adresse5;
End If
'------------------------------------6------------------
If Trim(recCptInfo.AdresseCP) <> "" _
Or Trim(recCptInfo.AdresseBD) <> "" Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = 4700
    XPrt.Print recCptInfo.AdresseCP & "  ";
    XPrt.Print recCptInfo.AdresseBD;
End If
'------------------------------------8------------------
If Trim(recCptInfo.AdressePays) <> "" Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = 5000
    XPrt.Print recCptInfo.AdressePays;
'----------------------------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3 - 50
XPrt.FontBold = True
XPrt.FontSize = 10
XPrt.CurrentX = 200
XPrt.Print Mid$(recOpCpt.Libellé, 28, 5);
XPrt.CurrentX = 1100
XPrt.Print Mid$(recOpCpt.Libellé, 33, 5);
XPrt.CurrentX = 2100
XPrt.Print Mid$(recOpCpt.Libellé, 38, 11);
XPrt.CurrentX = 3650
XPrt.Print Mid$(recOpCpt.Libellé, 49, 2);

XPrt.FontBold = False

XPrt.CurrentX = 5000
XPrt.Print recBic.Intitulé;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 5000
XPrt.Print recBic.Ville;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 5000
XPrt.Print recBic.Pays;


XPrt.FontSize = 8
XPrt.CurrentY = H + 1800
XPrt.CurrentX = 800
XPrt.Print recBdfG.BdfG20;


End If






End Sub

Public Sub prtVirFin2(H As Integer)
XPrt.FontSize = 5
XPrt.FontBold = False

XPrt.CurrentY = H + 4550
XPrt.CurrentX = 3200
XPrt.Print "MOTIFS";
XPrt.CurrentX = 4400
XPrt.Print "COORDONNEES ERRONEES";
XPrt.CurrentX = 5900
XPrt.Print "12";
XPrt.CurrentX = 6600
XPrt.Print "COMPTE SOLDE";
XPrt.CurrentX = 7500
XPrt.Print "14";
XPrt.CurrentX = 8200
XPrt.Print "AUTRE MOTIF";
'------------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 200
XPrt.Print "ORDONNATEUR";
XPrt.CurrentX = 2200
XPrt.Print "TRANSITAIRE";
XPrt.CurrentX = 3100
XPrt.Print "DE RETOUR ";
XPrt.CurrentX = 4400
XPrt.Print "DESTINATAIRE INCONNU";
XPrt.CurrentX = 5900
XPrt.Print "16";
XPrt.CurrentX = 6600
XPrt.Print "COMPTE VIRE";
XPrt.CurrentX = 7500
XPrt.Print "15";
'------------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 800
XPrt.Print "MARQUES D'AUTHENTIFICATION";

    For I = 3800 To 4800 Step 200
XPrt.Line (1800, I + H)-(1800, I + H + 100)
    Next I
XPrt.FontSize = 6

XPrt.CurrentY = H + 3350 + prtlineHeight / 2
XPrt.CurrentX = 900
XPrt.Print "MENTION BANQUE DE FRANCE";

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight / 2
XPrt.CurrentX = 1500
XPrt.Print "s'il y a lieu .";

'---------------------------BOXES--------------------
XPrt.DrawWidth = 1
Call frmElpPrt.prtTrame(4200, 4500 + H, 6100, 5000 + H, "B", 250)
Call frmElpPrt.prtTrame(6500, 4500 + H, 7700, 5000 + H, "B", 250)
Call frmElpPrt.prtTrame(8100, 4500 + H, 11000, 5000 + H, "B", 250)
Call frmElpPrt.prtTrame(3000, 4500 + H, 3900, 5000 + H, "B", 250)
Call frmElpPrt.prtTrame(6100, 4500 + H, 6300, 5000 + H, "B", 250)
Call frmElpPrt.prtTrame(7700, 4500 + H, 7900, 5000 + H, "B", 250)
Call frmElpPrt.prtTrame(7700, 4500 + H, 7900, 5000 + H, "B", 250)

'-----------------TRAITS HORIZONTAUX------------------

XPrt.Line (6500, 4750 + H)-(7900, 4750 + H)
XPrt.Line (4200, 4750 + H)-(6300, 4750 + H)

'-------------GUILLEMETS----------------->----------
XPrt.Line (3900, 4500 + H)-(4100, 4750 + H)
XPrt.Line (3900, 5000 + H)-(4100, 4750 + H)

End Sub

Public Sub prtVirFin3(H As Integer)
XPrt.FontSize = 6
XPrt.FontBold = False
XPrt.CurrentY = 9000
XPrt.CurrentX = 200
XPrt.Print "Date , cachet et signature de l'établissement chargé du règlement";

XPrt.FontSize = 8
XPrt.CurrentY = 9800
XPrt.CurrentX = 4800
XPrt.Print "Nous vous informons que nous portons au crédit de votre compte , valeur";
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 4800
XPrt.Print "le montant ci-dessus d'un virement émis en votre faveur.";

XPrt.Line (4000, 3800 + H)-(4000, 5200 + H)







End Sub

Public Sub prtVirFin4(H As Integer)


End Sub
