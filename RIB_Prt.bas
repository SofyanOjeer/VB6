Attribute VB_Name = "prtRIB"
Option Explicit
Dim mRib_Compte As String, mRib_Clé As String, mRib_IbanE As String
Dim mRib_COMPTEDEV As String
Dim meZADRESS0 As typeZADRESS0 ', selZADRESS0 As typeZADRESS0

Dim blnZADRESS0_Auto As Boolean
'----------------------------------
Public Sub prtRIB_Monitor(lCompte As String)
'---------------------------------------------------------

On Error GoTo prtError

blnZADRESS0_Auto = True
rsZADRESS0_Init meZADRESS0

Set XPrt = Printer

frmElpPrt.Show vbModeless

prtOrientation = vbPRORPortrait
prtTitleText = "RIB"
prtPgmName = "prtRib"
prtTitleUsr = usrName
prtRIB_A4 lCompte

DoEvents: If prtKillDoc Then frmElpPrt.Hide: Exit Sub


'frmElpPrt.prtLineY
Call frmElpPrt.prtEndDoc(1000)
frmElpPrt.Hide

Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub
'----------------------------------
Public Sub prtRIB_ZADRESS0(lCompte As String, lZADRESS0 As typeZADRESS0)
'---------------------------------------------------------

On Error GoTo prtError

meZADRESS0 = lZADRESS0
blnZADRESS0_Auto = False

Set XPrt = Printer

frmElpPrt.Show vbModeless

prtOrientation = vbPRORPortrait
prtTitleText = "RIB"
prtPgmName = "prtRib"
prtTitleUsr = usrName
prtRIB_A4 lCompte

DoEvents: If prtKillDoc Then frmElpPrt.Hide: Exit Sub


'frmElpPrt.prtLineY
Call frmElpPrt.prtEndDoc(1000)
frmElpPrt.Hide

Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub

'----------------------------------
Public Sub prtRIB_A4(lCompte As String)
'---------------------------------------------------------
Dim V, xSQL As String
Dim xRacine As String

Dim blnOk As Boolean

On Error GoTo prtError

mRib_Compte = Trim(lCompte)
meZADRESS0.ADRESSNUM = lCompte

blnOk = False
If Mid$(lCompte, 1, 4) = "3656" Or Mid$(lCompte, 1, 4) = "3616" Or Mid$(lCompte, 1, 4) = "3889" Or Mid$(lCompte, 1, 3) = "162" Then
    Set rsSab = Nothing

    xSQL = "select COMREFREF from " & paramIBM_Library_SAB & ".ZCOMREF0 where COMREFCOM = '" & lCompte & "' and COMREFCOR = 'SI'"
    Set rsSab = cnsab.Execute(xSQL)
    If Not rsSab.EOF Then
        blnOk = True
        mRib_Compte = Trim(rsSab("COMREFREF"))
        meZADRESS0.ADRESSRA1 = paramSOC_RS
        meZADRESS0.ADRESSAD1 = paramSOC_Adresse
        meZADRESS0.ADRESSVIL = paramSOC_Ville
    End If
Else
    If blnZADRESS0_Auto Then
        V = rsZADRESS0_Compte(meZADRESS0)
        If Trim(meZADRESS0.ADRESSVIL) = "" Then
            meZADRESS0.ADRESSNUM = lCompte
            V = rsZADRESS0_Titulaire(meZADRESS0)
        End If
    Else
        V = Null
    End If
End If
If IsNull(V) Then blnOk = True


If blnOk Then
    
    mRib_Clé = Format$(RibClé(strSocBdfE, strSocBdfG, mRib_Compte, mRib_IbanE), "00")
    
    xSQL = "select COMPTEDEV from " & paramIBM_Library_SAB & ".ZCOMPTE0" _
     & " where COMPTECOM = '" & lCompte & "'" _
     & " and COMPTEETA = " & currentZMNURUT0.MNURUTETB
    
    Set rsSab = cnsab.Execute(xSQL)
    If Not rsSab.EOF Then
        mRib_COMPTEDEV = rsSab("COMPTEDEV")
    Else
         mRib_COMPTEDEV = ""
   End If
    
    prtLineNb = 1
    prtlineHeight = 250
    prtHeaderHeight = 300
    
    prtFormType = ""
    frmElpPrt.prtInit
    
    XPrt.FontName = prtFontName_Arial
    
    XPrt.CurrentY = 0
    prtRibForm
    XPrt.CurrentY = 5100
    frmElpPrt.prtTiret
    XPrt.CurrentY = 5400
    
    prtRibForm
    XPrt.CurrentY = 10700
    frmElpPrt.prtTiret
    XPrt.CurrentY = 11000
    
    prtRibForm

End If


Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub


'----------------------------------
Public Sub prtRibForm()

'----------------------------------
Dim X As String, J As Integer, H2 As Integer, IbanE As String
Dim I As Long

XPrt.FontSize = 8
XPrt.FontBold = True
XPrt.DrawWidth = 2

J = XPrt.CurrentY
I = frmElpPrt.imgSocLogo.Width * 0.13
XPrt.PaintPicture frmElpPrt.imgSocLogo.Picture _
                , 3500, J + prtlineHeight _
                , I _
                , frmElpPrt.imgSocLogo.Height * 0.13




XPrt.CurrentY = J + prtlineHeight
XPrt.FontSize = 11
XPrt.FontBold = True
frmElpPrt.prtCentré 1500, paramSOC_RS
'---------------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.FontBold = False
XPrt.FontSize = 8
'---------------------------------------------------
frmElpPrt.prtCentré 1500, paramSOC_Adresse
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
frmElpPrt.prtCentré 1500, paramSOC_Ville
'---------------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
frmElpPrt.prtCentré 1500, "Tél : " & socTéléphone
XPrt.FontSize = 11
XPrt.FontBold = True
XPrt.CurrentX = 6750
XPrt.Print "RELEVE D'IDENTITE BANCAIRE / IBAN";
XPrt.FontBold = False
XPrt.FontSize = 8
'---------------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
frmElpPrt.prtCentré 1500, "Compensable à Paris - 27"
XPrt.FontSize = 7
XPrt.CurrentX = 7500
XPrt.Print "Cadre réservé au destinataire du relevé";
J = XPrt.CurrentY + prtlineHeight / 2 - 50
'---------------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Line (6800, J)-(7400, J), prtLineColor
XPrt.Line (10150, J)-(10800, J), prtLineColor

'---------------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.CurrentX = 150
XPrt.Print "Ce relevé est destiné à être remis, sur leur demande, à vos créanciers ou débiteurs français, ou";

'---------------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 150
XPrt.Print "étrangers, appelés à faire inscrire des opérations à votre compte(virements, paiement de quittances,...).";

'---------------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight + 150
XPrt.CurrentX = 150
XPrt.Print "This statement is intended to be delivered to those of your creditors or debtors who have";

'---------------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 150
XPrt.Print "transactions posted to your account (credit transfers, invoice payments, ...).";
'---------------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2

XPrt.Line (6800, XPrt.CurrentY)-(10800, XPrt.CurrentY), prtLineColor
XPrt.Line (6800, J)-(6800, XPrt.CurrentY), prtLineColor
XPrt.Line (10800, J)-(10800, XPrt.CurrentY), prtLineColor

'---------------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight - 200
J = XPrt.CurrentY
H2 = J + 80
XPrt.Line (150, H2)-(250, H2), prtLineColor
XPrt.Line (800, H2)-(1050, H2), prtLineColor
XPrt.Line (1650, H2)-(2150, H2), prtLineColor
XPrt.Line (2750, H2)-(3250, H2), prtLineColor
XPrt.Line (3550, H2)-(3850, H2), prtLineColor
XPrt.Line (4450, H2)-(6600, H2), prtLineColor

XPrt.Line (150, H2)-(150, H2 + 650), prtLineColor
XPrt.Line (6600, H2)-(6600, H2 + 650), prtLineColor
XPrt.Line (150, H2 + 650)-(6600, H2 + 650), prtLineColor

XPrt.CurrentY = J
XPrt.CurrentX = 300
XPrt.Print "Banque";
XPrt.CurrentX = 1100
XPrt.Print "Guichet";
XPrt.CurrentX = 2200
XPrt.Print "Compte";
XPrt.CurrentX = 3300
XPrt.Print "Clé";
XPrt.CurrentX = 3900
XPrt.Print "Domiciliation";
'---------------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight + 100
XPrt.FontSize = 9
XPrt.FontBold = True

'13/08/2010 - Denis
'If mRib_COMPTEDEV = "EUR" Then
    XPrt.CurrentX = 300
    XPrt.Print strSocBdfE;
    XPrt.CurrentX = 1100
    XPrt.Print strSocBdfG;
    XPrt.CurrentX = 1900
    XPrt.Print Format$(mRib_Compte, "@@@ @@@ @@@ @@@");
    XPrt.CurrentX = 3300
    XPrt.Print Format$(mRib_Clé, "@@");
    XPrt.CurrentX = 3900
    XPrt.Print SocRibDom;
'End If

XPrt.FontSize = 7
'XPrt.FontBold = True
XPrt.CurrentX = 7400
XPrt.Print meZADRESS0.ADRESSRA1;
XPrt.FontBold = False

'-------------------------------------------------
  
If Trim(meZADRESS0.ADRESSRA2) <> "" Then
    XPrt.CurrentY = XPrt.CurrentY + 200 'prtlineHeight
    XPrt.CurrentX = 7400
    XPrt.Print meZADRESS0.ADRESSRA2;
End If

'-------------------------------------------------
  
If Trim(meZADRESS0.ADRESSAD1) <> "" Then
    XPrt.CurrentY = XPrt.CurrentY + 200 'prtlineHeight
    XPrt.CurrentX = 7400
    XPrt.Print meZADRESS0.ADRESSAD1;
End If


'-------------------------------------------------
If Trim(meZADRESS0.ADRESSAD2) <> "" Then
    XPrt.CurrentY = XPrt.CurrentY + 200 'prtlineHeight
    XPrt.CurrentX = 7400
    XPrt.Print meZADRESS0.ADRESSAD2;
End If
'---------------------------------------------------
If Trim(meZADRESS0.ADRESSAD3) <> "" Then
    XPrt.CurrentY = XPrt.CurrentY + 200 'prtlineHeight
    XPrt.CurrentX = 7400
    XPrt.Print meZADRESS0.ADRESSAD3;
End If

'----------------------------------------------------
If Trim(meZADRESS0.ADRESSCOP) <> "" _
Or Trim(meZADRESS0.ADRESSVIL) <> "" Then
    XPrt.CurrentY = XPrt.CurrentY + 200 'prtlineHeight
    XPrt.CurrentX = 7400
    If Trim(meZADRESS0.ADRESSCOP) <> "" Then XPrt.Print meZADRESS0.ADRESSCOP & "  ";
    XPrt.Print meZADRESS0.ADRESSVIL;
End If
'-----------------------------------------------------
If Trim(meZADRESS0.ADRESSPAY) <> "" Then
    XPrt.CurrentY = XPrt.CurrentY + 200 'prtlineHeight
    XPrt.CurrentX = 7400
    XPrt.Print meZADRESS0.ADRESSPAY;
End If
'---------------------------------------------------
XPrt.CurrentY = J + prtlineHeight * 4
XPrt.CurrentX = 300
XPrt.Print "IBAN International Bank Account Number";
XPrt.CurrentX = 3900
XPrt.Print "Bank Identification Code";

J = XPrt.CurrentY
H2 = J + 80
XPrt.Line (150, H2)-(150, H2 + 650), prtLineColor
XPrt.Line (6600, H2)-(6600, H2 + 650), prtLineColor

XPrt.Line (150, H2)-(250, H2), prtLineColor
XPrt.Line (2950, H2)-(3850, H2), prtLineColor

XPrt.Line (5450, H2)-(6600, H2), prtLineColor
XPrt.Line (150, H2 + 650)-(6600, H2 + 650), prtLineColor

XPrt.CurrentY = J + prtlineHeight + 100
XPrt.FontSize = 9
XPrt.CurrentX = 300
XPrt.FontBold = True
XPrt.Print Iban_Print(mRib_IbanE);
XPrt.CurrentX = 3900
XPrt.Print SocBicId;
XPrt.CurrentX = 6700
XPrt.Print mRib_COMPTEDEV;

XPrt.FontBold = False

End Sub







