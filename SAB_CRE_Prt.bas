Attribute VB_Name = "prtSAB_CRE"
'-----------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim X As String, I As Integer, Height8_6 As Integer

Dim blnNewPage As Boolean, blnOpen As Boolean

 
Dim prtRéférenceY As Integer, prtCorpsY As Integer
Dim xDocRéférence As String

Dim meYBIACRE As typeYBIACRE

Dim meZCREAVI0 As typeZCREAVI0
Dim meZCREBIS0 As typeZCREBIS0
Dim meZCREPRE0 As typeZCREPRE0
Dim wREF As String
Dim xAMJ_Print As String
Dim xCompte_Print As String
Dim xNatureR As String, xNature As String
Public Sub prtSAB_CRE_ZCREAVI0(lZCREAVI0 As typeZCREAVI0, lYBIACRE As typeYBIACRE)
Dim I As Integer
Dim X As String
blnOpen = False

meYBIACRE = lYBIACRE
meZCREAVI0 = lZCREAVI0
meZCREPRE0 = meYBIACRE.ZCREPRE0(1)

prtSAB_CRE_ZCREAVI0_ZADRESS0

wREF = meZCREAVI0.CREAVINAT & " " & meZCREAVI0.CREAVIDOS & "_" & meZCREAVI0.CREAVIPRE & " du " & dateIBM10(meZCREPRE0.CREPREOUV, True)

xAMJ_Print = dateIBM10(meZCREAVI0.CREAVIDTC, True)
xCompte_Print = Trim(meZCREAVI0.CREAVICOM)
If xCompte_Print = "" Then xCompte_Print = Space$(30)
prtSAB_CRE_Nature

Select Case meZCREAVI0.CREAVITYP

    Case "00": prtSAB_CRE_MAD
    Case "02", "03", "04": prtSAB_CRE_ZCREAVI0_Avis_Echéance
    Case "??":  'prtSAB_CRE_ZCREAVI0_Confirmation
   
End Select
If blnOpen Then prtSAB_CRE_Close

End Sub

Public Sub prtSAB_CRE_ZCREBIS0(lZCREBIS0 As typeZCREBIS0, lYBIACRE As typeYBIACRE, optSelect_Confirmation As Boolean)
Dim I As Integer
Dim X As String
blnOpen = False

meYBIACRE = lYBIACRE
meZCREBIS0 = lZCREBIS0
meZCREPRE0 = meYBIACRE.ZCREPRE0(1)

wREF = meYBIACRE.ZCREPRE0(1).CREPRENAT & " " & meZCREBIS0.CREBISDOS & "_" & meZCREBIS0.CREBISPRE & " du " & dateIBM10(meZCREPRE0.CREPREOUV, True)
xAMJ_Print = dateImp10(DSys)
xCompte_Print = Trim(meZCREBIS0.CREBISCOM)
If xCompte_Print = "" Then xCompte_Print = Space$(30)
prtSAB_CRE_Nature

Select Case meZCREBIS0.CREBISTYP
    Case "00": prtSAB_CRE_ZCREBIS0_MAD
    Case "02", "03", "04": prtSAB_CRE_ZCREBIS0_Avis_Echéance optSelect_Confirmation
   
End Select
If blnOpen Then prtSAB_CRE_Close

End Sub

Public Sub prtSAB_CRE_Open()
On Error GoTo prtError
blnOpen = True
blnNewPage = False
Set XPrt = Printer
frmElpPrt.Show vbModeless
XPrt.FontItalic = False

Height8_6 = frmElpPrt.prtHeightDelta(8, 6)

blnFiligrane = False
prtPgmName = "prtSAB_CRE"
prtTitleUsr = usrName
prtOrientation = vbPRORPortrait
prtTitleText = "CRE_Courrier"
prtFontName = prtFontName_TimesNewRoman 'Arial


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


Public Sub prtSAB_CRE_Close()
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
Public Sub prtSAB_CRE_form()
'---------------------------------------------------------
Dim X As String

If Not blnOpen Then prtSAB_CRE_Open
If blnNewPage Then frmElpPrt.prtNewPage   'XPrt.NewPage
blnNewPage = True

XPrt.DrawWidth = 1
XPrt.FontSize = 10: XPrt.FontBold = False

XPrt.CurrentX = prtMinX + 6800
XPrt.CurrentY = prtMinY + prtlineHeight * 4

    XPrt.Print "Paris, le  " & xAMJ_Print;

Call prtAdresse_Enveloppe(meYBIACRE.CRE_ZADRESS0)

XPrt.FontSize = 8
XPrt.FontBold = False
XPrt.CurrentY = prtRéférenceY
XPrt.CurrentX = prtMinMarge: XPrt.Print "N/Référence";

XPrt.CurrentX = prtMinMarge + 1250: XPrt.Print ":";
XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1400: XPrt.Print wREF;
XPrt.FontBold = False


XPrt.CurrentY = prtRéférenceY + prtlineHeight * 2
XPrt.CurrentX = prtMinMarge: XPrt.Print "Votre contact";
XPrt.CurrentX = prtMinMarge + 1250: XPrt.Print ":";
XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1400: XPrt.Print "01 53 76 " & meYBIACRE.Contact;

XPrt.FontBold = False
XPrt.CurrentY = prtRéférenceY + prtlineHeight * 3
    XPrt.CurrentX = prtMinMarge: XPrt.Print "Télécopie";
XPrt.CurrentX = prtMinMarge + 1250: XPrt.Print ":";
XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1400: XPrt.Print "01 53 76 64 05";

XPrt.FontSize = 10: XPrt.FontBold = False

End Sub
'---------------------------------------------------------
Public Sub prtSAB_CRE_MAD()
'---------------------------------------------------------
Dim X As String

prtSAB_CRE_form

XPrt.CurrentY = prtCorpsY
XPrt.FontSize = 14: XPrt.FontBold = True: XPrt.FontUnderline = True
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
frmElpPrt.prtCentré prtMedX, "MISE A DISPOSITION "

XPrt.FontSize = 12: XPrt.FontBold = False: XPrt.FontUnderline = False


XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 4
XPrt.CurrentX = prtMinMarge
XPrt.Print "Messieurs,";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 4
XPrt.Print "Conformément à vos instructions, dans le cadre " & xNature & "par notre établissement,";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "nous vous confirmons avoir mis à votre disposition :";

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
'22/08/2011 DR le paramètre optionnel doit être passé
'Call frmElpPrt.prtTrame(prtMinMarge - 50, XPrt.CurrentY - 50, prtMaxMarge, XPrt.CurrentY + prtlineHeight * 5 + 50, , 245)
Call frmElpPrt.prtTrame(prtMinMarge - 50, XPrt.CurrentY - 50, prtMaxMarge, XPrt.CurrentY + prtlineHeight * 5 + 50, " ", 245)

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1
XPrt.FontBold = True
XPrt.FontSize = 14
X = meZCREAVI0.CREAVIDEV & "  " & Trim(Format$(meZCREAVI0.CREAVIMON, "### ### ### ##0.00"))
'XPrt.CurrentX = prtMaxMarge - 6000 - XPrt.TextWidth(X)
frmElpPrt.prtCentré prtMedX, X
XPrt.FontBold = False
XPrt.FontSize = 12


XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "au crédit de votre compte courant n°  ";
XPrt.FontBold = True
XPrt.Print meZCREAVI0.CREAVICOM;
XPrt.FontBold = False

XPrt.Print "   Valeur : ";
XPrt.FontBold = True
XPrt.Print dateIBM10(meZCREAVI0.CREAVIECH, True)
XPrt.FontBold = False

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 4
XPrt.Print "Veuillez agréer, Messieurs, nos salutations distinguées."


XPrt.CurrentX = prtMinMarge + 5000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 6
XPrt.FontBold = True
XPrt.Print paramSOC_RS;
XPrt.FontBold = False

End Sub

'---------------------------------------------------------
Public Sub prtSAB_CRE_ZCREBIS0_MAD()
'---------------------------------------------------------
Dim X As String

prtSAB_CRE_form

XPrt.CurrentY = prtCorpsY
XPrt.FontSize = 14: XPrt.FontBold = True: XPrt.FontUnderline = True
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
frmElpPrt.prtCentré prtMedX, "MISE A DISPOSITION "

XPrt.FontSize = 12: XPrt.FontBold = False: XPrt.FontUnderline = False


XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 4
XPrt.CurrentX = prtMinMarge
XPrt.Print "Messieurs,";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 4
XPrt.Print "Conformément à vos instructions, dans le cadre " & xNature & "par notre établissement,";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "nous vous confirmons avoir mis à votre disposition :";

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
'22/08/2011 DR le paramètre optionnel doit être passé
'Call frmElpPrt.prtTrame(prtMinMarge - 50, XPrt.CurrentY - 50, prtMaxMarge, XPrt.CurrentY + prtlineHeight * 5 + 50, , 245)
Call frmElpPrt.prtTrame(prtMinMarge - 50, XPrt.CurrentY - 50, prtMaxMarge, XPrt.CurrentY + prtlineHeight * 5 + 50, " ", 245)

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1
XPrt.FontBold = True
XPrt.FontSize = 14
X = meZCREBIS0.CREBISDRE & "  " & Trim(Format$(meZCREBIS0.CREBISMRE, "### ### ### ##0.00"))
'XPrt.CurrentX = prtMaxMarge - 6000 - XPrt.TextWidth(X)
frmElpPrt.prtCentré prtMedX, X
XPrt.FontBold = False
XPrt.FontSize = 12


XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "au crédit de votre compte courant n°  ";
XPrt.FontBold = True
XPrt.Print meZCREBIS0.CREBISCOM;
XPrt.FontBold = False

XPrt.Print "   Valeur : ";
XPrt.FontBold = True
XPrt.Print dateIBM10(meZCREBIS0.CREBISEMI, True)
XPrt.FontBold = False

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 4
XPrt.Print "Veuillez agréer, Messieurs, nos salutations distinguées."


XPrt.CurrentX = prtMinMarge + 5000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 6
XPrt.FontBold = True
XPrt.Print paramSOC_RS;
XPrt.FontBold = False

End Sub


'---------------------------------------------------------
Public Sub prtSAB_CRE_ZCREAVI0_Avis_Echéance()
'---------------------------------------------------------

Dim X As String
Dim y As String

prtSAB_CRE_form

XPrt.CurrentY = prtCorpsY
XPrt.FontSize = 14: XPrt.FontBold = True: XPrt.FontUnderline = True
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
frmElpPrt.prtCentré prtMedX, "AVIS D'ECHEANCE "

XPrt.FontSize = 12: XPrt.FontBold = False: XPrt.FontUnderline = False

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 4
XPrt.Print "Messieurs,";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.Print "Dans le cadre " & xNatureR;
'XPrt.FontBold = True
XPrt.Print wREF;
'XPrt.FontBold = False

XPrt.Print ", nous débitons, en date du ";
XPrt.FontBold = True
XPrt.Print dateIBM10(meZCREAVI0.CREAVIRGL, True);
XPrt.FontBold = False
XPrt.Print ",";
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge
XPrt.Print "votre compte courant n°  ";
XPrt.FontBold = True
XPrt.Print xCompte_Print;
XPrt.FontBold = False
XPrt.Print ", suivant le décompte ci-après :";

'=======================

XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3

XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.Print "Taux appliqué      :";
XPrt.FontBold = True
X = Format$(meZCREAVI0.CREAVITAU, "#0.000000")
XPrt.CurrentX = prtMinMarge + 5000 - XPrt.TextWidth(X)
XPrt.Print X & "  %";
XPrt.FontBold = False

XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.Print "Intérêts de la période du ";
'XPrt.FontBold = True
XPrt.Print dateIBM10(meZCREAVI0.CREAVIDEB, True);
'XPrt.FontBold = False
XPrt.Print " au ";
'XPrt.FontBold = True
XPrt.Print dateIBM10(meZCREAVI0.CREAVIFIN, True);
'XPrt.FontBold = False


X = Format$(meZCREAVI0.CREAVIMIN, "### ### ### ##0.00")
XPrt.CurrentX = prtMinMarge + 9000 - XPrt.TextWidth(X)
XPrt.Print X & "  " & meZCREAVI0.CREAVIDEV;

XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Amortissement ";

X = Format$(meZCREAVI0.CREAVIMON, "### ### ### ##0.00")
XPrt.CurrentX = prtMinMarge + 9000 - XPrt.TextWidth(X)

XPrt.Print X & "  " & meZCREAVI0.CREAVIDEV;

XPrt.FontBold = True

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
'Call frmElpPrt.prtTrame(prtMinMarge + 7500, XPrt.CurrentY - 50, prtMinMarge + 9800, XPrt.CurrentY + prtlineHeight + 50, , 245)
'22/08/2011 DR le paramètre optionnel doit être passé
'Call frmElpPrt.prtTrame(prtMinMarge, XPrt.CurrentY - 50, prtMinMarge + 9800, XPrt.CurrentY + prtlineHeight + 50, , 245)
Call frmElpPrt.prtTrame(prtMinMarge, XPrt.CurrentY - 50, prtMinMarge + 9800, XPrt.CurrentY + prtlineHeight + 50, " ", 245)
XPrt.Line (prtMinMarge + 7500, XPrt.CurrentY - 50)-(prtMinMarge + 9800, XPrt.CurrentY - 50), prtLineColor

XPrt.CurrentX = prtMinMarge
XPrt.Print "Montant net à payer ";
X = Format$(meZCREAVI0.CREAVIMDR, "### ### ### ##0.00")
XPrt.CurrentX = prtMinMarge + 9000 - XPrt.TextWidth(X)
XPrt.Print X & "  " & meZCREPRE0.CREPREDEV;

If meZCREAVI0.CREAVIMDR <> meZCREAVI0.CREAVIMON + meZCREAVI0.CREAVIMIN Then
    XPrt.FontSize = 14: XPrt.FontBold = True: XPrt.FontUnderline = True
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
    frmElpPrt.prtCentré prtMedX, "!!!!!!!!!!!!!!!!! ERREUR : TOTAL!!!!!!!!!!!!!!!!!! "
    
    XPrt.FontSize = 11: XPrt.FontBold = False: XPrt.FontUnderline = False
End If

XPrt.FontBold = False


XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 4
XPrt.Print "Veuillez agréer, Messieurs, nos salutations distinguées."


XPrt.CurrentX = prtMinMarge + 5000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 5
XPrt.FontBold = True
XPrt.Print paramSOC_RS;
XPrt.FontBold = False

End Sub

'---------------------------------------------------------
Public Sub prtSAB_CRE_ZCREBIS0_Avis_Echéance(optSelect_Confirmation As Boolean)
'---------------------------------------------------------

Dim X As String
Dim y As String

prtSAB_CRE_form

XPrt.CurrentY = prtCorpsY
XPrt.FontSize = 14: XPrt.FontBold = True: XPrt.FontUnderline = True
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
If optSelect_Confirmation Then
    X = "CONFIRMATION des CONDITIONS de CREDIT"
Else
    X = "AVIS D'ECHEANCE "
End If

frmElpPrt.prtCentré prtMedX, X

XPrt.FontSize = 12: XPrt.FontBold = False: XPrt.FontUnderline = False

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 4
XPrt.Print "Messieurs,";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.Print "Dans le cadre " & xNatureR;
'XPrt.FontBold = True
XPrt.Print wREF;
'XPrt.FontBold = False
If optSelect_Confirmation Then
    XPrt.Print ", veuillez prendre note des conditions appliquées";
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "pour la période en cours :";
Else
    XPrt.Print ", nous débitons, en date du ";
    XPrt.FontBold = True
    XPrt.Print dateIBM10(meZCREBIS0.CREBISEMI, True);
    XPrt.FontBold = False
    XPrt.Print ",";
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "votre compte courant n°  ";
    XPrt.FontBold = True
    XPrt.Print xCompte_Print;
    XPrt.FontBold = False
    XPrt.Print ", suivant le décompte ci-après :";
End If


'=======================

XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.Print "Capital restant dû :";

XPrt.FontBold = True
X = Format$(meZCREPRE0.CREPRECAP, "### ### ### ##0.00")
XPrt.CurrentX = prtMinMarge + 5000 - XPrt.TextWidth(X)
XPrt.Print X & "  " & meZCREPRE0.CREPREDEV;
XPrt.FontBold = False

XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.Print "Taux appliqué      :";
XPrt.FontBold = True
X = Format$(meZCREBIS0.CREBISTAU, "#0.000000")
XPrt.CurrentX = prtMinMarge + 5000 - XPrt.TextWidth(X)
XPrt.Print X & "  %";
XPrt.FontBold = False

XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.Print "Intérêts de la période du ";
XPrt.FontBold = True
XPrt.Print dateIBM10(meZCREBIS0.CREBISDEB, True);
XPrt.FontBold = False
XPrt.Print " au ";
XPrt.FontBold = True
XPrt.Print dateIBM10(meZCREBIS0.CREBISFIN, True);
XPrt.FontBold = False


X = Format$(meZCREBIS0.CREBISMIN, "### ### ### ##0.00")
XPrt.CurrentX = prtMinMarge + 9000 - XPrt.TextWidth(X)
XPrt.Print X & "  " & meZCREBIS0.CREBISDRE;

If meZCREBIS0.CREBISMAM <> 0 Then
    XPrt.CurrentX = prtMinMarge + 1000
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
    XPrt.Print "Amortissement ";
    
    X = Format$(meZCREBIS0.CREBISMAM, "### ### ### ##0.00")
    XPrt.CurrentX = prtMinMarge + 9000 - XPrt.TextWidth(X)
    XPrt.Print X & "  " & meZCREPRE0.CREPREDEV;
End If
If meZCREBIS0.CREBISASC <> 0 Then

    XPrt.CurrentX = prtMinMarge + 1000
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
    XPrt.Print "Assurance ";
    
    X = Format$(meZCREBIS0.CREBISASC, "### ### ### ##0.00")
    XPrt.CurrentX = prtMinMarge + 9000 - XPrt.TextWidth(X)
    
    XPrt.Print X & "  " & meZCREAVI0.CREAVIDEV;
End If

XPrt.FontBold = True

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
'Call frmElpPrt.prtTrame(prtMinMarge + 7500, XPrt.CurrentY - 50, prtMinMarge + 9800, XPrt.CurrentY + prtlineHeight + 50, , 245)
Call frmElpPrt.prtTrame(prtMinMarge, XPrt.CurrentY - 50, prtMinMarge + 9800, XPrt.CurrentY + prtlineHeight + 50, "B", 245)
XPrt.Line (prtMinMarge + 7500, XPrt.CurrentY - 50)-(prtMinMarge + 9800, XPrt.CurrentY - 50), prtLineColor

XPrt.CurrentX = prtMinMarge
If optSelect_Confirmation Then
    XPrt.Print "Total ";
Else
    XPrt.Print "Montant net à payer ";
End If
X = Format$(meZCREBIS0.CREBISMRE, "### ### ### ##0.00")
XPrt.CurrentX = prtMinMarge + 9000 - XPrt.TextWidth(X)
XPrt.Print X & "  " & meZCREPRE0.CREPREDEV;

If meZCREBIS0.CREBISMRE <> meZCREBIS0.CREBISMAM + meZCREBIS0.CREBISMIN + meZCREBIS0.CREBISASC Then
    XPrt.FontSize = 14: XPrt.FontBold = True: XPrt.FontUnderline = True
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
    frmElpPrt.prtCentré prtMedX, "!!!!!!!!!!!!!!!!! ERREUR : TOTAL!!!!!!!!!!!!!!!!!! "
    
    XPrt.FontSize = 11: XPrt.FontBold = False: XPrt.FontUnderline = False
End If

XPrt.FontBold = False


XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 4
XPrt.Print "Veuillez agréer, Messieurs, nos salutations distinguées."


XPrt.CurrentX = prtMinMarge + 5000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 5
XPrt.FontBold = True
XPrt.Print paramSOC_RS;
XPrt.FontBold = False

End Sub


'---------------------------------------------------------
Public Sub prtSAB_CRE_ZCREBIS0_Confirmation()
'---------------------------------------------------------
Dim X As String
'XPrt.Print "Dans le cadre du crédit en rubrique, veuillez prendre note des conditions appliquées pour la période en cours :";

prtSAB_CRE_form

XPrt.CurrentY = prtCorpsY
XPrt.FontSize = 14: XPrt.FontBold = True: XPrt.FontUnderline = True
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
frmElpPrt.prtCentré prtMedX, "CONFIRMATION DES CONDITIONS DE CREDIT"

XPrt.FontSize = 11: XPrt.FontBold = False: XPrt.FontUnderline = False

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.Print "Messieurs,";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.Print "Dans le cadre " & xNatureR;
XPrt.FontBold = True
XPrt.Print wREF;
XPrt.FontBold = False

XPrt.Print ", veuillez prendre note des conditions appliquées pour la période en cours :";
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.CurrentX = prtMinMarge
XPrt.Print "Débit du compte :";

XPrt.FontBold = True
XPrt.Print xCompte_Print & ",";
XPrt.FontBold = False

XPrt.Print " en date du : ";
XPrt.FontBold = True
XPrt.Print dateIBM10(meZCREBIS0.CREBISEMI, True);
XPrt.FontBold = False


'=======================


XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Intérêts le la période du ";
XPrt.FontBold = True
XPrt.Print dateIBM10(meZCREBIS0.CREBISDEB, True);
XPrt.FontBold = False
XPrt.Print " au ";
XPrt.FontBold = True
XPrt.Print dateIBM10(meZCREBIS0.CREBISFIN, True);
XPrt.FontBold = False

If meZCREBIS0.CREBISMIN <> 0 Then
    X = Format$(meZCREBIS0.CREBISMIN, "### ### ### ##0.00")
    XPrt.CurrentX = prtMinMarge + 9000 - XPrt.TextWidth(X)
    XPrt.Print X & "  " & meYBIACRE.ZCREPRE0(1).CREPREDEV;
End If

XPrt.CurrentX = prtMinMarge + 700
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.Print "Taux appliqué : ";
XPrt.FontBold = True
X = Format$(meZCREBIS0.CREBISTAU, "#0.000000")
'XPrt.CurrentX = prtMinMarge + 3500 - XPrt.TextWidth(X)
XPrt.Print X & " %";
XPrt.FontBold = False

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2

If meZCREBIS0.CREBISMAM <> 0 Then

    XPrt.Print "Amortissement ";
    
    X = Format$(meZCREBIS0.CREBISMAM, "### ### ### ##0.00")
    XPrt.CurrentX = prtMinMarge + 9000 - XPrt.TextWidth(X)
        XPrt.Print X & "  " & meYBIACRE.ZCREPRE0(1).CREPREDEV;
End If

If meZCREBIS0.CREBISASC <> 0 Then

    XPrt.CurrentX = prtMinMarge + 1000
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
    XPrt.Print "Assurance ";
    
    X = Format$(meZCREBIS0.CREBISASC, "### ### ### ##0.00")
    XPrt.CurrentX = prtMinMarge + 9000 - XPrt.TextWidth(X)
    
    XPrt.Print X & "  " & meZCREAVI0.CREAVIDEV;
End If

XPrt.FontBold = True

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2



XPrt.FontBold = False

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 4
XPrt.Print "Veuillez agréer, Messieurs, nos salutations distinguées."


XPrt.CurrentX = prtMinMarge + 5000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 5
XPrt.FontBold = True
XPrt.Print paramSOC_RS;
XPrt.FontBold = False

End Sub

Public Sub prtSAB_CRE_ZCREAVI0_ZADRESS0()
meYBIACRE.CRE_ZADRESS0.ADRESSETA = meZCREAVI0.CREAVIETA                     ' Etablissement
meYBIACRE.CRE_ZADRESS0.ADRESSTYP = "T"      ' String * 1                     ' 1 client , 2 compte
meYBIACRE.CRE_ZADRESS0.ADRESSPLA = 0       ' Long                           ' Numéro de plan
meYBIACRE.CRE_ZADRESS0.ADRESSNUM = meZCREAVI0.CREAVICOM     ' String * 20                    ' ou numéro de client
meYBIACRE.CRE_ZADRESS0.ADRESSCOA = ""      ' String * 2                     ' Code adresse
meYBIACRE.CRE_ZADRESS0.ADRESSDLI = 0       ' Long                           ' Date limite validité
meYBIACRE.CRE_ZADRESS0.ADRESSDDE = 0       ' Long                           ' Date début validité
meYBIACRE.CRE_ZADRESS0.ADRESSRA1 = meZCREAVI0.CREAVIRA1      ' String * 32                    ' ou raison sociale 1
meYBIACRE.CRE_ZADRESS0.ADRESSRA2 = meZCREAVI0.CREAVIRA2      ' String * 32                    ' ou raison sociale 2
meYBIACRE.CRE_ZADRESS0.ADRESSAD1 = meZCREAVI0.CREAVIAD1     ' String * 32                    ' Adresse 1
meYBIACRE.CRE_ZADRESS0.ADRESSAD2 = meZCREAVI0.CREAVIAD2     ' String * 32                    ' Adresse 2
meYBIACRE.CRE_ZADRESS0.ADRESSAD3 = meZCREAVI0.CREAVIAD3      ' String * 32                    ' Adresse 3
meYBIACRE.CRE_ZADRESS0.ADRESSCOP = meZCREAVI0.CREAVICOP    ' String * 6                     ' Code postal
meYBIACRE.CRE_ZADRESS0.ADRESSVIL = meZCREAVI0.CREAVIVIL      ' String * 25                    ' Ville
meYBIACRE.CRE_ZADRESS0.ADRESSPAY = meZCREAVI0.CREAVIPAY      ' String * 25                    ' Pays
meYBIACRE.CRE_ZADRESS0.ADRESSTEL = meZCREAVI0.CREAVITEL     ' String * 20                    ' No Tel.
meYBIACRE.CRE_ZADRESS0.ADRESSFAX = meZCREAVI0.CREAVIFAX       ' String * 20                    ' No Fax.
meYBIACRE.CRE_ZADRESS0.ADRESSTEX = ""       ' String * 20                    ' No Télex

End Sub

Public Sub prtSAB_CRE_Nature()
If meZCREPRE0.CREPRENAT = "PTR" Then
    xNatureR = "de l'avance référencée "
    xNature = "de l'avance qui vous a été consentie "
Else
    xNatureR = "du prêt référencé "
    xNature = "du prêt qui vous a été consenti "
End If

End Sub
