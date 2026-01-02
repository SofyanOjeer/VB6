Attribute VB_Name = "prtCDO_Courrier"

'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim X As String, I As Integer, Height8_6 As Integer

Dim blnNewPage As Boolean, blnOpen As Boolean

Dim meYCDODOS0 As typeZCDODOS0
Dim meYCDOMOD0 As typeZCDOMOD0
Dim meYCDOTC20 As typeZCDOTC20
Dim meYCDOCOM0 As typeZCDOCOM0
Dim meYCDOCO20 As typeZCDOCO20
Dim mecnfYBIACDOCOM0 As typeYBIACDOCOM0, menotYBIACDOCOM0 As typeYBIACDOCOM0
 
Dim prtRéférenceY As Integer, prtCorpsY As Integer

Dim W_Validité_AMJ As String
Dim wDate_Validité As String, wDate_Anc_Validité As String, wDate_Limite_Emb As String
Dim wDate_Validité_GB As String
Dim wDate_Valeur_CR As String, wDate_Echeance_CR As String
Dim wDate_Valeur_DB As String, wDate_Echeance_DB As String
Dim wDate_Valeur_DB_GB As String, wDate_Echeance_DB_GB As String
Dim wDate_Remise_Util As String
Dim wCompte_CR As String

Dim wAnnexe_Nb As String, wContact As String
Dim xDocRéférence As String

Dim Line2_Ecart As Integer

Dim Booleen_GB As Boolean
Dim Booleen_AR As Boolean

Type typeCDO_Courrier
    
    Garantie_Nb     As Integer   ' Pour OUVERTURE
    Garantie()      As String

    prtNb           As Integer
    Contact         As String
    Annexe_Nb       As Integer
    
    ATT             As String
    CrrGB_Mnt       As String
    CrrGB_Tx        As String
    CrrGB_BqRbt     As String
    
    IBAN            As String
    
    Exp_Par         As String
    Exp_Le          As String
    Exp_De          As String
    Exp_A           As String
    Exp_Nb          As Integer
    Exp()           As String
    Delai_Nb        As Integer
    Delai()         As String
    Irregul_Nb      As Integer
    Irregul()       As String
    
    Document_Nb     As Integer
    Document()      As String
    Document_Jeu1() As String
    Document_Jeu2() As String
    
    YCDOUTI0        As typeZCDOUTI0
    YCDOREG0_R_Nb   As Integer
    YCDOREG0_R      As typeZCDOREG0
    YCDOREG0_C_Nb   As Integer
    YCDOREG0_C      As typeZCDOREG0
    YCDOREG0_D_Nb   As Integer
    YCDOREG0_D      As typeZCDOREG0
    
    Com_Nb          As Integer
    CDOCOMCOM()     As String
    CDOCOMDEV()     As String
    CDOREGCRD()     As String
    CDOCOMMON()     As Currency
    CDOCOMMTV()     As Currency
    
    
    BQE_ZADRESS0 As typeZADRESS0
    BQE_Concat As String
    DON_ZADRESS0 As typeZADRESS0
    DON_Concat As String
    BEN_ZADRESS0 As typeZADRESS0
    BEN_Concat As String
    BED_ZADRESS0 As typeZADRESS0
    BED_Concat As String

End Type

Dim meCDO_Courrier As typeCDO_Courrier

Const paramSignature As String = "GESTION DES CREDITS DOCUMENTAIRES"

'---------------------------------------------------------
Public Sub prtCDO_Courrier_OUV_Annexe_09()
'---------------------------------------------------------
Dim X As String

XPrt.CurrentY = prtCorpsY
XPrt.FontSize = 14: XPrt.FontBold = True
frmElpPrt.prtCentré prtMedX, "ANNEXE FAISANT PARTIE INTEGRANTE"
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
frmElpPrt.prtCentré prtMedX, "DU CREDIT DOCUMENTAIRE"
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
frmElpPrt.prtCentré prtMedX, "N°  " & Trim(meYCDODOS0.CDODOSEXT)
XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 8
XPrt.CurrentX = prtMinMarge
XPrt.Print " GARANTIE DE BONNE EXECUTION :";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print " Conformément aux instructions de la banque émettrice, ce crédit ne deviendra opérationnel qu'après ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print " réception d'une caution bancaire de bonne exécution fixée à 10 pour cent de la valeur globale du ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print " contrat et son acceptation par l'ordonnateur. ";


XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print " L'instruction rendant le crédit opérationnel, vous sera notifiée par nos soins."; ";"


XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print " Par conséquent, veuillez considérer ce crédit non encore opérationnel. ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print " Les frais et commissions étant stipulés à votre charge nous vous réclamerons lors de l'utilisation ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print " ou à la péremption du crédit notre commission forfaitaire globale de : ";

XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "EUR 150,00 compte tenu du faible montant de cette opération. ";

XPrt.CurrentX = prtMinMarge + 5000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 6
XPrt.FontBold = True
XPrt.Print paramSignature;
XPrt.FontBold = False


End Sub

'---------------------------------------------------------
Public Sub prtCDO_Courrier_OUV_Annexe_10()
'---------------------------------------------------------
Dim X As String

XPrt.CurrentY = prtCorpsY
XPrt.FontSize = 14: XPrt.FontBold = True
frmElpPrt.prtCentré prtMedX, "ANNEXE FAISANT PARTIE INTEGRANTE"
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
frmElpPrt.prtCentré prtMedX, "DU CREDIT DOCUMENTAIRE"
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
frmElpPrt.prtCentré prtMedX, "N°  " & Trim(meYCDODOS0.CDODOSEXT)
XPrt.FontSize = 11: XPrt.FontBold = False



XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 8
XPrt.CurrentX = prtMinMarge
XPrt.Print " GARANTIE DE BONNE EXECUTION :";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print " Conformément aux instructions de la banque émettrice, ce crédit ne deviendra opérationnel qu'après ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print " réception d'une caution bancaire de bonne exécution fixée à 10 pour cent de la valeur globale du ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print " contrat et son acceptation par l'ordonnateur. ";


XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print " L'instruction rendant le crédit opérationnel, vous sera notifiée par nos soins."; ";"


XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print " Par conséquent, veuillez considérer ce crédit non encore opérationnel. ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print " Les frais et commissions étant stipulés à votre charge nous vous réclamerons lors de l'utilisation ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print " ou à la péremption du crédit notre commission forfaitaire globale de : ";

XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "EUR 250,00 compte tenu du faible montant de cette opération. ";

XPrt.CurrentX = prtMinMarge + 5000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 6
XPrt.FontBold = True
XPrt.Print paramSignature;
XPrt.FontBold = False





End Sub

'---------------------------------------------------------
Public Sub prtCDO_Courrier_OUV_C_NOP_11_AVueNonRecl()
'---------------------------------------------------------
Dim X As String
Dim Tx As Double
Dim C As Integer

XPrt.CurrentY = prtCorpsY
XPrt.FontSize = 14: XPrt.FontBold = True
frmElpPrt.prtCentré prtMedX, "ANNEXE FAISANT PARTIE INTEGRANTE"
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
frmElpPrt.prtCentré prtMedX, "DU CREDIT DOCUMENTAIRE"
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
frmElpPrt.prtCentré prtMedX, "N°  " & Trim(meYCDODOS0.CDODOSEXT)

XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 5

' >>>>>  Paragraphe concernant GARANTIE DE BONNE EXECUTION  <<<<<
For C = 1 To meCDO_Courrier.Garantie_Nb
    If Trim(meCDO_Courrier.Garantie(C)) <> "" Then
        XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
        XPrt.CurrentX = prtMinMarge
        XPrt.Print meCDO_Courrier.Garantie(C);
    End If
Next C

'XPrt.CurrentX = prtMinMarge
'XPrt.Print " GARANTIE DE BONNE EXECUTION :";

'XPrt.CurrentX = prtMinMarge
'XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'XPrt.Print " Conformément aux instructions de la banque émettrice, ce crédit ne deviendra opérationnel qu'après ";

'XPrt.CurrentX = prtMinMarge
'XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'XPrt.Print " réception d'une caution bancaire de bonne exécution fixée à 10 pour cent de la valeur globale du ";

'XPrt.CurrentX = prtMinMarge
'XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'XPrt.Print " contrat et son acceptation par l'ordonnateur. ";

'XPrt.CurrentX = prtMinMarge
'XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'XPrt.Print " L'instruction rendant le crédit opérationnel, vous sera notifiée par nos soins."; ";"

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2

XPrt.Print " Par conséquent, veuillez considérer ce crédit non encore opérationnel. ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print " Les frais et commissions étant stipulés à votre charge nous vous réclamerons lors de l'utilisation ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
Tx = mecnfYBIACDOCOM0.CDOCO2TX1 ' 2003.11.20  / 4       ' Taux de confirmation dans SAB : en annuel
Select Case mecnfYBIACDOCOM0.CDOCO2PER
    Case "M":   XPrt.Print " ou à la péremption du crédit notre commission de confirmation calculée au taux de : " & Format$(Tx, "### ##0.00") & " pour cent";
    Case "T":   XPrt.Print " ou à la péremption du crédit notre commission de confirmation calculée au taux de : ";
                XPrt.CurrentX = prtMinMarge + 2000
                XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
                XPrt.Print Format$(Tx, "### ##0.00") & " pour cent par trimestre indivisible";
    Case Else:  XPrt.Print " ou à la péremption du crédit notre commission de confirmation calculée au taux de : " & Format$(Tx, "### ##0.00") & " pour cent";
End Select

XPrt.CurrentX = prtMinMarge + 6000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.FontItalic = True: XPrt.FontBold = True
XPrt.Print " Devise               Montant ";
XPrt.FontBold = False: XPrt.FontItalic = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.FontItalic = True
XPrt.CurrentX = prtMinMarge
XPrt.Print " Commission de confirmation ";
Select Case mecnfYBIACDOCOM0.CDOCO2PER
    Case "M":   XPrt.CurrentX = prtMinMarge + 2700: XPrt.Print "pour le premier mois";
    Case "T":   XPrt.CurrentX = prtMinMarge + 2700: XPrt.Print "pour le premier trimestre";
    Case Else:  XPrt.CurrentX = prtMinMarge + 2700: XPrt.Print "";
End Select
XPrt.CurrentX = prtMinMarge + 6200: XPrt.Print mecnfYBIACDOCOM0.CDOCOMDEV & "            " & Format$(mecnfYBIACDOCOM0.CDOCOMMON, "### ### ### ##0.00");
XPrt.CurrentX = prtMinMarge

XPrt.FontItalic = False

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print " Par ailleurs, lors de l'utilisation nous vous décompterons : ";

XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "- notre commission de levée de documents au taux de : ";

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinMarge + 2000
XPrt.Print "1,5 pour mille sur le montant utilisé minimum EUR 106,71 plus TVA ";

XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "- frais de port et télex... ";

XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "- éventuellement une commission pour documents irréguliers de EUR 83,85 plus T.V.A. ";

XPrt.CurrentX = prtMinMarge + 5000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 4
XPrt.FontBold = True
XPrt.Print paramSignature;
XPrt.FontBold = False

End Sub


'---------------------------------------------------------
Public Sub prtCDO_Courrier_OUV_C_NOP_12_AVueRecl()
'---------------------------------------------------------
Dim X As String
Dim Tx As Double
Dim C As Integer

XPrt.CurrentY = prtCorpsY
XPrt.FontSize = 14: XPrt.FontBold = True
frmElpPrt.prtCentré prtMedX, "ANNEXE FAISANT PARTIE INTEGRANTE"
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
frmElpPrt.prtCentré prtMedX, "DU CREDIT DOCUMENTAIRE"
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
frmElpPrt.prtCentré prtMedX, "N°  " & Trim(meYCDODOS0.CDODOSEXT)

XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 5

' >>>>>  Paragraphe concernant GARANTIE DE BONNE EXECUTION  <<<<<
For C = 1 To meCDO_Courrier.Garantie_Nb
    If Trim(meCDO_Courrier.Garantie(C)) <> "" Then
        XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
        XPrt.CurrentX = prtMinMarge
        XPrt.Print meCDO_Courrier.Garantie(C);
    End If
Next C


'XPrt.CurrentX = prtMinMarge
'XPrt.Print " GARANTIE DE BONNE EXECUTION :";

'XPrt.CurrentX = prtMinMarge
'XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'XPrt.Print " Conformément aux instructions de la banque émettrice, ce crédit ne deviendra opérationnel qu'après ";

'XPrt.CurrentX = prtMinMarge
'XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'XPrt.Print " réception d'une caution bancaire de bonne exécution fixée à 10 pour cent de la valeur globale du ";

'XPrt.CurrentX = prtMinMarge
'XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'XPrt.Print " contrat et son acceptation par l'ordonnateur. ";

'XPrt.CurrentX = prtMinMarge
'XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'XPrt.Print " L'instruction rendant le crédit opérationnel, vous sera notifiée par nos soins."; ";"

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print " Par conséquent, veuillez considérer ce crédit non encore opérationnel. ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print " Compte tenu des instructions particulières de la banque émettrice au regard du paiement de nos ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print " commissions, veuillez nous reconnaître, à votre meilleure convenance de notre commission de ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
Tx = mecnfYBIACDOCOM0.CDOCO2TX1 ' 2003.11.20  / 4       ' Taux de confirmation dans SAB : en annuel
Select Case mecnfYBIACDOCOM0.CDOCO2PER
    Case "M":   XPrt.Print " confirmation calculée au taux de " & Format$(Tx, "### ##0.00") & " pour cent. ";
    Case "T":   XPrt.Print " confirmation calculée au taux de " & Format$(Tx, "### ##0.00") & " pour cent par trimestre indivisible. ";
    Case Else:  XPrt.Print " confirmation calculée au taux de " & Format$(Tx, "### ##0.00") & " pour cent. ";
End Select

XPrt.CurrentX = prtMinMarge + 6000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.FontItalic = True: XPrt.FontBold = True
XPrt.Print " Devise               Montant ";
XPrt.FontBold = False: XPrt.FontItalic = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.FontItalic = True
XPrt.CurrentX = prtMinMarge
XPrt.Print " Commission de confirmation ";
Select Case mecnfYBIACDOCOM0.CDOCO2PER
    Case "M":   XPrt.CurrentX = prtMinMarge + 2700: XPrt.Print "pour le premier mois";
    Case "T":   XPrt.CurrentX = prtMinMarge + 2700: XPrt.Print "pour le premier trimestre";
    Case Else:  XPrt.CurrentX = prtMinMarge + 2700: XPrt.Print "";
End Select
XPrt.CurrentX = prtMinMarge + 6200: XPrt.Print mecnfYBIACDOCOM0.CDOCOMDEV & "            " & Format$(mecnfYBIACDOCOM0.CDOCOMMON, "### ### ### ##0.00");
XPrt.CurrentX = prtMinMarge

XPrt.FontItalic = False

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print " Par ailleurs, lors de l'utilisation nous vous décompterons : ";

XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "- notre commission de levée de documents au taux de : ";

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinMarge + 2000
XPrt.Print "1,5 pour mille sur le montant utilisé minimum EUR 106,71 plus TVA ";

XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "- frais de port et télex... ";

XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "- éventuellement une commission pour documents irréguliers de EUR 83,85 plus T.V.A. ";

XPrt.CurrentX = prtMinMarge + 5000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 4
XPrt.FontBold = True
XPrt.Print paramSignature;
XPrt.FontBold = False

End Sub

'---------------------------------------------------------
Public Sub prtCDO_Courrier_OUV_C_NOP_13_PDifNonRecl()
'---------------------------------------------------------
Dim X As String
Dim Tx As Double
Dim C As Integer

XPrt.CurrentY = prtCorpsY
XPrt.FontSize = 14: XPrt.FontBold = True
frmElpPrt.prtCentré prtMedX, "ANNEXE FAISANT PARTIE INTEGRANTE"
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
frmElpPrt.prtCentré prtMedX, "DU CREDIT DOCUMENTAIRE"
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
frmElpPrt.prtCentré prtMedX, "N°  " & Trim(meYCDODOS0.CDODOSEXT)

XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2

' >>>>>  Paragraphe concernant GARANTIE DE BONNE EXECUTION  <<<<<
For C = 1 To meCDO_Courrier.Garantie_Nb
    If Trim(meCDO_Courrier.Garantie(C)) <> "" Then
        XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
        XPrt.CurrentX = prtMinMarge
        XPrt.Print meCDO_Courrier.Garantie(C);
    End If
Next C

'XPrt.CurrentX = prtMinMarge
'XPrt.Print " GARANTIE DE BONNE EXECUTION :";

'XPrt.CurrentX = prtMinMarge
'XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'XPrt.Print " Conformément aux instructions de la banque émettrice, ce crédit ne deviendra opérationnel qu'après ";

'XPrt.CurrentX = prtMinMarge
'XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'XPrt.Print " réception d'une caution bancaire de bonne exécution fixée à 10 pour cent de la valeur globale du ";

'XPrt.CurrentX = prtMinMarge
'XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'XPrt.Print " contrat et son acceptation par l'ordonnateur. ";

'XPrt.CurrentX = prtMinMarge
'XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'XPrt.Print " L'instruction rendant le crédit opérationnel, vous sera notifiée par nos soins."; ";"

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print " Par conséquent, veuillez considérer ce crédit non encore opérationnel. ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 4
XPrt.Print " Les frais et commissions étant stipulés à votre charge nous vous réclamerons lors de l'utilisation ou ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
Tx = mecnfYBIACDOCOM0.CDOCO2TX1 ' 2003.11.20  / 4       ' Taux de confirmation dans SAB : en annuel
Select Case mecnfYBIACDOCOM0.CDOCO2PER
    Case "M":   XPrt.Print " à la péremption du crédit notre commission de confirmation calculée au taux de : " & Format$(Tx, "### ##0.00") & " pour cent";
    Case "T":   XPrt.Print " à la péremption du crédit notre commission de confirmation calculée au taux de : ";
                XPrt.CurrentX = prtMinMarge + 2000
                XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
                XPrt.Print Format$(Tx, "### ##0.00") & " pour cent par trimestre indivisible";
    Case Else:  XPrt.Print " à la péremption du crédit notre commission de confirmation calculée au taux de : " & Format$(Tx, "### ##0.00") & " pour cent";
End Select

XPrt.CurrentX = prtMinMarge + 6000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.FontItalic = True: XPrt.FontBold = True
XPrt.Print " Devise               Montant ";
XPrt.FontBold = False: XPrt.FontItalic = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.FontItalic = True
XPrt.CurrentX = prtMinMarge
XPrt.Print " Commission de confirmation ";
Select Case mecnfYBIACDOCOM0.CDOCO2PER
    Case "M":   XPrt.CurrentX = prtMinMarge + 2700: XPrt.Print "pour le premier mois";
    Case "T":   XPrt.CurrentX = prtMinMarge + 2700: XPrt.Print "pour le premier trimestre";
    Case Else:  XPrt.CurrentX = prtMinMarge + 2700: XPrt.Print "";
End Select
XPrt.CurrentX = prtMinMarge + 6200: XPrt.Print mecnfYBIACDOCOM0.CDOCOMDEV & "            " & Format$(mecnfYBIACDOCOM0.CDOCOMMON, "### ### ### ##0.00");
XPrt.CurrentX = prtMinMarge

XPrt.FontItalic = False

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print " Par ailleurs, lors de l'utilisation nous vous décompterons : ";

XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "- notre commission de levée de documents au taux de : ";

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinMarge + 2000
XPrt.Print "1,5 pour mille sur le montant utilisé minimum EUR 106,71 plus TVA ";

XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
If meYCDODOS0.CDODOSMDI <> 0 Then
    XPrt.Print "- notre commission de paiement différé au taux de 0,10 % par mois ";
Else
    XPrt.Print "- notre commission d'acceptation au taux de 0,10 % par mois ";
End If

XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "- frais de port et télex. ";

XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "- éventuellement une commission pour documents irréguliers de EUR 83,85 plus T.V.A. ";

XPrt.CurrentX = prtMinMarge + 5000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 4
XPrt.FontBold = True
XPrt.Print paramSignature;
XPrt.FontBold = False

End Sub

'---------------------------------------------------------
Public Sub prtCDO_Courrier_OUV_C_NOP_14_PDifRecl()
'---------------------------------------------------------
Dim X As String
Dim Tx As Double
Dim C As Integer

XPrt.CurrentY = prtCorpsY
XPrt.FontSize = 14: XPrt.FontBold = True
frmElpPrt.prtCentré prtMedX, "ANNEXE FAISANT PARTIE INTEGRANTE"
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
frmElpPrt.prtCentré prtMedX, "DU CREDIT DOCUMENTAIRE"
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
frmElpPrt.prtCentré prtMedX, "N°  " & Trim(meYCDODOS0.CDODOSEXT)

XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2

' >>>>>  Paragraphe concernant GARANTIE DE BONNE EXECUTION  <<<<<
For C = 1 To meCDO_Courrier.Garantie_Nb
    If Trim(meCDO_Courrier.Garantie(C)) <> "" Then
        XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
        XPrt.CurrentX = prtMinMarge
        XPrt.Print meCDO_Courrier.Garantie(C);
    End If
Next C

'XPrt.CurrentX = prtMinMarge
'XPrt.Print " GARANTIE DE BONNE EXECUTION :";

'XPrt.CurrentX = prtMinMarge
'XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'XPrt.Print " Conformément aux instructions de la banque émettrice, ce crédit ne deviendra opérationnel qu'après ";

'XPrt.CurrentX = prtMinMarge
'XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'XPrt.Print " réception d'une caution bancaire de bonne exécution fixée à 10 pour cent de la valeur globale du ";

'XPrt.CurrentX = prtMinMarge
'XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'XPrt.Print " contrat et son acceptation par l'ordonnateur. ";

'XPrt.CurrentX = prtMinMarge
'XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'XPrt.Print " L'instruction rendant le crédit opérationnel, vous sera notifiée par nos soins."; ";"

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print " Par conséquent, veuillez considérer ce crédit non encore opérationnel. ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print " Compte tenu des instructions particulières de la banque émettrice au regard du paiement de nos ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print " commissions, veuillez nous reconnaître, à votre meilleure convenance de notre commission de ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
Tx = mecnfYBIACDOCOM0.CDOCO2TX1 ' 2003.11.20  / 4       ' Taux de confirmation dans SAB : en annuel
Select Case mecnfYBIACDOCOM0.CDOCO2PER
    Case "M":   XPrt.Print " confirmation calculée au taux de " & Format$(Tx, "### ##0.00") & " pour cent. ";
    Case "T":   XPrt.Print " confirmation calculée au taux de " & Format$(Tx, "### ##0.00") & " pour cent par trimestre indivisible. ";
    Case Else:  XPrt.Print " confirmation calculée au taux de " & Format$(Tx, "### ##0.00") & " pour cent. ";
End Select

XPrt.CurrentX = prtMinMarge + 6000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.FontItalic = True: XPrt.FontBold = True
XPrt.Print " Devise               Montant ";
XPrt.FontBold = False: XPrt.FontItalic = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.FontItalic = True
XPrt.CurrentX = prtMinMarge
XPrt.Print " Commission de confirmation ";
Select Case mecnfYBIACDOCOM0.CDOCO2PER
    Case "M":   XPrt.CurrentX = prtMinMarge + 2700: XPrt.Print "pour le premier mois";
    Case "T":   XPrt.CurrentX = prtMinMarge + 2700: XPrt.Print "pour le premier trimestre";
    Case Else:  XPrt.CurrentX = prtMinMarge + 2700: XPrt.Print "";
End Select
XPrt.CurrentX = prtMinMarge + 6200: XPrt.Print mecnfYBIACDOCOM0.CDOCOMDEV & "            " & Format$(mecnfYBIACDOCOM0.CDOCOMMON, "### ### ### ##0.00");
XPrt.CurrentX = prtMinMarge
XPrt.FontItalic = False

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print " Par ailleurs, lors de l'utilisation nous vous décompterons : ";

XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "- notre commission de levée de documents au taux de : ";

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinMarge + 2000
XPrt.Print "1,5 pour mille sur le montant utilisé minimum EUR 106,71 plus TVA ";

XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
If meYCDODOS0.CDODOSMDI <> 0 Then
    XPrt.Print "- notre commission de paiement différé au taux de 0,10 % par mois ";
Else
    XPrt.Print "- notre commission d'acceptation au taux de 0,10 % par mois ";
End If

XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "- frais de port et télex. ";

XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "- éventuellement une commission pour documents irréguliers de EUR 83,85 plus T.V.A. ";

XPrt.CurrentX = prtMinMarge + 5000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 4
XPrt.FontBold = True
XPrt.Print paramSignature;
XPrt.FontBold = False

End Sub

'---------------------------------------------------------
Public Sub prtCDO_Courrier_OUV_N_OP_20_NonRecl()
'---------------------------------------------------------
Dim X As String
Dim Tx As Double
Dim MntTTC As Currency

XPrt.CurrentY = prtCorpsY
XPrt.FontSize = 14: XPrt.FontBold = True
frmElpPrt.prtCentré prtMedX, "ANNEXE FAISANT PARTIE INTEGRANTE"
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
frmElpPrt.prtCentré prtMedX, "DU CREDIT DOCUMENTAIRE"
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
frmElpPrt.prtCentré prtMedX, "N°  " & Trim(meYCDODOS0.CDODOSEXT)

XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 5
XPrt.Print " Les frais et commissions étant stipulés à votre charge nous vous réclamerons lors de l'utilisation ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
Tx = menotYBIACDOCOM0.CDOCO2TX1 * 10      ' Taux de notification dans SAB : en % et non pour mille
XPrt.Print " ou à la péremption du crédit notre commission de notification calculée au taux de " & Format$(Tx, "### ##0.00") & " pour mille FLAT. ";

XPrt.FontItalic = True: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 3500
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.Print " Devise               Montant               T.V.A.";
XPrt.FontBold = False: XPrt.FontItalic = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge
XPrt.Print " Comm. notification HT (ou minimum) ";
XPrt.CurrentX = prtMinMarge + 3700: XPrt.Print menotYBIACDOCOM0.CDOCOMDEV;
XPrt.CurrentX = prtMinMarge + 5100: XPrt.Print Format$(menotYBIACDOCOM0.CDOCOMMON, "### ### ### ##0.00");

If menotYBIACDOCOM0.CDOCOMMTV <> 0 Then   ' SI montant TVA
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print " TVA 19,60 % ";
    XPrt.CurrentX = prtMinMarge + 3700: XPrt.Print menotYBIACDOCOM0.CDOCOMDEV;
    XPrt.CurrentX = prtMinMarge + 6500: XPrt.Print Format$(menotYBIACDOCOM0.CDOCOMMTV, "### ### ### ##0.00");
    
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    MntTTC = menotYBIACDOCOM0.CDOCOMMON + menotYBIACDOCOM0.CDOCOMMTV
    XPrt.Print " Montant TTC ";
    XPrt.CurrentX = prtMinMarge + 3700: XPrt.Print menotYBIACDOCOM0.CDOCOMDEV;
    XPrt.CurrentX = prtMinMarge + 5100: XPrt.Print Format$(MntTTC, "### ### ### ##0.00");
End If

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print " Par ailleurs, lors de l'utilisation nous vous décompterons : ";

XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "- notre commission de levée de documents au taux de : ";

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinMarge + 2500
XPrt.Print "1,5 pour mille sur le montant utilisé minimum EUR 106,71 plus TVA ";

XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "- frais de port et télex… ";

XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "- éventuellement une commission pour documents irréguliers de EUR 83,85 plus T.V.A. ";

XPrt.CurrentX = prtMinMarge + 5000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 4
XPrt.FontBold = True
XPrt.Print paramSignature;
XPrt.FontBold = False

End Sub

'---------------------------------------------------------
Public Sub prtCDO_Courrier_OUV_N_OP_21_Recl()
'---------------------------------------------------------
Dim X As String
Dim Tx As Double
Dim MntTTC As Currency

XPrt.CurrentY = prtCorpsY
XPrt.FontSize = 14: XPrt.FontBold = True
frmElpPrt.prtCentré prtMedX, "ANNEXE FAISANT PARTIE INTEGRANTE"
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
frmElpPrt.prtCentré prtMedX, "DU CREDIT DOCUMENTAIRE"
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
frmElpPrt.prtCentré prtMedX, "N°  " & Trim(meYCDODOS0.CDODOSEXT)

XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 5
XPrt.Print " Compte tenu des instructions particulières de la banque émettrice au regard du paiement de nos ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print " commissions, veuillez nous reconnaître, à votre meilleure convenance de notre commission de ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
Tx = menotYBIACDOCOM0.CDOCO2TX1 * 10      ' Taux de notification dans SAB : en % et non pour mille
XPrt.Print " notification calculée au taux de " & Format$(Tx, "### ##0.00") & " pour mille flat. ";

XPrt.FontItalic = True: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 3500
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.Print " Devise               Montant               T.V.A.";
XPrt.FontBold = False: XPrt.FontItalic = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge
XPrt.Print " Comm. notification HT (ou minimum) ";
XPrt.CurrentX = prtMinMarge + 3700: XPrt.Print menotYBIACDOCOM0.CDOCOMDEV;
XPrt.CurrentX = prtMinMarge + 5100: XPrt.Print Format$(menotYBIACDOCOM0.CDOCOMMON, "### ### ### ##0.00");

If menotYBIACDOCOM0.CDOCOMMTV <> 0 Then   ' SI montant TVA
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print " TVA 19,60 % ";
    XPrt.CurrentX = prtMinMarge + 3700: XPrt.Print menotYBIACDOCOM0.CDOCOMDEV;
    XPrt.CurrentX = prtMinMarge + 6500: XPrt.Print Format$(menotYBIACDOCOM0.CDOCOMMTV, "### ### ### ##0.00");
    
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    MntTTC = menotYBIACDOCOM0.CDOCOMMON + menotYBIACDOCOM0.CDOCOMMTV
    XPrt.Print " Montant TTC ";
    XPrt.CurrentX = prtMinMarge + 3700: XPrt.Print menotYBIACDOCOM0.CDOCOMDEV;
    XPrt.CurrentX = prtMinMarge + 5100: XPrt.Print Format$(MntTTC, "### ### ### ##0.00");
End If

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print " Par ailleurs, lors de l'utilisation nous vous décompterons : ";

XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "- notre commission de levée de documents au taux de : ";

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinMarge + 2500
XPrt.Print "1,5 pour mille sur le montant utilisé minimum EUR 106,71 plus TVA ";

XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "- frais de port et télex. ";

XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "- éventuellement une commission pour documents irréguliers de EUR 83,85 plus T.V.A. ";

XPrt.CurrentX = prtMinMarge + 5000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 4
XPrt.FontBold = True
XPrt.Print paramSignature;
XPrt.FontBold = False

End Sub

'---------------------------------------------------------
Public Sub prtCDO_Courrier_OUV_N_NOP_25_NonRecl()
'---------------------------------------------------------
Dim X As String
Dim Tx As Double
Dim MntTTC As Currency
Dim C As Integer

XPrt.CurrentY = prtCorpsY
XPrt.FontSize = 14: XPrt.FontBold = True
frmElpPrt.prtCentré prtMedX, "ANNEXE FAISANT PARTIE INTEGRANTE"
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
frmElpPrt.prtCentré prtMedX, "DU CREDIT DOCUMENTAIRE"
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
frmElpPrt.prtCentré prtMedX, "N°  " & Trim(meYCDODOS0.CDODOSEXT)

XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
' >>>>>  Paragraphe concernant GARANTIE DE BONNE EXECUTION  <<<<<
For C = 1 To meCDO_Courrier.Garantie_Nb
    If Trim(meCDO_Courrier.Garantie(C)) <> "" Then
        XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
        XPrt.CurrentX = prtMinMarge
        XPrt.Print meCDO_Courrier.Garantie(C);
    End If
Next C

'XPrt.CurrentX = prtMinMarge
'XPrt.Print " GARANTIE DE BONNE EXECUTION :";

'XPrt.CurrentX = prtMinMarge
'XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'XPrt.Print " Conformément aux instructions de la banque émettrice, ce crédit ne deviendra opérationnel qu'après ";

'XPrt.CurrentX = prtMinMarge
'XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'XPrt.Print " réception d'une caution bancaire de bonne exécution fixée à 10 pour cent de la valeur globale du ";

'XPrt.CurrentX = prtMinMarge
'XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'XPrt.Print " contrat et son acceptation par l'ordonnateur. ";

'XPrt.CurrentX = prtMinMarge
'XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'XPrt.Print " L'instruction rendant le crédit opérationnel, vous sera notifiée par nos soins."; ";"

XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print " Par conséquent, veuillez considérer ce crédit non encore opérationnel. ";
XPrt.FontBold = False

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print " Les frais et commissions étant stipulés à votre charge nous vous réclamerons lors de l'utilisation ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
Tx = menotYBIACDOCOM0.CDOCO2TX1 * 10      ' Taux de notification dans SAB : en % et non pour mille
XPrt.Print " ou à la péremption du crédit notre commission de notification calculée au taux de " & Format$(Tx, "### ##0.00") & " pour mille FLAT ";

XPrt.FontItalic = True: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 3500
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.Print " Devise               Montant               T.V.A.";
XPrt.FontBold = False: XPrt.FontItalic = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge
XPrt.Print " Comm. notification HT (ou minimum) ";
XPrt.CurrentX = prtMinMarge + 3700: XPrt.Print menotYBIACDOCOM0.CDOCOMDEV;
XPrt.CurrentX = prtMinMarge + 5100: XPrt.Print Format$(menotYBIACDOCOM0.CDOCOMMON, "### ### ### ##0.00");

If menotYBIACDOCOM0.CDOCOMMTV <> 0 Then   ' SI montant TVA
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print " TVA 19,60 % ";
    XPrt.CurrentX = prtMinMarge + 3700: XPrt.Print menotYBIACDOCOM0.CDOCOMDEV;
    XPrt.CurrentX = prtMinMarge + 6500: XPrt.Print Format$(menotYBIACDOCOM0.CDOCOMMTV, "### ### ### ##0.00");
    
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    MntTTC = menotYBIACDOCOM0.CDOCOMMON + menotYBIACDOCOM0.CDOCOMMTV
    XPrt.Print " Montant TTC ";
    XPrt.CurrentX = prtMinMarge + 3700: XPrt.Print menotYBIACDOCOM0.CDOCOMDEV;
    XPrt.CurrentX = prtMinMarge + 5100: XPrt.Print Format$(MntTTC, "### ### ### ##0.00");
End If

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print " Par ailleurs, lors de l'utilisation nous vous décompterons : ";

XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "- notre commission de levée de documents au taux de : ";

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinMarge + 2500
XPrt.Print "1,5 pour mille sur le montant utilisé minimum EUR 106,71 plus TVA ";

XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "- frais de port et télex. ";

XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "- éventuellement une commission pour documents irréguliers de EUR 83,85 plus T.V.A. ";

XPrt.CurrentX = prtMinMarge + 5000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 4
XPrt.FontBold = True
XPrt.Print paramSignature;
XPrt.FontBold = False

End Sub

'---------------------------------------------------------
Public Sub prtCDO_Courrier_OUV_N_NOP_26_Recl()
'---------------------------------------------------------
Dim X As String
Dim Tx As Double
Dim MntTTC As Currency
Dim C As Integer

XPrt.CurrentY = prtCorpsY
XPrt.FontSize = 14: XPrt.FontBold = True
frmElpPrt.prtCentré prtMedX, "ANNEXE FAISANT PARTIE INTEGRANTE"
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
frmElpPrt.prtCentré prtMedX, "DU CREDIT DOCUMENTAIRE"
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
frmElpPrt.prtCentré prtMedX, "N°  " & Trim(meYCDODOS0.CDODOSEXT)

XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
' >>>>>  Paragraphe concernant GARANTIE DE BONNE EXECUTION  <<<<<
For C = 1 To meCDO_Courrier.Garantie_Nb
    If Trim(meCDO_Courrier.Garantie(C)) <> "" Then
        XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
        XPrt.CurrentX = prtMinMarge
        XPrt.Print meCDO_Courrier.Garantie(C);
    End If
Next C

'XPrt.CurrentX = prtMinMarge
'XPrt.Print " GARANTIE DE BONNE EXECUTION :";

'XPrt.CurrentX = prtMinMarge
'XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'XPrt.Print " Conformément aux instructions de la banque émettrice, ce crédit ne deviendra opérationnel qu'après ";

'XPrt.CurrentX = prtMinMarge
'XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'XPrt.Print " réception d'une caution bancaire de bonne exécution fixée à 10 pour cent de la valeur globale du ";

'XPrt.CurrentX = prtMinMarge
'XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'XPrt.Print " contrat et son acceptation par l'ordonnateur. ";

'XPrt.CurrentX = prtMinMarge
'XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'XPrt.Print " L'instruction rendant le crédit opérationnel, vous sera notifiée par nos soins."; ";"

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print " Par conséquent, veuillez considérer ce crédit non encore opérationnel. ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print " Compte tenu des instructions particulières de la banque émettrice au regard du paiement de nos ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print " commissions, veuillez nous reconnaître, à votre meilleure convenance de notre commission de ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
Tx = menotYBIACDOCOM0.CDOCO2TX1 * 10      ' Taux de notification dans SAB : en % et non pour mille
XPrt.Print " notification calculée au taux de " & Format$(Tx, "### ##0.00") & " pour mille flat.";

XPrt.FontItalic = True: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 3500
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.Print " Devise               Montant               T.V.A.";
XPrt.FontBold = False: XPrt.FontItalic = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge
XPrt.Print " Comm. notification HT (ou minimum) ";
XPrt.CurrentX = prtMinMarge + 3700: XPrt.Print menotYBIACDOCOM0.CDOCOMDEV;
XPrt.CurrentX = prtMinMarge + 5100: XPrt.Print Format$(menotYBIACDOCOM0.CDOCOMMON, "### ### ### ##0.00");

If menotYBIACDOCOM0.CDOCOMMTV <> 0 Then   ' SI montant TVA
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print " TVA 19,60 % ";
    XPrt.CurrentX = prtMinMarge + 3700: XPrt.Print menotYBIACDOCOM0.CDOCOMDEV;
    XPrt.CurrentX = prtMinMarge + 6500: XPrt.Print Format$(menotYBIACDOCOM0.CDOCOMMTV, "### ### ### ##0.00");
    
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    MntTTC = menotYBIACDOCOM0.CDOCOMMON + menotYBIACDOCOM0.CDOCOMMTV
    XPrt.Print " Montant TTC ";
    XPrt.CurrentX = prtMinMarge + 3700: XPrt.Print menotYBIACDOCOM0.CDOCOMDEV;
    XPrt.CurrentX = prtMinMarge + 5100: XPrt.Print Format$(MntTTC, "### ### ### ##0.00");
End If

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print " Par ailleurs, lors de l'utilisation nous vous décompterons : ";

XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "- notre commission de levée de documents au taux de : ";

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinMarge + 2500
XPrt.Print "1,5 pour mille sur le montant utilisé minimum EUR 106,71 plus TVA ";

XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "- frais de port et télex… ";

XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "- éventuellement une commission pour documents irréguliers de EUR 83,85 plus T.V.A. ";

XPrt.CurrentX = prtMinMarge + 5000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 4
XPrt.FontBold = True
XPrt.Print paramSignature;
XPrt.FontBold = False

End Sub


'---------------------------------------------------------
Public Sub prtCDO_Courrier_OUV_P_OP_32_AVueNonRecl()
'---------------------------------------------------------
Dim X As String
Dim Tx As Double
 
XPrt.CurrentY = prtCorpsY
XPrt.FontSize = 14: XPrt.FontBold = True
frmElpPrt.prtCentré prtMedX, "ANNEXE FAISANT PARTIE INTEGRANTE"
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
frmElpPrt.prtCentré prtMedX, "DU CREDIT DOCUMENTAIRE"
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
frmElpPrt.prtCentré prtMedX, "N°  " & Trim(meYCDODOS0.CDODOSEXT)

XPrt.FontSize = 11: XPrt.FontBold = False
 
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 6
XPrt.CurrentX = prtMinMarge
XPrt.Print " Les frais et commissions étant stipulés à votre charge nous vous réclamerons lors de l'utilisation ou à ";
 
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print " la péremption du crédit : ";
 
XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
Tx = mecnfYBIACDOCOM0.CDOCO2TX1 ' 2003.11.20  / 4       ' Taux de confirmation dans SAB : en annuel
XPrt.Print " - notre commission de confirmation calculée au taux de " & Format$(Tx, "### ##0.00") & " pour cent par trimestre ";
 
XPrt.CurrentX = prtMinMarge + 1100
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print " indivisible sur la partie confirmée soit pour le premier trimestre " & mecnfYBIACDOCOM0.CDOCOMDEV & "  " & Format$(mecnfYBIACDOCOM0.CDOCOMMON, "### ### ### ##0.00");
 
XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
Tx = menotYBIACDOCOM0.CDOCO2TX1 * 10      ' Taux de notification dans SAB : en % et non pour mille
XPrt.Print " - notre commission de notification calculée au taux de " & Format$(Tx, "### ##0.00") & " pour mille flat plus TVA ";
 
XPrt.CurrentX = prtMinMarge + 1100
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print " sur la partie non confirmée soit " & menotYBIACDOCOM0.CDOCOMDEV & "  " & Format$(menotYBIACDOCOM0.CDOCOMMON, "### ### ### ##0.00");
 
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print " Par ailleurs, lors de l'utilisation nous vous décompterons : ";
 
XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "- notre commission de levée de documents au taux de : ";
 
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinMarge + 2000
XPrt.Print "1,5 pour mille sur le montant utilisé minimum EUR 106,71 plus TVA ";
 
XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "- frais de port et télex… ";
 
XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "- éventuellement une commission pour documents irréguliers de EUR 83,85 plus T.V.A. ";
 
XPrt.CurrentX = prtMinMarge + 5000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 4
XPrt.FontBold = True
XPrt.Print paramSignature;
XPrt.FontBold = False
 
End Sub

'---------------------------------------------------------
Public Sub prtCDO_Courrier_OUV_P_OP_33_AVueRecl()
'---------------------------------------------------------
Dim X As String
Dim Tx As Double
 
XPrt.CurrentY = prtCorpsY
XPrt.FontSize = 14: XPrt.FontBold = True
frmElpPrt.prtCentré prtMedX, "ANNEXE FAISANT PARTIE INTEGRANTE"
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
frmElpPrt.prtCentré prtMedX, "DU CREDIT DOCUMENTAIRE"
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
frmElpPrt.prtCentré prtMedX, "N°  " & Trim(meYCDODOS0.CDODOSEXT)
XPrt.FontSize = 11: XPrt.FontBold = False
 
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 6
XPrt.CurrentX = prtMinMarge
XPrt.Print " Compte tenu des instructions particulières de la banque émettrice au regard du paiement de nos commissions,";
 
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print " veuillez nous reconnaître, à votre meilleure convenance de " & meYCDODOS0.CDODOSDEV & "  " & Format$(mecnfYBIACDOCOM0.CDOCOMMON + menotYBIACDOCOM0.CDOCOMMON, "### ### ### ##0.00") & " représentant : ";
 
XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
Tx = mecnfYBIACDOCOM0.CDOCO2TX1 ' 2003.11.20  / 4       ' Taux de confirmation dans SAB : en annuel
XPrt.Print " - notre commission de confirmation calculée au taux de " & Format$(Tx, "### ##0.00") & " pour cent par trimestre ";
 
XPrt.CurrentX = prtMinMarge + 1100
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print " indivisible sur la partie confirmée ";
 
XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
Tx = menotYBIACDOCOM0.CDOCO2TX1 * 10      ' Taux de notification dans SAB : en % et non pour mille
XPrt.Print " - notre commission de notification calculée au taux de " & Format$(Tx, "### ##0.00") & " pour mille flat plus TVA ";
 
XPrt.CurrentX = prtMinMarge + 1100
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print " sur la partie non confirmée.";
 
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print " Par ailleurs, lors de l'utilisation nous vous décompterons : ";
 
XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "- notre commission de levée de documents au taux de : ";
 
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinMarge + 2000
XPrt.Print "1,5 pour mille sur le montant utilisé minimum EUR 106,71 plus TVA ";
 
XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "- frais de port et télex… ";
 
XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "- éventuellement une commission pour documents irréguliers de EUR 83,85 plus T.V.A. ";
 
XPrt.CurrentX = prtMinMarge + 5000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 4
XPrt.FontBold = True
XPrt.Print paramSignature;
XPrt.FontBold = False
 
End Sub

'---------------------------------------------------------
Public Sub prtCDO_Courrier_OUV_P_OP_34_PDifNonRecl()
'---------------------------------------------------------
Dim X As String
Dim Tx As Double
 
XPrt.CurrentY = prtCorpsY
XPrt.FontSize = 14: XPrt.FontBold = True
frmElpPrt.prtCentré prtMedX, "ANNEXE FAISANT PARTIE INTEGRANTE"
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
frmElpPrt.prtCentré prtMedX, "DU CREDIT DOCUMENTAIRE"
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
frmElpPrt.prtCentré prtMedX, "N°  " & Trim(meYCDODOS0.CDODOSEXT)
 
XPrt.FontSize = 11: XPrt.FontBold = False
 
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 6
XPrt.CurrentX = prtMinMarge
XPrt.Print " Les frais et commissions étant stipulés à votre charge nous vous réclamerons lors de l'utilisation ou à ";
 
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print " la péremption du crédit : ";
 
XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
Tx = mecnfYBIACDOCOM0.CDOCO2TX1 ' 2003.11.20  / 4       ' Taux de confirmation dans SAB : en annuel
XPrt.Print " - notre commission de confirmation calculée au taux de " & Format$(Tx, "### ##0.00") & " pour cent par trimestre ";
 
XPrt.CurrentX = prtMinMarge + 1100
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print " indivisible sur la partie confirmée soit pour le premier trimestre " & mecnfYBIACDOCOM0.CDOCOMDEV & "  " & Format$(mecnfYBIACDOCOM0.CDOCOMMON, "### ### ### ##0.00");
 
XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
Tx = menotYBIACDOCOM0.CDOCO2TX1 * 10      ' Taux de notification dans SAB : en % et non pour mille
XPrt.Print " - notre commission de notification calculée au taux de " & Format$(Tx, "### ##0.00") & " pour mille flat plus TVA sur la ";
 
XPrt.CurrentX = prtMinMarge + 1100
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print " partie non confirmée soit " & menotYBIACDOCOM0.CDOCOMDEV & "  " & Format$(menotYBIACDOCOM0.CDOCOMMON, "### ### ### ##0.00");
 
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print " Par ailleurs, lors de l'utilisation nous vous décompterons : ";
 
XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "- notre commission de levée de documents au taux de : ";
 
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinMarge + 2000
XPrt.Print "1,5 pour mille sur le montant utilisé minimum EUR 106,71 plus TVA ";
 
XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "- notre commission de paiement différé au taux de 0,10 % par mois, minimum EUR 152,44 ";
 
XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "- frais de port et télex… ";

XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "- éventuellement une commission pour documents irréguliers de EUR 83,85 plus T.V.A. ";
 
XPrt.CurrentX = prtMinMarge + 5000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 4
XPrt.FontBold = True
XPrt.Print paramSignature;
XPrt.FontBold = False
 
End Sub


'---------------------------------------------------------
Public Sub prtCDO_Courrier_OUV_P_OP_35_PDifRecl()
'---------------------------------------------------------
Dim X As String
Dim Tx As Double
 
XPrt.CurrentY = prtCorpsY
XPrt.FontSize = 14: XPrt.FontBold = True
frmElpPrt.prtCentré prtMedX, "ANNEXE FAISANT PARTIE INTEGRANTE"
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
frmElpPrt.prtCentré prtMedX, "DU CREDIT DOCUMENTAIRE"
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
frmElpPrt.prtCentré prtMedX, "N°  " & Trim(meYCDODOS0.CDODOSEXT)
 
XPrt.FontSize = 11: XPrt.FontBold = False
 
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 5
XPrt.CurrentX = prtMinMarge
XPrt.Print " Compte tenu des instructions particulières de la banque émettrice au regard du paiement de nos commissions, ";
 
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print " veuillez nous reconnaître, à votre meilleure convenance de " & meYCDODOS0.CDODOSDEV & "  " & Format$(mecnfYBIACDOCOM0.CDOCOMMON + menotYBIACDOCOM0.CDOCOMMON, "### ### ### ##0.00") & " représentant : ";
 
XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
Tx = mecnfYBIACDOCOM0.CDOCO2TX1 ' 2003.11.20  / 4       ' Taux de confirmation dans SAB : en annuel
XPrt.Print " - notre commission de confirmation calculée au taux de " & Format$(Tx, "### ##0.00") & " pour cent par trimestre ";
 
XPrt.CurrentX = prtMinMarge + 1100
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print " indivisible sur la partie confirmée ";
 
XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
Tx = menotYBIACDOCOM0.CDOCO2TX1 * 10      ' Taux de notification dans SAB : en % et non pour mille
XPrt.Print " - notre commission de notification calculée au taux de " & Format$(Tx, "### ##0.00") & " pour mille flat plus TVA ";
 
XPrt.CurrentX = prtMinMarge + 1100
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print " sur la partie non confirmée.";
 
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print " Par ailleurs, lors de l'utilisation nous vous décompterons : ";
 
XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "- notre commission de levée de documents au taux de : ";
 
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinMarge + 2000
XPrt.Print "1,5 pour mille sur le montant utilisé minimum EUR 106,71 plus TVA ";
 
XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "- notre commission de paiement différé au taux de 0.10 % par mois, minimum EUR 152,44 ";
 
XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "- frais de port et télex… ";
 
XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "- éventuellement une commission pour documents irréguliers de EUR 83,85 plus T.V.A. ";
 
XPrt.CurrentX = prtMinMarge + 5000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 4
XPrt.FontBold = True
XPrt.Print paramSignature;
XPrt.FontBold = False
 
End Sub

'---------------------------------------------------------
Public Sub prtCDO_Courrier_OUV_P_NOP_39_AVueNonRecl()
'---------------------------------------------------------
Dim X As String
Dim Tx As Double
Dim C As Integer

XPrt.CurrentY = prtCorpsY
XPrt.FontSize = 14: XPrt.FontBold = True
frmElpPrt.prtCentré prtMedX, "ANNEXE FAISANT PARTIE INTEGRANTE"
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
frmElpPrt.prtCentré prtMedX, "DU CREDIT DOCUMENTAIRE"
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
frmElpPrt.prtCentré prtMedX, "N°  " & Trim(meYCDODOS0.CDODOSEXT)

XPrt.FontSize = 11: XPrt.FontBold = False
 
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
' >>>>>  Paragraphe concernant GARANTIE DE BONNE EXECUTION  <<<<<
For C = 1 To meCDO_Courrier.Garantie_Nb
    If Trim(meCDO_Courrier.Garantie(C)) <> "" Then
        XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
        XPrt.CurrentX = prtMinMarge
        XPrt.Print meCDO_Courrier.Garantie(C);
    End If
Next C

'XPrt.CurrentX = prtMinMarge
'XPrt.Print " GARANTIE DE BONNE EXECUTION :";

'XPrt.CurrentX = prtMinMarge
'XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'XPrt.Print " Conformément aux instructions de la banque émettrice, ce crédit ne deviendra opérationnel qu'après ";

'XPrt.CurrentX = prtMinMarge
'XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'XPrt.Print " réception d'une caution bancaire de bonne exécution fixée à 10 pour cent de la valeur globale du ";

'XPrt.CurrentX = prtMinMarge
'XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'XPrt.Print " contrat et son acceptation par l'ordonnateur. ";

'XPrt.CurrentX = prtMinMarge
'XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'XPrt.Print " L'instruction rendant le crédit opérationnel, vous sera notifiée par nos soins."; ";"

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print " Par conséquent, veuillez considérer ce crédit non encore opérationnel. ";

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.CurrentX = prtMinMarge
XPrt.Print " Les frais et commissions étant stipulés à votre charge nous vous réclamerons lors de l'utilisation ";
 
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print " ou à la péremption du crédit : ";
 
XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
Tx = mecnfYBIACDOCOM0.CDOCO2TX1 ' 2003.11.20  / 4       ' Taux de confirmation dans SAB : en annuel
XPrt.Print " - notre commission de confirmation calculée au taux de " & Format$(Tx, "### ##0.00") & " pour cent par trimestre ";
 
XPrt.CurrentX = prtMinMarge + 1100
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print " indivisible sur la partie confirmée ";
 
 
XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
Tx = menotYBIACDOCOM0.CDOCO2TX1 * 10      ' Taux de notification dans SAB : en % et non pour mille
XPrt.Print " - notre commission de notification calculée au taux de " & Format$(Tx, "### ##0.00") & " pour mille flat plus TVA ";
 
XPrt.CurrentX = prtMinMarge + 1100
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print " sur la partie non confirmée.";
 
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print " Par ailleurs, lors de l'utilisation nous vous décompterons : ";
 
XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "- notre commission de levée de documents au taux de : ";
 
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinMarge + 2000
XPrt.Print "1,5 pour mille sur le montant utilisé minimum EUR 106,71 plus TVA ";
 
XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "- frais de port et télex… ";
 
XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "- éventuellement une commission pour documents irréguliers de EUR 83,85 plus T.V.A. ";
 
XPrt.CurrentX = prtMinMarge + 5000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 4
XPrt.FontBold = True
XPrt.Print paramSignature;
XPrt.FontBold = False
 
End Sub

'---------------------------------------------------------
Public Sub prtCDO_Courrier_OUV_P_NOP_40_AVueRecl()
'---------------------------------------------------------
Dim X As String
Dim Tx As Double
Dim C As Integer

XPrt.CurrentY = prtCorpsY
XPrt.FontSize = 14: XPrt.FontBold = True
frmElpPrt.prtCentré prtMedX, "ANNEXE FAISANT PARTIE INTEGRANTE"
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
frmElpPrt.prtCentré prtMedX, "DU CREDIT DOCUMENTAIRE"
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
frmElpPrt.prtCentré prtMedX, "N°  " & Trim(meYCDODOS0.CDODOSEXT)
 
XPrt.FontSize = 11: XPrt.FontBold = False
 
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
' >>>>>  Paragraphe concernant GARANTIE DE BONNE EXECUTION  <<<<<
For C = 1 To meCDO_Courrier.Garantie_Nb
    If Trim(meCDO_Courrier.Garantie(C)) <> "" Then
        XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
        XPrt.CurrentX = prtMinMarge
        XPrt.Print meCDO_Courrier.Garantie(C);
    End If
Next C

'XPrt.CurrentX = prtMinMarge
'XPrt.Print " GARANTIE DE BONNE EXECUTION :";

'XPrt.CurrentX = prtMinMarge
'XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'XPrt.Print " Conformément aux instructions de la banque émettrice, ce crédit ne deviendra opérationnel qu'après ";

'XPrt.CurrentX = prtMinMarge
'XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'XPrt.Print " réception d'une caution bancaire de bonne exécution fixée à 10 pour cent de la valeur globale du ";

'XPrt.CurrentX = prtMinMarge
'XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'XPrt.Print " contrat et son acceptation par l'ordonnateur. ";

'XPrt.CurrentX = prtMinMarge
'XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'XPrt.Print " L'instruction rendant le crédit opérationnel, vous sera notifiée par nos soins."; ";"


XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print " Par conséquent, veuillez considérer ce crédit non encore opérationnel. ";

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.CurrentX = prtMinMarge
XPrt.Print " Compte tenu des instructions particulières de la banque émettrice au regard du paiement de nos commissions, ";
 
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print " veuillez nous reconnaître, à votre meilleure convenance de " & meYCDODOS0.CDODOSDEV & "  " & Format$(mecnfYBIACDOCOM0.CDOCOMMON + menotYBIACDOCOM0.CDOCOMMON, "### ### ### ##0.00") & " représentant : ";
 
XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
Tx = mecnfYBIACDOCOM0.CDOCO2TX1 ' 2003.11.20  / 4       ' Taux de confirmation dans SAB : en annuel
XPrt.Print " - notre commission de confirmation calculée au taux de " & Format$(Tx, "### ##0.00") & " pour cent par trimestre ";
 
XPrt.CurrentX = prtMinMarge + 1100
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print " indivisible sur la partie confirmée ";
 
XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
Tx = menotYBIACDOCOM0.CDOCO2TX1 * 10      ' Taux de notification dans SAB : en % et non pour mille
XPrt.Print " - notre commission de notification calculée au taux de " & Format$(Tx, "### ##0.00") & " pour mille flat plus TVA ";
 
XPrt.CurrentX = prtMinMarge + 1100
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print " sur la partie non confirmée.";
 
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print " Par ailleurs, lors de l'utilisation nous vous décompterons : ";
 
XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "- notre commission de levée de documents au taux de : ";
 
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinMarge + 2000
XPrt.Print "1,5 pour mille sur le montant utilisé minimum EUR 106,71 plus TVA ";
 
XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "- frais de port et télex… ";
 
XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "- éventuellement une commission pour documents irréguliers de EUR 83,85 plus T.V.A. ";
 
XPrt.CurrentX = prtMinMarge + 5000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 4
XPrt.FontBold = True
XPrt.Print paramSignature;
XPrt.FontBold = False
 
End Sub

'---------------------------------------------------------
Public Sub prtCDO_Courrier_OUV_P_NOP_41_PDifNonRecl()
'---------------------------------------------------------
Dim X As String
Dim Tx As Double
Dim C As Integer

XPrt.CurrentY = prtCorpsY
XPrt.FontSize = 14: XPrt.FontBold = True
frmElpPrt.prtCentré prtMedX, "ANNEXE FAISANT PARTIE INTEGRANTE"
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
frmElpPrt.prtCentré prtMedX, "DU CREDIT DOCUMENTAIRE"
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
frmElpPrt.prtCentré prtMedX, "N°  " & Trim(meYCDODOS0.CDODOSEXT)

XPrt.FontSize = 11: XPrt.FontBold = False
 
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2

' >>>>>  Paragraphe concernant GARANTIE DE BONNE EXECUTION  <<<<<
For C = 1 To meCDO_Courrier.Garantie_Nb
    If Trim(meCDO_Courrier.Garantie(C)) <> "" Then
        XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
        XPrt.CurrentX = prtMinMarge
        XPrt.Print meCDO_Courrier.Garantie(C);
    End If
Next C

'XPrt.CurrentX = prtMinMarge
'XPrt.Print " GARANTIE DE BONNE EXECUTION :";

'XPrt.CurrentX = prtMinMarge
'XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'XPrt.Print " Conformément aux instructions de la banque émettrice, ce crédit ne deviendra opérationnel qu'après ";

'XPrt.CurrentX = prtMinMarge
'XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'XPrt.Print " réception d'une caution bancaire de bonne exécution fixée à 10 pour cent de la valeur globale du ";

'XPrt.CurrentX = prtMinMarge
'XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'XPrt.Print " contrat et son acceptation par l'ordonnateur. ";

'XPrt.CurrentX = prtMinMarge
'XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'XPrt.Print " L'instruction rendant le crédit opérationnel, vous sera notifiée par nos soins."; ";"


XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print " Par conséquent, veuillez considérer ce crédit non encore opérationnel. ";
 

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.CurrentX = prtMinMarge
XPrt.Print " Les frais et commissions étant stipulés à votre charge nous vous réclamerons lors de l'utilisation ";
 
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print " ou à la péremption du crédit : ";
 
XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
Tx = mecnfYBIACDOCOM0.CDOCO2TX1 ' 2003.11.20  / 4       ' Taux de confirmation dans SAB : en annuel
XPrt.Print " - notre commission de confirmation calculée au taux de " & Format$(Tx, "### ##0.00") & " pour cent par trimestre ";
 
XPrt.CurrentX = prtMinMarge + 1100
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print " indivisible sur la partie confirmée ";
 
 
XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
Tx = menotYBIACDOCOM0.CDOCO2TX1 * 10      ' Taux de notification dans SAB : en % et non pour mille
XPrt.Print " - notre commission de notification calculée au taux de " & Format$(Tx, "### ##0.00") & " pour mille flat plus TVA ";

XPrt.CurrentX = prtMinMarge + 1100
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print " sur la partie non confirmée.";
 
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print " Par ailleurs, lors de l'utilisation nous vous décompterons : ";
 
XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "- notre commission de levée de documents au taux de : ";
 
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinMarge + 2000
XPrt.Print "1,5 pour mille sur le montant utilisé minimum EUR 106,71 plus TVA ";
 
XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "- notre commission de paiement différé au taux de 0,10 % par mois minimum EUR 152,44";
 
XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "- frais de port et télex… ";

XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "- éventuellement une commission pour documents irréguliers de EUR 83,85 plus T.V.A. ";
 
XPrt.CurrentX = prtMinMarge + 5000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 4
XPrt.FontBold = True
XPrt.Print paramSignature;
XPrt.FontBold = False
 
End Sub


'---------------------------------------------------------
Public Sub prtCDO_Courrier_OUV_P_NOP_42_PDifRecl()
'---------------------------------------------------------
Dim X As String
Dim Tx As Double
Dim C As Integer

XPrt.CurrentY = prtCorpsY
XPrt.FontSize = 14: XPrt.FontBold = True
frmElpPrt.prtCentré prtMedX, "ANNEXE FAISANT PARTIE INTEGRANTE"
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
frmElpPrt.prtCentré prtMedX, "DU CREDIT DOCUMENTAIRE"
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
frmElpPrt.prtCentré prtMedX, "N°  " & Trim(meYCDODOS0.CDODOSEXT)

XPrt.FontSize = 11: XPrt.FontBold = False
 
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
' >>>>>  Paragraphe concernant GARANTIE DE BONNE EXECUTION  <<<<<
For C = 1 To meCDO_Courrier.Garantie_Nb
    If Trim(meCDO_Courrier.Garantie(C)) <> "" Then
        XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
        XPrt.CurrentX = prtMinMarge
        XPrt.Print meCDO_Courrier.Garantie(C);
    End If
Next C


'XPrt.CurrentX = prtMinMarge
'XPrt.Print " GARANTIE DE BONNE EXECUTION :";

'XPrt.CurrentX = prtMinMarge
'XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'XPrt.Print " Conformément aux instructions de la banque émettrice, ce crédit ne deviendra opérationnel qu'après ";

'XPrt.CurrentX = prtMinMarge
'XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'XPrt.Print " réception d'une caution bancaire de bonne exécution fixée à 10 pour cent de la valeur globale du ";

'XPrt.CurrentX = prtMinMarge
'XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'XPrt.Print " contrat et son acceptation par l'ordonnateur. ";

'XPrt.CurrentX = prtMinMarge
'XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'XPrt.Print " L'instruction rendant le crédit opérationnel, vous sera notifiée par nos soins."; ";"

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print " Par conséquent, veuillez considérer ce crédit non encore opérationnel. ";


XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.CurrentX = prtMinMarge
XPrt.Print " Compte tenu des instructions particulières de la banque émettrice au regard du paiement de nos ";
 
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print " commissions, veuillez nous reconnaître, à votre meilleure convenance de EUR :xxxx ";
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print " représentant : ";
 
XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
Tx = mecnfYBIACDOCOM0.CDOCO2TX1 ' 2003.11.20  / 4       ' Taux de confirmation dans SAB : en annuel
XPrt.Print " - notre commission de confirmation calculée au taux de " & Format$(Tx, "### ##0.00") & " pour cent par trimestre ";
 
XPrt.CurrentX = prtMinMarge + 1100
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print " indivisible sur la partie confirmée ";
 
XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
Tx = menotYBIACDOCOM0.CDOCO2TX1 * 10      ' Taux de notification dans SAB : en % et non pour mille
XPrt.Print " - notre commission de notification calculée au taux de " & Format$(Tx, "### ##0.00") & " pour mille flat plus TVA ";
 
XPrt.CurrentX = prtMinMarge + 1100
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print " sur la partie non confirmée.";
 
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print " Par ailleurs, lors de l'utilisation nous vous décompterons : ";
 
XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "- notre commission de levée de documents au taux de : ";
 
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinMarge + 2000
XPrt.Print "1,5 pour mille sur le montant utilisé minimum EUR 106,71 plus TVA ";
 
XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "- notre commission de paiement différé au taux de 0,10 % par mois minimum EUR 152,44";
 
XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "- frais de port et télex… ";

XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "- éventuellement une commission pour documents irréguliers de EUR 83,85 plus T.V.A. ";
 
XPrt.CurrentX = prtMinMarge + 5000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 4
XPrt.FontBold = True
XPrt.Print paramSignature;
XPrt.FontBold = False
 
End Sub

'---------------------------------------------------------
Public Sub prtCDO_Courrier_OUV_NOT_Page1()
'---------------------------------------------------------
Dim X As String
Line2_Ecart = 5: prtCDO_Courrier_Trame (Line2_Ecart)
'Call frmElpPrt.prtTrame(prtMinMarge - 50, prtCorpsY - 100, prtMaxMarge, prtCorpsY + prtlineHeight * 5 + 100, "B", 245)

XPrt.CurrentY = prtCorpsY
XPrt.FontSize = 14: XPrt.FontBold = True: XPrt.FontUnderline = True
frmElpPrt.prtCentré prtMedX, "CREDIT DOCUMENTAIRE"

XPrt.FontSize = 11: XPrt.FontBold = False: XPrt.FontUnderline = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
'Call frmElpPrt.prtTrame(prtMinMarge - 50, XPrt.CurrentY - 50, prtMaxMarge, XPrt.CurrentY + prtlineHeight * 3 + 50, , 245)
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Donneur d'ordre";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1700: XPrt.Print meCDO_Courrier.DON_Concat;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Bénéficiaire";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1700: XPrt.Print meCDO_Courrier.BEN_Concat;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Montant";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentX = prtMinMarge + 4100: XPrt.Print "Devise :";
XPrt.CurrentX = prtMinMarge + 7500: XPrt.Print "Validité :";
XPrt.FontBold = True
X = Trim(Format$(meYCDODOS0.CDODOSMON, "### ### ### ##0.00"))
XPrt.CurrentX = prtMinMarge + 1700 '4000 - XPrt.TextWidth(X)
XPrt.Print X;
XPrt.CurrentX = prtMinMarge + 5000: XPrt.Print meYCDODOS0.CDODOSDEV;
XPrt.CurrentX = prtMinMarge + 8500: XPrt.Print wDate_Validité;
XPrt.FontBold = False

XPrt.FontSize = 11
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 4
XPrt.CurrentX = prtMinMarge
XPrt.Print "Messieurs,";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Nous vous informons de l'ouverture du CREDIT DOCUMENTAIRE irrévocable N° ";
XPrt.FontBold = True
XPrt.Print meYCDODOS0.CDODOSEXT;
XPrt.FontBold = False

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "émis en votre faveur par notre correspondant :";
XPrt.CurrentX = XPrt.CurrentX + 500
prtAdresse meCDO_Courrier.BQE_ZADRESS0, False

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Valable à nos guichets jusqu'au :   ";
XPrt.FontBold = True
XPrt.Print wDate_Validité;
XPrt.FontBold = False
XPrt.Print " pour présentation des documents pour paiement dans les ";
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "conditions reprises sur ";
XPrt.FontBold = True
XPrt.Print wAnnexe_Nb;
XPrt.FontBold = False
XPrt.Print " faisant partie intégrante de ce crédit.";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print " Ce crédit documentaire n'étant pas confirmé par notre établissement il est bien entendu que cette opération ne ";
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print " comporte aucun engagement de notre part quant au règlement des documents que nous transmettons à la banque ";
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print " émettrice lors de l'utilisation.";
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print " Nous ne vous effectuerons le règlement qu'après réception de la couverture correspondante.";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Lors de l'utilisation,  nous vous remercions de bien vouloir accompagner ";
XPrt.FontBold = True: XPrt.ForeColor = vbBlue
XPrt.Print "IMPERATIVEMENT";
XPrt.FontBold = False: XPrt.ForeColor = prtForeColor
XPrt.Print " les documents";
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "d'un exemplaire supplémentaire de votre facture et d'un ";
XPrt.FontBold = True: XPrt.ForeColor = vbBlue
XPrt.Print "relevé d'identité bancaire IBAN";
XPrt.FontBold = False: XPrt.ForeColor = prtForeColor
XPrt.Print " pour nous permettre";
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "de vous en effectuer le règlement et de nous communiquer votre ";
XPrt.FontBold = True: XPrt.ForeColor = vbBlue
XPrt.Print "n° de TVA intracommunautaire.";
XPrt.FontBold = False: XPrt.ForeColor = prtForeColor

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Ce crédit documentaire est soumis aux Règles et Usances Uniformes relatives aux crédits documentaires";
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "(version révisée de 2007 - Publication N° 600 de la Chambre de Commerce Internationale).";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Si vous n'êtes pas d'accord sur les conditions de ce crédit, nous vous conseillons de vous mettre DIRECTEMENT";
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "en rapport avec vos acheteurs pour qu'ils donnent les instructions nécessaires de modification à notre correspondant.";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Veuillez agréer, Messieurs, nos salutations distinguées."


XPrt.CurrentX = prtMinMarge + 5000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.FontBold = True
XPrt.Print paramSignature;
XPrt.FontBold = False

End Sub

'---------------------------------------------------------
Public Sub prtCDO_Courrier_OUV_PAR_Page1()
'---------------------------------------------------------
Dim X As String, y As String
Dim T_Cnf As Double, T_Not As Double
Dim Mnt_Not As Currency
Line2_Ecart = 5: prtCDO_Courrier_Trame (Line2_Ecart)
'Call frmElpPrt.prtTrame(prtMinMarge - 50, prtCorpsY - 100, prtMaxMarge, prtCorpsY + prtlineHeight * 5 + 100, "B", 245)

XPrt.CurrentY = prtCorpsY
XPrt.FontSize = 14: XPrt.FontBold = True: XPrt.FontUnderline = True
frmElpPrt.prtCentré prtMedX, "CREDIT DOCUMENTAIRE"

XPrt.FontSize = 11: XPrt.FontBold = False: XPrt.FontUnderline = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
'Call frmElpPrt.prtTrame(prtMinMarge - 50, XPrt.CurrentY - 50, prtMaxMarge, XPrt.CurrentY + prtlineHeight * 3 + 50, , 245)
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Donneur d'ordre";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1700: XPrt.Print meCDO_Courrier.DON_Concat;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Bénéficiaire";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1700: XPrt.Print meCDO_Courrier.BEN_Concat;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Montant";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentX = prtMinMarge + 4100: XPrt.Print "Devise :";
XPrt.CurrentX = prtMinMarge + 7500: XPrt.Print "Validité :";
XPrt.FontBold = True
X = Trim(Format$(meYCDODOS0.CDODOSMON, "### ### ### ##0.00"))
XPrt.CurrentX = prtMinMarge + 1700 '4000 - XPrt.TextWidth(X)
XPrt.Print X;
XPrt.CurrentX = prtMinMarge + 5000: XPrt.Print meYCDODOS0.CDODOSDEV;
XPrt.CurrentX = prtMinMarge + 8500: XPrt.Print wDate_Validité;
XPrt.FontBold = False

XPrt.FontSize = 11
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.CurrentX = prtMinMarge
XPrt.Print "Messieurs,";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Nous vous informons de l'ouverture du CREDIT DOCUMENTAIRE irrévocable N° ";
XPrt.FontBold = True
XPrt.Print meYCDODOS0.CDODOSEXT;
XPrt.FontBold = False

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "émis en votre faveur par notre correspondant :";
XPrt.CurrentX = XPrt.CurrentX + 200
prtAdresse meCDO_Courrier.BQE_ZADRESS0, False

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.Print "Valable à nos guichets jusqu'au :   ";
XPrt.FontBold = True
XPrt.Print wDate_Validité;
XPrt.FontBold = False

XPrt.Print " pour présentation des documents pour paiement dans les";
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "conditions reprises sur ";
XPrt.FontBold = True
XPrt.Print wAnnexe_Nb;
XPrt.FontBold = False
XPrt.Print " faisant partie intégrante de ce crédit.";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.Print "Nous vous prions de noter que ce crédit documentaire comporte NOTRE CONFIRMATION à hauteur de ";
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
T_Cnf = meYCDODOS0.CDODOSMOC * 100 / meYCDODOS0.CDODOSMOT
T_Not = 100 - T_Cnf
Mnt_Not = meYCDODOS0.CDODOSMOT - meYCDODOS0.CDODOSMOC
X = Trim(Format$(meYCDODOS0.CDODOSMOC, "### ### ### ##0.00"))
y = Trim(Format$(Mnt_Not, "### ### ### ##0.00"))
XPrt.Print Round(T_Cnf, 2) & " % soit " & meYCDODOS0.CDODOSDEV & " " & X & ". Le solde de " & Round(T_Not, 2) & " % soit " & meYCDODOS0.CDODOSDEV & " " & y & " interviendra conformément aux ";
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "conditions du crédit sans engagement ni responsabilité de la part de notre établissement. ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.Print "En cas d'irrégularités constatées lors de l'utilisation, notre confirmation deviendra nulle et sans effet au prorata ";
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "de l'utilisation et le règlement ne s'effectuera qu'après réception par nos soins :";
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "    1) de l'accord de la banque émettrice si cet accord est reçu dans la validité du crédit ";
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "    2) de l'accord de la banque émettrice et de la réception de la couverture correspondante si cet accord est reçu ";
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "        après la péremption du crédit. ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Lors de l'utilisation,  nous vous remercions de bien vouloir accompagner ";
XPrt.FontBold = True: XPrt.ForeColor = vbBlue
XPrt.Print "IMPERATIVEMENT";
XPrt.FontBold = False: XPrt.ForeColor = prtForeColor
XPrt.Print " les documents";
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "d'un exemplaire supplémentaire de votre facture et d'un ";
XPrt.FontBold = True: XPrt.ForeColor = vbBlue
XPrt.Print "relevé d'identité bancaire IBAN";
XPrt.FontBold = False: XPrt.ForeColor = prtForeColor
XPrt.Print " pour nous permettre";
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "de vous en effectuer le règlement et de nous communiquer votre ";
XPrt.FontBold = True: XPrt.ForeColor = vbBlue
XPrt.Print "n° de TVA intracommunautaire.";
XPrt.FontBold = False: XPrt.ForeColor = prtForeColor

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.Print "Ce crédit documentaire est soumis aux Règles et Usances Uniformes relatives aux crédits documentaires";
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "(version révisée de 2007 - Publication N° 600 de la Chambre de Commerce Internationale).";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.Print "Si vous n'êtes pas d'accord sur les conditions de ce crédit, nous vous conseillons de vous mettre DIRECTEMENT";
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "en rapport avec vos acheteurs pour qu'ils donnent les instructions nécessaires de modification à notre correspondant.";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.Print "Veuillez agréer, Messieurs, nos salutations distinguées."


XPrt.CurrentX = prtMinMarge + 5000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.FontBold = True
XPrt.Print paramSignature;
XPrt.FontBold = False

End Sub

'---------------------------------------------------------
Public Sub prtCDO_Courrier_OUV_C_OP_06_PDifNonRecl()
'---------------------------------------------------------
Dim X As String
Dim Tx As Double

XPrt.CurrentY = prtCorpsY
XPrt.FontSize = 14: XPrt.FontBold = True
frmElpPrt.prtCentré prtMedX, "ANNEXE FAISANT PARTIE INTEGRANTE"
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
frmElpPrt.prtCentré prtMedX, "DU CREDIT DOCUMENTAIRE"
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
frmElpPrt.prtCentré prtMedX, "N°  " & Trim(meYCDODOS0.CDODOSEXT)

XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 8
XPrt.CurrentX = prtMinMarge
XPrt.Print "Les frais et commissions étant stipulés à votre charge nous vous réclamerons lors de l'utilisation";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
Tx = mecnfYBIACDOCOM0.CDOCO2TX1 ' 2003.11.20  / 4       ' Taux de confirmation dans SAB : en annuel
Select Case mecnfYBIACDOCOM0.CDOCO2PER
    Case "M":   XPrt.Print " ou à la péremption du crédit notre commission de confirmation calculée au taux de : " & Format$(Tx, "### ##0.00") & " pour cent";
    Case "T":   XPrt.Print " ou à la péremption du crédit notre commission de confirmation calculée au taux de : ";
                XPrt.CurrentX = prtMinMarge + 3000
                XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
                XPrt.Print Format$(Tx, "### ##0.00") & " pour cent par trimestre indivisible";
    Case Else:  XPrt.Print " ou à la péremption du crédit notre commission de confirmation calculée au taux de : " & Format$(Tx, "### ##0.00") & " pour cent";
End Select

XPrt.CurrentX = prtMinMarge + 6000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.FontBold = True
XPrt.Print " Devise               Montant ";
XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.FontItalic = True
XPrt.CurrentX = prtMinMarge
XPrt.Print " Commission de confirmation ";
Select Case mecnfYBIACDOCOM0.CDOCO2PER
    Case "M":   XPrt.CurrentX = prtMinMarge + 2700: XPrt.Print "pour le premier mois";
    Case "T":   XPrt.CurrentX = prtMinMarge + 2700: XPrt.Print "pour le premier trimestre";
    Case Else:  XPrt.CurrentX = prtMinMarge + 2700: XPrt.Print "";
End Select
XPrt.CurrentX = prtMinMarge + 6200: XPrt.Print mecnfYBIACDOCOM0.CDOCOMDEV & "            " & Format$(mecnfYBIACDOCOM0.CDOCOMMON, "### ### ### ##0.00");
XPrt.CurrentX = prtMinMarge
XPrt.FontItalic = False

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print " Par ailleurs, lors de l'utilisation nous vous décompterons : ";

XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "- notre commission de levée de documents au taux de : ";

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinMarge + 2000
XPrt.Print "1,5 pour mille sur le montant utilisé minimum EUR 106,71 plus TVA ";

XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
If meYCDODOS0.CDODOSMDI <> 0 Then
    XPrt.Print "- notre commission de paiement différé au taux de 0,10 % par mois ";
Else
    XPrt.Print "- notre commission d'acceptation au taux de 0,10 % par mois ";
End If

XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "- frais de port et télex... ";

XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "- éventuellement une commission pour documents irréguliers de EUR 83,85 plus T.V.A. ";

XPrt.CurrentX = prtMinMarge + 5000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 4
XPrt.FontBold = True
XPrt.Print paramSignature;
XPrt.FontBold = False

End Sub

'---------------------------------------------------------
Public Sub prtCDO_Courrier_OUV_C_OP_07_PDifRecl()
'---------------------------------------------------------
Dim X As String
Dim Tx As Double

XPrt.CurrentY = prtCorpsY
XPrt.FontSize = 14: XPrt.FontBold = True
frmElpPrt.prtCentré prtMedX, "ANNEXE FAISANT PARTIE INTEGRANTE"
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
frmElpPrt.prtCentré prtMedX, "DU CREDIT DOCUMENTAIRE"
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
frmElpPrt.prtCentré prtMedX, "N°  " & Trim(meYCDODOS0.CDODOSEXT)

XPrt.FontSize = 11: XPrt.FontBold = False

If meYCDODOS0.CDODOSNOT = "0011077" Then     'Paragraphe spécifique pour BDL - 15.06.2004

    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 7
    XPrt.CurrentX = prtMinMarge
    XPrt.Print " Compte tenu des instructions particulières de la banque émettrice, notre confirmation ne ";
    
    XPrt.CurrentX = prtMinMarge
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.Print " deviendra effective qu'après réception par nous-mêmes de votre accord sur les conditions de ce ";
    
    XPrt.CurrentX = prtMinMarge
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.Print " crédit documentaire et du règlement de notre commission de confirmation calculée au taux de ";
    
    XPrt.CurrentX = prtMinMarge
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    Tx = mecnfYBIACDOCOM0.CDOCO2TX1 ' 2003.11.20  / 4       ' Taux de confirmation dans SAB : en annuel
    Select Case mecnfYBIACDOCOM0.CDOCO2PER
        Case "M":   XPrt.Print Format$(Tx, "### ##0.00") & " pour cent.";
        Case "T":   XPrt.Print Format$(Tx, "### ##0.00") & " pour cent par trimestre indivisible.";
        Case Else:  XPrt.Print Format$(Tx, "### ##0.00") & " pour cent.";
    End Select

Else

    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 8
    XPrt.CurrentX = prtMinMarge
    XPrt.Print " Compte tenu des instructions particulières de la banque émettrice au regard du paiement de nos ";
    
    XPrt.CurrentX = prtMinMarge
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.Print " commissions, veuillez nous reconnaître, à votre meilleure convenance de notre commission de ";
    
    XPrt.CurrentX = prtMinMarge
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    Tx = mecnfYBIACDOCOM0.CDOCO2TX1 ' 2003.11.20  / 4       ' Taux de confirmation dans SAB : en annuel
    Select Case mecnfYBIACDOCOM0.CDOCO2PER
        Case "M":   XPrt.Print " confirmation calculée au taux de " & Format$(Tx, "### ##0.00") & " pour cent.";
        Case "T":   XPrt.Print " confirmation calculée au taux de " & Format$(Tx, "### ##0.00") & " pour cent par trimestre indivisible.";
        Case Else:  XPrt.Print " confirmation calculée au taux de " & Format$(Tx, "### ##0.00") & " pour cent.";
    End Select
  
End If

XPrt.CurrentX = prtMinMarge + 6000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.FontItalic = True
XPrt.FontBold = True
XPrt.Print " Devise               Montant ";
XPrt.FontBold = False
XPrt.FontItalic = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.FontItalic = True
XPrt.CurrentX = prtMinMarge
XPrt.Print " Commission de confirmation ";
Select Case mecnfYBIACDOCOM0.CDOCO2PER
    Case "M":   XPrt.CurrentX = prtMinMarge + 2700: XPrt.Print "pour le premier mois";
    Case "T":   XPrt.CurrentX = prtMinMarge + 2700: XPrt.Print "pour le premier trimestre";
    Case Else:  XPrt.CurrentX = prtMinMarge + 2700: XPrt.Print "";
End Select
XPrt.CurrentX = prtMinMarge + 6200: XPrt.Print mecnfYBIACDOCOM0.CDOCOMDEV & "            " & Format$(mecnfYBIACDOCOM0.CDOCOMMON, "### ### ### ##0.00");
XPrt.CurrentX = prtMinMarge

XPrt.FontItalic = False

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print " Notre confirmation ne deviendra effective qu'à réception de votre règlement. En cas de désaccord, ";
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print " veuillez nous en aviser dès que possible par retour de courrier. ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print " Par ailleurs, lors de l'utilisation nous vous décompterons : ";

XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "- notre commission de levée de documents au taux de : ";

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinMarge + 2000
XPrt.Print "1,5 pour mille sur le montant utilisé minimum EUR 106,71 plus TVA ";

XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
If meYCDODOS0.CDODOSMDI <> 0 Then
    XPrt.Print "- notre commission de paiement différé au taux de 0,10 % par mois ";
Else
    XPrt.Print "- notre commission d'acceptation au taux de 0,10 % par mois ";
End If

XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "- frais de port et télex... ";

XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "- éventuellement une commission pour documents irréguliers de EUR 83,85 plus T.V.A. ";

XPrt.CurrentX = prtMinMarge + 5000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 4
XPrt.FontBold = True
XPrt.Print paramSignature;
XPrt.FontBold = False

End Sub

'---------------------------------------------------------
Public Sub prtCDO_Courrier_OUV_08_NOP_Forfait()
'---------------------------------------------------------
Dim X As String

XPrt.CurrentY = prtCorpsY
XPrt.FontSize = 14: XPrt.FontBold = True
frmElpPrt.prtCentré prtMedX, "ANNEXE FAISANT PARTIE INTEGRANTE"
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
frmElpPrt.prtCentré prtMedX, "DU CREDIT DOCUMENTAIRE"
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
frmElpPrt.prtCentré prtMedX, "N°  " & Trim(meYCDODOS0.CDODOSEXT)

XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 8
XPrt.CurrentX = prtMinMarge
XPrt.Print " GARANTIE DE BONNE EXECUTION :";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print " Conformément aux instructions de la banque émettrice, ce crédit ne deviendra opérationnel qu'après ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print " réception d'une caution bancaire de bonne exécution fixée à 10 pour cent de la valeur globale du ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print " contrat et son acceptation par l'ordonnateur. ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print " L'instruction rendant le crédit opérationnel, vous sera notifiée par nos soins."; ";"

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print " Par conséquent, veuillez considérer ce crédit non encore opérationnel. ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print " Les frais et commissions étant stipulés à votre charge nous vous réclamerons lors de l'utilisation ou ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print " à la péremption du crédit notre commission forfaitaire globale de : ";

XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
If meYCDODOS0.CDODOSCON = "C" Then         'Confirmé
    XPrt.Print mecnfYBIACDOCOM0.CDOTC2DEV & " " & mecnfYBIACDOCOM0.CDOTC2MTF & " compte tenu du faible montant de cette opération. ";
Else
    XPrt.Print menotYBIACDOCOM0.CDOTC2DEV & " " & menotYBIACDOCOM0.CDOTC2MTF & " compte tenu du faible montant de cette opération. ";
End If
XPrt.CurrentX = prtMinMarge + 5000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 5
XPrt.FontBold = True
XPrt.Print paramSignature;
XPrt.FontBold = False

End Sub


'---------------------------------------------------------
Public Sub prtCDO_Courrier_OUV_C_OP_05_AVueRecl()
'---------------------------------------------------------
Dim X As String
Dim Tx As Double

XPrt.CurrentY = prtCorpsY
XPrt.FontSize = 14: XPrt.FontBold = True
frmElpPrt.prtCentré prtMedX, "ANNEXE FAISANT PARTIE INTEGRANTE"
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
frmElpPrt.prtCentré prtMedX, "DU CREDIT DOCUMENTAIRE"
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
frmElpPrt.prtCentré prtMedX, "N°  " & Trim(meYCDODOS0.CDODOSEXT)

XPrt.FontSize = 11: XPrt.FontBold = False

If meYCDODOS0.CDODOSNOT = "0011077" Then     'Paragraphe spécifique pour BDL - 15.06.2004

    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 7
    XPrt.CurrentX = prtMinMarge
    XPrt.Print " Compte tenu des instructions particulières de la banque émettrice, notre confirmation ne ";
    
    XPrt.CurrentX = prtMinMarge
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.Print " deviendra effective qu'après réception par nous-mêmes de votre accord sur les conditions de ce ";
    
    XPrt.CurrentX = prtMinMarge
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.Print " crédit documentaire et du règlement de notre commission de confirmation calculée au taux de ";
    
    XPrt.CurrentX = prtMinMarge
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    Tx = mecnfYBIACDOCOM0.CDOCO2TX1 ' 2003.11.20  / 4       ' Taux de confirmation dans SAB : en annuel
    Select Case mecnfYBIACDOCOM0.CDOCO2PER
        Case "M":   XPrt.Print Format$(Tx, "### ##0.00") & " pour cent.";
        Case "T":   XPrt.Print Format$(Tx, "### ##0.00") & " pour cent par trimestre indivisible.";
        Case Else:  XPrt.Print Format$(Tx, "### ##0.00") & " pour cent.";
    End Select

Else

    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 8
    XPrt.CurrentX = prtMinMarge
    XPrt.Print " Compte tenu des instructions particulières de la banque émettrice au regard du paiement de nos ";
    
    XPrt.CurrentX = prtMinMarge
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.Print " commissions, veuillez nous reconnaître, à votre meilleure convenance de notre commission de ";

    XPrt.CurrentX = prtMinMarge
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    Tx = mecnfYBIACDOCOM0.CDOCO2TX1 ' 2003.11.20  / 4       ' Taux de confirmation dans SAB : en annuel
    Select Case mecnfYBIACDOCOM0.CDOCO2PER
        Case "M":   XPrt.Print " confirmation calculée au taux de " & Format$(Tx, "### ##0.00") & " pour cent.";
        Case "T":   XPrt.Print " confirmation calculée au taux de " & Format$(Tx, "### ##0.00") & " pour cent par trimestre indivisible.";
        Case Else:  XPrt.Print " confirmation calculée au taux de " & Format$(Tx, "### ##0.00") & " pour cent.";
    End Select

End If

XPrt.CurrentX = prtMinMarge + 6000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.FontItalic = True
XPrt.FontBold = True
XPrt.Print " Devise               Montant ";
XPrt.FontBold = False
XPrt.FontItalic = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.FontItalic = True
XPrt.CurrentX = prtMinMarge
XPrt.Print " Commission de confirmation ";
Select Case mecnfYBIACDOCOM0.CDOCO2PER
    Case "M":   XPrt.CurrentX = prtMinMarge + 2700: XPrt.Print "pour le premier mois";
    Case "T":   XPrt.CurrentX = prtMinMarge + 2700: XPrt.Print "pour le premier trimestre";
    Case Else:  XPrt.CurrentX = prtMinMarge + 2700: XPrt.Print "";
End Select

' XPrt.CurrentX = prtMinMarge + 6200: XPrt.Print mecnfYBIACDOCOM0.CDOCOMDEV & "    -8C                " & mecnfYBIACDOCOM0.CDOCOMMON;
XPrt.CurrentX = prtMinMarge + 6200: XPrt.Print mecnfYBIACDOCOM0.CDOCOMDEV & "            " & Format$(mecnfYBIACDOCOM0.CDOCOMMON, "### ### ### ##0.00");
XPrt.CurrentX = prtMinMarge

XPrt.FontItalic = False

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print " Notre confirmation ne deviendra effective qu'à réception de votre règlement. En cas de désaccord, ";
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print " veuillez nous en aviser dès que possible par retour de courrier. ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print " Par ailleurs, lors de l'utilisation nous vous décompterons : ";

XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "- notre commission de levée de documents au taux de : ";

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinMarge + 2000
XPrt.Print "1,5 pour mille sur le montant utilisé minimum EUR 106,71 plus TVA ";

XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "- frais de port et télex... ";

XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "- éventuellement une commission pour documents irréguliers de EUR 83,85 plus T.V.A. ";

XPrt.CurrentX = prtMinMarge + 5000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 4
XPrt.FontBold = True
XPrt.Print paramSignature;
XPrt.FontBold = False

End Sub

'---------------------------------------------------------
Public Sub prtCDO_Courrier_OUV_Annexe_03()
'---------------------------------------------------------
Dim X As String

XPrt.CurrentY = prtCorpsY
XPrt.FontSize = 14: XPrt.FontBold = True
frmElpPrt.prtCentré prtMedX, "ANNEXE FAISANT PARTIE INTEGRANTE"
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
frmElpPrt.prtCentré prtMedX, "DU CREDIT DOCUMENTAIRE"
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
frmElpPrt.prtCentré prtMedX, "N°  " & Trim(meYCDODOS0.CDODOSEXT)
XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 8
XPrt.CurrentX = prtMinMarge
XPrt.Print "Les frais et commissions étant stipulés à votre charge nous vous réclamerons lors de l'utilisation";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.Print "ou à la péremption du crédit notre commission forfaitaire globale de :     EUR 250,00 ";
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.Print "compte tenu du faible montant de cette opération. ";

XPrt.CurrentX = prtMinMarge + 5000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 6
XPrt.FontBold = True
XPrt.Print paramSignature;
XPrt.FontBold = False

End Sub

'---------------------------------------------------------
Public Sub prtCDO_Courrier_OUV_01_OP_Forfait()
'---------------------------------------------------------
Dim X As String

XPrt.CurrentY = prtCorpsY
XPrt.FontSize = 14: XPrt.FontBold = True
frmElpPrt.prtCentré prtMedX, "ANNEXE FAISANT PARTIE INTEGRANTE"
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
frmElpPrt.prtCentré prtMedX, "DU CREDIT DOCUMENTAIRE"
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
frmElpPrt.prtCentré prtMedX, "N°  " & Trim(meYCDODOS0.CDODOSEXT)

XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 8
XPrt.CurrentX = prtMinMarge
XPrt.Print "Les frais et commissions étant stipulés à votre charge nous vous réclamerons lors de l'utilisation";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
If meYCDODOS0.CDODOSCON = "C" Then         'Confirmé
    XPrt.Print "ou à la péremption du crédit notre commission forfaitaire globale de : " & mecnfYBIACDOCOM0.CDOTC2DEV & " " & mecnfYBIACDOCOM0.CDOTC2MTF;
Else
    XPrt.Print "ou à la péremption du crédit notre commission forfaitaire globale de : " & menotYBIACDOCOM0.CDOTC2DEV & " " & menotYBIACDOCOM0.CDOTC2MTF;
End If
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "compte tenu du faible montant de cette opération. ";

XPrt.CurrentX = prtMinMarge + 5000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 5
XPrt.FontBold = True
XPrt.Print paramSignature;
XPrt.FontBold = False

End Sub

'---------------------------------------------------------
Public Sub prtCDO_Courrier_OUV_Annexe_02()
'---------------------------------------------------------
Dim X As String

XPrt.CurrentY = prtCorpsY
XPrt.FontSize = 14: XPrt.FontBold = True
frmElpPrt.prtCentré prtMedX, "ANNEXE FAISANT PARTIE INTEGRANTE"
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
frmElpPrt.prtCentré prtMedX, "DU CREDIT DOCUMENTAIRE"
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
frmElpPrt.prtCentré prtMedX, "N°  " & Trim(meYCDODOS0.CDODOSEXT)

XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 8
XPrt.CurrentX = prtMinMarge
XPrt.Print "Les frais et commissions étant stipulés à votre charge nous vous réclamerons lors de l'utilisation";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.Print "ou à la péremption du crédit notre commission forfaitaire globale de :     EUR 150,00 ";
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.Print "compte tenu du faible montant de cette opération. ";

XPrt.CurrentX = prtMinMarge + 5000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 6
XPrt.FontBold = True
XPrt.Print paramSignature;
XPrt.FontBold = False

End Sub

Public Sub prtCDO_Courrier_Monitor(lstOptions As ListBox, lYCDODOS0 As typeZCDODOS0, lYCDOMOD0 As typeZCDOMOD0, lcnfYBIACDOCOM0 As typeYBIACDOCOM0, lnotYBIACDOCOM0 As typeYBIACDOCOM0, lCDO_Courrier As typeCDO_Courrier)
Dim I As Integer
Dim X As String


meCDO_Courrier = lCDO_Courrier
meYCDODOS0 = lYCDODOS0

meYCDOMOD0 = lYCDOMOD0
X = ""

mecnfYBIACDOCOM0 = lcnfYBIACDOCOM0
menotYBIACDOCOM0 = lnotYBIACDOCOM0

W_Validité_AMJ = Format$((meYCDODOS0.CDODOSVAL + 19000000), "00000000")
wDate_Validité = dateIBM10(meYCDODOS0.CDODOSVAL, True)
wDate_Validité_GB = dateImp_ddMonthYYYY(meYCDODOS0.CDODOSVAL + 19000000)
wDate_Limite_Emb = dateIBM10(meYCDODOS0.CDODOSDLE, True)
wDate_Anc_Validité = dateIBM10(meYCDOMOD0.CDOMODVAL, True)
wDate_Valeur_CR = dateIBM10(meCDO_Courrier.YCDOREG0_C.CDOREGDVA, True)
wDate_Echeance_CR = dateIBM10(meCDO_Courrier.YCDOREG0_C.CDOREGDAE, True)
wDate_Valeur_DB = dateIBM10(meCDO_Courrier.YCDOREG0_D.CDOREGDVA, True)
wDate_Valeur_DB_GB = dateImp_ddMonthYYYY(meCDO_Courrier.YCDOREG0_D.CDOREGDVA + 19000000)
wDate_Echeance_DB = dateIBM10(meCDO_Courrier.YCDOREG0_D.CDOREGDAE, True)
wDate_Echeance_DB_GB = dateImp_ddMonthYYYY(meCDO_Courrier.YCDOREG0_D.CDOREGDAE + 19000000)
wDate_Remise_Util = dateIBM10(meCDO_Courrier.YCDOUTI0.CDOUTIDRE, True)
wCompte_CR = meCDO_Courrier.YCDOREG0_C.CDOREGCOM

wAnnexe_Nb = prtCDO_Courrier_Annexe_Nb(meCDO_Courrier.Annexe_Nb)
wContact = meCDO_Courrier.Contact

For I = 0 To lstOptions.ListCount - 1
    lstOptions.ListIndex = I
    If lstOptions.Selected(I) Then
        xDocRéférence = (lstOptions.Text)  ' Nom du document à imprimer
        Booleen_GB = False     ' Booléen pour courrier en Anglais
        Booleen_AR = False     ' Booléen pour AR en entête : "A l'attention de ..."
        Select Case xDocRéférence
        
            ' UTILISATIONS :
            Case "UTI_CONFORME_BED1_Pdif": prtCDO_Courrier_form "D": prtCDO_Courrier_UTI_CONFORME_BED1_Pdif
            Case "UTI_CONFORME_BED1_Pdif_GB": Booleen_GB = True
                                              prtCDO_Courrier_form "D"
                                              prtCDO_Courrier_UTI_CONFORME_BED1_Pdif_GB
                                              Booleen_GB = False
            Case "UTI_CONFORME_BED2_Avue": prtCDO_Courrier_form ("D"): prtCDO_Courrier_UTI_CONFORME_BED2_Avue
            Case "UTI_CONFORME_BED2_Avue_GB": Booleen_GB = True
                                              prtCDO_Courrier_form "D"
                                              prtCDO_Courrier_UTI_CONFORME_BED2_Avue_GB
                                              Booleen_GB = False
            Case "UTI_CONFORME_C_Avue_AR1": Booleen_AR = True
                                            prtCDO_Courrier_form ("B")
                                            prtCDO_Courrier_UTI_CONFORME_C_Avue_AR1
                                            Booleen_AR = False
            Case "UTI_CONFORME_C_Pdif_AR2": Booleen_AR = True
                                            prtCDO_Courrier_form ("B")
                                            prtCDO_Courrier_UTI_CONFORME_C_Pdif_AR2
                                            Booleen_AR = False
            Case "UTI_CONFORME_N_Avue_AR5": Booleen_AR = True
                                            prtCDO_Courrier_form ("B")
                                            prtCDO_Courrier_UTI_CONFORME_N_Avue_AR5
                                            Booleen_AR = False
            Case "UTI_CONFORME_N_Pdif_AR6": Booleen_AR = True
                                            prtCDO_Courrier_form ("B")
                                            prtCDO_Courrier_UTI_CONFORME_N_Pdif_AR6
                                            Booleen_AR = False
            Case "UTI_NCONFORME_AR": Booleen_AR = True
                                     prtCDO_Courrier_form ("B")
                                     prtCDO_Courrier_UTI_NCONFORME_AR
                                     Booleen_AR = False
            Case "UTI_NCONFORME_BED_Avue": prtCDO_Courrier_form ("D"): prtCDO_Courrier_UTI_NCONFORME_BED_Avue
            Case "UTI_NCONFORME_BED_Pdif": prtCDO_Courrier_form ("D"): prtCDO_Courrier_UTI_NCONFORME_BED_Pdif
            Case "UTI_ACCORDRECU_C_Avue_AR10": Booleen_AR = True
                                               prtCDO_Courrier_form ("B")
                                               prtCDO_Courrier_UTI_ACCORDRECU_C_Avue_AR10
                                               Booleen_AR = False
            Case "UTI_ACCORDRECU_C_Pdif_AR11": Booleen_AR = True
                                               prtCDO_Courrier_form ("B")
                                               prtCDO_Courrier_UTI_ACCORDRECU_C_Pdif_AR11
                                               Booleen_AR = False
            Case "UTI_ACCORDRECU_N_Avue_AR12": Booleen_AR = True
                                               prtCDO_Courrier_form ("B")
                                               prtCDO_Courrier_UTI_ACCORDRECU_N_Avue_AR12
                                               Booleen_AR = False
            Case "UTI_ACCORDRECU_N_Pdif_AR13": Booleen_AR = True
                                               prtCDO_Courrier_form ("B")
                                               prtCDO_Courrier_UTI_ACCORDRECU_N_Pdif_AR13
                                               Booleen_AR = False
            
            ' OUVERTURE : Les annexes fofaitaires
            Case "OUV_01_OP_Forfait": prtCDO_Courrier_form ("B"): prtCDO_Courrier_OUV_01_OP_Forfait
            Case "OUV_08_NOP_Forfait": prtCDO_Courrier_form ("B"): prtCDO_Courrier_OUV_08_NOP_Forfait
            
            ' OUVERTURE : Lettre et annexes CONFIRMES
            Case "OUV_CNF_Page1": prtCDO_Courrier_form ("B"): prtCDO_Courrier_OUV_CNF_Page1
            Case "OUV_C_NOP_11_AVueNonRecl": prtCDO_Courrier_form ("B"): prtCDO_Courrier_OUV_C_NOP_11_AVueNonRecl
            Case "OUV_C_NOP_12_AVueRecl": prtCDO_Courrier_form ("B"): prtCDO_Courrier_OUV_C_NOP_12_AVueRecl
            Case "OUV_C_NOP_13_PDifNonRecl": prtCDO_Courrier_form ("B"): prtCDO_Courrier_OUV_C_NOP_13_PDifNonRecl
            Case "OUV_C_NOP_14_PDifRecl": prtCDO_Courrier_form ("B"): prtCDO_Courrier_OUV_C_NOP_14_PDifRecl
            Case "OUV_C_OP_04_AVueNonRecl": prtCDO_Courrier_form ("B"): prtCDO_Courrier_OUV_C_OP_04_AVueNonRecl
            Case "OUV_C_OP_05_AVueRecl": prtCDO_Courrier_form ("B"): prtCDO_Courrier_OUV_C_OP_05_AVueRecl
            Case "OUV_C_OP_06_PDifNonRecl": prtCDO_Courrier_form ("B"): prtCDO_Courrier_OUV_C_OP_06_PDifNonRecl
            Case "OUV_C_OP_07_PDifRecl": prtCDO_Courrier_form ("B"): prtCDO_Courrier_OUV_C_OP_07_PDifRecl

            ' OUVERTURE : Lettre et annexes NOTIFIES
            Case "OUV_NOT_Page1": prtCDO_Courrier_form ("B"): prtCDO_Courrier_OUV_NOT_Page1
            Case "OUV_N_NOP_25_NonRecl": prtCDO_Courrier_form ("B"): prtCDO_Courrier_OUV_N_NOP_25_NonRecl
            Case "OUV_N_NOP_26_Recl": prtCDO_Courrier_form ("B"): prtCDO_Courrier_OUV_N_NOP_26_Recl
            Case "OUV_N_OP_20_NonRecl": prtCDO_Courrier_form ("B"): prtCDO_Courrier_OUV_N_OP_20_NonRecl
            Case "OUV_N_OP_21_Recl": prtCDO_Courrier_form ("B"): prtCDO_Courrier_OUV_N_OP_21_Recl
            
            ' OUVERTURE : Lettre et annexes PARTIELLEMENT CONFIRMES
            Case "OUV_PAR_Page1": prtCDO_Courrier_form ("B"): prtCDO_Courrier_OUV_PAR_Page1
            Case "OUV_P_OP_32_AVueNonRecl": prtCDO_Courrier_form ("B"): prtCDO_Courrier_OUV_P_OP_32_AVueNonRecl
            Case "OUV_P_OP_33_AVueRecl": prtCDO_Courrier_form ("B"): prtCDO_Courrier_OUV_P_OP_33_AVueRecl
            Case "OUV_P_OP_34_PDifNonRecl": prtCDO_Courrier_form ("B"): prtCDO_Courrier_OUV_P_OP_34_PDifNonRecl
            Case "OUV_P_OP_35_PDifRecl": prtCDO_Courrier_form ("B"): prtCDO_Courrier_OUV_P_OP_35_PDifRecl
            Case "OUV_P_NOP_39_AVueNonRecl": prtCDO_Courrier_form ("B"): prtCDO_Courrier_OUV_P_NOP_39_AVueNonRecl
            Case "OUV_P_NOP_40_AVueRecl": prtCDO_Courrier_form ("B"): prtCDO_Courrier_OUV_P_NOP_40_AVueRecl
            Case "OUV_P_NOP_41_PDifNonRecl": prtCDO_Courrier_form ("B"): prtCDO_Courrier_OUV_P_NOP_41_PDifNonRecl
            Case "OUV_P_NOP_42_PDifRecl": prtCDO_Courrier_form ("B"): prtCDO_Courrier_OUV_P_NOP_42_PDifRecl
           
            ' OUVERTURE : Annexes FIGES - A SUPPRIMER
            Case "OUV_Annexe_02": prtCDO_Courrier_form ("B"): prtCDO_Courrier_OUV_Annexe_02
            Case "OUV_Annexe_03": prtCDO_Courrier_form ("B"): prtCDO_Courrier_OUV_Annexe_03
            Case "OUV_Annexe_09": prtCDO_Courrier_form ("B"): prtCDO_Courrier_OUV_Annexe_09
            Case "OUV_Annexe_10": prtCDO_Courrier_form ("B"): prtCDO_Courrier_OUV_Annexe_10

            ' MODIFICATION : Diminution de montant de dossier FCB (Frais charge Benef) ou FCO (Frais charge D.O.)
            Case "MOD_19_Diminution": prtCDO_Courrier_form ("B"): prtCDO_Courrier_MOD_19_Diminution
            
            ' MODIFICATION : Lettres FCB (Frais charge Benef)
            Case "MOD_FCB_01_ValProrogée": prtCDO_Courrier_form ("B"): prtCDO_Courrier_MOD_FCB_01_ValProrogée
            Case "MOD_FCB_02_ValProrogée_Emb": prtCDO_Courrier_form ("B"): prtCDO_Courrier_MOD_FCB_02_ValProrogée_Emb
            Case "MOD_FCB_03_Annexe": prtCDO_Courrier_form ("B"): prtCDO_Courrier_MOD_FCB_03_Annexe
            Case "MOD_FCB_09_ValRaccourcie": prtCDO_Courrier_form ("B"): prtCDO_Courrier_MOD_FCB_09_ValRaccourcie
            Case "MOD_FCB_10_ValRaccourcie_Emb": prtCDO_Courrier_form ("B"): prtCDO_Courrier_MOD_FCB_10_ValRaccourcie_Emb
            Case "MOD_FCB_11_CNF_En_NOT": prtCDO_Courrier_form ("B"): prtCDO_Courrier_MOD_FCB_11_CNF_En_NOT
            Case "MOD_FCB_13_NOT_En_CNF": prtCDO_Courrier_form ("B"): prtCDO_Courrier_MOD_FCB_13_NOT_En_CNF
            Case "MOD_FCB_17_Augmentation_CNF": prtCDO_Courrier_form ("B"): prtCDO_Courrier_MOD_FCB_17_Augmentation_CNF
            Case "MOD_FCB_20_Augmentation_NOT": prtCDO_Courrier_form ("B"): prtCDO_Courrier_MOD_FCB_20_Augmentation_NOT
            
            ' MODIFICATION : Lettres FCO (Frais charge donneur d'ordre)
            Case "MOD_FCO_04_ValProrogée": prtCDO_Courrier_form ("B"): prtCDO_Courrier_MOD_FCO_04_ValProrogée
            Case "MOD_FCO_05_ValProrogée_Emb": prtCDO_Courrier_form ("B"): prtCDO_Courrier_MOD_FCO_05_ValProrogée_Emb
            Case "MOD_FCO_06_Annexe": prtCDO_Courrier_form ("B"): prtCDO_Courrier_MOD_FCO_06_Annexe
            Case "MOD_FCO_07_ValRaccourcie": prtCDO_Courrier_form ("B"): prtCDO_Courrier_MOD_FCO_07_ValRaccourcie
            Case "MOD_FCO_08_ValRaccourcie_Emb": prtCDO_Courrier_form ("B"): prtCDO_Courrier_MOD_FCO_08_ValRaccourcie_Emb
            Case "MOD_FCO_12_CNF_En_NOT": prtCDO_Courrier_form ("B"): prtCDO_Courrier_MOD_FCO_12_CNF_En_NOT
            Case "MOD_FCO_14_NOT_En_CNF": prtCDO_Courrier_form ("B"): prtCDO_Courrier_MOD_FCO_14_NOT_En_CNF
            Case "MOD_FCO_18_Augmentation": prtCDO_Courrier_form ("B"): prtCDO_Courrier_MOD_FCO_18_Augmentation
            
        End Select
    End If
Next I
If blnOpen Then prtCDO_Courrier_Close

End Sub

Public Sub prtCDO_Courrier_UTI_CONFORME_N_Avue_AR5()
'---------------------------------------------------------
Dim X As String
Dim W_Cumul_Comm As Currency
Dim W_Cumul_TVA As Currency
Dim W_MNT_NET As Currency
Dim C As Integer
Dim W_Code_CRD As String
Dim MNT As String, MTV As String

W_Cumul_Comm = 0
W_Cumul_TVA = 0
W_MNT_NET = 0

Line2_Ecart = 10: prtCDO_Courrier_Trame (Line2_Ecart)
'Call frmElpPrt.prtTrame(prtMinMarge - 50, prtCorpsY - 100, prtMaxMarge, prtCorpsY + prtlineHeight * 10 + 100, "B", 245)

XPrt.CurrentY = prtCorpsY
XPrt.FontSize = 14: XPrt.FontBold = True: XPrt.FontUnderline = True
frmElpPrt.prtCentré prtMedX, "CREDIT DOCUMENTAIRE"
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
frmElpPrt.prtCentré prtMedX, "ACCUSE DE RECEPTION"

XPrt.FontSize = 11: XPrt.FontBold = False: XPrt.FontUnderline = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
'Call frmElpPrt.prtTrame(prtMinMarge - 50, XPrt.CurrentY - 50, prtMaxMarge, XPrt.CurrentY + prtlineHeight * 3 + 50, , 245)

XPrt.CurrentX = prtMinMarge: XPrt.Print "  V/Référence :";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1700: XPrt.Print meCDO_Courrier.YCDOUTI0.CDOUTIRER;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Donneur d'ordre";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1700: XPrt.Print meCDO_Courrier.DON_Concat;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Bénéficiaire";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1700: XPrt.Print meCDO_Courrier.BEN_Concat;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Montant";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentX = prtMinMarge + 4100: XPrt.Print "Devise :";
XPrt.CurrentX = prtMinMarge + 7500: XPrt.Print "Validité :";
XPrt.FontBold = True
X = Trim(Format$(meYCDODOS0.CDODOSMON, "### ### ### ##0.00"))
XPrt.CurrentX = prtMinMarge + 1700
XPrt.Print X;
XPrt.CurrentX = prtMinMarge + 5000: XPrt.Print meYCDODOS0.CDODOSDEV;
XPrt.CurrentX = prtMinMarge + 8500: XPrt.Print wDate_Validité;
XPrt.FontBold = False

XPrt.FontSize = 11
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 4
XPrt.CurrentX = prtMinMarge
XPrt.Print "Messieurs,";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Nous accusons réception de votre lettre nous remettant les documents de ";
XPrt.Print meYCDODOS0.CDODOSDEV & "  " & Trim(Format$(meCDO_Courrier.YCDOUTI0.CDOUTIMPA, "### ### ### ##0.00"));
XPrt.Print "  en réalisation du ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "crédit documentaire précité. ";

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.CurrentX = prtMinMarge
XPrt.Print "Montant de l'utilisation ";
XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meYCDODOS0.CDODOSDEV & "            " & Format$(meCDO_Courrier.YCDOUTI0.CDOUTIMPA, "### ### ### ##0.00");

' >>>>> Lettre D'ACCUSE RECEPTION = Toujours destinataire BENEF donc Toujours comm. ligne CREDIT
W_Code_CRD = "C"


' La somme des commissions lues
For C = 1 To meCDO_Courrier.Com_Nb
    If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD Then
        W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    End If
Next C

If W_Cumul_Comm <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Nos frais et commissions selon détail ci-après déduits : ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print "Devise             Montant        TVA 19,60%";
    XPrt.CurrentX = prtMinMarge
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
End If

' >>>>> Remettre à zéro W_Cumul_Comm avant de passer dans la boucle lignes de commission
W_Cumul_Comm = 0
For C = 1 To meCDO_Courrier.Com_Nb

' Commission ECNF...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "ECNF  " And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission de confirmation ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission ECSIL...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "ECSIL " And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission de confirmation silencieuse ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission ENOTIF...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "ENOTIF" And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission de notification ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EPDIF...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EPDIF " And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission d'engagement paiement différé ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EPDIFD...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EPDIFD" And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission d'engagement paiement différé ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EACCEP...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EACCEP" And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission d'acceptation ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EACCED...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EACCED" And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission d'acceptation ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission ELVD...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "ELVD  " And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission de levée de documents ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission ERFA...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "ERFA  " And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Frais de ports et telex ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EDOCIR...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EDOCIR" And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission de documents irréguliers ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EACD...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EACD  " And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Autre commission documentaire ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EMODIF...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EMODIF" And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission de modification ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EFBEMT... Toujours en PLUS pour le bénéficiaire à l'accusé de réception
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EFBEMT" And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Frais banque émettrice ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' >>>>> Fin de boucle sur lignes de commissions
Next C

If W_Cumul_Comm <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
    XPrt.CurrentX = prtMinMarge
    XPrt.CurrentX = prtMinMarge + 3000: XPrt.Print "Montant  total ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(1);
    MNT = Format$(W_Cumul_Comm, "### ### ### ##0.00")
    MTV = Format$(W_Cumul_TVA, "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
End If

W_MNT_NET = meCDO_Courrier.YCDOUTI0.CDOUTIMPA - W_Cumul_Comm - W_Cumul_TVA
' If W_MNT_NET <> 0 Then
If W_Cumul_Comm <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
    XPrt.CurrentX = prtMinMarge
    Call frmElpPrt.prtTrame(prtMinMarge + 3000 - 50, XPrt.CurrentY - 100, prtMaxMarge - 2000, XPrt.CurrentY + prtlineHeight + 100, 245)
    XPrt.CurrentX = prtMinMarge + 3000: XPrt.Print "Montant  total  net ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(1);
    MNT = Format$(W_MNT_NET, "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
End If

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Ce crédit documentaire n'étant pas confirmé par notre banque, nous verserons ce montant en votre faveur ";
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "conformément à vos instructions,"
XPrt.CurrentX = prtMinMarge + 3000
XPrt.CurrentY = XPrt.CurrentY - prtlineHeight
XPrt.FontUnderline = True:  XPrt.Print "après réception de la couverture correspondante"
XPrt.CurrentX = prtMinMarge + 7250
XPrt.CurrentY = XPrt.CurrentY - prtlineHeight
XPrt.FontUnderline = False: XPrt.Print ", prévue pour le " & wDate_Valeur_CR & ",": XPrt.CurrentY = XPrt.CurrentY - prtlineHeight
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "selon les termes du crédit chez  " & meCDO_Courrier.IBAN;

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Veuillez agréer, Messieurs, nos salutations distinguées."

XPrt.CurrentX = prtMinMarge + 5000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.FontBold = True
XPrt.Print paramSignature;
XPrt.FontBold = False

End Sub

Public Sub prtCDO_Courrier_UTI_CONFORME_N_Pdif_AR6()
'---------------------------------------------------------
Dim X As String
Dim W_Cumul_Comm As Currency
Dim W_Cumul_TVA As Currency
Dim W_MNT_NET As Currency
Dim C As Integer
Dim W_Code_CRD As String
Dim MNT As String, MTV As String

W_Cumul_Comm = 0
W_Cumul_TVA = 0
W_MNT_NET = 0

Line2_Ecart = 10: prtCDO_Courrier_Trame (Line2_Ecart)
'Call frmElpPrt.prtTrame(prtMinMarge - 50, prtCorpsY - 100, prtMaxMarge, prtCorpsY + prtlineHeight * 10 + 100, "B", 245)

XPrt.CurrentY = prtCorpsY
XPrt.FontSize = 14: XPrt.FontBold = True: XPrt.FontUnderline = True
frmElpPrt.prtCentré prtMedX, "CREDIT DOCUMENTAIRE"
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
frmElpPrt.prtCentré prtMedX, "ACCUSE DE RECEPTION"

XPrt.FontSize = 11: XPrt.FontBold = False: XPrt.FontUnderline = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
'Call frmElpPrt.prtTrame(prtMinMarge - 50, XPrt.CurrentY - 50, prtMaxMarge, XPrt.CurrentY + prtlineHeight * 3 + 50, , 245)

XPrt.CurrentX = prtMinMarge: XPrt.Print "  V/Référence :";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1700: XPrt.Print meCDO_Courrier.YCDOUTI0.CDOUTIRER;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Donneur d'ordre";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1700: XPrt.Print meCDO_Courrier.DON_Concat;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Bénéficiaire";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1700: XPrt.Print meCDO_Courrier.BEN_Concat;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Montant";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentX = prtMinMarge + 4100: XPrt.Print "Devise :";
XPrt.CurrentX = prtMinMarge + 7500: XPrt.Print "Validité :";
XPrt.FontBold = True
X = Trim(Format$(meYCDODOS0.CDODOSMON, "### ### ### ##0.00"))
XPrt.CurrentX = prtMinMarge + 1700
XPrt.Print X;
XPrt.CurrentX = prtMinMarge + 5000: XPrt.Print meYCDODOS0.CDODOSDEV;
XPrt.CurrentX = prtMinMarge + 8500: XPrt.Print wDate_Validité;
XPrt.FontBold = False

XPrt.FontSize = 11
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 4
XPrt.CurrentX = prtMinMarge
XPrt.Print "Messieurs,";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Nous accusons réception de votre lettre nous remettant les documents de ";
XPrt.Print meYCDODOS0.CDODOSDEV & "  " & Trim(Format$(meCDO_Courrier.YCDOUTI0.CDOUTIMPA, "### ### ### ##0.00"));
XPrt.Print "  en réalisation du ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "crédit documentaire précité. ";

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.CurrentX = prtMinMarge
XPrt.Print "Montant de l'utilisation ";
XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meYCDODOS0.CDODOSDEV & "            " & Format$(meCDO_Courrier.YCDOUTI0.CDOUTIMPA, "### ### ### ##0.00");

' >>>>> Lettre D'ACCUSE RECEPTION = Toujours destinataire BENEF donc Toujours comm. ligne CREDIT
W_Code_CRD = "C"


' La somme des commissions lues
For C = 1 To meCDO_Courrier.Com_Nb
    If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD Then
        W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    End If
Next C

If W_Cumul_Comm <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Nos frais et commissions selon détail ci-après déduits : ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print "Devise             Montant        TVA 19,60%";
    XPrt.CurrentX = prtMinMarge
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
End If

' >>>>> Remettre à zéro W_Cumul_Comm avant de passer dans la boucle lignes de commission
W_Cumul_Comm = 0
For C = 1 To meCDO_Courrier.Com_Nb

' Commission ECNF...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "ECNF  " And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission de confirmation ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission ECSIL...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "ECSIL " And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission de confirmation silencieuse ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission ENOTIF...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "ENOTIF" And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission de notification ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EPDIF...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EPDIF " And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission d'engagement paiement différé ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EPDIFD...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EPDIFD" And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission d'engagement paiement différé ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EACCEP...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EACCEP" And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission d'acceptation ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EACCED...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EACCED" And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission d'acceptation ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission ELVD...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "ELVD  " And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission de levée de documents ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission ERFA...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "ERFA  " And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Frais de ports et telex ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EDOCIR...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EDOCIR" And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission de documents irréguliers ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EACD...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EACD  " And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Autre commission documentaire ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EMODIF...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EMODIF" And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission de modification ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EFBEMT... Toujours en PLUS pour le bénéficiaire à l'accusé de réception
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EFBEMT" And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Frais banque émettrice ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' >>>>> Fin de boucle sur lignes de commissions
Next C

If W_Cumul_Comm <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
    XPrt.CurrentX = prtMinMarge
    XPrt.CurrentX = prtMinMarge + 3000: XPrt.Print "Montant  total ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(1);
    MNT = Format$(W_Cumul_Comm, "### ### ### ##0.00")
    MTV = Format$(W_Cumul_TVA, "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
End If

W_MNT_NET = meCDO_Courrier.YCDOUTI0.CDOUTIMPA - W_Cumul_Comm - W_Cumul_TVA
' If W_MNT_NET <> 0 Then
If W_Cumul_Comm <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
    XPrt.CurrentX = prtMinMarge
    Call frmElpPrt.prtTrame(prtMinMarge + 3000 - 50, XPrt.CurrentY - 100, prtMaxMarge - 2000, XPrt.CurrentY + prtlineHeight + 100, 245)
    XPrt.CurrentX = prtMinMarge + 3000: XPrt.Print "Montant  total  net ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(1);
    MNT = Format$(W_MNT_NET, "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
End If

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Ce crédit documentaire n'étant pas confirmé par notre banque, nous verserons ce montant en votre faveur ";
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "conformément à vos instructions,"
XPrt.CurrentX = prtMinMarge + 3000
XPrt.CurrentY = XPrt.CurrentY - prtlineHeight
XPrt.FontUnderline = True:  XPrt.Print "après réception de la couverture correspondante"
XPrt.CurrentX = prtMinMarge + 7250
XPrt.CurrentY = XPrt.CurrentY - prtlineHeight
XPrt.FontUnderline = False: XPrt.Print ", prévue à l'échéance du ";
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print wDate_Echeance_CR & "," & " selon les termes du crédit.";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Veuillez agréer, Messieurs, nos salutations distinguées."

XPrt.CurrentX = prtMinMarge + 5000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.FontBold = True
XPrt.Print paramSignature;
XPrt.FontBold = False

End Sub

Public Sub prtCDO_Courrier_UTI_ACCORDRECU_C_Avue_AR10()
'---------------------------------------------------------
Dim X As String
Dim W_Cumul_Comm As Currency
Dim W_Cumul_TVA As Currency
Dim W_MNT_NET As Currency
Dim C As Integer
Dim W_Code_CRD As String
Dim MNT As String, MTV As String
Dim W_REGCOM_5C As String, W_DOSBEN_5C As String

W_Cumul_Comm = 0
W_Cumul_TVA = 0
W_MNT_NET = 0

Line2_Ecart = 10: prtCDO_Courrier_Trame (Line2_Ecart)
' Call frmElpPrt.prtTrame(prtMinMarge - 50, prtCorpsY - 100, prtMaxMarge, prtCorpsY + prtlineHeight * 10 + 100, "B", 245)

XPrt.CurrentY = prtCorpsY
XPrt.FontSize = 14: XPrt.FontBold = True: XPrt.FontUnderline = True
frmElpPrt.prtCentré prtMedX, "CREDIT DOCUMENTAIRE"
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
frmElpPrt.prtCentré prtMedX, "ACCUSE DE RECEPTION"

XPrt.FontSize = 11: XPrt.FontBold = False: XPrt.FontUnderline = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
'Call frmElpPrt.prtTrame(prtMinMarge - 50, XPrt.CurrentY - 50, prtMaxMarge, XPrt.CurrentY + prtlineHeight * 3 + 50, , 245)

XPrt.CurrentX = prtMinMarge: XPrt.Print "  V/Référence :";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1700: XPrt.Print meCDO_Courrier.YCDOUTI0.CDOUTIRER;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Donneur d'ordre";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1700: XPrt.Print meCDO_Courrier.DON_Concat;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Bénéficiaire";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1700: XPrt.Print meCDO_Courrier.BEN_Concat;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Montant";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentX = prtMinMarge + 4100: XPrt.Print "Devise :";
XPrt.CurrentX = prtMinMarge + 7500: XPrt.Print "Validité :";
XPrt.FontBold = True
X = Trim(Format$(meYCDODOS0.CDODOSMON, "### ### ### ##0.00"))
XPrt.CurrentX = prtMinMarge + 1700
XPrt.Print X;
XPrt.CurrentX = prtMinMarge + 5000: XPrt.Print meYCDODOS0.CDODOSDEV;
XPrt.CurrentX = prtMinMarge + 8500: XPrt.Print wDate_Validité;
XPrt.FontBold = False

XPrt.FontSize = 11
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.CurrentX = prtMinMarge
XPrt.Print "Messieurs,";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Nous nous référons à notre lettre du " & wDate_Remise_Util & " relative à l'envoi de vos documents à la banque émettrice pour ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "accord de paiement. ";

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.CurrentX = prtMinMarge
XPrt.Print "Nous vous informons avoir reçu l'autorisation de la banque émettrice pour procéder au règlement des documents ";

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinMarge
XPrt.Print "de  " & meYCDODOS0.CDODOSDEV & "  " & Format$(meCDO_Courrier.YCDOUTI0.CDOUTIMPA, "### ### ### ##0.00");

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.CurrentX = prtMinMarge
XPrt.Print "Montant de l'utilisation ";
XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meYCDODOS0.CDODOSDEV & "            " & Format$(meCDO_Courrier.YCDOUTI0.CDOUTIMPA, "### ### ### ##0.00");

' >>>>> Lettre D'ACCUSE RECEPTION = Toujours destinataire BENEF donc Toujours comm. ligne CREDIT
W_Code_CRD = "C"

' La somme des commissions lues
For C = 1 To meCDO_Courrier.Com_Nb
    If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD Then
        W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    End If
Next C

If W_Cumul_Comm <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Nos frais et commissions selon détail ci-après déduits : ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print "Devise             Montant        TVA 19,60%";
    XPrt.CurrentX = prtMinMarge
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
End If

' >>>>> Remettre à zéro W_Cumul_Comm avant de passer dans la boucle lignes de commission
W_Cumul_Comm = 0
For C = 1 To meCDO_Courrier.Com_Nb

' Commission ECNF...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "ECNF  " And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission de confirmation ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission ECSIL...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "ECSIL " And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission de confirmation silencieuse ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission ENOTIF...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "ENOTIF" And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission de notification ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EPDIF...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EPDIF " And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission d'engagement paiement différé ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EPDIFD...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EPDIFD" And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission d'engagement paiement différé ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EACCEP...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EACCEP" And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission d'acceptation ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EACCED...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EACCED" And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission d'acceptation ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission ELVD...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "ELVD  " And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission de levée de documents ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission ERFA...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "ERFA  " And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Frais de ports et telex ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EDOCIR...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EDOCIR" And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission de documents irréguliers ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EACD...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EACD  " And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Autre commission documentaire ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EMODIF...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EMODIF" And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission de modification ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EFBEMT... Toujours en PLUS pour le bénéficiaire à l'accusé de réception
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EFBEMT" And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Frais banque émettrice ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' >>>>> Fin de boucle sur lignes de commissions
Next C

If W_Cumul_Comm <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
    XPrt.CurrentX = prtMinMarge
    XPrt.CurrentX = prtMinMarge + 3000: XPrt.Print "Montant  total ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(1);
    MNT = Format$(W_Cumul_Comm, "### ### ### ##0.00")
    MTV = Format$(W_Cumul_TVA, "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
End If

W_MNT_NET = meCDO_Courrier.YCDOUTI0.CDOUTIMPA - W_Cumul_Comm - W_Cumul_TVA
' If W_MNT_NET <> 0 Then
If W_Cumul_Comm <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
    XPrt.CurrentX = prtMinMarge
    Call frmElpPrt.prtTrame(prtMinMarge + 3000 - 50, XPrt.CurrentY - 100, prtMaxMarge - 2000, XPrt.CurrentY + prtlineHeight + 100, 245)
    XPrt.CurrentX = prtMinMarge + 3000: XPrt.Print "Montant  total  net ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(1);
    MNT = Format$(W_MNT_NET, "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
End If

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Que nous verserons en votre faveur conformément à vos instructions valeur " & wDate_Valeur_CR & " au crédit de votre ";
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "compte ";

' Test du compte de la ligne CREDIT
Mid$(wCompte_CR, 1, 5) = W_REGCOM_5C
Mid$(meYCDODOS0.CDODOSBEN, 3, 5) = W_DOSBEN_5C
If W_REGCOM_5C = W_DOSBEN_5C And meYCDODOS0.CDODOSBER = " " Then
    XPrt.CurrentX = prtMinMarge + 800: XPrt.Print "sur nos livres. ";
Else
    XPrt.CurrentX = prtMinMarge + 800: XPrt.Print meCDO_Courrier.IBAN;
End If

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Veuillez agréer, Messieurs, nos salutations distinguées."

XPrt.CurrentX = prtMinMarge + 5000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.FontBold = True
XPrt.Print paramSignature;
XPrt.FontBold = False

End Sub

Public Sub prtCDO_Courrier_UTI_ACCORDRECU_C_Pdif_AR11()
'---------------------------------------------------------
Dim X As String
Dim W_Cumul_Comm As Currency
Dim W_Cumul_TVA As Currency
Dim W_MNT_NET As Currency
Dim C As Integer
Dim W_Code_CRD As String
Dim MNT As String, MTV As String
Dim W_REGCOM_5C As String, W_DOSBEN_5C As String

W_Cumul_Comm = 0
W_Cumul_TVA = 0
W_MNT_NET = 0

Line2_Ecart = 10: prtCDO_Courrier_Trame (Line2_Ecart)
' Call frmElpPrt.prtTrame(prtMinMarge - 50, prtCorpsY - 100, prtMaxMarge, prtCorpsY + prtlineHeight * 10 + 100, "B", 245)

XPrt.CurrentY = prtCorpsY
XPrt.FontSize = 14: XPrt.FontBold = True: XPrt.FontUnderline = True
frmElpPrt.prtCentré prtMedX, "CREDIT DOCUMENTAIRE"
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
frmElpPrt.prtCentré prtMedX, "ACCUSE DE RECEPTION"

XPrt.FontSize = 11: XPrt.FontBold = False: XPrt.FontUnderline = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
'Call frmElpPrt.prtTrame(prtMinMarge - 50, XPrt.CurrentY - 50, prtMaxMarge, XPrt.CurrentY + prtlineHeight * 3 + 50, , 245)

XPrt.CurrentX = prtMinMarge: XPrt.Print "  V/Référence :";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1700: XPrt.Print meCDO_Courrier.YCDOUTI0.CDOUTIRER;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Donneur d'ordre";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1700: XPrt.Print meCDO_Courrier.DON_Concat;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Bénéficiaire";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1700: XPrt.Print meCDO_Courrier.BEN_Concat;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Montant";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentX = prtMinMarge + 4100: XPrt.Print "Devise :";
XPrt.CurrentX = prtMinMarge + 7500: XPrt.Print "Validité :";
XPrt.FontBold = True
X = Trim(Format$(meYCDODOS0.CDODOSMON, "### ### ### ##0.00"))
XPrt.CurrentX = prtMinMarge + 1700
XPrt.Print X;
XPrt.CurrentX = prtMinMarge + 5000: XPrt.Print meYCDODOS0.CDODOSDEV;
XPrt.CurrentX = prtMinMarge + 8500: XPrt.Print wDate_Validité;
XPrt.FontBold = False

XPrt.FontSize = 11
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.CurrentX = prtMinMarge
XPrt.Print "Messieurs,";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Nous nous référons à notre lettre du " & wDate_Remise_Util & " relative à l'envoi de vos documents à la banque émettrice pour ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "accord de paiement. ";

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.CurrentX = prtMinMarge
XPrt.Print "Nous vous informons avoir reçu l'autorisation de la banque émettrice pour procéder au règlement des documents ";

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinMarge
XPrt.Print "de  " & meYCDODOS0.CDODOSDEV & "  " & Format$(meCDO_Courrier.YCDOUTI0.CDOUTIMPA, "### ### ### ##0.00");

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.CurrentX = prtMinMarge
XPrt.Print "Montant de l'utilisation ";
XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meYCDODOS0.CDODOSDEV & "            " & Format$(meCDO_Courrier.YCDOUTI0.CDOUTIMPA, "### ### ### ##0.00");

' >>>>> Lettre D'ACCUSE RECEPTION = Toujours destinataire BENEF donc Toujours comm. ligne CREDIT
W_Code_CRD = "C"

' La somme des commissions lues
For C = 1 To meCDO_Courrier.Com_Nb
    If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD Then
        W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    End If
Next C

If W_Cumul_Comm <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Nos frais et commissions selon détail ci-après déduits : ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print "Devise             Montant        TVA 19,60%";
    XPrt.CurrentX = prtMinMarge
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
End If

' >>>>> Remettre à zéro W_Cumul_Comm avant de passer dans la boucle lignes de commission
W_Cumul_Comm = 0
For C = 1 To meCDO_Courrier.Com_Nb

' Commission ECNF...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "ECNF  " And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission de confirmation ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission ECSIL...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "ECSIL " And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission de confirmation silencieuse ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission ENOTIF...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "ENOTIF" And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission de notification ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EPDIF...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EPDIF " And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission d'engagement paiement différé ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EPDIFD...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EPDIFD" And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission d'engagement paiement différé ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EACCEP...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EACCEP" And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission d'acceptation ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EACCED...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EACCED" And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission d'acceptation ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission ELVD...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "ELVD  " And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission de levée de documents ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission ERFA...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "ERFA  " And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Frais de ports et telex ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EDOCIR...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EDOCIR" And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission de documents irréguliers ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EACD...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EACD  " And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Autre commission documentaire ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EMODIF...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EMODIF" And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission de modification ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EFBEMT... Toujours en PLUS pour le bénéficiaire à l'accusé de réception
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EFBEMT" And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Frais banque émettrice ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' >>>>> Fin de boucle sur lignes de commissions
Next C

If W_Cumul_Comm <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
    XPrt.CurrentX = prtMinMarge
    XPrt.CurrentX = prtMinMarge + 3000: XPrt.Print "Montant  total ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(1);
    MNT = Format$(W_Cumul_Comm, "### ### ### ##0.00")
    MTV = Format$(W_Cumul_TVA, "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
End If

W_MNT_NET = meCDO_Courrier.YCDOUTI0.CDOUTIMPA - W_Cumul_Comm - W_Cumul_TVA
' If W_MNT_NET <> 0 Then
If W_Cumul_Comm <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
    XPrt.CurrentX = prtMinMarge
    Call frmElpPrt.prtTrame(prtMinMarge + 3000 - 50, XPrt.CurrentY - 100, prtMaxMarge - 2000, XPrt.CurrentY + prtlineHeight + 100, 245)
    XPrt.CurrentX = prtMinMarge + 3000: XPrt.Print "Montant  total  net ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(1);
    MNT = Format$(W_MNT_NET, "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
End If

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Que nous verserons en votre faveur conformément à vos instructions à l'échéance du " & wDate_Echeance_CR & " au crédit de ";
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "votre compte ";

' Test du compte de la ligne CREDIT
Mid$(wCompte_CR, 1, 5) = W_REGCOM_5C
Mid$(meYCDODOS0.CDODOSBEN, 3, 5) = W_DOSBEN_5C
If W_REGCOM_5C = W_DOSBEN_5C And meYCDODOS0.CDODOSBER = " " Then
    XPrt.CurrentX = prtMinMarge + 1350: XPrt.Print "sur nos livres. ";
Else
    XPrt.CurrentX = prtMinMarge + 1350: XPrt.Print meCDO_Courrier.IBAN;
End If

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Veuillez agréer, Messieurs, nos salutations distinguées."

XPrt.CurrentX = prtMinMarge + 5000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.FontBold = True
XPrt.Print paramSignature;
XPrt.FontBold = False

End Sub

Public Sub prtCDO_Courrier_UTI_ACCORDRECU_N_Avue_AR12()
'---------------------------------------------------------
Dim X As String
Dim W_Cumul_Comm As Currency
Dim W_Cumul_TVA As Currency
Dim W_MNT_NET As Currency
Dim C As Integer
Dim W_Code_CRD As String
Dim MNT As String, MTV As String

W_Cumul_Comm = 0
W_Cumul_TVA = 0
W_MNT_NET = 0

Line2_Ecart = 10: prtCDO_Courrier_Trame (Line2_Ecart)
' Call frmElpPrt.prtTrame(prtMinMarge - 50, prtCorpsY - 100, prtMaxMarge, prtCorpsY + prtlineHeight * 10 + 100, "B", 245)

XPrt.CurrentY = prtCorpsY
XPrt.FontSize = 14: XPrt.FontBold = True: XPrt.FontUnderline = True
frmElpPrt.prtCentré prtMedX, "CREDIT DOCUMENTAIRE"
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
frmElpPrt.prtCentré prtMedX, "ACCUSE DE RECEPTION"

XPrt.FontSize = 11: XPrt.FontBold = False: XPrt.FontUnderline = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
' Call frmElpPrt.prtTrame(prtMinMarge - 50, XPrt.CurrentY - 50, prtMaxMarge, XPrt.CurrentY + prtlineHeight * 3 + 50, , 245)

XPrt.CurrentX = prtMinMarge: XPrt.Print "  V/Référence :";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1700: XPrt.Print meCDO_Courrier.YCDOUTI0.CDOUTIRER;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Donneur d'ordre";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1700: XPrt.Print meCDO_Courrier.DON_Concat;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Bénéficiaire";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1700: XPrt.Print meCDO_Courrier.BEN_Concat;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Montant";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentX = prtMinMarge + 4100: XPrt.Print "Devise :";
XPrt.CurrentX = prtMinMarge + 7500: XPrt.Print "Validité :";
XPrt.FontBold = True
X = Trim(Format$(meYCDODOS0.CDODOSMON, "### ### ### ##0.00"))
XPrt.CurrentX = prtMinMarge + 1700
XPrt.Print X;
XPrt.CurrentX = prtMinMarge + 5000: XPrt.Print meYCDODOS0.CDODOSDEV;
XPrt.CurrentX = prtMinMarge + 8500: XPrt.Print wDate_Validité;
XPrt.FontBold = False

XPrt.FontSize = 11
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.CurrentX = prtMinMarge
XPrt.Print "Messieurs,";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Nous nous référons à notre lettre du " & wDate_Remise_Util & " relative à l'envoi de vos documents à la banque émettrice pour ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "accord de paiement. ";

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.CurrentX = prtMinMarge
XPrt.Print "Nous vous informons avoir reçu l'autorisation de la banque émettrice pour procéder au règlement des documents ";

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinMarge
XPrt.Print "de  " & meYCDODOS0.CDODOSDEV & "  " & Format$(meCDO_Courrier.YCDOUTI0.CDOUTIMPA, "### ### ### ##0.00");

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.CurrentX = prtMinMarge
XPrt.Print "Montant de l'utilisation ";
XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meYCDODOS0.CDODOSDEV & "            " & Format$(meCDO_Courrier.YCDOUTI0.CDOUTIMPA, "### ### ### ##0.00");

' >>>>> Lettre D'ACCUSE RECEPTION = Toujours destinataire BENEF donc Toujours comm. ligne CREDIT
W_Code_CRD = "C"

' La somme des commissions lues
For C = 1 To meCDO_Courrier.Com_Nb
    If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD Then
        W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    End If
Next C

If W_Cumul_Comm <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Nos frais et commissions selon détail ci-après déduits : ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print "Devise             Montant        TVA 19,60%";
    XPrt.CurrentX = prtMinMarge
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
End If

' >>>>> Remettre à zéro W_Cumul_Comm avant de passer dans la boucle lignes de commission
W_Cumul_Comm = 0
For C = 1 To meCDO_Courrier.Com_Nb

' Commission ECNF...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "ECNF  " And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission de confirmation ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission ECSIL...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "ECSIL " And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission de confirmation silencieuse ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission ENOTIF...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "ENOTIF" And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission de notification ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EPDIF...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EPDIF " And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission d'engagement paiement différé ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EPDIFD...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EPDIFD" And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission d'engagement paiement différé ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EACCEP...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EACCEP" And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission d'acceptation ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EACCED...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EACCED" And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission d'acceptation ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission ELVD...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "ELVD  " And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission de levée de documents ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission ERFA...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "ERFA  " And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Frais de ports et telex ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EDOCIR...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EDOCIR" And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission de documents irréguliers ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EACD...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EACD  " And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Autre commission documentaire ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EMODIF...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EMODIF" And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission de modification ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EFBEMT... Toujours en PLUS pour le bénéficiaire à l'accusé de réception
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EFBEMT" And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Frais banque émettrice ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' >>>>> Fin de boucle sur lignes de commissions
Next C

If W_Cumul_Comm <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
    XPrt.CurrentX = prtMinMarge
    XPrt.CurrentX = prtMinMarge + 3000: XPrt.Print "Montant  total ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(1);
    MNT = Format$(W_Cumul_Comm, "### ### ### ##0.00")
    MTV = Format$(W_Cumul_TVA, "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
End If

W_MNT_NET = meCDO_Courrier.YCDOUTI0.CDOUTIMPA - W_Cumul_Comm - W_Cumul_TVA
' If W_MNT_NET <> 0 Then
If W_Cumul_Comm <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
    XPrt.CurrentX = prtMinMarge
    Call frmElpPrt.prtTrame(prtMinMarge + 3000 - 50, XPrt.CurrentY - 100, prtMaxMarge - 2000, XPrt.CurrentY + prtlineHeight + 100, 245)
    XPrt.CurrentX = prtMinMarge + 3000: XPrt.Print "Montant  total  net ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(1);
    MNT = Format$(W_MNT_NET, "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
End If

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Ce crédit documentaire n'étant pas confirmé par notre banque, nous verserons ce montant en votre faveur, ";
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "conformément à vos instructions ";
XPrt.FontUnderline = True
XPrt.Print "après réception de la couverture correspondante";
XPrt.FontUnderline = False
XPrt.Print " prévue le " & wDate_Valeur_CR;
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "selon les termes du crédit chez  " & meCDO_Courrier.IBAN;

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Veuillez agréer, Messieurs, nos salutations distinguées."

XPrt.CurrentX = prtMinMarge + 5000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.FontBold = True
XPrt.Print paramSignature;
XPrt.FontBold = False

End Sub
Public Sub prtCDO_Courrier_UTI_ACCORDRECU_N_Pdif_AR13()
'---------------------------------------------------------
Dim X As String
Dim W_Cumul_Comm As Currency
Dim W_Cumul_TVA As Currency
Dim W_MNT_NET As Currency
Dim C As Integer
Dim W_Code_CRD As String
Dim MNT As String, MTV As String

W_Cumul_Comm = 0
W_Cumul_TVA = 0
W_MNT_NET = 0

Line2_Ecart = 10: prtCDO_Courrier_Trame (Line2_Ecart)
' Call frmElpPrt.prtTrame(prtMinMarge - 50, prtCorpsY - 100, prtMaxMarge, prtCorpsY + prtlineHeight * 10 + 100, "B", 245)

XPrt.CurrentY = prtCorpsY
XPrt.FontSize = 14: XPrt.FontBold = True: XPrt.FontUnderline = True
frmElpPrt.prtCentré prtMedX, "CREDIT DOCUMENTAIRE"
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
frmElpPrt.prtCentré prtMedX, "ACCUSE DE RECEPTION"

XPrt.FontSize = 11: XPrt.FontBold = False: XPrt.FontUnderline = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
'Call frmElpPrt.prtTrame(prtMinMarge - 50, XPrt.CurrentY - 50, prtMaxMarge, XPrt.CurrentY + prtlineHeight * 3 + 50, , 245)

XPrt.CurrentX = prtMinMarge: XPrt.Print "  V/Référence :";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1700: XPrt.Print meCDO_Courrier.YCDOUTI0.CDOUTIRER;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Donneur d'ordre";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1700: XPrt.Print meCDO_Courrier.DON_Concat;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Bénéficiaire";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1700: XPrt.Print meCDO_Courrier.BEN_Concat;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Montant";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentX = prtMinMarge + 4100: XPrt.Print "Devise :";
XPrt.CurrentX = prtMinMarge + 7500: XPrt.Print "Validité :";
XPrt.FontBold = True
X = Trim(Format$(meYCDODOS0.CDODOSMON, "### ### ### ##0.00"))
XPrt.CurrentX = prtMinMarge + 1700
XPrt.Print X;
XPrt.CurrentX = prtMinMarge + 5000: XPrt.Print meYCDODOS0.CDODOSDEV;
XPrt.CurrentX = prtMinMarge + 8500: XPrt.Print wDate_Validité;
XPrt.FontBold = False

XPrt.FontSize = 11
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.CurrentX = prtMinMarge
XPrt.Print "Messieurs,";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Nous nous référons à notre lettre du " & wDate_Remise_Util & " relative à l'envoi de vos documents à la banque émettrice pour ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "accord de paiement. ";

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.CurrentX = prtMinMarge
XPrt.Print "Nous vous informons avoir reçu l'autorisation de la banque émettrice pour procéder au règlement des documents ";

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinMarge
XPrt.Print "de  " & meYCDODOS0.CDODOSDEV & "  " & Format$(meCDO_Courrier.YCDOUTI0.CDOUTIMPA, "### ### ### ##0.00");

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.CurrentX = prtMinMarge
XPrt.Print "Montant de l'utilisation ";
XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meYCDODOS0.CDODOSDEV & "            " & Format$(meCDO_Courrier.YCDOUTI0.CDOUTIMPA, "### ### ### ##0.00");

' >>>>> Lettre D'ACCUSE RECEPTION = Toujours destinataire BENEF donc Toujours comm. ligne CREDIT
W_Code_CRD = "C"

' La somme des commissions lues
For C = 1 To meCDO_Courrier.Com_Nb
    If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD Then
        W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    End If
Next C

If W_Cumul_Comm <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Nos frais et commissions selon détail ci-après déduits : ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print "Devise             Montant        TVA 19,60%";
    XPrt.CurrentX = prtMinMarge
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
End If

' >>>>> Remettre à zéro W_Cumul_Comm avant de passer dans la boucle lignes de commission
W_Cumul_Comm = 0
For C = 1 To meCDO_Courrier.Com_Nb

' Commission ECNF...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "ECNF  " And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission de confirmation ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission ECSIL...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "ECSIL " And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission de confirmation silencieuse ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission ENOTIF...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "ENOTIF" And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission de notification ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EPDIF...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EPDIF " And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission d'engagement paiement différé ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EPDIFD...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EPDIFD" And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission d'engagement paiement différé ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EACCEP...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EACCEP" And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission d'acceptation ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EACCED...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EACCED" And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission d'acceptation ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission ELVD...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "ELVD  " And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission de levée de documents ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission ERFA...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "ERFA  " And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Frais de ports et telex ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EDOCIR...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EDOCIR" And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission de documents irréguliers ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EACD...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EACD  " And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Autre commission documentaire ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EMODIF...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EMODIF" And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission de modification ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EFBEMT... Toujours en PLUS pour le bénéficiaire à l'accusé de réception
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EFBEMT" And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Frais banque émettrice ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' >>>>> Fin de boucle sur lignes de commissions
Next C

If W_Cumul_Comm <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
    XPrt.CurrentX = prtMinMarge
    XPrt.CurrentX = prtMinMarge + 3000: XPrt.Print "Montant  total ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(1);
    MNT = Format$(W_Cumul_Comm, "### ### ### ##0.00")
    MTV = Format$(W_Cumul_TVA, "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
End If

W_MNT_NET = meCDO_Courrier.YCDOUTI0.CDOUTIMPA - W_Cumul_Comm - W_Cumul_TVA
' If W_MNT_NET <> 0 Then
If W_Cumul_Comm <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
    XPrt.CurrentX = prtMinMarge
    Call frmElpPrt.prtTrame(prtMinMarge + 3000 - 50, XPrt.CurrentY - 100, prtMaxMarge - 2000, XPrt.CurrentY + prtlineHeight + 100, 245)
    XPrt.CurrentX = prtMinMarge + 3000: XPrt.Print "Montant  total  net ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(1);
    MNT = Format$(W_MNT_NET, "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
End If

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Ce crédit documentaire n'étant pas confirmé par notre banque, nous verserons ce montant en votre faveur, ";
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "conformément à vos instructions ";
XPrt.FontUnderline = True
XPrt.Print "après réception de la couverture correspondante";
XPrt.FontUnderline = False
XPrt.Print " à l'échéance prévue le ";
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print wDate_Echeance_CR & " selon les termes du crédit chez  " & meCDO_Courrier.IBAN;

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Veuillez agréer, Messieurs, nos salutations distinguées."

XPrt.CurrentX = prtMinMarge + 5000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.FontBold = True
XPrt.Print paramSignature;
XPrt.FontBold = False

End Sub

Public Sub prtCDO_Courrier_UTI_CONFORME_C_Avue_AR1()
'---------------------------------------------------------
Dim X As String
Dim W_Cumul_Comm As Currency
Dim W_Cumul_TVA As Currency
Dim W_MNT_NET As Currency
Dim C As Integer
Dim W_Code_CRD As String
Dim MNT As String, MTV As String
Dim W_REGCOM_5C As String, W_DOSBEN_5C As String

W_Cumul_Comm = 0
W_Cumul_TVA = 0
W_MNT_NET = 0

Line2_Ecart = 10: prtCDO_Courrier_Trame (Line2_Ecart)
' Call frmElpPrt.prtTrame(prtMinMarge - 50, prtCorpsY - 100, prtMaxMarge, prtCorpsY + prtlineHeight * 10 + 100, "B", 245)

XPrt.CurrentY = prtCorpsY
XPrt.FontSize = 14: XPrt.FontBold = True: XPrt.FontUnderline = True
frmElpPrt.prtCentré prtMedX, "CREDIT DOCUMENTAIRE"
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
frmElpPrt.prtCentré prtMedX, "ACCUSE DE RECEPTION"

XPrt.FontSize = 11: XPrt.FontBold = False: XPrt.FontUnderline = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
'Call frmElpPrt.prtTrame(prtMinMarge - 50, XPrt.CurrentY - 50, prtMaxMarge, XPrt.CurrentY + prtlineHeight * 3 + 50, , 245)

XPrt.CurrentX = prtMinMarge: XPrt.Print "  V/Référence :";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1700: XPrt.Print meCDO_Courrier.YCDOUTI0.CDOUTIRER;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Donneur d'ordre";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1700: XPrt.Print meCDO_Courrier.DON_Concat;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Bénéficiaire";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1700: XPrt.Print meCDO_Courrier.BEN_Concat;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Montant";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentX = prtMinMarge + 4100: XPrt.Print "Devise :";
XPrt.CurrentX = prtMinMarge + 7500: XPrt.Print "Validité :";
XPrt.FontBold = True
X = Trim(Format$(meYCDODOS0.CDODOSMON, "### ### ### ##0.00"))
XPrt.CurrentX = prtMinMarge + 1700
XPrt.Print X;
XPrt.CurrentX = prtMinMarge + 5000: XPrt.Print meYCDODOS0.CDODOSDEV;
XPrt.CurrentX = prtMinMarge + 8500: XPrt.Print wDate_Validité;
XPrt.FontBold = False

XPrt.FontSize = 11
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 4
XPrt.CurrentX = prtMinMarge
XPrt.Print "Messieurs,";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Nous accusons réception de votre lettre nous remettant les documents de ";
XPrt.Print meYCDODOS0.CDODOSDEV & "  " & Trim(Format$(meCDO_Courrier.YCDOUTI0.CDOUTIMPA, "### ### ### ##0.00"));
XPrt.Print "  en réalisation du ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "crédit documentaire précité. ";

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.CurrentX = prtMinMarge
XPrt.Print "Montant de l'utilisation ";
XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meYCDODOS0.CDODOSDEV & "            " & Format$(meCDO_Courrier.YCDOUTI0.CDOUTIMPA, "### ### ### ##0.00");

' >>>>> Lettre D'ACCUSE RECEPTION = Toujours destinataire BENEF donc Toujours comm. ligne CREDIT
W_Code_CRD = "C"


' La somme des commissions lues
For C = 1 To meCDO_Courrier.Com_Nb
    If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD Then
        W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    End If
Next C

If W_Cumul_Comm <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Nos frais et commissions selon détail ci-après déduits : ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print "Devise             Montant        TVA 19,60%";
    XPrt.CurrentX = prtMinMarge
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
End If

' >>>>> Remettre à zéro W_Cumul_Comm avant de passer dans la boucle lignes de commission
W_Cumul_Comm = 0
For C = 1 To meCDO_Courrier.Com_Nb

' Commission ECNF...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "ECNF  " And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission de confirmation ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission ECSIL...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "ECSIL " And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission de confirmation silencieuse ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission ENOTIF...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "ENOTIF" And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission de notification ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EPDIF...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EPDIF " And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission d'engagement paiement différé ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EPDIFD...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EPDIFD" And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission d'engagement paiement différé ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EACCEP...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EACCEP" And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission d'acceptation ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EACCED...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EACCED" And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission d'acceptation ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission ELVD...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "ELVD  " And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission de levée de documents ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission ERFA...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "ERFA  " And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Frais de ports et telex ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EDOCIR...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EDOCIR" And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission de documents irréguliers ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EACD...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EACD  " And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Autre commission documentaire ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EMODIF...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EMODIF" And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission de modification ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EFBEMT... Toujours en PLUS pour le bénéficiaire à l'accusé de réception
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EFBEMT" And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Frais banque émettrice ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' >>>>> Fin de boucle sur lignes de commissions
Next C

If W_Cumul_Comm <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
    XPrt.CurrentX = prtMinMarge
    XPrt.CurrentX = prtMinMarge + 3000: XPrt.Print "Montant  total ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(1);
    MNT = Format$(W_Cumul_Comm, "### ### ### ##0.00")
    MTV = Format$(W_Cumul_TVA, "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
End If

W_MNT_NET = meCDO_Courrier.YCDOUTI0.CDOUTIMPA - W_Cumul_Comm - W_Cumul_TVA
' If W_MNT_NET <> 0 Then
If W_Cumul_Comm <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
    XPrt.CurrentX = prtMinMarge
    Call frmElpPrt.prtTrame(prtMinMarge + 3000 - 50, XPrt.CurrentY - 100, prtMaxMarge - 2000, XPrt.CurrentY + prtlineHeight + 100, 245)
    XPrt.CurrentX = prtMinMarge + 3000: XPrt.Print "Montant  total  net ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(1);
    MNT = Format$(W_MNT_NET, "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
End If

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Que selon les termes du crédit et conformément à vos instructions, nous verserons en votre faveur valeur ";
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print wDate_Valeur_CR & "  au crédit de votre compte ";

' Test du compte de la ligne CREDIT
Mid$(wCompte_CR, 1, 5) = W_REGCOM_5C
Mid$(meYCDODOS0.CDODOSBEN, 3, 5) = W_DOSBEN_5C
If W_REGCOM_5C = W_DOSBEN_5C And meYCDODOS0.CDODOSBER = " " Then
    XPrt.CurrentX = prtMinMarge + 3400: XPrt.Print "sur nos livres. ";
Else
    XPrt.CurrentX = prtMinMarge + 3400: XPrt.Print " " & meCDO_Courrier.IBAN;
End If

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Veuillez agréer, Messieurs, nos salutations distinguées."

XPrt.CurrentX = prtMinMarge + 5000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.FontBold = True
XPrt.Print paramSignature;
XPrt.FontBold = False

End Sub


Public Sub prtCDO_Courrier_UTI_CONFORME_C_Pdif_AR2()
'---------------------------------------------------------
Dim X As String
Dim W_Cumul_Comm As Currency
Dim W_Cumul_TVA As Currency
Dim W_MNT_NET As Currency
Dim C As Integer
Dim W_Code_CRD As String
Dim MNT As String, MTV As String
Dim W_REGCOM_5C As String, W_DOSBEN_5C As String

W_Cumul_Comm = 0
W_Cumul_TVA = 0
W_MNT_NET = 0

Line2_Ecart = 10: prtCDO_Courrier_Trame (Line2_Ecart)
' Call frmElpPrt.prtTrame(prtMinMarge - 50, prtCorpsY - 100, prtMaxMarge, prtCorpsY + prtlineHeight * 10 + 100, "B", 245)

XPrt.CurrentY = prtCorpsY
XPrt.FontSize = 14: XPrt.FontBold = True: XPrt.FontUnderline = True
frmElpPrt.prtCentré prtMedX, "CREDIT DOCUMENTAIRE"
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
frmElpPrt.prtCentré prtMedX, "ACCUSE DE RECEPTION"

XPrt.FontSize = 11: XPrt.FontBold = False: XPrt.FontUnderline = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
'Call frmElpPrt.prtTrame(prtMinMarge - 50, XPrt.CurrentY - 50, prtMaxMarge, XPrt.CurrentY + prtlineHeight * 3 + 50, , 245)

XPrt.CurrentX = prtMinMarge: XPrt.Print "  V/Référence :";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1700: XPrt.Print meCDO_Courrier.YCDOUTI0.CDOUTIRER;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Donneur d'ordre";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1700: XPrt.Print meCDO_Courrier.DON_Concat;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Bénéficiaire";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1700: XPrt.Print meCDO_Courrier.BEN_Concat;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Montant";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentX = prtMinMarge + 4100: XPrt.Print "Devise :";
XPrt.CurrentX = prtMinMarge + 7500: XPrt.Print "Validité :";
XPrt.FontBold = True
X = Trim(Format$(meYCDODOS0.CDODOSMON, "### ### ### ##0.00"))
XPrt.CurrentX = prtMinMarge + 1700
XPrt.Print X;
XPrt.CurrentX = prtMinMarge + 5000: XPrt.Print meYCDODOS0.CDODOSDEV;
XPrt.CurrentX = prtMinMarge + 8500: XPrt.Print wDate_Validité;
XPrt.FontBold = False

XPrt.FontSize = 11
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 4
XPrt.CurrentX = prtMinMarge
XPrt.Print "Messieurs,";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Nous accusons réception de votre lettre nous remettant les documents de ";
XPrt.Print meYCDODOS0.CDODOSDEV & "  " & Trim(Format$(meCDO_Courrier.YCDOUTI0.CDOUTIMPA, "### ### ### ##0.00"));
XPrt.Print "  en réalisation du ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "crédit documentaire précité. ";

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.CurrentX = prtMinMarge
XPrt.Print "Montant de l'utilisation ";
XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meYCDODOS0.CDODOSDEV & "            " & Format$(meCDO_Courrier.YCDOUTI0.CDOUTIMPA, "### ### ### ##0.00");

' >>>>> Lettre D'ACCUSE RECEPTION = Toujours destinataire BENEF donc Toujours comm. ligne CREDIT
W_Code_CRD = "C"


' La somme des commissions lues
For C = 1 To meCDO_Courrier.Com_Nb
    If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD Then
        W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    End If
Next C

If W_Cumul_Comm <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Nos frais et commissions selon détail ci-après déduits : ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print "Devise             Montant        TVA 19,60%";
    XPrt.CurrentX = prtMinMarge
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
End If

' >>>>> Remettre à zéro W_Cumul_Comm avant de passer dans la boucle lignes de commission
W_Cumul_Comm = 0
For C = 1 To meCDO_Courrier.Com_Nb

' Commission ECNF...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "ECNF  " And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission de confirmation ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission ECSIL...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "ECSIL " And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission de confirmation silencieuse ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission ENOTIF...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "ENOTIF" And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission de notification ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EPDIF...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EPDIF " And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission d'engagement paiement différé ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EPDIFD...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EPDIFD" And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission d'engagement paiement différé ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EACCEP...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EACCEP" And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission d'acceptation ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EACCED...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EACCED" And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission d'acceptation ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission ELVD...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "ELVD  " And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission de levée de documents ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission ERFA...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "ERFA  " And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Frais de ports et telex ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EDOCIR...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EDOCIR" And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission de documents irréguliers ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EACD...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EACD  " And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Autre commission documentaire ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EMODIF...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EMODIF" And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission de modification ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EFBEMT... Toujours en PLUS pour le bénéficiaire à l'accusé de réception
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EFBEMT" And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Frais banque émettrice ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' >>>>> Fin de boucle sur lignes de commissions
Next C

If W_Cumul_Comm <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
    XPrt.CurrentX = prtMinMarge
    XPrt.CurrentX = prtMinMarge + 3000: XPrt.Print "Montant  total ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(1);
    MNT = Format$(W_Cumul_Comm, "### ### ### ##0.00")
    MTV = Format$(W_Cumul_TVA, "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
End If

W_MNT_NET = meCDO_Courrier.YCDOUTI0.CDOUTIMPA - W_Cumul_Comm - W_Cumul_TVA
' If W_MNT_NET <> 0 Then
If W_Cumul_Comm <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
    XPrt.CurrentX = prtMinMarge
    Call frmElpPrt.prtTrame(prtMinMarge + 3000 - 50, XPrt.CurrentY - 100, prtMaxMarge - 2000, XPrt.CurrentY + prtlineHeight + 100, 245)
    XPrt.CurrentX = prtMinMarge + 3000: XPrt.Print "Montant  total  net ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(1);
    MNT = Format$(W_MNT_NET, "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
End If

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Que selon les termes du crédit et conformément à vos instructions, nous verserons en votre faveur à ";
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "l'échéance du " & wDate_Echeance_CR & "  au crédit de votre compte ";

' Test du compte de la ligne CREDIT
Mid$(wCompte_CR, 1, 5) = W_REGCOM_5C
Mid$(meYCDODOS0.CDODOSBEN, 3, 5) = W_DOSBEN_5C
If W_REGCOM_5C = W_DOSBEN_5C And meYCDODOS0.CDODOSBER = " " Then
    XPrt.CurrentX = prtMinMarge + 4700: XPrt.Print "sur nos livres. ";
Else
    XPrt.CurrentX = prtMinMarge + 4700: XPrt.Print " " & meCDO_Courrier.IBAN;
End If

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Veuillez agréer, Messieurs, nos salutations distinguées."

XPrt.CurrentX = prtMinMarge + 5000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.FontBold = True
XPrt.Print paramSignature;
XPrt.FontBold = False

End Sub

Public Sub prtCDO_Courrier_UTI_NCONFORME_AR()
'---------------------------------------------------------
Dim X As String
Dim C As Integer
Dim W_Code_CRD As String
Dim MNT As String, MTV As String

Line2_Ecart = 10: prtCDO_Courrier_Trame (Line2_Ecart)

XPrt.CurrentY = prtCorpsY
XPrt.FontSize = 14: XPrt.FontBold = True: XPrt.FontUnderline = True
frmElpPrt.prtCentré prtMedX, "CREDIT DOCUMENTAIRE"
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
frmElpPrt.prtCentré prtMedX, "ACCUSE DE RECEPTION"

XPrt.FontSize = 11: XPrt.FontBold = False: XPrt.FontUnderline = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3

XPrt.CurrentX = prtMinMarge: XPrt.Print "  V/Référence :";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1700: XPrt.Print meCDO_Courrier.YCDOUTI0.CDOUTIRER;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Donneur d'ordre";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1700: XPrt.Print meCDO_Courrier.DON_Concat;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Bénéficiaire";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1700: XPrt.Print meCDO_Courrier.BEN_Concat;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Montant";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentX = prtMinMarge + 4100: XPrt.Print "Devise :";
XPrt.CurrentX = prtMinMarge + 7500: XPrt.Print "Validité :";
XPrt.FontBold = True
X = Trim(Format$(meYCDODOS0.CDODOSMON, "### ### ### ##0.00"))
XPrt.CurrentX = prtMinMarge + 1700
XPrt.Print X;
XPrt.CurrentX = prtMinMarge + 5000: XPrt.Print meYCDODOS0.CDODOSDEV;
XPrt.CurrentX = prtMinMarge + 8500: XPrt.Print wDate_Validité;
XPrt.FontBold = False

XPrt.FontSize = 11
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.CurrentX = prtMinMarge
XPrt.Print "Messieurs,";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Nous accusons réception de votre lettre nous remettant les documents de ";
XPrt.Print meYCDODOS0.CDODOSDEV & "  " & Trim(Format$(meCDO_Courrier.YCDOUTI0.CDOUTIMPA, "### ### ### ##0.00"));
XPrt.Print "  en réalisation du ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "crédit documentaire précité. ";

' >>>>>  Paragraphe DIFFERENT SI Crédit échu ou Crédit encore valide  <<<<<
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Conformément à vos instructions par télécopie du " & wDate_Remise_Util & ", nous adressons les documents à notre ";
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
If W_Validité_AMJ >= YBIATAB0_DATE_CPT_JS1 Then  ' CREDIT VALIDE
    XPrt.Print "correspondant pour accord de paiement du fait de(s) irrégularité(s) suivante(s) : ";
Else                                             ' CREDIT ECHU
    XPrt.Print "correspondant sur base d'encaissement du fait de(s) irrégularité(s) suivante(s) : ";
End If

' >>>>>  Paragraphe de(s) irrégularité(s) : récupérer directement du ZCDOIRR0 par rapport au dossier et no utilisation  <<<<<
For C = 1 To meCDO_Courrier.Irregul_Nb
    XPrt.FontSize = 8
    If C = 1 Then
        XPrt.CurrentX = prtMinMarge
        XPrt.CurrentY = XPrt.CurrentY + 200   '+ prtlineHeight
        ' XPrt.CurrentY = XPrt.CurrentY + 400   '+ prtlineHeight
        ' If W_Validité_AMJ < YBIATAB0_DATE_CPT_JS1 Then  ' CREDIT ECHU
        '     XPrt.FontBold = True
        '     XPrt.Print "- CREDIT ECHU ";
        '     XPrt.FontBold = False
        '     XPrt.CurrentX = prtMinMarge
        '     XPrt.CurrentY = XPrt.CurrentY + 200   '+ prtlineHeight
        ' End If
        XPrt.Print meCDO_Courrier.Irregul(1);
    Else
        XPrt.CurrentX = prtMinMarge
        XPrt.CurrentY = XPrt.CurrentY + 200   '+ prtlineHeight
        XPrt.Print meCDO_Courrier.Irregul(C);
    End If
Next C

XPrt.FontSize = 11

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "Nous reviendrons dès que possible. ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "MONTANT  DE  L'UTILISATION ";
XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meYCDODOS0.CDODOSDEV & "            " & Format$(meCDO_Courrier.YCDOUTI0.CDOUTIMPA, "### ### ### ##0.00");

'>>>>> SI Accusé de réception pour utilisation non conforme avec Frais charge bénéficiaire
If meCDO_Courrier.YCDOUTI0.CDOUTIBEC = "O" Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Nos frais et ceux de la banque émettrice seront déduits lors du règlement.";
End If

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Veuillez agréer, Messieurs, nos salutations distinguées."

XPrt.CurrentX = prtMinMarge + 5000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.FontBold = True
XPrt.Print paramSignature;
XPrt.FontBold = False

End Sub

Public Sub prtCDO_Courrier_UTI_CONFORME_BED1_Pdif_GB()
'---------------------------------------------------------
Dim X As String
Dim C As Integer
Dim W_Code_CRD As String
Dim MNT As String

' >>>>>>  Ce bordereau est adressé à la banque émettrice -agence-

Line2_Ecart = 8: prtCDO_Courrier_Trame (Line2_Ecart)

XPrt.CurrentY = prtCorpsY
XPrt.FontSize = 14: XPrt.FontBold = True: XPrt.FontUnderline = True
frmElpPrt.prtCentré prtMedX, "DOCUMENTARY  CREDIT"

XPrt.FontSize = 11: XPrt.FontBold = False: XPrt.FontUnderline = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2

XPrt.CurrentX = prtMinMarge: XPrt.Print "  Y/Reference :";
XPrt.CurrentX = prtMinMarge + 1550: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1750: XPrt.Print meYCDODOS0.CDODOSEXT;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Applicant";
XPrt.CurrentX = prtMinMarge + 1550: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1750: XPrt.Print meCDO_Courrier.DON_Concat;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Beneficiary";
XPrt.CurrentX = prtMinMarge + 1550: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1750: XPrt.Print meCDO_Courrier.BEN_Concat;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Amount";
XPrt.CurrentX = prtMinMarge + 1550: XPrt.Print ":";
XPrt.CurrentX = prtMinMarge + 4100: XPrt.Print "Currency :";
XPrt.CurrentX = prtMinMarge + 7500: XPrt.Print "Validity :";
XPrt.FontBold = True
X = Trim(Format$(meYCDODOS0.CDODOSMON, "### ### ### ##0.00"))
XPrt.CurrentX = prtMinMarge + 1750
XPrt.Print X;
XPrt.CurrentX = prtMinMarge + 5100: XPrt.Print meYCDODOS0.CDODOSDEV;
XPrt.CurrentX = prtMinMarge + 8500: XPrt.Print wDate_Validité_GB;
XPrt.FontBold = False

XPrt.FontSize = 11
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 4
XPrt.CurrentX = prtMinMarge
XPrt.Print "Dear sirs,";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "We have pleasure to send you herewith, to our entire discharge, the documents amouting to ";
XPrt.Print meYCDODOS0.CDODOSDEV & "  " & Trim(Format$(meCDO_Courrier.CrrGB_Mnt, "### ### ### ##0.00"));
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "under above mentioned credit. ";

' >>>>> Pavé d'éditions des différents types de documents <<<<<
XPrt.FontBold = True
XPrt.FontSize = 8
For C = 1 To meCDO_Courrier.Document_Nb
    If C = 1 Then
        XPrt.CurrentX = prtMinMarge + 5000
        XPrt.CurrentY = XPrt.CurrentY + 200  ' prtlineHeight
        XPrt.FontUnderline = True
        XPrt.Print "First set";
        XPrt.CurrentX = prtMinMarge + 6000
        XPrt.Print "Second set";
        XPrt.FontUnderline = False
    End If
    XPrt.CurrentX = prtMinMarge
    XPrt.CurrentY = XPrt.CurrentY + 200  ' prtlineHeight
    XPrt.Print meCDO_Courrier.Document(C);
    XPrt.CurrentX = prtMinMarge + 5200
    XPrt.Print meCDO_Courrier.Document_Jeu1(C);
    XPrt.CurrentX = prtMinMarge + 6300
    XPrt.Print meCDO_Courrier.Document_Jeu2(C);
Next C

XPrt.FontBold = False

' >>>>>  Les éléments concernant l'expédition   <<<<<
For C = 1 To meCDO_Courrier.Exp_Nb
    If C = 1 Then
        XPrt.CurrentX = prtMinMarge
        XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
        XPrt.FontSize = 11
        XPrt.Print "Covering shipment of : ";
        XPrt.FontSize = 8
        XPrt.CurrentX = prtMinMarge + 3000: XPrt.Print meCDO_Courrier.Exp(1);
    Else
        XPrt.CurrentX = prtMinMarge + 3000
        XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
        XPrt.Print meCDO_Courrier.Exp(C);
    End If
Next C

XPrt.FontSize = 11

' If meCDO_Courrier.Exp_Par <> Space Then
    XPrt.CurrentX = prtMinMarge
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.Print "By " & meCDO_Courrier.Exp_Par & " on " & meCDO_Courrier.Exp_Le & " from " & meCDO_Courrier.Exp_De & " to " & meCDO_Courrier.Exp_A;
' End If

' >>>>>  Nouveau paragraphe  <<<<<
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2

' >>>>> BORDEREAU = Toujours destinataire Donneur d'ordre donc Toujours comm. ligne DEBIT
'       BED EN ANGLAIS : JAMAIS DE COMMISSION
W_Code_CRD = "D"

If Trim(meCDO_Courrier.YCDOREG0_D.CDOREGPAS) = "" Then  'Ligne DEBIT pour repérer la banque de remboursement
    XPrt.Print "As per L/C terms, we shall debit your account in our records at maturity date " & wDate_Echeance_DB_GB & " as indicated in our ";
    XPrt.CurrentX = prtMinMarge
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.Print "today's SWIFT MT754, ie utilization amount : ";
    XPrt.Print meYCDODOS0.CDODOSDEV & "  " & Trim(Format$(meCDO_Courrier.YCDOUTI0.CDOUTIMPA, "### ### ### ##0.00"));
Else
    XPrt.Print "As per L/C terms, we shall reimburse ourselves with the account of " & meCDO_Courrier.CrrGB_BqRbt & " in our records at maturity ";
    XPrt.CurrentX = prtMinMarge
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.Print "date " & wDate_Echeance_DB_GB & " as indicated in our today's SWIFT MT754, ie utilization amount : ";
    XPrt.Print meYCDODOS0.CDODOSDEV & "  " & Trim(Format$(meCDO_Courrier.YCDOUTI0.CDOUTIMPA, "### ### ### ##0.00"));
End If
If Trim(meCDO_Courrier.CrrGB_Tx) <> "" Then   'Tx si dossier partiellemet confirmé (à saisir)
    XPrt.CurrentX = prtMinMarge
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.Print "( " & Trim(Format$(meCDO_Courrier.CrrGB_Tx, "### ##0.00")) & "% of documents value )";
End If

' >>>>>  Phrase concernant délai de présentation  <<<<<
If Trim(meCDO_Courrier.Delai(1)) <> "" Then
    XPrt.CurrentX = prtMinMarge
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
    XPrt.Print UCase$(meCDO_Courrier.Delai(1));
End If

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Kindly acknowledge receipt of this letter. ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Yours faithfully. ";

XPrt.CurrentX = prtMinMarge + 5000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.FontBold = True
XPrt.Print paramSignature;
XPrt.FontBold = False

End Sub

Public Sub prtCDO_Courrier_UTI_CONFORME_BED1_Pdif()
'---------------------------------------------------------
Dim X As String
Dim W_Cumul_Comm As Currency
Dim W_Cumul_TVA As Currency
Dim W_MNT_NET As Currency
Dim W_FRS As Currency
Dim C As Integer
Dim W_Code_CRD As String
Dim MNT As String, MTV As String

W_Cumul_Comm = 0
W_Cumul_TVA = 0
W_MNT_NET = 0

' >>>>>>  Ce bordereau est adressé à la banque émettrice -agence-

Line2_Ecart = 8: prtCDO_Courrier_Trame (Line2_Ecart)
'Call frmElpPrt.prtTrame(prtMinMarge - 50, prtCorpsY - 100, prtMaxMarge, prtCorpsY + prtlineHeight * 8 + 100, "B", 245)

XPrt.CurrentY = prtCorpsY
XPrt.FontSize = 14: XPrt.FontBold = True: XPrt.FontUnderline = True
frmElpPrt.prtCentré prtMedX, "CREDIT DOCUMENTAIRE"

XPrt.FontSize = 11: XPrt.FontBold = False: XPrt.FontUnderline = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
'Call frmElpPrt.prtTrame(prtMinMarge - 50, XPrt.CurrentY - 50, prtMaxMarge, XPrt.CurrentY + prtlineHeight * 3 + 50, , 245)

XPrt.CurrentX = prtMinMarge: XPrt.Print "  V/Référence :";
XPrt.CurrentX = prtMinMarge + 1550: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1750: XPrt.Print meYCDODOS0.CDODOSEXT;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Donneur d'ordre";
XPrt.CurrentX = prtMinMarge + 1550: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1750: XPrt.Print meCDO_Courrier.DON_Concat;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Bénéficiaire";
XPrt.CurrentX = prtMinMarge + 1550: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1750: XPrt.Print meCDO_Courrier.BEN_Concat;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Montant";
XPrt.CurrentX = prtMinMarge + 1550: XPrt.Print ":";
XPrt.CurrentX = prtMinMarge + 4100: XPrt.Print "Devise :";
XPrt.CurrentX = prtMinMarge + 7500: XPrt.Print "Validité :";
XPrt.FontBold = True
X = Trim(Format$(meYCDODOS0.CDODOSMON, "### ### ### ##0.00"))
XPrt.CurrentX = prtMinMarge + 1750
XPrt.Print X;
XPrt.CurrentX = prtMinMarge + 5000: XPrt.Print meYCDODOS0.CDODOSDEV;
XPrt.CurrentX = prtMinMarge + 8500: XPrt.Print wDate_Validité;
XPrt.FontBold = False

XPrt.FontSize = 11
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.CurrentX = prtMinMarge
XPrt.Print "Messieurs,";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Nous avons l'honneur de vous remettre ci-joints, à notre décharge, les documents de ";
' XPrt.Print meYCDODOS0.CDODOSDEV & "  " & Trim(Format$(meCDO_Courrier.YCDOUTI0.CDOUTIMPA, "### ### ### ##0.00")) & " levés ";
XPrt.Print meYCDODOS0.CDODOSDEV & "  " & Trim(Format$(meCDO_Courrier.YCDOUTI0.CDOUTIMON, "### ### ### ##0.00")) & " levés ";
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "en vertu de ce crédit documentaire. ";

' >>>>> Pavé d'éditions des différents types de documents <<<<<
XPrt.FontBold = True
XPrt.FontSize = 8
For C = 1 To meCDO_Courrier.Document_Nb
    If C = 1 Then
        XPrt.CurrentX = prtMinMarge + 5000
        XPrt.CurrentY = XPrt.CurrentY + 200  ' prtlineHeight
        XPrt.FontUnderline = True
        XPrt.Print "1er jeu";
        XPrt.CurrentX = prtMinMarge + 6000
        XPrt.Print "2ème jeu";
        XPrt.FontUnderline = False
    End If
    XPrt.CurrentX = prtMinMarge
    XPrt.CurrentY = XPrt.CurrentY + 200  ' prtlineHeight
    XPrt.Print meCDO_Courrier.Document(C);
    XPrt.CurrentX = prtMinMarge + 5200
    XPrt.Print meCDO_Courrier.Document_Jeu1(C);
    XPrt.CurrentX = prtMinMarge + 6300
    XPrt.Print meCDO_Courrier.Document_Jeu2(C);
Next C

' XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.FontBold = False

' >>>>>  Les éléments concernant l'expédition   <<<<<
For C = 1 To meCDO_Courrier.Exp_Nb
    If C = 1 Then
        XPrt.CurrentX = prtMinMarge
        XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
        XPrt.FontSize = 11
        XPrt.Print "Couvrant l'expédition de : ";
        XPrt.FontSize = 8
        XPrt.CurrentX = prtMinMarge + 3000: XPrt.Print meCDO_Courrier.Exp(1);
    Else
        XPrt.CurrentX = prtMinMarge + 3000
        XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
        XPrt.Print meCDO_Courrier.Exp(C);
    End If
Next C

XPrt.FontSize = 11

' If meCDO_Courrier.Exp_Par <> Space Then
    XPrt.CurrentX = prtMinMarge
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.Print "Par " & meCDO_Courrier.Exp_Par & " le " & meCDO_Courrier.Exp_Le & " de " & meCDO_Courrier.Exp_De & " à " & meCDO_Courrier.Exp_A;
' End If

' >>>>>  Nouveau paragraphe  <<<<<
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2

If meYCDODOS0.CDODOSNOT = "0011001" Then   'Si BEA
    XPrt.Print "Conformément à vos instructions, nous débitons le compte de votre siège d'Alger sur nos livres à l'échéance du ";
    XPrt.CurrentX = prtMinMarge
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.Print wDate_Echeance_DB & " selon notre MT754 de ce jour du montant de l'utilisation : ";
    ' XPrt.Print meYCDODOS0.CDODOSDEV & "  " & Trim(Format$(meCDO_Courrier.YCDOUTI0.CDOUTIMPA, "### ### ### ##0.00"));
    XPrt.Print meYCDODOS0.CDODOSDEV & "  " & Trim(Format$(meCDO_Courrier.YCDOUTI0.CDOUTIMON, "### ### ### ##0.00"));
Else
    ' Si USD et (BDL ou ABC ou CPA ou BARAKA) CR compte du Crédit Suisse
    If meYCDODOS0.CDODOSDEV = "USD" And (meYCDODOS0.CDODOSNOT = "0011077" Or meYCDODOS0.CDODOSNOT = "0011189" Or meYCDODOS0.CDODOSNOT = "0011074" Or meYCDODOS0.CDODOSNOT = "0011082") Then
        XPrt.Print "En remboursement de notre paiement, vous voudrez bien faire créditer le compte du CREDIT SUISSE Zurich ";
        XPrt.CurrentX = prtMinMarge
        XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
        XPrt.Print "(CRESCHZZ80A) auprès de la Bank of New York, à New York (IRVTUS3N), à l'échéance du " & wDate_Echeance_DB;
        XPrt.CurrentX = prtMinMarge
        XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
        XPrt.Print "selon notre MT754 de ce jour du montant de l'utilisation : ";
     '   XPrt.Print meYCDODOS0.CDODOSDEV & "  " & Trim(Format$(meCDO_Courrier.YCDOUTI0.CDOUTIMPA, "### ### ### ##0.00"));
        XPrt.Print meYCDODOS0.CDODOSDEV & "  " & Trim(Format$(meCDO_Courrier.YCDOUTI0.CDOUTIMON, "### ### ### ##0.00"));
    Else
        ' Si CHF et (BDL ou ABC ou CPA ou BARAKA)
        If meYCDODOS0.CDODOSDEV = "CHF" And (meYCDODOS0.CDODOSNOT = "0011077" Or meYCDODOS0.CDODOSNOT = "0011189" Or meYCDODOS0.CDODOSNOT = "0011074" Or meYCDODOS0.CDODOSNOT = "0011082") Then
            XPrt.Print "En remboursement de notre paiement, vous voudrez bien faire créditer notre compte auprès du CREDIT SUISSE ";
            XPrt.CurrentX = prtMinMarge
            XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
            XPrt.Print "Zurich, à l'échéance du " & wDate_Echeance_DB & " selon notre MT754 de ce jour du montant de l'utilisation : ";
          '  XPrt.Print meYCDODOS0.CDODOSDEV & "  " & Trim(Format$(meCDO_Courrier.YCDOUTI0.CDOUTIMPA, "### ### ### ##0.00"));
            XPrt.Print meYCDODOS0.CDODOSDEV & "  " & Trim(Format$(meCDO_Courrier.YCDOUTI0.CDOUTIMON, "### ### ### ##0.00"));
        Else
            'Cas normal + Restitution Provision si besoin...
            XPrt.Print "Conformément à vos instructions, nous débitons votre compte sur nos livres à l'échéance du " & wDate_Echeance_DB & " selon ";
            XPrt.CurrentX = prtMinMarge
            XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
            XPrt.Print "notre MT754 de ce jour du montant de l'utilisation : ";
          '  XPrt.Print meYCDODOS0.CDODOSDEV & "  " & Trim(Format$(meCDO_Courrier.YCDOUTI0.CDOUTIMPA, "### ### ### ##0.00"));
            XPrt.Print meYCDODOS0.CDODOSDEV & "  " & Trim(Format$(meCDO_Courrier.YCDOUTI0.CDOUTIMON, "### ### ### ##0.00"));
            If meCDO_Courrier.YCDOREG0_R_Nb <> 0 Then
                XPrt.CurrentX = prtMinMarge
                XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
                XPrt.Print "Par ailleurs, nous reversons au crédit de votre compte sur nos livres, la provision constituée sous même date ";
                XPrt.CurrentX = prtMinMarge
                XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
                XPrt.Print "de valeur. ";
            End If
        End If
    End If
End If


' >>>>> BORDEREAU = Toujours destinataire Donneur d'ordre donc Toujours comm. ligne DEBIT
W_Code_CRD = "D"

' La somme des commissions lues
For C = 1 To meCDO_Courrier.Com_Nb
    If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD Then
        W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    End If
Next C

' Phrase avec Frais et commissions
If W_Cumul_Comm <> 0 Then
    If meYCDODOS0.CDODOSNOT = "0011001" Then   'Si BEA
        XPrt.Print "  tenant compte des ";
        XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
        XPrt.CurrentX = prtMinMarge
        XPrt.Print "frais et commissions selon détail ci-après : ";
    Else
        ' Si USD et (BDL ou ABC ou CPA ou BARAKA) CR compte du Crédit Suisse
        If meYCDODOS0.CDODOSDEV = "USD" And (meYCDODOS0.CDODOSNOT = "0011077" Or meYCDODOS0.CDODOSNOT = "0011189" Or meYCDODOS0.CDODOSNOT = "0011074" Or meYCDODOS0.CDODOSNOT = "0011082") Then
            XPrt.Print "  tenant compte des frais et ";
            XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
            XPrt.CurrentX = prtMinMarge
            XPrt.Print "commissions selon détail ci-après et sous avis à nous-même : ";
        Else
            ' Si CHF et (BDL ou ABC ou CPA ou BARAKA)
            If meYCDODOS0.CDODOSDEV = "CHF" And (meYCDODOS0.CDODOSNOT = "0011077" Or meYCDODOS0.CDODOSNOT = "0011189" Or meYCDODOS0.CDODOSNOT = "0011074" Or meYCDODOS0.CDODOSNOT = "0011082") Then
                XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
                XPrt.CurrentX = prtMinMarge
                XPrt.Print "tenant compte des frais et commissions selon détail ci-après et sous avis à nous-même : ";
            Else
                ' Cas normal
                XPrt.Print "  tenant compte des frais et commissions ";
                XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
                XPrt.CurrentX = prtMinMarge
                XPrt.Print "selon détail ci-après : ";
            End If
        End If
    End If
    XPrt.CurrentX = prtMinMarge
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.FontUnderline = True
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print "Devise             Montant        TVA 19,60%";
    XPrt.FontUnderline = False
End If

' >>>>> Remettre à zéro W_Cumul_Comm avant de passer sur chaque ligne de commission
W_Cumul_Comm = 0
For C = 1 To meCDO_Courrier.Com_Nb

' Commission ECNF...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "ECNF  " And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission de confirmation ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission ECSIL...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "ECSIL " And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission de confirmation silencieuse ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission ENOTIF...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "ENOTIF" And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission de notification ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EPDIF...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EPDIF " And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission d'engagement paiement différé ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EPDIFD...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EPDIFD" And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission d'engagement paiement différé ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EACCEP...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EACCEP" And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission d'acceptation ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EACCED...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EACCED" And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission d'acceptation ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission ELVD...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "ELVD  " And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission de levée de documents ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission ERFA...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "ERFA  " And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Frais de ports et telex ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EDOCIR...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EDOCIR" And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission de documents irréguliers ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EACD...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EACD  " And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Autre commission documentaire ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EMODIF...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EMODIF" And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission de modification ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EFBEMT... Toujours en MOINS pour la banque émettrice au bordereau adressé au Donneur d'Ordre
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EFBEMT" And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Frais banque émettrice ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm - meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA - meCDO_Courrier.CDOCOMMTV(C)
End If

' >>>>> Fin de boucle sur lignes de commissions
Next C

' Frais de notre correspondant : Résultat d'une soustraction entre montant à payer et montant de l'utilisation
W_FRS = meCDO_Courrier.YCDOUTI0.CDOUTIMPA - meCDO_Courrier.YCDOUTI0.CDOUTIMON
If W_FRS <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Frais de notre correspondant ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(1);
    MNT = Format$(W_FRS, "### ### ### ##0.00")
    MTV = 0
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + W_FRS
    W_Cumul_TVA = W_Cumul_TVA + 0
End If

If W_Cumul_Comm <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
    XPrt.CurrentX = prtMinMarge
    XPrt.CurrentX = prtMinMarge + 3000: XPrt.Print "Montant  total ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(1);
    MNT = Format$(W_Cumul_Comm, "### ### ### ##0.00")
    MTV = Format$(W_Cumul_TVA, "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
End If

W_MNT_NET = meCDO_Courrier.YCDOUTI0.CDOUTIMON + W_Cumul_Comm + W_Cumul_TVA
' If W_MNT_NET <> 0 Then
If W_Cumul_Comm <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
    XPrt.CurrentX = prtMinMarge
    Call frmElpPrt.prtTrame(prtMinMarge + 3000 - 50, XPrt.CurrentY - 100, prtMaxMarge - 2000, XPrt.CurrentY + prtlineHeight + 100, 245)
    XPrt.CurrentX = prtMinMarge + 3000: XPrt.Print "Montant  total  net ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(1);
    MNT = Format$(W_MNT_NET, "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
End If

' Si USD et (BDL ou ABC ou CPA ou BARAKA) CR compte du Crédit Suisse
' >>>>>>  Le 18/01/2005 Mise en commentaire de la phrase car levée d'embargo LIBYE
'If meYCDODOS0.CDODOSDEV = "USD" And (meYCDODOS0.CDODOSNOT = "0011077" Or meYCDODOS0.CDODOSNOT = "0011189" Or meYCDODOS0.CDODOSNOT = "0011074" Or meYCDODOS0.CDODOSNOT = "0011082") Then
'    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
'    Call frmElpPrt.prtTrame(prtMinMarge, XPrt.CurrentY, prtMaxMarge, XPrt.CurrentY + prtlineHeight, " ", 235)
'    XPrt.CurrentX = prtMinMarge
'    XPrt.CurrentY = XPrt.CurrentY + 50
'    XPrt.FontSize = 7
'    XPrt.FontBold = True
'    XPrt.Print "NOUS VOUS RAPPELONS QU'EN AUCUN CAS, LE NOM DE NOTRE BANQUE NE DOIT ÊTRE MENTIONNE SUR VOTRE AVIS DE COUVERTURE. ";
'    XPrt.FontSize = 11
'    XPrt.FontBold = False
'End If

' >>>>>  Phrase concernant délai de présentation  <<<<<
If Trim(meCDO_Courrier.Delai(1)) <> "" Then
    XPrt.CurrentX = prtMinMarge
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
    XPrt.Print UCase$(meCDO_Courrier.Delai(1));
End If

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.Print "Veuillez nous accuser réception de notre envoi et agréer, Messieurs, nos salutations distinguées. ";

XPrt.CurrentX = prtMinMarge + 5000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.FontBold = True
XPrt.Print paramSignature;
XPrt.FontBold = False

End Sub

Public Sub prtCDO_Courrier_UTI_CONFORME_BED2_Avue_GB()
'---------------------------------------------------------
Dim X As String
Dim C As Integer
Dim W_Code_CRD As String
Dim MNT As String

' >>>>>>  Ce bordereau est adressé à la banque émettrice -agence-

Line2_Ecart = 8: prtCDO_Courrier_Trame (Line2_Ecart)

XPrt.CurrentY = prtCorpsY
XPrt.FontSize = 14: XPrt.FontBold = True: XPrt.FontUnderline = True
frmElpPrt.prtCentré prtMedX, "DOCUMENTARY  CREDIT"

XPrt.FontSize = 11: XPrt.FontBold = False: XPrt.FontUnderline = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2

XPrt.CurrentX = prtMinMarge: XPrt.Print "  Y/Reference :";
XPrt.CurrentX = prtMinMarge + 1550: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1750: XPrt.Print meYCDODOS0.CDODOSEXT;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Applicant";
XPrt.CurrentX = prtMinMarge + 1550: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1750: XPrt.Print meCDO_Courrier.DON_Concat;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Beneficiary";
XPrt.CurrentX = prtMinMarge + 1550: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1750: XPrt.Print meCDO_Courrier.BEN_Concat;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Amount";
XPrt.CurrentX = prtMinMarge + 1550: XPrt.Print ":";
XPrt.CurrentX = prtMinMarge + 4100: XPrt.Print "Currency :";
XPrt.CurrentX = prtMinMarge + 7500: XPrt.Print "Validity :";
XPrt.FontBold = True
X = Trim(Format$(meYCDODOS0.CDODOSMON, "### ### ### ##0.00"))
XPrt.CurrentX = prtMinMarge + 1750
XPrt.Print X;
XPrt.CurrentX = prtMinMarge + 5100: XPrt.Print meYCDODOS0.CDODOSDEV;
XPrt.CurrentX = prtMinMarge + 8500: XPrt.Print wDate_Validité_GB;
XPrt.FontBold = False

XPrt.FontSize = 11
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 4
XPrt.CurrentX = prtMinMarge
XPrt.Print "Dear sirs,";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "We have pleasure to send you herewith, to our entire discharge, the documents amouting to ";
XPrt.Print meYCDODOS0.CDODOSDEV & "  " & Trim(Format$(meCDO_Courrier.CrrGB_Mnt, "### ### ### ##0.00"));
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "under above mentioned credit. ";

' >>>>> Pavé d'éditions des différents types de documents <<<<<
XPrt.FontBold = True
XPrt.FontSize = 8
For C = 1 To meCDO_Courrier.Document_Nb
    If C = 1 Then
        XPrt.CurrentX = prtMinMarge + 5000
        XPrt.CurrentY = XPrt.CurrentY + 200  ' prtlineHeight
        XPrt.FontUnderline = True
        XPrt.Print "First set";
        XPrt.CurrentX = prtMinMarge + 6000
        XPrt.Print "Second set";
        XPrt.FontUnderline = False
    End If
    XPrt.CurrentX = prtMinMarge
    XPrt.CurrentY = XPrt.CurrentY + 200  ' prtlineHeight
    XPrt.Print meCDO_Courrier.Document(C);
    XPrt.CurrentX = prtMinMarge + 5200
    XPrt.Print meCDO_Courrier.Document_Jeu1(C);
    XPrt.CurrentX = prtMinMarge + 6300
    XPrt.Print meCDO_Courrier.Document_Jeu2(C);
Next C

XPrt.FontBold = False

' >>>>>  Les éléments concernant l'expédition   <<<<<
For C = 1 To meCDO_Courrier.Exp_Nb
    If C = 1 Then
        XPrt.CurrentX = prtMinMarge
        XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
        XPrt.FontSize = 11
        XPrt.Print "Covering shipment of : ";
        XPrt.FontSize = 8
        XPrt.CurrentX = prtMinMarge + 3000: XPrt.Print meCDO_Courrier.Exp(1);
    Else
        XPrt.CurrentX = prtMinMarge + 3000
        XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
        XPrt.Print meCDO_Courrier.Exp(C);
    End If
Next C

XPrt.FontSize = 11

' If meCDO_Courrier.Exp_Par <> Space Then
    XPrt.CurrentX = prtMinMarge
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.Print "By " & meCDO_Courrier.Exp_Par & " on " & meCDO_Courrier.Exp_Le & " from " & meCDO_Courrier.Exp_De & " to " & meCDO_Courrier.Exp_A;
' End If

' >>>>>  Nouveau paragraphe  <<<<<
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2

' >>>>> BORDEREAU = Toujours destinataire Donneur d'ordre donc Toujours comm. ligne DEBIT
'       BED EN ANGLAIS : JAMAIS DE COMMISSION
W_Code_CRD = "D"

If Trim(meCDO_Courrier.YCDOREG0_D.CDOREGPAS) = "" Then  'Ligne DEBIT pour repérer la banque de remboursement
    XPrt.Print "As per L/C terms, we shall debit your account in our records value " & wDate_Valeur_DB_GB & " as indicated in our ";
    XPrt.CurrentX = prtMinMarge
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.Print "today's SWIFT MT754, ie utilization amount : ";
    XPrt.Print meYCDODOS0.CDODOSDEV & "  " & Trim(Format$(meCDO_Courrier.YCDOUTI0.CDOUTIMPA, "### ### ### ##0.00"));
Else
    XPrt.Print "As per L/C terms, we shall reimburse ourselves with the account of " & meCDO_Courrier.CrrGB_BqRbt & " in our records value ";
    XPrt.CurrentX = prtMinMarge
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.Print wDate_Valeur_DB_GB & " as indicated in our today's SWIFT MT754, ie utilization amount : ";
    XPrt.Print meYCDODOS0.CDODOSDEV & "  " & Trim(Format$(meCDO_Courrier.YCDOUTI0.CDOUTIMPA, "### ### ### ##0.00"));
End If
If Trim(meCDO_Courrier.CrrGB_Tx) <> "" Then   'Tx si dossier partiellemet confirmé (à saisir)
    XPrt.CurrentX = prtMinMarge
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.Print "( " & Trim(Format$(meCDO_Courrier.CrrGB_Tx, "### ##0.00")) & "% of documents value )";
End If

' >>>>>  Phrase concernant délai de présentation  <<<<<
If Trim(meCDO_Courrier.Delai(1)) <> "" Then
    XPrt.CurrentX = prtMinMarge
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
    XPrt.Print UCase$(meCDO_Courrier.Delai(1));
End If

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Kindly acknowledge receipt of this letter. ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Yours faithfully. ";

XPrt.CurrentX = prtMinMarge + 5000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.FontBold = True
XPrt.Print paramSignature;
XPrt.FontBold = False

End Sub
Public Sub prtCDO_Courrier_UTI_CONFORME_BED2_Avue()
'---------------------------------------------------------
Dim X As String
Dim W_Cumul_Comm As Currency
Dim W_Cumul_TVA As Currency
Dim W_MNT_NET As Currency
Dim W_FRS As Currency
Dim C As Integer
Dim W_Code_CRD As String
Dim MNT As String, MTV As String

W_Cumul_Comm = 0
W_Cumul_TVA = 0
W_MNT_NET = 0

' >>>>>>  Ce bordereau est adressé à la banque émettrice -agence-

Line2_Ecart = 8: prtCDO_Courrier_Trame (Line2_Ecart)

XPrt.CurrentY = prtCorpsY
XPrt.FontSize = 14: XPrt.FontBold = True: XPrt.FontUnderline = True
frmElpPrt.prtCentré prtMedX, "CREDIT DOCUMENTAIRE"

XPrt.FontSize = 11: XPrt.FontBold = False: XPrt.FontUnderline = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2

XPrt.CurrentX = prtMinMarge: XPrt.Print "  V/Référence  ";
XPrt.CurrentX = prtMinMarge + 1550: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1750: XPrt.Print meYCDODOS0.CDODOSEXT;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Donneur d'ordre";
XPrt.CurrentX = prtMinMarge + 1550: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1750: XPrt.Print meCDO_Courrier.DON_Concat;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Bénéficiaire";
XPrt.CurrentX = prtMinMarge + 1550: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1750: XPrt.Print meCDO_Courrier.BEN_Concat;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Montant";
XPrt.CurrentX = prtMinMarge + 1550: XPrt.Print ":";
XPrt.CurrentX = prtMinMarge + 4100: XPrt.Print "Devise :";
XPrt.CurrentX = prtMinMarge + 7500: XPrt.Print "Validité :";
XPrt.FontBold = True
X = Trim(Format$(meYCDODOS0.CDODOSMON, "### ### ### ##0.00"))
XPrt.CurrentX = prtMinMarge + 1750
XPrt.Print X;
XPrt.CurrentX = prtMinMarge + 5000: XPrt.Print meYCDODOS0.CDODOSDEV;
XPrt.CurrentX = prtMinMarge + 8500: XPrt.Print wDate_Validité;
XPrt.FontBold = False

XPrt.FontSize = 11
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.CurrentX = prtMinMarge
XPrt.Print "Messieurs,";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Nous avons l'honneur de vous remettre ci-joints, à notre décharge, les documents de ";
' XPrt.Print meYCDODOS0.CDODOSDEV & "  " & Trim(Format$(meCDO_Courrier.YCDOUTI0.CDOUTIMPA, "### ### ### ##0.00")) & " levés ";
XPrt.Print meYCDODOS0.CDODOSDEV & "  " & Trim(Format$(meCDO_Courrier.YCDOUTI0.CDOUTIMON, "### ### ### ##0.00")) & " levés ";
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "en vertu de ce crédit documentaire. ";

' >>>>> Pavé d'éditions des différents types de documents <<<<<

XPrt.FontBold = True
XPrt.FontSize = 8
For C = 1 To meCDO_Courrier.Document_Nb
    If C = 1 Then
        XPrt.CurrentX = prtMinMarge + 5000
        XPrt.CurrentY = XPrt.CurrentY + 200  ' prtlineHeight
        XPrt.FontUnderline = True
        XPrt.Print "1er jeu";
        XPrt.CurrentX = prtMinMarge + 6000
        XPrt.Print "2ème jeu";
        XPrt.FontUnderline = False
    End If
    XPrt.CurrentX = prtMinMarge
    XPrt.CurrentY = XPrt.CurrentY + 200  ' prtlineHeight
    XPrt.Print meCDO_Courrier.Document(C);
    XPrt.CurrentX = prtMinMarge + 5200
    XPrt.Print meCDO_Courrier.Document_Jeu1(C);
    XPrt.CurrentX = prtMinMarge + 6300
    XPrt.Print meCDO_Courrier.Document_Jeu2(C);
Next C

' XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.FontBold = False

' >>>>>  Les éléments concernant l'expédition   <<<<<
For C = 1 To meCDO_Courrier.Exp_Nb
    If C = 1 Then
        XPrt.CurrentX = prtMinMarge
        XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
        XPrt.FontSize = 11
        XPrt.Print "Couvrant l'expédition de : ";
        XPrt.FontSize = 8
        XPrt.CurrentX = prtMinMarge + 3000: XPrt.Print meCDO_Courrier.Exp(1);
    Else
        XPrt.CurrentX = prtMinMarge + 3000
        XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
        XPrt.Print meCDO_Courrier.Exp(C);
    End If
Next C

XPrt.FontSize = 11

' If meCDO_Courrier.Exp_Par <> Space Then
    XPrt.CurrentX = prtMinMarge
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.Print "Par " & meCDO_Courrier.Exp_Par & " le " & meCDO_Courrier.Exp_Le & " de " & meCDO_Courrier.Exp_De & " à " & meCDO_Courrier.Exp_A;
' End If

' >>>>>  Nouveau paragraphe  <<<<<
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
If meYCDODOS0.CDODOSNOT = "0011001" Then  'Si BEA
    XPrt.Print "Conformément à vos instructions, nous débitons le compte de votre siège d'Alger sur nos livres valeur " & wDate_Valeur_DB;
    XPrt.CurrentX = prtMinMarge
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.Print "selon notre MT754 de ce jour du montant de l'utilisation : ";
    ' XPrt.Print meYCDODOS0.CDODOSDEV & "  " & Trim(Format$(meCDO_Courrier.YCDOUTI0.CDOUTIMPA, "### ### ### ##0.00"));
    XPrt.Print meYCDODOS0.CDODOSDEV & "  " & Trim(Format$(meCDO_Courrier.YCDOUTI0.CDOUTIMON, "### ### ### ##0.00"));
Else
    ' Si USD et (BDL ou ABC ou CPA ou BARAKA) CR compte du Crédit Suisse
    If meYCDODOS0.CDODOSDEV = "USD" And (meYCDODOS0.CDODOSNOT = "0011077" Or meYCDODOS0.CDODOSNOT = "0011189" Or meYCDODOS0.CDODOSNOT = "0011074" Or meYCDODOS0.CDODOSNOT = "0011082") Then
        XPrt.Print "En remboursement de notre paiement, vous voudrez bien faire créditer le compte du CREDIT SUISSE Zurich ";
        XPrt.CurrentX = prtMinMarge
        XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
        XPrt.Print "(CRESCHZZ80A) auprès de la Bank of New York, à New York (IRVTUS3N), valeur " & wDate_Valeur_DB & " selon notre ";
        XPrt.CurrentX = prtMinMarge
        XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
        XPrt.Print "MT754 de ce jour du montant de l'utilisation : ";
        ' XPrt.Print meYCDODOS0.CDODOSDEV & "  " & Trim(Format$(meCDO_Courrier.YCDOUTI0.CDOUTIMPA, "### ### ### ##0.00"));
        XPrt.Print meYCDODOS0.CDODOSDEV & "  " & Trim(Format$(meCDO_Courrier.YCDOUTI0.CDOUTIMON, "### ### ### ##0.00"));
    Else
        ' Si CHF et (BDL ou ABC ou CPA ou BARAKA)
        If meYCDODOS0.CDODOSDEV = "CHF" And (meYCDODOS0.CDODOSNOT = "0011077" Or meYCDODOS0.CDODOSNOT = "0011189" Or meYCDODOS0.CDODOSNOT = "0011074" Or meYCDODOS0.CDODOSNOT = "0011082") Then
            XPrt.Print "En remboursement de notre paiement, vous voudrez bien faire créditer notre compte auprès du CREDIT SUISSE ";
            XPrt.CurrentX = prtMinMarge
            XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
            XPrt.Print "Zurich, valeur " & wDate_Valeur_DB & " selon notre MT754 de ce jour du montant de l'utilisation : ";
            ' XPrt.Print meYCDODOS0.CDODOSDEV & "  " & Trim(Format$(meCDO_Courrier.YCDOUTI0.CDOUTIMPA, "### ### ### ##0.00"));
            XPrt.Print meYCDODOS0.CDODOSDEV & "  " & Trim(Format$(meCDO_Courrier.YCDOUTI0.CDOUTIMON, "### ### ### ##0.00"));
        Else
            'Cas normal + Restitution Provision si besoin...
            XPrt.Print "Conformément à vos instructions, nous débitons votre compte sur nos livres valeur " & wDate_Valeur_DB & " selon notre ";
            XPrt.CurrentX = prtMinMarge
            XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
            XPrt.Print "MT754 de ce jour du montant de l'utilisation : ";
            ' XPrt.Print meYCDODOS0.CDODOSDEV & "  " & Trim(Format$(meCDO_Courrier.YCDOUTI0.CDOUTIMPA, "### ### ### ##0.00"));
            XPrt.Print meYCDODOS0.CDODOSDEV & "  " & Trim(Format$(meCDO_Courrier.YCDOUTI0.CDOUTIMON, "### ### ### ##0.00"));
            If meCDO_Courrier.YCDOREG0_R_Nb <> 0 Then
                XPrt.CurrentX = prtMinMarge
                XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
                XPrt.Print "Par ailleurs, nous reversons au crédit de votre compte sur nos livres, la provision constituée sous même date ";
                XPrt.CurrentX = prtMinMarge
                XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
                XPrt.Print "de valeur. ";
            End If
        End If
    End If
End If

' Phrase sur provision


' >>>>> BORDEREAU = Toujours destinataire Donneur d'ordre donc Toujours comm. ligne DEBIT
W_Code_CRD = "D"

' La somme des commissions lues
For C = 1 To meCDO_Courrier.Com_Nb
    If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD Then
        W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    End If
Next C

If W_Cumul_Comm <> 0 Then
    ' Si USD et (BDL ou ABC ou CPA ou BARAKA) CR compte du Crédit Suisse
    If meYCDODOS0.CDODOSDEV = "USD" And (meYCDODOS0.CDODOSNOT = "0011077" Or meYCDODOS0.CDODOSNOT = "0011189" Or meYCDODOS0.CDODOSNOT = "0011074" Or meYCDODOS0.CDODOSNOT = "0011082") Then
        XPrt.Print "  tenant compte des frais et commissions ";
        XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
        XPrt.CurrentX = prtMinMarge
        XPrt.Print "selon détail ci-après et sous avis à nous-même : ";
    Else
        ' Si CHF et (BDL ou ABC ou CPA ou BARAKA)
        If meYCDODOS0.CDODOSDEV = "CHF" And (meYCDODOS0.CDODOSNOT = "0011077" Or meYCDODOS0.CDODOSNOT = "0011189" Or meYCDODOS0.CDODOSNOT = "0011074" Or meYCDODOS0.CDODOSNOT = "0011082") Then
            XPrt.Print "  tenant ";
            XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
            XPrt.CurrentX = prtMinMarge
            XPrt.Print "compte des frais et commissions selon détail ci-après et sous avis à nous-même : ";
        Else
            ' Cas normal et CDODOSNOT="0011001" BEA
            XPrt.Print "  tenant compte des frais et commissions ";
            XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
            XPrt.CurrentX = prtMinMarge
            XPrt.Print "selon détail ci-après : ";
        End If
    End If
    XPrt.CurrentX = prtMinMarge
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.FontUnderline = True
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print "Devise             Montant        TVA 19,60%";
    XPrt.FontUnderline = False
End If

' >>>>> Remettre à zéro W_Cumul_Comm avant de passer sur chaque ligne de commission
W_Cumul_Comm = 0
For C = 1 To meCDO_Courrier.Com_Nb

' Commission ECNF...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "ECNF  " And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission de confirmation ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission ECSIL...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "ECSIL " And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission de confirmation silencieuse ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission ENOTIF...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "ENOTIF" And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission de notification ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EPDIF...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EPDIF " And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission d'engagement paiement différé ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EPDIFD...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EPDIFD" And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission d'engagement paiement différé ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EACCEP...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EACCEP" And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission d'acceptation ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EACCED...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EACCED" And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission d'acceptation ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission ELVD...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "ELVD  " And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission de levée de documents ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission ERFA...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "ERFA  " And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Frais de ports et telex ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EDOCIR...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EDOCIR" And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission de documents irréguliers ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EACD...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EACD  " And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Autre commission documentaire ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EMODIF...
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EMODIF" And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Commission de modification ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA + meCDO_Courrier.CDOCOMMTV(C)
End If

' Commission EFBEMT... Toujours en MOINS pour la banque émettrice au bordereau adressé au Donneur d'Ordre
If meCDO_Courrier.CDOREGCRD(C) = W_Code_CRD And meCDO_Courrier.CDOCOMCOM(C) = "EFBEMT" And meCDO_Courrier.CDOCOMMON(C) <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Frais banque émettrice ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(C);
    MNT = Format$(meCDO_Courrier.CDOCOMMON(C), "### ### ### ##0.00")
    MTV = Format$(meCDO_Courrier.CDOCOMMTV(C), "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm - meCDO_Courrier.CDOCOMMON(C)
    W_Cumul_TVA = W_Cumul_TVA - meCDO_Courrier.CDOCOMMTV(C)
End If

' >>>>> Fin de boucle sur lignes de commissions
Next C

' Frais de notre correspondant : Résultat d'une soustraction entre montant à payer et montant de l'utilisation
W_FRS = meCDO_Courrier.YCDOUTI0.CDOUTIMPA - meCDO_Courrier.YCDOUTI0.CDOUTIMON
If W_FRS <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinMarge
    XPrt.Print "Frais de notre correspondant ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(1);
    MNT = Format$(W_FRS, "### ### ### ##0.00")
    MTV = 0
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
    W_Cumul_Comm = W_Cumul_Comm + W_FRS
    W_Cumul_TVA = W_Cumul_TVA + 0
End If

If W_Cumul_Comm <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
    XPrt.CurrentX = prtMinMarge
    XPrt.CurrentX = prtMinMarge + 3000: XPrt.Print "Montant  total ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(1);
    MNT = Format$(W_Cumul_Comm, "### ### ### ##0.00")
    MTV = Format$(W_Cumul_TVA, "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
    XPrt.CurrentX = prtMaxMarge - 500 - XPrt.TextWidth(MTV): XPrt.Print MTV;
End If

W_MNT_NET = meCDO_Courrier.YCDOUTI0.CDOUTIMON + W_Cumul_Comm + W_Cumul_TVA
' If W_MNT_NET <> 0 Then
If W_Cumul_Comm <> 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
    XPrt.CurrentX = prtMinMarge
    Call frmElpPrt.prtTrame(prtMinMarge + 3000 - 50, XPrt.CurrentY - 100, prtMaxMarge - 2000, XPrt.CurrentY + prtlineHeight + 100, 245)
    XPrt.CurrentX = prtMinMarge + 3000: XPrt.Print "Montant  total  net ";
    XPrt.CurrentX = prtMinMarge + 5800: XPrt.Print meCDO_Courrier.CDOCOMDEV(1);
    MNT = Format$(W_MNT_NET, "### ### ### ##0.00")
    XPrt.CurrentX = prtMaxMarge - 2000 - XPrt.TextWidth(MNT): XPrt.Print MNT;
End If

' Si USD et (BDL ou ABC ou CPA ou BARAKA) CR compte du Crédit Suisse
' >>>>>>  Le 18/01/2005 Mise en commentaire de la phrase car levée d'embargo LIBYE
'If meYCDODOS0.CDODOSDEV = "USD" And (meYCDODOS0.CDODOSNOT = "0011077" Or meYCDODOS0.CDODOSNOT = "0011189" Or meYCDODOS0.CDODOSNOT = "0011074" Or meYCDODOS0.CDODOSNOT = "0011082") Then
'    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
'    Call frmElpPrt.prtTrame(prtMinMarge, XPrt.CurrentY, prtMaxMarge, XPrt.CurrentY + prtlineHeight, " ", 235)
'    XPrt.CurrentX = prtMinMarge
'    XPrt.CurrentY = XPrt.CurrentY + 50
'    XPrt.FontSize = 7
'    XPrt.FontBold = True
'    XPrt.Print "NOUS VOUS RAPPELONS QU'EN AUCUN CAS, LE NOM DE NOTRE BANQUE NE DOIT ÊTRE MENTIONNE SUR VOTRE AVIS DE COUVERTURE. ";
'    XPrt.FontSize = 11
'    XPrt.FontBold = False
'End If

' >>>>>  Phrase concernant délai de présentation  <<<<<
If Trim(meCDO_Courrier.Delai(1)) <> "" Then
    XPrt.CurrentX = prtMinMarge
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
    XPrt.Print UCase$(meCDO_Courrier.Delai(1));
End If

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.Print "Veuillez nous accuser réception de notre envoi et agréer, Messieurs, nos salutations distinguées. ";

XPrt.CurrentX = prtMinMarge + 5000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.FontBold = True
XPrt.Print paramSignature;
XPrt.FontBold = False

End Sub

Public Sub prtCDO_Courrier_UTI_NCONFORME_BED_Pdif()
'---------------------------------------------------------
Dim X As String
Dim C As Integer
Dim W_Code_CRD As String
Dim MNT As String, MTV As String

' >>>>>>  Ce bordereau est adressé à la banque émettrice -agence-

Line2_Ecart = 8: prtCDO_Courrier_Trame (Line2_Ecart)

XPrt.CurrentY = prtCorpsY
XPrt.FontSize = 14: XPrt.FontBold = True: XPrt.FontUnderline = True
frmElpPrt.prtCentré prtMedX, "CREDIT DOCUMENTAIRE"

XPrt.FontSize = 11: XPrt.FontBold = False: XPrt.FontUnderline = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2

XPrt.CurrentX = prtMinMarge: XPrt.Print "  V/Référence :";
XPrt.CurrentX = prtMinMarge + 1550: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1750: XPrt.Print meYCDODOS0.CDODOSEXT;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Donneur d'ordre";
XPrt.CurrentX = prtMinMarge + 1550: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1750: XPrt.Print meCDO_Courrier.DON_Concat;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Bénéficiaire";
XPrt.CurrentX = prtMinMarge + 1550: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1750: XPrt.Print meCDO_Courrier.BEN_Concat;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Montant";
XPrt.CurrentX = prtMinMarge + 1550: XPrt.Print ":";
XPrt.CurrentX = prtMinMarge + 4100: XPrt.Print "Devise :";
XPrt.CurrentX = prtMinMarge + 7500: XPrt.Print "Validité :";
XPrt.FontBold = True
X = Trim(Format$(meYCDODOS0.CDODOSMON, "### ### ### ##0.00"))
XPrt.CurrentX = prtMinMarge + 1750
XPrt.Print X;
XPrt.CurrentX = prtMinMarge + 5000: XPrt.Print meYCDODOS0.CDODOSDEV;
XPrt.CurrentX = prtMinMarge + 8500: XPrt.Print wDate_Validité;
XPrt.FontBold = False

XPrt.FontSize = 11
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.CurrentX = prtMinMarge
XPrt.Print "Messieurs,";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Nous avons l'honneur de vous remettre ci-joints, à notre décharge, les documents de ";
XPrt.Print meYCDODOS0.CDODOSDEV & "  " & Trim(Format$(meCDO_Courrier.YCDOUTI0.CDOUTIMPA, "### ### ### ##0.00")) & " présentés ";
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "en vertu de ce crédit documentaire. ";

' >>>>> Pavé d'éditions des différents types de documents <<<<<
XPrt.FontBold = True
XPrt.FontSize = 8
For C = 1 To meCDO_Courrier.Document_Nb
    If C = 1 Then
        XPrt.CurrentX = prtMinMarge + 5000
        XPrt.CurrentY = XPrt.CurrentY + 200  ' prtlineHeight
        XPrt.FontUnderline = True
        XPrt.Print "1er jeu";
        XPrt.CurrentX = prtMinMarge + 6000
        XPrt.Print "2ème jeu";
        XPrt.FontUnderline = False
    End If
    XPrt.CurrentX = prtMinMarge
    XPrt.CurrentY = XPrt.CurrentY + 200  ' prtlineHeight
    XPrt.Print meCDO_Courrier.Document(C);
    XPrt.CurrentX = prtMinMarge + 5200
    XPrt.Print meCDO_Courrier.Document_Jeu1(C);
    XPrt.CurrentX = prtMinMarge + 6300
    XPrt.Print meCDO_Courrier.Document_Jeu2(C);
Next C

' XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.FontBold = False

' >>>>>  Les éléments concernant l'expédition   <<<<<
For C = 1 To meCDO_Courrier.Exp_Nb
    If C = 1 Then
        XPrt.CurrentX = prtMinMarge
        XPrt.CurrentY = XPrt.CurrentY + 200   '+ prtlineHeight
        XPrt.FontSize = 11
        XPrt.Print "Couvrant l'expédition de : ";
        XPrt.FontSize = 8
        XPrt.CurrentX = prtMinMarge + 3000
        XPrt.Print meCDO_Courrier.Exp(1);
    Else
        XPrt.CurrentX = prtMinMarge + 3000
        XPrt.CurrentY = XPrt.CurrentY + 200   '+ prtlineHeight
        XPrt.Print meCDO_Courrier.Exp(C);
    End If
Next C

XPrt.FontSize = 11

' If meCDO_Courrier.Exp_Par <> Space Then
    XPrt.CurrentX = prtMinMarge
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.Print "Par " & meCDO_Courrier.Exp_Par & " le " & meCDO_Courrier.Exp_Le & " de " & meCDO_Courrier.Exp_De & " à " & meCDO_Courrier.Exp_A;
' End If

' >>>>>  Paragraphe DIFFERENT SI Crédit échu ou Crédit encore valide  <<<<<
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
If W_Validité_AMJ >= YBIATAB0_DATE_CPT_JS1 Then  ' CREDIT VALIDE
    XPrt.Print "Nous vous adressons ces documents pour accord de paiement en raison de(s) irrégularité(s) suivante(s) :";
Else                                ' CREDIT ECHU
    XPrt.Print "Nous vous adressons ces documents sur base d'encaissement en raison de(s) irrégularité(s) suivante(s) :";
    ' XPrt.CurrentX = prtMinMarge
    ' XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    ' XPrt.FontBold = True:     XPrt.FontSize = 8
    ' XPrt.Print "- CREDIT ECHU ";
    ' XPrt.FontBold = False
End If

' >>>>>  Paragraphe de(s) irrégularité(s) : récupérer directement du ZCDOIRR0 par rapport au dossier et no utilisation  <<<<<
For C = 1 To meCDO_Courrier.Irregul_Nb
    XPrt.FontSize = 8
    XPrt.CurrentX = prtMinMarge
    XPrt.CurrentY = XPrt.CurrentY + 200   '+ prtlineHeight
    XPrt.Print meCDO_Courrier.Irregul(C);
Next C

XPrt.FontSize = 11

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1
If Trim(meCDO_Courrier.IBAN) <> "" Then
    XPrt.Print "Vous voudrez bien nous autoriser à procéder au règlement ET débiter votre compte à l'échéance du " & wDate_Echeance_DB;
    XPrt.CurrentX = prtMinMarge
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.Print "OU nous faire part de vos instructions. ";
Else
    XPrt.Print "Vous voudrez bien nous autoriser à procéder au règlement ET débiter votre compte OU nous faire part de vos ";
    XPrt.CurrentX = prtMinMarge
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.Print "instructions. ";
End If

' Phrase pour charge D.O. seulement
If meCDO_Courrier.YCDOUTI0.CDOUTIBEC = "N" Then
    XPrt.CurrentX = prtMinMarge
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
    XPrt.Print "Le détail de nos frais et commissions vous sera communiqué ultérieurement.";
End If

' >>>>>  Phrase concernant délai de présentation  <<<<<
If Trim(meCDO_Courrier.Delai(1)) <> "" Then
    XPrt.CurrentX = prtMinMarge
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
    XPrt.Print UCase$(meCDO_Courrier.Delai(1));
End If

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.Print "Veuillez nous accuser réception de notre envoi et agréer, Messieurs, nos salutations distinguées. ";

XPrt.CurrentX = prtMinMarge + 5000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.FontBold = True
XPrt.Print paramSignature;
XPrt.FontBold = False

End Sub

'---------------------------------------------------------
Public Sub prtCDO_Courrier_OUV_C_OP_04_AVueNonRecl()
'---------------------------------------------------------
Dim X As String
Dim Tx As Double

XPrt.CurrentY = prtCorpsY
XPrt.FontSize = 14: XPrt.FontBold = True
frmElpPrt.prtCentré prtMedX, "ANNEXE FAISANT PARTIE INTEGRANTE"
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
frmElpPrt.prtCentré prtMedX, "DU CREDIT DOCUMENTAIRE"
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
frmElpPrt.prtCentré prtMedX, "N°  " & Trim(meYCDODOS0.CDODOSEXT)

XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 8
XPrt.CurrentX = prtMinMarge
XPrt.Print "Les frais et commissions étant stipulés à votre charge nous vous réclamerons lors de l'utilisation";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
Tx = mecnfYBIACDOCOM0.CDOCO2TX1 ' 2003.11.20  / 4       ' Taux de confirmation dans SAB : en annuel
Select Case mecnfYBIACDOCOM0.CDOCO2PER
    Case "M":   XPrt.Print " ou à la péremption du crédit notre commission de confirmation calculée au taux de : " & Format$(Tx, "### ##0.00") & " pour cent";
    Case "T":   XPrt.Print " ou à la péremption du crédit notre commission de confirmation calculée au taux de : ";
                XPrt.CurrentX = prtMinMarge + 3000
                XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
                XPrt.Print Format$(Tx, "### ##0.00") & " pour cent par trimestre indivisible";
    Case Else:  XPrt.Print " ou à la péremption du crédit notre commission de confirmation calculée au taux de : " & Format$(Tx, "### ##0.00") & " pour cent";
End Select

XPrt.CurrentX = prtMinMarge + 6000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.FontBold = True
XPrt.Print " Devise               Montant ";
XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.FontItalic = True
XPrt.CurrentX = prtMinMarge
XPrt.Print " Commission de confirmation ";
Select Case mecnfYBIACDOCOM0.CDOCO2PER
    Case "M":   XPrt.CurrentX = prtMinMarge + 2700: XPrt.Print "pour le premier mois";
    Case "T":   XPrt.CurrentX = prtMinMarge + 2700: XPrt.Print "pour le premier trimestre";
    Case Else:  XPrt.CurrentX = prtMinMarge + 2700: XPrt.Print "";
End Select
XPrt.CurrentX = prtMinMarge + 6200: XPrt.Print mecnfYBIACDOCOM0.CDOCOMDEV & "            " & Format$(mecnfYBIACDOCOM0.CDOCOMMON, "### ### ### ##0.00");
XPrt.CurrentX = prtMinMarge
XPrt.FontItalic = False

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print " Par ailleurs, lors de l'utilisation nous vous décompterons : ";

XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "- notre commission de levée de documents au taux de : ";

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinMarge + 2000
XPrt.Print "1,5 pour mille sur le montant utilisé minimum EUR 106,71 plus TVA ";

XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "- frais de port et télex... ";

XPrt.CurrentX = prtMinMarge + 1000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "- éventuellement une commission pour documents irréguliers de EUR 83,85 plus T.V.A. ";

'XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 4
'XPrt.FontSize = 14: XPrt.FontBold = True
'frmElpPrt.prtCentré prtMedX, paramSignature
'XPrt.FontBold = False

XPrt.CurrentX = prtMinMarge + 5000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 4
XPrt.FontBold = True
XPrt.Print paramSignature;
XPrt.FontBold = False

End Sub

Public Sub prtCDO_Courrier_UTI_NCONFORME_BED_Avue()
'---------------------------------------------------------
Dim X As String
Dim C As Integer
Dim W_Code_CRD As String
Dim MNT As String, MTV As String

' >>>>>>  Ce bordereau est adressé à la banque émettrice -agence-

Line2_Ecart = 8: prtCDO_Courrier_Trame (Line2_Ecart)

XPrt.CurrentY = prtCorpsY
XPrt.FontSize = 14: XPrt.FontBold = True: XPrt.FontUnderline = True
frmElpPrt.prtCentré prtMedX, "CREDIT DOCUMENTAIRE"

XPrt.FontSize = 11: XPrt.FontBold = False: XPrt.FontUnderline = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2

XPrt.CurrentX = prtMinMarge: XPrt.Print "  V/Référence :";
XPrt.CurrentX = prtMinMarge + 1550: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1750: XPrt.Print meYCDODOS0.CDODOSEXT;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Donneur d'ordre";
XPrt.CurrentX = prtMinMarge + 1550: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1750: XPrt.Print meCDO_Courrier.DON_Concat;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Bénéficiaire";
XPrt.CurrentX = prtMinMarge + 1550: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1750: XPrt.Print meCDO_Courrier.BEN_Concat;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Montant";
XPrt.CurrentX = prtMinMarge + 1550: XPrt.Print ":";
XPrt.CurrentX = prtMinMarge + 4100: XPrt.Print "Devise :";
XPrt.CurrentX = prtMinMarge + 7500: XPrt.Print "Validité :";
XPrt.FontBold = True
X = Trim(Format$(meYCDODOS0.CDODOSMON, "### ### ### ##0.00"))
XPrt.CurrentX = prtMinMarge + 1750
XPrt.Print X;
XPrt.CurrentX = prtMinMarge + 5000: XPrt.Print meYCDODOS0.CDODOSDEV;
XPrt.CurrentX = prtMinMarge + 8500: XPrt.Print wDate_Validité;
XPrt.FontBold = False

XPrt.FontSize = 11
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.CurrentX = prtMinMarge
XPrt.Print "Messieurs,";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Nous avons l'honneur de vous remettre ci-joints, à notre décharge, les documents de ";
XPrt.Print meYCDODOS0.CDODOSDEV & "  " & Trim(Format$(meCDO_Courrier.YCDOUTI0.CDOUTIMPA, "### ### ### ##0.00")) & " présentés ";
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "en vertu de ce crédit documentaire. ";

' >>>>> Pavé d'éditions des différents types de documents <<<<<
XPrt.FontBold = True
XPrt.FontSize = 8
For C = 1 To meCDO_Courrier.Document_Nb
    If C = 1 Then
        XPrt.CurrentX = prtMinMarge + 5000
        XPrt.CurrentY = XPrt.CurrentY + 200  ' prtlineHeight
        XPrt.FontUnderline = True
        XPrt.Print "1er jeu";
        XPrt.CurrentX = prtMinMarge + 6000
        XPrt.Print "2ème jeu";
        XPrt.FontUnderline = False
    End If
    XPrt.CurrentX = prtMinMarge
    XPrt.CurrentY = XPrt.CurrentY + 200  ' prtlineHeight
    XPrt.Print meCDO_Courrier.Document(C);
    XPrt.CurrentX = prtMinMarge + 5200
    XPrt.Print meCDO_Courrier.Document_Jeu1(C);
    XPrt.CurrentX = prtMinMarge + 6300
    XPrt.Print meCDO_Courrier.Document_Jeu2(C);
Next C

' XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.FontBold = False

' >>>>>  Les éléments concernant l'expédition   <<<<<
For C = 1 To meCDO_Courrier.Exp_Nb
    If C = 1 Then
        XPrt.CurrentX = prtMinMarge
        XPrt.CurrentY = XPrt.CurrentY + 200  '+ prtlineHeight
        XPrt.FontSize = 11
        XPrt.Print "Couvrant l'expédition de : ";
        XPrt.FontSize = 8
        XPrt.CurrentX = prtMinMarge + 3000
        XPrt.Print meCDO_Courrier.Exp(1);
    Else
        XPrt.CurrentX = prtMinMarge + 3000
        XPrt.CurrentY = XPrt.CurrentY + 200  '+ prtlineHeight
        XPrt.Print meCDO_Courrier.Exp(C);
    End If
Next C

XPrt.FontSize = 11

' If meCDO_Courrier.Exp_Par <> Space Then
    XPrt.CurrentX = prtMinMarge
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.Print "Par " & meCDO_Courrier.Exp_Par & " le " & meCDO_Courrier.Exp_Le & " de " & meCDO_Courrier.Exp_De & " à " & meCDO_Courrier.Exp_A;
' End If

' >>>>>  Paragraphe DIFFERENT SI Crédit échu ou Crédit encore valide  <<<<<
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
If W_Validité_AMJ >= YBIATAB0_DATE_CPT_JS1 Then  ' CREDIT VALIDE
    XPrt.Print "Nous vous adressons ces documents pour accord de paiement en raison de(s) irrégularité(s) suivante(s) :";
Else                                ' CREDIT ECHU
    XPrt.Print "Nous vous adressons ces documents sur base d'encaissement en raison de(s) irrégularité(s) suivante(s) :";
    ' XPrt.CurrentX = prtMinMarge
    ' XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    ' XPrt.FontBold = True:     XPrt.FontSize = 8
    ' XPrt.Print "- CREDIT ECHU ";
    ' XPrt.FontBold = False
End If

' >>>>>  Paragraphe de(s) irrégularité(s) : récupérer directement du ZCDOIRR0 par rapport au dossier et no utilisation  <<<<<
For C = 1 To meCDO_Courrier.Irregul_Nb
    XPrt.FontSize = 8
    XPrt.CurrentX = prtMinMarge
    XPrt.CurrentY = XPrt.CurrentY + 200   '+ prtlineHeight
    XPrt.Print meCDO_Courrier.Irregul(C);
Next C

XPrt.FontSize = 11

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1
XPrt.Print "Vous voudrez bien nous autoriser à procéder au règlement ET débiter votre compte OU nous faire part de vos ";
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "instructions. ";

' Phrase pour charge D.O. seulement
If meCDO_Courrier.YCDOUTI0.CDOUTIBEC = "N" Then
    XPrt.CurrentX = prtMinMarge
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
    XPrt.Print "Le détail de nos frais et commissions vous sera communiqué ultérieurement.";
End If

' >>>>>  Phrase concernant délai de présentation  <<<<<
If Trim(meCDO_Courrier.Delai(1)) <> "" Then
    XPrt.CurrentX = prtMinMarge
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
    XPrt.Print UCase$(meCDO_Courrier.Delai(1));
End If

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.Print "Veuillez nous accuser réception de notre envoi et agréer, Messieurs, nos salutations distinguées. ";

XPrt.CurrentX = prtMinMarge + 5000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.FontBold = True
XPrt.Print paramSignature;
XPrt.FontBold = False

End Sub

'---------------------------------------------------------
Public Sub prtCDO_Courrier_OUV_CNF_Page1()
'---------------------------------------------------------
Dim X As String
Line2_Ecart = 7: prtCDO_Courrier_Trame (Line2_Ecart)
'Call frmElpPrt.prtTrame(prtMinMarge - 50, prtCorpsY - 100, prtMaxMarge, prtCorpsY + prtlineHeight * 7 + 100, "B", 245)

XPrt.CurrentY = prtCorpsY
XPrt.FontSize = 14: XPrt.FontBold = True: XPrt.FontUnderline = True
frmElpPrt.prtCentré prtMedX, "CREDIT DOCUMENTAIRE"

XPrt.FontSize = 11: XPrt.FontBold = False: XPrt.FontUnderline = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
'Call frmElpPrt.prtTrame(prtMinMarge - 50, XPrt.CurrentY - 50, prtMaxMarge, XPrt.CurrentY + prtlineHeight * 3 + 50, , 245)
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Donneur d'ordre";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1700: XPrt.Print meCDO_Courrier.DON_Concat;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Bénéficiaire";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1700: XPrt.Print meCDO_Courrier.BEN_Concat;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Montant";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentX = prtMinMarge + 4100: XPrt.Print "Devise :";
XPrt.CurrentX = prtMinMarge + 7500: XPrt.Print "Validité :";
XPrt.FontBold = True
X = Trim(Format$(meYCDODOS0.CDODOSMON, "### ### ### ##0.00"))
XPrt.CurrentX = prtMinMarge + 1700 '4000 - XPrt.TextWidth(X)
XPrt.Print X;
XPrt.CurrentX = prtMinMarge + 5000: XPrt.Print meYCDODOS0.CDODOSDEV;
XPrt.CurrentX = prtMinMarge + 8500: XPrt.Print wDate_Validité;
XPrt.FontBold = False

XPrt.FontSize = 11
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.CurrentX = prtMinMarge
XPrt.Print "Messieurs,";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Nous vous informons de l'ouverture du CREDIT DOCUMENTAIRE irrévocable N° ";
XPrt.FontBold = True
XPrt.Print meYCDODOS0.CDODOSEXT;
XPrt.FontBold = False

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.Print "émis en votre faveur par notre correspondant :";
XPrt.CurrentX = XPrt.CurrentX + 200
prtAdresse meCDO_Courrier.BQE_ZADRESS0, False

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.Print "Valable à nos guichets jusqu'au ";
XPrt.FontBold = True
XPrt.Print wDate_Validité;
XPrt.FontBold = False
' Le 18/01/2005 : dépendre le type de paiement
If meYCDODOS0.CDODOSMCA = 0 And meYCDODOS0.CDODOSMOV <> 0 And meYCDODOS0.CDODOSMDI = 0 Then XPrt.Print " pour présentation des documents pour paiement à vue.";
If meYCDODOS0.CDODOSMCA = 0 And meYCDODOS0.CDODOSMOV <> 0 And meYCDODOS0.CDODOSMDI <> 0 Then XPrt.Print " pour présentation des documents pour paiement mixte.";
If meYCDODOS0.CDODOSMCA = 0 And meYCDODOS0.CDODOSMOV = 0 And meYCDODOS0.CDODOSMDI <> 0 Then XPrt.Print " pour présentation des documents pour paiement différé.";
If meYCDODOS0.CDODOSMCA <> 0 And meYCDODOS0.CDODOSMOV = 0 And meYCDODOS0.CDODOSMDI = 0 Then XPrt.Print " pour présentation des documents pour acceptation.";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.Print "Nous vous prions de noter que ce crédit documentaire comporte NOTRE CONFIRMATION aux conditions";
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "reprises sur ";
XPrt.FontBold = True
XPrt.Print wAnnexe_Nb;
XPrt.FontBold = False
XPrt.Print " faisant partie intégrante de ce crédit.";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.Print "En cas d'irrégularités constatées lors de l'utilisation, notre confirmation deviendra nulle et sans effet au prorata ";
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "de l'utilisation et le règlement ne s'effectuera qu'après réception par nos soins :";
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "    1) de l'accord de la banque émettrice si cet accord est reçu dans la validité du crédit ";
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "    2) de l'accord de la banque émettrice et de la réception de la couverture correspondante si cet accord est reçu ";
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "        après la péremption du crédit. ";


XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Lors de l'utilisation,  nous vous remercions de bien vouloir accompagner ";
XPrt.FontBold = True: XPrt.ForeColor = vbBlue
XPrt.Print "IMPERATIVEMENT";
XPrt.FontBold = False: XPrt.ForeColor = prtForeColor
XPrt.Print " les documents";
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "d'un exemplaire supplémentaire de votre facture et d'un ";
XPrt.FontBold = True: XPrt.ForeColor = vbBlue
XPrt.Print "relevé d'identité bancaire IBAN";
XPrt.FontBold = False: XPrt.ForeColor = prtForeColor
XPrt.Print " pour nous permettre";
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "de vous en effectuer le règlement et de nous communiquer votre ";
XPrt.FontBold = True: XPrt.ForeColor = vbBlue
XPrt.Print "n° de TVA intracommunautaire.";
XPrt.FontBold = False: XPrt.ForeColor = prtForeColor

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.Print "Ce crédit documentaire est soumis aux Règles et Usances Uniformes relatives aux crédits documentaires";
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "(version révisée de 2007 - Publication N° 600 de la Chambre de Commerce Internationale).";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.Print "Si vous n'êtes pas d'accord sur les conditions de ce crédit, nous vous conseillons de vous mettre DIRECTEMENT";
XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "en rapport avec vos acheteurs pour qu'ils donnent les instructions nécessaires de modification à notre correspondant.";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.Print "Veuillez agréer, Messieurs, nos salutations distinguées."


XPrt.CurrentX = prtMinMarge + 5000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.FontBold = True
XPrt.Print paramSignature;
XPrt.FontBold = False

End Sub

Public Sub prtCDO_Courrier_Close()
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



Public Sub prtCDO_Courrier_Open()
On Error GoTo prtError
blnOpen = True
blnNewPage = False
Set XPrt = Printer
frmElpPrt.Show vbModeless

Height8_6 = frmElpPrt.prtHeightDelta(8, 6)

blnFiligrane = False
prtPgmName = "prtCDO_Courrier"
prtTitleUsr = usrName
prtOrientation = vbPRORPortrait
prtTitleText = "CDO_Courrier"
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

Public Function prtCDO_Courrier_Annexe_Nb(lAnnexe_Nb As Integer)
Select Case lAnnexe_Nb
    Case 0: prtCDO_Courrier_Annexe_Nb = ""
    Case 1: prtCDO_Courrier_Annexe_Nb = "l'annexe ci-jointe"
    Case 2: prtCDO_Courrier_Annexe_Nb = "les deux annexes ci-jointes"
    Case 3: prtCDO_Courrier_Annexe_Nb = "les trois annexes ci-jointes"
    Case 4: prtCDO_Courrier_Annexe_Nb = "les quatre annexes ci-jointes"
    Case Else: prtCDO_Courrier_Annexe_Nb = "les " & lAnnexe_Nb & " annexes ci-jointes"
End Select

End Function

'---------------------------------------------------------
Public Sub prtCDO_Courrier_MOD_FCB_01_ValProrogée()
'---------------------------------------------------------
Dim X As String
Dim C As Integer

Line2_Ecart = 7: prtCDO_Courrier_Trame (Line2_Ecart)
'Call frmElpPrt.prtTrame(prtMinMarge - 50, prtCorpsY - 100, prtMaxMarge, prtCorpsY + prtlineHeight * 7 + 100, "B", 245)

XPrt.CurrentY = prtCorpsY
XPrt.FontSize = 14: XPrt.FontBold = True: XPrt.FontUnderline = True
frmElpPrt.prtCentré prtMedX, "CREDIT DOCUMENTAIRE"

XPrt.FontSize = 11: XPrt.FontBold = False: XPrt.FontUnderline = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
'Call frmElpPrt.prtTrame(prtMinMarge - 50, XPrt.CurrentY - 50, prtMaxMarge, XPrt.CurrentY + prtlineHeight * 3 + 50, , 245)
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Donneur d'ordre";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1700: XPrt.Print meCDO_Courrier.DON_Concat;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Bénéficiaire";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1700: XPrt.Print meCDO_Courrier.BEN_Concat;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Montant";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentX = prtMinMarge + 4100: XPrt.Print "Devise :";
XPrt.CurrentX = prtMinMarge + 7500: XPrt.Print "Validité :";
XPrt.FontBold = True
X = Trim(Format$(meYCDODOS0.CDODOSMON, "### ### ### ##0.00"))
XPrt.CurrentX = prtMinMarge + 1700 '4000 - XPrt.TextWidth(X)
XPrt.Print X;
XPrt.CurrentX = prtMinMarge + 5000: XPrt.Print meYCDODOS0.CDODOSDEV;
XPrt.CurrentX = prtMinMarge + 8500: XPrt.Print wDate_Validité;
XPrt.FontBold = False

XPrt.FontSize = 11
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 4
XPrt.CurrentX = prtMinMarge
XPrt.Print "Messieurs,";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Nous vous informons que par SWIFT la banque émettrice, suivant les instructions reçues des ordonnateurs, ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "modifie les conditions du crédit N° ";
XPrt.Print meYCDODOS0.CDODOSEXT;
XPrt.Print " ouvert en votre faveur comme suit : ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.Print "La validité du crédit est prorogée du ";
XPrt.FontBold = True
XPrt.Print wDate_Anc_Validité;  ' Validité du dossier dans ZCDOMOD0
XPrt.FontBold = False
XPrt.Print " au ";
XPrt.FontBold = True
XPrt.Print wDate_Validité;  ' Validité du dossier
XPrt.FontBold = False

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.Print "Tous autres termes et conditions inchangés. ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Nous vous prions de bien vouloir en prendre note. ";

' >>>>>  Paragraphe concernant FRAIS BANQUE EMETTRICE en utilisation la zone GARANTIE DE BONNE EXECUTION en cas -OUV- <<<<<
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
For C = 1 To meCDO_Courrier.Garantie_Nb
    If Trim(meCDO_Courrier.Garantie(C)) <> "" Then
        XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
        XPrt.CurrentX = prtMinMarge
        XPrt.Print meCDO_Courrier.Garantie(C);
    End If
Next C

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Veuillez noter que les frais relatifs à ce crédit étant à votre charge, nous nous réservons de retenir sur le ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "montant de notre règlement, notre commission de modification s'élevant à EUR 60,98 plus TVA 19,60%. ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Veuillez agréer, Messieurs, nos salutations distinguées."

XPrt.CurrentX = prtMinMarge + 5000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.FontBold = True
XPrt.Print paramSignature;
XPrt.FontBold = False

End Sub

'---------------------------------------------------------
Public Sub prtCDO_Courrier_MOD_FCB_02_ValProrogée_Emb()
'---------------------------------------------------------
Dim X As String
Dim C As Integer

Line2_Ecart = 7: prtCDO_Courrier_Trame (Line2_Ecart)
'Call frmElpPrt.prtTrame(prtMinMarge - 50, prtCorpsY - 100, prtMaxMarge, prtCorpsY + prtlineHeight * 7 + 100, "B", 245)

XPrt.CurrentY = prtCorpsY
XPrt.FontSize = 14: XPrt.FontBold = True: XPrt.FontUnderline = True
frmElpPrt.prtCentré prtMedX, "CREDIT DOCUMENTAIRE"

XPrt.FontSize = 11: XPrt.FontBold = False: XPrt.FontUnderline = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
'Call frmElpPrt.prtTrame(prtMinMarge - 50, XPrt.CurrentY - 50, prtMaxMarge, XPrt.CurrentY + prtlineHeight * 3 + 50, , 245)
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Donneur d'ordre";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1700: XPrt.Print meCDO_Courrier.DON_Concat;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Bénéficiaire";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1700: XPrt.Print meCDO_Courrier.BEN_Concat;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Montant";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentX = prtMinMarge + 4100: XPrt.Print "Devise :";
XPrt.CurrentX = prtMinMarge + 7500: XPrt.Print "Validité :";
XPrt.FontBold = True
X = Trim(Format$(meYCDODOS0.CDODOSMON, "### ### ### ##0.00"))
XPrt.CurrentX = prtMinMarge + 1700 '4000 - XPrt.TextWidth(X)
XPrt.Print X;
XPrt.CurrentX = prtMinMarge + 5000: XPrt.Print meYCDODOS0.CDODOSDEV;
XPrt.CurrentX = prtMinMarge + 8500: XPrt.Print wDate_Validité;
XPrt.FontBold = False

XPrt.FontSize = 11
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 4
XPrt.CurrentX = prtMinMarge
XPrt.Print "Messieurs,";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Nous vous informons que par SWIFT la banque émettrice, suivant les instructions reçues des ordonnateurs, ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "modifie les conditions du crédit N° ";
XPrt.Print meYCDODOS0.CDODOSEXT;
XPrt.Print " ouvert en votre faveur comme suit : ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.Print "La validité du crédit est prorogée du ";
XPrt.FontBold = True
XPrt.Print wDate_Anc_Validité;  ' Validité du dossier dans ZCDOMOD0
XPrt.FontBold = False
XPrt.Print " au ";
XPrt.FontBold = True
XPrt.Print wDate_Validité;      ' Validité du dossier
XPrt.FontBold = False

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "La date limite d'embarquement est reportée au ";
XPrt.FontBold = True
XPrt.Print wDate_Limite_Emb;      ' Date limite d'embarquement CDODOSDLE
XPrt.FontBold = False

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.Print "Tous autres termes et conditions inchangés. ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Nous vous prions de bien vouloir en prendre note. ";

' >>>>>  Paragraphe concernant FRAIS BANQUE EMETTRICE en utilisation la zone GARANTIE DE BONNE EXECUTION en cas -OUV- <<<<<
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
For C = 1 To meCDO_Courrier.Garantie_Nb
    If Trim(meCDO_Courrier.Garantie(C)) <> "" Then
        XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
        XPrt.CurrentX = prtMinMarge
        XPrt.Print meCDO_Courrier.Garantie(C);
    End If
Next C

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Veuillez noter que les frais relatifs à ce crédit étant à votre charge, nous nous réservons de retenir sur le ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "montant de notre règlement, notre commission de modification s'élevant à EUR 60,98 plus TVA 19,60%. ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Veuillez agréer, Messieurs, nos salutations distinguées."

XPrt.CurrentX = prtMinMarge + 5000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.FontBold = True
XPrt.Print paramSignature;
XPrt.FontBold = False

End Sub

'---------------------------------------------------------
Public Sub prtCDO_Courrier_MOD_FCB_03_Annexe()
'---------------------------------------------------------
Dim X As String
Dim C As Integer

Line2_Ecart = 7: prtCDO_Courrier_Trame (Line2_Ecart)
'Call frmElpPrt.prtTrame(prtMinMarge - 50, prtCorpsY - 100, prtMaxMarge, prtCorpsY + prtlineHeight * 7 + 100, "B", 245)

XPrt.CurrentY = prtCorpsY
XPrt.FontSize = 14: XPrt.FontBold = True: XPrt.FontUnderline = True
frmElpPrt.prtCentré prtMedX, "CREDIT DOCUMENTAIRE"

XPrt.FontSize = 11: XPrt.FontBold = False: XPrt.FontUnderline = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
'Call frmElpPrt.prtTrame(prtMinMarge - 50, XPrt.CurrentY - 50, prtMaxMarge, XPrt.CurrentY + prtlineHeight * 3 + 50, , 245)
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Donneur d'ordre";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1700: XPrt.Print meCDO_Courrier.DON_Concat;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Bénéficiaire";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1700: XPrt.Print meCDO_Courrier.BEN_Concat;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Montant";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentX = prtMinMarge + 4100: XPrt.Print "Devise :";
XPrt.CurrentX = prtMinMarge + 7500: XPrt.Print "Validité :";
XPrt.FontBold = True
X = Trim(Format$(meYCDODOS0.CDODOSMON, "### ### ### ##0.00"))
XPrt.CurrentX = prtMinMarge + 1700 '4000 - XPrt.TextWidth(X)
XPrt.Print X;
XPrt.CurrentX = prtMinMarge + 5000: XPrt.Print meYCDODOS0.CDODOSDEV;
XPrt.CurrentX = prtMinMarge + 8500: XPrt.Print wDate_Validité;
XPrt.FontBold = False

XPrt.FontSize = 11
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 4
XPrt.CurrentX = prtMinMarge
XPrt.Print "Messieurs,";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Nous vous informons que par SWIFT la banque émettrice, suivant les instructions reçues des ordonnateurs, ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "modifie les conditions du crédit N° ";
XPrt.Print meYCDODOS0.CDODOSEXT;
XPrt.Print " ouvert en votre faveur selon " & wAnnexe_Nb & " faisant ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "partie intégrante du crédit. ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Tous autres termes et conditions inchangés. ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Nous vous prions de bien vouloir en prendre note. ";

' >>>>>  Paragraphe concernant FRAIS BANQUE EMETTRICE en utilisation la zone GARANTIE DE BONNE EXECUTION en cas -OUV- <<<<<
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
For C = 1 To meCDO_Courrier.Garantie_Nb
    If Trim(meCDO_Courrier.Garantie(C)) <> "" Then
        XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
        XPrt.CurrentX = prtMinMarge
        XPrt.Print meCDO_Courrier.Garantie(C);
    End If
Next C

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Veuillez noter que les frais relatifs à ce crédit étant à votre charge, nous nous réservons de retenir sur le ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "montant de notre règlement, notre commission de modification s'élevant à EUR 60,98 plus TVA 19,60%. ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Veuillez agréer, Messieurs, nos salutations distinguées."

XPrt.CurrentX = prtMinMarge + 5000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.FontBold = True
XPrt.Print paramSignature;
XPrt.FontBold = False

End Sub

'---------------------------------------------------------
Public Sub prtCDO_Courrier_MOD_FCO_04_ValProrogée()
'---------------------------------------------------------
Dim X As String
Dim C As Integer

Line2_Ecart = 7: prtCDO_Courrier_Trame (Line2_Ecart)
'Call frmElpPrt.prtTrame(prtMinMarge - 50, prtCorpsY - 100, prtMaxMarge, prtCorpsY + prtlineHeight * 7 + 100, "B", 245)

XPrt.CurrentY = prtCorpsY
XPrt.FontSize = 14: XPrt.FontBold = True: XPrt.FontUnderline = True
frmElpPrt.prtCentré prtMedX, "CREDIT DOCUMENTAIRE"

XPrt.FontSize = 11: XPrt.FontBold = False: XPrt.FontUnderline = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
'Call frmElpPrt.prtTrame(prtMinMarge - 50, XPrt.CurrentY - 50, prtMaxMarge, XPrt.CurrentY + prtlineHeight * 3 + 50, , 245)
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Donneur d'ordre";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1700: XPrt.Print meCDO_Courrier.DON_Concat;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Bénéficiaire";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1700: XPrt.Print meCDO_Courrier.BEN_Concat;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Montant";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentX = prtMinMarge + 4100: XPrt.Print "Devise :";
XPrt.CurrentX = prtMinMarge + 7500: XPrt.Print "Validité :";
XPrt.FontBold = True
X = Trim(Format$(meYCDODOS0.CDODOSMON, "### ### ### ##0.00"))
XPrt.CurrentX = prtMinMarge + 1700 '4000 - XPrt.TextWidth(X)
XPrt.Print X;
XPrt.CurrentX = prtMinMarge + 5000: XPrt.Print meYCDODOS0.CDODOSDEV;
XPrt.CurrentX = prtMinMarge + 8500: XPrt.Print wDate_Validité;
XPrt.FontBold = False

XPrt.FontSize = 11
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 4
XPrt.CurrentX = prtMinMarge
XPrt.Print "Messieurs,";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Nous vous informons que par SWIFT la banque émettrice, suivant les instructions reçues des ordonnateurs, ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "modifie les conditions du crédit N° ";
XPrt.Print meYCDODOS0.CDODOSEXT;
XPrt.Print " ouvert en votre faveur comme suit : ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.Print "La validité du crédit est prorogée du ";
XPrt.FontBold = True
XPrt.Print wDate_Anc_Validité;  ' Validité du dossier dans ZCDOMOD0
XPrt.FontBold = False
XPrt.Print " au ";
XPrt.FontBold = True
XPrt.Print wDate_Validité;      ' Validité du dossier
XPrt.FontBold = False

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.Print "Tous autres termes et conditions inchangés. ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Nous vous prions de bien vouloir en prendre note. ";

' >>>>>  Paragraphe concernant FRAIS BANQUE EMETTRICE en utilisation la zone GARANTIE DE BONNE EXECUTION en cas -OUV- <<<<<
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
For C = 1 To meCDO_Courrier.Garantie_Nb
    If Trim(meCDO_Courrier.Garantie(C)) <> "" Then
        XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
        XPrt.CurrentX = prtMinMarge
        XPrt.Print meCDO_Courrier.Garantie(C);
    End If
Next C

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Veuillez agréer, Messieurs, nos salutations distinguées."

XPrt.CurrentX = prtMinMarge + 5000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.FontBold = True
XPrt.Print paramSignature;
XPrt.FontBold = False

End Sub

'---------------------------------------------------------
Public Sub prtCDO_Courrier_MOD_FCO_05_ValProrogée_Emb()
'---------------------------------------------------------
Dim X As String
Dim C As Integer

Line2_Ecart = 7: prtCDO_Courrier_Trame (Line2_Ecart)
'Call frmElpPrt.prtTrame(prtMinMarge - 50, prtCorpsY - 100, prtMaxMarge, prtCorpsY + prtlineHeight * 7 + 100, "B", 245)

XPrt.CurrentY = prtCorpsY
XPrt.FontSize = 14: XPrt.FontBold = True: XPrt.FontUnderline = True
frmElpPrt.prtCentré prtMedX, "CREDIT DOCUMENTAIRE"

XPrt.FontSize = 11: XPrt.FontBold = False: XPrt.FontUnderline = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
'Call frmElpPrt.prtTrame(prtMinMarge - 50, XPrt.CurrentY - 50, prtMaxMarge, XPrt.CurrentY + prtlineHeight * 3 + 50, , 245)
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Donneur d'ordre";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1700: XPrt.Print meCDO_Courrier.DON_Concat;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Bénéficiaire";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1700: XPrt.Print meCDO_Courrier.BEN_Concat;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Montant";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentX = prtMinMarge + 4100: XPrt.Print "Devise :";
XPrt.CurrentX = prtMinMarge + 7500: XPrt.Print "Validité :";
XPrt.FontBold = True
X = Trim(Format$(meYCDODOS0.CDODOSMON, "### ### ### ##0.00"))
XPrt.CurrentX = prtMinMarge + 1700 '4000 - XPrt.TextWidth(X)
XPrt.Print X;
XPrt.CurrentX = prtMinMarge + 5000: XPrt.Print meYCDODOS0.CDODOSDEV;
XPrt.CurrentX = prtMinMarge + 8500: XPrt.Print wDate_Validité;
XPrt.FontBold = False

XPrt.FontSize = 11
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 4
XPrt.CurrentX = prtMinMarge
XPrt.Print "Messieurs,";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Nous vous informons que par SWIFT la banque émettrice, suivant les instructions reçues des ordonnateurs, ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "modifie les conditions du crédit N° ";
XPrt.Print meYCDODOS0.CDODOSEXT;
XPrt.Print " ouvert en votre faveur comme suit : ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.Print "La validité du crédit est prorogée du ";
XPrt.FontBold = True
XPrt.Print wDate_Anc_Validité;  ' Validité du dossier dans ZCDOMOD0
XPrt.FontBold = False
XPrt.Print " au ";
XPrt.FontBold = True
XPrt.Print wDate_Validité;      ' Validité du dossier
XPrt.FontBold = False

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "La date limite d'embarquement est reportée au ";
XPrt.FontBold = True
XPrt.Print wDate_Limite_Emb;  ' Date limite d'embarquement CDODOSDLE
XPrt.FontBold = False

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.Print "Tous autres termes et conditions inchangés. ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Nous vous prions de bien vouloir en prendre note. ";

' >>>>>  Paragraphe concernant FRAIS BANQUE EMETTRICE en utilisation la zone GARANTIE DE BONNE EXECUTION en cas -OUV- <<<<<
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
For C = 1 To meCDO_Courrier.Garantie_Nb
    If Trim(meCDO_Courrier.Garantie(C)) <> "" Then
        XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
        XPrt.CurrentX = prtMinMarge
        XPrt.Print meCDO_Courrier.Garantie(C);
    End If
Next C

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Veuillez agréer, Messieurs, nos salutations distinguées."

XPrt.CurrentX = prtMinMarge + 5000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.FontBold = True
XPrt.Print paramSignature;
XPrt.FontBold = False

End Sub

'---------------------------------------------------------
Public Sub prtCDO_Courrier_MOD_FCO_06_Annexe()
'---------------------------------------------------------
Dim X As String
Dim C As Integer

Line2_Ecart = 7: prtCDO_Courrier_Trame (Line2_Ecart)
'Call frmElpPrt.prtTrame(prtMinMarge - 50, prtCorpsY - 100, prtMaxMarge, prtCorpsY + prtlineHeight * 7 + 100, "B", 245)

XPrt.CurrentY = prtCorpsY
XPrt.FontSize = 14: XPrt.FontBold = True: XPrt.FontUnderline = True
frmElpPrt.prtCentré prtMedX, "CREDIT DOCUMENTAIRE"

XPrt.FontSize = 11: XPrt.FontBold = False: XPrt.FontUnderline = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
'Call frmElpPrt.prtTrame(prtMinMarge - 50, XPrt.CurrentY - 50, prtMaxMarge, XPrt.CurrentY + prtlineHeight * 3 + 50, , 245)
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Donneur d'ordre";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1700: XPrt.Print meCDO_Courrier.DON_Concat;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Bénéficiaire";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1700: XPrt.Print meCDO_Courrier.BEN_Concat;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Montant";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentX = prtMinMarge + 4100: XPrt.Print "Devise :";
XPrt.CurrentX = prtMinMarge + 7500: XPrt.Print "Validité :";
XPrt.FontBold = True
X = Trim(Format$(meYCDODOS0.CDODOSMON, "### ### ### ##0.00"))
XPrt.CurrentX = prtMinMarge + 1700 '4000 - XPrt.TextWidth(X)
XPrt.Print X;
XPrt.CurrentX = prtMinMarge + 5000: XPrt.Print meYCDODOS0.CDODOSDEV;
XPrt.CurrentX = prtMinMarge + 8500: XPrt.Print wDate_Validité;
XPrt.FontBold = False

XPrt.FontSize = 11
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 4
XPrt.CurrentX = prtMinMarge
XPrt.Print "Messieurs,";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Nous vous informons que par SWIFT la banque émettrice, suivant les instructions reçues des ordonnateurs, ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "modifie les conditions du crédit N° ";
XPrt.Print meYCDODOS0.CDODOSEXT;
XPrt.Print " ouvert en votre faveur selon " & wAnnexe_Nb & " faisant ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "partie intégrante du crédit. ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Tous autres termes et conditions inchangés. ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Nous vous prions de bien vouloir en prendre note. ";

' >>>>>  Paragraphe concernant FRAIS BANQUE EMETTRICE en utilisation la zone GARANTIE DE BONNE EXECUTION en cas -OUV- <<<<<
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
For C = 1 To meCDO_Courrier.Garantie_Nb
    If Trim(meCDO_Courrier.Garantie(C)) <> "" Then
        XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
        XPrt.CurrentX = prtMinMarge
        XPrt.Print meCDO_Courrier.Garantie(C);
    End If
Next C

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Veuillez agréer, Messieurs, nos salutations distinguées."

XPrt.CurrentX = prtMinMarge + 5000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.FontBold = True
XPrt.Print paramSignature;
XPrt.FontBold = False

End Sub

'---------------------------------------------------------
Public Sub prtCDO_Courrier_MOD_FCO_07_ValRaccourcie()
'---------------------------------------------------------
Dim X As String
Dim C As Integer

Line2_Ecart = 7: prtCDO_Courrier_Trame (Line2_Ecart)
'Call frmElpPrt.prtTrame(prtMinMarge - 50, prtCorpsY - 100, prtMaxMarge, prtCorpsY + prtlineHeight * 7 + 100, "B", 245)

XPrt.CurrentY = prtCorpsY
XPrt.FontSize = 14: XPrt.FontBold = True: XPrt.FontUnderline = True
frmElpPrt.prtCentré prtMedX, "CREDIT DOCUMENTAIRE"

XPrt.FontSize = 11: XPrt.FontBold = False: XPrt.FontUnderline = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
'Call frmElpPrt.prtTrame(prtMinMarge - 50, XPrt.CurrentY - 50, prtMaxMarge, XPrt.CurrentY + prtlineHeight * 3 + 50, , 245)
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Donneur d'ordre";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1700: XPrt.Print meCDO_Courrier.DON_Concat;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Bénéficiaire";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1700: XPrt.Print meCDO_Courrier.BEN_Concat;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Montant";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentX = prtMinMarge + 4100: XPrt.Print "Devise :";
XPrt.CurrentX = prtMinMarge + 7500: XPrt.Print "Validité :";
XPrt.FontBold = True
X = Trim(Format$(meYCDODOS0.CDODOSMON, "### ### ### ##0.00"))
XPrt.CurrentX = prtMinMarge + 1700 '4000 - XPrt.TextWidth(X)
XPrt.Print X;
XPrt.CurrentX = prtMinMarge + 5000: XPrt.Print meYCDODOS0.CDODOSDEV;
XPrt.CurrentX = prtMinMarge + 8500: XPrt.Print wDate_Validité;
XPrt.FontBold = False

XPrt.FontSize = 11
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 4
XPrt.CurrentX = prtMinMarge
XPrt.Print "Messieurs,";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Nous vous informons que par SWIFT la banque émettrice, suivant les instructions reçues des ordonnateurs, ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "modifie les conditions du crédit N° ";
XPrt.Print meYCDODOS0.CDODOSEXT;
XPrt.Print " ouvert en votre faveur comme suit : ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.Print "La validité du crédit est ramenée du ";
XPrt.FontBold = True
XPrt.Print wDate_Anc_Validité;  ' Validité du dossier dans ZCDOMOD0
XPrt.FontBold = False
XPrt.Print " au ";
XPrt.FontBold = True
XPrt.Print wDate_Validité;      ' Validité du dossier
XPrt.FontBold = False

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.Print "Tous autres termes et conditions inchangés. ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Nous vous prions de bien vouloir en prendre note. ";

' >>>>>  Paragraphe concernant FRAIS BANQUE EMETTRICE en utilisation la zone GARANTIE DE BONNE EXECUTION en cas -OUV- <<<<<
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
For C = 1 To meCDO_Courrier.Garantie_Nb
    If Trim(meCDO_Courrier.Garantie(C)) <> "" Then
        XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
        XPrt.CurrentX = prtMinMarge
        XPrt.Print meCDO_Courrier.Garantie(C);
    End If
Next C

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Veuillez agréer, Messieurs, nos salutations distinguées."

XPrt.CurrentX = prtMinMarge + 5000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.FontBold = True
XPrt.Print paramSignature;
XPrt.FontBold = False

End Sub

'---------------------------------------------------------
Public Sub prtCDO_Courrier_MOD_FCO_08_ValRaccourcie_Emb()
'---------------------------------------------------------
Dim X As String
Dim C As Integer

Line2_Ecart = 7: prtCDO_Courrier_Trame (Line2_Ecart)
'Call frmElpPrt.prtTrame(prtMinMarge - 50, prtCorpsY - 100, prtMaxMarge, prtCorpsY + prtlineHeight * 7 + 100, "B", 245)

XPrt.CurrentY = prtCorpsY
XPrt.FontSize = 14: XPrt.FontBold = True: XPrt.FontUnderline = True
frmElpPrt.prtCentré prtMedX, "CREDIT DOCUMENTAIRE"

XPrt.FontSize = 11: XPrt.FontBold = False: XPrt.FontUnderline = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
'Call frmElpPrt.prtTrame(prtMinMarge - 50, XPrt.CurrentY - 50, prtMaxMarge, XPrt.CurrentY + prtlineHeight * 3 + 50, , 245)
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Donneur d'ordre";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1700: XPrt.Print meCDO_Courrier.DON_Concat;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Bénéficiaire";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1700: XPrt.Print meCDO_Courrier.BEN_Concat;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Montant";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentX = prtMinMarge + 4100: XPrt.Print "Devise :";
XPrt.CurrentX = prtMinMarge + 7500: XPrt.Print "Validité :";
XPrt.FontBold = True
X = Trim(Format$(meYCDODOS0.CDODOSMON, "### ### ### ##0.00"))
XPrt.CurrentX = prtMinMarge + 1700 '4000 - XPrt.TextWidth(X)
XPrt.Print X;
XPrt.CurrentX = prtMinMarge + 5000: XPrt.Print meYCDODOS0.CDODOSDEV;
XPrt.CurrentX = prtMinMarge + 8500: XPrt.Print wDate_Validité;
XPrt.FontBold = False

XPrt.FontSize = 11
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 4
XPrt.CurrentX = prtMinMarge
XPrt.Print "Messieurs,";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Nous vous informons que par SWIFT la banque émettrice, suivant les instructions reçues des ordonnateurs, ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "modifie les conditions du crédit N° ";
XPrt.Print meYCDODOS0.CDODOSEXT;
XPrt.Print " ouvert en votre faveur comme suit : ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.Print "La validité du crédit est ramenée du ";
XPrt.FontBold = True
XPrt.Print wDate_Anc_Validité;  ' Validité du dossier dans ZCDOMOD0
XPrt.FontBold = False
XPrt.Print " au ";
XPrt.FontBold = True
XPrt.Print wDate_Validité;      ' Validité du dossier
XPrt.FontBold = False

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "La date limite d'embarquement est ramenée au ";
XPrt.FontBold = True
XPrt.Print wDate_Limite_Emb;  ' Date limite d'embarquement CDODOSDLE
XPrt.FontBold = False

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.Print "Tous autres termes et conditions inchangés. ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Nous vous prions de bien vouloir en prendre note. ";

' >>>>>  Paragraphe concernant FRAIS BANQUE EMETTRICE en utilisation la zone GARANTIE DE BONNE EXECUTION en cas -OUV- <<<<<
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
For C = 1 To meCDO_Courrier.Garantie_Nb
    If Trim(meCDO_Courrier.Garantie(C)) <> "" Then
        XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
        XPrt.CurrentX = prtMinMarge
        XPrt.Print meCDO_Courrier.Garantie(C);
    End If
Next C

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Veuillez agréer, Messieurs, nos salutations distinguées."

XPrt.CurrentX = prtMinMarge + 5000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.FontBold = True
XPrt.Print paramSignature;
XPrt.FontBold = False

End Sub

'---------------------------------------------------------
Public Sub prtCDO_Courrier_MOD_FCB_09_ValRaccourcie()
'---------------------------------------------------------
Dim X As String
Dim C As Integer

Line2_Ecart = 7: prtCDO_Courrier_Trame (Line2_Ecart)
'Call frmElpPrt.prtTrame(prtMinMarge - 50, prtCorpsY - 100, prtMaxMarge, prtCorpsY + prtlineHeight * 7 + 100, "B", 245)

XPrt.CurrentY = prtCorpsY
XPrt.FontSize = 14: XPrt.FontBold = True: XPrt.FontUnderline = True
frmElpPrt.prtCentré prtMedX, "CREDIT DOCUMENTAIRE"

XPrt.FontSize = 11: XPrt.FontBold = False: XPrt.FontUnderline = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
'Call frmElpPrt.prtTrame(prtMinMarge - 50, XPrt.CurrentY - 50, prtMaxMarge, XPrt.CurrentY + prtlineHeight * 3 + 50, , 245)
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Donneur d'ordre";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1700: XPrt.Print meCDO_Courrier.DON_Concat;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Bénéficiaire";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1700: XPrt.Print meCDO_Courrier.BEN_Concat;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Montant";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentX = prtMinMarge + 4100: XPrt.Print "Devise :";
XPrt.CurrentX = prtMinMarge + 7500: XPrt.Print "Validité :";
XPrt.FontBold = True
X = Trim(Format$(meYCDODOS0.CDODOSMON, "### ### ### ##0.00"))
XPrt.CurrentX = prtMinMarge + 1700 '4000 - XPrt.TextWidth(X)
XPrt.Print X;
XPrt.CurrentX = prtMinMarge + 5000: XPrt.Print meYCDODOS0.CDODOSDEV;
XPrt.CurrentX = prtMinMarge + 8500: XPrt.Print wDate_Validité;
XPrt.FontBold = False

XPrt.FontSize = 11
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 4
XPrt.CurrentX = prtMinMarge
XPrt.Print "Messieurs,";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Nous vous informons que par SWIFT la banque émettrice, suivant les instructions reçues des ordonnateurs, ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "modifie les conditions du crédit N° ";
XPrt.Print meYCDODOS0.CDODOSEXT;
XPrt.Print " ouvert en votre faveur comme suit : ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.Print "La validité du crédit est ramenée du ";
XPrt.FontBold = True
XPrt.Print wDate_Anc_Validité;  ' Validité du dossier dans ZCDOMOD0
XPrt.FontBold = False
XPrt.Print " au ";
XPrt.FontBold = True
XPrt.Print wDate_Validité;      ' Validité du dossier
XPrt.FontBold = False

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.Print "Tous autres termes et conditions inchangés. ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Nous vous prions de bien vouloir en prendre note. ";

' >>>>>  Paragraphe concernant FRAIS BANQUE EMETTRICE en utilisation la zone GARANTIE DE BONNE EXECUTION en cas -OUV- <<<<<
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
For C = 1 To meCDO_Courrier.Garantie_Nb
    If Trim(meCDO_Courrier.Garantie(C)) <> "" Then
        XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
        XPrt.CurrentX = prtMinMarge
        XPrt.Print meCDO_Courrier.Garantie(C);
    End If
Next C

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Veuillez noter que les frais relatifs à ce crédit étant à votre charge, nous nous réservons de retenir sur le ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "montant de notre règlement, notre commission de modification s'élevant à EUR 60,98 plus TVA 19,60%. ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Veuillez agréer, Messieurs, nos salutations distinguées."

XPrt.CurrentX = prtMinMarge + 5000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.FontBold = True
XPrt.Print paramSignature;
XPrt.FontBold = False

End Sub

'---------------------------------------------------------
Public Sub prtCDO_Courrier_MOD_FCB_10_ValRaccourcie_Emb()
'---------------------------------------------------------
Dim X As String
Dim C As Integer

Line2_Ecart = 7: prtCDO_Courrier_Trame (Line2_Ecart)
'Call frmElpPrt.prtTrame(prtMinMarge - 50, prtCorpsY - 100, prtMaxMarge, prtCorpsY + prtlineHeight * 7 + 100, "B", 245)

XPrt.CurrentY = prtCorpsY
XPrt.FontSize = 14: XPrt.FontBold = True: XPrt.FontUnderline = True
frmElpPrt.prtCentré prtMedX, "CREDIT DOCUMENTAIRE"

XPrt.FontSize = 11: XPrt.FontBold = False: XPrt.FontUnderline = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
'Call frmElpPrt.prtTrame(prtMinMarge - 50, XPrt.CurrentY - 50, prtMaxMarge, XPrt.CurrentY + prtlineHeight * 3 + 50, , 245)
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Donneur d'ordre";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1700: XPrt.Print meCDO_Courrier.DON_Concat;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Bénéficiaire";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1700: XPrt.Print meCDO_Courrier.BEN_Concat;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Montant";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentX = prtMinMarge + 4100: XPrt.Print "Devise :";
XPrt.CurrentX = prtMinMarge + 7500: XPrt.Print "Validité :";
XPrt.FontBold = True
X = Trim(Format$(meYCDODOS0.CDODOSMON, "### ### ### ##0.00"))
XPrt.CurrentX = prtMinMarge + 1700 '4000 - XPrt.TextWidth(X)
XPrt.Print X;
XPrt.CurrentX = prtMinMarge + 5000: XPrt.Print meYCDODOS0.CDODOSDEV;
XPrt.CurrentX = prtMinMarge + 8500: XPrt.Print wDate_Validité;
XPrt.FontBold = False

XPrt.FontSize = 11
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 4
XPrt.CurrentX = prtMinMarge
XPrt.Print "Messieurs,";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Nous vous informons que par SWIFT la banque émettrice, suivant les instructions reçues des ordonnateurs, ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "modifie les conditions du crédit N° ";
XPrt.Print meYCDODOS0.CDODOSEXT;
XPrt.Print " ouvert en votre faveur comme suit : ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.Print "La validité du crédit est ramenée du ";
XPrt.FontBold = True
XPrt.Print wDate_Anc_Validité;  ' Validité du dossier dans ZCDOMOD0
XPrt.FontBold = False
XPrt.Print " au ";
XPrt.FontBold = True
XPrt.Print wDate_Validité;      ' Validité du dossier
XPrt.FontBold = False

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "La date limite d'embarquement est ramemée au ";
XPrt.FontBold = True
XPrt.Print wDate_Limite_Emb;  ' Date limite d'embarquement CDODOSDLE
XPrt.FontBold = False

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.Print "Tous autres termes et conditions inchangés. ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Nous vous prions de bien vouloir en prendre note. ";

' >>>>>  Paragraphe concernant FRAIS BANQUE EMETTRICE en utilisation la zone GARANTIE DE BONNE EXECUTION en cas -OUV- <<<<<
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
For C = 1 To meCDO_Courrier.Garantie_Nb
    If Trim(meCDO_Courrier.Garantie(C)) <> "" Then
        XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
        XPrt.CurrentX = prtMinMarge
        XPrt.Print meCDO_Courrier.Garantie(C);
    End If
Next C

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Veuillez noter que les frais relatifs à ce crédit étant à votre charge, nous nous réservons de retenir sur le ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "montant de notre règlement, notre commission de modification s'élevant à EUR 60,98 plus TVA 19,60%. ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Veuillez agréer, Messieurs, nos salutations distinguées."

XPrt.CurrentX = prtMinMarge + 5000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.FontBold = True
XPrt.Print paramSignature;
XPrt.FontBold = False

End Sub

'---------------------------------------------------------
Public Sub prtCDO_Courrier_MOD_FCB_11_CNF_En_NOT()
'---------------------------------------------------------
Dim X As String
Dim Tx As Double
Line2_Ecart = 7: prtCDO_Courrier_Trame (Line2_Ecart)

'Call frmElpPrt.prtTrame(prtMinMarge - 50, prtCorpsY - 100, prtMaxMarge, prtCorpsY + prtlineHeight * 7 + 100, "B", 245)

XPrt.CurrentY = prtCorpsY
XPrt.FontSize = 14: XPrt.FontBold = True: XPrt.FontUnderline = True
frmElpPrt.prtCentré prtMedX, "CREDIT DOCUMENTAIRE"

XPrt.FontSize = 11: XPrt.FontBold = False: XPrt.FontUnderline = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
'Call frmElpPrt.prtTrame(prtMinMarge - 50, XPrt.CurrentY - 50, prtMaxMarge, XPrt.CurrentY + prtlineHeight * 3 + 50, , 245)
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Donneur d'ordre";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1700: XPrt.Print meCDO_Courrier.DON_Concat;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Bénéficiaire";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1700: XPrt.Print meCDO_Courrier.BEN_Concat;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Montant";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentX = prtMinMarge + 4100: XPrt.Print "Devise :";
XPrt.CurrentX = prtMinMarge + 7500: XPrt.Print "Validité :";
XPrt.FontBold = True
X = Trim(Format$(meYCDODOS0.CDODOSMON, "### ### ### ##0.00"))
XPrt.CurrentX = prtMinMarge + 1700 '4000 - XPrt.TextWidth(X)
XPrt.Print X;
XPrt.CurrentX = prtMinMarge + 5000: XPrt.Print meYCDODOS0.CDODOSDEV;
XPrt.CurrentX = prtMinMarge + 8500: XPrt.Print wDate_Validité;
XPrt.FontBold = False

XPrt.FontSize = 11
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 4
XPrt.CurrentX = prtMinMarge
XPrt.Print "Messieurs,";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Nous vous informons que par SWIFT la banque émettrice, suivant les instructions des ordonnateurs, ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "modifie les conditions du crédit N° ";
XPrt.Print meYCDODOS0.CDODOSEXT;
XPrt.Print " ouvert en votre faveur selon l'annexe jointe faisant ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "partie intégrante du crédit. ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Tous autres termes et conditions inchangés. ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Veuillez nous marquer votre accord sur cette modification. ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Nous vous prions de bien vouloir noter que ce crédit n'étant plus confirmé par notre établissement, il est ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "bien entendu que cette opération ne comporte aucun engagement de notre part quant règlement des ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "documents que nous transmettrons à la banque émettrice lors de l'utilisation. ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "Nous ne vous effectuerons le règlement qu'après réception de la couverture correspondante. ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Par ailleurs, lors de l'utilisation ou à la péremption du crédit, nous vous réclamerons notre commission ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
Tx = menotYBIACDOCOM0.CDOCO2TX1 * 10      ' Taux de notification dans SAB : en % et non pour mille
XPrt.Print "de notification calculée au taux de " & Tx & " pour mille FLAT plus T.V.A. au lieu et place de notre commission ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "de confirmation réclamée sur l'annexe jointe à notre courrier. ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Veuillez agréer, Messieurs, nos salutations distinguées."

XPrt.CurrentX = prtMinMarge + 5000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.FontBold = True
XPrt.Print paramSignature;
XPrt.FontBold = False

End Sub

'---------------------------------------------------------
Public Sub prtCDO_Courrier_MOD_FCO_12_CNF_En_NOT()
'---------------------------------------------------------
Dim X As String
Line2_Ecart = 7: prtCDO_Courrier_Trame (Line2_Ecart)
'Call frmElpPrt.prtTrame(prtMinMarge - 50, prtCorpsY - 100, prtMaxMarge, prtCorpsY + prtlineHeight * 7 + 100, "B", 245)

XPrt.CurrentY = prtCorpsY
XPrt.FontSize = 14: XPrt.FontBold = True: XPrt.FontUnderline = True
frmElpPrt.prtCentré prtMedX, "CREDIT DOCUMENTAIRE"

XPrt.FontSize = 11: XPrt.FontBold = False: XPrt.FontUnderline = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
'Call frmElpPrt.prtTrame(prtMinMarge - 50, XPrt.CurrentY - 50, prtMaxMarge, XPrt.CurrentY + prtlineHeight * 3 + 50, , 245)
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Donneur d'ordre";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1700: XPrt.Print meCDO_Courrier.DON_Concat;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Bénéficiaire";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1700: XPrt.Print meCDO_Courrier.BEN_Concat;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Montant";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentX = prtMinMarge + 4100: XPrt.Print "Devise :";
XPrt.CurrentX = prtMinMarge + 7500: XPrt.Print "Validité :";
XPrt.FontBold = True
X = Trim(Format$(meYCDODOS0.CDODOSMON, "### ### ### ##0.00"))
XPrt.CurrentX = prtMinMarge + 1700 '4000 - XPrt.TextWidth(X)
XPrt.Print X;
XPrt.CurrentX = prtMinMarge + 5000: XPrt.Print meYCDODOS0.CDODOSDEV;
XPrt.CurrentX = prtMinMarge + 8500: XPrt.Print wDate_Validité;
XPrt.FontBold = False

XPrt.FontSize = 11
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 4
XPrt.CurrentX = prtMinMarge
XPrt.Print "Messieurs,";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Nous vous informons que par SWIFT la banque émettrice, suivant les instructions des ordonnateurs, ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "modifie les conditions du crédit N° ";
XPrt.Print meYCDODOS0.CDODOSEXT;
XPrt.Print " ouvert en votre faveur selon l'annexe jointe faisant ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "partie intégrante du crédit. ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Tous autres termes et conditions inchangés. ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Veuillez nous marquer votre accord sur cette modification. ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Nous vous prions de bien vouloir noter que ce crédit n'étant plus confirmé par notre établissement, il ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "est bien entendu que cette opération ne comporte aucun engagement de notre part quant règlement des ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "documents que nous transmettrons à la banque émettrice lors de l'utilisation. ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "Nous ne vous effectuerons le règlement qu'après réception de la couverture correspondante. ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Veuillez agréer, Messieurs, nos salutations distinguées."

XPrt.CurrentX = prtMinMarge + 5000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.FontBold = True
XPrt.Print paramSignature;
XPrt.FontBold = False

End Sub

'---------------------------------------------------------
Public Sub prtCDO_Courrier_MOD_FCB_13_NOT_En_CNF()
'---------------------------------------------------------
Dim X As String
Dim Tx As Double
Line2_Ecart = 7: prtCDO_Courrier_Trame (Line2_Ecart)
'Call frmElpPrt.prtTrame(prtMinMarge - 50, prtCorpsY - 100, prtMaxMarge, prtCorpsY + prtlineHeight * 7 + 100, "B", 245)

XPrt.CurrentY = prtCorpsY
XPrt.FontSize = 14: XPrt.FontBold = True: XPrt.FontUnderline = True
frmElpPrt.prtCentré prtMedX, "CREDIT DOCUMENTAIRE"

XPrt.FontSize = 11: XPrt.FontBold = False: XPrt.FontUnderline = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
'Call frmElpPrt.prtTrame(prtMinMarge - 50, XPrt.CurrentY - 50, prtMaxMarge, XPrt.CurrentY + prtlineHeight * 3 + 50, , 245)
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Donneur d'ordre";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1700: XPrt.Print meCDO_Courrier.DON_Concat;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Bénéficiaire";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1700: XPrt.Print meCDO_Courrier.BEN_Concat;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Montant";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentX = prtMinMarge + 4100: XPrt.Print "Devise :";
XPrt.CurrentX = prtMinMarge + 7500: XPrt.Print "Validité :";
XPrt.FontBold = True
X = Trim(Format$(meYCDODOS0.CDODOSMON, "### ### ### ##0.00"))
XPrt.CurrentX = prtMinMarge + 1700 '4000 - XPrt.TextWidth(X)
XPrt.Print X;
XPrt.CurrentX = prtMinMarge + 5000: XPrt.Print meYCDODOS0.CDODOSDEV;
XPrt.CurrentX = prtMinMarge + 8500: XPrt.Print wDate_Validité;
XPrt.FontBold = False

XPrt.FontSize = 11
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 4
XPrt.CurrentX = prtMinMarge
XPrt.Print "Messieurs,";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Nous vous informons que par SWIFT la banque émettrice, suivant les instructions des ordonnateurs, ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "modifie les conditions du crédit N° ";
XPrt.Print meYCDODOS0.CDODOSEXT;
XPrt.Print " ouvert en votre faveur selon l'annexe jointe faisant ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "partie intégrante du crédit. ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Tous autres termes et conditions inchangés. ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Veuillez nous marquer votre accord sur cette modification. ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Nous vous prions de bien vouloir noter que ce crédit documentaire comporte à présent notre confirmation. ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Par conséquent, les frais et commissions étant à votre charge, nous vous réclamerons ou vous prélèverons ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
Tx = mecnfYBIACDOCOM0.CDOCO2TX1 ' 2003.11.20  / 4       ' Taux de confirmation dans SAB : en annuel
XPrt.Print "lors de l'utilisation ou à la péremption du crédit, notre commission de confirmation au taux de " & Tx & " pour cent ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "par trimestre indivisible au lieu et place de notre commission de notification initialement prévue. ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Veuillez agréer, Messieurs, nos salutations distinguées."

XPrt.CurrentX = prtMinMarge + 5000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.FontBold = True
XPrt.Print paramSignature;
XPrt.FontBold = False

End Sub

'---------------------------------------------------------
Public Sub prtCDO_Courrier_MOD_FCO_14_NOT_En_CNF()
'---------------------------------------------------------
Dim X As String
Line2_Ecart = 7: prtCDO_Courrier_Trame (Line2_Ecart)
'Call frmElpPrt.prtTrame(prtMinMarge - 50, prtCorpsY - 100, prtMaxMarge, prtCorpsY + prtlineHeight * 7 + 100, "B", 245)

XPrt.CurrentY = prtCorpsY
XPrt.FontSize = 14: XPrt.FontBold = True: XPrt.FontUnderline = True
frmElpPrt.prtCentré prtMedX, "CREDIT DOCUMENTAIRE"

XPrt.FontSize = 11: XPrt.FontBold = False: XPrt.FontUnderline = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
'Call frmElpPrt.prtTrame(prtMinMarge - 50, XPrt.CurrentY - 50, prtMaxMarge, XPrt.CurrentY + prtlineHeight * 3 + 50, , 245)
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Donneur d'ordre";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1700: XPrt.Print meCDO_Courrier.DON_Concat;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Bénéficiaire";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1700: XPrt.Print meCDO_Courrier.BEN_Concat;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Montant";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentX = prtMinMarge + 4100: XPrt.Print "Devise :";
XPrt.CurrentX = prtMinMarge + 7500: XPrt.Print "Validité :";
XPrt.FontBold = True
X = Trim(Format$(meYCDODOS0.CDODOSMON, "### ### ### ##0.00"))
XPrt.CurrentX = prtMinMarge + 1700 '4000 - XPrt.TextWidth(X)
XPrt.Print X;
XPrt.CurrentX = prtMinMarge + 5000: XPrt.Print meYCDODOS0.CDODOSDEV;
XPrt.CurrentX = prtMinMarge + 8500: XPrt.Print wDate_Validité;
XPrt.FontBold = False

XPrt.FontSize = 11
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 4
XPrt.CurrentX = prtMinMarge
XPrt.Print "Messieurs,";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Nous vous informons que par SWIFT la banque émettrice, suivant les instructions des ordonnateurs, ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "modifie les conditions du crédit N° ";
XPrt.Print meYCDODOS0.CDODOSEXT;
XPrt.Print " ouvert en votre faveur selon l'annexe jointe faisant ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "partie intégrante du crédit. ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Tous autres termes et conditions inchangés. ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Veuillez nous marquer votre accord sur cette modification. ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Nous vous prions de bien vouloir noter que ce crédit documentaire comporte à présent notre confirmation. ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Veuillez agréer, Messieurs, nos salutations distinguées."

XPrt.CurrentX = prtMinMarge + 5000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.FontBold = True
XPrt.Print paramSignature;
XPrt.FontBold = False

End Sub

'---------------------------------------------------------
Public Sub prtCDO_Courrier_MOD_FCB_17_Augmentation_CNF()
'---------------------------------------------------------
Dim X As String
Dim C As Integer
Dim MNT As Currency
Line2_Ecart = 7: prtCDO_Courrier_Trame (Line2_Ecart)
'Call frmElpPrt.prtTrame(prtMinMarge - 50, prtCorpsY - 100, prtMaxMarge, prtCorpsY + prtlineHeight * 7 + 100, "B", 245)

XPrt.CurrentY = prtCorpsY
XPrt.FontSize = 14: XPrt.FontBold = True: XPrt.FontUnderline = True
frmElpPrt.prtCentré prtMedX, "CREDIT DOCUMENTAIRE"

XPrt.FontSize = 11: XPrt.FontBold = False: XPrt.FontUnderline = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Donneur d'ordre";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1700: XPrt.Print meCDO_Courrier.DON_Concat;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Bénéficiaire";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1700: XPrt.Print meCDO_Courrier.BEN_Concat;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Montant";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentX = prtMinMarge + 4100: XPrt.Print "Devise :";
XPrt.CurrentX = prtMinMarge + 7500: XPrt.Print "Validité :";
XPrt.FontBold = True
X = Trim(Format$(meYCDODOS0.CDODOSMON, "### ### ### ##0.00"))
XPrt.CurrentX = prtMinMarge + 1700 '4000 - XPrt.TextWidth(X)
XPrt.Print X;
XPrt.CurrentX = prtMinMarge + 5000: XPrt.Print meYCDODOS0.CDODOSDEV;
XPrt.CurrentX = prtMinMarge + 8500: XPrt.Print wDate_Validité;
XPrt.FontBold = False

XPrt.FontSize = 11
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 4
XPrt.CurrentX = prtMinMarge
XPrt.Print "Messieurs,";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Nous vous informons que par SWIFT la banque émettrice, suivant les instructions reçues des ordonnateurs, ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "modifie les conditions du crédit N° ";
XPrt.Print meYCDODOS0.CDODOSEXT;
XPrt.Print " ouvert en votre faveur comme suit : ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
MNT = meYCDODOS0.CDODOSMON - meYCDOMOD0.CDOMODMON
XPrt.Print "Le montant du crédit est augmenté de : " & meYCDODOS0.CDODOSDEV & "  " & Format$(MNT, "### ### ### ##0.00");

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.Print "Tous autres termes et conditions inchangés. ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Nous vous prions de bien vouloir en prendre note. ";

' >>>>>  Paragraphe concernant FRAIS BANQUE EMETTRICE en utilisation la zone GARANTIE DE BONNE EXECUTION en cas -OUV- <<<<<
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
For C = 1 To meCDO_Courrier.Garantie_Nb
    If Trim(meCDO_Courrier.Garantie(C)) <> "" Then
        XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
        XPrt.CurrentX = prtMinMarge
        XPrt.Print meCDO_Courrier.Garantie(C);
    End If
Next C

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Veuillez noter que les frais relatifs à ce crédit étant à votre charge, nous nous réservons de retenir sur le ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "montant de notre règlement,notre commission de confirmation sur le montant augmenté. ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Veuillez agréer, Messieurs, nos salutations distinguées."

XPrt.CurrentX = prtMinMarge + 5000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.FontBold = True
XPrt.Print paramSignature;
XPrt.FontBold = False

End Sub

'---------------------------------------------------------
Public Sub prtCDO_Courrier_MOD_FCB_20_Augmentation_NOT()
'---------------------------------------------------------
Dim X As String
Dim C As Integer
Dim MNT As Currency
Line2_Ecart = 7: prtCDO_Courrier_Trame (Line2_Ecart)
'Call frmElpPrt.prtTrame(prtMinMarge - 50, prtCorpsY - 100, prtMaxMarge, prtCorpsY + prtlineHeight * 7 + 100, "B", 245)

XPrt.CurrentY = prtCorpsY
XPrt.FontSize = 14: XPrt.FontBold = True: XPrt.FontUnderline = True
frmElpPrt.prtCentré prtMedX, "CREDIT DOCUMENTAIRE"

XPrt.FontSize = 11: XPrt.FontBold = False: XPrt.FontUnderline = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Donneur d'ordre";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1700: XPrt.Print meCDO_Courrier.DON_Concat;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Bénéficiaire";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1700: XPrt.Print meCDO_Courrier.BEN_Concat;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Montant";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentX = prtMinMarge + 4100: XPrt.Print "Devise :";
XPrt.CurrentX = prtMinMarge + 7500: XPrt.Print "Validité :";
XPrt.FontBold = True
X = Trim(Format$(meYCDODOS0.CDODOSMON, "### ### ### ##0.00"))
XPrt.CurrentX = prtMinMarge + 1700 '4000 - XPrt.TextWidth(X)
XPrt.Print X;
XPrt.CurrentX = prtMinMarge + 5000: XPrt.Print meYCDODOS0.CDODOSDEV;
XPrt.CurrentX = prtMinMarge + 8500: XPrt.Print wDate_Validité;
XPrt.FontBold = False

XPrt.FontSize = 11
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 4
XPrt.CurrentX = prtMinMarge
XPrt.Print "Messieurs,";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Nous vous informons que par SWIFT la banque émettrice, suivant les instructions reçues des ordonnateurs, ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "modifie les conditions du crédit N° ";
XPrt.Print meYCDODOS0.CDODOSEXT;
XPrt.Print " ouvert en votre faveur comme suit : ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
MNT = meYCDODOS0.CDODOSMON - meYCDOMOD0.CDOMODMON
XPrt.Print "Le montant du crédit est augmenté de : " & meYCDODOS0.CDODOSDEV & "  " & Format$(MNT, "### ### ### ##0.00");

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.Print "Tous autres termes et conditions inchangés. ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Nous vous prions de bien vouloir en prendre note. ";

' >>>>>  Paragraphe concernant FRAIS BANQUE EMETTRICE en utilisation la zone GARANTIE DE BONNE EXECUTION en cas -OUV- <<<<<
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
For C = 1 To meCDO_Courrier.Garantie_Nb
    If Trim(meCDO_Courrier.Garantie(C)) <> "" Then
        XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
        XPrt.CurrentX = prtMinMarge
        XPrt.Print meCDO_Courrier.Garantie(C);
    End If
Next C

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Veuillez noter que les frais relatifs à ce crédit étant à votre charge, nous nous réservons de retenir sur le ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "montant de notre règlement, notre commission de notification sur le montant augmenté. ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Veuillez agréer, Messieurs, nos salutations distinguées."

XPrt.CurrentX = prtMinMarge + 5000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.FontBold = True
XPrt.Print paramSignature;
XPrt.FontBold = False

End Sub

'---------------------------------------------------------
Public Sub prtCDO_Courrier_MOD_FCO_18_Augmentation()
'---------------------------------------------------------
Dim X As String
Dim C As Integer
Dim MNT As Currency
Line2_Ecart = 7: prtCDO_Courrier_Trame (Line2_Ecart)
'Call frmElpPrt.prtTrame(prtMinMarge - 50, prtCorpsY - 100, prtMaxMarge, prtCorpsY + prtlineHeight * 7 + 100, "B", 245)

XPrt.CurrentY = prtCorpsY
XPrt.FontSize = 14: XPrt.FontBold = True: XPrt.FontUnderline = True
frmElpPrt.prtCentré prtMedX, "CREDIT DOCUMENTAIRE"

XPrt.FontSize = 11: XPrt.FontBold = False: XPrt.FontUnderline = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Donneur d'ordre";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1700: XPrt.Print meCDO_Courrier.DON_Concat;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Bénéficiaire";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1700: XPrt.Print meCDO_Courrier.BEN_Concat;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Montant";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentX = prtMinMarge + 4100: XPrt.Print "Devise :";
XPrt.CurrentX = prtMinMarge + 7500: XPrt.Print "Validité :";
XPrt.FontBold = True
X = Trim(Format$(meYCDODOS0.CDODOSMON, "### ### ### ##0.00"))
XPrt.CurrentX = prtMinMarge + 1700 '4000 - XPrt.TextWidth(X)
XPrt.Print X;
XPrt.CurrentX = prtMinMarge + 5000: XPrt.Print meYCDODOS0.CDODOSDEV;
XPrt.CurrentX = prtMinMarge + 8500: XPrt.Print wDate_Validité;
XPrt.FontBold = False

XPrt.FontSize = 11
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 4
XPrt.CurrentX = prtMinMarge
XPrt.Print "Messieurs,";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Nous vous informons que par SWIFT la banque émettrice, suivant les instructions reçues des ordonnateurs, ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "modifie les conditions du crédit N° ";
XPrt.Print meYCDODOS0.CDODOSEXT;
XPrt.Print " ouvert en votre faveur comme suit : ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
MNT = meYCDODOS0.CDODOSMON - meYCDOMOD0.CDOMODMON
XPrt.Print "Le montant du crédit est augmenté de : " & meYCDODOS0.CDODOSDEV & "  " & Format$(MNT, "### ### ### ##0.00");

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.Print "Tous autres termes et conditions inchangés. ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Nous vous prions de bien vouloir en prendre note. ";

' >>>>>  Paragraphe concernant FRAIS BANQUE EMETTRICE en utilisation la zone GARANTIE DE BONNE EXECUTION en cas -OUV- <<<<<
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
For C = 1 To meCDO_Courrier.Garantie_Nb
    If Trim(meCDO_Courrier.Garantie(C)) <> "" Then
        XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
        XPrt.CurrentX = prtMinMarge
        XPrt.Print meCDO_Courrier.Garantie(C);
    End If
Next C

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Veuillez agréer, Messieurs, nos salutations distinguées."

XPrt.CurrentX = prtMinMarge + 5000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.FontBold = True
XPrt.Print paramSignature;
XPrt.FontBold = False

End Sub

'---------------------------------------------------------
Public Sub prtCDO_Courrier_MOD_19_Diminution()
'---------------------------------------------------------
Dim X As String
Dim C As Integer
Line2_Ecart = 7: prtCDO_Courrier_Trame (Line2_Ecart)
'Call frmElpPrt.prtTrame(prtMinMarge - 50, prtCorpsY - 100, prtMaxMarge, prtCorpsY + prtlineHeight * 7 + 100, "B", 245)

XPrt.CurrentY = prtCorpsY
XPrt.FontSize = 14: XPrt.FontBold = True: XPrt.FontUnderline = True
frmElpPrt.prtCentré prtMedX, "CREDIT DOCUMENTAIRE"

XPrt.FontSize = 11: XPrt.FontBold = False: XPrt.FontUnderline = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Donneur d'ordre";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1700: XPrt.Print meCDO_Courrier.DON_Concat;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Bénéficiaire";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 9: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1700: XPrt.Print meCDO_Courrier.BEN_Concat;
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 11: XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge: XPrt.Print "  Montant";
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print ":";
XPrt.CurrentX = prtMinMarge + 4100: XPrt.Print "Devise :";
XPrt.CurrentX = prtMinMarge + 7500: XPrt.Print "Validité :";
XPrt.FontBold = True
X = Trim(Format$(meYCDODOS0.CDODOSMON, "### ### ### ##0.00"))
XPrt.CurrentX = prtMinMarge + 1700 '4000 - XPrt.TextWidth(X)
XPrt.Print X;
XPrt.CurrentX = prtMinMarge + 5000: XPrt.Print meYCDODOS0.CDODOSDEV;
XPrt.CurrentX = prtMinMarge + 8500: XPrt.Print wDate_Validité;
XPrt.FontBold = False

XPrt.FontSize = 11
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 4
XPrt.CurrentX = prtMinMarge
XPrt.Print "Messieurs,";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Nous vous informons que par SWIFT la banque émettrice, suivant les instructions reçues des ordonnateurs, ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "modifie les conditions du crédit N° ";
XPrt.Print meYCDODOS0.CDODOSEXT;
XPrt.Print " ouvert en votre faveur comme suit : ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.Print "Le montant du crédit est ramené de :  " & meYCDOMOD0.CDOMODDEV & "  " & Format$(meYCDOMOD0.CDOMODMON, "### ### ### ##0.00") & "  à  " & meYCDODOS0.CDODOSDEV & "  " & Format$(meYCDODOS0.CDODOSMON, "### ### ### ##0.00");

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.Print "Tous autres termes et conditions inchangés. ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Nous vous prions de bien vouloir en prendre note et nous marquer votre accord sur cette modification. ";

XPrt.CurrentX = prtMinMarge
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Print "Veuillez agréer, Messieurs, nos salutations distinguées."

XPrt.CurrentX = prtMinMarge + 5000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.FontBold = True
XPrt.Print paramSignature;
XPrt.FontBold = False

End Sub


Public Sub prtCDO_Courrier_Trame(Line2_Ecart As Integer)
Dim Line1 As Integer, Line2 As Integer
Dim col1 As Integer, col2 As Integer

Line1 = prtCorpsY - 100
Line2 = prtCorpsY + prtlineHeight * Line2_Ecart + 100
col1 = prtMinMarge - 50
col2 = prtMaxMarge

Call frmElpPrt.prtTrame(col1, Line1, col2, Line2, " ", 245)
XPrt.DrawWidth = 2
XPrt.Line (col1 + 200, Line1)-(col2 - 200, Line1), prtLineColor
XPrt.Line (col1 + 200, Line2)-(col2 - 200, Line2), prtLineColor
XPrt.Line (col1, Line1 + 200)-(col1, Line2 - 200), prtLineColor
XPrt.Line (col2, Line1 + 200)-(col2, Line2 - 200), prtLineColor

XPrt.CurrentY = 0
XPrt.CurrentX = 0
XPrt.Circle Step(col1 + 200, Line1 + 200), 200, prtLineColor, 0.5 * Pi, Pi
XPrt.CurrentY = 0
XPrt.CurrentX = 0
XPrt.Circle Step(col1 + 200, Line2 - 200), 200, prtLineColor, Pi, 1.5 * Pi
XPrt.CurrentY = 0
XPrt.CurrentX = 0
XPrt.Circle Step(col2 - 200, Line1 + 200), 200, prtLineColor, 0, 0.5 * Pi
XPrt.CurrentY = 0
XPrt.CurrentX = 0
XPrt.Circle Step(col2 - 200, Line2 - 200), 200, prtLineColor, 1.5 * Pi, 0

End Sub

'---------------------------------------------------------
Public Sub prtCDO_Courrier_form(W_Adresse As String)
'---------------------------------------------------------
Dim X As String

If Not blnOpen Then prtCDO_Courrier_Open
If blnNewPage Then frmElpPrt.prtNewPage   'XPrt.NewPage
blnNewPage = True

XPrt.DrawWidth = 1
XPrt.FontSize = 10: XPrt.FontBold = False

XPrt.CurrentX = prtMinX + 7800
XPrt.CurrentY = prtMinY + prtlineHeight * 4

If Booleen_GB = True Then
    XPrt.Print "Paris, " & dateImp_ddMonthYYYY(DSys);
Else
    XPrt.Print "Paris, le  " & dateImp10(DSys);
End If

Select Case W_Adresse
    Case "B": Call prtAdresse_Enveloppe(meCDO_Courrier.BEN_ZADRESS0)
    Case "E": Call prtAdresse_Enveloppe(meCDO_Courrier.BQE_ZADRESS0)
    Case "D": Call prtAdresse_Enveloppe(meCDO_Courrier.BED_ZADRESS0)
End Select

XPrt.FontSize = 8
XPrt.FontBold = False
XPrt.CurrentY = prtRéférenceY
If Booleen_GB = True Then
    XPrt.CurrentX = prtMinMarge: XPrt.Print "O/Reference";
Else
    XPrt.CurrentX = prtMinMarge: XPrt.Print "N/Référence";
End If
XPrt.CurrentX = prtMinMarge + 1250: XPrt.Print ":";
XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1400: XPrt.Print meYCDODOS0.CDODOSCOP & " " & meYCDODOS0.CDODOSDOS;
XPrt.FontBold = False

'XPrt.FontSize = 6
'XPrt.CurrentY = XPrt.CurrentY + Height8_6
'XPrt.Print " / " & xDocRéférence;
'XPrt.FontSize = 8
'XPrt.CurrentY = XPrt.CurrentY - Height8_6

XPrt.CurrentY = prtRéférenceY + prtlineHeight * 1.5
If Booleen_GB = True Then
    XPrt.CurrentX = prtMinMarge: XPrt.Print "Your contact";
Else
    XPrt.CurrentX = prtMinMarge: XPrt.Print "Votre contact";
End If
XPrt.CurrentX = prtMinMarge + 1250: XPrt.Print ":";
XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1400: XPrt.Print "33 (0)1 53 76 " & wContact;

XPrt.FontBold = False
XPrt.CurrentY = prtRéférenceY + prtlineHeight * 3
If Booleen_GB = True Then
    XPrt.CurrentX = prtMinMarge: XPrt.Print "Fax number";
Else
    XPrt.CurrentX = prtMinMarge: XPrt.Print "Télécopie";
End If
XPrt.CurrentX = prtMinMarge + 1250: XPrt.Print ":";
XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1400: XPrt.Print "33 (0)1 53 76 62 58";

' Ligne "A l'attention de..." est réservée seulement pour -AR- aux courriers UTILISATIONS
If Booleen_AR = True Then
    If Trim(meCDO_Courrier.ATT) <> "" Then
        XPrt.FontBold = False
        XPrt.CurrentY = prtRéférenceY + prtlineHeight * 6
        If Booleen_GB = True Then
            XPrt.CurrentX = prtMinMarge: XPrt.Print "To the attention of ";
        Else
            XPrt.CurrentX = prtMinMarge: XPrt.Print "A l'attention de ";
        End If
        XPrt.CurrentX = prtMinMarge + 1250: XPrt.Print ":";
        XPrt.FontBold = True
        XPrt.CurrentX = prtMinMarge + 1400: XPrt.Print meCDO_Courrier.ATT;
    End If
End If

XPrt.FontSize = 10: XPrt.FontBold = False

End Sub



