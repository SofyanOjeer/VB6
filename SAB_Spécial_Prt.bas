Attribute VB_Name = "prtSAB_Spécial"
Option Explicit
Dim X As String, I As Integer, Height8_6 As Integer

Dim xIn As String
Dim K As Integer, L1 As Integer, L2 As Integer
Dim wZADRESS0 As typeZADRESS0
Dim curX As Currency, wSOMME As Currency
Dim blnAdresse As Boolean
Dim wRIB_BQ(100) As String, wRIB_GUI(100) As String, wRIB_CPT(100) As String, wRIB_Nb As Integer
Dim wTIREUR As String, wACC As String, wREF As String
Dim wECH As String, wENVOI As String, wRGL As String
Dim BOO_EDI_AVIS_1ER As String


Dim blnNewPage As Boolean, blnOpen As Boolean

Dim mCurrentY_RIB As Integer
Dim mprtSAB_SITTE003P1_Entete_Avis As String

Public Sub prtSAB_Spécial_Close()
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

Public Sub prtSAB_Spécial_Open(lOpen As String, lMsg As String)
On Error GoTo prtError
blnOpen = True
blnNewPage = False
Set XPrt = Printer
frmElpPrt.Show vbModeless

Height8_6 = frmElpPrt.prtHeightDelta(8, 6)

blnFiligrane = False
prtPgmName = "prtSAB_Spécial"
prtTitleUsr = usrName
prtTitleText = "SAB_Spécial"
prtFontName = prtFontName_TimesNewRoman 'Arial

prtLineNb = 1
prtlineHeight = 250
prtHeaderHeight = 50 ' 100

prtFormType = ""

If lOpen = "ECHEDI04P1" Then
    prtFontName = prtFontName_CourierNew
    prtTitleText = lMsg
    prtOrientation = vbPRORLandscape
    frmElpPrt.prtStdInit
    XPrt.FontSize = 9
    XPrt.FontBold = False

Else
    prtOrientation = vbPRORPortrait
    prtSocInit
End If
Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub
Public Sub prtSAB_SITTE003P1_NewLine()
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
If XPrt.CurrentY + 300 > prtMaxY Then
    Select Case mprtSAB_SITTE003P1_Entete_Avis
        Case "prtSAB_SITTE003P1_Entete_Avis01": prtSAB_SITTE003P1_Entete_Avis01
        Case "prtSAB_SITTE003P1_Entete_Avis02": prtSAB_SITTE003P1_Entete_Avis02
        Case Else:
                frmElpPrt.prtNewPage
                prtSAB_Form
    End Select
End If

End Sub

Public Sub prtSAB_Form()
Dim wId As String
Dim X As String


XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY), prtLineColor
XPrt.CurrentY = XPrt.CurrentY + 50


End Sub




Public Sub prtSAB_SITTE003P1(lFileName As String)
''On Error Resume Next
Dim wInstr As Integer
' >>>>>  Lecture du fichier spool pour découper les zones à imprimer - AVIS01 -
Dim blnSignature As Boolean, blnPageSuite As Boolean

blnPageSuite = False
Open lFileName For Input As #1
wRIB_Nb = 0
Do Until EOF(1)
    Line Input #1, xIn
    wInstr = InStr(114, xIn, "12179 00001")                       'nombreux effets
    If wInstr > 0 Then Mid$(xIn, 1, 4) = "032 "                 ' bricolage du 2005.11.22
    Select Case Mid$(xIn, 1, 4)
        Case "032 ": wRIB_Nb = wRIB_Nb + 1
                     wRIB_BQ(wRIB_Nb) = Mid$(xIn, 114, 5)
                     wRIB_GUI(wRIB_Nb) = Mid$(xIn, 120, 5)
                     wRIB_CPT(wRIB_Nb) = Mid$(xIn, 126, 11)
   End Select
Loop
Close



prtSAB_Spécial_Open "SITTE003P1", lFileName ' Ouvrir la partir EDITION
XPrt.DrawWidth = 1

BOO_EDI_AVIS_1ER = "O"
Open lFileName For Input As #1
wRIB_Nb = 0
Do Until EOF(1)
    Line Input #1, xIn
    ''Debug.Print xIn
    wInstr = InStr(1, xIn, "Signature :")                       'saut de ligne 2 ou 1 suivant nb lignes de l'adresse
    If wInstr > 0 Then Mid$(xIn, 1, 4) = "   1"                 ' bricolage du 2003.12.01
    wInstr = InStr(114, xIn, "12179 00001")                       'nombreux effets
    ' bricolage du 2005.11.22
    If wInstr > 0 Then Mid$(xIn, 1, 4) = "032 "                 ' bricolage du 2005.11.22
    
    Select Case Mid$(xIn, 1, 4)
        Case "011 ": L1 = 0: blnSignature = False
                     L2 = 0
                     K = 0
                     wENVOI = Mid$(xIn, 105, 8)
                     wRGL = Mid$(xIn, 92, 8)
                     blnAdresse = True
                     rsZADRESS0_Init wZADRESS0
                     wZADRESS0.ADRESSRA1 = Mid$(xIn, 129, 40)
                     K = 1
        Case "   2", "   3":
                     If blnAdresse Then
                        If K = 6 Then wZADRESS0.ADRESSRA2 = wZADRESS0.ADRESSAD1: wZADRESS0.ADRESSAD1 = ""
                        blnSignature = True
                        wSOMME = 0    ' RAZ après chaque en-tête  !! si effets > 10 !!!
                                        ' >>>>>  Début 1er avis  <<<<<
                        wRIB_Nb = wRIB_Nb + 1
                        If Not blnPageSuite Then prtSAB_SITTE003P1_Entete_Avis01: blnPageSuite = True
                        blnAdresse = False
                    End If
                    
                    L2 = L2 + 1
                     If L2 = 2 Then
                       wTIREUR = Mid$(xIn, 99, 25)
                       wREF = Mid$(xIn, 88, 6)
                       wACC = Mid$(xIn, 136, 1)
                       wECH = Mid$(xIn, 138, 2) & "/" & Mid$(xIn, 140, 2) & "/" & Mid$(xIn, 142, 2)
                       curX = CCur(Mid$(xIn, 144, 12) / 100)
                       ' Impression avis ligne effet
                       prtSAB_SITTE003P1_Detail
                    End If
        Case "032 ": 'wRIB_BQ = Mid$(xIn, 114, 5)
                     'wRIB_GUI = Mid$(xIn, 120, 5)
                     'wRIB_CPT = Mid$(xIn, 126, 11)
                     prtSAB_SITTE003P1_Fin_Avis01
                     blnPageSuite = False
        Case "   1":
                L1 = L1 + 1
               If Not blnSignature Then
                    Select Case K
                        Case 1: wZADRESS0.ADRESSAD1 = Mid$(xIn, 129, 40): K = 2
                        Case 2: wZADRESS0.ADRESSAD2 = Mid$(xIn, 129, 40): K = 3
                        Case 3: wZADRESS0.ADRESSAD3 = Mid$(xIn, 129, 40): K = 4
                        Case 4: wZADRESS0.ADRESSVIL = Mid$(xIn, 129, 40): K = 5
                        Case 5: wZADRESS0.ADRESSPAY = Mid$(xIn, 129, 40): K = 6
                    End Select
                Else
                    If L1 >= 3 Then
                       If Trim(Mid$(xIn, 89, 6)) <> "" Then
                            wTIREUR = Mid$(xIn, 99, 25)
                            wREF = Mid$(xIn, 88, 6)
                            wACC = Mid$(xIn, 136, 1)
                            wECH = Mid$(xIn, 138, 2) & "/" & Mid$(xIn, 140, 2) & "/" & Mid$(xIn, 142, 2)
                            curX = CCur(Mid$(xIn, 144, 12) / 100)
                            ' Impression avis ligne effet
                            prtSAB_SITTE003P1_Detail
                       End If
                    End If
                End If

   End Select
Loop
Close

prtSAB_Spécial_Close

' >>>>>  Lecture du fichier spool pour découper les zones à imprimer - AVIS02 -

prtSAB_Spécial_Open "SITTE003P1", lFileName ' Ouvrir la partir EDITION
XPrt.DrawWidth = 1
blnPageSuite = False

BOO_EDI_AVIS_1ER = "O"
Open lFileName For Input As #1
wRIB_Nb = 0
Do Until EOF(1)
    Line Input #1, xIn
    
    wInstr = InStr(1, xIn, "Signature :")                       'saut de ligne 2 ou 1 suivant nb lignes de l'adresse
    If wInstr > 0 Then Mid$(xIn, 1, 4) = "   1"                 ' bricolage du 2003.12.01
    wInstr = InStr(114, xIn, "12179 00001")                       'nombreux effets
    ' bricolage du 2005.11.22
    If wInstr > 0 Then Mid$(xIn, 1, 4) = "032 "                 ' bricolage du 2005.11.22
    Select Case Mid$(xIn, 1, 4)
        Case "011 ": L1 = 0: blnSignature = False
                     L2 = 0
                     K = 0
                     wENVOI = Mid$(xIn, 105, 8)
                     wRGL = Mid$(xIn, 92, 8)
                     blnAdresse = True
                     rsZADRESS0_Init wZADRESS0
                     wZADRESS0.ADRESSRA1 = Mid$(xIn, 129, 40)
                     K = 1
        Case "   2", "   3":
                     If blnAdresse Then
                        If K = 6 Then wZADRESS0.ADRESSRA2 = wZADRESS0.ADRESSAD1: wZADRESS0.ADRESSAD1 = ""
                        blnSignature = True
                        wSOMME = 0    ' RAZ après chaque en-tête  !! si effets > 10 !!!
                       wRIB_Nb = wRIB_Nb + 1
                       If Not blnPageSuite Then prtSAB_SITTE003P1_Entete_Avis02: blnPageSuite = True
                        blnAdresse = False
                    End If
                     L2 = L2 + 1
                     If L2 = 2 Then
                        wTIREUR = Mid$(xIn, 99, 25)
                        wREF = Mid$(xIn, 88, 6)
                        wACC = Mid$(xIn, 136, 1)
                        wECH = Mid$(xIn, 138, 2) & "/" & Mid$(xIn, 140, 2) & "/" & Mid$(xIn, 142, 2)
                        curX = CCur(Mid$(xIn, 144, 12) / 100)
                        ' Impression avis ligne effet
                        prtSAB_SITTE003P1_Detail
                     End If
        Case "032 ": 'wRIB_BQ = Mid$(xIn, 114, 5)
                     'wRIB_GUI = Mid$(xIn, 120, 5)
                     'wRIB_CPT = Mid$(xIn, 126, 11)
                     prtSAB_SITTE003P1_Fin_Avis02
                     blnPageSuite = False

        Case "   1":
                L1 = L1 + 1
               If Not blnSignature Then
                    Select Case K
                        Case 1: wZADRESS0.ADRESSAD1 = Mid$(xIn, 129, 40): K = 2
                        Case 2: wZADRESS0.ADRESSAD2 = Mid$(xIn, 129, 40): K = 3
                        Case 3: wZADRESS0.ADRESSAD3 = Mid$(xIn, 129, 40): K = 4
                        Case 4: wZADRESS0.ADRESSVIL = Mid$(xIn, 129, 40): K = 5
                        Case 5: wZADRESS0.ADRESSPAY = Mid$(xIn, 129, 40): K = 6
                    End Select
                Else
                    If L1 >= 3 Then
                       If Trim(Mid$(xIn, 89, 6)) <> "" Then
                            wTIREUR = Mid$(xIn, 99, 25)
                            wREF = Mid$(xIn, 88, 6)
                            wACC = Mid$(xIn, 136, 1)
                            wECH = Mid$(xIn, 138, 2) & "/" & Mid$(xIn, 140, 2) & "/" & Mid$(xIn, 142, 2)
                            curX = CCur(Mid$(xIn, 144, 12) / 100)
                            ' Impression avis ligne effet
                            prtSAB_SITTE003P1_Detail
                       End If
                    End If
                End If
   End Select
Loop
Close

prtSAB_Spécial_Close

End Sub

Public Sub prtSAB_ECHEDI04P1(lFileName_Input As String, lFileName_Output As String)
Dim mResponsable As String, wResponsable As String, blnOk As Boolean, blnRupture As Boolean
Dim X As String, xIN_Rupture As String
On Error GoTo Error_Handler
' >>>>>  Lecture du fichier spool pour découper les zones à imprimer
blnOk = False
blnRupture = False
mResponsable = ""
wResponsable = ""

Open lFileName_Input For Input As #1
Call FEU_ROUGE
Open lFileName_Output For Output As #2
Line Input #1, xIn
Print #2, xIn

Do Until EOF(1)
    Line Input #1, xIn
    If InStr(1, xIn, "0                                Responsable :") > 0 Then wResponsable = xIn
    If blnRupture Then
        blnRupture = False
'2015-10-01 jpl         If Mid$(xIn, 5, 5) = "    5" Then
        If Mid$(xIn, 5, 5) >= "    1" And Mid$(xIn, 5, 5) <= "    8" Then
            blnOk = True
            
            X = wResponsable

            If mResponsable = wResponsable Then
                Mid$(X, 1, 4) = "   1"
            Else
                Mid$(X, 1, 4) = "001 "
            End If
            Print #2, X
            mResponsable = wResponsable
            Print #2, xIN_Rupture           '2015-10-01 JPL

        End If
   End If
    
    If InStr(1, xIn, "Code Selection :") > 0 Then
        blnRupture = True
        xIN_Rupture = xIn
    Else
        blnRupture = False              '2015-10-01 JPL
    End If
    
    If blnOk Then Print #2, xIn
    
If InStr(1, xIn, "________________") > 0 Then blnOk = False

Loop
Close
Call FEU_VERT

Exit Sub
'---------------------------------------------------------
Error_Handler:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "prtSAB_ECHEDI04P1 ")


End Sub


Public Sub prtSAB_SITTE003P1_Entete_Avis01()
mprtSAB_SITTE003P1_Entete_Avis = "prtSAB_SITTE003P1_Entete_Avis01"

' Saut de page ...
If BOO_EDI_AVIS_1ER = "N" Then frmElpPrt.prtNewPage
BOO_EDI_AVIS_1ER = "N"

XPrt.FontSize = 10: XPrt.FontBold = False

XPrt.CurrentX = prtMinX + 7800
XPrt.CurrentY = prtMinY + prtlineHeight * 4

XPrt.Print "Paris, le  " & dateImp10(DSys);

Call prtAdresse_Enveloppe(wZADRESS0)

XPrt.FontSize = 8: XPrt.FontBold = False

XPrt.CurrentX = prtMinMarge - 300

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 5

XPrt.Print "Date d'envoi : ";
XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 800: XPrt.Print wENVOI;
XPrt.FontBold = False
XPrt.CurrentX = prtMinMarge - 300
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "Date règlement : ";
XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 800: XPrt.Print wRGL;
XPrt.FontBold = False
XPrt.CurrentX = prtMinMarge - 300
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "Code banque  Code guichet  No de compte";
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
mCurrentY_RIB = XPrt.CurrentY
XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge - 300: XPrt.Print wRIB_BQ(wRIB_Nb);
XPrt.CurrentX = prtMinMarge + 650: XPrt.Print wRIB_GUI(wRIB_Nb);
XPrt.CurrentX = prtMinMarge + 1550: XPrt.Print wRIB_CPT(wRIB_Nb);
XPrt.FontBold = False

XPrt.FontSize = 10: XPrt.FontBold = True

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
Call frmElpPrt.prtTrame(prtMinMarge - 350, XPrt.CurrentY, prtMaxMarge + 500, XPrt.CurrentY + prtlineHeight, " ", 235)
XPrt.CurrentX = prtMinMarge - 300: XPrt.Print "Tireur ";
XPrt.CurrentX = prtMinMarge + 2500: XPrt.Print "Référence ";
XPrt.CurrentX = prtMinMarge + 3500: XPrt.Print "  (1)  ";
XPrt.CurrentX = prtMinMarge + 4000: XPrt.Print "No LCR ";
XPrt.CurrentX = prtMinMarge + 4900: XPrt.Print "Echéance ";
XPrt.CurrentX = prtMinMarge + 6300: XPrt.Print "Montant ";
XPrt.CurrentX = prtMinMarge + 7500: XPrt.Print "Montant refusé (2) ";
XPrt.CurrentX = prtMinMarge + 9500: XPrt.Print "  (3)  ";

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight

End Sub

Public Sub prtSAB_SITTE003P1_Detail()

prtSAB_SITTE003P1_NewLine

' Somme d'une remise en fonction des différents effets
wSOMME = wSOMME + curX

XPrt.FontSize = 10: XPrt.FontBold = False
'XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.5
XPrt.CurrentX = prtMinMarge - 300: XPrt.Print wTIREUR;
XPrt.CurrentX = prtMinMarge + 2700: XPrt.Print wREF;
Select Case wACC
    Case "A":  wACC = "1"
End Select
XPrt.CurrentX = prtMinMarge + 3650: XPrt.Print wACC;
XPrt.CurrentX = prtMinMarge + 4900: XPrt.Print wECH;
X = Format$(curX, "### ### ##0.00")
XPrt.CurrentX = prtMinMarge + 6900 - XPrt.TextWidth(X): XPrt.Print X;
XPrt.FontSize = 8
XPrt.CurrentY = XPrt.CurrentY + 40
XPrt.CurrentX = prtMinMarge + 7050: XPrt.Print "EUR";
XPrt.CurrentY = XPrt.CurrentY - 40
Call frmElpPrt.prtTrame(prtMinMarge + 7400, XPrt.CurrentY, prtMinMarge + 9200, XPrt.CurrentY + prtlineHeight, " ", 235)

End Sub

Public Sub prtSAB_SITTE003P1_Fin_Avis01()
If XPrt.CurrentY + prtlineHeight * 7 > prtMaxY Then
    prtSAB_SITTE003P1_Entete_Avis01
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
End If

XPrt.FontSize = 10: XPrt.FontBold = True
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
Call frmElpPrt.prtTrame(prtMinMarge + 5900, XPrt.CurrentY, prtMinMarge + 9200, XPrt.CurrentY + (prtlineHeight * 3), " ", 235)
XPrt.FontUnderline = True
XPrt.CurrentX = prtMinMarge + 7500: XPrt.Print "Refus Pour   EUR ";
XPrt.FontUnderline = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 4300: XPrt.Print "Total du relevé ";
X = Format$(wSOMME, "### ### ##0.00")
XPrt.CurrentX = prtMinMarge + 6900 - XPrt.TextWidth(X): XPrt.Print X;
XPrt.FontSize = 8
XPrt.CurrentY = XPrt.CurrentY + 40
XPrt.CurrentX = prtMinMarge + 7050: XPrt.Print "EUR";
XPrt.CurrentY = XPrt.CurrentY - 40

XPrt.FontBold = False

XPrt.FontSize = 10
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 4300: XPrt.Print "Bon à payer pour ";
Call frmElpPrt.prtTrame(prtMinMarge + 5900, XPrt.CurrentY, prtMinMarge + 9200, XPrt.CurrentY + prtlineHeight, " ", 235)
XPrt.CurrentX = prtMinMarge + 8600: XPrt.Print "EUR";

' >>>>>  Bas de page  1er avis  <<<<<
If XPrt.CurrentY + 2900 > prtMaxY Then
    prtSAB_SITTE003P1_Entete_Avis01
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 4
Else
    XPrt.CurrentY = prtMaxY - 2700
End If
XPrt.FontSize = 8: XPrt.FontBold = True

XPrt.CurrentX = prtMinMarge - 300
XPrt.Print "L'enregistrement au débit de votre compte du montant inscrit dans la case << TOTAL DU RELEVE >> équivaudra à la détention par vous-même des effets ";
XPrt.CurrentX = prtMinMarge - 300
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 0.8
XPrt.Print "acquittés.";
XPrt.CurrentX = prtMinMarge - 300
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 0.8
XPrt.Print "Le présent document serait nul de plein droit si pour une raison quelconque une ou plusieurs lettres de changes étaient impayées.";
XPrt.CurrentX = prtMinMarge - 300
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 0.8
XPrt.Print "     (1) 0: Non accepté - 1: Accepté - 2: billet à ordre ";
XPrt.CurrentX = prtMinMarge - 300
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 0.8
XPrt.Print "     (2) Pour un paiement partiel, veuillez également préciser le montant qui doit être rejeté ";
XPrt.CurrentX = prtMinMarge - 300
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 0.8
XPrt.Print "     (3) Motif de refus ";

XPrt.FontSize = 8: XPrt.FontBold = False
    
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 0.8
'XPrt.CurrentY = prtMaxY - 1500
Call frmElpPrt.prtTrame(prtMinMarge + 1450, XPrt.CurrentY, prtMinMarge + 1700, XPrt.CurrentY + (prtlineHeight * 0.8), " ", 235)
Call frmElpPrt.prtTrame(prtMinMarge + 5950, XPrt.CurrentY, prtMinMarge + 6200, XPrt.CurrentY + (prtlineHeight * 0.8), " ", 235)
XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print "13 ";
XPrt.CurrentX = prtMinMarge + 6000: XPrt.Print "72 ";
XPrt.FontBold = False
XPrt.CurrentX = prtMinMarge + 2000: XPrt.Print "Créance non identifiable ";
XPrt.CurrentX = prtMinMarge + 6500: XPrt.Print "Code acceptation omis ";

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 0.8
Call frmElpPrt.prtTrame(prtMinMarge + 1450, XPrt.CurrentY, prtMinMarge + 1700, XPrt.CurrentY + (prtlineHeight * 0.8), " ", 235)
Call frmElpPrt.prtTrame(prtMinMarge + 5950, XPrt.CurrentY, prtMinMarge + 6200, XPrt.CurrentY + (prtlineHeight * 0.8), " ", 235)
XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print "18 ";
XPrt.CurrentX = prtMinMarge + 6000: XPrt.Print "73 ";
XPrt.FontBold = False
XPrt.CurrentX = prtMinMarge + 2000: XPrt.Print "Emetteur non identifiable ";
XPrt.CurrentX = prtMinMarge + 6500: XPrt.Print "Montant contesté ";
    
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 0.8
Call frmElpPrt.prtTrame(prtMinMarge + 1450, XPrt.CurrentY, prtMinMarge + 1700, XPrt.CurrentY + (prtlineHeight * 0.8), " ", 235)
Call frmElpPrt.prtTrame(prtMinMarge + 5950, XPrt.CurrentY, prtMinMarge + 6200, XPrt.CurrentY + (prtlineHeight * 0.8), " ", 235)
XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print "19 ";
XPrt.CurrentX = prtMinMarge + 6000: XPrt.Print "74 ";
XPrt.FontBold = False
XPrt.CurrentX = prtMinMarge + 2000: XPrt.Print "Créance cédée autre BQ ";
XPrt.CurrentX = prtMinMarge + 6500: XPrt.Print "Date d'échéance contestée ";
    
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 0.8
Call frmElpPrt.prtTrame(prtMinMarge + 1450, XPrt.CurrentY, prtMinMarge + 1700, XPrt.CurrentY + (prtlineHeight * 0.8), " ", 235)
Call frmElpPrt.prtTrame(prtMinMarge + 5950, XPrt.CurrentY, prtMinMarge + 6200, XPrt.CurrentY + (prtlineHeight * 0.8), " ", 235)
XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print "39 ";
XPrt.CurrentX = prtMinMarge + 6000: XPrt.Print "75 ";
XPrt.FontBold = False
XPrt.CurrentX = prtMinMarge + 2000: XPrt.Print "Ne paie qu'Effets acceptés ";
XPrt.CurrentX = prtMinMarge + 6500: XPrt.Print "Demande prorogation ";
    
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 0.8
Call frmElpPrt.prtTrame(prtMinMarge + 1450, XPrt.CurrentY, prtMinMarge + 1700, XPrt.CurrentY + (prtlineHeight * 0.8), " ", 235)
Call frmElpPrt.prtTrame(prtMinMarge + 5950, XPrt.CurrentY, prtMinMarge + 6200, XPrt.CurrentY + (prtlineHeight * 0.8), " ", 235)
XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print "70 ";
XPrt.CurrentX = prtMinMarge + 6000: XPrt.Print "76 ";
XPrt.FontBold = False
XPrt.CurrentX = prtMinMarge + 2000: XPrt.Print "Tirage contesté ";
XPrt.CurrentX = prtMinMarge + 6500: XPrt.Print "Réclamation tardive ";
    
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 0.8
Call frmElpPrt.prtTrame(prtMinMarge + 1450, XPrt.CurrentY, prtMinMarge + 1700, XPrt.CurrentY + (prtlineHeight * 0.8), " ", 235)
Call frmElpPrt.prtTrame(prtMinMarge + 5950, XPrt.CurrentY, prtMinMarge + 6200, XPrt.CurrentY + (prtlineHeight * 0.8), " ", 235)
XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 1500: XPrt.Print "71 ";
XPrt.CurrentX = prtMinMarge + 6000: XPrt.Print "90 ";
XPrt.FontBold = False
XPrt.CurrentX = prtMinMarge + 2000: XPrt.Print "Déjà réglée ";
XPrt.CurrentX = prtMinMarge + 6500: XPrt.Print "Paiement partiel ";

' Code Guichet... de l'entête

'XPrt.CurrentY = mCurrentY_RIB    '''prtMaxY -10300
'XPrt.FontSize = 8: XPrt.FontBold = True
'XPrt.CurrentX = prtMinMarge - 300: XPrt.Print wRIB_BQ;
'XPrt.CurrentX = prtMinMarge + 650: XPrt.Print wRIB_GUI;
'XPrt.CurrentX = prtMinMarge + 1550: XPrt.Print wRIB_CPT;
XPrt.FontBold = False

End Sub

Public Sub prtSAB_SITTE003P1_Entete_Avis02()

' >>>>>  Début 2ème avis  <<<<<
mprtSAB_SITTE003P1_Entete_Avis = "prtSAB_SITTE003P1_Entete_Avis02"

' Saut de page ...
If BOO_EDI_AVIS_1ER = "N" Then frmElpPrt.prtNewPage
BOO_EDI_AVIS_1ER = "N"

XPrt.FontSize = 10: XPrt.FontBold = False

XPrt.CurrentX = prtMinX + 7800
XPrt.CurrentY = prtMinY + prtlineHeight * 4

XPrt.Print "Paris, le  " & dateImp10(DSys);

' Adresse BIA en fixe

rsZADRESS0_Init wZADRESS0

wZADRESS0.ADRESSRA1 = paramSOC_RS
wZADRESS0.ADRESSRA2 = "        Service GDMP"
wZADRESS0.ADRESSAD1 = "."
wZADRESS0.ADRESSAD2 = paramSOC_Adresse
wZADRESS0.ADRESSCOP = "" '"75008"
wZADRESS0.ADRESSVIL = paramSOC_Ville

Call prtAdresse_Enveloppe(wZADRESS0)

XPrt.FontSize = 11: XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge - 300
XPrt.CurrentY = XPrt.CurrentY - prtlineHeight * 2
XPrt.Print "EXEMPLAIRE A RETOURNER A LA BANQUE ";

XPrt.FontSize = 8: XPrt.FontBold = False

XPrt.CurrentX = prtMinMarge - 300
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 8
XPrt.Print "Date règlement : ";
XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 800: XPrt.Print wRGL;
XPrt.FontBold = False
XPrt.CurrentX = prtMinMarge - 300
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "Code banque  Code guichet  No de compte";
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
mCurrentY_RIB = XPrt.CurrentY
XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge - 300: XPrt.Print wRIB_BQ(wRIB_Nb);
XPrt.CurrentX = prtMinMarge + 650: XPrt.Print wRIB_GUI(wRIB_Nb);
XPrt.CurrentX = prtMinMarge + 1550: XPrt.Print wRIB_CPT(wRIB_Nb);
XPrt.FontBold = False

XPrt.FontSize = 10: XPrt.FontBold = True

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
Call frmElpPrt.prtTrame(prtMinMarge - 350, XPrt.CurrentY, prtMaxMarge + 500, XPrt.CurrentY + prtlineHeight, " ", 235)
XPrt.CurrentX = prtMinMarge - 300: XPrt.Print "Tireur ";
XPrt.CurrentX = prtMinMarge + 2500: XPrt.Print "Référence ";
XPrt.CurrentX = prtMinMarge + 3500: XPrt.Print "  (1)  ";
XPrt.CurrentX = prtMinMarge + 4000: XPrt.Print "No LCR ";
XPrt.CurrentX = prtMinMarge + 4900: XPrt.Print "Echéance ";
XPrt.CurrentX = prtMinMarge + 6300: XPrt.Print "Montant ";
XPrt.CurrentX = prtMinMarge + 7500: XPrt.Print "Montant refusé (2) ";
XPrt.CurrentX = prtMinMarge + 9500: XPrt.Print "  (3)  ";

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight

End Sub

Public Sub prtSAB_SITTE003P1_Fin_Avis02()
'8
' >>>>>  Bas de page  2ème avis  <<<<<
If XPrt.CurrentY + prtlineHeight * 7 > prtMaxY Then
    prtSAB_SITTE003P1_Entete_Avis02
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
End If

XPrt.FontSize = 10: XPrt.FontBold = True
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
Call frmElpPrt.prtTrame(prtMinMarge + 5900, XPrt.CurrentY, prtMinMarge + 9200, XPrt.CurrentY + (prtlineHeight * 3), " ", 235)
XPrt.FontUnderline = True
XPrt.CurrentX = prtMinMarge + 7500: XPrt.Print "Refus Pour   EUR ";
XPrt.FontUnderline = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 4300: XPrt.Print "Total du relevé ";
X = Format$(wSOMME, "### ### ##0.00")
XPrt.CurrentX = prtMinMarge + 6900 - XPrt.TextWidth(X): XPrt.Print X;
XPrt.FontSize = 8
XPrt.CurrentY = XPrt.CurrentY + 40
XPrt.CurrentX = prtMinMarge + 7050: XPrt.Print "EUR";
XPrt.CurrentY = XPrt.CurrentY - 40

XPrt.FontBold = False

XPrt.FontSize = 10
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.FontBold = True
XPrt.CurrentX = prtMinMarge + 4300: XPrt.Print "Bon à payer pour ";
Call frmElpPrt.prtTrame(prtMinMarge + 5900, XPrt.CurrentY, prtMinMarge + 9200, XPrt.CurrentY + prtlineHeight, " ", 235)
XPrt.CurrentX = prtMinMarge + 8600: XPrt.Print "EUR";

If XPrt.CurrentY + 2900 > prtMaxY Then
    prtSAB_SITTE003P1_Entete_Avis02
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 4
Else
    If XPrt.CurrentY < prtMaxY - 5000 Then
        XPrt.CurrentY = prtMaxY - 3900
    Else
        XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
    End If
End If

XPrt.FontSize = 10: XPrt.FontBold = False
XPrt.CurrentX = prtMinMarge - 300

XPrt.Print "Nous vous donnons ordre de payer par le débit de notre compte, rappelé ci-dessus les effets décrits sur le présent relevé qui ne font pas ";
XPrt.CurrentX = prtMinMarge - 300
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Print "l'objet d'un refus de paiement pour le motif et la somme indiquée dans les colonnes ci-dessus. ";
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.CurrentX = prtMinMarge + 3700: XPrt.Print "Date ";
XPrt.CurrentX = prtMinMarge + 4900: XPrt.Print "Signature";
    
XPrt.FontSize = 8: XPrt.FontBold = True
    
XPrt.CurrentX = prtMinMarge - 300
'XPrt.CurrentY = prtMaxY - 1400
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.Print "Après avoir complété ce relevé, veuillez dater, signer et porter au bas de ce feuillet, le montant total à payer et le montant total refusé, l'ensemble formant ";
XPrt.CurrentX = prtMinMarge - 300
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 0.8
XPrt.Print "le total du relevé. Séparer les deux volets :";
XPrt.CurrentX = prtMinMarge - 300
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 0.8
XPrt.Print "    - 1er exemplaire à renvoyer à l'adresse ci-dessus avant la date limite indiquée, ";
XPrt.CurrentX = prtMinMarge - 300
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 0.8
XPrt.Print "    - 2ème exemplaire à conserver. ";
XPrt.CurrentX = prtMinMarge - 300
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 0.8
XPrt.Print "Sans réponse de votre part pour la date de règlement comme à défaut de provision, ces effets seront impayés en tout ou partie. Salutations distinguées. ";

XPrt.FontBold = False

' Code Guichet... de l'entête

'XPrt.CurrentY = mCurrentY_RIB   ''' prtMaxY - 10800
'XPrt.FontSize = 8: XPrt.FontBold = True
'XPrt.CurrentX = prtMinMarge - 300: XPrt.Print wRIB_BQ;
'XPrt.CurrentX = prtMinMarge + 650: XPrt.Print wRIB_GUI;
'XPrt.CurrentX = prtMinMarge + 1550: XPrt.Print wRIB_CPT;
'XPrt.FontBold = False

End Sub

