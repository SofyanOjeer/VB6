Attribute VB_Name = "prtCompteSolde"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------

Private recCompte As typeCompte
Dim I As Integer, Height8_6 As Integer
Dim K As Integer, K1 As Integer, K2 As Integer, arrCompte_K As Integer

Private Mt As Currency

Dim Col1 As Integer, Col1_X As Integer, Col2 As Integer, Col3 As Integer, Col2_X As Integer
Dim Col2_Db As Integer, Col2_Cr As Integer
Dim Col3_Db As Integer, Col3_Cr As Integer
Dim blnPrintSoldé As Boolean, blnPrintdécouvert As Boolean
Dim blnCompteSoldeLine As Boolean
Dim CompteSolde_LineNb As Integer, CompteSoldeRupture_LineNb As Integer
Dim arrCV2_Montant() As Currency, arrCV1_DeviseIso() As String * 3

Type typeCompteSoldeRupture
    IdKey       As String * 30
    Intitulé    As String * 40
    IndexDeb    As Integer
    IndexFin    As Integer
    DeviseN     As String * 3
    DeviseIso   As String * 3
    Db          As Currency
    Cr          As Currency
    Db_Cv       As Currency
    Cr_Cv       As Currency
    Nb          As Integer
    
End Type
    
Dim arrCompteSoldeRupture() As typeCompteSoldeRupture
Dim arrCompteSoldeRupture_Nb As Integer, arrCompteSoldeRupture_Index As Integer
Dim blnCompteSoldeRupture As Boolean, blnCompteSoldeRuptureRacine As Boolean


Type typeCompteSoldeTotal
    IdKey       As String * 10
    ''BiaTyp      As String * 3
    DeviseN     As String * 3
    DeviseIso   As String * 3
    Db          As Currency
    Cr          As Currency
    Db_Cv       As Currency
    Cr_Cv       As Currency
    Nb          As Integer
    
End Type
    
Dim arrCompteSoldeTotal() As typeCompteSoldeTotal
Dim arrCompteSoldeTotal_Nb As Integer, arrCompteSoldeTotal_Index As Integer, arrCompteSoldeTotal_NbMax As Integer
Dim blnCompteSoldeTotal As Boolean
Dim blnCompteSoldeBilan As Boolean

Dim CompteSoldeRupture_Len As Integer
Dim optEtat As String * 1, optEtatSortK As String * 2
Dim X10 As String * 10
Dim arrDevTotal() As typeCompteSoldeTotal
Dim arrDevTotalEur() As typeCompteSoldeTotal


'---------------------------------------------------------
Public Sub prtCompteSolde_Monitor(Msg As String)
'---------------------------------------------------------
blnCompteSoldeTotal = False
prtCompteSolde_Open Msg, "", ""
prtCompteSolde_Print Msg
prtCompteSolde_Close
End Sub
'---------------------------------------------------------
Public Sub prtCompteSolde_Open(Msg As String, xEnTete As String, xDestinataire As String)
'---------------------------------------------------------
Dim X As String, optAmj As String * 8

On Error GoTo prtError

Set XPrt = Printer
Height8_6 = frmElpPrt.prtHeightDelta(8, 6)
optAmj = mId$(Msg, 19, 8)

blnCompteSoldeLine = IIf(mId$(Msg, 29, 1) = "L", True, False)
blnCompteSoldeRupture = IIf(mId$(Msg, 30, 1) = "R", True, False)
blnCompteSoldeRuptureRacine = IIf(mId$(Msg, 34, 1) = "R", True, False)
blnCompteSoldeTotal = IIf(mId$(Msg, 31, 1) = "T", True, False)
blnPrintdécouvert = False
If xEnTete = "" Then xEnTete = "Etat des soldes"
Select Case mId$(Msg, 18, 1)
    Case "I": X = xEnTete & " ( ": blnPrintdécouvert = True
    Case "V": X = xEnTete & " (  ": blnPrintdécouvert = True
    Case "M": X = xEnTete & " ( fin de mois en date de traitement : "
    Case "O": X = xEnTete & " ( fin de mois en date d'opération : "
    Case "A": X = xEnTete & " ( fin d'année : "
    Case Else: X = xEnTete & "solde ?"
End Select
prtTitleText = X & dateImp(optAmj) & " ) "

If mId$(Msg, 27, 1) = "S" Then
    blnPrintSoldé = True
Else
    blnPrintSoldé = False
End If

optEtatSortK = mId$(Msg, 32, 2)
optEtat = mId$(Msg, 32, 1)

Select Case optEtatSortK
    Case "A1":
                If blnCompteSoldeRuptureRacine Then
                    CompteSoldeRupture_Len = 5: blnCompteSoldeRupture = True
                Else
                    CompteSoldeRupture_Len = 8
                End If
    Case "A2": CompteSoldeRupture_Len = 5
    Case "A3": CompteSoldeRupture_Len = 8
    Case "A4": CompteSoldeRupture_Len = 7
    Case "A5": CompteSoldeRupture_Len = 3
    Case "A6": CompteSoldeRupture_Len = 5
    Case "G1": CompteSoldeRupture_Len = 11
    Case "G2": CompteSoldeRupture_Len = 11
    Case "G3": CompteSoldeRupture_Len = 14
    Case "G4": CompteSoldeRupture_Len = 15
End Select

Call prtCompteSolde_CV_Init(optAmj, mId$(Msg, 14, 3))

prtLineNb = 1

frmElpPrt.Show vbModeless


prtOrientation = vbPRORLandscape
prtPgmName = "prtCompteSolde"

If xDestinataire <> "" Then
    prtTitleUsr = xDestinataire
Else
    prtTitleUsr = usrName
End If

prtlineHeight = 250
prtHeaderHeight = 300

frmElpPrt.prtStdInit

If mId$(Msg, 28, 1) = ">" Then prtMinX = prtMinX + 500 'reliure
Col1 = prtMinX + 1600: Col1_X = Col1 + 50
Col3_Cr = prtMaxX - 50: Col3_Db = Col3_Cr - 1500: Col3 = Col3_Db - 1500
Col2_Cr = Col3 - 350: Col2_Db = Col2_Cr - 1500: Col2 = Col2_Db - 1500
Col2_X = Col2 - 50
prtCompteSolde_Form

ReDim arrCompteSoldeTotal(100): arrCompteSoldeTotal_NbMax = 100: arrCompteSoldeTotal_Nb = 0
If optEtatSortK = "A6" Then prtCompteSolde_A6Init

Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide
End Sub

'---------------------------------------------------------
Public Sub prtCompteSolde_Form()
'---------------------------------------------------------
Dim X As String

XPrt.FontSize = 8
XPrt.FontBold = True
XPrt.DrawWidth = 3


Call frmElpPrt.prtTrame(prtMinX, prtMinY, prtMaxX, prtMinY + prtHeaderHeight, "B")
'Call frmElpPrt.prtTrame(prtMinX + 11100, prtMinY + prtHeaderHeight + 10, prtMinX + 13500, prtMaxY - 10, " ")


'---------------------------------------------------------

X = "N°de Compte"
XPrt.CurrentY = prtMinY + (prtHeaderHeight - XPrt.TextHeight(X)) / 2
XPrt.CurrentX = prtMinX + 300
XPrt.Print X;

XPrt.CurrentX = Col1_X
XPrt.Print "Intitulé";

If optEtatSortK = "A6" Then
    XPrt.CurrentX = Col2_Db - 400: XPrt.Print "Bilan  ( "; CV_X2.DeviseIso & " )";
    XPrt.CurrentX = Col3_Db - 400: XPrt.Print "Hors-Bilan  ( "; CV_X2.DeviseIso & " )";

Else

    XPrt.CurrentX = Col2_X - 1050
    XPrt.Print "Dernier Mvt le";
    
    XPrt.CurrentX = Col2_Db - 600: XPrt.Print "Débit";
    XPrt.CurrentX = Col2_Cr - 600: XPrt.Print "Crédit";
    
    XPrt.CurrentX = Col3_Db - 850: XPrt.Print CV_X2.DeviseIso & "  Débit ";
    XPrt.CurrentX = Col3_Cr - 850: XPrt.Print CV_X2.DeviseIso & "  Crédit ";
End If

If blnCompteSoldeRupture Then
    XPrt.CurrentY = prtMinY + prtHeaderHeight + 20 - prtlineHeight * 0.5
Else
    XPrt.CurrentY = prtMinY + prtHeaderHeight + 20 - prtlineHeight * 1
End If

CompteSolde_LineNb = 0
CompteSoldeRupture_LineNb = 0
End Sub




'---------------------------------------------------------
Public Sub prtCompteSolde_Line()
'---------------------------------------------------------
Dim X As String, wsdCurrentX As Integer, eurCurrentX As Integer
Dim Situation As String

If XPrt.CurrentY + prtlineHeight * 1.6 > prtMaxY Then
    prtCompteSolde_Trait
    frmElpPrt.prtNewPage
    prtCompteSolde_Form
End If

XPrt.FontBold = False
'------------------------------------------ligne 1--------------
'prtCurrentY = XPrt.CurrentY + prtlineHeight
CompteSolde_LineNb = CompteSolde_LineNb + 1
'If CompteSolde_LineNb = 4 Then CompteSolde_LineNb = 1: prtCurrentY = prtCurrentY + prtlineHeight * 0.5 'Call frmElpPrt.prtTrame(prtMinX, prtCurrentY - 10, prtMaxX, prtCurrentY + 10, " ", 240)
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
If CompteSolde_LineNb > 2 Then
    
    Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY, prtMaxX, XPrt.CurrentY + prtlineHeight - 10, " ", 245)
'Else
    If CompteSolde_LineNb = 4 Then CompteSolde_LineNb = 0
End If

'XPrt.CurrentY = prtCurrentY
'XPrt.FontSize = 8
XPrt.FontSize = 6
XPrt.CurrentY = XPrt.CurrentY + Height8_6


XPrt.CurrentX = prtMinX + 50
XPrt.Print Compte_Imp(recCompte.Numéro);
XPrt.CurrentX = Col1 - 600
XPrt.Print Format$(recCompte.Devise, "000") & "."; recCompte.Devisex;
XPrt.CurrentX = Col1_X
If Not blnCompteSoldeRupture Or optEtat = "G" Or optEtatSortK = "A4" Or optEtatSortK = "A5" Then
    If recCompte.Groupe <> "000000" Then
        XPrt.FontItalic = True: XPrt.FontUnderline = True
        XPrt.Print "Groupe : " & recCompte.Groupe & " ";
        XPrt.FontItalic = False: XPrt.FontUnderline = False
    End If
    XPrt.Print Trim(recCompte.Intitulé);
    If Trim(recCompte.Intitulé2) <> "" Then XPrt.Print " _ ";
End If

XPrt.Print Trim(recCompte.Intitulé2);
If recCompte.TypeGA = "A" Then XPrt.Print " " & Trim(DicLib(13, recCompte.BiaTyp));
XPrt.FontBold = True
Select Case recCompte.Situation
    Case " ": Situation = ""
    Case "A": Situation = " **Annulé**"
    Case "B": Situation = " **Bloqué**"
    Case Else: Situation = " ?? " & recCompte.Situation
End Select

XPrt.Print Situation;

XPrt.FontBold = False

X = dateImp(recCompte.MvtAmj)
XPrt.CurrentX = Col2_X - XPrt.TextWidth(X): wsdCurrentX = XPrt.CurrentX
XPrt.Print X;

If blnPrintdécouvert And recCompte.DécouvertMontant > 0 Then
    XPrt.FontUnderline = True
    X = "Découvert autorisé : " & Format$(recCompte.DécouvertMontant, "### ### ### ###") _
    & "  juqu'au : " & dateImp(recCompte.DécouvertAmj)
    If Val(recCompte.DécouvertAmj) < DSys Then: X = X & " !!! "
    If recCompte.SoldeInstantané + recCompte.DécouvertMontant < 0 Then XPrt.FontBold = True: X = X & " ????? "
    XPrt.CurrentX = wsdCurrentX - XPrt.TextWidth(X) - 200
    XPrt.Print X;
    XPrt.FontUnderline = False
End If
XPrt.FontBold = False
       
X = Format$(Abs(recCompte.SoldeInstantané), "#### ### ### ### ##0.00")
If recCompte.SoldeInstantané >= 0 Then
    wsdCurrentX = Col2_Cr: eurCurrentX = Col3_Cr
Else
    wsdCurrentX = Col2_Db: eurCurrentX = Col3_Db
End If

XPrt.CurrentX = wsdCurrentX - XPrt.TextWidth(X)

XPrt.Print X & " ";
XPrt.Print arrCV1_DeviseIso(arrCompte_K);

X = Format$(Abs(arrCV2_Montant(arrCompte_K)), "#### ### ### ### ##0.00")
XPrt.CurrentX = eurCurrentX - XPrt.TextWidth(X)
XPrt.Print X;
        
XPrt.FontBold = False
XPrt.FontSize = 8
XPrt.CurrentY = XPrt.CurrentY - Height8_6

End Sub

'---------------------------------------------------------
Public Sub prtCompteSoldeTotal_Line()
'---------------------------------------------------------
Dim X As String, wsdCurrentX As Integer, eurCurrentX As Integer
Dim kBilan As String, kBiaTyp As String, kDeviseIso As String, iTrame As Integer, kPays As String
Dim kRupture As String

If XPrt.CurrentY + prtlineHeight * 2 > prtMaxY Then
    prtCompteSolde_Trait
    frmElpPrt.prtNewPage
    prtCompteSolde_Form
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 0.5
End If

XPrt.FontBold = False
'------------------------------------------ligne 1--------------

X10 = arrCompteSoldeTotal(arrCompteSoldeTotal_Index).IdKey
iTrame = 0
kRupture = optEtat

Select Case optEtatSortK
    Case "A3", "G3"
                    kBilan = mId$(X10, 4, 1)
                    kBiaTyp = mId$(X10, 5, 3)
                    kDeviseIso = mId$(X10, 1, 3)
                    If kBiaTyp = "000" Then
                        iTrame = IIf(kBilan = "0", 220, 235)
                    End If
      Case "A4", "G4"
                    kBilan = mId$(X10, 1, 1)
                    kBiaTyp = mId$(X10, 6, 3)
                    kPays = mId$(X10, 2, 4)
                    If kBiaTyp = "000" Then iTrame = 220
                    If kBiaTyp = "   " Then
                        kRupture = "P"
                        iTrame = 235
                    End If
                    arrCompteSoldeTotal(arrCompteSoldeTotal_Index).Db = 0
                    arrCompteSoldeTotal(arrCompteSoldeTotal_Index).Cr = 0
               
   Case Else
                    kBilan = mId$(X10, 1, 1)
                    kBiaTyp = mId$(X10, 2, 3)
                    kDeviseIso = mId$(X10, 5, 3)
                    If kDeviseIso = "   " Then
                        iTrame = IIf(kBiaTyp = "000", 220, 235)
                    End If
End Select

prtCurrentY = XPrt.CurrentY + prtlineHeight
    
If iTrame > 0 Then
    CompteSolde_LineNb = 0 '-1
    Call frmElpPrt.prtTrame(prtMinX, prtCurrentY, prtMaxX, prtCurrentY + prtlineHeight - 10, " ", iTrame)
    XPrt.FontBold = True
Else
    XPrt.FontBold = False
End If

'CompteSolde_LineNb = CompteSolde_LineNb + 1
'If CompteSolde_LineNb = 4 Then CompteSolde_LineNb = 1: prtCurrentY = prtCurrentY + prtlineHeight * 0.5 'Call frmElpPrt.prtTrame(prtMinX, prtCurrentY - 10, prtMaxX, prtCurrentY + 10, " ", 240)

XPrt.CurrentY = prtCurrentY
XPrt.FontSize = 6
XPrt.CurrentY = XPrt.CurrentY + Height8_6

Select Case kBiaTyp
    Case "000"
            
            Select Case kBilan
                Case "B": XPrt.CurrentX = Col1_X - 400: XPrt.Print "Bilan";
                Case "H": XPrt.CurrentX = Col1_X - 700: XPrt.Print "Hors-Bilan";
            End Select
    Case Else
    
            XPrt.CurrentX = Col1_X - 400
            XPrt.Print "  " & kDeviseIso;
            XPrt.CurrentX = Col1_X
            Select Case kRupture
            
                Case "A": XPrt.Print Trim(DicLib(13, kBiaTyp));
                Case "P": XPrt.Print Trim(DicLib(919, kPays));
                Case Else
                    Select Case kBiaTyp
                        Case "1  ": XPrt.Print "Classe 1";
                        Case "2  ": XPrt.Print "Classe 2";
                        Case "3  ": XPrt.Print "Classe 3";
                        Case "4  ": XPrt.Print "Classe 4";
                        Case "5  ": XPrt.Print "Classe 5";
                        Case "6  ": XPrt.Print "Classe 6";
                        Case "7  ": XPrt.Print "Classe 7";
                        Case "8  ": XPrt.Print "Classe 8";
                        Case "9  ": XPrt.Print "Classe 9";
                    End Select
            End Select
            
            'XPrt.Print "  " & arrCompteSoldeTotal(arrCompteSoldeTotal_Index).IdKey;
End Select

If arrCompteSoldeTotal(arrCompteSoldeTotal_Index).Db <> 0 Then
    X = Format$(Abs(arrCompteSoldeTotal(arrCompteSoldeTotal_Index).Db), "#### ### ### ### ##0.00")
    XPrt.CurrentX = Col2_Db - XPrt.TextWidth(X)
    XPrt.Print X;
    XPrt.Print " " & arrCompteSoldeTotal(arrCompteSoldeTotal_Index).DeviseIso;
End If

If arrCompteSoldeTotal(arrCompteSoldeTotal_Index).Cr <> 0 Then
    X = Format$(Abs(arrCompteSoldeTotal(arrCompteSoldeTotal_Index).Cr), "#### ### ### ### ##0.00")
    XPrt.CurrentX = Col2_Cr - XPrt.TextWidth(X)
    XPrt.Print X;
    XPrt.Print " " & arrCompteSoldeTotal(arrCompteSoldeTotal_Index).DeviseIso;
End If
      
If arrCompteSoldeTotal(arrCompteSoldeTotal_Index).Db_Cv <> 0 Then
    X = Format$(Abs(arrCompteSoldeTotal(arrCompteSoldeTotal_Index).Db_Cv), "#### ### ### ### ##0.00")
    XPrt.CurrentX = Col3_Db - XPrt.TextWidth(X)
    XPrt.Print X;
End If
    
If arrCompteSoldeTotal(arrCompteSoldeTotal_Index).Cr_Cv <> 0 Then
    X = Format$(Abs(arrCompteSoldeTotal(arrCompteSoldeTotal_Index).Cr_Cv), "#### ### ### ### ##0.00")
    XPrt.CurrentX = Col3_Cr - XPrt.TextWidth(X)
    XPrt.Print X;
End If
XPrt.FontBold = False
XPrt.FontSize = 8
XPrt.CurrentY = XPrt.CurrentY - Height8_6

End Sub


'---------------------------------------------------------
Public Sub prtCompteSoldeRupture_Line()
'---------------------------------------------------------
Dim X As String, XIntitulé As String, X3 As String


If XPrt.CurrentY + prtlineHeight * 3 > prtMaxY Then
    prtCompteSolde_Trait
    frmElpPrt.prtNewPage
    prtCompteSolde_Form
End If

'------------------------------------------ligne 1--------------
If blnCompteSoldeLine Then
    prtCurrentY = XPrt.CurrentY + prtlineHeight * 1.25
    Call frmElpPrt.prtTrame(prtMinX, prtCurrentY, prtMaxX, prtCurrentY + prtlineHeight - 10, " ", 220)
    XPrt.CurrentY = prtCurrentY
Else
    CompteSoldeRupture_LineNb = CompteSoldeRupture_LineNb + 1
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'    If CompteSoldeRupture_LineNb = 3 Then
'        CompteSoldeRupture_LineNb = 0
'        Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY, prtMaxX, XPrt.CurrentY + prtlineHeight - 10, " ", 250)
    If CompteSoldeRupture_LineNb > 2 Then
        
        Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY, prtMaxX, XPrt.CurrentY + prtlineHeight - 10, " ", 245)
        'XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
 '   Else
        If CompteSoldeRupture_LineNb = 4 Then CompteSoldeRupture_LineNb = 0
        'XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    End If
End If

CompteSolde_LineNb = 0

XPrt.FontSize = 8

XPrt.CurrentX = prtMinX + 50
X = ""
XIntitulé = Trim(arrCompteSoldeRupture(arrCompteSoldeRupture_Index).Intitulé) & " _ "

Select Case optEtatSortK
    Case "A1": XPrt.Print mId$(arrCompteSoldeRupture(arrCompteSoldeRupture_Index).IdKey, 1, 5) & "." & mId$(arrCompteSoldeRupture(arrCompteSoldeRupture_Index).IdKey, 6, 3);
                X3 = mId$(arrCompteSoldeRupture(arrCompteSoldeRupture_Index).IdKey, 6, 3)
                X = Trim(DicLib(13, X3))
    Case "A2": XPrt.Print mId$(arrCompteSoldeRupture(arrCompteSoldeRupture_Index).IdKey, 1, 5);
    Case "A3": XPrt.Print mId$(arrCompteSoldeRupture(arrCompteSoldeRupture_Index).IdKey, 1, 3) & "." & mId$(arrCompteSoldeRupture(arrCompteSoldeRupture_Index).IdKey, 4, 5) & "." & mId$(arrCompteSoldeRupture(arrCompteSoldeRupture_Index).IdKey, 9, 3);
                X3 = mId$(arrCompteSoldeRupture(arrCompteSoldeRupture_Index).IdKey, 9, 3)
                X = Trim(DicLib(13, X3))
    Case "A4": XPrt.Print mId$(arrCompteSoldeRupture(arrCompteSoldeRupture_Index).IdKey, 1, 4) & "." & mId$(arrCompteSoldeRupture(arrCompteSoldeRupture_Index).IdKey, 5, 3);
                arrCompteSoldeRupture(arrCompteSoldeRupture_Index).Intitulé = ""
                X = Trim(DicLib(919, mId$(arrCompteSoldeRupture(arrCompteSoldeRupture_Index).IdKey, 1, 4))) _
                  & " _ " & Trim(DicLib(13, mId$(arrCompteSoldeRupture(arrCompteSoldeRupture_Index).IdKey, 5, 3)))
    Case "A5":  X3 = mId$(arrCompteSoldeRupture(arrCompteSoldeRupture_Index).IdKey, 1, 3)
                XPrt.Print X3;
                XIntitulé = "": X = Trim(DicLib(13, X3))
    Case "A6": XPrt.Print mId$(arrCompteSoldeRupture(arrCompteSoldeRupture_Index).IdKey, 1, 5);
    Case "G1": XPrt.Print Format$(mId$(arrCompteSoldeRupture(arrCompteSoldeRupture_Index).IdKey, 4, 8), "000 000 00");
    Case "G2": XPrt.Print Format$(mId$(arrCompteSoldeRupture(arrCompteSoldeRupture_Index).IdKey, 4, 8), "000 000 00");
    Case "G3": XPrt.Print mId$(arrCompteSoldeRupture(arrCompteSoldeRupture_Index).IdKey, 1, 3) & "." & Format$(mId$(arrCompteSoldeRupture(arrCompteSoldeRupture_Index).IdKey, 7, 8), "000 000 00");
    Case "G4": XPrt.Print mId$(arrCompteSoldeRupture(arrCompteSoldeRupture_Index).IdKey, 1, 4);
                arrCompteSoldeRupture(arrCompteSoldeRupture_Index).Intitulé = ""
                X = Trim(DicLib(919, mId$(arrCompteSoldeRupture(arrCompteSoldeRupture_Index).IdKey, 1, 4)))
    Case Else: XPrt.Print arrCompteSoldeRupture(arrCompteSoldeRupture_Index).IdKey;
End Select

XPrt.FontSize = 6
XPrt.CurrentY = XPrt.CurrentY + Height8_6
XPrt.CurrentX = Col1_X
XPrt.Print XIntitulé & X;
       
If optEtatSortK = "A3" Or optEtatSortK = "G3" Then
    XPrt.FontBold = XPrt.FontBold = False
    If arrCompteSoldeRupture(arrCompteSoldeRupture_Index).Cr <> 0 Then
        X = Format$(arrCompteSoldeRupture(arrCompteSoldeRupture_Index).Cr, "#### ### ### ### ##0.00")
        XPrt.CurrentX = Col2_Cr - XPrt.TextWidth(X)
        XPrt.Print X & " ";
        XPrt.Print arrCompteSoldeRupture(arrCompteSoldeRupture_Index).DeviseIso;
    End If
    
    If arrCompteSoldeRupture(arrCompteSoldeRupture_Index).Db <> 0 Then
        X = Format$(Abs(arrCompteSoldeRupture(arrCompteSoldeRupture_Index).Db), "#### ### ### ### ##0.00")
        XPrt.CurrentX = Col2_Db - XPrt.TextWidth(X)
        XPrt.Print X & " ";
        XPrt.Print arrCompteSoldeRupture(arrCompteSoldeRupture_Index).DeviseIso;
    End If
End If

XPrt.FontBold = True

If arrCompteSoldeRupture(arrCompteSoldeRupture_Index).Db_Cv <> 0 Then
    X = Format$(Abs(arrCompteSoldeRupture(arrCompteSoldeRupture_Index).Db_Cv), "#### ### ### ### ##0.00")
    XPrt.CurrentX = Col3_Db - XPrt.TextWidth(X)
    XPrt.Print X;
End If

If arrCompteSoldeRupture(arrCompteSoldeRupture_Index).Cr_Cv <> 0 Then
    X = Format$(arrCompteSoldeRupture(arrCompteSoldeRupture_Index).Cr_Cv, "#### ### ### ### ##0.00")
    XPrt.CurrentX = Col3_Cr - XPrt.TextWidth(X)
    XPrt.Print X;
End If

XPrt.FontSize = 8
XPrt.CurrentY = XPrt.CurrentY - Height8_6
XPrt.FontBold = False

End Sub

Public Sub prtCompteSolde_Trait()
XPrt.DrawWidth = 2

XPrt.Line (Col1, prtMinY)-(Col1, prtMaxY)
XPrt.Line (Col2, prtMinY)-(Col2, prtMaxY)
XPrt.Line (Col3, prtMinY)-(Col3, prtMaxY)

End Sub

Public Sub prtCompteSoldeRupture_AddItem(K As Integer, IdKey As String)
arrCompteSoldeRupture_Nb = arrCompteSoldeRupture_Nb + 1
arrCompteSoldeRupture(arrCompteSoldeRupture_Nb).IdKey = IdKey
arrCompteSoldeRupture(arrCompteSoldeRupture_Nb).Intitulé = arrCompte(K).Intitulé
arrCompteSoldeRupture(arrCompteSoldeRupture_Nb).IndexDeb = K
arrCompteSoldeRupture(arrCompteSoldeRupture_Nb).IndexFin = K
arrCompteSoldeRupture(arrCompteSoldeRupture_Nb).Nb = 0
arrCompteSoldeRupture(arrCompteSoldeRupture_Nb).Db = 0
arrCompteSoldeRupture(arrCompteSoldeRupture_Nb).Cr = 0
arrCompteSoldeRupture(arrCompteSoldeRupture_Nb).Db_Cv = 0
arrCompteSoldeRupture(arrCompteSoldeRupture_Nb).Cr_Cv = 0

End Sub

Public Sub prtCompteSoldeRupture_Scan()
Dim X As String, wIdKey As String * 30

arrCompteSoldeRupture_Nb = 0
arrCompteSoldeRupture(0).IdKey = 0

For arrCompte_K = K1 To K2

    CV_X1.DeviseN = Format$(arrCompte(arrCompte_K).Devise, "000")
    CV_X1.DeviseIso = ""
    CV_X1.Montant = arrCompte(arrCompte_K).SoldeInstantané
    Call CV_Transitoire(CV_X1, CV_X2, CV_X3, X)
        
    arrCV2_Montant(arrCompte_K) = CV_X2.Montant
    arrCV1_DeviseIso(arrCompte_K) = CV_X1.DeviseIso
    X = arrCompte(arrCompte_K).LibTyp
    
    If blnCompteSoldeTotal Then
        X10 = prtCompteSoldeTotal_IdKey(X)
        Call prtCompteSoldeTotal_Add(CV_X1, CV_X2, X10)
    End If
    
    wIdKey = mId$(X, 1, CompteSoldeRupture_Len)
    If arrCompteSoldeRupture(arrCompteSoldeRupture_Nb).IdKey <> wIdKey Then Call prtCompteSoldeRupture_AddItem(arrCompte_K, wIdKey)
    
    arrCompteSoldeRupture(arrCompteSoldeRupture_Nb).IndexFin = arrCompte_K
    If CV_X1.Montant <> 0 Then
        arrCompteSoldeRupture(arrCompteSoldeRupture_Nb).DeviseN = CV_X1.DeviseN
        arrCompteSoldeRupture(arrCompteSoldeRupture_Nb).DeviseIso = CV_X1.DeviseIso
        arrCompteSoldeRupture(arrCompteSoldeRupture_Nb).Nb = arrCompteSoldeRupture(arrCompteSoldeRupture_Nb).Nb + 1
        If CV_X1.Montant < 0 Then
            arrCompteSoldeRupture(arrCompteSoldeRupture_Nb).Db = arrCompteSoldeRupture(arrCompteSoldeRupture_Nb).Db + CV_X1.Montant
            arrCompteSoldeRupture(arrCompteSoldeRupture_Nb).Db_Cv = arrCompteSoldeRupture(arrCompteSoldeRupture_Nb).Db_Cv + CV_X2.Montant
        Else
            arrCompteSoldeRupture(arrCompteSoldeRupture_Nb).Cr = arrCompteSoldeRupture(arrCompteSoldeRupture_Nb).Cr + CV_X1.Montant
            arrCompteSoldeRupture(arrCompteSoldeRupture_Nb).Cr_Cv = arrCompteSoldeRupture(arrCompteSoldeRupture_Nb).Cr_Cv + CV_X2.Montant
       End If
    End If
Next arrCompte_K

End Sub

Public Sub prtCompteSolde_A6()
Dim X As String, wIdKey As String * 30
Dim curBilanDb As Currency, curBilanCr As Currency
Dim curHBilanDb As Currency, curHBilanCr As Currency
Dim IDev As Integer, blnPrintLine As Boolean

curBilanDb = 0: curBilanCr = 0
curHBilanDb = 0: curHBilanCr = 0
blnPrintLine = False

'arrCompteSoldeRupture_Nb = 0
'arrCompteSoldeRupture(0).IdKey = 0

For arrCompte_K = K1 To K2

    CV_X1.Montant = arrCompte(arrCompte_K).SoldeInstantané
    If CV_X1.Montant <> 0 Then
        blnPrintLine = True
         CV_X1.DeviseN = arrCompte(arrCompte_K).Devise
         CV_X1.DeviseIso = ""
         Call CV_Transitoire(CV_X1, CV_X2, CV_X3, X)
             
         IDev = CInt(CV_X1.DeviseN)
         arrDevTotal(IDev).DeviseN = CV_X1.DeviseN
         arrDevTotal(IDev).Nb = arrDevTotal(IDev).Nb + 1
         X = mId$(arrCompte(arrCompte_K).LibTyp, 6, 1)
         If X = "B" Then
             If CV_X1.Montant < 0 Then
                 curBilanDb = curBilanDb + CV_X2.Montant
                 arrDevTotal(IDev).Db = arrDevTotal(IDev).Db + CV_X1.Montant
                  arrDevTotal(0).Db = arrDevTotal(0).Db + CV_X2.Montant
           Else
                 curBilanCr = curBilanCr + CV_X2.Montant
                 arrDevTotal(IDev).Cr = arrDevTotal(IDev).Cr + CV_X1.Montant
                  arrDevTotal(0).Cr = arrDevTotal(0).Cr + CV_X2.Montant
            End If
         Else
              If CV_X1.Montant < 0 Then
                 curHBilanDb = curHBilanDb + CV_X2.Montant
                 arrDevTotal(IDev).Db_Cv = arrDevTotal(IDev).Db_Cv + CV_X1.Montant
                  arrDevTotal(0).Db_Cv = arrDevTotal(0).Db_Cv + CV_X2.Montant
           Else
                 curHBilanCr = curHBilanCr + CV_X2.Montant
                 arrDevTotal(IDev).Cr_Cv = arrDevTotal(IDev).Cr_Cv + CV_X1.Montant
                 arrDevTotal(0).Cr_Cv = arrDevTotal(0).Cr_Cv + CV_X2.Montant
            End If
        End If
    End If
Next arrCompte_K

If Not blnPrintLine Then Exit Sub

If XPrt.CurrentY + prtlineHeight * 1.6 > prtMaxY Then
    prtCompteSolde_Trait
    frmElpPrt.prtNewPage
    prtCompteSolde_Form
End If

XPrt.FontBold = False
'------------------------------------------ligne 1--------------
CompteSolde_LineNb = CompteSolde_LineNb + 1
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
If CompteSolde_LineNb > 2 Then
    
    Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY, prtMaxX, XPrt.CurrentY + prtlineHeight - 10, " ", 245)
'Else
    If CompteSolde_LineNb = 4 Then CompteSolde_LineNb = 0
End If

XPrt.FontSize = 6
XPrt.CurrentY = XPrt.CurrentY + Height8_6

XPrt.CurrentX = prtMinX + 50
XPrt.Print mId$(arrCompte(1).Numéro, 1, 5);
XPrt.CurrentX = Col1 + 50
XPrt.Print Trim(arrCompte(1).Intitulé);

If curBilanDb <> 0 Then
    X = Format$(Abs(curBilanDb), "#### ### ### ### ##0.00")
    XPrt.CurrentX = Col2_Db - XPrt.TextWidth(X)
    XPrt.Print X & " ";
End If

If curBilanCr <> 0 Then
    X = Format$(Abs(curBilanCr), "#### ### ### ### ##0.00")
    XPrt.CurrentX = Col2_Cr - XPrt.TextWidth(X)
    XPrt.Print X & " ";
End If
    
If curHBilanDb <> 0 Then
    X = Format$(Abs(curHBilanDb), "#### ### ### ### ##0.00")
    XPrt.CurrentX = Col3_Db - XPrt.TextWidth(X)
    XPrt.Print X & " ";
 End If
   
If curHBilanCr <> 0 Then
    X = Format$(Abs(curHBilanCr), "#### ### ### ### ##0.00")
    XPrt.CurrentX = Col3_Cr - XPrt.TextWidth(X)
    XPrt.Print X & " ";
End If
       
XPrt.FontBold = False
XPrt.FontSize = 8
XPrt.CurrentY = XPrt.CurrentY - Height8_6

End Sub

Public Sub prtCompteSolde_A6Total()
Dim X As String

CV_X2.DeviseIso = "dev"
XPrt.CurrentY = prtMaxY
For I = 1 To 999
    If arrDevTotal(I).Nb <> 0 Then
        CV_X1.DeviseN = arrDevTotal(I).DeviseN
        CV_AttributN CV_X1
        prtCompteSolde_A6TotalLine I
        End If
Next I

arrDevTotal(0).Nb = 1: arrDevTotal(0).DeviseN = CV_X3.DeviseN
CV_X2.DeviseN = "400"
CV_X2.DeviseIso = ""
For I = 0 To 999
    If arrDevTotal(I).Nb <> 0 Then
        arrDevTotalEur(I) = arrDevTotal(I)
        CV_X1.DeviseN = arrDevTotal(I).DeviseN
        CV_X1.DeviseIso = ""
        
        CV_X1.Montant = arrDevTotal(I).Cr
        Call CV_Transitoire(CV_X1, CV_X2, CV_X3, X)
        arrDevTotal(I).Cr = CV_X2.Montant
        arrDevTotalEur(I).Cr = CV_X3.Montant
       
        CV_X1.Montant = arrDevTotal(I).Db
        Call CV_Transitoire(CV_X1, CV_X2, CV_X3, X)
        arrDevTotal(I).Db = CV_X2.Montant
        arrDevTotalEur(I).Db = CV_X3.Montant
        
        CV_X1.Montant = arrDevTotal(I).Cr_Cv
        Call CV_Transitoire(CV_X1, CV_X2, CV_X3, X)
        arrDevTotal(I).Cr_Cv = CV_X2.Montant
        arrDevTotalEur(I).Cr_Cv = CV_X3.Montant
        
        CV_X1.Montant = arrDevTotal(I).Db_Cv
        Call CV_Transitoire(CV_X1, CV_X2, CV_X3, X)
        arrDevTotal(I).Db_Cv = CV_X2.Montant
        arrDevTotalEur(I).Db_Cv = CV_X3.Montant
    End If
Next I

XPrt.CurrentY = prtMaxY
For I = 1 To 999
    If arrDevTotal(I).Nb <> 0 Then
        CV_X1.DeviseN = arrDevTotal(I).DeviseN
        CV_AttributN CV_X1
        prtCompteSolde_A6TotalLine I
        End If
Next I

CV_X1.DeviseN = ""
CV_X1.DeviseLibellé = "Total"
CompteSolde_LineNb = 2
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
prtCompteSolde_A6TotalLine 0


CV_X2 = CV_X3
XPrt.CurrentY = prtMaxY
For I = 1 To 999
    arrDevTotal(I) = arrDevTotalEur(I)
    If arrDevTotal(I).Nb <> 0 Then
        CV_X1.DeviseN = arrDevTotal(I).DeviseN
        CV_AttributN CV_X1
        prtCompteSolde_A6TotalLine I
        End If
Next I

CV_X1.DeviseN = ""
CV_X1.DeviseLibellé = "Total"
CompteSolde_LineNb = 2
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
arrDevTotal(0) = arrDevTotalEur(0)
prtCompteSolde_A6TotalLine 0

ReDim arrDevTotal(1), arrDevTotalEur(1)
End Sub



Public Sub prtCompteSolde_Print(Msg As String)
On Error GoTo prtError
K1 = Val(mId$(Msg, 1, 6))
K2 = Val(mId$(Msg, 7, 6))

If optEtatSortK = "A6" Then prtCompteSolde_A6: Exit Sub

ReDim arrCV2_Montant(K2), arrCV1_DeviseIso(K2)
ReDim arrCompteSoldeRupture(K2)
prtCompteSoldeRupture_Scan

For arrCompteSoldeRupture_Index = 1 To arrCompteSoldeRupture_Nb
    If arrCompteSoldeRupture(arrCompteSoldeRupture_Index).Nb > 0 Or blnPrintSoldé Then
        If mId$(arrCompteSoldeRupture(arrCompteSoldeRupture_Index).IdKey, 1, CompteSoldeRupture_Len) = mId$(arrCompteSoldeRupture(arrCompteSoldeRupture_Index - 1).IdKey, 1, CompteSoldeRupture_Len) Then
            XPrt.FontBold = False
        Else
            XPrt.FontBold = True
        End If
        
        If blnCompteSoldeRupture Then prtCompteSoldeRupture_Line
        If blnCompteSoldeLine Then
            For arrCompte_K = arrCompteSoldeRupture(arrCompteSoldeRupture_Index).IndexDeb To arrCompteSoldeRupture(arrCompteSoldeRupture_Index).IndexFin
                recCompte = arrCompte(arrCompte_K)
                If arrCompte(arrCompte_K).SoldeInstantané <> 0 Or blnPrintSoldé Then prtCompteSolde_Line
                
                DoEvents: If prtKillDoc Then frmElpPrt.Hide: Exit Sub
            
            Next arrCompte_K
        End If
    End If
Next arrCompteSoldeRupture_Index

Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub

Public Sub prtCompteSolde_Close()
Dim I As Integer

If blnCompteSoldeTotal Then
    If optEtatSortK = "A6" Then
        prtCompteSolde_A6Total
    Else
        prtCompteSoldeTotal_Cumul
        prtCompteSoldeTotal_Print
    '''blnCompteSoldeBilan = True: prtCompteSoldeTotal_Print
    ''''blnCompteSoldeBilan = False: prtCompteSoldeTotal_Print
    End If
End If

'XPrt.DrawWidth = 5
'frmElpPrt.prtLineY

'K = 0
'XPrt.CurrentY = XPrt.CurrentY + 50
'XPrt.FontBold = True
'prtCurrentY = XPrt.CurrentY + prtlineHeight + 50
'If K > 0 Then: XPrt.Line (prtMinX + 11100, prtCurrentY)-(prtMaxX, prtCurrentY)
prtCompteSolde_Trait
frmElpPrt.prtEndDoc
frmElpPrt.Hide

End Sub

Public Sub prtCompteSoldeTotal_Add(CV_X1 As typeCV, CV_X2 As typeCV, xIdKey As String)
Dim I As Integer, blnAddNew As Boolean
blnAddNew = True
For I = 1 To arrCompteSoldeTotal_Nb
    If arrCompteSoldeTotal(I).IdKey = xIdKey Then blnAddNew = False: Exit For
Next I
        
If blnAddNew Then
    arrCompteSoldeTotal_Nb = arrCompteSoldeTotal_Nb + 1
    If arrCompteSoldeTotal_Nb > arrCompteSoldeTotal_NbMax Then
        arrCompteSoldeTotal_NbMax = arrCompteSoldeTotal_NbMax + 10
        ReDim Preserve arrCompteSoldeTotal(arrCompteSoldeTotal_NbMax)
    End If
    I = arrCompteSoldeTotal_Nb
    arrCompteSoldeTotal(I).IdKey = xIdKey
    arrCompteSoldeTotal(I).DeviseIso = CV_X1.DeviseIso
    arrCompteSoldeTotal(I).DeviseN = CV_X1.DeviseN
    arrCompteSoldeTotal(I).Db = 0
    arrCompteSoldeTotal(I).Db_Cv = 0
    arrCompteSoldeTotal(I).Cr = 0
    arrCompteSoldeTotal(I).Cr_Cv = 0
End If

If CV_X1.Montant <> 0 Or CV_X2.Montant <> 0 Then
    arrCompteSoldeTotal(I).Nb = arrCompteSoldeTotal(I).Nb + 1
    If CV_X2.Montant < 0 Then
        arrCompteSoldeTotal(I).Db = arrCompteSoldeTotal(I).Db + CV_X1.Montant
        arrCompteSoldeTotal(I).Db_Cv = arrCompteSoldeTotal(I).Db_Cv + CV_X2.Montant
    Else
        arrCompteSoldeTotal(I).Cr = arrCompteSoldeTotal(I).Cr + CV_X1.Montant
        arrCompteSoldeTotal(I).Cr_Cv = arrCompteSoldeTotal(I).Cr_Cv + CV_X2.Montant
    End If
End If

End Sub

Public Sub prtCompteSoldeTotal_Print()
Dim curDb_Cv As Currency, curCr_Cv As Currency, X As String

curDb_Cv = 0: curCr_Cv = 0
MDB.Execute "delete * from CptP0"
mdbCptP0.tableCptP0_Open

recCptP0_Init reccptp0
reccptp0.Method = "AddNew"

For arrCompteSoldeTotal_Index = 1 To arrCompteSoldeTotal_Nb
    reccptp0.Id = arrCompteSoldeTotal(arrCompteSoldeTotal_Index).IdKey
    reccptp0.Text = Format(arrCompteSoldeTotal_Index, "000000000")
    dbCptP0_Update reccptp0
Next arrCompteSoldeTotal_Index


''prtCompteSolde_Form

prtCurrentY = XPrt.CurrentY + prtlineHeight * 1.25
'Call frmElpPrt.prtTrame(prtMinX, prtCurrentY, prtMaxX, prtCurrentY + prtlineHeight - 10, " ", 220)
CompteSolde_LineNb = 0

XPrt.CurrentY = prtCurrentY
XPrt.FontBold = True
XPrt.FontSize = 8
XPrt.CurrentX = prtMinX + 50

reccptp0.Method = "MoveFirst"

Call dbCptP0_ReadE(reccptp0)
Do While reccptp0.Err = 0
    arrCompteSoldeTotal_Index = Val(mId$(reccptp0.Text, 1, 9))
    
    If arrCompteSoldeTotal(arrCompteSoldeTotal_Index).Nb > 0 Then prtCompteSoldeTotal_Line
    reccptp0.Method = "MoveNext    "
    reccptp0.Err = tableCptP0_Read(reccptp0)
Loop

mdbCptP0.tableCptP0_Close

End Sub


Public Function prtCompteSoldeTotal_IdKey(xRupture As String) As String
Dim X As String
Select Case optEtatSortK
    Case "A1": X = mId$(xRupture, 26, 1) & mId$(xRupture, 6, 3) & mId$(xRupture, 12, 3)
    Case "A2": X = mId$(xRupture, 26, 1) & mId$(xRupture, 12, 3) & mId$(xRupture, 9, 3)
    Case "A3": X = mId$(xRupture, 1, 3) & mId$(xRupture, 26, 1) & mId$(xRupture, 9, 3)
    Case "A4": X = mId$(xRupture, 19, 1) & mId$(xRupture, 1, 4) & mId$(xRupture, 5, 3)
    Case "A5": X = mId$(xRupture, 14, 1) & mId$(xRupture, 1, 3) & mId$(xRupture, 11, 3)
    Case "A6": X = mId$(xRupture, 6, 1) & mId$(xRupture, 1, 5)
    Case "G1": X = mId$(xRupture, 26, 1) & mId$(xRupture, 4, 1) & "  " & mId$(xRupture, 12, 3)
    Case "G2": X = mId$(xRupture, 26, 1) & mId$(xRupture, 4, 1) & "  " & mId$(xRupture, 23, 3)
    Case "G3": X = mId$(xRupture, 1, 3) & mId$(xRupture, 26, 1) & mId$(xRupture, 7, 1) & "  "
    Case "G4": X = mId$(xRupture, 30, 1) & mId$(xRupture, 1, 4)

End Select
prtCompteSoldeTotal_IdKey = X
End Function

Public Sub prtCompteSoldeTotal_Cumul()
Dim Nb As Integer, I As Integer, blnDeviseCumul As Boolean
Dim wIdKey0 As String, wIdKey1 As String

Nb = arrCompteSoldeTotal_Nb
blnDeviseCumul = True
For I = 1 To Nb
    If arrCompteSoldeTotal(I).DeviseIso <> arrCompteSoldeTotal(1).DeviseIso Then blnDeviseCumul = False: Exit For
Next I

For I = 1 To Nb
    
    wIdKey0 = arrCompteSoldeTotal(I).IdKey
    wIdKey1 = wIdKey0
    Select Case optEtatSortK
        Case "A3", "G3": Mid$(wIdKey0, 4, 4) = "0000": Mid$(wIdKey1, 5, 3) = "000"
        Case "A4", "G4": Mid$(wIdKey0, 2, 7) = "0000000": Mid$(wIdKey1, 6, 3) = "   "
        Case Else: Mid$(wIdKey0, 2, 6) = "000   ": Mid$(wIdKey1, 5, 3) = "   "
    End Select
    
    CV_X1.DeviseIso = arrCompteSoldeTotal(I).DeviseIso
    CV_X1.DeviseN = arrCompteSoldeTotal(I).DeviseN
    
    CV_X1.Montant = IIf(blnDeviseCumul, arrCompteSoldeTotal(I).Db, 0)
    CV_X2.Montant = arrCompteSoldeTotal(I).Db_Cv
    Call prtCompteSoldeTotal_Add(CV_X1, CV_X2, wIdKey0)
    Call prtCompteSoldeTotal_Add(CV_X1, CV_X2, wIdKey1)
    
    CV_X1.Montant = IIf(blnDeviseCumul, arrCompteSoldeTotal(I).Cr, 0)
    CV_X2.Montant = arrCompteSoldeTotal(I).Cr_Cv
    Call prtCompteSoldeTotal_Add(CV_X1, CV_X2, wIdKey0)
    Call prtCompteSoldeTotal_Add(CV_X1, CV_X2, wIdKey1)
Next I

End Sub

Public Sub prtCompteSolde_CV_Init(optAmj As String, XDevise As String)
CV_X2 = CV_Euro
CV_X1.OpéAmj = optAmj: CV_X1.CoursCompta = "C"
CV_X2.OpéAmj = optAmj: CV_X2.CoursCompta = "C"
CV_X3.OpéAmj = optAmj: CV_X3.CoursCompta = "C"
Call CV_AttributS(XDevise, CV_X2)

End Sub

Public Sub prtCompteSolde_A6Init()
ReDim arrDevTotal(999)
ReDim arrDevTotalEur(999)
For I = 0 To 999
    arrDevTotal(I).Nb = 0
    arrDevTotal(I).Cr = 0: arrDevTotal(I).Db = 0
    arrDevTotal(I).Cr_Cv = 0: arrDevTotal(I).Db_Cv = 0
Next I

End Sub

Public Sub prtCompteSolde_A6TotalLine(I As Integer)
Dim X As String
If XPrt.CurrentY + prtlineHeight * 1.6 > prtMaxY Then
    prtCompteSolde_Trait
    frmElpPrt.prtNewPage
    prtCompteSolde_Form
End If

XPrt.FontBold = False
'------------------------------------------ligne 1--------------
CompteSolde_LineNb = CompteSolde_LineNb + 1
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
If CompteSolde_LineNb > 2 Then
    
    Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY, prtMaxX, XPrt.CurrentY + prtlineHeight - 10, " ", 245)
'Else
    If CompteSolde_LineNb = 4 Then CompteSolde_LineNb = 0
End If

XPrt.FontSize = 6
XPrt.CurrentY = XPrt.CurrentY + Height8_6
XPrt.CurrentX = Col1 - 300
XPrt.Print CV_X1.DeviseN;

XPrt.CurrentX = Col1 + 50
XPrt.Print CV_X1.DeviseLibellé;

If arrDevTotal(I).Db <> 0 Then
    X = Format$(Abs(arrDevTotal(I).Db), "#### ### ### ### ##0.00")
    XPrt.CurrentX = Col2_Db - XPrt.TextWidth(X)
    XPrt.Print X & " ";
End If

If arrDevTotal(I).Cr <> 0 Then
    X = Format$(Abs(arrDevTotal(I).Cr), "#### ### ### ### ##0.00")
    XPrt.CurrentX = Col2_Cr - XPrt.TextWidth(X)
    XPrt.Print X & " ";
End If
    
If arrDevTotal(I).Db_Cv <> 0 Then
    X = Format$(Abs(arrDevTotal(I).Db_Cv), "#### ### ### ### ##0.00")
    XPrt.CurrentX = Col3_Db - XPrt.TextWidth(X)
    XPrt.Print X & " ";
 End If
   
If arrDevTotal(I).Cr_Cv <> 0 Then
    X = Format$(Abs(arrDevTotal(I).Cr_Cv), "#### ### ### ### ##0.00")
    XPrt.CurrentX = Col3_Cr - XPrt.TextWidth(X)
    XPrt.Print X & " ";
End If
   
XPrt.FontBold = False
XPrt.FontSize = 8
XPrt.CurrentY = XPrt.CurrentY - Height8_6

End Sub
