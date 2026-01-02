Attribute VB_Name = "prtYBIAMVT0_A4"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim I As Integer, solde As Currency, mCurrenty As Integer, Height8_6 As Integer
Dim Line1 As Integer, Line2 As Integer, Line3 As Integer, Line4 As Integer, Line5 As Integer
Dim col1 As Integer, col2 As Integer, col3 As Integer
Dim Col4 As Integer, Col5 As Integer, Col6 As Integer, Col7 As Integer, Col8 As Integer
Dim Col As Integer
Dim K As Integer, K1 As Integer, K2 As Integer, K3 As Integer
Dim X As String
Dim nbLigne As Integer, NbPage As Integer
Dim NbLigneMax As Integer, NbPageMax As Integer
Dim NbImprimé As Integer

Dim valAmjMin As String, valAmjMax As String
Dim IbmAmjMin As String, IbmAmjMax As String

Dim curCumulDébit As Currency, curCumulCrédit As Currency
Dim blnA4_Form As Boolean
Dim blnMsgInfo As Boolean, mMsgInfo As String, mExtraitNuméro As String

Dim xYBIAMVT0 As typeYBIAMVT0, mYBIAMVT0 As typeYBIAMVT0
Dim xYRELEVE0 As typeYRELEVE0
Dim xYCLIENA0 As typeYCLIENA0
Dim xYTITULA0 As typeYTITULA0

Dim xMvtP0 As typeMvtP0
Dim blnCptOrdinaire As Boolean, blnRIB As Boolean, blnMédiateur As Boolean
Dim blnConvention_Print As Boolean
Dim mRib_Compte As String, mRib_Clé As String, mRib_IbanE As String
Dim mResponsable As String
Dim zYADRESS0 As typeYADRESS0, xYADRESS0 As typeYADRESS0, fiscalYADRESS0 As typeYADRESS0

Public Sub prtYBIAMVT0_A4_OpenX()
'---------------------------------------------------------
On Error GoTo prtError

Set XPrt = Printer
frmElpPrt.Show vbModeless

prtYBIAMVT0_A4_OpenX_Reset



Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub

Public Sub prtYBIAMVT0_A4_Close()
'---------------------------------------------------------
On Error GoTo prtError

frmElpPrt.prtEndDoc
frmElpPrt.Hide
Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub



'---------------------------------------------------------
Public Sub prtYBIAMVT0_A4_Form(Msg As String, lRéférence As String)
'---------------------------------------------------------
Dim X As String
Dim mCurrenty

prtYBIAMVT0_A4_RIB

Call frmElpPrt.prtTrame(Col4, Line3, Col5, Line4, " ", 250)
Call frmElpPrt.prtTrame(col1, Line2, Col8, Line3, " ", 240)
XPrt.CurrentY = prtMinY + prtlineHeight * 4


XPrt.DrawWidth = 3
XPrt.Line (Col4 + 200, Line1)-(Col6 - 200, Line1)
XPrt.DrawWidth = 2
XPrt.Line (col1 + 200, Line2)-(Col8, Line2)
XPrt.Line (col1, Line3)-(Col8, Line3)
XPrt.Line (col1 + 200, Line4)-(Col8, Line4)
XPrt.DrawWidth = 3
XPrt.Line (Col4 + 200, Line5)-(Col6 - 200, Line5)
XPrt.DrawWidth = 2
XPrt.Line (col1, Line2 + 200)-(col1, Line4 - 200)
XPrt.DrawWidth = 1
XPrt.Line (col2, Line2)-(col2, Line4)
XPrt.DrawWidth = 1
XPrt.Line (col3, Line2)-(col3, Line4)
XPrt.DrawWidth = 3
XPrt.Line (Col4, Line1 + 200)-(Col4, Line5 - 200)
XPrt.DrawWidth = 1
XPrt.Line (Col5, Line1)-(Col5, Line5)
XPrt.DrawWidth = 3
XPrt.Line (Col6, Line1 + 200)-(Col6, Line5 - 200)

XPrt.CurrentY = Line2 + 50
XPrt.FontBold = True

XPrt.FontSize = prtFontSize
frmElpPrt.prtCentré (col1 + col2) / 2, "Date"
frmElpPrt.prtCentré (col2 + col3) / 2, "Libellé"
frmElpPrt.prtCentré (col3 + Col4) / 2, "Date Valeur"
frmElpPrt.prtCentré (Col4 + Col5) / 2, "Débit"
frmElpPrt.prtCentré (Col5 + Col6) / 2, "Crédit"

'------------------------
XPrt.DrawWidth = 2

XPrt.CurrentY = 0
XPrt.CurrentX = 0
XPrt.Circle Step(col1 + 200, Line2 + 200), 200, 0, 0.5 * Pi, Pi
XPrt.DrawWidth = 3

XPrt.CurrentY = 0
XPrt.CurrentX = 0
XPrt.Circle Step(Col6 - 200, Line1 + 200), 200, 0, 0, 0.5 * Pi

XPrt.CurrentY = 0
XPrt.CurrentX = 0
XPrt.Circle Step(Col4 + 200, Line1 + 200), 200, 0, 0.5 * Pi, Pi

XPrt.DrawWidth = 2
XPrt.CurrentY = 0
XPrt.CurrentX = 0
XPrt.Circle Step(col1 + 200, Line4 - 200), 200, 0, Pi, 1.5 * Pi

XPrt.DrawWidth = 3
XPrt.CurrentY = 0
XPrt.CurrentX = 0
XPrt.Circle Step(Col4 + 200, Line5 - 200), 200, 0, Pi, 1.5 * Pi



XPrt.CurrentY = 0
XPrt.CurrentX = 0
XPrt.Circle Step(Col6 - 200, Line5 - 200), 200, 0, 1.5 * Pi, 2 * Pi

'----------------------------------------ligne 1-----------------
XPrt.FontSize = 10
XPrt.CurrentY = prtMinY + prtlineHeight * 10 - XPrt.TextHeight("test")
'----------------------------------1------------
XPrt.FontBold = True

XPrt.CurrentX = 5800
XPrt.Print xYADRESS0.ADRESSRA1;
XPrt.FontBold = False
'-----------------------------------2-------------
''If Trim(xYADRESS0.ADRESSRA2) <> "" Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = 5800
    XPrt.Print xYADRESS0.ADRESSRA2;
''End If
'------------------------------------3---------------
If Trim(xYADRESS0.ADRESSAD1) <> "" Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = 5800
    XPrt.Print xYADRESS0.ADRESSAD1;
End If
'----------------------------------4-------------------
If Trim(xYADRESS0.ADRESSAD2) <> "" Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = 5800
    XPrt.Print xYADRESS0.ADRESSAD2;
End If

'-----------------------------------5------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 5800
XPrt.Print xYADRESS0.ADRESSAD3;
'------------------------------------6------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 5800
If Trim(xYADRESS0.ADRESSCOP) <> "" Then XPrt.Print xYADRESS0.ADRESSCOP & "  ";
XPrt.Print xYADRESS0.ADRESSVIL;
'------------------------------------8------------------
If Trim(xYADRESS0.ADRESSPAY) <> "FRANCE" Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = 5800
    XPrt.Print xYADRESS0.ADRESSPAY;
End If
'XPrt.FontSize = 6
'XPrt.CurrentX = Col8 - 350
'XPrt.Print "  G " '& recCptInfo.Gestionnaire & "-" & recCptInfo.Courrier;

XPrt.FontSize = 8

XPrt.CurrentY = Line1 - prtlineHeight * 3 + 50
XPrt.FontBold = True

'$$ jpl X = "RELEVE DE COMPTE   " & mYBIAMVT0.COMPTEDEV & "   " & mExtraitNuméro
'$$ jpl Col = Col4 + (Col8 - Col4 - XPrt.TextWidth(X)) / 2
'$$ jpl Call frmElpPrt.prtTrame(Col, XPrt.CurrentY, Col + XPrt.TextWidth(X) + 100, XPrt.CurrentY + prtlineHeight, " ", 240)
'$$ jpl XPrt.CurrentX = Col + 50
'$$ jpl XPrt.Print X;
'$$ jpl XPrt.CurrentY = XPrt.CurrentY + Height8_6
'$$ jpl XPrt.FontBold = False
'$$ jpl XPrt.FontSize = 6

Call frmElpPrt.prtTrame(Col4, XPrt.CurrentY - 100, Col6, XPrt.CurrentY + prtlineHeight, " ", 240)
XPrt.CurrentX = Col4 + 200

XPrt.CurrentY = XPrt.CurrentY - Height8_6
XPrt.FontSize = 10
XPrt.Print mYBIAMVT0.COMPTEDEV;
XPrt.CurrentY = XPrt.CurrentY + Height8_6
XPrt.FontBold = False
XPrt.FontSize = 8

XPrt.Print "  -  RELEVE DE COMPTE : ";
XPrt.FontBold = True
XPrt.CurrentY = XPrt.CurrentY - Height8_6
XPrt.FontSize = 10
XPrt.Print mExtraitNuméro;
XPrt.CurrentY = XPrt.CurrentY + Height8_6
XPrt.FontBold = False
XPrt.FontSize = 8
XPrt.Print " / " & Format$(NbPage, "###");
XPrt.CurrentX = Col6 + 20
XPrt.Print lRéférence;

'XPrt.Print " -" & Format$(NbPage, "###");  '$$ jpl 2003.03.31$   & " / " & Format$(NbPageMax, "###");
XPrt.CurrentY = XPrt.CurrentY - Height8_6
XPrt.FontSize = 8
'----------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
'XPrt.CurrentX = 800
'XPrt.Print "Numéro ";
XPrt.CurrentX = col1 + 50
'XPrt.Print ": ";
XPrt.FontBold = True
'XPrt.Print Format$(recCptInfo.Numéro, "@@@@@.@@@.@@.@") ;
''XPrt.Print "recCptInfo.Intitulé2";
'-------------------------------------------------------
XPrt.FontBold = False
'Call DevX("recCptInfo.Devise")
XPrt.FontSize = 8
''frmElpPrt.prtCentré (Col4 + Col6) / 2, Trim(mYBIAMVT0.COMPTEDEV)

XPrt.FontSize = 8
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'XPrt.CurrentX = 800
'XPrt.Print "Type";
XPrt.CurrentX = col1 + 50
'XPrt.Print ": ";
XPrt.FontBold = True
'XPrt.Print Trim(DicLib(13, recCptInfo.BiaTyp)) & "-" & Trim(xDevise.DevLib);

'---------------------------------------
'XPrt.FontBold = False

'XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'XPrt.CurrentX = 400
'XPrt.Print "Devise";
'XPrt.CurrentX = Col2
'XPrt.Print ": ";
'XPrt.FontBold = True
'XPrt.Print Format$(recCptInfo.Devise, "000") & "-" & XDevise.DevLib;

'------------------------------------9--------------
'---------------------------------------
XPrt.FontBold = False


XPrt.FontSize = prtFontSize

XPrt.CurrentY = Line1 + 50
XPrt.CurrentX = Col4 - 100 - XPrt.TextWidth(Msg)
XPrt.Print Msg;
prtYBIAMVT0_A4_Montant (solde)

nbLigne = 0
blnA4_Form = True

XPrt.CurrentY = Line3 - prtlineHeight + 50

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
' 20050330 Mettre l'instruction suivante en commentaire : le message date du 31.03.2005
'                                                         A garder jusqu'à 30.09.2005 inclu
' blnConvention_Print = False
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
If blnConvention_Print Then
    nbLigne = 5
    blnConvention_Print = False
    XPrt.FontBold = True
    XPrt.FontSize = 8
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
   ' Call frmElpPrt.prtTrame(col2, XPrt.CurrentY, col3, XPrt.CurrentY + prtlineHeight * 2, " ", 0)
    'XPrt.ForeColor = RGB(255, 255, 255)
    XPrt.CurrentX = col2 + 50
    XPrt.Print "La BIA mettra à jour les conditions générales appliquées";
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = col2 + 50
    XPrt.Print "à la clientèle le 1er juillet 2005, date à laquelle la";
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = col2 + 50
    XPrt.Print "nouvelle grille tarifaire sera à votre disposition.";
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.FontBold = False
    XPrt.FontSize = prtFontSize
   ' XPrt.ForeColor = RGB(0, 0, 0)
End If

End Sub

'---------------------------------------------------------
Public Sub prtYBIAMVT0_A4_Line()
'---------------------------------------------------------
Dim X As String, I As Integer, libCV As String, blnCV As Boolean
Dim blnLine2 As Boolean, xLine1 As String, xLine2 As String

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
 
XPrt.FontSize = prtFontSize
XPrt.FontBold = False


XPrt.CurrentX = col1 + 50
XPrt.Print dateImp(Val(xYBIAMVT0.MOUVEMDTR) + 19000000);

XPrt.CurrentX = col3 + 50
XPrt.Print dateImp(Val(xYBIAMVT0.MOUVEMDVA) + 19000000);
prtYBIAMVT0_A4_Montant (xYBIAMVT0.MOUVEMMON)

XPrt.CurrentX = col2 + 50
If xYBIAMVT0.MOUVEMOPE = "-RM" Then Mid$(xYBIAMVT0.LIBELLIB2, 13, 18) = Space$(18)
xLine1 = Trim(xYBIAMVT0.LIBELLIB1) & " " & Trim(xYBIAMVT0.LIBELLIB2)
xLine2 = Trim(xYBIAMVT0.LIBELLIB3) & " " & Trim(xYBIAMVT0.LIBELLIB4)
X = xLine1 & " " & xLine2

blnLine2 = True

For I = prtFontSize To 8 Step -1  ' 6
    XPrt.FontSize = I
    If XPrt.TextWidth(X) <= (col3 - XPrt.CurrentX - 50) Then blnLine2 = False: Exit For
Next I

If Not blnLine2 Then
    XPrt.Print X;
Else
    XPrt.FontSize = prtFontSize
    If XPrt.TextWidth(xLine1) > (col3 - XPrt.CurrentX - 50) Then XPrt.FontSize = prtFontSize - 1
    XPrt.Print xLine1;
    If nbLigne = NbLigneMax Then prtYBIAMVT0_A4_Report
    nbLigne = nbLigne + 1
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = col2 + 50
    If XPrt.TextWidth(xLine2) > (col3 - XPrt.CurrentX - 50) Then XPrt.FontSize = prtFontSize - 1
    XPrt.Print xLine2;
End If

XPrt.FontSize = prtFontSize

End Sub


'---------------------------------------------------------
Public Sub prtYBIAMVT0_A4_Montant(MT As Currency)
'---------------------------------------------------------
Dim X As String

XPrt.FontBold = True
X = Format$(Abs(MT), "## ### ### ### ### ##0.00")
XPrt.CurrentX = IIf(MT < 0, Col6, Col5) - 100 - XPrt.TextWidth(X)
XPrt.Print X;
XPrt.FontBold = False

End Sub

'---------------------------------------------------------
Public Sub prtYBIAMVT0_A4_Médiateur()
'---------------------------------------------------------
Dim X As String
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight + 50
Call frmElpPrt.prtTrame(col1, XPrt.CurrentY, Col8, XPrt.CurrentY + prtlineHeight * 2.3, " ", 225)

XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + 100
XPrt.CurrentX = col1 + 200
XPrt.Print "Nous vous informons qu'un médiateur est à votre disposition à l'adresse suivante : ";
XPrt.FontBold = True
XPrt.Print "  M. le MEDIATEUR   -   B.P. 151   -   75422 PARIS CEDEX 09";

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = col2 + 50
XPrt.FontBold = False
frmElpPrt.prtCentré prtMedX, "pour tout problème que vous n'avez pu résoudre préalablement avec la banque."



End Sub

Public Sub prtYBIAMVT0_A4_Relevé(lMvtP0 As typeMvtP0, lAmjMin As String, lAmjMax As String, lRELEVEREL As String, lRéférence As String)
'---------------------------------------------------------

'$$$$$$$$$$$$$ ancienne version à remplacer par prtYBIAMVT0_A4_Extrait
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$


Dim CTLAMJ As String

prtYBIAMVT0_A4_OpenX_Reset
recYADRESS0_Init zYADRESS0

blnA4_Form = False
If lRELEVEREL = "M" Then
    mExtraitNuméro = libMois(mId$(lAmjMax, 5, 2))
Else
    mExtraitNuméro = "____"
End If

MsgTxt = Space$(34) & lMvtP0.Text
MsgTxtIndex = 0

srvYBIAMVT0_GetBuffer mYBIAMVT0

prtYBIAMVT0_A4_Compte lRELEVEREL
blnConvention_Print = blnMédiateur
' 20050330 filtre des PCI pour impression d'un message sur les extraits de compte
Select Case mId$(mYBIAMVT0.COMPTEOBL, 1, 5)
    Case "25111", "25113", "25112", "25114", "25115", "25117": blnConvention_Print = True
End Select

CTLAMJ = mYBIAMVT0.MOUVEMDTR

K1 = 1 ' Val(Mid$(Msg, 1, 6))
K2 = 10 ' Val(Mid$(Msg, 7, 6))
valAmjMin = lAmjMin
valAmjMax = lAmjMax
IbmAmjMin = dateIBM(lAmjMin)
IbmAmjMax = dateIBM(lAmjMax)
prtFontSize = 8

NbPageMax = 0 'Fix((Abs(K2 - K1)) / NbLigneMax) + 1
NbPage = 1

xMvtP0 = lMvtP0
xMvtP0.Method = "Seek>="
intReturn = tableMvtP0_Read(xMvtP0)

Do
    If intReturn = 0 Then
        MsgTxt = Space$(34) & xMvtP0.Text
        MsgTxtIndex = 0
    
        srvYBIAMVT0_GetBuffer xYBIAMVT0
        
        If mYBIAMVT0.MOUVEMCOM <> xYBIAMVT0.MOUVEMCOM Then
            intReturn = -1
        Else
        
           If xYBIAMVT0.MOUVEMDTR > IbmAmjMax Then Exit Do
               
               If CTLAMJ <> xYBIAMVT0.MOUVEMDTR Then
                       If solde <> xYBIAMVT0.BIAMVTSD0 Then
                           XPrt.CurrentX = col2
                           MsgBox "erreur Solde .........", vbCritical, "prtCptMvt"
                           XPrt.FontSize = 14
                           XPrt.Print "ERREUR SOLDE ............."
                           Exit Do
                       End If
                   CTLAMJ = xYBIAMVT0.MOUVEMDTR
               End If
               
               
                If xYBIAMVT0.MOUVEMDTR >= IbmAmjMin Then
           
                    If Not blnA4_Form Then prtYBIAMVT0_A4_Form "Solde au : " & dateImp(dateElp("Jour", -1, valAmjMin)), lRéférence
                    
                    If nbLigne = NbLigneMax Then prtYBIAMVT0_A4_Report
                    nbLigne = nbLigne + 1
                    
                    prtYBIAMVT0_A4_Line
                    
                    DoEvents: If prtKillDoc Then frmElpPrt.Hide: Exit Sub
                End If
            
            solde = solde + xYBIAMVT0.MOUVEMMON
           
           xMvtP0.Method = "Seek>"
           intReturn = tableMvtP0_Read(xMvtP0)
           If mId$(xMvtP0.ID, 1, 8) <> constYBIAMVT0 Then intReturn = -1
           'Nb = Nb + 1
           'If Nb Mod 500 = 0 Then Call lstErr_ChangeLastItem(Me.lstErr, Me.cmdContext, "affichage : " & Nb)
           
          ' If Nb = 10 Then intReturn = -1
        End If
    End If
Loop Until intReturn <> 0


'Pas de mouvement dans la période : imprimer un extrait (sauf mensuel)
If Not blnA4_Form Then
    If lRELEVEREL <> "M" Then prtYBIAMVT0_A4_Form "Solde au : " & dateImp(dateElp("Jour", -1, valAmjMin)), lRéférence
End If

If blnA4_Form Then
    XPrt.CurrentY = Line4 + 50
    X = "Solde au : " & dateImp(valAmjMax)
    XPrt.CurrentX = Col4 - XPrt.TextWidth(X) - 200
    XPrt.Print X;
    XPrt.CurrentX = 5000
    prtYBIAMVT0_A4_Montant (solde)
    
    '$JPL 2002.12.26 médiateur
    If blnMédiateur Then
        prtYBIAMVT0_A4_Médiateur
    Else
        If blnMsgInfo Then
            XPrt.FontBold = True: XPrt.FontSize = 10
            XPrt.CurrentY = XPrt.CurrentY + prtlineHeight ''* 2
            Call frmElpPrt.prtTrame(col1, XPrt.CurrentY, Col8, XPrt.CurrentY + prtlineHeight - 10, " ", 245)
            frmElpPrt.prtCentré 5500, mMsgInfo
        End If
    End If
End If

End Sub

Public Sub prtYBIAMVT0_A4_Select(lRELEVEREL As String, lAmjMin As String, lAmjMax As String)
Dim xFileName As String, X As String
Dim xYRELEVE0 As typeYRELEVE0
Dim mYBIARELH As typeYBIARELH, xYBIARELH As typeYBIARELH
Dim mYBIAMVT0 As typeYBIAMVT0
Dim xMvtP0 As typeMvtP0, mMvtP0_YBIAMVT0 As typeMvtP0
Dim wAMJHMS As String
Dim blnPrinter_Open As Boolean


On Error GoTo Error_Handler

blnPrinter_Open = False
wAMJHMS = DSys & "_" & time_Hms & "_"

IbmAmjMin = dateIBM(lAmjMin)
IbmAmjMax = dateIBM(lAmjMax)

xFileName = paramYBase_DataF & "log\" & wAMJHMS & "RELEVE_YBIARELH_" & lRELEVEREL & paramYBase_Data_ExtensionP
Open xFileName For Output As #3
Print #3, Time & " : " & lRELEVEREL & lAmjMin & " " & lAmjMax
recYBIARELH_Init xYBIARELH

recYBIAMVT0_Init mYBIAMVT0
mMvtP0_YBIAMVT0.ID = constYBIAMVT0
mMvtP0_YBIAMVT0.Method = "Seek>="
intReturn = tableMvtP0_Read(mMvtP0_YBIAMVT0)
Do
    If mId$(mMvtP0_YBIAMVT0.ID, 1, 8) <> constYBIAMVT0 Then intReturn = -1
    
    If intReturn = 0 Then
    
         If mId$(mMvtP0_YBIAMVT0.Text, 72, 7) >= IbmAmjMin And mId$(mMvtP0_YBIAMVT0.Text, 72, 7) <= IbmAmjMax Then 'perfomance
         'If xYBIAMVT0.MOUVEMDTR < lAmjMin Or xYBIAMVT0.MOUVEMDTR > lAmjMax Then
                 
                 If mYBIAMVT0.MOUVEMCOM <> mId$(mMvtP0_YBIAMVT0.Text, 10, 20) Then
                 
                     
                     MsgTxt = Space$(34) & mMvtP0_YBIAMVT0.Text
                     MsgTxtIndex = 0
                               
                     srvYBIAMVT0_GetBuffer mYBIAMVT0
                            
                     xMvtP0.ID = constYRELEVE0 & mYBIAMVT0.MOUVEMCOM & lRELEVEREL
                     xMvtP0.Method = "Seek="
                     If tableMvtP0_Read(xMvtP0) = 0 Then
                         MsgTxt = Space$(34) & xMvtP0.Text
                         MsgTxtIndex = 0
                         srvYRELEVE0_GetBuffer xYRELEVE0
                         
                         If Not blnPrinter_Open Then prtYBIAMVT0_A4_OpenX: blnPrinter_Open = True
                         
                              X = lRELEVEREL & mYBIAMVT0.MOUVEMCOM
                              If IsNull(srvYBIARELH_Import_Read(X, xYBIARELH)) Then
                                  xYBIARELH.Method = "RELEVE_OK"
                              Else
                                  recYBIARELH_Init xYBIARELH
                                  xYBIARELH.BIARELCOM = mYBIAMVT0.MOUVEMCOM
                                  xYBIARELH.BIARELREL = lRELEVEREL
                                  xYBIARELH.Method = constAddNew
                            End If

                         xMvtP0 = mMvtP0_YBIAMVT0
                         Call prtYBIAMVT0_A4_Relevé(xMvtP0, lAmjMin, lAmjMax, lRELEVEREL, "")
                         frmElpPrt.prtNewPage

                            MsgTxtLen = 0
                            srvYBIARELH_PutBuffer xYBIARELH
                            
                           ' If blnRelevéA4W_Update Then
                           '     srvYBIARELH_Update xYBIARELH
                           ' Else
                                Mid$(MsgTxt, 25, 10) = "sans màj"
                            'End If
                            Print #3, mId$(MsgTxt, 1, recYBIARELHLen)

                     End If
                End If
                
             End If
        mMvtP0_YBIAMVT0.Method = "Seek>"
        intReturn = tableMvtP0_Read(mMvtP0_YBIAMVT0)
   End If
    
Loop Until intReturn <> 0


Print #3, Time & " :================================================================="
Close

If blnPrinter_Open Then prtYBIAMVT0_A4_Close

'frmElpPrt.Shell_Print paramMT950_YBIARELH
Exit Sub

Error_Handler:

Close
Shell_MsgBox Error, vbCritical, "prtYBIAMVT0_A4_Select", False

End Sub

Public Sub prtYBIAMVT0_A4_RIB()
Dim iY As Integer
'--------------------------TRAME---------------------------
Dim X As String

XPrt.DrawWidth = 1
iY = 1500
Call frmElpPrt.prtTrame(200, iY, 4750, iY + 250, "", 240)

Call frmElpPrt.prtTrame(200, iY + 1450, 4750, iY + 1700, "B", 240)

'------------------------verticaux avec arrondi
XPrt.Line (200, iY + 200)-(200, iY + 2800)
XPrt.Line (1100, iY + 1450)-(1100, iY + 2100)
XPrt.Line (2000, iY + 1450)-(2000, iY + 2100)
XPrt.Line (4200, iY + 1450)-(4200, iY + 2100)
XPrt.Line (4750, iY + 200)-(4750, iY + 2800)
'------------------------horizontaux
XPrt.Line (400, iY)-(4550, iY)
XPrt.Line (200, iY + 250)-(4750, iY + 250)

XPrt.Line (200, iY + 2100)-(4750, iY + 2100)
XPrt.Line (400, iY + 3000)-(4550, iY + 3000)
'------------------------
XPrt.DrawWidth = 1

XPrt.CurrentY = 0
XPrt.CurrentX = 0
XPrt.Circle Step(200 + 200, iY + 200), 200, 0, 0.5 * Pi, Pi

XPrt.CurrentY = 0
XPrt.CurrentX = 0
XPrt.Circle Step(4750 - 200, iY + 200), 200, 0, 0, 0.5 * Pi

XPrt.CurrentY = 0
XPrt.CurrentX = 0
XPrt.Circle Step(200 + 200, iY + 3000 - 200), 200, 0, Pi, 1.5 * Pi

XPrt.CurrentY = 0
XPrt.CurrentX = 0
XPrt.Circle Step(4750 - 200, iY + 3000 - 200), 200, 0, 1.5 * Pi, 2 * Pi

XPrt.CurrentY = iY + prtlineHeight - 200
XPrt.FontSize = 8
XPrt.FontBold = True
If blnRIB Then frmElpPrt.prtCentré 2500, "RELEVE D'IDENTITE BANCAIRE"
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight - 50
XPrt.FontBold = False
XPrt.FontSize = 6
If blnRIB Then frmElpPrt.prtCentré 2500, "Cadre réservé au destinataire du R.I.B"
'------------------------------------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 5
XPrt.FontBold = False
XPrt.FontSize = 6
XPrt.CurrentX = 250
XPrt.Print "Code Banque";
XPrt.CurrentX = 1200
XPrt.Print "Code Guichet";
XPrt.CurrentX = 2600
XPrt.Print "Numéro de compte";
XPrt.CurrentX = 4250
XPrt.Print "clé R.I.B";
'----------------------------------------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight '+ 100
XPrt.FontBold = True
XPrt.FontSize = 9
XPrt.CurrentX = 400
XPrt.Print strSocBdfE;
XPrt.CurrentX = 1300
XPrt.Print strSocBdfG;
If blnRIB Then
    XPrt.CurrentX = 2400
    XPrt.Print Format$(mRib_Compte, "@@@  @@@  @@@  @@@");
    XPrt.CurrentX = 4400
    XPrt.Print Format$(mRib_Clé, "@@");
Else
    frmElpPrt.prtCentré 3100, Trim(mRib_Compte)
End If
'------------------------------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight + 150
If blnRIB Then
    XPrt.FontBold = True
    XPrt.CurrentX = 300
    XPrt.Print "IBAN";
    XPrt.CurrentX = 1050
    XPrt.Print ":";
    XPrt.CurrentX = 1200
    XPrt.Print Iban_Print(mRib_IbanE);
End If
'------------------------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight '+ 50
XPrt.FontSize = 8
XPrt.CurrentX = 300
XPrt.Print SocRibDom;
XPrt.FontBold = False
XPrt.CurrentX = 3300
XPrt.Print socTéléphone;
XPrt.FontBold = False
'--------------------------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight '+ 50
XPrt.CurrentX = 300
XPrt.Print "Titulaire";
XPrt.CurrentX = 900
XPrt.Print ":";
XPrt.CurrentX = 1100
XPrt.Print mYBIAMVT0.COMPTEINT;
XPrt.CurrentX = 4350
XPrt.FontBold = True
XPrt.Print mResponsable;
XPrt.FontBold = False
End Sub


Public Sub prtYBIAMVT0_A4_Compte(lRELEVEREL As String)
Dim X20 As String * 20, X As String
Dim xId As String
Dim wCompte As String

solde = mYBIAMVT0.BIAMVTSD0
xYADRESS0 = zYADRESS0
fiscalYADRESS0 = zYADRESS0

Call fctPCEC_Atribut(mYBIAMVT0.COMPTEOBL, mYBIAMVT0.COMPTEDEV, blnCptOrdinaire, blnRIB, blnMédiateur)

mRib_Compte = Trim(mYBIAMVT0.MOUVEMCOM)
wCompte = mRib_Compte
mRib_Clé = Format$(RibClé(strSocBdfE, strSocBdfG, wCompte, mRib_IbanE), "00")

If mRib_Clé = 99 Then blnRIB = False: blnMédiateur = False
    
mResponsable = "   "

xId = constYRELEVE0 & mYBIAMVT0.MOUVEMCOM & lRELEVEREL  '' " " 'espace => 1 ère occurence J D M
If Not IsNull(srvYRELEVE0_Import_Read(xId, xYRELEVE0)) Then
    xId = constYRELEVE0 & mYBIAMVT0.MOUVEMCOM & "M"
    If Not IsNull(srvYRELEVE0_Import_Read(xId, xYRELEVE0)) Then
        xId = constYRELEVE0 & mYBIAMVT0.MOUVEMCOM & "D"
        If Not IsNull(srvYRELEVE0_Import_Read(xId, xYRELEVE0)) Then
            recYRELEVE0_Init xYRELEVE0
             xYRELEVE0.RELEVETYP = "1"
             xYRELEVE0.RELEVEADR = "CO"
        End If
    End If
End If
 
X = mYBIAMVT0.MOUVEMCOM
If Not IsNull(srvYTITULA0_Import_Read(X, xYTITULA0)) Then
    xYTITULA0.TITULACLI = ""                    ''''''''"00" & mId$(mYBIAMVT0.MOUVEMCOM, 1, 5)
End If

xId = constYCLIENA0 & xYTITULA0.TITULACLI
If IsNull(srvYCLIENA0_Import_Read(xId, xYCLIENA0)) Then
    mResponsable = xYCLIENA0.CLIENARES
Else
    mResponsable = ""
    xYCLIENA0.CLIENARA1 = mYBIAMVT0.COMPTEINT
    xYCLIENA0.CLIENARA2 = ""
End If

X = "1 " & xYTITULA0.TITULACLI

If Not IsNull(srvYADRESS0_Import_Read(X, fiscalYADRESS0)) Then recYADRESS0_Init fiscalYADRESS0


'$2003.07.15 JPL adresse de la racine associé au compte dans ZRELEVE0

If Trim(xYRELEVE0.RELEVENUM) = "" Then
    X20 = " " & xYTITULA0.TITULACLI
Else
    X20 = xYRELEVE0.RELEVENUM
End If
    
X = xYRELEVE0.RELEVETYP & X20 & xYRELEVE0.RELEVEADR
If Not IsNull(srvYADRESS0_Import_Read(X, xYADRESS0)) Then
    If xYRELEVE0.RELEVEADR <> "  " Then
        xMvtP0.ID = "1" & X20 & "  "
        '''If Not IsNull(srvYADRESS0_Import_Read(X, xYADRESS0)) Then xYADRESS0 = fiscalYADRESS0
        If Not IsNull(srvYADRESS0_Import_Read(X, xYADRESS0)) Then xYADRESS0 = fiscalYADRESS0
    End If
End If
If Trim(xYADRESS0.ADRESSRA1) = "" Then
    xYADRESS0.ADRESSRA1 = Trim(xYCLIENA0.CLIENARA1) & " " & Trim(xYCLIENA0.CLIENARA2) '  Trim(xYCLIENA0.CLIENAETA) & " " &
    If Trim(xYADRESS0.ADRESSRA1) = "" Then
        xYADRESS0.ADRESSRA1 = mYBIAMVT0.COMPTEINT
    End If
End If
End Sub

Public Sub prtYBIAMVT0_A4_Report()
XPrt.CurrentY = Line4 + 50
prtYBIAMVT0_A4_Montant (solde)
NbPage = NbPage + 1
frmElpPrt.prtNewPage
prtYBIAMVT0_A4_Form "Report", ""

End Sub

Public Sub prtYBIAMVT0_A4_OpenX_Reset()
Height8_6 = frmElpPrt.prtHeightDelta(8, 6)


prtTitleText = "Extrait de Compte"
prtPgmName = "prtYBIAMVT0_A4"
prtTitleUsr = usrName
prtFontName = prtFontName_Arial

prtLineNb = 1
prtlineHeight = 250

prtHeaderHeight = 300
prtOrientation = vbPRORPortrait

prtFormType = ""
prtSocInit
NbLigneMax = 35
'prtInit
col1 = prtMinX
col2 = col1 + 1100 '1325
col3 = col1 + 6100 '6025
Col4 = col1 + 7250 '6950
Col5 = col1 + 9075 '8925
Col6 = col1 + 10900
Col7 = col1 + 10900
Col8 = col1 + 10900

Line1 = prtlineHeight * 21

Line2 = Line1 + prtlineHeight + 50
Line3 = Line2 + prtlineHeight + 50
Line4 = Line3 + prtlineHeight * NbLigneMax + 50
Line5 = Line4 + prtlineHeight + 50

End Sub
