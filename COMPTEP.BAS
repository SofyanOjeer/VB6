Attribute VB_Name = "prtCompte"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------

Private recCompte As typeCompte
Dim I As Integer, Height8_6 As Integer

Dim chkNuméroAncien As Boolean
Dim chkDeviseCV As Boolean
Dim mDevCode As Integer

Private recCptInfo As typeCptInfo
'Private CV       As typeDevise
Private Mt As Currency
Private totalCV As Currency, totalDev As Currency

Dim blnBalance As Boolean

Dim mAMJCours As String * 8
'---------------------------------------------------------
Public Sub recCompte_Display(Xcompte As typeCompte, XpicBox As PictureBox, Optional vSolde)
'---------------------------------------------------------
Dim X As String, XX As String
Dim SituationColor As Long, Situation As String
Dim DécouvertColor As Long
Dim I As Integer, Mt As Currency
Dim blnSolde As Boolean
Dim wLineHeight As Integer, col1 As Integer, col2 As Integer, col3 As Integer, Col4 As Integer, Col5 As Integer
Dim wD As Double
Dim xCV As typeCV

If IsMissing(vSolde) Then
    blnSolde = True
Else
    blnSolde = vSolde
End If

wLineHeight = (XpicBox.Height - 20) / 4
col1 = 0: col2 = 450: col3 = 1900: Col4 = 8500: Col5 = 9200

If XpicBox.Width > 10000 Then
    wD = XpicBox.Width / 9350
    col2 = col2 * wD: col3 = col3 * wD: Col4 = Col4 * wD: Col5 = Col5 * wD
End If


XpicBox.Visible = True
XpicBox.Cls
'XpicBox.BackColor = frmUsr.BackColor

        If Xcompte.Situation = " " Then
            Situation = "": SituationColor = libUsr.ForeColor
        Else
            SituationColor = errUsr.ForeColor
            Select Case Xcompte.Situation
                Case "A"
                    Situation = "Annulé"
                Case "B"
                    Situation = "Bloqué"
                Case "E"
                    Situation = "Bascule EURO"
                Case "F"
                    Situation = "Autorisé Devise IN"
                Case "G"
                    Situation = "Réservé comptaG"
                Case Else
                    Situation = "? " & Xcompte.Situation
            End Select
        End If
        XpicBox.FontBold = True
    
        XpicBox.ForeColor = SituationColor
        
        XpicBox.CurrentY = 0
        XpicBox.CurrentX = 50: XpicBox.Print Format$(Xcompte.Devise, "000") & ".";
        
        If Xcompte.TypeGA = "A" Then
        
            XpicBox.Print Format$(Xcompte.Numéro, "@@@@@.@@@.@@.@");
        Else
            XpicBox.Print Format$(mId$(Xcompte.Numéro, 4, 8), "@@@ @@@ @.@");
        End If
        
        XpicBox.CurrentX = col3:
        If Xcompte.Groupe <> "000000" Then
            ''XpicBox.CurrentX = Col1
            XpicBox.FontBold = True
            XpicBox.ForeColor = errUsr.ForeColor
            XpicBox.Print "Groupe : " & Xcompte.Groupe & " ";
            XpicBox.FontBold = False
        End If
        XpicBox.ForeColor = libUsr.ForeColor
        XpicBox.Print Trim(Xcompte.Intitulé);

        CV_X1.DeviseN = Format$(Xcompte.Devise, "000")
        CV_AttributN CV_X1
       ' XX = DevX(Xcompte.Devise)
        If Not blnSolde Then Exit Sub
        
        If Xcompte.SoldeInstantané < 0 Then XpicBox.ForeColor = errUsr.ForeColor
          
        
        X = Format$(Xcompte.SoldeInstantané, "#### ### ### ### ##0.00")
        XpicBox.CurrentX = Col4 - XpicBox.TextWidth(X)
        XpicBox.Print X;
        XpicBox.CurrentX = Col4 + 200
        XpicBox.FontBold = False
        XpicBox.Print CV_X1.DeviseIso;
        XpicBox.ForeColor = &H4000&
        XpicBox.Print " " & Xcompte.MvtceJour;

        XpicBox.ForeColor = libUsr.ForeColor
        XpicBox.CurrentY = wLineHeight '200
        
        If frmCompte.chkDeviseCV Then
            XpicBox.FontBold = True
            XpicBox.ForeColor = warnUsrColor
 '           Mt = Xcompte.SoldeInstantané * XDevise.Cours / CV.Cours
 
            CV_X1.Montant = Xcompte.SoldeInstantané
            xCV = CV_X2
            If CV_X1.DeviseN = 978 Then xCV.DeviseIso = "FRF"
            
            Call CV_Transitoire(CV_X1, xCV, CV_X3, X)
            X = Format$(xCV.Montant, "#### ### ### ### ##0.00")
            XpicBox.CurrentX = Col4 - XpicBox.TextWidth(X)
            XpicBox.Print X;
            XpicBox.CurrentX = Col4 + 200
            XpicBox.FontBold = False
            XpicBox.Print xCV.DeviseIso;
            
        End If
        XpicBox.FontBold = False
       
        
        If frmCompte.chkNuméroAncien Then
            XpicBox.ForeColor = SituationColor
            XpicBox.CurrentX = col2: XpicBox.Print Format$(mId$(Xcompte.NuméroAncien, 1, 8), "@@@@@@.@@");
        End If
        
        XpicBox.ForeColor = libUsr.ForeColor
        XpicBox.CurrentX = col2: XpicBox.Print Xcompte.Intitulé2;

        XpicBox.CurrentY = wLineHeight * 2 '390
        XpicBox.ForeColor = SituationColor
        XpicBox.CurrentX = col2: XpicBox.Print Situation;
        XpicBox.ForeColor = txtUsr.ForeColor 'libUsr.forecolor
        XpicBox.CurrentX = col3
        If Xcompte.TypeGA = "A" Then
            XpicBox.Print Trim(DicLib(13, Xcompte.BiaTyp)) & "  ";
        End If
        XpicBox.Print " ( " & Trim(CV_X1.DeviseLibellé) & " ) ";
            XpicBox.CurrentY = wLineHeight * 3  '570
            If Xcompte.DécouvertMontant > 0 Then
            XpicBox.ForeColor = txtUsr.ForeColor
            XpicBox.CurrentX = col3:
            If Val(Xcompte.DécouvertAmj) < DSys Then
                XpicBox.ForeColor = warnUsrColor
            Else
                If Xcompte.SoldeInstantané + Xcompte.DécouvertMontant < 0 Then
                    XpicBox.ForeColor = errUsr.ForeColor
                End If
            End If
            X = "Découvert autorisé : " & Trim(Format$(Xcompte.DécouvertMontant, "### ### ### ###")) _
                & " juqu'au : " & dateImp(Xcompte.DécouvertAmj)
            XpicBox.Print X;
        End If
        
        
        If Xcompte.ChéquierInterdit <> "0" Then
            XpicBox.CurrentX = 0
            XpicBox.FontBold = True
            XpicBox.ForeColor = errUsr.ForeColor
            XpicBox.Print "Interdit de chéquier";
            XpicBox.FontBold = False
        End If

        XpicBox.ForeColor = libUsr.ForeColor
        X = dateImp(Xcompte.MvtAmj)
        XpicBox.CurrentX = Col4 - XpicBox.TextWidth(X)
        XpicBox.Print X;
End Sub





'---------------------------------------------------------
 Public Sub prtCompteX(Msg As String)
'---------------------------------------------------------
Dim K As Integer, K1 As Integer, K2 As Integer
Dim X As String
On Error GoTo prtError

Set XPrt = Printer
Height8_6 = frmElpPrt.prtHeightDelta(8, 6)
K1 = Val(mId$(Msg, 1, 6))
K2 = Val(mId$(Msg, 7, 6))

If mId$(Msg, 13, 1) = " " Then
    chkNuméroAncien = False
Else
    chkNuméroAncien = True
End If

If mId$(Msg, 14, 3) = "   " Then
    chkDeviseCV = False
Else
    chkDeviseCV = True
'99-01-20    Call DevCode(Mid$(Msg, 14, 3))
'99-01-20    CV = XDevise

    Call CV_AttributS(mId$(Msg, 14, 3), CV_X2)
End If

Select Case mId$(Msg, 18, 1)
    Case "I": X = "solde instantané": mAMJCours = DSys
    Case "V": X = "solde veille": mAMJCours = DSys
    Case "M": X = "solde fin de mois (date de traitement)": mAMJCours = dateElp("FinDeMoisP", 0, DSys)
    Case "O": X = "solde fin de mois (date d'opération)": mAMJCours = dateElp("FinDeMoisP", 0, DSys)
    Case "A": X = "solde fin d'année": mAMJCours = dateElp("FinDAnnéeP", 0, DSys)
    Case Else: X = "solde ?"
End Select
CV_X1.OpéAmj = mAMJCours
CV_X2.OpéAmj = mAMJCours
CV_X3.OpéAmj = mAMJCours

If mId$(Msg, 17, 1) = "B" Then
    prtTitleText = "Balance de comptes - " & X
    prtLineNb = 1
    blnBalance = True
Else
    prtTitleText = "Liste de comptes - " & X
    prtLineNb = 3
    blnBalance = False
End If

mDevCode = 0: totalDev = 0: totalCV = 0

frmElpPrt.Show vbModeless

recCompteInit recCompte

prtOrientation = vbPRORLandscape
prtPgmName = "prtCompte"
prtTitleUsr = usrName

prtlineHeight = 250
prtHeaderHeight = 300

frmElpPrt.prtStdInit
prtCompteForm

For K = K1 To K2
recCompte = arrCompte(K)
prtCompteLine

DoEvents: If prtKillDoc Then frmElpPrt.Hide: Exit Sub

Next K
XPrt.DrawWidth = 5
frmElpPrt.prtLineY

K = 0
XPrt.CurrentY = XPrt.CurrentY + 50
XPrt.FontBold = True

If mDevCode > 0 Then
    K = 1
    X = Format$(totalDev, "#### ### ### ### ##0.00")
    If totalDev >= 0 Then
        XPrt.CurrentX = prtMinX + 15300 - XPrt.TextWidth(X)
    Else
        XPrt.CurrentX = prtMinX + 13000 - XPrt.TextWidth(X)
    End If

    XPrt.Print X & " ";
End If

If chkDeviseCV Then
    K = 1
    X = Format$(totalCV, "#### ### ### ### ##0.00")
    XPrt.CurrentX = prtMinX + 9700 - XPrt.TextWidth(X)
    XPrt.Print X;
End If

prtCurrentY = XPrt.CurrentY + prtlineHeight + 50
If K > 0 Then: XPrt.Line (prtMinX + 11100, prtCurrentY)-(prtMaxX, prtCurrentY)

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
Public Sub prtCptInfoX(Msg As String)
'---------------------------------------------------------
Dim K As Integer, K1 As Integer, K2 As Integer
Dim X As String
On Error GoTo prtError


Set XPrt = Printer

K1 = Val(mId$(Msg, 1, 6))
K2 = Val(mId$(Msg, 7, 6))


frmElpPrt.Show vbModeless


prtOrientation = vbPRORPortrait
prtTitleText = "Fiche compte"
prtPgmName = "prtCptInfo"
prtTitleUsr = usrName

prtLineNb = 1
prtlineHeight = 250
prtHeaderHeight = 300

frmElpPrt.prtStdInit
prtCptInfoForm

For K = K1 To K2
recCptInfo = arrCptInfo(K)
prtCptInfoLine

DoEvents: If prtKillDoc Then frmElpPrt.Hide: Exit Sub

Next K

'frmElpPrt.prtLineY
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
Public Sub prtCompteForm()
'---------------------------------------------------------
Dim X As String

XPrt.FontSize = 8
XPrt.FontBold = True
XPrt.DrawWidth = 3

'XPrt.FillStyle = 0
'XPrt.ForeColor = RGB(0, 0, 0)
'XPrt.FillStyle = 1

Call frmElpPrt.prtTrame(prtMinX, prtMinY, prtMaxX, prtMinY + prtHeaderHeight, "B")
Call frmElpPrt.prtTrame(prtMinX + 11100, prtMinY + prtHeaderHeight + 10, prtMinX + 13500, prtMaxY - 10, " ")
'XPrt.Line (prtMinX, prtMinY)-(prtMinX, prtMaxY)
'XPrt.Line (prtMaxX, prtMinY)-(prtMaxX, prtMaxY)

XPrt.DrawWidth = 2

XPrt.Line (prtMinX + 11100, prtMinY)-(prtMinX + 11100, prtMaxY)

XPrt.Line (prtMinX + 1600, prtMinY)-(prtMinX + 1600, prtMaxY)
'---------------------------------------------------------

X = "N°de Compte"
XPrt.CurrentY = prtMinY + (prtHeaderHeight - XPrt.TextHeight(X)) / 2
XPrt.CurrentX = prtMinX + 300
XPrt.Print X;

X = "Intitulé"
XPrt.CurrentX = prtMinX + 1750
XPrt.Print X;

XPrt.CurrentX = 10200
XPrt.Print "Dernier Mvt le";

If chkDeviseCV Then
    XPrt.CurrentX = 9000
    XPrt.Print "Contre-valeur";
End If

X = "Débit"
XPrt.CurrentX = 12950
XPrt.Print X;

X = "Crédit"
XPrt.CurrentX = 15250
XPrt.Print X;
prtCurrentY = prtMinY + prtHeaderHeight

End Sub




'---------------------------------------------------------
Public Sub prtCompteLine()
'---------------------------------------------------------
Dim X As String, K As Integer, wsdCurrentX As Integer
Dim Situation As String

If prtCurrentY + prtParagraphHeight > prtMaxY Then
    frmElpPrt.prtNewPage
    prtCompteForm
'Else
    'frmElpPrt.prtLineY
End If

XPrt.FontBold = False
'------------------------------------------ligne 1--------------

XPrt.CurrentY = prtCurrentY + prtlineHeight - XPrt.TextHeight("test")
XPrt.FontSize = 8

XPrt.CurrentX = prtMinX + 50
XPrt.Print Format$(recCompte.Devise, "000") & ".";

If recCompte.TypeGA = "A" Then
    XPrt.Print Format$(recCompte.Numéro, "@@@@@.@@@.@@.@");
Else
    XPrt.Print Format$(recCompte.Numéro, "@@@@ @@@ @.@");
End If

XPrt.FontSize = 6
XPrt.CurrentY = XPrt.CurrentY + Height8_6
XPrt.CurrentX = prtMinX + 1750
    If recCompte.Groupe <> "000000" Then
        XPrt.FontItalic = True: XPrt.FontUnderline = True
        XPrt.Print "Groupe : " & recCompte.Groupe & " ";
        XPrt.FontItalic = False: XPrt.FontUnderline = False
    End If
XPrt.Print recCompte.Intitulé;

XPrt.CurrentX = prtMinX + 5500
If recCompte.TypeGA = "A" Then
    XPrt.Print Trim(DicLib(13, recCompte.BiaTyp));
    
End If
Select Case recCompte.Situation
    Case " ": Situation = ""
    Case "A": Situation = " **Annulé**"
    Case "B": Situation = " **Bloqué**"
    Case Else: Situation = " ?? " & recCompte.Situation
End Select

XPrt.FontBold = True
XPrt.Print Situation;

XPrt.FontSize = 8
XPrt.CurrentY = XPrt.CurrentY - Height8_6
       
X = Format$(recCompte.SoldeInstantané, "#### ### ### ### ##0.00")
If recCompte.SoldeInstantané >= 0 Then
    wsdCurrentX = prtMinX + 15300
Else
    wsdCurrentX = prtMinX + 13000
End If

XPrt.CurrentX = wsdCurrentX - XPrt.TextWidth(X)

K = Val(recCompte.Devise)
Call DevX(K)
If K <> mDevCode Then
    mDevCode = IIf(mDevCode = 0, K, -1)
End If
totalDev = totalDev + recCompte.SoldeInstantané
XPrt.Print X & " ";
XPrt.FontBold = False
XPrt.FontSize = 6
XPrt.CurrentY = XPrt.CurrentY + Height8_6
XPrt.Print XDevise.DevX;
    XPrt.FontSize = 8
    XPrt.CurrentY = XPrt.CurrentY - Height8_6
XPrt.CurrentX = wsdCurrentX + 400
XPrt.Print recCompte.MvtceJour;

XPrt.FontSize = 6
XPrt.CurrentY = XPrt.CurrentY + Height8_6
X = dateImp(recCompte.MvtAmj)
XPrt.CurrentX = 11200 - XPrt.TextWidth(X)
XPrt.Print X;

If chkDeviseCV Then
    XPrt.FontBold = True
'    Mt = recCompte.SoldeVeille * XDevise.Cours / CV.Cours
'    X = Format$(Mt, "#### ### ### ### ##0.00")
    CV_X1.DeviseN = Format$(recCompte.Devise, "000")
    CV_X1.DeviseIso = ""
    CV_X1.Montant = recCompte.SoldeInstantané
    Call CV_Transitoire(CV_X1, CV_X2, CV_X3, X)
    X = Format$(CV_X2.Montant, "#### ### ### ### ##0.00")
    XPrt.CurrentX = prtMinX + 9700 - XPrt.TextWidth(X)
    XPrt.Print X & " ";
    XPrt.FontBold = False
    XPrt.Print CV_X2.DeviseIso;
    totalCV = totalCV + CV_X2.Montant
End If
    
XPrt.FontSize = 8
XPrt.CurrentY = XPrt.CurrentY - Height8_6
    
XPrt.FontBold = False

If blnBalance Then prtCurrentY = prtCurrentY + prtParagraphHeight: Exit Sub

'------------------------------------------ligne 2--------------

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
       
    
XPrt.FontBold = True
XPrt.CurrentX = prtMinX + 400
XPrt.Print Situation;
XPrt.FontBold = False

XPrt.CurrentX = prtMinX + 1850
XPrt.FontSize = 6
XPrt.CurrentY = XPrt.CurrentY + Height8_6

XPrt.Print recCompte.Intitulé2;
If recCompte.DécouvertMontant > 0 Then
    XPrt.CurrentX = prtMinX + 5500:
            X = "Découvert autorisé : " & Format$(recCompte.DécouvertMontant, "### ### ### ###") _
          & "  juqu'au : " & dateImp(recCompte.DécouvertAmj)
        XPrt.Print X;
End If
'------------------------------------------ligne 3--------------

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight - Height8_6


If chkNuméroAncien Then
    XPrt.CurrentX = prtMinX + 50
    XPrt.Print Format$(mId$(recCompte.NuméroAncien, 1, 8), "@@@@@@.@@");
End If
prtCurrentY = prtCurrentY + prtParagraphHeight
End Sub




'---------------------------------------------------------
Public Sub prtCptInfoForm()
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

XPrt.CurrentX = 7700
XPrt.Print "Agence";
XPrt.CurrentX = 9100
XPrt.Print ":";
XPrt.CurrentX = 9400
XPrt.Print strSocBdfE & " " & strSocBdfG;

prtCurrentY = prtMinY + prtHeaderHeight


End Sub

'---------------------------------------------------------
Public Sub prtCptInfoLine()
'---------------------------------------------------------
Dim X As String, K As Integer, mCurrenty As Integer

If prtCurrentY + prtParagraphHeight > prtMaxY Then
    frmElpPrt.prtNewPage
    prtCompteForm
'Else
    'frmElpPrt.prtLineY
End If
XPrt.DrawWidth = 1
XPrt.FontBold = False

'------------------------------------------ligne 1--------------
XPrt.CurrentY = prtCurrentY + prtlineHeight + 50 - XPrt.TextHeight("test")
XPrt.FontSize = 8
XPrt.CurrentX = 400
XPrt.Print "Numéro de compte";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.FontBold = True
XPrt.CurrentX = 2450
XPrt.Print Format$(recCptInfo.Devise, "000") & ".";
XPrt.Print Compte_Imp(recCptInfo.Numéro);

XPrt.FontBold = False
XPrt.CurrentX = 7700
XPrt.Print "Ancien Compte";
XPrt.CurrentX = 9100
XPrt.Print ":";
XPrt.FontBold = True
XPrt.CurrentX = 9400
XPrt.Print Format$(recCptInfo.NuméroAncien, "@@@@@@.@@.@@@");

'---------------------------------- 2--------------
XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Intitulé";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.FontBold = True
XPrt.CurrentX = 2450
XPrt.Print recCptInfo.Intitulé;
XPrt.FontBold = False

XPrt.CurrentX = 7700
XPrt.Print "Recherche Alpha";
XPrt.CurrentX = 9100
XPrt.Print ":";
XPrt.FontBold = True
XPrt.CurrentX = 9400
XPrt.Print recCptInfo.Alpha;
     
'--------------------------------- 3------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.FontBold = False

XPrt.CurrentX = 400
XPrt.Print "Intitulé Complémentaire";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.CurrentX = 2450
XPrt.Print recCptInfo.Intitulé2;

XPrt.CurrentX = 7700
XPrt.Print "Rib";
XPrt.CurrentX = 9100
XPrt.Print ":";
XPrt.FontBold = True

XPrt.CurrentX = 9400
XPrt.Print recCptInfo.CléRib;

'--------------------------------- 3------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.FontBold = False
XPrt.CurrentX = 2450
If recCptInfo.TypeGA = "A" Then XPrt.Print DicLib(13, recCptInfo.BiaTyp);

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Situation";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.CurrentX = 2450
XPrt.Print recCptInfo.Situation;
XPrt.CurrentX = 3000
XPrt.FontBold = True
Select Case recCptInfo.Situation
    Case "A"
        XPrt.Print "Annulé";
    Case "B"
        XPrt.Print "Bloqué";
End Select
'-----------------------4----
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY)
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight

XPrt.FontBold = False
XPrt.CurrentX = 400
XPrt.Print "Adresse ligne 1";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.CurrentX = 2450
XPrt.Print recCptInfo.Adresse1;

'--------------------------5--------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Adresse ligne 2";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.CurrentX = 2450
XPrt.Print recCptInfo.Adresse2;
'-------------------------6----------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Adresse ligne 3";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.CurrentX = 2450
XPrt.Print recCptInfo.Adresse3;
'-------------------------7--------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Adresse ligne 4";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.CurrentX = 2450
XPrt.Print recCptInfo.Adresse4;
'---------------------------8---------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Adresse ligne 5";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.CurrentX = 2450
XPrt.Print recCptInfo.Adresse5;
'------------------------------9-------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Adresse ligne 6";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.CurrentX = 2450
XPrt.Print recCptInfo.AdresseCP;
'--------------------------10-------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Adresse ligne 7";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.CurrentX = 2450
XPrt.Print recCptInfo.AdresseBD;
'---------------------------11------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Adresse ligne 8";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.CurrentX = 2450
XPrt.Print recCptInfo.AdressePays;

'--------------------12----------COMPTABILITE--------------------

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY)
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Résident BDF";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.CurrentX = 2450
XPrt.FontBold = True
XPrt.Print recCptInfo.Résident;
'-------------------------------------
XPrt.FontBold = False
XPrt.CurrentX = 3000
XPrt.Print DicLib(23, recCptInfo.Résident);

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Actionnaire";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.CurrentX = 2450
XPrt.FontBold = True
XPrt.Print recCptInfo.Actionnaire;
'-------------------------------------
XPrt.FontBold = False
XPrt.CurrentX = 3000
XPrt.Print DicLib(7, recCptInfo.Actionnaire);

'-------------------------------------
If recCptInfo.TypeGA = "A" Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = 400
    XPrt.Print "Nature de titulaire";
    XPrt.CurrentX = 2200
    XPrt.Print ":";
    XPrt.CurrentX = 2450
    XPrt.FontBold = True
    XPrt.Print recCptInfo.NatureTitulaire;
    XPrt.FontBold = False
    XPrt.CurrentX = 3000
    K = IIf(recCptInfo.Numéro < 30000000000#, 18, 62)
    XPrt.Print DicLib(K, recCptInfo.NatureTitulaire);
End If

'-------------------------13---------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
 
XPrt.CurrentX = 400
XPrt.Print "Type de Compte";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.FontBold = True
XPrt.CurrentX = 2450
XPrt.Print recCptInfo.TypeGA;
XPrt.FontBold = False

XPrt.CurrentX = 3000
Select Case recCptInfo.TypeGA
  Case "A"
     XPrt.Print "Auxiliaire";
  Case "g"
        XPrt.Print "Général";
  Case "R"
        XPrt.Print "Regroupement";
End Select

'--------------------------- 14-----------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.FontBold = False
     
XPrt.CurrentX = 400
XPrt.Print "Compte Général";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.FontBold = True
XPrt.CurrentX = 3000
XPrt.Print Compte_Imp(recCptInfo.CompteGénéral);
If recCompte.Numéro <> recCptInfo.CompteGénéral Then
    recCompteInit recCompte
    recCompte.Method = "SeekL1"
    recCompte.Société = recCptInfo.Société
    recCompte.Agence = recCptInfo.Agence
    recCompte.Devise = recCptInfo.Devise
    recCompte.Numéro = recCptInfo.CompteGénéral
    recCompte.BiaTyp = "000"
    recCompte.BiaNum = "00"
    
    srvCompteMon recCompte
End If
XPrt.FontBold = False

XPrt.Print "  " & recCompte.Intitulé;

'----------------------------14bis-------------
XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Sens";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.FontBold = True

XPrt.CurrentX = 2450
XPrt.Print recCptInfo.Sens;
XPrt.FontBold = False

XPrt.CurrentX = 3000
XPrt.Print DicLib(2, recCptInfo.Sens);

'--------------------------15
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Nostro/Loro";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.FontBold = True

XPrt.CurrentX = 2450
XPrt.Print recCptInfo.Nostro;
XPrt.FontBold = False

XPrt.CurrentX = 3000
Select Case recCptInfo.Nostro
  Case "L"
     XPrt.Print "Loro";
  Case "N"
        XPrt.Print "Nostro";
End Select

'------------------------------16
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
 
XPrt.CurrentX = 400
XPrt.Print "Service responsable";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.FontBold = True
XPrt.CurrentX = 2450
XPrt.Print recCptInfo.ServiceResponsable;
XPrt.FontBold = False
XPrt.CurrentX = 3000
If recCptInfo.ServiceResponsable <> "000" Then XPrt.Print DicLib(4, recCptInfo.ServiceResponsable);

'------------------------------17
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Service autorisé 1";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.FontBold = True
XPrt.CurrentX = 2450
XPrt.Print recCptInfo.ServiceAutorisé1;
XPrt.FontBold = False
XPrt.CurrentX = 3000
If recCptInfo.ServiceAutorisé1 <> "000" Then XPrt.Print DicLib(4, recCptInfo.ServiceAutorisé1);

'-------------------------------18
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.FontBold = False
XPrt.CurrentX = 400
XPrt.Print "Service autorisé 2";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.FontBold = True
XPrt.CurrentX = 2450
XPrt.Print recCptInfo.ServiceAutorisé2;
XPrt.FontBold = False
XPrt.CurrentX = 3000
If recCptInfo.ServiceAutorisé2 <> "000" Then XPrt.Print DicLib(4, recCptInfo.ServiceAutorisé2);

'-------------------------------19
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.FontBold = False
XPrt.CurrentX = 400
XPrt.Print "Service autorisé 3";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.FontBold = True
XPrt.CurrentX = 2450
XPrt.Print recCptInfo.ServiceAutorisé3;
XPrt.FontBold = False
XPrt.CurrentX = 3000
If recCptInfo.ServiceAutorisé3 <> "000" Then XPrt.Print DicLib(4, recCptInfo.ServiceAutorisé3);

'--------------------------------20
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.FontBold = False
XPrt.CurrentX = 400
XPrt.Print "Service autorisé 4";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.FontBold = True
XPrt.CurrentX = 2450
XPrt.Print recCptInfo.ServiceAutorisé4;
XPrt.FontBold = False
XPrt.CurrentX = 3000
If recCptInfo.ServiceAutorisé4 <> "000" Then XPrt.Print DicLib(4, recCptInfo.ServiceAutorisé4);

'---------------------------------21
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.FontBold = False
XPrt.CurrentX = 400
XPrt.Print "Service autorisé 5";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.FontBold = True
XPrt.CurrentX = 2450
XPrt.Print recCptInfo.ServiceAutorisé5;
XPrt.FontBold = False
XPrt.CurrentX = 3000
If recCptInfo.ServiceAutorisé5 <> "000" Then XPrt.Print DicLib(4, recCptInfo.ServiceAutorisé5);

'----------------------------22--GESTION-----------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY)
'-----------------------------23-------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight

XPrt.CurrentX = 400
XPrt.Print "Conditions";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.FontBold = True
XPrt.CurrentX = 2450
XPrt.Print recCptInfo.Conditions;
XPrt.FontBold = False
XPrt.CurrentX = 3000
XPrt.Print DicLib(522, recCptInfo.Conditions);

'------------------------------24-------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Prélèvement libératoire";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.FontBold = True
XPrt.CurrentX = 2450
XPrt.Print recCptInfo.PrélèvementLibératoire;
XPrt.FontBold = False
XPrt.CurrentX = 3000
XPrt.Print DicLib(7, recCptInfo.PrélèvementLibératoire);

'-------------------------------25-----------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Courrier";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.FontBold = True
XPrt.CurrentX = 2450
XPrt.Print recCptInfo.Courrier;
XPrt.FontBold = False
XPrt.CurrentX = 3000
XPrt.Print DicLib(61, recCptInfo.Courrier);

'--------------------------------26----------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Gestionnaire";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.FontBold = True
XPrt.CurrentX = 2450
XPrt.Print recCptInfo.Gestionnaire;
XPrt.FontBold = False
XPrt.CurrentX = 3000
XPrt.Print DicLib(60, recCptInfo.Gestionnaire);

'--------------------------------27-------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Apporteur";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.FontBold = True
XPrt.CurrentX = 2450
XPrt.Print recCptInfo.Apporteur;
XPrt.FontBold = False
XPrt.CurrentX = 3000
XPrt.Print DicLib(58, recCptInfo.Apporteur);

'-------------------------------28--
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Regroupement Client";
XPrt.CurrentX = 2200
XPrt.Print ":";
If recCptInfo.GroupeIdentification <> "000000" Then
    XPrt.FontBold = True
    XPrt.CurrentX = 2450
    XPrt.Print recCptInfo.GroupeIdentification;
    XPrt.FontBold = False
    XPrt.CurrentX = 3000
    XPrt.Print DicLib(65, recCptInfo.GroupeIdentification);
End If
'-----------------------------29 bis------------
XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.CurrentX = 400
XPrt.Print "Extrait";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.FontBold = True
XPrt.CurrentX = 2450
XPrt.Print recCptInfo.Extrait;
XPrt.FontBold = False
XPrt.CurrentX = 3000
XPrt.Print DicLib(6, recCptInfo.Extrait);

XPrt.CurrentX = 8200
XPrt.Print "Report";
XPrt.CurrentX = 8700
XPrt.Print ":";
XPrt.FontBold = True

X = Format$(recCptInfo.ExtraitSolde, "#### ### ### ### ##0.00")
XPrt.CurrentX = prtMinX + 10200 - XPrt.TextWidth(X)
XPrt.Print X;
XPrt.FontBold = False
X = IIf(recCptInfo.ExtraitSolde < 0, "Db", "Cr")
XPrt.CurrentX = 10600
XPrt.Print X;
XPrt.CurrentX = 4800
XPrt.Print "Date Dernier Extrait";
XPrt.CurrentX = 6400
XPrt.Print ":";
XPrt.FontBold = True
XPrt.CurrentX = 6550
XPrt.Print dateImp(recCptInfo.ExtraitAmj);
XPrt.FontBold = False
XPrt.CurrentX = 4000
XPrt.Print "N°";
XPrt.CurrentX = 4200
XPrt.Print ":";
XPrt.FontBold = True
XPrt.CurrentX = 4300
XPrt.Print recCptInfo.ExtraitNuméro;

'---------------------------30------------------
XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Echelles";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.FontBold = True
XPrt.CurrentX = 2450
XPrt.Print recCptInfo.Echelle;
XPrt.FontBold = False
XPrt.CurrentX = 3000
XPrt.Print DicLib(7, recCptInfo.Echelle);

XPrt.CurrentX = 8200
XPrt.Print "S.E.V.";
XPrt.CurrentX = 8700
XPrt.Print ":";
XPrt.FontBold = True
X = Format$(recCptInfo.EchelleSolde, "#### ### ### ### ##0.00")
XPrt.CurrentX = prtMinX + 10200 - XPrt.TextWidth(X)
XPrt.Print X;
 
XPrt.FontBold = False
X = IIf(recCptInfo.EchelleSolde < 0, "Db", "Cr")
XPrt.CurrentX = 10600
XPrt.Print X;

XPrt.CurrentX = 4800
XPrt.Print "Date arrêté échelles";
XPrt.CurrentX = 6400
XPrt.Print ":";
XPrt.FontBold = True
XPrt.CurrentX = 6550
XPrt.Print dateImp(recCptInfo.EchelleAmj);

'---------------------------------29----------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.FontBold = False
XPrt.CurrentX = 400
XPrt.Print "Découvert Autorisé";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.FontBold = True

X = Format$(recCptInfo.DécouvertMontant, "#### ### ### ### ##0.00")
XPrt.CurrentX = prtMinX + 10200 - XPrt.TextWidth(X)
XPrt.Print X;
XPrt.FontBold = False

XPrt.CurrentX = 4800
XPrt.Print "Echéance";
XPrt.CurrentX = 6400
XPrt.Print ":";
XPrt.FontBold = True
XPrt.CurrentX = 6550
XPrt.Print dateImp(recCptInfo.DécouvertAmj);

'---------------------------------29----------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.DrawWidth = 2
mCurrenty = XPrt.CurrentY

XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY)
XPrt.CurrentY = XPrt.CurrentY + 50
XPrt.FontBold = True
XPrt.CurrentX = 400
XPrt.Print DevX(recCptInfo.Devise);

XPrt.FontBold = False
XPrt.CurrentX = 3000
XPrt.Print "Cumul Débit";
XPrt.CurrentX = 5000
XPrt.Print "Cumul Crédit";
XPrt.CurrentX = 7400
XPrt.Print "Solde Débiteur";
XPrt.CurrentX = 9400
XPrt.Print "Solde Créditeur";

'---------------------------------34-----------------
    
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY)
XPrt.CurrentY = XPrt.CurrentY + 50
XPrt.CurrentX = 400
XPrt.Print "Solde Instantané";
XPrt.FontBold = True
    
K = IIf(recCptInfo.SoldeInstantané < 0, 8200, 10200)
X = Format$(recCptInfo.SoldeInstantané, "#### ### ### ### ##0.00")
XPrt.CurrentX = prtMinX + K - XPrt.TextWidth(X)
XPrt.Print X & " " & recCptInfo.MvtceJour;
'---------------------------41-----------
XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Solde en valeur";
'Xprt.FontBold = True
 
K = IIf(recCptInfo.SoldeEnValeurVeille < 0, 8200, 10200)
X = Format$(recCptInfo.SoldeEnValeurVeille, "#### ### ### ### ##0.00")
XPrt.CurrentX = prtMinX + K - XPrt.TextWidth(X)
XPrt.Print X;
'----------------------------------------
'Xprt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight + 50
XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY)
XPrt.CurrentY = XPrt.CurrentY + 50
XPrt.CurrentX = 400
XPrt.Print "Veille";
X = Format$(recCptInfo.DébitExercice, "#### ### ### ### ##0.00")
XPrt.CurrentX = prtMinX + 3800 - XPrt.TextWidth(X)
XPrt.Print X;
   
X = Format$(recCptInfo.CréditExercice, "#### ### ### ### ##0.00")
XPrt.CurrentX = prtMinX + 5800 - XPrt.TextWidth(X)
XPrt.Print X;

'Xprt.FontBold = True

K = IIf(recCptInfo.SoldeInstantané < 0, 8200, 10200)
X = Format$(recCptInfo.SoldeInstantané, "#### ### ### ### ##0.00")
XPrt.CurrentX = prtMinX + K - XPrt.TextWidth(X)
XPrt.Print X;

'-----------------------------38------------
'Xprt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Fin de Mois";

X = Format$(recCptInfo.DébitFindeMois, "#### ### ### ### ##0.00")
XPrt.CurrentX = prtMinX + 3800 - XPrt.TextWidth(X)
XPrt.Print X;

X = Format$(recCptInfo.CréditFindeMois, "#### ### ### ### ##0.00")
XPrt.CurrentX = prtMinX + 5800 - XPrt.TextWidth(X)
XPrt.Print X;
'Xprt.FontBold = True

Mt = recCptInfo.DébitFindeMois + recCptInfo.CréditFindeMois
K = IIf(Mt < 0, 8200, 10200)
X = Format$(Mt, "#### ### ### ### ##0.00")
XPrt.CurrentX = prtMinX + K - XPrt.TextWidth(X)
XPrt.Print X;

'---------------------------------------40
'Xprt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Exercice antérieur";
X = Format$(recCptInfo.DébitExerciceAntérieur, "#### ### ### ### ##0.00")
XPrt.CurrentX = prtMinX + 3800 - XPrt.TextWidth(X)
XPrt.Print X;

X = Format$(recCptInfo.CréditExerciceAntérieur, "#### ### ### ### ##0.00")
XPrt.CurrentX = prtMinX + 5800 - XPrt.TextWidth(X)
XPrt.Print X;
'Xprt.FontBold = True

Mt = recCptInfo.DébitExerciceAntérieur + recCptInfo.CréditExerciceAntérieur
K = IIf(Mt < 0, 8200, 10200)
X = Format$(Mt, "#### ### ### ### ##0.00")
XPrt.CurrentX = prtMinX + K - XPrt.TextWidth(X)
XPrt.Print X;


XPrt.CurrentY = XPrt.CurrentY + prtlineHeight + 50

XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY)
XPrt.Line (2200, mCurrenty)-(2200, XPrt.CurrentY)
XPrt.Line (6400, mCurrenty)-(6400, XPrt.CurrentY)

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.FontBold = False
XPrt.CurrentX = 400
XPrt.Print "Dernier mouvement";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.FontBold = True
XPrt.CurrentX = 2450
XPrt.Print dateImp(recCptInfo.AmjDernierMouvement);
        
XPrt.DrawWidth = 2

'------------------------45-----------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.FontBold = False
 
XPrt.CurrentX = 400
XPrt.Print "Date Création";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.FontBold = True
XPrt.CurrentX = 2450
XPrt.Print dateImp(recCptInfo.AmjCréation);

'---------------------------46---------------
XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Date Modification";
XPrt.CurrentX = 2200
XPrt.Print ":";

XPrt.FontBold = True
XPrt.CurrentX = 2450
XPrt.Print dateImp(recCptInfo.AmjModification);

'------------------------47---------
XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Date Annulation";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.FontBold = True
XPrt.CurrentX = 2450
XPrt.Print dateImp(recCptInfo.AmjAnnulation);
'--------------------------48------------
XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Date Réactivation";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.CurrentX = 2450
XPrt.Print dateImp(recCptInfo.AmjRéactivation);
'-------------------------49-------------------

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 400
XPrt.Print "Nom opérateur";
XPrt.CurrentX = 2200
XPrt.Print ":";
XPrt.FontBold = True
XPrt.CurrentX = 2450
XPrt.Print recCptInfo.NomOpérateur;
'----------------------------------

prtCurrentY = prtCurrentY + prtParagraphHeight
        

End Sub



'----------------------------------
Public Sub prtRib(Msg As String)
'---------------------------------------------------------
Dim K As Integer, K1 As Integer, K2 As Integer, I As Integer
Dim X As String
On Error GoTo prtError

Set XPrt = Printer
K1 = Val(mId$(Msg, 1, 6))
K2 = Val(mId$(Msg, 7, 6))

frmElpPrt.Show vbModeless

prtOrientation = vbPRORPortrait
prtTitleText = "RIB"
prtPgmName = "prtRib"
prtTitleUsr = usrName

prtLineNb = 1
prtlineHeight = 250
prtHeaderHeight = 300

prtFormType = ""
frmElpPrt.prtInit

For K = K1 To K2


recCptInfo = arrCptInfo(K)
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


DoEvents: If prtKillDoc Then frmElpPrt.Hide: Exit Sub

Next K

'frmElpPrt.prtLineY
frmElpPrt.prtEndDoc
frmElpPrt.Hide

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

XPrt.FontSize = 8
XPrt.FontBold = True
XPrt.DrawWidth = 2

J = XPrt.CurrentY
XPrt.PaintPicture frmElpPrt.imgSocLogo.Picture _
                , 4800, J _
                , frmElpPrt.imgSocLogo.Width * 0.5 _
                , frmElpPrt.imgSocLogo.Height * 0.5
XPrt.PaintPicture frmElpPrt.imgSocSigle.Picture _
                , 7500, J + 200 _
                , frmElpPrt.imgSocSigle.Width * 0.3 _
              , frmElpPrt.imgSocSigle.Height * 0.3



XPrt.CurrentY = J + prtlineHeight
XPrt.FontSize = 11
XPrt.FontBold = True
frmElpPrt.prtCentré 2500, "BANQUE INTERCONTINENTALE ARABE"
'---------------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.FontBold = False
XPrt.FontSize = 8
'---------------------------------------------------
frmElpPrt.prtCentré 2500, "67, Avenue Franklin Roosevelt"
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
frmElpPrt.prtCentré 2500, "75008 PARIS"
'---------------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
frmElpPrt.prtCentré 2500, "Tél : 01.53.76.62.03"
XPrt.FontSize = 11
XPrt.FontBold = True
XPrt.CurrentX = 6750
XPrt.Print "RELEVE D'IDENTITE BANCAIRE / IBAN";
XPrt.FontBold = False
XPrt.FontSize = 8
'---------------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
frmElpPrt.prtCentré 2500, "Compensable à Paris - 27"
XPrt.FontSize = 7
XPrt.CurrentX = 7500
XPrt.Print "Cadre réservé au destinataire du relevé";
J = XPrt.CurrentY + prtlineHeight / 2 - 50
'---------------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Line (6800, J)-(7400, J)
XPrt.Line (10150, J)-(10800, J)

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

XPrt.Line (6800, XPrt.CurrentY)-(10800, XPrt.CurrentY)
XPrt.Line (6800, J)-(6800, XPrt.CurrentY)
XPrt.Line (10800, J)-(10800, XPrt.CurrentY)

'---------------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight - 200
J = XPrt.CurrentY
H2 = J + 80
XPrt.Line (150, H2)-(250, H2)
XPrt.Line (800, H2)-(1050, H2)
XPrt.Line (1650, H2)-(2150, H2)
XPrt.Line (2750, H2)-(3250, H2)
XPrt.Line (3550, H2)-(3850, H2)
XPrt.Line (4450, H2)-(6600, H2)

XPrt.Line (150, H2)-(150, H2 + 650)
XPrt.Line (6600, H2)-(6600, H2 + 650)
XPrt.Line (150, H2 + 650)-(6600, H2 + 650)

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
XPrt.CurrentX = 300
XPrt.Print strSocBdfE;
XPrt.CurrentX = 1100
XPrt.Print strSocBdfG;
XPrt.CurrentX = 1900
XPrt.Print Format$(recCptInfo.Numéro, "@@@@@@@@@@@");
XPrt.CurrentX = 3300
XPrt.Print recCptInfo.CléRib;
XPrt.CurrentX = 3900
XPrt.Print SocRibDom;
XPrt.FontSize = 7
XPrt.FontBold = True
XPrt.CurrentX = 7400
XPrt.Print recCptInfo.Intitulé;
XPrt.FontBold = False
  
If Trim(recCptInfo.Adresse2) <> "" Then
    XPrt.CurrentY = XPrt.CurrentY + 200 'prtlineHeight
    XPrt.CurrentX = 7400
    XPrt.Print recCptInfo.Adresse2;
End If


'-------------------------------------------------
If Trim(recCptInfo.Adresse3) <> "" Then
    XPrt.CurrentY = XPrt.CurrentY + 200 'prtlineHeight
    XPrt.CurrentX = 7400
    XPrt.Print recCptInfo.Adresse3;
End If
'---------------------------------------------------
If Trim(recCptInfo.Adresse4) <> "" Then
    XPrt.CurrentY = XPrt.CurrentY + 200 'prtlineHeight
    XPrt.CurrentX = 7400
    XPrt.Print recCptInfo.Adresse4;
End If

'---------------------------------------------------
If Trim(recCptInfo.Adresse5) <> "" Then
    XPrt.CurrentY = XPrt.CurrentY + 200 'prtlineHeight
    XPrt.CurrentX = 7400
    XPrt.Print recCptInfo.Adresse5;
End If
'----------------------------------------------------
If Trim(recCptInfo.AdresseCP) <> "" _
Or Trim(recCptInfo.AdresseBD) <> "" Then
    XPrt.CurrentY = XPrt.CurrentY + 200 'prtlineHeight
    XPrt.CurrentX = 7400
    If Trim(recCptInfo.AdresseCP) <> "" Then XPrt.Print recCptInfo.AdresseCP & "  ";
    XPrt.Print recCptInfo.AdresseBD;
End If
'-----------------------------------------------------
If Trim(recCptInfo.AdressePays) <> "" Then
    XPrt.CurrentY = XPrt.CurrentY + 200 'prtlineHeight
    XPrt.CurrentX = 7400
    XPrt.Print recCptInfo.AdressePays;
End If
'---------------------------------------------------
XPrt.CurrentY = J + prtlineHeight * 4
XPrt.CurrentX = 300
XPrt.Print "IBAN International Bank Account Number";
XPrt.CurrentX = 3900
XPrt.Print "Bank Identification Code";

J = XPrt.CurrentY
H2 = J + 80
XPrt.Line (150, H2)-(150, H2 + 650)
XPrt.Line (6600, H2)-(6600, H2 + 650)

XPrt.Line (150, H2)-(250, H2)
XPrt.Line (2950, H2)-(3850, H2)

XPrt.Line (5450, H2)-(6600, H2)
XPrt.Line (150, H2 + 650)-(6600, H2 + 650)

XPrt.CurrentY = J + prtlineHeight + 100
XPrt.FontSize = 10
XPrt.CurrentX = 300
X = "FR00" & strSocBdfE & strSocBdfG & Format$(recCptInfo.Numéro, "@@@@@@@@@@@") & Format$(recCptInfo.CléRib, "@@")
Call Iban_Calc(X, IbanE)
XPrt.Print Iban_Print(IbanE);
XPrt.CurrentX = 3900
XPrt.Print SocBicId;


End Sub





