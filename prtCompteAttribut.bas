Attribute VB_Name = "Module1"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------

Private recCompte As typeCompte
Dim I As Integer
Dim chkNuméroAncien As Boolean
Dim chkDeviseCV As Boolean
Dim mDevCode As Integer

Private recCptInfo As typeCptInfo
Private CV       As typeDevise
Private Mt As Currency
Private totalCV As Currency, totalDev As Currency
'---------------------------------------------------------
Public Sub prtCompteAttributForm()
'---------------------------------------------------------
Dim x As String

XPrt.FontSize = 8
XPrt.FontBold = True

XPrt.FillStyle = 0
XPrt.DrawWidth = 3
XPrt.ForeColor = RGB(0, 0, 0)
XPrt.FillStyle = 1

XPrt.Line (prtMinX, prtMinY)-(prtMinX, prtMaxY)
XPrt.Line (prtMaxX, prtMinY)-(prtMaxX, prtMaxY)
XPrt.Line (prtMinX, prtMinY + prtHeaderHeight)-(prtMaxX, prtMinY + prtHeaderHeight)

XPrt.DrawWidth = 2

XPrt.Line (prtMinX + 11100, prtMinY)-(prtMinX + 11100, prtMaxY)

XPrt.Line (prtMinX + 1600, prtMinY)-(prtMinX + 1600, prtMaxY)
'---------------------------------------------------------

x = "N°de Compte"
XPrt.CurrentY = prtMinY + (prtHeaderHeight - XPrt.TextHeight(x)) / 2
XPrt.CurrentX = prtMinX + 300
XPrt.Print x;

x = "Intitulé"
XPrt.CurrentX = 3000
XPrt.Print x;

If chkDeviseCV Then
    XPrt.CurrentX = 9950
    XPrt.Print "Contre-valeur";
End If

x = "Débit"
XPrt.CurrentX = 12950
XPrt.Print x;

x = "Crédit"
XPrt.CurrentX = 15250
XPrt.Print x;
prtCurrentY = prtMinY + prtHeaderHeight

End Sub

'---------------------------------------------------------
 Public Sub prtCompteAttributX(Msg As String)
'---------------------------------------------------------
Dim K As Integer, K1 As Integer, K2 As Integer
Dim x As String

Set XPrt = Printer
K1 = Val(Mid$(Msg, 1, 6))
K2 = Val(Mid$(Msg, 7, 6))

If Mid$(Msg, 13, 1) = " " Then
    chkNuméroAncien = False
Else
    chkNuméroAncien = True
End If

If Mid$(Msg, 14, 3) = "   " Then
    chkDeviseCV = False
Else
    chkDeviseCV = True
    Call DevCode(Mid$(Msg, 14, 3))
    CV = XDevise
End If

mDevCode = 0: totalDev = 0: totalCV = 0

frmElpPrt.Show vbModeless

recCompteInit recCompte

prtOrientation = vbPRORLandscape
prtTitleText = "Interrogation de Comptes"
prtPgmName = "prtCompteAttribut"
prtTitleUsr = usrName

prtLineNb = 3
prtlineHeight = 250
prtHeaderHeight = 300

frmElpPrt.prtStdinit
prtCompteAttributForm

For K = K1 To K2
recCompte = arrCompte(K)
prtCompteAttributLine

DoEvents: If prtKillDoc Then frmElpPrt.Hide: Exit Sub

Next K

frmElpPrt.prtLineY

K = 0
XPrt.CurrentY = XPrt.CurrentY + 50
XPrt.FontBold = True

If mDevCode > 0 Then
    K = 1
    x = Format$(totalDev, "#### ### ### ### ##0.00")
    If totalDev >= 0 Then
        XPrt.CurrentX = prtMinX + 15200 - XPrt.TextWidth(x)
    Else
        XPrt.CurrentX = prtMinX + 12900 - XPrt.TextWidth(x)
    End If

    XPrt.Print x & " ";
    XPrt.FontBold = True
    XPrt.Print XDevise.DevX;
End If

If chkDeviseCV Then
    K = 1
    x = Format$(totalCV, "#### ### ### ### ##0.00")
    XPrt.CurrentX = prtMinX + 10500 - XPrt.TextWidth(x)
    XPrt.Print x & " "; CV.DevX;
End If

prtCurrentY = XPrt.CurrentY + prtlineHeight + 50
If K > 0 Then: frmElpPrt.prtLineY

frmElpPrt.prtEndDoc
frmElpPrt.Hide

End Sub




'---------------------------------------------------------
Public Sub prtCompteAttributLine()
'---------------------------------------------------------
Dim x As String, K As Integer
Dim Situation As String

If prtCurrentY + prtParagraphHeight > prtMaxY Then
    frmElpPrt.prtNewPage
    prtCompteAttributForm
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
    XPrt.CurrentX = prtMinX + 5500
    XPrt.Print DicLib(13, recCompte.BiaTyp);
Else
    XPrt.Print Format$(recCompte.Numéro, "@@@@ @@@ @.@");
End If

XPrt.CurrentX = prtMinX + 1850
XPrt.Print recCompte.Intitulé;

XPrt.FontBold = True

       
x = Format$(recCompte.SoldeVeille, "#### ### ### ### ##0.00")
If recCompte.SoldeVeille >= 0 Then
    XPrt.CurrentX = prtMinX + 15200 - XPrt.TextWidth(x)
Else
    XPrt.CurrentX = prtMinX + 12900 - XPrt.TextWidth(x)
End If


K = Val(recCompte.Devise)
Call DevX(K)
If K <> mDevCode Then
    mDevCode = IIf(mDevCode = 0, K, -1)
End If
totalDev = totalDev + recCompte.SoldeVeille

XPrt.Print x & " ";
XPrt.FontBold = False
XPrt.Print XDevise.DevX & " " & recCompte.MvtceJour;

If chkDeviseCV Then
    XPrt.FontBold = True
    Mt = recCompte.SoldeVeille * XDevise.Cours / CV.Cours
    x = Format$(Mt, "#### ### ### ### ##0.00")
    XPrt.CurrentX = prtMinX + 10500 - XPrt.TextWidth(x)
    XPrt.Print x & " ";
    XPrt.FontBold = False
    XPrt.Print CV.DevX;
    totalCV = totalCV + Mt
End If
    
XPrt.FontBold = False


'------------------------------------------ligne 2--------------

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
       
If recCompte.Situation <> " " Then
    Select Case recCompte.Situation
        Case "A"
           x = "Annulé"
         Case "B"
           x = "Bloqué"
         Case Else
         x = "? " & recCompte.Situation
    End Select
    
    XPrt.FontBold = True
    XPrt.CurrentX = prtMinX + 400
    XPrt.Print x;
    XPrt.FontBold = False
End If

XPrt.CurrentX = prtMinX + 1850

XPrt.Print recCompte.Intitulé2;
If recCompte.DécouvertMontant > 0 Then
    XPrt.CurrentX = prtMinX + 5500:
            x = "Découvert autorisé : " & Format$(recCompte.DécouvertMontant, "### ### ### ###") _
          & "  juqu'au : " & dateImp(recCompte.DécouvertAmj)
        XPrt.Print x;
End If
'------------------------------------------ligne 3--------------

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight

If chkNuméroAncien Then
    XPrt.CurrentX = prtMinX + 5500
    XPrt.Print Format$(Mid$(recCompte.NuméroAncien, 1, 8), "@@@@@@.@@");
End If

prtCurrentY = prtCurrentY + prtParagraphHeight
End Sub






