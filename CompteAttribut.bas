Attribute VB_Name = "prtCompteAttribut"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------

Private recCompte As typeCompte
Dim I As Integer
Dim chkNuméroAncien As Boolean
Dim chkDeviseCV As Boolean
Dim mDevCode As Integer
Dim NbImprimé As Integer

Private recCptInfo As typeCptInfo
Private CV       As typeDevise
Private Mt As Currency
Private totalCV As Currency, totalDev As Currency
Public mNuméro As String * 11

'---------------------------------------------------------
Public Sub prtCompteAttributForm()
'---------------------------------------------------------
Dim X As String
NbImprimé = 0
XPrt.FontSize = 8
XPrt.FontBold = True

XPrt.FillStyle = 0
XPrt.DrawWidth = 3
XPrt.ForeColor = RGB(0, 0, 0)
XPrt.FillStyle = 1

XPrt.Line (prtMinX, prtMinY + prtHeaderHeight)-(prtMaxX, prtMinY + prtHeaderHeight)

XPrt.DrawWidth = 2


XPrt.Line (prtMinX, prtMinY)-(prtMinX, prtMaxY)
XPrt.Line (prtMaxX, prtMinY)-(prtMaxX, prtMaxY)


'---------------------------------------------------------

X = "N°de Compte"
XPrt.CurrentY = prtMinY + (prtHeaderHeight - XPrt.TextHeight(X)) / 2
XPrt.CurrentX = prtMinX + 300
XPrt.Print X;

X = "Intitulé"
XPrt.CurrentX = 3000
XPrt.Print X;

X = "Rés BDF"
XPrt.CurrentX = 5100
XPrt.Print X;

'x = "Act."
'XPrt.CurrentX = 6000
'XPrt.Print x;

X = "Situation"
XPrt.CurrentX = 6200
XPrt.Print X;

X = "Cpt Général"
XPrt.CurrentX = 7075
XPrt.Print X;

X = "Sens"
XPrt.CurrentX = 8125
XPrt.Print X;

X = "Cond."
XPrt.CurrentX = 8800
XPrt.Print X;

X = "Ech."
XPrt.CurrentX = 9700
XPrt.Print X;

X = "Gest."
XPrt.CurrentX = 10400
XPrt.Print X;

X = "Serv.Resp."
XPrt.CurrentX = 11000
XPrt.Print X;

X = "Services Autorisés"
XPrt.CurrentX = 13100
XPrt.Print X;

prtCurrentY = prtMinY + prtHeaderHeight

End Sub

'---------------------------------------------------------
 Public Sub prtCompteAttributX(Msg As String)
'---------------------------------------------------------
Dim K As Integer, K1 As Integer, K2 As Integer
Dim X As String

K1 = Val(Mid$(Msg, 1, 6))
K2 = Val(Mid$(Msg, 7, 6))



For K = K1 To K2
If mNuméro <> arrCptInfo(K).Numéro Then
    recCptInfo = arrCptInfo(K)
    prtCompteAttributLine
    mNuméro = recCptInfo.Numéro
End If

DoEvents: If prtKillDoc Then frmElpPrt.Hide: Exit Sub

Next K
End Sub




'---------------------------------------------------------
Public Sub prtCompteAttributLine()
'---------------------------------------------------------
Dim X As String, K As Integer
Dim Situation As String

If prtCurrentY + prtParagraphHeight > prtMaxY Then
   prtCompteAttributTrait
   frmElpPrt.prtNewPage
   prtCompteAttributForm
End If
XPrt.FontSize = 8
XPrt.CurrentY = prtCurrentY + prtlineHeight - XPrt.TextHeight("test")
NbImprimé = NbImprimé + 1

If NbImprimé = 1 Then
   Call frmElpPrt.prtTrame(prtMinX + 20, XPrt.CurrentY - 50, prtMaxX - 20, XPrt.CurrentY + prtlineHeight * 3 - 50, " ")
   NbImprimé = -5
End If


XPrt.FontBold = False
If recCptInfo.TypeGA = "R" Then XPrt.Line (prtMinX, XPrt.CurrentY - 20)-(prtMaxX, XPrt.CurrentY - 20)
XPrt.CurrentY = XPrt.CurrentY + 20

XPrt.CurrentX = prtMinX + 50
XPrt.Print Format$(recCptInfo.Devise, "000") & ".";

XPrt.Print Compte_Imp(recCptInfo.Numéro);
XPrt.FontSize = 7

XPrt.CurrentX = prtMinX + 1750
XPrt.Print recCptInfo.Intitulé;

XPrt.FontSize = 8

XPrt.CurrentX = 5500
XPrt.Print recCptInfo.Résident;

XPrt.CurrentX = 6100
    If recCptInfo.Actionnaire <> "0" Then
    XPrt.Print recCptInfo.Actionnaire;

End If

XPrt.CurrentX = 6400
Select Case recCptInfo.Situation
    Case "B": XPrt.Print "Bloqué";
    Case "A": XPrt.Print "Annulé";

End Select

XPrt.CurrentX = 7200
XPrt.Print Compte_Imp(recCptInfo.CompteGénéral);

XPrt.CurrentX = 8300
XPrt.Print recCptInfo.Sens;

XPrt.CurrentX = 8900
XPrt.Print recCptInfo.Conditions;

XPrt.CurrentX = 9700
XPrt.Print DicLib(7, recCptInfo.Echelle);

XPrt.CurrentX = 10500
XPrt.Print recCptInfo.Gestionnaire;

XPrt.CurrentX = 11300
XPrt.Print recCptInfo.ServiceResponsable;


If Trim(recCptInfo.ServiceAutorisé1) <> "000" Then
    XPrt.CurrentX = 12100
    XPrt.Print recCptInfo.ServiceAutorisé1;
End If

If Trim(recCptInfo.ServiceAutorisé2) <> "000" Then
    XPrt.CurrentX = 12900
    XPrt.Print recCptInfo.ServiceAutorisé2;
End If


If Trim(recCptInfo.ServiceAutorisé3) <> "000" Then
    XPrt.CurrentX = 13700
    XPrt.Print recCptInfo.ServiceAutorisé3;
End If

If Trim(recCptInfo.ServiceAutorisé4) <> "000" Then
    XPrt.CurrentX = 14400
    XPrt.Print recCptInfo.ServiceAutorisé4;
End If

If Trim(recCptInfo.ServiceAutorisé5) <> "000" Then
    XPrt.CurrentX = 15300
    XPrt.Print recCptInfo.ServiceAutorisé5;

End If


prtCurrentY = prtCurrentY + prtlineHeight
End Sub






Public Sub prtCompteAttributTrait()

XPrt.Line (prtMinX + 1600, prtMinY)-(prtMinX + 1600, prtMaxY)
XPrt.Line (prtMinX + 8400, prtMinY)-(prtMinX + 8400, prtMaxY)
XPrt.Line (prtMinX + 11700, prtMinY)-(prtMinX + 11700, prtMaxY)

End Sub

Public Sub prtCompteAttribut_Open()
Set XPrt = Printer
frmElpPrt.Show vbModeless


prtOrientation = vbPRORLandscape
prtTitleText = "Attributs des Comptes"
prtPgmName = "prtCompteAttribut"
prtTitleUsr = usrName

prtLineNb = 1
prtlineHeight = 300
prtHeaderHeight = 300

frmElpPrt.prtStdinit
prtCompteAttributForm


End Sub

Public Sub prtCptGenLst_Open()
Set XPrt = Printer
frmElpPrt.Show vbModeless


prtOrientation = vbPRORPortrait
prtTitleText = "Liste des Comptes généraux"
prtPgmName = "prtCptGenLst"
prtTitleUsr = usrName

prtLineNb = 1
prtlineHeight = 300
prtHeaderHeight = 300

frmElpPrt.prtStdinit

XPrt.CurrentY = prtMinY + 50
 NbImprimé = 0
End Sub

Public Sub prtCptGenLst_Line()
Dim X As String, K As Integer
Dim Situation As String

If XPrt.CurrentY + prtlineHeight > prtMaxY Then
   frmElpPrt.prtNewPage
   NbImprimé = 0
End If
XPrt.FontSize = 8
NbImprimé = NbImprimé + 1

If NbImprimé = 1 Then
   Call frmElpPrt.prtTrame(prtMinX + 20, XPrt.CurrentY - 50, prtMaxX - 20, XPrt.CurrentY + prtlineHeight * 3 - 50, " ")
   NbImprimé = -5
End If


XPrt.FontBold = False
XPrt.CurrentX = prtMinX
XPrt.Print arrCptInfo(arrCptInfoIndex).Numéro;
XPrt.CurrentX = prtMinX + 1500
XPrt.Print arrCptInfo(arrCptInfoIndex).Intitulé;


XPrt.CurrentY = XPrt.CurrentY + prtlineHeight

End Sub

