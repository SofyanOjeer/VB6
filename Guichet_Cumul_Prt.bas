Attribute VB_Name = "prtGuichet_Cumul"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim Xcompte As typeCompte, xCompteMax As typeCompte

Public arrCumulEspèces() As typeG_CumulEspèces
Public arrCumulEspèces_Nb As Integer

Dim I As Integer, Height8_6 As Integer

Type typeG_CumulEspèces
    Devise As String * 3
    CodeOpération As String * 4
    Nb As Integer
    Montant As Currency
End Type

Dim curDB As Currency, curCR As Currency
Dim xCV As typeCV
Dim mY1 As Integer, mY2 As Integer
'---------------------------------------------------------
Public Sub prtGuichet_CumulForm()
'---------------------------------------------------------
Dim X As String

XPrt.FontSize = 8
XPrt.FontBold = True
XPrt.DrawWidth = 3

Call frmElpPrt.prtTrame(prtMinX, prtMinY, prtMaxX, prtMinY + prtHeaderHeight, "B")
Call frmElpPrt.prtTrame(prtMinX + 8100, prtMinY + prtHeaderHeight + 10, prtMinX + 10500, prtMaxY - 10, " ")

XPrt.DrawWidth = 2

XPrt.Line (prtMinX + 5600, prtMinY)-(prtMinX + 5600, prtMaxY)
XPrt.Line (prtMinX + 8100, prtMinY)-(prtMinX + 8100, prtMaxY)
XPrt.Line (prtMinX + 12700, prtMinY)-(prtMinX + 12700, prtMaxY)

XPrt.Line (prtMinX + 1600, prtMinY)-(prtMinX + 1600, prtMaxY)
'---------------------------------------------------------

X = "N°de compte"
XPrt.CurrentY = prtMinY + (prtHeaderHeight - XPrt.TextHeight(X)) / 2
XPrt.CurrentX = prtMinX + 300
XPrt.Print X;

X = "Intitulé"
XPrt.CurrentX = prtMinX + 1750
XPrt.Print X;

XPrt.CurrentX = 4700
XPrt.Print "Dernier mvt le";

XPrt.CurrentX = 6800
XPrt.Print "Solde précédent";

X = "Versements"
XPrt.CurrentX = 9200
XPrt.Print X;

X = "Retraits"
XPrt.CurrentX = 12000
XPrt.Print X;

XPrt.CurrentX = 13800
XPrt.Print "Nouveau solde";

prtCurrentY = prtMinY + prtHeaderHeight

End Sub


Public Sub prtGuichet_CumulLine()
'---------------------------------------------------------
Dim X As String, K As Integer, wsdCurrentX As Integer
Dim Situation As String

If prtCurrentY + prtParagraphHeight > prtMaxY Then
    frmElpPrt.prtNewPage
    prtGuichet_CumulForm
'Else
    'frmElpPrt.prtLineY
End If

XPrt.FontBold = False
'------------------------------------------ligne 1--------------

XPrt.CurrentY = prtCurrentY + prtlineHeight - XPrt.TextHeight("test")
XPrt.FontSize = 8

XPrt.CurrentX = prtMinX + 50
XPrt.Print Format$(Xcompte.Devise, "000") & ".";

    XPrt.Print Compte_Imp(Xcompte.Numéro);

XPrt.FontSize = 6
XPrt.CurrentY = XPrt.CurrentY + Height8_6
XPrt.CurrentX = prtMinX + 1750
XPrt.Print Xcompte.Intitulé;

Select Case Xcompte.Situation
    Case " ": Situation = ""
    Case "A": Situation = " **Annulé**"
    Case "B": Situation = " **Bloqué**"
    Case Else: Situation = " ?? " & Xcompte.Situation
End Select

XPrt.Print Situation;

XPrt.FontSize = 8
XPrt.CurrentY = XPrt.CurrentY - Height8_6
       
X = Format$(Xcompte.SoldeInstantané, "#### ### ### ### ##0.00")
XPrt.CurrentX = 8100 - XPrt.TextWidth(X)
XPrt.Print X;

XPrt.FontSize = 6
XPrt.CurrentY = XPrt.CurrentY + Height8_6
X = dateImp(Xcompte.MvtAmj)
XPrt.CurrentX = 5600 - XPrt.TextWidth(X)
XPrt.Print X;

    
arrCumulEspèces_Print

XPrt.FontSize = 8
XPrt.CurrentY = XPrt.CurrentY - Height8_6

X = Format$(curDB, "#### ### ### ### ##0.00")
XPrt.CurrentX = 10500 - XPrt.TextWidth(X)
XPrt.Print X;
              
X = Format$(curCR, "#### ### ### ### ##0.00")
XPrt.CurrentX = 12800 - XPrt.TextWidth(X)
XPrt.Print X;

XPrt.FontBold = True
X = Format$(Xcompte.SoldeInstantané + curDB + curCR, "#### ### ### ### ##0.00")
XPrt.CurrentX = 15100 - XPrt.TextWidth(X)
XPrt.Print X;
XPrt.Print "    " & xCV.DeviseIso;
XPrt.FontBold = False



prtCurrentY = mY2 + prtlineHeight

End Sub


'---------------------------------------------------------
Public Sub prtGuichet_CumulX(Msg As String)
'---------------------------------------------------------
Dim K As Integer, K1 As Integer, K2 As Integer
Dim X As String
On Error GoTo prtError

arrCumulEspèces_Init

Set XPrt = Printer
Height8_6 = frmElpPrt.prtHeightDelta(8, 6)

    prtTitleText = "Guichet : situation espèces "
    prtLineNb = 1

frmElpPrt.Show vbModeless

prtOrientation = vbPRORLandscape
prtPgmName = "prtGuichet_Cumul"
prtTitleUsr = usrName

prtlineHeight = 250
prtHeaderHeight = 300

frmElpPrt.prtStdInit

prtGuichet_CumulForm

For K = 1 To selCompte_Nb
    Xcompte = selCompte(K)
    xCV.DeviseN = Xcompte.Devise: CV_AttributN xCV
    prtGuichet_CumulLine
    
    DoEvents: If prtKillDoc Then frmElpPrt.Hide: Exit Sub

Next K
XPrt.DrawWidth = 5
frmElpPrt.prtLineY

frmElpPrt.prtEndDoc
frmElpPrt.Hide
Call selCompte_Load(Xcompte, xCompteMax, "End")
'$$$$$$$$$$$$$$$$$$$$$$$

Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide
End Sub





Public Sub arrCumulEspèces_Load()
Dim I As Integer, K As Integer
arrCumulEspèces_Nb = Val(mId$(MsgTxt, 35, 3))
ReDim arrCumulEspèces(arrCumulEspèces_Nb)
K = 37

For I = 1 To arrCumulEspèces_Nb
    arrCumulEspèces(I).Devise = mId$(MsgTxt, K + 1, 3)
    arrCumulEspèces(I).CodeOpération = mId$(MsgTxt, K + 4, 4)
    arrCumulEspèces(I).Nb = CInt(Val(mId$(MsgTxt, K + 8, 6)))
    arrCumulEspèces(I).Montant = CCur(Val(mId$(MsgTxt, K + 14, 17)) / 100)
     K = K + 30
Next I

End Sub

Public Sub arrCumulEspèces_Print()
Dim I As Integer, strNb As String, X As String

curDB = 0: curCR = 0
mY1 = XPrt.CurrentY

For I = 1 To arrCumulEspèces_Nb
    If Xcompte.Devise = arrCumulEspèces(I).Devise Then
        XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
        
        strNb = IIf(arrCumulEspèces(I).Nb > 1, "s", "")
        X = Format$(arrCumulEspèces(I).Montant, "#### ### ### ### ##0.00")
        XPrt.CurrentX = 3000 - XPrt.TextWidth(X)
        XPrt.Print X & " " & xCV.DeviseIso & "   ";
  
        XPrt.Print Trim(Format$(arrCumulEspèces(I).Nb, "#### ##0  "));
       
        Select Case arrCumulEspèces(I).CodeOpération
            Case "G001"
                XPrt.Print " Versement" & strNb;
                curDB = curDB - arrCumulEspèces(I).Montant
            Case "G002"
                XPrt.Print " Retrait" & strNb;
                curCR = curCR + arrCumulEspèces(I).Montant
            Case "G005"
                XPrt.Print " Retrait" & strNb & " de devises";
                curCR = curCR + arrCumulEspèces(I).Montant
            Case "G006"
                XPrt.Print " Versement" & strNb & " de devises";
                curDB = curDB - arrCumulEspèces(I).Montant
            Case "G007"
                XPrt.Print " Change" & strNb;
                curCR = curCR + arrCumulEspèces(I).Montant
            Case "X007"
                XPrt.Print " Change" & strNb;
                curDB = curDB - arrCumulEspèces(I).Montant
       End Select


    End If
Next I

mY2 = XPrt.CurrentY
XPrt.CurrentY = mY1

End Sub
Public Sub arrCumulEspèces_Total(xDev As String, xDB As Currency, xCR As Currency)
Dim I As Integer, strNb As String, X As String

xDB = 0: xCR = 0

For I = 1 To arrCumulEspèces_Nb
    If xDev = arrCumulEspèces(I).Devise Then
        Select Case arrCumulEspèces(I).CodeOpération
            Case "G001"
                xDB = xDB - arrCumulEspèces(I).Montant
            Case "G002"
                xCR = xCR + arrCumulEspèces(I).Montant
            Case "G005"
                xCR = xCR + arrCumulEspèces(I).Montant
            Case "G006"
                xDB = xDB - arrCumulEspèces(I).Montant
            Case "G007"
                xCR = xCR + arrCumulEspèces(I).Montant
            Case "X007"
                xDB = xDB - arrCumulEspèces(I).Montant
       End Select


    End If
Next I
End Sub


Public Sub arrCumulEspèces_Init()

arrCumulEspèces_Load

recCompteInit Xcompte
Xcompte.Method = "SnapL5"
Xcompte.Société = SocId$
Xcompte.Agence = SocAgence$
Xcompte.Numéro = paramGuichetBillets_In
Xcompte.Devise = "000"
Xcompte.MvtceJour = " "
Xcompte.chkAnnul = "0"

xCompteMax = Xcompte
xCompteMax.Devise = "999"
If Not IsNull(selCompte_Load(Xcompte, xCompteMax, "Init")) Then Exit Sub

Xcompte.Numéro = paramGuichetBillets_Out
Xcompte.Devise = "000"

xCompteMax = Xcompte
xCompteMax.Devise = "999"
If Not IsNull(selCompte_Load(Xcompte, xCompteMax, "Add")) Then Exit Sub

End Sub
