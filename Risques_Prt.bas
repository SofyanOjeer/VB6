Attribute VB_Name = "prtRisques"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim Colonne As Integer
Dim Height8_6 As Integer
Dim NbLg As Integer

Public biaRisques   As typeRisques
Public bdfRisques As typeRisques

'---------------------------------------------------------
Public Sub prtRisquesForm()
'---------------------------------------------------------
Dim X As String
XPrt.FontSize = 8
XPrt.FontBold = False

XPrt.FillStyle = 0
XPrt.DrawWidth = 3
XPrt.ForeColor = RGB(0, 0, 0)
XPrt.FillStyle = 1

XPrt.Line (prtMinX, prtMinY + 1100 + prtHeaderHeight)-(prtMaxX, prtMinY + 1100 + prtHeaderHeight)
XPrt.DrawWidth = 3
Call frmElpPrt.prtTrame(prtMinX, prtMinY + 1100, prtMaxX, prtMinY + 1100 + prtHeaderHeight, "B")

'-----------------------------------------------------
XPrt.CurrentY = prtMinY + 1100 + (prtHeaderHeight - XPrt.TextHeight("X")) / 2
XPrt.FontBold = True
XPrt.FontSize = 10
frmElpPrt.prtCentré 5600, "REPERTOIRE TELEPHONIQUE"
XPrt.FontBold = False

XPrt.DrawWidth = 1
XPrt.Line (5600, prtMinY + 1400)-(5600, prtMaxY - 800)  '

'---------------------------------------------------------
XPrt.CurrentY = prtMinY + 1100 + prtHeaderHeight + prtlineHeight * 0.5 - XPrt.TextHeight(X)

Colonne = 0: NbLg = 0


End Sub


'---------------------------------------------------------
Public Sub prtRisquesFormDétail()
'---------------------------------------------------------
Dim X As String
XPrt.FontSize = 8
XPrt.FontBold = False

XPrt.FillStyle = 0
XPrt.DrawWidth = 3
XPrt.ForeColor = RGB(0, 0, 0)
XPrt.FillStyle = 1
XPrt.DrawWidth = 1


XPrt.DrawWidth = 3
Call frmElpPrt.prtTrame(prtMinX, prtMinY + 1100, prtMaxX, prtMinY + 1100 + prtHeaderHeight, "B")

'-----------------------------------------------------

XPrt.CurrentY = prtMinY + 1100 + (prtHeaderHeight - XPrt.TextHeight(X) + 100) / 2
XPrt.CurrentX = prtMinX + 100
XPrt.Print X;
XPrt.CurrentX = 1200: XPrt.Print "Intitulé";
XPrt.CurrentX = 3000: XPrt.Print "Téléphone";
XPrt.CurrentX = 4000: XPrt.Print "Autres Postes";
XPrt.CurrentX = 6700: XPrt.Print "S/N";
XPrt.CurrentX = 8000: XPrt.Print "AdresseIP";
XPrt.CurrentX = 9300: XPrt.Print "Service";
XPrt.CurrentX = 10400: XPrt.Print "Bureau";


XPrt.DrawWidth = 1
'---------------------------------------------------------
XPrt.FontSize = 8
XPrt.CurrentY = prtMinY + 1100 + prtHeaderHeight + prtlineHeight * 0.5 - XPrt.TextHeight(X)

Colonne = 0: NbLg = 0


End Sub



'---------------------------------------------------------
 Public Sub prtRisquesX(Msg As String)
'---------------------------------------------------------
Dim K As Integer, K1 As Integer, K2 As Integer
Dim X As String
NbLg = 0


Set XPrt = Printer
K1 = Val(Mid$(Msg, 1, 6))
K2 = Val(Mid$(Msg, 7, 6))
frmElpPrt.Show vbModeless

prtFormType = "+"
prtOrientation = vbPRORPortrait
prtTitleText = ""
prtPgmName = "prtRisques"
prtTitleUsr = usrName

prtLineNb = 1
prtlineHeight = 250
prtHeaderHeight = 300
prtFormType = "SOC"
Colonne = 0
XPrt.FontSize = 8
Height8_6 = XPrt.TextHeight("X")
XPrt.FontSize = 6
Height8_6 = Height8_6 - XPrt.TextHeight("X")

prtRisquesForm
For K = 1 To 12
    biaRisques = arrRisques(K)
    bdfRisques = totalRisques(K)
   
Next K
frmElpPrt.prtEndDoc
frmElpPrt.Hide

End Sub
'---------------------------------------------------------
Public Sub prtRisquesLine()
'---------------------------------------------------------
Dim X As String, K As Integer
Dim Situation As String
Dim Civilité As String
Civilité = ""
'-------------------------------------------
If XPrt.CurrentY + prtlineHeight * 5 > prtMaxY Then
    NbLg = 0
    If Colonne = 0 Then
        Colonne = 5600
       XPrt.CurrentY = prtMinY + 1100 + prtHeaderHeight + prtlineHeight * 0.5 - XPrt.TextHeight("x")
    Else
      
        frmElpPrt.prtNewPage
        prtRisquesForm
    End If
End If
'------------------------------------------------
XPrt.FontBold = False
'------------------------------------------1-----------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.FontSize = 8
If NbLg = 3 Then
    Call frmElpPrt.prtTrame(150 + Colonne, XPrt.CurrentY - 20, 5450 + Colonne, XPrt.CurrentY + prtlineHeight * 3 + 20, " ")
    XPrt.CurrentY = XPrt.CurrentY + 20
Else

    If NbLg = 6 Then
       NbLg = 0
    End If
End If

NbLg = NbLg + 1

Select Case recRisques.Civilité
       Case "1": Civilité = "Mr "
       Case "2": Civilité = "Mme "
       Case "3": Civilité = "Mlle "
       Case "4": Civilité = ""
 End Select
XPrt.CurrentX = 300 + Colonne
XPrt.FontBold = False
XPrt.Print Civilité;
XPrt.CurrentX = 800 + Colonne
XPrt.FontBold = True
XPrt.Print Trim(recRisques.Nom) & " ";
XPrt.FontSize = 6
XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + Height8_6
XPrt.Print Trim(recRisques.Prénoms);
XPrt.CurrentY = XPrt.CurrentY - Height8_6


XPrt.FontSize = 8
XPrt.CurrentX = 3200 + Colonne
XPrt.FontBold = True
XPrt.Print recRisques.Tél1;
XPrt.CurrentX = 4100 + Colonne
XPrt.FontBold = False
XPrt.Print recRisques.Tél2;
If Trim(recRisques.Tél3) <> "" Then
    XPrt.CurrentX = 4450 + Colonne
    XPrt.Print "-";
    XPrt.CurrentX = 4600 + Colonne
    XPrt.Print recRisques.Tél3;
End If







End Sub
'---------------------------------------------------------
Public Sub prtRisquesLineDétail()
'---------------------------------------------------------
Dim X As String, K As Integer
Dim Situation As String
Dim Civilité As String
Civilité = ""
'-------------------------------------------
'------------------------------------------------

If XPrt.CurrentY + prtlineHeight * 3 > prtMaxY Then
    prtRisquesTraitDétail
   frmElpPrt.prtNewPage
   prtRisquesFormDétail
End If

XPrt.FontBold = False
XPrt.FontSize = 8
If NbLg = 3 Then
    Call frmElpPrt.prtTrame(150 + Colonne, XPrt.CurrentY - 20, 11000 + Colonne, XPrt.CurrentY + prtlineHeight * 3 + 20, " ")
    XPrt.CurrentY = XPrt.CurrentY + 20
Else

    If NbLg = 6 Then
       NbLg = 0
    End If
End If

NbLg = NbLg + 1


XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
Select Case recRisques.Civilité
       Case "1": Civilité = "Mr "
       Case "2": Civilité = "Mme "
       Case "3": Civilité = "Mlle "
       Case "4": Civilité = ""
 End Select

XPrt.CurrentX = 300 + Colonne
XPrt.FontBold = False
XPrt.Print Civilité;
XPrt.CurrentX = 800 + Colonne
XPrt.FontBold = True
XPrt.Print Trim(recRisques.Nom) & " ";
XPrt.FontSize = 6
XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + Height8_6
XPrt.Print Trim(recRisques.Prénoms);
XPrt.CurrentY = XPrt.CurrentY - Height8_6

XPrt.FontSize = 8
XPrt.CurrentX = 3200 + Colonne
XPrt.FontBold = True
XPrt.Print recRisques.Tél1;
XPrt.CurrentX = 4100 + Colonne
XPrt.FontBold = False
XPrt.Print recRisques.Tél2;
If Trim(recRisques.Tél3) <> "" Then
    XPrt.CurrentX = 4450 + Colonne
    XPrt.Print "-";
    XPrt.CurrentX = 4600 + Colonne
    XPrt.Print recRisques.Tél3;
End If

XPrt.CurrentX = 6600
XPrt.Print recRisques.MicroSN;
XPrt.CurrentX = 8000
XPrt.Print recRisques.MicroIP;
XPrt.CurrentX = 9500
XPrt.Print recRisques.Service;
XPrt.CurrentX = 10600
XPrt.Print recRisques.Bureau;


End Sub


Public Sub prtRisquesTraitDétail()

XPrt.Line (6300, prtMinY + 1400)-(6300, prtMaxY - 350)  '
XPrt.Line (9000, prtMinY + 1400)-(9000, prtMaxY - 350)  '
XPrt.Line (9900, prtMinY + 1400)-(9900, prtMaxY - 350)  '

End Sub
