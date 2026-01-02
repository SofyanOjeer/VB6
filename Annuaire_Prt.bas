Attribute VB_Name = "prtAnnuaire"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim Colonne As Integer
Dim Height8_6 As Integer
Dim NbLg As Integer

'---------------------------------------------------------
Public Sub prtAnnuaireForm()
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
Public Sub prtAnnuaireFormDétail()
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
 Public Sub prtAnnuaireX(Msg As String)
'---------------------------------------------------------
Dim K As Integer, K1 As Integer, K2 As Integer
Dim X As String
NbLg = 0
On Error GoTo prtError



Set XPrt = Printer
Height8_6 = frmElpPrt.prtHeightDelta(8, 6)
K1 = Val(mId$(Msg, 1, 6))
K2 = Val(mId$(Msg, 7, 6))
frmElpPrt.Show vbModeless

prtFormType = "+"
prtOrientation = vbPRORPortrait
prtTitleText = ""
prtPgmName = "prtAnnuaire"
prtTitleUsr = usrName

prtLineNb = 1
prtlineHeight = 250
prtHeaderHeight = 300
prtFormType = "SOC"
Colonne = 0

If mId$(Msg, 13, 1) = "L" Then
    frmElpPrt.prtInit
    prtSoc
    prtAnnuaireForm
    For K = 0 To frmElp.lstAnnuaire.ListCount - 1
        frmElp.lstAnnuaire.ListIndex = K
        arrAnnuaire_Scan frmElp.lstAnnuaire.Text
        If arrAnnuaireIndex > 0 Then
            recAnnuaire = arrAnnuaire(arrAnnuaireIndex)
            prtAnnuaireLine
        End If
        DoEvents: If prtKillDoc Then frmElpPrt.Hide: Exit Sub
    
    Next K
    frmElpPrt.prtEndDoc
Else
    frmElpPrt.prtInit
    prtSoc
    prtAnnuaireFormDétail
    For K = 0 To frmElp.lstAnnuaire.ListCount - 1
        frmElp.lstAnnuaire.ListIndex = K
        arrAnnuaire_Scan frmElp.lstAnnuaire.Text
        If arrAnnuaireIndex > 0 Then
            recAnnuaire = arrAnnuaire(arrAnnuaireIndex)
            prtAnnuaireLineDétail
        End If
        
        DoEvents: If prtKillDoc Then frmElpPrt.Hide: Exit Sub
    
    Next K
    prtAnnuaireTraitDétail
    frmElpPrt.prtEndDoc
End If

frmElpPrt.Hide
Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub
'---------------------------------------------------------
Public Sub prtAnnuaireLine()
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
        prtAnnuaireForm
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

Select Case recAnnuaire.Civilité
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
XPrt.Print Trim(recAnnuaire.Nom) & " ";
XPrt.FontSize = 6
XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + Height8_6
XPrt.Print Trim(recAnnuaire.Prénoms);
XPrt.CurrentY = XPrt.CurrentY - Height8_6


XPrt.FontSize = 8
XPrt.CurrentX = 3200 + Colonne
XPrt.FontBold = True
XPrt.Print recAnnuaire.Tél1;
XPrt.CurrentX = 4100 + Colonne
XPrt.FontBold = False
XPrt.Print recAnnuaire.Tél2;
If Trim(recAnnuaire.Tél3) <> "" Then
    XPrt.CurrentX = 4450 + Colonne
    XPrt.Print "-";
    XPrt.CurrentX = 4600 + Colonne
    XPrt.Print recAnnuaire.Tél3;
End If







End Sub
'---------------------------------------------------------
Public Sub prtAnnuaireLineDétail()
'---------------------------------------------------------
Dim X As String, K As Integer
Dim Situation As String
Dim Civilité As String
Civilité = ""
'-------------------------------------------
'------------------------------------------------

If XPrt.CurrentY + prtlineHeight * 3 > prtMaxY Then
    prtAnnuaireTraitDétail
   frmElpPrt.prtNewPage
   prtAnnuaireFormDétail
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
Select Case recAnnuaire.Civilité
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
XPrt.Print Trim(recAnnuaire.Nom) & " ";
XPrt.FontSize = 6
XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + Height8_6
XPrt.Print Trim(recAnnuaire.Prénoms);
XPrt.CurrentY = XPrt.CurrentY - Height8_6

XPrt.FontSize = 8
XPrt.CurrentX = 3200 + Colonne
XPrt.FontBold = True
XPrt.Print recAnnuaire.Tél1;
XPrt.CurrentX = 4100 + Colonne
XPrt.FontBold = False
XPrt.Print recAnnuaire.Tél2;
If Trim(recAnnuaire.Tél3) <> "" Then
    XPrt.CurrentX = 4450 + Colonne
    XPrt.Print "-";
    XPrt.CurrentX = 4600 + Colonne
    XPrt.Print recAnnuaire.Tél3;
End If

XPrt.CurrentX = 6600
XPrt.Print recAnnuaire.MicroSN;
XPrt.CurrentX = 8000
XPrt.Print recAnnuaire.MicroIP;
XPrt.CurrentX = 9500
XPrt.Print recAnnuaire.Service;
XPrt.CurrentX = 10600
XPrt.Print recAnnuaire.Bureau;


End Sub


Public Sub prtAnnuaireTraitDétail()

XPrt.Line (6300, prtMinY + 1400)-(6300, prtMaxY - 350)  '
XPrt.Line (9000, prtMinY + 1400)-(9000, prtMaxY - 350)  '
XPrt.Line (9900, prtMinY + 1400)-(9900, prtMaxY - 350)  '

End Sub
