Attribute VB_Name = "prtOpTrf"
Option Explicit
'---------------------------------------------------------
Private recBic As typeBic
Private recBicIban As typeBicIban
Dim I As Integer, mCurrenty As Integer, maxCurrentY As Integer
Dim blnTest As Boolean
Dim strValidation As String
'---------------------------------------------------------
Public Sub prtOpTrfX(Msg As String, Optional Dest As Variant)
'---------------------------------------------------------
On Error GoTo prtError

Set XPrt = Printer
If prtShow Then frmElpPrt.Show vbModeless
prtTitleUsr = IIf(Not IsMissing(Dest), Dest, usrName)
prtOrientation = vbPRORLandscape
prtPgmName = "prtOpTrfP"

prtLineNb = 3
prtlineHeight = 250
prtHeaderHeight = 300
prtTitleText = "Opération de Transfert :  " & XopCpt(1).Référence & "  du  " & dateImp(DSys)
strValidation = ""

frmElpPrt.prtStdInit 'optFontSize:=9

prtOpTrf_Form Msg
maxCurrentY = XPrt.CurrentY
mCurrenty = maxCurrentY + prtlineHeight
XPrt.CurrentY = mCurrenty: Xopswift_MT100S: prtOpSwift prtMinX
If XopCpt(2).optSwift > 0 Then
    XPrt.CurrentY = mCurrenty: Xopswift_MT100R: prtOpSwift prtMinX + 8100
End If
XPrt.CurrentY = maxCurrentY + prtlineHeight

prtOpCpt_Form
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
maxCurrentY = XPrt.CurrentY
prtOpCpt
XPrt.CurrentY = maxCurrentY + prtlineHeight: mCurrenty = XPrt.CurrentY

prtOpEch_Form
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
prtOpEch
If XopCpt(2).optSwift = 2 Then
    XPrt.CurrentY = mCurrenty + prtlineHeight: XopSwift_MT202: prtOpSwift prtMinX + 8100
End If
If Trim(Msg) = constDemandeDeValidation Then
    Call frmElpPrt.prtTrame(9000, prtMaxY - prtlineHeight * 10, 14000, prtMaxY - prtlineHeight * 1, " ", "230")
    XPrt.CurrentY = prtMaxY - prtlineHeight * 9
    XPrt.FontSize = 10
    XPrt.FontUnderline = True
    frmElpPrt.prtCentré 11500, "VALIDATION"
    XPrt.FontUnderline = False
    XPrt.CurrentY = prtMaxY - prtlineHeight * 7
    frmElpPrt.prtCentré 11500, strValidation
End If



DoEvents: If prtKillDoc Then frmElpPrt.Hide: Exit Sub

frmElpPrt.prtEndDoc
If prtShow Then frmElpPrt.Hide
Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide
End Sub
'---------------------------------------------------------
Public Sub prtOpTrf_Form(Msg As String)
'---------------------------------------------------------
Dim X As String

XPrt.FontSize = 7
XPrt.FontBold = True
XPrt.DrawWidth = 2
XPrt.Line (prtMinX + 8000, prtMinY)-(prtMinX + 8000, prtMaxY)
XPrt.Line (prtMinX, prtMinY)-(prtMaxX, prtMinY)
Call frmElpPrt.prtTrame(prtMinX, prtMinY, prtMaxX, prtMinY + prtHeaderHeight, " ")
'---------------------------------------------------------

X = Trim(XopBicSender.BicId) & " : " & Trim(XopBicSender.Intitulé)
XPrt.CurrentY = prtMinY + (prtHeaderHeight - XPrt.TextHeight(X)) / 2
XPrt.CurrentX = prtMinX + (8000 - prtMinX - XPrt.TextWidth(X)) / 2
XPrt.Print X;
XPrt.CurrentX = prtMinX + (prtMaxX - prtMinX - XPrt.TextWidth(Trim(Msg))) / 2
XPrt.Print Trim(Msg);
X = Trim(XopBicReceiver.BicId) & " : " & Trim(XopBicReceiver.Intitulé)
XPrt.CurrentX = 8000 + (prtMaxX - 8000 - XPrt.TextWidth(X)) / 2
XPrt.Print X;
prtCurrentY = prtMinY + prtHeaderHeight
XPrt.FontBold = False

End Sub

'---------------------------------------------------------
Public Sub prtOpCpt_Form()
'---------------------------------------------------------
Dim X As String

XPrt.FontSize = 8
XPrt.FontBold = True

XPrt.DrawWidth = 2
mCurrenty = XPrt.CurrentY + prtlineHeight
Call frmElpPrt.prtTrame(prtMinX, mCurrenty, prtMaxX, mCurrenty + prtHeaderHeight, " ")
XPrt.Line (prtMinX, mCurrenty)-(prtMaxX, mCurrenty)

XPrt.CurrentY = mCurrenty + (prtHeaderHeight - XPrt.TextHeight("D")) / 2
X = "Débit : " & DevX(XopCpt(1).Devise) & num_Display(XopCpt(1).Brut + XopCpt(1).ComMontantDevise, 15, 2, Lx, X, "0")
XPrt.CurrentX = 7550 + prtMinX - XPrt.TextWidth(X)
XPrt.Print X;
X = "Crédit : " & DevX(XopCpt(2).Devise) & num_Display(XopCpt(2).Brut - XopCpt(2).ComMontantDevise, 15, 2, Lx, X, "0")
XPrt.CurrentX = 15650 + prtMinX - XPrt.TextWidth(X)
XPrt.Print X;
End Sub
'---------------------------------------------------------
Public Sub prtOpEch_Form()
'---------------------------------------------------------
Dim X As String
XPrt.FontBold = True
XPrt.FontSize = 8
XPrt.DrawWidth = 2
mCurrenty = XPrt.CurrentY + prtlineHeight
Call frmElpPrt.prtTrame(prtMinX, mCurrenty, prtMaxX, mCurrenty + prtHeaderHeight, " ")
XPrt.Line (prtMinX, mCurrenty)-(prtMaxX, mCurrenty)
X = "Suivi"
XPrt.CurrentY = mCurrenty + (prtHeaderHeight - XPrt.TextHeight(X)) / 2
XPrt.CurrentX = prtMinX + 3500
XPrt.Print X;
'XPrt.CurrentX = prtMinX + 11500

X = ""
If XopCpt(2).optSwift = 2 Then
    X = "MT202 : " & mId$(XopSwift(1).FieldText, 13, 11)
Else
    Select Case XopCpt(2).optVirement
        Case 1: X = "Virement Papier"
        Case 2: X = "Virement Aristote"
        Case 3: X = "Virement CRI"
        Case 4: X = ""
    End Select
End If
XPrt.CurrentX = 8000 + (prtMaxX - 8000 - XPrt.TextWidth(X)) / 2
XPrt.Print X;

XPrt.FontBold = True
End Sub


Public Sub prtOpSwift(IX0 As Integer)
Dim I As Integer, K As Integer, X As String, X2 As String
X2 = ""
XPrt.FontBold = False
For I = 1 To arrOpSwiftNb
    X = mId$(arrOpSwift(I).FieldNature, 1, 2)
    If X > "09" Then
       XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
       XPrt.FontBold = False
        If X <> X2 Then
            K = 0
            X2 = mId$(arrOpSwift(I).FieldNature, 1, 2)
            XPrt.FontSize = 7
            XPrt.CurrentX = IX0
            XPrt.Print lblOpTrf("Swift", X2);
            If X2 = "32" Or X2 = "57" Or X2 = "58" Or X2 = "71" Then XPrt.FontBold = True
            XPrt.FontSize = 8
        End If
        K = K + 1: arrOpSwift(I).FieldLine = K
        XPrt.CurrentX = IX0 + 2300
        XPrt.Print arrOpSwift(I).FieldText;
        If XPrt.CurrentY > maxCurrentY Then maxCurrentY = XPrt.CurrentY
    End If
Next I

End Sub

Public Sub prtOpEch()
Dim I As Integer, K As Integer, X As String, X2 As String
XPrt.FontBold = False
XPrt.FontSize = 7

For I = 1 To XopEchNb
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinX
    XPrt.Print XopEch(I).EchéanceNature;
    XPrt.CurrentX = 1000
    XPrt.Print dateImp(XopEch(I).EchéanceAmj);
    XPrt.CurrentX = 2200
    XPrt.Print timeImp(XopEch(I).EchéanceHms);
    XPrt.CurrentX = 3500
    XPrt.Print " / " & XopEch(I).ActionOpérateur;
    XPrt.CurrentX = 4700
    XPrt.Print XopEch(I).ActionNature;
    XPrt.CurrentX = 5700
    XPrt.Print dateImp(XopEch(I).ActionAmj);
    XPrt.CurrentX = 6900
    XPrt.Print timeImp(XopEch(I).ActionHms);
Next I
End Sub

Public Sub prtOpCpt()
Dim I As Integer, IX0 As Integer, X As String, X2 As String
Dim curMt As Currency
XPrt.FontBold = False
mCurrenty = XPrt.CurrentY
For I = 1 To XopCptNb
    XPrt.CurrentY = mCurrenty
    IX0 = IIf(I = 1, 0, 8100) + prtMinX
    X = ""
    If XopCpt(I).chkCompte <> "0" Then X = X & " Compte bloqué /"
    If XopCpt(I).chkSolde <> "0" Then X = X & " Solde débiteur /"
    If XopCpt(I).chkAmjOpération <> "0" Then X = X & " date opération /"
    If XopCpt(I).chkAmjValeur <> "0" Then X = X & " date valeur /"
    If X <> "" Then
        strValidation = strValidation & X
        X2 = "A vérifier :" & X
        XPrt.FontUnderline = True
        XPrt.FontSize = 8
        XPrt.CurrentX = IX0 + 8000 - XPrt.TextWidth(X2)
        XPrt.Print X2;
        XPrt.FontUnderline = False
    End If
    XPrt.FontBold = False: XPrt.FontSize = 8
    XPrt.CurrentY = mCurrenty + prtlineHeight
    XPrt.CurrentX = IX0
    XPrt.Print XopCpt(I).Devise & ".";
    XPrt.Print Compte_Imp(XopCpt(I).Compte);
    XPrt.CurrentX = IX0 + 2300
    If I = 1 Then
        XPrt.Print XopCompteSender.Intitulé;
        If Trim(XopCompteSender.Intitulé2) <> "" Then
            XPrt.CurrentX = IX0 + 2300
            XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
            XPrt.Print XopCompteSender.Intitulé2;
        End If
    Else
        XPrt.Print XopCompteNostro.Intitulé;
        If Trim(XopCompteNostro.Intitulé2) <> "" Then
            XPrt.CurrentX = IX0 + 2300
            XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
            XPrt.Print XopCompteNostro.Intitulé2;
        End If
    End If
'    If XopCpt(I).Sens = "C" Then
'        curMt = XopCpt(I).Brut - XopCpt(I).ComMontantDevise
'    Else
'        curMt = XopCpt(I).Brut + XopCpt(I).ComMontantDevise
'    End If
    curMt = XopCpt(I).Brut
    X = num_Display(curMt, 15, XopDevise(I).maxD, Lx, X, "0")
    XPrt.CurrentX = IX0 + 7550 - XPrt.TextWidth(X)
    XPrt.Print X;
    XPrt.CurrentX = IX0 + 7600
    XPrt.Print XopCpt(I).Sens;
    
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.FontSize = 7
    XPrt.CurrentX = IX0
    XPrt.Print "Date valeur: ";
    XPrt.CurrentX = IX0 + 800
    XPrt.Print dateImp(XopCpt(I).AmjValeur);
    XPrt.CurrentX = IX0 + 2300
    XPrt.Print XopCpt(I).Libellé;

    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    If XopCpt(I).AmjOpération <> XopCpt(I).AmjValeur Then
        XPrt.CurrentX = IX0
        XPrt.Print "Date opération: ";
        XPrt.CurrentX = IX0 + 1000
        XPrt.Print dateImp(XopCpt(I).AmjOpération);
    End If
    XPrt.CurrentX = IX0 + 2300:    XPrt.Print "Commission:" & " (" & XopCpt(I).ComPays & ") ";
    XPrt.Print lblOpTrf("optComImputation", XopCpt(I).optComImputation) & " ";
    If Val(XopCpt(I).ComMontantFRF) <> 0 Then
        XPrt.Print lblOpTrf("optComShare", XopCpt(I).optComShare) & " barème ";
        XPrt.Print lblOpTrf("optComBarème", XopCpt(I).optComBarème) & " ";
        If XopCpt(I).Devise <> "001" Then XPrt.Print num_Display(XopCpt(I).CoursDevise, 2, 7, Lx, X, "0");
        XPrt.FontSize = 8
        X = num_Display(XopCpt(I).ComMontantDevise, 15, XopDevise(I).maxD, Lx, X, "0")
        XPrt.CurrentX = IX0 + 7550 - XPrt.TextWidth(X)
        XPrt.Print X;
        XPrt.CurrentX = IX0 + 7600
        XPrt.Print "D";
        XPrt.FontSize = 7
        If XopCpt(I).chkCom1 <> "0" Then
            XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
            XPrt.CurrentX = IX0 + 2300
            XPrt.Print lblOpTrf("chkCom", "1");
            X = num_Display(XopCpt(I).ComMontant1, 12, 2, Lx, X, "0")
            XPrt.CurrentX = IX0 + 5000 - XPrt.TextWidth(X)
            XPrt.Print X;
        End If
        
        If XopCpt(I).chkCom2 <> "0" Then
            XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
            XPrt.CurrentX = IX0 + 2300
            XPrt.Print lblOpTrf("chkCom", "2");
            X = num_Display(XopCpt(I).ComMontant2, 12, 2, Lx, X, "0")
            XPrt.CurrentX = IX0 + 5000 - XPrt.TextWidth(X)
            XPrt.Print X;
        End If
        
        If XopCpt(I).chkCom3 <> "0" Then
            XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
            XPrt.CurrentX = IX0 + 2300
            XPrt.Print lblOpTrf("chkCom", "3");
            X = num_Display(XopCpt(I).ComMontant3, 12, 2, Lx, X, "0")
            XPrt.CurrentX = IX0 + 5000 - XPrt.TextWidth(X)
            XPrt.Print X;
        End If
        
        If XopCpt(I).chkCom4 <> "0" Then
            XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
            XPrt.CurrentX = IX0 + 2300
            XPrt.Print lblOpTrf("chkCom", "4");
            X = num_Display(XopCpt(I).ComMontant4, 12, 2, Lx, X, "0")
            XPrt.CurrentX = IX0 + 5000 - XPrt.TextWidth(X)
            XPrt.Print X;
        End If
        
        If XopCpt(I).chkCom5 <> "0" Then
            XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
            XPrt.CurrentX = IX0 + 2300
            XPrt.Print lblOpTrf("chkCom", "5");
            X = num_Display(XopCpt(I).ComMontant5, 12, 2, Lx, X, "0")
            XPrt.CurrentX = IX0 + 5000 - XPrt.TextWidth(X)
            XPrt.Print X;
        End If
    End If
    If XPrt.CurrentY > maxCurrentY Then maxCurrentY = XPrt.CurrentY
Next I

End Sub
