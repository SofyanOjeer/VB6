Attribute VB_Name = "prtGuichetList"
Option Explicit
Dim recCptInfo As typeCptInfo
Dim recGuichet As typeGuichet
'---------------------------------------------------------
Public Sub recGuichet_Display(recGuichet As typeGuichet, XpicBox As PictureBox)
'---------------------------------------------------------
Dim X As String, recCompte As typeCompte

XpicBox.Cls
recCompteInit recCompte
recCompte.Société = recGuichet.Société
recCompte.Agence = recGuichet.Agence
recCompte.Devise = recGuichet.Devise
recCompte.Numéro = recGuichet.Compte
recCompte.BiaTyp = "000"
recCompte.BiaNum = "00"

srvCompteFind recCompte

XpicBox.FontBold = True
XpicBox.ForeColor = libUsr.ForeColor
XpicBox.CurrentX = 50
XpicBox.CurrentY = 50
XpicBox.Print Compte_Imp(recGuichet.Compte)

XpicBox.FontBold = False
XpicBox.CurrentX = 2000
XpicBox.CurrentY = 50
XpicBox.Print recCompte.Intitulé;

XpicBox.CurrentY = 300
XpicBox.ForeColor = lblUsr.ForeColor
XpicBox.CurrentX = 50
XpicBox.Print "Valeur : ";
XpicBox.ForeColor = warnUsrColor
XpicBox.Print dateImpS(recGuichet.AmjValeur);
XpicBox.CurrentX = 2000
XpicBox.Print recGuichet.Libellé;

If recGuichet.Sens = "C" Then
    XpicBox.ForeColor = libUsr.ForeColor
Else
    XpicBox.ForeColor = errUsr.ForeColor
End If
X = num_Display(recGuichet.Montant, 15, 2, Lx, X, "0")
XpicBox.CurrentX = 8700 - XpicBox.TextWidth(X)
XpicBox.Print X;

XpicBox.ForeColor = libUsr.ForeColor
XpicBox.CurrentX = 8800
XpicBox.Print DevX(recGuichet.Devise);

XpicBox.CurrentY = 550

XpicBox.ForeColor = lblUsr.ForeColor
XpicBox.CurrentX = 50
XpicBox.Print "Pièce : ";
X = Format$(recGuichet.CptMvtPièce, "####") & "." & Format$(recGuichet.CptMvtLigne, "0000")
XpicBox.CurrentX = 1400 - XpicBox.TextWidth(X)
XpicBox.ForeColor = libUsr.ForeColor
XpicBox.Print X;

If Trim(recGuichet.CoursChange) <> 1 Or Trim(recGuichet.CoursChangeEspèces) <> 1 Then
    XpicBox.ForeColor = lblUsr.ForeColor
    XpicBox.CurrentX = 2000
    XpicBox.Print "Cours :";
    
    XpicBox.ForeColor = libUsr.ForeColor
    X = num_Display(recGuichet.CoursChange, 10, 5, Lx, X, "#") & "  / "
    XpicBox.CurrentX = 3900 - XpicBox.TextWidth(X)
    XpicBox.Print X;

    X = num_Display(recGuichet.CoursChangeEspèces, 10, 5, Lx, X, "#")
    XpicBox.CurrentX = 4600 - XpicBox.TextWidth(X)
    XpicBox.Print X;

End If


End Sub

'---------------------------------------------------------
Public Sub recGuichet_Détail(recGuichet As typeGuichet, XpicBox As PictureBox)
'---------------------------------------------------------
Dim X As String, recCompte As typeCompte, chkX As String
'XpicBox.Height = 3780
XpicBox.Cls
recCompteInit recCompte
recCompte.Société = recGuichet.Société
recCompte.Agence = recGuichet.Agence
recCompte.Devise = recGuichet.Devise
recCompte.Numéro = recGuichet.Compte
recCompte.BiaTyp = "000"
recCompte.BiaNum = "00"

srvCompteFind recCompte

XpicBox.FontBold = False
XpicBox.ForeColor = libUsr.ForeColor
XpicBox.CurrentX = 50
XpicBox.CurrentY = 50
XpicBox.Print Compte_Imp(recGuichet.Compte);


XpicBox.CurrentX = 1500
XpicBox.Print recCompte.Intitulé;
If recGuichet.Sens = "C" Then
        XpicBox.ForeColor = libUsr.ForeColor
Else
        XpicBox.ForeColor = errUsr.ForeColor
End If
X = num_Display(recGuichet.Montant, 15, 2, Lx, X, "0")
XpicBox.CurrentX = 6500 - XpicBox.TextWidth(X)
XpicBox.Print X;
XpicBox.ForeColor = libUsr.ForeColor
XpicBox.CurrentX = 6550
XpicBox.Print DevX(recGuichet.Devise);


XpicBox.CurrentY = XpicBox.CurrentY + 350
XpicBox.ForeColor = lblUsr.ForeColor
XpicBox.CurrentX = 50
XpicBox.Print "Valeur :";
XpicBox.ForeColor = warnUsrColor
XpicBox.CurrentX = 600
XpicBox.Print dateImpS(recGuichet.AmjValeur);
XpicBox.CurrentX = 1500
XpicBox.Print recGuichet.Libellé;
XpicBox.ForeColor = errUsr.ForeColor
chkX = ""
If recGuichet.chkCompte <> "0" Then chkX = "bloqué "
If recGuichet.chkSolde <> "0" Then chkX = chkX & "débiteur "

XpicBox.CurrentX = 7000 - XpicBox.TextWidth(chkX)
XpicBox.Print chkX;
XpicBox.ForeColor = libUsr.ForeColor


XpicBox.ForeColor = libUsr.ForeColor
If Trim(recGuichet.Identité) <> "" Then
    XpicBox.CurrentY = XpicBox.CurrentY + 350
    XpicBox.CurrentX = 1500
    XpicBox.Print recGuichet.Identité;
End If

If Trim(recGuichet.Complément1) <> "" Then
    XpicBox.CurrentX = 1500
    XpicBox.CurrentY = XpicBox.CurrentY + 350
    XpicBox.Print recGuichet.Complément1;
End If

If Trim(recGuichet.Complément2) <> "" Then
    XpicBox.CurrentX = 1500
    XpicBox.CurrentY = XpicBox.CurrentY + 350
    XpicBox.Print recGuichet.Complément2;
End If


If Trim(recGuichet.Complément3) <> "" Then
    XpicBox.CurrentX = 1500
    XpicBox.CurrentY = XpicBox.CurrentY + 350
    XpicBox.Print recGuichet.Complément3;
End If
'------------------------------------------------------------------
XpicBox.CurrentY = XpicBox.CurrentY + 350
XpicBox.ForeColor = lblUsr.ForeColor
XpicBox.Line (0, XpicBox.CurrentY)-(7000, XpicBox.CurrentY)
XpicBox.CurrentY = XpicBox.CurrentY + 100

If recGuichet.CoursChange <> 1 Or recGuichet.CoursChangeEspèces <> 1 Then
    XpicBox.ForeColor = lblUsr.ForeColor
    XpicBox.CurrentX = 50
    XpicBox.Print "Cours :";
    XpicBox.ForeColor = libUsr.ForeColor
    X = num_Display(recGuichet.CoursChange, 10, 5, Lx, X, "#") & "  / "
    XpicBox.CurrentX = 1700 - XpicBox.TextWidth(X)
    XpicBox.Print X;
    X = num_Display(recGuichet.CoursChangeEspèces, 10, 5, Lx, X, "#")
    XpicBox.CurrentX = 2400 - XpicBox.TextWidth(X)
    XpicBox.Print X;
End If

XpicBox.ForeColor = lblUsr.ForeColor
XpicBox.CurrentX = 3900
XpicBox.Print "Espèces :";
XpicBox.ForeColor = libUsr.ForeColor
X = num_Display(recGuichet.MontantEspèces, 15, 2, Lx, X, "0")
XpicBox.CurrentX = 6500 - XpicBox.TextWidth(X)
XpicBox.Print X;
XpicBox.CurrentX = 6550
XpicBox.Print DevX(recGuichet.DeviseEspèces);


XpicBox.CurrentY = XpicBox.CurrentY + 350
XpicBox.ForeColor = lblUsr.ForeColor
XpicBox.Line (0, XpicBox.CurrentY)-(7000, XpicBox.CurrentY)
XpicBox.CurrentY = XpicBox.CurrentY + 100

XpicBox.ForeColor = lblUsr.ForeColor
XpicBox.FontSize = 7
XpicBox.CurrentX = 50
XpicBox.Print "Saisi      : ";
XpicBox.ForeColor = libUsr.ForeColor
XpicBox.Print recGuichet.SaisieUsr;
XpicBox.CurrentX = 2000
XpicBox.Print dateImpS(recGuichet.SaisieAmj);
XpicBox.CurrentX = 2700
XpicBox.Print timeImpHM(recGuichet.SaisieHMS);

XpicBox.ForeColor = lblUsr.ForeColor
XpicBox.CurrentX = 3900
XpicBox.Print "Opération :";
XpicBox.ForeColor = libUsr.ForeColor
XpicBox.CurrentX = 5300
XpicBox.Print recGuichet.CodeOpération;

XpicBox.CurrentX = 50
XpicBox.CurrentY = XpicBox.CurrentY + 350
XpicBox.ForeColor = lblUsr.ForeColor
XpicBox.Print "Validé   : ";
XpicBox.ForeColor = libUsr.ForeColor
XpicBox.Print recGuichet.ValidationUsr;
XpicBox.CurrentX = 2000
XpicBox.Print dateImpS(recGuichet.ValidationAMJ);
XpicBox.CurrentX = 2700
XpicBox.Print timeImpHM(recGuichet.ValidationHMS);

XpicBox.ForeColor = lblUsr.ForeColor
XpicBox.CurrentX = 3900
XpicBox.Print "Séquence:";
XpicBox.ForeColor = libUsr.ForeColor
XpicBox.CurrentX = 5300
XpicBox.Print recGuichet.Référence;


XpicBox.CurrentX = 50
XpicBox.CurrentY = XpicBox.CurrentY + 350
XpicBox.ForeColor = lblUsr.ForeColor
XpicBox.Print "Compta : ";
XpicBox.ForeColor = libUsr.ForeColor
XpicBox.Print recGuichet.ComptaUsr;
XpicBox.CurrentX = 2000
XpicBox.Print dateImpS(recGuichet.ComptaAMJ);
XpicBox.CurrentX = 2700
XpicBox.Print timeImpHM(recGuichet.ComptaHMS);

XpicBox.ForeColor = lblUsr.ForeColor
XpicBox.CurrentX = 3900
XpicBox.Print "N° Pièce   :";
XpicBox.ForeColor = libUsr.ForeColor
X = Format$(recGuichet.CptMvtPièce, "####") & "." & Format$(recGuichet.CptMvtLigne, "0000")
XpicBox.CurrentX = 6000 - XpicBox.TextWidth(X)
XpicBox.Print X;

If recGuichet.CptMvtPièceEspèces <> 0 Then
    XpicBox.ForeColor = libUsr.ForeColor
    X = " / " & Format$(recGuichet.CptMvtPièceEspèces, "####") & "." & Format$(recGuichet.CptMvtLigneEspèces, "0000")
    XpicBox.CurrentX = 6700 - XpicBox.TextWidth(X)
    XpicBox.Print X;
End If

'XpicBox.Height = XpicBox.CurrentY + 350
End Sub



'---------------------------------------------------------
Private Sub prtGuichetListForm()
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'---------------------------------------------------------
Dim X As String

XPrt.FontSize = 8

XPrt.Line (prtMinX, prtMinY)-(prtMaxX, prtMaxY), , B
Call frmElpPrt.prtTrame(prtMinX, prtMinY, prtMaxX, prtMinY + prtHeaderHeight, "B")
XPrt.DrawWidth = 1
XPrt.Line (2600, prtMinY)-(2600, prtMaxY)
XPrt.Line (12000, prtMinY)-(12000, prtMaxY)
XPrt.Line (8500, prtMinY)-(8500, prtMaxY)


XPrt.CurrentY = prtMinY + (prtHeaderHeight - XPrt.TextHeight(X)) / 2
XPrt.CurrentX = 500
Printer.Print "Opération";
XPrt.CurrentX = 2700
Printer.Print "Compte";
XPrt.CurrentX = 3900
Printer.Print "Intitulé";
XPrt.CurrentX = 9400
Printer.Print "Montant";
XPrt.CurrentX = 11500
Printer.Print "cours";
XPrt.CurrentX = 14000
Printer.Print "Espèces";
XPrt.CurrentX = 15030
Printer.Print "cours";

XPrt.CurrentY = prtMinY + prtHeaderHeight + 50

End Sub

'---------------------------------------------------------
Private Sub prtGuichetListLine()
'---------------------------------------------------------
 Dim X As String, chkX As String

If XPrt.CurrentY + prtlineHeight > prtMaxY Then
    frmElpPrt.prtNewPage
    prtGuichetListForm
'Else
 '   frmElpPrt.prtLineY
End If

XPrt.FontBold = False


'------------------------------------ligne 1
XPrt.FontSize = 8

XPrt.CurrentX = 300
XPrt.Print recGuichet.Référence;
XPrt.FontBold = True

XPrt.CurrentX = 2700
XPrt.Print Compte_Imp(recGuichet.Compte);
XPrt.FontBold = False
XPrt.CurrentX = 3900
XPrt.Print recCptInfo.Intitulé;


'-----------------------------------Impression  du Montant--(cadré à droite )---
XPrt.FontBold = False
X = num_Display(recGuichet.Montant, 15, 2, Lx, X, "0") & " " & recGuichet.Sens
    XPrt.CurrentX = 10000 - XPrt.TextWidth(X)
    XPrt.Print X;
    
     
 '---------------------------------------------------------
X = num_Display(recGuichet.MontantEspèces, 15, 2, Lx, X, "0")
XPrt.CurrentX = 14000 - XPrt.TextWidth(X)
XPrt.Print X;

XPrt.FontBold = True
XPrt.CurrentX = 10100
XPrt.Print DevX(recGuichet.Devise);

XPrt.CurrentX = 14100
XPrt.Print DevX(recGuichet.DeviseEspèces);
XPrt.FontBold = False
XPrt.FontItalic = True

X = num_Display(recGuichet.CoursChange, 10, 5, Lx, X, "#")
XPrt.CurrentX = 11900 - XPrt.TextWidth(X)
XPrt.Print X;

X = num_Display(recGuichet.CoursChangeEspèces, 10, 5, Lx, X, "#")
    XPrt.CurrentX = 15500 - XPrt.TextWidth(X)
    XPrt.Print X;
XPrt.FontItalic = False
 

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 300
XPrt.Print Trim(DicLib(27, recGuichet.CodeOpération));

XPrt.FontSize = 6

XPrt.CurrentX = 3900
XPrt.Print recGuichet.Libellé;
XPrt.FontBold = True
XPrt.FontSize = 8

    
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY)
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
End Sub
'---------------------------------------------------------
Private Sub prtGuichetListLine_20010726()
'---------------------------------------------------------
 Dim X As String, chkX As String

If XPrt.CurrentY + prtlineHeight > prtMaxY Then
    frmElpPrt.prtNewPage
    prtGuichetListForm
'Else
 '   frmElpPrt.prtLineY
End If

XPrt.FontBold = False


'------------------------------------ligne 1
XPrt.FontSize = 8

XPrt.CurrentX = 300
XPrt.Print recGuichet.Référence;
XPrt.CurrentX = 900
XPrt.Print Trim(DicLib(27, recGuichet.CodeOpération));

X = Format$(recGuichet.CptMvtPièce, "####") & "." & Format$(recGuichet.CptMvtLigne, "0000")
    XPrt.CurrentX = 2200 - XPrt.TextWidth(X)
    XPrt.Print X;
       
XPrt.FontBold = True

XPrt.CurrentX = 2700
XPrt.Print Compte_Imp(recGuichet.Compte);
XPrt.FontBold = False
XPrt.CurrentX = 3900
XPrt.Print recCptInfo.Intitulé;
     X = num_Display(recGuichet.CoursChange, 10, 5, Lx, X, "#")
     XPrt.CurrentX = 9000 - XPrt.TextWidth(X)
     XPrt.Print X;
 
     X = num_Display(recGuichet.MontantAjustement, 6, 2, Lx, X, "#")
     XPrt.CurrentX = 9800 - XPrt.TextWidth(X)
     XPrt.Print X;



'-----------------------------------Impression  du Montant--(cadré à droite )---
XPrt.FontBold = False
X = num_Display(recGuichet.Montant, 15, 2, Lx, X, "0") & " " & recGuichet.Sens
    XPrt.CurrentX = 11400 - XPrt.TextWidth(X)
    XPrt.Print X;
    
    XPrt.CurrentX = 11600
XPrt.Print DevX(recGuichet.Devise);
      
 '---------------------------------------------------------
    X = num_Display(recGuichet.MontantEspèces, 15, 2, Lx, X, "0")
    XPrt.CurrentX = 15400 - XPrt.TextWidth(X)
    XPrt.Print X;


X = Format$(recGuichet.CptMvtPièceEspèces, "####") & "." & Format$(recGuichet.CptMvtLigneEspèces, "0000")
   XPrt.CurrentX = 12900 - XPrt.TextWidth(X)
   XPrt.Print X;
   
XPrt.CurrentX = 15600
XPrt.Print DevX(recGuichet.DeviseEspèces);


X = num_Display(recGuichet.CoursChangeEspèces, 10, 5, Lx, X, "#")
    XPrt.CurrentX = 13800 - XPrt.TextWidth(X)
    XPrt.Print X;
    
XPrt.CurrentX = 13900
Select Case recGuichet.optCours
    
    Case 1: XPrt.Print "N";
    Case 2: XPrt.Print "P";
    Case 3: XPrt.Print "M";
End Select

''----------------------------------------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.FontSize = 7
XPrt.CurrentX = 300
XPrt.Print "S : ";
XPrt.Print recGuichet.SaisieUsr;
XPrt.CurrentX = 1400
XPrt.Print dateImpS(recGuichet.SaisieAmj);
XPrt.CurrentX = 2050
XPrt.Print timeImpHM(recGuichet.SaisieHMS);
   
XPrt.CurrentX = 3900
XPrt.Print recGuichet.Libellé;
XPrt.FontBold = True
XPrt.FontSize = 8

chkX = ""
If recGuichet.chkCompte <> "0" Then chkX = "bloqué "
If recGuichet.chkSolde <> "0" Then chkX = chkX & "débiteur "

XPrt.CurrentX = 11900 - XPrt.TextWidth(chkX)
XPrt.Print chkX;
XPrt.FontBold = False

XPrt.FontSize = 7

     '---------------------------------------------------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 300
XPrt.Print "V : ";
XPrt.Print recGuichet.ValidationUsr;
XPrt.CurrentX = 1400
XPrt.Print dateImpS(recGuichet.ValidationAMJ);
XPrt.CurrentX = 2050
XPrt.Print timeImpHM(recGuichet.ValidationHMS);

If Trim(recGuichet.Identité) <> "" Then
    XPrt.CurrentX = 3900
    XPrt.Print recGuichet.Identité;

End If
'----------------------------------------------------------------------------

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 300
XPrt.Print "C : ";
XPrt.Print recGuichet.ComptaUsr;
XPrt.CurrentX = 1400
XPrt.Print dateImpS(recGuichet.ComptaAMJ);
XPrt.CurrentX = 2050
XPrt.Print timeImpHM(recGuichet.ComptaHMS);

XPrt.CurrentX = 2700
XPrt.Print "Val :";
XPrt.CurrentX = 3100
XPrt.Print dateImpS(recGuichet.AmjValeur);


  
If Trim(recGuichet.Complément1) <> "" Then
    XPrt.CurrentX = 3900
    XPrt.Print recGuichet.Complément1;
End If

If Trim(recGuichet.Complément2) <> "" Then
    XPrt.CurrentX = 3900
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.Print recGuichet.Complément2;
End If


If Trim(recGuichet.Complément3) <> "" Then
    XPrt.CurrentX = 3900
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.Print recGuichet.Complément3;
End If


         
              
      
 XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY)
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
End Sub

'---------------------------------------------------------
Public Sub prtGuichetListX(Msg As String)
'---------------------------------------------------------
Dim K As Integer, K1 As Integer, K2 As Integer
On Error GoTo prtError


Set XPrt = Printer
If prtShow Then frmElpPrt.Show vbModeless

prtOrientation = vbPRORLandscape
prtTitleText = "liste Guichet "
prtPgmName = "prtGuichet"
prtTitleUsr = usrName

prtLineNb = 1
prtlineHeight = 250
prtHeaderHeight = 300

frmElpPrt.prtStdInit
prtGuichetListForm

K1 = Val(mId$(Msg, 1, 6))
K2 = Val(mId$(Msg, 7, 6))
For K = K1 To K2
    recGuichet = arrGuichet(K)
    If mId$(recGuichet.ValidationUsr, 1, 6) <> "$TOTAL" Then
        If recGuichet.Devise <> recGuichet.DeviseEspèces Then
            
            If recGuichet.Devise = "978" And recGuichet.DeviseEspèces = "001" Then
            Else
                 If recGuichet.Devise = "001" And recGuichet.DeviseEspèces = "978" Then
                 Else
                    If Trim(recGuichet.ValidationUsr) <> Trim(constAnnulé) Then

                        recGuichet_mdbCptInfo recGuichet, recCptInfo '2001.08.23 jpl '''à modifier et remplacer par reccompte'
                        prtGuichetListLine
                    End If
                    DoEvents: If prtKillDoc Then frmElpPrt.Hide: Exit Sub
                End If
            End If
        End If
    End If
Next K

'frmElpPrt.prtLineY
frmElpPrt.prtEndDoc
If prtShow Then frmElpPrt.Hide
Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub





