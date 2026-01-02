Attribute VB_Name = "prtGuichet"
Option Explicit
Dim mCurrenty As Integer, mCurrentX As Integer
Dim X As String, curX As Currency
Dim recGuichet As typeGuichet
Dim recCptInfo As typeCptInfo
Dim libOpération As String, libMsg As String
Dim libIdentité As String, libComplément As String
Dim libMontantEspèces As String, libMontantRendu As String
Dim libContreValeur As String
Dim xChange As Double
'----------------------------------
Public Sub prtguichetX(Exemplaire As String, xGuichet As typeGuichet, recCV1 As typeCV, recCV2 As typeCV)
'----------------------------------
Dim X As String, xCodeOpération As String
On Error GoTo prtError


Set XPrt = Printer
recGuichet = xGuichet
recGuichet_CptInfo recGuichet, recCptInfo
'If prtShow Then frmElpPrt.Show vbModeless
prtOrientation = vbPRORPortrait
prtTitleText = ""
prtPgmName = "prtGuichet"
prtTitleUsr = usrName

prtLineNb = 0
prtlineHeight = 250
prtHeaderHeight = 0

prtFormType = ""
frmElpPrt.prtInit

If arrDeviseCoupuresNb = 0 Then
    srvDeviseCoupures_Load recCV1.DeviseIso
Else
    If arrDeviseCoupures(1).Id <> recCV1.DeviseIso Then
        srvDeviseCoupures_Load recCV1.DeviseIso
    End If
End If

prtSocMini 0, recGuichet.SaisieAmj
XPrt.CurrentY = 7800
prtSocMiniFin
xCodeOpération = Trim(recGuichet.CodeOpération)
Select Case xCodeOpération
    Case "G001", "G006": libOpération = "Versement Espèces"
                 libMsg = "Nous créditons"
                 libIdentité = "Déposant :"
                 libComplément = ""
                 libMontantEspèces = " versés :"
                 libMontantRendu = " rendus :"
                 libContreValeur = "contre-valeur du versement de "
    Case "G002", "G005":
                If recGuichet.Devise = recGuichet.DeviseEspèces Or recCV1.DeviseIso = "FRF" Then
                    libOpération = "Retrait Espèces"
                Else
                    libOpération = "Délivrance de devises"
                End If
                libMsg = "Nous débitons"
                libIdentité = "Bénéficiaire : "
                libMontantEspèces = " retirés :"
                libMontantRendu = " rendus  :"
                libComplément = ""
                libContreValeur = "contre-valeur de la délivrance de "
  Case "G007": libOpération = "Change"
                libMsg = ""
                libIdentité = "Identité :"
                libComplément = ""
                libMontantEspèces = " versés :"
                libMontantRendu = " rendus :"
                libContreValeur = "contre-valeur du versement de "
     Case "G008": libOpération = "Arbitrage"
                libMsg = "Nous débitons"
                libIdentité = ""
                libComplément = ""
                libMontantEspèces = " versés :"
                libMontantRendu = " rendus :"
                libContreValeur = "pour créditer votre compte en " & Trim(recCV1.DeviseLibellé) & " de "
Case Else: libOpération = "?"
                 libMsg = "? "
End Select
If Exemplaire = 2 Then libOpération = libOpération & " (Pièce comptable )"
If recGuichet.CodeOpération <> "G008" Then prtGuichetForm recGuichet

XPrt.FontBold = True
XPrt.CurrentY = 5350: XPrt.CurrentX = 200
XPrt.Print "Signature Client";
XPrt.CurrentY = 6350: XPrt.CurrentX = 200
XPrt.Print "Visa BIA";
XPrt.FontBold = False

If recGuichet.chkCoupureEspèces = "1" Then
    XPrt.CurrentY = 4350: XPrt.CurrentX = 0
    XPrt.Print recCV1.DeviseIso & libMontantEspèces;
    X = num_Display(recGuichet.MontantEspèces + recGuichet.MontantRendu, 15, recCV1.maxD, lX, X, "0")
    XPrt.CurrentX = 2350 - XPrt.TextWidth(X)
    XPrt.Print X;
    If recGuichet.MontantRendu <> 0 Then
        XPrt.CurrentY = XPrt.CurrentY + prtlineHeight: XPrt.CurrentX = 0
        XPrt.Print recCV1.DeviseIso & libMontantRendu;
        X = num_Display(recGuichet.MontantRendu, 15, recCV1.maxD, lX, X, "0")
        XPrt.CurrentX = 2350 - XPrt.TextWidth(X)
        XPrt.Print X;
    End If
End If

prtAdresse 0, recCptInfo

XPrt.FontBold = True
XPrt.FontSize = 14
XPrt.CurrentY = 4800: mCurrenty = XPrt.CurrentY + XPrt.TextHeight(libOpération)
frmElpPrt.prtCentré 6750, libOpération

XPrt.FontSize = 9
XPrt.FontBold = False
If recGuichet.chkChèque > "0" Then
     XPrt.CurrentY = mCurrenty - XPrt.TextHeight("12345")
     Select Case recGuichet.chkChèque
        Case Is = "1": XPrt.Print " (chèque n° " & Format$(recGuichet.NoChèque, "0000000") & ")"
        Case Is = "2": XPrt.Print " (chèque guichet n° " & Format$(recGuichet.NoChèque, "0000000") & ")"
    End Select
End If
'------------------------------------------------------
XPrt.CurrentY = 5350: XPrt.CurrentX = 2500
If xCodeOpération = "G007" Then
    XPrt.Print "retrait en espèces";
Else
    XPrt.Print libMsg & " votre compte numéro ";
    XPrt.FontBold = True
    XPrt.Print Compte_Imp(recGuichet.Compte);
    XPrt.FontBold = False
    XPrt.Print ", valeur ";
    XPrt.FontBold = True
    XPrt.Print dateImp(recGuichet.AmjValeur);
    XPrt.FontBold = False
End If

XPrt.Print ", de ";
XPrt.FontBold = True
X = recCV2.DeviseIso & num_Display(recGuichet.Montant, 15, recCV2.maxD, lX, X, "0")
XPrt.CurrentX = 10950 - XPrt.TextWidth(X)
XPrt.Print X;
XPrt.FontBold = False
If recGuichet.Devise <> recGuichet.DeviseEspèces Then
XPrt.FontSize = 8
 XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
 XPrt.CurrentX = 2500
 XPrt.Print libContreValeur;
 XPrt.FontBold = True
 XPrt.Print recCV1.DeviseIso & " " & Trim(num_Display(recGuichet.MontantEspèces, 15, recCV1.maxD, lX, X, "0"));
 XPrt.FontBold = False
 XPrt.Print " au cours de ";
 If recCV1.DeviseIso = "FRF" Then
     xChange = recGuichet.CoursChangeEspèces / recGuichet.CoursChange
  Else
     xChange = recGuichet.CoursChange / recGuichet.CoursChangeEspèces
End If
    XPrt.Print Trim(num_Display(xChange, 12, 5, lX, X, "0"));
'    If recGuichet.CoursChangeEspèces <> 1 Then XPrt.Print Trim(num_Display(recGuichet.CoursChangeEspèces, 12, 5, Lx, X, "0"));
'    If recGuichet.CoursChange <> 1 Then XPrt.Print " / " & Trim(num_Display(recGuichet.CoursChange, 12, 5, Lx, X, "0"));
End If

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 2500
mCurrentX = XPrt.CurrentX
If Trim(recGuichet.Identité) <> "" Then
    XPrt.Print libIdentité;
    mCurrentX = XPrt.CurrentX
    XPrt.Print recGuichet.Identité;
    XPrt.CurrentX = XPrt.CurrentX + 300
End If
If Trim(recGuichet.Complément1) <> "" Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = mCurrentX: XPrt.Print recGuichet.Complément1;
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = mCurrentX: XPrt.Print recGuichet.Complément2;
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = mCurrentX: XPrt.Print recGuichet.Complément3;
End If

XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 2500
X = "(" & MontantEnLettres(recGuichet.Montant, recCV2.DeviseLibellé) & ")"
XPrt.FontSize = 7
XPrt.CurrentX = 10950 - XPrt.TextWidth(X)
XPrt.Print X;

XPrt.FontSize = 9

XPrt.Line (2400, 5300)-(11000, XPrt.CurrentY + prtlineHeight), , B

XPrt.CurrentY = 7550
XPrt.CurrentX = 200
XPrt.Print "Service émetteur : Caisse (tél: 01 53 76 62 62)";
XPrt.CurrentX = 5000
XPrt.Print "Notre référence : " & Format$(recGuichet.CptMvtPièce, "#####") & "." & Format$(recGuichet.CptMvtLigne, "0000");
'XPrt.CurrentX = 9000
'XPrt.Print "( Euros : xxxxxx  )";

If Exemplaire = "2" Then
    If recGuichet.chkCompte <> "0" Or recGuichet.chkSolde <> "0" Then prtGuichetAutorisation
End If
DoEvents: If prtKillDoc Then frmElpPrt.Hide: Exit Sub

frmElpPrt.prtEndDoc
'If prtShow Then frmElpPrt.Hide
Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide
End Sub
Public Sub prtGuichetForm(recGuichet As typeGuichet)
Dim I As Integer, K As Integer, Kb As Integer, Kp As Integer, Kc As Integer

XPrt.DrawStyle = 0
'------------------------------------------------
XPrt.DrawWidth = 3
XPrt.Line (500, 1550)-(500, 4300)
XPrt.Line (1200, 1550)-(1200, 4300)
XPrt.Line (1700, 1550)-(1700, 4300)
XPrt.DrawWidth = 1
'----------------------------------------lignes-------------
XPrt.Line (0, 1550)-(2400, 1550)
XPrt.Line (0, 1800)-(2400, 1800)
XPrt.Line (0, 2050)-(2400, 2050)
XPrt.Line (0, 2300)-(2400, 2300)
XPrt.Line (0, 2550)-(2400, 2550)
XPrt.Line (0, 2800)-(2400, 2800)
XPrt.Line (0, 3050)-(2400, 3050)
XPrt.Line (0, 3300)-(2400, 3300)
XPrt.Line (0, 3550)-(2400, 3550)
XPrt.Line (0, 3800)-(2400, 3800)
XPrt.Line (0, 4050)-(2400, 4050)
XPrt.Line (0, 4300)-(2400, 4300)
'-------------------------------------------------
'-------------------------------------------------------
XPrt.FontSize = 8
XPrt.FontBold = True
XPrt.CurrentY = 1350: XPrt.CurrentX = 700
XPrt.Print "Billets";
XPrt.CurrentY = 1350: XPrt.CurrentX = 1850
XPrt.Print "Pièces";


XPrt.FontBold = False
K = 0: Kb = -1: Kp = -1
For I = 1 To arrDeviseCoupuresNb
    If arrDeviseCoupures(I).Actif = " " Then
        XPrt.FontBold = False
        If arrDeviseCoupures(I).Nature = "B" Then
            Kb = Kb + 1: XPrt.CurrentY = 1600 + 250 * Kb
            K = 450
        Else
            Kp = Kp + 1: XPrt.CurrentY = 1600 + 250 * Kp
            K = 1650
        End If
        curX = arrDeviseCoupures(I).Nominal
        If curX = Fix(curX) Then
            X = Format$(curX, "###")
        Else
             X = Format$(curX, "##0.00")
       End If
        XPrt.CurrentX = K - XPrt.TextWidth(X)
        XPrt.Print X;
        Kc = arrDeviseCoupures(I).Séquence * 4 - 3
        If Kc > 0 Then
            XPrt.FontBold = True
            X = Format$(Val(mId$(recGuichet.CoupureEspèces, Kc, 4)), "####")
            XPrt.CurrentX = K + 700 - XPrt.TextWidth(X)
            XPrt.Print X;
        End If
    End If
Next I

End Sub




Public Sub prtGuichetAutorisation()
Dim strSens As String
Call frmElpPrt.prtTrame(2600, 1550, 5400, 4300, " ", "220")
XPrt.FontSize = 7
XPrt.CurrentX = 2650: XPrt.CurrentY = 1600
XPrt.FontBold = True
frmElpPrt.prtCentré 4000, "Autorisation"
XPrt.FontBold = False
If recCptInfo.Situation <> " " Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight: XPrt.CurrentX = 2650
    Select Case recCptInfo.Situation
        Case "B": XPrt.Print "Compte bloqué"
        Case "A": XPrt.Print "Compte annulé"
        Case Else: XPrt.Print "Compte : " & recCptInfo.Situation
    End Select
End If

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight: XPrt.CurrentX = 2650
XPrt.Print "Solde (" & timeImpHM(recGuichet.SaisieHMS) & ")";
X = num_Display(recCptInfo.SoldeInstantané, 15, 2, lX, strSens, "0")
XPrt.CurrentX = 5050 - XPrt.TextWidth(X)
XPrt.Print X & " " & strSens;
curX = 0
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight: XPrt.CurrentX = 2650
XPrt.Print "Déc autorisé ";
If recCptInfo.DécouvertMontant > 0 Then
    If Val(recCptInfo.DécouvertAmj) < DSys Then
        XPrt.Print "échu ";
    Else
        curX = recCptInfo.DécouvertMontant
        X = num_Display(recCptInfo.DécouvertMontant, 15, 2, lX, strSens, "0")
        XPrt.CurrentX = 5050 - XPrt.TextWidth(X)
        XPrt.Print X & " " & strSens;
        XPrt.CurrentY = XPrt.CurrentY + prtlineHeight: XPrt.CurrentX = 2650
        XPrt.Print "juqu'au : " & dateImp(recCptInfo.DécouvertAmj);
    End If
End If

curX = curX + recCptInfo.SoldeInstantané - recGuichet.MontantEspèces
If curX < 0 Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight: XPrt.CurrentX = 2650
    XPrt.Print "Dépassement";
    X = num_Display(curX, 15, 2, lX, strSens, "0")
    XPrt.CurrentX = 5050 - XPrt.TextWidth(X)
    XPrt.Print X & " " & strSens;
End If
End Sub
