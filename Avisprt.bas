Attribute VB_Name = "prtAvis"
Option Explicit
'---------------------------------------------------------
Dim recCptInfo As typeCptInfo
Dim recOpCpt As typeOpCpt, mOpCpt As typeOpCpt
Dim recDevise As typeDevise
Dim I As Integer, lX As Long, optAvisLangue As String * 1
Dim Net As Currency
Dim libTitre As String, libInfo1 As String, libInfo2 As String
Dim libBrut As String, libCompte As String
Dim libDate As String, libNet As String
Dim libComMontant1 As String, libComMontant2 As String, libComMontant3 As String
Dim libComMontant4 As String, libComMontant5 As String
Dim wOpCpt As typeOpCpt

Public Function lblOpTrf(ByVal Nat As String, ByVal X As String) As String
Select Case Trim(Nat)
    Case "Swift"
        Select Case Trim(X)
            Case "00": lblOpTrf = "00 : Emetteur"
            Case "01": lblOpTrf = "01 : Destinataire"
            Case "20": lblOpTrf = "20 : Numéro de transaction"
            Case "21": lblOpTrf = "21 : Référence message d'origine"
            Case "50": lblOpTrf = "50 : Client donneur d'ordre"
            Case "32": lblOpTrf = "32A: Date valeur, Devise, Montant"
            Case "52": lblOpTrf = "52 : Banque Ordonnatrice"
            Case "53": lblOpTrf = "53 : "
            Case "54": lblOpTrf = "54 : Correspondant"
            Case "57": lblOpTrf = "57 : Compte auprès de"
            Case "58": lblOpTrf = "58 : Banque bénéficiaire"
            Case "59": lblOpTrf = "59 : Client bénéficiaire"
            Case "70": lblOpTrf = "70 : Information du versement"
            Case "71": lblOpTrf = "71A: Détails des charges"
            Case "72": lblOpTrf = "72 : Info de l'émetteur au récepteur"
        End Select
    Case "optComShare"
        Select Case Trim(X)
            Case "1": lblOpTrf = "100 %"
            Case "2": lblOpTrf = "50 %"
            Case "3": lblOpTrf = "x %"
            Case Else: lblOpTrf = ""
        End Select
    Case "optComBarème"
        Select Case Trim(X)
            Case "1": lblOpTrf = "correspondant"
            Case "2": lblOpTrf = "entreprise"
            Case "3": lblOpTrf = "particulier"
            Case "3": lblOpTrf = "autre"
            Case Else: lblOpTrf = ""
        End Select
    Case "optComImputation"
        Select Case Trim(X)
            Case "1": lblOpTrf = "bénéficiaire"
            Case "2": lblOpTrf = "partage"
            Case "3": lblOpTrf = "donneur d'ordre"
            Case "4": lblOpTrf = "sans commission"
            Case Else: lblOpTrf = ""
        End Select
    Case "chkCom"
        Select Case Trim(X)
            Case "1": lblOpTrf = "commission de transfert"
            Case "2": lblOpTrf = "frais de Télex"
            Case "3": lblOpTrf = "frais Swift"
            Case "4": lblOpTrf = "commission de change"
            Case "5": lblOpTrf = "TVA"
            Case Else: lblOpTrf = ""
        End Select
    Case "chkComA"
        Select Case Trim(X)
            Case "1": lblOpTrf = "transfert commission"
            Case "2": lblOpTrf = "telex charge"
            Case "3": lblOpTrf = "swift charge"
            Case "4": lblOpTrf = "exchange commission"
            Case "5": lblOpTrf = "VAT"
            Case Else: lblOpTrf = ""
        End Select
End Select
End Function



'----------------------------------
Public Sub prtAvisX(XopCpt As typeOpCpt, lMsg As String)
'---------------------------------------------------------
On Error GoTo ErrorHandler
recOpCpt = XopCpt
prtAvis_Open
prtAvis_recOpCpt 0, lMsg
prtAvis_Close
ErrorHandler:
End Sub
Public Sub prtAvis_CptMvt(lCptMvt As typeCptMvt, lMsg As String)

prtAvis_CptMvt_OpCpt lCptMvt

prtAvisX wOpCpt, lMsg

End Sub

Public Sub prtAvis_CptEar(lCptMvt As typeCptMvt, lMsg As String)
On Error GoTo ErrorHandler

prtAvis_CptMvt_OpCpt lCptMvt

recOpCpt = wOpCpt
prtAvis_Open

prtAvis_recOpCpt 0, lMsg


ErrorHandler:

End Sub

Public Sub prtAvis_CptMvt_OpCpt(lCptMvt As typeCptMvt)

recOpCpt_Init wOpCpt
wOpCpt.Référence = Format(Val(lCptMvt.Pièce), "000000")
wOpCpt.CodeOpération = lCptMvt.CodeOpération
wOpCpt.Société = lCptMvt.Société
wOpCpt.Agence = lCptMvt.Agence
wOpCpt.Devise = Format(Val(lCptMvt.Devise), "000")
wOpCpt.Compte = lCptMvt.Compte
wOpCpt.Brut = lCptMvt.MT
wOpCpt.Sens = IIf(lCptMvt.MT < 0, "D", "C")
wOpCpt.AmjOpération = lCptMvt.AmjOpération
wOpCpt.AmjValeur = lCptMvt.AmjValeur
wOpCpt.Libellé = lCptMvt.Libellé
wOpCpt.optAvis = "1"

End Sub

'----------------------------------
Public Sub prtAvis(mCurrenty As Integer)
'----------------------------------
Dim H2 As Integer, I As Integer, X As String
XPrt.FontBold = True
XPrt.FontSize = 10
XPrt.CurrentY = mCurrenty + 1300
X = libTitre & " - " & recOpCpt.Référence
XPrt.CurrentX = (5000 - prtMinX - XPrt.TextWidth(X)) / 2
XPrt.Print X;

XPrt.FontSize = 7
XPrt.FontBold = False
XPrt.CurrentY = mCurrenty + 2300 - prtlineHeight * 2
XPrt.CurrentX = prtMinX
XPrt.Print libInfo2;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinX
XPrt.Print libInfo1;
I = 7
If recOpCpt.ComMontant1 <> 0 Then I = I + 1
If recOpCpt.ComMontant2 <> 0 Then I = I + 1
If recOpCpt.ComMontant3 <> 0 Then I = I + 1
If recOpCpt.ComMontant4 <> 0 Then I = I + 1
If recOpCpt.ComMontant5 <> 0 Then I = I + 1
H2 = 2600 / I

XPrt.FontSize = 7
XPrt.CurrentY = mCurrenty + 2300 + H2
XPrt.CurrentX = prtMinX + 100
XPrt.Print libCompte;
XPrt.FontBold = True
XPrt.Print Trim(XDevise.DevX) & " " & Compte_Imp(recOpCpt.Compte);

XPrt.CurrentY = XPrt.CurrentY + H2
XPrt.FontBold = False
XPrt.CurrentX = prtMinX + 100
XPrt.Print libDate & " " & dateImp(recOpCpt.AmjValeur);

XPrt.CurrentY = XPrt.CurrentY + H2 * 1.5

If Net <> recOpCpt.Brut Then
    XPrt.FontBold = False
    XPrt.CurrentX = prtMinX + 100
    XPrt.Print libBrut;
    X = num_Display(recOpCpt.Brut, 15, recDevise.maxD, lX, X, "0")
    XPrt.CurrentX = 4750 - XPrt.TextWidth(X)
    XPrt.Print X;
    XPrt.CurrentX = 4800: XPrt.Print recOpCpt.Sens;
    If recOpCpt.ComMontant1 <> 0 Then
        XPrt.CurrentY = XPrt.CurrentY + H2
        XPrt.CurrentX = prtMinX + 100
        XPrt.Print libComMontant1;
        X = num_Display(recOpCpt.ComMontant1, 15, recDevise.maxD, lX, X, "0")
        XPrt.CurrentX = 4750 - XPrt.TextWidth(X)
        XPrt.Print X;
        XPrt.CurrentX = 4800: XPrt.Print "D";
    End If
    
    
    If recOpCpt.ComMontant2 <> 0 Then
        XPrt.CurrentY = XPrt.CurrentY + H2
        XPrt.CurrentX = prtMinX + 100
        XPrt.Print libComMontant2;
        X = num_Display(recOpCpt.ComMontant2, 15, recDevise.maxD, lX, X, "0")
        XPrt.CurrentX = 4750 - XPrt.TextWidth(X)
        XPrt.Print X;
        XPrt.CurrentX = 4800: XPrt.Print "D";
    End If
    
    If recOpCpt.ComMontant3 <> 0 Then
        XPrt.CurrentY = XPrt.CurrentY + H2
        XPrt.CurrentX = prtMinX + 100
        XPrt.Print libComMontant3;
        X = num_Display(recOpCpt.ComMontant3, 15, recDevise.maxD, lX, X, "0")
        XPrt.CurrentX = 4750 - XPrt.TextWidth(X)
        XPrt.Print X;
        XPrt.CurrentX = 4800: XPrt.Print "D";
    End If
    
    If recOpCpt.ComMontant4 <> 0 Then
        XPrt.CurrentY = XPrt.CurrentY + H2
        XPrt.CurrentX = prtMinX + 100
        XPrt.Print libComMontant4;
        X = num_Display(recOpCpt.ComMontant4, 15, recDevise.maxD, lX, X, "0")
        XPrt.CurrentX = 4750 - XPrt.TextWidth(X)
        XPrt.Print X;
        XPrt.CurrentX = 4800: XPrt.Print "D";
    End If
    
    If recOpCpt.ComMontant5 <> 0 Then
        XPrt.CurrentY = XPrt.CurrentY + H2
        XPrt.CurrentX = prtMinX + 100
        XPrt.Print libComMontant5;
        X = num_Display(recOpCpt.ComMontant5, 15, recDevise.maxD, lX, X, "0")
        XPrt.CurrentX = 4750 - XPrt.TextWidth(X)
        XPrt.Print X;
        XPrt.CurrentX = 4800: XPrt.Print "D";
    End If
    XPrt.CurrentY = XPrt.CurrentY + H2 + 50
    XPrt.Line (3550, XPrt.CurrentY)-(4900, XPrt.CurrentY)
End If

XPrt.FontBold = True
XPrt.CurrentY = XPrt.CurrentY + 50
XPrt.CurrentX = prtMinX + 100
XPrt.Print libNet;
X = num_Display(Net, 15, recDevise.maxD, lX, X, "0")
XPrt.CurrentX = 4750 - XPrt.TextWidth(X)
XPrt.Print X;
XPrt.CurrentX = 4800: XPrt.Print recOpCpt.Sens;
XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + H2
XPrt.CurrentX = prtMinX + 100
XPrt.Print recOpCpt.Libellé;

End Sub
Public Sub prtAvisLib(Sens As String, optLangue As String)
libInfo2 = ""
Select Case optLangue
    Case "1"
        If Sens = "D" Then
            libTitre = "AVIS DE DEBIT"
            libInfo1 = "Nous avons l'honneur de vous informer que nous débitons selon détail ci-après :"
        Else
            libTitre = "AVIS DE CREDIT"
            libInfo1 = "Nous avons l'honneur de vous informer que nous créditons selon détail ci-après :"
        End If
        libCompte = "votre compte "
        libBrut = "montant transféré"
        libNet = "montant net"
        libDate = "date valeur:"
        libComMontant1 = lblOpTrf("chkCom", "1")
        libComMontant2 = lblOpTrf("chkCom", "2")
        libComMontant3 = lblOpTrf("chkCom", "3")
        libComMontant4 = lblOpTrf("chkCom", "4")
        libComMontant5 = lblOpTrf("chkCom", "5")
    Case "2"
        If Sens = "D" Then
            libTitre = "DEBIT ADVICE"
            libInfo1 = "We beg to inform you that we are debiting as follows :"
        Else
            libTitre = "CREDIT ADVICE"
            libInfo1 = "We beg to inform you that we are crediting as follows :"
        End If
        libCompte = "your account "
        libBrut = "transferred amount"
        libNet = "net amount"
        libDate = "value date :"
        libComMontant1 = lblOpTrf("chkComA", "1")
        libComMontant2 = lblOpTrf("chkComA", "2")
        libComMontant3 = lblOpTrf("chkComA", "3")
        libComMontant4 = lblOpTrf("chkComA", "4")
        libComMontant5 = lblOpTrf("chkComA", "5")
    Case Else
        If Sens = "D" Then
            libTitre = "AVIS DE DEBIT / DEBIT ADVICE"
            libInfo1 = "Nous avons l'honneur de vous informer que nous débitons selon détail ci-après :"
            libInfo2 = "We beg to inform you that we are debiting as follows :"
        Else
            libTitre = "AVIS DE CREDIT / CREDIT ADVICE"
            libInfo1 = "Nous avons l'honneur de vous informer que nous créditons selon détail ci-après :"
            libInfo2 = "We beg to inform you that we are crediting as follows :"
        End If
        libCompte = "votre compte numéro / your account number : "
        libBrut = "montant transféré / transferred amount"
        libNet = "montant net / net amount"
        libDate = "valeur / value date :"
        libComMontant1 = lblOpTrf("chkCom", "1") & " / " & lblOpTrf("chkComA", "1")
        libComMontant2 = lblOpTrf("chkCom", "2") & " / " & lblOpTrf("chkComA", "2")
        libComMontant3 = lblOpTrf("chkCom", "3") & " / " & lblOpTrf("chkComA", "3")
        libComMontant4 = lblOpTrf("chkCom", "4") & " / " & lblOpTrf("chkComA", "4")
        libComMontant5 = lblOpTrf("chkCom", "5") & " / " & lblOpTrf("chkComA", "5")
End Select
End Sub

Public Sub prtAvis_Open()
On Error GoTo prtError

Set XPrt = Printer
frmElpPrt.Show vbModeless

prtOrientation = vbPRORPortrait
prtTitleText = "Avis"
prtPgmName = "prtAvis"
prtTitleUsr = usrName

prtLineNb = 1
prtlineHeight = 250
prtHeaderHeight = 300

prtFormType = ""
frmElpPrt.prtInit

Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide
End Sub

Public Sub prtAvis_Close()
frmElpPrt.prtEndDoc
frmElpPrt.Hide

End Sub

Public Sub prtAvis_recOpCpt(mCurrenty As Integer, lMsg As String)
DevX (recOpCpt.Devise)
recDevise = XDevise

recCptInfoInit recCptInfo
recCptInfo.Method = "JoinL1"
recCptInfo.Société = recOpCpt.Société
recCptInfo.Agence = recOpCpt.Agence
recCptInfo.Devise = recOpCpt.Devise
recCptInfo.Numéro = recOpCpt.Compte
recCptInfo.BiaTyp = "000"
recCptInfo.BiaNum = "00000"
recCptInfo.NuméroAncien = "00000000000"
If Not IsNull(srvCptInfoFind(recCptInfo)) Then recCptInfoInit recCptInfo
    XPrt.CurrentY = mCurrenty
    If recOpCpt.Sens = "D" Then
        Net = recOpCpt.Brut + recOpCpt.ComMontantDevise
    Else
        Net = recOpCpt.Brut - recOpCpt.ComMontantDevise
    End If
    Call prtAvisLib(recOpCpt.Sens, recOpCpt.optAvisLangue)
    prtSocMini mCurrenty, recOpCpt.AmjOpération
    prtAdresse mCurrenty, recCptInfo
    Call frmElpPrt.prtTrame(prtMinX, mCurrenty + 2300, 5000, mCurrenty + 4900, "B")
    prtAvis mCurrenty
    If Trim(lMsg) <> "" Then
        XPrt.FontSize = 6
        XPrt.FontBold = False
        XPrt.CurrentY = mCurrenty + 2300 - prtlineHeight
        XPrt.CurrentX = 11000 - XPrt.TextWidth(lMsg)
        XPrt.Print lMsg;
    End If
    XPrt.CurrentY = mCurrenty + 5000
    prtSocMiniFin
    If mCurrenty < 10000 Then
        XPrt.CurrentY = mCurrenty + 5100
        frmElpPrt.prtTiret
    End If
''End If
End Sub

Public Sub prtAvis_Global(lstSort As ListBox)
Dim K As Integer, mCurrenty As Integer
Dim blnTest As Boolean
On Error GoTo ErrorHandler

prtAvis_Open
mCurrenty = 0
blnTest = False

For K = 0 To lstSort.ListCount - 1
    lstSort.ListIndex = K
    arrOpCptIndex = Val(mId$(lstSort.Text, 24, 5))
    
    recOpCpt = arrOpCpt(arrOpCptIndex)
    If blnTest Then
        If recOpCpt.Devise <> mOpCpt.Devise _
        Or recOpCpt.Compte <> mOpCpt.Compte Then
            XPrt.NewPage
            mCurrenty = 0
        Else
            Select Case mCurrenty
                Case 0: mCurrenty = 5400
                Case 5400: mCurrenty = 11000
                Case 11000: XPrt.NewPage: mCurrenty = 0
            End Select
        End If
    End If
    blnTest = True
    mOpCpt = recOpCpt
    
    Call prtAvis_recOpCpt(mCurrenty, " ")
    DoEvents: If prtKillDoc Then frmElpPrt.Hide: Exit Sub
Next K
prtAvis_Close
ErrorHandler:
End Sub
