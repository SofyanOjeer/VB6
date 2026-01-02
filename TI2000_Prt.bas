Attribute VB_Name = "prtTI2000"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim x As String, I As Integer, Height8_6 As Integer

Dim recCDDossier As typeCDDossier
Dim xAmjSituation As String, xAmjMin As String, xAmjMax  As String

Dim Nb1 As Integer, Nb2 As Integer
Dim curSolde As Currency, curS36Solde As Currency
Dim prtEtat As String
Dim blnDevise As Boolean
'---------------------------------------------------------
 Public Sub prtTI2000_Monitor(Msg As String)
'---------------------------------------------------------
Dim K As Long, xK1 As Long, xK2 As Long
Dim x As String, curX As Currency

On Error GoTo prtError

Set XPrt = Printer
Height8_6 = frmElpPrt.prtHeightDelta(8, 6)
xK1 = mId$(Msg, 1, 6)
xK2 = mId$(Msg, 7, 6)
prtEtat = mId$(Msg, 13, 2)

Select Case prtEtat
    Case "SG": x = "Crédits documentaires "
    Case "SD": x = "Crédits documentaires(diff TI / S36) "
    Case "SA": x = "Crédits documentaires soldés (TI = S36 = 0) "
    Case "SI": x = "Crédits documentaires non soldés (TI = S36 <> 0)"
    Case "UP": x = "Crédoc :( Util - Paiement < 5 %) "
    Case Else: x = "Crédits documentaires "
End Select

prtTitleText = x & " : Validité au " & dateImp(paramTI2000DB2_AMJValidité)
If Trim(selCDComD_Devise) = "" Then
    blnDevise = False
Else
    blnDevise = True
End If

prtLineNb = 1

frmElpPrt.Show vbModeless


prtOrientation = vbPRORPortrait
prtPgmName = "prtTI2000"
prtTitleUsr = usrName

prtlineHeight = 250
prtHeaderHeight = 300

frmElpPrt.prtStdInit

mdbCDDossier.tableCDDossier_Open
recCDDossier_Init recCDDossier
recCDDossier.Method = "Seek>="
recCDDossier.Dossier = xK1

prtTI2000_Form
Nb1 = 0

Do
    intReturn = tableCDDossier_Read(recCDDossier)

    If intReturn = 0 Then
        If recCDDossier.Dossier > xK2 Then
            intReturn = -1
        Else
            If Not blnDevise Or (blnDevise And selCDComD_Devise = recCDDossier.Devise) Then
                
                If Trim(recCDDossier.AMJSituation) = "" And recCDDossier.AMJValidité <= paramTI2000DB2_AMJValidité Then
                    curSolde = recCDDossier.MontantEngagement - recCDDossier.MontantUtilisé
                     curS36Solde = recCDDossier.S36Engagement - recCDDossier.S36Utilisé
                     Select Case prtEtat
                        Case "SG": prtTI2000_Line
                        
                        Case "SA": If curSolde = 0 And curS36Solde = curSolde Then prtTI2000_Line
                        
                        Case "SI": If curSolde <> 0 And curS36Solde = curSolde Then prtTI2000_Line
                        
                        Case "SD":
                             'If recCDDossier.s36Engagement <> recCDDossier.MontantEngagement _
                             'Or recCDDossier.curS36Solde <> curSolde Then
                                 If curS36Solde <> curSolde Then prtTI2000_Line
                        Case "UP":
                                curX = recCDDossier.TIMt226 - recCDDossier.MontantUtilisé
                             If Abs(curX) > (recCDDossier.MontantUtilisé * 5 / 100) Then
                             'Or recCDDossier.curS36Solde <> curSolde Then
                               ' If Trim(recCDDossier.AMJSituation) = "" Then
                                      prtTI2000_LineUP
                               '   End If
                                End If
                          
                    End Select
                End If
            End If
        recCDDossier.Method = "MoveNext"
    End If
    End If
Loop While intReturn = 0
        
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY)
XPrt.CurrentY = XPrt.CurrentY + 50
XPrt.FontBold = True
XPrt.CurrentX = prtMinX: XPrt.Print Nb1 & " dossiers";
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY)

DoEvents: If prtKillDoc Then frmElpPrt.Hide: Exit Sub

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
Public Sub prtTI2000_Form()
'---------------------------------------------------------
Dim x As String
XPrt.FontSize = 6

XPrt.FontBold = True
XPrt.DrawWidth = 3
XPrt.CurrentY = prtMinY
prtCurrentY = prtMinY + prtlineHeight
Call frmElpPrt.prtTrame(prtMinX, prtMinY, prtMaxX, prtCurrentY, "B", 250)
'XPrt.Line (prtMinX, prtCurrentY)-(prtMaxX, prtCurrentY)

XPrt.DrawWidth = 1
Call frmElpPrt.prtTrame(prtMinX + 4620, prtCurrentY + 20, prtMinX + 7480, prtMaxY - 20, " ", 250)

XPrt.Line (prtMinX + 2350, prtMinY)-(prtMinX + 2350, prtMaxY)
XPrt.Line (prtMinX + 4600, prtMinY)-(prtMinX + 4600, prtMaxY)
XPrt.Line (prtMinX + 7500, prtMinY)-(prtMinX + 7500, prtMaxY)

'---------------------------------------------------------

x = "Dossier"
XPrt.CurrentY = prtMinY + (prtHeaderHeight - XPrt.TextHeight(x)) / 2
XPrt.CurrentX = prtMinX: XPrt.Print " Dossier";
XPrt.CurrentX = prtMinX + 600: XPrt.Print "    Ouverture";
XPrt.CurrentX = prtMinX + 1500: XPrt.Print "    Validité";
XPrt.CurrentX = prtMinX + 3000: XPrt.Print "Correspondant";

Select Case prtEtat
    Case "UP"
        XPrt.CurrentX = prtMinX + 5200: XPrt.Print "Engagement";
        XPrt.CurrentX = prtMinX + 6800: XPrt.Print "Utilisation";
        XPrt.CurrentX = prtMinX + 8500: XPrt.Print " Paiement";
        XPrt.CurrentX = prtMinX + 10000: XPrt.Print "Différence";

    Case Else
        XPrt.CurrentX = prtMinX + 5000: XPrt.Print "Engagement";
        XPrt.CurrentX = prtMinX + 7000: XPrt.Print "Solde";
        XPrt.CurrentX = prtMinX + 8000: XPrt.Print " S36 Engagement";
        XPrt.CurrentX = prtMinX + 10000: XPrt.Print "S36 Solde";

End Select


XPrt.CurrentY = prtMinY + prtHeaderHeight - XPrt.TextHeight("X")

XPrt.FontSize = 6

End Sub

'---------------------------------------------------------
Public Sub prtTI2000_Line()
'---------------------------------------------------------

If XPrt.CurrentY + prtlineHeight * 1.5 > prtMaxY Then
    frmElpPrt.prtNewPage
    prtTI2000_Form
End If

XPrt.FontBold = False
'_______________________________________________________________ligne 1-

Nb1 = Nb1 + 1

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.FontBold = True
XPrt.CurrentX = prtMinX: XPrt.Print recCDDossier.Dossier;
XPrt.FontBold = False
XPrt.CurrentX = prtMinX + 600: XPrt.Print dateImp(recCDDossier.AMJOuverture);
XPrt.CurrentX = prtMinX + 1500: XPrt.Print dateImp(recCDDossier.AMJValidité);
XPrt.CurrentX = prtMinX + 2500: XPrt.Print recCDDossier.Compte;

If recCDDossier.MontantEngagement <> 0 Then
    x = Format$(recCDDossier.MontantEngagement, "## ### ### ### ### ##0.00")
    XPrt.CurrentX = prtMinX + 5800 - XPrt.TextWidth(x)
    XPrt.Print x;
End If
XPrt.CurrentX = prtMinX + 6100: XPrt.Print recCDDossier.Devise;
If curSolde <> 0 Then
    x = Format$(curSolde, "## ### ### ### ### ##0.00")
    XPrt.CurrentX = prtMinX + 7450 - XPrt.TextWidth(x)
    XPrt.Print x;
End If


If recCDDossier.S36Engagement <> 0 Then
    x = Format$(recCDDossier.S36Engagement, "## ### ### ### ### ##0.00")
    XPrt.CurrentX = prtMinX + 9000 - XPrt.TextWidth(x)
    XPrt.Print x;
End If
If curS36Solde <> 0 Then
    x = Format$(curS36Solde, "## ### ### ### ### ##0.00")
    XPrt.CurrentX = prtMinX + 10500 - XPrt.TextWidth(x)
    XPrt.Print x;
    If curS36Solde < 0 Then XPrt.Print " -";

End If
End Sub

'---------------------------------------------------------
Public Sub prtTI2000_LineUP()
'---------------------------------------------------------
Dim curX As Currency

If XPrt.CurrentY + prtlineHeight * 1.5 > prtMaxY Then
    frmElpPrt.prtNewPage
    prtTI2000_Form
End If

XPrt.FontBold = False
'_______________________________________________________________ligne 1-


XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.FontBold = True
XPrt.CurrentX = prtMinX: XPrt.Print recCDDossier.Dossier;
XPrt.FontBold = False
XPrt.CurrentX = prtMinX + 600: XPrt.Print dateImp(recCDDossier.AMJOuverture);
XPrt.CurrentX = prtMinX + 1500: XPrt.Print dateImp(recCDDossier.AMJValidité);
XPrt.CurrentX = prtMinX + 2500: XPrt.Print recCDDossier.Compte;

If recCDDossier.MontantEngagement <> 0 Then
    x = Format$(recCDDossier.MontantEngagement, "## ### ### ### ### ##0.00")
    XPrt.CurrentX = prtMinX + 6000 - XPrt.TextWidth(x)
    XPrt.Print x;
End If
XPrt.CurrentX = prtMinX + 6100: XPrt.Print recCDDossier.Devise;
If recCDDossier.MontantUtilisé <> 0 Then
    x = Format$(recCDDossier.MontantUtilisé, "## ### ### ### ### ##0.00")
    XPrt.CurrentX = prtMinX + 7450 - XPrt.TextWidth(x)
    XPrt.Print x;
End If


If recCDDossier.TIMt226 <> 0 Then
    x = Format$(recCDDossier.TIMt226, "## ### ### ### ### ##0.00")
    XPrt.CurrentX = prtMinX + 9000 - XPrt.TextWidth(x)
    XPrt.Print x;
End If
curX = recCDDossier.MontantUtilisé - recCDDossier.TIMt226

If curX <> 0 Then
    x = Format$(Abs(curX), "## ### ### ### ### ##0.00")
    XPrt.CurrentX = prtMinX + 10500 - XPrt.TextWidth(x)
    XPrt.Print x;
    If curX < 0 Then XPrt.Print " -";

End If
XPrt.CurrentX = prtMinX + 10800: XPrt.Print Trim(recCDDossier.AMJSituation);
End Sub



