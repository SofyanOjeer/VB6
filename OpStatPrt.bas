Attribute VB_Name = "prtOpStat"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------

Private recOpStat As typeOpStat
Dim I As Integer, mTrame As Integer
Dim NbImprimé As Integer

Type typeTotalOpstat
    Nb      As Integer
    Brut    As Currency
    Com     As Currency
    K1      As String * 11
    K2      As String * 11
End Type

Private totalZ As typeTotalOpstat
Private totalX As typeTotalOpstat
Private totalK(3) As typeTotalOpstat
Dim blndeviseK1 As Boolean, blndeviseK2 As Boolean

Dim currentXK1 As Integer, currentXK2 As Integer
Dim currentXK1B As Integer, currentXK2B As Integer

Dim currentXDevise As Integer, currentXCorrespondant As Integer, _
    currentXSender As Integer, currentXReceiver As Integer, _
    currentXRéférence As Integer
'---------------------------------------------------------
Public Sub prtOpStat_Form()
'---------------------------------------------------------
Dim X As String
NbImprimé = 0
XPrt.FontSize = 6
XPrt.FontBold = True

XPrt.DrawWidth = 2

Call frmElpPrt.prtTrame(prtMinX, prtMinY, prtMaxX, prtMinY + prtHeaderHeight, "B")
'Call frmElpPrt.prtTrame(currentXK1, prtMinY + prtHeaderHeight + 10, currentXK1B, prtMaxY - 10, " ", 250)
'Call frmElpPrt.prtTrame(currentXK2, prtMinY + prtHeaderHeight + 10, currentXK2B, prtMaxY - 10, " ", 250)

XPrt.DrawWidth = 1

XPrt.Line (prtMinX, prtMinY)-(prtMinX, prtMaxY)
XPrt.Line (prtMaxX, prtMinY)-(prtMaxX, prtMaxY)


'---------------------------------------------------------

XPrt.CurrentY = prtMinY + (prtHeaderHeight - XPrt.TextHeight("X")) / 2

XPrt.CurrentX = 1100: XPrt.Print "Montant brut";
XPrt.CurrentX = currentXDevise: XPrt.Print "Devise";
XPrt.CurrentX = currentXCorrespondant: XPrt.Print "Correspondant";
XPrt.CurrentX = currentXSender: XPrt.Print "Emetteur";
XPrt.CurrentX = currentXReceiver: XPrt.Print "Destinataire";
XPrt.CurrentX = currentXRéférence: XPrt.Print "Référence";
XPrt.CurrentX = 8500: XPrt.Print "Date compta";
XPrt.CurrentX = 9800: XPrt.Print "Commissions FRF";
XPrt.FontBold = False
XPrt.CurrentY = prtMinY + prtHeaderHeight - XPrt.TextHeight("X")

End Sub

'---------------------------------------------------------
 Public Sub prtOpStatX(Msg As String)
'---------------------------------------------------------
Dim K As Integer, K1 As Integer, K2 As Integer
Dim X As String
On Error GoTo prtError

K1 = Val(mId$(Msg, 1, 6))
K2 = Val(mId$(Msg, 7, 6))

Set XPrt = Printer
frmElpPrt.Show vbModeless

prtOrientation = vbPRORPortrait
prtTitleText = "Opérations de tranfert"
prtPgmName = "prtOpStat"
prtTitleUsr = usrName

prtLineNb = 1
prtlineHeight = 250
prtHeaderHeight = 300

frmElpPrt.prtStdInit
totalK_Init
prtOpStat_Form

mTrame = IIf(paramOpStat.prtDétail, 225, 235)
If Not paramOpStat.prtK1 Then mTrame = mTrame + 10
If Not paramOpStat.prtK2 Then mTrame = mTrame + 10

totalZ.Nb = 0: totalZ.Brut = 0: totalZ.Com = 0: totalZ.K1 = "": totalZ.K2 = ""
totalX = totalZ: totalK(0) = totalZ: totalK(1) = totalZ: totalK(2) = totalZ: totalK(3) = totalZ

If Trim(paramOpStat.sortK1) = "Devise" Then
    blndeviseK1 = True: blndeviseK2 = True
Else
    blndeviseK1 = False
    If Trim(paramOpStat.sortK2) = "Devise" Then
        blndeviseK2 = True
    Else
        blndeviseK2 = False
    End If
End If

recOpStat.Method = "MoveFirst   "
recOpStat.Err = tableOpStat_Read(recOpStat)
totalX_Init
totalK(2) = totalX
totalK(1) = totalX

Do While recOpStat.Err = 0
    totalX_Init
    If totalX.K1 <> totalK(1).K1 Then
        totalK_Print 2
        totalK_Print 1
    End If
    If totalX.K2 <> totalK(2).K2 Then
        totalK_Print 2
    End If
    
    If paramOpStat.prtDétail Then prtOpStat_Line
    
    totalK(2).Nb = totalK(2).Nb + 1
    If blndeviseK2 Then totalK(2).Brut = totalK(2).Brut + recOpStat.Brut
    totalK(2).Com = totalK(2).Com + recOpStat.ComMontantFRF

    DoEvents: If prtKillDoc Then frmElpPrt.Hide: Exit Sub
    
    recOpStat.Method = "MoveNext    "
    recOpStat.Err = tableOpStat_Read(recOpStat)

Loop

totalK_Print 2
totalK_Print 1
totalK(0).Brut = 0
totalK_Print 0
prtOpStat_Trait
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
Public Sub prtOpStat_Line()
'---------------------------------------------------------
Dim X As String, K As Integer

If XPrt.CurrentY + prtlineHeight * 2 > prtMaxY Then
   prtOpStat_Trait
   frmElpPrt.prtNewPage
   prtOpStat_Form
End If
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
NbImprimé = NbImprimé + 1
If NbImprimé = 6 Then
    XPrt.CurrentY = XPrt.CurrentY + 100
    NbImprimé = 1
End If


XPrt.CurrentX = currentXDevise: XPrt.Print recOpStat.Devise;

X = num_Display(recOpStat.Brut, 15, 2, Lx, X, "0")
XPrt.CurrentX = 2000 - XPrt.TextWidth(X)
XPrt.Print X;

XPrt.CurrentX = currentXCorrespondant: XPrt.Print recOpStat.BicIdCorrespondant;
XPrt.CurrentX = currentXSender: XPrt.Print recOpStat.BicIdSender;
XPrt.CurrentX = currentXReceiver: XPrt.Print recOpStat.BicIdReceiver;
XPrt.CurrentX = currentXRéférence: XPrt.Print recOpStat.Référence;

Select Case Trim(recOpStat.xNature)
    Case constàValider: XPrt.Print " -v";
    Case constàCompta: XPrt.Print " -c";
End Select

XPrt.CurrentX = 8500: XPrt.Print dateImp(recOpStat.xAMJ);

X = num_Display(recOpStat.ComMontantFRF, 15, 2, Lx, X, "0")
XPrt.CurrentX = 11000 - XPrt.TextWidth(X)
XPrt.Print X;

End Sub
'---------------------------------------------------------
Public Sub totalK_Print(I As Integer)
'---------------------------------------------------------
Dim X As String, K As Integer, xBox As String, iTrame As Integer
Dim prtTotal As Boolean
    
prtTotal = True
Select Case I
    Case 2: prtTotal = paramOpStat.prtK2
    Case 1: prtTotal = paramOpStat.prtK1
End Select

If prtTotal Then
    If XPrt.CurrentY + prtlineHeight * 2 > prtMaxY Then
       prtOpStat_Trait
       frmElpPrt.prtNewPage
       prtOpStat_Form
    End If
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight

    XPrt.FontBold = True
    iTrame = mTrame + I * 10
    Call frmElpPrt.prtTrame(prtMinX + 20, XPrt.CurrentY - 30, prtMaxX - 20, XPrt.CurrentY + prtlineHeight - 30, " ", iTrame)

    X = num_Display(totalK(I).Brut, 15, 2, Lx, X, " ")
    XPrt.CurrentX = 2000 - XPrt.TextWidth(X)
    XPrt.Print X;
    
    If I > 1 Then XPrt.CurrentX = currentXK2: XPrt.Print totalK(I).K2;
    If I > 0 Then XPrt.CurrentX = currentXK1: XPrt.Print totalK(I).K1;
    
    X = num_Display(totalK(I).Nb, 8, 0, Lx, X, "0")
    XPrt.CurrentX = 8400 - XPrt.TextWidth(X)
    XPrt.Print X;
    'If I = 2 Then
        XPrt.Print " dossier";
        If totalK(I).Nb > 1 Then XPrt.Print "s";
    'End If
    
    X = num_Display(totalK(I).Com, 15, 2, Lx, X, "0")
    XPrt.CurrentX = 11000 - XPrt.TextWidth(X)
    XPrt.Print X;
    XPrt.FontBold = False
    
    NbImprimé = 0
End If

K = I - 1
If K >= 0 Then
    totalK(K).Nb = totalK(K).Nb + totalK(I).Nb
    If blndeviseK1 Then totalK(K).Brut = totalK(K).Brut + totalK(I).Brut
    totalK(K).Com = totalK(K).Com + totalK(I).Com
End If
totalK(I) = totalX

End Sub

Public Sub totalX_Init()

Select Case Trim(paramOpStat.sortK1)
    Case "Devise": totalX.K1 = recOpStat.Devise
    Case "Correspondant": totalX.K1 = recOpStat.BicIdCorrespondant
    Case "Sender": totalX.K1 = recOpStat.BicIdSender
    Case "Receiver": totalX.K1 = recOpStat.BicIdReceiver
    Case "Référence": totalX.K1 = recOpStat.Référence
End Select
Select Case Trim(paramOpStat.sortK2)
    Case "Devise": totalX.K2 = recOpStat.Devise
    Case "Correspondant": totalX.K2 = recOpStat.BicIdCorrespondant
    Case "Sender": totalX.K2 = recOpStat.BicIdSender
    Case "Receiver": totalX.K2 = recOpStat.BicIdReceiver
    Case "Référence": totalX.K2 = recOpStat.Référence
End Select

End Sub
Public Sub totalK_Init()

currentXDevise = 2300
currentXCorrespondant = 3000
currentXSender = 4500
currentXReceiver = 6000
currentXRéférence = 7500

Select Case Trim(paramOpStat.sortK1)
    Case "Devise": currentXK1 = currentXDevise: currentXK1B = currentXDevise + 700
    Case "Correspondant": currentXK1 = currentXCorrespondant: currentXK1B = currentXCorrespondant + 1500
    Case "Sender": currentXK1 = currentXSender: currentXK1B = currentXSender + 1500
    Case "Receiver": currentXK1 = currentXReceiver: currentXK1B = currentXReceiver + 1500
    Case "Référence": currentXK1 = currentXRéférence: currentXK1B = currentXRéférence + 1000
End Select
Select Case Trim(paramOpStat.sortK2)
    Case "Devise": currentXK2 = currentXDevise: currentXK2B = currentXDevise + 700
    Case "Correspondant": currentXK2 = currentXCorrespondant: currentXK2B = currentXCorrespondant + 1500
    Case "Sender": currentXK2 = currentXSender: currentXK2B = currentXSender + 1500
    Case "Receiver": currentXK2 = currentXReceiver: currentXK2B = currentXReceiver + 1500
    Case "Référence": currentXK2 = currentXRéférence: currentXK2B = currentXRéférence + 1000
End Select

End Sub


Public Sub prtOpStat_Trait()
XPrt.Line (currentXCorrespondant - 50, prtMinY)-(currentXCorrespondant - 50, prtMaxY)
XPrt.Line (currentXRéférence - 50, prtMinY)-(currentXRéférence - 50, prtMaxY)

End Sub
