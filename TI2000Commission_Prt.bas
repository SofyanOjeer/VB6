Attribute VB_Name = "prtTI2000Commission"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim X As String, I As Integer, Height8_6 As Integer

Dim wCDComD As typeCDComD, xCDComD As typeCDComD
Dim xAmjSituation As String, xAmjMin As String, xAmjMax  As String

Dim Nb1 As Integer, Nb2 As Integer
Dim curComPerçue As Currency, curComDue As Currency, curComAnte As Currency, curComProrata As Currency, curComSituation As Currency
Dim curComSolde As Currency, curUtilisation As Currency
Dim curX As Currency, curS36 As Currency

Dim mCDDossier As typeCDDossier
Dim mAmjOuvertureMin As String * 8
Dim arrCDComD_KS As Integer

Dim tComEur As Currency, tComDev As Currency, tComDevX As String
Dim tMvtEngagement As Currency, tMvtUtilisation As Currency
Dim curMvtEngagement As Currency, curMvtUtilisation As Currency
Dim sMvtEngagement As Currency, sMvtUtilisation As Currency

Public selCDComD_Type As String, selCDComD_Devise As String, selprtStatut As String
Public mCDDossier_AMJSituation As String
Dim mRupture As String, mType As String

Dim blnExportTICom As Boolean
'---------------------------------------------------------
 Public Sub prtTI2000Commission_Monitor(Msg As String)
'---------------------------------------------------------
Dim K As Long, K1 As Long, K2 As Long
Dim X As String

On Error GoTo prtError

Set XPrt = Printer
Height8_6 = frmElpPrt.prtHeightDelta(8, 6)
K1 = Val(mId$(Msg, 1, 6))
K2 = Val(mId$(Msg, 7, 6))

xAmjSituation = dateImp(paramTIDB2_AMJSituation)
mAmjOuvertureMin = dateElp("Jour", -15, paramTIDB2_AMJSituation)

prtTitleText = "Crédits documentaires : Commissions au " & xAmjSituation & "  " & selprtStatut

prtLineNb = 1

frmElpPrt.Show vbModeless


prtOrientation = vbPRORLandscape
prtPgmName = "prtTI2000Commission"
prtTitleUsr = usrName

prtlineHeight = 250
prtHeaderHeight = 300

frmElpPrt.prtStdInit

mdbMvtP0.tableMvtP0_Open
recMvtP0_Init recMvtp0
recMvtp0.Method = "MoveFirst"
intReturn = tableMvtP0_Read(recMvtp0)


mdbCDDossier.tableCDDossier_Open
recCDDossier_Init mCDDossier

mdbCDComD.tableCDComD_Open
recCDComD_Init xCDComD
xCDComD.Method = "Seek="

prtTI2000Commission_Form

''For I = 1 To 26
''    XPrt.CurrentY = XPrt.CurrentY + 400
''XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY)

''Next I

'''frmElpPrt.prtEndDoc
''GoTo prtError

'dbCDComD_ReadE xCDComD
'wCDComD = xCDComD
'prtTI2000Commission_Line_Init
mCDDossier_AMJSituation = "CAN"
tComEur = 0: tComDev = 0
tMvtEngagement = 0: tMvtUtilisation = 0
sMvtEngagement = 0: sMvtUtilisation = 0
tComDevX = mId$(recMvtp0.Id, 3, 3)
mRupture = mId$(recMvtp0.Id, 1, 5)
mType = mId$(recMvtp0.Id, 1, 2)

''''blnExportTICom = True



If blnExportTICom Then Open "C:\Temp\TICom.txt" For Output As #1

Do
    xCDComD.Type = mId$(recMvtp0.Id, 1, 2)
    xCDComD.Dossier = CLng(mId$(recMvtp0.Id, 6, 11))
    xCDComD.AmjD = mId$(recMvtp0.Id, 17, 8)
    

    dbCDComD_ReadE xCDComD
    If wCDComD.Dossier <> xCDComD.Dossier Then
        If mCDDossier_AMJSituation = selprtStatut Then
            'curX = curS36 - curMvtEngagement + curMvtUtilisation
            'If curX <> 0 Or curS36 <> 0 Then
                If blnExportTICom Then
                    prtTI2000Commission_Export
                Else
                    prtTI2000Commission_Line
                End If
            'End If
        End If
        
        wCDComD = xCDComD
        prtTI2000Commission_Line_Init
       ''' If mCDDossier.AMJValidité > "20001200" Then mCDDossier.AMJSituation = "x"
        mCDDossier_AMJSituation = Trim(mCDDossier.AMJSituation)
        
        If mRupture <> mId$(recMvtp0.Id, 1, 5) Then
            prtTI2000Commission_Total
            frmElpPrt.prtNewPage
            prtTI2000Commission_Form

            mRupture = mId$(recMvtp0.Id, 1, 5)
            tComDevX = mId$(recMvtp0.Id, 3, 3)
            mType = mId$(recMvtp0.Id, 1, 2)
        End If

   End If
    
    ''wCDComD.AmjF = xCDComD.AmjF
   
   ' If xCDComD.AmjD < paramTIDB2_AMJSituation Then
        arrCDComD_Nb = arrCDComD_Nb + 1
        curUtilisation = curUtilisation + xCDComD.MvtUtilisé
        If mCDDossier.AMJOuverture > mAmjOuvertureMin Then xCDComD.CommissionD = 0
        wCDComD.CommissionTaux = xCDComD.CommissionTaux
        xAmjMin = dateImp(xCDComD.AmjD)
        xAmjMax = dateImp(xCDComD.AmjF)
        If xCDComD.CommissionPAmj <= paramTIDB2_AMJSituation Then curComPerçue = curComPerçue + xCDComD.CommissionP
        If xCDComD.AmjF <= paramTIDB2_AMJSituation Then
            curComAnte = curComAnte + xCDComD.CommissionD
        Else
            arrCDComD_KS = arrCDComD_Nb
            Nb1 = DateDiff("d", xAmjMin, xAmjSituation)
            Nb2 = DateDiff("d", xAmjMin, xAmjMax)
            curComProrata = curComProrata + Round(xCDComD.CommissionD * Nb1 / Nb2, 2)
            curComSituation = xCDComD.CommissionD
        End If
        curMvtUtilisation = curMvtUtilisation + xCDComD.MvtUtilisé
        curMvtEngagement = curMvtEngagement + xCDComD.MvtEngagement
        arrCDComD(arrCDComD_Nb) = xCDComD

    'End If

    recMvtp0.Method = "MoveNext"
    intReturn = tableMvtP0_Read(recMvtp0)

Loop While intReturn = 0
If blnExportTICom Then prtTI2000Commission_Export: Close
        
If mCDDossier_AMJSituation = selprtStatut Then prtTI2000Commission_Line
prtTI2000Commission_Total

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
Public Sub prtTI2000Commission_Form()
'---------------------------------------------------------
Dim X As String
XPrt.FontSize = 6

XPrt.FontBold = True
XPrt.DrawWidth = 3
XPrt.CurrentY = prtMinY
prtCurrentY = prtMinY + prtlineHeight
Call frmElpPrt.prtTrame(prtMinX, prtMinY, prtMaxX, prtCurrentY, "B", 250)
'XPrt.Line (prtMinX, prtCurrentY)-(prtMaxX, prtCurrentY)

XPrt.DrawWidth = 1
Call frmElpPrt.prtTrame(prtMinX + 4620, prtCurrentY + 20, prtMinX + 7480, prtMaxY - 20, " ", 250)

XPrt.Line (prtMinX + 2500, prtMinY)-(prtMinX + 2500, prtMaxY)
XPrt.Line (prtMinX + 4600, prtMinY)-(prtMinX + 4600, prtMaxY)
XPrt.Line (prtMinX + 7500, prtMinY)-(prtMinX + 7500, prtMaxY)

'---------------------------------------------------------

X = "Dossier"
XPrt.CurrentY = prtMinY + (prtHeaderHeight - XPrt.TextHeight(X)) / 2
XPrt.CurrentX = prtMinX: XPrt.Print " Dossier";
XPrt.CurrentX = prtMinX + 700: XPrt.Print "    Ouverture";
XPrt.CurrentX = prtMinX + 1600: XPrt.Print "    Validité";
XPrt.CurrentX = prtMinX + 3500: XPrt.Print "-EUR- Com perçues";
XPrt.CurrentX = prtMinX + 2500: XPrt.Print "    Com dues";
XPrt.CurrentX = prtMinX + 5300: XPrt.Print "Commissions non perçues";
XPrt.CurrentX = prtMinX + 8000: XPrt.Print "  Période";
XPrt.CurrentX = prtMinX + 10000: XPrt.Print "Base";
XPrt.CurrentX = prtMinX + 11000: XPrt.Print " Taux";
XPrt.CurrentX = prtMinX + 11500: XPrt.Print "  Com période";
XPrt.CurrentX = prtMinX + 13100: XPrt.Print "Prorata";
XPrt.CurrentX = prtMinX + 13900: XPrt.Print "Engagement";
XPrt.CurrentX = prtMinX + 15100: XPrt.Print "Utilisation";


XPrt.CurrentY = prtMinY + prtHeaderHeight - XPrt.TextHeight("X")

XPrt.FontSize = 6

End Sub

'---------------------------------------------------------
Public Sub prtTI2000Commission_Line()
'---------------------------------------------------------

If XPrt.CurrentY + prtlineHeight * (arrCDComD_Nb + 1.5) > prtMaxY Then
    frmElpPrt.prtNewPage
    prtTI2000Commission_Form
End If

'jpl 20010206 If tComDevX <> wCDComD.Devise Then
'jpl 20010206    prtTI2000Commission_Total
'jpl 20010206    tComDevX = wCDComD.Devise
'jpl 20010206End If

XPrt.FontBold = False
'_______________________________________________________________ligne 1-



XPrt.CurrentY = XPrt.CurrentY + prtlineHeight

tMvtEngagement = tMvtEngagement + curMvtEngagement
tMvtUtilisation = tMvtUtilisation + curMvtUtilisation


XPrt.FontBold = True
XPrt.CurrentX = prtMinX: XPrt.Print wCDComD.Dossier;
XPrt.FontBold = False
XPrt.Print " " & wCDComD.Type & " " & mCDDossier_AMJSituation;
XPrt.CurrentX = prtMinX + 800: XPrt.Print dateImp(wCDComD.AmjD);
XPrt.CurrentX = prtMinX + 1650: XPrt.Print dateImp(mCDDossier.AMJValidité);

If curComPerçue <> 0 Then
    X = Format$(curComPerçue, "## ### ### ### ### ##0.00")
    XPrt.CurrentX = prtMinX + 4500 - XPrt.TextWidth(X)
    XPrt.Print X;
End If

    X = Format$(curComAnte + curComProrata, "## ### ### ### ### ##0.00")
    XPrt.CurrentX = prtMinX + 3300 - XPrt.TextWidth(X)
    XPrt.Print X;
    
curComSolde = curComAnte + curComProrata - curComPerçue

If curUtilisation <> 0 Then
    If curComSolde > 10 Then XPrt.CurrentX = prtMinX + 5800: XPrt.Print "*";
Else
    If curComSolde > 10 Then
        XPrt.FontBold = True
        tComEur = tComEur + curComSolde
        X = Format$(curComSolde, "## ### ### ### ### ##0.00")
        XPrt.CurrentX = prtMinX + 5800 - XPrt.TextWidth(X)
        XPrt.Print X;
        curComSolde = Round(curComSolde * wCDComD.CoursEur, 2)
        tComDev = tComDev + curComSolde
        X = Format$(curComSolde, "## ### ### ### ### ##0.00")
        XPrt.CurrentX = prtMinX + 7100 - XPrt.TextWidth(X)
        XPrt.Print X & "  " & wCDComD.Devise; ;
    
        XPrt.FontBold = False
        sMvtEngagement = sMvtEngagement + curMvtEngagement
        sMvtUtilisation = sMvtUtilisation + curMvtUtilisation
   End If
End If


prtTI2000Commission_Line_Détail

End Sub

'---------------------------------------------------------
Public Sub prtTI2000Commission_Export()
'---------------------------------------------------------
Dim X1 As String, X2 As String, X3 As String, curX As Currency

'_______________________________________________________________ligne 1-


X1 = IIf(curMvtEngagement < 0, "-", "+")
X1 = X1 & Format$(Abs(curMvtEngagement), "00000000000000000000000000.00")
X2 = IIf(curMvtUtilisation < 0, "-", "+")
X2 = X2 & Format$(Abs(curMvtUtilisation), "00000000000000000000000000.00")
curX = curMvtEngagement - curMvtUtilisation
X3 = IIf(curX < 0, "-", "+")
X3 = X3 & Format$(Abs(curX), "00000000000000000000000000.00")

tMvtEngagement = tMvtEngagement + curMvtEngagement
tMvtUtilisation = tMvtUtilisation + curMvtUtilisation


Print #1, wCDComD.Dossier; ";"; wCDComD.Type; ";"; wCDComD.Devise; ";"; wCDComD.AmjD; ";"; mCDDossier.AMJValidité; ";"; X1; ";"; X2; ";"; X3

'        sMvtEngagement = sMvtEngagement + curMvtEngagement
'        sMvtUtilisation = sMvtUtilisation + curMvtUtilisation
End Sub



'---------------------------------------------------------
Public Sub prtTI2000Commission_Line_Détail()
'---------------------------------------------------------
Dim I As Integer, kTrame As Integer

If XPrt.CurrentY + prtlineHeight * 1.5 > prtMaxY Then
    frmElpPrt.prtNewPage
    XPrt.CurrentY = prtMinX + prtlineHeight * 3
    prtTI2000Commission_Form
End If

XPrt.CurrentY = XPrt.CurrentY - prtlineHeight
'_______________________________________________________________ligne 1-
For I = 1 To arrCDComD_Nb

    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    If I = 1 Then XPrt.CurrentX = prtMinX + 7550: XPrt.Print dateImp(arrCDComD(I).AmjD);
    XPrt.CurrentX = prtMinX + 8450: XPrt.Print dateImp(arrCDComD(I).AmjF);

    
    X = Format$(arrCDComD(I).MontantBase, "## ### ### ### ### ##0.00")
    XPrt.CurrentX = prtMinX + 10500 - XPrt.TextWidth(X)
    XPrt.Print X;
    XPrt.CurrentX = prtMinX + 10600: XPrt.Print wCDComD.Devise;
    X = Format$(arrCDComD(I).CommissionTaux, "##0.00")
    XPrt.CurrentX = prtMinX + 11300 - XPrt.TextWidth(X)
    XPrt.Print X;
     
    X = Format$(arrCDComD(I).CommissionD, "## ### ### ### ### ##0.00")
    XPrt.CurrentX = prtMinX + 12300 - XPrt.TextWidth(X)
    XPrt.Print X;
   
    If I = arrCDComD_KS Then
        If curComProrata <> 0 Then
            XPrt.CurrentX = prtMinX + 12500: XPrt.Print Nb1 & "  / " & Nb2;
            X = Format$(curComProrata, "## ### ### ### ### ##0.00")
            XPrt.CurrentX = prtMinX + 13600 - XPrt.TextWidth(X)
            XPrt.Print X;
        End If
    End If
    
    If arrCDComD(I).MvtEngagement <> 0 Then
        X = Format$(arrCDComD(I).MvtEngagement, "## ### ### ### ### ##0.00")
        XPrt.CurrentX = prtMinX + 14700 - XPrt.TextWidth(X)
        XPrt.Print X;
    End If
    X = Format$(arrCDComD(I).MvtUtilisé, "## ### ### ### ### ##0.00")
    XPrt.CurrentX = prtMinX + 15800 - XPrt.TextWidth(X)
    XPrt.Print X;
    
Next I

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight

'curX = mCDDossier.S36Engagement - mCDDossier.S36Utilisé - curMvtEngagement + curMvtUtilisation
curX = curS36 - curMvtEngagement + curMvtUtilisation
If curX <> 0 Then

    Call frmElpPrt.prtTrame(prtMinX + 12500, XPrt.CurrentY - 50, prtMinX + 15800, XPrt.CurrentY + prtlineHeight - 50, " ", 230)
    XPrt.CurrentX = prtMinX + 12500: XPrt.Print "Contrôle S36 / TI :";
    X = Format$(curS36, "## ### ### ### ### ##0.00")
    XPrt.CurrentX = prtMinX + 14700 - XPrt.TextWidth(X)
    XPrt.Print X;
Else
    Call frmElpPrt.prtTrame(prtMinX + 14700, XPrt.CurrentY - 50, prtMinX + 15800, XPrt.CurrentY + prtlineHeight - 50, " ", 245)
End If
XPrt.FontBold = True
X = Format$(curMvtEngagement - curMvtUtilisation, "## ### ### ### ### ##0.00")
XPrt.CurrentX = prtMinX + 15800 - XPrt.TextWidth(X)
XPrt.Print X;
XPrt.FontBold = False

prtCurrentY = XPrt.CurrentY + prtlineHeight - 20
'XPrt.Line (prtMinX + 7500, prtCurrentY)-(prtMaxX, prtCurrentY)
XPrt.Line (prtMinX, prtCurrentY)-(prtMaxX, prtCurrentY)
XPrt.CurrentY = XPrt.CurrentY - prtlineHeight + 50
End Sub






Public Sub prtTI2000Commission_Line_Init()
curMvtEngagement = 0: curMvtUtilisation = 0
curComSituation = 0: curComAnte = 0: curComPerçue = 0: curComDue = 0: curComProrata = 0
curUtilisation = 0
mCDDossier.Method = "Seek="
mCDDossier.Dossier = wCDComD.Dossier

dbCDDossier_ReadE mCDDossier
arrCDComD_Nb = 0: arrCDComD_KS = -1
Select Case mType
    Case "RC": curS36 = mCDDossier.S36RC
    Case "RE": curS36 = mCDDossier.S36RE
    Case "RI": curS36 = mCDDossier.S36RI
    Case "RA": curS36 = mCDDossier.S36RA
End Select

End Sub

Public Sub prtTI2000Commission_Total()

If Trim(tComDevX) <> "" Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    
    XPrt.FontBold = True
    X = Format$(tComEur, "## ### ### ### ### ##0.00")
    XPrt.CurrentX = prtMinX + 5800 - XPrt.TextWidth(X)
    XPrt.Print X;
    X = Format$(tComDev, "## ### ### ### ### ##0.00")
    XPrt.CurrentX = prtMinX + 7100 - XPrt.TextWidth(X)
    XPrt.Print X & "  " & tComDevX; ;
    
    X = Format$(tMvtEngagement, "## ### ### ### ### ##0.00")
    XPrt.CurrentX = prtMinX + 14700 - XPrt.TextWidth(X)
    XPrt.Print X;
    X = Format$(tMvtUtilisation, "## ### ### ### ### ##0.00")
    XPrt.CurrentX = prtMinX + 15800 - XPrt.TextWidth(X)
    XPrt.Print X;
    
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY)
    
    XPrt.CurrentX = prtMinX + 11500
    XPrt.Print "Solde des engagements : " & selprtStatut & "  " & mId$(mRupture, 1, 2);
    
    X = Format$(tMvtEngagement - tMvtUtilisation, "## ### ### ### ### ##0.00")
    XPrt.CurrentX = prtMinX + 14700 - XPrt.TextWidth(X)
    XPrt.Print X & " " & tComDevX;

    
    XPrt.FontBold = False
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY)
    
End If

tComEur = 0: tComDev = 0
tMvtEngagement = 0: tMvtUtilisation = 0
sMvtEngagement = 0: sMvtUtilisation = 0

End Sub
