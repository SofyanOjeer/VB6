Attribute VB_Name = "prtBIA_Balance"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim X As String, I As Integer, Height8_6 As Integer

Dim blnPage As Boolean

Dim xYbase As typeYBase
Dim curX As Currency, curCumul_Db As Currency, curCumul_Cr As Currency
Dim curClient_Db As Currency, curClient_Cr As Currency, nbClient_Line As Long
Dim curListe_Db As Currency, curListe_Cr As Currency, nbListe_Line As Long
Dim curW_Db As Currency, curW_Cr As Currency
Dim IbmAmjMin As String, IbmAmjMax As String
Dim meYBIACPT0 As typeYBIACPT0, prevYBIACPT0 As typeYBIACPT0

Dim blnCompte As Boolean
Dim prtY As Integer
Dim meCV1 As typeCV, meCV2 As typeCV

Dim blnSoldeZ As Boolean
Dim blnClient_Line As Boolean
Public Sub prtBIA_Balance_Monitor(lFct As String, lAmjMin As String, lstW As ListBox)
Dim wIndex As Long
Dim mFct1 As String

mFct1 = mId$(lFct, 1, 1)

IbmAmjMin = dateIBM(lAmjMin)
meCV1.DeviseN = 0
meCV1.Montant = 0
meCV1.OpéAmj = YBIATAB0_DATE_CPT_J
meCV2.OpéAmj = YBIATAB0_DATE_CPT_J
blnMOUVEMDCO = False
blnRésidence = False: mRésidence = "-"
blnCompte = False:    curCumul_Db = 0: curCumul_Cr = 0
blnClient_Line = False
nbClient_Line = 0: nbListe_Line = 0
curClient_Db = 0: curClient_Cr = 0
curListe_Db = 0: curListe_Cr = 0
curW_Db = 0: curW_Cr = 0
recYBIACPT0_Init prevYBIACPT0

Select Case mFct1
    Case "B": prtTitleText = "Balance"
                If mId$(lFct, 2, 1) = "T" Then blnClient_Line = True
End Select

If mId$(lFct, 6, 1) = "Z" Then
    blnSoldeZ = True
Else
    blnSoldeZ = False
End If

prtTitleText = prtTitleText & " au " & dateImp10(lAmjMin)
prtFontName = prtFontName_Arial
prtBIA_Balance_Open
prtHeaderHeight = 300
prtBIA_Balance_Form
XPrt.CurrentY = XPrt.CurrentY - prtlineHeight

XPrt.FontSize = 8
For I = 1 To lstW.ListCount - 1
    
    lstW.ListIndex = I
    X = mId$(lstW.Text, 14, 20)
    V = srvYBIACPT0_Import_Read(X, meYBIACPT0)
    If IsNull(V) Then
        meYBIACPT0 = larrYBIACPT0(wIndex)
        meCV1.DeviseIso = meYBIACPT0.COMPTEDEV
        
        Select Case mFct1
            Case Else: prtBIA_Balance_B_Line
         End Select
         
         prevYBIACPT0 = meYBIACPT0
    End If
Next I


If mFct1 = "B" Then
    prtBIA_Balance_B_Rupture
    If curListe_Db <> curW_Db Or curListe_Cr <> curW_Cr Then
        prtBIA_Balance_NewLine
        XPrt.FontSize = 12: XPrt.FontBold = True
        frmElpPrt.prtCentré prtMedX, "ERREUR TOTALISATION"
        XPrt.FontSize = 8: XPrt.FontBold = False
    End If
    XPrt.DrawWidth = 10
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.Line (prtMinX + 12000, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY)
    XPrt.CurrentY = XPrt.CurrentY - prtlineHeight
    prevYBIACPT0.CLIENACLI = ""
    prevYBIACPT0.CLIENASIG = ""
    nbClient_Line = nbListe_Line
    curClient_Db = curListe_Db
    curClient_Cr = curListe_Cr
    prtBIA_Balance_B_Rupture
Else

End If

prtBIA_Balance_Close

End Sub

'---------------------------------------------------------
Public Sub prtBIA_Balance_Form()
'---------------------------------------------------------
Dim X As String

XPrt.DrawWidth = 1
XPrt.FontSize = 7: XPrt.FontBold = True

Call frmElpPrt.prtTrame(prtMinX, prtMinY, prtMaxX, prtMinY + prtHeaderHeight, "B", 240)
XPrt.Line (prtMinX + 12000, prtMinY)-(prtMinX + 12000, prtMaxY)
XPrt.Line (prtMinX + 7500, prtMinY)-(prtMinX + 7500, prtMaxY)
XPrt.Line (prtMaxX, prtMinY)-(prtMaxX, prtMaxY)
'---------------------------------------------------------
XPrt.CurrentY = prtMinY + 50

XPrt.CurrentX = prtMinX + 100: XPrt.Print " ";
XPrt.CurrentX = prtMinX + 400: XPrt.Print "Compte ";
XPrt.CurrentX = prtMinX + 2000: XPrt.Print "Intitulé";
'XPrt.CurrentX = prtMinX + 10500: XPrt.Print "Devise";
XPrt.CurrentX = prtMinX + 8900: XPrt.Print "Débit";
XPrt.CurrentX = prtMinX + 10900: XPrt.Print "Crédit";
XPrt.CurrentX = prtMinX + 13100: XPrt.Print "Débit";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 6
XPrt.FontBold = True
XPrt.CurrentX = prtMinX + 14100: XPrt.Print "/ EUR /";
XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 8

XPrt.CurrentX = prtMinX + 15100: XPrt.Print "Crédit";

XPrt.CurrentY = prtMinY + prtHeaderHeight + 100
XPrt.FontBold = False

End Sub


Public Sub prtBIA_Balance_Close()
On Error GoTo prtError


frmElpPrt.prtEndDoc
frmElpPrt.Hide
Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub



Public Sub prtBIA_Balance_Open()
On Error GoTo prtError

Set XPrt = Printer
frmElpPrt.Show vbModeless

Height8_6 = frmElpPrt.prtHeightDelta(8, 6)

blnFiligrane = False
prtOrientation = vbPRORLandscape '
prtPgmName = "prtBIA_Balance"
prtTitleUsr = usrName

prtLineNb = 1
prtlineHeight = 250
prtHeaderHeight = 50 ' 100


prtFormType = ""
frmElpPrt.prtStdInit

Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub





Public Sub prtBIA_Balance_B(lFct As String, lcurX As Currency)

prtBIA_Balance_NewLine
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 6

Select Case lFct
    Case "R0"
     '   Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY, prtMinX + 7480, XPrt.CurrentY + prtlineHeight - 50, " ", 240)
    '    Call frmElpPrt.prtTrame(prtMinX + 7520, XPrt.CurrentY, prtMinX + 11980, XPrt.CurrentY- 50, " ", 240)
    '    Call frmElpPrt.prtTrame(prtMinX + 12020, XPrt.CurrentY , prtMaxX - 20, XPrt.CurrentY- 50, " ", 240)
        XPrt.FontBold = True: XPrt.ForeColor = prtForeColor_Header

        prtBIA_Balance_Montant lcurX
        XPrt.FontBold = False: XPrt.ForeColor = prtForeColor
    Case "R1"
        prtY = XPrt.CurrentY + prtlineHeight
       ' Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY, prtMinX + 7480, prtY - 50, " ", 240)
        Call frmElpPrt.prtTrame(prtMinX + 7520, XPrt.CurrentY, prtMinX + 11980, prtY - 50, " ", 240)
       ' Call frmElpPrt.prtTrame(prtMinX + 12020, XPrt.CurrentY, prtMaxX - 20, prtY - 50, " ", 240)
        XPrt.FontBold = True
        prtBIA_Balance_Montant lcurX
        XPrt.FontBold = False
    Case "C1"
        prtBIA_Balance_Montant lcurX
    Case Else
        prtBIA_Balance_Montant lcurX
End Select

XPrt.CurrentX = prtMinX + 100: XPrt.Print meYBIACPT0.PLANCOPRO;
XPrt.FontBold = True

''Dim mRib_IbanE As String, mRib_Clé As String
''mRib_Clé = Format$(RibClé(strSocBdfE, strSocBdfG, Trim(meYBIACPT0.COMPTECOM), mRib_IbanE), "00")

XPrt.CurrentX = prtMinX + 400: XPrt.Print meYBIACPT0.COMPTECOM;  '''"& "    " & mRib_Clé;

XPrt.CurrentX = prtMinX + 2000: XPrt.Print meYBIACPT0.COMPTEINT;
XPrt.FontBold = False
If lFct = "B" Then
    XPrt.CurrentX = prtMinX + 6800: XPrt.Print dateIBM10(meYBIACPT0.SOLDEDMO, True);
End If
XPrt.CurrentX = prtMinX + 11600: XPrt.Print meYBIACPT0.COMPTEDEV;

XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 8

End Sub

Public Sub prtBIA_Balance_Montant(lcurX As Currency)
Dim X As String
X = Format$(Abs(lcurX), "### ### ### ### ##0.00")
If lcurX > 0 Then
    XPrt.CurrentX = prtMinX + 9400 - XPrt.TextWidth(X)
Else
    XPrt.CurrentX = prtMinX + 11400 - XPrt.TextWidth(X)
End If
XPrt.Print X;

If meCV1.DeviseIso <> "EUR" Then
    meCV1.Montant = lcurX
       
    Call CV_Calc(meCV1, meCV2)
Else
    meCV2.Montant = lcurX
End If

X = Format$(Abs(meCV2.Montant), "### ### ### ### ##0.00")
If meCV2.Montant > 0 Then
    XPrt.CurrentX = prtMaxX - 2100 - XPrt.TextWidth(X)
Else
    XPrt.CurrentX = prtMaxX - 100 - XPrt.TextWidth(X)
End If
XPrt.Print X;

End Sub

Public Sub prtBIA_Balance_Montant_Cumul(lcurDB As Currency, lcurCR As Currency)
Dim X As String
X = Format$(Abs(lcurDB), "### ### ### ### ##0.00")
XPrt.CurrentX = prtMaxX - 2100 - XPrt.TextWidth(X)
XPrt.Print X;
X = Format$(Abs(lcurCR), "### ### ### ### ##0.00")
XPrt.CurrentX = prtMaxX - 100 - XPrt.TextWidth(X)
XPrt.Print X;

End Sub


Public Sub prtBIA_Balance_NewLine()
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
If XPrt.CurrentY + 300 > prtMaxY Then
    frmElpPrt.prtNewPage
    prtBIA_Balance_Form
End If

End Sub

Public Sub prtBIA_Balance_B_Rupture()

nbListe_Line = nbListe_Line + 1

If blnClient_Line Then
    prtBIA_Balance_NewLine
    XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 6: XPrt.FontBold = True

    XPrt.CurrentX = prtMinX + 6000: XPrt.Print prevYBIACPT0.CLIENACLI & " " & prevYBIACPT0.CLIENASIG;

    Call frmElpPrt.prtTrame(prtMinX + 12020, XPrt.CurrentY, prtMaxX - 20, XPrt.CurrentY + prtlineHeight - 50, " ", 240)
    If curClient_Db <> 0 Then
        X = Format$(Abs(curClient_Db), "### ### ### ### ##0.00")
        XPrt.CurrentX = prtMaxX - 2100 - XPrt.TextWidth(X)
        XPrt.Print X;
    End If
    If curClient_Cr <> 0 Then
        X = Format$(Abs(curClient_Cr), "### ### ### ### ##0.00")
        XPrt.CurrentX = prtMaxX - 100 - XPrt.TextWidth(X)
        XPrt.Print X;
    End If
    XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 8:: XPrt.FontBold = False

End If

curW_Db = curW_Db + curClient_Db
curW_Cr = curW_Cr + curClient_Cr

nbClient_Line = 0
curClient_Db = 0: curClient_Cr = 0

End Sub
Public Sub prtBIA_Balance_B_Line()

If prevYBIACPT0.CLIENACLI <> meYBIACPT0.CLIENACLI Then
    If Trim(prevYBIACPT0.COMPTECOM) <> "" Then
        prtBIA_Balance_B_Rupture
    End If
End If

prtBIA_Balance_B "B", meYBIACPT0.SOLDECEN

nbClient_Line = nbClient_Line + 1

If meCV2.Montant > 0 Then
    curClient_Db = curClient_Db + meCV2.Montant
    curListe_Db = curListe_Db + meCV2.Montant
Else
    curClient_Cr = curClient_Cr + meCV2.Montant
    curListe_Cr = curListe_Cr + meCV2.Montant
End If

End Sub


