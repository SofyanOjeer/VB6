Attribute VB_Name = "prtCompteCapMoy"

'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public paramCompteCapMoy_Cpt_Import As String, paramCompteCapMoy_Cpt_Export As String, paramCompteCapMoy_Mvt_Import As String

Private reccptp0 As typeCptP0
Dim X As String, I As Integer, Height8_6 As Integer

'Private recCptInfo As typeCptInfo

Dim prtLineNb As Integer

Dim X250 As String * 250, wDec, wDecCV   ' As decimal
Dim xAmjMin As String * 10, xAmjMax As String * 10
Dim mDbMt As Currency, mDbNb As Long
Dim mCrMt As Currency, mCrNb As Long
Dim mID14 As String * 14, wMt As Currency
Dim mNbj As Long

Dim totalDbSd, totalCRSd, totalDbNbj, totalCrNbj
Dim totalDbMt As Currency, totalDbNb
Dim totalCrMt As Currency, totalCrNb
Dim totalCompteNb As Long

Dim wIntitulé As String
'---------------------------------------------------------
 Public Sub prtCompteCapMoy_Monitor(Msg As String)
'---------------------------------------------------------
Dim K As Integer, K1 As Integer, K2 As Integer, Kmin As Integer, Kmax As Integer
Dim X As String, wL As Long
Dim blnOk As Boolean

On Error GoTo prtError

Set XPrt = Printer
Height8_6 = frmElpPrt.prtHeightDelta(8, 6)
K1 = Val(mId$(Msg, 1, 6))
K2 = Val(mId$(Msg, 7, 6))
Open paramCompteCapMoy_Cpt_Export For Input As #1
Line Input #1, X250
xAmjMin = mId$(X250, 17, 10)
xAmjMax = mId$(X250, 28, 10)

recElpTable_Init recElpTable
xElpTable = recElpTable
recElpTable.Method = "Seek="
recElpTable.Id = "BiaPgm"
prtTitleText = "Etat des capitaux moyens :  " & xAmjMin & "  au  " & xAmjMax

prtLineNb = 1

frmElpPrt.Show vbModeless


prtOrientation = vbPRORLandscape
prtPgmName = "prtCompteCapMoy"
prtTitleUsr = usrName

prtlineHeight = 300
prtHeaderHeight = 300

frmElpPrt.prtStdInit

'recCompteInit recCompte
'recCompte.Société = SocId$
'recCompte.Agence = SocAgence$
'recCompte.Devise = "001"
'recCompte.BiaTyp = "000"
'recCompte.BiaNum = "00"
'recCompte.Method = "SeekL1"

 recCptP0_Init reccptp0
tableCptP0_Open

prtCompteCapMoy_Form
Close
Open paramCompteCapMoy_Cpt_Export For Input As #1
wDec = CDec(0)
totalDbSd = CDec(0): totalCRSd = CDec(0)
totalDbMt = CDec(0)
totalCrMt = CDec(0)
totalDbNbj = CDec(0): totalCrNbj = CDec(0)
totalCompteNb = 0
totalDbNb = CDec(0): totalCrNb = CDec(0)
'For K = 1 To 5 'K1 To K2

Do Until EOF(1)
    Line Input #1, X250

blnOk = True


If blnOk Then
    'recCompte.Devise = mId$(X250, 1, 3)
    'recCompte.Numéro = mId$(X250, 4, 11)
'    If Not IsNull(srvCompteFind(recCompte)) Then recCompte.Intitulé = "????"
    reccptp0.Id = mId$(X250, 1, 3) & mId$(X250, 5, 11)
    reccptp0.Method = "Seek="
    If tableCptP0_Read(reccptp0) = 0 Then
        wIntitulé = mId$(reccptp0.Text, 35, 40)
    Else
        wIntitulé = "????"
    End If
   
    mNbj = CLng(mId$(X250, 39, 10))
    wDecCV = CDec(mId$(X250, 50, 30))
    wDec = CDec(mId$(X250, 81, 30))
    If wDecCV < 0 Then
        totalDbSd = totalDbSd + wDecCV / mNbj: totalDbNbj = totalDbNbj + mNbj
    Else
        totalCRSd = totalCRSd + wDecCV / mNbj: totalCrNbj = totalCrNbj + mNbj
   End If
    
    mDbNb = CLng(mId$(X250, 112, 10)): totalDbNb = totalDbNb + mDbNb
    mDbMt = CCur(mId$(X250, 123, 19)): totalDbMt = totalDbMt + mDbMt
    mCrNb = CLng(mId$(X250, 143, 10)): totalCrNb = totalCrNb + mCrNb
    mCrMt = CCur(mId$(X250, 154, 19)): totalCrMt = totalCrMt + mCrMt
    totalCompteNb = totalCompteNb + 1
    prtCompteCapMoy_Line
    DoEvents: If prtKillDoc Then frmElpPrt.Hide: Exit Sub
End If
'Next K
Loop


XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY)

XPrt.FontBold = True
XPrt.CurrentY = XPrt.CurrentY + 100
If totalDbSd Then
    wMt = CCur(totalDbSd) '' / totalDbNbj)
    X = Format$(Abs(wMt), "## ### ### ### ### ##0.00")
    XPrt.CurrentX = prtMinX + 6000 - XPrt.TextWidth(X)
    XPrt.Print X;
End If
If totalCRSd Then
    wMt = CCur(totalCRSd)   ''/ totalCrNbj)
    X = Format$(Abs(wMt), "## ### ### ### ### ##0.00")
    XPrt.CurrentX = prtMinX + 7500 - XPrt.TextWidth(X)
    XPrt.Print X;
End If

reccptp0.Id = ""
wIntitulé = "Total : " & Trim(Format$(totalCompteNb, "### ### ###")) & " comptes"
mDbNb = totalDbNb
mDbMt = totalDbMt
mCrNb = totalCrNb
mCrMt = totalCrMt
mNbj = totalDbNbj + totalCrNbj
wDec = 0
wDecCV = 0
prtCompteCapMoy_Line

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY)


Close
tableCptP0_Close
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
Public Sub prtCompteCapMoy_Form()
'---------------------------------------------------------
Dim X As String
prtCurrentY = XPrt.CurrentY
XPrt.FontSize = 8

XPrt.FontBold = True
XPrt.DrawWidth = 3

Call frmElpPrt.prtTrame(prtMinX, prtCurrentY, prtMaxX, prtCurrentY + prtlineHeight, "B", 250)

XPrt.DrawWidth = 1


XPrt.Line (prtMinX + 4550, prtMinY)-(prtMinX + 4550, prtMaxY)
XPrt.Line (prtMinX + 7550, prtMinY)-(prtMinX + 7550, prtMaxY)
XPrt.Line (prtMinX + 10900, prtMinY)-(prtMinX + 10900, prtMaxY)

'---------------------------------------------------------

XPrt.CurrentY = prtCurrentY + 50
XPrt.CurrentX = prtMinX + 100: XPrt.Print "Compte";
XPrt.CurrentX = prtMinX + 1500: XPrt.Print "Intitulé";
frmElpPrt.prtCentré prtMinX + 6000, "Solde moyen (Euro)"
frmElpPrt.prtCentré prtMinX + 9000, "Solde moyen (devise)"

XPrt.CurrentX = prtMinX + 11800: XPrt.Print "Mvt Db : Nb /Capitaux Eur";
XPrt.CurrentX = prtMinX + 13800: XPrt.Print "Mvt Cr : Nb / Capitaux Eur";


XPrt.FontBold = False
XPrt.FontSize = 6
End Sub
Public Sub prtCompteCapMoy_Line()
'------------------------------------------------------ligne 1---
Dim iReturn As Integer

If XPrt.CurrentY + prtlineHeight * 1.5 > prtMaxY Then
    frmElpPrt.prtNewPage
    prtCompteCapMoy_Form
End If

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight

XPrt.CurrentX = prtMinX + 100: XPrt.Print mId$(reccptp0.Id, 1, 3) & " " & Compte_Imp(mId$(reccptp0.Id, 4, 11));
XPrt.CurrentX = prtMinX + 1500: XPrt.Print wIntitulé;
If mNbj <> 0 Then
    wMt = CCur(wDecCV / mNbj)
    X = Format$(Abs(wMt), "## ### ### ### ### ##0.00")
    XPrt.CurrentX = prtMinX + IIf(wMt < 0, 6000, 7500) - XPrt.TextWidth(X)
    XPrt.Print X;
End If
If mNbj <> 0 Then
    wMt = CCur(wDec / mNbj)
    X = Format$(Abs(wMt), "## ### ### ### ### ##0.00")
    XPrt.CurrentX = prtMinX + IIf(wMt < 0, 9000, 10500) - XPrt.TextWidth(X)
    XPrt.Print X;
End If
'XPrt.CurrentY = XPrt.CurrentY + Height8_6
XPrt.FontBold = True
XPrt.CurrentX = prtMinX + 10600
XPrt.Print DevX(mId$(reccptp0.Id, 1, 3));
XPrt.FontBold = False
'XPrt.CurrentY = XPrt.CurrentY - Height8_6

If mDbNb > 0 Then
    X = Format$(Abs(mDbNb), "### ### ##0")
    XPrt.CurrentX = prtMinX + 12000 - XPrt.TextWidth(X)
    XPrt.Print X;
    X = Format$(Abs(mDbMt), "## ### ### ### ### ##0.00")
'    X = Format$(Abs(mDbMt) / mDbNb, "## ### ### ### ### ##0.00")
    XPrt.CurrentX = prtMinX + 13500 - XPrt.TextWidth(X)
    XPrt.Print X;
End If

If mCrNb > 0 Then
    X = Format$(Abs(mCrNb), "### ### ##0")
    XPrt.CurrentX = prtMinX + 14000 - XPrt.TextWidth(X)
    XPrt.Print X;
    X = Format$(Abs(mCrMt), "## ### ### ### ### ##0.00")
'    X = Format$(Abs(mCrMt) / mCrNb, "## ### ### ### ### ##0.00")
    XPrt.CurrentX = prtMinX + 15500 - XPrt.TextWidth(X)
    XPrt.Print X;
End If

End Sub

