Attribute VB_Name = "prtDictio"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------

Private recDevise As typeDevise
Dim NbImprimÈ As Integer
Dim mOrigine As String, mSort As Integer, mAmj As String * 8, mHHMM As String * 4, mMethod As String
Dim recDeviseChange As typeDeviseChange
Dim vReturn
Dim nbLigne As Integer, Height8_6 As Integer
'---------------------------------------------------------
Public Sub prtDeviseX(Msg As String)
'---------------------------------------------------------
Dim K As Integer, K1 As Integer, K2 As Integer
Dim X
On Error GoTo prtError

Set XPrt = Printer
K1 = Val(mId$(Msg, 1, 6))
K2 = Val(mId$(Msg, 7, 6))

If mId$(Msg, 21, 1) = "C" Then
    mMethod = "LastLC"
    mOrigine = "C"
    mHHMM = "9999"
    prtTitleText = "cours de change ComptabilitÈ au "
Else
    mMethod = "LastLT"
    mOrigine = "T"
    mHHMM = "0000"
    prtTitleText = "cours de change TrÈsorerie au "
End If

mSort = IIf(mId$(Msg, 22, 1) = "N", 888, 889)
mAmj = mId$(Msg, 13, 8)
If Val(mAmj) = 0 Then mAmj = DSys

recDeviseChange_Init recDeviseChange

'frmElpPrt.Show vbModal 'vbModeless
Height8_6 = frmElpPrt.prtHeightDelta(8, 6)

prtOrientation = vbPRORPortrait
prtTitleText = prtTitleText & dateImp(mAmj)
prtPgmName = "prtDevise"
prtTitleUsr = usrName

prtLineNb = 1
prtlineHeight = 300
prtHeaderHeight = 300

frmElpPrt.prtStdInit
prtDeviseForm

XDictio.Method = "Seek>=      "
XDictio.DicCode = String$(20, Chr$(0))
XDictio.DicRub = mSort

X = dbDictioRead(XDictio)
XDictio.Method = "MoveNext    "

Do While XDictio.Err = 0 _
     And XDictio.DicRub = mSort
        
    If mSort = 888 Then
        recDevise.DevCode = mId$(XDictio.DicCode, 1, 3)
        recDevise.DevX = mId$(XDictio.DicTxt, 1, 3)
    Else
        recDevise.DevX = mId$(XDictio.DicCode, 1, 3)
        recDevise.DevCode = mId$(XDictio.DicTxt, 1, 3)
    End If
    recDevise.DevLib = XDictio.DicLib
    
    recDeviseChange.Method = mMethod
    recDeviseChange.Id2 = recDevise.DevX
    recDeviseChange.Id1 = "EUR"
    recDeviseChange.Amj = mAmj
    recDeviseChange.Origine = mOrigine
    recDeviseChange.HHMM = mHHMM
    vReturn = srvDeviseChange_Monitor(recDeviseChange)


'    recDevise.Cours = Val(Mid$(XDictio.DicTxt, 4, 13))
 '   recDevise.DevAmj = Mid$(XDictio.DicTxt, 17, 8)

    prtDeviseLine

    DoEvents: If prtKillDoc Then frmElpPrt.Hide: Exit Sub

    X = dbDictioRead(XDictio)

Loop

frmElpPrt.prtEndDoc

frmElpPrt.Hide
Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.prtEndDoc
frmElpPrt.Hide

End Sub

'---------------------------------------------------------
Public Sub prtDictioX(Msg As String)
'---------------------------------------------------------
Dim K As Integer, K1 As Integer, K2 As Integer, mDicrub As Integer
Dim X
On Error GoTo prtError

Set XPrt = Printer
mDicrub = Val(Msg)


'frmElpPrt.Show vbModal 'vbModeless

prtOrientation = vbPRORPortrait
prtTitleText = "RÈfÈrentiel Pays"
prtPgmName = "prtDictio"
prtTitleUsr = usrName

prtLineNb = 1
prtlineHeight = 250
prtHeaderHeight = 300

frmElpPrt.prtStdInit
prtDictioForm

XDictio.Method = "Seek>=      "
XDictio.DicCode = String$(20, Chr$(0))
XDictio.DicRub = mDicrub

X = dbDictioRead(XDictio)
XDictio.Method = "MoveNext    "
K = 0
Do While XDictio.Err = 0 _
     And XDictio.DicRub = mDicrub
        
    prtDictioLine
    
'Test    If K > 57 Then Exit Do
    K = K + 1
    
    DoEvents: If prtKillDoc Then frmElpPrt.Hide: Exit Sub

    X = dbDictioRead(XDictio)

Loop

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
Public Sub prtDeviseForm()
'---------------------------------------------------------
Dim X As String

XPrt.FontSize = 8
XPrt.FontBold = True

XPrt.FillStyle = 0
XPrt.DrawWidth = 3
XPrt.ForeColor = RGB(0, 0, 0)
XPrt.FillStyle = 1

XPrt.Line (prtMinX, prtMinY)-(prtMaxX, prtMaxY), , B
'XPrt.Line (prtMinX, prtMinY + prtHeaderHeight)-(prtMaxX, prtMinY + prtHeaderHeight)
Call frmElpPrt.prtTrame(prtMinX, prtMinY, prtMaxX, prtMinY + prtHeaderHeight, "B", 240)
XPrt.DrawWidth = 2


'----------------------------------------ligne 1-----------------

XPrt.CurrentY = prtMinY + (prtHeaderHeight - XPrt.TextHeight(X)) / 2
XPrt.CurrentX = 2700
XPrt.Print "1";

XPrt.CurrentX = 1600
XPrt.Print "EUR";

XPrt.CurrentX = 4200
XPrt.Print "LibellÈ";

XPrt.CurrentX = prtMinX + 100
XPrt.Print "Date du cours";

'XPrt.CurrentX = 9000
'XPrt.Print "Agence";

'XPrt.CurrentX = 9650
'XPrt.Print ":";
'XPrt.CurrentX = 9800
'XPrt.Print strSocBdfE;
'XPrt.Print " ""  & strSocBdfG;"

prtCurrentY = prtMinY + prtHeaderHeight + 100
nbLigne = 0

End Sub

'---------------------------------------------------------
Public Sub prtDictioForm()
'---------------------------------------------------------
Dim X As String

XPrt.FontSize = 8
XPrt.FontBold = True

XPrt.DrawWidth = 3
XPrt.ForeColor = RGB(0, 0, 0)

XPrt.Line (prtMinX, prtMinY)-(prtMaxX, prtMaxY), , B
XPrt.Line (prtMinX, prtMinY + prtHeaderHeight)-(prtMaxX, prtMinY + prtHeaderHeight)


XPrt.DrawWidth = 1
'----------------------------------------ligne 1-----------------

XPrt.FontSize = 6
XPrt.CurrentY = prtMinY + (prtHeaderHeight - XPrt.TextHeight(X)) / 2
XPrt.CurrentX = 400
XPrt.Print "Code";
XPrt.CurrentX = 1500
XPrt.Print "LibellÈ";
XPrt.CurrentX = 3900
XPrt.Print "Union";
XPrt.CurrentX = 4650
XPrt.Print "Euro";
XPrt.CurrentX = 5350
XPrt.Print "Mbres OCDE";
XPrt.CurrentX = 6400
XPrt.Print "CB:zone A/B";
XPrt.CurrentX = 7650
XPrt.Print "Zone Pays";
XPrt.CurrentX = 8800
XPrt.Print "Risque Pays";


XPrt.CurrentY = prtMinY + prtHeaderHeight - XPrt.TextHeight("test")
   NbImprimÈ = 0

XPrt.FontSize = 8

End Sub


'---------------------------------------------------------
Public Sub prtDictioLine()
'---------------------------------------------------------
Dim X As String, K As Integer, mCurrenty As Integer

If XPrt.CurrentY + prtlineHeight * 2 > prtMaxY Then
    frmElpPrt.prtNewPage
    prtDictioForm
'Else
    'frmElpPrt.prtLineY
End If

NbImprimÈ = NbImprimÈ + 1
If NbImprimÈ = 7 Then
'   Call frmElpPrt.prtTrame(prtMinX + 20, XPrt.CurrentY - prtlineHeight * 2 - 50, prtMaxX - 20, XPrt.CurrentY + prtlineHeight - 50, " ")
   NbImprimÈ = 1
End If

'------------------------------------------ligne 1--------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'----------------------------------ligne2
XPrt.FontBold = True

XPrt.CurrentX = 400

XPrt.Print Format$(mId$(XDictio.DicCode, 1, 4), "0000");
XPrt.FontBold = False
XPrt.CurrentX = 1000
XPrt.Print mId$(XDictio.DicTxt, 7, 2);
XPrt.CurrentX = 1500
XPrt.Print XDictio.DicLib;

Select Case mId$(XDictio.DicTxt, 1, 1)
    Case "1"
        XPrt.CurrentX = 4000
        XPrt.Print "UE";
End Select
Select Case mId$(XDictio.DicTxt, 9, 1)
     Case "O"
     XPrt.CurrentX = 4700
     XPrt.Print "IN";
 End Select


Select Case mId$(XDictio.DicTxt, 3, 1)
    Case "1"
        XPrt.CurrentX = 5500
        XPrt.Print "OCDE";
End Select

Select Case mId$(XDictio.DicTxt, 10, 1)
     Case "A"
     XPrt.CurrentX = 6500
     XPrt.Print "Zone A";
 End Select

Select Case mId$(XDictio.DicTxt, 10, 1)
     Case "B"
     XPrt.CurrentX = 6500
     XPrt.Print "Zone B";
 End Select

Select Case mId$(XDictio.DicTxt, 4, 3)
     Case "AUT"
     XPrt.CurrentX = 7800
     XPrt.Print "Autre";
 End Select

Select Case mId$(XDictio.DicTxt, 4, 3)
     Case "ARB"
     XPrt.CurrentX = 7800
     XPrt.Print "Arabe";
 End Select
 
Select Case mId$(XDictio.DicTxt, 4, 3)
     Case "AFR"
     XPrt.CurrentX = 7800
     XPrt.Print "Afrique";
 End Select



Select Case mId$(XDictio.DicTxt, 11, 1)
     Case "O"
     XPrt.CurrentX = 8800
     XPrt.Print "Risque";
 End Select















End Sub



'---------------------------------------------------------
Public Sub prtDeviseLine()
'---------------------------------------------------------
Dim X As String, K As Integer, mCurrenty As Integer

If prtCurrentY + prtParagraphHeight > prtMaxY Then
    frmElpPrt.prtNewPage
    prtDeviseForm
'Else
    'frmElpPrt.prtLineY
End If
XPrt.DrawWidth = 1
XPrt.FontBold = False

'------------------------------------------ligne 1--------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight

nbLigne = nbLigne + 1
If nbLigne > 2 Then
    Call frmElpPrt.prtTrame(prtMinX + 20, XPrt.CurrentY - 50, prtMaxX - 20, XPrt.CurrentY + prtlineHeight - 50, " ")
    If nbLigne = 4 Then nbLigne = 0
End If

XPrt.FontSize = 8
'----------------------------------ligne2
XPrt.FontBold = True

XPrt.CurrentX = 3700
XPrt.Print Format$(recDevise.DevCode, "000");
XPrt.CurrentX = 1600
XPrt.Print Format$(recDevise.DevX, "00000");
XPrt.FontBold = False
XPrt.CurrentX = 4200
XPrt.Print recDevise.DevLib;


If IsNull(vReturn) Then
   If Trim(recDeviseChange.ValidationUsr) = "" _
   Or recDeviseChange.ValidationUsr = const‡Valider Then XPrt.CurrentX = prtMinX: XPrt.Print "???";
   
   If recDeviseChange.QD1 <> 0 Then
        'X = Format$(recDevise.Cours, "### ##0.000 000")
        X = Format$(recDeviseChange.QD2CoursPivot / recDeviseChange.QD1, "### ##0.000 00")
        XPrt.CurrentX = 3300 - XPrt.TextWidth(X)
        XPrt.Print X;
        
        XPrt.CurrentY = XPrt.CurrentY + Height8_6
        XPrt.FontSize = 6
        XPrt.CurrentX = prtMinX + 300
        'XPrt.Print dateImp(recDevise.DevAmj);
        XPrt.Print dateImp(recDeviseChange.Amj);
        XPrt.CurrentY = XPrt.CurrentY - Height8_6
        XPrt.FontSize = 8
    End If
End If

End Sub






