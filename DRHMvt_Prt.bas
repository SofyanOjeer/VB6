Attribute VB_Name = "prtDRHMvt"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public constDRHFérié As String, constDRHService  As String, constDRHMvt As String
Dim X As String, I As Integer, K As Integer, Height8_6 As Integer

Public recDRH As typeDRH
Public recDRHMvt As typeDRHMvt
Public xTable As typeElpTable

Dim iReturn As Integer

Dim prtEnTête As String, prtDestinataire As String
Dim prtDocument As String * 2, prtSort As String * 1, prtDébutAmj As String * 8, prtFinAmj As String * 8
Dim prtSelectK As String * 1, prtSelect As String, prtSelectMvtK As String * 1, prtSelectMvt As String

Public arrDRHNbjCivils() As Double, arrDRHNbjOuvrés() As Double, arrDRHNbj_Index As Integer
Public arrDRHNbjX(12) As String * 6
Public arrDRHTR() As Double

Dim wAmj As String * 8

Dim prtDRHMvt_LineNb As Integer, Trame_MinX As Integer, Trame_MaxX As Integer

Dim mService As String * 4, mMatricule As String * 5

Dim recLine As String
Public mTR_Nbj As Double
'---------------------------------------------------------
Public Sub prtDRHTR_Monitor(Msg As String)
'---------------------------------------------------------
Dim K As Integer, K1 As Integer, K2 As Integer, mDicrub As Integer
Dim X
Set XPrt = Printer


'frmElpPrt.Show vbModal 'vbModeless

prtOrientation = vbPRORPortrait
prtTitleText = "DRH : Etat des tickets restaurant"
prtPgmName = "prtDRHTR"
prtTitleUsr = usrName
prtFontName = "Courier"
prtLineNb = 1
prtlineHeight = 250
prtHeaderHeight = 300

frmElpPrt.prtStdInit
prtDRHTR_Form

Open Msg For Input As #1
Do Until EOF(1)
    Input #1, recLine
    prtDRHTR_Line
Loop
Close #1
prtFontName = prtFontNameZ
frmElpPrt.prtEndDoc

frmElpPrt.Hide

End Sub



'---------------------------------------------------------
Public Sub prtDRHTR_Form()
'---------------------------------------------------------
Dim X As String

XPrt.FontSize = 7
XPrt.FontBold = True

XPrt.DrawWidth = 3
XPrt.ForeColor = RGB(0, 0, 0)

XPrt.Line (prtMinX, prtMinY)-(prtMaxX, prtMaxY), , B
XPrt.Line (prtMinX, prtMinY + prtHeaderHeight)-(prtMaxX, prtMinY + prtHeaderHeight)

Call frmElpPrt.prtTrame(prtMinX + 7950, prtMinY + prtHeaderHeight + 20, prtMaxX - 20, prtMaxY - 20, " ", 245)

XPrt.DrawWidth = 1
'----------------------------------------ligne 1-----------------

XPrt.CurrentY = prtMinY + (prtHeaderHeight - XPrt.TextHeight(X)) / 2
XPrt.CurrentX = prtMinX + 50
XPrt.Print "Id";
XPrt.CurrentX = prtMinX + 1000
XPrt.Print "Matricule";
XPrt.CurrentX = XPrt.CurrentX + 700
XPrt.Print "Nom";
XPrt.CurrentX = XPrt.CurrentX + 2000
XPrt.Print "NB";
XPrt.CurrentX = XPrt.CurrentX + 100
XPrt.Print "Prix";
XPrt.CurrentX = XPrt.CurrentX + 50
XPrt.Print "Part Patronale";
XPrt.CurrentX = prtMinX + 8000
XPrt.Print "Contrôle BIA";

XPrt.CurrentY = prtMinY + prtHeaderHeight - XPrt.TextHeight("test")
XPrt.FontBold = False


End Sub


'---------------------------------------------------------
Public Sub prtDRHTR_Line()
'---------------------------------------------------------
Dim X As String, K As Integer, mCurrenty As Integer

If XPrt.CurrentY + prtlineHeight * 2 > prtMaxY Then
    frmElpPrt.prtNewPage
    prtDRHTR_Form
End If

'------------------------------------------ligne 1--------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'----------------------------------ligne2

XPrt.CurrentX = prtMinX + 50

XPrt.Print mId$(recLine, 1, 53);
XPrt.CurrentX = prtMinX + 3750
XPrt.Print mId$(recLine, 54, 2);
XPrt.CurrentX = prtMinX + 4750
XPrt.Print mId$(recLine, 56, 24);

XPrt.CurrentX = prtMinX + 8000

XPrt.Print mId$(recLine, 19, 4) & " :   "; '  mId$(xDRH.Matricule, 2, 4)
'XPrt.Print mId$(recLine, 30, 6) = Format$(I, "000000")
XPrt.Print mId$(recLine, 56, 2) & "    "; ' Format$(wNbj, "00")
XPrt.Print Trim(mId$(recLine, 36, 17)) & " "; 'mId$(xDRH.Nom, 1, 17)
XPrt.Print mId$(recLine, 54, 2); '  mId$(xDRH.Prénom, 1, 2)

End Sub




Public Sub prtDRHMvt_Nbj_Init(lAmj As String)

ReDim arrDRHNbjCivils(arrDRH_NB, 12)
ReDim arrDRHNbjOuvrés(arrDRH_NB, 12)
For I = 0 To arrDRH_NB
    For K = 0 To 12
        arrDRHNbjCivils(I, K) = 0
        arrDRHNbjOuvrés(I, K) = 0
    Next K
Next I

wAmj = lAmj
For K = 1 To 12
    arrDRHNbjX(K) = mId$(wAmj, 1, 6)
    wAmj = dateElp("MoisAdd", 1, wAmj)
Next K

End Sub

Public Sub prtDRHTR_Init()

ReDim arrDRHTR(arrDRH_NB)
For I = 0 To arrDRH_NB
    arrDRHTR(I) = 0
Next I

End Sub


'---------------------------------------------------------
 Public Sub prtDRHMvt_Monitor(Msg As String)
'---------------------------------------------------------
Dim K As Integer, K1 As Integer, K2 As Integer, Kmin As Integer, Kmax As Integer
Dim X As String

On Error GoTo prtError

Set XPrt = Printer
Height8_6 = frmElpPrt.prtHeightDelta(8, 6)
K1 = Val(mId$(Msg, 1, 6))
K2 = Val(mId$(Msg, 7, 6))

recElpTable_Init recElpTable
xElpTable = recElpTable
recElpTable.Method = "Seek="
recElpTable.Id = "BiaPgm"
prtTitleText = "Liste des prêts"

prtLineNb = 1

frmElpPrt.Show vbModeless


prtOrientation = vbPRORLandscape
prtPgmName = "prtDRHMvt"
prtTitleUsr = usrName

prtlineHeight = 300
prtHeaderHeight = 300

frmElpPrt.prtStdInit


prtDRHMvt_Form
'For K = K1 To 5 'K2
'    recDRHMvt = P_arrDRHMvt(K)
'    'recElpTable.K1 = recDRHMvt.Nature
'    tableElpTable_Read recElpTable
'
'    prtDRHMvt_Line
'
    DoEvents: If prtKillDoc Then frmElpPrt.Hide: Exit Sub

'Next K
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
 Public Sub prtDRHMvt_Close()
'---------------------------------------------------------

On Error GoTo prtError

If prtDocument = "10" Then
    XPrt.DrawWidth = 3
    prtDRHMvt_Line10_Trait
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY)
End If
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
 Public Sub prtDRHMvt_Open(Msg As String, lEnTête As String, lDestinataire As String)
'---------------------------------------------------------
Dim K As Integer, K1 As Integer, K2 As Integer, Kmin As Integer, Kmax As Integer
Dim X As String

On Error GoTo prtError
prtDocument = mId$(Msg, 14, 2)
prtSort = mId$(Msg, 16, 1)
prtDébutAmj = mId$(Msg, 17, 8)
prtFinAmj = mId$(Msg, 25, 8)
prtSelectK = mId$(Msg, 33, 1)
prtSelect = mId$(Msg, 34, 16)
prtSelectMvtK = mId$(Msg, 50, 1)
prtSelectMvt = mId$(Msg, 51, 4)

Set XPrt = Printer
Height8_6 = frmElpPrt.prtHeightDelta(8, 6)
prtTitleText = lEnTête

prtLineNb = 1

frmElpPrt.Show vbModeless
prtlineHeight = 300
prtHeaderHeight = 300

Select Case prtDocument
    Case "02": prtOrientation = vbPRORPortrait: prtlineHeight = 350: prtHeaderHeight = 350
    Case "10": prtOrientation = vbPRORPortrait: prtlineHeight = 250: prtHeaderHeight = 350

    Case Else: prtOrientation = vbPRORLandscape
End Select

prtPgmName = "prtDRHMvt"

If lDestinataire <> "" Then
    prtTitleUsr = lDestinataire
Else
    prtTitleUsr = usrName
End If



frmElpPrt.prtStdInit

recElpTable_Init xTable
xTable.Method = "Seek="
xTable.Id = "DRH"
mService = ""
mMatricule = ""
prtDRHMvt_Form
Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide
End Sub

'---------------------------------------------------------
Public Sub prtDRHMvt_Form()
'---------------------------------------------------------
Dim X As String, K As Integer
prtCurrentY = XPrt.CurrentY
XPrt.FontSize = 8

XPrt.FontBold = True
XPrt.DrawWidth = 3

Call frmElpPrt.prtTrame(prtMinX, prtCurrentY, prtMaxX, prtCurrentY + prtlineHeight, " ", 230)
XPrt.Line (prtMinX, prtMinY)-(prtMaxX, prtMinY)

'XPrt.DrawWidth = 1

'XPrt.Line (prtMinX, prtMinY)-(prtMinX, prtMaxY)
'XPrt.Line (prtMaxX, prtMinY)-(prtMaxX, prtMaxY)

'---------------------------------------------------------
XPrt.CurrentY = prtCurrentY + 50
Select Case prtDocument
     Case "01", "05"
        Trame_MinX = prtMinX + 8000: Trame_MaxX = prtMinX + 11000
        Call frmElpPrt.prtTrame(Trame_MinX, prtMinY + prtHeaderHeight, Trame_MaxX, prtMaxY - 10, " ", 245)
        XPrt.CurrentX = prtMinX + 400: XPrt.Print "Nom";
        XPrt.CurrentX = prtMinX + 8000: XPrt.Print " du " & dateImp(prtDébutAmj);
        XPrt.CurrentX = prtMinX + 9500: XPrt.Print " au " & dateImp(prtFinAmj);
        XPrt.CurrentX = prtMinX + 5000: XPrt.Print "Motif de l'absence";
        XPrt.FontSize = 6
        XPrt.CurrentY = XPrt.CurrentY + Height8_6
        XPrt.CurrentX = prtMinX + 11500: XPrt.Print "Civils";
        XPrt.CurrentX = prtMinX + 12300: XPrt.Print "Ouvrés";
        XPrt.CurrentX = prtMinX + 13500: XPrt.Print "Réf";
        XPrt.CurrentX = prtMinX + 14800: XPrt.Print "màj";
        XPrt.CurrentX = prtMinX + 15200: XPrt.Print "matricule";
        XPrt.CurrentY = XPrt.CurrentY - Height8_6
        XPrt.FontSize = 8
     Case "02"
        Trame_MinX = prtMaxX: Trame_MaxX = prtMaxX
        'Trame_MinX = prtMinX + 8000: Trame_MaxX = prtMaxX
        'Call frmElpPrt.prtTrame(Trame_MinX, prtMinY + prtHeaderHeight, Trame_MaxX, prtMaxY - 10, " ", 245)
        XPrt.CurrentX = prtMinX + 400: XPrt.Print "Nom";
        'jpl2001.11.16 XPrt.CurrentX = prtMinX + 3500: XPrt.Print "du" & dateImp(prtDébutAmj);
        'jpl2001.11.16 XPrt.CurrentX = prtMinX + 5000: XPrt.Print "au " & dateImp(prtFinAmj);
        XPrt.CurrentX = prtMinX + 7000: XPrt.Print "Motif de l'absence";
    Case "03", "04"
        Trame_MinX = prtMinX + 3400: Trame_MaxX = prtMinX + 5600
        Call frmElpPrt.prtTrame(Trame_MinX, prtMinY + prtHeaderHeight, prtMinX + 5100, prtMaxY - 10, " ", 250)
        Call frmElpPrt.prtTrame(prtMinX + 5100, prtMinY + prtHeaderHeight, Trame_MaxX, prtMaxY - 10, " ", 240)
        Call frmElpPrt.prtTrame(prtMinX + 15400, prtMinY + prtHeaderHeight, prtMaxX, prtMaxY - 10, " ", 250)
        XPrt.CurrentX = prtMinX + 400: XPrt.Print "Nom";
        XPrt.FontSize = 6
        XPrt.CurrentY = XPrt.CurrentY + Height8_6
        XPrt.CurrentX = prtMinX + 3500: XPrt.Print "Date d'entrée";
        XPrt.CurrentX = prtMinX + 4500: XPrt.Print "Nb enfants";
        XPrt.CurrentX = prtMinX + 5200: XPrt.Print "Total";
        XPrt.CurrentX = prtMinX + 15200: XPrt.Print "matricule";
        For K = 1 To 12
            XPrt.CurrentX = prtMinX + 5000 + 800 * K
            XPrt.Print Format$(arrDRHNbjX(K), "@@@@ - @@");
        Next K
        
        XPrt.CurrentY = XPrt.CurrentY - Height8_6
        XPrt.FontSize = 8
     Case "06", "07"
        Trame_MinX = prtMinX + 11500: Trame_MaxX = prtMinX + 13000
        Call frmElpPrt.prtTrame(Trame_MinX, prtMinY + prtHeaderHeight, Trame_MaxX, prtMaxY - 10, " ", 250)
        XPrt.CurrentX = prtMinX + 400: XPrt.Print "Nom";
        XPrt.CurrentX = prtMinX + 8000: XPrt.Print "du" & dateImp(prtDébutAmj);
        XPrt.CurrentX = prtMinX + 9500: XPrt.Print "au " & dateImp(prtFinAmj);
        XPrt.CurrentX = prtMinX + 5000: XPrt.Print "Motif de l'absence";
        XPrt.FontBold = True
        If prtDocument = "06" Then
            XPrt.CurrentX = prtMinX + 11200:
            XPrt.Print "TR : " & mTR_Nbj;
        Else
            XPrt.CurrentX = prtMinX + 11600:
            XPrt.Print "Total";
        End If
        XPrt.FontBold = False
        XPrt.FontSize = 6
        XPrt.CurrentY = XPrt.CurrentY + Height8_6
        XPrt.CurrentX = prtMinX + 12300: XPrt.Print "Abs. Ouvrées";
        XPrt.CurrentX = prtMinX + 13600: XPrt.Print "Réf";
        XPrt.CurrentX = prtMinX + 14800: XPrt.Print "màj";
        XPrt.CurrentX = prtMinX + 15200: XPrt.Print "matricule";
        XPrt.CurrentY = XPrt.CurrentY - Height8_6
        XPrt.FontSize = 8
     Case "10"
        Trame_MinX = prtMaxX: Trame_MaxX = prtMaxX
        'Trame_MinX = prtMinX + 8000: Trame_MaxX = prtMaxX
        'Call frmElpPrt.prtTrame(Trame_MinX, prtMinY + prtHeaderHeight, Trame_MaxX, prtMaxY - 10, " ", 245)
        XPrt.CurrentX = prtMinX + 50: XPrt.Print "Nom";
        XPrt.CurrentX = prtMinX + 4800: XPrt.Print "Droits";
        XPrt.CurrentX = prtMinX + 5400: XPrt.Print "Absences";
        XPrt.CurrentX = prtMinX + 6300: XPrt.Print "Solde";
        If prtSelectMvtK = "1" Then
            XPrt.CurrentX = prtMinX + 6900: XPrt.Print "du " & dateImp(prtDébutAmj);
            XPrt.CurrentX = prtMinX + 8300: XPrt.Print "au " & dateImp(prtFinAmj);
            XPrt.CurrentX = prtMinX + 9700: XPrt.Print "Mouvement";
        End If
        Call frmElpPrt.prtTrame(prtMinX + 4700, prtMinY + prtHeaderHeight, prtMinX + 6100, prtMaxY - 20, " ", 250)
        Call frmElpPrt.prtTrame(prtMinX + 6100, prtMinY + prtHeaderHeight, prtMinX + 6700, prtMaxY - 20, " ", 240)
        XPrt.CurrentY = prtMinY + 50
        XPrt.DrawWidth = 1

End Select

prtDRHMvt_LineNb = 0
End Sub

'---------------------------------------------------------
Public Sub prtDRHMvt_Line()
'---------------------------------------------------------

If XPrt.CurrentY + prtlineHeight * 1.9 > prtMaxY Then
    If prtDocument = "10" Then prtDRHMvt_Line10_Trait
    frmElpPrt.prtNewPage
    prtDRHMvt_Form
End If

If prtSort = "1" Then
    If mService <> recDRH.Service Then
        If Trim(mService) = "" Then XPrt.CurrentY = XPrt.CurrentY - prtlineHeight
        mService = recDRH.Service: prtDRHMvt_Service_Rupture
    End If
End If


XPrt.FontBold = False
prtDRHMvt_LineNb = prtDRHMvt_LineNb + 1
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
If prtDRHMvt_LineNb > 2 Then
    Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY, Trame_MinX, XPrt.CurrentY + prtlineHeight - 10, " ", 250)
    Call frmElpPrt.prtTrame(Trame_MaxX, XPrt.CurrentY, prtMaxX, XPrt.CurrentY + prtlineHeight - 10, " ", 250)
   If prtDRHMvt_LineNb = 4 Then prtDRHMvt_LineNb = 0
End If

Select Case prtDocument
    Case "01": prtDRHMvt_Line01
    Case "02": prtDRHMvt_Line02: prtDRHMvt_LineNb = 0
    Case "03", "04": prtDRHMvt_Line03
    Case "05": prtDRHMvt_Line01
    Case "06": prtDRHMvt_Line01: prtDRHMvt_LineNb = 0
    Case "07": prtDRHMvt_Line01: prtDRHMvt_LineNb = 0
    Case "10": prtDRHMvt_Line10: prtDRHMvt_LineNb = 0
End Select

End Sub
'---------------------------------------------------------
Public Sub prtDRHMvt_Print(lDRHMvt As typeDRHMvt, lDRH As typeDRH)
'---------------------------------------------------------
recDRHMvt = lDRHMvt
recDRH = lDRH
prtDRHMvt_Line

End Sub







Public Sub prtDRHMvt_Line01()

XPrt.CurrentX = prtMinX + 400
If mMatricule <> recDRH.Matricule Then
    mMatricule = recDRH.Matricule
    If prtDocument = "06" Or prtDocument = "07" Then prtDRHMvt_Line06
    XPrt.CurrentX = prtMinX + 400
    XPrt.Print srvDRH_Identité(recDRH);
End If
Call prtDRHMvt_Amj(recDRHMvt.DébutAmj, recDRHMvt.DébutAmjK, prtMinX + 8100)
If recDRHMvt.NbjChk = "0" Then
    XPrt.CurrentX = prtMinX + 9600
    XPrt.Print " à préciser";
Else
    Call prtDRHMvt_Amj(recDRHMvt.RepriseAmj, recDRHMvt.RepriseAmjK, prtMinX + 9600)
End If

If recDRHMvt.MvtCO = "C" Then Call prtDRHMvt_Nbj(recDRHMvt.Nbj, prtMinX + 11800)
Call prtDRHMvt_Nbj(recDRHMvt.NbjOuvré, prtMinX + 12700)
Select Case recDRHMvt.MvtSens
    Case "-"
    Case "C": XPrt.Print "+";
    Case "D", "P": XPrt.Print "-";
End Select
xTable.K1 = constDRHMvt
xTable.K2 = recDRHMvt.MvtCode
iReturn = tableElpTable_Read(xTable)
If iReturn <> 0 Then xTable.Name = "? " & recDRHMvt.MvtCode
XPrt.CurrentX = prtMinX + 5000
XPrt.Print xTable.Name;

XPrt.FontSize = 6
XPrt.CurrentY = XPrt.CurrentY + Height8_6
XPrt.CurrentX = prtMinX + 13500: XPrt.Print recDRHMvt.RéfInterne & " " & recDRHMvt.Statut;
XPrt.CurrentX = prtMinX + 14500: XPrt.Print dateImp(recDRHMvt.UpdAmj);
XPrt.CurrentX = prtMinX + 15400: XPrt.Print recDRHMvt.Matricule;

XPrt.CurrentY = XPrt.CurrentY - Height8_6
XPrt.FontSize = 8

End Sub
Public Sub prtDRHMvt_Line10()

XPrt.CurrentX = prtMinX + 400
If mMatricule <> recDRH.Matricule Then
    XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY)
    XPrt.CurrentY = XPrt.CurrentY + 10
    Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY, prtMinX + 4700, XPrt.CurrentY + prtlineHeight, " ", 250)
    mMatricule = recDRH.Matricule
    XPrt.CurrentX = prtMinX + 50
    XPrt.Print srvDRH_Identité(recDRH);
    prtDRHMvt_Line10_Total
    If prtSelectMvtK = "1" Then XPrt.Line (prtMinX + 6700, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY)
End If
XPrt.FontBold = False

If prtSelectMvtK = "1" Then
    prtDRHMvt_Line10_Détail
Else
    XPrt.CurrentY = XPrt.CurrentY - prtlineHeight
End If
End Sub

Public Sub prtDRHMvt_Line02()
Dim I02 As Integer, blnLine As Boolean

XPrt.CurrentX = prtMinX + 400
If mMatricule <> recDRH.Matricule Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    mMatricule = recDRH.Matricule
    XPrt.Print srvDRH_Identité(recDRH);
End If
Call prtDRHMvt_Amj(recDRHMvt.DébutAmj, recDRHMvt.DébutAmjK, prtMinX + 3500)
'Call prtDRHMvt_Amj(recDRHMvt.RepriseAmj, recDRHMvt.RepriseAmjK, prtMinX + 9600)
XPrt.FontBold = True
If recDRHMvt.NbjChk = "0" Then
    XPrt.CurrentX = prtMinX + 5000
    XPrt.Print " à préciser";
Else
    Call prtDRHMvt_Amj(recDRHMvt.RepriseAmj, recDRHMvt.RepriseAmjK, prtMinX + 5000)
End If

blnLine = False
XPrt.FontBold = False
XPrt.FontSize = 6
XPrt.CurrentY = XPrt.CurrentY + Height8_6
For I02 = 1 To arrDRHMvt_NB
    If arrDRHMvt(I02).DébutAmj >= recDRHMvt.DébutAmj _
    And arrDRHMvt(I02).RepriseAmj <= recDRHMvt.RepriseAmj Then
        If blnLine Then XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
        blnLine = True
        Call prtDRHMvt_Amj(arrDRHMvt(I02).DébutAmj, arrDRHMvt(I02).DébutAmjK, prtMinX + 7000)
        If arrDRHMvt(I02).NbjChk <> "0" Then Call prtDRHMvt_Amj(arrDRHMvt(I02).RepriseAmj, arrDRHMvt(I02).RepriseAmjK, prtMinX + 8200)
        xTable.K1 = constDRHMvt
        xTable.K2 = arrDRHMvt(I02).MvtCode
        iReturn = tableElpTable_Read(xTable)
        If iReturn <> 0 Then xTable.Name = "? " & arrDRHMvt(I02).MvtCode
        XPrt.CurrentX = prtMinX + 9400
        XPrt.Print xTable.Name;
    End If
Next I02

XPrt.CurrentY = XPrt.CurrentY - Height8_6
XPrt.FontSize = 8

End Sub

Public Sub prtDRHMvt_Line03()
Dim wNbj As Double

XPrt.CurrentX = prtMinX + 400: XPrt.Print srvDRH_Identité(recDRH);
XPrt.CurrentX = prtMinX + 3500: XPrt.Print dateImp(recDRH.EntréeAmj);
If recDRH.EnfantNb <> 0 Then
    X = Format$(recDRH.EnfantNb, "##0")
    XPrt.CurrentX = prtMinX + 5000 - XPrt.TextWidth(X)
    XPrt.Print X;
End If
For K = 0 To 12
    If prtDocument = "03" Then
        wNbj = arrDRHNbjOuvrés(arrDRH_Index, K)
    Else
        wNbj = arrDRHNbjCivils(arrDRH_Index, K)
    End If
    If wNbj = 0 Then
        XPrt.CurrentX = prtMinX + 5450 + 800 * K: XPrt.Print ".";
    Else
        Call prtDRHMvt_Nbj(wNbj, prtMinX + 5500 + 800 * K)
    End If
Next K
    
'Call prtDRHMvt_Nbj(wNbj, prtMinX + 5500)
XPrt.FontSize = 6
XPrt.CurrentY = XPrt.CurrentY + Height8_6
XPrt.CurrentX = prtMinX + 15400: XPrt.Print recDRH.Matricule;
XPrt.CurrentY = XPrt.CurrentY - Height8_6
XPrt.FontSize = 8

End Sub

Public Sub prtDRHMvt_Nbj(lNbj As Double, lCurrentX As Integer)
Dim X As String, D As Double, absNbj As Double

absNbj = Abs(lNbj)
    X = Format$(Fix(absNbj), "##0")
    XPrt.CurrentX = lCurrentX - XPrt.TextWidth(X)
    XPrt.Print X;
    D = absNbj - Fix(absNbj)
    If D <> 0 Then XPrt.Print Format$(D, "#.0");
    If lNbj < 0 Then XPrt.Print "-";
End Sub
Public Sub prtDRHMvt_Amj(lAmj As String, lAmjK As String, lCurrentX As Integer)
Dim X As String, D As Double
XPrt.CurrentX = lCurrentX
XPrt.Print dateImp(lAmj);

If lAmjK <> "0" Then
    If XPrt.FontSize <> 6 Then
        XPrt.FontSize = 6
        XPrt.CurrentY = XPrt.CurrentY + Height8_6
        XPrt.Print " midi";
        XPrt.CurrentY = XPrt.CurrentY - Height8_6
        XPrt.FontSize = 8
    Else
        XPrt.Print " midi";
    End If
    
End If

End Sub


Public Sub prtDRHMvt_Service_Rupture()
prtDRHMvt_LineNb = 0
If XPrt.CurrentY + prtlineHeight * 4 > prtMaxY Then
    frmElpPrt.prtNewPage
    prtDRHMvt_Form
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
Else
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
End If
    
Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY - 50, Trame_MinX, XPrt.CurrentY + prtlineHeight - 50, " ", 245)
xTable.K1 = constDRHService
xTable.K2 = mService
iReturn = tableElpTable_Read(xTable)
XPrt.FontBold = True
XPrt.FontSize = 10
XPrt.CurrentX = prtMinX + 25
If iReturn <> 0 Then xTable.Memo = "? " & recDRH.Service
XPrt.Print xTable.Memo;
XPrt.FontBold = False
XPrt.FontSize = 8
'XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 0.5
End Sub

Public Sub prtDRHMvt_Line05()
Dim I As Integer, X As String

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY, prtMaxX, XPrt.CurrentY + prtlineHeight, " ", 230)
'XPrt.CurrentX = prtMinX + 5000: XPrt.Print "Récapitulatif";
XPrt.CurrentX = prtMinX + 8500: XPrt.Print "Absences";
XPrt.CurrentX = prtMinX + 10500: XPrt.Print "Droits ";

For I = 1 To frmDRH.fgTotal.Rows - 1
    If XPrt.CurrentY + prtlineHeight * 1.5 > prtMaxY Then
        frmElpPrt.prtNewPage
        XPrt.CurrentY = prtMinX + prtlineHeight * 3
        prtDRHMvt_Form
    End If
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    frmDRH.fgTotal.Row = I
    frmDRH.fgTotal.Col = 0: XPrt.CurrentX = prtMinX + 5000: XPrt.Print Trim(frmDRH.fgTotal.Text);
    frmDRH.fgTotal.Col = 1: X = Trim(frmDRH.fgTotal.Text)
    XPrt.CurrentX = prtMinX + 9000 - XPrt.TextWidth(X): XPrt.Print X;
    frmDRH.fgTotal.Col = 2: X = Trim(frmDRH.fgTotal.Text)
    XPrt.CurrentX = prtMinX + 10800 - XPrt.TextWidth(X): XPrt.Print X;
Next I
End Sub

Public Sub prtDRHMvt_Line06()
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight '* 0.5
Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY - 50, Trame_MinX, XPrt.CurrentY + prtlineHeight - 50, " ", 250)
'Call frmElpPrt.prtTrame(Trame_MaxX, XPrt.CurrentY - 50, prtMaxX, XPrt.CurrentY + prtlineHeight - 50, " ", 240)
XPrt.FontBold = True
Call prtDRHMvt_Nbj(arrDRHTR(arrDRH_Index), prtMinX + 11800)
XPrt.FontBold = False
End Sub

Public Sub prtDRHMvt_Line10_Total()
Dim I As Integer, Nb As Double

For I = 1 To 99 ' fgTotal.Rows - 1
    If arrDRHMvt_Absences_Nb(I) <> 0 Or arrDRHMvt_Droits_Nb(I) <> 0 Then
        XPrt.CurrentX = prtMinX + 6800
        XPrt.Print arrDRHMvt_Libellé(I);
        If arrDRHMvt_Droits_Nb(I) <> 0 Then Call prtDRHMvt_Nbj(arrDRHMvt_Droits_Nb(I), prtMinX + 5300)
        If arrDRHMvt_Absences_Nb(I) <> 0 Then Call prtDRHMvt_Nbj(arrDRHMvt_Absences_Nb(I), prtMinX + 5900)
        Nb = arrDRHMvt_Droits_Nb(I) - arrDRHMvt_Absences_Nb(I)
        Call prtDRHMvt_Nbj(Nb, prtMinX + 6500)
        XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    End If
Next I

End Sub

Public Sub prtDRHMvt_Line10_Détail()
XPrt.FontSize = 6
XPrt.CurrentY = XPrt.CurrentY + Height8_6

If recDRHMvt.MvtCO = "C" Then
    Call prtDRHMvt_Nbj(recDRHMvt.Nbj, prtMinX + 7000)
Else
    Call prtDRHMvt_Nbj(recDRHMvt.NbjOuvré, prtMinX + 7000)
End If
Select Case recDRHMvt.MvtSens
    Case "-"
    Case "C": XPrt.Print "+";
    Case "D", "P": XPrt.Print "-";
End Select

Call prtDRHMvt_Amj(recDRHMvt.DébutAmj, recDRHMvt.DébutAmjK, prtMinX + 7300)
If recDRHMvt.NbjChk = "0" Then
    XPrt.CurrentX = prtMinX + 8400
    XPrt.Print " à préciser";
Else
    Call prtDRHMvt_Amj(recDRHMvt.RepriseAmj, recDRHMvt.RepriseAmjK, prtMinX + 8400)
End If

xTable.K1 = constDRHMvt
xTable.K2 = recDRHMvt.MvtCode
iReturn = tableElpTable_Read(xTable)
If iReturn <> 0 Then xTable.Name = "? " & recDRHMvt.MvtCode
XPrt.CurrentX = prtMinX + 9500
XPrt.Print xTable.Name;

XPrt.CurrentY = XPrt.CurrentY - Height8_6
XPrt.FontSize = 8

End Sub

Public Sub prtDRHMvt_Line10_Trait()
'XPrt.Line (prtMinX + 4700, prtMinY)-(prtMinX + 4700, prtMaxY)
'XPrt.Line (prtMinX + 6700, prtMinY)-(prtMinX + 6700, prtMaxY)

End Sub
