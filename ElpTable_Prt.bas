Attribute VB_Name = "prtElpTable"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------

Dim recElpTable As typeElpTable
Dim nbLigne As Integer, Height8_6 As Integer
Dim intReturn As Integer, xId As String * 12, I As Integer

'---------------------------------------------------------
Public Sub prtElpTableX(Msg As String)
'---------------------------------------------------------
Dim X As String
On Error GoTo prtError

Set XPrt = Printer
xId = mId$(Msg, 1, 12)
prtFontName = prtFontName_CourierNew
frmElpPrt.Show vbModeless
Height8_6 = frmElpPrt.prtHeightDelta(8, 6)

prtOrientation = vbPRORLandscape
prtPgmName = "prtElpTable"
prtTitleUsr = usrName
prtTitleText = "Table : " & xId

prtLineNb = 1
prtlineHeight = 250
prtHeaderHeight = 300

frmElpPrt.prtStdInit
prtElpTable_Form

X = "select * from ElpTable where SNN = 0" _
    & " and id = '" & xId & "' order by K1,K2"
    
Set rsMDB = cnMDB.Execute(X)
Do While Not rsMDB.EOF
    Call rsElpTable_GetBuffer(rsMDB, recElpTable)

    prtElpTable_Line

    DoEvents: If prtKillDoc Then frmElpPrt.Hide: Exit Sub
    rsMDB.MoveNext
Loop

frmElpPrt.prtEndDoc 1000

frmElpPrt.Hide
Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.prtEndDoc 1000
frmElpPrt.Hide

End Sub

'---------------------------------------------------------
Public Sub prtElpTable_Form()
'---------------------------------------------------------
Dim X As String

XPrt.FontSize = 7
XPrt.FontBold = True
XPrt.DrawWidth = 3

XPrt.Line (prtMinX, prtMinY)-(prtMaxX, prtMaxY), prtLineColor, B
'XPrt.Line (prtMinX, prtMinY + prtHeaderHeight)-(prtMaxX, prtMinY + prtHeaderHeight)
Call frmElpPrt.prtTrame(prtMinX, prtMinY, prtMaxX, prtMinY + prtHeaderHeight, "B", 240)
XPrt.DrawWidth = 1


'----------------------------------------ligne 1-----------------

XPrt.CurrentY = prtMinY + (prtHeaderHeight - XPrt.TextHeight(X)) / 2

XPrt.CurrentX = prtMinX + 100
XPrt.Print "Clé 1";

XPrt.CurrentX = prtMinX + 2000
XPrt.Print "Clé 2";

XPrt.CurrentX = prtMinX + 4000
XPrt.Print "Intitulé";
XPrt.CurrentX = prtMinX + 8000
XPrt.Print "Mémo";


prtCurrentY = prtMinY + prtHeaderHeight + 100
nbLigne = 0

End Sub

'---------------------------------------------------------
Public Sub prtElpTable_Line()
'---------------------------------------------------------
Dim X As String, K As Integer, mCurrenty As Integer
Dim xLine1 As String, xLine2 As String
Dim kMax As Integer, kJust As Integer
If XPrt.CurrentY + prtParagraphHeight > prtMaxY - 100 Then
    frmElpPrt.prtNewPage
    prtElpTable_Form
'Else
    'frmElpPrt.prtLineY
End If
XPrt.DrawWidth = 1
XPrt.FontBold = False

'------------------------------------------ligne 1--------------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight

nbLigne = nbLigne + 1
If nbLigne > 2 Then
    Call frmElpPrt.prtTrame(prtMinX + 10, XPrt.CurrentY - 50, prtMaxX - 10, XPrt.CurrentY + prtlineHeight - 50, "240")
    If nbLigne = 4 Then nbLigne = 0
End If

XPrt.FontSize = 8
'----------------------------------

XPrt.CurrentX = prtMinX + 100
XPrt.Print recElpTable.K1;
XPrt.CurrentX = prtMinX + 2000
XPrt.Print recElpTable.K2;
XPrt.CurrentX = prtMinX + 4000
XPrt.Print recElpTable.Name;
If Not IsNull(recElpTable.Memo) Then
    XPrt.CurrentX = prtMinX + 8000
    X = recElpTable.Memo
    If XPrt.TextWidth(X) <= 7500 Then
        XPrt.Print X;
    Else
    For kMax = Len(X) To 1 Step -1
        xLine1 = mId$(X, 1, kMax)
        If XPrt.TextWidth(xLine1) <= 7500 Then Exit For
    Next kMax
  
    kJust = kMax
    For I = kMax To kMax - 10 Step -1
        If mId$(X, I, 1) = " " Then kJust = I: Exit For
    Next I
    xLine1 = mId$(X, 1, kJust)
    xLine2 = mId$(X, kJust + 1, Len(X) - kJust)
        
    XPrt.Print xLine1;
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = prtMinX + 8000
    XPrt.Print xLine2;
    End If
    
End If


End Sub







