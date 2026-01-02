Attribute VB_Name = "prtBiaPgm"
Option Explicit
Dim I As Integer, mCurrenty As Integer
Dim colP As Integer, colA As Integer, colV As Integer, colX As Integer

Dim V, Height8_6 As Integer
Dim mK1 As String * 12
'---------------------------------------------------------
 Public Sub prtBiaPgm_Open(Msg As String)
'---------------------------------------------------------
Dim K As Integer, K1 As Integer, K2 As Integer
Dim X As String

Set XPrt = Printer
Height8_6 = frmElpPrt.prtHeightDelta(8, 6)


frmElpPrt.Show vbModeless

prtOrientation = vbPRORPortrait
prtTitleText = Msg
prtPgmName = "prtBiaPgm"
prtTitleUsr = usrName

prtLineNb = 1
prtlineHeight = 250
prtHeaderHeight = 300

frmElpPrt.prtStdInit
colP = prtMinX
colA = prtMinX + 6000
colV = prtMinX + 8500
colX = prtMaxX
prtBiaPgm_Form
mK1 = ""

End Sub
'---------------------------------------------------------
 Public Sub prtBiaPgm_Close()
'---------------------------------------------------------
prtBiaPgm_Form_End
XPrt.DrawWidth = 6
XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMinX, XPrt.CurrentY), prtLineColor
frmElpPrt.prtEndDoc 1000
frmElpPrt.Hide

End Sub

'---------------------------------------------------------
Public Sub prtBiaPgm_Form()
'---------------------------------------------------------
Dim X As String, K As Integer

XPrt.FontSize = 8
XPrt.FontBold = False

XPrt.DrawWidth = 3
Call frmElpPrt.prtTrame(prtMinX, prtMinY, prtMaxX, prtMinY + prtHeaderHeight, "B")

XPrt.Line (prtMinX, prtMinY)-(prtMinX, prtMaxY), prtLineColor
XPrt.Line (prtMaxX, prtMinY)-(prtMaxX, prtMaxY), prtLineColor
K = prtMinY + prtHeaderHeight + 10

'---------------------------------------------------------
XPrt.FontBold = True
XPrt.CurrentY = prtMinY + 50
XPrt.CurrentX = prtMinX + 50
XPrt.Print "Identifiant";

XPrt.CurrentX = prtMinX + 1500
XPrt.Print "Intitulé";
XPrt.CurrentX = colA + 100

XPrt.Print "Autorisations";
XPrt.CurrentX = colV + 100
XPrt.Print "Projet.vbp";
XPrt.FontBold = False

XPrt.CurrentY = prtMinY + prtHeaderHeight + prtlineHeight - XPrt.TextHeight("test")

End Sub

'---------------------------------------------------------
Public Sub prtBiaPgm_Line(recElpTable As typeElpTable, Msg As String)
'---------------------------------------------------------
Dim X As String, K As Integer
Dim Situation As String, xSens As String

If XPrt.CurrentY + prtlineHeight > prtMaxY Then
    prtBiaPgm_Form_End
    frmElpPrt.prtNewPage
    prtBiaPgm_Form
'Else
    'frmElpPrt.prtLineY
End If

XPrt.FontBold = False
'------------------------------------------ligne 1--------------

XPrt.FontSize = 8
If Msg = "1" Then Call frmElpPrt.prtTrame(prtMinX + 10, XPrt.CurrentY - 60, prtMaxX - 10, XPrt.CurrentY + prtlineHeight - 50, " ")
If Msg <= "1" Then
    XPrt.CurrentX = colP + 100
    XPrt.Print recElpTable.K1;
    XPrt.CurrentX = colP + 1500
    XPrt.Print recElpTable.Name;
    XPrt.CurrentX = colV + 100
    XPrt.Print mId$(recElpTable.Memo, 21, 20);
End If

If Msg = "2" Then
    XPrt.CurrentX = colA - 2000
    XPrt.Print recElpTable.K1;
    XPrt.CurrentX = colV + 100
    XPrt.Print dateImp(mId$(recElpTable.Memo, 21, 8)) & "    " & timeImp(mId$(recElpTable.Memo, 29, 6));
End If

XPrt.CurrentX = colA + 100
X = prtBiaMsg_Aut(mId$(recElpTable.Memo, 1, 20))
XPrt.Print X;

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight

End Sub

'---------------------------------------------------------
Public Sub prtBiaPgm_LineAut(recElpTable As typeElpTable)
'---------------------------------------------------------
Dim X As String, K As Integer
Dim Situation As String, xSens As String

If XPrt.CurrentY + prtlineHeight > prtMaxY Then
    prtBiaPgm_Form_End
    frmElpPrt.prtNewPage
    prtBiaPgm_Form
'Else
    'frmElpPrt.prtLineY
End If

XPrt.FontBold = False
'------------------------------------------ligne 1--------------

XPrt.FontSize = 8
If mK1 <> recElpTable.K1 Then
    mK1 = recElpTable.K1
    Call frmElpPrt.prtTrame(prtMinX + 10, XPrt.CurrentY - 60, prtMaxX - 10, XPrt.CurrentY + prtlineHeight - 50, " ")
    XPrt.CurrentX = colP + 100
    XPrt.Print recElpTable.K1;
End If
XPrt.CurrentX = colP + 1500
XPrt.Print recElpTable.K2;
XPrt.CurrentX = colP + 3000
XPrt.Print recElpTable.Name;
XPrt.CurrentX = colV + 100
XPrt.Print dateImp(mId$(recElpTable.Memo, 21, 8)) & "    " & timeImp(mId$(recElpTable.Memo, 29, 6));

XPrt.CurrentX = colA + 100
X = prtBiaMsg_Aut(mId$(recElpTable.Memo, 1, 20))
XPrt.Print X;

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight

End Sub


Public Sub prtBiaPgm_Form_End()
XPrt.DrawWidth = 2
XPrt.Line (colA, prtMinY)-(colA, prtMaxY), prtLineColor
XPrt.Line (colV, prtMinY)-(colV, prtMaxY), prtLineColor

End Sub


Public Function prtBiaMsg_Aut(X20 As String)
Dim I As Integer
prtBiaMsg_Aut = Space$(36)
Mid$(prtBiaMsg_Aut, 1, 1) = IIf(mId$(X20, 1, 1) = "X", "C", "-")
Mid$(prtBiaMsg_Aut, 5, 1) = IIf(mId$(X20, 2, 1) = "X", "S", "-")
Mid$(prtBiaMsg_Aut, 9, 1) = IIf(mId$(X20, 3, 1) = "X", "V", "-")
Mid$(prtBiaMsg_Aut, 13, 1) = IIf(mId$(X20, 4, 1) = "X", "M", "-")
Mid$(prtBiaMsg_Aut, 17, 1) = IIf(mId$(X20, 5, 1) = "X", "R", "-")
Mid$(prtBiaMsg_Aut, 21, 1) = IIf(mId$(X20, 6, 1) = "X", "S", "-")
Mid$(prtBiaMsg_Aut, 25, 1) = IIf(mId$(X20, 7, 1) = "X", "V", "-")
Mid$(prtBiaMsg_Aut, 29, 1) = IIf(mId$(X20, 8, 1) = "X", "A", "-")
Mid$(prtBiaMsg_Aut, 33, 1) = IIf(mId$(X20, 9, 1) = "X", "X", "-")

End Function
