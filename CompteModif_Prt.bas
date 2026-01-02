Attribute VB_Name = "prtCompteModif"
Option Explicit
Dim I As Integer, mCurrenty As Integer
Dim colP As Integer, colA As Integer, colV As Integer, colX As Integer

Dim V, Height8_6 As Integer
'---------------------------------------------------------
 Public Sub prtCompteModif_Open()
'---------------------------------------------------------
Dim K As Integer, K1 As Integer, K2 As Integer
Dim X As String

Set XPrt = Printer
Height8_6 = frmElpPrt.prtHeightDelta(8, 6)


frmElpPrt.Show vbModeless

prtOrientation = vbPRORLandscape
prtTitleText = "Modification de compte "
prtPgmName = "prtCompteModif"
prtTitleUsr = usrName

prtLineNb = 1
prtlineHeight = 250
prtHeaderHeight = 300

frmElpPrt.prtStdInit
colP = prtMinX
colA = prtMinX + 11000
colV = prtMinX + 13500
colX = prtMaxX
prtCompteModif_Form

End Sub
'---------------------------------------------------------
 Public Sub prtCompteModif_Close()
'---------------------------------------------------------
prtCompteModif_Form_End
frmElpPrt.prtEndDoc
frmElpPrt.Hide

End Sub

'---------------------------------------------------------
Public Sub prtCompteModif_Form()
'---------------------------------------------------------
Dim X As String, K As Integer

XPrt.FontSize = 8
XPrt.FontBold = False

XPrt.DrawWidth = 3
Call frmElpPrt.prtTrame(prtMinX, prtMinY, prtMaxX, prtMinY + prtHeaderHeight, "B")

XPrt.Line (prtMinX, prtMinY)-(prtMinX, prtMaxY)
XPrt.Line (prtMaxX, prtMinY)-(prtMaxX, prtMaxY)
K = prtMinY + prtHeaderHeight + 10

'---------------------------------------------------------
XPrt.FontBold = True
XPrt.CurrentY = prtMinY + 50
XPrt.CurrentX = prtMinX + 50
XPrt.Print "Devise";

XPrt.CurrentX = prtMinX + 700
XPrt.Print "Compte";
XPrt.CurrentX = colA + 100

XPrt.Print "Périodicité extrait";
XPrt.CurrentX = colV + 100
XPrt.Print "Retenue courrier";
XPrt.FontBold = False

XPrt.CurrentY = prtMinY + prtHeaderHeight + prtlineHeight - XPrt.TextHeight("test")

End Sub

'---------------------------------------------------------
Public Sub prtCompteModif_Line(strDevise As String, strCompte As String, strIntitulé As String, strSituation As String, strExtrait As String, strCourrier As String)
'---------------------------------------------------------
Dim X As String, K As Integer
Dim Situation As String, xSens As String

If XPrt.CurrentY + prtlineHeight > prtMaxY Then
    prtCompteModif_Form_End
    frmElpPrt.prtNewPage
    prtCompteModif_Form
'Else
    'frmElpPrt.prtLineY
End If

XPrt.FontBold = False
'------------------------------------------ligne 1--------------

XPrt.FontSize = 8
XPrt.CurrentX = colP + 100
XPrt.Print strDevise;
XPrt.CurrentX = colP + 450
XPrt.Print strCompte;

XPrt.CurrentX = colP + 1700
XPrt.Print strIntitulé;
XPrt.CurrentX = colA - 600
XPrt.Print strSituation;
XPrt.CurrentX = colA + 100
XPrt.Print strExtrait;
XPrt.CurrentX = colV + 100
XPrt.Print strCourrier;


XPrt.CurrentY = XPrt.CurrentY + prtlineHeight

End Sub









Public Sub prtCompteModif_Form_End()
XPrt.DrawWidth = 2
XPrt.Line (colA, prtMinY)-(colA, prtMaxY)
XPrt.Line (colV, prtMinY)-(colV, prtMaxY)

End Sub

