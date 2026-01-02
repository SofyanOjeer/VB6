Attribute VB_Name = "prtDRH"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Private recCompte As typeCompte
Dim X As String, I As Integer, Height8_6 As Integer

Public P_arrDRH() As typeDRH
Public recDRH As typeDRH
Public CV1 As typeCV

'---------------------------------------------------------
 Public Sub prtDRH_Monitor(Msg As String)
'---------------------------------------------------------
Dim K As Integer, K1 As Integer, K2 As Integer, Kmin As Integer, Kmax As Integer
Dim X As String

On Error GoTo prtError

Set XPrt = Printer
Height8_6 = frmElpPrt.prtHeightDelta(8, 6)
K1 = Val(mId$(Msg, 1, 6))
K2 = Val(mId$(Msg, 7, 6))

prtTitleText = "Liste des salariés"

prtLineNb = 1

frmElpPrt.Show vbModeless


prtOrientation = vbPRORLandscape
prtPgmName = "prtDRH"
prtTitleUsr = usrName

prtlineHeight = 300
prtHeaderHeight = 300

frmElpPrt.prtStdInit

recCompteInit recCompte
recCompte.Société = SocId$
recCompte.Agence = SocAgence$
recCompte.Devise = CV1.DeviseN
recCompte.BiaTyp = "000"
recCompte.BiaNum = "00"
recCompte.Method = "SeekL1"

prtDRH_Form
For K = 1 To arrDRH_NB
    recDRH = arrDRH(K)
    
    prtDRH_Line
    
    DoEvents: If prtKillDoc Then frmElpPrt.Hide: Exit Sub

Next K
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
Public Sub prtDRH_Form()
'---------------------------------------------------------
Dim X As String
prtCurrentY = XPrt.CurrentY
'!!!! Xprt.currenty à définir avant appel proc
XPrt.FontSize = 6
XPrt.FontBold = False
XPrt.CurrentY = prtMaxY + 100
XPrt.CurrentX = prtMaxX - 1200
XPrt.Print Now; Space$(5); XPrt.Page;
XPrt.FontSize = 8

XPrt.FontBold = True
XPrt.DrawWidth = 3

Call frmElpPrt.prtTrame(prtMinX, prtCurrentY, prtMaxX, prtCurrentY + prtlineHeight, " ", 250)

XPrt.DrawWidth = 1

'XPrt.Line (prtMinX, prtCurrentY)-(prtMinX, prtMaxY)
'XPrt.Line (prtMinX, prtMaxY)-(prtMaxX, prtMaxY)

'---------------------------------------------------------

XPrt.CurrentY = prtCurrentY + 50
XPrt.CurrentX = prtMinX + 300: XPrt.Print "Statut";

X = "Serv.": XPrt.CurrentX = prtMinX + 1500 - XPrt.TextWidth(X)
XPrt.Print X;


XPrt.FontSize = 8

End Sub

'---------------------------------------------------------
Public Sub prtDRH_Line()
'---------------------------------------------------------

If XPrt.CurrentY + prtlineHeight * 1.5 > prtMaxY Then
    frmElpPrt.prtNewPage
    XPrt.CurrentY = prtMinX + prtlineHeight * 3
    prtDRH_Form
End If

XPrt.FontBold = False
'_______________________________________________________________ligne 1-

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight

XPrt.CurrentX = prtMinX + 250
'XPrt.Print (recDRH.Statut);

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight


End Sub






