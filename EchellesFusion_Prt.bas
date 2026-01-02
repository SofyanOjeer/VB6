Attribute VB_Name = "prtEchellesFusion"
Option Explicit
Dim Dev1 As typeDevise, Dev2 As typeDevise
Private recEchellesFusion As typeEchellesFusion, mEchellesFusion As typeEchellesFusion
Dim I As Integer, mCurrenty As Integer
Dim colP As Integer, colA As Integer, colV As Integer, colX As Integer
Dim blnList As Boolean
Public sortEchellesfusion()
Dim recCptInfo As typeCptInfo
Dim CV1 As typeCV, CV2 As typeCV, CV3 As typeCV

Dim V, Height8_6 As Integer
'---------------------------------------------------------
 Public Sub prtEchellesFusionX(Msg As String, Text As String)
'---------------------------------------------------------
Dim K As Integer, K1 As Integer, K2 As Integer
Dim X As String

Set XPrt = Printer
Height8_6 = frmElpPrt.prtHeightDelta(8, 6)
K1 = Val(Mid$(Msg, 1, 6))
K2 = Val(Mid$(Msg, 7, 6))


frmElpPrt.Show vbModeless

recEchellesFusion_Init recEchellesFusion
prtOrientation = vbPRORLandscape
prtTitleText = "Echelles fusionnées : liste des comptes" & Text
prtPgmName = "prtEchellesFusion"
prtTitleUsr = usrName

prtLineNb = 1
prtlineHeight = 250
prtHeaderHeight = 300

frmElpPrt.prtStdinit
CV_Init CV1: CV2 = CV1: CV3 = CV1
colP = prtMinX
colA = prtMinX + 9000
colV = prtMinX + 13500
colX = prtMaxX
Dev1.DevX = arrEchellesFusion(1).DeviseOrigine
prtEchellesFusion_Form
If K2 > 0 Then
    mEchellesFusion = arrEchellesFusion(sortEchellesfusion(1))
Else
    recEchellesFusion_Init mEchellesFusion
End If


For K = K1 To K2
    recEchellesFusion = arrEchellesFusion(sortEchellesfusion(K))
    
    If recEchellesFusion.DeviseFusion <> mEchellesFusion.DeviseFusion _
    Or recEchellesFusion.CompteFusion <> mEchellesFusion.CompteFusion Then
        prtEchellesFusion_Line_End
        mEchellesFusion = recEchellesFusion
    End If
    
    prtEchellesFusion_Line recEchellesFusion
    DoEvents: If prtKillDoc Then frmElpPrt.Hide: Exit Sub

Next K
        
prtEchellesFusion_Line_End
prtEchellesFusion_Form_End
frmElpPrt.prtEndDoc
frmElpPrt.Hide

End Sub
'---------------------------------------------------------
Public Sub prtEchellesFusion_Form()
'---------------------------------------------------------
Dim X As String, K As Integer

XPrt.FontSize = 8
XPrt.FontBold = False

XPrt.DrawWidth = 3
Call frmElpPrt.prtTrame(prtMinX, prtMinY, prtMaxX, prtMinY + prtHeaderHeight, "B")

XPrt.Line (prtMinX, prtMinY)-(prtMinX, prtMaxY)
XPrt.Line (prtMaxX, prtMinY)-(prtMaxX, prtMaxY)
K = prtMinY + prtHeaderHeight + 10
'Call frmElpPrt.prtTrame(colP + 1600, K, colP + 3000, prtMaxY - 10, " ")
'Call frmElpPrt.prtTrame(colA + 1100, K, colA + 2100, prtMaxY - 10, " ")
'Call frmElpPrt.prtTrame(colV + 1100, K, colV + 2100, prtMaxY - 10, " ")

'---------------------------------------------------------
XPrt.FontBold = True
XPrt.CurrentY = prtMinY + 50

XPrt.CurrentX = prtMinX + 50
XPrt.Print "Compte";
XPrt.CurrentX = colA - 1400
XPrt.Print "Fusion depuis le";

X = "Solde en valeur"
XPrt.CurrentX = (colA + colV - XPrt.TextWidth(X)) / 2: XPrt.Print X;
X = "Contre-valeur"
XPrt.CurrentX = (colV + colX - XPrt.TextWidth(X)) / 2: XPrt.Print X;
XPrt.FontBold = False

XPrt.CurrentY = prtMinY + prtHeaderHeight + prtlineHeight - XPrt.TextHeight("test")

End Sub

'---------------------------------------------------------
Public Sub prtEchellesFusion_Line(recEchellesFusion As typeEchellesFusion)
'---------------------------------------------------------
Dim X As String, K As Integer
Dim Situation As String, xSens As String

If prtCurrentY + prtParagraphHeight > prtMaxY Then
    prtEchellesFusion_Form_End
    frmElpPrt.prtNewPage
    prtEchellesFusion_Form
'Else
    'frmElpPrt.prtLineY
End If

XPrt.FontBold = False
'------------------------------------------ligne 1--------------

XPrt.FontSize = 8

XPrt.CurrentX = colP + 50
recCptInfoInit recCptInfo
recCptInfo.Method = "JoinL1"
recCptInfo.Société = SocId$
recCptInfo.Agence = SocAgence$
recCptInfo.Devise = recEchellesFusion.DeviseOrigine
recCptInfo.Numéro = recEchellesFusion.CompteOrigine
V = srvCptInfoMon(recCptInfo)
If Not IsNull(V) Then
    XPrt.Print "Compte inexistant : " & recCptInfo.Devise & "." & recCptInfo.Numéro
Else
    CV1.DeviseIso = ""
    CV2.DeviseIso = ""
    CV1.DeviseN = recEchellesFusion.DeviseOrigine
    CV2.DeviseN = recEchellesFusion.DeviseFusion
    CV1.Montant = recCptInfo.EchelleSolde
    Call CV_Transitoire(CV1, CV2, CV3, X)
    
    
    XPrt.Print recEchellesFusion.DeviseOrigine & "." & Compte_Imp(recEchellesFusion.CompteOrigine);
    XPrt.CurrentX = colP + 1700
    XPrt.Print recCptInfo.Intitulé;
    XPrt.CurrentX = colA - 1100
    XPrt.Print dateImp(recEchellesFusion.AmjDébut);

    X = Format$(recCptInfo.EchelleSolde, "## ### ### ### ### ##0.00")

    XPrt.CurrentX = colA + 2600 - XPrt.TextWidth(X)
    XPrt.Print X;
   
    XPrt.CurrentX = colA + 2800
    XPrt.Print CV1.DeviseIso;
    XPrt.CurrentX = colA + 3300
    XPrt.Print dateImp(recCptInfo.EchelleAmj);

     X = Format$(CV2.Montant, "## ### ### ### ### ##0.00")
    XPrt.CurrentX = colV + 1800 - XPrt.TextWidth(X)
    XPrt.Print X;

    XPrt.CurrentX = colV + 1950
    XPrt.Print CV2.DeviseIso;

    If recCptInfo.EchelleSolde < 0 Then
        xSens = " db"
    Else
        If recCptInfo.EchelleSolde = 0 Then
            xSens = "   "
        Else
            xSens = " cr"
        End If
    End If
    XPrt.FontSize = 6
    XPrt.CurrentY = XPrt.CurrentY + Height8_6
    XPrt.CurrentX = colA + 2620
    XPrt.Print xSens;
    XPrt.CurrentX = colV + 1820
    XPrt.Print xSens;

    If recCptInfo.Echelle <> "1" Then
        XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
        XPrt.CurrentX = colP + 50
        XPrt.FontBold = True
        XPrt.Print "!!! L'indicateur [ COMPTE soumis à ECHELLE ] n'est pas en fonction ";
    End If
End If

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight

End Sub









Public Sub prtEchellesFusion_Line_End()

mEchellesFusion.DeviseOrigine = mEchellesFusion.DeviseFusion
mEchellesFusion.CompteOrigine = mEchellesFusion.CompteFusion
mEchellesFusion.AmjDébut = "00000000"
Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY + 50, prtMaxX, XPrt.CurrentY + 50 + prtlineHeight, "B")
XPrt.CurrentY = XPrt.CurrentY + 100
prtEchellesFusion_Line mEchellesFusion
XPrt.CurrentY = XPrt.CurrentY + 100
End Sub

Public Sub prtEchellesFusion_Form_End()
XPrt.DrawWidth = 2
XPrt.Line (colA, prtMinY)-(colA, prtMaxY)
XPrt.Line (colV, prtMinY)-(colV, prtMaxY)

End Sub
