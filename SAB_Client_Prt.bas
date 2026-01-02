Attribute VB_Name = "prtSAB_Client"
Option Explicit

Dim Height8_6 As Integer
 
Type typeWECHISB0
    ECHISBCOM As String
    ECHISBCDM As Currency
    ECHISBICR As Currency
    ECHISBIDE As Currency
    ECHISBPRE As Currency
    ECHISBPFD As Currency
    ECHISBTDC As Currency
    
    ECHISBCDM_AUT As String
    ECHISBICR_AUT As String
    ECHISBIDE_AUT As String
    ECHISBPRE_AUT As String
    ECHISBPFD_AUT As String
    ECHISBTDC_AUT As String

End Type
    

Type typeXgsop
    X           As Integer
    y           As Integer
    CLIENACLI   As String
    CLIENAETA   As String
    CLIENACAT   As String
    CLIENARA1   As String
    CLIENARA2   As String
    CLIENARES   As String
    CLIENARSD   As String
    CLIENANAT   As String
    CLIENACOL   As Integer
    CAV_Client  As Integer
    CAV_Tiers   As Integer
    CAV_Clos    As Integer
    Tech_Client As Integer
    Tech_Tiers  As Integer
    Tech_Clos   As Integer
    COMPTECLO   As Long
    PLANCOPRO As String
End Type

Public Sub prtSAB_Client_Liste_Close()
Dim X As String
On Error GoTo prtError
XPrt.FontBold = False

prtSAB_Client_Liste_Colonne
Call frmElpPrt.prtEndDoc(1000)
frmElpPrt.Hide
Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub


Public Sub prtSAB_Client_Liste_Form()
Dim wId As String
Dim X As String

XPrt.FontSize = 7
XPrt.FontBold = True
XPrt.DrawWidth = 7
'XPrt.Line (prtMaxX, prtMinY)-(prtMaxX, prtMaxY)

XPrt.CurrentY = prtMinY
Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY, prtMaxX, prtMinY + prtHeaderHeight, "B", 245)

XPrt.CurrentY = prtMinY + 50

XPrt.CurrentX = prtMinX
XPrt.Print "Racine";
XPrt.CurrentX = prtMinX + 1000
XPrt.Print "Mandant";
XPrt.CurrentX = prtMinX + 3000
XPrt.Print "Lien";
XPrt.CurrentX = prtMinX + 3700
XPrt.Print "Racine";

XPrt.CurrentX = prtMinX + 4700
XPrt.Print "Mandataire";

XPrt.CurrentX = prtMinX + 6500
XPrt.Print "Produit";
XPrt.CurrentX = prtMinX + 7000
XPrt.Print "Compte";

XPrt.CurrentX = prtMinX + 9000
XPrt.Print "Intitulé";



'XPrt.FontSize = 8
XPrt.FontBold = False

XPrt.CurrentY = prtMinX + prtHeaderHeight + 100


End Sub
Public Sub prtSAB_Client_Liste_Colonne()
Dim wId As String
Dim X As String

XPrt.DrawWidth = 3
XPrt.Line (prtMinX + 2900, prtMinY)-(prtMinX + 2900, prtMaxY), prtLineColor
XPrt.Line (prtMinX + 6400, prtMinY)-(prtMinX + 6400, prtMaxY), prtLineColor

End Sub





Public Sub prtSAB_Client_Liste_Line(larrX() As String)
Dim X As String, curX As Currency



    XPrt.CurrentX = prtMinX + 50: XPrt.Print larrX(0);
    XPrt.CurrentX = prtMinX + 1000: XPrt.Print larrX(1);
    XPrt.CurrentX = prtMinX + 3000: XPrt.Print larrX(2);
    XPrt.CurrentX = prtMinX + 3700: XPrt.Print larrX(3);
    XPrt.CurrentX = prtMinX + 4700: XPrt.Print larrX(4);
    XPrt.CurrentX = prtMinX + 6500: XPrt.Print larrX(5);
    XPrt.CurrentX = prtMinX + 7200: XPrt.Print larrX(6);
    XPrt.CurrentX = prtMinX + 9000: XPrt.Print larrX(7);
   
    XPrt.FontBold = False

End Sub

Public Sub prtSAB_Client_Liste_NewLine()
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
If XPrt.CurrentY + 300 > prtMaxY Then
    prtSAB_Client_Liste_Colonne
    frmElpPrt.prtNewPage
    prtSAB_Client_Liste_Form
End If

End Sub



Public Sub prtSAB_Client_Liste_Open()
On Error GoTo prtError

Set XPrt = Printer
frmElpPrt.Show vbModeless

Height8_6 = frmElpPrt.prtHeightDelta(8, 6)

blnFiligrane = False
prtOrientation = vbPRORPortrait 'Landscape '
prtPgmName = "prtCHQ_SCAN"
prtTitleUsr = usrName
prtTitleText = "Liste des Mandants / Mandataires / Comptes"

prtLineNb = 1
prtlineHeight = 300
prtHeaderHeight = 400


prtFormType = ""
frmElpPrt.prtStdInit
prtSAB_Client_Liste_Form


Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub









Public Sub typeXgsop_Init(lXgsop As typeXgsop)
lXgsop.X = 0
lXgsop.y = 0
lXgsop.CLIENACLI = ""
lXgsop.CLIENAETA = ""
lXgsop.CLIENACAT = ""
lXgsop.CLIENARA1 = ""
lXgsop.CLIENARA2 = ""
lXgsop.CLIENARES = ""
lXgsop.CLIENARSD = ""
lXgsop.CLIENANAT = ""
lXgsop.PLANCOPRO = ""
lXgsop.CLIENACOL = 0
lXgsop.CAV_Client = 0
lXgsop.CAV_Tiers = 0
lXgsop.CAV_Clos = 0
lXgsop.Tech_Client = 0
lXgsop.Tech_Tiers = 0
lXgsop.Tech_Clos = 0
lXgsop.COMPTECLO = 0
End Sub
