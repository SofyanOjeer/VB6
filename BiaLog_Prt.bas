Attribute VB_Name = "prtBiaLog"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim X As String, I As Integer, Height8_6 As Integer
Dim V
Dim Nb1 As Integer, Nb2 As Integer

Dim meCompte As typeCompte
Dim meCV1 As typeCV, meCV2 As typeCV, meCV3  As typeCV
Dim X1 As String, X2 As String
Dim mLog_Compte      As String * 11
'---------------------------------------------------------
 Public Sub prtBiaLog_Open(Msg As String)
'---------------------------------------------------------

On Error GoTo prtError

Set XPrt = Printer
Height8_6 = frmElpPrt.prtHeightDelta(8, 6)


prtTitleText = "Journal des événements"

prtLineNb = 1

frmElpPrt.Show vbModeless


prtOrientation = vbPRORLandscape
prtPgmName = "prtBiaLog"
prtTitleUsr = usrName

prtlineHeight = 250
prtHeaderHeight = 300

frmElpPrt.prtStdInit
prtBiaLog_Form

recCompteInit meCompte
mLog_Compte = ""
meCV1 = CV_Euro
meCV1.OpéAmj = DSys
meCV1.Normal = "P"
meCV1.AchatVente = " "
meCV2 = meCV1: meCV3 = meCV1

Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "prtBiaLog_Open")
frmElpPrt.Hide
End Sub
'---------------------------------------------------------
 Public Sub prtBiaLog_Close()
'---------------------------------------------------------
                        
On Error GoTo prtError
        
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY)
XPrt.CurrentY = XPrt.CurrentY + 50
XPrt.FontBold = True
XPrt.CurrentX = prtMinX: XPrt.Print Nb1 & " messages";
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY)

DoEvents: If prtKillDoc Then frmElpPrt.Hide: Exit Sub

frmElpPrt.prtEndDoc
frmElpPrt.Hide

Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "prtBiaLog_Close")
frmElpPrt.Hide
End Sub

'---------------------------------------------------------
Public Sub prtBiaLog_Form()
'---------------------------------------------------------
Dim X As String
XPrt.FontSize = 6

XPrt.FontBold = True
XPrt.DrawWidth = 3
XPrt.CurrentY = prtMinY
prtCurrentY = prtMinY + prtlineHeight
Call frmElpPrt.prtTrame(prtMinX, prtMinY + 20, prtMaxX, prtCurrentY, " ", 230)
XPrt.FontSize = 6
XPrt.CurrentY = prtMinY + (prtHeaderHeight - XPrt.TextHeight(X)) / 2

XPrt.CurrentX = prtMinX + 500: XPrt.Print "Dossier";
XPrt.CurrentX = prtMinX + 2000: XPrt.Print "Compte / Message";
XPrt.CurrentX = prtMinX + 12000: XPrt.Print "  Date Cpt";
XPrt.CurrentX = prtMinX + 12700: XPrt.Print "Service";
XPrt.CurrentX = prtMinX + 13200: XPrt.Print "Profil";
XPrt.CurrentX = prtMinX + 14200: XPrt.Print "Programme";
XPrt.CurrentX = prtMinX + 15100: XPrt.Print "Code Erreur";


'---------------------------------------------------------

XPrt.CurrentY = prtMinY + prtHeaderHeight - XPrt.TextHeight("X")

XPrt.FontSize = 8

End Sub

'---------------------------------------------------------
Public Sub prtBiaLog_Line(lBiaLog As typeBiaLog)
'---------------------------------------------------------
Dim lX As Long, lMax As Long

If XPrt.CurrentY + prtlineHeight * 2.5 > prtMaxY Then
    frmElpPrt.prtNewPage
    prtBiaLog_Form
End If

Nb1 = Nb1 + 1

XPrt.FontSize = 8
XPrt.FontBold = False
'_______________________________________________________________ligne 1-

CV_AttributS lBiaLog.Log_Devise, meCV1
If mLog_Compte <> lBiaLog.Log_Compte Then
    mLog_Compte = lBiaLog.Log_Compte
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight

    If IsNumeric(lBiaLog.Log_Compte) Then
        If lBiaLog.Log_Compte > 0 Then
            XPrt.FontBold = True
            meCompte.Devise = meCV1.DeviseN
            meCompte.Numéro = lBiaLog.Log_Compte
            V = mdbCptP0_Find(meCompte)
            XPrt.CurrentX = prtMinX: XPrt.Print meCV1.DeviseIso;
            XPrt.CurrentX = prtMinX + 500: XPrt.Print Compte_Imp(lBiaLog.Log_Compte);
            XPrt.FontBold = False
            XPrt.CurrentX = prtMinX + 2000: XPrt.Print meCompte.Intitulé;
       End If
    End If
End If
XPrt.FontSize = 6
XPrt.CurrentY = XPrt.CurrentY + Height8_6
''XPrt.CurrentX = prtMinX + 12500: XPrt.Print lBiaLog.Log_CodErr & "_" & Trim(DicLib(524, lBiaLog.Log_CodErr));

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.FontBold = True
XPrt.CurrentX = prtMinX + 500: XPrt.Print lBiaLog.Log_RefCon;
XPrt.FontBold = False

XPrt.CurrentX = prtMinX + 12000: XPrt.Print dateImp(lBiaLog.Log_CptAmj);
XPrt.CurrentX = prtMinX + 12800: XPrt.Print lBiaLog.Log_Servic;
XPrt.CurrentX = prtMinX + 13200: XPrt.Print lBiaLog.Log_Profil;
XPrt.CurrentX = prtMinX + 14200: XPrt.Print lBiaLog.Log_Progr;
XPrt.CurrentX = prtMinX + 15100: XPrt.Print lBiaLog.Log_CodErr;

X1 = Trim(lBiaLog.Log_Texte1)
X2 = Trim(lBiaLog.Log_Texte2)
If X1 = "" Then X1 = X2:       X2 = ""

XPrt.CurrentX = prtMinX + 2000: XPrt.Print X1 & "  ";
If X2 <> "" Then
    lX = XPrt.TextWidth(X2)
    lMax = 12000 - XPrt.CurrentX
    If lX > lMax Then XPrt.CurrentY = XPrt.CurrentY + prtlineHeight: XPrt.CurrentX = prtMinX + 2500
    XPrt.Print X2;
End If

XPrt.CurrentY = XPrt.CurrentY - Height8_6
End Sub

