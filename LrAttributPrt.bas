Attribute VB_Name = "prtLrAttribut"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------

Private recLrAttribut As typeLrAttribut
Dim I As Integer
Dim NbImprimé As Integer
Dim kPage As Integer

'---------------------------------------------------------
Public Sub prtLrAttributForm()
'---------------------------------------------------------
Dim X As String
NbImprimé = 0
XPrt.FontSize = 6
XPrt.FontBold = True

XPrt.DrawWidth = 2
prtLrAttributTrait
XPrt.Line (prtMinX, prtMinY + prtHeaderHeight)-(prtMaxX, prtMinY + prtHeaderHeight)

XPrt.DrawWidth = 1


XPrt.Line (prtMinX, prtMinY)-(prtMinX, prtMaxY)
XPrt.Line (prtMaxX, prtMinY)-(prtMaxX, prtMaxY)


'---------------------------------------------------------

X = "Référence"
XPrt.CurrentY = prtMinY + (prtHeaderHeight - XPrt.TextHeight(X)) / 2
XPrt.CurrentX = prtMinX + 20
XPrt.Print X;

XPrt.CurrentX = 1000: XPrt.Print "Intitulé";

Select Case kPage
    Case 1
        XPrt.CurrentX = 3400: XPrt.Print "AFFPU";
        XPrt.CurrentX = 4000: XPrt.Print "AGEMT";
        XPrt.CurrentX = 4600: XPrt.Print "AGENT";
        XPrt.CurrentX = 5200: XPrt.Print "APPAR";
        XPrt.CurrentX = 5800: XPrt.Print "AREFR";
        XPrt.CurrentX = 6400: XPrt.Print "ATTCF";
        XPrt.CurrentX = 7000: XPrt.Print "AUTDV";
        XPrt.CurrentX = 7600: XPrt.Print "BONIF";
        XPrt.CurrentX = 8200: XPrt.Print "CAROB";
        XPrt.CurrentX = 8800: XPrt.Print "CATET";
        XPrt.CurrentX = 9400: XPrt.Print "CDRES";
        XPrt.CurrentX = 10000: XPrt.Print "CDZON";
        XPrt.CurrentX = 10600: XPrt.Print "CLCRC";
        XPrt.CurrentX = 11200: XPrt.Print "COTIT";
        XPrt.CurrentX = 11800: XPrt.Print "CPEMS";
        XPrt.CurrentX = 12400: XPrt.Print "CRDIV";
        XPrt.CurrentX = 13000: XPrt.Print "CREIM";
        XPrt.CurrentX = 13600: XPrt.Print "CREOR";
        XPrt.CurrentX = 14200: XPrt.Print "CRETC";
        XPrt.CurrentX = 14800: XPrt.Print "CRHYP";
        XPrt.CurrentX = 15400: XPrt.Print "DCTOM";
    Case 2
        XPrt.CurrentX = 3400: XPrt.Print "DRAC ";
        XPrt.CurrentX = 4000: XPrt.Print "DURIN";
        XPrt.CurrentX = 4600: XPrt.Print "DUROM";
        XPrt.CurrentX = 5200: XPrt.Print "DVOPR";
        XPrt.CurrentX = 5800: XPrt.Print "ECART";
        XPrt.CurrentX = 6400: XPrt.Print "ECFIN";
        XPrt.CurrentX = 7000: XPrt.Print "ELIGB";
        XPrt.CurrentX = 7600: XPrt.Print "FAMDV";
        XPrt.CurrentX = 8200: XPrt.Print "FOPIF";
        XPrt.CurrentX = 8800: XPrt.Print "FPRBG";
        XPrt.CurrentX = 9400: XPrt.Print "GARCF";
        XPrt.CurrentX = 10000: XPrt.Print "MLFCE";
        XPrt.CurrentX = 10600: XPrt.Print "MONDV";
        XPrt.CurrentX = 11200: XPrt.Print "MUTFG";
        XPrt.CurrentX = 11800: XPrt.Print "NACGA";
        XPrt.CurrentX = 12400: XPrt.Print "NACGR";
        XPrt.CurrentX = 13000: XPrt.Print "NACPS";
        XPrt.CurrentX = 13600: XPrt.Print "NAEGA";
        XPrt.CurrentX = 14200: XPrt.Print "NAIMO";
        XPrt.CurrentX = 14800: XPrt.Print "NAOCB";
        XPrt.CurrentX = 15400: XPrt.Print "NAPRO";
    Case 3
        XPrt.CurrentX = 3400: XPrt.Print "NARCP";
        XPrt.CurrentX = 4000: XPrt.Print "NATCP";
        XPrt.CurrentX = 4600: XPrt.Print "NATCR";
        XPrt.CurrentX = 5200: XPrt.Print "NATCS";
        XPrt.CurrentX = 5800: XPrt.Print "NATDD";
        XPrt.CurrentX = 6400: XPrt.Print "NATER";
        XPrt.CurrentX = 7000: XPrt.Print "NATIF";
        XPrt.CurrentX = 7600: XPrt.Print "NATIT";
        XPrt.CurrentX = 8200: XPrt.Print "NATMA";
        XPrt.CurrentX = 8800: XPrt.Print "NATOF";
        XPrt.CurrentX = 9400: XPrt.Print "NATRS";
        XPrt.CurrentX = 10000: XPrt.Print "NRAST";
        XPrt.CurrentX = 10600: XPrt.Print "NREHB";
        XPrt.CurrentX = 11200: XPrt.Print "OPCVM";
        XPrt.CurrentX = 11800: XPrt.Print "OPEFC";
        XPrt.CurrentX = 12400: XPrt.Print "OPFDH";
        XPrt.CurrentX = 13000: XPrt.Print "OPREC";
        XPrt.CurrentX = 13600: XPrt.Print "PAACT";
        XPrt.CurrentX = 14200: XPrt.Print "PERIO";
        XPrt.CurrentX = 14800: XPrt.Print "PRIMP";
        XPrt.CurrentX = 15400: XPrt.Print "PROCB";
    Case 4
        XPrt.CurrentX = 3400: XPrt.Print "REDES";
        XPrt.CurrentX = 4000: XPrt.Print "REDHB";
        XPrt.CurrentX = 4600: XPrt.Print "RESET";
        XPrt.CurrentX = 5200: XPrt.Print "REZON";
        XPrt.CurrentX = 5800: XPrt.Print "RISPA";
        XPrt.CurrentX = 6400: XPrt.Print "SENOP";
        XPrt.CurrentX = 7000: XPrt.Print "TCFPE";
        XPrt.CurrentX = 7600: XPrt.Print "TOPIF";
        XPrt.CurrentX = 8200: XPrt.Print "TYCGR";
        XPrt.CurrentX = 8800: XPrt.Print "TYCOM";
        XPrt.CurrentX = 9400: XPrt.Print "TYDSU";
        XPrt.CurrentX = 10000: XPrt.Print "TYETS";
        XPrt.CurrentX = 10600: XPrt.Print "TYPOR";
        XPrt.CurrentX = 11200: XPrt.Print "TYPSU";
        XPrt.CurrentX = 11800: XPrt.Print "TYRES";
        XPrt.CurrentX = 12400: XPrt.Print "ZACTI";
        XPrt.CurrentX = 13000: XPrt.Print "ZAGDT";
        XPrt.CurrentX = 13600: XPrt.Print "REESC1";
        XPrt.CurrentX = 14200: XPrt.Print "REESC6";
    Case 5
        XPrt.CurrentX = 3400: XPrt.Print "CDCPCO/JO";
'       XPrt.CurrentX = 4000: XPrt.Print "CDCPJO";
        XPrt.CurrentX = 4000: XPrt.Print "CDCPFU";
        XPrt.CurrentX = 4600: XPrt.Print "CDAGCO";
        XPrt.CurrentX = 5200: XPrt.Print "CDREME";
        XPrt.CurrentX = 5800: XPrt.Print "TYMTDV";
        XPrt.CurrentX = 6400: XPrt.Print "TYVENT";
        XPrt.CurrentX = 7000: XPrt.Print "CRVENT";
        XPrt.CurrentX = 7600: XPrt.Print "CDDURE";
        XPrt.CurrentX = 8200: XPrt.Print "DUINIT";
        XPrt.CurrentX = 8800: XPrt.Print "CDCRTI";
        XPrt.CurrentX = 9400: XPrt.Print "CDCRAC";
        XPrt.CurrentX = 10000: XPrt.Print "CDBIOR";
        XPrt.CurrentX = 10600: XPrt.Print "CDDEIN";
        XPrt.CurrentX = 11200: XPrt.Print "CDCRIM";
        XPrt.CurrentX = 11800: XPrt.Print "CDCRCO";
        XPrt.CurrentX = 12400: XPrt.Print "CDCREF";
        XPrt.CurrentX = 13000: XPrt.Print "CDLODA";
        XPrt.CurrentX = 13600: XPrt.Print "CDCRET";
        XPrt.CurrentX = 14200: XPrt.Print "CDOMPO";
        XPrt.CurrentX = 14800: XPrt.Print "CDOPIM";
        XPrt.CurrentX = 15400: XPrt.Print "CDSWAP";
End Select
XPrt.CurrentY = prtMinY + prtHeaderHeight - XPrt.TextHeight("X")

End Sub

'---------------------------------------------------------
 Public Sub prtLrAttributX(Msg As String)
'---------------------------------------------------------
Dim K As Integer, K1 As Integer, K2 As Integer
Dim X As String
Dim Kmin As Integer, Kmax As Integer

K1 = Val(Mid$(Msg, 1, 6))
K2 = Val(Mid$(Msg, 7, 6))
Kmin = 1: Kmax = 5
Select Case Mid$(Msg, 13, 1)
    Case "B": Kmax = 4
    Case "R": Kmin = 5
End Select

Set XPrt = Printer
frmElpPrt.Show vbModeless

prtOrientation = vbPRORLandscape
prtTitleText = "Liste des attributs (Lr97 - BAFI)"
prtPgmName = "prtLrAttribut"
prtTitleUsr = usrName

prtLineNb = 1
prtlineHeight = 300
prtHeaderHeight = 300

frmElpPrt.prtStdinit

For kPage = Kmin To Kmax
    prtLrAttributForm
    
    For K = K1 To K2
        recLrAttribut = arrLrAttribut(K)
        prtLrAttributLine
        
        DoEvents: If prtKillDoc Then frmElpPrt.Hide: Exit Sub
    Next K
    If kPage <> 5 Then
        frmElpPrt.prtNewPage
    Else
        frmElpPrt.prtEndDoc
    End If
Next kPage

frmElpPrt.Hide
End Sub




'---------------------------------------------------------
Public Sub prtLrAttributLine()
'---------------------------------------------------------
Dim X As String, K As Integer
Dim Situation As String

If XPrt.CurrentY + prtlineHeight * 2 > prtMaxY Then
   frmElpPrt.prtNewPage
   prtLrAttributForm
End If
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
NbImprimé = NbImprimé + 1
If NbImprimé = 4 Then
'   Call frmElpPrt.prtTrame(prtMinX + 20, XPrt.CurrentY, prtMaxX - 20, XPrt.CurrentY + prtlineHeight, recLrAttribut. recLrAttribut., 255)
'   XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.Line (prtMinX, XPrt.CurrentY - 50)-(prtMaxX, XPrt.CurrentY - 50)
    XPrt.CurrentY = XPrt.CurrentY + 50
    NbImprimé = 1
End If

XPrt.CurrentX = prtMinX + 10
XPrt.Print recLrAttribut.Référence;

XPrt.CurrentX = 1000: XPrt.Print DicLib(13, Trim(recLrAttribut.Référence));

Select Case kPage
    Case 1
        XPrt.CurrentX = 3400: XPrt.Print recLrAttribut.AFFPU;
        XPrt.CurrentX = 4000: XPrt.Print recLrAttribut.AGEMT;
        XPrt.CurrentX = 4600: XPrt.Print recLrAttribut.AGENT;
        XPrt.CurrentX = 5200: XPrt.Print recLrAttribut.APPAR;
        XPrt.CurrentX = 5800: XPrt.Print recLrAttribut.AREFR;
        XPrt.CurrentX = 6400: XPrt.Print recLrAttribut.ATTCF;
        XPrt.CurrentX = 7000: XPrt.Print recLrAttribut.AUTDV;
        XPrt.CurrentX = 7600: XPrt.Print recLrAttribut.BONIF;
        XPrt.CurrentX = 8200: XPrt.Print recLrAttribut.CAROB;
        XPrt.CurrentX = 8800: XPrt.Print recLrAttribut.CATET;
        XPrt.CurrentX = 9400: XPrt.Print recLrAttribut.CDRES;
        XPrt.CurrentX = 10000: XPrt.Print recLrAttribut.CDZON;
        XPrt.CurrentX = 10600: XPrt.Print recLrAttribut.CLCRC;
        XPrt.CurrentX = 11200: XPrt.Print recLrAttribut.COTIT;
        XPrt.CurrentX = 11800: XPrt.Print recLrAttribut.CPEMS;
        XPrt.CurrentX = 12400: XPrt.Print recLrAttribut.CRDIV;
        XPrt.CurrentX = 13000: XPrt.Print recLrAttribut.CREIM;
        XPrt.CurrentX = 13600: XPrt.Print recLrAttribut.CREOR;
        XPrt.CurrentX = 14200: XPrt.Print recLrAttribut.CRETC;
        XPrt.CurrentX = 14800: XPrt.Print recLrAttribut.CRHYP;
        XPrt.CurrentX = 15400: XPrt.Print recLrAttribut.DCTOM;
    Case 2
        XPrt.CurrentX = 3400: XPrt.Print recLrAttribut.DRAC;
        XPrt.CurrentX = 4000: XPrt.Print recLrAttribut.DURIN;
        XPrt.CurrentX = 4600: XPrt.Print recLrAttribut.DUROM;
        XPrt.CurrentX = 5200: XPrt.Print recLrAttribut.DVOPR;
        XPrt.CurrentX = 5800: XPrt.Print recLrAttribut.ECART;
        XPrt.CurrentX = 6400: XPrt.Print recLrAttribut.ECFIN;
        XPrt.CurrentX = 7000: XPrt.Print recLrAttribut.ELIGB;
        XPrt.CurrentX = 7600: XPrt.Print recLrAttribut.FAMDV;
        XPrt.CurrentX = 8200: XPrt.Print recLrAttribut.FOPIF;
        XPrt.CurrentX = 8800: XPrt.Print recLrAttribut.FPRBG;
        XPrt.CurrentX = 9400: XPrt.Print recLrAttribut.GARCF;
        XPrt.CurrentX = 10000: XPrt.Print recLrAttribut.MLFCE;
        XPrt.CurrentX = 10600: XPrt.Print recLrAttribut.MONDV;
        XPrt.CurrentX = 11200: XPrt.Print recLrAttribut.MUTFG;
        XPrt.CurrentX = 11800: XPrt.Print recLrAttribut.NACGA;
        XPrt.CurrentX = 12400: XPrt.Print recLrAttribut.NACGR;
        XPrt.CurrentX = 13000: XPrt.Print recLrAttribut.NACPS;
        XPrt.CurrentX = 13600: XPrt.Print recLrAttribut.NAEGA;
        XPrt.CurrentX = 14200: XPrt.Print recLrAttribut.NAIMO;
        XPrt.CurrentX = 14800: XPrt.Print recLrAttribut.NAOCB;
        XPrt.CurrentX = 15400: XPrt.Print recLrAttribut.NAPRO;
    Case 3
        XPrt.CurrentX = 3400: XPrt.Print recLrAttribut.NARCP;
        XPrt.CurrentX = 4000: XPrt.Print recLrAttribut.NATCP;
        XPrt.CurrentX = 4600: XPrt.Print recLrAttribut.NATCR;
        XPrt.CurrentX = 5200: XPrt.Print recLrAttribut.NATCS;
        XPrt.CurrentX = 5800: XPrt.Print recLrAttribut.NATDD;
        XPrt.CurrentX = 6400: XPrt.Print recLrAttribut.NATER;
        XPrt.CurrentX = 7000: XPrt.Print recLrAttribut.NATIF;
        XPrt.CurrentX = 7600: XPrt.Print recLrAttribut.NATIT;
        XPrt.CurrentX = 8200: XPrt.Print recLrAttribut.NATMA;
        XPrt.CurrentX = 8800: XPrt.Print recLrAttribut.NATOF;
        XPrt.CurrentX = 9400: XPrt.Print recLrAttribut.NATRS;
        XPrt.CurrentX = 10000: XPrt.Print recLrAttribut.NRAST;
        XPrt.CurrentX = 10600: XPrt.Print recLrAttribut.NREHB;
        XPrt.CurrentX = 11200: XPrt.Print recLrAttribut.OPCVM;
        XPrt.CurrentX = 11800: XPrt.Print recLrAttribut.OPEFC;
        XPrt.CurrentX = 12400: XPrt.Print recLrAttribut.OPFDH;
        XPrt.CurrentX = 13000: XPrt.Print recLrAttribut.OPREC;
        XPrt.CurrentX = 13600: XPrt.Print recLrAttribut.PAACT;
        XPrt.CurrentX = 14200: XPrt.Print recLrAttribut.PERIO;
        XPrt.CurrentX = 14800: XPrt.Print recLrAttribut.PRIMP;
        XPrt.CurrentX = 15400: XPrt.Print recLrAttribut.PROCB;
    Case 4
        XPrt.CurrentX = 3400: XPrt.Print recLrAttribut.REDES;
        XPrt.CurrentX = 4000: XPrt.Print recLrAttribut.REDHB;
        XPrt.CurrentX = 4600: XPrt.Print recLrAttribut.RESET;
        XPrt.CurrentX = 5200: XPrt.Print recLrAttribut.REZON;
        XPrt.CurrentX = 5800: XPrt.Print recLrAttribut.RISPA;
        XPrt.CurrentX = 6400: XPrt.Print recLrAttribut.SENOP;
        XPrt.CurrentX = 7000: XPrt.Print recLrAttribut.TCFPE;
        XPrt.CurrentX = 7600: XPrt.Print recLrAttribut.TOPIF;
        XPrt.CurrentX = 8200: XPrt.Print recLrAttribut.TYCGR;
        XPrt.CurrentX = 8800: XPrt.Print recLrAttribut.TYCOM;
        XPrt.CurrentX = 9400: XPrt.Print recLrAttribut.TYDSU;
        XPrt.CurrentX = 10000: XPrt.Print recLrAttribut.TYETS;
        XPrt.CurrentX = 10600: XPrt.Print recLrAttribut.TYPOR;
        XPrt.CurrentX = 11200: XPrt.Print recLrAttribut.TYPSU;
        XPrt.CurrentX = 11800: XPrt.Print recLrAttribut.TYRES;
        XPrt.CurrentX = 12400: XPrt.Print recLrAttribut.ZACTI;
        XPrt.CurrentX = 13000: XPrt.Print recLrAttribut.ZAGDT;
        XPrt.CurrentX = 13600: XPrt.Print recLrAttribut.REESC1;
        XPrt.CurrentX = 14200: XPrt.Print recLrAttribut.REESC6;
   Case 5
        XPrt.CurrentX = 3400: XPrt.Print recLrAttribut.CDCPCO & "-"; recLrAttribut.CDCPJO;
'       XPrt.CurrentX = 4000: XPrt.Print recLrAttribut.CDCPJO;
        XPrt.CurrentX = 4000: XPrt.Print recLrAttribut.CDCPFU;
        XPrt.CurrentX = 4600: XPrt.Print recLrAttribut.CDAGCO;
        XPrt.CurrentX = 5200: XPrt.Print recLrAttribut.CDREME;
        XPrt.CurrentX = 5800: XPrt.Print recLrAttribut.TYMTDV;
        XPrt.CurrentX = 6400: XPrt.Print recLrAttribut.TYVENT;
        XPrt.CurrentX = 7000: XPrt.Print recLrAttribut.CRVENT;
        XPrt.CurrentX = 7600: XPrt.Print recLrAttribut.CDDURE;
        XPrt.CurrentX = 8200: XPrt.Print recLrAttribut.DUINIT;
        XPrt.CurrentX = 8800: XPrt.Print recLrAttribut.CDCRTI;
        XPrt.CurrentX = 9400: XPrt.Print recLrAttribut.CDCRAC;
        XPrt.CurrentX = 10000: XPrt.Print recLrAttribut.CDBIOR;
        XPrt.CurrentX = 10600: XPrt.Print recLrAttribut.CDDEIN;
        XPrt.CurrentX = 11200: XPrt.Print recLrAttribut.CDCRIM;
        XPrt.CurrentX = 11800: XPrt.Print recLrAttribut.CDCRCO;
        XPrt.CurrentX = 12400: XPrt.Print recLrAttribut.CDCREF;
        XPrt.CurrentX = 13000: XPrt.Print recLrAttribut.CDLODA;
        XPrt.CurrentX = 13600: XPrt.Print recLrAttribut.CDCRET;
        XPrt.CurrentX = 14200: XPrt.Print recLrAttribut.CDOMPO;
        XPrt.CurrentX = 14800: XPrt.Print recLrAttribut.CDOPIM;
        XPrt.CurrentX = 15400: XPrt.Print recLrAttribut.CDSWAP;
End Select

End Sub






Public Sub prtLrAttributTrait()

Call frmElpPrt.prtTrame(3400 - 30, prtMinY + 10, 5200 - 30, prtMaxY - 10, " ")
Call frmElpPrt.prtTrame(7000 - 30, prtMinY + 10, 8800 - 30, prtMaxY - 10, " ")
Call frmElpPrt.prtTrame(10600 - 30, prtMinY + 10, 12400 - 30, prtMaxY - 10, " ")
Call frmElpPrt.prtTrame(14200 - 30, prtMinY + 10, 16000 - 30, prtMaxY - 10, " ")

End Sub

