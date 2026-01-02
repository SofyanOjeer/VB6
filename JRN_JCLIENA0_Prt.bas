Attribute VB_Name = "prtJRN_JCLIENA0"
Option Explicit
Dim mFct1 As String

Dim X As String, I As Integer, Height8_6 As Integer
'Dim curX As Currency, curX1 As Currency, curX2 As Currency

Dim blnDossier_Line As Boolean

Dim mJCLIENA0 As typeJCLIENA0, xJCLIENA0 As typeJCLIENA0

Dim blnFontBold_True As Boolean, blnFontBold_false As Boolean
Public Sub prtJRN_JCLIENA0_Line(lJCLIENA0 As typeJCLIENA0, lJRNENT0 As typeJRNENT0)
Dim X As String
Dim wFontBold_1 As Boolean, wFontBold_2 As Boolean
Dim blnOk  As Boolean
Dim blnNew As Boolean

On Error Resume Next
XPrt.FontSize = 7
xJCLIENA0 = lJCLIENA0
prtJRN_JCLIENA0_NewLine

XPrt.FontBold = blnFontBold_True
XPrt.CurrentX = prtMinX + 50: XPrt.Print lJCLIENA0.CLIENACLI;
XPrt.FontBold = blnFontBold_false
XPrt.CurrentX = prtMinX + 2000: XPrt.Print lJCLIENA0.CLIENARA1;
XPrt.CurrentX = prtMinX + 7000: XPrt.Print lJCLIENA0.CLIENAETA;
XPrt.CurrentX = prtMinX + 8000: XPrt.Print lJCLIENA0.CLIENANAT;

XPrt.FontBold = blnFontBold_false
wFontBold_1 = blnFontBold_false
Select Case lJRNENT0.JOENTT
    Case "UB": X = "*"
    Case "UP": X = "Màj":  wFontBold_1 = blnFontBold_True
    Case "PX", "PT": X = "Cre"
    Case "DL": X = "Sup"
    Case Else: X = lJRNENT0.JOENTT
End Select
XPrt.CurrentX = prtMinX + 1600: XPrt.Print X;
XPrt.CurrentX = prtMinX + 13300: XPrt.Print dateJma6_Imp10(lJRNENT0.JODATE);
XPrt.CurrentX = prtMinX + 14000: XPrt.Print timeImp(lJRNENT0.JOTIME);
XPrt.CurrentX = prtMinX + 14800: XPrt.Print lJRNENT0.JOUSER;


mJCLIENA0 = lJCLIENA0
'=======================
End Sub

Public Sub prtJRN_JCLIENA0_Open(lText As String)
On Error GoTo prtError

Set XPrt = Printer
frmElpPrt.Show vbModeless

Height8_6 = frmElpPrt.prtHeightDelta(8, 6)

blnFiligrane = False
 prtOrientation = vbPRORLandscape '
prtPgmName = "prtJRN_JCLIENA0"
prtTitleUsr = usrName
prtTitleText = lText


If Trim(Printer.Devicename) = "Easy PDF Creator" Then
    blnFontBold_True = False
    blnFontBold_false = True
Else
    blnFontBold_True = True
    blnFontBold_false = False
End If

prtLineNb = 1
prtlineHeight = 250
prtHeaderHeight = 300


prtFormType = ""
frmElpPrt.prtStdInit

prtFontName = prtFontName_Arial
prtJRN_JCLIENA0_Form
XPrt.CurrentY = XPrt.CurrentY - prtlineHeight

srvJCLIENA0_Init mJCLIENA0


Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub


Public Sub prtJRN_JCLIENA0_Close()
On Error GoTo prtError

XPrt.DrawWidth = 5

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY), prtLineColor

frmElpPrt.prtEndDoc 1000
frmElpPrt.Hide
Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub


Public Sub prtJRN_JCLIENA0_NewLine()
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
If XPrt.CurrentY + 300 > prtMaxY Then
    frmElpPrt.prtNewPage
    prtJRN_JCLIENA0_Form
End If

End Sub




Public Sub prtJRN_JCLIENA0_Form()
Dim wId As String
Dim X As String

XPrt.FontSize = 7
XPrt.FontBold = blnFontBold_True
XPrt.DrawWidth = 5

XPrt.CurrentY = prtMinY + 50
Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY - 20, prtMaxX, XPrt.CurrentY + prtlineHeight * 2 - 100, " ", 235)

XPrt.CurrentX = prtMinX + 50: XPrt.Print "Client";
XPrt.CurrentX = prtMinX + 2000: XPrt.Print "Intitulé";
XPrt.CurrentX = prtMinX + 7000: XPrt.Print "Etat";

XPrt.CurrentX = prtMinX + 8000: XPrt.Print "Pays";

XPrt.CurrentX = prtMinX + 13300: XPrt.Print "Date";
XPrt.CurrentX = prtMinX + 14000: XPrt.Print "Heure";
XPrt.CurrentX = prtMinX + 14800: XPrt.Print "Utilisateur";



XPrt.FontSize = 7
XPrt.FontBold = blnFontBold_false
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight - 100

XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY), prtLineColor
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight + 50

blnDossier_Line = False
XPrt.DrawWidth = 1

End Sub









