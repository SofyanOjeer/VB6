Attribute VB_Name = "prtSAB_Ordonnanceur"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim X As String, I As Integer, Height8_6 As Integer

Dim blnPage As Boolean

Dim meYCLIENA0 As typeYCLIENA0, xYCLIENA0 As typeYCLIENA0
Dim meYADRESS0 As typeYADRESS0, xYADRESS0 As typeYADRESS0
Dim meYSWIBIC0 As typeYSWIBIC0
Dim meMVTP0 As typeMvtP0
Dim xYCLIREF0 As typeYCLIREF0

Public Sub prtSAB_Ordonnanceur(fgW As MSFlexGrid)



prtTitleText = "SAB : Ordonnanceur"
prtFontName = prtFontName_Arial
prtSAB_Ordonnanceur_Open
prtHeaderHeight = 300
prtSAB_Ordonnanceur_Form
XPrt.CurrentY = XPrt.CurrentY - prtlineHeight

XPrt.FontSize = 7
For I = 1 To fgW.Rows - 1
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    If XPrt.CurrentY + 500 > prtMaxY Then
        frmElpPrt.prtNewPage
        prtSAB_Ordonnanceur_Form
    End If
    
    fgW.Row = I
    fgW.Col = 6
    If fgW.Text = "YCOMTAC0" Then
        fgW.Col = 0: XPrt.CurrentX = prtMinX + 100: XPrt.Print fgW.Text;
        fgW.Col = 1: XPrt.CurrentX = prtMinX + 500: XPrt.Print fgW.Text;
        fgW.Col = 2: XPrt.CurrentX = prtMinX + 1500: XPrt.Print fgW.Text;
        fgW.Col = 3: XPrt.CurrentX = prtMinX + 2200: XPrt.Print fgW.Text;
    End If
    fgW.Col = 4: XPrt.CurrentX = prtMinX + 3000: XPrt.Print fgW.Text;

Next I

prtSAB_Ordonnanceur_Close

End Sub

'---------------------------------------------------------
Public Sub prtSAB_Ordonnanceur_Form()
'---------------------------------------------------------
Dim X As String

XPrt.DrawWidth = 1
XPrt.FontSize = 7: XPrt.FontBold = True

Call frmElpPrt.prtTrame(prtMinX, prtMinY, prtMaxX, prtMinY + prtHeaderHeight, "B", 240)
'---------------------------------------------------------
XPrt.CurrentY = prtMinY + 50

XPrt.CurrentX = prtMinX + 100: XPrt.Print "Per";
XPrt.CurrentX = prtMinX + 500: XPrt.Print "Traitement";
XPrt.CurrentX = prtMinX + 1500: XPrt.Print "Séquence";
XPrt.CurrentX = prtMinX + 2200: XPrt.Print "Option";
XPrt.CurrentX = prtMinX + 3000: XPrt.Print "Libellé";
XPrt.CurrentY = prtMinY + 50 + prtHeaderHeight
XPrt.FontBold = False

End Sub


Public Sub prtSAB_Ordonnanceur_Close()
On Error GoTo prtError


frmElpPrt.prtEndDoc
frmElpPrt.Hide
Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub



Public Sub prtSAB_Ordonnanceur_Open()
On Error GoTo prtError

Set XPrt = Printer
frmElpPrt.Show vbModeless

Height8_6 = frmElpPrt.prtHeightDelta(8, 6)

blnFiligrane = False
prtOrientation = vbPRORPortrait '
prtPgmName = "prtSAB_Ordonnanceur"
prtTitleUsr = usrName

prtLineNb = 1
prtlineHeight = 250
prtHeaderHeight = 50 ' 100


prtFormType = ""
frmElpPrt.prtStdInit

Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub



