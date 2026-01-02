Attribute VB_Name = "prtNovaBank"
Option Explicit

Dim I As Integer, mCurrenty As Integer
Dim colP As Integer, colA As Integer, colV As Integer, colX As Integer
Dim V, Height8_6 As Integer
Dim curTotal As Currency, curMtOpération As Currency
Dim xIn As String, mDevise As String, mAMJEchange As String, mAMJRéglement As String, NbOpération As Long
Dim mCodeOpération As String

'**************************************************************************
' Procédure de gestion des paramètres généraux à l'ouverture de l'état    *
'**************************************************************************

Public Sub prtNovaBank_Open()

Dim K As Integer, K1 As Integer, K2 As Integer
Dim X As String

Set XPrt = Printer
Height8_6 = frmElpPrt.prtHeightDelta(8, 6)


frmElpPrt.Show vbModeless

prtPgmName = "prtNovaBank"
prtTitleUsr = usrName

prtLineNb = 1
prtlineHeight = 250
prtHeaderHeight = 300

frmElpPrt.prtStdInit

End Sub
'***************************************************
'* Procédure de fermeture du fichier d'impression  *
'***************************************************
Public Sub prtNovaBank_Close()
    frmElpPrt.prtEndDoc
    frmElpPrt.Hide
End Sub

'**********************************************************************
'* Procédure d'impression de l'entête de page et entêtes de colonnes  *
'**********************************************************************

Public Sub prtNovaBank_SIT_Form()

Dim X As String, K As Integer

XPrt.FontSize = 10
XPrt.FontBold = False
XPrt.DrawWidth = 3

Call frmElpPrt.prtTrame(prtMinX, prtMinY, prtMaxX, prtMinY + prtHeaderHeight, "B")

K = prtMinY + prtHeaderHeight + 10

XPrt.FontBold = True
XPrt.CurrentY = prtMinY + 50
XPrt.CurrentX = prtMinX + 60: XPrt.Print "Code opération";
XPrt.CurrentX = prtMinX + 2000: XPrt.Print "Libellé de l'opération";
XPrt.CurrentX = prtMinX + 5150: XPrt.Print "Date réglement";
XPrt.CurrentX = prtMinX + 7150: XPrt.Print "Date échange";
XPrt.CurrentX = prtMinX + 9300: XPrt.Print "Nombre d'opérations";
XPrt.CurrentX = prtMinX + 13000: XPrt.Print "Montant SIT des opérations";

XPrt.FontBold = False

XPrt.CurrentY = prtMinY + prtHeaderHeight + prtlineHeight - XPrt.TextHeight("test")

End Sub

Public Sub prtNovaBank_SIT_Line()
'************************************************************************
' Procédure de gestion de l'édition des lignes Détail                   *
' (qui est en réalité un nombre total et un montant total d'opérations) *
'************************************************************************

Dim X As String
Dim NuméroLigne As Integer

'Gestion des sauts de pages
'--------------------------
If XPrt.CurrentY + prtlineHeight > prtMaxY Then
    prtNovaBank_SIT_Form_End
    frmElpPrt.prtNewPage
    prtNovaBank_SIT_Form
End If

XPrt.FontBold = False
XPrt.FontSize = 10
XPrt.CurrentX = prtMinX + 120: XPrt.Print mCodeOpération;
XPrt.CurrentX = prtMinX + 2000: XPrt.Print "VIREMENT ORDINAIRE";
XPrt.CurrentX = prtMinX + 5400: XPrt.Print Format$(mId$(mAMJRéglement, 5, 2) & mId$(mAMJRéglement, 3, 2) & mId$(mAMJRéglement, 1, 2), "@@.@@.@@");
XPrt.CurrentX = prtMinX + 7400: XPrt.Print Format$(mId$(mAMJEchange, 5, 2) & mId$(mAMJEchange, 3, 2) & mId$(mAMJEchange, 1, 2), "@@.@@.@@");
X = Format$(NbOpération, "### ##0")
XPrt.CurrentX = prtMinX + 10500 - XPrt.TextWidth(X)
XPrt.Print X;
X = Format$(Abs(curMtOpération), "### ### ### ### ##0.00")
XPrt.CurrentX = prtMinX + 15000 - XPrt.TextWidth(X)
XPrt.Print X;
XPrt.CurrentX = prtMinX + 15250: XPrt.Print Format$(mDevise, "@@@");

NuméroLigne = XPrt.CurrentY
XPrt.Line (prtMinX, NuméroLigne)-(prtMinX, NuméroLigne + prtlineHeight)                         'Ligne verticale gauche
XPrt.Line (prtMinX + 1800, NuméroLigne)-(prtMinX + 1800, NuméroLigne + prtlineHeight)           'Ligne verticale intérieure
XPrt.Line (prtMinX + 5000, NuméroLigne)-(prtMinX + 5000, NuméroLigne + prtlineHeight)           'Ligne verticale intérieure
XPrt.Line (prtMinX + 7000, NuméroLigne)-(prtMinX + 7000, NuméroLigne + prtlineHeight)           'Ligne verticale intérieure
XPrt.Line (prtMinX + 9000, NuméroLigne)-(prtMinX + 9000, NuméroLigne + prtlineHeight)           'Ligne verticale intérieure
XPrt.Line (prtMinX + 12500, NuméroLigne)-(prtMinX + 12500, NuméroLigne + prtlineHeight)         'Ligne verticale intérieure
XPrt.Line (prtMaxX, NuméroLigne)-(prtMaxX, NuméroLigne + prtlineHeight)                         'Ligne verticale droite

'Incrément du numéro ligne (position verticale)
'----------------------------------------------
'XPrt.CurrentY = XPrt.CurrentY + prtlineHeight

End Sub

'*****************************************************************
' Procédure de traitement du fichier à imprimer                  *
'*****************************************************************

Public Sub prtNovaBank_SIT(lFileName As String)

Dim X As String, Nb As Long
On Error GoTo Error_Handler

Open lFileName For Input As #1

prtOrientation = vbPRORLandscape
prtTitleText = "Récapitulatif des virements SIT (Comptabilité)"

'Appel des procédures de gestion des entêtes
'-------------------------------------------
prtNovaBank_Open
prtNovaBank_SIT_Form

xIn = ""
curTotal = 0

'Lecture du fichier SIT + apple procédure de gestion lignes détail
'-----------------------------------------------------------------
Do Until EOF(1)
    DoEvents
    Line Input #1, xIn
    
    Select Case mId$(xIn, 7, 2)
        Case "01": mDevise = mId$(xIn, 56, 3)
                   mAMJEchange = mId$(xIn, 35, 6)
        Case "02": mCodeOpération = mId$(xIn, 9, 3) & mId$(xIn, 81, 4)
                   mAMJRéglement = mId$(xIn, 53, 6)
        Case "09": NbOpération = CLng(mId$(xIn, 9, 8))
                   curMtOpération = CCur(Val(mId$(xIn, 17, 18)) / 100)
                    Call prtNovaBank_SIT_Line
                   curTotal = curTotal + curMtOpération
       End Select
    
Loop

Close
prtNovaBank_SIT_Form_End
prtNovaBank_Close
Exit Sub
Error_Handler:
Call MsgBox(Err & " : " & Error(Err), vbCritical, "prtNovaBank_SIT")

End Sub

'*******************************************************************
' Procédure des gestion de l'impression des lignes total           *
'*******************************************************************

Public Sub prtNovaBank_SIT_Form_End()
Dim X As String
XPrt.DrawWidth = 2
XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY)

XPrt.FontBold = True
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.CurrentX = prtMinX + 9000: XPrt.Print "Total général";
X = Format$(Abs(curTotal), "### ### ### ### ##0.00")
XPrt.CurrentX = prtMinX + 15000 - XPrt.TextWidth(X)
XPrt.Print X;

XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.CurrentX = prtMinX: XPrt.Print "BON A PAYER LE  " & dateImp(DSys) & "  à  " & Time;
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.CurrentX = prtMinX: XPrt.Print "12179   00001   BANQUE INTERCONTINENTALE ARABE";
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.CurrentX = prtMinX: XPrt.Print "FAX :    01 42 89 09 59";
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.CurrentX = prtMinX: XPrt.Print "TEL :    01 53 76 62 62";
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.CurrentX = prtMinX: XPrt.Print "POUR  :  Mme BELEM  (SG  A.D.B.)";
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2
XPrt.CurrentX = prtMinX: XPrt.Print "SIGNATURE : ";

End Sub


