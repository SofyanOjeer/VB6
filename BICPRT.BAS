Attribute VB_Name = "prtBic"
Option Explicit

'---------------------------------------------------------
Private recBic As typeBic
Private recBicIban As typeBicIban
Dim I As Integer
Dim blnTest As Boolean


'---------------------------------------------------------
Public Sub prtBicX(Msg As String)
'---------------------------------------------------------
Dim K As Integer, K1 As Integer, K2 As Integer
On Error GoTo prtError


Set XPrt = Printer
If prtShow Then frmElpPrt.Show vbModeless

prtOrientation = vbPRORLandscape
prtTitleText = "liste des Correspondants "
prtPgmName = "prtBic"
prtTitleUsr = usrName

prtLineNb = 3
prtlineHeight = 250
prtHeaderHeight = 300

frmElpPrt.prtStdInit
prtBicForm
If mId$(Msg, 13, 12) = "frmBic      " Then
    prtBicAll Msg
Else

    K1 = Val(mId$(Msg, 1, 6))
    K2 = Val(mId$(Msg, 7, 6))
    For K = K1 To K2
        recBic = arrBic(K)
        prtBicLine
        prtBicIbanX
        DoEvents: If prtKillDoc Then frmElpPrt.Hide: Exit Sub
    Next K
End If
frmElpPrt.prtLineY
prtTraitV
frmElpPrt.prtEndDoc
If prtShow Then frmElpPrt.Hide
Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub

'---------------------------------------------------------
Private Sub prtBicForm()
'---------------------------------------------------------
Dim X As String

XPrt.FontSize = 8
XPrt.FontBold = True
XPrt.DrawWidth = 3

'XPrt.FillStyle = 0
'XPrt.ForeColor = RGB(0, 0, 0)
'XPrt.FillStyle = 1

XPrt.Line (prtMinX, prtMaxY)-(prtMaxX, prtMaxY)
Call frmElpPrt.prtTrame(prtMinX, prtMinY, prtMaxX, prtMinY + prtHeaderHeight, "B")
'---------------------------------------------------------
X = "Code"
XPrt.CurrentY = prtMinY + (prtHeaderHeight - XPrt.TextHeight(X)) / 2
XPrt.CurrentX = prtMinX + 300
XPrt.Print X;
XPrt.CurrentX = 1900: XPrt.Print "Swift";
XPrt.CurrentX = 3400: XPrt.Print "Télex";
XPrt.CurrentX = 7200: XPrt.Print "Intitulé";
XPrt.CurrentX = 14400: XPrt.Print "Validité";
XPrt.CurrentX = 10100: XPrt.Print "Banque";
XPrt.CurrentX = 10900: XPrt.Print "Guichet";
XPrt.CurrentX = 12400: XPrt.Print "Avis";
prtCurrentY = prtMinY + prtHeaderHeight
XPrt.DrawWidth = 1
End Sub
'---------------------------------------------------------
Private Sub prtBicLine()
'---------------------------------------------------------
Dim X As String, I As Integer

If prtCurrentY + prtParagraphHeight > prtMaxY Then
    prtTraitV
    frmElpPrt.prtNewPage
    prtBicForm
Else
    frmElpPrt.prtLineY
End If
'I = XPrt.CurrentY

XPrt.FontBold = False

XPrt.CurrentY = prtCurrentY + prtlineHeight - XPrt.TextHeight("test")
Call frmElpPrt.prtTrame(prtMinX, prtCurrentY, prtMaxX, prtCurrentY + prtlineHeight, " ")

'------------------------------------ligne 1
XPrt.FontSize = 8
XPrt.CurrentX = prtMinX + 100: XPrt.Print recBic.BicId;
XPrt.FontBold = True
XPrt.CurrentX = 1700: XPrt.Print recBic.SwiftId;
XPrt.FontBold = False
XPrt.CurrentX = 3200: XPrt.Print recBic.TélexId;
XPrt.FontBold = True
XPrt.CurrentX = 5100: XPrt.Print recBic.Intitulé;
XPrt.FontBold = False
XPrt.CurrentX = 12200
Select Case recBic.optAvis
  Case "0": XPrt.Print "Pas d'Avis";
  Case "1": XPrt.Print "Au fil de l'eau";
  Case "2": XPrt.Print "Envoi Périodique";
End Select

XPrt.FontBold = False
XPrt.CurrentX = 10100
If recBic.BdfEtablissement <> "00000" Then XPrt.Print (recBic.BdfEtablissement);
XPrt.CurrentX = 11000
If recBic.BdfGuichet <> "00000" Then XPrt.Print (recBic.BdfGuichet);
XPrt.CurrentX = 13700: XPrt.Print dateImp(recBic.AmjMin);
XPrt.CurrentX = 14900: XPrt.Print dateImp(recBic.AmjMax);
'-----------------------------------ligne 2----------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 5100: XPrt.Print recBic.VilleId;
XPrt.CurrentX = 5600: XPrt.Print recBic.Ville;
XPrt.CurrentX = 12200
If recBic.optAvis <> "0" Then
  Select Case recBic.optAvisLangue
  Case "0": XPrt.Print "Français/Anglais";
  Case "1": XPrt.Print "Français";
  Case "2": XPrt.Print "Anglais";
End Select
End If
'------------------------------------ligne 3---------
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = 5100: XPrt.Print recBic.PaysId;
XPrt.CurrentX = 5600: XPrt.Print recBic.Pays;

prtCurrentY = prtCurrentY + prtParagraphHeight

End Sub



Public Sub prtBicIbanX()
blnTest = True
recBicIbanInit recBicIban
recBicIban.Method = "SnapP0"
recBicIban.BicId = recBic.BicId
arrBicIban(0) = recBicIban
arrBicIban(0).IbanDevise = "999"
Do
    arrBicIbanNb = 0
    arrBicIbanIndex = 0
    If IsNull(srvBicIbanMon(recBicIban)) Then
        For arrBicIbanIndex = 1 To arrBicIbanNb
            recBicIban = arrBicIban(arrBicIbanIndex)
            prtBicIbanLine
            DoEvents: If prtKillDoc Then frmElpPrt.Hide: Exit Sub
        Next arrBicIbanIndex
        recBicIban.Method = "SnapP0+"
    End If
        
Loop While arrBicIbanNext

End Sub

Public Sub prtBicIbanLine()
Dim X As String, I As Integer

If prtCurrentY + prtlineHeight > prtMaxY Then
    prtTraitV
    frmElpPrt.prtNewPage
    prtBicForm
End If

XPrt.FontBold = False

I = prtCurrentY + prtlineHeight - XPrt.TextHeight("test")
'If blnTest Then blnTest = False: XPrt.Line (5000, I - 100)-(prtMaxX, I - 100)
'Call frmElpPrt.prtTrame(5000, I, prtMaxX, I + prtlineHeight, " ")
'------------------------------------ligne 1
XPrt.CurrentY = I
XPrt.FontSize = 8
XPrt.FontBold = False
XPrt.CurrentX = 7500: XPrt.Print recBicIban.IbanBicId;

XPrt.CurrentX = 9000: XPrt.Print recBicIban.IbanDevise;
XPrt.CurrentX = 9300: XPrt.Print recBicIban.IbanPays;
XPrt.CurrentX = 9600: XPrt.Print recBicIban.IbanClé;
If recBicIban.IbanPays = "FR" Then
    XPrt.CurrentX = 10100: XPrt.Print mId$(recBicIban.IbanCompte, 1, 5);
    XPrt.CurrentX = 11000: XPrt.Print mId$(recBicIban.IbanCompte, 6, 5);
    XPrt.CurrentX = 12000: XPrt.Print Format$(mId$(recBicIban.IbanCompte, 11, 11), "@@ @@@ @@@ @@@");
     XPrt.Print "   " & mId$(recBicIban.IbanCompte, 22, 2);
Else

    XPrt.CurrentX = 11500: XPrt.Print recBicIban.IbanCompte;
End If
XPrt.CurrentX = 13700: XPrt.Print dateImp(recBicIban.AmjMin);
XPrt.CurrentX = 14900: XPrt.Print dateImp(recBicIban.AmjMax);
prtCurrentY = prtCurrentY + prtlineHeight


End Sub

Public Sub prtBicAll(Msg As String)

recBicInit recBic
If mId$(Msg, 25, 2) = "P0" Then
    prtTitleText = "liste des Correspondants (Identifiant)"
    recBic.Method = "SnapP0"
    recBic.BicId = Trim(mId$(Msg, 1, 12))
    arrBic(0) = recBic
    arrBic(0).BicId = Trim(mId$(Msg, 1, 12)) & "99999999999"
Else
    prtTitleText = "liste des Correspondants (Intitulé)"
    recBic.Method = "SnapL5"
    recBic.Intitulé = Trim(mId$(Msg, 1, 12))
    arrBic(0) = recBic
    arrBic(0).Intitulé = Trim(mId$(Msg, 1, 12)) & "99999999999"
End If
Do
    arrBicNb = 0
    arrBicIndex = 0
    If IsNull(srvBicMon(recBic)) Then
        For arrBicIndex = 1 To arrBicNb
            recBic = arrBic(arrBicIndex)
            prtBicLine
            prtBicIbanX
            DoEvents: If prtKillDoc Then frmElpPrt.Hide: Exit Sub
        Next arrBicIndex
        recBic.Method = mId$(recBic.Method, 1, 5) + "+"
    End If
        
Loop While arrBicNext

End Sub

Public Sub prtTraitV()
XPrt.DrawWidth = 1
XPrt.Line (5000, prtMinY)-(5000, prtMaxY)
XPrt.Line (10000, prtMinY)-(10000, prtMaxY)
XPrt.Line (13600, prtMinY)-(13600, prtMaxY)

End Sub
