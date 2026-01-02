Attribute VB_Name = "prtAdresseSub"
Option Explicit
Dim recAdresse As typeAdresse
Dim recRacine As typeRacine
Dim colA As Integer, colV As Integer
Dim mY As Integer, mRacineY As Integer
'---------------------------------------------------------
Public Sub prtAdresseX(Msg As String)
'---------------------------------------------------------
Dim K As Integer, K1 As Integer, K2 As Integer
Dim X As String
K1 = Val(Mid$(Msg, 1, 6))
K2 = Val(Mid$(Msg, 7, 6))


For K = K1 To K2
    recRacine = arrRacine(K)
    prtAdresse_Racine
    mRacineY = XPrt.CurrentY
    prtAdresse_Compte_dbSnap
    If arrAdresseNb > 0 Then
        XPrt.CurrentY = mY - prtlineHeight
        arrAdresseIndex = 1
        prtAdresse_Compte arrAdresse(arrAdresseIndex)
        recAdresse = arrAdresse(1)
        For arrAdresseIndex = 2 To arrAdresseNb
            If recAdresse.Adresse1 <> arrAdresse(arrAdresseIndex).Adresse1 _
            Or recAdresse.Adresse2 <> arrAdresse(arrAdresseIndex).Adresse2 _
            Or recAdresse.Adresse3 <> arrAdresse(arrAdresseIndex).Adresse3 _
            Or recAdresse.Adresse4 <> arrAdresse(arrAdresseIndex).Adresse4 _
            Or recAdresse.Adresse5 <> arrAdresse(arrAdresseIndex).Adresse5 _
            Or recAdresse.AdresseCP <> arrAdresse(arrAdresseIndex).AdresseCP _
            Or recAdresse.AdresseBD <> arrAdresse(arrAdresseIndex).AdresseBD _
            Or recAdresse.AdressePays <> arrAdresse(arrAdresseIndex).AdressePays _
                    Then
                    prtAdresse_Compte arrAdresse(arrAdresseIndex)
            Else
                    prtAdresse_Numéro (arrAdresseIndex)
            End If
        Next arrAdresseIndex
    End If
    If XPrt.CurrentY < mRacineY Then XPrt.CurrentY = mRacineY
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.Line (prtMinX, XPrt.CurrentY - 50)-(prtMaxX, XPrt.CurrentY - 50)
    DoEvents: If prtKillDoc Then frmElpPrt.Hide: Exit Sub
Next K

'frmElpPrt.prtLineY

End Sub


'---------------------------------------------------------
Public Sub prtAdresseForm()
'---------------------------------------------------------
Dim X As String
XPrt.FontSize = 7
XPrt.FontBold = False

XPrt.DrawWidth = 1
Call frmElpPrt.prtTrame(prtMinX, prtMinY, prtMaxX, prtMinY + prtHeaderHeight, "B", 250)

XPrt.Line (colA, prtMinY)-(colA, prtMaxY)
XPrt.Line (colV, prtMinY)-(colV, prtMaxY)
'---------------------------------------------------------
XPrt.CurrentY = prtMinY + 50
X = "Adresse de la racine"
XPrt.CurrentX = (colA + colV - XPrt.TextWidth(X)) / 2: XPrt.Print X;
X = "Adresse des comptes"
XPrt.CurrentX = (colV + prtMaxX - XPrt.TextWidth(X)) / 2: XPrt.Print X;

XPrt.CurrentY = prtMinY + prtHeaderHeight + 50
mRacineY = XPrt.CurrentY
End Sub

'---------------------------------------------------------
Public Sub prtAdresse_Racine()
'---------------------------------------------------------
Dim X As String, K As Integer, mCurrenty As Integer

If XPrt.CurrentY + prtParagraphHeight > prtMaxY Then
    frmElpPrt.prtNewPage
    prtAdresseForm
'else
    'frmElpPrt.prtLineY
End If
XPrt.FontBold = False

'------------------------------------------ligne 1--------------
mY = XPrt.CurrentY
XPrt.FontSize = 7
XPrt.CurrentX = prtMinX
XPrt.Print "Racine ";
XPrt.CurrentX = 800
XPrt.Print ":";
XPrt.CurrentX = 900
XPrt.FontBold = True
XPrt.Print Format$(recRacine.Numéro, "00000");

XPrt.CurrentX = colA + 100
XPrt.Print recRacine.Intitulé;
XPrt.FontBold = False

'----------------------------------ligne2
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinX
XPrt.Print "Tél. ";
XPrt.CurrentX = 800
XPrt.Print ":";
XPrt.CurrentX = 900
XPrt.Print recRacine.Téléphone1 & " " & recRacine.Téléphone2;
XPrt.CurrentX = colA + 100
XPrt.Print recRacine.Adresse1;

 '----------------------------------ligne3
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinX
XPrt.Print "Fax. ";
XPrt.CurrentX = 800
XPrt.Print ":";
XPrt.CurrentX = 900
XPrt.Print recRacine.Fax;
XPrt.CurrentX = colA + 100
XPrt.Print recRacine.Adresse2;

'----------------------------------ligne2
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMinX
XPrt.Print "Swift/Tx ";
XPrt.CurrentX = 800
XPrt.Print ":";
XPrt.CurrentX = 900
XPrt.Print recRacine.Swift & "   " & recRacine.Fax;
XPrt.CurrentX = colA + 100
XPrt.Print recRacine.AdresseCP & "    " & recRacine.Adresse3;

'---------------------------------------------
End Sub






Public Sub prtAdresse_Open()
On Error GoTo prtError

Set XPrt = Printer


frmElpPrt.Show vbModeless

prtOrientation = vbPRORPortrait
prtTitleText = "Liste des adresses"
prtPgmName = "prtAdresse"
prtTitleUsr = usrName

prtLineNb = 5
prtlineHeight = 250
prtHeaderHeight = 300
colA = 3000
colV = 7000

frmElpPrt.prtStdInit

prtAdresseForm
ReDim arrAdresse(1): arrAdresseNbMax = 1
Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub

Public Sub prtAdresse_Close()
frmElpPrt.prtEndDoc
frmElpPrt.Hide

End Sub

Public Sub prtAdresse_Compte_dbSnap()
srvAdresse.Init recAdresse
recAdresse.Method = "SnapL0"
recAdresse.Numéro = recRacine.Numéro & "000000"
arrAdresse(0) = recAdresse
arrAdresse(0).Numéro = recRacine.Numéro & "999999"
arrAdressesuite = True
arrAdresseNb = 0: arrAdresseIndex = 0
Do Until Not arrAdressesuite
    srvAdresse.Monitor recAdresse
    recAdresse = arrAdresse(arrAdresseNb)
    recAdresse.Method = "SnapL0+"
Loop

End Sub

Public Sub prtAdresse_Compte(xAdresse As typeAdresse)
Dim X As String, K As Integer, mCurrenty As Integer

If XPrt.CurrentY + prtParagraphHeight > prtMaxY Then
    frmElpPrt.prtNewPage
    prtAdresseForm
'else
    'frmElpPrt.prtLineY
End If
XPrt.FontBold = False

'------------------------------------------ligne 1--------------
'XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
prtAdresse_Numéro (arrAdresseIndex)
mCurrenty = XPrt.CurrentY
XPrt.CurrentX = colV + 100
XPrt.Print xAdresse.Adresse1;
Call frmElpPrt.prtTrame(colV + 10, mCurrenty, prtMaxX, mCurrenty + prtlineHeight, " ")

If Trim(xAdresse.Adresse2) <> "" Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = colV + 100
    XPrt.Print xAdresse.Adresse2;
End If
If Trim(xAdresse.Adresse3) <> "" Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = colV + 100
    XPrt.Print xAdresse.Adresse3;
End If
If Trim(xAdresse.Adresse4) <> "" Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = colV + 100
    XPrt.Print xAdresse.Adresse4;
End If

If Trim(xAdresse.Adresse5) <> "" Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = colV + 100
    XPrt.Print xAdresse.Adresse5;
End If

X = xAdresse.AdresseCP & " " & xAdresse.AdresseBD
If Trim(X) <> "" Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = colV + 100
    XPrt.Print X;
End If
If Trim(xAdresse.AdressePays) <> "" Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    XPrt.CurrentX = colV + 100
    XPrt.Print xAdresse.AdressePays;
End If


End Sub

Public Sub prtAdresse_Numéro(I As Integer)
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = prtMaxX - 1100
XPrt.Print Compte_Imp(arrAdresse(I).Numéro);
End Sub
