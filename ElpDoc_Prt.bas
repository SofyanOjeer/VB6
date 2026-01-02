Attribute VB_Name = "prtElpDoc"
Option Explicit
Public arrPlan() As typeElpTable, arrPlan_Index As Integer, arrPlan_Nb As Integer
Public lnkPlan() As typeElpTable
Public arrTable() As typeElpTable, lnkTable() As typeElpTable
Public arrDoc() As typeElpDoc, recDoc As typeElpDoc, arrDoc_Index As Integer
Public arrSelectFilter() As String
Public arrSelectSN() As String * 12
Dim arrSelectSN_Nb As Integer
Dim xElpDoc As typeElpDoc, xElpTable As typeElpTable

Dim arrDoc_Nb As Integer
Dim colDocument As Integer, linDocument As Integer
Dim colDiffusion As Integer, linDiffusion As Integer

Dim memoId As String, lenX As Integer

Dim Line1 As Integer, Line2 As Integer
Dim Col1 As Integer, Col2 As Integer
Dim nbLigne As Integer
Dim X As String
'---------------------------------------------------------
 Public Sub prtElpDocX(Msg As String)
'---------------------------------------------------------
Dim I As Integer, xId As String

On Error GoTo prtError


Set XPrt = Printer

arrDoc_Nb = UBound(arrDoc) - 1
arrSelectSN_Nb = UBound(arrSelectSN) - 1
memoId = ""

frmElpPrt.Show vbModeless

prtOrientation = vbPRORPortrait
prtTitleText = ""
prtPgmName = "prtElpDoc"
prtTitleUsr = usrName

prtLineNb = 1
prtlineHeight = 300
prtHeaderHeight = 0

recElpDoc_Init xElpDoc
recElpTable_Init xElpTable

Select Case Msg
    Case "Dossier"
                    prtlineHeight = 300
                    prtHeaderHeight = 0
                    prtOrientation = vbPRORPortrait
                    prtSocInit
                    prtMinX = 800
                    prtElpDoc_Form
                    prtElpDoc_Dossier
    Case "ElpDoc"
                    prtlineHeight = 300
                    prtHeaderHeight = 300
                    prtOrientation = vbPRORLandscape
                    prtTitleText = "Fiche technique du dossier : " & arrDoc(I).K1
                    frmElpPrt.prtStdInit
                    prtElpDoc_Technique_Form
                    prtElpDoc_Technique
    Case "SelectSN"
                    prtlineHeight = 300
                    prtHeaderHeight = 300
                    prtOrientation = vbPRORLandscape
                    prtTitleText = "Filtre / Sélection"
                    frmElpPrt.prtStdInit
                    prtElpDoc_SelectSN_Form
                    prtElpDoc_SelectFilter
                    prtElpDoc_SelectSN
End Select


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
Public Sub prtElpDoc_Line(recDoc As typeElpDoc)
'---------------------------------------------------------
XPrt.CurrentX = prtMinX + 2000
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.FontBold = False
XPrt.FontSize = 8

XPrt.Print recDoc.Memo;

End Sub

'---------------------------------------------------------
Public Sub prtElpDoc_Form()
'---------------------------------------------------------
Dim strAv As String
linDocument = 2200
Line1 = linDocument - 200
Line2 = Line1
Col1 = prtMinX - 50
Col2 = prtMaxX

XPrt.FontBold = False
XPrt.FontSize = 8

XPrt.FontBold = True
XPrt.CurrentX = prtMinX
XPrt.CurrentY = linDocument
XPrt.Print "Références : ";
XPrt.FontBold = False
colDocument = XPrt.CurrentX
XPrt.Print Trim(lnkTable(1).Name);

XPrt.FontBold = True
XPrt.CurrentX = 7000
XPrt.Print "Diffusion : ";
XPrt.FontBold = False
colDiffusion = XPrt.CurrentX
linDiffusion = XPrt.CurrentY

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.CurrentX = colDocument
XPrt.Print arrDoc(0).Memo;
linDocument = XPrt.CurrentY

XPrt.FontBold = False
XPrt.FontSize = 8

End Sub



Public Sub prtElpDoc_Intitulé(X As String)
Dim I As Integer
XPrt.FontBold = True

For I = 10 To 4 Step -1
    XPrt.FontSize = I
    If XPrt.TextWidth(X) <= (prtMaxX - prtMinX - 100) Then Exit For
Next I

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
frmElpPrt.prtTrame prtMinX - 50, XPrt.CurrentY, prtMaxX, XPrt.CurrentY + prtlineHeight * 2, " ", 245
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight / 2
frmElpPrt.prtCentré (prtMaxX + prtMinX) / 2, X
XPrt.FontBold = False

End Sub
Public Sub prtElpDoc_Cadre()

If linDiffusion > linDocument Then
    Line2 = linDiffusion
Else
    Line2 = linDocument
End If


Line2 = Line2 + prtlineHeight + 100

XPrt.DrawWidth = 2
XPrt.Line (Col1 + 200, Line1)-(Col2 - 200, Line1)
XPrt.Line (Col1 + 200, Line2)-(Col2 - 200, Line2)
XPrt.Line (Col1, Line1 + 200)-(Col1, Line2 - 200)
XPrt.Line (Col2, Line1 + 200)-(Col2, Line2 - 200)

XPrt.CurrentY = 0
XPrt.CurrentX = 0
XPrt.Circle Step(Col1 + 200, Line1 + 200), 200, 0, 0.5 * Pi, Pi
XPrt.DrawWidth = 3

XPrt.CurrentY = 0
XPrt.CurrentX = 0
XPrt.Circle Step(Col2 - 200, Line1 + 200), 200, 0, 0, 0.5 * Pi


XPrt.DrawWidth = 3
XPrt.CurrentY = 0
XPrt.CurrentX = 0
XPrt.Circle Step(Col1 + 200, Line2 - 200), 200, 0, Pi, 1.5 * Pi



XPrt.CurrentY = 0
XPrt.CurrentX = 0
XPrt.Circle Step(Col2 - 200, Line2 - 200), 200, 0, 1.5 * Pi, 2 * Pi

End Sub



Public Sub prtElpDoc_Dossier()
Dim I As Integer, xId As String

arrPlan_Nb = UBound(arrPlan) - 1
For arrPlan_Index = 1 To arrPlan_Nb
   If Trim(arrPlan(arrPlan_Index).K2) = "Intitulé" Then Exit For
Next arrPlan_Index

For I = 2 To arrDoc_Nb
    xId = Trim(arrDoc(I).Id)
    Select Case xId
        Case "Version"
            XPrt.FontSize = 8
            XPrt.CurrentX = colDocument
            XPrt.CurrentY = linDocument + prtlineHeight
            XPrt.Print arrDoc(I).Memo;
            linDocument = XPrt.CurrentY
        Case "Rédacteur"
            XPrt.FontSize = 8
            XPrt.CurrentX = colDocument
            XPrt.CurrentY = linDocument + prtlineHeight
            XPrt.Print lnkTable(I).Name;
            linDocument = XPrt.CurrentY
        Case "Diffusion"
            XPrt.FontSize = 8
            XPrt.CurrentX = colDiffusion
            XPrt.CurrentY = linDiffusion: linDiffusion = XPrt.CurrentY + prtlineHeight
            XPrt.Print lnkTable(I).Name;
        Case "Intitulé"
            If memoId <> xId Then prtElpDoc_Cadre: XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
            prtElpDoc_Intitulé Trim(CStr(arrDoc(I).Memo))
        Case "Confidential", "MotClé"
        Case Else
            If memoId <> xId Then
                prtElpDoc_Dossier_Paragraphe_Void xId
                
                prtElpDoc_Dossier_Paragraphe Trim(lnkPlan(I).Name)
            End If
            XPrt.FontSize = 10
            XPrt.CurrentX = prtMinX + 600
            XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
            XPrt.Print arrDoc(I).Memo;
    End Select
    memoId = xId
    
Next I
prtElpDoc_Dossier_Paragraphe_Void Trim(arrPlan(arrPlan_Nb).K2)

End Sub
'---------------------------------------------------------
Public Sub prtElpDoc_Technique_Form()
'---------------------------------------------------------
Dim X As String

XPrt.FontSize = 8
XPrt.FontBold = True
XPrt.DrawWidth = 3

XPrt.Line (prtMinX, prtMinY)-(prtMaxX, prtMaxY), , B
'XPrt.Line (prtMinX, prtMinY + prtHeaderHeight)-(prtMaxX, prtMinY + prtHeaderHeight)
Call frmElpPrt.prtTrame(prtMinX, prtMinY, prtMaxX, prtMinY + prtHeaderHeight, "B", 240)
XPrt.DrawWidth = 1


'----------------------------------------ligne 1-----------------

XPrt.CurrentY = prtMinY + (prtHeaderHeight - XPrt.TextHeight(X)) / 2

XPrt.CurrentX = prtMinX + 100
XPrt.Print "Plan";

XPrt.CurrentX = prtMinX + 1500
XPrt.Print "Code";
XPrt.CurrentX = prtMinX + 3100
XPrt.Print "Informations";

XPrt.CurrentX = prtMinX + 9000
XPrt.Print "Intitulé";
XPrt.CurrentX = prtMinX + 12000
XPrt.Print "Complément";

XPrt.FontBold = False

prtCurrentY = prtMinY + prtHeaderHeight + 100

End Sub


'---------------------------------------------------------
Public Sub prtElpDoc_SelectSN_Form()
'---------------------------------------------------------
Dim X As String

XPrt.FontSize = 8
XPrt.FontBold = True
XPrt.DrawWidth = 3

XPrt.Line (prtMinX, prtMinY)-(prtMaxX, prtMaxY), , B
'XPrt.Line (prtMinX, prtMinY + prtHeaderHeight)-(prtMaxX, prtMinY + prtHeaderHeight)
Call frmElpPrt.prtTrame(prtMinX, prtMinY, prtMaxX, prtMinY + prtHeaderHeight, "B", 240)
XPrt.DrawWidth = 1


'----------------------------------------ligne 1-----------------

XPrt.CurrentY = prtMinY + (prtHeaderHeight - XPrt.TextHeight(X)) / 2

XPrt.CurrentX = prtMinX + 100
XPrt.Print "Service";

XPrt.CurrentX = prtMinX + 3000
XPrt.Print "Document";

XPrt.CurrentX = prtMinX + 8000
XPrt.Print "Intitulé";

XPrt.FontBold = False

prtCurrentY = prtMinY + prtHeaderHeight + 100
nbLigne = 0
End Sub



Public Sub prtElpDoc_Technique()
Dim I As Integer
For I = 0 To arrDoc_Nb
    If prtCurrentY + prtParagraphHeight > prtMaxY Then
        frmElpPrt.prtNewPage
        prtElpDoc_Technique_Form
    End If
    
    '------------------------------------------ligne 1--------------
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    
    'NbLigne = NbLigne + 1
    'If NbLigne > 2 Then
    '    Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY - 50, prtMaxX, XPrt.CurrentY + prtlineHeight - 50, " ")
    '    If NbLigne = 4 Then NbLigne = 0
    'End If
    
    '----------------------------------
    If Trim(lnkTable(I).Id) <> "" Then
       Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY - 50, prtMaxX, XPrt.CurrentY + prtlineHeight - 50, " ")
       XPrt.CurrentX = prtMinX + 9000
       XPrt.Print lnkTable(I).Name;
       If Not IsNull(lnkTable(I).Memo) Then XPrt.CurrentX = prtMinX + 12000: XPrt.Print lnkTable(I).Memo;
    End If
    
    XPrt.CurrentX = prtMinX + 100
    XPrt.Print arrDoc(I).Id;
    XPrt.CurrentX = prtMinX + 1500
    XPrt.Print arrDoc(I).K2;
    If Not IsNull(arrDoc(I).Memo) Then XPrt.CurrentX = prtMinX + 3000: XPrt.Print arrDoc(I).Memo;
Next I

End Sub
Public Sub prtElpDoc_SelectFilter()
Dim I As Integer, K As Integer
For I = 0 To UBound(arrSelectFilter) - 1
    If prtCurrentY + prtParagraphHeight > prtMaxY Then
        frmElpPrt.prtNewPage
        prtElpDoc_SelectSN_Form
    End If
    
    '------------------------------------------ligne 1--------------
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    
    X = arrSelectFilter(I)
    K = InStr(X, Chr$(9))
    XPrt.CurrentX = prtMinX + 2000
    XPrt.Print mId$(X, 1, K - 1);
    XPrt.CurrentX = prtMinX + 3000
    XPrt.Print mId$(X, K + 2, 40);
Next I
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight

XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY)
XPrt.CurrentY = XPrt.CurrentY + 20
XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY)
XPrt.CurrentY = XPrt.CurrentY - prtlineHeight + 50

End Sub

Public Sub prtElpDoc_SelectSN()
Dim I As Integer
For I = 0 To arrSelectSN_Nb
    If prtCurrentY + prtParagraphHeight > prtMaxY Then
        frmElpPrt.prtNewPage
        prtElpDoc_SelectSN_Form
    End If
    
    '------------------------------------------ligne 1--------------
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    
    nbLigne = nbLigne + 1
    If nbLigne > 2 Then
        Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY - 50, prtMaxX, XPrt.CurrentY + prtlineHeight - 50, " ")
        If nbLigne = 4 Then nbLigne = 0
    End If
    
    xElpDoc.Method = "Seek="
      
    xElpDoc.Id = "Document"
    xElpDoc.K2 = ""
    xElpDoc.K1 = arrSelectSN(I)
    If tableElpDoc_Read(xElpDoc) <> 0 Then xElpDoc.Memo = "???? err " & arrSelectSN(I)

    XPrt.CurrentX = prtMinX + 3000
    If Not IsNull(xElpDoc.Memo) Then XPrt.Print xElpDoc.Memo;
    
    xElpDoc.Id = "Intitulé"
    xElpDoc.K2 = "000000000001"
    xElpDoc.K1 = arrSelectSN(I)
    Call tableElpDoc_Read(xElpDoc)
    XPrt.CurrentX = prtMinX + 8000
    If Not IsNull(xElpDoc.Memo) Then XPrt.Print xElpDoc.Memo;
    
    xElpDoc.Method = "Seek>="

    xElpDoc.Id = "Doc_Service"
    xElpDoc.K2 = ""
    xElpDoc.K1 = arrSelectSN(I)
    If tableElpDoc_Read(xElpDoc) = 0 Then
        If xElpDoc.K1 = arrSelectSN(I) Then
            xElpTable.Method = "Seek="
            xElpTable.Id = "Doc_Service"
            xElpTable.K2 = ""
            xElpTable.K1 = xElpDoc.K2
            If tableElpTable_Read(xElpTable) = 0 Then
                XPrt.CurrentX = prtMinX + 100
                XPrt.Print xElpTable.Name;
            End If
        End If
    End If
Next I


End Sub


Public Sub prtElpDoc_Dossier_Paragraphe(strX As String)

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.FontBold = True
XPrt.FontSize = 8
lenX = XPrt.TextWidth(strX) + 150
frmElpPrt.prtTrame prtMinX, XPrt.CurrentY, prtMinX + lenX, XPrt.CurrentY + prtlineHeight, " ", 245
XPrt.CurrentY = XPrt.CurrentY + 50
XPrt.CurrentX = prtMinX + 50
XPrt.Print strX;
XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight

End Sub

Public Sub prtElpDoc_Dossier_Paragraphe_Void(strX As String)
Dim IPlan As Integer
IPlan = arrPlan_Index + 1
For arrPlan_Index = IPlan To arrPlan_Nb
    If strX = Trim((arrPlan(arrPlan_Index).K2)) Then Exit For
    prtElpDoc_Dossier_Paragraphe Trim(arrPlan(arrPlan_Index).Name)
Next arrPlan_Index

End Sub
