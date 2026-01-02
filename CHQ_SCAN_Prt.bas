Attribute VB_Name = "prtCHQ_SCAN"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim xCHQ_SCAN As typeCHQ_SCAN
Dim mRefInterne As String
Dim nbR As Long, curR As Currency
Dim Height8_6 As Integer

Dim nbT As Long, curT As Currency
Dim nbChqR As Long, nbChqT As Long

Dim Page_No As Integer
Dim blnErreur_Total As Boolean
Dim blnErreur_Ajustement As Boolean

Type typeCHQ_Stat
 
      Date              As String
      Nature            As String
      
      Remise_SG         As Long
      Chèque_SG        As Long
      Montant_SG        As Currency
      
      Remise_BIA         As Long
      Chèque_BIA        As Long
      Montant_BIA        As Currency
      
      Remise_Divers         As Long
      Chèque_Divers       As Long
      Montant_Divers       As Currency
      
      Remise_Devise        As Long
      Chèque_Devise       As Long
     
      Remise_Nb1         As Long
      Remise_Nb2         As Long
      Remise_Nb3         As Long
     
End Type


Public Sub prtCHQ_SCAN_List1_Close()
Dim X As String
On Error GoTo prtError
prtCHQ_SCAN_List1_Rupture
XPrt.CurrentY = XPrt.CurrentY + 100
XPrt.FontBold = True
prtCHQ_SCAN_List1_NewLine
Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY, prtMaxX, XPrt.CurrentY + prtlineHeight + 100, "B", 245)
XPrt.CurrentY = XPrt.CurrentY + 75
X = Format$(curT, "### ### ### ###.00")
XPrt.CurrentX = prtMinX + 10000 - XPrt.TextWidth(X)
XPrt.Print X;
XPrt.CurrentX = prtMinX + 5000
    If nbT <= 1 Then
        XPrt.Print nbT & "     bordereau de remise en banque";
    Else
        XPrt.Print nbT & "     bordereaux de remise en banque";
    End If
XPrt.CurrentX = prtMinX + 1500
XPrt.Print "TOTAL";
X = Format$(nbChqT, "### ### ##0")
XPrt.CurrentX = prtMinX + 12000 - XPrt.TextWidth(X)
XPrt.Print X;
 nbChqT = 0
nbT = 0: curT = 0

If blnErreur_Total Then
    XPrt.FontSize = 14
    prtCHQ_SCAN_List1_NewLine
    Call frmElpPrt.prtTrame(prtMinX + 3420, XPrt.CurrentY, prtMinX + 10780, XPrt.CurrentY + prtlineHeight + 100, "b", 240)
    frmElpPrt.prtCentré prtMedX, "??? ERREUR : Remise / total des chèques ???"
End If

If blnErreur_Ajustement Then
    XPrt.FontSize = 14
    prtCHQ_SCAN_List1_NewLine
    Call frmElpPrt.prtTrame(prtMinX + 3420, XPrt.CurrentY, prtMinX + 10780, XPrt.CurrentY + prtlineHeight + 100, "b", 240)
    frmElpPrt.prtCentré prtMedX, "??? ERREUR : Remise non ajustée ???"
End If
XPrt.FontBold = False

prtCHQ_SCAN_List1_Colonne
Call frmElpPrt.prtEndDoc(1000)
frmElpPrt.Hide
Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub


Public Sub prtCHQ_SCAN_List1_Open()
On Error GoTo prtError

Set XPrt = Printer
frmElpPrt.Show vbModeless

Height8_6 = frmElpPrt.prtHeightDelta(8, 6)

blnFiligrane = False
prtOrientation = vbPRORLandscape '
prtPgmName = "prtCHQ_SCAN"
prtTitleUsr = usrName
prtTitleText = "BIA : Liste des remises en banque"

prtLineNb = 1
prtlineHeight = 300
prtHeaderHeight = 50 ' 100

nbR = 0: curR = 0
nbT = 0: curT = 0
nbChqR = 0: nbChqT = 0
mRefInterne = ""
blnErreur_Total = False
blnErreur_Ajustement = False

prtFormType = ""
frmElpPrt.prtStdInit
prtCHQ_SCAN_List1_Form


Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub



Public Sub prtCHQ_SCAN_List1_Form()
Dim wId As String
Dim X As String

XPrt.FontSize = 8
XPrt.FontBold = True
XPrt.DrawWidth = 2
'XPrt.Line (prtMaxX, prtMinY)-(prtMaxX, prtMaxY), prtLineColor

XPrt.CurrentY = prtMinY + 50

XPrt.CurrentX = prtMinX
XPrt.Print "Date";
XPrt.CurrentX = prtMinX + 1500
XPrt.Print "Réf Interne";

XPrt.CurrentX = prtMinX + 2800
XPrt.Print "N° lot";
XPrt.CurrentX = prtMinX + 3500
XPrt.Print "Compte";
XPrt.CurrentX = prtMinX + 5000
XPrt.Print "Intitulé";

XPrt.CurrentX = prtMinX + 8500
XPrt.Print "Total de la remise";
XPrt.CurrentX = prtMinX + 10200
XPrt.Print "Dev";
XPrt.CurrentX = prtMinX + 11100
XPrt.Print "Nb Chèques";
XPrt.CurrentX = prtMinX + 12200
XPrt.Print "Nature";
XPrt.CurrentX = prtMinX + 12900
XPrt.Print "Réf Client";

'XPrt.FontSize = 8
XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY), prtLineColor
XPrt.CurrentY = XPrt.CurrentY + 100


End Sub

Public Sub prtCHQ_SCAN_List1_Colonne()
Dim wId As String
Dim X As String

XPrt.DrawWidth = 2
XPrt.Line (prtMinX + 3400, prtMinY)-(prtMinX + 3400, prtMaxY), prtLineColor
XPrt.Line (prtMinX + 10800, prtMinY)-(prtMinX + 10800, prtMaxY), prtLineColor
XPrt.Line (prtMinX + 12100, prtMinY)-(prtMinX + 12100, prtMaxY), prtLineColor

End Sub


Public Sub prtCHQ_SCAN_List1_Line(lCHQ_SCAN As typeCHQ_SCAN, lCOMPTEINT As String, lNb As Long, lcurTotal As Currency)
Dim X As String, curX As Currency

If mRefInterne <> lCHQ_SCAN.RefInterne Then
    prtCHQ_SCAN_List1_Rupture
    mRefInterne = lCHQ_SCAN.RefInterne
End If

prtCHQ_SCAN_List1_NewLine

XPrt.CurrentX = prtMinX
XPrt.Print dateImp10(lCHQ_SCAN.Date);
XPrt.CurrentX = prtMinX + 1500
XPrt.Print lCHQ_SCAN.RefInterne;
If IsNumeric(lCHQ_SCAN.CRem) Then
    X = Format$(CLng(lCHQ_SCAN.CRem), "### ### ###")
    XPrt.CurrentX = prtMinX + 3300 - XPrt.TextWidth(X)

Else
    XPrt.CurrentX = prtMinX + 2500
    X = lCHQ_SCAN.CRem
End If
XPrt.Print X;

XPrt.FontBold = True
XPrt.CurrentX = prtMinX + 3500
If IsNumeric(lCHQ_SCAN.COMPTE) Then
    X = Format$(Val(lCHQ_SCAN.COMPTE), "##### ### ###")
Else
    X = lCHQ_SCAN.COMPTE
End If
XPrt.Print X;

XPrt.FontBold = False
curX = CCur(lCHQ_SCAN.Zone1) / 100
X = Format$(curX, "### ### ### ###.00")
XPrt.CurrentX = prtMinX + 10000 - XPrt.TextWidth(X)
XPrt.Print X;

XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 6

If lCHQ_SCAN.Id = "R" Then
    X = lCOMPTEINT
Else
    X = "???? " & lCHQ_SCAN.Cmc7
End If
XPrt.CurrentX = prtMinX + 5000
XPrt.Print X;

XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 8
XPrt.CurrentX = prtMinX + 10200
XPrt.Print lCHQ_SCAN.Devise;

X = Format$(lNb, "### ### ##0")
XPrt.CurrentX = prtMinX + 12000 - XPrt.TextWidth(X)
XPrt.Print X;
nbChqR = nbChqR + lNb
nbChqT = nbChqT + lNb

XPrt.CurrentX = prtMinX + 12200
XPrt.Print lCHQ_SCAN.Nature;
XPrt.CurrentX = prtMinX + 12900
XPrt.Print lCHQ_SCAN.RefClient;
If curX <> lcurTotal Then
    prtCHQ_SCAN_List1_NewLine
    XPrt.FontBold = True
    XPrt.CurrentX = prtMinX + 5000: XPrt.Print "###### ECART : total cumulé des chèques = " & lcurTotal
    XPrt.FontBold = False
    blnErreur_Total = True
End If

If lCHQ_SCAN.StatutRem <> "AJ" Then
    prtCHQ_SCAN_List1_NewLine
    XPrt.FontBold = True
    XPrt.CurrentX = prtMinX + 5000: XPrt.Print "###### REMISE NON AJUSTEE";
    XPrt.FontBold = False
    blnErreur_Total = True
    blnErreur_Ajustement = True
End If

curR = curR + curX: nbR = nbR + 1
curT = curT + curX: nbT = nbT + 1
End Sub





Public Sub prtCHQ_SCAN_List1_NewLine()
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
If XPrt.CurrentY + 300 > prtMaxY Then
    prtCHQ_SCAN_List1_Colonne
    frmElpPrt.prtNewPage
    prtCHQ_SCAN_List1_Form
End If

End Sub



Public Sub prtCHQ_SCAN_List1_Rupture()
Dim X As String
If nbR > 0 Then
    XPrt.FontBold = True
    XPrt.CurrentY = XPrt.CurrentY + 100
    prtCHQ_SCAN_List1_NewLine
    Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY - 20, prtMaxX, XPrt.CurrentY + prtlineHeight - 50, " ", 240)
    X = Format$(curR, "### ### ### ###.00")
    XPrt.CurrentX = prtMinX + 10000 - XPrt.TextWidth(X)
    XPrt.Print X;
    XPrt.CurrentX = prtMinX + 5000
    If nbR <= 1 Then
        XPrt.Print nbR & "     bordereau de remise en banque";
    Else
        XPrt.Print nbR & "     bordereaux de remise en banque";
    End If
    XPrt.CurrentX = prtMinX + 1500
    XPrt.Print mRefInterne;
    X = Format$(nbChqR, "### ### ##0")
    XPrt.CurrentX = prtMinX + 12000 - XPrt.TextWidth(X)
    XPrt.Print X;

    XPrt.FontBold = False
   prtCHQ_SCAN_List1_NewLine
End If
nbR = 0: curR = 0
nbChqR = 0
End Sub
Public Sub prtCHQ_SCAN_Rapprochement_Close()
Dim X As String
On Error GoTo prtError
XPrt.FontBold = False

prtCHQ_SCAN_Rapprochement_Colonne
Call frmElpPrt.prtEndDoc(1000)
frmElpPrt.Hide
Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub

Public Sub prtCHQ_SCAN_Rapprochement_Colonne()
Dim wId As String
Dim X As String

XPrt.DrawWidth = 2
XPrt.Line (prtMinX + 6000, prtMinY)-(prtMinX + 6000, prtMaxY), prtLineColor
XPrt.Line (prtMinX + 10100, prtMinY)-(prtMinX + 10100, prtMaxY), prtLineColor
XPrt.DrawWidth = 10
XPrt.Line (prtMinX + 8100, prtMinY)-(prtMinX + 8100, prtMaxY), prtLineColor

End Sub




Public Sub prtCHQ_SCAN_Rapprochement_Form()
Dim wId As String
Dim X As String

XPrt.FontSize = 8
XPrt.FontBold = True
XPrt.DrawWidth = 7
'XPrt.Line (prtMaxX, prtMinY)-(prtMaxX, prtMaxY), prtLineColor

XPrt.CurrentY = prtMinY
Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY, prtMaxX, prtMinY + prtHeaderHeight, "B", 245)

XPrt.CurrentY = prtMinY + 100

XPrt.CurrentX = prtMinX
XPrt.Print "Sta";
XPrt.CurrentX = prtMinX + 500
XPrt.Print "Service";
XPrt.CurrentX = prtMinX + 1200
XPrt.Print "Nature";

XPrt.CurrentX = prtMinX + 2000
XPrt.Print "Dossier";
XPrt.CurrentX = prtMinX + 3300
XPrt.Print "Date";
XPrt.CurrentX = prtMinX + 4000
XPrt.Print "Compte";

XPrt.CurrentX = prtMinX + 6200
XPrt.Print "Total remise SAB";


XPrt.CurrentX = prtMinX
XPrt.Print "Sta";
XPrt.CurrentX = prtMinX + 8200
XPrt.Print "Total remise SCAN";
XPrt.CurrentX = prtMinX + 14000
XPrt.Print "Nb Chèques";
XPrt.CurrentX = prtMinX + 15000
XPrt.Print "N° remise";
XPrt.CurrentX = prtMinX + 13300
XPrt.Print "Date";
XPrt.CurrentX = prtMinX + 10500
XPrt.Print "Compte";

'XPrt.FontSize = 8
XPrt.FontBold = False

XPrt.CurrentY = prtMinX + prtHeaderHeight + 100


End Sub
Public Sub prtCHQ_SCAN_Rapprochement_Line(lYCHQMON0 As typeYCHQMON0)
Dim X As String, curX As Currency



If lYCHQMON0.CHQRC1ETA = 1 Then
    XPrt.CurrentX = prtMinX + 50: XPrt.Print lYCHQMON0.CHQMONSTA;
    XPrt.CurrentX = prtMinX + 500: XPrt.Print lYCHQMON0.CHQRC1SER & " " & lYCHQMON0.CHQRC1SSE;
    XPrt.CurrentX = prtMinX + 2000: XPrt.Print Val(lYCHQMON0.CHQRC1DOS);
    XPrt.CurrentX = prtMinX + 3000: XPrt.Print dateIBM10(lYCHQMON0.CHQRC1DCR, True);
    XPrt.CurrentX = prtMinX + 4000: XPrt.Print lYCHQMON0.CHQCOMPTE;
    
    XPrt.FontBold = True
    XPrt.CurrentX = prtMinX + 1500: XPrt.Print lYCHQMON0.CHQCREM;
    curX = CCur(lYCHQMON0.CHQMONTANT)
    X = Format$(curX, "### ### ### ###.00")
    XPrt.CurrentX = prtMinX + 7500 - XPrt.TextWidth(X)
    XPrt.Print X;
    XPrt.CurrentX = prtMinX + 7600: XPrt.Print lYCHQMON0.CHQDEVISE;
    XPrt.FontBold = False
Else
    XPrt.CurrentX = prtMinX + 8050: XPrt.Print lYCHQMON0.CHQMONSTA;
    XPrt.CurrentX = prtMinX + 15000: XPrt.Print Val(lYCHQMON0.CHQCREM);

    'xxxx correction 25/11/2009 - la date d'origine est au format aaaammjj
    'XPrt.CurrentX = prtMinX + 13000: XPrt.Print dateIBM10(lYCHQMON0.CHQDATE, True);
    XPrt.CurrentX = prtMinX + 13000: XPrt.Print Mid(lYCHQMON0.CHQDATE, 7, 2) & "." & Mid(lYCHQMON0.CHQDATE, 5, 2) & "." & Left(lYCHQMON0.CHQDATE, 4);
    'xxxxFIN correction 25/11/2009
    XPrt.CurrentX = prtMinX + 10500: XPrt.Print lYCHQMON0.CHQCOMPTE;
    
    XPrt.FontBold = True
    X = Format$(Val(lYCHQMON0.CHQNB), "### ### ##0")
    XPrt.CurrentX = prtMinX + 14700 - XPrt.TextWidth(X)
    XPrt.Print X;
    curX = CCur(lYCHQMON0.CHQMONTANT)
    X = Format$(curX, "### ### ### ###.00")
    XPrt.CurrentX = prtMinX + 9500 - XPrt.TextWidth(X)
    XPrt.Print X;
    XPrt.CurrentX = prtMinX + 9600: XPrt.Print lYCHQMON0.CHQDEVISE;
    XPrt.FontBold = False
End If

End Sub

Public Sub prtCHQ_SCAN_Rapprochement_NewLine()
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
If XPrt.CurrentY + 300 > prtMaxY Then
    prtCHQ_SCAN_Rapprochement_Colonne
    frmElpPrt.prtNewPage
    prtCHQ_SCAN_Rapprochement_Form
End If

End Sub









Public Sub prtCHQ_SCAN_Rapprochement_Open()
On Error GoTo prtError

Set XPrt = Printer
frmElpPrt.Show vbModeless

Height8_6 = frmElpPrt.prtHeightDelta(8, 6)

blnFiligrane = False
prtOrientation = vbPRORLandscape '
prtPgmName = "prtCHQ_SCAN"
prtTitleUsr = usrName
prtTitleText = "Liste de rapprochement des remises en banque"

prtLineNb = 1
prtlineHeight = 300
prtHeaderHeight = 400

nbR = 0: curR = 0
nbT = 0: curT = 0
nbChqR = 0: nbChqT = 0
mRefInterne = ""
blnErreur_Total = False
blnErreur_Ajustement = False

prtFormType = ""
frmElpPrt.prtStdInit
prtCHQ_SCAN_Rapprochement_Form


Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub

Public Sub prtCHQ_SCAN_List2_Close()
Dim X As String
On Error GoTo prtError
prtCHQ_SCAN_List2_Rupture
'XPrt.CurrentY = XPrt.CurrentY + 100
'XPrt.FontBold = True
'prtCHQ_SCAN_List2_NewLine
'Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY, prtMaxX, XPrt.CurrentY + prtlineHeight + 100, "B", 245)
'XPrt.CurrentY = XPrt.CurrentY + 100
'X = Format$(curT, "### ### ### ###.00")
'XPrt.CurrentX = prtMinX + 10500 - XPrt.TextWidth(X)
'XPrt.Print X;
'XPrt.CurrentX = prtMinX + 6000
'    If nbT <= 1 Then
'        XPrt.Print nbT & "     chèque remis en banque";
'    Else
'        XPrt.Print nbT & "     chèques remis en banque";
'    End If
'XPrt.CurrentX = prtMinX + 1500
'XPrt.Print "TOTAL;"
'nbT = 0: curT = 0:
'XPrt.FontBold = False

prtCHQ_SCAN_List2_Colonne
Call frmElpPrt.prtEndDoc(1000)
frmElpPrt.Hide
Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub


Public Sub prtCHQ_SCAN_List2_Open()
On Error GoTo prtError

Set XPrt = Printer
frmElpPrt.Show vbModeless

'Height8_6 = frmElpPrt.prtHeightDelta(8, 6)

blnFiligrane = False
prtOrientation = vbPRORPortrait '
prtPgmName = "prtCHQ_SCAN"
prtTitleUsr = usrName
prtTitleText = "SG: Remise en banque"

prtLineNb = 1
prtlineHeight = 250
prtHeaderHeight = 50 ' 100

nbR = 0: curR = 0
nbT = 0: curT = 0
Page_No = 0
mRefInterne = ""

prtFormType = ""
'frmElpPrt.prtStdInit
prtSocInit

prtCHQ_SCAN_List2_Form

Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub



Public Sub prtCHQ_SCAN_List2_Form()
Dim wId As String
Dim X As String

Page_No = Page_No + 1

XPrt.FontBold = True
XPrt.DrawWidth = 2
'XPrt.Line (prtMaxX, prtMinY)-(prtMaxX, prtMaxY), prtLineColor

'XPrt.CurrentY = prtMinY + 50
XPrt.CurrentY = 1700
XPrt.FontSize = 12
XPrt.FontUnderline = True
frmElpPrt.prtCentré prtMedX, "SOCIETE GENERALE - Remise de chèques"
XPrt.FontUnderline = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3
XPrt.CurrentX = prtMinX + 1200
XPrt.Print "30003";
XPrt.CurrentX = prtMinX + 3600
XPrt.Print "04970";
XPrt.CurrentX = prtMinX + 6100
XPrt.Print "00001080042";
XPrt.CurrentX = prtMinX + 8500
XPrt.Print "54";

XPrt.FontUnderline = True
XPrt.CurrentY = XPrt.CurrentY + 50
XPrt.FontSize = 8
XPrt.CurrentX = prtMinX
XPrt.Print "code banque";
XPrt.CurrentX = prtMinX + 2300
XPrt.Print "code guichet";
XPrt.CurrentX = prtMinX + 4500
XPrt.Print "numéro de compte";
XPrt.CurrentX = prtMinX + 7700
XPrt.Print "clé RIB";

XPrt.FontUnderline = False
XPrt.CurrentX = prtMinX + 10000
XPrt.Print "Page : " & Page_No;

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 3

Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY - 20, prtMaxX, XPrt.CurrentY + prtlineHeight - 50, "B", 245)

XPrt.CurrentX = prtMinX
XPrt.Print " Lot";
XPrt.CurrentX = prtMinX + 1500
XPrt.Print "Image";

XPrt.CurrentX = prtMinX + 3000
XPrt.Print "N° Chèque";
XPrt.CurrentX = prtMinX + 4900
XPrt.Print "ZIN";
XPrt.CurrentX = prtMinX + 6400
XPrt.Print "ZIB";
XPrt.CurrentX = prtMinX + 9900
XPrt.Print "Montant";

'XPrt.FontSize = 8
XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
'XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY), prtLineColor
XPrt.CurrentY = XPrt.CurrentY + 50


End Sub

Public Sub prtCHQ_SCAN_List2_Colonne()
Dim wId As String
Dim X As String

'XPrt.DrawWidth = 2
'XPrt.Line (prtMinX + 3400, prtMinY)-(prtMinX + 3400, prtMaxY)
'XPrt.Line (prtMinX + 9800, prtMinY)-(prtMinX + 9800, prtMaxY)
'XPrt.Line (prtMinX + 11100, prtMinY)-(prtMinX + 11100, prtMaxY)

End Sub


Public Sub prtCHQ_SCAN_List2_Line(lCHQ_SCAN As typeCHQ_SCAN, lCHQ_SCAN_Remise As typeCHQ_SCAN)
Dim X As String, curX As Currency

If mRefInterne <> lCHQ_SCAN_Remise.RefInterne Then
    prtCHQ_SCAN_List2_Rupture
    mRefInterne = lCHQ_SCAN_Remise.RefInterne
End If

prtCHQ_SCAN_List2_NewLine

XPrt.CurrentX = prtMinX
XPrt.Print lCHQ_SCAN.CRem;
XPrt.CurrentX = prtMinX + 1500
XPrt.Print lCHQ_SCAN.IMAGE;
XPrt.CurrentX = prtMinX + 3000
XPrt.Print lCHQ_SCAN.Zone4;
XPrt.CurrentX = prtMinX + 4500
XPrt.Print lCHQ_SCAN.Zone3;
XPrt.CurrentX = prtMinX + 6000
XPrt.Print lCHQ_SCAN.Zone2;

curX = CCur(lCHQ_SCAN.Zone1) / 100
X = Format$(curX, "### ### ### ###.00")
XPrt.CurrentX = prtMinX + 10500 - XPrt.TextWidth(X)
XPrt.Print X;

curR = curR + curX: nbR = nbR + 1
curT = curT + curX: nbT = nbT + 1
End Sub





Public Sub prtCHQ_SCAN_List2_NewLine()
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
If XPrt.CurrentY + 300 > prtMaxY Then
    prtCHQ_SCAN_List2_Colonne
    frmElpPrt.prtNewPage
    prtCHQ_SCAN_List2_Form
End If

End Sub



Public Sub prtCHQ_SCAN_List2_Rupture()
Dim X As String
If nbR > 0 Then
     XPrt.FontBold = True
    prtCHQ_SCAN_List2_NewLine
    prtCHQ_SCAN_List2_NewLine
    Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY - 20, prtMaxX, XPrt.CurrentY + prtlineHeight - 50, " ", 240)
    X = Format$(curR, "### ### ### ###.00")
    XPrt.CurrentX = prtMinX + 10500 - XPrt.TextWidth(X)
    XPrt.Print X;
    XPrt.CurrentX = prtMinX + 6000
    If nbR <= 1 Then
        XPrt.Print nbR & "     chèque remis en banque";
    Else
        XPrt.Print nbR & "     chèques remis en banque";
    End If
    XPrt.CurrentX = prtMinX
    XPrt.Print "Référence bordereau SG : " & mRefInterne;
    XPrt.FontBold = False
    
    XPrt.CurrentY = prtMaxY
   prtCHQ_SCAN_List2_NewLine
End If
nbR = 0: curR = 0

End Sub

Public Sub prtCHQ_Stat_Close(lCHQ_Stat As typeCHQ_Stat)
Dim X As String
On Error GoTo prtError

XPrt.CurrentY = XPrt.CurrentY + 100
XPrt.FontBold = True
prtCHQ_Stat_NewLine

XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY), prtLineColor
Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY + 20, prtMaxX, XPrt.CurrentY + prtlineHeight + 100, " ", 245)
XPrt.CurrentY = XPrt.CurrentY - prtlineHeight + 100

prtCHQ_Stat_Line lCHQ_Stat

prtCHQ_Stat_Colonne
Call frmElpPrt.prtEndDoc(1000)
frmElpPrt.Hide
Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub

Public Sub prtCHQ_imgCHQ_Close()
Dim X As String
On Error GoTo prtError

Call frmElpPrt.prtEndDoc(1000)
frmElpPrt.Hide
Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub

Public Sub prtCHQ_Stat_Open(lText As String)
On Error GoTo prtError

Set XPrt = Printer
frmElpPrt.Show vbModeless

Height8_6 = frmElpPrt.prtHeightDelta(8, 6)

blnFiligrane = False
prtOrientation = vbPRORLandscape '
prtPgmName = "prtCHQ_Stat"
prtTitleUsr = usrName
prtTitleText = lText
prtLineNb = 1
prtlineHeight = 300
prtHeaderHeight = 500


prtFormType = ""
frmElpPrt.prtStdInit
prtCHQ_Stat_Form


Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub


Public Sub prtCHQ_imgCHQ_Open(lText As String)
On Error GoTo prtError

Set XPrt = Printer
frmElpPrt.Show vbModeless

Height8_6 = frmElpPrt.prtHeightDelta(8, 6)

blnFiligrane = False
prtOrientation = vbPRORPortrait '
prtPgmName = "prtCHQ_imgCHQ"
prtTitleUsr = usrName
prtTitleText = lText
prtLineNb = 1
prtlineHeight = 300
prtHeaderHeight = 500


prtFormType = ""
frmElpPrt.prtStdInit

Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub

Public Sub prtCHQ_Stat_Form()
Dim wId As String
Dim X As String

XPrt.FontSize = 8
XPrt.FontBold = True
XPrt.DrawWidth = 7
'XPrt.Line (prtMaxX, prtMinY)-(prtMaxX, prtMaxY)

XPrt.CurrentY = prtMinY
Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY, prtMaxX, prtMinY + prtHeaderHeight, "B", 245)
XPrt.CurrentY = prtMinY + 50

XPrt.CurrentX = prtMinX + 200
XPrt.Print "Date";
XPrt.CurrentX = prtMinX + 2000
XPrt.Print "Total des chèques en EUR";
XPrt.CurrentX = prtMinX + 5700
XPrt.Print "Répartition des montants";
XPrt.CurrentX = prtMinX + 9000
XPrt.Print "S.G.";
XPrt.CurrentX = prtMinX + 11000
XPrt.Print "B.I.A.";
XPrt.CurrentX = prtMinX + 12600
XPrt.Print "Autres natures";

XPrt.CurrentX = prtMinX + 14100
XPrt.Print "CHEQUES en DEVISE";

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight - Height8_6
XPrt.FontSize = 6
XPrt.CurrentX = prtMinX + 1500
XPrt.Print "Remises";
XPrt.CurrentX = prtMinX + 2500
XPrt.Print "Chèques";
XPrt.CurrentX = prtMinX + 4500
XPrt.Print "Montant";

XPrt.CurrentX = prtMinX + 5600
XPrt.Print "< " & CStr(getSeuil1YEICGCC0);
XPrt.CurrentX = prtMinX + 6600
XPrt.Print "< " & CStr(getSeuil2YEICGCC0);
XPrt.CurrentX = prtMinX + 7800
XPrt.Print "=>";

XPrt.CurrentX = prtMinX + 8500
XPrt.Print "Remises";
XPrt.CurrentX = prtMinX + 9500
XPrt.Print "Chèques";
XPrt.CurrentX = prtMinX + 10500
XPrt.Print "Remises";
XPrt.CurrentX = prtMinX + 11500
XPrt.Print "Chèques";


XPrt.CurrentX = prtMinX + 12500
XPrt.Print "Remises";
XPrt.CurrentX = prtMinX + 13500
XPrt.Print "Chèques";
XPrt.CurrentX = prtMinX + 14300
XPrt.Print "Remises";
XPrt.CurrentX = prtMinX + 15200
XPrt.Print "Chèques";


XPrt.FontBold = False
XPrt.FontSize = 8
XPrt.DrawWidth = 2

'XPrt.CurrentY = XPrt.CurrentY + prtlineHeight - Height8_6
'XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY)
XPrt.CurrentY = prtMinY + prtHeaderHeight + 100


End Sub

Public Sub prtCHQ_Stat_Colonne()
Dim wId As String
Dim X As String

XPrt.DrawWidth = 2
XPrt.Line (prtMinX + 1100, prtMinY)-(prtMinX + 1100, prtMaxY), prtLineColor
XPrt.Line (prtMinX + 5100, prtMinY)-(prtMinX + 5100, prtMaxY), prtLineColor
XPrt.Line (prtMinX + 8100, prtMinY)-(prtMinX + 8100, prtMaxY), prtLineColor
XPrt.Line (prtMinX + 10100, prtMinY)-(prtMinX + 10100, prtMaxY), prtLineColor
XPrt.Line (prtMinX + 12100, prtMinY)-(prtMinX + 12100, prtMaxY), prtLineColor
XPrt.Line (prtMinX + 14100, prtMinY)-(prtMinX + 14100, prtMaxY), prtLineColor

End Sub


Public Sub prtCHQ_Stat_Line(lCHQ_Stat As typeCHQ_Stat)
Dim X As String, curX As Currency


prtCHQ_Stat_NewLine

XPrt.CurrentX = prtMinX
XPrt.Print dateImp10(lCHQ_Stat.Date);

nbR = lCHQ_Stat.Remise_SG + lCHQ_Stat.Remise_BIA + lCHQ_Stat.Remise_Divers
X = Format$(nbR, "### ### ###")
XPrt.CurrentX = prtMinX + 2000 - XPrt.TextWidth(X)
XPrt.Print X;
nbChqR = lCHQ_Stat.Chèque_SG + lCHQ_Stat.Chèque_BIA + lCHQ_Stat.Chèque_Divers
X = Format$(nbChqR, "### ### ###")
XPrt.CurrentX = prtMinX + 3000 - XPrt.TextWidth(X)
XPrt.Print X;
curR = lCHQ_Stat.Montant_SG + lCHQ_Stat.Montant_BIA + lCHQ_Stat.Montant_Divers
X = Format$(curR, "### ### ### ###.00")
XPrt.CurrentX = prtMinX + 5000 - XPrt.TextWidth(X)
XPrt.Print X;

X = Format$(lCHQ_Stat.Remise_Nb1, "### ### ###")
XPrt.CurrentX = prtMinX + 6000 - XPrt.TextWidth(X)
XPrt.Print X;
X = Format$(lCHQ_Stat.Remise_Nb2, "### ### ###")
XPrt.CurrentX = prtMinX + 7000 - XPrt.TextWidth(X)
XPrt.Print X;
X = Format$(lCHQ_Stat.Remise_Nb3, "### ### ###")
XPrt.CurrentX = prtMinX + 8000 - XPrt.TextWidth(X)
XPrt.Print X;

X = Format$(lCHQ_Stat.Remise_SG, "### ### ###")
XPrt.CurrentX = prtMinX + 9000 - XPrt.TextWidth(X)
XPrt.Print X;
X = Format$(lCHQ_Stat.Chèque_SG, "### ### ###")
XPrt.CurrentX = prtMinX + 10000 - XPrt.TextWidth(X)
XPrt.Print X;

X = Format$(lCHQ_Stat.Remise_BIA, "### ### ###")
XPrt.CurrentX = prtMinX + 11000 - XPrt.TextWidth(X)
XPrt.Print X;
X = Format$(lCHQ_Stat.Chèque_BIA, "### ### ###")
XPrt.CurrentX = prtMinX + 12000 - XPrt.TextWidth(X)
XPrt.Print X;

X = Format$(lCHQ_Stat.Remise_Divers, "### ### ###")
XPrt.CurrentX = prtMinX + 13000 - XPrt.TextWidth(X)
XPrt.Print X;
X = Format$(lCHQ_Stat.Chèque_Divers, "### ### ###")
XPrt.CurrentX = prtMinX + 14000 - XPrt.TextWidth(X)
XPrt.Print X;

X = Format$(lCHQ_Stat.Remise_Devise, "### ### ###")
XPrt.CurrentX = prtMinX + 14800 - XPrt.TextWidth(X)
XPrt.Print X;
X = Format$(lCHQ_Stat.Chèque_Devise, "### ### ###")
XPrt.CurrentX = prtMaxX - XPrt.TextWidth(X) - 50
XPrt.Print X;

End Sub



Public Sub prtCHQ_Stat_Z(lCHQ_Stat As typeCHQ_Stat)
lCHQ_Stat.Date = ""
lCHQ_Stat.Nature = ""
      
lCHQ_Stat.Remise_SG = 0
lCHQ_Stat.Chèque_SG = 0
lCHQ_Stat.Montant_SG = 0

lCHQ_Stat.Remise_BIA = 0
lCHQ_Stat.Chèque_BIA = 0
lCHQ_Stat.Montant_BIA = 0

lCHQ_Stat.Remise_Divers = 0
lCHQ_Stat.Chèque_Divers = 0
lCHQ_Stat.Montant_Divers = 0

lCHQ_Stat.Remise_Devise = 0
lCHQ_Stat.Chèque_Devise = 0

lCHQ_Stat.Remise_Nb1 = 0
lCHQ_Stat.Remise_Nb2 = 0
lCHQ_Stat.Remise_Nb3 = 0


End Sub



Public Sub prtCHQ_Stat_NewLine()
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
If XPrt.CurrentY + 300 > prtMaxY Then
    prtCHQ_Stat_Colonne
    frmElpPrt.prtNewPage
    prtCHQ_Stat_Form
End If

End Sub



