Attribute VB_Name = "prtSWI_Messages"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------


Dim nbR As Long

Dim wBoo As Boolean

Public Sub prtSWI_Messages_List6_Close(l_unit_name As String, lNrequest As String)

Dim X As String
On Error GoTo prtError

wBoo = True
prtSWI_Messages_List6_Rupture l_unit_name, lNrequest

frmElpPrt.prtEndDoc 1000
frmElpPrt.Hide
Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub


Public Sub prtSWI_Messages_List6_Close_xlsManual(l_unit_name As String, lNrequest As String, ByRef currentrow As Long, wsExcel As Excel.Worksheet, ByRef comptageRows As Long, maxRows As Long, maxRowsPlus As Long)
Dim X As String
On Error GoTo prtError

wBoo = True
Call prtSWI_Messages_List6_Rupture_xlsManual(l_unit_name, lNrequest, currentrow, wsExcel, comptageRows, maxRows, maxRowsPlus)

Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")

End Sub

Public Sub prtSWI_Messages_List6_Line_xlsManual(lNrequest As String, lrMesg As typerMesg, lrAppe_E As typerAppe, lrAppe_R As typerAppe, lrInst As typerInst, ByRef currentrow As Long, wsExcel As Excel.Worksheet, ByRef comptageRows As Long, maxRows As Long, maxRowsPlus As Long)
Dim X As String, xDev As String, xCur As String
Dim curX As Currency, K As Integer

wBoo = False
If currentrow >= maxRows + maxRowsPlus Then
    If comptageRows >= maxRows Then
        Call insere_entete_page(wsExcel, "1:3", 3, currentrow)
        comptageRows = 3
        currentrow = currentrow + 3
    End If
End If
comptageRows = comptageRows + 1
currentrow = currentrow + 1
Range("5:5").Select
Selection.Copy
Range("A" & CStr(currentrow)).Select
ActiveSheet.Paste
wsExcel.Cells(currentrow, 1) = lrMesg.mesg_type
wsExcel.Cells(currentrow, 2) = "'" & lrMesg.mesg_crea_date_time
wsExcel.Cells(currentrow, 3) = lrMesg.mesg_trn_ref
K = 0
X = lrMesg.mesg_fin_ccy_amount
xDev = Space_Scan(X, K)
wsExcel.Cells(currentrow, 4) = xDev
xCur = num_CDec_USA(Space_Scan(X, K))
X = Format$(xCur, "### ### ### ###.00")
If xCur <> 0 Then
    wsExcel.Cells(currentrow, 5) = "'" & X
End If
If Not IsNull(lrMesg.mesg_fin_value_date) Then
    wsExcel.Cells(currentrow, 6) = dateAMJ6_Imp10(lrMesg.mesg_fin_value_date)
End If
wsExcel.Cells(currentrow, 7) = lrMesg.mesg_receiver_swift_address
Select Case lNrequest
    Case "6 ": wsExcel.Cells(currentrow, 8) = lrMesg.mesg_crea_oper_nickname
    Case "7 ": wsExcel.Cells(currentrow, 8) = lrMesg.mesg_mod_oper_nickname
End Select
wsExcel.Cells(currentrow, 9) = lrInst.inst_auth_oper_nickname
X = lrMesg.mesg_status
If IsNull(lrAppe_E.appe_network_delivery_status) Then
    X = X & " " & lrAppe_R.appe_network_delivery_status
Else
    X = X & " " & lrAppe_E.appe_network_delivery_status
End If
wsExcel.Cells(currentrow, 10) = X
' + 1 dans nbR
nbR = nbR + 1
End Sub

Public Sub prtSWI_Messages_List6_Open(lNrequest As String, C1 As Control, C2 As Control)
Dim Amj_Deb As String, Amj_Fin As String

On Error GoTo prtError

Set XPrt = Printer
frmElpPrt.Show vbModeless

'Height8_6 = frmElpPrt.prtHeightDelta(8, 6)

blnFiligrane = False
prtOrientation = vbPRORLandscape
prtPgmName = "prtSWI_Messages"
prtTitleUsr = usrName

' Période de sélection
Call DTPicker_Control(C1, Amj_Deb)
Call DTPicker_Control(C2, Amj_Fin)

' Attention au no de la requête pour le titre de la liste à imprimer
Select Case lNrequest
    Case "6 ": prtTitleText = "Liste des Messages créés dans SWIFT ALLIANCE du " & dateImp10(Amj_Deb) & " au " & dateImp10(Amj_Fin)
    Case "7 ": prtTitleText = "Liste des Messages automatiques modifiés dans SWIFT ALLIANCE du " & dateImp10(Amj_Deb) & " au " & dateImp10(Amj_Fin)
    Case "8 ": prtTitleText = "Répertition des messages SWIFT ALLIANCE du " & dateImp10(Amj_Deb) & " au " & dateImp10(Amj_Fin)
End Select

prtLineNb = 1
prtlineHeight = 250
prtHeaderHeight = 50 ' 100

prtFormType = ""
frmElpPrt.prtStdInit
Select Case lNrequest
    Case "8 ": prtHeaderHeight = 300: prtSWI_Messages_List8_Form
    Case Else: prtSWI_Messages_List6_Form lNrequest
End Select

'XPrt.CurrentY = XPrt.CurrentY + prtlineHeight

nbR = 0
Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub

Public Sub prtSWI_Messages_List6_Form(lNrequest As String)
Dim wId As String
Dim X As String

XPrt.FontSize = 8
XPrt.FontBold = True
XPrt.DrawWidth = 2
'XPrt.Line (prtMaxX, prtMinY)-(prtMaxX, prtMaxY)

XPrt.CurrentY = prtMinY + 50

XPrt.CurrentX = prtMinX
XPrt.Print " MT ";
XPrt.CurrentX = prtMinX + 550
XPrt.Print "Crée le ...";

XPrt.CurrentX = prtMinX + 2500
XPrt.Print "Référence";
XPrt.CurrentX = prtMinX + 4200
XPrt.Print "DEV                    Montant";
XPrt.CurrentX = prtMinX + 7000
XPrt.Print "Valeur";

XPrt.CurrentX = prtMinX + 8500
XPrt.Print "Destinataire";

XPrt.CurrentX = prtMinX + 10200     ' Attention au no de la requête
Select Case lNrequest
    Case "6 ": XPrt.Print "Créé par";
    Case "7 ": XPrt.Print "Modifié par";
End Select

XPrt.CurrentX = prtMinX + 11400
XPrt.Print "Validé par";
XPrt.CurrentX = prtMinX + 12600
XPrt.Print "Etat";

'XPrt.FontSize = 8
XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY), prtLineColor
XPrt.CurrentY = XPrt.CurrentY + 50


End Sub


Public Sub prtSWI_Messages_List6_Open_xlsManual(lNrequest As String, C1 As Control, C2 As Control, wsExcel As Excel.Worksheet)
Dim Amj_Deb As String, Amj_Fin As String

On Error GoTo prtError

' Période de sélection
Call DTPicker_Control(C1, Amj_Deb)
Call DTPicker_Control(C2, Amj_Fin)

' Attention au no de la requête pour le titre de la liste à imprimer
Select Case lNrequest
    Case "6 ": prtTitleText = "Liste des Messages créés dans SWIFT ALLIANCE du " & dateImp10(Amj_Deb) & " au " & dateImp10(Amj_Fin)
    Case "7 ": prtTitleText = "Liste des Messages automatiques modifiés dans SWIFT ALLIANCE du " & dateImp10(Amj_Deb) & " au " & dateImp10(Amj_Fin)
    Case "8 ": prtTitleText = "Répertition des messages SWIFT ALLIANCE du " & dateImp10(Amj_Deb) & " au " & dateImp10(Amj_Fin) 'non utilisé
End Select
wsExcel.Cells(1, 4) = prtTitleText
nbR = 0
Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide
End Sub

Public Sub prtSWI_Messages_List6_Rupture_xlsManual(l_unit_name As String, lNrequest As String, ByRef currentrow As Long, wsExcel As Excel.Worksheet, ByRef comptageRows As Long, maxRows As Long, maxRowsPlus As Long)
Dim X As String

If nbR > 0 Then
    If currentrow >= maxRows + maxRowsPlus Then
        If comptageRows >= maxRows Then
            Call insere_entete_page(wsExcel, "1:3", 3, currentrow)
            comptageRows = 3
            currentrow = currentrow + 3
        End If
    End If
    comptageRows = comptageRows + 1
    currentrow = currentrow + 1
    'on saute 2 lignes
    If currentrow >= maxRows + maxRowsPlus Then
        If comptageRows >= maxRows Then
            Call insere_entete_page(wsExcel, "1:3", 3, currentrow)
            comptageRows = 3
            currentrow = currentrow + 3
        End If
    End If
    comptageRows = comptageRows + 1
    currentrow = currentrow + 1
    Range("7:7").Select
    Selection.Copy
    Range("A" & CStr(currentrow)).Select
    ActiveSheet.Paste
    If nbR > 1 Then
        wsExcel.Cells(currentrow, 6) = nbR & "      messages"
    Else
        wsExcel.Cells(currentrow, 6) = nbR & "      message"
    End If
    wsExcel.Cells(currentrow, 3) = l_unit_name
    If wBoo = False Then     ' Si prtSWI_Messages_List6_Close
        comptageRows = comptageRows + 1
        currentrow = currentrow + 1
        Call insere_entete_page(wsExcel, "1:3", 3, currentrow)
        comptageRows = 3
        currentrow = currentrow + 3
    End If
End If
nbR = 0

End Sub

Public Sub prtSWI_Messages_List8_Form()
Dim I As Integer
Dim X As String

XPrt.FontSize = 8
XPrt.FontBold = True
XPrt.DrawWidth = 2
'XPrt.Line (prtMaxX, prtMinY)-(prtMaxX, prtMaxY)

XPrt.CurrentY = prtMinY + 50

XPrt.CurrentX = prtMinX
XPrt.ForeColor = vbBlack
XPrt.Print " MT ";
XPrt.ForeColor = vbRed
XPrt.Print " Emis";
XPrt.ForeColor = vbBlue
XPrt.Print " Reçus";
XPrt.ForeColor = vbBlack

XPrt.CurrentX = prtMinX + 1300 * 1 + 500: XPrt.Print "BOTC";
XPrt.CurrentX = prtMinX + 1300 * 2 + 500: XPrt.Print "CSOP";
XPrt.CurrentX = prtMinX + 1300 * 3 + 500: XPrt.Print "DAFI";
XPrt.CurrentX = prtMinX + 1300 * 4 + 500: XPrt.Print "DCOM";
XPrt.CurrentX = prtMinX + 1300 * 5 + 500: XPrt.Print "ORPA";
XPrt.CurrentX = prtMinX + 1300 * 6 + 500: XPrt.Print "SCLE";
XPrt.CurrentX = prtMinX + 1300 * 7 + 500: XPrt.Print "SOBF";
XPrt.CurrentX = prtMinX + 1300 * 8 + 500: XPrt.Print "SOBI";
XPrt.CurrentX = prtMinX + 1300 * 9 + 500: XPrt.Print "Autres";
XPrt.CurrentX = prtMinX + 1300 * 10 + 500: XPrt.Print "None";
XPrt.CurrentX = prtMinX + 1300 * 11 + 500: XPrt.Print "Total";
'XPrt.FontSize = 8
XPrt.FontBold = False
XPrt.CurrentY = prtMinY + prtHeaderHeight
XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY), prtLineColor
XPrt.CurrentY = XPrt.CurrentY + 50
For I = 1 To 11
    XPrt.Line (prtMinX + 1300 * I, prtMinY)-(prtMinX + 1300 * I, prtMaxY), prtLineColor

Next I
XPrt.CurrentY = prtMinY + prtHeaderHeight + 50



End Sub

Public Sub prtSWI_Messages_List6_Line(lNrequest As String, lrMesg As typerMesg, lrAppe_E As typerAppe, lrAppe_R As typerAppe, lrInst As typerInst)

Dim X As String, xDev As String, xCur As String
Dim curX As Currency, K As Integer

wBoo = False


' Si fin de page

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
If XPrt.CurrentY + 300 > prtMaxY Then
    frmElpPrt.prtNewPage
    prtSWI_Messages_List6_Form lNrequest
End If

' Impression de la ligne courante

XPrt.CurrentX = prtMinX + 50
XPrt.Print lrMesg.mesg_type;
XPrt.CurrentX = prtMinX + 550
XPrt.Print lrMesg.mesg_crea_date_time;
XPrt.CurrentX = prtMinX + 2500
XPrt.Print lrMesg.mesg_trn_ref;

XPrt.CurrentX = prtMinX + 4200
K = 0
X = lrMesg.mesg_fin_ccy_amount
xDev = Space_Scan(X, K)
XPrt.Print xDev;
xCur = num_CDec_USA(Space_Scan(X, K))
X = Format$(xCur, "### ### ### ###.00")
XPrt.CurrentX = prtMinX + 6500 - XPrt.TextWidth(X)
If xCur <> 0 Then XPrt.Print X;

If Not IsNull(lrMesg.mesg_fin_value_date) Then
    XPrt.CurrentX = prtMinX + 7000
    XPrt.Print dateAMJ6_Imp10(lrMesg.mesg_fin_value_date);
End If

XPrt.CurrentX = prtMinX + 8500
XPrt.Print lrMesg.mesg_receiver_swift_address;

XPrt.CurrentX = prtMinX + 10200     ' Attention no de la requête
Select Case lNrequest
    Case "6 ": XPrt.Print lrMesg.mesg_crea_oper_nickname;
    Case "7 ": XPrt.Print lrMesg.mesg_mod_oper_nickname;
End Select

XPrt.CurrentX = prtMinX + 11400
XPrt.Print lrInst.inst_auth_oper_nickname;
XPrt.CurrentX = prtMinX + 12600
XPrt.Print lrMesg.mesg_status;

XPrt.CurrentX = prtMinX + 13800
If IsNull(lrAppe_E.appe_network_delivery_status) Then
    XPrt.Print lrAppe_R.appe_network_delivery_status;
Else
    XPrt.Print lrAppe_E.appe_network_delivery_status;
End If

' + 1 dans nbR
nbR = nbR + 1

End Sub







Public Sub prtSWI_Messages_List6_Rupture(l_unit_name As String, lNrequest As String)
Dim X As String

If nbR > 0 Then
    
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 2

    XPrt.FontBold = True
    Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY - 20, prtMaxX, XPrt.CurrentY + prtlineHeight - 50, " ", 240)
    XPrt.CurrentX = prtMinX + 7000
    XPrt.Print nbR & "      message(s)";
    XPrt.CurrentX = prtMinX + 2500
    XPrt.Print l_unit_name;
    XPrt.FontBold = False
    
    If wBoo = False Then     ' Si prtSWI_Messages_List6_Close
        XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
        frmElpPrt.prtNewPage
        prtSWI_Messages_List6_Form lNrequest
    End If
        
End If

nbR = 0


End Sub

