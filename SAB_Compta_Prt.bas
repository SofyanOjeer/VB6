Attribute VB_Name = "prtSAB_Compta"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim V
Dim X As String, I As Integer, Height8_6 As Integer

Dim blnPage As Boolean


Dim meYBIAMVT0 As typeYBIAMVT0, xYBIAMVT0 As typeYBIAMVT0
Dim colDb As Integer, colCr As Integer
Dim arrDev(100) As String, arrDev_Nb As Integer
Dim arrDb_N(100) As Currency, arrCr_N(100) As Currency
Dim arrDb_S(100) As Currency, arrCr_S(100) As Currency
Dim arrDb_T(100) As Currency, arrCr_T(100) As Currency
Dim arrDb_Recap(100) As Currency, arrCr_Recap(100) As Currency

Dim meDateComptable As String, meDateTraitement As String
Dim blnDateComptableMultiple As Boolean
Dim meService As String, xService As String
Dim meDevise As String, xDevise As String
Dim blnSéparateur As Boolean
Dim blnDossier As Boolean, blnDossierTrame As Boolean
Dim xLIBELLIB As String

Dim mTotal_Trame1 As Integer, mTotal_Trame2 As Integer, mTotal_Trame3 As Integer
Dim blnTotal As Boolean
 
 
Public Sub prtSAB_Compta_Compte_xlsManual(larrYBIAMVT0() As typeYBIAMVT0, fgW As MSFlexGrid, fgW_arrIndex As Integer, lMe As Form, blnDétail As Boolean, ByRef wbExcel As Excel.Workbook)
Dim wIndex As Long
'On Error Resume Next:
Call lstErr_Clear(lMe.lstErr, lMe.cmdContext, "Journal Devise> "): DoEvents
Call lstErr_AddItem(lMe.lstErr, lMe.cmdContext, "Journal Devise...."): DoEvents

prtSAB_Compta_Open blnDétail
blnSéparateur = False
blnDossier = False
blnDossierTrame = False

For I = 1 To fgW.Rows - 1
    
    fgW.Row = I
    fgW.Col = 1: xService = fgW.Text
    fgW.Col = 5: xDevise = fgW.Text
    fgW.Col = fgW_arrIndex
    wIndex = Val(fgW.Text)
                  
        xYBIAMVT0 = larrYBIAMVT0(wIndex)
        If meYBIAMVT0.MOUVEMDTR <> xYBIAMVT0.MOUVEMDTR _
        Or meYBIAMVT0.MOUVEMDCO <> xYBIAMVT0.MOUVEMDCO _
                                                                Then
          If arrDev_Nb <> 0 Then
                If meYBIAMVT0.MOUVEMDTR <> xYBIAMVT0.MOUVEMDTR Then blnDateComptableMultiple = True
                'prtSab_Compta_Total_Print 1
                prtSab_Compta_Total_Print 2
            End If
            
            blnDossier = False
            blnSéparateur = False
            meDateComptable = dateImp10(xYBIAMVT0.MOUVEMDCO + 19000000)
            meDateTraitement = dateImp10(xYBIAMVT0.MOUVEMDTR + 19000000)
            meDevise = xDevise
            prtTitleText = "Mouvements comptables" _
                        & " en date du " & meDateComptable _
                        & " - traitement du " & dateImp10(xYBIAMVT0.MOUVEMDTR + 19000000)

            
            If blnDétail Then
                If arrDev_Nb <> 0 Then
                    prtSAB_Compta_Form_Line
                    frmElpPrt.prtNewPage
                End If
                prtSAB_Compta_Form
            Else
               If arrDev_Nb = 0 Then prtSAB_Compta_Form
            End If
            Call lstErr_ChangeLastItem(lMe.lstErr, lMe.cmdContext, meDevise & xYBIAMVT0.MOUVEMOPE): DoEvents

        End If
        meYBIAMVT0 = xYBIAMVT0
        
        
         prtSab_Compta_Total_Add ' après lecture du compte : DEV
         
        If blnDétail Then prtSAB_Compta_Line
              

Next I
'prtSab_Compta_Total_Print 1
prtSab_Compta_Total_Print 2
prtSab_Compta_Total_Print 3

prtSAB_Compta_Form_Line
prtSAB_Compta_Close
Call lstErr_AddItem(lMe.lstErr, lMe.cmdContext, "Journal < "): DoEvents

End Sub

Public Sub prtSAB_Compta_Devise_xlsManual(larrYBIAMVT0() As typeYBIAMVT0, fgW As MSFlexGrid, fgW_arrIndex As Integer, lMe As Form, blnDétail As Boolean, ByRef wbExcel As Excel.Workbook)
Dim maxRows As Long
Dim maxRowsPlus As Long
Dim comptageRows As Long
Dim curSheet As Long
Dim curModele As Long
Dim wIndex As Long
Dim currentrow As Long
'On Error Resume Next:
Call lstErr_Clear(lMe.lstErr, lMe.cmdContext, "Journal Devise> "): DoEvents
Call lstErr_AddItem(lMe.lstErr, lMe.cmdContext, "Journal Devise...."): DoEvents

maxRows = 55
maxRowsPlus = 4
If fgW.Rows > 0 Then
    wbExcel.Sheets.Add
    curSheet = indice_nouvelle_feuille(wbExcel)
    'on recopie les 7 premières lignes de JOURNAL_D vers la nouvelle feuille
    curModele = indice_feuille_modele("AAJOURNAL_D", wbExcel)
    wbExcel.Sheets(curModele).Select
    Range("1:7").Select
    Selection.Copy
    wbExcel.Sheets(curSheet).Activate
    Range("A1").PasteSpecial Paste:=xlPasteColumnWidths
    ActiveSheet.Paste
    Range("A8").Select
    currentrow = 7
    comptageRows = currentrow
    fgW.Row = 1
    fgW.Col = fgW_arrIndex
    wIndex = Val(fgW.Text)
    fgW.Col = 5
    prtTitleText = "Journal comptable : " & fgW.Text _
               & " en date comptable du " & dateImp10(larrYBIAMVT0(wIndex).MOUVEMDCO + 19000000) _
               & " - traitement du " & dateImp10(larrYBIAMVT0(wIndex).MOUVEMDTR + 19000000)
               wbExcel.Sheets(curSheet).Cells(1, 3) = prtTitleText
               wbExcel.Sheets(curSheet).Cells(3, 6) = dateImp10(larrYBIAMVT0(wIndex).MOUVEMDCO + 19000000)
End If
Call prtSAB_Compta_Open_xlsManual(blnDétail)
blnDossierTrame = False
For I = 1 To fgW.Rows - 1
    fgW.Row = I
    fgW.Col = 1: xService = fgW.Text
    fgW.Col = 5: xDevise = fgW.Text
    fgW.Col = fgW_arrIndex
    wIndex = Val(fgW.Text)
    xYBIAMVT0 = larrYBIAMVT0(wIndex)
        If meYBIAMVT0.MOUVEMDTR <> xYBIAMVT0.MOUVEMDTR _
        Or meYBIAMVT0.MOUVEMDCO <> xYBIAMVT0.MOUVEMDCO _
        Or meDevise <> xDevise _
                                                                Then
            If arrDev_Nb <> 0 Then
                If meYBIAMVT0.MOUVEMDTR <> xYBIAMVT0.MOUVEMDTR Then blnDateComptableMultiple = True
                Call prtSab_Compta_Total_Print_xlsManual(1, wbExcel.Sheets(curSheet), currentrow, comptageRows, maxRows, maxRowsPlus)
                Call prtSAB_Compta_NewLine_xlsManual(currentrow, wbExcel.Sheets(curSheet), comptageRows, maxRows, maxRowsPlus)
                Range("6:6").Select
                Selection.Copy
                Range("A" & CStr(currentrow)).Select
                ActiveSheet.Paste
                Call prtSab_Compta_Total_Print_xlsManual(2, wbExcel.Sheets(curSheet), currentrow, comptageRows, maxRows, maxRowsPlus)
                Call prtSAB_Compta_NewLine_xlsManual(currentrow, wbExcel.Sheets(curSheet), comptageRows, maxRows, maxRowsPlus)
                Range("6:6").Select
                Selection.Copy
                Range("A" & CStr(currentrow)).Select
                ActiveSheet.Paste
            End If
            blnDossier = False
            blnSéparateur = False
            meDateComptable = dateImp10(xYBIAMVT0.MOUVEMDCO + 19000000)
            meDateTraitement = dateImp10(xYBIAMVT0.MOUVEMDTR + 19000000)
            If meDevise <> xDevise And meDevise <> "" Then
               wbExcel.Sheets(curSheet).Name = meDevise
            End If
            meDevise = xDevise
            If blnDétail Then
                 prtTitleText = "Journal comptable : " & xDevise _
                            & " en date comptable du " & meDateComptable _
                            & " - traitement du " & dateImp10(xYBIAMVT0.MOUVEMDTR + 19000000)
               If arrDev_Nb <> 0 Then
                    'suppression des lignes modèles
                    Rows("4:7").Select
                    Selection.Delete
                    currentrow = currentrow - 3
                    wbExcel.Sheets(curSheet).Cells(currentrow, 1) = "END_OF_SHEET"
                    wbExcel.Sheets.Add
                    curSheet = indice_nouvelle_feuille(wbExcel)
                    'on recopie les 7 premières lignes de JOURNAL_D vers la nouvelle feuille
                    curModele = indice_feuille_modele("AAJOURNAL_D", wbExcel)
                    wbExcel.Sheets(curModele).Select
                    Range("1:7").Select
                    Selection.Copy
                    wbExcel.Sheets(curSheet).Activate
                    Range("A1").PasteSpecial Paste:=xlPasteColumnWidths
                    ActiveSheet.Paste
                    Range("A8").Select
                    currentrow = 7
                    comptageRows = 7
                    wbExcel.Sheets(curSheet).Cells(1, 3) = prtTitleText
                    wbExcel.Sheets(curSheet).Cells(3, 6) = meDateComptable
                End If
           Else
                prtTitleText = "Journal comptable par devise" _
                            & " en date comptable du " & meDateComptable _
                            & " - traitement du " & dateImp10(xYBIAMVT0.MOUVEMDTR + 19000000)
           End If
            Call lstErr_ChangeLastItem(lMe.lstErr, lMe.cmdContext, meDevise & xYBIAMVT0.MOUVEMOPE)
        Else
            If meYBIAMVT0.MOUVEMOPE <> xYBIAMVT0.MOUVEMOPE Then
                Call prtSab_Compta_Total_Print_xlsManual(1, wbExcel.Sheets(curSheet), currentrow, comptageRows, maxRows, maxRowsPlus)
                blnDossier = False
              Call lstErr_ChangeLastItem(lMe.lstErr, lMe.cmdContext, meDevise & xYBIAMVT0.MOUVEMOPE)
          Else
                If meYBIAMVT0.MOUVEMNUM <> xYBIAMVT0.MOUVEMNUM _
                Or meYBIAMVT0.MOUVEMEVE <> xYBIAMVT0.MOUVEMEVE Then
                    blnDossier = False
                Else
                    blnDossier = True
                    blnSéparateur = True
                End If
            End If
        End If
        meYBIAMVT0 = xYBIAMVT0
         prtSab_Compta_Total_Add ' après lecture du compte : DEV
        If blnDétail Then
            Call prtSAB_Compta_Line_xlsManual(wbExcel.Sheets(curSheet), currentrow, comptageRows, maxRows, maxRowsPlus)
        End If
Next I
Call prtSab_Compta_Total_Print_xlsManual(1, wbExcel.Sheets(curSheet), currentrow, comptageRows, maxRows, maxRowsPlus)
Call prtSAB_Compta_NewLine_xlsManual(currentrow, wbExcel.Sheets(curSheet), comptageRows, maxRows, maxRowsPlus)
Range("6:6").Select
Selection.Copy
Range("A" & CStr(currentrow)).Select
ActiveSheet.Paste
Call prtSab_Compta_Total_Print_xlsManual(2, wbExcel.Sheets(curSheet), currentrow, comptageRows, maxRows, maxRowsPlus)
Call prtSAB_Compta_NewLine_xlsManual(currentrow, wbExcel.Sheets(curSheet), comptageRows, maxRows, maxRowsPlus)
Range("6:6").Select
Selection.Copy
Range("A" & CStr(currentrow)).Select
ActiveSheet.Paste
Call prtSab_Compta_Total_Print_xlsManual(3, wbExcel.Sheets(curSheet), currentrow, comptageRows, maxRows, maxRowsPlus)
Call prtSAB_Compta_NewLine_xlsManual(currentrow, wbExcel.Sheets(curSheet), comptageRows, maxRows, maxRowsPlus)
Range("6:6").Select
Selection.Copy
Range("A" & CStr(currentrow)).Select
ActiveSheet.Paste
wbExcel.Sheets(curSheet).Name = meDevise
'suppression des lignes modèles
Rows("4:7").Select
Selection.Delete
currentrow = currentrow - 3
wbExcel.Sheets(curSheet).Cells(currentrow, 1) = "END_OF_SHEET"

'RECAP devise
curSheet = indice_feuille_modele("ZZRECAPITULATIF", wbExcel)
wbExcel.Sheets(curSheet).Activate
prtTitleText = "Journal comptable : Récapitulatif par devises" _
           & " - traitement du " & dateImp10(xYBIAMVT0.MOUVEMDTR + 19000000)
wbExcel.Sheets(curSheet).Cells(1, 4) = prtTitleText
wbExcel.Sheets(curSheet).Cells(3, 6) = meDateComptable
currentrow = 6
comptageRows = 6
Call prtSab_Compta_Total_Print_xlsManual(4, wbExcel.Sheets(curSheet), currentrow, comptageRows, maxRows, maxRowsPlus)
'suppression des lignes modèles
Rows("4:5").Select
Selection.Delete
currentrow = currentrow - 2
wbExcel.Sheets(curSheet).Cells(currentrow, 1) = "END_OF_SHEET"

End Sub




Public Sub prtSAB_Compta_Line_xlsManual(ByRef wsexcel As Excel.Worksheet, ByRef currentrow As Long, ByRef comptageRows As Long, maxRows As Long, maxRowsPlus As Long)
Dim xStr As String

Call prtSAB_Compta_NewLine_xlsManual(currentrow, wsexcel, comptageRows, maxRows, maxRowsPlus)
If blnDossier Then
    Range("5:5").Select
Else
    Range("6:6").Select
End If
Selection.Copy
Range("A" & CStr(currentrow)).Select
ActiveSheet.Paste
wsexcel.Cells(currentrow, 2) = xYBIAMVT0.MOUVEMCOM
wsexcel.Cells(currentrow, 3) = "'" & dateImp10(xYBIAMVT0.MOUVEMDVA + 19000000)
wsexcel.Cells(currentrow, 4) = xYBIAMVT0.COMPTEDEV
Call prtSAB_Compta_Mt_xlsManual(xYBIAMVT0.MOUVEMMON, 5, 7, wsexcel, currentrow, 8)
wsexcel.Cells(currentrow, 8) = xYBIAMVT0.COMPTEINT
If Not blnDossier Then
    wsexcel.Cells(currentrow, 1) = xYBIAMVT0.MOUVEMOPE & "    " & Format$(xYBIAMVT0.MOUVEMNUM, "### ### ##0") & " " & xYBIAMVT0.MOUVEMEVE
    wsexcel.Cells(currentrow, 9) = xService & " " & xYBIAMVT0.MOUVEMSER & xYBIAMVT0.MOUVEMSSE
Else
    wsexcel.Cells(currentrow, 1) = ""
End If
wsexcel.Cells(currentrow, 10) = Format$(xYBIAMVT0.MOUVEMPIE, "### ### ##0") & "_" & xYBIAMVT0.MOUVEMECR
xStr = xYBIAMVT0.MOUVEMBDF
If xYBIAMVT0.MOUVEMANU <> "0" Then
    xStr = xStr & xYBIAMVT0.MOUVEMANU
End If
If xYBIAMVT0.MOUVEMEXO <> "N" Then
    xStr = xStr & xYBIAMVT0.MOUVEMEXO
End If
If xYBIAMVT0.MOUVEMRET <> " " Then
    xStr = xStr & xYBIAMVT0.MOUVEMRET
End If
wsexcel.Cells(currentrow, 11) = xStr
' Libellé 1 & 2
Call prtSAB_Compta_NewLine_xlsManual(currentrow, wsexcel, comptageRows, maxRows, maxRowsPlus)
Range("7:7").Select
Selection.Copy
Range("A" & CStr(currentrow)).Select
ActiveSheet.Paste
xStr = Trim(xYBIAMVT0.LIBELLIB1) & " " & Trim(xYBIAMVT0.LIBELLIB2) & " " & Trim(xYBIAMVT0.LIBELLIB3) & " " & Trim(xYBIAMVT0.LIBELLIB4)
wsexcel.Cells(currentrow, 8) = xStr

End Sub


Private Sub prtSAB_Compta_NewSheet_xlsManual(ByRef currentSheet As Long)

    currentSheet = currentSheet + 1
    
End Sub

Public Sub prtSAB_Compta_Open_xlsManual(blnDétail As Boolean)
Dim ii As Integer
'On Error GoTo prtError

prtPgmName = "prtSAB_Compta"
prtTitleUsr = usrName
prtTitleText = ""

rsYBIAMVT0_Init meYBIAMVT0
arrDev_Nb = 0
meService = "": meDevise = ""
blnDateComptableMultiple = False

blnTotal = Not blnDétail
If blnDétail Then
    mTotal_Trame1 = 240: mTotal_Trame2 = 220: mTotal_Trame3 = 200
Else
    mTotal_Trame1 = 255: mTotal_Trame2 = 235: mTotal_Trame3 = 215
End If

For ii = 1 To 100
    arrDb_N(ii) = 0: arrCr_N(ii) = 0
    arrDb_S(ii) = 0: arrCr_S(ii) = 0
    arrDb_T(ii) = 0: arrCr_T(ii) = 0
    arrDb_Recap(ii) = 0: arrCr_Recap(ii) = 0
Next ii

Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "iimpression")
frmElpPrt.Hide

End Sub

Public Sub prtSab_Compta_Total_Print_xlsManual(lK As Integer, ByRef wsexcel As Excel.Worksheet, ByRef currentrow As Long, ByRef comptageRows As Long, maxRows As Long, maxRowsPlus As Long)
Dim ii As Integer
Dim curDB As Currency, curCR As Currency
Dim intTrame As Integer
Dim blnLine As Boolean
Dim J As Long

blnLine = False
If lK = 3 And Not blnDateComptableMultiple Then Exit Sub
For ii = 1 To arrDev_Nb
    Select Case lK
        Case 1
            curDB = arrDb_N(ii): arrDb_N(ii) = 0
            curCR = arrCr_N(ii): arrCr_N(ii) = 0
        Case 2
            curDB = arrDb_S(ii): arrDb_S(ii) = 0
            curCR = arrCr_S(ii): arrCr_S(ii) = 0
       Case 3
            curDB = arrDb_T(ii): arrDb_T(ii) = 0
            curCR = arrCr_T(ii): arrCr_T(ii) = 0
        Case 4
            curDB = arrDb_Recap(ii): arrDb_Recap(ii) = 0
            curCR = arrCr_Recap(ii): arrCr_Recap(ii) = 0
   End Select
    
    If curDB <> 0 And curCR <> 0 Then
            Call prtSAB_Compta_NewLine_xlsManual(currentrow, wsexcel, comptageRows, maxRows, maxRowsPlus)
            wsexcel.Activate
            If InStr(wsexcel.Name, "Feuil") > 0 Or InStr(wsexcel.Name, "sheet") > 0 Then
                Range("6:6").Select
            Else
                Range("5:5").Select
            End If
            Selection.Copy
            wsexcel.Activate
            Range("A" & CStr(currentrow)).Select
            ActiveSheet.Paste
         If Not blnLine Then
            blnLine = True
        End If
         Call prtSAB_Compta_Mt_xlsManual(curDB, 5, 7, wsexcel, currentrow, 8)
         Call prtSAB_Compta_Mt_xlsManual(curCR, 5, 7, wsexcel, currentrow, 8)
         wsexcel.Cells(currentrow, 4) = arrDev(ii)
        If lK = 1 Then
            wsexcel.Cells(currentrow, 1) = meYBIAMVT0.MOUVEMOPE
        End If
        If lK = 3 Then
            wsexcel.Cells(currentrow, 2) = "Traitement du " & meDateTraitement
        Else
            wsexcel.Cells(currentrow, 2) = meService & "    Compta du"
            wsexcel.Cells(currentrow, 3) = meDateComptable
        End If
        If lK = 4 Then
            If curDB + curCR <> 0 Then
                wsexcel.Cells(currentrow, 8) = "********** ERREUR **************"
            End If
            Call prtSAB_Compta_NewLine_xlsManual(currentrow, wsexcel, comptageRows, maxRows, maxRowsPlus)
        End If
    End If
Next ii

End Sub

Public Sub prtSAB_Compta_Unit(larrYBIAMVT0() As typeYBIAMVT0, fgW As MSFlexGrid, fgW_arrIndex As Integer, lMe As Form, blnDétail As Boolean)
'On Error Resume Next:
Dim wIndex As Long

Call lstErr_Clear(lMe.lstErr, lMe.cmdContext, "Journal > " & Time): DoEvents
Call lstErr_AddItem(lMe.lstErr, lMe.cmdContext, "Journal : " & Time): DoEvents

prtSAB_Compta_Open blnDétail
blnDossierTrame = True

For I = 1 To fgW.Rows - 1
    
    fgW.Row = I
    fgW.Col = 1: xService = fgW.Text
    fgW.Col = fgW_arrIndex
    wIndex = Val(fgW.Text)
                  
        xYBIAMVT0 = larrYBIAMVT0(wIndex)
        If meYBIAMVT0.MOUVEMDTR <> xYBIAMVT0.MOUVEMDTR _
        Or meYBIAMVT0.MOUVEMDCO <> xYBIAMVT0.MOUVEMDCO _
        Or meService <> xService _
                                                                Then
            If arrDev_Nb <> 0 Then
                If meYBIAMVT0.MOUVEMDTR <> xYBIAMVT0.MOUVEMDTR Then blnDateComptableMultiple = True
                prtSab_Compta_Total_Print 1
                prtSab_Compta_Total_Print 2
            End If
            
            blnDossier = False
            blnSéparateur = False
            meDateComptable = dateImp10(xYBIAMVT0.MOUVEMDCO + 19000000)
            meDateTraitement = dateImp10(xYBIAMVT0.MOUVEMDTR + 19000000)
            meService = xService    'Table_Ope_Unit(xYBIAMVT0.MOUVEMSER & xYBIAMVT0.MOUVEMSSE & xYBIAMVT0.MOUVEMOPE)
            
            If blnDétail Then
                prtTitleText = "Journal comptable du service : " & meService _
                           & " en date comptable du " & meDateComptable _
                           & " - traitement du " & dateImp10(xYBIAMVT0.MOUVEMDTR + 19000000)
               If arrDev_Nb <> 0 Then
                    prtSAB_Compta_Form_Line
                    frmElpPrt.prtNewPage
                End If
                prtSAB_Compta_Form
           Else
                prtTitleText = "Journal comptable par service " _
                           & " en date comptable du " & meDateComptable _
                           & " - traitement du " & dateImp10(xYBIAMVT0.MOUVEMDTR + 19000000)
               If arrDev_Nb = 0 Then prtSAB_Compta_Form
            End If
                
            Call lstErr_ChangeLastItem(lMe.lstErr, lMe.cmdContext, meService & xYBIAMVT0.MOUVEMOPE & Time): DoEvents

        Else
            If meYBIAMVT0.MOUVEMOPE <> xYBIAMVT0.MOUVEMOPE Then
                prtSab_Compta_Total_Print 1
                blnDossier = False
                Call lstErr_ChangeLastItem(lMe.lstErr, lMe.cmdContext, meService & xYBIAMVT0.MOUVEMOPE & Time): DoEvents
            Else

                If meYBIAMVT0.MOUVEMNUM <> xYBIAMVT0.MOUVEMNUM _
                Or meYBIAMVT0.MOUVEMEVE <> xYBIAMVT0.MOUVEMEVE Then
                    blnDossier = False
                Else
                    blnDossier = True
                    blnSéparateur = True
                End If
            End If
        End If
        meYBIAMVT0 = xYBIAMVT0
        
        
         prtSab_Compta_Total_Add ' après lecture du compte : DEV
         
        If blnDétail Then prtSAB_Compta_Line
              
Next I
prtSab_Compta_Total_Print 1
prtSab_Compta_Total_Print 2
prtSab_Compta_Total_Print 3

prtSAB_Compta_Form_Line
prtSAB_Compta_Close
Call lstErr_AddItem(lMe.lstErr, lMe.cmdContext, "Journal < " & Time): DoEvents

End Sub

Public Sub prtSAB_Compta_Devise(larrYBIAMVT0() As typeYBIAMVT0, fgW As MSFlexGrid, fgW_arrIndex As Integer, lMe As Form, blnDétail As Boolean)
Dim wIndex As Long
'On Error Resume Next:
Call lstErr_Clear(lMe.lstErr, lMe.cmdContext, "Journal Devise> "): DoEvents
Call lstErr_AddItem(lMe.lstErr, lMe.cmdContext, "Journal Devise...."): DoEvents


prtSAB_Compta_Open blnDétail
blnDossierTrame = False

For I = 1 To fgW.Rows - 1
    
    fgW.Row = I
    fgW.Col = 1: xService = fgW.Text
    fgW.Col = 5: xDevise = fgW.Text
    fgW.Col = fgW_arrIndex
    wIndex = Val(fgW.Text)
                  
        xYBIAMVT0 = larrYBIAMVT0(wIndex)
        If meYBIAMVT0.MOUVEMDTR <> xYBIAMVT0.MOUVEMDTR _
        Or meYBIAMVT0.MOUVEMDCO <> xYBIAMVT0.MOUVEMDCO _
        Or meDevise <> xDevise _
                                                                Then
            If arrDev_Nb <> 0 Then
                If meYBIAMVT0.MOUVEMDTR <> xYBIAMVT0.MOUVEMDTR Then blnDateComptableMultiple = True
                prtSab_Compta_Total_Print 1
                prtSab_Compta_Total_Print 2
            End If
            
            blnDossier = False
            blnSéparateur = False
            meDateComptable = dateImp10(xYBIAMVT0.MOUVEMDCO + 19000000)
            meDateTraitement = dateImp10(xYBIAMVT0.MOUVEMDTR + 19000000)
            meDevise = xDevise
            
            If blnDétail Then
                 prtTitleText = "Journal comptable : " & meDevise _
                            & " en date comptable du " & meDateComptable _
                            & " - traitement du " & dateImp10(xYBIAMVT0.MOUVEMDTR + 19000000)
               If arrDev_Nb <> 0 Then
                    prtSAB_Compta_Form_Line
                    frmElpPrt.prtNewPage
                End If
                prtSAB_Compta_Form
           Else
                prtTitleText = "Journal comptable par devise" _
                            & " en date comptable du " & meDateComptable _
                            & " - traitement du " & dateImp10(xYBIAMVT0.MOUVEMDTR + 19000000)
                If arrDev_Nb = 0 Then prtSAB_Compta_Form
           End If
            
            
            Call lstErr_ChangeLastItem(lMe.lstErr, lMe.cmdContext, meDevise & xYBIAMVT0.MOUVEMOPE): DoEvents

        Else
            If meYBIAMVT0.MOUVEMOPE <> xYBIAMVT0.MOUVEMOPE Then
                prtSab_Compta_Total_Print 1
                blnDossier = False
              Call lstErr_ChangeLastItem(lMe.lstErr, lMe.cmdContext, meDevise & xYBIAMVT0.MOUVEMOPE): DoEvents
          Else

                If meYBIAMVT0.MOUVEMNUM <> xYBIAMVT0.MOUVEMNUM _
                Or meYBIAMVT0.MOUVEMEVE <> xYBIAMVT0.MOUVEMEVE Then
                    blnDossier = False
                Else
                    blnDossier = True
                    blnSéparateur = True
                End If
            End If
        End If
        meYBIAMVT0 = xYBIAMVT0
        
        
         prtSab_Compta_Total_Add ' après lecture du compte : DEV
         
        If blnDétail Then prtSAB_Compta_Line
              

Next I
prtSab_Compta_Total_Print 1
prtSab_Compta_Total_Print 2
prtSab_Compta_Total_Print 3

prtSAB_Compta_Form_Line

'RECAP devise

prtSAB_Compta_Form_Line
frmElpPrt.prtNewPage

prtTitleText = "Journal comptable : Récapitulatif par devises" _
           & " - traitement du " & dateImp10(xYBIAMVT0.MOUVEMDTR + 19000000)

prtSAB_Compta_Form
prtSab_Compta_Total_Print 4

prtSAB_Compta_Close
Call lstErr_AddItem(lMe.lstErr, lMe.cmdContext, "Journal < "): DoEvents

End Sub

Public Sub prtSAB_Compta_Compte(larrYBIAMVT0() As typeYBIAMVT0, fgW As MSFlexGrid, fgW_arrIndex As Integer, lMe As Form, blnDétail As Boolean)
Dim wIndex As Long
'On Error Resume Next:
Call lstErr_Clear(lMe.lstErr, lMe.cmdContext, "Journal Devise> "): DoEvents
Call lstErr_AddItem(lMe.lstErr, lMe.cmdContext, "Journal Devise...."): DoEvents

prtSAB_Compta_Open blnDétail
blnSéparateur = False
blnDossier = False
blnDossierTrame = False

For I = 1 To fgW.Rows - 1
    
    fgW.Row = I
    fgW.Col = 1: xService = fgW.Text
    fgW.Col = 5: xDevise = fgW.Text
    fgW.Col = fgW_arrIndex
    wIndex = Val(fgW.Text)
                  
        xYBIAMVT0 = larrYBIAMVT0(wIndex)
        If meYBIAMVT0.MOUVEMDTR <> xYBIAMVT0.MOUVEMDTR _
        Or meYBIAMVT0.MOUVEMDCO <> xYBIAMVT0.MOUVEMDCO _
                                                                Then
          If arrDev_Nb <> 0 Then
                If meYBIAMVT0.MOUVEMDTR <> xYBIAMVT0.MOUVEMDTR Then blnDateComptableMultiple = True
                'prtSab_Compta_Total_Print 1
                prtSab_Compta_Total_Print 2
            End If
            
            blnDossier = False
            blnSéparateur = False
            meDateComptable = dateImp10(xYBIAMVT0.MOUVEMDCO + 19000000)
            meDateTraitement = dateImp10(xYBIAMVT0.MOUVEMDTR + 19000000)
            meDevise = xDevise
            prtTitleText = "Mouvements comptables" _
                        & " en date du " & meDateComptable _
                        & " - traitement du " & dateImp10(xYBIAMVT0.MOUVEMDTR + 19000000)

            
            If blnDétail Then
                If arrDev_Nb <> 0 Then
                    prtSAB_Compta_Form_Line
                    frmElpPrt.prtNewPage
                End If
                prtSAB_Compta_Form
            Else
               If arrDev_Nb = 0 Then prtSAB_Compta_Form
            End If
            Call lstErr_ChangeLastItem(lMe.lstErr, lMe.cmdContext, meDevise & xYBIAMVT0.MOUVEMOPE): DoEvents

        End If
        meYBIAMVT0 = xYBIAMVT0
        
        
         prtSab_Compta_Total_Add ' après lecture du compte : DEV
         
        If blnDétail Then prtSAB_Compta_Line
              

Next I
'prtSab_Compta_Total_Print 1
prtSab_Compta_Total_Print 2
prtSab_Compta_Total_Print 3

prtSAB_Compta_Form_Line
prtSAB_Compta_Close
Call lstErr_AddItem(lMe.lstErr, lMe.cmdContext, "Journal < "): DoEvents

End Sub

'---------------------------------------------------------
Public Sub prtSAB_Compta_Form()
'---------------------------------------------------------
Dim X As String

frmElpPrt.prtStdTop

XPrt.DrawWidth = 1
XPrt.FontBold = True

Call frmElpPrt.prtTrame(prtMinX, prtMinY, prtMaxX, prtMinY + prtHeaderHeight, "B", 240)
'---------------------------------------------------------
XPrt.CurrentY = prtMinY + 50

XPrt.FontSize = 7

XPrt.CurrentX = prtMinX + 20: XPrt.Print "Dossier Evénement";
XPrt.CurrentX = prtMinX + 1450: XPrt.Print "Compte";
XPrt.CurrentX = prtMinX + 3400: XPrt.Print "D.valeur";
XPrt.CurrentX = prtMinX + 5750: XPrt.Print "Débit";
XPrt.CurrentX = prtMinX + 7700: XPrt.Print "Crédit";
XPrt.CurrentX = prtMinX + 8600: XPrt.Print "Intitulé du compte / Libellé du mouvement comptable";
XPrt.CurrentX = prtMinX + 14300: XPrt.Print "Réf pièce";
XPrt.CurrentX = prtMinX + 15200: XPrt.Print "Bdf AER";

XPrt.FontUnderline = True
XPrt.CurrentX = prtMinX + 6600: XPrt.Print meDateComptable;
XPrt.FontUnderline = False

XPrt.CurrentY = prtMinY + 50 + prtHeaderHeight
XPrt.FontBold = False
blnSéparateur = False
End Sub

Private Sub prtSAB_Compta_NewLine_xlsManual(ByRef currentrow As Long, ByRef wsexcel As Excel.Worksheet, ByRef comptageRows As Long, maxRows As Long, maxRowsPlus As Long)

    If currentrow >= maxRows + maxRowsPlus Then
        If comptageRows >= maxRows Then
            Call insere_entete_page(wsexcel, "1:4", 4, currentrow)
            comptageRows = 4
            currentrow = currentrow + 4
        End If
    End If
    comptageRows = comptageRows + 1
    currentrow = currentrow + 1
    
End Sub

Public Sub prtSAB_Compta_Close()
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



Public Sub prtSAB_Compta_Open(blnDétail As Boolean)
Dim I As Integer
On Error GoTo prtError

Set XPrt = Printer
frmElpPrt.Show vbModeless

blnFiligrane = False
prtOrientation = vbPRORLandscape '
prtPgmName = "prtSAB_Compta"
prtTitleUsr = usrName
prtTitleText = ""
prtFontName = prtFontName_Arial

prtLineNb = 1
prtlineHeight = 250
prtHeaderHeight = 50 ' 100

prtFormType = ""
frmElpPrt.prtStdInit
Height8_6 = frmElpPrt.prtHeightDelta(8, 6)

prtHeaderHeight = 300
 XPrt.CurrentY = XPrt.CurrentY - prtlineHeight

rsYBIAMVT0_Init meYBIAMVT0
arrDev_Nb = 0
colDb = 6500: colCr = 8500
meService = "": meDevise = ""
blnDateComptableMultiple = False

blnTotal = Not blnDétail
If blnDétail Then
    mTotal_Trame1 = 240: mTotal_Trame2 = 220: mTotal_Trame3 = 200
Else
    mTotal_Trame1 = 255: mTotal_Trame2 = 235: mTotal_Trame3 = 215
End If

For I = 1 To 100
    arrDb_N(I) = 0: arrCr_N(I) = 0
    arrDb_S(I) = 0: arrCr_S(I) = 0
    arrDb_T(I) = 0: arrCr_T(I) = 0
    arrDb_Recap(I) = 0: arrCr_Recap(I) = 0
Next I

Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub



Public Sub prtSab_Compta_Total_Add() ' à faire redim  / uBound
Dim I As Integer

For I = 1 To arrDev_Nb
    If arrDev(I) = xYBIAMVT0.COMPTEDEV Then
        If xYBIAMVT0.MOUVEMMON > 0 Then
            arrDb_N(I) = arrDb_N(I) + xYBIAMVT0.MOUVEMMON
            arrDb_S(I) = arrDb_S(I) + xYBIAMVT0.MOUVEMMON
            arrDb_T(I) = arrDb_T(I) + xYBIAMVT0.MOUVEMMON
            arrDb_Recap(I) = arrDb_Recap(I) + xYBIAMVT0.MOUVEMMON
        Else
            arrCr_N(I) = arrCr_N(I) + xYBIAMVT0.MOUVEMMON
            arrCr_S(I) = arrCr_S(I) + xYBIAMVT0.MOUVEMMON
            arrCr_T(I) = arrCr_T(I) + xYBIAMVT0.MOUVEMMON
            arrCr_Recap(I) = arrCr_Recap(I) + xYBIAMVT0.MOUVEMMON
        End If
        
        Exit Sub
    End If
Next I

arrDev_Nb = arrDev_Nb + 1
arrDev(arrDev_Nb) = xYBIAMVT0.COMPTEDEV
If xYBIAMVT0.MOUVEMMON > 0 Then
    arrDb_N(arrDev_Nb) = xYBIAMVT0.MOUVEMMON
    arrDb_S(arrDev_Nb) = xYBIAMVT0.MOUVEMMON
    arrDb_T(arrDev_Nb) = xYBIAMVT0.MOUVEMMON
    arrDb_Recap(arrDev_Nb) = xYBIAMVT0.MOUVEMMON
Else
    arrCr_N(arrDev_Nb) = xYBIAMVT0.MOUVEMMON
    arrCr_S(arrDev_Nb) = xYBIAMVT0.MOUVEMMON
    arrCr_T(arrDev_Nb) = xYBIAMVT0.MOUVEMMON
    arrCr_Recap(arrDev_Nb) = xYBIAMVT0.MOUVEMMON
End If
 
End Sub
Public Sub prtSab_Compta_Total_Print(lK As Integer) ' à faire redim  / uBound
Dim I As Integer
Dim curDB As Currency, curCR As Currency
Dim intTrame As Integer
Dim blnLine As Boolean
Dim wMinx As Integer, wMaxX As Integer
XPrt.FontSize = 8

'XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
blnLine = False

If lK = 3 And Not blnDateComptableMultiple Then Exit Sub

For I = 1 To arrDev_Nb
    Select Case lK
        Case 1
            curDB = arrDb_N(I): arrDb_N(I) = 0
            curCR = arrCr_N(I): arrCr_N(I) = 0
            intTrame = mTotal_Trame1
            wMinx = prtMinX: wMaxX = prtMaxX
        Case 2
            curDB = arrDb_S(I): arrDb_S(I) = 0
            curCR = arrCr_S(I): arrCr_S(I) = 0
            intTrame = mTotal_Trame2
             wMinx = prtMinX + 1350: wMaxX = prtMinX + 8500
       Case 3
            curDB = arrDb_T(I): arrDb_T(I) = 0
            curCR = arrCr_T(I): arrCr_T(I) = 0
            intTrame = mTotal_Trame3
            wMinx = prtMinX + 4600: wMaxX = prtMinX + 8500
        Case 4
            curDB = arrDb_Recap(I): arrDb_Recap(I) = 0
            curCR = arrCr_Recap(I): arrCr_Recap(I) = 0
            intTrame = mTotal_Trame3
            wMinx = prtMinX + 4600: wMaxX = prtMinX + 8500
   End Select
    
    If curDB <> 0 And curCR <> 0 Then
         XPrt.CurrentY = XPrt.CurrentY + prtlineHeight + 50
        If XPrt.CurrentY + 300 > prtMaxY Then
            prtSAB_Compta_Form_Line
            frmElpPrt.prtNewPage
            prtSAB_Compta_Form
        End If

         If Not blnLine Then
        '    XPrt.Line (wMinx, XPrt.CurrentY)-(wMaxX, XPrt.CurrentY), prtLineColor
            blnLine = True
        '    XPrt.CurrentY = XPrt.CurrentY + 10
        End If
         'Call frmElpPrt.prtTrame(prtMinX + 4600 + 10, XPrt.CurrentY, prtMinX + 8500 - 10, XPrt.CurrentY + prtlineHeight, " ", intTrame)
        Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY - 20, prtMaxX, XPrt.CurrentY + prtlineHeight - 20, " ", intTrame)
         XPrt.CurrentY = XPrt.CurrentY + 20
         XPrt.FontSize = 8
         Call prtSAB_Compta_Mt(curDB, colDb, colCr)
         Call prtSAB_Compta_Mt(curCR, colDb, colCr)
         XPrt.FontBold = True
         XPrt.FontSize = 6
         XPrt.CurrentX = prtMinX + 4200: XPrt.Print arrDev(I);
        If lK = 1 Then XPrt.CurrentX = prtMinX + 100: XPrt.Print meYBIAMVT0.MOUVEMOPE;
        If lK = 3 Then
            XPrt.FontBold = False
            XPrt.CurrentX = prtMinX + 2600: XPrt.Print "Traitement du";
            XPrt.CurrentX = prtMinX + 3500: XPrt.Print meDateTraitement;
        Else
            XPrt.CurrentX = prtMinX + 1400: XPrt.Print meService;
            XPrt.FontBold = False
            XPrt.CurrentX = prtMinX + 2800: XPrt.Print "Compta du";
            XPrt.CurrentX = prtMinX + 3500: XPrt.Print meDateComptable;
        End If
        If lK = 4 Then
            If curDB + curCR <> 0 Then
                XPrt.FontBold = True: XPrt.FontSize = 10
                XPrt.CurrentX = prtMinX + 8600: XPrt.Print "********** ERREUR **************";
            End If
        End If

    End If
Next I

If Not blnTotal And blnLine Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
End If

End Sub

Public Sub prtSAB_Compta_Form_Line()
XPrt.Line (prtMinX + 1350, prtMinY)-(prtMinX + 1350, prtMaxY), prtLineColor
XPrt.Line (prtMinX + 4600, prtMinY)-(prtMinX + 4600, prtMaxY), prtLineColor
XPrt.Line (prtMinX + 8500, prtMinY)-(prtMinX + 8500, prtMaxY), prtLineColor

End Sub

Public Sub prtSAB_Compta_Line()

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
If XPrt.CurrentY + 510 > prtMaxY Then
    prtSAB_Compta_Form_Line
    frmElpPrt.prtNewPage
    prtSAB_Compta_Form
End If
         
If Not blnDossier And blnDossierTrame Then Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY - 20, prtMaxX, XPrt.CurrentY + prtlineHeight - 20, " ", 240)

XPrt.FontSize = 8
Call prtSAB_Compta_Mt(xYBIAMVT0.MOUVEMMON, colDb, colCr)

XPrt.FontSize = 6: XPrt.CurrentY = XPrt.CurrentY + Height8_6
XPrt.FontBold = True
XPrt.CurrentX = prtMinX + 4200: XPrt.Print xYBIAMVT0.COMPTEDEV;
XPrt.FontBold = False
XPrt.CurrentX = prtMinX + 1450: XPrt.Print xYBIAMVT0.MOUVEMCOM;
 
XPrt.FontBold = True
XPrt.CurrentX = prtMinX + 8600: XPrt.Print xYBIAMVT0.COMPTEINT;

If Not blnDossier Then
    XPrt.CurrentX = prtMinX + 50: XPrt.Print xYBIAMVT0.MOUVEMOPE;
    X = Format$(xYBIAMVT0.MOUVEMNUM, "### ### ##0")
    XPrt.CurrentX = prtMinX + 950 - XPrt.TextWidth(X):         XPrt.Print X;
    XPrt.FontBold = False
    XPrt.CurrentX = prtMinX + 1000: XPrt.Print xYBIAMVT0.MOUVEMEVE;
    XPrt.CurrentX = prtMinX + 13500: XPrt.Print xService & " " & xYBIAMVT0.MOUVEMSER & xYBIAMVT0.MOUVEMSSE;
    
End If
XPrt.FontBold = False

 XPrt.CurrentX = prtMinX + 3400: XPrt.Print dateImp10(xYBIAMVT0.MOUVEMDVA + 19000000);
 
 X = Format$(xYBIAMVT0.MOUVEMPIE, "### ### ##0")
 XPrt.CurrentX = 15000 - XPrt.TextWidth(X)
 XPrt.Print X & "_" & xYBIAMVT0.MOUVEMECR;
 XPrt.CurrentX = prtMinX + 15400: XPrt.Print xYBIAMVT0.MOUVEMBDF;
 If xYBIAMVT0.MOUVEMANU <> "0" Then XPrt.CurrentX = prtMinX + 15700: XPrt.Print xYBIAMVT0.MOUVEMANU;
 If xYBIAMVT0.MOUVEMEXO <> "N" Then XPrt.CurrentX = prtMinX + 15800: XPrt.Print xYBIAMVT0.MOUVEMEXO;
 If xYBIAMVT0.MOUVEMRET <> " " Then XPrt.CurrentX = prtMinX + 15900: XPrt.Print xYBIAMVT0.MOUVEMRET;
          
 XPrt.CurrentY = XPrt.CurrentY - Height8_6

' Libellé 1 & 2
  XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.FontItalic = True
xLIBELLIB = Trim(xYBIAMVT0.LIBELLIB1) & " " & Trim(xYBIAMVT0.LIBELLIB2) & " " & Trim(xYBIAMVT0.LIBELLIB3) & " " & Trim(xYBIAMVT0.LIBELLIB4)
XPrt.CurrentX = prtMinX + 8600: XPrt.Print xLIBELLIB;

XPrt.FontItalic = False

End Sub

Public Sub prtSAB_Compta_Unit_xlsManual(larrYBIAMVT0() As typeYBIAMVT0, fgW As MSFlexGrid, fgW_arrIndex As Integer, lMe As Form, blnDétail As Boolean, wsexcel As Excel.Worksheet)
Dim maxRows As Long
Dim maxRowsPlus As Long
Dim comptageRows As Long
Dim currentrow As Long
Dim xStr As String
Dim laColonne As Long
'On Error Resume Next
Dim wIndex As Long
Dim nbSheetRows As Long

maxRows = 47
maxRowsPlus = 4
If fgW.Rows > 0 Then
    fgW.Row = 1
    fgW.Col = fgW_arrIndex
    wIndex = Val(fgW.Text)
    fgW.Col = 1
    prtTitleText = "Journal comptable du service : " & fgW.Text _
               & " en date comptable du " & meDateComptable _
               & " - traitement du " & dateImp10(larrYBIAMVT0(wIndex).MOUVEMDTR + 19000000)
    wsexcel.Cells(1, 3) = prtTitleText
    wsexcel.Cells(3, 6) = dateImp10(larrYBIAMVT0(wIndex).MOUVEMDTR + 19000000)
    currentrow = 7
    comptageRows = currentrow
End If
Call prtSAB_Compta_Open_xlsManual(blnDétail)
blnDossierTrame = True
For I = 1 To fgW.Rows - 1
    fgW.Row = I
    fgW.Col = 1: xService = fgW.Text
    fgW.Col = fgW_arrIndex
    wIndex = Val(fgW.Text)
        xYBIAMVT0 = larrYBIAMVT0(wIndex)
        If meYBIAMVT0.MOUVEMDTR <> xYBIAMVT0.MOUVEMDTR _
        Or meYBIAMVT0.MOUVEMDCO <> xYBIAMVT0.MOUVEMDCO _
        Or meService <> xService _
        Then
            If arrDev_Nb <> 0 Then
                If meYBIAMVT0.MOUVEMDTR <> xYBIAMVT0.MOUVEMDTR Then
                    blnDateComptableMultiple = True
                End If
                Call prtSab_Compta_Total_Print_xlsManual(1, wsexcel, currentrow, comptageRows, maxRows, maxRowsPlus)
                Call prtSAB_Compta_NewLine_xlsManual(currentrow, wsexcel, comptageRows, maxRows, maxRowsPlus)
                Range("7:7").Select
                Selection.Copy
                Range("A" & CStr(currentrow)).Select
                ActiveSheet.Paste
                Call prtSab_Compta_Total_Print_xlsManual(2, wsexcel, currentrow, comptageRows, maxRows, maxRowsPlus)
                Call prtSAB_Compta_NewLine_xlsManual(currentrow, wsexcel, comptageRows, maxRows, maxRowsPlus)
                Range("7:7").Select
                Selection.Copy
                Range("A" & CStr(currentrow)).Select
                ActiveSheet.Paste
            End If
            blnDossier = False
            blnSéparateur = False
            meDateComptable = dateImp10(xYBIAMVT0.MOUVEMDCO + 19000000)
            meDateTraitement = dateImp10(xYBIAMVT0.MOUVEMDTR + 19000000)
            meService = xService
            If blnDétail Then
                prtTitleText = "Journal comptable du service : " & xService _
                           & " en date comptable du " & meDateComptable _
                           & " - traitement du " & dateImp10(xYBIAMVT0.MOUVEMDTR + 19000000)
                wsexcel.Cells(1, 3) = prtTitleText
                wsexcel.Cells(3, 6) = meDateComptable
                currentrow = 7
                comptageRows = currentrow
           Else
                prtTitleText = "Journal comptable par service " _
                           & " en date comptable du " & meDateComptable _
                           & " - traitement du " & dateImp10(xYBIAMVT0.MOUVEMDTR + 19000000)
           End If
        Else
            If meYBIAMVT0.MOUVEMOPE <> xYBIAMVT0.MOUVEMOPE Then
                Call prtSab_Compta_Total_Print_xlsManual(1, wsexcel, currentrow, comptageRows, maxRows, maxRowsPlus)
                blnDossier = False
            Else
                If meYBIAMVT0.MOUVEMNUM <> xYBIAMVT0.MOUVEMNUM _
                Or meYBIAMVT0.MOUVEMEVE <> xYBIAMVT0.MOUVEMEVE Then
                    blnDossier = False
                Else
                    blnDossier = True
                    blnSéparateur = True
                End If
            End If
        End If
        meYBIAMVT0 = xYBIAMVT0
        prtSab_Compta_Total_Add
        If blnDétail Then
            Call prtSAB_Compta_Line_xlsManual(wsexcel, currentrow, comptageRows, maxRows, maxRowsPlus)
        End If
Next I
Call prtSab_Compta_Total_Print_xlsManual(1, wsexcel, currentrow, comptageRows, maxRows, maxRowsPlus)
Call prtSAB_Compta_NewLine_xlsManual(currentrow, wsexcel, comptageRows, maxRows, maxRowsPlus)
Range("7:7").Select
Selection.Copy
Range("A" & CStr(currentrow)).Select
ActiveSheet.Paste
Call prtSab_Compta_Total_Print_xlsManual(2, wsexcel, currentrow, comptageRows, maxRows, maxRowsPlus)
Call prtSAB_Compta_NewLine_xlsManual(currentrow, wsexcel, comptageRows, maxRows, maxRowsPlus)
Range("7:7").Select
Selection.Copy
Range("A" & CStr(currentrow)).Select
ActiveSheet.Paste
Call prtSab_Compta_Total_Print_xlsManual(3, wsexcel, currentrow, comptageRows, maxRows, maxRowsPlus)
Call prtSAB_Compta_NewLine_xlsManual(currentrow, wsexcel, comptageRows, maxRows, maxRowsPlus)
Range("7:7").Select
Selection.Copy
Range("A" & CStr(currentrow)).Select
ActiveSheet.Paste
wsexcel.Name = meService
'on supprime les 3 lignes modèles
Rows("5:7").Select
Selection.Delete
currentrow = currentrow - 2
wsexcel.Cells(currentrow + 1, 1) = "END_OF_SHEET"
nbSheetRows = retourne_fin_de_sheet(wsexcel)
Call frmSAB_Compta.zoneImpression_xlsManual(wsexcel.Name, nbSheetRows, wsexcel)
Call wsexcel.ExportAsFixedFormat(xlTypePDF, paramIMP_PDF_Path & "\" & paramEditionNoPaper_Auto_PgmName & ".pdf")
'sauvegarde du fichier
Call impressions_xlsManual.prtIMP_PDF_Monitor_xlsManual
'wsexcel.Delete

End Sub


