Attribute VB_Name = "rsYBIAMVT0"
'---------------------------------------------------------
Option Explicit
Public Const constYBIAMVT0 = "YBIAMVT0"

Dim appExcel As Excel.Application 'Application Excel
Dim wbExcel As Excel.Workbook 'Classeur Excel
Dim wsExcel As Excel.Worksheet 'Feuille Excel
Dim mXls2_Col As Long, mXls2_Row As Long

Type typeYBIAMVT0
    
    MOUVEMETA       As Integer                        ' ETABLISSEMENT
    MOUVEMPLA       As Long                           ' NUMERO PLAN
    MOUVEMCOM       As String * 20                    ' NUMERO COMPTE
    MOUVEMMON       As Currency                       ' MONTANT
    MOUVEMDOP       As Long                           ' DATE D'OPERATION
    MOUVEMDVA       As Long                           ' DATE DE VALEUR
    MOUVEMDCO       As Long                           ' DATE COMPTABLE
    MOUVEMDTR       As Long                           ' DATE DE TRAITEMENT
    MOUVEMPIE       As Long                           ' NUMERO DE PIECE
    MOUVEMECR       As Long                           ' NUMERO D'ECRITURE
    MOUVEMOPE       As String * 3                     ' CODE OPERATION
    MOUVEMNUM       As Long                           ' NUMERO OPERATION
    MOUVEMSCH       As Integer                        ' CODE SCHEMA
    MOUVEMUTI       As Integer                        ' UTILISATEUR
    MOUVEMAGE       As Integer                        ' AGENCE OPERATRICE
    MOUVEMSER       As String * 2                     ' SERVICE OPERATEUR
    MOUVEMSSE       As String * 2                     ' S/SERVICE OPERATEUR
    MOUVEMEXO       As String * 1                     ' CODE EXONERATION
    MOUVEMANA       As String * 6                     ' CODE ANALYTIQUE
    MOUVEMBDF       As String * 3                     ' CODE BANQUE DE FR.
    MOUVEMANU       As String * 1                     ' CODE ANNULATION
    MOUVEMRET       As String * 1                     ' MOUVEMENT RETRO
    MOUVEMEVE       As String * 3                     ' EVENEMENT
    MOUVEMSAN       As String * 6                     ' STRUCT ANALY-CODE
    MOUVEMSAD       As String * 80                    ' STRUCT ANALY-DONNEES
    
    LIBELLIB1       As String * 30                    ' Libellé 1
    LIBELLIB2       As String * 30                    ' Libellé 2
    LIBELLIB3       As String * 30                    ' Libellé 3
    LIBELLIB4       As String * 30                    ' Libellé 4
    
    COMPTEOBL       As String * 10                    ' COMPTE OBLIGATOIRE
    COMPTEINT       As String * 32                    ' INTITULE
    COMPTEDEV       As String * 3                     ' TABLES BASE 013
    COMPTELOR       As String * 1                     ' Lori/Nostri/AUTRE
    COMPTECLA       As Long                           ' CLASSE SECURITE
    
    BIAMVTSD0       As Currency                       ' solde
    BIAMVTID        As Long                           ' référence

End Type

Type typeYBIAMVT9
   MOUVEMCOM       As String * 20                    ' NUMERO COMPTE
   MOUVEMDTR       As Long                           ' DATE DE TRAITEMENT
   MOUVEMDOP       As Long                           ' DATE D'OPERATION
   MOUVEMPIE       As Long                           ' NUMERO DE PIECE
   MOUVEMECR       As Long                           ' NUMERO D'ECRITURE
   MOUVEMMON       As Currency                       ' MONTANT
   MOUVEMDVA       As Long                           ' DATE DE VALEUR
   MOUVEMSER       As String * 2                     ' SERVICE OPERATEUR
   MOUVEMSSE       As String * 2                     ' S/SERVICE OPERATEUR
   MOUVEMOPE       As String * 3                     ' CODE OPERATION
   MOUVEMNUM       As Long                           ' NUMERO OPERATION
   MOUVEMEVE       As String * 3                     ' EVENEMENT
   CLIENACLI       As String * 7
   CLIENARSD       As String * 3                     ' CDE PAYS DE RESIDENC
   COMPTEDEV       As String * 3
   COMPTEINT       As String * 32                    ' INTITULE
   LIBELLIB1       As String * 30                    ' Libellé 1
   LIBELLIB2       As String * 30                    ' Libellé 2
   LIBELLIB3       As String * 30                    ' Libellé 3
End Type
Public Function rsYBIAMVT9_GetBuffer(rsAdo As ADODB.Recordset, rsYBIAMVT9 As typeYBIAMVT9)
On Error GoTo Error_Handler
rsYBIAMVT9_GetBuffer = Null

    rsYBIAMVT9.MOUVEMCOM = rsAdo("MOUVEMCOM")
    rsYBIAMVT9.MOUVEMDTR = rsAdo("MOUVEMDTR")
    rsYBIAMVT9.MOUVEMDOP = rsAdo("MOUVEMDOP")
    rsYBIAMVT9.MOUVEMPIE = rsAdo("MOUVEMPIE")
    rsYBIAMVT9.MOUVEMECR = rsAdo("MOUVEMECR")
    rsYBIAMVT9.MOUVEMMON = rsAdo("MOUVEMMON")
    rsYBIAMVT9.MOUVEMDVA = rsAdo("MOUVEMDVA")
    rsYBIAMVT9.MOUVEMSER = rsAdo("MOUVEMSER")
    rsYBIAMVT9.MOUVEMSSE = rsAdo("MOUVEMSSE")
    rsYBIAMVT9.MOUVEMOPE = rsAdo("MOUVEMOPE")
    rsYBIAMVT9.MOUVEMNUM = rsAdo("MOUVEMNUM")
    rsYBIAMVT9.MOUVEMEVE = rsAdo("MOUVEMEVE")
    rsYBIAMVT9.CLIENACLI = rsAdo("CLIENACLI")
    rsYBIAMVT9.CLIENARSD = rsAdo("CLIENARSD")
    rsYBIAMVT9.COMPTEDEV = rsAdo("COMPTEDEV")
    rsYBIAMVT9.COMPTEINT = rsAdo("COMPTEINT")
    rsYBIAMVT9.LIBELLIB1 = rsAdo("LIBELLIB1")
    rsYBIAMVT9.LIBELLIB2 = rsAdo("LIBELLIB2")
    rsYBIAMVT9.LIBELLIB3 = rsAdo("LIBELLIB3")

Exit Function

Error_Handler:
rsYBIAMVT9_GetBuffer = Error

End Function

Public Sub YBIAMVTH_Exportation(mSQL_Extrait_Fct As String, lAMJMin As String, lAMJMax As String, lstErr As ListBox, lYBIAMVTH As typeYBIAMVT0, mSQL_Dossier_YBIAMVTHN As String, mSQL_Dossier_YDOSXODN As String, mSQL_Dossier_Pièce As String, mSQL_Extrait_YBIAMVTHN As String, mSQL_Extrait_Pièce As String)

On Error GoTo Error_Handler
Dim X As String, K As Long, K2 As Long, xWhere As String, wNum As Long
Dim wFile As String, wFilex As String
Dim blnCALCS As Boolean
Dim xLib As String
Dim mSheet_Nb As Integer
On Error GoTo Error_Handler
'===================================================================================
'If blnAuto Then
'    X = paramServer("\\CPT_Archive\")
'Else
    X = ""
'End If
If X = "" Then X = "C:\Temp\"
If Mid$(X, Len(X), 1) <> "\" Then X = X & "\"

blnCALCS = False
If Dir(X & "CALCS.jpl") <> "" Then blnCALCS = True
blnCALCS = True

If lYBIAMVTH.MOUVEMNUM <> 0 Then
    xLib = "Dossier " & lYBIAMVTH.MOUVEMSER & " " & lYBIAMVTH.MOUVEMSSE & " " & lYBIAMVTH.MOUVEMOPE & " " & lYBIAMVTH.MOUVEMNUM
Else
    xLib = "Mvts comptables "
End If

wFile = X & xLib & " " & DSys & "_" & time_Hms & ".xlsx"

'If Not blnAuto Then
    X = InputBox("par défaut : " _
        & vbCrLf & "     =========================" & vbCrLf & wFile _
        & vbCrLf & "     =========================", "Mvts comptables : nom du fichier d'exportation", wFile)
    If Trim(X) = "" Then Exit Sub
    wFilex = Trim(X)
    '______________________________________________
    If wFile <> wFilex Then
        wFile = wFilex
    End If
'End If

If Dir(wFile) <> "" Then msFileSystem.DeleteFile wFile
'_________________________________________


Set appExcel = CreateObject("Excel.Application")
appExcel.Workbooks.Add
Set wbExcel = appExcel.ActiveWorkbook
With wbExcel
    .Title = "YBIAMVTH"
    .Subject = ""
End With
If mSQL_Extrait_Pièce <> "" Then appExcel.Worksheets.Add

If mSQL_Dossier_YBIAMVTHN <> "" Then
    mSheet_Nb = mSheet_Nb + 1
    X = "Ecritures comptables du dossier : " & lYBIAMVTH.MOUVEMOPE & " " & lYBIAMVTH.MOUVEMNUM
    Call YBIAMVTH_Exportation_Page("", mSheet_Nb, lYBIAMVTH.MOUVEMOPE & " " & lYBIAMVTH.MOUVEMNUM, X)
    Call YBIAMVTH_Exportation_Detail(" ", mSQL_Dossier_YDOSXODN)
    Call YBIAMVTH_Exportation_Detail(" ", mSQL_Dossier_YBIAMVTHN)
End If

If mSQL_Dossier_Pièce <> "" Then
    wNum = 0
    K = InStr(1, mSQL_Dossier_Pièce, "MOUVEMPIE =")
    K2 = InStr(K + 10, mSQL_Dossier_Pièce, "and")
    If K2 > 0 Then wNum = Val(Mid$(mSQL_Dossier_Pièce, K + 11, K2 - K - 11))
    mSheet_Nb = mSheet_Nb + 1
    X = "Ecritures de la pièce comptable : " & wNum
    Call YBIAMVTH_Exportation_Page("", mSheet_Nb, "P_" & wNum, X)
    Call YBIAMVTH_Exportation_Detail(" ", mSQL_Dossier_Pièce)
End If

If mSQL_Extrait_YBIAMVTHN <> "" Then
    mSheet_Nb = mSheet_Nb + 1
    If mSQL_Extrait_Fct = "" Then
        X = "Extrait du compte : " & lYBIAMVTH.MOUVEMCOM
        Call YBIAMVTH_Exportation_Page("E", mSheet_Nb, lYBIAMVTH.MOUVEMCOM, X)
        Call YBIAMVTH_Exportation_Detail("E", mSQL_Extrait_YBIAMVTHN)
    Else
        X = "Extrait en date de valeur du compte  : " & lYBIAMVTH.MOUVEMCOM
        Call YBIAMVTH_Exportation_Page("E", mSheet_Nb, lYBIAMVTH.MOUVEMCOM, X)
        Call YBIAMVTH_Exportation_Detail_MOUVEMDVA(lYBIAMVTH.MOUVEMCOM, lAMJMin, lAMJMax, mSQL_Extrait_YBIAMVTHN)
    End If
    
End If

If mSQL_Extrait_Pièce <> "" Then
    wNum = 0
    K = InStr(1, mSQL_Extrait_Pièce, "MOUVEMPIE =")
    K2 = InStr(K + 10, mSQL_Extrait_Pièce, "and")
    If K2 > 0 Then wNum = Val(Mid$(mSQL_Extrait_Pièce, K + 11, K2 - K - 11))
    mSheet_Nb = mSheet_Nb + 1
    X = "Ecritures de la pièce comptable : " & wNum
    Call YBIAMVTH_Exportation_Page("", mSheet_Nb, "P_" & wNum, X)
    Call YBIAMVTH_Exportation_Detail(" ", mSQL_Extrait_Pièce)
End If


wbExcel.SaveAs wFile
wbExcel.Close
appExcel.Quit
'===================================================================================================
Exit_sub:
'__________________________________________________________________________________

Set rsSab = Nothing


Set wsExcel = Nothing
Set wbExcel = Nothing
Set appExcel = Nothing
Call lstErr_AddItem(lstErr, frmElp.cmdContext, "< Exportation terminée"): DoEvents
'_____________________________
Exit Sub

Error_Handler:

If Not blnCALCS Then
    X = "C:\Temp\"
    Resume Next
End If
MsgBox Error, vbCritical, "YBIAMVTH_Exportation"
Call lstErr_AddItem(lstErr, frmElp.cmdContext, "< Exportation terminée"): DoEvents

End Sub




Public Sub YBIAMVTH_Exportation_Page(lFct As String, lSheet As Integer, lName As String, lHeader As String)

On Error GoTo Error_Handler
Dim K As Integer

'==========================================================================================================

Set wsExcel = wbExcel.Sheets(lSheet)
wsExcel.Name = lSheet & "-" & lName

'__________________________________________________________________________________

With wsExcel.Cells
    .Borders(xlInsideHorizontal).Weight = xlThin ' xlMedium
    .Borders(xlInsideHorizontal).Color = RGB(160, 160, 160)
    .Borders(xlInsideVertical).Weight = xlThin
    .Borders(xlInsideVertical).Color = RGB(220, 220, 220)
    .VerticalAlignment = Excel.xlVAlignCenter
    .HorizontalAlignment = Excel.xlHAlignLeft
    .WrapText = True
    .Font.Size = 8
    .Font.Name = "Calibri"
    .Font.Color = RGB(0, 64, 128)
    .RowHeight = 17
End With

wsExcel.PageSetup.Orientation = vbPRORLandscape
wsExcel.PageSetup.HeaderMargin = 5
wsExcel.PageSetup.CenterHeader = vbCr & "&B&U&14 " & lHeader _
                                & vbCr & "&B&U&10(édité le " & dateImp10(DSys) & " " & Time & ")" & vbCr
wsExcel.PageSetup.LeftFooter = "&F - &B&A"
wsExcel.PageSetup.RightFooter = "&P / &N"
wsExcel.PageSetup.CenterHorizontally = True
wsExcel.PageSetup.PrintTitleRows = "$A1:$G1"

wsExcel.PageSetup.Zoom = 80

'Call lstErr_AddItem(lstErr, cmdContext, "Exportation en cours : "): DoEvents

mXls2_Col = 10
mXls2_Row = 1

wsExcel.Cells(1, 1) = "Date TRT": wsExcel.Columns(1).ColumnWidth = 10 ': wsExcel.Columns(1).NumberFormat = "mm/dd/yyyy"
wsExcel.Columns(1).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Cells(1, 2) = "Opération": wsExcel.Columns(2).ColumnWidth = 10
wsExcel.Cells(1, 3) = "numéro": wsExcel.Columns(3).ColumnWidth = 9: wsExcel.Columns(3).NumberFormat = "### ### ### ###"
wsExcel.Columns(3).HorizontalAlignment = Excel.xlHAlignRight
wsExcel.Cells(1, 4) = "O.D.": wsExcel.Columns(4).ColumnWidth = 8
wsExcel.Cells(1, 5) = "Date valeur": wsExcel.Columns(5).ColumnWidth = 10 ': wsExcel.Columns(5).NumberFormat = "mm/dd/yyyy"
wsExcel.Columns(5).HorizontalAlignment = Excel.xlHAlignCenter
wsExcel.Cells(1, 7) = "Devise": wsExcel.Columns(7).ColumnWidth = 5
wsExcel.Columns(7).HorizontalAlignment = Excel.xlHAlignCenter
If lFct = "E" Then
    wsExcel.Cells(1, 6) = "Montant dev": wsExcel.Columns(6).ColumnWidth = 15: wsExcel.Columns(6).NumberFormat = "### ### ### ##0.00;[Red]-### ### ### ##0.00"
    wsExcel.Columns(6).HorizontalAlignment = Excel.xlHAlignRight
    wsExcel.Cells(1, 8) = "Solde ": wsExcel.Columns(8).ColumnWidth = 15: wsExcel.Columns(8).NumberFormat = "### ### ### ##0.00;[Red]-### ### ### ##0.00"
    wsExcel.Columns(8).HorizontalAlignment = Excel.xlHAlignRight
Else
    wsExcel.Cells(1, 6) = "Montant dev": wsExcel.Columns(6).ColumnWidth = 15: wsExcel.Columns(6).NumberFormat = "### ### ### ##0.00;[Red]-### ### ### ##0.00"
    wsExcel.Columns(6).HorizontalAlignment = Excel.xlHAlignRight
    wsExcel.Cells(1, 8) = "Compte": wsExcel.Columns(8).ColumnWidth = 20
End If
wsExcel.Cells(1, 9) = "Libellé": wsExcel.Columns(9).ColumnWidth = 70
wsExcel.Cells(1, 10) = "Intitulé": wsExcel.Columns(10).ColumnWidth = 40

For K = 1 To mXls2_Col
    wsExcel.Cells(1, K).Interior.Color = mColor_GB
    wsExcel.Cells(1, K).Font.Color = mColor_Z0
Next

'==========================================================================================================
Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, "YBIAMVTH_Exportation_Detail"


End Sub

Public Sub YBIAMVTH_Exportation_Detail(lFct As String, lSQL As String)

On Error GoTo Error_Handler
Dim blnYDOSXOD0 As Boolean, K As Integer
Dim mMOUVEMDTR As Long, xMOUVEMDTR As Long, mSolde As Currency
Dim mXls2_Row_0 As Long, mXls2_Row_1 As Long
'==========================================================================================================
If lFct = "E" Then mXls2_Row = mXls2_Row + 1
'_________________________________________________________________________________________________
If InStr(1, lSQL, "YDOSXOD") > 0 Then
    blnYDOSXOD0 = True
Else
    blnYDOSXOD0 = False
End If

Set rsSab = cnsab.Execute(lSQL)

Do While Not rsSab.EOF

    If fctUser_Classe_Aut(rsSab("COMPTECLA")) Then
    
        mXls2_Row = mXls2_Row + 1
        xMOUVEMDTR = rsSab("MOUVEMDTR") + 19000000
            wsExcel.Cells(mXls2_Row, 1) = Format(xMOUVEMDTR, "0000/00/00")
            wsExcel.Cells(mXls2_Row, 2) = rsSab("MOUVEMSER") & " " & rsSab("MOUVEMSSE") & " " & rsSab("MOUVEMOPE") & " " & rsSab("MOUVEMEVE")
            wsExcel.Cells(mXls2_Row, 3) = rsSab("MOUVEMNUM")
           If blnYDOSXOD0 Then
                If Not IsNull(rsSab("DOSXODNUM")) Then
                    wsExcel.Cells(mXls2_Row, 2) = rsSab("MOUVEMSER") & " " & rsSab("MOUVEMSSE") & " " & rsSab("DOSXODOPE") & " " & rsSab("MOUVEMEVE")
                    wsExcel.Cells(mXls2_Row, 3) = rsSab("DOSXODNUM")
                    wsExcel.Cells(mXls2_Row, 4) = rsSab("MOUVEMOPE") & " " & rsSab("MOUVEMNUM")
                    wsExcel.Cells(mXls2_Row, 4).Font.Color = vbMagenta
                End If
            End If
            If lFct = "E" Then
                If mMOUVEMDTR <> xMOUVEMDTR Then
                    If mMOUVEMDTR = 0 Then
                        mSolde = rsSab("BIAMVTSD0")
                        'wsExcel.Cells(mXls2_Row, 1) = Format(dateElp("Jour", -1, xMOUVEMDTR), "0000/00/00")
                    Else
                        If mSolde <> rsSab("BIAMVTSD0") Then Call MsgBox("Erreur solde", vbCritical, "YBIAMVTH_Exportation_Detail")
                    End If
                    mXls2_Row_1 = mXls2_Row - 1
                    mMOUVEMDTR = xMOUVEMDTR
                    wsExcel.Cells(mXls2_Row_1, 8) = -rsSab("BIAMVTSD0")
                    wsExcel.Cells(mXls2_Row_1, 8).Font.Bold = True
                    wsExcel.Cells(mXls2_Row_1, 1).Font.Bold = True
                    If mSolde < 0 Then
                        wsExcel.Cells(mXls2_Row_1, 1).Interior.Color = mColor_B0
                        wsExcel.Cells(mXls2_Row_1, 8).Interior.Color = mColor_B0
                    Else
                        wsExcel.Cells(mXls2_Row_1, 1).Interior.Color = mColor_W0
                        wsExcel.Cells(mXls2_Row_1, 8).Interior.Color = mColor_W0
                    End If
                    
                End If
                wsExcel.Cells(mXls2_Row, 6) = -rsSab("MOUVEMMON")
                wsExcel.Cells(mXls2_Row, 8).Interior.Color = RGB(240, 240, 240)
            Else
                wsExcel.Cells(mXls2_Row, 8) = rsSab("MOUVEMCOM")
                wsExcel.Cells(mXls2_Row, 6) = -rsSab("MOUVEMMON")
            End If
            mSolde = mSolde + rsSab("MOUVEMMON")
            wsExcel.Cells(mXls2_Row, 7) = rsSab("COMPTEDEV")
            wsExcel.Cells(mXls2_Row, 5) = Format(rsSab("MOUVEMDVA") + 19000000, "0000/00/00")
            wsExcel.Cells(mXls2_Row, 10) = rsSab("COMPTEINT")
            wsExcel.Cells(mXls2_Row, 9) = Trim(rsSab("LIBELLIB1")) & " " & Trim(rsSab("LIBELLIB2")) & " " & Trim(rsSab("LIBELLIB3")) & " " & Trim(rsSab("LIBELLIB4"))
    End If
    rsSab.MoveNext
Loop

If lFct = "E" Then

    mXls2_Row = mXls2_Row + 1
    wsExcel.Cells(mXls2_Row, 8) = -mSolde
    wsExcel.Cells(mXls2_Row, 8).Font.Bold = True
    For K = 1 To mXls2_Col
        wsExcel.Cells(mXls2_Row, K).Interior.Color = RGB(230, 230, 230)
    Next
    If mSolde < 0 Then
        wsExcel.Cells(mXls2_Row, 8).Interior.Color = mColor_B1
    Else
        wsExcel.Cells(mXls2_Row, 8).Interior.Color = mColor_W1
    End If

End If

'==========================================================================================================
Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, "YBIAMVTH_Exportation_Detail"


End Sub
Public Sub YBIAMVTH_Exportation_Detail_MOUVEMDVA(lMOUVEMCOM As String, lAMJMin As String, lAMJMax As String, lSQL As String)

On Error GoTo Error_Handler
Dim blnYDOSXOD0 As Boolean, K As Integer
Dim mMOUVEMDVA As Long, xMOUVEMDVA As Long, mSolde As Currency
Dim wAmjMin7 As Long, wAmjMax7 As Long
Dim xSQL As String
Dim mXls2_Row_0 As Long, mXls2_Row_1 As Long
Dim dateMOUVEMDTR As Date, dateMOUVEMDVA As Date, blnSolde_Initial As Boolean, blnOk As Boolean, blnRupture As Boolean
'==========================================================================================================
'mXls2_Row = mXls2_Row + 1

wsExcel.Cells(1, 10) = "":: wsExcel.Columns(10).ColumnWidth = 1
mXls2_Col = 9


mXls2_Row_0 = mXls2_Row
blnSolde_Initial = False
    'If fctUser_Classe_Aut(rsSab("COMPTECLA")) Then


xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIAMVTH " _
     & " where MOUVEMCOM = '" & lMOUVEMCOM & "'" _
     & " order by MOUVEMDTR, MOUVEMPIE, MOUVEMECR "

Set rsSab = cnsab.Execute(xSQL)

If Not rsSab.EOF Then
    mSolde = rsSab("BIAMVTSD0")
    mMOUVEMDVA = rsSab("MOUVEMDVA")
Else
    Error = "Erreur lecture : " & xSQL
    GoTo Error_Handler
End If

wAmjMin7 = lAMJMin - 19000000
wAmjMax7 = lAMJMax - 19000000

'_________________________________________________________________________________________________
If InStr(1, lSQL, "YDOSXOD") > 0 Then
    blnYDOSXOD0 = True
Else
    blnYDOSXOD0 = False
End If

Set rsSab = cnsab.Execute(lSQL)

Do While Not rsSab.EOF

    If rsSab("MOUVEMDVA") > wAmjMax7 Then Exit Do
    
    If rsSab("MOUVEMDVA") < wAmjMin7 Then
         If rsSab("MOUVEMDTR") < wAmjMin7 Then
            mSolde = mSolde + rsSab("MOUVEMMON")
            mMOUVEMDVA = rsSab("MOUVEMDVA")
            blnOk = False
        Else
            blnOk = True: blnRupture = False
        End If
    Else
        blnOk = True: blnRupture = True
    End If
    
    If blnOk Then
        mXls2_Row = mXls2_Row + 1
        xMOUVEMDVA = rsSab("MOUVEMDVA")
                
        If blnRupture Then
            If mMOUVEMDVA <> xMOUVEMDVA Then
                
                If Not blnSolde_Initial Then
                    blnSolde_Initial = True
                        For K = 1 To mXls2_Col
                            wsExcel.Cells(mXls2_Row, K).Interior.Color = RGB(230, 230, 230)
                        Next

                    wsExcel.Cells(mXls2_Row, 5) = Format(mMOUVEMDVA + 19000000, "0000/00/00")
                    wsExcel.Cells(mXls2_Row, 8) = -mSolde
                    mXls2_Row_1 = mXls2_Row
                    mXls2_Row = mXls2_Row + 1

                Else
                    mXls2_Row_1 = mXls2_Row - 1
                    wsExcel.Cells(mXls2_Row_1, 8).FormulaLocal = "=SOMME(F" & mXls2_Row_0 + 1 & ":F" & mXls2_Row_1 & ")+H" & mXls2_Row_0
                End If
                'mMOUVEMDVA = xMOUVEMDVA
                wsExcel.Cells(mXls2_Row_1, 8).Font.Bold = True
                wsExcel.Cells(mXls2_Row_1, 5).Font.Bold = True
                If mSolde < 0 Then
                     wsExcel.Cells(mXls2_Row_1, 5).Interior.Color = mColor_B0
                     wsExcel.Cells(mXls2_Row_1, 8).Interior.Color = mColor_B0
                Else
                     wsExcel.Cells(mXls2_Row_1, 5).Interior.Color = mColor_W0
                     wsExcel.Cells(mXls2_Row_1, 8).Interior.Color = mColor_W0
                End If
                mXls2_Row_0 = mXls2_Row_1
            End If
        End If
        
        
        wsExcel.Cells(mXls2_Row, 8).Interior.Color = RGB(240, 240, 240)
        mSolde = mSolde + rsSab("MOUVEMMON")
        mMOUVEMDVA = xMOUVEMDVA
        
        wsExcel.Cells(mXls2_Row, 1) = Format(rsSab("MOUVEMDTR") + 19000000, "0000/00/00")
        wsExcel.Cells(mXls2_Row, 2) = rsSab("MOUVEMSER") & " " & rsSab("MOUVEMSSE") & " " & rsSab("MOUVEMOPE") & " " & rsSab("MOUVEMEVE")
        wsExcel.Cells(mXls2_Row, 3) = rsSab("MOUVEMNUM")
        If blnYDOSXOD0 Then
             If Not IsNull(rsSab("DOSXODNUM")) Then
                 wsExcel.Cells(mXls2_Row, 2) = rsSab("MOUVEMSER") & " " & rsSab("MOUVEMSSE") & " " & rsSab("DOSXODOPE") & " " & rsSab("MOUVEMEVE")
                 wsExcel.Cells(mXls2_Row, 3) = rsSab("DOSXODNUM")
                 wsExcel.Cells(mXls2_Row, 4) = rsSab("MOUVEMOPE") & " " & rsSab("MOUVEMNUM")
                 wsExcel.Cells(mXls2_Row, 4).Font.Color = vbMagenta
             End If
         End If
        wsExcel.Cells(mXls2_Row, 6) = -rsSab("MOUVEMMON")
        wsExcel.Cells(mXls2_Row, 7) = rsSab("COMPTEDEV")
        wsExcel.Cells(mXls2_Row, 5) = Format(rsSab("MOUVEMDVA") + 19000000, "0000/00/00")
        ''''wsExcel.Cells(mXls2_Row, 10) = rsSab("COMPTEINT")
        wsExcel.Cells(mXls2_Row, 9) = Trim(rsSab("LIBELLIB1")) & " " & Trim(rsSab("LIBELLIB2")) & " " & Trim(rsSab("LIBELLIB3")) & " " & Trim(rsSab("LIBELLIB4"))
    
        If Not blnRupture Then
            wsExcel.Cells(mXls2_Row, 5).Interior.Color = mColor_Y2
            wsExcel.Cells(mXls2_Row, 1).Interior.Color = mColor_Y2
        Else
            If rsSab("MOUVEMDTR") > wAmjMax7 Then
                wsExcel.Cells(mXls2_Row, 1).Interior.Color = mColor_Y2
            Else
                dateMOUVEMDTR = Date_VB(CLng(rsSab("MOUVEMDTR") + 19000000), 0)
                dateMOUVEMDVA = Date_VB(CLng(rsSab("MOUVEMDVA") + 19000000), 0)
                
                If Abs(DateDiff("d", dateMOUVEMDVA, dateMOUVEMDTR)) > 7 Then wsExcel.Cells(mXls2_Row, 1).Interior.Color = mColor_Y2
            End If
        End If

    End If
    rsSab.MoveNext
Loop


    mXls2_Row = mXls2_Row + 1
    'x = "=SOMME(F" & mXls2_Row_0 + 1 & ":F" & mXls2_Row - 1 & ")+H" & mXls2_Row_0
    wsExcel.Cells(mXls2_Row, 8).FormulaLocal = "=SOMME(F" & mXls2_Row_0 + 1 & ":F" & mXls2_Row - 1 & ")+H" & mXls2_Row_0
    wsExcel.Cells(mXls2_Row, 8).Font.Bold = True
    For K = 1 To mXls2_Col
        wsExcel.Cells(mXls2_Row, K).Interior.Color = RGB(230, 230, 230)
    Next
    If mSolde < 0 Then
        wsExcel.Cells(mXls2_Row, 8).Interior.Color = mColor_B1
    Else
        wsExcel.Cells(mXls2_Row, 8).Interior.Color = mColor_W1
    End If

If wsExcel.Cells(mXls2_Row, 8) <> -mSolde Then
    Call MsgBox("Erreur solde final : " & wsExcel.Cells(mXls2_Row, 8) & " # " & -mSolde, vbCritical, "Exportation Date de VAleur")
End If

'==========================================================================================================
Exit Sub

Error_Handler:
    MsgBox Error, vbCritical, "YBIAMVTH_Exportation_Detail"


End Sub




Public Sub srvYBIAMVT0_fgDisplay(recYBIAMVT0 As typeYBIAMVT0, fgDisplay As MSFlexGrid)
fgDisplay.Rows = 37
fgDisplay.Row = 1
fgDisplay.Col = 0: fgDisplay = "MOUVEMETA    5A"
fgDisplay.Col = 1: fgDisplay = "ETABLISSEMENT"
fgDisplay.Col = 2: fgDisplay = recYBIAMVT0.MOUVEMETA
fgDisplay.Row = 2
fgDisplay.Col = 0: fgDisplay = "MOUVEMPLA    4A"
fgDisplay.Col = 1: fgDisplay = "NUMERO PLAN"
fgDisplay.Col = 2: fgDisplay = recYBIAMVT0.MOUVEMPLA
fgDisplay.Row = 3
fgDisplay.Col = 0: fgDisplay = "MOUVEMCOM   20A"
fgDisplay.Col = 1: fgDisplay = "NUMERO COMPTE"
fgDisplay.Col = 2: fgDisplay = recYBIAMVT0.MOUVEMCOM
fgDisplay.Row = 4
fgDisplay.Col = 0: fgDisplay = "MOUVEMMON   18A"
fgDisplay.Col = 1: fgDisplay = "MONTANT"
fgDisplay.Col = 2: fgDisplay = recYBIAMVT0.MOUVEMMON
fgDisplay.Row = 5
fgDisplay.Col = 0: fgDisplay = "MOUVEMDOP    8A"
fgDisplay.Col = 1: fgDisplay = "DATE D OPERATION"
fgDisplay.Col = 2: fgDisplay = recYBIAMVT0.MOUVEMDOP
fgDisplay.Row = 6
fgDisplay.Col = 0: fgDisplay = "MOUVEMDVA    8A"
fgDisplay.Col = 1: fgDisplay = "DATE DE VALEUR"
fgDisplay.Col = 2: fgDisplay = recYBIAMVT0.MOUVEMDVA
fgDisplay.Row = 7
fgDisplay.Col = 0: fgDisplay = "MOUVEMDCO    8A"
fgDisplay.Col = 1: fgDisplay = "DATE COMPTABLE"
fgDisplay.Col = 2: fgDisplay = recYBIAMVT0.MOUVEMDCO
fgDisplay.Row = 8
fgDisplay.Col = 0: fgDisplay = "MOUVEMDTR    8A"
fgDisplay.Col = 1: fgDisplay = "DATE DE TRAITEMEN"
fgDisplay.Col = 2: fgDisplay = recYBIAMVT0.MOUVEMDTR
fgDisplay.Row = 9
fgDisplay.Col = 0: fgDisplay = "MOUVEMPIE   10A"
fgDisplay.Col = 1: fgDisplay = "NUMERO DE PIECE"
fgDisplay.Col = 2: fgDisplay = recYBIAMVT0.MOUVEMPIE
fgDisplay.Row = 10
fgDisplay.Col = 0: fgDisplay = "MOUVEMECR    8A"
fgDisplay.Col = 1: fgDisplay = "NUMERO D ECRITURE"
fgDisplay.Col = 2: fgDisplay = recYBIAMVT0.MOUVEMECR
fgDisplay.Row = 11
fgDisplay.Col = 0: fgDisplay = "MOUVEMOPE    3A"
fgDisplay.Col = 1: fgDisplay = "CODE OPERATION"
fgDisplay.Col = 2: fgDisplay = recYBIAMVT0.MOUVEMOPE
fgDisplay.Row = 12
fgDisplay.Col = 0: fgDisplay = "MOUVEMNUM   10A"
fgDisplay.Col = 1: fgDisplay = "NUMERO OPERATION"
fgDisplay.Col = 2: fgDisplay = recYBIAMVT0.MOUVEMNUM
fgDisplay.Row = 13
fgDisplay.Col = 0: fgDisplay = "MOUVEMSCH    5A"
fgDisplay.Col = 1: fgDisplay = "CODE SCHEMA"
fgDisplay.Col = 2: fgDisplay = recYBIAMVT0.MOUVEMSCH
fgDisplay.Row = 14
fgDisplay.Col = 0: fgDisplay = "MOUVEMUTI    5A"
fgDisplay.Col = 1: fgDisplay = "UTILISATEUR"
fgDisplay.Col = 2: fgDisplay = recYBIAMVT0.MOUVEMUTI
fgDisplay.Row = 15
fgDisplay.Col = 0: fgDisplay = "MOUVEMAGE    5A"
fgDisplay.Col = 1: fgDisplay = "AGENCE OPERATRICE"
fgDisplay.Col = 2: fgDisplay = recYBIAMVT0.MOUVEMAGE
fgDisplay.Row = 16
fgDisplay.Col = 0: fgDisplay = "MOUVEMSER    2A"
fgDisplay.Col = 1: fgDisplay = "SERVICE OPERATEUR"
fgDisplay.Col = 2: fgDisplay = recYBIAMVT0.MOUVEMSER
fgDisplay.Row = 17
fgDisplay.Col = 0: fgDisplay = "MOUVEMSSE    2A"
fgDisplay.Col = 1: fgDisplay = "S/SERVICE OPERATE"
fgDisplay.Col = 2: fgDisplay = recYBIAMVT0.MOUVEMSSE
fgDisplay.Row = 18
fgDisplay.Col = 0: fgDisplay = "MOUVEMEXO    1A"
fgDisplay.Col = 1: fgDisplay = "CODE EXONERATION"
fgDisplay.Col = 2: fgDisplay = recYBIAMVT0.MOUVEMEXO
fgDisplay.Row = 19
fgDisplay.Col = 0: fgDisplay = "MOUVEMANA    6A"
fgDisplay.Col = 1: fgDisplay = "CODE ANALYTIQUE"
fgDisplay.Col = 2: fgDisplay = recYBIAMVT0.MOUVEMANA
fgDisplay.Row = 20
fgDisplay.Col = 0: fgDisplay = "MOUVEMBDF    3A"
fgDisplay.Col = 1: fgDisplay = "CODE BANQUE DE FR"
fgDisplay.Col = 2: fgDisplay = recYBIAMVT0.MOUVEMBDF
fgDisplay.Row = 21
fgDisplay.Col = 0: fgDisplay = "MOUVEMANU    1A"
fgDisplay.Col = 1: fgDisplay = "CODE ANNULATION"
fgDisplay.Col = 2: fgDisplay = recYBIAMVT0.MOUVEMANU
fgDisplay.Row = 22
fgDisplay.Col = 0: fgDisplay = "MOUVEMRET    1A"
fgDisplay.Col = 1: fgDisplay = "MOUVEMENT RETRO"
fgDisplay.Col = 2: fgDisplay = recYBIAMVT0.MOUVEMRET
fgDisplay.Row = 23
fgDisplay.Col = 0: fgDisplay = "MOUVEMEVE    3A"
fgDisplay.Col = 1: fgDisplay = "EVENEMENT"
fgDisplay.Col = 2: fgDisplay = recYBIAMVT0.MOUVEMEVE
fgDisplay.Row = 24
fgDisplay.Col = 0: fgDisplay = "MOUVEMSAN    6A"
fgDisplay.Col = 1: fgDisplay = "STRUCT ANALY-CODE"
fgDisplay.Col = 2: fgDisplay = recYBIAMVT0.MOUVEMSAN
fgDisplay.Row = 25
fgDisplay.Col = 0: fgDisplay = "MOUVEMSAD   80A"
fgDisplay.Col = 1: fgDisplay = "STRUCT ANALY-DONN"
fgDisplay.Col = 2: fgDisplay = recYBIAMVT0.MOUVEMSAD
fgDisplay.Row = 26
fgDisplay.Col = 0: fgDisplay = "LIBELLIB1   30A"
fgDisplay.Col = 1: fgDisplay = "LIBELLE"
fgDisplay.Col = 2: fgDisplay = recYBIAMVT0.LIBELLIB1
fgDisplay.Row = 27
fgDisplay.Col = 0: fgDisplay = "LIBELLIB2   30A"
fgDisplay.Col = 1: fgDisplay = "LIBELLE"
fgDisplay.Col = 2: fgDisplay = recYBIAMVT0.LIBELLIB2
fgDisplay.Row = 28
fgDisplay.Col = 0: fgDisplay = "LIBELLIB3   30A"
fgDisplay.Col = 1: fgDisplay = "LIBELLE"
fgDisplay.Col = 2: fgDisplay = recYBIAMVT0.LIBELLIB3
fgDisplay.Row = 29
fgDisplay.Col = 0: fgDisplay = "LIBELLIB4   30A"
fgDisplay.Col = 1: fgDisplay = "LIBELLE"
fgDisplay.Col = 2: fgDisplay = recYBIAMVT0.LIBELLIB4
fgDisplay.Row = 30
fgDisplay.Col = 0: fgDisplay = "COMPTEOBL   10A"
fgDisplay.Col = 1: fgDisplay = "COMPTE OBLIGATOIR"
fgDisplay.Col = 2: fgDisplay = recYBIAMVT0.COMPTEOBL
fgDisplay.Row = 31
fgDisplay.Col = 0: fgDisplay = "COMPTEINT   32A"
fgDisplay.Col = 1: fgDisplay = "INTITULE"
fgDisplay.Col = 2: fgDisplay = recYBIAMVT0.COMPTEINT
fgDisplay.Row = 32
fgDisplay.Col = 0: fgDisplay = "COMPTEDEV    3A"
fgDisplay.Col = 1: fgDisplay = "DEVISE"
fgDisplay.Col = 2: fgDisplay = recYBIAMVT0.COMPTEDEV
fgDisplay.Row = 33
fgDisplay.Col = 0: fgDisplay = "COMPTELOR    1A"
fgDisplay.Col = 1: fgDisplay = "Lori/Nostri/AUTRE"
fgDisplay.Col = 2: fgDisplay = recYBIAMVT0.COMPTELOR
fgDisplay.Row = 34
fgDisplay.Col = 0: fgDisplay = "COMPTECLA    3A"
fgDisplay.Col = 1: fgDisplay = "CLASSE SECURITE"
fgDisplay.Col = 2: fgDisplay = recYBIAMVT0.COMPTECLA
fgDisplay.Row = 35
fgDisplay.Col = 0: fgDisplay = "BIAMVTSD0   19A"
fgDisplay.Col = 1: fgDisplay = "SOLDE PRECEDENT"
fgDisplay.Col = 2: fgDisplay = recYBIAMVT0.BIAMVTSD0
fgDisplay.Row = 36
fgDisplay.Col = 0: fgDisplay = "BIAMVTID   11A"
fgDisplay.Col = 1: fgDisplay = "IDENTIFICATION"
fgDisplay.Col = 2: fgDisplay = recYBIAMVT0.BIAMVTID
End Sub

'---------------------------------------------------------
Public Function rsYBIAMVT0_GetBuffer(rsAdo As ADODB.Recordset, rsYBIAMVT0 As typeYBIAMVT0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsYBIAMVT0_GetBuffer = Null

rsYBIAMVT0.MOUVEMETA = rsAdo("MOUVEMETA")
rsYBIAMVT0.MOUVEMPLA = rsAdo("MOUVEMPLA")
rsYBIAMVT0.MOUVEMCOM = rsAdo("MOUVEMCOM")
rsYBIAMVT0.MOUVEMMON = rsAdo("MOUVEMMON")
rsYBIAMVT0.MOUVEMDOP = rsAdo("MOUVEMDOP")
rsYBIAMVT0.MOUVEMDVA = rsAdo("MOUVEMDVA")
rsYBIAMVT0.MOUVEMDCO = rsAdo("MOUVEMDCO")
rsYBIAMVT0.MOUVEMDTR = rsAdo("MOUVEMDTR")
rsYBIAMVT0.MOUVEMPIE = rsAdo("MOUVEMPIE")
rsYBIAMVT0.MOUVEMECR = rsAdo("MOUVEMECR")
rsYBIAMVT0.MOUVEMOPE = rsAdo("MOUVEMOPE")
rsYBIAMVT0.MOUVEMNUM = rsAdo("MOUVEMNUM")
rsYBIAMVT0.MOUVEMSCH = rsAdo("MOUVEMSCH")
rsYBIAMVT0.MOUVEMUTI = rsAdo("MOUVEMUTI")
rsYBIAMVT0.MOUVEMAGE = rsAdo("MOUVEMAGE")
rsYBIAMVT0.MOUVEMSER = rsAdo("MOUVEMSER")
rsYBIAMVT0.MOUVEMSSE = rsAdo("MOUVEMSSE")
rsYBIAMVT0.MOUVEMEXO = rsAdo("MOUVEMEXO")
rsYBIAMVT0.MOUVEMANA = rsAdo("MOUVEMANA")
rsYBIAMVT0.MOUVEMBDF = rsAdo("MOUVEMBDF")
rsYBIAMVT0.MOUVEMANU = rsAdo("MOUVEMANU")
rsYBIAMVT0.MOUVEMRET = rsAdo("MOUVEMRET")
rsYBIAMVT0.MOUVEMEVE = rsAdo("MOUVEMEVE")
rsYBIAMVT0.MOUVEMSAN = rsAdo("MOUVEMSAN")
rsYBIAMVT0.MOUVEMSAD = rsAdo("MOUVEMSAD")

rsYBIAMVT0.LIBELLIB1 = rsAdo("LIBELLIB1")
rsYBIAMVT0.LIBELLIB2 = rsAdo("LIBELLIB2")
rsYBIAMVT0.LIBELLIB3 = rsAdo("LIBELLIB3")
rsYBIAMVT0.LIBELLIB4 = rsAdo("LIBELLIB4")
    
rsYBIAMVT0.COMPTEOBL = rsAdo("COMPTEOBL")
rsYBIAMVT0.COMPTEINT = rsAdo("COMPTEINT")
rsYBIAMVT0.COMPTEDEV = rsAdo("COMPTEDEV")
rsYBIAMVT0.COMPTELOR = rsAdo("COMPTELOR")
rsYBIAMVT0.COMPTECLA = rsAdo("COMPTECLA")
    
rsYBIAMVT0.BIAMVTSD0 = rsAdo("BIAMVTSD0")
rsYBIAMVT0.BIAMVTID = rsAdo("BIAMVTID")

Exit Function

Error_Handler:

rsYBIAMVT0_GetBuffer = Error

End Function

'---------------------------------------------------------
'---------------------------------------------------------
Public Sub rsYBIAMVT0_Init(rsYBIAMVT0 As typeYBIAMVT0)
'---------------------------------------------------------
On Error GoTo Error_Handler

rsYBIAMVT0.MOUVEMETA = 0
rsYBIAMVT0.MOUVEMPLA = 0
rsYBIAMVT0.MOUVEMCOM = ""
rsYBIAMVT0.MOUVEMMON = 0
rsYBIAMVT0.MOUVEMDOP = 0
rsYBIAMVT0.MOUVEMDVA = 0
rsYBIAMVT0.MOUVEMDCO = 0
rsYBIAMVT0.MOUVEMDTR = 0
rsYBIAMVT0.MOUVEMPIE = 0
rsYBIAMVT0.MOUVEMECR = 0
rsYBIAMVT0.MOUVEMOPE = ""
rsYBIAMVT0.MOUVEMNUM = 0
rsYBIAMVT0.MOUVEMSCH = 0
rsYBIAMVT0.MOUVEMUTI = 0
rsYBIAMVT0.MOUVEMAGE = 0
rsYBIAMVT0.MOUVEMSER = ""
rsYBIAMVT0.MOUVEMSSE = ""
rsYBIAMVT0.MOUVEMEXO = ""
rsYBIAMVT0.MOUVEMANA = ""
rsYBIAMVT0.MOUVEMBDF = ""
rsYBIAMVT0.MOUVEMANU = ""
rsYBIAMVT0.MOUVEMRET = ""
rsYBIAMVT0.MOUVEMEVE = ""
rsYBIAMVT0.MOUVEMSAN = ""
rsYBIAMVT0.MOUVEMSAD = ""
   
rsYBIAMVT0.LIBELLIB1 = ""
rsYBIAMVT0.LIBELLIB2 = ""
rsYBIAMVT0.LIBELLIB3 = ""
rsYBIAMVT0.LIBELLIB4 = ""
    
rsYBIAMVT0.COMPTEOBL = ""
rsYBIAMVT0.COMPTEINT = ""
rsYBIAMVT0.COMPTEDEV = ""
rsYBIAMVT0.COMPTELOR = ""
rsYBIAMVT0.COMPTECLA = ""
    
rsYBIAMVT0.BIAMVTSD0 = 0
rsYBIAMVT0.BIAMVTID = 0
   
Exit Sub

Error_Handler:


End Sub


Public Function sqlYBIAMVTHP(lMOUVEMETA As Long, lMOUVEMPIE As Long, lMOUVEMECR As Long, lYBIAMVT0 As typeYBIAMVT0)
Dim xSQL As String
Dim V

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIAMVTH " _
     & " where MOUVEMETA = " & lMOUVEMETA _
     & " and MOUVEMPIE = " & lMOUVEMPIE _
     & " and MOUVEMECR = " & lMOUVEMECR
     
Set rsSab = cnsab.Execute(xSQL)
If Not rsSab.EOF Then
    sqlYBIAMVTHP = rsYBIAMVT0_GetBuffer(rsSab, lYBIAMVT0)
Else
    Call rsYBIAMVT0_Init(lYBIAMVT0)
    sqlYBIAMVTHP = "? écriture inconnue : " & lMOUVEMPIE & " / " & lMOUVEMECR
End If
End Function
