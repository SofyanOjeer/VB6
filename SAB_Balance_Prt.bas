Attribute VB_Name = "prtSAB_Balance"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Type typeSAB_BALANCE
    Obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    CLIENARSD       As String * 3                     ' Pays de résidence
    COMPTEDEV       As String * 3                     ' TABLES BASE 013
    COMPTEOBL       As String * 10                    ' COMPTE OBLIGATOIRE
    COMPTECOM       As String * 20                    ' NUMERO COMPTE
    COMPTEINT       As String * 32                    ' INTITULE
    DB              As Currency                       '
    CR              As Currency                       '
    DB_EUR          As Currency                       '
    CR_EUR          As Currency
    blnPrint        As Boolean          '
    blnPrint_Fontbold        As Boolean          '
    blnPrint_Line        As Boolean          '
    iPrint_Trame        As Integer          '
End Type

Dim arrSAB_BALANCE() As typeSAB_BALANCE
Dim arrSAB_BALANCE_Nb As Long, arrSAB_BALANCE_NbMax As Long
Dim zSAB_BALANCE As typeSAB_BALANCE
Dim totalSAB_BALANCE(10) As typeSAB_BALANCE
Dim prevSAB_BALANCE As typeSAB_BALANCE
Dim détailSAB_BALANCE As typeSAB_BALANCE

Dim arrDEV_C1A5(100) As typeSAB_BALANCE
Dim arrDEV_C6A8(100) As typeSAB_BALANCE
Dim arrDEV_C9(100) As typeSAB_BALANCE
Dim arrDev_Nb As Long, arrDev_K As Integer


Dim X As String, I As Integer, Height8_6 As Integer

Dim blnPage As Boolean

Dim curX As Currency, curCumul_Db As Currency, curCumul_Cr As Currency
Dim curClient_Db As Currency, curClient_Cr As Currency, nbClient_Line As Long
Dim curListe_Db As Currency, curListe_Cr As Currency, nbListe_Line As Long
Dim curW_Db As Currency, curW_Cr As Currency
Dim IbmAmjMin As String, IbmAmjMax As String
Dim meYBIACPT0 As typeYBIACPT0, prevYBIACPT0 As typeYBIACPT0

Dim blnCompte As Boolean
Dim prtY As Integer
Dim mBalance_CV As String '* 3
Dim meCV1 As typeCV, meCV2 As typeCV

Dim blnMOUVEMDCO As Boolean, blnRésidence As Boolean
Dim mRésidence As String
Dim blnSoldeZ As Boolean
Dim blnClient_Line As Boolean

Dim blnBalance_B_COMPTEDEV As Boolean, blnBalance_B_Détail As Boolean, blnBalance_B_Récap As Boolean
Dim curCOMPTEDEV_Db As Currency, curCOMPTEDEV_Cr As Currency, nbCOMPTEDEV_Line As Long
Dim curCOMPTEDEV_Db_EUR As Currency, curCOMPTEDEV_Cr_EUR As Currency
Dim curCOMPTEOBL_Db As Currency, curCOMPTEOBL_Cr As Currency, nbCOMPTEOBL_Line As Long
Dim curCOMPTEOBL_Db_EUR As Currency, curCOMPTEOBL_Cr_EUR As Currency
Dim mBalance_ZSOLDE0 As Integer
Dim blnBalance_Compte_Soldé As Boolean
Dim blnBalance_Pays As Boolean, xPays As String
Dim blnBalance_Récap_Bilan As Boolean

Dim meZPLAN0 As typeZPLAN0

Dim idFile_CSV As Integer, blnFile_CSV As Boolean

Type typeSAB_Client_Stat
    Nb_Client           As Long
    Nb_Compte           As Long
    Nb_Compte_Annulé    As Long
    Solde_DB           As Currency
    Solde_CR           As Currency
End Type

Dim SAB_Client_Stat As typeSAB_Client_Stat
Dim SAB_Client_Stat_Actif As typeSAB_Client_Stat
Dim SAB_Client_Stat_Annulé As typeSAB_Client_Stat
Dim SAB_Client_Stat_Produit As typeSAB_Client_Stat
Dim SAB_Client_Stat_Produit_Lib As String

Dim meYBIATAB0 As typeYBIATAB0
Dim xTitleText As String
Dim mMsg As String
Dim blnChk_BalanceEquilibrée As Boolean, blnChk_BalanceStock As Boolean
Dim blnChk_COMPTEOUV As Boolean

Dim meYSTOMON As Currency, meDORCPTDMV As Long

Dim wAMJ_6M_00 As Long, wAMJ_6M_99 As Long
Dim wbExcel2 As Excel.Workbook 'Classeur Excel

Public blnPrint_Relevé_Total_Mvt As Boolean

Public Sub prtSAB_Balance_B_COMPTEDEV_xlsManual(ByRef currentRow As Long, wsExcel As Excel.Worksheet, ByRef comptageRows As Long, maxRows As Long, maxRowsPlus As Long)

Call prtSAB_Balance_B_COMPTEOBL_xlsManual(currentRow, wsExcel, comptageRows, maxRows, maxRowsPlus)
If blnBalance_Pays Then curCOMPTEDEV_Db = 0: curCOMPTEDEV_Cr = 0
If blnBalance_B_Détail And nbCOMPTEDEV_Line > 0 Then
    Call prtSAB_Balance_NewLine_xlsManual(currentRow, wsExcel, comptageRows, maxRows, maxRowsPlus)
    Range("A7:L7").Select
    Selection.Copy
    Range("A" & CStr(currentRow)).Select
    ActiveSheet.Paste
    If blnBalance_B_COMPTEDEV Then
        wsExcel.Cells(currentRow, 8) = prevYBIACPT0.COMPTEDEV
        If curCOMPTEDEV_Db <> 0 Then
            X = Format$(Abs(curCOMPTEDEV_Db), "### ### ### ### ##0.00")
            wsExcel.Cells(currentRow, 6) = X
        End If
        If curCOMPTEDEV_Cr <> 0 Then
             X = Format$(Abs(curCOMPTEDEV_Cr), "### ### ### ### ##0.00")
            wsExcel.Cells(currentRow, 7) = X
        End If
    End If
    If curCOMPTEDEV_Db_EUR <> 0 Then
        X = Format$(Abs(curCOMPTEDEV_Db_EUR), "### ### ### ### ##0.00")
            wsExcel.Cells(currentRow, 10) = X
    End If
    If curCOMPTEDEV_Cr_EUR <> 0 Then
        X = Format$(Abs(curCOMPTEDEV_Cr_EUR), "### ### ### ### ##0.00")
            wsExcel.Cells(currentRow, 12) = X
    End If
End If
curW_Db = curW_Db + curCOMPTEDEV_Db
curW_Cr = curW_Cr + curCOMPTEDEV_Cr
curCOMPTEDEV_Db = 0: curCOMPTEDEV_Cr = 0
curCOMPTEDEV_Db_EUR = 0: curCOMPTEDEV_Cr_EUR = 0
nbCOMPTEDEV_Line = 0
End Sub

Public Sub prtSAB_Balance_B_COMPTEOBL_xlsManual(ByRef currentRow As Long, wsExcel As Excel.Worksheet, ByRef comptageRows As Long, maxRows As Long, maxRowsPlus As Long)

Call rsZPLAN0_Read(prevYBIACPT0.COMPTEOBL, meZPLAN0)
If blnBalance_B_Détail And détailSAB_BALANCE.blnPrint And nbCOMPTEOBL_Line > 0 Then
    Call prtSAB_Balance_NewLine_xlsManual(currentRow, wsExcel, comptageRows, maxRows, maxRowsPlus)
    If détailSAB_BALANCE.iPrint_Trame <> 255 Then
        Range("A5:L5").Select
        Selection.Copy
        Range("A" & CStr(currentRow)).Select
        ActiveSheet.Paste
    Else
        Range("A7:L7").Select
        Selection.Copy
        Range("A" & CStr(currentRow)).Select
        ActiveSheet.Paste
    End If
    wsExcel.Cells(currentRow, 3) = meZPLAN0.PLANINTIT
    wsExcel.Cells(currentRow, 5) = prevYBIACPT0.COMPTEDEV & " " & prevYBIACPT0.COMPTEOBL
    If blnBalance_B_COMPTEDEV Then
        wsExcel.Cells(currentRow, 8) = prevYBIACPT0.COMPTEDEV
        If curCOMPTEOBL_Db <> 0 Then
            X = Format$(Abs(curCOMPTEOBL_Db), "### ### ### ### ##0.00")
            wsExcel.Cells(currentRow, 6) = X
        End If
        If curCOMPTEOBL_Cr <> 0 Then
             X = Format$(Abs(curCOMPTEOBL_Cr), "### ### ### ### ##0.00")
            wsExcel.Cells(currentRow, 7) = X
        End If
    End If
    If curCOMPTEOBL_Db_EUR <> 0 Then
        X = Format$(Abs(curCOMPTEOBL_Db_EUR), "### ### ### ### ##0.00")
            wsExcel.Cells(currentRow, 10) = X
    End If
    If curCOMPTEOBL_Cr_EUR <> 0 Then
        X = Format$(Abs(curCOMPTEOBL_Cr_EUR), "### ### ### ### ##0.00")
            wsExcel.Cells(currentRow, 12) = X
    End If
End If
curW_Db = curW_Db + curCOMPTEOBL_Db
curW_Cr = curW_Cr + curCOMPTEOBL_Cr
If arrSAB_BALANCE_Nb = arrSAB_BALANCE_NbMax Then
    arrSAB_BALANCE_NbMax = arrSAB_BALANCE_NbMax + 100
    ReDim Preserve arrSAB_BALANCE(arrSAB_BALANCE_NbMax)
End If
arrSAB_BALANCE_Nb = arrSAB_BALANCE_Nb + 1
arrSAB_BALANCE(arrSAB_BALANCE_Nb) = totalSAB_BALANCE(6)
arrSAB_BALANCE(arrSAB_BALANCE_Nb).CLIENARSD = prevYBIACPT0.CLIENARSD
arrSAB_BALANCE(arrSAB_BALANCE_Nb).COMPTEDEV = prevYBIACPT0.COMPTEDEV
arrSAB_BALANCE(arrSAB_BALANCE_Nb).COMPTEOBL = prevYBIACPT0.COMPTEOBL
arrSAB_BALANCE(arrSAB_BALANCE_Nb).COMPTECOM = ""
arrSAB_BALANCE(arrSAB_BALANCE_Nb).COMPTEINT = meZPLAN0.PLANINTIT
arrSAB_BALANCE(arrSAB_BALANCE_Nb).DB = arrSAB_BALANCE(arrSAB_BALANCE_Nb).DB + curCOMPTEOBL_Db
arrSAB_BALANCE(arrSAB_BALANCE_Nb).CR = arrSAB_BALANCE(arrSAB_BALANCE_Nb).CR + curCOMPTEOBL_Cr
arrSAB_BALANCE(arrSAB_BALANCE_Nb).DB_EUR = arrSAB_BALANCE(arrSAB_BALANCE_Nb).DB_EUR + curCOMPTEOBL_Db_EUR
arrSAB_BALANCE(arrSAB_BALANCE_Nb).CR_EUR = arrSAB_BALANCE(arrSAB_BALANCE_Nb).CR_EUR + curCOMPTEOBL_Cr_EUR
curCOMPTEOBL_Db = 0: curCOMPTEOBL_Cr = 0
curCOMPTEOBL_Db_EUR = 0: curCOMPTEOBL_Cr_EUR = 0
nbCOMPTEOBL_Line = 0
End Sub

Public Sub prtSAB_Balance_B_Dev_Classe_Prt_xlsManual(ByRef currentRow As Long, wsExcel As Excel.Worksheet, ByRef comptageRows As Long, maxRows As Long, maxRowsPlus As Long)
Dim curX As Currency, X As String
Dim curDB As Currency, curCR As Currency
Dim blnSaut As Boolean, blnErreur As Boolean
Dim wNb As Integer

blnErreur = False
wNb = arrDev_Nb + 2
arrDEV_C1A5(wNb) = zSAB_BALANCE
arrDEV_C6A8(wNb) = zSAB_BALANCE
arrDEV_C9(wNb) = zSAB_BALANCE

arrDEV_C1A5(arrDev_Nb + 1).COMPTEDEV = "???"
arrDEV_C6A8(arrDev_Nb + 1).COMPTEDEV = "???"
arrDEV_C9(arrDev_Nb + 1).COMPTEDEV = "???"
arrDEV_C1A5(arrDev_Nb + 2).COMPTEDEV = "***"
arrDEV_C6A8(arrDev_Nb + 2).COMPTEDEV = "***"
arrDEV_C9(arrDev_Nb + 2).COMPTEDEV = "***"

prtTitleText = "Récapitulatif Bilan / Hors Bilan " & xTitleText
wsExcel.Cells(1, 5) = prtTitleText
For arrDev_K = 1 To wNb
    blnSaut = False
    If arrDEV_C1A5(arrDev_K).DB_EUR <> 0 Or arrDEV_C1A5(arrDev_K).CR_EUR <> 0 Then
        blnSaut = True
        Call prtSAB_Balance_NewLine_xlsManual(currentRow, wsExcel, comptageRows, maxRows, maxRowsPlus)
        Range("A5:J5").Select
        Selection.Copy
        Range("A" & CStr(currentRow)).Select
        ActiveSheet.Paste
        wsExcel.Cells(currentRow, 4) = arrDEV_C1A5(arrDev_K).COMPTEDEV & " - Total classes 1 à 5 "
        wsExcel.Cells(currentRow, 7) = arrDEV_C1A5(arrDev_K).COMPTEDEV
        curX = Abs(arrDEV_C1A5(arrDev_K).DB)
        If curX <> 0 Then
            X = Format$(curX, "### ### ### ### ##0.00")
            wsExcel.Cells(currentRow, 5) = X
        End If
        
        curX = Abs(arrDEV_C1A5(arrDev_K).CR)
        If curX <> 0 Then
            X = Format$(curX, "### ### ### ### ##0.00")
            wsExcel.Cells(currentRow, 6) = X
        End If
        
        curX = Abs(arrDEV_C1A5(arrDev_K).DB_EUR)
        If curX <> 0 Then
            X = Format$(curX, "### ### ### ### ##0.00")
            wsExcel.Cells(currentRow, 8) = X
        End If
        
        curX = Abs(arrDEV_C1A5(arrDev_K).CR_EUR)
        If curX <> 0 Then
            X = Format$(curX, "### ### ### ### ##0.00")
            wsExcel.Cells(currentRow, 10) = X
        End If
    End If

    If arrDEV_C6A8(arrDev_K).DB_EUR <> 0 Or arrDEV_C6A8(arrDev_K).CR_EUR <> 0 Then
        Call prtSAB_Balance_NewLine_xlsManual(currentRow, wsExcel, comptageRows, maxRows, maxRowsPlus)
        Range("A5:J5").Select
        Selection.Copy
        Range("A" & CStr(currentRow)).Select
        ActiveSheet.Paste
        wsExcel.Cells(currentRow, 4) = arrDEV_C6A8(arrDev_K).COMPTEDEV & " - Total classes 6 à 8 "
        wsExcel.Cells(currentRow, 7) = arrDEV_C6A8(arrDev_K).COMPTEDEV
        curX = Abs(arrDEV_C6A8(arrDev_K).DB)
        If curX <> 0 Then
            X = Format$(curX, "### ### ### ### ##0.00")
            wsExcel.Cells(currentRow, 5) = X
        End If
        
        curX = Abs(arrDEV_C6A8(arrDev_K).CR)
        If curX <> 0 Then
            X = Format$(curX, "### ### ### ### ##0.00")
            wsExcel.Cells(currentRow, 6) = X
        End If
        
        curX = Abs(arrDEV_C6A8(arrDev_K).DB_EUR)
        If curX <> 0 Then
            X = Format$(curX, "### ### ### ### ##0.00")
            wsExcel.Cells(currentRow, 8) = X
        End If
        
        curX = Abs(arrDEV_C6A8(arrDev_K).CR_EUR)
        If curX <> 0 Then
            X = Format$(curX, "### ### ### ### ##0.00")
            wsExcel.Cells(currentRow, 10) = X
        End If
    End If
    If arrDEV_C9(arrDev_K).DB_EUR <> 0 Or arrDEV_C9(arrDev_K).CR_EUR <> 0 Then
        Call prtSAB_Balance_NewLine_xlsManual(currentRow, wsExcel, comptageRows, maxRows, maxRowsPlus)
        Range("A6:J6").Select
        Selection.Copy
        Range("A" & CStr(currentRow)).Select
        ActiveSheet.Paste
        wsExcel.Cells(currentRow, 4) = arrDEV_C9(arrDev_K).COMPTEDEV & " - Total classe 9 "
        wsExcel.Cells(currentRow, 7) = arrDEV_C9(arrDev_K).COMPTEDEV
        curX = Abs(arrDEV_C9(arrDev_K).DB)
        If curX <> 0 Then
            X = Format$(curX, "### ### ### ### ##0.00")
            wsExcel.Cells(currentRow, 5) = X
        End If
        curX = Abs(arrDEV_C9(arrDev_K).CR)
        If curX <> 0 Then
            X = Format$(curX, "### ### ### ### ##0.00")
            wsExcel.Cells(currentRow, 6) = X
        End If
        curX = Abs(arrDEV_C9(arrDev_K).DB_EUR)
        If curX <> 0 Then
            X = Format$(curX, "### ### ### ### ##0.00")
            wsExcel.Cells(currentRow, 8) = X
        End If
        curX = Abs(arrDEV_C9(arrDev_K).CR_EUR)
        If curX <> 0 Then
            X = Format$(curX, "### ### ### ### ##0.00")
            wsExcel.Cells(currentRow, 10) = X
        End If
    End If
    
    curX = arrDEV_C1A5(arrDev_K).DB + arrDEV_C1A5(arrDev_K).CR _
            + arrDEV_C6A8(arrDev_K).DB + arrDEV_C6A8(arrDev_K).CR
    If curX <> 0 And blnChk_BalanceEquilibrée Then
        blnErreur = True
        Call prtSAB_Balance_NewLine_xlsManual(currentRow, wsExcel, comptageRows, maxRows, maxRowsPlus)
        Range("A5:J5").Select
        Selection.Copy
        Range("A" & CStr(currentRow)).Select
        ActiveSheet.Paste
        X = Format$(curX, "### ### ### ### ##0.00")
        wsExcel.Cells(currentRow, 3) = X & " " & arrDEV_C1A5(arrDev_K).COMPTEDEV & "   ?????????? ERREUR BILAN "
        wsExcel.Cells(currentRow, 3).Font.Color = vbMagenta
        wsExcel.Cells(currentRow, 4) = ""
    End If
    curX = arrDEV_C9(arrDev_K).DB + arrDEV_C9(arrDev_K).CR
    If curX <> 0 And blnChk_BalanceEquilibrée Then
        blnErreur = True
        Call prtSAB_Balance_NewLine_xlsManual(currentRow, wsExcel, comptageRows, maxRows, maxRowsPlus)
        Range("A5:J5").Select
        Selection.Copy
        Range("A" & CStr(currentRow)).Select
        ActiveSheet.Paste
        X = Format$(curX, "### ### ### ### ##0.00")
        wsExcel.Cells(currentRow, 3) = X & " " & arrDEV_C9(arrDev_K).COMPTEDEV & "   ?????????? ERREUR HORS-BILAN "
        wsExcel.Cells(currentRow, 3).Font.Color = vbMagenta
        wsExcel.Cells(currentRow, 4) = ""
    End If
   If arrDev_K < wNb Then
        arrDEV_C1A5(wNb).DB_EUR = arrDEV_C1A5(wNb).DB_EUR + arrDEV_C1A5(arrDev_K).DB_EUR
        arrDEV_C1A5(wNb).CR_EUR = arrDEV_C1A5(wNb).CR_EUR + arrDEV_C1A5(arrDev_K).CR_EUR
        arrDEV_C6A8(wNb).DB_EUR = arrDEV_C6A8(wNb).DB_EUR + arrDEV_C6A8(arrDev_K).DB_EUR
        arrDEV_C6A8(wNb).CR_EUR = arrDEV_C6A8(wNb).CR_EUR + arrDEV_C6A8(arrDev_K).CR_EUR
        arrDEV_C9(wNb).DB_EUR = arrDEV_C9(wNb).DB_EUR + arrDEV_C9(arrDev_K).DB_EUR
        arrDEV_C9(wNb).CR_EUR = arrDEV_C9(wNb).CR_EUR + arrDEV_C9(arrDev_K).CR_EUR
   End If
   
   If blnSaut Then
        Call prtSAB_Balance_NewLine_xlsManual(currentRow, wsExcel, comptageRows, maxRows, maxRowsPlus)
        Range("A8:J8").Select
        Selection.Copy
        Range("A" & CStr(currentRow)).Select
        ActiveSheet.Paste
    End If

Next arrDev_K

blnSaut = True
Call prtSAB_Balance_NewLine_xlsManual(currentRow, wsExcel, comptageRows, maxRows, maxRowsPlus)
Range("A8:J8").Select
Selection.Copy
Range("A" & CStr(currentRow)).Select
ActiveSheet.Paste

Call prtSAB_Balance_NewLine_xlsManual(currentRow, wsExcel, comptageRows, maxRows, maxRowsPlus)
Range("A7:J7").Select
Selection.Copy
Range("A" & CStr(currentRow)).Select
ActiveSheet.Paste
If blnChk_BalanceEquilibrée Then
    If blnErreur Then
        X = "??? ERREUR ???"
        wsExcel.Cells(currentRow, 5) = X
        wsExcel.Cells(currentRow, 5).Font.Color = vbMagenta
    Else
        X = "Balance équilibrée"
        wsExcel.Cells(currentRow, 5) = X
    End If
End If
wsExcel.Cells(currentRow, 4) = arrDEV_C9(wNb).COMPTEDEV & " - Total classes 1-8 "
curDB = Abs(arrDEV_C1A5(wNb).DB_EUR + arrDEV_C6A8(wNb).DB_EUR)
If curDB <> 0 Then
            X = Format$(curDB, "### ### ### ### ##0.00")
            wsExcel.Cells(currentRow, 8) = X
End If

curCR = Abs(arrDEV_C1A5(wNb).CR_EUR + arrDEV_C6A8(wNb).CR_EUR)
If curCR <> 0 Then
            X = Format$(curCR, "### ### ### ### ##0.00")
            wsExcel.Cells(currentRow, 10) = X
End If
End Sub

Public Sub prtSAB_Balance_B_Fin_xlsManual(ByRef currentRow As Long, wsExcel As Excel.Worksheet, ByRef comptageRows As Long, maxRows As Long, maxRowsPlus As Long)
Dim I As Long, K As Long
Dim blnOk As Boolean

Call prtSAB_Balance_B_COMPTEOBL_xlsManual(currentRow, wsExcel, comptageRows, maxRows, maxRowsPlus)

If blnBalance_Pays Then                                  ' Tri Pays / PCi / Compte / Dev
    prtSAB_Balance_B_COMPTEOBL
Else
    Call prtSAB_Balance_B_COMPTEDEV_xlsManual(currentRow, wsExcel, comptageRows, maxRows, maxRowsPlus)
    Call prtSAB_Balance_B_COMPTEOBL_xlsManual(currentRow, wsExcel, comptageRows, maxRows, maxRowsPlus)
End If

If blnBalance_B_Récap Then
    prtTitleText = "Récapitulatif de la " & xTitleText
    If blnBalance_B_Détail Then
        wsExcel.Cells(1, 5) = prtTitleText
    End If
    
    prevSAB_BALANCE = arrSAB_BALANCE(1)
    
    For I = 1 To arrSAB_BALANCE_Nb
        blnOk = True
        
        If Not blnBalance_Pays Then
            If prevSAB_BALANCE.COMPTEDEV <> arrSAB_BALANCE(I).COMPTEDEV Then
            
                blnOk = False
                prtSAB_Balance_B_Total_Cumul 0
            End If
        End If
        
        If blnOk Then
            For K = 1 To 5
                If Mid$(prevSAB_BALANCE.COMPTEOBL, 1, K) <> Mid$(arrSAB_BALANCE(I).COMPTEOBL, 1, K) Then
                    prtSAB_Balance_B_Total_Cumul K
                    Exit For
                End If
            Next K
        End If
        
        If totalSAB_BALANCE(6).blnPrint Then
            If arrSAB_BALANCE(I).DB <> 0 Or arrSAB_BALANCE(I).CR <> 0 Then
                prtSAB_Balance_B_Total_Prt 6, arrSAB_BALANCE(I)
            End If
        End If
        
        prevSAB_BALANCE = arrSAB_BALANCE(I)
        totalSAB_BALANCE(5).DB = totalSAB_BALANCE(5).DB + prevSAB_BALANCE.DB
        totalSAB_BALANCE(5).CR = totalSAB_BALANCE(5).CR + prevSAB_BALANCE.CR
        totalSAB_BALANCE(5).DB_EUR = totalSAB_BALANCE(5).DB_EUR + prevSAB_BALANCE.DB_EUR
        totalSAB_BALANCE(5).CR_EUR = totalSAB_BALANCE(5).CR_EUR + prevSAB_BALANCE.CR_EUR
    
    Next I
    If blnBalance_Pays Then
        prevSAB_BALANCE.COMPTEDEV = "***"
        For K = 0 To 6
                totalSAB_BALANCE(K).COMPTEDEV = ""
                totalSAB_BALANCE(K).DB = 0: totalSAB_BALANCE(K).CR = 0
        Next K
    End If
    prtSAB_Balance_B_Total_Cumul 0

End If

If blnBalance_Récap_Bilan Then
    Call prtSAB_Balance_B_Dev_Classe_Prt_xlsManual(currentRow, wsExcel, comptageRows, maxRows, maxRowsPlus)
End If
End Sub

Public Sub prtSAB_Balance_B_Line_xlsManual(ByRef currentRow As Long, wsExcel As Excel.Worksheet, ByRef comptageRows As Long, maxRows As Long, maxRowsPlus As Long)
Dim curX As Currency, curX1 As Currency, curX2 As Currency
Dim X As String, XS As String, xText As String

Dim V
Dim blnOk As Boolean
Dim strX As String

If blnBalance_Pays Then                                  ' Tri Pays / PCi / Compte / Dev
    If prevYBIACPT0.CLIENARSD <> meYBIACPT0.CLIENARSD Then
        prtSAB_Balance_B_COMPTEOBL
        Call rsYBIATAB0_Read("SAB", "CLIENAPAY", "CLI" & meYBIACPT0.CLIENARSD, X)
        xPays = Trim(Mid$(X, 15, 30))
        prtTitleText = xPays & " - " & xTitleText
        If blnBalance_B_Détail Then
            frmElpPrt.prtNewPage
            prtSAB_Balance_Form
        End If
    End If
    If prevYBIACPT0.COMPTEOBL <> meYBIACPT0.COMPTEOBL Or prevYBIACPT0.COMPTEDEV <> meYBIACPT0.COMPTEDEV Then
        If Trim(prevYBIACPT0.COMPTEOBL) <> "" Then
            Call prtSAB_Balance_B_COMPTEOBL_xlsManual(currentRow, wsExcel, comptageRows, maxRows, maxRowsPlus)
        End If
    End If
Else
    If prevYBIACPT0.COMPTEDEV <> meYBIACPT0.COMPTEDEV Then
        If Trim(prevYBIACPT0.COMPTEDEV) <> "" Then
            Call prtSAB_Balance_B_COMPTEDEV_xlsManual(currentRow, wsExcel, comptageRows, maxRows, maxRowsPlus)
        End If
    End If
    If prevYBIACPT0.COMPTEOBL <> meYBIACPT0.COMPTEOBL Then
        If Trim(prevYBIACPT0.COMPTEOBL) <> "" Then
            Call prtSAB_Balance_B_COMPTEOBL_xlsManual(currentRow, wsExcel, comptageRows, maxRows, maxRowsPlus)
        End If
    End If
End If
    If prevYBIACPT0.COMPTEOBL <> meYBIACPT0.COMPTEOBL Then
        If Trim(prevYBIACPT0.COMPTEOBL) <> "" Then
            Call prtSAB_Balance_B_COMPTEOBL_xlsManual(currentRow, wsExcel, comptageRows, maxRows, maxRowsPlus)
        End If
    End If

blnOk = True
If mBalance_ZSOLDE0 = 0 Then
    curX = meYBIACPT0.SOLDECEN
Else
    V = rsZSOLDE0_Read(meYBIACPT0.COMPTECOM, mBalance_ZSOLDE0, curX)
    If Not IsNull(V) Then blnOk = False
End If

''If Not blnBalance_Compte_Soldé Or curX <> 0 Then
If blnBalance_Compte_Soldé And curX = 0 Then blnOk = False
If blnOk Then

    nbCOMPTEOBL_Line = nbCOMPTEOBL_Line + 1
    nbCOMPTEDEV_Line = nbCOMPTEDEV_Line + 1
    If blnBalance_B_Détail Then
    
        If blnChk_BalanceStock Then
            Call prtSAB_Balance_NewLine_xlsManual(currentRow, wsExcel, comptageRows, maxRows, maxRowsPlus)
            Range("A6:L6").Select
            Selection.Copy
            Range("A" & CStr(currentRow)).Select
            ActiveSheet.Paste
        End If
        Call prtSAB_Balance_L_xlsManual("B", curX, currentRow, wsExcel)
        If blnChk_BalanceStock Then
           If meDORCPTDMV > 0 Then
                strX = dateIBM10(meDORCPTDMV, True)
                If meDORCPTDMV <> meYBIACPT0.SOLDEDMO Then
                    strX = strX & " +"
                End If
                wsExcel.Cells(currentRow, 5) = strX
            Else
                wsExcel.Cells(currentRow, 5) = dateIBM10(meYBIACPT0.SOLDEDMO, True)
            End If
            If meYSTOMON <> -2 Then
                If meYSTOMON = -1 Then
                    X = "??"
                    wsExcel.Cells(currentRow, 9) = X
                    wsExcel.Cells(currentRow, 9).Font.Color = vbMagenta
                Else
                    curX1 = Abs(meYBIACPT0.SOLDECEN)
                    curX2 = Abs(meYSTOMON)
                    curX = Abs(curX1 - curX2)
                    If curX = 0 Then
                        X = "'="
                        wsExcel.Cells(currentRow, 9) = X
                    Else
                        X = Format$(Abs(curX2), "### ### ### ### ##0.00")
                        wsExcel.Cells(currentRow, 6) = X
                        wsExcel.Cells(currentRow, 6).Font.Color = vbMagenta
                        X = "##"
                        wsExcel.Cells(currentRow, 9) = X
                        wsExcel.Cells(currentRow, 9).Font.Color = vbMagenta
                    End If
                End If
            End If
        End If
        
        If meYBIACPT0.PLANCOPRO = "CAV" Or meYBIACPT0.PLANCOPRO = "LOR" Then
            If meYBIACPT0.COMPTEOUV > wAMJ_6M_00 And meYBIACPT0.COMPTEOUV < wAMJ_6M_99 Then
                wsExcel.Cells(currentRow, 9) = "6M"
                wsExcel.Cells(currentRow, 9).Font.Color = vbMagenta
            End If
        End If
        
        If blnBalance_Pays Then
            XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 6
            XPrt.CurrentX = prtMinX - 180: XPrt.Print meYBIACPT0.CLIENARSD;
            XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 8
        End If
    Else
        prtSAB_Balance_CV curX
    End If
    
    nbClient_Line = nbClient_Line + 1
    
    prtSAB_Balance_B_Dev_Classe_Cumul
    
    If meCV2.Montant > 0 Then
        curCOMPTEOBL_Db = curCOMPTEOBL_Db + meCV1.Montant
        curCOMPTEDEV_Db = curCOMPTEDEV_Db + meCV1.Montant
        curCOMPTEOBL_Db_EUR = curCOMPTEOBL_Db_EUR + meCV2.Montant
        curCOMPTEDEV_Db_EUR = curCOMPTEDEV_Db_EUR + meCV2.Montant
        curListe_Db = curListe_Db + meCV2.Montant
    Else
        curCOMPTEOBL_Cr = curCOMPTEOBL_Cr + meCV1.Montant
        curCOMPTEDEV_Cr = curCOMPTEDEV_Cr + meCV1.Montant
        curCOMPTEOBL_Cr_EUR = curCOMPTEOBL_Cr_EUR + meCV2.Montant
        curCOMPTEDEV_Cr_EUR = curCOMPTEDEV_Cr_EUR + meCV2.Montant
        curListe_Cr = curListe_Cr + meCV2.Montant
    End If
    If blnFile_CSV Then
        X = meYBIACPT0.COMPTEDEV & ";" & meYBIACPT0.COMPTEOBL & ";" & meYBIACPT0.COMPTECOM & ";" & meYBIACPT0.COMPTEINT & ";"
        If meCV2.Montant > 0 Then
            XS = cur_AbsV(meCV1.Montant) & "; ;" & cur_AbsV(meCV2.Montant) & "; "
        Else
            XS = " ;" & cur_AbsV(meCV1.Montant) & "; ;" & cur_AbsV(meCV2.Montant)
        End If
        If meYBIACPT0.COMPTEOUV > 0 Then
            xText = ";" & meYBIACPT0.COMPTEOUV + 19000000 & ";" & meYBIACPT0.COMPTEFON & ";"
        Else
            xText = ";" & ";" & meYBIACPT0.COMPTEFON & ";"
        End If
        If meYBIACPT0.SOLDEDMO > 0 Then xText = xText & meYBIACPT0.SOLDEDMO + 19000000
       
        Call File_Export_Monitor("Print", idFile_CSV, X & XS & xText)
    End If
End If
End Sub

Public Sub prtSAB_Balance_C_csvManual()
Dim xYBIAMVT0 As typeYBIAMVT0, wSolde As Currency
Dim curDB As Currency, curCR As Currency
Dim xDate As String
Dim xSQL As String

curDB = 0: curCR = 0
wSolde = meYBIACPT0.SOLDECEN
rsYBIAMVT0_Init xYBIAMVT0
If blnMOUVEMDCO Then
    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIAMVTH " _
         & " where MOUVEMCOM = '" & meYBIACPT0.COMPTECOM & "'" _
         & " and MOUVEMDCO >= " & IbmAmjMin _
         & " and MOUVEMDCO <= " & IbmAmjMax _
         & " order by MOUVEMDTR,MOUVEMPIE,MOUVEMECR"
Else
    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIAMVTH " _
         & " where MOUVEMCOM = '" & meYBIACPT0.COMPTECOM & "'" _
         & " and MOUVEMDTR >= " & IbmAmjMin _
         & " and MOUVEMDTR <= " & IbmAmjMax _
         & " order by MOUVEMDTR,MOUVEMPIE,MOUVEMECR"
End If
Set rsSab = cnsab.Execute(xSQL)
Do While Not rsSab.EOF
    Call rsYBIAMVT0_GetBuffer(rsSab, xYBIAMVT0)
    If xYBIAMVT0.MOUVEMMON < 0 Then
        curCR = curCR + xYBIAMVT0.MOUVEMMON
    Else
        curDB = curDB + xYBIAMVT0.MOUVEMMON
    End If
    rsSab.MoveNext
Loop
If blnSoldeZ Or curDB <> 0 Or curCR <> 0 Then
    blnCompte = True
    curCumul_Db = curCumul_Db + curDB
    curCumul_Cr = curCumul_Cr + curCR
    Call prtSAB_Balance_L_csvManual(curDB, curCR)
End If

End Sub

Public Sub prtSAB_Balance_C_xlsManual(ByRef currentRowStock As Long, wsexcelStock As Excel.Worksheet)
Dim xYBIAMVT0 As typeYBIAMVT0, wSolde As Currency
Dim curDB As Currency, curCR As Currency
Dim xDate As String
Dim xSQL As String

If blnRésidence Then
    If Mid$(meYBIACPT0.COMPTECOM, 10, 1) <> mRésidence Then
        Call prtSAB_Balance_C_xlsManual(currentRowStock, wsexcelStock)
        mRésidence = Mid$(meYBIACPT0.COMPTECOM, 10, 1)
        blnCompte = False:    curCumul_Db = 0: curCumul_Cr = 0
    End If
End If

curDB = 0: curCR = 0
wSolde = meYBIACPT0.SOLDECEN
rsYBIAMVT0_Init xYBIAMVT0
If blnMOUVEMDCO Then
    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIAMVTH " _
         & " where MOUVEMCOM = '" & meYBIACPT0.COMPTECOM & "'" _
         & " and MOUVEMDCO >= " & IbmAmjMin _
         & " and MOUVEMDCO <= " & IbmAmjMax _
         & " order by MOUVEMDTR,MOUVEMPIE,MOUVEMECR"
Else
    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIAMVTH " _
         & " where MOUVEMCOM = '" & meYBIACPT0.COMPTECOM & "'" _
         & " and MOUVEMDTR >= " & IbmAmjMin _
         & " and MOUVEMDTR <= " & IbmAmjMax _
         & " order by MOUVEMDTR,MOUVEMPIE,MOUVEMECR"
End If

     
Set rsSab = cnsab.Execute(xSQL)
Do While Not rsSab.EOF
    Call rsYBIAMVT0_GetBuffer(rsSab, xYBIAMVT0)

    If xYBIAMVT0.MOUVEMMON < 0 Then
        curCR = curCR + xYBIAMVT0.MOUVEMMON
    Else
        curDB = curDB + xYBIAMVT0.MOUVEMMON
    End If

    rsSab.MoveNext
Loop

        
If blnSoldeZ Or curDB <> 0 Or curCR <> 0 Then
'''If curDB = 0 And curCR = 0 Then
    blnCompte = True
    curCumul_Db = curCumul_Db + curDB
    curCumul_Cr = curCumul_Cr + curCR
    
    Call prtSAB_Balance_L_xlsManual("C1", curCR, currentRowStock, wsexcelStock)
    If curDB <> 0 Then Call prtSAB_Balance_Montant_xlsManual(curDB, currentRowStock, wsexcelStock)
End If
End Sub

Public Sub prtSAB_Balance_L_RELEVE_FOTC(lFct As String, lcurX As Currency, ByRef wsExcel As Excel.Worksheet, ByRef currentRow As Long)
Dim wForecolor As Long

currentRow = currentRow + 1

Select Case lFct
    Case "R0":
        Call prtSAB_Balance_Montant_RELEVE_FOTC(lcurX, wsExcel, currentRow)
    Case "R1"
        Call prtSAB_Balance_Montant_RELEVE_FOTC(lcurX, wsExcel, currentRow)
    Case "C1"
        Call prtSAB_Balance_Montant_RELEVE_FOTC(lcurX, wsExcel, currentRow)
    Case Else
        Call prtSAB_Balance_Montant_RELEVE_FOTC(lcurX, wsExcel, currentRow)
End Select
Select Case meYBIACPT0.COMPTEFON
    Case 4: wsExcel.Cells(currentRow, 4) = dateIBM10(meYBIACPT0.COMPTECLO, True)
            wsExcel.Cells(currentRow, 4).Font.Color = vbRed
End Select

wsExcel.Cells(currentRow, 1) = meYBIACPT0.PLANCOPRO
wsExcel.Cells(currentRow, 2) = meYBIACPT0.COMPTECOM
wsExcel.Cells(currentRow, 2).Font.Size = 8
wsExcel.Cells(currentRow, 2).Font.Name = "Arial"
wsExcel.Cells(currentRow, 3) = meYBIACPT0.COMPTEINT
wsExcel.Cells(currentRow, 3).Font.Size = 8
wsExcel.Cells(currentRow, 3).Font.Name = "Arial"
wsExcel.Cells(currentRow, 7) = meYBIACPT0.COMPTEDEV
wsExcel.Cells(currentRow, 7).Font.Size = 8
wsExcel.Cells(currentRow, 7).Font.Name = "Arial"
End Sub

Public Sub prtSAB_Balance_Monitor_csvManual(lFct As String, lAMJMin As String, lAMJMax As String, fgW As MSFlexGrid, larrYBIACPT0() As typeYBIACPT0, larrYBIACPT0_Nb As Long, lMsg As String, lYSTOMON() As Currency, lDORCPTDMV() As Long)
Dim nomModele As String
Dim nomCsv As String

Dim wIndex As Long, I As Integer
Dim mFct1 As String

nomCsv = ""
nomModele = ""
blnChk_BalanceEquilibrée = True
blnChk_BalanceStock = False
blnChk_COMPTEOUV = False

mFct1 = Mid$(lFct, 1, 1)
mMsg = lMsg
IbmAmjMin = dateIBM(lAMJMin)
IbmAmjMax = dateIBM(lAMJMax)
wAMJ_6M_00 = Fix((dateElp("MoisAdd", -6, YBIATAB0_DATE_CPT_J) - 19000000) / 100) * 100
wAMJ_6M_99 = wAMJ_6M_00 + 99
meCV1.DeviseN = 0
meCV1.Montant = 0
mBalance_CV = "J  "
meCV1.OpéAmj = YBIATAB0_DATE_CPT_J
meCV2.OpéAmj = YBIATAB0_DATE_CPT_J
If lAMJMax = YBIATAB0_DATE_CPT_MP1 Then
        mBalance_CV = "MP1"
        meCV1.OpéAmj = YBIATAB0_DATE_CPT_MP1
        meCV2.OpéAmj = YBIATAB0_DATE_CPT_MP1
End If
If lAMJMax = YBIATAB0_DATE_CPT_AP1 Then
        mBalance_CV = "AP1"
        meCV1.OpéAmj = YBIATAB0_DATE_CPT_AP1
        meCV2.OpéAmj = YBIATAB0_DATE_CPT_AP1
End If
blnMOUVEMDCO = False
blnRésidence = False: mRésidence = "-"
blnCompte = False:    curCumul_Db = 0: curCumul_Cr = 0
blnClient_Line = False
nbClient_Line = 0: nbListe_Line = 0
curClient_Db = 0: curClient_Cr = 0
curListe_Db = 0: curListe_Cr = 0
curCOMPTEDEV_Db = 0: curCOMPTEDEV_Cr = 0
curCOMPTEDEV_Db_EUR = 0: curCOMPTEDEV_Cr_EUR = 0
curCOMPTEOBL_Db = 0: curCOMPTEOBL_Cr = 0
curCOMPTEOBL_Db_EUR = 0: curCOMPTEOBL_Cr_EUR = 0
curW_Db = 0: curW_Cr = 0
rsYBIACPT0_Init prevYBIACPT0
prtSAB_Balance_Total_Init
blnBalance_B_Détail = False
blnBalance_B_Récap = False
blnBalance_B_COMPTEDEV = False
blnBalance_Compte_Soldé = False
blnBalance_Pays = False
blnBalance_Récap_Bilan = False
nbCOMPTEOBL_Line = 0: nbCOMPTEDEV_Line = 0
mBalance_ZSOLDE0 = 0
idFile_CSV = 0
blnFile_CSV = False
Select Case mFct1
    Case "C": prtTitleText = "Cumul de mouvements du " & dateImp10(lAMJMin) & " au " & dateImp10(lAMJMax)
               If Mid$(lFct, 2, 3) = "DCO" Then blnMOUVEMDCO = True
               If Mid$(lFct, 5, 1) = "R" Then blnRésidence = True
               nomCsv = Format(Now, "yyyymmdd_hhnnss") & "_BALANCE_Mvts.csv"
               nomModele = paramFolder_Local & "\Modeles\modele_BALANCE_Mvts.csv"
End Select
If Mid$(lFct, 6, 1) = "Z" Then
    blnSoldeZ = True
Else
    blnSoldeZ = False
End If
'On recopie le fichier csv modèle de c:\BIASRV vers c:\temp\imp_pdf
FileCopy nomModele, paramTemp_Folder & "\" & nomCsv
csvFic = FreeFile
Open paramTemp_Folder & "\" & nomCsv For Append As #csvFic
'                       '
For I = 1 To fgW.Rows - 1
    fgW.Row = I
    fgW.Col = fgW.Cols - 1: wIndex = Val(fgW.Text)
    meYBIACPT0 = larrYBIACPT0(wIndex)
    meCV1.DeviseIso = meYBIACPT0.COMPTEDEV
    If blnChk_BalanceStock Then meYSTOMON = lYSTOMON(wIndex): meDORCPTDMV = lDORCPTDMV(wIndex)
    Call prtSAB_Balance_C_csvManual
     prevYBIACPT0 = meYBIACPT0
Next I
Close #csvFic
Call MsgBox("Fin de l'impression de la balance." & vbCrLf & paramTemp_Folder & "\" & nomCsv)

End Sub




Public Sub prtSAB_Balance_Monitor_RELEVE_FOTC(lAMJMin As String, lAMJMax As String, fgW As MSFlexGrid, larrYBIACPT0() As typeYBIACPT0, larrYBIACPT0_Nb As Long, lMsg As String, lYSTOMON() As Currency, lDORCPTDMV() As Long)
Dim nomModele As String
Dim nomExcel As String
Dim currentRow As Long
Dim wbExcel As Excel.Workbook
Dim wsExcel As Excel.Worksheet
Dim wIndex As Long
Dim I As Integer

    nomModele = paramFolder_Local & "\Modeles\modele_RELEVE_FOTC.xlsx"
    nomExcel = "FOTC.BIA-BAL-RELEVE-FOTC_(S54).xlsx"
    Set appExcelPublic = CreateObject("Excel.Application")
    appExcelPublic.Visible = False
    appExcelPublic.ControlCharacters = False
    appExcelPublic.Interactive = False
    'On recopie le classeur modèle de c:\BIASRV vers c:\temp\imp_pdf
    FileCopy nomModele, paramIMP_PDF_Path_Temp & "\" & nomExcel
    Call appExcelPublic.Workbooks.Open(paramIMP_PDF_Path_Temp & "\" & nomExcel)
    Set wbExcel = appExcelPublic.ActiveWorkbook
    Set wsExcel = wbExcel.ActiveSheet
    currentRow = 3
    
blnChk_BalanceEquilibrée = True
blnChk_BalanceStock = False
blnChk_COMPTEOUV = False
mMsg = lMsg
IbmAmjMin = dateIBM(lAMJMin)
IbmAmjMax = dateIBM(lAMJMax)
wAMJ_6M_00 = Fix((dateElp("MoisAdd", -6, YBIATAB0_DATE_CPT_J) - 19000000) / 100) * 100
wAMJ_6M_99 = wAMJ_6M_00 + 99
meCV1.DeviseN = 0
meCV1.Montant = 0
mBalance_CV = "J  "
meCV1.OpéAmj = YBIATAB0_DATE_CPT_J
meCV2.OpéAmj = YBIATAB0_DATE_CPT_J
If lAMJMax = YBIATAB0_DATE_CPT_MP1 Then
        mBalance_CV = "MP1"
        meCV1.OpéAmj = YBIATAB0_DATE_CPT_MP1
        meCV2.OpéAmj = YBIATAB0_DATE_CPT_MP1
End If
If lAMJMax = YBIATAB0_DATE_CPT_AP1 Then
        mBalance_CV = "AP1"
        meCV1.OpéAmj = YBIATAB0_DATE_CPT_AP1
        meCV2.OpéAmj = YBIATAB0_DATE_CPT_AP1
End If
blnMOUVEMDCO = False
blnRésidence = False: mRésidence = "-"
blnCompte = False:    curCumul_Db = 0: curCumul_Cr = 0
blnClient_Line = False
nbClient_Line = 0: nbListe_Line = 0
curClient_Db = 0: curClient_Cr = 0
curListe_Db = 0: curListe_Cr = 0
curCOMPTEDEV_Db = 0: curCOMPTEDEV_Cr = 0
curCOMPTEDEV_Db_EUR = 0: curCOMPTEDEV_Cr_EUR = 0
curCOMPTEOBL_Db = 0: curCOMPTEOBL_Cr = 0
curCOMPTEOBL_Db_EUR = 0: curCOMPTEOBL_Cr_EUR = 0
curW_Db = 0: curW_Cr = 0
rsYBIACPT0_Init prevYBIACPT0
prtSAB_Balance_Total_Init
blnBalance_B_Détail = False
blnBalance_B_Récap = False
blnBalance_B_COMPTEDEV = False
blnBalance_Compte_Soldé = False
blnBalance_Pays = False
blnBalance_Récap_Bilan = False
nbCOMPTEOBL_Line = 0: nbCOMPTEDEV_Line = 0
mBalance_ZSOLDE0 = 0
idFile_CSV = 0
blnFile_CSV = False
prtTitleText = "Relevé FOTC du " & dateImp10(lAMJMin) & " au " & dateImp10(lAMJMax)
wsExcel.Cells(1, 4) = prtTitleText
wsExcel.Cells(currentRow, 5).Font.Color = vbBlue
wsExcel.Cells(currentRow, 5).Font.Size = 8
wsExcel.Cells(currentRow, 5).Font.Name = "Arial"
blnSoldeZ = False
For I = 1 To fgW.Rows - 1
    fgW.Row = I
    fgW.Col = fgW.Cols - 1: wIndex = Val(fgW.Text)
    meYBIACPT0 = larrYBIACPT0(wIndex)
    meCV1.DeviseIso = meYBIACPT0.COMPTEDEV
    Call prtSAB_Balance_R_RELEVE_FOTC(wsExcel, currentRow)
    prevYBIACPT0 = meYBIACPT0
Next I
Call wbExcel.Close(True)
Set wbExcel = Nothing
appExcelPublic.Quit
Set appExcelPublic = Nothing
Call frmElpPrt.prtIMP_PDF_NoPaper_CopyFile("S54", paramIMP_PDF_Path_Temp & "\" & nomExcel, "Archive", "BIA-BAL-RELEVE-FOTC")
Call frmElpPrt.prtIMP_PDF_NoPaper_Mail_RELEVE_FOTC("BIA_RELEVE", "@FOTC", "Relevé FOTC du " & dateImp10(lAMJMin), prtIMP_PDF_FileName)

End Sub

Public Sub prtSAB_Balance_Montant_csvManual(lcurX As Currency, lcurxcr As Currency, devise As String, ByRef maligne As String)
Dim X As String

prtSAB_Balance_CV lcurX

X = Abs(lcurX)
If lcurX > 0 Then
    maligne = maligne & X & ";0.00;" & devise & ";"
Else
    maligne = maligne & "0.00;" & X & ";" & devise & ";"
End If

X = Abs(meCV2.Montant)
If meCV2.Montant > 0 Then
    maligne = maligne & X & ";;0.00;"
Else
    maligne = maligne & "0.00;;" & X & ";"
End If

End Sub

Public Sub prtSAB_Balance_Montant_RELEVE_FOTC(lcurX As Currency, ByRef wsExcel As Excel.Worksheet, currentRow As Long)
Dim X As String, mColor As Long

prtSAB_Balance_CV lcurX

X = Format$(Abs(lcurX), "### ### ### ### ##0.00")
If lcurX > 0 Then
    wsExcel.Cells(currentRow, 5) = X
    wsExcel.Cells(currentRow, 5).Font.Color = vbRed
    wsExcel.Cells(currentRow, 5).HorizontalAlignment = xlHAlignRight
    wsExcel.Cells(currentRow, 5).Font.Size = 8
    wsExcel.Cells(currentRow, 5).Font.Name = "Arial"
Else
    wsExcel.Cells(currentRow, 6) = X
    wsExcel.Cells(currentRow, 6).Font.Color = vbBlue
    wsExcel.Cells(currentRow, 6).HorizontalAlignment = xlHAlignRight
    wsExcel.Cells(currentRow, 6).Font.Size = 8
    wsExcel.Cells(currentRow, 6).Font.Name = "Arial"
End If

X = Format$(Abs(meCV2.Montant), "### ### ### ### ##0.00")
If meCV2.Montant > 0 Then
    wsExcel.Cells(currentRow, 8) = X
    wsExcel.Cells(currentRow, 8).Font.Color = vbRed
    wsExcel.Cells(currentRow, 8).HorizontalAlignment = xlHAlignRight
    wsExcel.Cells(currentRow, 8).Font.Size = 8
    wsExcel.Cells(currentRow, 8).Font.Name = "Arial"
Else
    wsExcel.Cells(currentRow, 10) = X
    wsExcel.Cells(currentRow, 10).Font.Color = vbBlue
    wsExcel.Cells(currentRow, 10).HorizontalAlignment = xlHAlignRight
    wsExcel.Cells(currentRow, 10).Font.Size = 8
    wsExcel.Cells(currentRow, 10).Font.Name = "Arial"
End If

End Sub

Public Sub prtSAB_Balance_Mvt_RELEVE_FOTC(lYBIAMVT0 As typeYBIAMVT0, ByRef wsExcel As Excel.Worksheet, ByRef currentRow As Long)
    
    currentRow = currentRow + 1
    Call prtSAB_Balance_Montant_RELEVE_FOTC(lYBIAMVT0.MOUVEMMON, wsExcel, currentRow)
    wsExcel.Cells(currentRow, 2) = lYBIAMVT0.MOUVEMOPE & " " & lYBIAMVT0.MOUVEMNUM & " " & lYBIAMVT0.MOUVEMEVE
    wsExcel.Cells(currentRow, 2).Font.Italic = True
    wsExcel.Cells(currentRow, 3) = Trim(lYBIAMVT0.LIBELLIB1) & " " & Trim(lYBIAMVT0.LIBELLIB2) & " " & Trim(lYBIAMVT0.LIBELLIB3)
    wsExcel.Cells(currentRow, 3).Font.Italic = True
    wsExcel.Cells(currentRow, 4) = dateIBM10(lYBIAMVT0.MOUVEMDTR, True)
    wsExcel.Cells(currentRow, 4).Font.Italic = True

End Sub

Public Sub prtSAB_Balance_NewLine_xlsManual(ByRef currentRow As Long, ByRef wsExcel As Excel.Worksheet, ByRef comptageRows As Long, maxRows As Long, maxRowsPlus As Long)

    If currentRow >= maxRows + maxRowsPlus Then
        If comptageRows >= maxRows Then
            Call insere_entete_page(wsExcel, "1:3", 3, currentRow)
            comptageRows = 3
            currentRow = currentRow + 3
        End If
    End If
    comptageRows = comptageRows + 1
    currentRow = currentRow + 1
    
End Sub

Public Sub prtSAB_Balance_B_Total_Prt_xlsManual(lK As Long, lSAB_BALANCE As typeSAB_BALANCE, ByRef currentRow As Long, wsExcel As Excel.Worksheet, ByRef comptageRows As Long, maxRows As Long, maxRowsPlus As Long)
Dim curX As Currency, X As String

Call prtSAB_Balance_NewLine_xlsManual(currentRow, wsExcel, comptageRows, maxRows, maxRowsPlus)
Range("A5:J5").Select
Selection.Copy
Range("A" & CStr(currentRow)).Select
ActiveSheet.Paste

If lK = 6 Then
    wsExcel.Cells(currentRow, 5) = lSAB_BALANCE.COMPTEOBL
    If blnBalance_Pays Then
        XPrt.FontSize = 6: XPrt.FontBold = False
        XPrt.CurrentX = prtMinX - 180: XPrt.Print lSAB_BALANCE.CLIENARSD;
    End If

Else
    XPrt.CurrentX = prtMinX + 7000: XPrt.Print lSAB_BALANCE.COMPTEOBL;
End If



XPrt.CurrentX = prtMinX + 2000: XPrt.Print lSAB_BALANCE.COMPTEINT;
XPrt.CurrentX = prtMinX + 11600: XPrt.Print lSAB_BALANCE.COMPTEDEV;

curX = Abs(lSAB_BALANCE.DB)
If curX <> 0 Then
    X = Format$(curX, "### ### ### ### ##0.00")
    XPrt.CurrentX = prtMinX + 9400 - XPrt.TextWidth(X)
    XPrt.Print X;
End If

curX = Abs(lSAB_BALANCE.CR)
If curX <> 0 Then
    X = Format$(curX, "### ### ### ### ##0.00")
    XPrt.CurrentX = prtMinX + 11400 - XPrt.TextWidth(X)
    XPrt.Print X;
End If

curX = Abs(lSAB_BALANCE.DB_EUR)
If curX <> 0 Then
    X = Format$(curX, "### ### ### ### ##0.00")
    XPrt.CurrentX = prtMaxX - 2100 - XPrt.TextWidth(X)
    XPrt.Print X;
End If

curX = Abs(lSAB_BALANCE.CR_EUR)
If curX <> 0 Then
    X = Format$(curX, "### ### ### ### ##0.00")
    XPrt.CurrentX = prtMaxX - 100 - XPrt.TextWidth(X)
    XPrt.Print X;
End If



If XPrt.FontSize = 6 Then XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 8
End Sub

Public Sub prtSAB_Balance_Close_xlsManual(ByRef currentRow As Long, wsExcel As Excel.Worksheet, ByRef comptageRows As Long, maxRows As Long, maxRowsPlus As Long)

Call prtSAB_Balance_NewLine_xlsManual(currentRow, wsExcel, comptageRows, maxRows, maxRowsPlus)
Range("A6:L6").Select
Selection.Copy
Range("A" & CStr(currentRow)).Select
ActiveSheet.Paste
Call prtSAB_Balance_NewLine_xlsManual(currentRow, wsExcel, comptageRows, maxRows, maxRowsPlus)
Range("A6:L6").Select
Selection.Copy
Range("A" & CStr(currentRow)).Select
ActiveSheet.Paste
Call prtSAB_Balance_NewLine_xlsManual(currentRow, wsExcel, comptageRows, maxRows, maxRowsPlus)
Range("A6:L6").Select
Selection.Copy
Range("A" & CStr(currentRow)).Select
ActiveSheet.Paste
wsExcel.Cells(currentRow, 9) = "'="
wsExcel.Cells(currentRow, 9).Font.Color = vbMagenta
wsExcel.Cells(currentRow, 10) = "Solde balance = cumul des contrats"
wsExcel.Cells(currentRow, 10).Font.Color = vbMagenta
wsExcel.Cells(currentRow, 10).HorizontalAlignment = Excel.xlHAlignLeft
Range("J" & CStr(currentRow) & ":L" & CStr(currentRow)).Select
Selection.Merge
Call prtSAB_Balance_NewLine_xlsManual(currentRow, wsExcel, comptageRows, maxRows, maxRowsPlus)
Range("A6:L6").Select
Selection.Copy
Range("A" & CStr(currentRow)).Select
ActiveSheet.Paste
wsExcel.Cells(currentRow, 9) = "##"
wsExcel.Cells(currentRow, 9).Font.Color = vbMagenta
wsExcel.Cells(currentRow, 10) = "Solde balance <> cumul des contrats"
wsExcel.Cells(currentRow, 10).Font.Color = vbMagenta
wsExcel.Cells(currentRow, 10).HorizontalAlignment = Excel.xlHAlignLeft
Range("J" & CStr(currentRow) & ":L" & CStr(currentRow)).Select
Selection.Merge
Call prtSAB_Balance_NewLine_xlsManual(currentRow, wsExcel, comptageRows, maxRows, maxRowsPlus)
Range("A6:L6").Select
Selection.Copy
Range("A" & CStr(currentRow)).Select
ActiveSheet.Paste
wsExcel.Cells(currentRow, 9) = "??"
wsExcel.Cells(currentRow, 9).Font.Color = vbMagenta
wsExcel.Cells(currentRow, 10) = "Solde balance, aucun contrat rattaché??"
wsExcel.Cells(currentRow, 10).Font.Color = vbMagenta
wsExcel.Cells(currentRow, 10).HorizontalAlignment = Excel.xlHAlignLeft
Range("J" & CStr(currentRow) & ":L" & CStr(currentRow)).Select
Selection.Merge
Call prtSAB_Balance_NewLine_xlsManual(currentRow, wsExcel, comptageRows, maxRows, maxRowsPlus)
Range("A6:L6").Select
Selection.Copy
Range("A" & CStr(currentRow)).Select
ActiveSheet.Paste
wsExcel.Cells(currentRow, 3) = "CAV et LOR : date de dernier mouvement hors échelles, facturation ...(+) indique qu'il y a des mouvements postérieurs initiés par la banque"
wsExcel.Cells(currentRow, 3).Font.Color = vbMagenta
Range("C" & CStr(currentRow) & ":G" & CStr(currentRow)).Select
Selection.Merge
wsExcel.Cells(currentRow, 9) = "6M"
wsExcel.Cells(currentRow, 9).Font.Color = vbMagenta
wsExcel.Cells(currentRow, 10) = "Compte ouvert depuis 6 mois"
wsExcel.Cells(currentRow, 10).Font.Color = vbMagenta
wsExcel.Cells(currentRow, 10).HorizontalAlignment = Excel.xlHAlignLeft
Range("J" & CStr(currentRow) & ":L" & CStr(currentRow)).Select
Selection.Merge

End Sub


Public Sub prtSAB_Balance_L_xlsManual(lFct As String, lcurX As Currency, ByRef currentRow As Long, wsExcel As Excel.Worksheet)
Dim wForecolor As Long

If lcurX > 0 Then
    wForecolor = vbRed
Else
    wForecolor = prtForeColor
End If

Select Case lFct
    Case "R0", "R1":
        Range("A5:J5").Select
        Selection.Copy
        Range("A" & CStr(currentRow)).Select
        ActiveSheet.Paste
        Call prtSAB_Balance_Montant_xlsManual(lcurX, currentRow, wsExcel)
    Case "C1"
        prtSAB_Balance_Montant lcurX
    Case Else
        Call prtSAB_Balance_Montant_xlsManual(lcurX, currentRow, wsExcel)
End Select

Select Case meYBIACPT0.COMPTEFON
    Case 0: If Mid$(lFct, 1, 1) <> "R" Then wForecolor = prtForeColor
    Case 4: wForecolor = vbRed: wsExcel.Cells(currentRow, 4) = dateIBM10(meYBIACPT0.COMPTECLO, True)
            wsExcel.Cells(currentRow, 3).Font.Color = vbRed
    Case Else: wForecolor = vbMagenta
End Select
wsExcel.Cells(currentRow, 1) = meYBIACPT0.PLANCOPRO
wsExcel.Cells(currentRow, 1).Font.Color = wForecolor
wsExcel.Cells(currentRow, 2) = meYBIACPT0.COMPTECOM
wsExcel.Cells(currentRow, 2).Font.Color = wForecolor
'If lFct = "B" Then
'    XPrt.FontBold = False
'Else
'    XPrt.FontBold = True
'End If
wsExcel.Cells(currentRow, 3) = meYBIACPT0.COMPTEINT
wsExcel.Cells(currentRow, 3).Font.Color = wForecolor
'XPrt.FontBold = False
If lFct = "L" Then
'_____________________________
           If meDORCPTDMV > 0 Then
                XPrt.CurrentX = prtMinX + 6800: XPrt.Print dateIBM10(meDORCPTDMV, True);
                If meDORCPTDMV <> meYBIACPT0.SOLDEDMO Then XPrt.Print " +";
            Else
                XPrt.CurrentX = prtMinX + 6800: XPrt.Print dateIBM10(meYBIACPT0.SOLDEDMO, True);
            End If
'____________________
End If

If Mid$(lFct, 1, 1) <> "R" Then
    wForecolor = prtForeColor
End If
If blnChk_BalanceStock Then
    wsExcel.Cells(currentRow, 8) = meYBIACPT0.COMPTEDEV
    wsExcel.Cells(currentRow, 8).Font.Color = wForecolor
Else
    wsExcel.Cells(currentRow, 7) = meYBIACPT0.COMPTEDEV
    wsExcel.Cells(currentRow, 7).Font.Color = wForecolor
End If
End Sub

Public Sub prtSAB_Balance_Monitor(lFct As String, lAMJMin As String, lAMJMax As String, fgW As MSFlexGrid, larrYBIACPT0() As typeYBIACPT0, larrYBIACPT0_Nb As Long, lMsg As String, lYSTOMON() As Currency, lDORCPTDMV() As Long)

' B : balance 1,1 : "B"
'             2,1 : "D" (Rupture Devise/ PCEC)              blnBalance_B_COMPTEDEV
'                 : "P" (Rupture PCEC/Devise)
'             3,1 : "D" (imprimer lignes détail des compes  blnBalance_Détail
'             4,1 : "V" (solde Veille)
'                 : "M" (solde fin de mois)
'                 : "2" (solde fin de mois -2)
'                 : "A" (solde fin d'année)
'             5,1 : "1" Recap
'             6,1 : "1" CSV
'             7,3 : N0 Fichier
'            10,1 : "1" Ignorer les comptes soldés
'            11,1 : "1" Balance par pays
'            12,1 : "1" Recap Bilan /Hors-Bilan

'            16,6 : "1"Recap niveau 0 (DEV),Gras,Souliné,Trame
'            22,6 : "1" Recap niveau 1
'            28,6 : "1" Recap niveau 2
'            34,6 : "1" Recap niveau 3
'            40,6 : "1" Recap niveau 4
'            46,6 : "1" Recap niveau 5
'            52,6 : "1" Recap niveau 6
'            58,6 : "1"  total de la balance Détail niveau 7

Dim wIndex As Long, I As Integer
Dim mFct1 As String

blnChk_BalanceEquilibrée = True
blnChk_BalanceStock = False
blnChk_COMPTEOUV = False

mFct1 = Mid$(lFct, 1, 1)
If mFct1 = "S" Then
    mFct1 = "B"
    blnChk_BalanceEquilibrée = False
    blnChk_BalanceStock = True
    
End If

mMsg = lMsg
IbmAmjMin = dateIBM(lAMJMin)
IbmAmjMax = dateIBM(lAMJMax)

wAMJ_6M_00 = Fix((dateElp("MoisAdd", -6, YBIATAB0_DATE_CPT_J) - 19000000) / 100) * 100
wAMJ_6M_99 = wAMJ_6M_00 + 99

meCV1.DeviseN = 0
meCV1.Montant = 0

mBalance_CV = "J  "
meCV1.OpéAmj = YBIATAB0_DATE_CPT_J
meCV2.OpéAmj = YBIATAB0_DATE_CPT_J
If lAMJMax = YBIATAB0_DATE_CPT_MP1 Then
        mBalance_CV = "MP1"
        meCV1.OpéAmj = YBIATAB0_DATE_CPT_MP1
        meCV2.OpéAmj = YBIATAB0_DATE_CPT_MP1
End If
If lAMJMax = YBIATAB0_DATE_CPT_AP1 Then
        mBalance_CV = "AP1"
        meCV1.OpéAmj = YBIATAB0_DATE_CPT_AP1
        meCV2.OpéAmj = YBIATAB0_DATE_CPT_AP1
End If



blnMOUVEMDCO = False
blnRésidence = False: mRésidence = "-"
blnCompte = False:    curCumul_Db = 0: curCumul_Cr = 0
blnClient_Line = False
nbClient_Line = 0: nbListe_Line = 0
curClient_Db = 0: curClient_Cr = 0
curListe_Db = 0: curListe_Cr = 0
curCOMPTEDEV_Db = 0: curCOMPTEDEV_Cr = 0
curCOMPTEDEV_Db_EUR = 0: curCOMPTEDEV_Cr_EUR = 0
curCOMPTEOBL_Db = 0: curCOMPTEOBL_Cr = 0
curCOMPTEOBL_Db_EUR = 0: curCOMPTEOBL_Cr_EUR = 0
curW_Db = 0: curW_Cr = 0
rsYBIACPT0_Init prevYBIACPT0

prtSAB_Balance_Total_Init
blnBalance_B_Détail = False
blnBalance_B_Récap = False
blnBalance_B_COMPTEDEV = False
blnBalance_Compte_Soldé = False
blnBalance_Pays = False
blnBalance_Récap_Bilan = False
nbCOMPTEOBL_Line = 0: nbCOMPTEDEV_Line = 0

mBalance_ZSOLDE0 = 0
idFile_CSV = 0
blnFile_CSV = False

Select Case mFct1
    Case "L": prtTitleText = "Liste au " & dateImp10(lAMJMax)
                If Mid$(lFct, 2, 1) = "T" Then blnClient_Line = True
                Select Case Mid$(lFct, 1, 7)
                    Case "LT-CPT-": prtTitleText = "Position comptable du groupe " & Mid$(lFct, 8, 3) & " au " & dateImp10(lAMJMax)
                                    meCV1.OpéAmj = lAMJMin
                    Case "LT-ENG-": prtTitleText = "Etat des engagements du groupe " & Mid$(lFct, 8, 3) & " au " & dateImp10(lAMJMax)
                                    meCV1.OpéAmj = lAMJMin
                End Select
                
    Case "B":
                If Mid$(lFct, 2, 1) = "D" Then blnBalance_B_COMPTEDEV = True
                If Mid$(lFct, 3, 1) = "1" Then blnBalance_B_Détail = True
                If Mid$(lFct, 5, 1) = "1" Then blnBalance_B_Récap = True
                If Mid$(lFct, 10, 1) = "1" Then blnBalance_Compte_Soldé = True
                If Mid$(lFct, 12, 1) = "1" Then blnBalance_Récap_Bilan = True
                
                Select Case Mid$(lFct, 4, 1)
                    Case "M":   mBalance_ZSOLDE0 = 1
                                mBalance_CV = "MP1"
                                meCV1.OpéAmj = YBIATAB0_DATE_CPT_MP1
                                meCV2.OpéAmj = YBIATAB0_DATE_CPT_MP1
                    Case "2":   mBalance_ZSOLDE0 = 2
                                mBalance_CV = YBIATAB0_DATE_CPT_MP2
                                meCV1.OpéAmj = YBIATAB0_DATE_CPT_MP2
                                meCV2.OpéAmj = YBIATAB0_DATE_CPT_MP2
                    Case "A":   mBalance_ZSOLDE0 = 1 + Val(Mid$(YBIATAB0_DATE_CPT_MP1, 5, 2))
                                mBalance_CV = "AP1"
                                meCV1.OpéAmj = YBIATAB0_DATE_CPT_AP1
                                meCV2.OpéAmj = YBIATAB0_DATE_CPT_AP1
                End Select
                
                If Mid$(lFct, 6, 1) = "1" Then blnFile_CSV = True
                idFile_CSV = Val(Mid$(lFct, 7, 3))

                prtSAB_Balance_B_Total_Init Mid$(lFct, 16, 6), totalSAB_BALANCE(0)
                prtSAB_Balance_B_Total_Init Mid$(lFct, 22, 6), totalSAB_BALANCE(1)
                prtSAB_Balance_B_Total_Init Mid$(lFct, 28, 6), totalSAB_BALANCE(2)
                prtSAB_Balance_B_Total_Init Mid$(lFct, 34, 6), totalSAB_BALANCE(3)
                prtSAB_Balance_B_Total_Init Mid$(lFct, 40, 6), totalSAB_BALANCE(4)
                prtSAB_Balance_B_Total_Init Mid$(lFct, 46, 6), totalSAB_BALANCE(5)
                prtSAB_Balance_B_Total_Init Mid$(lFct, 52, 6), totalSAB_BALANCE(6)
                prtSAB_Balance_B_Total_Init Mid$(lFct, 58, 6), détailSAB_BALANCE
                                
                If blnChk_BalanceStock Then
                    prtTitleText = mMsg & " - Balance / Stock au " & dateImp10(meCV1.OpéAmj)
                Else
                     prtTitleText = "Balance au " & dateImp10(meCV1.OpéAmj)
               End If
                
                If Mid$(lFct, 11, 1) = "1" Then blnBalance_Pays = True: prtTitleText = mMsg & " Balance par PAYS de résidence au " & dateImp10(meCV1.OpéAmj)
                xTitleText = prtTitleText
    Case "C": prtTitleText = "Cumul de mouvements du " & dateImp10(lAMJMin) & " au " & dateImp10(lAMJMax)
               If Mid$(lFct, 2, 3) = "DCO" Then blnMOUVEMDCO = True
               If Mid$(lFct, 5, 1) = "R" Then blnRésidence = True
    Case "R": prtTitleText = "Relevé du " & dateImp10(lAMJMin) & " au " & dateImp10(lAMJMax)

End Select

If Mid$(lFct, 6, 1) = "Z" Then
    blnSoldeZ = True
Else
    blnSoldeZ = False
End If

prtFontName = prtFontName_Arial
prtSAB_Balance_Open
prtHeaderHeight = 300
prtSAB_Balance_Form
XPrt.CurrentY = XPrt.CurrentY - prtlineHeight

XPrt.FontSize = 8
For I = 1 To fgW.Rows - 1
    
    fgW.Row = I
    fgW.Col = fgW.Cols - 1: wIndex = Val(fgW.Text)
    meYBIACPT0 = larrYBIACPT0(wIndex)
    meCV1.DeviseIso = meYBIACPT0.COMPTEDEV
    If blnChk_BalanceStock Then meYSTOMON = lYSTOMON(wIndex): meDORCPTDMV = lDORCPTDMV(wIndex)
    Select Case mFct1
        Case "B": prtSAB_Balance_B_Line
        Case "R": prtSAB_Balance_R
        Case "C": prtSAB_Balance_C
        Case Else: meDORCPTDMV = lDORCPTDMV(wIndex): prtSAB_Balance_L_Line
     End Select
     
     prevYBIACPT0 = meYBIACPT0
Next I

If mFct1 = "C" And blnRésidence Then prtSAB_Balance_C_Cumul


If mFct1 = "B" Then
    prtSAB_Balance_B_Fin

Else
    If mFct1 = "L" Then
        
        prtSAB_Balance_L_Rupture
        If curListe_Db <> curW_Db Or curListe_Cr <> curW_Cr Then
            prtSAB_Balance_NewLine
            XPrt.FontSize = 12: XPrt.FontBold = True: XPrt.ForeColor = vbMagenta
            frmElpPrt.prtCentré prtMedX, "ERREUR TOTALISATION"
            XPrt.FontSize = 8: XPrt.FontBold = False: XPrt.ForeColor = vbBlack
        End If
        XPrt.DrawWidth = 10
        XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
        XPrt.Line (prtMinX + 12000, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY), prtLineColor
        XPrt.CurrentY = XPrt.CurrentY - prtlineHeight
        prevYBIACPT0.CLIENACLI = ""
        prevYBIACPT0.CLIENASIG = ""
        nbClient_Line = nbListe_Line
        curClient_Db = curListe_Db
        curClient_Cr = curListe_Cr
        prtSAB_Balance_L_Rupture
    Else
    
        XPrt.DrawWidth = 10
        XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
        XPrt.Line (prtMinX + 7500, XPrt.CurrentY)-(prtMinX + 12000, XPrt.CurrentY), prtLineColor
    End If
End If

prtSAB_Balance_Close

End Sub

Public Sub prtSAB_Balance_Monitor_xlsManual(lFct As String, lAMJMin As String, lAMJMax As String, fgW As MSFlexGrid, larrYBIACPT0() As typeYBIACPT0, larrYBIACPT0_Nb As Long, lMsg As String, lYSTOMON() As Currency, lDORCPTDMV() As Long, ByRef currentRowStock As Long, wsexcelStock As Excel.Worksheet)
Dim nomModele As String
Dim nomExcel As String
Dim currentSheet As Long
Dim currentRow As Long
Dim comptageRows As Long
Dim maxRows As Long
Dim maxRowsPlus As Long
Dim xFile As String

' B : balance 1,1 : "B"
'             2,1 : "D" (Rupture Devise/ PCEC)              blnBalance_B_COMPTEDEV
'                 : "P" (Rupture PCEC/Devise)
'             3,1 : "D" (imprimer lignes détail des compes  blnBalance_Détail
'             4,1 : "V" (solde Veille)
'                 : "M" (solde fin de mois)
'                 : "2" (solde fin de mois -2)
'                 : "A" (solde fin d'année)
'             5,1 : "1" Recap
'             6,1 : "1" CSV
'             7,3 : N0 Fichier
'            10,1 : "1" Ignorer les comptes soldés
'            11,1 : "1" Balance par pays
'            12,1 : "1" Recap Bilan /Hors-Bilan

'            16,6 : "1"Recap niveau 0 (DEV),Gras,Souliné,Trame
'            22,6 : "1" Recap niveau 1
'            28,6 : "1" Recap niveau 2
'            34,6 : "1" Recap niveau 3
'            40,6 : "1" Recap niveau 4
'            46,6 : "1" Recap niveau 5
'            52,6 : "1" Recap niveau 6
'            58,6 : "1"  total de la balance Détail niveau 7

Dim wIndex As Long, I As Integer
Dim mFct1 As String

blnChk_BalanceEquilibrée = True
blnChk_BalanceStock = False
blnChk_COMPTEOUV = False

mFct1 = Mid$(lFct, 1, 1)
If mFct1 = "S" Then
    mFct1 = "B"
    blnChk_BalanceEquilibrée = False
    blnChk_BalanceStock = True
    
End If

mMsg = lMsg
IbmAmjMin = dateIBM(lAMJMin)
IbmAmjMax = dateIBM(lAMJMax)

wAMJ_6M_00 = Fix((dateElp("MoisAdd", -6, YBIATAB0_DATE_CPT_J) - 19000000) / 100) * 100
wAMJ_6M_99 = wAMJ_6M_00 + 99

meCV1.DeviseN = 0
meCV1.Montant = 0

mBalance_CV = "J  "
meCV1.OpéAmj = YBIATAB0_DATE_CPT_J
meCV2.OpéAmj = YBIATAB0_DATE_CPT_J

If lAMJMax = YBIATAB0_DATE_CPT_MP1 Then
        mBalance_CV = "MP1"
        meCV1.OpéAmj = YBIATAB0_DATE_CPT_MP1
        meCV2.OpéAmj = YBIATAB0_DATE_CPT_MP1
End If
If lAMJMax = YBIATAB0_DATE_CPT_AP1 Then
        mBalance_CV = "AP1"
        meCV1.OpéAmj = YBIATAB0_DATE_CPT_AP1
        meCV2.OpéAmj = YBIATAB0_DATE_CPT_AP1
End If



blnMOUVEMDCO = False
blnRésidence = False: mRésidence = "-"
blnCompte = False:    curCumul_Db = 0: curCumul_Cr = 0
blnClient_Line = False
nbClient_Line = 0: nbListe_Line = 0
curClient_Db = 0: curClient_Cr = 0
curListe_Db = 0: curListe_Cr = 0
curCOMPTEDEV_Db = 0: curCOMPTEDEV_Cr = 0
curCOMPTEDEV_Db_EUR = 0: curCOMPTEDEV_Cr_EUR = 0
curCOMPTEOBL_Db = 0: curCOMPTEOBL_Cr = 0
curCOMPTEOBL_Db_EUR = 0: curCOMPTEOBL_Cr_EUR = 0
curW_Db = 0: curW_Cr = 0
rsYBIACPT0_Init prevYBIACPT0

prtSAB_Balance_Total_Init
blnBalance_B_Détail = False
blnBalance_B_Récap = False
blnBalance_B_COMPTEDEV = False
blnBalance_Compte_Soldé = False
blnBalance_Pays = False
blnBalance_Récap_Bilan = False
nbCOMPTEOBL_Line = 0: nbCOMPTEDEV_Line = 0

mBalance_ZSOLDE0 = 0
idFile_CSV = 0
blnFile_CSV = False
maxRows = 37

Select Case mFct1
    Case "L": prtTitleText = "Liste au " & dateImp10(lAMJMax)
                If Mid$(lFct, 2, 1) = "T" Then blnClient_Line = True
                Select Case Mid$(lFct, 1, 7)
                    Case "LT-CPT-": prtTitleText = "Position comptable du groupe " & Mid$(lFct, 8, 3) & " au " & dateImp10(lAMJMax)
                                    meCV1.OpéAmj = lAMJMin
                    Case "LT-ENG-": prtTitleText = "Etat des engagements du groupe " & Mid$(lFct, 8, 3) & " au " & dateImp10(lAMJMax)
                                    meCV1.OpéAmj = lAMJMin
                End Select
                
    Case "B":
                If Mid$(lFct, 2, 1) = "D" Then blnBalance_B_COMPTEDEV = True 'ok
                If Mid$(lFct, 3, 1) = "1" Then blnBalance_B_Détail = True
                If Mid$(lFct, 5, 1) = "1" Then blnBalance_B_Récap = True
                If Mid$(lFct, 10, 1) = "1" Then blnBalance_Compte_Soldé = True 'ok
                If Mid$(lFct, 12, 1) = "1" Then blnBalance_Récap_Bilan = True 'ok
                
                Select Case Mid$(lFct, 4, 1)
                    Case "M":   mBalance_ZSOLDE0 = 1
                                mBalance_CV = "MP1"
                                meCV1.OpéAmj = YBIATAB0_DATE_CPT_MP1
                                meCV2.OpéAmj = YBIATAB0_DATE_CPT_MP1
                    Case "2":   mBalance_ZSOLDE0 = 2
                                mBalance_CV = YBIATAB0_DATE_CPT_MP2
                                meCV1.OpéAmj = YBIATAB0_DATE_CPT_MP2
                                meCV2.OpéAmj = YBIATAB0_DATE_CPT_MP2
                    Case "A":   mBalance_ZSOLDE0 = 1 + Val(Mid$(YBIATAB0_DATE_CPT_MP1, 5, 2))
                                mBalance_CV = "AP1"
                                meCV1.OpéAmj = YBIATAB0_DATE_CPT_AP1
                                meCV2.OpéAmj = YBIATAB0_DATE_CPT_AP1
                End Select
                
                If Mid$(lFct, 6, 1) = "1" Then blnFile_CSV = True
                idFile_CSV = Val(Mid$(lFct, 7, 3))

                prtSAB_Balance_B_Total_Init Mid$(lFct, 16, 6), totalSAB_BALANCE(0)
                prtSAB_Balance_B_Total_Init Mid$(lFct, 22, 6), totalSAB_BALANCE(1)
                prtSAB_Balance_B_Total_Init Mid$(lFct, 28, 6), totalSAB_BALANCE(2)
                prtSAB_Balance_B_Total_Init Mid$(lFct, 34, 6), totalSAB_BALANCE(3)
                prtSAB_Balance_B_Total_Init Mid$(lFct, 40, 6), totalSAB_BALANCE(4)
                prtSAB_Balance_B_Total_Init Mid$(lFct, 46, 6), totalSAB_BALANCE(5)
                prtSAB_Balance_B_Total_Init Mid$(lFct, 52, 6), totalSAB_BALANCE(6)
                prtSAB_Balance_B_Total_Init Mid$(lFct, 58, 6), détailSAB_BALANCE
                                
                If blnChk_BalanceStock Then
                    prtTitleText = mMsg & " - Balance / Stock au " & dateImp10(meCV1.OpéAmj)
                    nomExcel = "modele_BALANCE_Stock.xlsx"
                    nomModele = paramFolder_Local & "\Modeles\" & nomExcel
                    currentRow = 7
                    comptageRows = currentRow
                    maxRowsPlus = 4
                    maxRows = 38
                Else
                    prtTitleText = "Balance au " & dateImp10(meCV1.OpéAmj)
                    nomExcel = "modele_BAL_B_HB.xlsx"
                    nomModele = paramFolder_Local & "\Modeles\" & nomExcel
                    currentRow = 8
                    comptageRows = currentRow
                    maxRowsPlus = 5
                    maxRows = 35
                End If
                
                If Mid$(lFct, 11, 1) = "1" Then
                    blnBalance_Pays = True
                    prtTitleText = mMsg & " Balance par PAYS de résidence au " & dateImp10(meCV1.OpéAmj)
                End If
                xTitleText = prtTitleText
    Case "C": prtTitleText = "Cumul de mouvements du " & dateImp10(lAMJMin) & " au " & dateImp10(lAMJMax)
               If Mid$(lFct, 2, 3) = "DCO" Then blnMOUVEMDCO = True
               If Mid$(lFct, 5, 1) = "R" Then blnRésidence = True
                nomExcel = "modele_BALANCE_Mvts.xlsx"
                nomModele = paramFolder_Local & "\Modeles\" & nomExcel
                currentRow = 5
                comptageRows = currentRow
                maxRowsPlus = 4
                maxRows = 35
    Case "R": prtTitleText = "Relevé du " & dateImp10(lAMJMin) & " au " & dateImp10(lAMJMax)
                nomExcel = "modele_BAL6000.xlsx"
                nomModele = paramFolder_Local & "\Modeles\" & nomExcel
                currentRow = 8
                comptageRows = currentRow
                maxRowsPlus = 5
                maxRows = 33

End Select

If Mid$(lFct, 6, 1) = "Z" Then
    blnSoldeZ = True
Else
    blnSoldeZ = False
End If
    If fgW.Rows > 1 And Not blnChk_BalanceStock Then
        If appExcelPublic Is Nothing Then
            Set appExcelPublic = CreateObject("Excel.Application")
            appExcelPublic.Visible = False
            appExcelPublic.ControlCharacters = False
            appExcelPublic.Interactive = False
        End If
        'On recopie le classeur modèle de c:\BIASRV vers c:\temp\imp_pdf
        FileCopy nomModele, paramIMP_PDF_Path_Temp & "\" & nomExcel
        'on charge CE classeur dans Excel
        Call init_xlsManual
        Call appExcelPublic.Workbooks.Open(paramIMP_PDF_Path_Temp & "\" & nomExcel)
        Set wbExcel2 = appExcelPublic.ActiveWorkbook
        With wbExcel2
            .Title = .Sheets(1).Name
            .Subject = .Sheets(1).Name
        End With
        currentSheet = 1
        wbExcel2.Sheets(currentSheet).Cells(1, 4) = prtTitleText
    ElseIf fgW.Rows > 1 And blnChk_BalanceStock Then
        wsexcelStock.Cells(1, 5) = prtTitleText
    End If
For I = 1 To fgW.Rows - 1
    fgW.Row = I
    fgW.Col = fgW.Cols - 1: wIndex = Val(fgW.Text)
    meYBIACPT0 = larrYBIACPT0(wIndex)
    meCV1.DeviseIso = meYBIACPT0.COMPTEDEV
    If blnChk_BalanceStock Then
        meYSTOMON = lYSTOMON(wIndex)
        meDORCPTDMV = lDORCPTDMV(wIndex)
    End If
    Select Case mFct1
        Case "B":
            If blnChk_BalanceStock Then
                Call prtSAB_Balance_B_Line_xlsManual(currentRowStock, wsexcelStock, comptageRows, maxRows, maxRowsPlus)
            Else
                Call prtSAB_Balance_B_Line_xlsManual(currentRow, wbExcel2.Sheets(currentSheet), comptageRows, maxRows, maxRowsPlus)
            End If
        Case "R": Call prtSAB_Balance_R_xlsManual(currentRow, wbExcel2.Sheets(currentSheet), comptageRows, maxRows, maxRowsPlus)
        Case "C": Call prtSAB_Balance_C_xlsManual(currentRow, wbExcel2.Sheets(currentSheet))
        Case Else: meDORCPTDMV = lDORCPTDMV(wIndex): prtSAB_Balance_L_Line
     End Select
     
     prevYBIACPT0 = meYBIACPT0
Next I

If mFct1 = "C" And blnRésidence Then prtSAB_Balance_C_Cumul


If fgW.Rows > 1 Then
    If mFct1 = "B" Then
        If blnChk_BalanceStock Then
            Call prtSAB_Balance_B_Fin_xlsManual(currentRowStock, wsexcelStock, comptageRows, maxRows, maxRowsPlus)
        Else
            Call prtSAB_Balance_B_Fin_xlsManual(currentRow, wbExcel2.Sheets(currentSheet), comptageRows, maxRows, maxRowsPlus)
        End If
    Else
        If mFct1 = "L" Then
            
            prtSAB_Balance_L_Rupture
            If curListe_Db <> curW_Db Or curListe_Cr <> curW_Cr Then
                prtSAB_Balance_NewLine
                XPrt.FontSize = 12: XPrt.FontBold = True: XPrt.ForeColor = vbMagenta
                frmElpPrt.prtCentré prtMedX, "ERREUR TOTALISATION"
                XPrt.FontSize = 8: XPrt.FontBold = False: XPrt.ForeColor = vbBlack
            End If
            XPrt.DrawWidth = 10
            XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
            XPrt.Line (prtMinX + 12000, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY), prtLineColor
            XPrt.CurrentY = XPrt.CurrentY - prtlineHeight
            prevYBIACPT0.CLIENACLI = ""
            prevYBIACPT0.CLIENASIG = ""
            nbClient_Line = nbListe_Line
            curClient_Db = curListe_Db
            curClient_Cr = curListe_Cr
            prtSAB_Balance_L_Rupture
        Else
            Range("A7:J7").Select
            Selection.Copy
            Range("A" & CStr(currentRow)).Select
            ActiveSheet.Paste
        End If
    End If
    If blnChk_BalanceStock Then
        Call prtSAB_Balance_Close_xlsManual(currentRowStock, wsexcelStock, comptageRows, maxRows, maxRowsPlus)
    Else
        If mFct1 = "C" Then
            appExcelPublic.Quit
            Set appExcelPublic = Nothing
    End If
    End If
    'Pas de PDF pour la balance_Mvts
    If mFct1 <> "C" Then
        'on supprime les 4 ou 5 lignes modèles
        If blnChk_BalanceStock Then
            Rows("4:7").Select
            Selection.Delete
            currentRowStock = currentRowStock - 4
            Call frmSAB_Balance.zoneImpression_xlsManual("BALANCE_Stock", currentRowStock, wsexcelStock)
            Call ActiveSheet.ExportAsFixedFormat(xlTypePDF, paramIMP_PDF_Path_Temp & "\" & paramEditionNoPaper_Auto_PgmName & ".pdf")
            Call impressions_xlsManual.prtIMP_PDF_Monitor_xlsManual
            'wsexcelStock.Delete
        Else
            Rows("4:8").Select
            Selection.Delete
            currentRow = currentRow - 5
            Call frmSAB_Balance.zoneImpression_xlsManual(wbExcel2.Sheets(currentSheet).Name, currentRow, wbExcel2.Sheets(currentSheet))
            If Mid(lFct, 2, 1) = "V" Then
                'on stocke le fichier Excel sur DOCSRV2013 et on l'envoie en fichier attaché à SSIMEL0 SSIMELUIDX="BIA_RELEVE.@FOTC" (DAOUD)
                Call ActiveSheet.SaveAs(paramIMP_PDF_Path_Temp & "\" & paramEditionNoPaper_Auto_PgmName & ".xlsx")
                xFile = paramIMP_PDF_Path_Temp & "\" & paramEditionNoPaper_Auto_PgmName & ".xls"
                Call wbExcel2.Close(False)
                Set wbExcel2 = Nothing
                Call frmElpPrt.prtIMP_PDF_NoPaper_CopyFile("S54", xFile, "Archive", "BIA-BAL-RELEVE-FOTC")
                Call frmElpPrt.prtIMP_PDF_NoPaper_Mail("BIA_RELEVE", "@FOTC", "")
            Else
                Call ActiveSheet.ExportAsFixedFormat(xlTypePDF, paramIMP_PDF_Path_Temp & "\" & paramEditionNoPaper_Auto_PgmName & ".pdf")
                Call impressions_xlsManual.prtIMP_PDF_Monitor_xlsManual
                Call wbExcel2.Close(False)
                Set wbExcel2 = Nothing
            End If
            Kill paramIMP_PDF_Path_Temp & "\" & nomExcel
        End If
    End If
End If
End Sub
Public Sub prtSAB_Balance_Montant_xlsManual(lcurX As Currency, ByRef currentRow As Long, wsExcel As Excel.Worksheet)
Dim X As String

prtSAB_Balance_CV lcurX

If blnChk_BalanceStock Then
    X = "'" & Format$(Abs(lcurX), "### ### ### ### ##0.00")
    If lcurX > 0 Then
        wsExcel.Cells(currentRow, 6) = X
        wsExcel.Cells(currentRow, 6).Font.Color = vbRed
    Else
        wsExcel.Cells(currentRow, 7) = X
        wsExcel.Cells(currentRow, 7).Font.Color = vbBlue
    End If
    
    X = "'" & Format$(Abs(meCV2.Montant), "### ### ### ### ##0.00")
    If meCV2.Montant > 0 Then
        wsExcel.Cells(currentRow, 10) = X
        wsExcel.Cells(currentRow, 10).Font.Color = vbRed
    Else
        wsExcel.Cells(currentRow, 12) = X
        wsExcel.Cells(currentRow, 12).Font.Color = vbBlue
    End If
Else
    X = "'" & Format$(Abs(lcurX), "### ### ### ### ##0.00")
    If lcurX > 0 Then
        wsExcel.Cells(currentRow, 5) = X
        wsExcel.Cells(currentRow, 5).Font.Color = vbRed
    Else
        wsExcel.Cells(currentRow, 6) = X
        wsExcel.Cells(currentRow, 6).Font.Color = vbBlue
    End If
    
    X = "'" & Format$(Abs(meCV2.Montant), "### ### ### ### ##0.00")
    If meCV2.Montant > 0 Then
        wsExcel.Cells(currentRow, 8) = X
        wsExcel.Cells(currentRow, 8).Font.Color = vbRed
    Else
        wsExcel.Cells(currentRow, 10) = X
        wsExcel.Cells(currentRow, 10).Font.Color = vbBlue
    End If
End If

End Sub

Public Sub prtSAB_Balance_Mvt_xlsManual(lYBIAMVT0 As typeYBIAMVT0, ByRef currentRow As Long, wsExcel As Excel.Worksheet, ByRef comptageRows As Long, maxRows As Long, maxRowsPlus As Long)

Call prtSAB_Balance_NewLine_xlsManual(currentRow, wsExcel, comptageRows, maxRows, maxRowsPlus)
Range("A6:J6").Select
Selection.Copy
Range("A" & CStr(currentRow)).Select
ActiveSheet.Paste

Call prtSAB_Balance_Montant_xlsManual(lYBIAMVT0.MOUVEMMON, currentRow, wsExcel)
wsExcel.Cells(currentRow, 2) = lYBIAMVT0.MOUVEMOPE & " " & lYBIAMVT0.MOUVEMNUM & " " & lYBIAMVT0.MOUVEMEVE
wsExcel.Cells(currentRow, 3) = Trim(lYBIAMVT0.LIBELLIB1) & " " & Trim(lYBIAMVT0.LIBELLIB2) & " " & Trim(lYBIAMVT0.LIBELLIB3)
wsExcel.Cells(currentRow, 4) = dateIBM10(lYBIAMVT0.MOUVEMDTR, True)

End Sub

Public Sub prtSAB_Balance_R_RELEVE_FOTC(ByRef wsExcel As Excel.Worksheet, ByRef currentRow As Long)
Dim X As String
Dim xYBIAMVT0 As typeYBIAMVT0, wSolde As Currency, wMvt_DB As Currency, wMvt_CR As Currency
Dim xSQL As String
Dim wMvt_Nb As Long

wSolde = meYBIACPT0.SOLDECEN
blnCompte = False
rsYBIAMVT0_Init xYBIAMVT0
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIAMVTH " _
     & " where MOUVEMCOM = '" & meYBIACPT0.COMPTECOM & "'" _
     & " and MOUVEMDTR >= " & IbmAmjMin _
     & " and MOUVEMDTR <= " & IbmAmjMax _
     & " order by MOUVEMDTR,MOUVEMPIE,MOUVEMECR"
Set rsSab = cnsab.Execute(xSQL)
Do While Not rsSab.EOF
    Call rsYBIAMVT0_GetBuffer(rsSab, xYBIAMVT0)
    If Not blnCompte Then
        wMvt_DB = 0: wMvt_CR = 0: wMvt_Nb = 0
        blnCompte = True
        wSolde = xYBIAMVT0.BIAMVTSD0
        Call prtSAB_Balance_L_RELEVE_FOTC("R0", wSolde, wsExcel, currentRow)
    End If
    
    Call prtSAB_Balance_Mvt_RELEVE_FOTC(xYBIAMVT0, wsExcel, currentRow)
    wSolde = wSolde + xYBIAMVT0.MOUVEMMON
    If xYBIAMVT0.MOUVEMOPE <> "RPC" Then
        wMvt_Nb = wMvt_Nb + 1
        If xYBIAMVT0.MOUVEMMON > 0 Then
            wMvt_DB = wMvt_DB + xYBIAMVT0.MOUVEMMON
        Else
            wMvt_CR = wMvt_CR + xYBIAMVT0.MOUVEMMON
        End If
    End If
    
    rsSab.MoveNext
Loop




If Not blnCompte Then
    currentRow = currentRow + 1
    Call prtSAB_Balance_L_RELEVE_FOTC("R1", wSolde, wsExcel, currentRow)
Else
    If prtSAB_Balance.blnPrint_Relevé_Total_Mvt Then
         currentRow = currentRow + 1
         wsExcel.Cells(currentRow, 2) = meYBIACPT0.COMPTECOM
         wsExcel.Cells(currentRow, 3) = "cumul des mouvements (hors RPC)"
         wsExcel.Cells(currentRow, 3).HorizontalAlignment = xlHAlignRight
         wsExcel.Cells(currentRow, 3).Font.Size = 8
         wsExcel.Cells(currentRow, 3).Font.Name = "Arial"
         If wMvt_DB <> 0 Then
             wsExcel.Cells(currentRow, 5) = X
             wsExcel.Cells(currentRow, 5).HorizontalAlignment = xlHAlignRight
             wsExcel.Cells(currentRow, 5).Font.Size = 8
             wsExcel.Cells(currentRow, 5).Font.Name = "Arial"
         End If
         If wMvt_CR <> 0 Then
             X = Format$(Abs(wMvt_CR), "### ### ### ### ##0.00")
             wsExcel.Cells(currentRow, 6) = X
             wsExcel.Cells(currentRow, 6).HorizontalAlignment = xlHAlignRight
             wsExcel.Cells(currentRow, 6).Font.Size = 8
             wsExcel.Cells(currentRow, 6).Font.Name = "Arial"
         End If
         currentRow = currentRow + 1
         Call prtSAB_Balance_Montant_RELEVE_FOTC(wSolde, wsExcel, currentRow)
         wsExcel.Cells(currentRow, 7) = meYBIACPT0.COMPTEDEV
         wsExcel.Cells(currentRow, 7).Font.Color = vbBlue
         wsExcel.Cells(currentRow, 7).HorizontalAlignment = xlHAlignRight
         wsExcel.Cells(currentRow, 7).Font.Size = 8
         wsExcel.Cells(currentRow, 7).Font.Name = "Arial"
    End If
End If
End Sub

Public Sub prtSAB_Balance_R_xlsManual(ByRef currentRow As Long, wsExcel As Excel.Worksheet, ByRef comptageRows As Long, maxRows As Long, maxRowsPlus As Long)
Dim X As String
Dim xYBIAMVT0 As typeYBIAMVT0, wSolde As Currency, wMvt_DB As Currency, wMvt_CR As Currency
Dim xSQL As String
Dim wMvt_Nb As Long
wSolde = meYBIACPT0.SOLDECEN
blnCompte = False
rsYBIAMVT0_Init xYBIAMVT0

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIAMVTH " _
     & " where MOUVEMCOM = '" & meYBIACPT0.COMPTECOM & "'" _
     & " and MOUVEMDTR >= " & IbmAmjMin _
     & " and MOUVEMDTR <= " & IbmAmjMax _
     & " order by MOUVEMDTR,MOUVEMPIE,MOUVEMECR"
     
Set rsSab = cnsab.Execute(xSQL)
Do While Not rsSab.EOF
    Call rsYBIAMVT0_GetBuffer(rsSab, xYBIAMVT0)

    If Not blnCompte Then
        wMvt_DB = 0: wMvt_CR = 0: wMvt_Nb = 0
        blnCompte = True
        wSolde = xYBIAMVT0.BIAMVTSD0
        currentRow = currentRow + 1
        Call prtSAB_Balance_L_xlsManual("R0", wSolde, currentRow, wsExcel)
    End If
    
    Call prtSAB_Balance_Mvt_xlsManual(xYBIAMVT0, currentRow, wsExcel, comptageRows, maxRows, maxRowsPlus)
    wSolde = wSolde + xYBIAMVT0.MOUVEMMON
    If xYBIAMVT0.MOUVEMOPE <> "RPC" Then
        wMvt_Nb = wMvt_Nb + 1
        If xYBIAMVT0.MOUVEMMON > 0 Then
            wMvt_DB = wMvt_DB + xYBIAMVT0.MOUVEMMON
        Else
            wMvt_CR = wMvt_CR + xYBIAMVT0.MOUVEMMON
        End If
    End If
    
    rsSab.MoveNext
Loop

If Not blnCompte Then
    Call prtSAB_Balance_NewLine_xlsManual(currentRow, wsExcel, comptageRows, maxRows, maxRowsPlus)
    Range("A8:J8").Select
    Selection.Copy
    Range("A" & CStr(currentRow)).Select
    ActiveSheet.Paste
    Call prtSAB_Balance_NewLine_xlsManual(currentRow, wsExcel, comptageRows, maxRows, maxRowsPlus)
    Call prtSAB_Balance_L_xlsManual("R1", wSolde, currentRow, wsExcel)
Else
    If prtSAB_Balance.blnPrint_Relevé_Total_Mvt Then
         Call prtSAB_Balance_NewLine_xlsManual(currentRow, wsExcel, comptageRows, maxRows, maxRowsPlus)
         meYBIACPT0.COMPTEINT = "cumul des mouvements (hors RPC)"
         Call prtSAB_Balance_Mvt_xlsManual(xYBIAMVT0, currentRow, wsExcel, comptageRows, maxRows, maxRowsPlus)
         If wMvt_DB <> 0 Then
             X = Format$(Abs(wMvt_DB), "### ### ### ### ##0.00")
             Call prtSAB_Balance_Montant_xlsManual(Abs(wMvt_DB), currentRow, wsExcel)
         End If
         If wMvt_CR <> 0 Then
             X = Format$(Abs(wMvt_CR), "### ### ### ### ##0.00")
             Call prtSAB_Balance_Montant_xlsManual(Abs(wMvt_CR), currentRow, wsExcel)
         End If
            Call prtSAB_Balance_NewLine_xlsManual(currentRow, wsExcel, comptageRows, maxRows, maxRowsPlus)
            Range("A8:J8").Select
            Selection.Copy
            Range("A" & CStr(currentRow)).Select
            ActiveSheet.Paste
         Call prtSAB_Balance_Montant_xlsManual(wSolde, currentRow, wsExcel)
    End If
End If
End Sub

Public Sub prtSAB_Liste_Monitor(lFct As String, lAMJMin As String, lAMJMax As String, fgW As MSFlexGrid, larrYBIACPT0() As typeYBIACPT0, larrYBIACPT0_Nb As Long, lMsg As String, lYSTOMON() As Currency, lDORCPTDMV() As Long)

Dim wIndex As Long, I As Integer
Dim mFct1 As String


mMsg = lMsg
IbmAmjMin = dateIBM(lAMJMin)
IbmAmjMax = dateIBM(lAMJMax)

meCV1.DeviseN = 0
meCV1.Montant = 0

mBalance_CV = "J  "
meCV1.OpéAmj = YBIATAB0_DATE_CPT_J
meCV2.OpéAmj = YBIATAB0_DATE_CPT_J



curCumul_Db = 0: curCumul_Cr = 0
curW_Db = 0: curW_Cr = 0
curCOMPTEDEV_Db_EUR = 0: curCOMPTEDEV_Cr_EUR = 0
rsYBIACPT0_Init prevYBIACPT0

If lFct = "FOTC_CHAPRO" Then
    prtTitleText = "Etat CHA|PRO (hors 'intérêts') JPY|GBP|USD au " & dateImp10(lAMJMax)
Else
    prtTitleText = "Liste au " & dateImp10(lAMJMax)
End If
prtFontName = prtFontName_Arial
prtSAB_Balance_Open
prtHeaderHeight = 300
prtSAB_Balance_Form
XPrt.CurrentY = XPrt.CurrentY - prtlineHeight

XPrt.FontSize = 8
For I = 1 To fgW.Rows - 1
    
    fgW.Row = I
    fgW.Col = fgW.Cols - 1: wIndex = Val(fgW.Text)
    meYBIACPT0 = larrYBIACPT0(wIndex)
    If meYBIACPT0.PLANCOPRO <> prevYBIACPT0.PLANCOPRO _
    Or meYBIACPT0.COMPTEDEV <> prevYBIACPT0.COMPTEDEV Then
        If I > 1 Then Call prtSAB_Liste_Total
    End If
    
    meCV1.DeviseIso = meYBIACPT0.COMPTEDEV
    prtSAB_Balance_L "L", meYBIACPT0.SOLDECEN
    If meYBIACPT0.SOLDECEN > 0 Then
        curW_Db = curW_Db + meYBIACPT0.SOLDECEN
        curCumul_Db = curCumul_Db + meYBIACPT0.SOLDECEN
        curCOMPTEDEV_Db_EUR = curCOMPTEDEV_Db_EUR + meCV2.Montant
    Else
        curW_Cr = curW_Cr + meYBIACPT0.SOLDECEN
        curCumul_Cr = curCumul_Cr + meYBIACPT0.SOLDECEN
        curCOMPTEDEV_Cr_EUR = curCOMPTEDEV_Cr_EUR + meCV2.Montant
    End If
    
     prevYBIACPT0 = meYBIACPT0
Next I

meYBIACPT0.COMPTEDEV = ""
Call prtSAB_Liste_Total
        

prtSAB_Balance_Close

End Sub

Public Sub prtSAB_Client_Stat(lFct As String, lAMJMin As String, lAMJMax As String, fgW As MSFlexGrid, larrYBIACPT0() As typeYBIACPT0, larrYBIACPT0_Nb As Long)
Dim wIndex As Long, I As Integer
Dim blnCumul As Boolean

prtTitleText = "Répartition de la clientèle par catégorie"
meCV1.DeviseN = 0
meCV1.Montant = 0

mBalance_CV = "J  "
meCV1.OpéAmj = YBIATAB0_DATE_CPT_J
meCV2.OpéAmj = YBIATAB0_DATE_CPT_J

prtFontName = prtFontName_Arial
prtSAB_Balance_Open
prtHeaderHeight = 300
XPrt.CurrentY = XPrt.CurrentY - prtlineHeight
prtSAB_Client_Stat_Form

prtSAB_Client_Stat_Z SAB_Client_Stat
prtSAB_Client_Stat_Z SAB_Client_Stat_Actif
prtSAB_Client_Stat_Z SAB_Client_Stat_Annulé
prtSAB_Client_Stat_Z SAB_Client_Stat_Produit

rsYBIACPT0_Init prevYBIACPT0
blnCumul = False

fgW.Row = 1
fgW.Col = fgW.Cols - 1: wIndex = Val(fgW.Text)
prevYBIACPT0 = larrYBIACPT0(wIndex)
prtSAB_Client_Stat_PLANCOPRO prevYBIACPT0.PLANCOPRO

XPrt.FontSize = 8
For I = 1 To fgW.Rows - 1
    
    fgW.Row = I
    fgW.Col = fgW.Cols - 1: wIndex = Val(fgW.Text)
    meYBIACPT0 = larrYBIACPT0(wIndex)
    
    If meYBIACPT0.PLANCOPRO <> prevYBIACPT0.PLANCOPRO Then
        prtSAB_Client_Stat_Cumul
        prtSAB_Client_Stat_Line
        prtSAB_Client_Stat_Produit
        prtSAB_Client_Stat_PLANCOPRO meYBIACPT0.PLANCOPRO

    Else
         If meYBIACPT0.CLIENACAT <> prevYBIACPT0.CLIENACAT Then
             prtSAB_Client_Stat_Cumul
             prtSAB_Client_Stat_Line
         Else
             If meYBIACPT0.CLIENACLI <> prevYBIACPT0.CLIENACLI Then
                 prtSAB_Client_Stat_Cumul
             End If
        End If
    End If
    
    If meYBIACPT0.COMPTEFON = "4" Then
        SAB_Client_Stat.Nb_Compte_Annulé = SAB_Client_Stat.Nb_Compte_Annulé + 1
    Else
        SAB_Client_Stat.Nb_Compte = SAB_Client_Stat.Nb_Compte + 1
    End If
    meCV1.DeviseIso = meYBIACPT0.COMPTEDEV
    prtSAB_Balance_CV meYBIACPT0.SOLDECEN
    If meCV2.Montant > 0 Then
        SAB_Client_Stat.Solde_DB = SAB_Client_Stat.Solde_DB + meCV2.Montant
    Else
        SAB_Client_Stat.Solde_CR = SAB_Client_Stat.Solde_CR - meCV2.Montant
    End If
    prevYBIACPT0 = meYBIACPT0
Next I

prtSAB_Client_Stat_Cumul
prtSAB_Client_Stat_Line
prtSAB_Client_Stat_Produit

prtSAB_Client_Stat_Form_End

prtSAB_Balance_Close

End Sub


'---------------------------------------------------------
Public Sub prtSAB_Balance_Form()
'---------------------------------------------------------
Dim X As String

XPrt.DrawWidth = 1
XPrt.FontSize = 7: XPrt.FontBold = True

Call frmElpPrt.prtTrame(prtMinX, prtMinY, prtMaxX, prtMinY + prtHeaderHeight, "B", 240)
XPrt.Line (prtMinX + 12000, prtMinY)-(prtMinX + 12000, prtMaxY), prtLineColor
XPrt.Line (prtMinX + 7500, prtMinY)-(prtMinX + 7500, prtMaxY), prtLineColor
XPrt.Line (prtMaxX, prtMinY)-(prtMaxX, prtMaxY), prtLineColor
'---------------------------------------------------------
XPrt.CurrentY = prtMinY + 50

XPrt.CurrentX = prtMinX + 100: XPrt.Print " ";
XPrt.CurrentX = prtMinX + 400: XPrt.Print "Compte ";
XPrt.CurrentX = prtMinX + 2000: XPrt.Print "Intitulé";
If blnChk_BalanceStock Then
    XPrt.CurrentX = prtMinX + 5700: XPrt.Print "Stock opé";
    XPrt.CurrentX = prtMinX + 6500: XPrt.Print "Der Mvt";
    XPrt.CurrentX = prtMinX + 12100: XPrt.Print "Contrôle";
End If

'XPrt.CurrentX = prtMinX + 10500: XPrt.Print "Devise";
XPrt.CurrentX = prtMinX + 9000: XPrt.Print "Débit";
XPrt.CurrentX = prtMinX + 10900: XPrt.Print "Crédit";
XPrt.CurrentX = prtMinX + 13100: XPrt.Print "Débit";
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 6
XPrt.FontBold = True
XPrt.CurrentX = prtMinX + 14100: XPrt.Print "/ EUR /";
XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 8

XPrt.CurrentX = prtMinX + 15100: XPrt.Print "Crédit";

XPrt.CurrentY = prtMinY + prtHeaderHeight + 100
XPrt.FontBold = False

End Sub


Public Sub prtSAB_Balance_Close()
On Error GoTo prtError

If blnChk_BalanceStock Then
    XPrt.FontBold = True
    XPrt.FontSize = 6
    
    prtSAB_Balance_NewLine
    prtSAB_Balance_NewLine
    XPrt.ForeColor = vbMagenta
    XPrt.CurrentX = prtMinX + 12100: XPrt.Print "=";
    XPrt.CurrentX = prtMinX + 12400: XPrt.Print "Solde balance = cumul des contrats";
    prtSAB_Balance_NewLine
    XPrt.CurrentX = prtMinX + 12100: XPrt.Print "##";
    XPrt.CurrentX = prtMinX + 12400: XPrt.Print "Solde balance <> cumul des contrats ";
    prtSAB_Balance_NewLine
    XPrt.CurrentX = prtMinX + 12100: XPrt.Print "??";
    XPrt.CurrentX = prtMinX + 12400: XPrt.Print "Solde balance, aucun contrat rattaché??";
    prtSAB_Balance_NewLine
    XPrt.CurrentX = prtMinX + 2000: XPrt.Print "CAV et LOR : date de dernier mouvement hors échelles, facturation ...(+) indique qu'il y a des mouvements postérieurs initiés par la banque";
    XPrt.CurrentX = prtMinX + 12100: XPrt.Print "6M";
    XPrt.CurrentX = prtMinX + 12400: XPrt.Print "Compte ouvert depuis 6 mois ";
    XPrt.DrawWidth = 1: XPrt.Line (prtMinX + 12300, prtMinY + prtHeaderHeight)-(prtMinX + 12300, prtMaxY), prtLineColor
    XPrt.ForeColor = prtForeColor
End If

Call frmElpPrt.prtEndDoc(1000)
frmElpPrt.Hide
Exit Sub
'---------------------------------------------------------
prtError:
'---------------------------------------------------------

Call MsgBox(Err & " : " & Error(Err), vbCritical, "Impression")
frmElpPrt.Hide

End Sub


Public Sub prtSAB_Balance_Open()
On Error GoTo prtError

Set XPrt = Printer
frmElpPrt.Show vbModeless

Height8_6 = frmElpPrt.prtHeightDelta(8, 6)

blnFiligrane = False
prtOrientation = vbPRORLandscape '
prtPgmName = "prtSAB_Balance"
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





Public Sub prtSAB_Balance_R()
Dim X As String
Dim xYBIAMVT0 As typeYBIAMVT0, wSolde As Currency, wMvt_DB As Currency, wMvt_CR As Currency
Dim xSQL As String
Dim wMvt_Nb As Long
wSolde = meYBIACPT0.SOLDECEN
blnCompte = False
rsYBIAMVT0_Init xYBIAMVT0


xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIAMVTH " _
     & " where MOUVEMCOM = '" & meYBIACPT0.COMPTECOM & "'" _
     & " and MOUVEMDTR >= " & IbmAmjMin _
     & " and MOUVEMDTR <= " & IbmAmjMax _
     & " order by MOUVEMDTR,MOUVEMPIE,MOUVEMECR"
     
Set rsSab = cnsab.Execute(xSQL)
Do While Not rsSab.EOF
    Call rsYBIAMVT0_GetBuffer(rsSab, xYBIAMVT0)

    If Not blnCompte Then
        wMvt_DB = 0: wMvt_CR = 0: wMvt_Nb = 0
        blnCompte = True
        wSolde = xYBIAMVT0.BIAMVTSD0
        XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 0.5
        prtSAB_Balance_L "R0", wSolde
    End If
    
    prtSAB_Balance_Mvt xYBIAMVT0
    wSolde = wSolde + xYBIAMVT0.MOUVEMMON
    If xYBIAMVT0.MOUVEMOPE <> "RPC" Then
        wMvt_Nb = wMvt_Nb + 1
        If xYBIAMVT0.MOUVEMMON > 0 Then
            wMvt_DB = wMvt_DB + xYBIAMVT0.MOUVEMMON
        Else
            wMvt_CR = wMvt_CR + xYBIAMVT0.MOUVEMMON
        End If
    End If
    
    rsSab.MoveNext
Loop




If Not blnCompte Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 0.5
    prtSAB_Balance_L "R1", wSolde
Else
    If prtSAB_Balance.blnPrint_Relevé_Total_Mvt Then
         prtSAB_Balance_NewLine
         XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 6
         XPrt.FontItalic = True
         XPrt.ForeColor = vbMagenta
         XPrt.CurrentX = prtMinX + 400: XPrt.Print meYBIACPT0.COMPTECOM;
         XPrt.CurrentX = prtMinX + 2000: XPrt.Print "cumul des mouvements (hors RPC)";
         If wMvt_DB <> 0 Then
             X = Format$(Abs(wMvt_DB), "### ### ### ### ##0.00")
             XPrt.CurrentX = prtMinX + 9400 - XPrt.TextWidth(X)
             XPrt.Print X;
         End If
         If wMvt_CR <> 0 Then
             X = Format$(Abs(wMvt_CR), "### ### ### ### ##0.00")
             XPrt.CurrentX = prtMinX + 11400 - XPrt.TextWidth(X)
             XPrt.Print X;
         End If
         XPrt.FontItalic = False
         prtSAB_Balance_NewLine
         prtY = XPrt.CurrentY + prtlineHeight
         Call frmElpPrt.prtTrame(prtMinX + 7520, XPrt.CurrentY, prtMinX + 11980, prtY - 50, " ", 240)
        ' Call frmElpPrt.prtTrame(prtMinX + 12020, XPrt.CurrentY, prtMaxX - 20, prtY - 50, " ", 240)
          XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 6
         XPrt.ForeColor = prtForeColor_Header
         XPrt.FontBold = True
         prtSAB_Balance_Montant wSolde
         XPrt.FontBold = False
         XPrt.CurrentX = prtMinX + 11600: XPrt.Print meYBIACPT0.COMPTEDEV;
         XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 8
    End If
End If
'XPrt.Line (prtMinX + 7500, prtY)-(prtMinX + 12000, prtY), prtLineColor
'XPrt.CurrentY = XPrt.CurrentY - prtlineHeight

End Sub
Public Sub prtSAB_Balance_C()
Dim xYBIAMVT0 As typeYBIAMVT0, wSolde As Currency
Dim curDB As Currency, curCR As Currency
Dim xDate As String
Dim xSQL As String

If blnRésidence Then
    If Mid$(meYBIACPT0.COMPTECOM, 10, 1) <> mRésidence Then
        prtSAB_Balance_C_Cumul
        mRésidence = Mid$(meYBIACPT0.COMPTECOM, 10, 1)
        blnCompte = False:    curCumul_Db = 0: curCumul_Cr = 0
    End If
End If

curDB = 0: curCR = 0
wSolde = meYBIACPT0.SOLDECEN
rsYBIAMVT0_Init xYBIAMVT0
If blnMOUVEMDCO Then
    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIAMVTH " _
         & " where MOUVEMCOM = '" & meYBIACPT0.COMPTECOM & "'" _
         & " and MOUVEMDCO >= " & IbmAmjMin _
         & " and MOUVEMDCO <= " & IbmAmjMax _
         & " order by MOUVEMDTR,MOUVEMPIE,MOUVEMECR"
Else
    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIAMVTH " _
         & " where MOUVEMCOM = '" & meYBIACPT0.COMPTECOM & "'" _
         & " and MOUVEMDTR >= " & IbmAmjMin _
         & " and MOUVEMDTR <= " & IbmAmjMax _
         & " order by MOUVEMDTR,MOUVEMPIE,MOUVEMECR"
End If

     
Set rsSab = cnsab.Execute(xSQL)
Do While Not rsSab.EOF
    Call rsYBIAMVT0_GetBuffer(rsSab, xYBIAMVT0)

    If xYBIAMVT0.MOUVEMMON < 0 Then
        curCR = curCR + xYBIAMVT0.MOUVEMMON
    Else
        curDB = curDB + xYBIAMVT0.MOUVEMMON
    End If

    rsSab.MoveNext
Loop

        
If blnSoldeZ Or curDB <> 0 Or curCR <> 0 Then
'''If curDB = 0 And curCR = 0 Then
    blnCompte = True
    curCumul_Db = curCumul_Db + curDB
    curCumul_Cr = curCumul_Cr + curCR
    
    prtSAB_Balance_L "C1", curCR
    If curDB <> 0 Then prtSAB_Balance_Montant curDB
End If

End Sub


Public Sub prtSAB_Balance_L(lFct As String, lcurX As Currency)

Dim wForecolor As Long
prtSAB_Balance_NewLine
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 6

If lcurX > 0 Then
    XPrt.ForeColor = vbRed
Else
    XPrt.ForeColor = prtForeColor
End If

Select Case lFct
    Case "R0"
     '   Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY, prtMinX + 7480, XPrt.CurrentY + prtlineHeight - 50, " ", 240)
    '    Call frmElpPrt.prtTrame(prtMinX + 7520, XPrt.CurrentY, prtMinX + 11980, XPrt.CurrentY- 50, " ", 240)
    '    Call frmElpPrt.prtTrame(prtMinX + 12020, XPrt.CurrentY , prtMaxX - 20, XPrt.CurrentY- 50, " ", 240)
        XPrt.FontBold = True: XPrt.ForeColor = prtForeColor_Header

        prtSAB_Balance_Montant lcurX
    Case "R1"
       prtY = XPrt.CurrentY + prtlineHeight
       ' Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY, prtMinX + 7480, prtY - 50, " ", 240)
        Call frmElpPrt.prtTrame(prtMinX + 7520, XPrt.CurrentY, prtMinX + 11980, prtY - 50, " ", 240)
       ' Call frmElpPrt.prtTrame(prtMinX + 12020, XPrt.CurrentY, prtMaxX - 20, prtY - 50, " ", 240)
        XPrt.FontBold = True
        XPrt.FontBold = True: XPrt.ForeColor = prtForeColor_Header
        prtSAB_Balance_Montant lcurX
        XPrt.FontBold = False
    Case "C1"
        prtSAB_Balance_Montant lcurX
    Case Else
        prtSAB_Balance_Montant lcurX
End Select

Select Case meYBIACPT0.COMPTEFON
    Case 0: If Mid$(lFct, 1, 1) <> "R" Then XPrt.ForeColor = prtForeColor
    Case 4: XPrt.ForeColor = vbRed: XPrt.CurrentX = prtMinX + 7600: XPrt.Print dateIBM10(meYBIACPT0.COMPTECLO, True);
    Case Else: XPrt.ForeColor = vbMagenta
End Select

XPrt.CurrentX = prtMinX + 100: XPrt.Print meYBIACPT0.PLANCOPRO;
If lFct = "B" Then
    XPrt.FontBold = False
Else
    XPrt.FontBold = True
End If

''Dim mRib_IbanE As String, mRib_Clé As String
''mRib_Clé = Format$(RibClé(strSocBdfE, strSocBdfG, Trim(meYBIACPT0.COMPTECOM), mRib_IbanE), "00")
    
XPrt.CurrentX = prtMinX + 400: XPrt.Print meYBIACPT0.COMPTECOM;  '''"& "    " & mRib_Clé;

XPrt.CurrentX = prtMinX + 2000: XPrt.Print meYBIACPT0.COMPTEINT;
XPrt.FontBold = False
If lFct = "L" Then
'_____________________________
           If meDORCPTDMV > 0 Then
                XPrt.CurrentX = prtMinX + 6800: XPrt.Print dateIBM10(meDORCPTDMV, True);
                If meDORCPTDMV <> meYBIACPT0.SOLDEDMO Then XPrt.Print " +";
            Else
                XPrt.CurrentX = prtMinX + 6800: XPrt.Print dateIBM10(meYBIACPT0.SOLDEDMO, True);
            End If
'____________________

'    XPrt.CurrentX = prtMinX + 6800: XPrt.Print dateIBM10(meYBIACPT0.SOLDEDMO, True);
End If

If Mid$(lFct, 1, 1) <> "R" Then XPrt.ForeColor = prtForeColor
XPrt.CurrentX = prtMinX + 11600: XPrt.Print meYBIACPT0.COMPTEDEV;

XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 8
XPrt.ForeColor = prtForeColor
End Sub
Public Sub prtSAB_Balance_L_csvManual(lcurXDB As Currency, lcurxcr As Currency)
Dim maligne As String

maligne = meYBIACPT0.PLANCOPRO & " " & meYBIACPT0.COMPTECOM & ";" & meYBIACPT0.COMPTEINT & ";"
Call prtSAB_Balance_Montant_csvManual(lcurXDB, lcurxcr, meYBIACPT0.COMPTEDEV, maligne)
maligne = maligne & Mid$(meYBIACPT0.COMPTECOM, 10, 1) 'code résidence
Print #csvFic, maligne
maligne = ""

End Sub

Public Sub prtSAB_Balance_B_Total_Prt(lK As Long, lSAB_BALANCE As typeSAB_BALANCE)
Dim curX As Currency, X As String

prtSAB_Balance_NewLine
If lSAB_BALANCE.iPrint_Trame <> 255 Then
    XPrt.CurrentY = XPrt.CurrentY - 20
        prtFillColor = RGB(240, 240, 220)
        Call frmElpPrt.prtTrame_Color(prtMinX + 6800, XPrt.CurrentY, prtMinX + 7500 - 20, XPrt.CurrentY + prtlineHeight, " ")
        Call frmElpPrt.prtTrame_Color(prtMinX + 7520, XPrt.CurrentY, prtMinX + 12000 - 20, XPrt.CurrentY + prtlineHeight, " ")
        Call frmElpPrt.prtTrame_Color(prtMinX + 12020, XPrt.CurrentY, prtMaxX - 20, XPrt.CurrentY + prtlineHeight, " ")
        prtFillColor = prtFillColor_Standard

    'Call frmElpPrt.prtTrame(prtMinX + 6800, XPrt.CurrentY, prtMinX + 7500 - 20, XPrt.CurrentY + prtlineHeight, " ", lSAB_BALANCE.iPrint_Trame)
    'Call frmElpPrt.prtTrame(prtMinX + 7520, XPrt.CurrentY, prtMinX + 12000 - 20, XPrt.CurrentY + prtlineHeight, " ", lSAB_BALANCE.iPrint_Trame)
    'Call frmElpPrt.prtTrame(prtMinX + 12020, XPrt.CurrentY, prtMaxX - 20, XPrt.CurrentY + prtlineHeight, " ", lSAB_BALANCE.iPrint_Trame)
    XPrt.CurrentY = XPrt.CurrentY + 20
End If

If lSAB_BALANCE.blnPrint_Line Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight - 30
    XPrt.Line (prtMinX + 6800, XPrt.CurrentY)-(prtMaxX - 20, XPrt.CurrentY), prtLineColor
    XPrt.CurrentY = XPrt.CurrentY - prtlineHeight + 30
End If

XPrt.FontSize = 6
XPrt.FontBold = lSAB_BALANCE.blnPrint_Fontbold

If lK = 6 Then
    XPrt.CurrentX = prtMinX + 400: XPrt.Print lSAB_BALANCE.COMPTEOBL;
    If blnBalance_Pays Then
        XPrt.FontSize = 6: XPrt.FontBold = False
        XPrt.CurrentX = prtMinX - 180: XPrt.Print lSAB_BALANCE.CLIENARSD;
    End If

Else
    XPrt.CurrentX = prtMinX + 7000: XPrt.Print lSAB_BALANCE.COMPTEOBL;
End If



XPrt.CurrentX = prtMinX + 2000: XPrt.Print lSAB_BALANCE.COMPTEINT;
XPrt.CurrentX = prtMinX + 11600: XPrt.Print lSAB_BALANCE.COMPTEDEV;

curX = Abs(lSAB_BALANCE.DB)
If curX <> 0 Then
    X = Format$(curX, "### ### ### ### ##0.00")
    XPrt.CurrentX = prtMinX + 9400 - XPrt.TextWidth(X)
    XPrt.Print X;
End If

curX = Abs(lSAB_BALANCE.CR)
If curX <> 0 Then
    X = Format$(curX, "### ### ### ### ##0.00")
    XPrt.CurrentX = prtMinX + 11400 - XPrt.TextWidth(X)
    XPrt.Print X;
End If

curX = Abs(lSAB_BALANCE.DB_EUR)
If curX <> 0 Then
    X = Format$(curX, "### ### ### ### ##0.00")
    XPrt.CurrentX = prtMaxX - 2100 - XPrt.TextWidth(X)
    XPrt.Print X;
End If

curX = Abs(lSAB_BALANCE.CR_EUR)
If curX <> 0 Then
    X = Format$(curX, "### ### ### ### ##0.00")
    XPrt.CurrentX = prtMaxX - 100 - XPrt.TextWidth(X)
    XPrt.Print X;
End If



If XPrt.FontSize = 6 Then XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 8

End Sub

Public Sub prtSAB_Balance_B_Dev_Classe_Prt()
Dim curX As Currency, X As String
Dim curDB As Currency, curCR As Currency
Dim blnSaut As Boolean, blnErreur As Boolean
Dim wNb As Integer

blnErreur = False
wNb = arrDev_Nb + 2
arrDEV_C1A5(wNb) = zSAB_BALANCE
arrDEV_C6A8(wNb) = zSAB_BALANCE
arrDEV_C9(wNb) = zSAB_BALANCE

arrDEV_C1A5(arrDev_Nb + 1).COMPTEDEV = "???"
arrDEV_C6A8(arrDev_Nb + 1).COMPTEDEV = "???"
arrDEV_C9(arrDev_Nb + 1).COMPTEDEV = "???"
arrDEV_C1A5(arrDev_Nb + 2).COMPTEDEV = "***"
arrDEV_C6A8(arrDev_Nb + 2).COMPTEDEV = "***"
arrDEV_C9(arrDev_Nb + 2).COMPTEDEV = "***"

prtTitleText = "Récapitulatif Bilan / Hors Bilan " & xTitleText
frmElpPrt.prtNewPage
prtSAB_Balance_Form
XPrt.FontSize = 6

For arrDev_K = 1 To wNb
    blnSaut = False
    If arrDEV_C1A5(arrDev_K).DB_EUR <> 0 Or arrDEV_C1A5(arrDev_K).CR_EUR <> 0 Then
        XPrt.FontBold = True
        blnSaut = True
        prtSAB_Balance_NewLine
        XPrt.CurrentX = prtMinX + 5000: XPrt.Print arrDEV_C1A5(arrDev_K).COMPTEDEV & " - Total classes 1 à 5 ";
        XPrt.CurrentX = prtMinX + 11600: XPrt.Print arrDEV_C1A5(arrDev_K).COMPTEDEV;
    
        curX = Abs(arrDEV_C1A5(arrDev_K).DB)
        If curX <> 0 Then
            X = Format$(curX, "### ### ### ### ##0.00")
            XPrt.CurrentX = prtMinX + 9400 - XPrt.TextWidth(X)
            XPrt.Print X;
        End If
        
        curX = Abs(arrDEV_C1A5(arrDev_K).CR)
        If curX <> 0 Then
            X = Format$(curX, "### ### ### ### ##0.00")
            XPrt.CurrentX = prtMinX + 11400 - XPrt.TextWidth(X)
            XPrt.Print X;
        End If
        
        curX = Abs(arrDEV_C1A5(arrDev_K).DB_EUR)
        If curX <> 0 Then
            X = Format$(curX, "### ### ### ### ##0.00")
            XPrt.CurrentX = prtMaxX - 2100 - XPrt.TextWidth(X)
            XPrt.Print X;
        End If
        
        curX = Abs(arrDEV_C1A5(arrDev_K).CR_EUR)
        If curX <> 0 Then
            X = Format$(curX, "### ### ### ### ##0.00")
            XPrt.CurrentX = prtMaxX - 100 - XPrt.TextWidth(X)
            XPrt.Print X;
        End If
    End If

    If arrDEV_C6A8(arrDev_K).DB_EUR <> 0 Or arrDEV_C6A8(arrDev_K).CR_EUR <> 0 Then
        XPrt.FontBold = True
        blnSaut = True
        prtSAB_Balance_NewLine
        XPrt.CurrentX = prtMinX + 5000: XPrt.Print arrDEV_C6A8(arrDev_K).COMPTEDEV & " - Total classes 6 à 8 ";
        XPrt.CurrentX = prtMinX + 11600: XPrt.Print arrDEV_C6A8(arrDev_K).COMPTEDEV;
    
        curX = Abs(arrDEV_C6A8(arrDev_K).DB)
        If curX <> 0 Then
            X = Format$(curX, "### ### ### ### ##0.00")
            XPrt.CurrentX = prtMinX + 9400 - XPrt.TextWidth(X)
            XPrt.Print X;
        End If
        
        curX = Abs(arrDEV_C6A8(arrDev_K).CR)
        If curX <> 0 Then
            X = Format$(curX, "### ### ### ### ##0.00")
            XPrt.CurrentX = prtMinX + 11400 - XPrt.TextWidth(X)
            XPrt.Print X;
        End If
        
        curX = Abs(arrDEV_C6A8(arrDev_K).DB_EUR)
        If curX <> 0 Then
            X = Format$(curX, "### ### ### ### ##0.00")
            XPrt.CurrentX = prtMaxX - 2100 - XPrt.TextWidth(X)
            XPrt.Print X;
        End If
        
        curX = Abs(arrDEV_C6A8(arrDev_K).CR_EUR)
        If curX <> 0 Then
            X = Format$(curX, "### ### ### ### ##0.00")
            XPrt.CurrentX = prtMaxX - 100 - XPrt.TextWidth(X)
            XPrt.Print X;
        End If
    End If
    If arrDEV_C9(arrDev_K).DB_EUR <> 0 Or arrDEV_C9(arrDev_K).CR_EUR <> 0 Then
        XPrt.FontBold = False
        blnSaut = True
        prtSAB_Balance_NewLine
        XPrt.CurrentX = prtMinX + 5000: XPrt.Print arrDEV_C9(arrDev_K).COMPTEDEV & " - Total classe 9 ";
        XPrt.CurrentX = prtMinX + 11600: XPrt.Print arrDEV_C9(arrDev_K).COMPTEDEV;
    
        curX = Abs(arrDEV_C9(arrDev_K).DB)
        If curX <> 0 Then
            X = Format$(curX, "### ### ### ### ##0.00")
            XPrt.CurrentX = prtMinX + 9400 - XPrt.TextWidth(X)
            XPrt.Print X;
        End If
        
        curX = Abs(arrDEV_C9(arrDev_K).CR)
        If curX <> 0 Then
            X = Format$(curX, "### ### ### ### ##0.00")
            XPrt.CurrentX = prtMinX + 11400 - XPrt.TextWidth(X)
            XPrt.Print X;
        End If
        
        curX = Abs(arrDEV_C9(arrDev_K).DB_EUR)
        If curX <> 0 Then
            X = Format$(curX, "### ### ### ### ##0.00")
            XPrt.CurrentX = prtMaxX - 2100 - XPrt.TextWidth(X)
            XPrt.Print X;
        End If
        
        curX = Abs(arrDEV_C9(arrDev_K).CR_EUR)
        If curX <> 0 Then
            X = Format$(curX, "### ### ### ### ##0.00")
            XPrt.CurrentX = prtMaxX - 100 - XPrt.TextWidth(X)
            XPrt.Print X;
        End If
    End If
    
    curX = arrDEV_C1A5(arrDev_K).DB + arrDEV_C1A5(arrDev_K).CR _
            + arrDEV_C6A8(arrDev_K).DB + arrDEV_C6A8(arrDev_K).CR
    If curX <> 0 And blnChk_BalanceEquilibrée Then
        blnErreur = True
        prtSAB_Balance_NewLine
        XPrt.ForeColor = vbMagenta
        XPrt.CurrentX = prtMinX + 2000: XPrt.Print arrDEV_C1A5(arrDev_K).COMPTEDEV & "   ?????????? ERREUR BILAN ";
        X = Format$(curX, "### ### ### ### ##0.00")
        XPrt.CurrentX = prtMinX + 1900 - XPrt.TextWidth(X)
        XPrt.Print X;
        XPrt.ForeColor = vbBlack

    End If
    
    
    curX = arrDEV_C9(arrDev_K).DB + arrDEV_C9(arrDev_K).CR
    If curX <> 0 And blnChk_BalanceEquilibrée Then
        blnErreur = True
        prtSAB_Balance_NewLine
        XPrt.ForeColor = vbMagenta
        XPrt.CurrentX = prtMinX + 2000: XPrt.Print arrDEV_C9(arrDev_K).COMPTEDEV & "   ?????????? ERREUR HORS-BILAN ";
        X = Format$(curX, "### ### ### ### ##0.00")
        XPrt.CurrentX = prtMinX + 1900 - XPrt.TextWidth(X)
        XPrt.Print X;
        XPrt.ForeColor = vbBlack
    End If
   If arrDev_K < wNb Then
        arrDEV_C1A5(wNb).DB_EUR = arrDEV_C1A5(wNb).DB_EUR + arrDEV_C1A5(arrDev_K).DB_EUR
        arrDEV_C1A5(wNb).CR_EUR = arrDEV_C1A5(wNb).CR_EUR + arrDEV_C1A5(arrDev_K).CR_EUR
        arrDEV_C6A8(wNb).DB_EUR = arrDEV_C6A8(wNb).DB_EUR + arrDEV_C6A8(arrDev_K).DB_EUR
        arrDEV_C6A8(wNb).CR_EUR = arrDEV_C6A8(wNb).CR_EUR + arrDEV_C6A8(arrDev_K).CR_EUR
        arrDEV_C9(wNb).DB_EUR = arrDEV_C9(wNb).DB_EUR + arrDEV_C9(arrDev_K).DB_EUR
        arrDEV_C9(wNb).CR_EUR = arrDEV_C9(wNb).CR_EUR + arrDEV_C9(arrDev_K).CR_EUR
   End If
   
   If blnSaut Then prtSAB_Balance_NewLine

Next arrDev_K

XPrt.FontBold = True
blnSaut = True
prtSAB_Balance_NewLine
XPrt.CurrentX = prtMinX + 5000: XPrt.Print arrDEV_C9(wNb).COMPTEDEV & " - Total classes 1-8 ";

curDB = Abs(arrDEV_C1A5(wNb).DB_EUR + arrDEV_C6A8(wNb).DB_EUR)
If curDB <> 0 Then
    X = Format$(curDB, "### ### ### ### ##0.00")
    XPrt.CurrentX = prtMaxX - 2100 - XPrt.TextWidth(X)
    XPrt.Print X;
End If

curCR = Abs(arrDEV_C1A5(wNb).CR_EUR + arrDEV_C6A8(wNb).CR_EUR)
If curCR <> 0 Then
    X = Format$(curCR, "### ### ### ### ##0.00")
    XPrt.CurrentX = prtMaxX - 100 - XPrt.TextWidth(X)
    XPrt.Print X;
End If

If blnChk_BalanceEquilibrée Then
    Call frmElpPrt.prtTrame(prtMinX + 7500 + 20, XPrt.CurrentY, prtMinX + 12000 - 20, XPrt.CurrentY + prtlineHeight, " ", 240)
    If blnErreur Then
        XPrt.ForeColor = vbMagenta
        X = "??? ERREUR ???"
    Else
        X = "Balance équilibrée"
    End If
    XPrt.FontSize = 12
    frmElpPrt.prtCentré prtMinX + 9750, X
    XPrt.ForeColor = vbBlack
End If

End Sub


Public Sub prtSAB_Balance_Mvt(lYBIAMVT0 As typeYBIAMVT0)
prtSAB_Balance_NewLine

XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 6
XPrt.FontItalic = True
 prtSAB_Balance_Montant lYBIAMVT0.MOUVEMMON
XPrt.CurrentX = prtMinX + 400: XPrt.Print lYBIAMVT0.MOUVEMOPE & " " & lYBIAMVT0.MOUVEMNUM&; " " & lYBIAMVT0.MOUVEMEVE;
XPrt.CurrentX = prtMinX + 2000: XPrt.Print Trim(lYBIAMVT0.LIBELLIB1) & " " & Trim(lYBIAMVT0.LIBELLIB2) & " " & Trim(lYBIAMVT0.LIBELLIB3);
XPrt.CurrentX = prtMinX + 6800: XPrt.Print dateIBM10(lYBIAMVT0.MOUVEMDTR, True);
XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 8
XPrt.FontItalic = False
End Sub

Public Sub prtSAB_Balance_Montant(lcurX As Currency)
Dim X As String, mColor As Long

mColor = XPrt.ForeColor
prtSAB_Balance_CV lcurX

X = Format$(Abs(lcurX), "### ### ### ### ##0.00")
If lcurX > 0 Then
    XPrt.CurrentX = prtMinX + 9400 - XPrt.TextWidth(X)
    XPrt.ForeColor = vbRed
Else
    XPrt.CurrentX = prtMinX + 11400 - XPrt.TextWidth(X)
    XPrt.ForeColor = vbBlue
End If
XPrt.Print X;


X = Format$(Abs(meCV2.Montant), "### ### ### ### ##0.00")
If meCV2.Montant > 0 Then
    XPrt.CurrentX = prtMaxX - 2100 - XPrt.TextWidth(X)
    XPrt.ForeColor = vbRed
Else
    XPrt.CurrentX = prtMaxX - 100 - XPrt.TextWidth(X)
    XPrt.ForeColor = vbBlue
End If
XPrt.Print X;
XPrt.ForeColor = mColor

End Sub
Public Sub prtSAB_Balance_CV(lcurX As Currency)

meCV1.Montant = lcurX
If meCV1.DeviseIso <> "EUR" Then
    Call CV_Calc(mBalance_CV, meCV1, meCV2)
Else
    meCV2.Montant = lcurX
End If

End Sub

Public Sub prtSAB_Balance_Montant_Cumul(lcurDB As Currency, lcurCR As Currency)
Dim X As String
X = Format$(Abs(lcurDB), "### ### ### ### ##0.00")
XPrt.CurrentX = prtMaxX - 2100 - XPrt.TextWidth(X)
XPrt.Print X;
X = Format$(Abs(lcurCR), "### ### ### ### ##0.00")
XPrt.CurrentX = prtMaxX - 100 - XPrt.TextWidth(X)
XPrt.Print X;

End Sub


Public Sub prtSAB_Balance_NewLine()
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
If XPrt.CurrentY + 300 > prtMaxY Then
    If blnChk_BalanceStock Then XPrt.DrawWidth = 1: XPrt.Line (prtMinX + 12300, prtMinY + prtHeaderHeight)-(prtMinX + 12300, prtMaxY), prtLineColor
    frmElpPrt.prtNewPage
    prtSAB_Balance_Form
End If

End Sub

Public Sub prtSAb_Client_Stat_NewLine()
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
If XPrt.CurrentY + 300 > prtMaxY Then
    prtSAB_Client_Stat_Form_End
    frmElpPrt.prtNewPage
    prtSAB_Client_Stat_Form
End If

End Sub

Public Sub prtSAB_Balance_C_Cumul()
If blnCompte Then
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
    Call frmElpPrt.prtTrame(prtMinX + 400, XPrt.CurrentY - 50, 2000, XPrt.CurrentY + prtlineHeight, " ", 240)
    Call frmElpPrt.prtTrame(prtMinX + 12020, XPrt.CurrentY - 50, prtMaxX - 20, XPrt.CurrentY + prtlineHeight, " ", 240)
    XPrt.FontBold = True
    XPrt.CurrentX = prtMinX + 400
    XPrt.Print "Code résidence : " & mRésidence;
    prtSAB_Balance_Montant_Cumul curCumul_Db, curCumul_Cr
    XPrt.FontBold = False
End If

End Sub

Public Sub prtSAB_Balance_L_Rupture()

nbListe_Line = nbListe_Line + 1

If blnClient_Line Then
    prtSAB_Balance_NewLine
    XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 6: XPrt.FontBold = True

    XPrt.CurrentX = prtMinX + 6000: XPrt.Print prevYBIACPT0.CLIENACLI & " " & prevYBIACPT0.CLIENASIG;

    Call frmElpPrt.prtTrame(prtMinX + 12020, XPrt.CurrentY, prtMaxX - 20, XPrt.CurrentY + prtlineHeight - 50, " ", 240)
    If curClient_Db <> 0 Then
        XPrt.ForeColor = vbRed
        X = Format$(Abs(curClient_Db), "### ### ### ### ##0.00")
        XPrt.CurrentX = prtMaxX - 2100 - XPrt.TextWidth(X)
        XPrt.Print X;
    End If
    If curClient_Cr <> 0 Then
        XPrt.ForeColor = prtForeColor
        X = Format$(Abs(curClient_Cr), "### ### ### ### ##0.00")
        XPrt.CurrentX = prtMaxX - 100 - XPrt.TextWidth(X)
        XPrt.Print X;
    End If
    XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 8:: XPrt.FontBold = False

End If

curW_Db = curW_Db + curClient_Db
curW_Cr = curW_Cr + curClient_Cr

nbClient_Line = 0
curClient_Db = 0: curClient_Cr = 0

End Sub
Public Sub prtSAB_Balance_B_COMPTEDEV()

prtSAB_Balance_B_COMPTEOBL
If blnBalance_Pays Then curCOMPTEDEV_Db = 0: curCOMPTEDEV_Cr = 0

If blnBalance_B_Détail And nbCOMPTEDEV_Line > 0 Then

    prtSAB_Balance_NewLine
    XPrt.FontBold = True

    'Call frmElpPrt.prtTrame(prtMinX + 7520, XPrt.CurrentY, prtMinX + 12000 - 20, XPrt.CurrentY + prtlineHeight, " ", 240)
        prtFillColor = RGB(255, 255, 190)
        Call frmElpPrt.prtTrame_Color(prtMinX + 7520, XPrt.CurrentY, prtMinX + 12000 - 20, XPrt.CurrentY + prtlineHeight, " ")
        Call frmElpPrt.prtTrame_Color(prtMinX + 12020, XPrt.CurrentY, prtMaxX - 20, XPrt.CurrentY + prtlineHeight, " ")
        prtFillColor = prtFillColor_Standard
        'Call frmElpPrt.prtTrame_Color(prtMinX, XPrt.CurrentY, prtMaxX, XPrt.CurrentY + prtHeaderHeight, "B")
        'Call frmElpPrt.prtTrame(prtMinX + 12020, XPrt.CurrentY, prtMaxX - 20, XPrt.CurrentY + prtlineHeight, " ", 240)

    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight + Height8_6: XPrt.FontSize = 6
    XPrt.Line (prtMinX + 7500, XPrt.CurrentY)-(prtMaxX - 20, XPrt.CurrentY), prtLineColor
    XPrt.CurrentY = XPrt.CurrentY - prtlineHeight

    
    If blnBalance_B_COMPTEDEV Then
        XPrt.CurrentX = prtMinX + 11600: XPrt.Print prevYBIACPT0.COMPTEDEV;
        If curCOMPTEDEV_Db <> 0 Then
            X = Format$(Abs(curCOMPTEDEV_Db), "### ### ### ### ##0.00")
            XPrt.CurrentX = prtMinX + 9400 - XPrt.TextWidth(X)
            XPrt.Print X;
        End If
        If curCOMPTEDEV_Cr <> 0 Then
             X = Format$(Abs(curCOMPTEDEV_Cr), "### ### ### ### ##0.00")
            XPrt.CurrentX = prtMinX + 11400 - XPrt.TextWidth(X)
            XPrt.Print X;
        End If
    End If

    If curCOMPTEDEV_Db_EUR <> 0 Then
        X = Format$(Abs(curCOMPTEDEV_Db_EUR), "### ### ### ### ##0.00")
        XPrt.CurrentX = prtMaxX - 2100 - XPrt.TextWidth(X)
        XPrt.Print X;
    End If
    If curCOMPTEDEV_Cr_EUR <> 0 Then
        X = Format$(Abs(curCOMPTEDEV_Cr_EUR), "### ### ### ### ##0.00")
        XPrt.CurrentX = prtMaxX - 100 - XPrt.TextWidth(X)
        XPrt.Print X;
    End If

     XPrt.FontBold = False
    XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 8
End If

curW_Db = curW_Db + curCOMPTEDEV_Db
curW_Cr = curW_Cr + curCOMPTEDEV_Cr

curCOMPTEDEV_Db = 0: curCOMPTEDEV_Cr = 0
curCOMPTEDEV_Db_EUR = 0: curCOMPTEDEV_Cr_EUR = 0
nbCOMPTEDEV_Line = 0
End Sub


Public Sub prtSAB_Balance_B_COMPTEOBL()


Call rsZPLAN0_Read(prevYBIACPT0.COMPTEOBL, meZPLAN0)

If blnBalance_B_Détail And détailSAB_BALANCE.blnPrint And nbCOMPTEOBL_Line > 0 Then
    
    prtSAB_Balance_NewLine
    XPrt.FontBold = détailSAB_BALANCE.blnPrint_Fontbold
    
    If détailSAB_BALANCE.iPrint_Trame <> 255 Then
        prtFillColor = RGB(240, 240, 220)
        Call frmElpPrt.prtTrame_Color(prtMinX + 2000, XPrt.CurrentY, prtMinX + 7500 - 20, XPrt.CurrentY + prtlineHeight, " ")
        Call frmElpPrt.prtTrame_Color(prtMinX + 7520, XPrt.CurrentY, prtMinX + 12000 - 20, XPrt.CurrentY + prtlineHeight, " ")
        Call frmElpPrt.prtTrame_Color(prtMinX + 12020, XPrt.CurrentY, prtMaxX - 20, XPrt.CurrentY + prtlineHeight, " ")
        prtFillColor = prtFillColor_Standard

        'Call frmElpPrt.prtTrame(prtMinX + 2000, XPrt.CurrentY, prtMinX + 7500 - 20, XPrt.CurrentY + prtlineHeight, " ", détailSAB_BALANCE.iPrint_Trame)
        'Call frmElpPrt.prtTrame(prtMinX + 7520, XPrt.CurrentY, prtMinX + 12000 - 20, XPrt.CurrentY + prtlineHeight, " ", détailSAB_BALANCE.iPrint_Trame)
        'Call frmElpPrt.prtTrame(prtMinX + 12020, XPrt.CurrentY, prtMaxX - 20, XPrt.CurrentY + prtlineHeight, " ", détailSAB_BALANCE.iPrint_Trame)
    End If
    If détailSAB_BALANCE.blnPrint_Line Then
        XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
        XPrt.Line (prtMinX + 2000, XPrt.CurrentY)-(prtMaxX - 20, XPrt.CurrentY), prtLineColor
        XPrt.CurrentY = XPrt.CurrentY - prtlineHeight
    End If
    
    XPrt.FontSize = 6: XPrt.CurrentY = XPrt.CurrentY + Height8_6

    XPrt.CurrentX = prtMinX + 2000: XPrt.Print meZPLAN0.PLANINTIT;
    XPrt.CurrentX = prtMinX + 6500: XPrt.Print prevYBIACPT0.COMPTEDEV & " " & prevYBIACPT0.COMPTEOBL;
    

    If blnBalance_B_COMPTEDEV Then
        XPrt.CurrentX = prtMinX + 11600: XPrt.Print prevYBIACPT0.COMPTEDEV;
        If curCOMPTEOBL_Db <> 0 Then
            X = Format$(Abs(curCOMPTEOBL_Db), "### ### ### ### ##0.00")
            XPrt.CurrentX = prtMinX + 9400 - XPrt.TextWidth(X)
            XPrt.Print X;
        End If
        If curCOMPTEOBL_Cr <> 0 Then
             X = Format$(Abs(curCOMPTEOBL_Cr), "### ### ### ### ##0.00")
            XPrt.CurrentX = prtMinX + 11400 - XPrt.TextWidth(X)
            XPrt.Print X;
        End If
    End If

    If curCOMPTEOBL_Db_EUR <> 0 Then
        X = Format$(Abs(curCOMPTEOBL_Db_EUR), "### ### ### ### ##0.00")
        XPrt.CurrentX = prtMaxX - 2100 - XPrt.TextWidth(X)
        XPrt.Print X;
    End If
    If curCOMPTEOBL_Cr_EUR <> 0 Then
        X = Format$(Abs(curCOMPTEOBL_Cr_EUR), "### ### ### ### ##0.00")
        XPrt.CurrentX = prtMaxX - 100 - XPrt.TextWidth(X)
        XPrt.Print X;
    End If

     XPrt.FontBold = False
    XPrt.CurrentY = XPrt.CurrentY - Height8_6
End If

curW_Db = curW_Db + curCOMPTEOBL_Db
curW_Cr = curW_Cr + curCOMPTEOBL_Cr

If arrSAB_BALANCE_Nb = arrSAB_BALANCE_NbMax Then
    arrSAB_BALANCE_NbMax = arrSAB_BALANCE_NbMax + 100
    ReDim Preserve arrSAB_BALANCE(arrSAB_BALANCE_NbMax)
End If
arrSAB_BALANCE_Nb = arrSAB_BALANCE_Nb + 1
arrSAB_BALANCE(arrSAB_BALANCE_Nb) = totalSAB_BALANCE(6)
arrSAB_BALANCE(arrSAB_BALANCE_Nb).CLIENARSD = prevYBIACPT0.CLIENARSD
arrSAB_BALANCE(arrSAB_BALANCE_Nb).COMPTEDEV = prevYBIACPT0.COMPTEDEV
arrSAB_BALANCE(arrSAB_BALANCE_Nb).COMPTEOBL = prevYBIACPT0.COMPTEOBL
arrSAB_BALANCE(arrSAB_BALANCE_Nb).COMPTECOM = ""
arrSAB_BALANCE(arrSAB_BALANCE_Nb).COMPTEINT = meZPLAN0.PLANINTIT
arrSAB_BALANCE(arrSAB_BALANCE_Nb).DB = arrSAB_BALANCE(arrSAB_BALANCE_Nb).DB + curCOMPTEOBL_Db
arrSAB_BALANCE(arrSAB_BALANCE_Nb).CR = arrSAB_BALANCE(arrSAB_BALANCE_Nb).CR + curCOMPTEOBL_Cr
arrSAB_BALANCE(arrSAB_BALANCE_Nb).DB_EUR = arrSAB_BALANCE(arrSAB_BALANCE_Nb).DB_EUR + curCOMPTEOBL_Db_EUR
arrSAB_BALANCE(arrSAB_BALANCE_Nb).CR_EUR = arrSAB_BALANCE(arrSAB_BALANCE_Nb).CR_EUR + curCOMPTEOBL_Cr_EUR


curCOMPTEOBL_Db = 0: curCOMPTEOBL_Cr = 0
curCOMPTEOBL_Db_EUR = 0: curCOMPTEOBL_Cr_EUR = 0
nbCOMPTEOBL_Line = 0
End Sub

Public Sub prtSAB_Balance_L_Line()

If prevYBIACPT0.CLIENACLI <> meYBIACPT0.CLIENACLI Then
    If Trim(prevYBIACPT0.COMPTECOM) <> "" Then
        prtSAB_Balance_L_Rupture
    End If
End If

prtSAB_Balance_L "L", meYBIACPT0.SOLDECEN

nbClient_Line = nbClient_Line + 1

If meCV2.Montant > 0 Then
    curClient_Db = curClient_Db + meCV2.Montant
    curListe_Db = curListe_Db + meCV2.Montant
Else
    curClient_Cr = curClient_Cr + meCV2.Montant
    curListe_Cr = curListe_Cr + meCV2.Montant
End If

End Sub

Public Sub prtSAB_Balance_B_Line()
Dim curX As Currency, curX1 As Currency, curX2 As Currency
Dim X As String, XS As String, xText As String

Dim V
Dim blnOk As Boolean

If blnBalance_Pays Then                                  ' Tri Pays / PCi / Compte / Dev
    If prevYBIACPT0.CLIENARSD <> meYBIACPT0.CLIENARSD Then
        prtSAB_Balance_B_COMPTEOBL
        Call rsYBIATAB0_Read("SAB", "CLIENAPAY", "CLI" & meYBIACPT0.CLIENARSD, X)
        xPays = Trim(Mid$(X, 15, 30))
        prtTitleText = xPays & " - " & xTitleText
        If blnBalance_B_Détail Then
            frmElpPrt.prtNewPage
            prtSAB_Balance_Form
        End If
    End If
    If prevYBIACPT0.COMPTEOBL <> meYBIACPT0.COMPTEOBL Or prevYBIACPT0.COMPTEDEV <> meYBIACPT0.COMPTEDEV Then
        If Trim(prevYBIACPT0.COMPTEOBL) <> "" Then
            prtSAB_Balance_B_COMPTEOBL
        End If
    End If
Else
    If prevYBIACPT0.COMPTEDEV <> meYBIACPT0.COMPTEDEV Then
        If Trim(prevYBIACPT0.COMPTEDEV) <> "" Then
            prtSAB_Balance_B_COMPTEDEV
        End If
    End If
    If prevYBIACPT0.COMPTEOBL <> meYBIACPT0.COMPTEOBL Then
        If Trim(prevYBIACPT0.COMPTEOBL) <> "" Then
            prtSAB_Balance_B_COMPTEOBL
        End If
    End If
End If
    If prevYBIACPT0.COMPTEOBL <> meYBIACPT0.COMPTEOBL Then
        If Trim(prevYBIACPT0.COMPTEOBL) <> "" Then
            prtSAB_Balance_B_COMPTEOBL
        End If
    End If

blnOk = True
If mBalance_ZSOLDE0 = 0 Then
    curX = meYBIACPT0.SOLDECEN
Else
    V = rsZSOLDE0_Read(meYBIACPT0.COMPTECOM, mBalance_ZSOLDE0, curX)
    If Not IsNull(V) Then blnOk = False
End If

''If Not blnBalance_Compte_Soldé Or curX <> 0 Then
If blnBalance_Compte_Soldé And curX = 0 Then blnOk = False
If blnOk Then

    nbCOMPTEOBL_Line = nbCOMPTEOBL_Line + 1
    nbCOMPTEDEV_Line = nbCOMPTEDEV_Line + 1
    If blnBalance_B_Détail Then
    
        prtSAB_Balance_L "B", curX
        
        If blnChk_BalanceStock Then
            XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 6
           If meDORCPTDMV > 0 Then
                XPrt.CurrentX = prtMinX + 6500: XPrt.Print dateIBM10(meDORCPTDMV, True);
                If meDORCPTDMV <> meYBIACPT0.SOLDEDMO Then XPrt.Print " +";
            Else
                XPrt.CurrentX = prtMinX + 6500: XPrt.Print dateIBM10(meYBIACPT0.SOLDEDMO, True);
            End If
            XPrt.FontBold = True
            XPrt.ForeColor = vbMagenta
            If meYSTOMON <> -2 Then
                XPrt.FontBold = True
                If meYSTOMON = -1 Then
                    X = "??"
                Else
                    curX1 = Abs(meYBIACPT0.SOLDECEN)
                    curX2 = Abs(meYSTOMON)
                    curX = Abs(curX1 - curX2)
                    If curX = 0 Then
                        X = "="
                        XPrt.ForeColor = vbBlue
                    Else
                        X = Format$(Abs(curX2), "### ### ### ### ##0.00")
                        XPrt.CurrentX = prtMinX + 6400 - XPrt.TextWidth(X)
                        XPrt.Print X;
    
                        X = "##"
                    End If
                End If
                XPrt.CurrentX = prtMinX + 12100: XPrt.Print X;
                XPrt.FontBold = False
    
            End If
            XPrt.ForeColor = prtForeColor
            XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 8
       End If
        
        If meYBIACPT0.PLANCOPRO = "CAV" Or meYBIACPT0.PLANCOPRO = "LOR" Then
            If meYBIACPT0.COMPTEOUV > wAMJ_6M_00 And meYBIACPT0.COMPTEOUV < wAMJ_6M_99 Then
                XPrt.FontBold = True
                XPrt.CurrentX = prtMinX + 12050: XPrt.Print "6M";
                XPrt.FontBold = False
            End If
        End If
        
        If blnBalance_Pays Then
            XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 6
            XPrt.CurrentX = prtMinX - 180: XPrt.Print meYBIACPT0.CLIENARSD;
            XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 8
        End If
    Else
        prtSAB_Balance_CV curX
    End If
    
    nbClient_Line = nbClient_Line + 1
    
    prtSAB_Balance_B_Dev_Classe_Cumul
    
    If meCV2.Montant > 0 Then
        curCOMPTEOBL_Db = curCOMPTEOBL_Db + meCV1.Montant
        curCOMPTEDEV_Db = curCOMPTEDEV_Db + meCV1.Montant
        curCOMPTEOBL_Db_EUR = curCOMPTEOBL_Db_EUR + meCV2.Montant
        curCOMPTEDEV_Db_EUR = curCOMPTEDEV_Db_EUR + meCV2.Montant
        curListe_Db = curListe_Db + meCV2.Montant
    Else
        curCOMPTEOBL_Cr = curCOMPTEOBL_Cr + meCV1.Montant
        curCOMPTEDEV_Cr = curCOMPTEDEV_Cr + meCV1.Montant
        curCOMPTEOBL_Cr_EUR = curCOMPTEOBL_Cr_EUR + meCV2.Montant
        curCOMPTEDEV_Cr_EUR = curCOMPTEDEV_Cr_EUR + meCV2.Montant
        curListe_Cr = curListe_Cr + meCV2.Montant
    End If
    If blnFile_CSV Then
    
        X = meYBIACPT0.COMPTEDEV & ";" & meYBIACPT0.COMPTEOBL & ";" & meYBIACPT0.COMPTECOM & ";" & meYBIACPT0.COMPTEINT & ";"
        If meCV2.Montant > 0 Then
            XS = cur_AbsV(meCV1.Montant) & "; ;" & cur_AbsV(meCV2.Montant) & "; "
        Else
            XS = " ;" & cur_AbsV(meCV1.Montant) & "; ;" & cur_AbsV(meCV2.Montant)
        End If
        If meYBIACPT0.COMPTEOUV > 0 Then
            xText = ";" & meYBIACPT0.COMPTEOUV + 19000000 & ";" & meYBIACPT0.COMPTEFON & ";"
        Else
            xText = ";" & ";" & meYBIACPT0.COMPTEFON & ";"
        End If
        If meYBIACPT0.SOLDEDMO > 0 Then xText = xText & meYBIACPT0.SOLDEDMO + 19000000
       
        Call File_Export_Monitor("Print", idFile_CSV, X & XS & xText)
    End If
End If

End Sub


Public Sub prtSAB_Balance_B_Fin()
Dim I As Long, K As Long
Dim blnOk As Boolean

prtSAB_Balance_B_COMPTEOBL
If blnBalance_B_Détail Then
    XPrt.DrawWidth = 10
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight + 20
    XPrt.Line (prtMinX + 7500, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY), prtLineColor
    XPrt.DrawWidth = 3
End If

If blnBalance_Pays Then                                  ' Tri Pays / PCi / Compte / Dev
    prtSAB_Balance_B_COMPTEOBL
Else
    prtSAB_Balance_B_COMPTEDEV
    prtSAB_Balance_B_COMPTEOBL
End If

If blnBalance_B_Récap Then
    prtTitleText = "Récapitulatif de la " & xTitleText
    
    If blnBalance_B_Détail Then
            frmElpPrt.prtNewPage
            prtSAB_Balance_Form
    End If
    
    prevSAB_BALANCE = arrSAB_BALANCE(1)
    
    For I = 1 To arrSAB_BALANCE_Nb
        blnOk = True
        
        If Not blnBalance_Pays Then
            If prevSAB_BALANCE.COMPTEDEV <> arrSAB_BALANCE(I).COMPTEDEV Then
            
                blnOk = False
                prtSAB_Balance_B_Total_Cumul 0
            End If
        End If
        
        If blnOk Then
            For K = 1 To 5
                If Mid$(prevSAB_BALANCE.COMPTEOBL, 1, K) <> Mid$(arrSAB_BALANCE(I).COMPTEOBL, 1, K) Then
                    prtSAB_Balance_B_Total_Cumul K
                    Exit For
                End If
            Next K
        End If
        
        If totalSAB_BALANCE(6).blnPrint Then
            If arrSAB_BALANCE(I).DB <> 0 Or arrSAB_BALANCE(I).CR <> 0 Then
                prtSAB_Balance_B_Total_Prt 6, arrSAB_BALANCE(I)
            End If
        End If
        
        prevSAB_BALANCE = arrSAB_BALANCE(I)
        totalSAB_BALANCE(5).DB = totalSAB_BALANCE(5).DB + prevSAB_BALANCE.DB
        totalSAB_BALANCE(5).CR = totalSAB_BALANCE(5).CR + prevSAB_BALANCE.CR
        totalSAB_BALANCE(5).DB_EUR = totalSAB_BALANCE(5).DB_EUR + prevSAB_BALANCE.DB_EUR
        totalSAB_BALANCE(5).CR_EUR = totalSAB_BALANCE(5).CR_EUR + prevSAB_BALANCE.CR_EUR
    
    Next I
    If blnBalance_Pays Then
        prevSAB_BALANCE.COMPTEDEV = "***"
        For K = 0 To 6
                totalSAB_BALANCE(K).COMPTEDEV = ""
                totalSAB_BALANCE(K).DB = 0: totalSAB_BALANCE(K).CR = 0
        Next K
    End If
    prtSAB_Balance_B_Total_Cumul 0

End If

If blnBalance_Récap_Bilan Then
    prtSAB_Balance_B_Dev_Classe_Prt
End If
End Sub
Public Sub prtSAB_Balance_Total_Init()
Dim I As Integer
Dim xSQL As String, xDevise As String

ReDim arrSAB_BALANCE(101) As typeSAB_BALANCE
arrSAB_BALANCE_NbMax = 100: arrSAB_BALANCE_Nb = 0
zSAB_BALANCE.COMPTEDEV = ""
zSAB_BALANCE.COMPTEOBL = ""
zSAB_BALANCE.COMPTECOM = ""
zSAB_BALANCE.COMPTEINT = ""
zSAB_BALANCE.DB = 0
zSAB_BALANCE.CR = 0
zSAB_BALANCE.DB_EUR = 0
zSAB_BALANCE.CR_EUR = 0
zSAB_BALANCE.blnPrint = False
zSAB_BALANCE.blnPrint_Fontbold = False
zSAB_BALANCE.blnPrint_Line = False
zSAB_BALANCE.iPrint_Trame = 0

prevSAB_BALANCE = zSAB_BALANCE
détailSAB_BALANCE = zSAB_BALANCE
For I = 0 To 10
    totalSAB_BALANCE(I) = zSAB_BALANCE
Next I

arrDev_Nb = 0
xSQL = "select * from YBIATAB0 where" _
    & " BIATABID = 'DEVISE'" _
    & " and BIATABK1 = 'ISO'" _
    & " order by BIATABK2"
    
Set rsMDB = cnMDB.Execute(xSQL)
Do While Not rsMDB.EOF
    xDevise = rsMDB("BIATABK2")
        arrDev_Nb = arrDev_Nb + 1
        arrDEV_C1A5(arrDev_Nb) = zSAB_BALANCE
        arrDEV_C1A5(arrDev_Nb).COMPTEDEV = xDevise
        arrDEV_C6A8(arrDev_Nb) = zSAB_BALANCE
        arrDEV_C6A8(arrDev_Nb).COMPTEDEV = xDevise
        arrDEV_C9(arrDev_Nb) = zSAB_BALANCE
        arrDEV_C9(arrDev_Nb).COMPTEDEV = xDevise
    rsMDB.MoveNext
Loop


' Code devise non trouvé
arrDEV_C1A5(arrDev_Nb + 1) = zSAB_BALANCE
arrDEV_C6A8(arrDev_Nb + 1) = zSAB_BALANCE
arrDEV_C9(arrDev_Nb + 1) = zSAB_BALANCE


End Sub

Public Sub prtSAB_Balance_B_Total_Cumul(lK As Long)
Dim K As Long

For K = 5 To lK Step -1
    totalSAB_BALANCE(K).COMPTEDEV = prevSAB_BALANCE.COMPTEDEV
    totalSAB_BALANCE(K).COMPTEOBL = Mid$(prevSAB_BALANCE.COMPTEOBL, 1, K)
    
    If totalSAB_BALANCE(K).blnPrint Then prtSAB_Balance_B_Total_Prt K, totalSAB_BALANCE(K)
    
    If K > 0 Then
        totalSAB_BALANCE(K - 1).DB = totalSAB_BALANCE(K - 1).DB + totalSAB_BALANCE(K).DB
        totalSAB_BALANCE(K - 1).CR = totalSAB_BALANCE(K - 1).CR + totalSAB_BALANCE(K).CR
        totalSAB_BALANCE(K - 1).DB_EUR = totalSAB_BALANCE(K - 1).DB_EUR + totalSAB_BALANCE(K).DB_EUR
        totalSAB_BALANCE(K - 1).CR_EUR = totalSAB_BALANCE(K - 1).CR_EUR + totalSAB_BALANCE(K).CR_EUR
    End If
    totalSAB_BALANCE(K).DB = 0
    totalSAB_BALANCE(K).CR = 0
    totalSAB_BALANCE(K).DB_EUR = 0
    totalSAB_BALANCE(K).CR_EUR = 0
Next K
End Sub

Public Sub prtSAB_Balance_B_Total_Init(lX As String, lSAB_BALANCE As typeSAB_BALANCE)

If Mid$(lX, 1, 1) = "1" Then lSAB_BALANCE.blnPrint = True  'Imprimer ce niveau
If Mid$(lX, 2, 1) = "1" Then lSAB_BALANCE.blnPrint_Fontbold = True  'gras
If Mid$(lX, 3, 1) = "1" Then lSAB_BALANCE.blnPrint_Line = True  'ligne séparation
lSAB_BALANCE.iPrint_Trame = 255 - Val(Mid$(lX, 4, 3)) 'trame

End Sub

Public Sub prtSAB_Client_Stat_Z(lSAB_Client_Stat As typeSAB_Client_Stat)
lSAB_Client_Stat.Nb_Client = 0
lSAB_Client_Stat.Nb_Compte = 0
lSAB_Client_Stat.Nb_Compte_Annulé = 0
lSAB_Client_Stat.Solde_DB = 0
lSAB_Client_Stat.Solde_CR = 0
End Sub

Public Sub prtSAB_Client_Stat_Cumul()
If SAB_Client_Stat.Nb_Compte <> 0 Then
    SAB_Client_Stat_Actif.Nb_Client = SAB_Client_Stat_Actif.Nb_Client + 1
    SAB_Client_Stat_Actif.Nb_Compte = SAB_Client_Stat_Actif.Nb_Compte + SAB_Client_Stat.Nb_Compte
    SAB_Client_Stat_Actif.Nb_Compte_Annulé = SAB_Client_Stat_Actif.Nb_Compte_Annulé + SAB_Client_Stat.Nb_Compte_Annulé
    SAB_Client_Stat_Actif.Solde_DB = SAB_Client_Stat_Actif.Solde_DB + SAB_Client_Stat.Solde_DB
    SAB_Client_Stat_Actif.Solde_CR = SAB_Client_Stat_Actif.Solde_CR + SAB_Client_Stat.Solde_CR

    SAB_Client_Stat_Produit.Nb_Client = SAB_Client_Stat_Produit.Nb_Client + 1
    SAB_Client_Stat_Produit.Nb_Compte = SAB_Client_Stat_Produit.Nb_Compte + SAB_Client_Stat.Nb_Compte
    SAB_Client_Stat_Produit.Nb_Compte_Annulé = SAB_Client_Stat_Produit.Nb_Compte_Annulé + SAB_Client_Stat.Nb_Compte_Annulé
    SAB_Client_Stat_Produit.Solde_DB = SAB_Client_Stat_Produit.Solde_DB + SAB_Client_Stat.Solde_DB
    SAB_Client_Stat_Produit.Solde_CR = SAB_Client_Stat_Produit.Solde_CR + SAB_Client_Stat.Solde_CR

Else
    If SAB_Client_Stat.Nb_Compte_Annulé <> 0 Then
        SAB_Client_Stat_Annulé.Nb_Client = SAB_Client_Stat_Annulé.Nb_Client + 1
        SAB_Client_Stat_Annulé.Nb_Compte = SAB_Client_Stat_Annulé.Nb_Compte + SAB_Client_Stat.Nb_Compte
        SAB_Client_Stat_Annulé.Nb_Compte_Annulé = SAB_Client_Stat_Annulé.Nb_Compte_Annulé + SAB_Client_Stat.Nb_Compte_Annulé

    End If
End If
prtSAB_Client_Stat_Z SAB_Client_Stat

End Sub

Public Sub prtSAB_Client_Stat_Line()
Dim wId As String
Dim X As String

'wId = "SAB         CLIENACAT   CLI" & prevYBIACPT0.CLIENACAT
'Call srvYBIATAB0_Import_Read(wId, meYBIATAB0)
Dim wBIATABTEXT As String

Call rsYBIATAB0_Read("SAB", "CLIENACAT", "CLI" & prevYBIACPT0.CLIENACAT, wBIATABTEXT)

prtSAb_Client_Stat_NewLine

XPrt.CurrentX = prtMinX
XPrt.Print prevYBIACPT0.CLIENACAT;
XPrt.CurrentX = prtMinX + 400
XPrt.Print "- " & Mid$(wBIATABTEXT, 13, 30);

num_XPrt_Long SAB_Client_Stat_Actif.Nb_Client, prtMinX + 5000
num_XPrt_Currency SAB_Client_Stat_Actif.Solde_DB, prtMinX + 7000
num_XPrt_Currency SAB_Client_Stat_Actif.Solde_CR, prtMinX + 9000

num_XPrt_Long SAB_Client_Stat_Actif.Nb_Compte, prtMinX + 11000
num_XPrt_Long SAB_Client_Stat_Actif.Nb_Compte_Annulé, prtMinX + 12000

num_XPrt_Long SAB_Client_Stat_Annulé.Nb_Client, prtMinX + 14500
num_XPrt_Long SAB_Client_Stat_Annulé.Nb_Compte_Annulé, prtMinX + 15500


prtSAB_Client_Stat_Z SAB_Client_Stat_Actif
prtSAB_Client_Stat_Z SAB_Client_Stat_Annulé

End Sub
Public Sub prtSAB_Client_Stat_Form()
Dim wId As String
Dim X As String

XPrt.FontSize = 7
XPrt.FontBold = True
XPrt.DrawWidth = 2

XPrt.CurrentY = prtMinY + 50

XPrt.CurrentX = prtMinX
XPrt.Print "Produit / catégorie clientèle";

X = "Nb clients"
XPrt.CurrentX = prtMinX + 5000 - XPrt.TextWidth(X)
XPrt.Print X;

X = "Soldes débiteurs"
XPrt.CurrentX = prtMinX + 7000 - XPrt.TextWidth(X)
XPrt.Print X;

X = "Soldes créditeurs"
XPrt.CurrentX = prtMinX + 9000 - XPrt.TextWidth(X)
XPrt.Print X;

X = "Comptes .... Actifs"
XPrt.CurrentX = prtMinX + 11000 - XPrt.TextWidth(X)
XPrt.Print X;

X = "...Annulés"
XPrt.CurrentX = prtMinX + 12000 - XPrt.TextWidth(X)
XPrt.Print X;

X = "Clients non actifs"
XPrt.CurrentX = prtMinX + 14500 - XPrt.TextWidth(X)
XPrt.Print X;

X = "Cpt annulés"
XPrt.CurrentX = prtMinX + 15500 - XPrt.TextWidth(X)
XPrt.Print X;

XPrt.FontSize = 8
XPrt.FontBold = False

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight
XPrt.Line (prtMinX, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY), prtLineColor
XPrt.CurrentY = XPrt.CurrentY + 50


End Sub

Public Sub prtSAB_Client_Stat_Produit()
Dim wId As String
Dim X As String

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 1.25
XPrt.FontSize = 7
XPrt.FontBold = True

Call frmElpPrt.prtTrame(prtMinX + 390, XPrt.CurrentY - 50, prtMaxX, XPrt.CurrentY + prtlineHeight, " ", 240)

XPrt.CurrentX = prtMinX + 400
XPrt.Print "  " & SAB_Client_Stat_Produit_Lib;

num_XPrt_Long SAB_Client_Stat_Produit.Nb_Client, prtMinX + 5000
num_XPrt_Currency SAB_Client_Stat_Produit.Solde_DB, prtMinX + 7000
num_XPrt_Currency SAB_Client_Stat_Produit.Solde_CR, prtMinX + 9000

num_XPrt_Long SAB_Client_Stat_Produit.Nb_Compte, prtMinX + 11000
num_XPrt_Long SAB_Client_Stat_Produit.Nb_Compte_Annulé, prtMinX + 12000

XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 0.5
XPrt.FontSize = 8
XPrt.FontBold = False

prtSAB_Client_Stat_Z SAB_Client_Stat_Produit

End Sub

Public Sub prtSAB_Client_Stat_PLANCOPRO(lX As String)
Dim wId As String
Dim wBIATABTEXT As String

''wId = "SAB         PLANCOPRO   " & lX
Call rsYBIATAB0_Read("SAB", "PLANCOPRO", lX, wBIATABTEXT)
XPrt.FontBold = True

prtSAb_Client_Stat_NewLine

Call frmElpPrt.prtTrame(prtMinX, XPrt.CurrentY - 50, prtMaxX, XPrt.CurrentY + prtlineHeight, " ", 240)

XPrt.CurrentX = prtMinX
XPrt.Print lX;
XPrt.CurrentX = prtMinX + 400
SAB_Client_Stat_Produit_Lib = Mid$(wBIATABTEXT, 13, 30)
XPrt.Print "- " & SAB_Client_Stat_Produit_Lib;
XPrt.FontBold = False
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight * 0.5


End Sub



Public Sub prtSAB_Client_Stat_Form_End()
Dim mCurrenty As Long

mCurrenty = XPrt.CurrentY + prtlineHeight
XPrt.Line (prtMinX + 5100, prtMinY)-(prtMinX + 5100, mCurrenty), prtLineColor
XPrt.Line (prtMinX + 9100, prtMinY)-(prtMinX + 9100, mCurrenty), prtLineColor
XPrt.DrawWidth = 5

XPrt.Line (prtMinX + 12100, prtMinY)-(prtMinX + 12100, mCurrenty), prtLineColor

End Sub

Public Sub prtSAB_Balance_B_Dev_Classe_Cumul()
For arrDev_K = 1 To arrDev_Nb
    If meYBIACPT0.COMPTEDEV = arrDEV_C1A5(arrDev_K).COMPTEDEV Then Exit For
Next arrDev_K
Select Case Mid$(meYBIACPT0.COMPTEOBL, 1, 1)
    Case Is <= 5
        
        If meCV2.Montant > 0 Then
            arrDEV_C1A5(arrDev_K).DB = arrDEV_C1A5(arrDev_K).DB + meCV1.Montant
            arrDEV_C1A5(arrDev_K).DB_EUR = arrDEV_C1A5(arrDev_K).DB_EUR + meCV2.Montant
        Else
            arrDEV_C1A5(arrDev_K).CR = arrDEV_C1A5(arrDev_K).CR + meCV1.Montant
            arrDEV_C1A5(arrDev_K).CR_EUR = arrDEV_C1A5(arrDev_K).CR_EUR + meCV2.Montant
        End If
    Case Is <= 8
        If meCV2.Montant > 0 Then
            arrDEV_C6A8(arrDev_K).DB = arrDEV_C6A8(arrDev_K).DB + meCV1.Montant
            arrDEV_C6A8(arrDev_K).DB_EUR = arrDEV_C6A8(arrDev_K).DB_EUR + meCV2.Montant
        Else
            arrDEV_C6A8(arrDev_K).CR = arrDEV_C6A8(arrDev_K).CR + meCV1.Montant
            arrDEV_C6A8(arrDev_K).CR_EUR = arrDEV_C6A8(arrDev_K).CR_EUR + meCV2.Montant
        End If
    Case Else
        If meCV2.Montant > 0 Then
            arrDEV_C9(arrDev_K).DB = arrDEV_C9(arrDev_K).DB + meCV1.Montant
            arrDEV_C9(arrDev_K).DB_EUR = arrDEV_C9(arrDev_K).DB_EUR + meCV2.Montant
        Else
            arrDEV_C9(arrDev_K).CR = arrDEV_C9(arrDev_K).CR + meCV1.Montant
            arrDEV_C9(arrDev_K).CR_EUR = arrDEV_C9(arrDev_K).CR_EUR + meCV2.Montant
        End If
End Select
End Sub

Public Sub prtSAB_Liste_Total()
Dim mCurrenty As Integer
Dim curX As Currency

prtSAB_Balance_NewLine

XPrt.DrawWidth = 10
'_______________________
XPrt.CurrentY = XPrt.CurrentY + Height8_6: XPrt.FontSize = 6: XPrt.FontBold = True

prtFillColor = RGB(240, 240, 200)
Call frmElpPrt.prtTrame_Color(prtMinX + 7520, XPrt.CurrentY, prtMaxX - 20, XPrt.CurrentY + prtlineHeight, " ")
prtFillColor = prtFillColor_Standard


XPrt.Line (prtMinX + 7500, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY), prtLineColor
XPrt.CurrentY = XPrt.CurrentY + 50

XPrt.ForeColor = vbRed
X = Format$(Abs(curW_Db), "### ### ### ### ##0.00")
XPrt.CurrentX = prtMinX + 9400 - XPrt.TextWidth(X)
XPrt.Print X;
X = Format$(Abs(curCOMPTEDEV_Db_EUR), "### ### ### ### ##0.00")
XPrt.CurrentX = prtMaxX - 2100 - XPrt.TextWidth(X)
XPrt.Print X;

XPrt.ForeColor = prtForeColor
X = Format$(Abs(curW_Cr), "### ### ### ### ##0.00")
XPrt.CurrentX = prtMinX + 11400 - XPrt.TextWidth(X)
XPrt.Print X;
X = Format$(Abs(curCOMPTEDEV_Cr_EUR), "### ### ### ### ##0.00")
XPrt.CurrentX = prtMaxX - 100 - XPrt.TextWidth(X)
XPrt.Print X;

XPrt.CurrentX = prtMinX + 11600: XPrt.Print prevYBIACPT0.COMPTEDEV;

curW_Db = 0: curW_Cr = 0
curCOMPTEDEV_Db_EUR = 0: curCOMPTEDEV_Cr_EUR = 0

XPrt.DrawWidth = 10
XPrt.CurrentY = XPrt.CurrentY + prtlineHeight - 50
XPrt.Line (prtMinX + 7500, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY), prtLineColor

'_____________________________________________________________________________________

If meYBIACPT0.COMPTEDEV <> prevYBIACPT0.COMPTEDEV Then
    'prtSAB_Balance_NewLine
    prtFillColor = RGB(230, 230, 164)
    Call frmElpPrt.prtTrame_Color(prtMinX + 7520, XPrt.CurrentY, prtMaxX - 20, XPrt.CurrentY + prtlineHeight, " ")
    prtFillColor = prtFillColor_Standard
    curX = curCumul_Db + curCumul_Cr
    X = Format$(Abs(curX), "### ### ### ### ##0.00")
    If curX > 0 Then
        XPrt.ForeColor = vbRed
        XPrt.CurrentX = prtMinX + 9400 - XPrt.TextWidth(X)
    Else
        XPrt.ForeColor = vbBlue
        XPrt.CurrentX = prtMinX + 11400 - XPrt.TextWidth(X)
   End If
    XPrt.CurrentY = XPrt.CurrentY + 50
    XPrt.FontBold = True
    XPrt.Print X;
    XPrt.ForeColor = vbBlue
    XPrt.FontUnderline = True
    XPrt.CurrentX = prtMinX + 11600: XPrt.Print prevYBIACPT0.COMPTEDEV;
    XPrt.CurrentX = prtMinX + 5700: XPrt.Print "Solde CHA | PRO en " & prevYBIACPT0.COMPTEDEV;
    XPrt.CurrentY = XPrt.CurrentY + prtlineHeight - 50
    XPrt.Line (prtMinX + 7500, XPrt.CurrentY)-(prtMaxX, XPrt.CurrentY), prtLineColor
    XPrt.FontBold = False
    XPrt.FontUnderline = False
    curCumul_Db = 0: curCumul_Cr = 0
End If


XPrt.CurrentY = XPrt.CurrentY - Height8_6: XPrt.FontSize = 8:: XPrt.FontBold = False


End Sub
