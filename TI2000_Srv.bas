Attribute VB_Name = "srvTI2000"
Option Explicit
Dim xIn As String, xOut As String, xIn1 As String

Public paramTI2000DB2_AMJSituation As String, paramTI2000DB2_AMJValidité As String
Public paramTI2000DB2_Input As String, paramTI2000DB2_Output As String, paramTI2000DB2_Table As String
Public paramTI2000DB2_Xls As Boolean, paramTI2000DB2_mdb As Boolean
Public paramTI2000DB2_DossierAExclure As String
Dim arrField_Pos(300) As Integer, arrField_Len(300) As Integer, arrField_Nb As Integer

Dim arrField_Name(300) As String
Dim wCDDossier As typeCDDossier, xCDDossier As typeCDDossier
Dim wCDTIMaster As typeCDTIMaster
Dim wCDTICom As typeCDTICom, zCDTICom As typeCDTICom, blnCDTICom_Update As Boolean, blnCHPE As Boolean

Dim curX As Currency, X As String, X8 As String * 8


Dim kTIMasterKey As Integer, kDossier As Integer
Dim kTIPostingKey As Integer, kTIPostingBASIC_NUM As Integer
Dim kMontant As Integer, kAMT_O_S As Integer, wMontant As Currency
Dim kTIChargeKey As Integer
Dim kAMJPosting As Integer, wAMJPosting As String
Dim kACC_TYPE As Integer, wACC_TYPE As String
Dim kMATURED As Integer
Dim kTRAN_CODE As Integer, wTRAN_CODE As String
Dim kPosted_As As Integer, wPosted_As As Long
Dim kPPAY_KEY As Integer, wPPAY_KEY As Long
Dim kSK_CODE As Integer, wSK_CODE As String
Dim kSens As Integer, wSens As String
Dim kDevise As Integer, wDevise As String
Dim kAMJOuverture As Integer, kAMJValidité As Integer
Dim kSTATUS As Integer, wSTATUS As String

Dim kNature As Integer
Dim kComTaux As Integer

Dim wCDComD As typeCDComD, zCDComD As typeCDComD, blnCDComD_Update As Boolean
Public mCDComD_Type As String, mCDComD_Dossier As String
Public mCDComD_Nb As Long
Dim xCDComD As typeCDComD

Public arrCDComD(40) As typeCDComD
Public arrCDComD_Nb As Integer, arrCDComD_Index  As Integer
Public arrCDComD_NbMax As Integer
Public blnCDTICom_Read As Boolean

Dim recMvtp0 As typeMvtP0
Dim wCDPosting As typeCDPosting, zCDPosting As typeCDPosting

Dim mTIMt226 As Currency, mTIMt651 As Currency, mTIMt760 As Currency
Public Sub TIDB2_Load()
Dim xFileName As String, I As Integer
Dim wNb As Integer, wLen As Integer

On Error GoTo Error_Handle

CV_X1 = CV_Euro: CV_X2 = CV_Euro: CV_X3 = CV_Euro

arrField_Pos(0) = 1: arrField_Len(0) = 0: arrField_Nb = 0: wNb = 0
Open paramTI2000DB2_Input For Input As #1
Call lstErr_Clear(frmTI2000.lstErr, frmTI2000.cmdContext, "TIDB2 : début ..." & paramTI2000DB2_Table)

'Line Input #1, xIn1
'If Trim(xIn1) = "" Then Line Input #1, xIn1
'Line Input #1, xIn
Do
    Line Input #1, xIn
    If mId$(xIn, 1, 2) = "--" Then Exit Do
    xIn1 = xIn
Loop

For I = 1 To Len(xIn)
    If mId$(xIn, I, 1) = " " Then
        arrField_Len(arrField_Nb) = wLen
        arrField_Nb = arrField_Nb + 1
        arrField_Pos(arrField_Nb) = I + 1
        wLen = 0
    Else
        wLen = wLen + 1
    End If
Next I
arrField_Len(arrField_Nb) = wLen
For I = 0 To arrField_Nb
    arrField_Name(I) = Trim(mId$(xIn1, arrField_Pos(I), arrField_Len(I)))
Next I

If paramTI2000DB2_Xls Then TIDB2_Xls
If paramTI2000DB2_mdb Then
    Select Case mId$(paramTI2000DB2_Table, 1, 6)
        Case "MASTER": TIDB2_Master
        Case "CALCTE": TIDB2_CalcText
        Case "POSTIN": TIDB2_Posting
   End Select
End If
    
Call lstErr_AddItem(frmTI2000.lstErr, frmTI2000.cmdContext, "OK : " & paramTI2000DB2_Table): DoEvents
Close

Exit Sub

Error_Handle:
 MsgBox "erreur " & xIn
Close
End Sub

Public Sub TIDB2_Separator(lX As String)
Dim I As Integer

For I = 1 To arrField_Nb
    Mid$(lX, arrField_Pos(I) - 1, 1) = ";"
Next I

End Sub


Public Sub TIDB2_Xls()

Call TIDB2_Separator(xIn1)
Open paramTI2000DB2_Output For Output As #2
Print #2, xIn1

Do Until EOF(1)
    DoEvents
    Line Input #1, xIn
    Call TIDB2_Separator(xIn)
    Print #2, xIn
Loop

End Sub

Public Sub TIDB2_Master()
Dim X As String, X2 As String

kTIMasterKey = TIDB2_FieldName_Scan("KEY97")
If kTIMasterKey < 0 Then Call MsgBox("champ 'KEY97' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub

'kDossier = TIDB2_FieldName_Scan("MASTER_REF")
'If kDossier < 0 Then Call MsgBox("champ 'MASTER_REF' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub


kDossier = TIDB2_FieldName_Scan("REFNO_SERL")
If kDossier < 0 Then Call MsgBox("champ 'REFNO_SERL' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub

kMontant = TIDB2_FieldName_Scan("AMOUNT")
If kMontant < 0 Then Call MsgBox("champ 'AMOUNT' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub

kAMT_O_S = TIDB2_FieldName_Scan("AMT_O_S")
If kAMT_O_S < 0 Then Call MsgBox("champ 'AMT_O_S' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub

kDevise = TIDB2_FieldName_Scan("CCY")
If kDevise < 0 Then Call MsgBox("champ 'CCY' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub

kAMJValidité = TIDB2_FieldName_Scan("EXPIRY_DAT")
If kAMJValidité < 0 Then Call MsgBox("champ 'EXPIRY_DAT' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub

kAMJOuverture = TIDB2_FieldName_Scan("CTRCT_DATE")
If kAMJOuverture < 0 Then Call MsgBox("champ 'CTRCT_DATE' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub

kSTATUS = TIDB2_FieldName_Scan("STATUS")
If kSTATUS < 0 Then Call MsgBox("champ 'STATUS' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub

MDB.Execute "delete * from CDTIMaster"
mdbCDTIMaster.tableCDTIMaster_Open
recCDTIMaster_Init wCDTIMaster
wCDTIMaster.Method = "AddNew"

MDB.Execute "delete * from CDDossier"
mdbCDDossier.tableCDDossier_Open
recCDDossier_Init wCDDossier
wCDDossier.Method = "AddNew"


Do Until EOF(1)
    DoEvents
    Line Input #1, xIn
    If Trim(xIn) <> "" Then
        X = mId$(xIn, arrField_Pos(kMontant), arrField_Len(kMontant))
        If IsNumeric(X) Then
            wCDDossier.Devise = Trim(mId$(xIn, arrField_Pos(kDevise), arrField_Len(kDevise)))
            Call CV_AttributS(wCDDossier.Devise, CV_X1)
            curX = CCur(X)
            wCDDossier.Montant = curMaxD(curX, CV_X1.maxD)
            wCDTIMaster.TIMasterKey = CLng(mId$(xIn, arrField_Pos(kTIMasterKey), arrField_Len(kTIMasterKey)))
            wCDTIMaster.Dossier = Trim(mId$(xIn, arrField_Pos(kDossier), arrField_Len(kDossier)))
            dbCDTIMaster_Update wCDTIMaster
            
            wCDDossier.TIMasterKey = CLng(mId$(xIn, arrField_Pos(kTIMasterKey), arrField_Len(kTIMasterKey)))
            wCDDossier.Dossier = Trim(mId$(xIn, arrField_Pos(kDossier), arrField_Len(kDossier)))
            Call dateJMA_AMJ(Trim(mId$(xIn, arrField_Pos(kAMJOuverture), arrField_Len(kAMJOuverture))), X8)
            wCDDossier.AMJOuverture = X8
            Call dateJMA_AMJ(Trim(mId$(xIn, arrField_Pos(kAMJValidité), arrField_Len(kAMJValidité))), X8)
            wCDDossier.AMJValidité = X8
            wSTATUS = Trim(mId$(xIn, arrField_Pos(kSTATUS), arrField_Len(kSTATUS)))
            If wSTATUS = "LIV" Then
                wCDDossier.AMJSituation = ""
            Else
                wCDDossier.AMJSituation = wSTATUS
            End If
            
'2001.03.01 JPL            'wCDDossier.TIMasterEngagement = wCDDossier.Montant
'2001.03.01 JPL            'X = mId$(xIn, arrField_Pos(kAMT_O_S), arrField_Len(kAMT_O_S))
'2001.03.01 JPL            'If IsNumeric(X) Then
'2001.03.01 JPL            '    curX = CCur(X)
'2001.03.01 JPL            '    wCDDossier.TIMasterSolde = curMaxD(curX, CV_X1.maxD)
'2001.03.01 JPL            'End If
            dbCDDossier_Update wCDDossier
        End If
    End If
Loop

If paramTI2000DB2_DossierAExclure <> "" Then
    Open paramTI2000DB2_DossierAExclure For Input As #2
    Call lstErr_Clear(frmTI2000.lstErr, frmTI2000.cmdContext, "TIDB2 : début ..." & paramTI2000DB2_Table)
    Do Until EOF(2)
        DoEvents
        Line Input #2, xIn
        
        If Trim(xIn) <> "" Then
            X = Trim(xIn)
            X2 = mId$(X, 1, 2)
            Mid$(X, 1, 2) = "  "
            wCDDossier.Method = "Seek="
            wCDDossier.Dossier = Trim(X)
            If IsNull(dbCDDossier_ReadE(wCDDossier)) Then
                wCDDossier.AMJSituation = X2
                wCDDossier.Method = "Update"
                dbCDDossier_Update wCDDossier
            End If
        End If
    Loop
End If

mdbCDTIMaster.tableCDTIMaster_Close
mdbCDDossier.tableCDDossier_Close

End Sub
Public Sub TIDB2_Posting()

kTRAN_CODE = TIDB2_FieldName_Scan("TRAN_CODE")
If kTRAN_CODE < 0 Then Call MsgBox("champ 'TRAN_CODE' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub

kTIPostingKey = TIDB2_FieldName_Scan("KEY97")
If kTIPostingKey < 0 Then Call MsgBox("champ 'KEY97' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub

kDossier = TIDB2_FieldName_Scan("RECN_REF")
If kDossier < 0 Then Call MsgBox("champ 'RECN_REF' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub

kDevise = TIDB2_FieldName_Scan("CCY")
If kDevise < 0 Then Call MsgBox("champ 'CCY' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub

kMontant = TIDB2_FieldName_Scan("AMOUNT")
If kMontant < 0 Then Call MsgBox("champ 'AMOUNT' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub

kSens = TIDB2_FieldName_Scan("DR_CR_FLG")
If kSens < 0 Then Call MsgBox("champ 'DR_CR' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub

kAMJPosting = TIDB2_FieldName_Scan("VALUEDATE")
If kAMJPosting < 0 Then Call MsgBox("champ 'VALUEDATE' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub

kTIChargeKey = TIDB2_FieldName_Scan("CHARGE")
If kTIChargeKey < 0 Then Call MsgBox("champ 'CHARGE' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub

kPosted_As = TIDB2_FieldName_Scan("POSTED_AS")
If kPosted_As < 0 Then Call MsgBox("champ 'POSTED_AS' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub

kPPAY_KEY = TIDB2_FieldName_Scan("PPAY_KEY")
If kPPAY_KEY < 0 Then Call MsgBox("champ 'PPAY_KEY' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub

kACC_TYPE = TIDB2_FieldName_Scan("ACC_TYPE")
If kTIChargeKey < 0 Then Call MsgBox("champ 'ACC_TYPE' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub

kSK_CODE = TIDB2_FieldName_Scan("SK_CODE")
If kSK_CODE < 0 Then Call MsgBox("champ 'SK_CODE' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub


kTIPostingBASIC_NUM = TIDB2_FieldName_Scan("BASIC_NUM")
If kTIPostingBASIC_NUM < 0 Then Call MsgBox("champ 'BASIC_NUM' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub

kMATURED = TIDB2_FieldName_Scan("MATURED")
If kTIChargeKey < 0 Then Call MsgBox("champ 'MATURED' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub

mdbCDTIMaster.tableCDTIMaster_Open
recCDTIMaster_Init wCDTIMaster
wCDTIMaster.Method = "Seek="

mdbCDDossier.tableCDDossier_Open
recCDDossier_Init wCDDossier
wCDDossier.Method = "Seek="
xCDDossier = wCDDossier

mdbCDTICom.tableCDTICom_Open
recCDTICom_Init wCDTICom
wCDTICom.Method = "Seek="


MDB.Execute "delete * from CDComD"
mdbCDComD.tableCDComD_Open
recCDComD_Init zCDComD
zCDComD.Method = "AddNew"
wCDComD = zCDComD
mCDComD_Type = ""
mCDComD_Dossier = ""
blnCDComD_Update = False


MDB.Execute "delete * from CDPosting"
mdbCDPosting.tableCDPosting_Open
recCDPosting_Init zCDPosting
zCDPosting.Method = "AddNew"
wCDPosting = zCDPosting

'jpl 20010204 : extraction pour tri par type de dossier MVTP0

MDB.Execute "delete * from MVTP0"
tableMvtP0_Open
recMvtP0_Init recMvtp0
recMvtp0.Method = "AddNew"
mCDComD_Nb = 0

Do Until EOF(1)
    DoEvents
    Line Input #1, xIn
    If mId$(xIn, 1, 3) <> "CDE" And mId$(xIn, 1, 3) <> "CDI" Then GoTo Read_Next
    If Trim(xIn) <> "" Then
         Mid$(xIn, 4, 1) = "-"
        mCDComD_Dossier = Trim(mId$(xIn, arrField_Pos(kDossier), arrField_Len(kDossier)))
        
        wTRAN_CODE = Trim(mId$(xIn, arrField_Pos(kTRAN_CODE), arrField_Len(kTRAN_CODE)))
        If wTRAN_CODE = "760" Then
                wSK_CODE = Trim(mId$(xIn, arrField_Pos(kSK_CODE), arrField_Len(kSK_CODE)))
                Select Case wSK_CODE
                    Case "CK": mCDComD_Type = "RC": TIDB2_Posting_MvtP0_AddNew
                    Case "CF": mCDComD_Type = "RC": TIDB2_Posting_MvtP0_AddNew
                    Case "CN": mCDComD_Type = "RE": TIDB2_Posting_MvtP0_AddNew
                End Select
        Else
            X = Trim(mId$(xIn, arrField_Pos(kMATURED), arrField_Len(kMATURED)))
            If X = "Y" Then
                wACC_TYPE = Trim(mId$(xIn, arrField_Pos(kACC_TYPE), arrField_Len(kACC_TYPE)))
                mCDComD_Type = wACC_TYPE
                   Select Case wTRAN_CODE
                        Case "290": If mId$(wACC_TYPE, 1, 1) = "R" Then TIDB2_Posting_MvtP0_AddNew
                        ''''"290", "291", "292"
                        Case "790": If mId$(wACC_TYPE, 1, 1) = "R" Then TIDB2_Posting_MvtP0_AddNew
                        '''"790", "791", "792"
                        Case "226", "651", "380", "880": mCDComD_Type = "ZZ": TIDB2_Posting_MvtP0_AddNew
                        Case Else
                    End Select
        '''            End If
         '       End If
            End If
        End If
    End If
Read_Next:
Loop

recMvtp0.Method = "MoveFirst"
intReturn = tableMvtP0_Read(recMvtp0)


Do
    
    xIn = Trim(recMvtp0.Text)
    ''If Trim(xIn) <> "" Then
        'If mId$(xIn, 1, 4) = "CDE-" Then
          wTRAN_CODE = Trim(mId$(xIn, arrField_Pos(kTRAN_CODE), arrField_Len(kTRAN_CODE)))
        Call dateJMA_AMJ(Trim(mId$(xIn, arrField_Pos(kAMJPosting), arrField_Len(kAMJPosting))), X8)
         If mId$(xIn, 1, 10) = "CDE-060910" Then
       '     If wTRAN_CODE = "226" Then
                Mid$(xIn, 1, 4) = "CDE-"
       '     End If
           End If
        If wTRAN_CODE = "760" Then
            X = Trim(mId$(xIn, arrField_Pos(kMATURED), arrField_Len(kMATURED)))
            If X = "Y" Then TIDB2_Posting_Amount_CV: mTIMt760 = mTIMt760 + wMontant
            
                wSK_CODE = Trim(mId$(xIn, arrField_Pos(kSK_CODE), arrField_Len(kSK_CODE)))
                Select Case wSK_CODE
                    Case "CK": mCDComD_Type = "RC": TIDB2_Posting_OK
                    Case "CF": mCDComD_Type = "RC": TIDB2_Posting_OK
                    Case "CN": mCDComD_Type = "RE": TIDB2_Posting_OK
                End Select
        Else
            X = Trim(mId$(xIn, arrField_Pos(kMATURED), arrField_Len(kMATURED)))
            If X = "Y" Then
                wACC_TYPE = Trim(mId$(xIn, arrField_Pos(kACC_TYPE), arrField_Len(kACC_TYPE)))
                mCDComD_Type = wACC_TYPE
   ''''            If X8 <= paramTI2000DB2_AMJSituation Then
                   Select Case wTRAN_CODE
                       ''' Case "290", "291", "292"
                        Case "290"
                                If mId$(wACC_TYPE, 1, 1) = "R" Then wTRAN_CODE = "290": TIDB2_Posting_OK
                        
                        ''''Case "790", "791", "792": wTRAN_CODE = "790"
                        Case "790": wTRAN_CODE = "790"
                                If mId$(wACC_TYPE, 1, 1) = "R" Then TIDB2_Posting_OK
        
                        Case "226", "380":
TIDB2_Posting_Amount_CV:                             mTIMt226 = mTIMt226 + wMontant
                        Case "651", "880": TIDB2_Posting_Amount_CV: mTIMt651 = mTIMt651 + wMontant
                                         
                        Case Else
                    End Select
        '''            End If
         '       End If
            End If
        End If
    ''End If

    recMvtp0.Method = "MoveNext"
    intReturn = tableMvtP0_Read(recMvtp0)

Loop While intReturn = 0

TIDB2_Posting_Dossier_Update wCDDossier

mdbCDTIMaster.tableCDTIMaster_Close
mdbCDDossier.tableCDDossier_Close
mdbCDTICom.tableCDTICom_Close
mdbCDComD.tableCDComD_Close
tableMvtP0_Close
tableCDPosting_Close

'2000.01.09            X = Trim(mId$(xIn, arrField_Pos(kPosted_As), arrField_Len(kPosted_As)))
'2000.01.09            If IsNumeric(X) Then
'2000.01.09                wPosted_As = CLng(X)
'2000.01.09                If wPosted_As > 0 Then
'2000.01.09                    wTRAN_CODE = Trim(mId$(xIn, arrField_Pos(kTRAN_CODE), arrField_Len(kTRAN_CODE)))
 '2000.01.09                   Call dateJMA_AMJ(Trim(mId$(xIn, arrField_Pos(kAMJPosting), arrField_Len(kAMJPosting))), X8)
 '2000.01.09       ''''            If X8 <= paramTI2000DB2_AMJSituation Then
 '2000.01.09                      Select Case wTRAN_CODE
 '2000.01.09                           Case "290", "291", "292": wTRAN_CODE = "290": TIDB2_Posting_OK
 '2000.01.09
 '2000.01.09                           Case "790", "791", "792": wTRAN_CODE = "790":
 '2000.01.09                               X = Trim(mId$(xIn, arrField_Pos(kPPAY_KEY), arrField_Len(kPPAY_KEY)))
'2000.01.09                                If IsNumeric(X) Then TIDB2_Posting_OK
'2000.01.09                            Case "760": TIDB2_Posting_OK
'2000.01.09
'2000.01.09                            Case "226": TIDB2_Posting_226
'2000.01.09
'2000.01.09                            Case Else
'2000.01.09                        End Select
'2000.01.09        '''            End If
'2000.01.09                End If
'2000.01.09            End If


End Sub


Public Sub TIDB2_CalcText()

kDossier = TIDB2_FieldName_Scan("REFNO_SERL")
If kDossier < 0 Then Call MsgBox("champ 'REFNO_SERL' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub

kTIMasterKey = TIDB2_FieldName_Scan("MASTER_KEY")
If kTIMasterKey < 0 Then Call MsgBox("champ 'MASTER_KEY' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub

kTIChargeKey = kTIMasterKey + 1

kNature = TIDB2_FieldName_Scan("CALCTEXT")
If kNature < 0 Then Call MsgBox("champ 'CALCTEXT' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub

MDB.Execute "delete * from CDTICom"
mdbCDTICom.tableCDTICom_Open
recCDTICom_Init zCDTICom
zCDTICom.Method = "AddNew"

mdbCDTIMaster.tableCDTIMaster_Open
recCDTIMaster_Init wCDTIMaster
wCDTIMaster.Method = "Seek="

mdbCDDossier.tableCDDossier_Open
recCDDossier_Init wCDDossier
wCDDossier.Method = "Seek="
blnCDTICom_Update = False

Do Until EOF(1)
    DoEvents
    Line Input #1, xIn
    If Trim(xIn) <> "" Then
    
        If mId$(xIn, 1, 4) = "    " Then
            If blnCDTICom_Update Then dbCDTICom_Update wCDTICom
            wCDTICom.Nature = ""
            blnCDTICom_Update = False
            blnCHPE = False
            TIDB2_CalcText_Commission
            
        Else
            Select Case wCDTICom.Nature
                Case "CK ": TIDB2_CalcText_Commission_CK
                Case "CN ": TIDB2_CalcText_Commission_CN
''''                Case "CF ":TIDB2_CalcText_Commission_CK
            End Select
               
        End If
    End If
Loop

If blnCDTICom_Update Then dbCDTICom_Update wCDTICom

mdbCDTICom.tableCDTICom_Close

End Sub


Public Function TIDB2_FieldName_Scan(lName As String)
Dim I As Integer
TIDB2_FieldName_Scan = -1
For I = 0 To arrField_Nb
    If arrField_Name(I) = lName Then TIDB2_FieldName_Scan = I: Exit Function
Next I

End Function

Public Sub TIDB2_CalcText_Commission()
Dim X As String

If Len(Trim(xIn)) > arrField_Pos(kNature) Then
    X = Trim(mId$(xIn, arrField_Pos(kNature), arrField_Len(kNature)))
    Select Case X
        Case "Commission de confirmation":
                            TIDB2_CalcText_Commission_Init
                            wCDTICom.Nature = "CK "
                            'TIDB2_CalcText_Commission_CK
        Case "Commission confirmation silencieuse":
                            TIDB2_CalcText_Commission_Init
                            wCDTICom.Nature = "CF "
        'Case "Levée de documents taxable": TIDB2_CalcText_Commission_Init: wCDTICom.Nature = "LD": TIDB2_CalcText_Commission
        'Case "Commission de Modification": TIDB2_CalcText_Commission_Init: wCDTICom.Nature = "CM": TIDB2_CalcText_Commission_Montant
        'Case "Frais de Port et Telex": TIDB2_CalcText_Commission_Init: wCDTICom.Nature = "TLX": TIDB2_CalcText_Commission_Montant
        Case "Commission de notification (non taxable)"
                            TIDB2_CalcText_Commission_Init
                            wCDTICom.Nature = "CN "
        Case "Commission de notification (taxable)"
                            TIDB2_CalcText_Commission_Init
                            wCDTICom.Nature = "CN "
        'Case Else: MsgBox "non traité : " & X, vbInformation, "TIDB2_CalcText"
    End Select
End If
'Do Until EOF(1)
'    DoEvents
'    Line Input #1, xIn
'    Select Case mId$(xIn, 1, 8)
'        Case "<<CHPE>>": TIDB2_CalcText_Commission_NbJours
'        Case "<<CHAP>>": TIDB2_CalcText_Commission_Montant:   Exit Sub
'    End Select
    
'Loop
End Sub

Public Sub TIDB2_CalcText_Commission_Montant()
Dim I1 As Integer, I2 As Integer

'Line Input #1, xIn

'I2 = Len(Trim(xIn))
'wCDTICom.Devise = Trim(mId$(xIn, I2 - 2, 3))
'I1 = InStr(1, xIn, "EUR")
'If I1 <= 0 Then MsgBox "manque 'EUR'" & Trim(xIn), vbInformation, "TIDB2_CalcText_Commission_Montant": Exit Sub
'wCDTICom.MontantCVEur = CCur(mId$(xIn, 1, I1 - 1))
'I1 = InStr(1, xIn, "=")
'If I1 <= 0 Then MsgBox "manque '='" & Trim(xIn), vbInformation, "TIDB2_CalcText_Commission_Montant": Exit Sub
'wCDTICom.Montant = CCur(mId$(xIn, I1 + 1, I2 - 3 - I1))

'dbCDTICom_Update wCDTICom

End Sub
Public Sub TIDB2_CalcText_Commission_NbJours()
Dim I1 As Integer, I2 As Integer
I1 = InStr(1, xIn, "- ")
If I1 <= 0 Then MsgBox "manque '- '" & Trim(xIn), vbInformation, "TIDB2_CalcText_Commission_Montant": Exit Sub
I1 = InStr(I1 + 3, xIn, " ")
If I1 <= 0 Then MsgBox "manque 'espace 1'" & Trim(xIn), vbInformation, "TIDB2_CalcText_Commission_Montant": Exit Sub
I2 = InStr(I1, xIn, "<")
If I2 <= 0 Then MsgBox "manque '<<DAY2>>'" & Trim(xIn), vbInformation, "TIDB2_CalcText_Commission_Montant": Exit Sub
''wCDTICom.NbJours = CLng(mId$(xIn, I1, I2 - I1))

End Sub


Public Sub TIDB2_CalcText_Commission_CK()
If mId$(xIn, 1, 2) = "<<" Then blnCHPE = False
Select Case mId$(xIn, 1, 8)
    Case "<<CHCA>>": TIDB2_CalcText_Commission_CHCA
    Case "<<CHBA>>": TIDB2_CalcText_Commission_CHBA
    Case "<<CHPE>>": blnCHPE = True: kComTaux = 0
    Case "<<CHAP>>": TIDB2_CalcText_Commission_CHAP
    Case "<<CHAM>>": TIDB2_CalcText_Commission_CHAM
    Case Else:
            If blnCHPE Then TIDB2_CalcText_Commission_CK_ComTaux
End Select
    
End Sub
Public Sub TIDB2_CalcText_Commission_CN()
If mId$(xIn, 1, 2) = "<<" Then blnCHPE = False
Select Case mId$(xIn, 1, 8)
    Case "<<CHCA>>": TIDB2_CalcText_Commission_CHCA
    Case "<<CHBA>>": TIDB2_CalcText_Commission_CHBA
    Case "<<CHPE>>": blnCHPE = True: kComTaux = 0
    Case "<<CHAP>>": TIDB2_CalcText_Commission_CHAP
    Case "<<CHAM>>": TIDB2_CalcText_Commission_CHAM
    Case Else:
            TIDB2_CalcText_Commission_CN_ComTaux
End Select
    
End Sub

Public Sub TIDB2_Posting_OK()
Dim curX2 As Currency

X = Trim(mId$(xIn, arrField_Pos(kDossier), arrField_Len(kDossier)))

If mId$(X, 1, 3) = "CDI" Then
    xCDDossier.Dossier = mId$(X, 5, 6)
Else
    xCDDossier.Dossier = mId$(X, 6, 6)
End If
If xCDDossier.Dossier <> wCDDossier.Dossier Or mCDComD_Type <> wCDComD.Type Then
    TIDB2_Posting_Dossier_Update wCDDossier
    wCDDossier.Dossier = xCDDossier.Dossier
    TIDB2_Posting_Dossier_Init wCDDossier
End If

TIDB2_Posting_Amount

Call dateJMA_AMJ(Trim(mId$(xIn, arrField_Pos(kAMJPosting), arrField_Len(kAMJPosting))), wAMJPosting)

For arrCDComD_Index = 0 To arrCDComD_Nb
        If wAMJPosting <= arrCDComD(arrCDComD_Index).AmjF Then Exit For
Next arrCDComD_Index

If arrCDComD_Index > arrCDComD_Nb Then arrCDComD_Index = arrCDComD_Nb

Select Case wTRAN_CODE
    Case 290: wCDDossier.MontantEngagement = wCDDossier.MontantEngagement + wMontant
              arrCDComD(arrCDComD_Index).MvtEngagement = arrCDComD(arrCDComD_Index).MvtEngagement + wMontant
    Case 790: wCDDossier.MontantUtilisé = wCDDossier.MontantUtilisé + wMontant
              arrCDComD(arrCDComD_Index).MvtUtilisé = arrCDComD(arrCDComD_Index).MvtUtilisé + wMontant
    Case 760:
            If TIDB2_Posting_760 Then
                ''' 2001.03.06   wCDDossier.CommissionP = wCDDossier.CommissionP + wMontant
                arrCDComD(arrCDComD_Index).CommissionP = arrCDComD(arrCDComD_Index).CommissionP + wMontant
                arrCDComD(arrCDComD_Index).CommissionPAmj = wAMJPosting
                arrCDComD(arrCDComD_Index).TIChargeKey = wCDTICom.TIChargeKey
            End If
End Select
    

End Sub



Public Sub TIDB2_Posting_Dossier_Init(lCDDossier As typeCDDossier)
Dim blnOk As Boolean

blnCDTICom_Read = False
lCDDossier.Method = "Seek="
If IsNull(dbCDDossier_ReadE(lCDDossier)) Then
    blnCDComD_Update = True
'jpl 20010204     lCDDossier.MontantEngagement = 0
'jpl 20010204     lCDDossier.MontantUtilisé = 0
'jpl 20010204     lCDDossier.CommissionP = 0

    wCDTICom.Dossier = lCDDossier.Dossier

    wCDComD = zCDComD
    wCDComD.Dossier = lCDDossier.Dossier
    wCDComD.Type = mCDComD_Type
    
    mTIMt651 = 0: mTIMt226 = 0: mTIMt760 = 0
    CV_X2.DeviseIso = lCDDossier.Devise

    arrCDComD_Nb = 0
    wCDComD.AmjD = lCDDossier.AMJOuverture
    wCDComD.Devise = lCDDossier.Devise
    blnOk = False
    
    Do
        wCDComD.AmjF = dateElp("MoisAdd", 3, wCDComD.AmjD)
        arrCDComD(arrCDComD_Nb) = wCDComD
        If wCDComD.AmjF >= lCDDossier.AMJValidité Then
            blnOk = True
        Else
            arrCDComD_Nb = arrCDComD_Nb + 1
            wCDComD.AmjD = wCDComD.AmjF
            If arrCDComD_Nb > 30 Then blnOk = True
        End If
        
    Loop Until blnOk
       
End If

End Sub


Public Sub TIDB2_CalcText_Commission_Init()
wCDTICom = zCDTICom
wCDTICom.Dossier = CLng(mId$(xIn, arrField_Pos(kDossier), arrField_Len(kDossier)))
wCDTICom.TIMasterKey = CLng(mId$(xIn, arrField_Pos(kTIMasterKey), arrField_Len(kTIMasterKey)))
wCDTICom.TIChargeKey = CLng(mId$(xIn, arrField_Pos(kTIChargeKey), arrField_Len(kTIChargeKey)))
blnCDTICom_Update = True
End Sub

Public Sub TIDB2_CalcText_Commission_CHCA()
Dim I1 As Integer, I2 As Integer

I2 = Len(Trim(xIn))
wCDTICom.CHCA_Devise = Trim(mId$(xIn, I2 - 2, 3))
I1 = InStr(1, xIn, ">>")
If I1 <= 0 Then MsgBox "manque '>>'" & Trim(xIn), vbInformation, "TIDB2_CalcText_Commission_CHCA": Exit Sub
If I2 > I1 + 3 Then wCDTICom.CHCA = CCur(mId$(xIn, I1 + 3, I2 - I1 - 5))

End Sub
Public Sub TIDB2_CalcText_Commission_CHAP()
Dim I1 As Integer, I2 As Integer

'I2 = Len(Trim(xIn))
'wCDTICom.CHAP_Devise = Trim(mId$(xIn, I2 - 2, 3))
I1 = InStr(1, xIn, ">>")
If I1 <= 0 Then MsgBox "manque '>>'" & Trim(xIn), vbInformation, "TIDB2_CalcText_Commission_CHAP": Exit Sub
I2 = InStr(I1 + 3, xIn, ",")
If I1 <= 0 Then MsgBox "manque '>>'" & Trim(xIn), vbInformation, "TIDB2_CalcText_Commission_CHAP": Exit Sub

wCDTICom.CHAP = CCur(mId$(xIn, I1 + 3, I2 - I1 - 2))

'I2 = Len(Trim(xIn))
'wCDTICom.CHAP_Devise = Trim(mId$(xIn, I2 - 2, 3))

End Sub

Public Sub TIDB2_CalcText_Commission_CHAM()
Dim I1 As Integer, I2 As Integer

I2 = Len(Trim(xIn))
wCDTICom.CHAM_Devise = Trim(mId$(xIn, I2 - 2, 3))
I1 = InStr(1, xIn, ">>")
If I1 <= 0 Then MsgBox "manque '>>'" & Trim(xIn), vbInformation, "TIDB2_CalcText_Commission_CHAM": Exit Sub
wCDTICom.CHAM = CCur(mId$(xIn, I1 + 3, I2 - I1 - 5))

End Sub


Public Sub TIDB2_CalcText_Commission_CHBA()
Dim I1 As Integer, I2 As Integer
I2 = Len(Trim(xIn))
I1 = InStr(1, xIn, "@")
If I1 <= 0 Then MsgBox "manque '@'" & Trim(xIn), vbInformation, "TIDB2_CalcText_Commission_CHBA": Exit Sub
wCDTICom.CoursEur = CDbl(mId$(xIn, I1 + 1, I2 - I1 + 1))

I2 = I1 - 4
wCDTICom.CHBA_Devise = Trim(mId$(xIn, I2, 3))
I1 = InStr(1, xIn, ">>")
If I1 <= 0 Then MsgBox "manque '>>'" & Trim(xIn), vbInformation, "TIDB2_CalcText_Commission_CHCA": Exit Sub
wCDTICom.CHBA = CCur(mId$(xIn, I1 + 3, I2 - I1 - 4))

End Sub


Public Sub TIDB2_CalcText_Commission_CK_ComTaux()
Dim I1 As Integer, I2 As Integer, wComTaux As Double
Dim wAMJD As String * 8, wAMJF As String * 8

I1 = InStr(1, xIn, "%")
If I1 <= 0 Then MsgBox "manque '%'" & Trim(xIn), vbInformation, "TIDB2_CalcText_Commission_ComTaux": Exit Sub
wComTaux = CDbl(mId$(xIn, 1, I1 - 1))
I2 = I1 + 1
I1 = InStr(I2, xIn, "/")
If I1 <= 0 Then MsgBox "manque '/'" & Trim(xIn), vbInformation, "TIDB2_CalcText_Commission_ComTaux": Exit Sub
Mid$(wAMJD, 7, 2) = Format$(CInt(mId$(xIn, I2, I1 - I2)), "00")
I2 = I1 + 1
I1 = InStr(I2, xIn, "/")
If I1 <= 0 Then MsgBox "manque '/'" & Trim(xIn), vbInformation, "TIDB2_CalcText_Commission_ComTaux": Exit Sub
Mid$(wAMJD, 5, 2) = Format$(CInt(mId$(xIn, I2, I1 - I2)), "00")
I2 = I1 + 1
I1 = InStr(I2, xIn, " ")
If I1 <= 0 Then MsgBox "manque '/'" & Trim(xIn), vbInformation, "TIDB2_CalcText_Commission_ComTaux": Exit Sub
Mid$(wAMJD, 1, 4) = Format$(CInt(mId$(xIn, I2, I1 - I2)), "0000")


I1 = InStr(1, xIn, "-")
I2 = I1 + 1
I1 = InStr(I2, xIn, "/")
If I1 <= 0 Then MsgBox "manque '/'" & Trim(xIn), vbInformation, "TIDB2_CalcText_Commission_ComTaux": Exit Sub
Mid$(wAMJF, 7, 2) = Format$(CInt(mId$(xIn, I2, I1 - I2)), "00")
I2 = I1 + 1
I1 = InStr(I2, xIn, "/")
If I1 <= 0 Then MsgBox "manque '/'" & Trim(xIn), vbInformation, "TIDB2_CalcText_Commission_ComTaux": Exit Sub
Mid$(wAMJF, 5, 2) = Format$(CInt(mId$(xIn, I2, I1 - I2)), "00")
I2 = I1 + 1
I1 = InStr(I2, xIn, " ")
If I1 <= 0 Then MsgBox "manque '/'" & Trim(xIn), vbInformation, "TIDB2_CalcText_Commission_ComTaux": Exit Sub
Mid$(wAMJF, 1, 4) = Format$(CInt(mId$(xIn, I2, I1 - I2)), "0000")

kComTaux = kComTaux + 1
Select Case kComTaux
    Case 1
            wCDTICom.ComTaux1 = wComTaux
            wCDTICom.ComTaux2 = wComTaux
            wCDTICom.ComTaux3 = wComTaux
            wCDTICom.ComAMJD1 = wAMJD
            wCDTICom.ComAMJF1 = wAMJF
    Case 2
            wCDTICom.ComTaux2 = wComTaux
            wCDTICom.ComTaux3 = wComTaux
            wCDTICom.ComAMJD2 = wAMJD
            wCDTICom.ComAMJF2 = wAMJF
    Case 3
            wCDTICom.ComTaux3 = wComTaux
            wCDTICom.ComAMJD3 = wAMJD
            wCDTICom.ComAMJF3 = wAMJF
End Select

End Sub

Public Sub TIDB2_CalcText_Commission_CN_ComTaux()
Dim I1 As Integer, I2 As Integer, wComTaux As Double
Dim wAMJD As String * 8, wAMJF As String * 8


I1 = InStr(1, xIn, "%")
If I1 <= 0 Then
    ''''MsgBox "manque '%'" & Trim(xIn), vbInformation, "TIDB2_CalcText_Commission_ComTaux": Exit Sub
Else
    wComTaux = CDbl(mId$(xIn, 1, I1 - 1))
    wCDTICom.ComTaux1 = wComTaux
    wCDTICom.ComTaux2 = wComTaux
    wCDTICom.ComTaux3 = wComTaux
    wCDTICom.ComAMJD1 = "00000000"
    wCDTICom.ComAMJF1 = "99999999"
End If
End Sub


Public Sub TIDB2_Posting_Dossier_Update(lCDDossier As typeCDDossier)

If blnCDComD_Update Then
    lCDDossier.TIMt226 = mTIMt226
    lCDDossier.TIMt651 = mTIMt651
    lCDDossier.CommissionP = mTIMt760
    
    lCDDossier.Method = "Update"
    dbCDDossier_Update lCDDossier
    
    If Not blnCDTICom_Read Then
        TIDB2_CDTICom_RN arrCDComD(0).Type
        If blnCDTICom_Read Then TIDB2_Posting_760_Ok
    End If
    
    TIDB2_CDComD_Update
End If
End Sub

Public Sub TIDB2_CDComD_Update()
Dim I As Integer, curX As Currency, blnUtilisé As Boolean, blnEngagement As Boolean
    
        curX = 0 'wCDDossier.MontantEngagement
        'arrCDComD(0).MvtEngagement = arrCDComD(0).MvtEngagement - curX
        blnUtilisé = False: blnEngagement = False
        For I = 0 To arrCDComD_Nb
        
            If Not blnEngagement Then
                If arrCDComD(I).MvtEngagement > 0 Then
                    blnEngagement = True
                    arrCDComD(I).MontantBase = arrCDComD(I).MvtEngagement 'jpl 20010206wCDDossier.MontantEngagement
                Else
                    arrCDComD(I).CommissionTaux = 0
                End If
            Else
                arrCDComD(I).MontantBase = curX
                If arrCDComD(I).Type = "RE" Then arrCDComD(I).CommissionTaux = 0
            End If
            
           If arrCDComD(I).CoursEur = 0 Then arrCDComD(I).CoursEur = 1
           
            arrCDComD(I).CommissionD = (arrCDComD(I).MontantBase / arrCDComD(I).CoursEur) * arrCDComD(I).CommissionTaux / 100
            
            If Not blnUtilisé And arrCDComD(I).CommissionD > 0 Then
                If arrCDComD(I).Type = "RE" Then
                    If arrCDComD(I).CommissionD < 76.22 Then arrCDComD(I).CommissionD = 76.22
                Else
                     If arrCDComD(I).CommissionD < 152.44 Then arrCDComD(I).CommissionD = 152.44
               End If
            End If
            curX = curX + arrCDComD(I).MvtEngagement - arrCDComD(I).MvtUtilisé
            If arrCDComD(I).MvtUtilisé <> 0 Then blnUtilisé = True
            If arrCDComD(I).Method = constUpdate Then
                wCDComD = arrCDComD(I)
                wCDComD.Method = "Seek="
                dbCDComD_ReadE wCDComD
            End If
            
            dbCDComD_Update arrCDComD(I)
        Next I
End Sub


Public Function TIDB2_Posting_760() As Boolean
TIDB2_Posting_760 = False
X = Trim(mId$(xIn, arrField_Pos(kPosted_As), arrField_Len(kPosted_As)))
If IsNumeric(X) Then
    wPosted_As = CLng(X)
    If wPosted_As > 0 Then
        If wSK_CODE = "CF" Then  ''' Or wSK_CODE = "CN" Then
            TIDB2_Posting_760 = True
        Else
            wCDTICom.Dossier = wCDDossier.Dossier
            wCDTICom.TIChargeKey = CLng(mId$(xIn, arrField_Pos(kTIChargeKey), arrField_Len(kTIChargeKey)))
            wCDTICom.Method = "Seek="
            
            If tableCDTICom_Read(wCDTICom) = 0 Then
                   TIDB2_Posting_760 = True: TIDB2_Posting_760_Ok: blnCDTICom_Read = True
            End If
        End If
    End If
End If


End Function
Public Sub TIDB2_CDTICom_RN(lType)
Dim mNature As String, iReturn As Integer

If lType = "RE" Then
    mNature = "CN "
Else
    mNature = "CK "
End If
'If wCDDossier.Dossier = 58725 Then
'    iReturn = 0
'End If

blnCDTICom_Read = False
wCDTICom.Dossier = wCDDossier.Dossier
wCDTICom.TIChargeKey = 0
wCDTICom.Method = "Seek>="
Do
    iReturn = tableCDTICom_Read(wCDTICom)
    If iReturn = 0 Then
        If wCDTICom.Dossier <> wCDDossier.Dossier Then
            iReturn = 1
        Else
            If wCDTICom.Nature = mNature Then blnCDTICom_Read = True: iReturn = 1
        End If
    End If
wCDTICom.Method = "MoveNext"
Loop While iReturn = 0
End Sub


Public Sub TIDB2_Posting_760_Ok()
Dim I As Integer
For I = 0 To arrCDComD_Nb
    arrCDComD(I).CoursEur = wCDTICom.CoursEur
    If arrCDComD(I).AmjF <= wCDTICom.ComAMJF2 Then
        arrCDComD(I).CommissionTaux = wCDTICom.ComTaux2
    Else
        arrCDComD(I).CommissionTaux = wCDTICom.ComTaux3
    End If
Next I
End Sub

Public Sub TIDB2_Posting_MvtP0_AddNew()
mCDComD_Nb = mCDComD_Nb + 1
recMvtp0.Id = Trim(mCDComD_Dossier) & Trim(mCDComD_Type) & Format$(mCDComD_Nb, "000000000")
recMvtp0.Text = Trim(xIn)
dbMvtP0_Update recMvtp0

wCDPosting = zCDPosting
If mId$(xIn, 1, 3) = "CDI" Then
    wCDPosting.Dossier = CLng(mId$(xIn, 5, 6))
Else
    wCDPosting.Dossier = CLng(mId$(xIn, 6, 6))
End If

wCDPosting.Seq = mCDComD_Nb
wCDPosting.TRAN_CODE = Trim(mId$(xIn, arrField_Pos(kTRAN_CODE), arrField_Len(kTRAN_CODE)))
Call dateJMA_AMJ(Trim(mId$(xIn, arrField_Pos(kAMJPosting), arrField_Len(kAMJPosting))), wCDPosting.VALUEDATE)
wCDPosting.CCY = Trim(mId$(xIn, arrField_Pos(kDevise), arrField_Len(kDevise)))
wCDPosting.AMOUNT = CCur(mId$(xIn, arrField_Pos(kMontant), arrField_Len(kMontant)))

wCDPosting.SK_CODE = Trim(mId$(xIn, arrField_Pos(kSK_CODE), arrField_Len(kSK_CODE)))
wCDPosting.ACC_TYPE = Trim(mId$(xIn, arrField_Pos(kACC_TYPE), arrField_Len(kACC_TYPE)))
'''If Trim(wCDPosting.ACC_TYPE) = "-" Then wCDPosting.ACC_TYPE = "ZZ"

X = Trim(mId$(xIn, arrField_Pos(kPosted_As), arrField_Len(kPosted_As)))
If IsNumeric(X) Then wCDPosting.POSTED_AS = CLng(X)
wCDPosting.KEY97 = CLng(mId$(xIn, arrField_Pos(kTIPostingKey), arrField_Len(kTIPostingKey)))
 X = mId$(xIn, arrField_Pos(kTIChargeKey), arrField_Len(kTIChargeKey))
If IsNumeric(X) Then wCDPosting.CHARGE = CLng(X)

dbCDPosting_Update wCDPosting



End Sub

Public Sub TIDB2_CDComD_Reprise(lMsg As String)
Dim intReturn  As Integer, intReturn2 As Integer

mdbCDComD.tableCDComD_Open
recCDComD_Init wCDComD
xCDComD = wCDComD

xCDComD.Method = "MoveFirst"
intReturn = tableCDComD_Read(xCDComD)

Do
    wCDComD = xCDComD
    wCDComD.Method = "Seek="
    intReturn = tableCDComD_Read(wCDComD)
    If intReturn = 0 Then
        arrCDComD_Nb = -1
        intReturn2 = 0
        Do

            If intReturn2 = 0 Then
            
                wCDComD.Method = constUpdate
                arrCDComD_Nb = arrCDComD_Nb + 1
                arrCDComD(arrCDComD_Nb) = wCDComD
                
                wCDComD.Method = "MoveNext"
                intReturn = tableCDComD_Read(wCDComD)
                If wCDComD.Dossier <> xCDComD.Dossier _
                Or wCDComD.Type <> xCDComD.Type Then xCDComD = wCDComD: intReturn2 = 1

            End If
         Loop While intReturn = 0 And intReturn2 = 0
         
        If arrCDComD_Nb >= 0 Then TIDB2_CDComD_Update

    End If
Loop While intReturn = 0
        
lMsg = xCDComD.Dossier
End Sub

Public Sub TIDB2_Posting_Amount()

wDevise = Trim(mId$(xIn, arrField_Pos(kDevise), arrField_Len(kDevise)))
Call CV_AttributS(wDevise, CV_X1)
wMontant = CCur(mId$(xIn, arrField_Pos(kMontant), arrField_Len(kMontant)))
wMontant = curMaxD(wMontant, CV_X1.maxD)

End Sub

Public Sub TIDB2_Posting_Amount_CV()
Dim X As String

TIDB2_Posting_Amount
If wDevise <> CV_X2.DeviseIso Then
    CV_X1.Montant = wMontant
    
    CV_X1.CoursAmj = wAMJPosting
    CV_X1.DeviseIso = wDevise
    Call CV_Transitoire(CV_X1, CV_X2, CV_X3, X)
    wMontant = CV_X2.Montant
End If
End Sub
