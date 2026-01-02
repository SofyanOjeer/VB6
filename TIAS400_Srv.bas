Attribute VB_Name = "srvTIAS400"
Option Explicit
Public paramAS400IN As String
Public paramTIDB2_Master As String
Public paramTIDB2_Posting As String
Public paramTIDB2_PayDiff  As String
Public paramTIDB2_ComEnc As String
Public paramTIDB2_RegTrans As String
Public paramTIDB2_OK As String
Public paramTIDB2_TIAS400 As String

Public paramTIDB2_CDOESC As String
Public paramTIDB2_CDOUTI As String
Public paramTIDB2_CDOFRS As String

Public paramTIDB2_FullPosting As String

Dim xIn As String, xOut As String, xIn1 As String
Dim xIn_nb As Long, xOut_nb As Long
Dim wNb As Integer, wLen As Integer
Dim curX As Currency, X As String

Dim arrField_Pos(100) As Integer, arrField_Len(100) As Integer, arrField_Nb As Integer
Dim arrField_Name(100) As String

Dim DateAMJ As String
Dim Mnt As Currency

Dim X1 As String, X2 As String, X3 As String, X4 As String, X5 As String, X6 As String, X7 As String, X8 As String, X9 As String
Dim X10 As String, X11 As String, X12 As String, X13 As String, X14 As String, X15 As String, X16 As String, X17 As String, X18 As String, X19 As String
Dim X20 As String, X21 As String, X22 As String, X23 As String, X24 As String, X25 As String, X26 As String, X27 As String, X28 As String, X29 As String
Dim X30 As String, X31 As String, X32 As String, X33 As String, X34 As String, X35 As String, X36 As String, X37 As String, X38 As String, X39 As String
Dim X40 As String, X41 As String, X42 As String, X43 As String, X44 As String, X45 As String, X46 As String, X47 As String, X48 As String, X49 As String

Dim kMaster_KEY97 As Integer
Dim kMaster_REFNO_PFIX As Integer
Dim kMaster_REFNO_SERL As Integer
Dim kMaster_ORIG_REF As Integer
Dim kMaster_INPUT_BRN As Integer
Dim kMaster_BHALF_BRN As Integer
Dim kMaster_CTRCT_DATE As Integer
Dim kMaster_EXPIRY_DAT As Integer
Dim kMaster_STATUS As Integer
Dim kMaster_USERCODE1 As Integer
Dim kMaster_USERCODE2 As Integer
Dim kMaster_USERCODE3 As Integer
Dim kMaster_EV_COUNT As Integer
Dim kMaster_AMOUNT As Integer
Dim kMaster_CCY As Integer
Dim kMaster_AMT_O_S As Integer
Dim kMaster_LIAB_AMT As Integer
Dim kMaster_LIAB_CCY As Integer
Dim kMaster_REVOCABLE As Integer
Dim kMaster_CONF_INSTR As Integer
Dim kMaster_MGN_PCTAMT As Integer
Dim kMaster_MGN_AMT As Integer
Dim kMaster_MGN_CCY As Integer
Dim kMaster_MGN_BRN As Integer
Dim kMaster_MGN_BASIC As Integer
Dim kMaster_MGN_SFIX As Integer
Dim kMaster_CFM_PCTAMT As Integer
Dim kMaster_PCT_PLUS As Integer
Dim kMaster_PCT_MINUS As Integer
Dim kMaster_QUALIFIER As Integer
Dim kMaster_BEN_PTY As Integer
Dim kMaster_APP_PTY As Integer
Dim kMaster_RCVD_PTY As Integer
Dim kMaster_ISSPTY_PTY As Integer
Dim kMaster_RELMSTRKEY As Integer
Dim kMaster_RELMSTRREF As Integer

Dim kMaster_EXPIRY_LOC As Integer
Dim kMaster_OPERATIVE As Integer
Dim kMaster_TRANSFER As Integer
Dim kMaster_REVOLVING As Integer
Dim kMaster_REV_CUM As Integer
Dim kMaster_AVAIL_BY As Integer
Dim kMaster_SHIP_FROM As Integer
Dim kMaster_SHIP_TO As Integer
Dim kMaster_SHIP_DATE As Integer
Dim kMaster_PART_SHIP As Integer
Dim kMaster_TRANS_SHIP As Integer
Dim kMaster_DFR_APP As Integer
Dim kMaster_DFR_BEN As Integer

Dim kPosting_BRANCH_NUM As Integer
Dim kPosting_BASIC_NUM As Integer
Dim kPosting_ACC_SUFFIX As Integer
Dim kPosting_TRAN_CODE As Integer
Dim kPosting_VALUEDATE As Integer
Dim kPosting_DR_CR_FLG As Integer
Dim kPosting_AMOUNT As Integer
Dim kPosting_CCY As Integer
Dim kPosting_ACC_TYPE As Integer
Dim kPosting_SP_CODE As Integer
Dim kPosting_SK_CODE As Integer
Dim kPosting_BRANCH As Integer
Dim kPosting_MATURED As Integer
Dim kPosting_EQ3SEQNO As Integer
Dim kPosting_EQ3RECNREF As Integer
Dim kPosting_KEY97 As Integer
Dim kPosting_PARTY As Integer

Dim kPayDiff_PAY_AMT As Integer
Dim kPayDiff_PAY_AMTCCY As Integer
Dim kPayDiff_KEY97 As Integer
Dim kPayDiff_PAYEV_KEY As Integer
Dim kPayDiff_START_DATE As Integer
Dim kPayDiff_VALUE_DAT As Integer
Dim kPayDiff_PERIOD_NO As Integer
Dim kPayDiff_TYPE As Integer

Dim kComEnc_KEY97 As Integer
Dim kComEnc_VALUEDATE As Integer
Dim kComEnc_TRAN_CODE As Integer
Dim kComEnc_AMOUNT As Integer
Dim kComEnc_CCY As Integer
Dim kComEnc_CHARGE As Integer
Dim kComEnc_POSTINGTYP As Integer
Dim kComEnc_CHG_SCH As Integer
Dim kComEnc_CHG_TYPE As Integer
Dim kComEnc_CH_CODE As Integer
Dim kComEnc_STATUS As Integer
Dim kComEnc_CHG_FOR As Integer

Dim kRegTrans_KEY97 As Integer
Dim kRegTrans_SCHED_TYPE As Integer
Dim kRegTrans_CHG_TYPE As Integer
Dim kRegTrans_T1_NO As Integer
Dim kRegTrans_T1_UNIT As Integer
Dim kRegTrans_T1_PERCENT As Integer
Dim kRegTrans_TIERAMT1 As Integer
Dim kRegTrans_T2_NO As Integer
Dim kRegTrans_T2_UNIT As Integer
Dim kRegTrans_T2_PERCENT As Integer
Dim kRegTrans_TIERAMT2 As Integer
Dim kRegTrans_T3_NO As Integer
Dim kRegTrans_T3_UNIT As Integer
Dim kRegTrans_T3_PERCENT As Integer
Dim kRegTrans_TIERAMT3 As Integer
Dim kRegTrans_T4_NO As Integer
Dim kRegTrans_T4_UNIT As Integer
Dim kRegTrans_T4_PERCENT As Integer
Dim kRegTrans_TIERAMT4 As Integer
Dim kRegTrans_OVERALLMIN As Integer
Dim kRegTrans_CCY As Integer

Dim kCDOESC_MASTER_REFNO_PFIX As Integer
Dim kCDOESC_MASTER_REFNO_SERL As Integer
Dim kCDOESC_EVENT_REFNO_PFIX As Integer
Dim kCDOESC_EVENT_REFNO_SERL As Integer
Dim kCDOESC_PAY_AMT As Integer
Dim kCDOESC_PAY_AMTCCY As Integer
Dim kCDOESC_PARTPAYMNT_KEY97 As Integer
Dim kCDOESC_START_DATE As Integer
Dim kCDOESC_VALUE_DAT As Integer
Dim kCDOESC_PERIOD_NO As Integer
Dim kCDOESC_TYPE As Integer
Dim kCDOESC_REFERENCE As Integer
Dim kCDOESC_DEAL_PTY As Integer
Dim kCDOESC_DEAL_TYPE As Integer
Dim kCDOESC_KEY97 As Integer
Dim kCDOESC_RATE As Integer
Dim kCDOESC_DEAL_AMT As Integer
Dim kCDOESC_AMT_CCY As Integer
Dim kCDOESC_SPREAD As Integer
Dim kCDOESC_IDB As Integer
Dim kCDOESC_STARTDATE As Integer
Dim kCDOESC_MATURITY As Integer
Dim kCDOESC_DISC_AMT As Integer
Dim kCDOESC_DISC_CCY As Integer
Dim kCDOESC_NET_AMT As Integer
Dim kCDOESC_NET_CCY As Integer
Dim kCDOESC_DISC_FOR As Integer

Dim kCDOUTI_MASTER_REFNO_PFIX As Integer
Dim kCDOUTI_MASTER_REFNO_SERL As Integer
Dim kCDOUTI_EVENT_REFNO_PFIX As Integer
Dim kCDOUTI_EVENT_REFNO_SERL As Integer
Dim kCDOUTI_KEY97 As Integer
Dim kCDOUTI_MIXEDPAY As Integer
Dim kCDOUTI_PRSPTY_PTY As Integer
Dim kCDOUTI_PRES_DATE As Integer
Dim kCDOUTI_PRESAMT As Integer
Dim kCDOUTI_PRES_CCY As Integer
Dim kCDOUTI_SENT_DATE As Integer
Dim kCDOUTI_DOC_COUNT As Integer
Dim kCDOUTI_HOLD_DOC As Integer
Dim kCDOUTI_DOCSINORDR As Integer

Dim kCDOFRS_KEY97 As Integer
Dim kCDOFRS_VALUEDATE As Integer
Dim kCDOFRS_TRAN_CODE As Integer
Dim kCDOFRS_AMOUNT As Integer
Dim kCDOFRS_CCY As Integer
Dim kCDOFRS_CHARGE As Integer
Dim kCDOFRS_POSTINGTYP As Integer
Dim kCDOFRS_CHG_SCH As Integer
Dim kCDOFRS_CHG_TYPE As Integer
Dim kCDOFRS_CH_CODE As Integer
Dim kCDOFRS_STATUS As Integer
Dim kCDOFRS_CHG_FOR As Integer


Public Function param_Init()
Dim V
param_Init = Null

paramAS400IN = paramServer("\\AS400_IN\")

recElpTable_Init recElpTable
recElpTable.Method = "Seek="
recElpTable.Id = "TI"
Call lstErr_Clear(frmTIAS400.lstErr, frmTIAS400.cmdContext, "BIA.mdb : table : " & recElpTable.Id)

recElpTable.K1 = "DB2"

recElpTable.K2 = "Master"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramTIDB2_Master = paramServer(recElpTable.Memo)
Call lstErr_AddItem(frmTIAS400.lstErr, frmTIAS400.cmdContext, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))

recElpTable.K2 = "Posting"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramTIDB2_Posting = paramServer(recElpTable.Memo)
Call lstErr_AddItem(frmTIAS400.lstErr, frmTIAS400.cmdContext, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))

recElpTable.K2 = "FullPosting"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramTIDB2_FullPosting = paramServer(recElpTable.Memo)
Call lstErr_AddItem(frmTIAS400.lstErr, frmTIAS400.cmdContext, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))

recElpTable.K2 = "PayDiff"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramTIDB2_PayDiff = paramServer(recElpTable.Memo)
Call lstErr_AddItem(frmTIAS400.lstErr, frmTIAS400.cmdContext, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))

recElpTable.K2 = "ComEnc"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramTIDB2_ComEnc = paramServer(recElpTable.Memo)
Call lstErr_AddItem(frmTIAS400.lstErr, frmTIAS400.cmdContext, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))

recElpTable.K2 = "RegTrans"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramTIDB2_RegTrans = paramServer(recElpTable.Memo)
Call lstErr_AddItem(frmTIAS400.lstErr, frmTIAS400.cmdContext, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))

recElpTable.K2 = "Cdoesc"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
 paramTIDB2_CDOESC = paramServer(recElpTable.Memo)
'paramTIDB2_CDOESC = "\\Fr90524099\TI.DAT\Log\TIExtract\TI_Cdoesc.txt"
Call lstErr_AddItem(frmTIAS400.lstErr, frmTIAS400.cmdContext, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))

recElpTable.K2 = "Cdouti"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
 paramTIDB2_CDOUTI = paramServer(recElpTable.Memo)
'paramTIDB2_CDOUTI = "\\Fr90524099\TI.DAT\Log\TIExtract\TI_Cdouti.txt"
Call lstErr_AddItem(frmTIAS400.lstErr, frmTIAS400.cmdContext, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))

recElpTable.K2 = "Cdofrs"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
 paramTIDB2_CDOFRS = paramServer(recElpTable.Memo)
'paramTIDB2_CDOFRS = "\\Fr90524099\TI.DAT\Log\TIExtract\TI_Cdofrs.txt"
Call lstErr_AddItem(frmTIAS400.lstErr, frmTIAS400.cmdContext, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))

recElpTable.K2 = "OK"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramTIDB2_OK = paramServer(recElpTable.Memo)
Call lstErr_AddItem(frmTIAS400.lstErr, frmTIAS400.cmdContext, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))

recElpTable.K2 = "TIAS400"
V = dbElpTable_ReadE(recElpTable)
If Not IsNull(V) Then GoTo Table_Error
If IsNull(recElpTable.Memo) Then GoTo Memo_Error
paramTIDB2_TIAS400 = paramServer(recElpTable.Memo)
Call lstErr_AddItem(frmTIAS400.lstErr, frmTIAS400.cmdContext, Trim(recElpTable.K1) & "_ " & recElpTable.K2 & ": " & Trim(recElpTable.Memo))


Call lstErr_Clear(frmTIAS400.lstErr, frmTIAS400.cmdContext, "BIA.mdb : table : " & recElpTable.Id & ": ok ")


Exit Function

Table_Error:
param_Init = V
Exit Function

Memo_Error:
param_Init = "Memo"
MsgBox recElpTable.Id & " : " & recElpTable.K1 & " : " & recElpTable.K2 & " : Mémo absent", vbCritical, "srvTI.Param_Init"
Exit Function

Num_Error:
param_Init = "Num"
MsgBox recElpTable.Id & " : " & recElpTable.K1 & " : " & recElpTable.K2 & " : " & recElpTable.Memo & " :Mémo non numérique", vbCritical, "srvTI.Param_Init"
End Function


Public Sub TIDB2_Load_Init()
Dim I As Integer

On Error GoTo Error_Handle

arrField_Pos(0) = 1: arrField_Len(0) = 0: arrField_Nb = 0: wNb = 0

Do
    Line Input #1, xIn
    If Trim(xIn) <> "" Then
        If mId$(xIn, 1, 2) = "--" Then Exit Do
        xIn1 = xIn
    End If
Loop

wLen = 0
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

    
Exit Sub

Error_Handle:
 MsgBox "erreur " & xIn
Close
End Sub

Public Sub TIDB2_Master(lstMsg As ListBox)
Dim I1 As Integer
On Error GoTo Error_Handle


Open paramTIDB2_Master For Input As #1

TIDB2_Load_Init

'===============

kMaster_KEY97 = TIDB2_FieldName_Scan("KEY97")
kMaster_REFNO_PFIX = TIDB2_FieldName_Scan("REFNO_PFIX")
kMaster_REFNO_SERL = TIDB2_FieldName_Scan("REFNO_SERL")
kMaster_ORIG_REF = TIDB2_FieldName_Scan("ORIG_REF")
kMaster_INPUT_BRN = TIDB2_FieldName_Scan("INPUT_BRN")
kMaster_BHALF_BRN = TIDB2_FieldName_Scan("BHALF_BRN")
kMaster_CTRCT_DATE = TIDB2_FieldName_Scan("CTRCT_DATE")
kMaster_EXPIRY_DAT = TIDB2_FieldName_Scan("EXPIRY_DAT")
kMaster_STATUS = TIDB2_FieldName_Scan("STATUS")
kMaster_USERCODE1 = TIDB2_FieldName_Scan("USERCODE1")
kMaster_USERCODE2 = TIDB2_FieldName_Scan("USERCODE2")
kMaster_USERCODE3 = TIDB2_FieldName_Scan("USERCODE3")
kMaster_EV_COUNT = TIDB2_FieldName_Scan("EV_COUNT")
kMaster_AMOUNT = TIDB2_FieldName_Scan("AMOUNT")
kMaster_CCY = TIDB2_FieldName_Scan("CCY")
kMaster_AMT_O_S = TIDB2_FieldName_Scan("AMT_O_S")
kMaster_LIAB_AMT = TIDB2_FieldName_Scan("LIAB_AMT")
kMaster_LIAB_CCY = TIDB2_FieldName_Scan("LIAB_CCY")
kMaster_REVOCABLE = TIDB2_FieldName_Scan("REVOCABLE")
kMaster_CONF_INSTR = TIDB2_FieldName_Scan("CONF_INSTR")
kMaster_MGN_PCTAMT = TIDB2_FieldName_Scan("MGN_PCTAMT")
kMaster_MGN_AMT = TIDB2_FieldName_Scan("MGN_AMT")
kMaster_MGN_CCY = TIDB2_FieldName_Scan("MGN_CCY")
kMaster_MGN_BRN = TIDB2_FieldName_Scan("MGN_BRN")
kMaster_MGN_BASIC = TIDB2_FieldName_Scan("MGN_BASIC")
kMaster_MGN_SFIX = TIDB2_FieldName_Scan("MGN_SFIX")
kMaster_CFM_PCTAMT = TIDB2_FieldName_Scan("CFM_PCTAMT")
kMaster_PCT_PLUS = TIDB2_FieldName_Scan("PCT_PLUS")
kMaster_PCT_MINUS = TIDB2_FieldName_Scan("PCT_MINUS")
kMaster_QUALIFIER = TIDB2_FieldName_Scan("QUALIFIER")
kMaster_BEN_PTY = TIDB2_FieldName_Scan("BEN_PTY")
kMaster_APP_PTY = TIDB2_FieldName_Scan("APP_PTY")
kMaster_RCVD_PTY = TIDB2_FieldName_Scan("RCVD_PTY")
kMaster_ISSPTY_PTY = TIDB2_FieldName_Scan("ISSPTY_PTY")
kMaster_RELMSTRKEY = TIDB2_FieldName_Scan("RELMSTRKEY")
kMaster_RELMSTRREF = TIDB2_FieldName_Scan("RELMSTRREF")

kMaster_EXPIRY_LOC = TIDB2_FieldName_Scan("EXPIRY_LOC")
kMaster_OPERATIVE = TIDB2_FieldName_Scan("OPERATIVE")
kMaster_TRANSFER = TIDB2_FieldName_Scan("TRANSFER")
kMaster_REVOLVING = TIDB2_FieldName_Scan("REVOLVING")
kMaster_REV_CUM = TIDB2_FieldName_Scan("REV_CUM")
kMaster_AVAIL_BY = TIDB2_FieldName_Scan("AVAIL_BY")
kMaster_SHIP_FROM = TIDB2_FieldName_Scan("SHIP_FROM")
kMaster_SHIP_TO = TIDB2_FieldName_Scan("SHIP_TO")
kMaster_SHIP_DATE = TIDB2_FieldName_Scan("SHIP_DATE")
kMaster_PART_SHIP = TIDB2_FieldName_Scan("PART_SHIP")
kMaster_TRANS_SHIP = TIDB2_FieldName_Scan("TRANS_SHIP")
kMaster_DFR_APP = TIDB2_FieldName_Scan("DFR_APP")
kMaster_DFR_BEN = TIDB2_FieldName_Scan("DFR_BEN")

'==============================================================

If kMaster_KEY97 < 0 Then Call MsgBox("champ 'KEY97' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub
If kMaster_REFNO_PFIX < 0 Then Call MsgBox("champ 'REFNO_PFIX' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub
If kMaster_REFNO_SERL < 0 Then Call MsgBox("champ 'REFNO_SERL' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub
If kMaster_ORIG_REF < 0 Then Call MsgBox("champ 'ORIG_REF' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub
If kMaster_INPUT_BRN < 0 Then Call MsgBox("champ 'INPUT_BRN' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub
If kMaster_BHALF_BRN < 0 Then Call MsgBox("champ 'BHALF_BRN' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub
If kMaster_CTRCT_DATE < 0 Then Call MsgBox("champ 'CTRCT_DATE' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub
If kMaster_EXPIRY_DAT < 0 Then Call MsgBox("champ 'EXPIRY_DAT' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub
If kMaster_STATUS < 0 Then Call MsgBox("champ 'STATUS' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub
If kMaster_USERCODE1 < 0 Then Call MsgBox("champ 'USERCODE1' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub
If kMaster_USERCODE2 < 0 Then Call MsgBox("champ 'USERCODE2' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub
If kMaster_USERCODE3 < 0 Then Call MsgBox("champ 'USERCODE3' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub
If kMaster_EV_COUNT < 0 Then Call MsgBox("champ 'EV_COUNT' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub
If kMaster_AMOUNT < 0 Then Call MsgBox("champ 'AMOUNT' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub
If kMaster_CCY < 0 Then Call MsgBox("champ 'CCY' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub
If kMaster_AMT_O_S < 0 Then Call MsgBox("champ 'AMT_O_S' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub
If kMaster_LIAB_AMT < 0 Then Call MsgBox("champ 'LIAB_AMT' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub
If kMaster_LIAB_CCY < 0 Then Call MsgBox("champ 'LIAB_CCY' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub
If kMaster_REVOCABLE < 0 Then Call MsgBox("champ 'REVOCABLE' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub
If kMaster_CONF_INSTR < 0 Then Call MsgBox("champ 'CONF_INSTR' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub
If kMaster_MGN_PCTAMT < 0 Then Call MsgBox("champ 'MGN_PCTAMT' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub
If kMaster_MGN_AMT < 0 Then Call MsgBox("champ 'MGN_AMT' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub
If kMaster_MGN_CCY < 0 Then Call MsgBox("champ 'MGN_CCY' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub
If kMaster_MGN_BRN < 0 Then Call MsgBox("champ 'MGN_BRN' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub
If kMaster_MGN_BASIC < 0 Then Call MsgBox("champ 'MGN_BASIC' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub
If kMaster_MGN_SFIX < 0 Then Call MsgBox("champ 'MGN_SFIX' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub
If kMaster_CFM_PCTAMT < 0 Then Call MsgBox("champ 'CFM_PCTAMT' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub
If kMaster_PCT_PLUS < 0 Then Call MsgBox("champ 'PCT_PLUS' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub
If kMaster_PCT_MINUS < 0 Then Call MsgBox("champ 'PCT_MINUS' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub
If kMaster_QUALIFIER < 0 Then Call MsgBox("champ 'QUALIFIER' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub
If kMaster_BEN_PTY < 0 Then Call MsgBox("champ 'BEN_PTY' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub
If kMaster_APP_PTY < 0 Then Call MsgBox("champ 'APP_PTY' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub
If kMaster_RCVD_PTY < 0 Then Call MsgBox("champ 'RCVD_PTY' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub
If kMaster_ISSPTY_PTY < 0 Then Call MsgBox("champ 'ISSPTY_PTY' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub
If kMaster_RELMSTRKEY < 0 Then Call MsgBox("champ 'RELMSTRKEY' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub
If kMaster_RELMSTRREF < 0 Then Call MsgBox("champ 'RELMSTRREF' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub

If kMaster_EXPIRY_LOC < 0 Then Call MsgBox("champ 'EXPIRY_LOC' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub
If kMaster_OPERATIVE < 0 Then Call MsgBox("champ 'OPERATIVE' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub
If kMaster_TRANSFER < 0 Then Call MsgBox("champ 'TRANSFER' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub
If kMaster_REVOLVING < 0 Then Call MsgBox("champ 'REVOLVING' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub
If kMaster_REV_CUM < 0 Then Call MsgBox("champ 'REV_CUM' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub
If kMaster_AVAIL_BY < 0 Then Call MsgBox("champ 'AVAIL_BY' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub
If kMaster_SHIP_FROM < 0 Then Call MsgBox("champ 'SHIP_FROM' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub
If kMaster_SHIP_TO < 0 Then Call MsgBox("champ 'SHIP_TO' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub
If kMaster_SHIP_DATE < 0 Then Call MsgBox("champ 'SHIP_DATE' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub
If kMaster_PART_SHIP < 0 Then Call MsgBox("champ 'PART_SHIP' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub
If kMaster_TRANS_SHIP < 0 Then Call MsgBox("champ 'TRANS_SHIP' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub
If kMaster_DFR_APP < 0 Then Call MsgBox("champ 'DFR_APP' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub
If kMaster_DFR_BEN < 0 Then Call MsgBox("champ 'DFR_BEN' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub

''paramTIDB2_Output = paramAS400IN & "CDDOSW0"
''Open paramTIDB2_Output For Output As #2
Open "\\Fr11024427\As400_IN\CDDOSW0" For Output As #2

xIn_nb = 0
xOut_nb = 0

Do Until EOF(1)
    DoEvents
    Line Input #1, xIn
    If Trim(xIn) <> "" Then
        I1 = InStr(1, xIn, "enregistrement(s) sélectionné(s)")
        If I1 > 0 Then
            xIn_nb = CLng(Val(mId$(xIn, 1, I1 - 1)))
            Exit Do
        End If
        
        X1 = mId$(xIn, arrField_Pos(kMaster_KEY97), arrField_Len(kMaster_KEY97))
        X2 = mId$(xIn, arrField_Pos(kMaster_REFNO_PFIX), arrField_Len(kMaster_REFNO_PFIX))
        X3 = mId$(xIn, arrField_Pos(kMaster_REFNO_SERL), arrField_Len(kMaster_REFNO_SERL))
        X4 = mId$(xIn, arrField_Pos(kMaster_ORIG_REF), arrField_Len(kMaster_ORIG_REF))
        X5 = mId$(xIn, arrField_Pos(kMaster_INPUT_BRN), arrField_Len(kMaster_INPUT_BRN))
        X6 = mId$(xIn, arrField_Pos(kMaster_BHALF_BRN), arrField_Len(kMaster_BHALF_BRN))
        X7 = mId$(xIn, arrField_Pos(kMaster_CTRCT_DATE), arrField_Len(kMaster_CTRCT_DATE))
        X8 = mId$(xIn, arrField_Pos(kMaster_EXPIRY_DAT), arrField_Len(kMaster_EXPIRY_DAT))
        X9 = mId$(xIn, arrField_Pos(kMaster_STATUS), arrField_Len(kMaster_STATUS))
        X10 = mId$(xIn, arrField_Pos(kMaster_USERCODE1), arrField_Len(kMaster_USERCODE1))
        X11 = mId$(xIn, arrField_Pos(kMaster_USERCODE2), arrField_Len(kMaster_USERCODE2))
        X12 = mId$(xIn, arrField_Pos(kMaster_USERCODE3), arrField_Len(kMaster_USERCODE3))
        X13 = mId$(xIn, arrField_Pos(kMaster_EV_COUNT), arrField_Len(kMaster_EV_COUNT))
        X14 = mId$(xIn, arrField_Pos(kMaster_AMOUNT), arrField_Len(kMaster_AMOUNT))
        X15 = mId$(xIn, arrField_Pos(kMaster_CCY), arrField_Len(kMaster_CCY))
        X16 = mId$(xIn, arrField_Pos(kMaster_AMT_O_S), arrField_Len(kMaster_AMT_O_S))
        X17 = mId$(xIn, arrField_Pos(kMaster_LIAB_AMT), arrField_Len(kMaster_LIAB_AMT))
        X18 = mId$(xIn, arrField_Pos(kMaster_LIAB_CCY), arrField_Len(kMaster_LIAB_CCY))
        X19 = mId$(xIn, arrField_Pos(kMaster_REVOCABLE), arrField_Len(kMaster_REVOCABLE))
        X20 = mId$(xIn, arrField_Pos(kMaster_CONF_INSTR), arrField_Len(kMaster_CONF_INSTR))
        X21 = mId$(xIn, arrField_Pos(kMaster_MGN_PCTAMT), arrField_Len(kMaster_MGN_PCTAMT))
        X22 = mId$(xIn, arrField_Pos(kMaster_MGN_AMT), arrField_Len(kMaster_MGN_AMT))
        X23 = mId$(xIn, arrField_Pos(kMaster_MGN_CCY), arrField_Len(kMaster_MGN_CCY))
        X24 = mId$(xIn, arrField_Pos(kMaster_MGN_BRN), arrField_Len(kMaster_MGN_BRN))
        X25 = mId$(xIn, arrField_Pos(kMaster_MGN_BASIC), arrField_Len(kMaster_MGN_BASIC))
        X26 = mId$(xIn, arrField_Pos(kMaster_MGN_SFIX), arrField_Len(kMaster_MGN_SFIX))
        X27 = mId$(xIn, arrField_Pos(kMaster_CFM_PCTAMT), arrField_Len(kMaster_CFM_PCTAMT))
        X28 = mId$(xIn, arrField_Pos(kMaster_PCT_PLUS), arrField_Len(kMaster_PCT_PLUS))
        X29 = mId$(xIn, arrField_Pos(kMaster_PCT_MINUS), arrField_Len(kMaster_PCT_MINUS))
        X30 = mId$(xIn, arrField_Pos(kMaster_QUALIFIER), arrField_Len(kMaster_QUALIFIER))
        X31 = mId$(xIn, arrField_Pos(kMaster_BEN_PTY), arrField_Len(kMaster_BEN_PTY))
        X32 = mId$(xIn, arrField_Pos(kMaster_APP_PTY), arrField_Len(kMaster_APP_PTY))
        X33 = mId$(xIn, arrField_Pos(kMaster_RCVD_PTY), arrField_Len(kMaster_RCVD_PTY))
        X34 = mId$(xIn, arrField_Pos(kMaster_ISSPTY_PTY), arrField_Len(kMaster_ISSPTY_PTY))
        X35 = mId$(xIn, arrField_Pos(kMaster_RELMSTRKEY), arrField_Len(kMaster_RELMSTRKEY))
        X36 = mId$(xIn, arrField_Pos(kMaster_RELMSTRREF), arrField_Len(kMaster_RELMSTRREF))
      
        X37 = mId$(xIn, arrField_Pos(kMaster_EXPIRY_LOC), arrField_Len(kMaster_EXPIRY_LOC))
        X38 = mId$(xIn, arrField_Pos(kMaster_OPERATIVE), arrField_Len(kMaster_OPERATIVE))
        X39 = mId$(xIn, arrField_Pos(kMaster_TRANSFER), arrField_Len(kMaster_TRANSFER))
        X40 = mId$(xIn, arrField_Pos(kMaster_REVOLVING), arrField_Len(kMaster_REVOLVING))
        X41 = mId$(xIn, arrField_Pos(kMaster_REV_CUM), arrField_Len(kMaster_REV_CUM))
        X42 = mId$(xIn, arrField_Pos(kMaster_AVAIL_BY), arrField_Len(kMaster_AVAIL_BY))
        X43 = mId$(xIn, arrField_Pos(kMaster_SHIP_FROM), arrField_Len(kMaster_SHIP_FROM))
        X44 = mId$(xIn, arrField_Pos(kMaster_SHIP_TO), arrField_Len(kMaster_SHIP_TO))
        X45 = mId$(xIn, arrField_Pos(kMaster_SHIP_DATE), arrField_Len(kMaster_SHIP_DATE))
        X46 = mId$(xIn, arrField_Pos(kMaster_PART_SHIP), arrField_Len(kMaster_PART_SHIP))
        X47 = mId$(xIn, arrField_Pos(kMaster_TRANS_SHIP), arrField_Len(kMaster_TRANS_SHIP))
        X48 = mId$(xIn, arrField_Pos(kMaster_DFR_APP), arrField_Len(kMaster_DFR_APP))
        X49 = mId$(xIn, arrField_Pos(kMaster_DFR_BEN), arrField_Len(kMaster_DFR_BEN))
      
        xOut = Space$(493)        ' xOut = Space$(387)
        ' Références DOSSIER
        Mid$(xOut, 1, 12) = Format$(X1, "000000000000")             '1
        Mid$(xOut, 13, 3) = Format$(X2, "@@@")                      '2
        Mid$(xOut, 16, 6) = Format$(X3, "000000")                   '3
        If Trim(X4) = "-" Then                                      '4
            Mid$(xOut, 22, 20) = "                    "
        Else
            Mid$(xOut, 22, 20) = Format$(X4, "@@@@@@@@@@@@@@@@@@@@")
        End If
        ' Dates
        dateJma10_Amj X7, DateAMJ                              '7
        Mid$(xOut, 42, 8) = DateAMJ
        If Trim(X8) = "-" Then
            Mid$(xOut, 50, 8) = "00000000"
        Else
            dateJma10_Amj X8, DateAMJ                              '8
            Mid$(xOut, 50, 8) = DateAMJ
        End If
        ' Codes
        Mid$(xOut, 58, 4) = Format$(X9, "@@@@")                     '9
        Mid$(xOut, 62, 1) = " "
        Mid$(xOut, 63, 1) = Format$(X19, "@")                       '19
        Mid$(xOut, 64, 1) = Format$(X20, "@")
        Mid$(xOut, 65, 2) = Format$(X10, "@@")
        Mid$(xOut, 67, 3) = Format$(X11, "@@@")
        Mid$(xOut, 70, 3) = Format$(X12, "@@@")
        Mid$(xOut, 73, 3) = Format$(X13, "000")
        ' Pourcentage et montant garantie ou provision
        If X21 = "" Then
            Mid$(xOut, 76, 7) = "0000000"
        Else
            Mnt = Val(X21)
            Mid$(xOut, 76, 7) = Format$(Abs(Mnt) * 100, "0000000")
        End If
        Mnt = Val(X22)
        If X23 = "ITL" Or X23 = "GRD" Or X23 = "PTE" _
           Or X23 = "ESP" Or X23 = "BEF" Or X23 = "LUF" _
           Or X23 = "JPY" Or X23 = "CFA" Or X23 = "XAF" _
           Or X23 = "XOF" Then
           Mnt = curMaxD(Mnt, 0)
        Else
           Mnt = curMaxD(Mnt, 2)
        End If
        Mid$(xOut, 83, 17) = Format$(Abs(Mnt) * 100, "00000000000000000")
        ' Devise et compte garantie
        If Trim(X23) = "-" Then
            Mid$(xOut, 100, 3) = "   "
        Else
            Mid$(xOut, 100, 3) = Format$(X23, "@@@")
        End If
        If Trim(X24) = "-" Then
            Mid$(xOut, 103, 4) = "    "
        Else
            Mid$(xOut, 103, 4) = Format$(X24, "@@@@")
        End If
        If Trim(X25) = "-" Then
            Mid$(xOut, 107, 6) = "      "
        Else
            Mid$(xOut, 107, 6) = Format$(X25, "@@@@@@")
        End If
        If Trim(X26) = "-" Then
            Mid$(xOut, 113, 3) = "   "
        Else
            Mid$(xOut, 113, 3) = Format$(X26, "@@@")
        End If
        Mid$(xOut, 116, 25) = "                         "
        ' Montant dossier
        Mnt = Val(X14)
        If X15 = "ITL" Or X15 = "GRD" Or X15 = "PTE" _
           Or X15 = "ESP" Or X15 = "BEF" Or X15 = "LUF" _
           Or X15 = "JPY" Or X15 = "CFA" Or X15 = "XAF" _
           Or X15 = "XOF" Then
           Mnt = curMaxD(Mnt, 0)
        Else
           Mnt = curMaxD(Mnt, 2)
        End If
        Mid$(xOut, 141, 17) = Format$(Abs(Mnt) * 100, "00000000000000000")
        Mid$(xOut, 158, 3) = Format$(X15, "@@@")
    
        ' Pourcentages
        If X27 = "" Then
            Mid$(xOut, 161, 7) = "0000000"
        Else
            Mnt = Val(X27)
            Mid$(xOut, 161, 7) = Format$(Abs(Mnt) * 100, "0000000")
        End If
        If X28 = "" Then
            Mid$(xOut, 168, 7) = "0000000"
        Else
            Mnt = Val(X28)
            Mid$(xOut, 168, 7) = Format$(Abs(Mnt) * 100, "0000000")
        End If
        If X29 = "" Then
            Mid$(xOut, 175, 7) = "0000000"
        Else
            Mnt = Val(X29)
            Mid$(xOut, 175, 7) = Format$(Abs(Mnt) * 100, "0000000")
        End If
        If Trim(X30) = "-" Then
            Mid$(xOut, 182, 1) = " "
        Else
            Mid$(xOut, 182, 1) = Format$(X30, "@")
        End If
        ' Outstanding et Liability
        Mnt = Val(X16)
        If X15 = "ITL" Or X15 = "GRD" Or X15 = "PTE" _
           Or X15 = "ESP" Or X15 = "BEF" Or X15 = "LUF" _
           Or X15 = "JPY" Or X15 = "CFA" Or X15 = "XAF" _
           Or X15 = "XOF" Then
           Mnt = curMaxD(Mnt, 0)
        Else
           Mnt = curMaxD(Mnt, 2)
        End If
        Mid$(xOut, 183, 17) = Format$(Abs(Mnt) * 100, "00000000000000000")
        
       ' If mId$(xOut, 16, 6) = "064324" Then
       '     MsgBox "Dossier :" & X3
       ' End If
        
        Mnt = Val(X17)
        If Trim(X18) = "ITL" Or Trim(X18) = "GRD" Or Trim(X18) = "PTE" _
           Or Trim(X18) = "ESP" Or Trim(X18) = "BEF" Or Trim(X18) = "LUF" _
           Or Trim(X18) = "JPY" Or Trim(X18) = "CFA" Or Trim(X18) = "XAF" _
           Or Trim(X18) = "XOF" Then
           Mnt = curMaxD(Mnt, 0)
        Else
           Mnt = curMaxD(Mnt, 2)
        End If
        Mid$(xOut, 200, 17) = Format$(Abs(Mnt) * 100, "00000000000000000")
        Mid$(xOut, 217, 3) = Format$(X18, "@@@")
        ' Les différents intervenants du DOSSIER
        If X31 = "" Then
            Mid$(xOut, 220, 12) = "000000000000"
        Else
            Mid$(xOut, 220, 12) = Format$(X31, "000000000000")
        End If
        Mid$(xOut, 232, 6) = "000000"
        If X32 = "" Then
            Mid$(xOut, 238, 12) = "000000000000"
        Else
            Mid$(xOut, 238, 12) = Format$(X32, "000000000000")
        End If
        Mid$(xOut, 250, 6) = "000000"
        If X33 = "" Then
            Mid$(xOut, 256, 12) = "000000000000"
        Else
            Mid$(xOut, 256, 12) = Format$(X33, "000000000000")
        End If
        Mid$(xOut, 268, 6) = "000000"
        If X34 = "" Then
            Mid$(xOut, 274, 12) = "000000000000"
        Else
            Mid$(xOut, 274, 12) = Format$(X34, "000000000000")
        End If
        Mid$(xOut, 286, 6) = "000000"
        ' Infos de mise à jour DOSSIER
        If Trim(X5) = "-" Then
            Mid$(xOut, 292, 4) = "    "
        Else
            Mid$(xOut, 292, 4) = Format$(X5, "@@@@")
        End If
        If Trim(X6) = "-" Then
            Mid$(xOut, 296, 4) = "    "
        Else
            Mid$(xOut, 296, 4) = Format$(X6, "@@@@")
        End If
        Mid$(xOut, 300, 20) = "                    "
        Mid$(xOut, 320, 8) = "00000000"
        Mid$(xOut, 328, 8) = "00000000"
        If Trim(X35) = "-" Then
            Mid$(xOut, 336, 12) = "000000000000"
        Else
            Mid$(xOut, 336, 12) = Format$(X35, "000000000000")
        End If
        If Trim(X36) = "-" Then
            Mid$(xOut, 348, 20) = "                    "
        Else
            Mid$(xOut, 348, 20) = Format$(X36, "@@@@@@@@@@@@@@@@@@@@")
        End If
        Mid$(xOut, 368, 20) = "                    "
        Mid$(xOut, 388, 29) = Trim(X37)
        Mid$(xOut, 417, 1) = Format$(X38, "@")
        Mid$(xOut, 418, 1) = Format$(X39, "@")
        Mid$(xOut, 419, 1) = Format$(X40, "@")
        Mid$(xOut, 420, 1) = Format$(X41, "@")
        Mid$(xOut, 421, 1) = Format$(X42, "@")
        Mid$(xOut, 422, 30) = Format$(mId$(X43, 1, 30), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
        Mid$(xOut, 452, 30) = Format$(mId$(X44, 1, 30), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
        If Trim(X45) = "-" Then
            Mid$(xOut, 482, 8) = "00000000"
        Else
            dateJma10_Amj X45, DateAMJ
            Mid$(xOut, 482, 8) = DateAMJ
        End If
        Mid$(xOut, 490, 1) = Format$(X46, "@")
        Mid$(xOut, 491, 1) = Format$(X47, "@")
        Mid$(xOut, 492, 1) = Format$(X48, "@")
        Mid$(xOut, 493, 1) = Format$(X49, "@")
      
        xOut_nb = xOut_nb + 1
        Print #2, xOut
    End If
Loop

If xOut_nb <> xIn_nb Then Call MsgBox("enregistrements traités / enregistrements attendus : " & xOut_nb & " / " & xIn_nb, vbCritical, "TIDB2_MASTER")
lstMsg.AddItem "srvTIAS400.TIDB2_Master : " & xOut_nb & " / " & xIn_nb
Close
Exit Sub

Error_Handle:
     lstMsg.AddItem "srvTIAS400.TIDB2_Master : " & xOut_nb & " : " & Error
     lstMsg.AddItem xIn
    Close

End Sub

Public Sub TIDB2_PayDiff(lstMsg As ListBox)
Dim I1 As Integer
On Error GoTo Error_Handle


Open paramTIDB2_PayDiff For Input As #1

TIDB2_Load_Init

'===============

kPayDiff_PAY_AMT = TIDB2_FieldName_Scan("PAY_AMT")
kPayDiff_PAY_AMTCCY = TIDB2_FieldName_Scan("PAY_AMTCCY")
kPayDiff_KEY97 = TIDB2_FieldName_Scan("KEY97")
kPayDiff_PAYEV_KEY = TIDB2_FieldName_Scan("PAYEV_KEY")
kPayDiff_START_DATE = TIDB2_FieldName_Scan("START_DATE")
kPayDiff_VALUE_DAT = TIDB2_FieldName_Scan("VALUE_DAT")
kPayDiff_PERIOD_NO = TIDB2_FieldName_Scan("PERIOD_NO")
kPayDiff_TYPE = TIDB2_FieldName_Scan("TYPE")

'==============================================================

If kPayDiff_PAY_AMT < 0 Then Call MsgBox("champ 'PAY_AMT' non trouvé", vbCritical, "TIDB2_PayDiff"): Exit Sub
If kPayDiff_PAY_AMTCCY < 0 Then Call MsgBox("champ 'PAY_AMTCCY' non trouvé", vbCritical, "TIDB2_PayDiff"): Exit Sub
If kPayDiff_KEY97 < 0 Then Call MsgBox("champ 'KEY97' non trouvé", vbCritical, "TIDB2_PayDiff"): Exit Sub
If kPayDiff_PAYEV_KEY < 0 Then Call MsgBox("champ 'PAYEV_KEY' non trouvé", vbCritical, "TIDB2_PayDiff"): Exit Sub
If kPayDiff_START_DATE < 0 Then Call MsgBox("champ 'START_DATE' non trouvé", vbCritical, "TIDB2_PayDiff"): Exit Sub
If kPayDiff_VALUE_DAT < 0 Then Call MsgBox("champ 'VALUE_DAT' non trouvé", vbCritical, "TIDB2_PayDiff"): Exit Sub
If kPayDiff_PERIOD_NO < 0 Then Call MsgBox("champ 'PERIOD_NO' non trouvé", vbCritical, "TIDB2_PayDiff"): Exit Sub
If kPayDiff_TYPE < 0 Then Call MsgBox("champ 'TYPE' non trouvé", vbCritical, "TIDB2_PayDiff"): Exit Sub

''paramTIDB2_Output = paramAS400IN & "CDPAYW0"
''Open paramTIDB2_Output For Output As #2
Open "\\Fr11024427\As400_IN\CDPAYW0" For Output As #2

xIn_nb = 0
xOut_nb = 0

Do Until EOF(1)
    DoEvents
    Line Input #1, xIn
    If Trim(xIn) <> "" Then
        I1 = InStr(1, xIn, "enregistrement(s) sélectionné(s)")
        If I1 > 0 Then
            xIn_nb = CLng(Val(mId$(xIn, 1, I1 - 1)))
            Exit Do
        End If
        
        X1 = mId$(xIn, arrField_Pos(kPayDiff_PAY_AMT), arrField_Len(kPayDiff_PAY_AMT))
        X2 = mId$(xIn, arrField_Pos(kPayDiff_PAY_AMTCCY), arrField_Len(kPayDiff_PAY_AMTCCY))
        X3 = mId$(xIn, arrField_Pos(kPayDiff_KEY97), arrField_Len(kPayDiff_KEY97))
        X4 = mId$(xIn, arrField_Pos(kPayDiff_PAYEV_KEY), arrField_Len(kPayDiff_PAYEV_KEY))
        X5 = mId$(xIn, arrField_Pos(kPayDiff_START_DATE), arrField_Len(kPayDiff_START_DATE))
        X6 = mId$(xIn, arrField_Pos(kPayDiff_VALUE_DAT), arrField_Len(kPayDiff_VALUE_DAT))
        X7 = mId$(xIn, arrField_Pos(kPayDiff_PERIOD_NO), arrField_Len(kPayDiff_PERIOD_NO))
        X8 = mId$(xIn, arrField_Pos(kPayDiff_TYPE), arrField_Len(kPayDiff_TYPE))
      
        xOut = Space$(96)
    Mid$(xOut, 1, 12) = Format$(X3, "000000000000")
    Mid$(xOut, 13, 12) = Format$(X4, "000000000000")
    Mid$(xOut, 25, 3) = "   "
    Mid$(xOut, 28, 6) = "000000"
    Mid$(xOut, 34, 12) = "000000000000"
    Mid$(xOut, 46, 3) = "   "
    Mid$(xOut, 49, 6) = "000000"
   ' Dates
    dateJma10_Amj X5, DateAMJ
    Mid$(xOut, 55, 8) = DateAMJ
    If Trim(X6) = "-" Then
        Mid$(xOut, 63, 8) = "00000000"
    Else
        dateJma10_Amj X6, DateAMJ
        Mid$(xOut, 63, 8) = DateAMJ
    End If
    If Trim(X7) = "-" Then
        Mid$(xOut, 71, 5) = "00000"
    Else
        Mid$(xOut, 71, 5) = Format$(X7, "00000")
    End If
    If Trim(X8) = "-" Then
        Mid$(xOut, 76, 1) = " "
    Else
        Mid$(xOut, 76, 1) = Format$(X8, "@")
    End If
    Mnt = Val(X1)
    If X2 = "ITL" Or X2 = "GRD" Or X2 = "PTE" _
       Or X2 = "ESP" Or X2 = "BEF" Or X2 = "LUF" _
       Or X2 = "JPY" Or X2 = "CFA" Or X2 = "XAF" _
       Or X2 = "XOF" Then
       Mnt = curMaxD(Mnt, 0)
    Else
       Mnt = curMaxD(Mnt, 2)
    End If
    Mid$(xOut, 77, 17) = Format$(Abs(Mnt) * 100, "00000000000000000")
    Mid$(xOut, 94, 3) = Format$(X2, "@@@")
      
        xOut_nb = xOut_nb + 1
        Print #2, xOut
    End If
Loop

If xOut_nb <> xIn_nb Then Call MsgBox("enregistrements traités / enregistrements attendus : " & xOut_nb & " / " & xIn_nb, vbCritical, "TIDB2_PayDiff")
lstMsg.AddItem "srvTIAS400.TIDB2_PayDiff : " & xOut_nb & " / " & xIn_nb
Close
Exit Sub

Error_Handle:
     lstMsg.AddItem "srvTIAS400.TIDB2_PayDiff : " & xOut_nb & " : " & Error
     lstMsg.AddItem xIn
    Close

End Sub

Public Sub TIDB2_CDOESC(lstMsg As ListBox)
Dim I1 As Integer
Dim MntTx As Double
Dim Mnt_7D As String

On Error GoTo Error_Handle

Open paramTIDB2_CDOESC For Input As #1

TIDB2_Load_Init

'===============

kCDOESC_MASTER_REFNO_PFIX = TIDB2_FieldName_Scan("REFNO_PFIX")
kCDOESC_MASTER_REFNO_SERL = TIDB2_FieldName_Scan("REFNO_SERL")
kCDOESC_EVENT_REFNO_PFIX = TIDB2_FieldName_Scan_Duplicate("REFNO_PFIX", 2)
kCDOESC_EVENT_REFNO_SERL = TIDB2_FieldName_Scan_Duplicate("REFNO_SERL", 2)
kCDOESC_PAY_AMT = TIDB2_FieldName_Scan("PAY_AMT")
kCDOESC_PAY_AMTCCY = TIDB2_FieldName_Scan("PAY_AMTCCY")
kCDOESC_PARTPAYMNT_KEY97 = TIDB2_FieldName_Scan("KEY97")
kCDOESC_START_DATE = TIDB2_FieldName_Scan("START_DATE")
kCDOESC_VALUE_DAT = TIDB2_FieldName_Scan("VALUE_DAT")
kCDOESC_PERIOD_NO = TIDB2_FieldName_Scan("PERIOD_NO")
kCDOESC_TYPE = TIDB2_FieldName_Scan("TYPE")
kCDOESC_REFERENCE = TIDB2_FieldName_Scan("REFERENCE")
kCDOESC_DEAL_PTY = TIDB2_FieldName_Scan("DEAL_PTY")
kCDOESC_DEAL_TYPE = TIDB2_FieldName_Scan("DEAL_TYPE")
kCDOESC_KEY97 = TIDB2_FieldName_Scan_Duplicate("KEY97", 2)
kCDOESC_RATE = TIDB2_FieldName_Scan("RATE")
kCDOESC_DEAL_AMT = TIDB2_FieldName_Scan("DEAL_AMT")
kCDOESC_AMT_CCY = TIDB2_FieldName_Scan("AMT_CCY")
kCDOESC_SPREAD = TIDB2_FieldName_Scan("SPREAD")
kCDOESC_IDB = TIDB2_FieldName_Scan("IDB")
kCDOESC_STARTDATE = TIDB2_FieldName_Scan("STARTDATE")
kCDOESC_MATURITY = TIDB2_FieldName_Scan("MATURITY")
kCDOESC_DISC_AMT = TIDB2_FieldName_Scan("DISC_AMT")
kCDOESC_DISC_CCY = TIDB2_FieldName_Scan("DISC_CCY")
kCDOESC_NET_AMT = TIDB2_FieldName_Scan("NET_AMT")
kCDOESC_NET_CCY = TIDB2_FieldName_Scan("NET_CCY")
kCDOESC_DISC_FOR = TIDB2_FieldName_Scan("DISC_FOR")

'==============================================================

If kCDOESC_MASTER_REFNO_PFIX < 0 Then Call MsgBox("champ 'MASTER.REFNO_PFIX' non trouvé", vbCritical, "TIDB2_CDOESC"): Exit Sub
If kCDOESC_MASTER_REFNO_SERL < 0 Then Call MsgBox("champ 'MASTER.REFNO_SERL' non trouvé", vbCritical, "TIDB2_CDOESC"): Exit Sub
If kCDOESC_EVENT_REFNO_PFIX < 0 Then Call MsgBox("champ 'EVENT.REFNO_PFIX' non trouvé", vbCritical, "TIDB2_CDOESC"): Exit Sub
If kCDOESC_EVENT_REFNO_SERL < 0 Then Call MsgBox("champ 'EVENT.REFNO_SERL' non trouvé", vbCritical, "TIDB2_CDOESC"): Exit Sub
If kCDOESC_PAY_AMT < 0 Then Call MsgBox("champ 'PAY_AMT' non trouvé", vbCritical, "TIDB2_CDOESC"): Exit Sub
If kCDOESC_PAY_AMTCCY < 0 Then Call MsgBox("champ 'PAY_AMTCCY' non trouvé", vbCritical, "TIDB2_CDOESC"): Exit Sub
If kCDOESC_PARTPAYMNT_KEY97 < 0 Then Call MsgBox("champ 'PARTPAYMNT.KEY97' non trouvé", vbCritical, "TIDB2_CDOESC"): Exit Sub
If kCDOESC_START_DATE < 0 Then Call MsgBox("champ 'START_DATE' non trouvé", vbCritical, "TIDB2_CDOESC"): Exit Sub
If kCDOESC_VALUE_DAT < 0 Then Call MsgBox("champ 'VALUE_DAT' non trouvé", vbCritical, "TIDB2_CDOESC"): Exit Sub
If kCDOESC_PERIOD_NO < 0 Then Call MsgBox("champ 'PERIOD_NO' non trouvé", vbCritical, "TIDB2_CDOESC"): Exit Sub
If kCDOESC_TYPE < 0 Then Call MsgBox("champ 'TYPE' non trouvé", vbCritical, "TIDB2_CDOESC"): Exit Sub
If kCDOESC_REFERENCE < 0 Then Call MsgBox("champ 'REFERENCE' non trouvé", vbCritical, "TIDB2_CDOESC"): Exit Sub
If kCDOESC_DEAL_PTY < 0 Then Call MsgBox("champ 'DEAL_PTY' non trouvé", vbCritical, "TIDB2_CDOESC"): Exit Sub
If kCDOESC_DEAL_TYPE < 0 Then Call MsgBox("champ 'DEAL_TYPE' non trouvé", vbCritical, "TIDB2_CDOESC"): Exit Sub
If kCDOESC_KEY97 < 0 Then Call MsgBox("champ 'KEY97' non trouvé", vbCritical, "TIDB2_CDOESC"): Exit Sub
If kCDOESC_RATE < 0 Then Call MsgBox("champ 'RATE' non trouvé", vbCritical, "TIDB2_CDOESC"): Exit Sub
If kCDOESC_DEAL_AMT < 0 Then Call MsgBox("champ 'DEAL_AMT' non trouvé", vbCritical, "TIDB2_CDOESC"): Exit Sub
If kCDOESC_AMT_CCY < 0 Then Call MsgBox("champ 'AMT_CCY' non trouvé", vbCritical, "TIDB2_CDOESC"): Exit Sub
If kCDOESC_SPREAD < 0 Then Call MsgBox("champ 'SPREAD' non trouvé", vbCritical, "TIDB2_CDOESC"): Exit Sub
If kCDOESC_IDB < 0 Then Call MsgBox("champ 'IDB' non trouvé", vbCritical, "TIDB2_CDOESC"): Exit Sub
If kCDOESC_STARTDATE < 0 Then Call MsgBox("champ 'STARTDATE' non trouvé", vbCritical, "TIDB2_CDOESC"): Exit Sub
If kCDOESC_MATURITY < 0 Then Call MsgBox("champ 'MATURITY' non trouvé", vbCritical, "TIDB2_CDOESC"): Exit Sub
If kCDOESC_DISC_AMT < 0 Then Call MsgBox("champ 'DISC_AMT' non trouvé", vbCritical, "TIDB2_CDOESC"): Exit Sub
If kCDOESC_DISC_CCY < 0 Then Call MsgBox("champ 'DISC_CCY' non trouvé", vbCritical, "TIDB2_CDOESC"): Exit Sub
If kCDOESC_NET_AMT < 0 Then Call MsgBox("champ 'NET_AMT' non trouvé", vbCritical, "TIDB2_CDOESC"): Exit Sub
If kCDOESC_NET_CCY < 0 Then Call MsgBox("champ 'NET_CCY' non trouvé", vbCritical, "TIDB2_CDOESC"): Exit Sub
If kCDOESC_DISC_FOR < 0 Then Call MsgBox("champ 'DISC_FOR' non trouvé", vbCritical, "TIDB2_CDOESC"): Exit Sub

''paramTIDB2_Output = paramAS400IN & "CDOESCW0"
''Open paramTIDB2_Output For Output As #2
Open "\\Fr11024427\As400_IN\CDOESCW0" For Output As #2

xIn_nb = 0
xOut_nb = 0

Do Until EOF(1)
    DoEvents
    Line Input #1, xIn
    If Trim(xIn) <> "" Then
        I1 = InStr(1, xIn, "enregistrement(s) sélectionné(s)")
        If I1 > 0 Then
            xIn_nb = CLng(Val(mId$(xIn, 1, I1 - 1)))
            Exit Do
        End If
        
        X1 = mId$(xIn, arrField_Pos(kCDOESC_MASTER_REFNO_PFIX), arrField_Len(kCDOESC_MASTER_REFNO_PFIX))
        X2 = mId$(xIn, arrField_Pos(kCDOESC_MASTER_REFNO_SERL), arrField_Len(kCDOESC_MASTER_REFNO_SERL))
        X3 = mId$(xIn, arrField_Pos(kCDOESC_EVENT_REFNO_PFIX), arrField_Len(kCDOESC_EVENT_REFNO_PFIX))
        X4 = mId$(xIn, arrField_Pos(kCDOESC_EVENT_REFNO_SERL), arrField_Len(kCDOESC_EVENT_REFNO_SERL))
        X5 = mId$(xIn, arrField_Pos(kCDOESC_PAY_AMT), arrField_Len(kCDOESC_PAY_AMT))
        X6 = mId$(xIn, arrField_Pos(kCDOESC_PAY_AMTCCY), arrField_Len(kCDOESC_PAY_AMTCCY))
        X7 = mId$(xIn, arrField_Pos(kCDOESC_PARTPAYMNT_KEY97), arrField_Len(kCDOESC_PARTPAYMNT_KEY97))
        X8 = mId$(xIn, arrField_Pos(kCDOESC_START_DATE), arrField_Len(kCDOESC_START_DATE))
        X9 = mId$(xIn, arrField_Pos(kCDOESC_VALUE_DAT), arrField_Len(kCDOESC_VALUE_DAT))
        X10 = mId$(xIn, arrField_Pos(kCDOESC_PERIOD_NO), arrField_Len(kCDOESC_PERIOD_NO))
        X11 = mId$(xIn, arrField_Pos(kCDOESC_TYPE), arrField_Len(kCDOESC_TYPE))
        X12 = mId$(xIn, arrField_Pos(kCDOESC_REFERENCE), arrField_Len(kCDOESC_REFERENCE))
        X13 = mId$(xIn, arrField_Pos(kCDOESC_DEAL_PTY), arrField_Len(kCDOESC_DEAL_PTY))
        X14 = mId$(xIn, arrField_Pos(kCDOESC_DEAL_TYPE), arrField_Len(kCDOESC_DEAL_TYPE))
        X15 = mId$(xIn, arrField_Pos(kCDOESC_KEY97), arrField_Len(kCDOESC_KEY97))
        X16 = mId$(xIn, arrField_Pos(kCDOESC_RATE), arrField_Len(kCDOESC_RATE))
        X17 = mId$(xIn, arrField_Pos(kCDOESC_DEAL_AMT), arrField_Len(kCDOESC_DEAL_AMT))
        X18 = mId$(xIn, arrField_Pos(kCDOESC_AMT_CCY), arrField_Len(kCDOESC_AMT_CCY))
        X19 = mId$(xIn, arrField_Pos(kCDOESC_SPREAD), arrField_Len(kCDOESC_SPREAD))
        X20 = mId$(xIn, arrField_Pos(kCDOESC_IDB), arrField_Len(kCDOESC_IDB))
        X21 = mId$(xIn, arrField_Pos(kCDOESC_STARTDATE), arrField_Len(kCDOESC_STARTDATE))
        X22 = mId$(xIn, arrField_Pos(kCDOESC_MATURITY), arrField_Len(kCDOESC_MATURITY))
        X23 = mId$(xIn, arrField_Pos(kCDOESC_DISC_AMT), arrField_Len(kCDOESC_DISC_AMT))
        X24 = mId$(xIn, arrField_Pos(kCDOESC_DISC_CCY), arrField_Len(kCDOESC_DISC_CCY))
        X25 = mId$(xIn, arrField_Pos(kCDOESC_NET_AMT), arrField_Len(kCDOESC_NET_AMT))
        X26 = mId$(xIn, arrField_Pos(kCDOESC_NET_CCY), arrField_Len(kCDOESC_NET_CCY))
        X27 = mId$(xIn, arrField_Pos(kCDOESC_DISC_FOR), arrField_Len(kCDOESC_DISC_FOR))
      
        xOut = Space$(215)
        
    Mid$(xOut, 1, 12) = Format$(X15, "000000000000")
    Mid$(xOut, 13, 3) = Format$(X3, "@@@")
    Mid$(xOut, 16, 6) = Format$(X4, "000000")
    Mid$(xOut, 22, 3) = Format$(X1, "@@@")
    Mid$(xOut, 25, 6) = Format$(X2, "000000")
    Mid$(xOut, 31, 12) = Format$(X7, "000000000000")

   ' Dates
   
    If Trim(X8) = "-" Then
        Mid$(xOut, 43, 8) = "00000000"
    Else
        dateJma10_Amj X8, DateAMJ
        Mid$(xOut, 43, 8) = DateAMJ
    End If
    If Trim(X9) = "-" Then
        Mid$(xOut, 51, 8) = "00000000"
    Else
        dateJma10_Amj X9, DateAMJ
        Mid$(xOut, 51, 8) = DateAMJ
    End If
    If Trim(X21) = "-" Then
        Mid$(xOut, 159, 8) = "00000000"
    Else
        dateJma10_Amj X21, DateAMJ
        Mid$(xOut, 159, 8) = DateAMJ
    End If
    If Trim(X22) = "-" Then
        Mid$(xOut, 167, 8) = "00000000"
    Else
        dateJma10_Amj X22, DateAMJ
        Mid$(xOut, 167, 8) = DateAMJ
    End If
    
    If Trim(X10) = "-" Then
        Mid$(xOut, 59, 5) = "00000"
    Else
        Mid$(xOut, 59, 5) = Format$(X10, "00000")
    End If
    If Trim(X11) = "-" Then
        Mid$(xOut, 64, 1) = " "
    Else
        Mid$(xOut, 64, 1) = Format$(X11, "@")
    End If
    
    ' Montants
    
    Mnt = Val(X5)
    If X6 = "ITL" Or X6 = "GRD" Or X6 = "PTE" _
       Or X6 = "ESP" Or X6 = "BEF" Or X6 = "LUF" _
       Or X6 = "JPY" Or X6 = "CFA" Or X6 = "XAF" _
       Or X6 = "XOF" Then
       Mnt = curMaxD(Mnt, 0)
    Else
       Mnt = curMaxD(Mnt, 2)
    End If
    Mid$(xOut, 65, 17) = Format$(Abs(Mnt) * 100, "00000000000000000")
    Mid$(xOut, 82, 3) = Format$(X6, "@@@")
    
    If Trim(X12) = "-" Then
        Mid$(xOut, 85, 16) = "                 "
    Else
        Mid$(xOut, 85, 16) = Format$(X12, "@@@@@@@@@@@@@@@@")
    End If
    Mid$(xOut, 101, 12) = Format$(X13, "000000000000")
    If Trim(X14) = "-" Then
        Mid$(xOut, 113, 1) = "   "
    Else
        Mid$(xOut, 113, 1) = Format$(X14, "@@@")
    End If
    
    Mnt = Val(X17)
    If X18 = "ITL" Or X18 = "GRD" Or X18 = "PTE" _
       Or X18 = "ESP" Or X18 = "BEF" Or X18 = "LUF" _
       Or X18 = "JPY" Or X18 = "CFA" Or X18 = "XAF" _
       Or X18 = "XOF" Then
       Mnt = curMaxD(Mnt, 0)
    Else
       Mnt = curMaxD(Mnt, 2)
    End If
    Mid$(xOut, 127, 17) = Format$(Abs(Mnt) * 100, "00000000000000000")
    Mid$(xOut, 144, 3) = Format$(X18, "@@@")
    If Trim(X20) = "-" Then
        Mid$(xOut, 158, 1) = " "
    Else
        Mid$(xOut, 158, 1) = Format$(X20, "@")
    End If
    
    Mnt = Val(X23)
    If X24 = "ITL" Or X24 = "GRD" Or X24 = "PTE" _
       Or X24 = "ESP" Or X24 = "BEF" Or X24 = "LUF" _
       Or X24 = "JPY" Or X24 = "CFA" Or X24 = "XAF" _
       Or X24 = "XOF" Then
       Mnt = curMaxD(Mnt, 0)
    Else
       Mnt = curMaxD(Mnt, 2)
    End If
    Mid$(xOut, 175, 17) = Format$(Abs(Mnt) * 100, "00000000000000000")
    Mid$(xOut, 192, 3) = Format$(X24, "@@@")
    Mnt = Val(X25)
    If X26 = "ITL" Or X26 = "GRD" Or X26 = "PTE" _
       Or X26 = "ESP" Or X26 = "BEF" Or X26 = "LUF" _
       Or X26 = "JPY" Or X26 = "CFA" Or X26 = "XAF" _
       Or X26 = "XOF" Then
       Mnt = curMaxD(Mnt, 0)
    Else
       Mnt = curMaxD(Mnt, 2)
    End If
    Mid$(xOut, 195, 17) = Format$(Abs(Mnt) * 100, "00000000000000000")
    Mid$(xOut, 212, 3) = Format$(X26, "@@@")
      
    If Trim(X27) = "-" Then
        Mid$(xOut, 215, 1) = " "
    Else
        Mid$(xOut, 215, 1) = Format$(X27, "@")
    End If
          
    ' Taux / Marge
    If Trim(X16) = "-" Then
       MntTx = 0
    Else
       Mnt_7D = mId$(X16, 7, 7)
       MntTx = Val(Mnt_7D) / 10000000
       MntTx = MntTx + Val(X16)
    End If
    Mid$(xOut, 116, 11) = Format$(Abs(MntTx * 10000000), "00000000000")
    
    If Trim(X19) = "-" Then
       MntTx = 0
    Else
       Mnt_7D = mId$(X19, 7, 7)
       MntTx = Val(Mnt_7D) / 10000000
       MntTx = MntTx + Val(X19)
    End If
    Mid$(xOut, 147, 11) = Format$(Abs(MntTx * 10000000), "00000000000")
         
    xOut_nb = xOut_nb + 1
    Print #2, xOut
    
    End If
Loop

If xOut_nb <> xIn_nb Then Call MsgBox("enregistrements traités / enregistrements attendus : " & xOut_nb & " / " & xIn_nb, vbCritical, "TIDB2_CDOESC")
lstMsg.AddItem "srvTIAS400.TIDB2_CDOESC : " & xOut_nb & " / " & xIn_nb
Close
Exit Sub

Error_Handle:
     lstMsg.AddItem "srvTIAS400.TIDB2_CDOESC : " & xOut_nb & " : " & Error
     lstMsg.AddItem xIn
    Close

End Sub

Public Sub TIDB2_CDOUTI(lstMsg As ListBox)
Dim I1 As Integer
On Error GoTo Error_Handle

Open paramTIDB2_CDOUTI For Input As #1

TIDB2_Load_Init

'===============

kCDOUTI_MASTER_REFNO_PFIX = TIDB2_FieldName_Scan("REFNO_PFIX")
kCDOUTI_MASTER_REFNO_SERL = TIDB2_FieldName_Scan("REFNO_SERL")
kCDOUTI_EVENT_REFNO_PFIX = TIDB2_FieldName_Scan_Duplicate("REFNO_PFIX", 2)
kCDOUTI_EVENT_REFNO_SERL = TIDB2_FieldName_Scan_Duplicate("REFNO_SERL", 2)
kCDOUTI_KEY97 = TIDB2_FieldName_Scan("KEY97")
kCDOUTI_MIXEDPAY = TIDB2_FieldName_Scan("MIXEDPAY")
kCDOUTI_PRSPTY_PTY = TIDB2_FieldName_Scan("PRSPTY_PTY")
kCDOUTI_PRES_DATE = TIDB2_FieldName_Scan("PRES_DATE")
kCDOUTI_PRESAMT = TIDB2_FieldName_Scan("PRESAMT")
kCDOUTI_PRES_CCY = TIDB2_FieldName_Scan("PRES_CCY")
kCDOUTI_SENT_DATE = TIDB2_FieldName_Scan("SENT_DATE")
kCDOUTI_DOC_COUNT = TIDB2_FieldName_Scan("DOC_COUNT")
kCDOUTI_HOLD_DOC = TIDB2_FieldName_Scan("HOLD_DOC")
kCDOUTI_DOCSINORDR = TIDB2_FieldName_Scan("DOCSINORDR")

'==============================================================

If kCDOUTI_MASTER_REFNO_PFIX < 0 Then Call MsgBox("champ 'MASTER.REFNO_PFIX' non trouvé", vbCritical, "TIDB2_CDOUTI"): Exit Sub
If kCDOUTI_MASTER_REFNO_SERL < 0 Then Call MsgBox("champ 'MASTER.REFNO_SERL' non trouvé", vbCritical, "TIDB2_CDOUTI"): Exit Sub
If kCDOUTI_EVENT_REFNO_PFIX < 0 Then Call MsgBox("champ 'EVENT.REFNO_PFIX' non trouvé", vbCritical, "TIDB2_CDOUTI"): Exit Sub
If kCDOUTI_EVENT_REFNO_SERL < 0 Then Call MsgBox("champ 'EVENT.REFNO_SERL' non trouvé", vbCritical, "TIDB2_CDOUTI"): Exit Sub
If kCDOUTI_KEY97 < 0 Then Call MsgBox("champ 'KEY97' non trouvé", vbCritical, "TIDB2_CDOUTI"): Exit Sub
If kCDOUTI_MIXEDPAY < 0 Then Call MsgBox("champ 'MIXEDPAY' non trouvé", vbCritical, "TIDB2_CDOUTI"): Exit Sub
If kCDOUTI_PRSPTY_PTY < 0 Then Call MsgBox("champ 'PRSPTY_PTY' non trouvé", vbCritical, "TIDB2_CDOUTI"): Exit Sub
If kCDOUTI_PRES_DATE < 0 Then Call MsgBox("champ 'PRES_DATE' non trouvé", vbCritical, "TIDB2_CDOUTI"): Exit Sub
If kCDOUTI_PRESAMT < 0 Then Call MsgBox("champ 'PRESAMT' non trouvé", vbCritical, "TIDB2_CDOUTI"): Exit Sub
If kCDOUTI_PRES_CCY < 0 Then Call MsgBox("champ 'PRES_CCY' non trouvé", vbCritical, "TIDB2_CDOUTI"): Exit Sub
If kCDOUTI_SENT_DATE < 0 Then Call MsgBox("champ 'SENT_DATE' non trouvé", vbCritical, "TIDB2_CDOUTI"): Exit Sub
If kCDOUTI_DOC_COUNT < 0 Then Call MsgBox("champ 'DOC_COUNT' non trouvé", vbCritical, "TIDB2_CDOUTI"): Exit Sub
If kCDOUTI_HOLD_DOC < 0 Then Call MsgBox("champ 'HOLD_DOC' non trouvé", vbCritical, "TIDB2_CDOUTI"): Exit Sub
If kCDOUTI_DOCSINORDR < 0 Then Call MsgBox("champ 'DOCSINORDR' non trouvé", vbCritical, "TIDB2_CDOUTI"): Exit Sub

''paramTIDB2_Output = paramAS400IN & "CDOUTIW0 ""
''Open paramTIDB2_Output For Output As #2
Open "\\Fr11024427\As400_IN\CDOUTIW0" For Output As #2

xIn_nb = 0
xOut_nb = 0

Do Until EOF(1)
    DoEvents
    Line Input #1, xIn
    If Trim(xIn) <> "" Then
        I1 = InStr(1, xIn, "enregistrement(s) sélectionné(s)")
        If I1 > 0 Then
            xIn_nb = CLng(Val(mId$(xIn, 1, I1 - 1)))
            Exit Do
        End If
        
        X1 = mId$(xIn, arrField_Pos(kCDOUTI_MASTER_REFNO_PFIX), arrField_Len(kCDOUTI_MASTER_REFNO_PFIX))
        X2 = mId$(xIn, arrField_Pos(kCDOUTI_MASTER_REFNO_SERL), arrField_Len(kCDOUTI_MASTER_REFNO_SERL))
        X3 = mId$(xIn, arrField_Pos(kCDOUTI_EVENT_REFNO_PFIX), arrField_Len(kCDOUTI_EVENT_REFNO_PFIX))
        X4 = mId$(xIn, arrField_Pos(kCDOUTI_EVENT_REFNO_SERL), arrField_Len(kCDOUTI_EVENT_REFNO_SERL))
        X5 = mId$(xIn, arrField_Pos(kCDOUTI_KEY97), arrField_Len(kCDOUTI_KEY97))
        X6 = mId$(xIn, arrField_Pos(kCDOUTI_MIXEDPAY), arrField_Len(kCDOUTI_MIXEDPAY))
        X7 = mId$(xIn, arrField_Pos(kCDOUTI_PRSPTY_PTY), arrField_Len(kCDOUTI_PRSPTY_PTY))
        X8 = mId$(xIn, arrField_Pos(kCDOUTI_PRES_DATE), arrField_Len(kCDOUTI_PRES_DATE))
        X9 = mId$(xIn, arrField_Pos(kCDOUTI_PRESAMT), arrField_Len(kCDOUTI_PRESAMT))
        X10 = mId$(xIn, arrField_Pos(kCDOUTI_PRES_CCY), arrField_Len(kCDOUTI_PRES_CCY))
        X11 = mId$(xIn, arrField_Pos(kCDOUTI_SENT_DATE), arrField_Len(kCDOUTI_SENT_DATE))
        X12 = mId$(xIn, arrField_Pos(kCDOUTI_DOC_COUNT), arrField_Len(kCDOUTI_DOC_COUNT))
        X13 = mId$(xIn, arrField_Pos(kCDOUTI_HOLD_DOC), arrField_Len(kCDOUTI_HOLD_DOC))
        X14 = mId$(xIn, arrField_Pos(kCDOUTI_DOCSINORDR), arrField_Len(kCDOUTI_DOCSINORDR))
      
        xOut = Space$(104)
        
    Mid$(xOut, 1, 12) = Format$(X5, "000000000000")
    Mid$(xOut, 13, 3) = Format$(X3, "@@@")
    Mid$(xOut, 16, 6) = Format$(X4, "000000")
    Mid$(xOut, 22, 3) = Format$(X1, "@@@")
    Mid$(xOut, 25, 6) = Format$(X2, "000000")

    If Trim(X6) = "-" Then
        Mid$(xOut, 31, 1) = " "
    Else
        Mid$(xOut, 31, 1) = Format$(X6, "@")
    End If
    If Trim(X7) = "-" Then
        Mid$(xOut, 32, 3) = "   "
    Else
        Mid$(xOut, 32, 3) = mId$(X7, 1, 3)
    End If
    
   ' Dates
    If Trim(X8) = "-" Then
        Mid$(xOut, 35, 8) = "00000000"
    Else
        dateJma10_Amj X8, DateAMJ
        Mid$(xOut, 35, 8) = DateAMJ
    End If
    If Trim(X11) = "-" Then
        Mid$(xOut, 63, 8) = "00000000"
    Else
        dateJma10_Amj X11, DateAMJ
        Mid$(xOut, 63, 8) = DateAMJ
    End If
    
    ' Montants
    Mnt = Val(X9)
    If X10 = "ITL" Or X10 = "GRD" Or X10 = "PTE" _
       Or X10 = "ESP" Or X10 = "BEF" Or X10 = "LUF" _
       Or X10 = "JPY" Or X10 = "CFA" Or X10 = "XAF" _
       Or X10 = "XOF" Then
       Mnt = curMaxD(Mnt, 0)
    Else
       Mnt = curMaxD(Mnt, 2)
    End If
    Mid$(xOut, 43, 17) = Format$(Abs(Mnt) * 100, "00000000000000000")
    Mid$(xOut, 60, 3) = Format$(X10, "@@@")
    
    If Trim(X12) = "-" Then
        Mid$(xOut, 71, 5) = "00000"
    Else
        Mid$(xOut, 71, 5) = Format$(X12, "00000")
    End If
    
    If Trim(X13) = "-" Then
        Mid$(xOut, 76, 1) = " "
    Else
        Mid$(xOut, 76, 1) = Format$(X13, "@")
    End If
    If Trim(X14) = "-" Then
        Mid$(xOut, 77, 1) = " "
    Else
        Mid$(xOut, 77, 1) = Format$(X14, "@")
    End If
          
    Mid$(xOut, 78, 20) = "                    "
    Mid$(xOut, 98, 7) = "0000000"
          
    xOut_nb = xOut_nb + 1
    Print #2, xOut
    
    End If
Loop

If xOut_nb <> xIn_nb Then Call MsgBox("enregistrements traités / enregistrements attendus : " & xOut_nb & " / " & xIn_nb, vbCritical, "TIDB2_CDOUTI")
lstMsg.AddItem "srvTIAS400.TIDB2_CDOUTI : " & xOut_nb & " / " & xIn_nb
Close
Exit Sub

Error_Handle:
    lstMsg.AddItem "srvTIAS400.TIDB2_CDOUTI : " & xOut_nb & " : " & Error
    lstMsg.AddItem xIn
    Close

End Sub

Public Sub TIDB2_CDOFRS(lstMsg As ListBox)
' Extraction des frais banque Emettrice et Notificatrice en devises du Trade Innovation
Dim I1 As Integer
On Error GoTo Error_Handle


Open paramTIDB2_CDOFRS For Input As #1

TIDB2_Load_Init

'===============

kCDOFRS_KEY97 = TIDB2_FieldName_Scan("KEY97")
kCDOFRS_VALUEDATE = TIDB2_FieldName_Scan("VALUEDATE")
kCDOFRS_TRAN_CODE = TIDB2_FieldName_Scan("TRAN_CODE")
kCDOFRS_AMOUNT = TIDB2_FieldName_Scan("AMOUNT")
kCDOFRS_CCY = TIDB2_FieldName_Scan("CCY")
kCDOFRS_CHARGE = TIDB2_FieldName_Scan("CHARGE")
kCDOFRS_POSTINGTYP = TIDB2_FieldName_Scan("POSTINGTYP")
kCDOFRS_CHG_SCH = TIDB2_FieldName_Scan("CHG_SCH")
kCDOFRS_CHG_TYPE = TIDB2_FieldName_Scan("CHG_TYPE")
kCDOFRS_CH_CODE = TIDB2_FieldName_Scan("CH_CODE")
kCDOFRS_STATUS = TIDB2_FieldName_Scan("STATUS")
kCDOFRS_CHG_FOR = TIDB2_FieldName_Scan("CHG_FOR")

'==============================================================

If kCDOFRS_KEY97 < 0 Then Call MsgBox("champ 'KEY97' non trouvé", vbCritical, "TIDB2_CDOFRS"): Exit Sub
If kCDOFRS_VALUEDATE < 0 Then Call MsgBox("champ 'VALUEDATE' non trouvé", vbCritical, "TIDB2_CDOFRS"): Exit Sub
If kCDOFRS_TRAN_CODE < 0 Then Call MsgBox("champ 'TRAN_CODE' non trouvé", vbCritical, "TIDB2_CDOFRS"): Exit Sub
If kCDOFRS_AMOUNT < 0 Then Call MsgBox("champ 'AMOUNT' non trouvé", vbCritical, "TIDB2_CDOFRS"): Exit Sub
If kCDOFRS_CCY < 0 Then Call MsgBox("champ 'CCY' non trouvé", vbCritical, "TIDB2_CDOFRS"): Exit Sub
If kCDOFRS_CHARGE < 0 Then Call MsgBox("champ 'CHARGE' non trouvé", vbCritical, "TIDB2_CDOFRS"): Exit Sub
If kCDOFRS_POSTINGTYP < 0 Then Call MsgBox("champ 'POSTINGTYP' non trouvé", vbCritical, "TIDB2_CDOFRS"): Exit Sub
If kCDOFRS_CHG_SCH < 0 Then Call MsgBox("champ 'CHG_SCH' non trouvé", vbCritical, "TIDB2_CDOFRS"): Exit Sub
If kCDOFRS_CHG_TYPE < 0 Then Call MsgBox("champ 'CHG_TYPE' non trouvé", vbCritical, "TIDB2_CDOFRS"): Exit Sub
If kCDOFRS_CH_CODE < 0 Then Call MsgBox("champ 'CH_CODE' non trouvé", vbCritical, "TIDB2_CDOFRS"): Exit Sub
If kCDOFRS_STATUS < 0 Then Call MsgBox("champ 'STATUS' non trouvé", vbCritical, "TIDB2_CDOFRS"): Exit Sub
If kCDOFRS_CHG_FOR < 0 Then Call MsgBox("champ 'CHG_FOR' non trouvé", vbCritical, "TIDB2_CDOFRS"): Exit Sub

''paramTIDB2_Output = paramAS400IN & "CDCPEW0"
''Open paramTIDB2_Output For Output As #2
Open "\\Fr11024427\As400_IN\CDOFRSW0" For Output As #2

xIn_nb = 0
xOut_nb = 0

Do Until EOF(1)
    DoEvents
    Line Input #1, xIn
    If Trim(xIn) <> "" Then
        I1 = InStr(1, xIn, "enregistrement(s) sélectionné(s)")
        If I1 > 0 Then
            xIn_nb = CLng(Val(mId$(xIn, 1, I1 - 1)))
            Exit Do
        End If
        
        X1 = mId$(xIn, arrField_Pos(kCDOFRS_KEY97), arrField_Len(kCDOFRS_KEY97))
        X2 = mId$(xIn, arrField_Pos(kCDOFRS_VALUEDATE), arrField_Len(kCDOFRS_VALUEDATE))
        X3 = mId$(xIn, arrField_Pos(kCDOFRS_TRAN_CODE), arrField_Len(kCDOFRS_TRAN_CODE))
        X4 = mId$(xIn, arrField_Pos(kCDOFRS_AMOUNT), arrField_Len(kCDOFRS_AMOUNT))
        X5 = mId$(xIn, arrField_Pos(kCDOFRS_CCY), arrField_Len(kCDOFRS_CCY))
        X6 = mId$(xIn, arrField_Pos(kCDOFRS_CHARGE), arrField_Len(kCDOFRS_CHARGE))
        X7 = mId$(xIn, arrField_Pos(kCDOFRS_POSTINGTYP), arrField_Len(kCDOFRS_POSTINGTYP))
        X8 = mId$(xIn, arrField_Pos(kCDOFRS_CHG_SCH), arrField_Len(kCDOFRS_CHG_SCH))
        X9 = mId$(xIn, arrField_Pos(kCDOFRS_CHG_TYPE), arrField_Len(kCDOFRS_CHG_TYPE))
        X10 = mId$(xIn, arrField_Pos(kCDOFRS_CH_CODE), arrField_Len(kCDOFRS_CH_CODE))
        X11 = mId$(xIn, arrField_Pos(kCDOFRS_STATUS), arrField_Len(kCDOFRS_STATUS))
        X12 = mId$(xIn, arrField_Pos(kCDOFRS_CHG_FOR), arrField_Len(kCDOFRS_CHG_FOR))
    
        xOut = Space$(135)
    Mid$(xOut, 1, 12) = Format$(X1, "000000000000")
    Mid$(xOut, 13, 12) = "000000000000"
    Mid$(xOut, 25, 3) = "   "
    Mid$(xOut, 28, 6) = "000000"
    Mid$(xOut, 34, 12) = "000000000000"
    Mid$(xOut, 46, 3) = "   "
    Mid$(xOut, 49, 6) = "000000"
   ' Dates
    dateJma10_Amj X2, DateAMJ
    Mid$(xOut, 55, 8) = DateAMJ
    
    Mid$(xOut, 63, 3) = Format$(X3, "010")
   
   ' Devise
    Mid$(xOut, 83, 3) = Format$(X5, "@@@")
   ' Montant
    Mnt = Val(X4)
    If X5 = "ITL" Or X5 = "GRD" Or X5 = "PTE" _
       Or X5 = "ESP" Or X5 = "BEF" Or X5 = "LUF" _
       Or X5 = "JPY" Or X5 = "CFA" Or X5 = "XAF" _
       Or X5 = "XOF" Then
       Mnt = curMaxD(Mnt, 0)
    Else
       Mnt = curMaxD(Mnt, 2)
    End If
    Mid$(xOut, 66, 17) = Format$(Abs(Mnt) * 100, "00000000000000000")
    Mid$(xOut, 86, 12) = Format$(X6, "000000000000")
    Mid$(xOut, 98, 10) = Format$(X7, "@@@@@@@@@@")
    Mid$(xOut, 108, 12) = Format$(X8, "000000000000")
    Mid$(xOut, 120, 12) = Format$(X9, "000000000000")
    Mid$(xOut, 132, 2) = Format$(X10, "@@")
    Mid$(xOut, 134, 1) = Format$(X11, "@")
    Mid$(xOut, 135, 1) = Format$(X12, "@")
      
        xOut_nb = xOut_nb + 1
        Print #2, xOut
    End If
Loop

If xOut_nb <> xIn_nb Then Call MsgBox("enregistrements traités / enregistrements attendus : " & xOut_nb & " / " & xIn_nb, vbCritical, "TIDB2_CDOFRS")
lstMsg.AddItem "srvTIAS400.TIDB2_CDOFRS : " & xOut_nb & " / " & xIn_nb
Close
Exit Sub

Error_Handle:
     lstMsg.AddItem "srvTIAS400.TIDB2_CDOFRS : " & xOut_nb & " : " & Error
     lstMsg.AddItem xIn
    Close

End Sub

Public Sub TIDB2_ComEnc(lstMsg As ListBox)
' Extraction des commissions encaissées en devises du Trade Innovation
Dim I1 As Integer
On Error GoTo Error_Handle


Open paramTIDB2_ComEnc For Input As #1

TIDB2_Load_Init

'===============

kComEnc_KEY97 = TIDB2_FieldName_Scan("KEY97")
kComEnc_VALUEDATE = TIDB2_FieldName_Scan("VALUEDATE")
kComEnc_TRAN_CODE = TIDB2_FieldName_Scan("TRAN_CODE")
kComEnc_AMOUNT = TIDB2_FieldName_Scan("AMOUNT")
kComEnc_CCY = TIDB2_FieldName_Scan("CCY")
kComEnc_CHARGE = TIDB2_FieldName_Scan("CHARGE")
kComEnc_POSTINGTYP = TIDB2_FieldName_Scan("POSTINGTYP")
kComEnc_CHG_SCH = TIDB2_FieldName_Scan("CHG_SCH")
kComEnc_CHG_TYPE = TIDB2_FieldName_Scan("CHG_TYPE")
kComEnc_CH_CODE = TIDB2_FieldName_Scan("CH_CODE")
kComEnc_STATUS = TIDB2_FieldName_Scan("STATUS")
kComEnc_CHG_FOR = TIDB2_FieldName_Scan("CHG_FOR")

'==============================================================

If kComEnc_KEY97 < 0 Then Call MsgBox("champ 'KEY97' non trouvé", vbCritical, "TIDB2_ComEnc"): Exit Sub
If kComEnc_VALUEDATE < 0 Then Call MsgBox("champ 'VALUEDATE' non trouvé", vbCritical, "TIDB2_ComEnc"): Exit Sub
If kComEnc_TRAN_CODE < 0 Then Call MsgBox("champ 'TRAN_CODE' non trouvé", vbCritical, "TIDB2_ComEnc"): Exit Sub
If kComEnc_AMOUNT < 0 Then Call MsgBox("champ 'AMOUNT' non trouvé", vbCritical, "TIDB2_ComEnc"): Exit Sub
If kComEnc_CCY < 0 Then Call MsgBox("champ 'CCY' non trouvé", vbCritical, "TIDB2_ComEnc"): Exit Sub
If kComEnc_CHARGE < 0 Then Call MsgBox("champ 'CHARGE' non trouvé", vbCritical, "TIDB2_ComEnc"): Exit Sub
If kComEnc_POSTINGTYP < 0 Then Call MsgBox("champ 'POSTINGTYP' non trouvé", vbCritical, "TIDB2_ComEnc"): Exit Sub
If kComEnc_CHG_SCH < 0 Then Call MsgBox("champ 'CHG_SCH' non trouvé", vbCritical, "TIDB2_ComEnc"): Exit Sub
If kComEnc_CHG_TYPE < 0 Then Call MsgBox("champ 'CHG_TYPE' non trouvé", vbCritical, "TIDB2_ComEnc"): Exit Sub
If kComEnc_CH_CODE < 0 Then Call MsgBox("champ 'CH_CODE' non trouvé", vbCritical, "TIDB2_ComEnc"): Exit Sub
If kComEnc_STATUS < 0 Then Call MsgBox("champ 'STATUS' non trouvé", vbCritical, "TIDB2_ComEnc"): Exit Sub
If kComEnc_CHG_FOR < 0 Then Call MsgBox("champ 'CHG_FOR' non trouvé", vbCritical, "TIDB2_ComEnc"): Exit Sub

''paramTIDB2_Output = paramAS400IN & "CDCPEW0"
''Open paramTIDB2_Output For Output As #2
Open "\\Fr11024427\As400_IN\CDCPEW0" For Output As #2

xIn_nb = 0
xOut_nb = 0

Do Until EOF(1)
    DoEvents
    Line Input #1, xIn
    If Trim(xIn) <> "" Then
        I1 = InStr(1, xIn, "enregistrement(s) sélectionné(s)")
        If I1 > 0 Then
            xIn_nb = CLng(Val(mId$(xIn, 1, I1 - 1)))
            Exit Do
        End If
        
        X1 = mId$(xIn, arrField_Pos(kComEnc_KEY97), arrField_Len(kComEnc_KEY97))
        X2 = mId$(xIn, arrField_Pos(kComEnc_VALUEDATE), arrField_Len(kComEnc_VALUEDATE))
        X3 = mId$(xIn, arrField_Pos(kComEnc_TRAN_CODE), arrField_Len(kComEnc_TRAN_CODE))
        X4 = mId$(xIn, arrField_Pos(kComEnc_AMOUNT), arrField_Len(kComEnc_AMOUNT))
        X5 = mId$(xIn, arrField_Pos(kComEnc_CCY), arrField_Len(kComEnc_CCY))
        X6 = mId$(xIn, arrField_Pos(kComEnc_CHARGE), arrField_Len(kComEnc_CHARGE))
        X7 = mId$(xIn, arrField_Pos(kComEnc_POSTINGTYP), arrField_Len(kComEnc_POSTINGTYP))
        X8 = mId$(xIn, arrField_Pos(kComEnc_CHG_SCH), arrField_Len(kComEnc_CHG_SCH))
        X9 = mId$(xIn, arrField_Pos(kComEnc_CHG_TYPE), arrField_Len(kComEnc_CHG_TYPE))
        X10 = mId$(xIn, arrField_Pos(kComEnc_CH_CODE), arrField_Len(kComEnc_CH_CODE))
        X11 = mId$(xIn, arrField_Pos(kComEnc_STATUS), arrField_Len(kComEnc_STATUS))
        X12 = mId$(xIn, arrField_Pos(kComEnc_CHG_FOR), arrField_Len(kComEnc_CHG_FOR))
    
        xOut = Space$(135)
        
    Mid$(xOut, 1, 12) = Format$(X1, "000000000000")
    Mid$(xOut, 13, 12) = "000000000000"
    Mid$(xOut, 25, 3) = "   "
    Mid$(xOut, 28, 6) = "000000"
    Mid$(xOut, 34, 12) = "000000000000"
    Mid$(xOut, 46, 3) = "   "
    Mid$(xOut, 49, 6) = "000000"
   ' Dates
    dateJma10_Amj X2, DateAMJ
    Mid$(xOut, 55, 8) = DateAMJ
    
    Mid$(xOut, 63, 3) = Format$(X3, "010")
   
   ' Devise
    Mid$(xOut, 83, 3) = Format$(X5, "@@@")
   ' Montant
    Mnt = Val(X4)
    If X5 = "ITL" Or X5 = "GRD" Or X5 = "PTE" _
       Or X5 = "ESP" Or X5 = "BEF" Or X5 = "LUF" _
       Or X5 = "JPY" Or X5 = "CFA" Or X5 = "XAF" _
       Or X5 = "XOF" Then
       Mnt = curMaxD(Mnt, 0)
    Else
       Mnt = curMaxD(Mnt, 2)
    End If
    Mid$(xOut, 66, 17) = Format$(Abs(Mnt) * 100, "00000000000000000")
    Mid$(xOut, 86, 12) = Format$(X6, "000000000000")
    Mid$(xOut, 98, 10) = Format$(X7, "@@@@@@@@@@")
    Mid$(xOut, 108, 12) = Format$(X8, "000000000000")
    Mid$(xOut, 120, 12) = Format$(X9, "000000000000")
    Mid$(xOut, 132, 2) = Format$(X10, "@@")
    Mid$(xOut, 134, 1) = Format$(X11, "@")
    Mid$(xOut, 135, 1) = Format$(X12, "@")
      
        xOut_nb = xOut_nb + 1
        Print #2, xOut
    End If
Loop

If xOut_nb <> xIn_nb Then Call MsgBox("enregistrements traités / enregistrements attendus : " & xOut_nb & " / " & xIn_nb, vbCritical, "TIDB2_ComEnc")
lstMsg.AddItem "srvTIAS400.TIDB2_ComEnc : " & xOut_nb & " / " & xIn_nb
Close
Exit Sub

Error_Handle:
     lstMsg.AddItem "srvTIAS400.TIDB2_ComEnc : " & xOut_nb & " : " & Error
     lstMsg.AddItem xIn
    Close

End Sub


Public Sub TIDB2_RegTrans(lstMsg As ListBox)
' Extraction des règles spécifiques de commissions pour les transactions du Trade Innovation
Dim I1 As Integer
On Error GoTo Error_Handle

Open paramTIDB2_RegTrans For Input As #1

TIDB2_Load_Init

'===============

kRegTrans_KEY97 = TIDB2_FieldName_Scan("KEY97")
kRegTrans_SCHED_TYPE = TIDB2_FieldName_Scan("SCHED_TYPE")
kRegTrans_CHG_TYPE = TIDB2_FieldName_Scan("CHG_TYPE")
kRegTrans_T1_NO = TIDB2_FieldName_Scan("T1_NO")
kRegTrans_T1_UNIT = TIDB2_FieldName_Scan("T1_UNIT")
kRegTrans_T1_PERCENT = TIDB2_FieldName_Scan("T1_PERCENT")
kRegTrans_TIERAMT1 = TIDB2_FieldName_Scan("TIERAMT1")
kRegTrans_T2_NO = TIDB2_FieldName_Scan("T2_NO")
kRegTrans_T2_UNIT = TIDB2_FieldName_Scan("T2_UNIT")
kRegTrans_T2_PERCENT = TIDB2_FieldName_Scan("T2_PERCENT")
kRegTrans_TIERAMT2 = TIDB2_FieldName_Scan("TIERAMT2")
kRegTrans_T3_NO = TIDB2_FieldName_Scan("T3_NO")
kRegTrans_T3_UNIT = TIDB2_FieldName_Scan("T3_UNIT")
kRegTrans_T3_PERCENT = TIDB2_FieldName_Scan("T3_PERCENT")
kRegTrans_TIERAMT3 = TIDB2_FieldName_Scan("TIERAMT3")
kRegTrans_T4_NO = TIDB2_FieldName_Scan("T4_NO")
kRegTrans_T4_UNIT = TIDB2_FieldName_Scan("T4_UNIT")
kRegTrans_T4_PERCENT = TIDB2_FieldName_Scan("T4_PERCENT")
kRegTrans_TIERAMT4 = TIDB2_FieldName_Scan("TIERAMT4")
kRegTrans_OVERALLMIN = TIDB2_FieldName_Scan("OVERALLMIN")
kRegTrans_CCY = TIDB2_FieldName_Scan("CCY")

'==============================================================

If kRegTrans_KEY97 < 0 Then Call MsgBox("champ 'KEY97' non trouvé", vbCritical, "TIDB2_RegTrans"): Exit Sub
If kRegTrans_SCHED_TYPE < 0 Then Call MsgBox("champ 'SCHED_TYPE ' non trouvé", vbCritical, "TIDB2_RegTrans"): Exit Sub
If kRegTrans_CHG_TYPE < 0 Then Call MsgBox("champ 'CHG_TYPE' non trouvé", vbCritical, "TIDB2_RegTrans"): Exit Sub
If kRegTrans_T1_NO < 0 Then Call MsgBox("champ 'T1_NO' non trouvé", vbCritical, "TIDB2_RegTrans"): Exit Sub
If kRegTrans_T1_UNIT < 0 Then Call MsgBox("champ 'T1_UNIT' non trouvé", vbCritical, "TIDB2_RegTrans"): Exit Sub
If kRegTrans_T1_PERCENT < 0 Then Call MsgBox("champ 'T1_PERCENT' non trouvé", vbCritical, "TIDB2_RegTrans"): Exit Sub
If kRegTrans_TIERAMT1 < 0 Then Call MsgBox("champ 'TIERAMT1' non trouvé", vbCritical, "TIDB2_RegTrans"): Exit Sub
If kRegTrans_T2_NO < 0 Then Call MsgBox("champ 'T2_NO' non trouvé", vbCritical, "TIDB2_RegTrans"): Exit Sub
If kRegTrans_T2_UNIT < 0 Then Call MsgBox("champ 'T2_UNIT' non trouvé", vbCritical, "TIDB2_RegTrans"): Exit Sub
If kRegTrans_T2_PERCENT < 0 Then Call MsgBox("champ 'T2_PERCENT' non trouvé", vbCritical, "TIDB2_RegTrans"): Exit Sub
If kRegTrans_TIERAMT2 < 0 Then Call MsgBox("champ 'TIERAMT2' non trouvé", vbCritical, "TIDB2_RegTrans"): Exit Sub
If kRegTrans_T3_NO < 0 Then Call MsgBox("champ 'T3_NO' non trouvé", vbCritical, "TIDB2_RegTrans"): Exit Sub
If kRegTrans_T3_UNIT < 0 Then Call MsgBox("champ 'T3_UNIT' non trouvé", vbCritical, "TIDB2_RegTrans"): Exit Sub
If kRegTrans_T3_PERCENT < 0 Then Call MsgBox("champ 'T3_PERCENT' non trouvé", vbCritical, "TIDB2_RegTrans"): Exit Sub
If kRegTrans_TIERAMT3 < 0 Then Call MsgBox("champ 'TIERAMT3' non trouvé", vbCritical, "TIDB2_RegTrans"): Exit Sub
If kRegTrans_T4_NO < 0 Then Call MsgBox("champ 'T4_NO' non trouvé", vbCritical, "TIDB2_RegTrans"): Exit Sub
If kRegTrans_T4_UNIT < 0 Then Call MsgBox("champ 'T4_UNIT' non trouvé", vbCritical, "TIDB2_RegTrans"): Exit Sub
If kRegTrans_T4_PERCENT < 0 Then Call MsgBox("champ 'T4_PERCENT' non trouvé", vbCritical, "TIDB2_RegTrans"): Exit Sub
If kRegTrans_TIERAMT4 < 0 Then Call MsgBox("champ 'TIERAMT4' non trouvé", vbCritical, "TIDB2_RegTrans"): Exit Sub
If kRegTrans_OVERALLMIN < 0 Then Call MsgBox("champ 'OVERALLMIN' non trouvé", vbCritical, "TIDB2_RegTrans"): Exit Sub
If kRegTrans_CCY < 0 Then Call MsgBox("champ 'CCY' non trouvé", vbCritical, "TIDB2_RegTrans"): Exit Sub

''paramTIDB2_Output = paramAS400IN & "CDRTRW0"
''Open paramTIDB2_Output For Output As #2
Open "\\Fr11024427\As400_IN\CDRTRW0" For Output As #2

xIn_nb = 0
xOut_nb = 0

Do Until EOF(1)
    DoEvents
    Line Input #1, xIn
    If Trim(xIn) <> "" Then
        I1 = InStr(1, xIn, "enregistrement(s) sélectionné(s)")
        If I1 > 0 Then
            xIn_nb = CLng(Val(mId$(xIn, 1, I1 - 1)))
            Exit Do
        End If
        
        X1 = mId$(xIn, arrField_Pos(kRegTrans_KEY97), arrField_Len(kRegTrans_KEY97))
        X2 = mId$(xIn, arrField_Pos(kRegTrans_SCHED_TYPE), arrField_Len(kRegTrans_SCHED_TYPE))
        X3 = mId$(xIn, arrField_Pos(kRegTrans_CHG_TYPE), arrField_Len(kRegTrans_CHG_TYPE))
        X4 = mId$(xIn, arrField_Pos(kRegTrans_T1_NO), arrField_Len(kRegTrans_T1_NO))
        X5 = mId$(xIn, arrField_Pos(kRegTrans_T1_UNIT), arrField_Len(kRegTrans_T1_UNIT))
        X6 = mId$(xIn, arrField_Pos(kRegTrans_T1_PERCENT), arrField_Len(kRegTrans_T1_PERCENT))
        X7 = mId$(xIn, arrField_Pos(kRegTrans_TIERAMT1), arrField_Len(kRegTrans_TIERAMT1))
        X8 = mId$(xIn, arrField_Pos(kRegTrans_T2_NO), arrField_Len(kRegTrans_T2_NO))
        X9 = mId$(xIn, arrField_Pos(kRegTrans_T2_UNIT), arrField_Len(kRegTrans_T2_UNIT))
        X10 = mId$(xIn, arrField_Pos(kRegTrans_T2_PERCENT), arrField_Len(kRegTrans_T2_PERCENT))
        X11 = mId$(xIn, arrField_Pos(kRegTrans_TIERAMT2), arrField_Len(kRegTrans_TIERAMT2))
        X12 = mId$(xIn, arrField_Pos(kRegTrans_T3_NO), arrField_Len(kRegTrans_T3_NO))
        X13 = mId$(xIn, arrField_Pos(kRegTrans_T3_UNIT), arrField_Len(kRegTrans_T3_UNIT))
        X14 = mId$(xIn, arrField_Pos(kRegTrans_T3_PERCENT), arrField_Len(kRegTrans_T3_PERCENT))
        X15 = mId$(xIn, arrField_Pos(kRegTrans_TIERAMT3), arrField_Len(kRegTrans_TIERAMT3))
        X16 = mId$(xIn, arrField_Pos(kRegTrans_T4_NO), arrField_Len(kRegTrans_T4_NO))
        X17 = mId$(xIn, arrField_Pos(kRegTrans_T4_UNIT), arrField_Len(kRegTrans_T4_UNIT))
        X18 = mId$(xIn, arrField_Pos(kRegTrans_T4_PERCENT), arrField_Len(kRegTrans_T4_PERCENT))
        X19 = mId$(xIn, arrField_Pos(kRegTrans_TIERAMT4), arrField_Len(kRegTrans_TIERAMT4))
        X20 = mId$(xIn, arrField_Pos(kRegTrans_OVERALLMIN), arrField_Len(kRegTrans_OVERALLMIN))
        X21 = mId$(xIn, arrField_Pos(kRegTrans_CCY), arrField_Len(kRegTrans_CCY))
    
        xOut = Space$(175)
    Mid$(xOut, 1, 12) = Format$(X1, "000000000000")
    Mid$(xOut, 13, 1) = Format$(X2, "@")
    Mid$(xOut, 14, 12) = Format$(X3, "000000000000")
    Mid$(xOut, 26, 3) = Format$(X21, "@@@")
    ' T1
    If Trim(X4) = "-" Then
        Mid$(xOut, 29, 3) = "000"
    Else
        Mid$(xOut, 29, 3) = Format$(X4, "000")
    End If
    If Trim(X5) = "-" Then
        Mid$(xOut, 32, 1) = " "
    Else
        Mid$(xOut, 32, 1) = Format$(X5, "@")
    End If
    ' Taux
    If Trim(X6) = "-" Then
        Mid$(xOut, 33, 11) = "00000000000"
    Else
        Mid$(xOut, 33, 11) = Format$(CDbl(X6) * 10000000, "00000000000")
    End If
    ' Montant
    If Trim(X7) = "-" Then
        Mid$(xOut, 44, 17) = "00000000000000000"
    Else
        Mnt = Val(X7)
        Mnt = curMaxD(Mnt, 2)
        Mid$(xOut, 44, 17) = Format$(Abs(Mnt) * 100, "00000000000000000")
    End If
    
    ' T2
    If Trim(X8) = "-" Then
        Mid$(xOut, 61, 3) = "000"
    Else
        Mid$(xOut, 61, 3) = Format$(X8, "000")
    End If
    If Trim(X9) = "-" Then
        Mid$(xOut, 64, 1) = " "
    Else
        Mid$(xOut, 64, 1) = Format$(X9, "@")
    End If
    If Trim(X10) = "-" Then
        Mid$(xOut, 65, 11) = "00000000000"
    Else
        Mid$(xOut, 65, 11) = Format$(CDbl(X10) * 10000000, "00000000000")
    End If
    If Trim(X11) = "-" Then
        Mid$(xOut, 76, 17) = "00000000000000000"
    Else
        Mnt = Val(X11)
        Mnt = curMaxD(Mnt, 2)
        Mid$(xOut, 76, 17) = Format$(Abs(Mnt) * 100, "00000000000000000")
    End If
    
    ' T3
    If Trim(X12) = "-" Then
        Mid$(xOut, 93, 3) = "000"
    Else
        Mid$(xOut, 93, 3) = Format$(X12, "000")
    End If
    If Trim(X13) = "-" Then
        Mid$(xOut, 96, 1) = " "
    Else
        Mid$(xOut, 96, 1) = Format$(X13, "@")
    End If
    If Trim(X14) = "-" Then
        Mid$(xOut, 97, 11) = "00000000000"
    Else
        Mid$(xOut, 97, 11) = Format$(CDbl(X14) * 10000000, "00000000000")
    End If
    If Trim(X15) = "-" Then
        Mid$(xOut, 108, 17) = "00000000000000000"
    Else
        Mnt = Val(X15)
        Mnt = curMaxD(Mnt, 2)
        Mid$(xOut, 108, 17) = Format$(Abs(Mnt) * 100, "00000000000000000")
    End If
    
    ' T4
    If Trim(X16) = "-" Then
        Mid$(xOut, 125, 3) = "000"
    Else
        Mid$(xOut, 125, 3) = Format$(X16, "000")
    End If
    If Trim(X17) = "-" Then
        Mid$(xOut, 128, 1) = " "
    Else
        Mid$(xOut, 128, 1) = Format$(X17, "@")
    End If
    If Trim(X18) = "-" Then
        Mid$(xOut, 129, 11) = "00000000000"
    Else
        Mid$(xOut, 129, 11) = Format$(CDbl(X18) * 10000000, "00000000000")
    End If
    If Trim(X19) = "-" Then
        Mid$(xOut, 140, 17) = "00000000000000000"
    Else
        Mnt = Val(X19)
        Mnt = curMaxD(Mnt, 2)
        Mid$(xOut, 140, 17) = Format$(Abs(Mnt) * 100, "00000000000000000")
    End If
    
    ' Montant commission minimum
    Mnt = Val(X20)
    Mnt = curMaxD(Mnt, 2)
    Mid$(xOut, 157, 17) = Format$(Abs(Mnt) * 100, "00000000000000000")
    ' Si comm. flat alors méthode -01-
    If X2 = "F" Then
        Mid$(xOut, 174, 2) = "01"
    Else
        Mid$(xOut, 174, 2) = "02"
    End If
    ' Si commission au taux annuel alors méthode -05-
    If Trim(X5) = "Y" Or Trim(X9) = "Y" Then
        Mid$(xOut, 174, 2) = "05"
    End If
      
        xOut_nb = xOut_nb + 1
        Print #2, xOut
    End If
Loop

If xOut_nb <> xIn_nb Then Call MsgBox("enregistrements traités / enregistrements attendus : " & xOut_nb & " / " & xIn_nb, vbCritical, "TIDB2_RegTrans")
lstMsg.AddItem "srvTIAS400.TIDB2_ComEnc : " & xOut_nb & " / " & xIn_nb
Close
Exit Sub

Error_Handle:
     lstMsg.AddItem "srvTIAS400.TIDB2_RegTrans : " & xOut_nb & " : " & Error
     lstMsg.AddItem xIn
    Close

End Sub

Public Sub TIDB2_FullPosting(lstMsg As ListBox)
Dim I1 As Integer
On Error GoTo Error_Handle

Open paramTIDB2_FullPosting For Input As #1

TIDB2_Load_Init

'===============
kPosting_BRANCH_NUM = TIDB2_FieldName_Scan("BRANCH_NUM")
kPosting_BASIC_NUM = TIDB2_FieldName_Scan("BASIC_NUM")
kPosting_ACC_SUFFIX = TIDB2_FieldName_Scan("ACC_SUFFIX")
kPosting_TRAN_CODE = TIDB2_FieldName_Scan("TRAN_CODE")
kPosting_VALUEDATE = TIDB2_FieldName_Scan("VALUEDATE")
kPosting_DR_CR_FLG = TIDB2_FieldName_Scan("DR_CR_FLG")
kPosting_AMOUNT = TIDB2_FieldName_Scan("AMOUNT")
kPosting_CCY = TIDB2_FieldName_Scan("CCY")
kPosting_ACC_TYPE = TIDB2_FieldName_Scan("ACC_TYPE")
kPosting_SP_CODE = TIDB2_FieldName_Scan("SP_CODE")
kPosting_SK_CODE = TIDB2_FieldName_Scan("SK_CODE")
kPosting_BRANCH = TIDB2_FieldName_Scan("BRANCH")
kPosting_MATURED = TIDB2_FieldName_Scan("MATURED")
kPosting_EQ3SEQNO = TIDB2_FieldName_Scan("EQ3SEQNO")
kPosting_EQ3RECNREF = TIDB2_FieldName_Scan("EQ3RECNREF")
kPosting_KEY97 = TIDB2_FieldName_Scan("KEY97")
kPosting_PARTY = TIDB2_FieldName_Scan("PARTY")

'==============================================================

If kPosting_BRANCH_NUM < 0 Then Call MsgBox("champ 'BRANCH_NUM' non trouvé", vbCritical, "TIDB2_Posting"): Exit Sub
If kPosting_BASIC_NUM < 0 Then Call MsgBox("champ 'BASIC_NUM' non trouvé", vbCritical, "TIDB2_Posting"): Exit Sub
If kPosting_ACC_SUFFIX < 0 Then Call MsgBox("champ 'ACC_SUFFIX' non trouvé", vbCritical, "TIDB2_Posting"): Exit Sub
If kPosting_TRAN_CODE < 0 Then Call MsgBox("champ 'TRAN_CODE' non trouvé", vbCritical, "TIDB2_Posting"): Exit Sub
If kPosting_VALUEDATE < 0 Then Call MsgBox("champ 'VALUEDATE' non trouvé", vbCritical, "TIDB2_Posting"): Exit Sub
If kPosting_DR_CR_FLG < 0 Then Call MsgBox("champ 'DR_CR_FLG' non trouvé", vbCritical, "TIDB2_Posting"): Exit Sub
If kPosting_AMOUNT < 0 Then Call MsgBox("champ 'AMOUNT' non trouvé", vbCritical, "TIDB2_Posting"): Exit Sub
If kPosting_CCY < 0 Then Call MsgBox("champ 'CCY' non trouvé", vbCritical, "TIDB2_Posting"): Exit Sub
If kPosting_ACC_TYPE < 0 Then Call MsgBox("champ 'ACC_TYPE' non trouvé", vbCritical, "TIDB2_Posting"): Exit Sub
If kPosting_SP_CODE < 0 Then Call MsgBox("champ 'SP_CODE' non trouvé", vbCritical, "TIDB2_Posting"): Exit Sub
If kPosting_SK_CODE < 0 Then Call MsgBox("champ 'SK_CODE' non trouvé", vbCritical, "TIDB2_Posting"): Exit Sub
If kPosting_BRANCH < 0 Then Call MsgBox("champ 'BRANCH' non trouvé", vbCritical, "TIDB2_Posting"): Exit Sub
If kPosting_MATURED < 0 Then Call MsgBox("champ 'MATURED' non trouvé", vbCritical, "TIDB2_Posting"): Exit Sub
If kPosting_EQ3SEQNO < 0 Then Call MsgBox("champ 'EQ3SEQNO' non trouvé", vbCritical, "TIDB2_Posting"): Exit Sub
If kPosting_EQ3RECNREF < 0 Then Call MsgBox("champ 'EQ3RECNREF' non trouvé", vbCritical, "TIDB2_Posting"): Exit Sub
If kPosting_KEY97 < 0 Then Call MsgBox("champ 'KEY97' non trouvé", vbCritical, "TIDB2_Posting"): Exit Sub
If kPosting_PARTY < 0 Then Call MsgBox("champ 'PARTY' non trouvé", vbCritical, "TIDB2_Posting"): Exit Sub

''paramTIDB2_Output = paramAS400IN & "CDPOSW0"
''Open paramTIDB2_Output For Output As #2
Open "\\Fr11024427\As400_IN\CDOFPOW0" For Output As #2

xIn_nb = 0
xOut_nb = 0

Do Until EOF(1)
    DoEvents
    Line Input #1, xIn
    If Trim(xIn) <> "" Then
        I1 = InStr(1, xIn, "enregistrement(s) sélectionné(s)")
        If I1 > 0 Then
            xIn_nb = CLng(Val(mId$(xIn, 1, I1 - 1)))
            Exit Do
        End If
        
        X1 = mId$(xIn, arrField_Pos(kPosting_BRANCH_NUM), arrField_Len(kPosting_BRANCH_NUM))
        X2 = mId$(xIn, arrField_Pos(kPosting_BASIC_NUM), arrField_Len(kPosting_BASIC_NUM))
        X3 = mId$(xIn, arrField_Pos(kPosting_ACC_SUFFIX), arrField_Len(kPosting_ACC_SUFFIX))
        X4 = mId$(xIn, arrField_Pos(kPosting_TRAN_CODE), arrField_Len(kPosting_TRAN_CODE))
        X5 = mId$(xIn, arrField_Pos(kPosting_VALUEDATE), arrField_Len(kPosting_VALUEDATE))
        X6 = mId$(xIn, arrField_Pos(kPosting_DR_CR_FLG), arrField_Len(kPosting_DR_CR_FLG))
        X7 = mId$(xIn, arrField_Pos(kPosting_AMOUNT), arrField_Len(kPosting_AMOUNT))
        X8 = mId$(xIn, arrField_Pos(kPosting_CCY), arrField_Len(kPosting_CCY))
        X9 = mId$(xIn, arrField_Pos(kPosting_ACC_TYPE), arrField_Len(kPosting_ACC_TYPE))
        X10 = mId$(xIn, arrField_Pos(kPosting_SP_CODE), arrField_Len(kPosting_SP_CODE))
        X11 = mId$(xIn, arrField_Pos(kPosting_SK_CODE), arrField_Len(kPosting_SK_CODE))
        X12 = mId$(xIn, arrField_Pos(kPosting_BRANCH), arrField_Len(kPosting_BRANCH))
        X13 = mId$(xIn, arrField_Pos(kPosting_MATURED), arrField_Len(kPosting_MATURED))
        X14 = mId$(xIn, arrField_Pos(kPosting_EQ3SEQNO), arrField_Len(kPosting_EQ3SEQNO))
        X15 = mId$(xIn, arrField_Pos(kPosting_EQ3RECNREF), arrField_Len(kPosting_EQ3RECNREF))
        X16 = mId$(xIn, arrField_Pos(kPosting_KEY97), arrField_Len(kPosting_KEY97))
        X17 = mId$(xIn, arrField_Pos(kPosting_PARTY), arrField_Len(kPosting_PARTY))
      
        xOut = Space$(151)
        
    Mid$(xOut, 1, 12) = Format$(X16, "000000000000")
    Mid$(xOut, 13, 12) = "000000000000"
    Mid$(xOut, 25, 3) = "   "
    Mid$(xOut, 28, 6) = "000000"
    Mid$(xOut, 34, 12) = "000000000000"
    Mid$(xOut, 46, 3) = "   "
    Mid$(xOut, 49, 6) = Format(Trim(X15), "000000")
    Mid$(xOut, 55, 4) = X12
  'Date
    dateJma10_Amj X5, DateAMJ
    Mid$(xOut, 59, 8) = DateAMJ
 
    If Trim(X1) = "-" Then
        Mid$(xOut, 67, 4) = "    "
    Else
        Mid$(xOut, 67, 4) = Format$(X1, "@@@@")
    End If
    If Trim(X2) = "-" Then
        Mid$(xOut, 71, 6) = "      "
    Else
        Mid$(xOut, 71, 6) = Format$(X2, "@@@@@@")
    End If
    If Trim(X3) = "-" Then
        Mid$(xOut, 77, 3) = "   "
    Else
        Mid$(xOut, 77, 3) = Format$(X3, "@@@")
    End If
    Mid$(xOut, 80, 25) = "                         "
    Mid$(xOut, 105, 3) = X4
    Mid$(xOut, 108, 1) = X6
 ' Devise
    Mid$(xOut, 126, 3) = X8
 
 ' Montant
    Mnt = Val(X7)
    If X8 = "ITL" Or X8 = "GRD" Or X8 = "PTE" _
       Or X8 = "ESP" Or X8 = "BEF" Or X8 = "LUF" _
       Or X8 = "JPY" Or X8 = "CFA" Or X8 = "XAF" _
       Or X8 = "XOF" Then
       Mnt = curMaxD(Mnt, 0)
    Else
       Mnt = curMaxD(Mnt, 2)
    End If
    Mid$(xOut, 109, 17) = Format$(Abs(Mnt) * 100, "00000000000000000")

    If Trim(X9) = "-" Then
        Mid$(xOut, 129, 2) = "  "
    Else
        Mid$(xOut, 129, 2) = Format$(X9, "@@")
    End If
    If Trim(X10) = "-" Then
        Mid$(xOut, 131, 6) = "      "
    Else
        Mid$(xOut, 131, 6) = Format$(X10, "@@@@@@")
    End If
    If Trim(X11) = "-" Then
        Mid$(xOut, 137, 2) = "  "
    Else
        Mid$(xOut, 137, 2) = Format$(X11, "@@")
    End If
    If Trim(X17) = "-" Then
        Mid$(xOut, 139, 12) = "000000000000"
    Else
        Mid$(xOut, 139, 12) = Format$(X17, "000000000000")
    End If
    Mid$(xOut, 151, 1) = X13
    
        xOut_nb = xOut_nb + 1
        Print #2, xOut
    End If
Loop

If xOut_nb <> xIn_nb Then Call MsgBox("enregistrements traités / enregistrements attendus : " & xOut_nb & " / " & xIn_nb, vbCritical, "TIDB2_FullPosting")
lstMsg.AddItem "srvTIAS400.TIDB2_FullPosting : " & xOut_nb & " / " & xIn_nb
Close
Exit Sub

Error_Handle:
     lstMsg.AddItem "srvTIAS400.TIDB2_FullPosting : " & xOut_nb & " : " & Error
     lstMsg.AddItem xIn
    Close

End Sub

Public Sub TIDB2_Posting(lstMsg As ListBox)
Dim I1 As Integer
On Error GoTo Error_Handle


Open paramTIDB2_Posting For Input As #1

TIDB2_Load_Init

'===============
kPosting_BRANCH_NUM = TIDB2_FieldName_Scan("BRANCH_NUM")
kPosting_BASIC_NUM = TIDB2_FieldName_Scan("BASIC_NUM")
kPosting_ACC_SUFFIX = TIDB2_FieldName_Scan("ACC_SUFFIX")
kPosting_TRAN_CODE = TIDB2_FieldName_Scan("TRAN_CODE")
kPosting_VALUEDATE = TIDB2_FieldName_Scan("VALUEDATE")
kPosting_DR_CR_FLG = TIDB2_FieldName_Scan("DR_CR_FLG")
kPosting_AMOUNT = TIDB2_FieldName_Scan("AMOUNT")
kPosting_CCY = TIDB2_FieldName_Scan("CCY")
kPosting_ACC_TYPE = TIDB2_FieldName_Scan("ACC_TYPE")
kPosting_SP_CODE = TIDB2_FieldName_Scan("SP_CODE")
kPosting_SK_CODE = TIDB2_FieldName_Scan("SK_CODE")
kPosting_BRANCH = TIDB2_FieldName_Scan("BRANCH")
kPosting_MATURED = TIDB2_FieldName_Scan("MATURED")
kPosting_EQ3SEQNO = TIDB2_FieldName_Scan("EQ3SEQNO")
kPosting_EQ3RECNREF = TIDB2_FieldName_Scan("EQ3RECNREF")
kPosting_KEY97 = TIDB2_FieldName_Scan("KEY97")
kPosting_PARTY = TIDB2_FieldName_Scan("PARTY")

'==============================================================

If kPosting_BRANCH_NUM < 0 Then Call MsgBox("champ 'BRANCH_NUM' non trouvé", vbCritical, "TIDB2_Posting"): Exit Sub
If kPosting_BASIC_NUM < 0 Then Call MsgBox("champ 'BASIC_NUM' non trouvé", vbCritical, "TIDB2_Posting"): Exit Sub
If kPosting_ACC_SUFFIX < 0 Then Call MsgBox("champ 'ACC_SUFFIX' non trouvé", vbCritical, "TIDB2_Posting"): Exit Sub
If kPosting_TRAN_CODE < 0 Then Call MsgBox("champ 'TRAN_CODE' non trouvé", vbCritical, "TIDB2_Posting"): Exit Sub
If kPosting_VALUEDATE < 0 Then Call MsgBox("champ 'VALUEDATE' non trouvé", vbCritical, "TIDB2_Posting"): Exit Sub
If kPosting_DR_CR_FLG < 0 Then Call MsgBox("champ 'DR_CR_FLG' non trouvé", vbCritical, "TIDB2_Posting"): Exit Sub
If kPosting_AMOUNT < 0 Then Call MsgBox("champ 'AMOUNT' non trouvé", vbCritical, "TIDB2_Posting"): Exit Sub
If kPosting_CCY < 0 Then Call MsgBox("champ 'CCY' non trouvé", vbCritical, "TIDB2_Posting"): Exit Sub
If kPosting_ACC_TYPE < 0 Then Call MsgBox("champ 'ACC_TYPE' non trouvé", vbCritical, "TIDB2_Posting"): Exit Sub
If kPosting_SP_CODE < 0 Then Call MsgBox("champ 'SP_CODE' non trouvé", vbCritical, "TIDB2_Posting"): Exit Sub
If kPosting_SK_CODE < 0 Then Call MsgBox("champ 'SK_CODE' non trouvé", vbCritical, "TIDB2_Posting"): Exit Sub
If kPosting_BRANCH < 0 Then Call MsgBox("champ 'BRANCH' non trouvé", vbCritical, "TIDB2_Posting"): Exit Sub
If kPosting_MATURED < 0 Then Call MsgBox("champ 'MATURED' non trouvé", vbCritical, "TIDB2_Posting"): Exit Sub
If kPosting_EQ3SEQNO < 0 Then Call MsgBox("champ 'EQ3SEQNO' non trouvé", vbCritical, "TIDB2_Posting"): Exit Sub
If kPosting_EQ3RECNREF < 0 Then Call MsgBox("champ 'EQ3RECNREF' non trouvé", vbCritical, "TIDB2_Posting"): Exit Sub
If kPosting_KEY97 < 0 Then Call MsgBox("champ 'KEY97' non trouvé", vbCritical, "TIDB2_Posting"): Exit Sub
If kPosting_PARTY < 0 Then Call MsgBox("champ 'PARTY' non trouvé", vbCritical, "TIDB2_Posting"): Exit Sub

''paramTIDB2_Output = paramAS400IN & "CDPOSW0"
''Open paramTIDB2_Output For Output As #2
Open "\\Fr11024427\As400_IN\CDPOSW0" For Output As #2

xIn_nb = 0
xOut_nb = 0

Do Until EOF(1)
    DoEvents
    Line Input #1, xIn
    If Trim(xIn) <> "" Then
        I1 = InStr(1, xIn, "enregistrement(s) sélectionné(s)")
        If I1 > 0 Then
            xIn_nb = CLng(Val(mId$(xIn, 1, I1 - 1)))
            Exit Do
        End If
        
        X1 = mId$(xIn, arrField_Pos(kPosting_BRANCH_NUM), arrField_Len(kPosting_BRANCH_NUM))
        X2 = mId$(xIn, arrField_Pos(kPosting_BASIC_NUM), arrField_Len(kPosting_BASIC_NUM))
        X3 = mId$(xIn, arrField_Pos(kPosting_ACC_SUFFIX), arrField_Len(kPosting_ACC_SUFFIX))
        X4 = mId$(xIn, arrField_Pos(kPosting_TRAN_CODE), arrField_Len(kPosting_TRAN_CODE))
        X5 = mId$(xIn, arrField_Pos(kPosting_VALUEDATE), arrField_Len(kPosting_VALUEDATE))
        X6 = mId$(xIn, arrField_Pos(kPosting_DR_CR_FLG), arrField_Len(kPosting_DR_CR_FLG))
        X7 = mId$(xIn, arrField_Pos(kPosting_AMOUNT), arrField_Len(kPosting_AMOUNT))
        X8 = mId$(xIn, arrField_Pos(kPosting_CCY), arrField_Len(kPosting_CCY))
        X9 = mId$(xIn, arrField_Pos(kPosting_ACC_TYPE), arrField_Len(kPosting_ACC_TYPE))
        X10 = mId$(xIn, arrField_Pos(kPosting_SP_CODE), arrField_Len(kPosting_SP_CODE))
        X11 = mId$(xIn, arrField_Pos(kPosting_SK_CODE), arrField_Len(kPosting_SK_CODE))
        X12 = mId$(xIn, arrField_Pos(kPosting_BRANCH), arrField_Len(kPosting_BRANCH))
        X13 = mId$(xIn, arrField_Pos(kPosting_MATURED), arrField_Len(kPosting_MATURED))
        X14 = mId$(xIn, arrField_Pos(kPosting_EQ3SEQNO), arrField_Len(kPosting_EQ3SEQNO))
        X15 = mId$(xIn, arrField_Pos(kPosting_EQ3RECNREF), arrField_Len(kPosting_EQ3RECNREF))
        X16 = mId$(xIn, arrField_Pos(kPosting_KEY97), arrField_Len(kPosting_KEY97))
        X17 = mId$(xIn, arrField_Pos(kPosting_PARTY), arrField_Len(kPosting_PARTY))
      
        xOut = Space$(150)
    Mid$(xOut, 1, 12) = Format$(X16, "000000000000")
    Mid$(xOut, 13, 12) = "000000000000"
    Mid$(xOut, 25, 3) = "   "
    Mid$(xOut, 28, 6) = "000000"
    Mid$(xOut, 34, 12) = "000000000000"
    Mid$(xOut, 46, 3) = "   "
    Mid$(xOut, 49, 6) = Format(Trim(X15), "000000")
    Mid$(xOut, 55, 4) = X12
  'Date
    dateJma10_Amj X5, DateAMJ
    Mid$(xOut, 59, 8) = DateAMJ
 
    If Trim(X1) = "-" Then
        Mid$(xOut, 67, 4) = "    "
    Else
        Mid$(xOut, 67, 4) = Format$(X1, "@@@@")
    End If
    If Trim(X2) = "-" Then
        Mid$(xOut, 71, 6) = "      "
    Else
        Mid$(xOut, 71, 6) = Format$(X2, "@@@@@@")
    End If
    If Trim(X3) = "-" Then
        Mid$(xOut, 77, 3) = "   "
    Else
        Mid$(xOut, 77, 3) = Format$(X3, "@@@")
    End If
    Mid$(xOut, 80, 25) = "                         "
    Mid$(xOut, 105, 3) = X4
    Mid$(xOut, 108, 1) = X6
 ' Devise
    Mid$(xOut, 126, 3) = X8
 
 ' Montant
    Mnt = Val(X7)
    If X8 = "ITL" Or X8 = "GRD" Or X8 = "PTE" _
       Or X8 = "ESP" Or X8 = "BEF" Or X8 = "LUF" _
       Or X8 = "JPY" Or X8 = "CFA" Or X8 = "XAF" _
       Or X8 = "XOF" Then
       Mnt = curMaxD(Mnt, 0)
    Else
       Mnt = curMaxD(Mnt, 2)
    End If
    Mid$(xOut, 109, 17) = Format$(Abs(Mnt) * 100, "00000000000000000")

    If Trim(X9) = "-" Then
        Mid$(xOut, 129, 2) = "  "
    Else
        Mid$(xOut, 129, 2) = Format$(X9, "@@")
    End If
    If Trim(X10) = "-" Then
        Mid$(xOut, 131, 6) = "      "
    Else
        Mid$(xOut, 131, 6) = Format$(X10, "@@@@@@")
    End If
    If Trim(X11) = "-" Then
        Mid$(xOut, 137, 2) = "  "
    Else
        Mid$(xOut, 137, 2) = Format$(X11, "@@")
    End If
    If Trim(X17) = "-" Then
        Mid$(xOut, 139, 12) = "000000000000"
    Else
        Mid$(xOut, 139, 12) = Format$(X17, "000000000000")
    End If
      
        xOut_nb = xOut_nb + 1
        Print #2, xOut
    End If
Loop

If xOut_nb <> xIn_nb Then Call MsgBox("enregistrements traités / enregistrements attendus : " & xOut_nb & " / " & xIn_nb, vbCritical, "TIDB2_Posting")
lstMsg.AddItem "srvTIAS400.TIDB2_Posting : " & xOut_nb & " / " & xIn_nb
Close
Exit Sub

Error_Handle:
     lstMsg.AddItem "srvTIAS400.TIDB2_Posting : " & xOut_nb & " : " & Error
     lstMsg.AddItem xIn
    Close

End Sub

Public Function TIDB2_FieldName_Scan(lName As String)
Dim I As Integer
TIDB2_FieldName_Scan = -1
For I = 0 To arrField_Nb
    If arrField_Name(I) = lName Then TIDB2_FieldName_Scan = I: Exit Function
Next I

End Function

Public Function TIDB2_FieldName_Scan_Duplicate(lName As String, lNb As Integer)
Dim I As Integer, wNb As Integer
wNb = 0
TIDB2_FieldName_Scan_Duplicate = -1
For I = 0 To arrField_Nb
    If arrField_Name(I) = lName Then
        wNb = wNb + 1
        If wNb = lNb Then TIDB2_FieldName_Scan_Duplicate = I: Exit Function
    End If
Next I

End Function

