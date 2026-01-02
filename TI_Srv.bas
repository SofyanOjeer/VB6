Attribute VB_Name = "srvTI"
Option Explicit
Dim xIn As String, xOut As String, xIn1 As String
Dim xIn_Nb As Long, xIn_DB2 As Long
Dim blnIN_Control As Boolean

Public paramTIDB2_AMJSituation As String, paramTIDB2_AMJValidité As String
Public paramTIDB2_Input As String, paramTIDB2_Output As String, paramTIDB2_Table As String
Dim arrField_Pos(300) As Integer, arrField_Len(300) As Integer, arrField_Nb As Integer

Dim arrField_Name(300) As String

Dim curX As Currency, x As String, X8 As String * 8
Dim recMvtp0 As typeMvtP0

Dim kPosting_KEY97 As Integer
Dim kPosting_BRANCH As Integer
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
Dim kPosting_MATURED As Integer
Dim kPosting_EQ3SEQNO As Integer
Dim kPosting_EQ3RECNREF As Integer

Dim kPartyDtls_KEY97 As Integer
Dim kPartyDtls_ADDRESS1 As Integer
Dim kPartyDtls_CUS_MNM As Integer

Dim kLcMaster_KEY97 As Integer
Dim kLcMaster_REVOCABLE As Integer
Dim kLcMaster_CONF_INSTR As Integer
Dim kLcMaster_MGN_PCTAMT As Integer
Dim kLcMaster_MGN_AMT As Integer
Dim kLcMaster_MGN_CCY As Integer
Dim kLcMaster_MGN_BRN As Integer
Dim kLcMaster_MGN_BASIC As Integer
Dim kLcMaster_MGN_SFIX As Integer
Dim kLcMaster_CFM_PCTAMT As Integer
Dim kLcMaster_PCT_PLUS As Integer
Dim kLcMaster_PCT_MINUS As Integer
Dim kLcMaster_QUALIFIER As Integer
Dim kLcMaster_BEN_PTY As Integer
Dim kLcMaster_APP_PTY As Integer
Dim kLcMaster_RCVD_PTY As Integer
Dim kLcMaster_ISSPTY_PTY As Integer

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

Public Sub TIDB2_Load()
Dim xFileName As String, I As Integer
Dim wNb As Integer, wLen As Integer

On Error GoTo Error_Handle

CV_X1 = CV_Euro: CV_X2 = CV_Euro: CV_X3 = CV_Euro

arrField_Pos(0) = 1: arrField_Len(0) = 0: arrField_Nb = 0: wNb = 0
Open paramTIDB2_Input For Input As #1
Call lstErr_AddItem(frmTI.lstErr, frmTI.cmdContext, "TIDB2 : début ..." & paramTIDB2_Table)

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

xIn_Nb = 0: blnIN_Control = False

Open paramTIDB2_Output For Output As #2

''On Error GoTo 0

 Select Case UCase$(paramTIDB2_Table)
        Case "MASTER": TIDB2_Master
        Case "LCMASTER": TIDB2_LcMaster
'        Case "CALCTE": TIDB2_CalcText
        Case "POSTING": TIDB2_Posting
        Case "PARTYDTLS": TIDB2_PartyDtls
        
End Select
    
Call lstErr_AddItem(frmTI.lstErr, frmTI.cmdContext, xIn_Nb & " : " & paramTIDB2_Table): DoEvents
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


Public Sub TIDB2_Posting()
Dim I1 As Integer

On Error GoTo Error_Handle

kPosting_KEY97 = TIDB2_FieldName_Scan("KEY97")
If kPosting_KEY97 < 0 Then Call MsgBox("champ 'KEY97' non trouvé", vbCritical, "TIDB2_Posting"): Exit Sub

kPosting_BRANCH = TIDB2_FieldName_Scan("BRANCH")
If kPosting_BRANCH < 0 Then Call MsgBox("champ 'BRANCH' non trouvé", vbCritical, "TIDB2_Posting"): Exit Sub

kPosting_BRANCH_NUM = TIDB2_FieldName_Scan("BRANCH_NUM")
If kPosting_BRANCH_NUM < 0 Then Call MsgBox("champ 'BRANCH_NUM' non trouvé", vbCritical, "TIDB2_Posting"): Exit Sub

kPosting_BASIC_NUM = TIDB2_FieldName_Scan("BASIC_NUM")
If kPosting_BASIC_NUM < 0 Then Call MsgBox("champ 'BASIC_NUM' non trouvé", vbCritical, "TIDB2_Posting"): Exit Sub

kPosting_ACC_SUFFIX = TIDB2_FieldName_Scan("ACC_SUFFIX")
If kPosting_ACC_SUFFIX < 0 Then Call MsgBox("champ 'ACC_SUFFIX' non trouvé", vbCritical, "TIDB2_Posting"): Exit Sub

kPosting_TRAN_CODE = TIDB2_FieldName_Scan("TRAN_CODE")
If kPosting_TRAN_CODE < 0 Then Call MsgBox("champ 'TRAN_CODE' non trouvé", vbCritical, "TIDB2_Posting"): Exit Sub

kPosting_VALUEDATE = TIDB2_FieldName_Scan("VALUEDATE")
If kPosting_VALUEDATE < 0 Then Call MsgBox("champ 'VALUEDATE' non trouvé", vbCritical, "TIDB2_Posting"): Exit Sub

kPosting_DR_CR_FLG = TIDB2_FieldName_Scan("DR_CR_FLG")
If kPosting_DR_CR_FLG < 0 Then Call MsgBox("champ 'DR_CR' non trouvé", vbCritical, "TIDB2_Posting"): Exit Sub

kPosting_AMOUNT = TIDB2_FieldName_Scan("AMOUNT")
If kPosting_AMOUNT < 0 Then Call MsgBox("champ 'AMOUNT' non trouvé", vbCritical, "TIDB2_Posting"): Exit Sub

kPosting_CCY = TIDB2_FieldName_Scan("CCY")
If kPosting_CCY < 0 Then Call MsgBox("champ 'CCY' non trouvé", vbCritical, "TIDB2_Posting"): Exit Sub

kPosting_ACC_TYPE = TIDB2_FieldName_Scan("ACC_TYPE")
If kPosting_ACC_TYPE < 0 Then Call MsgBox("champ 'ACC_TYPE' non trouvé", vbCritical, "TIDB2_Posting"): Exit Sub

kPosting_SK_CODE = TIDB2_FieldName_Scan("SK_CODE")
If kPosting_SK_CODE < 0 Then Call MsgBox("champ 'SK_CODE' non trouvé", vbCritical, "TIDB2_Posting"): Exit Sub

kPosting_SP_CODE = TIDB2_FieldName_Scan("SP_CODE")
If kPosting_SP_CODE < 0 Then Call MsgBox("champ 'SP_CODE' non trouvé", vbCritical, "TIDB2_Posting"): Exit Sub

kPosting_MATURED = TIDB2_FieldName_Scan("MATURED")
If kPosting_MATURED < 0 Then Call MsgBox("champ 'MATURED' non trouvé", vbCritical, "TIDB2_Posting"): Exit Sub

kPosting_EQ3SEQNO = TIDB2_FieldName_Scan("EQ3SEQNO")
If kPosting_EQ3SEQNO < 0 Then Call MsgBox("champ 'EQ3SEQNO' non trouvé", vbCritical, "TIDB2_Posting"): Exit Sub

kPosting_EQ3RECNREF = TIDB2_FieldName_Scan("EQ3RECNREF")
If kPosting_EQ3RECNREF < 0 Then Call MsgBox("champ 'EQ3RECNREF' non trouvé", vbCritical, "TIDB2_Posting"): Exit Sub

MsgTxt = Space$(recCDPosPfLen)
recCDPosPf_Init xCDPosPf

Do Until EOF(1)
    DoEvents
    Line Input #1, xIn
    If Trim(xIn) = "" Then
        blnIN_Control = True
    Else
        If blnIN_Control Then
            I1 = InStr(1, xIn, "enregistrement(s) sélectionné(s).")
            If I1 > 0 Then
                xIn_DB2 = Val(mId$(xIn, 1, I1 - 1))
                If xIn_DB2 <> xIn_Nb Then Call MsgBox("erreur : DB2_Nb = " & xIn_DB2 & "  Lus_Nb = " & xIn_Nb, vbCritical, "TIDB2_Posting")
                Exit Sub
            End If
            
        End If
        
        For I1 = 1 To Len(xIn)
            If mId$(xIn, I1, 1) = "-" Then Mid$(xIn, I1, 1) = " "
        Next I1

        xIn_Nb = xIn_Nb + 1
        
        xCDPosPf.POPKEY = CLng(Val(Trim(mId$(xIn, arrField_Pos(kPosting_KEY97), arrField_Len(kPosting_KEY97)))))
        xCDPosPf.PODNUM = CLng(Val(Trim(mId$(xIn, arrField_Pos(kPosting_EQ3RECNREF), arrField_Len(kPosting_EQ3RECNREF)))))
        xCDPosPf.POBRC = Trim(mId$(xIn, arrField_Pos(kPosting_BRANCH), arrField_Len(kPosting_BRANCH)))
        Call dateJma10_Amj(Trim(mId$(xIn, arrField_Pos(kPosting_VALUEDATE), arrField_Len(kPosting_VALUEDATE))), xCDPosPf.PODVAL)
        xCDPosPf.POATIB = Trim(mId$(xIn, arrField_Pos(kPosting_BRANCH_NUM), arrField_Len(kPosting_BRANCH_NUM)))
        xCDPosPf.POATIN = Trim(mId$(xIn, arrField_Pos(kPosting_BASIC_NUM), arrField_Len(kPosting_BASIC_NUM)))
        xCDPosPf.POATIS = Trim(mId$(xIn, arrField_Pos(kPosting_ACC_SUFFIX), arrField_Len(kPosting_ACC_SUFFIX)))
        xCDPosPf.POTRCD = Trim(mId$(xIn, arrField_Pos(kPosting_TRAN_CODE), arrField_Len(kPosting_TRAN_CODE)))
        xCDPosPf.PODBCR = Trim(mId$(xIn, arrField_Pos(kPosting_DR_CR_FLG), arrField_Len(kPosting_DR_CR_FLG)))
        xCDPosPf.POCCY = Trim(mId$(xIn, arrField_Pos(kPosting_CCY), arrField_Len(kPosting_CCY)))
        xCDPosPf.POAMT = CCur(Val(Trim(mId$(xIn, arrField_Pos(kPosting_AMOUNT), arrField_Len(kPosting_AMOUNT)))))
        Call TIDB2_Amount(xCDPosPf.POCCY, xCDPosPf.POAMT)
        xCDPosPf.POACTY = Trim(mId$(xIn, arrField_Pos(kPosting_ACC_TYPE), arrField_Len(kPosting_ACC_TYPE)))
        xCDPosPf.POSPCD = Trim(mId$(xIn, arrField_Pos(kPosting_SP_CODE), arrField_Len(kPosting_SP_CODE)))
        xCDPosPf.POSKCD = Trim(mId$(xIn, arrField_Pos(kPosting_SK_CODE), arrField_Len(kPosting_SK_CODE)))
        MsgTxtLen = 0
        srvCDPosPf_PutBuffer xCDPosPf
        Print #2, mId$(MsgTxt, 35, MemoCDPosPfLen)
    End If
Read_Next:
Loop

Exit Sub

Error_Handle:

Call MsgBox("erreur " & xIn, vbCritical, "TIDB2_Posting")

End Sub

Public Sub TIDB2_PartyDtls()
Dim I1 As Integer

On Error GoTo Error_Handle

kPartyDtls_KEY97 = TIDB2_FieldName_Scan("KEY97")
If kPartyDtls_KEY97 < 0 Then Call MsgBox("champ 'KEY97' non trouvé", vbCritical, "TIDB2_Posting"): Exit Sub

kPartyDtls_ADDRESS1 = TIDB2_FieldName_Scan("ADDRESS1")
If kPartyDtls_ADDRESS1 < 0 Then Call MsgBox("champ 'ADDRESS1' non trouvé", vbCritical, "TIDB2_Posting"): Exit Sub

kPartyDtls_CUS_MNM = TIDB2_FieldName_Scan("CUS_MNM")
If kPartyDtls_CUS_MNM < 0 Then Call MsgBox("champ 'CUS_MNM' non trouvé", vbCritical, "TIDB2_Posting"): Exit Sub


MsgTxt = Space$(recCDPtyPfLen)
recCDPtyPf_Init xCDPtyPf

Do Until EOF(1)
    DoEvents
    Line Input #1, xIn
    If Trim(xIn) = "" Then
        blnIN_Control = True
    Else
        If blnIN_Control Then
            I1 = InStr(1, xIn, "enregistrement(s) sélectionné(s).")
            If I1 > 0 Then
                xIn_DB2 = Val(mId$(xIn, 1, I1 - 1))
                If xIn_DB2 <> xIn_Nb Then Call MsgBox("erreur : DB2_Nb = " & xIn_DB2 & "  Lus_Nb = " & xIn_Nb, vbCritical, "TIDB2_Posting")
                Exit Sub
            End If
            
        End If
        
        For I1 = 1 To Len(xIn)
            If mId$(xIn, I1, 1) = "-" Then Mid$(xIn, I1, 1) = " "
        Next I1

        xIn_Nb = xIn_Nb + 1
        
        xCDPtyPf.PTKEY = CLng(Val(Trim(mId$(xIn, arrField_Pos(kPartyDtls_KEY97), arrField_Len(kPartyDtls_KEY97)))))
        xCDPtyPf.PTNOM = Trim(mId$(xIn, arrField_Pos(kPartyDtls_ADDRESS1), arrField_Len(kPartyDtls_ADDRESS1)))
        xCDPtyPf.PTMNM = Trim(mId$(xIn, arrField_Pos(kPartyDtls_CUS_MNM), arrField_Len(kPartyDtls_CUS_MNM)))
        MsgTxtLen = 0
        srvCDPtyPf_PutBuffer xCDPtyPf
        Print #2, mId$(MsgTxt, 35, MemoCDPtyPfLen)
    End If
Read_Next:
Loop

Exit Sub

Error_Handle:

Call MsgBox("erreur " & xIn, vbCritical, "TIDB2_Posting")

End Sub


Public Sub TIDB2_Master()
Dim I1 As Integer, xKEY97 As String

On Error GoTo Error_Handle

kMaster_KEY97 = TIDB2_FieldName_Scan("KEY97")
If kMaster_KEY97 < 0 Then Call MsgBox("champ 'KEY97' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub

kMaster_REFNO_PFIX = TIDB2_FieldName_Scan("REFNO_PFIX")
If kMaster_REFNO_PFIX < 0 Then Call MsgBox("champ 'REFNO_PFIX' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub

kMaster_REFNO_SERL = TIDB2_FieldName_Scan("REFNO_SERL")
If kMaster_REFNO_SERL < 0 Then Call MsgBox("champ 'REFNO_SERL' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub

kMaster_ORIG_REF = TIDB2_FieldName_Scan("ORIG_REF")
If kMaster_ORIG_REF < 0 Then Call MsgBox("champ 'ORIG_REF' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub

kMaster_INPUT_BRN = TIDB2_FieldName_Scan("INPUT_BRN")
If kMaster_INPUT_BRN < 0 Then Call MsgBox("champ 'INPUT_BRN' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub

kMaster_BHALF_BRN = TIDB2_FieldName_Scan("BHALF_BRN")
If kMaster_BHALF_BRN < 0 Then Call MsgBox("champ 'BHALF_BRN' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub

kMaster_CTRCT_DATE = TIDB2_FieldName_Scan("CTRCT_DATE")
If kMaster_CTRCT_DATE < 0 Then Call MsgBox("champ 'CTRCT_DATE' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub

kMaster_EXPIRY_DAT = TIDB2_FieldName_Scan("EXPIRY_DAT")
If kMaster_EXPIRY_DAT < 0 Then Call MsgBox("champ 'EXPIRY_DAT' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub

kMaster_STATUS = TIDB2_FieldName_Scan("STATUS")
If kMaster_STATUS < 0 Then Call MsgBox("champ 'STATUS' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub

kMaster_USERCODE1 = TIDB2_FieldName_Scan("USERCODE1")
If kMaster_USERCODE1 < 0 Then Call MsgBox("champ 'USERCODE1' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub

kMaster_USERCODE2 = TIDB2_FieldName_Scan("USERCODE2")
If kMaster_USERCODE2 < 0 Then Call MsgBox("champ 'USERCODE2' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub

kMaster_USERCODE3 = TIDB2_FieldName_Scan("USERCODE3")
If kMaster_USERCODE3 < 0 Then Call MsgBox("champ '' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub

kMaster_EV_COUNT = TIDB2_FieldName_Scan("EV_COUNT")
If kMaster_EV_COUNT < 0 Then Call MsgBox("champ 'EV_COUNT' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub

kMaster_AMOUNT = TIDB2_FieldName_Scan("AMOUNT")
If kMaster_AMOUNT < 0 Then Call MsgBox("champ 'AMOUNT' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub

kMaster_CCY = TIDB2_FieldName_Scan("CCY")
If kMaster_CCY < 0 Then Call MsgBox("champ 'CCY' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub

kMaster_AMT_O_S = TIDB2_FieldName_Scan("AMT_O_S")
If kMaster_AMT_O_S < 0 Then Call MsgBox("champ 'AMT_O_S' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub

kMaster_LIAB_AMT = TIDB2_FieldName_Scan("LIAB_AMT")
If kMaster_LIAB_AMT < 0 Then Call MsgBox("champ 'LIAB_AMT' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub

kMaster_LIAB_CCY = TIDB2_FieldName_Scan("LIAB_CCY")
If kMaster_LIAB_CCY < 0 Then Call MsgBox("champ 'LIAB_CCY' non trouvé", vbCritical, "TIDB2_Master"): Exit Sub

tableMvtP0_Open
recMvtP0_Init recMvtp0
recMvtp0.Method = "Seek="

Do Until EOF(1)
    DoEvents
    Line Input #1, xIn
    If Trim(xIn) = "" Then
        blnIN_Control = True
    Else
        If blnIN_Control Then
            I1 = InStr(1, xIn, "enregistrement(s) sélectionné(s).")
            If I1 > 0 Then
                xIn_DB2 = Val(mId$(xIn, 1, I1 - 1))
                If xIn_DB2 <> xIn_Nb Then Call MsgBox("erreur : DB2_Nb = " & xIn_DB2 & "  Lus_Nb = " & xIn_Nb, vbCritical, "TIDB2_Master")
                Exit Sub
            End If
            
        End If
        
        For I1 = 1 To Len(xIn)
            If mId$(xIn, I1, 1) = "-" Then Mid$(xIn, I1, 1) = " "
        Next I1

        xIn_Nb = xIn_Nb + 1
        
        xKEY97 = Trim(mId$(xIn, arrField_Pos(kMaster_KEY97), arrField_Len(kMaster_KEY97)))
        recMvtp0.Id = xKEY97
      
        If tableMvtP0_Read(recMvtp0) = 0 Then
            Mid$(MsgTxt, 1, recCDDosPfLen) = mId$(recMvtp0.Text, 1, recCDDosPfLen)
            MsgTxtIndex = 0
            srvCDDosPf_GetBuffer xCDDosPf

        Else
            Call MsgBox("Manque LcMaster : " & xKEY97 & " : " & xIn, vbCritical, "TIDB2_Master")
            MsgTxt = Space$(recCDDosPfLen)
            recCDDosPf_Init xCDDosPf
        End If
        
        xCDDosPf.DODKEY = CLng(Val(xKEY97))
        xCDDosPf.DODPFX = Trim(mId$(xIn, arrField_Pos(kMaster_REFNO_PFIX), arrField_Len(kMaster_REFNO_PFIX)))
        xCDDosPf.DODNUM = CLng(Val(Trim(mId$(xIn, arrField_Pos(kMaster_REFNO_SERL), arrField_Len(kMaster_REFNO_SERL)))))
        xCDDosPf.DOREF = Trim(mId$(xIn, arrField_Pos(kMaster_ORIG_REF), arrField_Len(kMaster_ORIG_REF)))
        xCDDosPf.DOIBRC = Trim(mId$(xIn, arrField_Pos(kMaster_INPUT_BRN), arrField_Len(kMaster_INPUT_BRN)))
        xCDDosPf.DOBBRC = Trim(mId$(xIn, arrField_Pos(kMaster_BHALF_BRN), arrField_Len(kMaster_BHALF_BRN)))
        Call dateJma10_Amj(Trim(mId$(xIn, arrField_Pos(kMaster_CTRCT_DATE), arrField_Len(kMaster_CTRCT_DATE))), xCDDosPf.DODCTR)
        Call dateJma10_Amj(Trim(mId$(xIn, arrField_Pos(kMaster_EXPIRY_DAT), arrField_Len(kMaster_EXPIRY_DAT))), xCDDosPf.DODEXP)
        xCDDosPf.DOSTAT = Trim(mId$(xIn, arrField_Pos(kMaster_STATUS), arrField_Len(kMaster_STATUS)))
        xCDDosPf.DOUSC1 = Trim(mId$(xIn, arrField_Pos(kMaster_USERCODE1), arrField_Len(kMaster_USERCODE1)))
        xCDDosPf.DOUSC2 = Trim(mId$(xIn, arrField_Pos(kMaster_USERCODE2), arrField_Len(kMaster_USERCODE2)))
        xCDDosPf.DOUSC3 = Trim(mId$(xIn, arrField_Pos(kMaster_USERCODE3), arrField_Len(kMaster_USERCODE3)))
        xCDDosPf.DONBEV = CLng(Val(Trim(mId$(xIn, arrField_Pos(kMaster_EV_COUNT), arrField_Len(kMaster_EV_COUNT)))))
        xCDDosPf.DOAMT = CCur(Val(Trim(mId$(xIn, arrField_Pos(kMaster_AMOUNT), arrField_Len(kMaster_AMOUNT)))))
        xCDDosPf.DOCCY = Trim(mId$(xIn, arrField_Pos(kMaster_CCY), arrField_Len(kMaster_CCY)))
        Call TIDB2_Amount(xCDDosPf.DOCCY, xCDDosPf.DOAMT)
        xCDDosPf.DOOUTS = CCur(Val(Trim(mId$(xIn, arrField_Pos(kMaster_AMT_O_S), arrField_Len(kMaster_AMT_O_S)))))
        xCDDosPf.DOLIAB = CCur(Val(Trim(mId$(xIn, arrField_Pos(kMaster_LIAB_AMT), arrField_Len(kMaster_LIAB_AMT)))))
        xCDDosPf.DOLCCY = Trim(mId$(xIn, arrField_Pos(kMaster_LIAB_CCY), arrField_Len(kMaster_LIAB_CCY)))
        Call TIDB2_Amount(xCDDosPf.DOLCCY, xCDDosPf.DOOUTS)
        Call TIDB2_Amount(xCDDosPf.DOLCCY, xCDDosPf.DOLIAB)
      
        MsgTxtLen = 0
        srvCDDosPf_PutBuffer xCDDosPf
        Print #2, mId$(MsgTxt, 35, MemoCDDosPfLen)
    End If
Read_Next:
Loop

Exit Sub

Error_Handle:

Call MsgBox("erreur " & xIn, vbCritical, "TIDB2_Master")

End Sub



Public Sub TIDB2_LcMaster()
Dim I1 As Integer, xKEY97 As String

On Error GoTo Error_Handle

kLcMaster_KEY97 = TIDB2_FieldName_Scan("KEY97")
If kLcMaster_KEY97 < 0 Then Call MsgBox("champ 'KEY97' non trouvé", vbCritical, "TIDB2_LcMaster"): Exit Sub

kLcMaster_REVOCABLE = TIDB2_FieldName_Scan("REVOCABLE")
If kLcMaster_REVOCABLE < 0 Then Call MsgBox("champ 'REVOCABLE' non trouvé", vbCritical, "TIDB2_LcMaster"): Exit Sub

kLcMaster_CONF_INSTR = TIDB2_FieldName_Scan("CONF_INSTR")
If kLcMaster_CONF_INSTR < 0 Then Call MsgBox("champ 'CONF_INSTR' non trouvé", vbCritical, "TIDB2_LcMaster"): Exit Sub

kLcMaster_MGN_PCTAMT = TIDB2_FieldName_Scan("MGN_PCTAMT")
If kLcMaster_MGN_PCTAMT < 0 Then Call MsgBox("champ 'MGN_PCTAMT' non trouvé", vbCritical, "TIDB2_LcMaster"): Exit Sub

kLcMaster_MGN_AMT = TIDB2_FieldName_Scan("MGN_AMT")
If kLcMaster_MGN_AMT < 0 Then Call MsgBox("champ 'MGN_AMT' non trouvé", vbCritical, "TIDB2_LcMaster"): Exit Sub

kLcMaster_MGN_CCY = TIDB2_FieldName_Scan("MGN_CCY")
If kLcMaster_MGN_CCY < 0 Then Call MsgBox("champ 'MGN_CCY' non trouvé", vbCritical, "TIDB2_LcMaster"): Exit Sub

kLcMaster_MGN_BRN = TIDB2_FieldName_Scan("MGN_BRN")
If kLcMaster_MGN_BRN < 0 Then Call MsgBox("champ 'MGN_BRN' non trouvé", vbCritical, "TIDB2_LcMaster"): Exit Sub

kLcMaster_MGN_BASIC = TIDB2_FieldName_Scan("MGN_BASIC")
If kLcMaster_MGN_BASIC < 0 Then Call MsgBox("champ 'MGN_BASIC' non trouvé", vbCritical, "TIDB2_LcMaster"): Exit Sub

kLcMaster_MGN_SFIX = TIDB2_FieldName_Scan("MGN_SFIX")
If kLcMaster_MGN_SFIX < 0 Then Call MsgBox("champ 'MGN_SFIX' non trouvé", vbCritical, "TIDB2_LcMaster"): Exit Sub

kLcMaster_CFM_PCTAMT = TIDB2_FieldName_Scan("CFM_PCTAMT")
If kLcMaster_CFM_PCTAMT < 0 Then Call MsgBox("champ 'CFM_PCTAMT' non trouvé", vbCritical, "TIDB2_LcMaster"): Exit Sub

kLcMaster_PCT_PLUS = TIDB2_FieldName_Scan("PCT_PLUS")
If kLcMaster_PCT_PLUS < 0 Then Call MsgBox("champ 'PCT_PLUS' non trouvé", vbCritical, "TIDB2_LcMaster"): Exit Sub

kLcMaster_PCT_MINUS = TIDB2_FieldName_Scan("PCT_MINUS")
If kLcMaster_PCT_MINUS < 0 Then Call MsgBox("champ 'PCT_MINUS' non trouvé", vbCritical, "TIDB2_LcMaster"): Exit Sub

kLcMaster_QUALIFIER = TIDB2_FieldName_Scan("QUALIFIER")
If kLcMaster_QUALIFIER < 0 Then Call MsgBox("champ 'QUALIFIER' non trouvé", vbCritical, "TIDB2_LcMaster"): Exit Sub

kLcMaster_BEN_PTY = TIDB2_FieldName_Scan("BEN_PTY")
If kLcMaster_BEN_PTY < 0 Then Call MsgBox("champ 'BEN_PTY' non trouvé", vbCritical, "TIDB2_LcMaster"): Exit Sub

kLcMaster_APP_PTY = TIDB2_FieldName_Scan("APP_PTY")
If kLcMaster_APP_PTY < 0 Then Call MsgBox("champ 'APP_PTY' non trouvé", vbCritical, "TIDB2_LcMaster"): Exit Sub

kLcMaster_RCVD_PTY = TIDB2_FieldName_Scan("RCVD_PTY")
If kLcMaster_RCVD_PTY < 0 Then Call MsgBox("champ 'RCVD_PTY' non trouvé", vbCritical, "TIDB2_LcMaster"): Exit Sub

kLcMaster_ISSPTY_PTY = TIDB2_FieldName_Scan("ISSPTY_PTY")
If kLcMaster_ISSPTY_PTY < 0 Then Call MsgBox("champ 'ISSPTY_PTY' non trouvé", vbCritical, "TIDB2_LcMaster"): Exit Sub

MsgTxt = Space$(recCDDosPfLen)
recCDDosPf_Init xCDDosPf

tableMvtP0_Close
MDB.Execute "delete * from MVTP0"
tableMvtP0_Open
recMvtP0_Init recMvtp0
recMvtp0.Method = "AddNew"

Do Until EOF(1)
    DoEvents
    Line Input #1, xIn
    If Trim(xIn) = "" Then
        blnIN_Control = True
    Else
        If blnIN_Control Then
            I1 = InStr(1, xIn, "enregistrement(s) sélectionné(s).")
            If I1 > 0 Then
                xIn_DB2 = Val(mId$(xIn, 1, I1 - 1))
                If xIn_DB2 <> xIn_Nb Then Call MsgBox("erreur : DB2_Nb = " & xIn_DB2 & "  Lus_Nb = " & xIn_Nb, vbCritical, "TIDB2_LcMaster")
                Exit Sub
            End If
            
        End If
        
        For I1 = 1 To Len(xIn)
            If mId$(xIn, I1, 1) = "-" Then Mid$(xIn, I1, 1) = " "
        Next I1

        xIn_Nb = xIn_Nb + 1
        
        xKEY97 = Trim(mId$(xIn, arrField_Pos(kLcMaster_KEY97), arrField_Len(kLcMaster_KEY97)))
        xCDDosPf.DODKEY = CLng(Val(xKEY97))
        xCDDosPf.DOREV = Trim(mId$(xIn, arrField_Pos(kLcMaster_REVOCABLE), arrField_Len(kLcMaster_REVOCABLE)))
        xCDDosPf.DONAT = Trim(mId$(xIn, arrField_Pos(kLcMaster_CONF_INSTR), arrField_Len(kLcMaster_CONF_INSTR)))
        xCDDosPf.DOGPER = CDbl(Val(Trim(mId$(xIn, arrField_Pos(kLcMaster_MGN_PCTAMT), arrField_Len(kLcMaster_MGN_PCTAMT)))))
        xCDDosPf.DOGAMT = CCur(Val(Trim(mId$(xIn, arrField_Pos(kLcMaster_MGN_AMT), arrField_Len(kLcMaster_MGN_AMT)))))
        xCDDosPf.DOGCCY = Trim(mId$(xIn, arrField_Pos(kLcMaster_MGN_CCY), arrField_Len(kLcMaster_MGN_CCY)))
        If xCDDosPf.DOGAMT <> 0 Then Call TIDB2_Amount(xCDDosPf.DOGCCY, xCDDosPf.DOGAMT)
        xCDDosPf.DOGTIB = Trim(mId$(xIn, arrField_Pos(kLcMaster_MGN_BRN), arrField_Len(kLcMaster_MGN_BRN)))
        xCDDosPf.DOGTIN = Trim(mId$(xIn, arrField_Pos(kLcMaster_MGN_BASIC), arrField_Len(kLcMaster_MGN_BASIC)))
        xCDDosPf.DOGTIS = Trim(mId$(xIn, arrField_Pos(kLcMaster_MGN_SFIX), arrField_Len(kLcMaster_MGN_SFIX)))
        xCDDosPf.DOCPER = CDbl(Val(Trim(mId$(xIn, arrField_Pos(kLcMaster_CFM_PCTAMT), arrField_Len(kLcMaster_CFM_PCTAMT)))))
        xCDDosPf.DOPLUS = CDbl(Val(Trim(mId$(xIn, arrField_Pos(kLcMaster_PCT_PLUS), arrField_Len(kLcMaster_PCT_PLUS)))))
        xCDDosPf.DOMINS = CDbl(Val(Trim(mId$(xIn, arrField_Pos(kLcMaster_PCT_MINUS), arrField_Len(kLcMaster_PCT_MINUS)))))
        xCDDosPf.DOQUA = Trim(mId$(xIn, arrField_Pos(kLcMaster_QUALIFIER), arrField_Len(kLcMaster_QUALIFIER)))
        xCDDosPf.DOBNKY = CLng(Val(Trim(mId$(xIn, arrField_Pos(kLcMaster_BEN_PTY), arrField_Len(kLcMaster_BEN_PTY)))))
        xCDDosPf.DOAPKY = CLng(Val(Trim(mId$(xIn, arrField_Pos(kLcMaster_APP_PTY), arrField_Len(kLcMaster_APP_PTY)))))
        xCDDosPf.DORCKY = CLng(Val(Trim(mId$(xIn, arrField_Pos(kLcMaster_RCVD_PTY), arrField_Len(kLcMaster_RCVD_PTY)))))
        xCDDosPf.DOISKY = CLng(Val(Trim(mId$(xIn, arrField_Pos(kLcMaster_ISSPTY_PTY), arrField_Len(kLcMaster_ISSPTY_PTY)))))
        
        MsgTxtLen = 0
        srvCDDosPf_PutBuffer xCDDosPf
       ' Print #2, mId$(MsgTxt, 35, MemoCDDosPfLen)
       
        recMvtp0.Id = xKEY97
        recMvtp0.Text = mId$(MsgTxt, 1, recCDDosPfLen)
        dbMvtP0_Update recMvtp0

    End If
Read_Next:
Loop

Exit Sub

Error_Handle:

Call MsgBox("erreur " & xIn, vbCritical, "TIDB2_LcMaster")

End Sub




Public Function TIDB2_FieldName_Scan(lName As String)
Dim I As Integer
TIDB2_FieldName_Scan = -1
For I = 0 To arrField_Nb
    If arrField_Name(I) = lName Then TIDB2_FieldName_Scan = I: Exit Function
Next I

End Function

Public Sub TIDB2_Amount(lDeviseISO As String, lMontant As Currency)
CV_X1.DeviseIso = lDeviseISO
Call CV_Attribut(CV_X1)
lMontant = curMaxD(lMontant, CV_X1.maxD)

End Sub

