Attribute VB_Name = "KHALDZAL"
Option Explicit

Public cnSAB073Y As New ADODB.Connection
Public rsKHALDZAL As New ADODB.Recordset
Public rsYSAAMSG0 As New ADODB.Recordset
Public rsYSAAMSG1 As New ADODB.Recordset
Public rsYSAAMVTLNK As New ADODB.Recordset

Type typeKHALDZAL
    Id           As Long
    HISMVTDEV    As String * 3  'devise
    HISMVTCPT    As String * 11 'compte
    HISMVTOPEC   As String * 4  'code opération
    HISMVTOPEN   As Long        ' n° opération
    HISMVTMTD    As Currency    'Montant en devise
    HISMVTLIB1   As String      '* 50 ' libellé 1
    HISMVTLIB2   As String      '* 50 ' libellé 2
    HISMVTDTRT   As Long        '  date de traitement
    HISMVTDVAL   As Long        '  date de valeur
    HISMVTPIEN   As Long        '  n° pièce
    HISMVTPIES   As Long        '  n° séquence
    HISMVTXBEN   As String      '* 16 ' bénéficiaire
    HISMVTXREF   As String
    
    SAAMSGID    As Long         ' lien YSAAMSG0
    SAAMSGXDO   As String       ' donneur d'ordre
    SAAMSGXBEN  As String       ' bénéficiaire
    SAAMSGXPAY  As String       ' pays banque ben
    SAAMSGTXT   As String
    
    SAAMATCH      As String       ' flag matching
    SAAMSGTYPE    As String       ' type de message
    SAAMSGDTRT    As Long         ' date réception
    SCANLINK     As Variant      ' lien hypertexte
    SCANLINK2     As Variant      ' lien hypertexte
    
    
End Type


Type typeYSAAMVTLNK
    HISMVTID    As Long
    SAAMSGID    As Long         ' lien YSAAMSG0
    
End Type

Type typeKHALDZAL_Match
    Id           As Long
    HISMVTMTD    As Currency    'Montant en devise
    HISMVTDVAL   As Long        '  date de valeur
    HISMVTXREF   As String
    
    SAAMSGID    As Long         ' lien YSAAMSG0
    SAAMATCH      As String       ' flag matching
    SAAMSGTYPE    As String       ' type de message
    SAAMSGDTRT    As Long         ' date réception
    
    End Type

Public Sub cnSAB073Y_Close()
On Error Resume Next

cnSAB073Y.Close
Set cnSAB073Y = Nothing


End Sub

Public Sub cnSAB073Y_Open()
On Error GoTo Error_Handler
Dim X As String

cnSAB073Y.Open paramODBC_DSN_SAB073Y

Exit Sub

Error_Handler:

End Sub

Public Sub HISMVTP0_Import()
Dim V, xSql As String
Dim xIn As String, K As Integer, K2 As Integer, lenX As Integer
Dim kIn As Integer, Seq As Integer
On Error GoTo Error_Handle
Dim X As String
Dim mSeq As Integer
Dim K1 As Integer, I1 As Integer, I As Integer
Dim blnOk As Boolean, blnPrint As Boolean, blnSwift As Boolean, kPrint As Integer

Dim wKHALDZAL As typeKHALDZAL

cnSAB073Y_Open

xSql = "delete * from KHALDZAL"
Call FEU_ROUGE
Set rsKHALDZAL = cnSAB073Y.Execute(xSql)
Call FEU_VERT
rsKHALDZAL.Open "select * from KHALDZAL", cnSAB073Y, , adLockOptimistic

Call lstErr_AddItem(frmKHALDZAL.lstErr, frmYGOSDOS0.cmdContext, "Import : ")
xSql = "select * from BIAFIL.HISMVTP0 where COMPTE = 25272001014 " _
     & " order by DEVISE,COMPTE,TRAITA , TRAITM, TRAITJ,NUMPIE,NOLIGN"
Set rsSab = cnsab.Execute(xSql)
mSeq = 0
Do While Not rsSab.EOF
    Call rsKHALDZAL_Init(wKHALDZAL)
    mSeq = mSeq + 1
    wKHALDZAL.Id = mSeq
    K1 = rsSab("DEVISE")
    Select Case K1
        Case 978: wKHALDZAL.HISMVTDEV = "EUR"
        Case 400: wKHALDZAL.HISMVTDEV = "USD"
        Case 732: wKHALDZAL.HISMVTDEV = "JPY"
        Case 8: wKHALDZAL.HISMVTDEV = "DKK"
        Case Else: wKHALDZAL.HISMVTDEV = rsSab("DEVISE")
    End Select
    
    wKHALDZAL.HISMVTCPT = rsSab("COMPTE")
    wKHALDZAL.HISMVTMTD = rsSab("MONDEV")
    wKHALDZAL.HISMVTOPEC = rsSab("BIACOP")
    wKHALDZAL.HISMVTLIB1 = Trim(rsSab("LIBELE"))
    wKHALDZAL.HISMVTLIB2 = Trim(rsSab("REFCON"))
    wKHALDZAL.HISMVTDTRT = rsSab("TRAITA") & Format(rsSab("TRAITM"), "00") & Format(rsSab("TRAITJ"), "00")
    wKHALDZAL.HISMVTDVAL = rsSab("AAVAL") & Format(rsSab("MMVAL"), "00") & Format(rsSab("JJVAL"), "00")
    wKHALDZAL.HISMVTPIEN = rsSab("NUMPIE")
    wKHALDZAL.HISMVTPIES = rsSab("NOLIGN")
    K = InStr(wKHALDZAL.HISMVTLIB1, "(OpTrf")
    If K > 0 Then
        X = Mid$(wKHALDZAL.HISMVTLIB1, 1, K - 1)
        X = Replace(X, " ", "")
        X = Replace(X, ".", "")
        wKHALDZAL.HISMVTXREF = Replace(X, "/", "")
        If Len(wKHALDZAL.HISMVTXREF) > 16 Then wKHALDZAL.HISMVTXREF = Mid$(wKHALDZAL.HISMVTXREF, 1, 16)
        K = InStr(K + 6, wKHALDZAL.HISMVTLIB1, ":")
        If K > 0 Then
            K2 = InStr(K + 1, wKHALDZAL.HISMVTLIB1, ")")
            wKHALDZAL.HISMVTOPEN = Val(Mid$(wKHALDZAL.HISMVTLIB1, K + 1, K2 - K - 1))
            lenX = Len(wKHALDZAL.HISMVTLIB1)
            If lenX > K2 Then
                wKHALDZAL.HISMVTXBEN = Mid$(wKHALDZAL.HISMVTLIB1, K2 + 1, lenX - K2)
                wKHALDZAL.HISMVTXBEN = Replace(wKHALDZAL.HISMVTXBEN, "O/", "")
                wKHALDZAL.HISMVTXBEN = Replace(wKHALDZAL.HISMVTXBEN, "O:", "")
                wKHALDZAL.HISMVTXBEN = Replace(wKHALDZAL.HISMVTXBEN, "ORDRE", "")
                wKHALDZAL.HISMVTXBEN = Replace(wKHALDZAL.HISMVTXBEN, "RECU DE", "")
                wKHALDZAL.HISMVTXBEN = Replace(wKHALDZAL.HISMVTXBEN, "RECU D:", "")
                wKHALDZAL.HISMVTXBEN = Replace(wKHALDZAL.HISMVTXBEN, "RODRE", "")

            End If
        End If
    Else
        K = InStr(wKHALDZAL.HISMVTLIB2, "BIA  CDE-")
        If K > 0 Then
            wKHALDZAL.HISMVTOPEN = Val(Mid$(wKHALDZAL.HISMVTLIB2, K + 9, Len(wKHALDZAL.HISMVTLIB2) - K - 9))
        End If
    End If
    
    V = adoKHALDZAL_AddNew(rsKHALDZAL, wKHALDZAL)
    If Not IsNull(V) Then MsgBox V, vbCritical, "erreur : YSAAMSG_Import_AddNew " & mSeq

    rsSab.MoveNext
Loop

Close
Call lstErr_AddItem(frmKHALDZAL.lstErr, frmYGOSDOS0.cmdContext, "Nb : " & Seq)
cnSAB073Y_Close
Exit Sub

Error_Handle:
 MsgBox "erreur : HISMVTP0_Import" & xIn, vbCritical, Error
Close
cnSAB073Y_Close

End Sub
Public Sub HISMVTP0_Match()
Dim V, xSql As String
Dim xIn As String, K As Integer, K2 As Integer, lenX As Integer, Nb_ok As Long, Nb_Lu As Long
Dim kIn As Integer, Seq As Integer
On Error GoTo Error_Handle
Dim X As String
Dim mSeq As Integer
Dim K1 As Integer, I1 As Integer, I As Integer
Dim blnOk As Boolean, blnPrint As Boolean, blnSwift As Boolean, kPrint As Integer

Dim arrMatch(16000) As typeKHALDZAL_Match, arrMatch_Nb As Long
Dim curX As Currency, xTRN As String
Dim K_Match As Long, blnMultiple As Boolean
cnSAB073Y_Open

'rsKHALDZAL.Open "select * from KHALDZAL", cnSAB073Y, , adLockOptimistic

Call lstErr_AddItem(frmKHALDZAL.lstErr, frmYGOSDOS0.cmdContext, "match : load KHALDZAL")
xSql = "select Id,HISMVTMTD,HISMVTDVAL,HISMVTXREF,SAAMSGID from KHALDZAL  where HISMVTOPEC in ('V002','V003')" _
     & "  and HISMVTDTRT > 20001107 order by HISMVTMTD,HISMVTDVAL,HISMVTXREF"
Set rsKHALDZAL = cnSAB073Y.Execute(xSql)
arrMatch_Nb = 0
Do While Not rsKHALDZAL.EOF
    arrMatch_Nb = arrMatch_Nb + 1
    arrMatch(arrMatch_Nb).Id = rsKHALDZAL("Id")
    arrMatch(arrMatch_Nb).HISMVTMTD = -CCur(rsKHALDZAL("HISMVTMTD"))
    arrMatch(arrMatch_Nb).HISMVTDVAL = rsKHALDZAL("HISMVTDVAL")
    arrMatch(arrMatch_Nb).HISMVTXREF = rsKHALDZAL("HISMVTXREF")
    If Not IsNull(rsKHALDZAL("SAAMSGID")) Then
        arrMatch(arrMatch_Nb).SAAMSGID = rsKHALDZAL("SAAMSGID")
    Else
        arrMatch(arrMatch_Nb).SAAMSGID = 0
    End If
    

    rsKHALDZAL.MoveNext
Loop
'_________________________________________________________________________________
Call lstErr_AddItem(frmKHALDZAL.lstErr, frmYGOSDOS0.cmdContext, "match : load YSAAMSG0")
Nb_ok = 0: Nb_Lu = 0
xSql = "select * from YSAAMSG0 where SAAMsgType in('100','103','200','2*1','202') " _
     & " order by SAAMsgMt desc,SAAMsgDVal"
Set rsYSAAMSG0 = cnSAB073Y.Execute(xSql)

Do While Not rsYSAAMSG0.EOF
    Nb_Lu = Nb_Lu + 1
    curX = rsYSAAMSG0("SAAMsgMt")
    xTRN = Replace(rsYSAAMSG0("SAAMsgTRN"), " ", "")
    blnOk = False
    K_Match = 0: blnMultiple = False
    For K = 1 To arrMatch_Nb
        If curX > arrMatch(K).HISMVTMTD Then Exit For

        If curX = arrMatch(K).HISMVTMTD And xTRN = arrMatch(K).HISMVTXREF And arrMatch(K).SAAMSGID = 0 Then
            If blnOk Then
                blnMultiple = True
            Else
                K_Match = K
                blnOk = True
            End If
            
        End If
    Next K
'-----------------------------------------------------------
    If blnOk Then
    
        If Not blnMultiple Then
            Nb_ok = Nb_ok + 1
            arrMatch(K_Match).SAAMSGID = rsYSAAMSG0("SAAMsgId")
            arrMatch(K_Match).SAAMATCH = "="
            arrMatch(K_Match).SAAMSGTYPE = rsYSAAMSG0("SAAMsgType")
            arrMatch(K_Match).SAAMSGDTRT = rsYSAAMSG0("SAAMsgDtrt")
        Else
           'Debug.Print "Multi MT + TRN :"; rsYSAAMSG0("SAAMsgId"); rsYSAAMSG0("SAAMsgTYpe"); curX; xTRN
            Nb_ok = Nb_ok + 1
            arrMatch(K_Match).SAAMSGID = rsYSAAMSG0("SAAMsgId")
            arrMatch(K_Match).SAAMATCH = "+"
            arrMatch(K_Match).SAAMSGTYPE = rsYSAAMSG0("SAAMsgType")
            arrMatch(K_Match).SAAMSGDTRT = rsYSAAMSG0("SAAMsgDtrt")
        End If
    Else
'-----------------------------------------------------------
        blnOk = False
        K_Match = 0: blnMultiple = False
        For K = 1 To arrMatch_Nb
            If curX > arrMatch(K).HISMVTMTD Then Exit For
    
            If curX = arrMatch(K).HISMVTMTD And arrMatch(K).SAAMSGID = 0 Then
                If blnOk Then
                    blnMultiple = True
                Else
                    K_Match = K
                    blnOk = True
                End If
                
            End If
        Next K
'-----------------------------------------------------------
        If blnOk Then
             If Not blnMultiple Then
                Nb_ok = Nb_ok + 1
                arrMatch(K_Match).SAAMSGID = rsYSAAMSG0("SAAMsgId")
                arrMatch(K_Match).SAAMATCH = "#"
                arrMatch(K_Match).SAAMSGTYPE = rsYSAAMSG0("SAAMsgType")
                arrMatch(K_Match).SAAMSGDTRT = rsYSAAMSG0("SAAMsgDtrt")


            Else
                Debug.Print "Multi MT ? :"; rsYSAAMSG0("SAAMsgId"); rsYSAAMSG0("SAAMsgTYpe"); curX; xTRN
            End If
        Else
            Debug.Print "id ? :"; rsYSAAMSG0("SAAMsgId"), rsYSAAMSG0("SAAMsgId0"), rsYSAAMSG0("SAAMsgTYpe"), dateImp10_S(rsYSAAMSG0("SAAMsgDTrt")), Format$(curX, "### ### ### ##0.00"), xTRN
        End If

    End If
    rsYSAAMSG0.MoveNext
Loop
'___________________________________________________________________________
Call lstErr_AddItem(frmKHALDZAL.lstErr, frmYGOSDOS0.cmdContext, "MàJ Nb : " & Nb_ok & " / " & Nb_Lu)

For K = 1 To arrMatch_Nb
    If arrMatch(K).SAAMSGID > 0 Then
        xSql = "update KHALDZAL  set SAAMsgId = " & arrMatch(K).SAAMSGID _
             & " , SAAMsgType = '" & arrMatch(K).SAAMSGTYPE & "'" _
             & " , SAAMsgDTrt = " & arrMatch(K).SAAMSGDTRT _
             & " , SAAMATCH = '" & arrMatch(K).SAAMATCH & "'" _
             & " where id = " & arrMatch(K).Id
        Call FEU_ROUGE
        Set rsYSAAMSG0 = cnSAB073Y.Execute(xSql)
        Call FEU_VERT
    End If
Next K


    

'___________________________________________________________________________
Close
Call lstErr_AddItem(frmKHALDZAL.lstErr, frmYGOSDOS0.cmdContext, "Terminé")
cnSAB073Y_Close
Exit Sub

Error_Handle:
 MsgBox Error, vbCritical, "erreur : HISMVTP0_match"
Close
cnSAB073Y_Close

End Sub
Public Sub YSAAMVTLNK_KHALDZAL()
Dim V, xSql As String
Dim xIn As String, K As Integer, K2 As Integer, lenX As Integer, Nb_ok As Long, Nb_Lu As Long
Dim kIn As Integer, Seq As Integer
On Error GoTo Error_Handle
Dim X As String
Dim mSeq As Integer
Dim K1 As Integer, I1 As Integer, I As Integer
Dim blnOk As Boolean, blnPrint As Boolean, blnSwift As Boolean, kPrint As Integer

Dim curX As Currency, xTRN As String
Dim K_Match As Long, blnMultiple As Boolean
cnSAB073Y_Open

'rsKHALDZAL.Open "select * from KHALDZAL", cnSAB073Y, , adLockOptimistic

Call lstErr_AddItem(frmKHALDZAL.lstErr, frmYGOSDOS0.cmdContext, "match : load KHALDZAL")
xSql = "select * from  YSAAMVTLNK " _
     & "  where HISMVTID > 0"
Set rsYSAAMVTLNK = cnSAB073Y.Execute(xSql)

Do While Not rsYSAAMVTLNK.EOF
    Nb_Lu = Nb_Lu + 1
    xSql = "select * from YSAAMSG0 where SAAMsgId = " & rsYSAAMVTLNK("SAAMSGID")
    Set rsYSAAMSG0 = cnSAB073Y.Execute(xSql)
    
    If Not rsYSAAMSG0.EOF Then

        xSql = "update KHALDZAL  set SAAMsgId = " & rsYSAAMSG0("SAAMSGID") _
             & " , SAAMsgType = '" & rsYSAAMSG0("SAAMSGTYPE") & "'" _
             & " , SAAMsgDTrt = " & rsYSAAMSG0("SAAMSGDTRT") _
             & " , SAAMATCH = 'M'" _
             & " where id = " & rsYSAAMVTLNK("HISMVTID")
        Call FEU_ROUGE
        Set rsYSAAMSG0 = cnSAB073Y.Execute(xSql)
        Call FEU_VERT
        Nb_ok = Nb_ok + 1
    End If

    rsYSAAMVTLNK.MoveNext
Loop
'_________________________________________________________________________________
'___________________________________________________________________________
Call lstErr_AddItem(frmKHALDZAL.lstErr, frmYGOSDOS0.cmdContext, "MàJ Nb : " & Nb_ok & " / " & Nb_Lu)

    

'___________________________________________________________________________
Close
Call lstErr_AddItem(frmKHALDZAL.lstErr, frmYGOSDOS0.cmdContext, "Terminé")
cnSAB073Y_Close
Exit Sub

Error_Handle:
 MsgBox Error, vbCritical, "erreur : HISMVTP0_match"
Close
cnSAB073Y_Close

End Sub

Public Sub HISMVTP0_match_NOK()
Dim V, xSql As String
Dim xIn As String, K As Integer, K2 As Integer, lenX As Integer, Nb_ok As Long, Nb_Lu As Long
Dim kIn As Integer, Seq As Integer
On Error GoTo Error_Handle
Dim X As String
Dim mSeq As Integer
Dim K1 As Integer, I1 As Integer, I As Integer
Dim blnOk As Boolean, blnPrint As Boolean, blnSwift As Boolean, kPrint As Integer

Dim arrMatch(16000) As typeKHALDZAL_Match, arrMatch_Nb As Long
Dim curX As Currency, xTRN As String
Dim K_Match As Long, blnMultiple As Boolean
Dim wSAAMsgId As Long
Dim wYSAAMVTLNK As typeYSAAMVTLNK

cnSAB073Y_Open

rsYSAAMVTLNK.Open "select * from YSAAMVTLNK", cnSAB073Y, , adLockOptimistic

Call lstErr_AddItem(frmKHALDZAL.lstErr, frmYGOSDOS0.cmdContext, "match : load KHALDZAL")
xSql = "select Id,HISMVTMTD,HISMVTDVAL,HISMVTXREF,SAAMSGID from KHALDZAL  where HISMVTOPEC in ('V002','V003')" _
     & "  and HISMVTDTRT > 20001107 and SAAMSGID > 0 order by SAAMSGID"
Set rsKHALDZAL = cnSAB073Y.Execute(xSql)
arrMatch_Nb = 0
Do While Not rsKHALDZAL.EOF
    arrMatch_Nb = arrMatch_Nb + 1
    arrMatch(arrMatch_Nb).Id = rsKHALDZAL("Id")
    arrMatch(arrMatch_Nb).HISMVTMTD = -CCur(rsKHALDZAL("HISMVTMTD"))
    arrMatch(arrMatch_Nb).HISMVTDVAL = rsKHALDZAL("HISMVTDVAL")
    arrMatch(arrMatch_Nb).HISMVTXREF = rsKHALDZAL("HISMVTXREF")
    If Not IsNull(rsKHALDZAL("SAAMSGID")) Then
        arrMatch(arrMatch_Nb).SAAMSGID = rsKHALDZAL("SAAMSGID")
    Else
        arrMatch(arrMatch_Nb).SAAMSGID = 0
    End If
    

    rsKHALDZAL.MoveNext
Loop
'_________________________________________________________________________________
Call lstErr_AddItem(frmKHALDZAL.lstErr, frmYGOSDOS0.cmdContext, "match : load YSAAMSG0")
Nb_ok = 0: Nb_Lu = 0
xSql = "select * from YSAAMSG0 where SAAMsgType in('100','103','200','2*1','202') " _
     & " order by SAAMsgMt desc,SAAMsgDVal"
Set rsYSAAMSG0 = cnSAB073Y.Execute(xSql)

Do While Not rsYSAAMSG0.EOF
    Nb_Lu = Nb_Lu + 1
    wSAAMsgId = rsYSAAMSG0("SAAMsgId")
   blnOk = False
    K_Match = 0: blnMultiple = False
    For K = 1 To arrMatch_Nb
        If wSAAMsgId < arrMatch(K).SAAMSGID Then Exit For

        If wSAAMsgId = arrMatch(K).SAAMSGID Then
                blnOk = True
                Exit For
            End If
            
    Next K
'-----------------------------------------------------------
    If Not blnOk Then
    
        wYSAAMVTLNK.HISMVTID = 0
        wYSAAMVTLNK.SAAMSGID = wSAAMsgId
        V = adoYSAAMVTLNK_AddNew(rsYSAAMVTLNK, wYSAAMVTLNK)
        If Not IsNull(V) Then
            If InStr(V, "risque de doublons") = 0 Then MsgBox V, vbCritical, "erreur : HISMVTP0_match_NOK " & wSAAMsgId
        End If
    End If
    rsYSAAMSG0.MoveNext
Loop
'___________________________________________________________________________
Call lstErr_AddItem(frmKHALDZAL.lstErr, frmYGOSDOS0.cmdContext, "MàJ Nb : " & Nb_ok & " / " & Nb_Lu)


    

'___________________________________________________________________________
Close
Call lstErr_AddItem(frmKHALDZAL.lstErr, frmYGOSDOS0.cmdContext, "Terminé")
cnSAB073Y_Close
Exit Sub

Error_Handle:
 MsgBox Error, vbCritical, "erreur : HISMVTP0_match_NOK"
Close
cnSAB073Y_Close

End Sub

Public Sub YSAAMSG_201()
Dim V, xSql As String
Dim xIn As String, K As Integer, K2 As Integer, lenX As Integer, Nb_ok As Long, Nb_Lu As Long
Dim kIn As Integer, Seq As Integer
On Error GoTo Error_Handle
Dim X As String
Dim mSeq As Long
Dim K1 As Integer, I1 As Integer, I As Integer
Dim blnOk As Boolean, blnPrint As Boolean, blnSwift As Boolean, kPrint As Integer
Dim oldYSAAMSG0 As typeYSAAMSG0, newYSAAMSG0 As typeYSAAMSG0
Dim oldYSAAMSG1 As typeYSAAMSG1, newYSAAMSG1 As typeYSAAMSG1
Dim blnYSAAMSG0 As Boolean, blnYSAAMSG1 As Boolean

Dim addYSAAMSG0 As New ADODB.Recordset
Dim addYSAAMSG1 As New ADODB.Recordset

cnSAB073Y_Open

addYSAAMSG0.Open "select * from YSAAMSG0", cnSAB073Y, , adLockOptimistic
addYSAAMSG1.Open "select * from YSAAMSG1", cnSAB073Y, , adLockOptimistic

mSeq = 9000000

xSql = "select * from YSAAMSG0 where SAAMsgType = '201' " _
     & " order by SAAMsgMt desc,SAAMsgDVal"
Set rsYSAAMSG0 = cnSAB073Y.Execute(xSql)

Do While Not rsYSAAMSG0.EOF
    V = rsYSAAMSG0_GetBuffer(rsYSAAMSG0, oldYSAAMSG0)
    newYSAAMSG0 = oldYSAAMSG0
    blnYSAAMSG0 = False
    blnYSAAMSG1 = False
    xSql = "select * from YSAAMSG1  where SAAMsgId = " & oldYSAAMSG0.SAAMSGID _
         & " order by SAAMsgSeq"
    Set rsYSAAMSG1 = cnSAB073Y.Execute(xSql)
    
    Do While Not rsYSAAMSG1.EOF
    
        V = rsYSAAMSG1_GetBuffer(rsYSAAMSG1, oldYSAAMSG1)
        If oldYSAAMSG1.SAAMsgFld = "20" Then
            If blnYSAAMSG0 Then V = adoYSAAMSG0_AddNew(addYSAAMSG0, newYSAAMSG0)

            blnYSAAMSG0 = True: blnYSAAMSG1 = True
            mSeq = mSeq + 1
            newYSAAMSG0.SAAMSGID0 = oldYSAAMSG0.SAAMSGID
            newYSAAMSG0.SAAMSGID = mSeq
            newYSAAMSG0.SAAMSGTYPE = "2*1"
            newYSAAMSG0.SAAMsgMt = 0
            newYSAAMSG0.SAAMsgTRN = oldYSAAMSG1.SAAMSGTXT
        End If
        
        If oldYSAAMSG1.SAAMsgFld = "32" Then
            X = oldYSAAMSG1.SAAMSGTXT
            newYSAAMSG0.SAAMsgMt = CCur(Mid$(X, 4, Len(X) - 3))
        End If
       
        If blnYSAAMSG1 Then
            newYSAAMSG1 = oldYSAAMSG1
            newYSAAMSG1.SAAMSGID = mSeq
            V = adoYSAAMSG1_AddNew(addYSAAMSG1, newYSAAMSG1)
        End If
        
        rsYSAAMSG1.MoveNext
    Loop
    
    If blnYSAAMSG0 Then V = adoYSAAMSG0_AddNew(addYSAAMSG0, newYSAAMSG0)
    rsYSAAMSG0.MoveNext

Loop

'___________________________________________________________________________
Close
Call lstErr_AddItem(frmKHALDZAL.lstErr, frmYGOSDOS0.cmdContext, "Terminé")
cnSAB073Y_Close
Exit Sub

Error_Handle:
 MsgBox Error, vbCritical, "erreur : HISMVTP0_match"
Close
cnSAB073Y_Close

End Sub

Public Sub HISMVTP0_YSAAMSG1()
Dim V, xSql As String
Dim xIn As String, K As Integer, K2 As Integer, lenX As Integer, Nb_ok As Long, Nb_Lu As Long
Dim kIn As Integer, Seq As Integer
On Error GoTo Error_Handle
Dim X As String
Dim mSeq As Integer
Dim I1 As Integer, I As Integer
Dim blnOk As Boolean, blnPrint As Boolean, blnSwift As Boolean, kPrint As Integer

Dim arrMatch(16000) As typeKHALDZAL_Match, arrMatch_Nb As Long
Dim curX As Currency, xTRN As String
Dim K_Match As Long, blnMultiple As Boolean

Dim arrSAA(500) As typeYSAAMSG1, arrSAA_Nb As Long

Dim xKHALDZAL As typeKHALDZAL, xYSAAMSG1 As typeYSAAMSG1

Dim w57D As String, iLen As Integer
cnSAB073Y_Open


Call lstErr_AddItem(frmKHALDZAL.lstErr, frmYGOSDOS0.cmdContext, "match : load KHALDZAL")
xSql = "select Id,SAAMSGID from KHALDZAL  where SAAMSGID > 0 " _
     & " order by Id"
Set rsKHALDZAL = cnSAB073Y.Execute(xSql)
arrMatch_Nb = 0
Do While Not rsKHALDZAL.EOF
    arrMatch_Nb = arrMatch_Nb + 1
    arrMatch(arrMatch_Nb).Id = rsKHALDZAL("Id")
    arrMatch(arrMatch_Nb).SAAMSGID = rsKHALDZAL("SAAMSGID")
    
    rsKHALDZAL.MoveNext
Loop
'_________________________________________________________________________________
Call lstErr_AddItem(frmKHALDZAL.lstErr, frmYGOSDOS0.cmdContext, "match : load YSAAMSG1")
Nb_ok = 0: Nb_Lu = 0
'___________________________________________________________________________
For K = 1 To arrMatch_Nb
    arrSAA_Nb = 0
    rsKHALDZAL_Init xKHALDZAL
    w57D = ""
    xSql = "select * from YSAAMSG1  where SAAMsgId = " & arrMatch(K).SAAMSGID _
         & " order by SAAMsgSeq"
    Set rsYSAAMSG1 = cnSAB073Y.Execute(xSql)
    
    Do While Not rsYSAAMSG1.EOF
        'arrSAA_Nb = arrSAA_Nb + 1
        V = rsYSAAMSG1_GetBuffer(rsYSAAMSG1, xYSAAMSG1)
        X = xYSAAMSG1.SAAMsgFld & xYSAAMSG1.SAAMsgFldX & ": " & xYSAAMSG1.SAAMSGTXT
        If xKHALDZAL.SAAMSGTXT = "" Then
            xKHALDZAL.SAAMSGTXT = X
        Else
            xKHALDZAL.SAAMSGTXT = xKHALDZAL.SAAMSGTXT & vbCrLf & X
        End If
        Select Case xYSAAMSG1.SAAMsgFld
            Case "50": xKHALDZAL.SAAMSGXDO = xKHALDZAL.SAAMSGXDO & xYSAAMSG1.SAAMSGTXT
            Case "59": xKHALDZAL.SAAMSGXBEN = xKHALDZAL.SAAMSGXBEN & xYSAAMSG1.SAAMSGTXT
            Case "57":
                        If xYSAAMSG1.SAAMsgFldX = "A" Then xKHALDZAL.SAAMSGXPAY = HISMVTP0_YSAAMSG1_Pays(xYSAAMSG1.SAAMSGTXT)
                        If xYSAAMSG1.SAAMsgFldX = "D" Then w57D = xYSAAMSG1.SAAMSGTXT
            Case "58":  xKHALDZAL.SAAMSGXPAY = ""
                        If xYSAAMSG1.SAAMsgFldX = "A" Then xKHALDZAL.SAAMSGXPAY = HISMVTP0_YSAAMSG1_Pays(xYSAAMSG1.SAAMSGTXT)
                        If xYSAAMSG1.SAAMsgFldX = "D" Then w57D = xYSAAMSG1.SAAMSGTXT
            Case "72":
                I = InStr(xYSAAMSG1.SAAMSGTXT, "//DRAWEE")
                If I > 0 Then
                    I1 = InStr(I + 10, xYSAAMSG1.SAAMSGTXT, "_//")
                    If I1 > 0 Then xKHALDZAL.SAAMSGXDO = Trim(Mid$(xYSAAMSG1.SAAMSGTXT, I + 9, I1 - I - 9))
                    I = InStr(xYSAAMSG1.SAAMSGTXT, "//DRAWER")
                    If I > 0 Then
                        I1 = InStr(I + 10, xYSAAMSG1.SAAMSGTXT, "_//")
                        If I1 > 0 Then xKHALDZAL.SAAMSGXBEN = Trim(Mid$(xYSAAMSG1.SAAMSGTXT, I + 9, I1 - I - 9))
                    End If
                End If
        End Select
        
        rsYSAAMSG1.MoveNext
    Loop
'    If arrMatch(K).SAAMSGID = 320795 Then
'        Debug.Print "320795"
'    End If
If xKHALDZAL.SAAMSGXPAY = "" Then
        If arrMatch(K).SAAMSGID = 302632 Then xKHALDZAL.SAAMSGXPAY = "FR"
        If arrMatch(K).SAAMSGID = 320795 Then xKHALDZAL.SAAMSGXPAY = "FR"
        If arrMatch(K).SAAMSGID = 304333 Then xKHALDZAL.SAAMSGXPAY = "FR"
        If arrMatch(K).SAAMSGID = 309918 Then xKHALDZAL.SAAMSGXPAY = "FR"
        If arrMatch(K).SAAMSGID = 312944 Then xKHALDZAL.SAAMSGXPAY = "FR"
        If arrMatch(K).SAAMSGID = 309918 Then xKHALDZAL.SAAMSGXPAY = "FR"
        If arrMatch(K).SAAMSGID = 309931 Then xKHALDZAL.SAAMSGXPAY = "FR"
        If arrMatch(K).SAAMSGID = 312910 Then xKHALDZAL.SAAMSGXPAY = "FR"
        If arrMatch(K).SAAMSGID = 311100 Then xKHALDZAL.SAAMSGXPAY = "FR"
        If arrMatch(K).SAAMSGID = 310194 Then xKHALDZAL.SAAMSGXPAY = "FR"
        If arrMatch(K).SAAMSGID = 309798 Then xKHALDZAL.SAAMSGXPAY = "FR"
        If arrMatch(K).SAAMSGID = 302931 Then xKHALDZAL.SAAMSGXPAY = "FR"
        If arrMatch(K).SAAMSGID = 299918 Then xKHALDZAL.SAAMSGXPAY = "FR"
        If arrMatch(K).SAAMSGID = 321139 Then xKHALDZAL.SAAMSGXPAY = "FR"
        If arrMatch(K).SAAMSGID = 379514 Then xKHALDZAL.SAAMSGXPAY = "FR"
        If arrMatch(K).SAAMSGID = 322876 Then xKHALDZAL.SAAMSGXPAY = "FR"
        If arrMatch(K).SAAMSGID = 326380 Then xKHALDZAL.SAAMSGXPAY = "FR"
        If arrMatch(K).SAAMSGID = 330116 Then xKHALDZAL.SAAMSGXPAY = "FR"
        If arrMatch(K).SAAMSGID = 330144 Then xKHALDZAL.SAAMSGXPAY = "FR"
        If arrMatch(K).SAAMSGID = 336300 Then xKHALDZAL.SAAMSGXPAY = "FR"
        If arrMatch(K).SAAMSGID = 339428 Then xKHALDZAL.SAAMSGXPAY = "FR"
        If arrMatch(K).SAAMSGID = 341689 Then xKHALDZAL.SAAMSGXPAY = "FR"
        If arrMatch(K).SAAMSGID = 367557 Then xKHALDZAL.SAAMSGXPAY = "FR"
        If arrMatch(K).SAAMSGID = 337691 Then xKHALDZAL.SAAMSGXPAY = "FR"
        If arrMatch(K).SAAMSGID = 346433 Then xKHALDZAL.SAAMSGXPAY = "FR"
        If arrMatch(K).SAAMSGID = 348054 Then xKHALDZAL.SAAMSGXPAY = "FR"
        
        If arrMatch(K).SAAMSGID = 351905 Then xKHALDZAL.SAAMSGXPAY = "FR"
        If arrMatch(K).SAAMSGID = 354084 Then xKHALDZAL.SAAMSGXPAY = "FR"
        If arrMatch(K).SAAMSGID = 358073 Then xKHALDZAL.SAAMSGXPAY = "FR"
        If arrMatch(K).SAAMSGID = 359053 Then xKHALDZAL.SAAMSGXPAY = "FR"
        If arrMatch(K).SAAMSGID = 360090 Then xKHALDZAL.SAAMSGXPAY = "FR"
        If arrMatch(K).SAAMSGID = 364518 Then xKHALDZAL.SAAMSGXPAY = "FR"
        If arrMatch(K).SAAMSGID = 375590 Then xKHALDZAL.SAAMSGXPAY = "FR"
        
        If arrMatch(K).SAAMSGID = 375540 Then xKHALDZAL.SAAMSGXPAY = "FR"
        If arrMatch(K).SAAMSGID = 375588 Then xKHALDZAL.SAAMSGXPAY = "FR"
        If arrMatch(K).SAAMSGID = 376166 Then xKHALDZAL.SAAMSGXPAY = "FR"
        If arrMatch(K).SAAMSGID = 379529 Then xKHALDZAL.SAAMSGXPAY = "FR"
        If arrMatch(K).SAAMSGID = 379742 Then xKHALDZAL.SAAMSGXPAY = "FR"
        If arrMatch(K).SAAMSGID = 379505 Then xKHALDZAL.SAAMSGXPAY = "FR"
        If arrMatch(K).SAAMSGID = 382961 Then xKHALDZAL.SAAMSGXPAY = "FR"
        If arrMatch(K).SAAMSGID = 331964 Then xKHALDZAL.SAAMSGXPAY = "FR"
        
        If arrMatch(K).SAAMSGID = 364991 Then xKHALDZAL.SAAMSGXPAY = "FR"
        If arrMatch(K).SAAMSGID = 389287 Then xKHALDZAL.SAAMSGXPAY = "RO"
        If arrMatch(K).SAAMSGID = 387710 Then xKHALDZAL.SAAMSGXPAY = "EG"
        If arrMatch(K).SAAMSGID = 389305 Then xKHALDZAL.SAAMSGXPAY = "HU"
        If arrMatch(K).SAAMSGID = 352836 Then xKHALDZAL.SAAMSGXPAY = "PL"
        If arrMatch(K).SAAMSGID = 346427 Then xKHALDZAL.SAAMSGXPAY = "PL"
        If arrMatch(K).SAAMSGID = 346428 Then xKHALDZAL.SAAMSGXPAY = "PL"
        If arrMatch(K).SAAMSGID = 351804 Then xKHALDZAL.SAAMSGXPAY = "LB"
        If arrMatch(K).SAAMSGID = 349727 Then xKHALDZAL.SAAMSGXPAY = "CN"
        
        If arrMatch(K).SAAMSGID = 341725 Then xKHALDZAL.SAAMSGXPAY = "TR"
        If arrMatch(K).SAAMSGID = 343758 Then xKHALDZAL.SAAMSGXPAY = "TR"
        If arrMatch(K).SAAMSGID = 342705 Then xKHALDZAL.SAAMSGXPAY = "GB"
        If arrMatch(K).SAAMSGID = 339431 Then xKHALDZAL.SAAMSGXPAY = "TR"
        If arrMatch(K).SAAMSGID = 338906 Then xKHALDZAL.SAAMSGXPAY = "TR"
        If arrMatch(K).SAAMSGID = 336130 Then xKHALDZAL.SAAMSGXPAY = "PL"
        If arrMatch(K).SAAMSGID = 336292 Then xKHALDZAL.SAAMSGXPAY = "TR"
        If arrMatch(K).SAAMSGID = 323245 Then xKHALDZAL.SAAMSGXPAY = "TR"
        If arrMatch(K).SAAMSGID = 390264 Then xKHALDZAL.SAAMSGXPAY = "CH"
        
        If arrMatch(K).SAAMSGID = 372157 Then xKHALDZAL.SAAMSGXPAY = "TR"
        If arrMatch(K).SAAMSGID = 364244 Then xKHALDZAL.SAAMSGXPAY = "TR"
        If arrMatch(K).SAAMSGID = 364312 Then xKHALDZAL.SAAMSGXPAY = "TN"
        If arrMatch(K).SAAMSGID = 360092 Then xKHALDZAL.SAAMSGXPAY = "TN"
        If arrMatch(K).SAAMSGID = 351655 Then xKHALDZAL.SAAMSGXPAY = "CN"
        
        If arrMatch(K).SAAMSGID = 349723 Then xKHALDZAL.SAAMSGXPAY = "LB"
        If arrMatch(K).SAAMSGID = 331761 Then xKHALDZAL.SAAMSGXPAY = "TR"
        If arrMatch(K).SAAMSGID = 299959 Then xKHALDZAL.SAAMSGXPAY = "CH"
        If arrMatch(K).SAAMSGID = 303086 Then xKHALDZAL.SAAMSGXPAY = "AE"
        If arrMatch(K).SAAMSGID = 303086 Then xKHALDZAL.SAAMSGXPAY = "FR"
        If arrMatch(K).SAAMSGID = 303123 Then xKHALDZAL.SAAMSGXPAY = "TR"
        If arrMatch(K).SAAMSGID = 304295 Then xKHALDZAL.SAAMSGXPAY = "ES"
        If arrMatch(K).SAAMSGID = 304645 Then xKHALDZAL.SAAMSGXPAY = "ES"
        If arrMatch(K).SAAMSGID = 304624 Then xKHALDZAL.SAAMSGXPAY = "PT"
        If arrMatch(K).SAAMSGID = 304282 Then xKHALDZAL.SAAMSGXPAY = "ES"
        If arrMatch(K).SAAMSGID = 297019 Then xKHALDZAL.SAAMSGXPAY = "TR"
End If
        If xKHALDZAL.SAAMSGXPAY = "" Then
            iLen = Len(w57D)
            If iLen > 10 Then
            
                If Mid$(w57D, iLen - 2, 3) = " FR" Then xKHALDZAL.SAAMSGXPAY = "FR"
                If Mid$(w57D, iLen - 2, 3) = " TN" Then xKHALDZAL.SAAMSGXPAY = "TN"
                If Mid$(w57D, iLen - 2, 3) = " UK" Then xKHALDZAL.SAAMSGXPAY = "GB"
                If Mid$(w57D, iLen - 2, 3) = " QA" Then xKHALDZAL.SAAMSGXPAY = "QA"
                
                If Mid$(w57D, iLen - 3, 4) = " UAE" Then xKHALDZAL.SAAMSGXPAY = "AE"
                If Mid$(w57D, iLen - 3, 4) = "_UAE" Then xKHALDZAL.SAAMSGXPAY = "AE"
 
                If Mid$(w57D, iLen - 5, 6) = "FRANCE" Then xKHALDZAL.SAAMSGXPAY = "FR"
                If Mid$(w57D, iLen - 5, 6) = "CANADA" Then xKHALDZAL.SAAMSGXPAY = "CA"
                If Mid$(w57D, iLen - 5, 6) = "ITALIE" Then xKHALDZAL.SAAMSGXPAY = "IT"
                If Mid$(w57D, iLen - 5, 6) = "EGYPTE" Then xKHALDZAL.SAAMSGXPAY = "EG"
                If Mid$(w57D, iLen - 5, 6) = "POLAND" Then xKHALDZAL.SAAMSGXPAY = "PL"
                If Mid$(w57D, iLen - 5, 6) = "TURKEY" Then xKHALDZAL.SAAMSGXPAY = "TR"
                If Mid$(w57D, iLen - 5, 6) = "GREECE" Then xKHALDZAL.SAAMSGXPAY = "GR"
                If Mid$(w57D, iLen - 5, 6) = "GENEVA" Then xKHALDZAL.SAAMSGXPAY = "CH"
                If Mid$(w57D, iLen - 5, 6) = "JORDAN" Then xKHALDZAL.SAAMSGXPAY = "JO"
                
                If Mid$(w57D, iLen - 4, 5) = "PARIS" Then xKHALDZAL.SAAMSGXPAY = "FR"
                If Mid$(w57D, iLen - 4, 5) = "EGYPT" Then xKHALDZAL.SAAMSGXPAY = "EG"
                If Mid$(w57D, iLen - 4, 5) = "CHINA" Then xKHALDZAL.SAAMSGXPAY = "CN"
                If Mid$(w57D, iLen - 4, 5) = "CHINE" Then xKHALDZAL.SAAMSGXPAY = "CN"
                If Mid$(w57D, iLen - 4, 5) = "SPAIN" Then xKHALDZAL.SAAMSGXPAY = "ES"
                If Mid$(w57D, iLen - 4, 5) = "TUNIS" Then xKHALDZAL.SAAMSGXPAY = "TN"
                If Mid$(w57D, iLen - 4, 5) = "MALTA" Then xKHALDZAL.SAAMSGXPAY = "MT"
                If Mid$(w57D, iLen - 4, 5) = " WIEN" Then xKHALDZAL.SAAMSGXPAY = "AT"
                If Mid$(w57D, iLen - 4, 5) = "ITALY" Then xKHALDZAL.SAAMSGXPAY = "IT"
                If Mid$(w57D, iLen - 4, 5) = "MAROC" Then xKHALDZAL.SAAMSGXPAY = "MA"
                If Mid$(w57D, iLen - 4, 5) = "CAIRO" Then xKHALDZAL.SAAMSGXPAY = "EG"
                If Mid$(w57D, iLen - 4, 5) = "U A E" Then xKHALDZAL.SAAMSGXPAY = "AE"
                If Mid$(w57D, iLen - 4, 5) = "LIBAN" Then xKHALDZAL.SAAMSGXPAY = "LB"
                
                If Mid$(w57D, iLen - 6, 7) = "DENMARK" Then xKHALDZAL.SAAMSGXPAY = "DK"
                If Mid$(w57D, iLen - 6, 7) = "TURQUIE" Then xKHALDZAL.SAAMSGXPAY = "TR"
                If Mid$(w57D, iLen - 6, 7) = "ENGLAND" Then xKHALDZAL.SAAMSGXPAY = "GB"
                If Mid$(w57D, iLen - 6, 7) = "POLOGNE" Then xKHALDZAL.SAAMSGXPAY = "PL"
                If Mid$(w57D, iLen - 6, 7) = "AVIGNON" Then xKHALDZAL.SAAMSGXPAY = "FR"
                If Mid$(w57D, iLen - 6, 7) = "SLOVENIE" Then xKHALDZAL.SAAMSGXPAY = "SI"
                If Mid$(w57D, iLen - 6, 7) = "PAYERNE" Then xKHALDZAL.SAAMSGXPAY = "CH"
                If Mid$(w57D, iLen - 6, 7) = "ROMANIA" Then xKHALDZAL.SAAMSGXPAY = "RO"
                If Mid$(w57D, iLen - 6, 7) = "TCHEQUE" Then xKHALDZAL.SAAMSGXPAY = "CZ"
                If Mid$(w57D, iLen - 6, 7) = "TUNISIA" Then xKHALDZAL.SAAMSGXPAY = "TN"
                If Mid$(w57D, iLen - 6, 7) = "HONGRIE" Then xKHALDZAL.SAAMSGXPAY = "HU"
                If Mid$(w57D, iLen - 6, 7) = "ESPAGNE" Then xKHALDZAL.SAAMSGXPAY = "ES"
 
                If Mid$(w57D, iLen - 7, 8) = "DENEMARK" Then xKHALDZAL.SAAMSGXPAY = "DK"
                If Mid$(w57D, iLen - 7, 8) = "HONGKONG" Then xKHALDZAL.SAAMSGXPAY = "HK"
                If Mid$(w57D, iLen - 7, 8) = "LAUSANNE" Then xKHALDZAL.SAAMSGXPAY = "CH"
                If Mid$(w57D, iLen - 7, 8) = "ISTAMBUL" Then xKHALDZAL.SAAMSGXPAY = "TR"
                If Mid$(w57D, iLen - 7, 8) = "SLOVENIE" Then xKHALDZAL.SAAMSGXPAY = "SI"
                If Mid$(w57D, iLen - 7, 8) = "ESPAGNE" Then xKHALDZAL.SAAMSGXPAY = "ES"
                If Mid$(w57D, iLen - 7, 8) = "EMIRATES" Then xKHALDZAL.SAAMSGXPAY = "AE"
                If Mid$(w57D, iLen - 7, 8) = "THAILAND" Then xKHALDZAL.SAAMSGXPAY = "TH"

                If Mid$(w57D, iLen - 9, 10) = "LUXEMBOURG" Then xKHALDZAL.SAAMSGXPAY = "LU"
                If Mid$(w57D, iLen - 10, 11) = "SWITZERLAND" Then xKHALDZAL.SAAMSGXPAY = "CH"
                
                
                If xKHALDZAL.SAAMSGXPAY = "" And iLen > 20 Then
                    If InStr(iLen - 21, w57D, "STRASBOURG") > 0 Then xKHALDZAL.SAAMSGXPAY = "FR"
                    If InStr(iLen - 21, w57D, "MARSEILLE") > 0 Then xKHALDZAL.SAAMSGXPAY = "FR"
                    If InStr(iLen - 21, w57D, "CEDEX") > 0 Then xKHALDZAL.SAAMSGXPAY = "FR"
                    If InStr(iLen - 21, w57D, "GRENOBLE") > 0 Then xKHALDZAL.SAAMSGXPAY = "FR"
                    If InStr(iLen - 21, w57D, "LYON") > 0 Then xKHALDZAL.SAAMSGXPAY = "FR"
                    If InStr(iLen - 21, w57D, "TROYES") > 0 Then xKHALDZAL.SAAMSGXPAY = "FR"
                    If InStr(iLen - 21, w57D, "PARIS") > 0 Then xKHALDZAL.SAAMSGXPAY = "FR"
                    If InStr(iLen - 21, w57D, "NANTES") > 0 Then xKHALDZAL.SAAMSGXPAY = "FR"
                    If InStr(iLen - 21, w57D, "LA BRIE") > 0 Then xKHALDZAL.SAAMSGXPAY = "FR"
                    If InStr(iLen - 21, w57D, "VARIN BERNI") > 0 Then xKHALDZAL.SAAMSGXPAY = "FR"
                    If InStr(iLen - 21, w57D, "LES ABRETS") > 0 Then xKHALDZAL.SAAMSGXPAY = "FR"
                    If InStr(iLen - 21, w57D, "COMPIEGNE") > 0 Then xKHALDZAL.SAAMSGXPAY = "FR"
                    If InStr(iLen - 21, w57D, "PERONNE") > 0 Then xKHALDZAL.SAAMSGXPAY = "FR"
                    If InStr(iLen - 21, w57D, "france") > 0 Then xKHALDZAL.SAAMSGXPAY = "FR"
                    If InStr(iLen - 21, w57D, "CREDIT DU NORD") > 0 Then xKHALDZAL.SAAMSGXPAY = "FR"
                    If InStr(iLen - 21, w57D, "SAINTES") > 0 Then xKHALDZAL.SAAMSGXPAY = "FR"
                    If InStr(iLen - 21, w57D, "MIRECOURT") > 0 Then xKHALDZAL.SAAMSGXPAY = "FR"
                    If InStr(iLen - 21, w57D, "DIJON") > 0 Then xKHALDZAL.SAAMSGXPAY = "FR"
                    If InStr(iLen - 21, w57D, "SALON DE PROVENCE") > 0 Then xKHALDZAL.SAAMSGXPAY = "FR"
                    If InStr(iLen - 21, w57D, "SAINT ETIENNE") > 0 Then xKHALDZAL.SAAMSGXPAY = "FR"
                    If InStr(iLen - 21, w57D, "COURBEVOIE") > 0 Then xKHALDZAL.SAAMSGXPAY = "FR"
                    If InStr(iLen - 21, w57D, "IVRY SUR SEINE") > 0 Then xKHALDZAL.SAAMSGXPAY = "FR"
                    If InStr(iLen - 21, w57D, "ROUBAIX") > 0 Then xKHALDZAL.SAAMSGXPAY = "FR"
                    If InStr(iLen - 21, w57D, "NICE") > 0 Then xKHALDZAL.SAAMSGXPAY = "FR"
                    If InStr(iLen - 21, w57D, "AUBERVILLIERS") > 0 Then xKHALDZAL.SAAMSGXPAY = "FR"
                    If InStr(iLen - 21, w57D, "VALENCIENNES") > 0 Then xKHALDZAL.SAAMSGXPAY = "FR"
                    If InStr(iLen - 21, w57D, "NANCY") > 0 Then xKHALDZAL.SAAMSGXPAY = "FR"
                    If InStr(iLen - 21, w57D, "BOURG EN BRESSE") > 0 Then xKHALDZAL.SAAMSGXPAY = "FR"
                    If InStr(iLen - 21, w57D, "BORDEAUX") > 0 Then xKHALDZAL.SAAMSGXPAY = "FR"
                    If InStr(iLen - 21, w57D, "CERGY") > 0 Then xKHALDZAL.SAAMSGXPAY = "FR"
                    If InStr(iLen - 21, w57D, "VALENCE  FR") > 0 Then xKHALDZAL.SAAMSGXPAY = "FR"
                    If InStr(iLen - 21, w57D, "MONTPELLIER") > 0 Then xKHALDZAL.SAAMSGXPAY = "FR"
                    If InStr(iLen - 21, w57D, "REIMS") > 0 Then xKHALDZAL.SAAMSGXPAY = "FR"
                    If InStr(iLen - 21, w57D, "MUTUEL DE BRETAGNE") > 0 Then xKHALDZAL.SAAMSGXPAY = "FR"
                    If InStr(iLen - 21, w57D, "CHAVORNAY") > 0 Then xKHALDZAL.SAAMSGXPAY = "FR"
                    If InStr(iLen - 21, w57D, "DE PICARDIE") > 0 Then xKHALDZAL.SAAMSGXPAY = "FR"
                    If InStr(iLen - 21, w57D, "SADI-CARNOT") > 0 Then xKHALDZAL.SAAMSGXPAY = "FR"
                    If InStr(iLen - 21, w57D, "SADI CARNOT") > 0 Then xKHALDZAL.SAAMSGXPAY = "FR"
                    If InStr(iLen - 21, w57D, "BRETONNEUX") > 0 Then xKHALDZAL.SAAMSGXPAY = "FR"
                    If InStr(iLen - 21, w57D, "MONTROUGE") > 0 Then xKHALDZAL.SAAMSGXPAY = "FR"
                    If InStr(iLen - 21, w57D, "GERCY") > 0 Then xKHALDZAL.SAAMSGXPAY = "FR"
                    If InStr(iLen - 21, w57D, "HAGUENAU") > 0 Then xKHALDZAL.SAAMSGXPAY = "FR"
                    If InStr(iLen - 21, w57D, "TAVERNY") > 0 Then xKHALDZAL.SAAMSGXPAY = "FR"
                    If InStr(iLen - 21, w57D, "EPINAY SUR SEINE") > 0 Then xKHALDZAL.SAAMSGXPAY = "FR"
                    If InStr(iLen - 21, w57D, "BOUTHEON") > 0 Then xKHALDZAL.SAAMSGXPAY = "FR"
                    If InStr(iLen - 21, w57D, "AUBAGNE") > 0 Then xKHALDZAL.SAAMSGXPAY = "FR"
                    If InStr(iLen - 21, w57D, "ROISSY") > 0 Then xKHALDZAL.SAAMSGXPAY = "FR"
                    If InStr(iLen - 21, w57D, "SAINT-ETIENNE") > 0 Then xKHALDZAL.SAAMSGXPAY = "FR"
                    If InStr(iLen - 21, w57D, "RENNES") > 0 Then xKHALDZAL.SAAMSGXPAY = "FR"
                    If InStr(iLen - 21, w57D, "VIRIAT") > 0 Then xKHALDZAL.SAAMSGXPAY = "FR"
                    If InStr(iLen - 21, w57D, "PERPIGNAN") > 0 Then xKHALDZAL.SAAMSGXPAY = "FR"
                    If InStr(iLen - 21, w57D, "MOUVAUX") > 0 Then xKHALDZAL.SAAMSGXPAY = "FR"
                    If InStr(iLen - 21, w57D, "ANNECY") > 0 Then xKHALDZAL.SAAMSGXPAY = "FR"
                    If InStr(iLen - 21, w57D, "SENLIS") > 0 Then xKHALDZAL.SAAMSGXPAY = "FR"
                    If InStr(iLen - 21, w57D, "ALPES DU SUD") > 0 Then xKHALDZAL.SAAMSGXPAY = "FR"
                    
                    
                    If InStr(iLen - 21, w57D, "MONACO") > 0 Then xKHALDZAL.SAAMSGXPAY = "MC"
                    If InStr(iLen - 21, w57D, "MONTE CARLO") > 0 Then xKHALDZAL.SAAMSGXPAY = "MC"
                    If InStr(iLen - 21, w57D, "AUSTRIA") > 0 Then xKHALDZAL.SAAMSGXPAY = "AT"
                    If InStr(iLen - 21, w57D, "LONDON") > 0 Then xKHALDZAL.SAAMSGXPAY = "GB"
                    If InStr(iLen - 21, w57D, "BEYROUT") > 0 Then xKHALDZAL.SAAMSGXPAY = "LB"
                    If InStr(iLen - 21, w57D, "BEIRUT") > 0 Then xKHALDZAL.SAAMSGXPAY = "LB"
                    If InStr(iLen - 21, w57D, "HONG KONG") > 0 Then xKHALDZAL.SAAMSGXPAY = "HK"
                    If InStr(iLen - 21, w57D, "MADRID") > 0 Then xKHALDZAL.SAAMSGXPAY = "ES"
                    If InStr(iLen - 21, w57D, "COPENHAGEN") > 0 Then xKHALDZAL.SAAMSGXPAY = "DK"
                    If InStr(iLen - 21, w57D, "GERMANY") > 0 Then xKHALDZAL.SAAMSGXPAY = "DE"
                    If InStr(iLen - 21, w57D, "PORTUGAL") > 0 Then xKHALDZAL.SAAMSGXPAY = "PT"
                    If InStr(iLen - 21, w57D, "ALGERIE") > 0 Then xKHALDZAL.SAAMSGXPAY = "DZ"
                    If InStr(iLen - 21, w57D, "HOLLAND") > 0 Then xKHALDZAL.SAAMSGXPAY = "NL"
                    If InStr(iLen - 21, w57D, "CANADA") > 0 Then xKHALDZAL.SAAMSGXPAY = "CA"
                    If InStr(iLen - 21, w57D, "TUNISIE") > 0 Then xKHALDZAL.SAAMSGXPAY = "TN"
                    If InStr(iLen - 21, w57D, "BELGIQUE") > 0 Then xKHALDZAL.SAAMSGXPAY = "BE"
                    If InStr(iLen - 21, w57D, "ENG 6-8,A-1090_VIENNE") > 0 Then xKHALDZAL.SAAMSGXPAY = "AT"
                    If InStr(iLen - 21, w57D, "RADOM") > 0 Then xKHALDZAL.SAAMSGXPAY = "PL"
                    If InStr(iLen - 21, w57D, "TUNIS BELVEDERE") > 0 Then xKHALDZAL.SAAMSGXPAY = "TN"
                    If InStr(iLen - 21, w57D, "LISBOA") > 0 Then xKHALDZAL.SAAMSGXPAY = "PT"
                    If InStr(iLen - 21, w57D, "DANEMARK") > 0 Then xKHALDZAL.SAAMSGXPAY = "DK"
                    If InStr(iLen - 21, w57D, "ISTANBUL") > 0 Then xKHALDZAL.SAAMSGXPAY = "TR"
                    If InStr(iLen - 21, w57D, "STOCKHOLM") > 0 Then xKHALDZAL.SAAMSGXPAY = "SE"
                    If InStr(iLen - 21, w57D, "LEBANON") > 0 Then xKHALDZAL.SAAMSGXPAY = "LB"
                    If InStr(iLen - 21, w57D, "BRUXELLES") > 0 Then xKHALDZAL.SAAMSGXPAY = "BE"
                    If InStr(iLen - 21, w57D, "SINGAPORE") > 0 Then xKHALDZAL.SAAMSGXPAY = "SG"
                    If InStr(iLen - 21, w57D, "UNITED ARAB EMIRATES") > 0 Then xKHALDZAL.SAAMSGXPAY = "AE"
                    If InStr(iLen - 21, w57D, "DUBAI") > 0 Then xKHALDZAL.SAAMSGXPAY = "AE"
                    If InStr(iLen - 21, w57D, "CASABLANCA") > 0 Then xKHALDZAL.SAAMSGXPAY = "MA"
                    If InStr(iLen - 21, w57D, "TURKEY") > 0 Then xKHALDZAL.SAAMSGXPAY = "TR"
                    If InStr(iLen - 21, w57D, "SAUDI ARABIA") > 0 Then xKHALDZAL.SAAMSGXPAY = "SA"
                    If InStr(iLen - 21, w57D, "JADDAH") > 0 Then xKHALDZAL.SAAMSGXPAY = "SA"
                    If InStr(iLen - 21, w57D, "TUNIS") > 0 Then xKHALDZAL.SAAMSGXPAY = "TN"
                    If InStr(iLen - 21, w57D, "SYRIE") > 0 Then xKHALDZAL.SAAMSGXPAY = "SY"
                    If InStr(iLen - 21, w57D, "GENEVA") > 0 Then xKHALDZAL.SAAMSGXPAY = "CH"
                    If InStr(iLen - 21, w57D, "BANGKOK") > 0 Then xKHALDZAL.SAAMSGXPAY = "TH"
                    If InStr(iLen - 21, w57D, "GENEVE") > 0 Then xKHALDZAL.SAAMSGXPAY = "CH"
                    If InStr(iLen - 21, w57D, "TURKEY") > 0 Then xKHALDZAL.SAAMSGXPAY = "TR"
                    

                    If iLen > 30 And xKHALDZAL.SAAMSGXPAY = "" Then
                        If InStr(iLen - 30, w57D, "ISTANBUL") > 0 Then xKHALDZAL.SAAMSGXPAY = "TR"
                        If InStr(iLen - 30, w57D, "KATOWICE") > 0 Then xKHALDZAL.SAAMSGXPAY = "PL"
                        If InStr(iLen - 30, w57D, "ZURICH") > 0 Then xKHALDZAL.SAAMSGXPAY = "CH"
                        If InStr(iLen - 30, w57D, "GREECE") > 0 Then xKHALDZAL.SAAMSGXPAY = "GR"
                        If InStr(iLen - 30, w57D, "CHALONS SUR CHAMPAGNE") > 0 Then xKHALDZAL.SAAMSGXPAY = "FR"
                        If InStr(iLen - 30, w57D, "BEYROUTH") > 0 Then xKHALDZAL.SAAMSGXPAY = "LB"
                        If InStr(iLen - 30, w57D, "BEYROTH") > 0 Then xKHALDZAL.SAAMSGXPAY = "LB"
                   End If
                End If
              
                
           End If
        End If
If xKHALDZAL.SAAMSGXPAY = "" Then
    'Debug.Print arrMatch(K).SAAMSGID; ":"; Mid$(w57D, iLen - 20, 21)
    Debug.Print arrMatch(K).SAAMSGID; ":"; w57D
End If
        
        xKHALDZAL.SAAMSGXDO = Replace(xKHALDZAL.SAAMSGXDO, "'", "''")
        If Mid$(xKHALDZAL.SAAMSGXDO, 1, 1) = "/" Then
            I = InStr(xKHALDZAL.SAAMSGXDO, "_")
            If I > 0 Then xKHALDZAL.SAAMSGXDO = Mid$(xKHALDZAL.SAAMSGXDO, I + 1, Len(xKHALDZAL.SAAMSGXDO) - I)
        End If
        
            
        xKHALDZAL.SAAMSGXBEN = Replace(xKHALDZAL.SAAMSGXBEN, "'", "''")
        If Mid$(xKHALDZAL.SAAMSGXBEN, 1, 1) = "/" Then
            I = InStr(xKHALDZAL.SAAMSGXBEN, "_")
            If I > 0 Then xKHALDZAL.SAAMSGXBEN = Mid$(xKHALDZAL.SAAMSGXBEN, I + 1, Len(xKHALDZAL.SAAMSGXBEN) - I)
        End If
        
        xKHALDZAL.SAAMSGTXT = Replace(xKHALDZAL.SAAMSGTXT, "'", "''")
        
        xSql = "update KHALDZAL  set SAAMSGXDO = '" & xKHALDZAL.SAAMSGXDO & "'" _
             & " , SAAMSGXBEN = '" & xKHALDZAL.SAAMSGXBEN & "'" _
             & " , SAAMSGXPAY = '" & xKHALDZAL.SAAMSGXPAY & "'" _
             & " , SAAMSGTXT = '" & xKHALDZAL.SAAMSGTXT & "'" _
             & " where id = " & arrMatch(K).Id
        Call FEU_ROUGE
        Set rsYSAAMSG0 = cnSAB073Y.Execute(xSql) '
        Call FEU_VERT

Next K


    

'___________________________________________________________________________
Close
Call lstErr_AddItem(frmKHALDZAL.lstErr, frmYGOSDOS0.cmdContext, "Nb : " & Nb_ok & " / " & Nb_Lu)
cnSAB073Y_Close
Exit Sub

Error_Handle:
 MsgBox Error, vbCritical, "erreur : HISMVTP0_YSAAMSG1"
Close
cnSAB073Y_Close



End Sub

Public Sub rsKHALDZAL_Init(rsKHALDZAL As typeKHALDZAL)
rsKHALDZAL.Id = 0
rsKHALDZAL.HISMVTDEV = ""
rsKHALDZAL.HISMVTCPT = ""
rsKHALDZAL.HISMVTOPEC = ""
rsKHALDZAL.HISMVTOPEN = 0
rsKHALDZAL.HISMVTMTD = 0
rsKHALDZAL.HISMVTLIB1 = ""
rsKHALDZAL.HISMVTLIB2 = ""
rsKHALDZAL.HISMVTDTRT = 0
rsKHALDZAL.HISMVTDVAL = 0
rsKHALDZAL.HISMVTPIEN = 0
rsKHALDZAL.HISMVTPIES = 0
rsKHALDZAL.HISMVTXBEN = ""
rsKHALDZAL.HISMVTXREF = ""
rsKHALDZAL.SAAMSGID = 0
rsKHALDZAL.SAAMSGXDO = ""
rsKHALDZAL.SAAMSGXBEN = ""
rsKHALDZAL.SAAMSGXPAY = ""
rsKHALDZAL.SAAMSGTXT = ""

rsKHALDZAL.SAAMATCH = ""
rsKHALDZAL.SAAMSGTYPE = ""
rsKHALDZAL.SAAMSGDTRT = 0
rsKHALDZAL.SCANLINK = ""
rsKHALDZAL.SCANLINK2 = ""

End Sub

'---------------------------------------------------------
Public Function adoKHALDZAL_AddNew(rsADO As ADODB.Recordset, rsKHALDZAL As typeKHALDZAL)
'---------------------------------------------------------

On Error GoTo Error_Handler

adoKHALDZAL_AddNew = Null
rsADO.AddNew
adoKHALDZAL_AddNew = rsKHALDZAL_PutBuffer(rsADO, rsKHALDZAL)
rsADO.Update

Exit Function

Error_Handler:

adoKHALDZAL_AddNew = Error


End Function

'---------------------------------------------------------
Public Function rsKHALDZAL_PutBuffer(rsADO As ADODB.Recordset, rsKHALDZAL As typeKHALDZAL)
'---------------------------------------------------------
On Error GoTo Error_Handler



rsADO("Id") = rsKHALDZAL.Id
rsADO("HISMVTDEV") = rsKHALDZAL.HISMVTDEV
rsADO("HISMVTCPT") = rsKHALDZAL.HISMVTCPT
rsADO("HISMVTOPEC") = rsKHALDZAL.HISMVTOPEC
rsADO("HISMVTMTD") = rsKHALDZAL.HISMVTMTD
rsADO("HISMVTOPEN") = rsKHALDZAL.HISMVTOPEN
rsADO("HISMVTLIB1") = rsKHALDZAL.HISMVTLIB1
rsADO("HISMVTLIB2") = rsKHALDZAL.HISMVTLIB2
rsADO("HISMVTDTRT") = rsKHALDZAL.HISMVTDTRT
rsADO("HISMVTDVAL") = rsKHALDZAL.HISMVTDVAL
rsADO("HISMVTPIEN") = rsKHALDZAL.HISMVTPIEN
rsADO("HISMVTPIES") = rsKHALDZAL.HISMVTPIES
rsADO("HISMVTXBEN") = rsKHALDZAL.HISMVTXBEN
rsADO("HISMVTXREF") = rsKHALDZAL.HISMVTXREF
rsADO("SAAMSGID") = rsKHALDZAL.SAAMSGID
rsADO("SAAMSGXDO") = rsKHALDZAL.SAAMSGXDO
rsADO("SAAMSGXBEN") = rsKHALDZAL.SAAMSGXBEN
rsADO("SAAMSGXPAY") = rsKHALDZAL.SAAMSGXPAY
rsADO("SAAMSGTXT") = rsKHALDZAL.SAAMSGTXT

rsADO("SAAMATCH") = rsKHALDZAL.SAAMATCH
rsADO("SAAMSGTYPE") = rsKHALDZAL.SAAMSGTYPE
rsADO("SAAMSGDTRT") = rsKHALDZAL.SAAMSGDTRT
rsADO("SCANLINK") = rsKHALDZAL.SCANLINK
rsADO("SCANLINK2") = rsKHALDZAL.SCANLINK2


rsKHALDZAL_PutBuffer = Null
Exit Function

Error_Handler:

rsKHALDZAL_PutBuffer = Error
End Function

'---------------------------------------------------------
Public Function rsKHALDZAL_GetBuffer(rsADO As ADODB.Recordset, rsKHALDZAL As typeKHALDZAL)
'---------------------------------------------------------
On Error GoTo Error_Handler



rsKHALDZAL.Id = rsADO("Id")
rsKHALDZAL.HISMVTDEV = rsADO("HISMVTDEV")
rsKHALDZAL.HISMVTCPT = rsADO("HISMVTCPT")
rsKHALDZAL.HISMVTOPEC = rsADO("HISMVTOPEC")
rsKHALDZAL.HISMVTMTD = rsADO("HISMVTMTD")
rsKHALDZAL.HISMVTOPEN = rsADO("HISMVTOPEN")
rsKHALDZAL.HISMVTLIB1 = rsADO("HISMVTLIB1")
rsKHALDZAL.HISMVTLIB2 = rsADO("HISMVTLIB2")
rsKHALDZAL.HISMVTDTRT = rsADO("HISMVTDTRT")
rsKHALDZAL.HISMVTDVAL = rsADO("HISMVTDVAL")
rsKHALDZAL.HISMVTPIEN = rsADO("HISMVTPIEN")
rsKHALDZAL.HISMVTPIES = rsADO("HISMVTPIES")
rsKHALDZAL.HISMVTXBEN = rsADO("HISMVTXBEN")
rsKHALDZAL.HISMVTXREF = rsADO("HISMVTXREF")
rsKHALDZAL.SAAMSGID = rsADO("SAAMSGID")
rsKHALDZAL.SAAMSGXDO = rsADO("SAAMSGXDO")
rsKHALDZAL.SAAMSGXBEN = rsADO("SAAMSGXBEN")
rsKHALDZAL.SAAMSGXPAY = rsADO("SAAMSGXPAY")
rsKHALDZAL.SAAMSGTXT = rsADO("SAAMSGTXT")

rsKHALDZAL.SAAMATCH = rsADO("SAAMATCH")
rsKHALDZAL.SAAMSGTYPE = rsADO("SAAMSGTYPE")
rsKHALDZAL.SAAMSGDTRT = rsADO("SAAMSGDTRT")
rsKHALDZAL.SCANLINK = rsADO("SCANLINK")
rsKHALDZAL.SCANLINK2 = rsADO("SCANLINK2")


rsKHALDZAL_GetBuffer = Null
Exit Function

Error_Handler:

rsKHALDZAL_GetBuffer = Error
End Function


'---------------------------------------------------------
Public Function adoYSAAMVTLNK_AddNew(rsADO As ADODB.Recordset, rsYSAAMVTLNK As typeYSAAMVTLNK)
'---------------------------------------------------------

On Error GoTo Error_Handler

adoYSAAMVTLNK_AddNew = Null
rsADO.AddNew
adoYSAAMVTLNK_AddNew = rsYSAAMVTLNK_PutBuffer(rsADO, rsYSAAMVTLNK)
rsADO.Update

Exit Function

Error_Handler:

adoYSAAMVTLNK_AddNew = Error


End Function

'---------------------------------------------------------
Public Function rsYSAAMVTLNK_PutBuffer(rsADO As ADODB.Recordset, rsYSAAMVTLNK As typeYSAAMVTLNK)
'---------------------------------------------------------
On Error GoTo Error_Handler



rsADO("HISMVTID") = rsYSAAMVTLNK.HISMVTID

rsADO("SAAMSGID") = rsYSAAMVTLNK.SAAMSGID



rsYSAAMVTLNK_PutBuffer = Null
Exit Function

Error_Handler:

rsYSAAMVTLNK_PutBuffer = Error
End Function

'---------------------------------------------------------
Public Function rsYSAAMVTLNK_GetBuffer(rsADO As ADODB.Recordset, rsYSAAMVTLNK As typeYSAAMVTLNK)
'---------------------------------------------------------
On Error GoTo Error_Handler



rsYSAAMVTLNK.HISMVTID = rsADO("HISMVTID")

rsYSAAMVTLNK.SAAMSGID = rsADO("SAAMSGID")

rsYSAAMVTLNK_GetBuffer = Null
Exit Function

Error_Handler:

rsYSAAMVTLNK_GetBuffer = Error
End Function



Public Sub YSAAMSG_Init()
Dim xSql As String
cnSAB073Y_Open
Call lstErr_AddItem(frmKHALDZAL.lstErr, frmYGOSDOS0.cmdContext, "YSAAMSG_Init")

xSql = "delete * from YSAAMSG0"
Call FEU_ROUGE
Set rsYSAAMSG0 = cnSAB073Y.Execute(xSql)
xSql = "delete * from YSAAMSG1"
Set rsYSAAMSG1 = cnSAB073Y.Execute(xSql)
Call FEU_VERT
Call lstErr_AddItem(frmKHALDZAL.lstErr, frmYGOSDOS0.cmdContext, "K01.txt")
Call YSAAMSG_Import("c:\temp\KHALDZAL\SAAMSG\K01.txt")

Call lstErr_AddItem(frmKHALDZAL.lstErr, frmYGOSDOS0.cmdContext, "K02.txt")
Call YSAAMSG_Import("c:\temp\KHALDZAL\SAAMSG\K02.txt")

Call lstErr_AddItem(frmKHALDZAL.lstErr, frmYGOSDOS0.cmdContext, "K03.txt")
Call YSAAMSG_Import("c:\temp\KHALDZAL\SAAMSG\K03.txt")

Call lstErr_AddItem(frmKHALDZAL.lstErr, frmYGOSDOS0.cmdContext, "K04.txt")
Call YSAAMSG_Import("c:\temp\KHALDZAL\SAAMSG\K04.txt")

Call lstErr_AddItem(frmKHALDZAL.lstErr, frmYGOSDOS0.cmdContext, "K05.txt")
Call YSAAMSG_Import("c:\temp\KHALDZAL\SAAMSG\K05.txt")

Call lstErr_AddItem(frmKHALDZAL.lstErr, frmYGOSDOS0.cmdContext, "K06.txt")
Call YSAAMSG_Import("c:\temp\KHALDZAL\SAAMSG\K06.txt")

Call lstErr_AddItem(frmKHALDZAL.lstErr, frmYGOSDOS0.cmdContext, "K07.txt")
Call YSAAMSG_Import("c:\temp\KHALDZAL\SAAMSG\K07.txt")

Call lstErr_AddItem(frmKHALDZAL.lstErr, frmYGOSDOS0.cmdContext, "K08.txt")
Call YSAAMSG_Import("c:\temp\KHALDZAL\SAAMSG\K08.txt")

Call lstErr_AddItem(frmKHALDZAL.lstErr, frmYGOSDOS0.cmdContext, "K09.txt")
Call YSAAMSG_Import("c:\temp\KHALDZAL\SAAMSG\K09.txt")

Call lstErr_AddItem(frmKHALDZAL.lstErr, frmYGOSDOS0.cmdContext, "K10.txt")
Call YSAAMSG_Import("c:\temp\KHALDZAL\SAAMSG\K10.txt")

Call lstErr_AddItem(frmKHALDZAL.lstErr, frmYGOSDOS0.cmdContext, "K11.txt")
Call YSAAMSG_Import("c:\temp\KHALDZAL\SAAMSG\K11.txt")

Call lstErr_AddItem(frmKHALDZAL.lstErr, frmYGOSDOS0.cmdContext, "K12.txt")
Call YSAAMSG_Import("c:\temp\KHALDZAL\SAAMSG\K12.txt")

Call lstErr_AddItem(frmKHALDZAL.lstErr, frmYGOSDOS0.cmdContext, "K13.txt")
Call YSAAMSG_Import("c:\temp\KHALDZAL\SAAMSG\K13.txt")

Call lstErr_AddItem(frmKHALDZAL.lstErr, frmYGOSDOS0.cmdContext, "K14.txt")
Call YSAAMSG_Import("c:\temp\KHALDZAL\SAAMSG\K14.txt")

Call lstErr_AddItem(frmKHALDZAL.lstErr, frmYGOSDOS0.cmdContext, "K15.txt")
Call YSAAMSG_Import("c:\temp\KHALDZAL\SAAMSG\K15.txt")

Call lstErr_AddItem(frmKHALDZAL.lstErr, frmYGOSDOS0.cmdContext, "K16.txt")
Call YSAAMSG_Import("c:\temp\KHALDZAL\SAAMSG\K16.txt")

Call lstErr_AddItem(frmKHALDZAL.lstErr, frmYGOSDOS0.cmdContext, "K17.txt")
Call YSAAMSG_Import("c:\temp\KHALDZAL\SAAMSG\K17.txt")

End Sub

Public Function HISMVTP0_YSAAMSG1_Pays(lTxt) As String
Dim K As Integer
If Len(lTxt) <= 11 Then
    HISMVTP0_YSAAMSG1_Pays = Mid$(lTxt, 5, 2)
Else
    K = InStr(lTxt, "_")
    If K > 0 Then HISMVTP0_YSAAMSG1_Pays = Mid$(lTxt, K + 5, 2)
End If

End Function

Public Sub ScanLink_Init()
Dim objFolder, objFiles_Open, objFiles_Close
Dim fsoFile As File
Dim X As String, K As Integer, K1 As Integer
Dim wHISMVTOPEN As Long, xSql As String, wId As Long
Dim mFolder_Old As String, mFolder_New As String
Dim xFile As String, xFile2 As String, xFile_New As String

On Error GoTo Error_Handle

cnSAB073Y_Open

mFolder_Old = "C:\Temp\KHALDZAL\PDF_X"
mFolder_New = "C:\Temp\KHALDZAL\PDF\"
Set objFolder = msFileSystem.GetFolder(mFolder_Old)
Set objFiles_Close = objFolder.Files
For Each fsoFile In objFiles_Close
    xFile = fsoFile.Name
    xFile2 = Replace(xFile, "..", ".")
    K = InStr(xFile, "Optrf")
    If K > 0 Then
        K = K + 5
        K1 = InStr(K, xFile2, ".PDF")
        If K1 > 0 Then
            If K1 - K < 7 Then
'_______________________________________________________________________________________
                wHISMVTOPEN = Val(Mid$(xFile2, K, K1 - K))
                xSql = "select Id,HISMVTOPEN,SCANLINK from KHALDZAL  where HISMVTOPEN = " & wHISMVTOPEN
                    Set rsKHALDZAL = cnSAB073Y.Execute(xSql)
                If rsKHALDZAL.EOF Then
                     Debug.Print "SCANLINK inconnu"; wHISMVTOPEN
                Else
                    wId = rsKHALDZAL("Id")
                    If rsKHALDZAL("SCANLINK") <> "" Then
                         Debug.Print "SCANLINK doublon"; wHISMVTOPEN; wId
                     Else
                        xFile_New = mFolder_New & wHISMVTOPEN & ".PDF"
                        xSql = "update KHALDZAL  set SCANLINK = '" & wHISMVTOPEN & "#" & xFile_New & "'" _
                         & " where id = " & wId
                        Call FEU_ROUGE
                        Set rsYSAAMSG0 = cnSAB073Y.Execute(xSql)
                        DoEvents
                        msFileSystem.MoveFile mFolder_Old & "\" & xFile, xFile_New
                        Call FEU_VERT
                        DoEvents: DoEvents: DoEvents
                     End If
                End If
            Else
'_______________________________________________________________________________________
                wHISMVTOPEN = Val(Mid$(xFile2, K, 6))
                xSql = "select Id,HISMVTOPEN,SCANLINK2 from KHALDZAL  where HISMVTOPEN = " & wHISMVTOPEN
                    Set rsKHALDZAL = cnSAB073Y.Execute(xSql)
                If rsKHALDZAL.EOF Then
                     Debug.Print "SCANLINK2 inconnu"; wHISMVTOPEN
                Else
                    wId = rsKHALDZAL("Id")
                    If rsKHALDZAL("SCANLINK2") <> "" Then
                         Debug.Print "SCANLINK2 doublon"; wHISMVTOPEN; wId
                     Else
                        xFile_New = mFolder_New & wHISMVTOPEN & "_X.PDF"
                        xSql = "update KHALDZAL  set SCANLINK2 = '" & wHISMVTOPEN & "#" & xFile_New & "'" _
                         & " where id = " & wId
                        Call FEU_ROUGE
                        Set rsYSAAMSG0 = cnSAB073Y.Execute(xSql)
                        DoEvents
                        msFileSystem.MoveFile mFolder_Old & "\" & xFile, xFile_New
                        Call FEU_VERT
                        DoEvents: DoEvents: DoEvents
                     End If
                End If
            
            End If
        End If
    End If
    

Next
Close
Call lstErr_AddItem(frmKHALDZAL.lstErr, frmYGOSDOS0.cmdContext, "ScanLink")
cnSAB073Y_Close
Exit Sub

Error_Handle:
 MsgBox Error, vbCritical, "erreur : ScanLink"
Close
cnSAB073Y_Close

End Sub
Public Sub ScanLink_Reprise()
Dim objFolder, objFiles_Open, objFiles_Close
Dim fsoFile As File
Dim X As String, K As Integer, K1 As Integer
Dim wHISMVTOPEN As Long, xSql As String, wId As Long
Dim mFolder_New As String
Dim xFile As String, xFile_New As String

On Error GoTo Error_Handle

cnSAB073Y_Open

        X = "update KHALDZAL  set SCANLINK = ''" _
             & " where SCANLINK <> ''"
        Set rsYSAAMSG0 = cnSAB073Y.Execute(X) '
        
        X = "update KHALDZAL  set SCANLINK2 = ''" _
             & " where SCANLINK2 <> ''"
        Set rsYSAAMSG0 = cnSAB073Y.Execute(X) '


mFolder_New = "C:\Temp\KHALDZAL\PDF\"
Set objFolder = msFileSystem.GetFolder(mFolder_New)
Set objFiles_Close = objFolder.Files
Call FEU_ROUGE
For Each fsoFile In objFiles_Close
    xFile = fsoFile.Name
    K = 1
    K1 = InStr(K, xFile, ".PDF")
    If K1 > 0 Then
        If K1 - K < 7 Then
            wHISMVTOPEN = Val(Mid$(xFile, K, K1 - K))
            xSql = "select Id,HISMVTOPEN,SCANLINK from KHALDZAL  where HISMVTOPEN = " & wHISMVTOPEN
                Set rsKHALDZAL = cnSAB073Y.Execute(xSql)
            If rsKHALDZAL.EOF Then
                 Debug.Print "SCANLINK inconnu"; wHISMVTOPEN
            Else
                wId = rsKHALDZAL("Id")
                If rsKHALDZAL("SCANLINK") <> "" Then
                     Debug.Print "SCANLINK doublon"; wHISMVTOPEN; wId
                 Else
                    xFile_New = mFolder_New & wHISMVTOPEN & ".PDF"
                    xSql = "update KHALDZAL  set SCANLINK = '" & wHISMVTOPEN & "#" & xFile_New & "'" _
                     & " where id = " & wId
                    Set rsYSAAMSG0 = cnSAB073Y.Execute(xSql)
                        DoEvents
                 End If
            End If
'_______________________________________________________________________________________
            Else
                wHISMVTOPEN = Val(Mid$(xFile, K, 6))
                xSql = "select Id,HISMVTOPEN,SCANLINK2 from KHALDZAL  where HISMVTOPEN = " & wHISMVTOPEN
                    Set rsKHALDZAL = cnSAB073Y.Execute(xSql)
                If rsKHALDZAL.EOF Then
                     Debug.Print "SCANLINK2 inconnu"; wHISMVTOPEN
                Else
                    wId = rsKHALDZAL("Id")
                    If rsKHALDZAL("SCANLINK2") <> "" Then
                         Debug.Print "SCANLINK2 doublon"; wHISMVTOPEN; wId
                     Else
                        xFile_New = mFolder_New & wHISMVTOPEN & "_X.PDF"
                        xSql = "update KHALDZAL  set SCANLINK2 = '" & wHISMVTOPEN & "#" & xFile_New & "'" _
                         & " where id = " & wId
                        Set rsYSAAMSG0 = cnSAB073Y.Execute(xSql)
                        DoEvents
                     End If
                End If

        End If
    End If
Next
Close
Call FEU_VERT
Call lstErr_AddItem(frmKHALDZAL.lstErr, frmYGOSDOS0.cmdContext, "ScanLink")
cnSAB073Y_Close
Exit Sub

Error_Handle:
 MsgBox Error, vbCritical, "erreur : ScanLink"
Close
cnSAB073Y_Close

End Sub



Public Sub YSAAMSG_201_Origine()
Dim V, xSql As String
Dim xIn As String, K As Integer, K2 As Integer, lenX As Integer, Nb_ok As Long, Nb_Lu As Long
Dim kIn As Integer, Seq As Integer
On Error GoTo Error_Handle
Dim X As String
Dim mSeq As Long
Dim K1 As Integer, I1 As Integer, I As Integer
Dim blnOk As Boolean, blnPrint As Boolean, blnSwift As Boolean, kPrint As Integer
Dim xKHALDZAL As typeKHALDZAL
Dim xYSAAMSG1 As typeYSAAMSG1
Dim blnYSAAMSG0 As Boolean, blnYSAAMSG1 As Boolean
Dim wSAAMsgId As Long, wSAAMsgId0 As Long

cnSAB073Y_Open

xSql = "select * from YSAAMSG0 where SAAMsgId0 > 0 " _
     & " order by SAAMsgId0"
Set rsYSAAMSG0 = cnSAB073Y.Execute(xSql)

Do While Not rsYSAAMSG0.EOF
    wSAAMsgId = rsYSAAMSG0("SAAMsgId")
    wSAAMsgId0 = rsYSAAMSG0("SAAMsgId0")
    blnOk = False
'___________________________________________________________________________

 xSql = "select * from YSAAMSG1  where SAAMsgId = " & wSAAMsgId0 _
         & " order by SAAMsgSeq"
    Set rsYSAAMSG1 = cnSAB073Y.Execute(xSql)
    
    Do While Not rsYSAAMSG1.EOF
        V = rsYSAAMSG1_GetBuffer(rsYSAAMSG1, xYSAAMSG1)
        blnOk = True
        X = xYSAAMSG1.SAAMsgFld & xYSAAMSG1.SAAMsgFldX & ": " & xYSAAMSG1.SAAMSGTXT
        If xKHALDZAL.SAAMSGTXT = "" Then
            xKHALDZAL.SAAMSGTXT = X
        Else
            xKHALDZAL.SAAMSGTXT = xKHALDZAL.SAAMSGTXT & vbCrLf & X
        End If
        rsYSAAMSG1.MoveNext
    Loop
'___________________________________________________________________________
 If blnOk Then
        xKHALDZAL.SAAMSGTXT = Replace(xKHALDZAL.SAAMSGTXT, "'", "''")

        xSql = "update KHALDZAL  set SAAMSGID = " & wSAAMsgId0 _
             & " , SAAMSGTXT = '" & xKHALDZAL.SAAMSGTXT & "'" _
             & " where SAAMSGID = " & wSAAMsgId
    Call FEU_ROUGE
    Set rsKHALDZAL = cnSAB073Y.Execute(xSql)
    Call FEU_VERT
End If

    rsYSAAMSG0.MoveNext

Loop

'___________________________________________________________________________
Close
Call lstErr_AddItem(frmKHALDZAL.lstErr, frmYGOSDOS0.cmdContext, "Terminé")
cnSAB073Y_Close
Exit Sub

Error_Handle:
 MsgBox Error, vbCritical, "erreur : HISMVTP0_match"
Close
cnSAB073Y_Close






End Sub
