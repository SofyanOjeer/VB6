Attribute VB_Name = "srvYSAAMSG"
Option Explicit

Dim cnSAB073Y As New ADODB.Connection
Dim rsYSAAMSG0 As New ADODB.Recordset
Dim rsYSAAMSG1 As New ADODB.Recordset
Dim arrYSAAMSG1(1000) As typeYSAAMSG1, arrYSAAMSG1_Nb As Integer

Type typeYSAAMSG0
    SAAMSGID           As Long
    SAAMsgBICS         As String * 11
    SAAMsgBICR         As String * 11
    SAAMSGTYPE         As String * 3
    SAAMsgTRN          As String * 16
    SAAMsgTRNR         As String * 16
    SAAMsgMt           As Currency
    SAAMsgDev          As String * 3
    SAAMsgDVal         As String * 8
    SAAMSGDTRT         As String * 8
    SAAMSGID0          As Long

End Type

Type typeYSAAMSG1
    SAAMSGID           As Long
    SAAMsgSeq          As Long
    SAAMsgFld          As String * 2
    SAAMsgFldX         As String * 1
    SAAMSGTXT          As Variant

End Type

'---------------------------------------------------------
Public Function rsYSAAMSG0_PutBuffer(rsADO As ADODB.Recordset, rsYSAAMSG0 As typeYSAAMSG0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsYSAAMSG0_PutBuffer = Null

rsADO("SAAMsgId") = rsYSAAMSG0.SAAMSGID
rsADO("SAAMsgBICS") = rsYSAAMSG0.SAAMsgBICS
rsADO("SAAMsgBICR") = rsYSAAMSG0.SAAMsgBICR
rsADO("SAAMsgType") = rsYSAAMSG0.SAAMSGTYPE
rsADO("SAAMsgTRN") = rsYSAAMSG0.SAAMsgTRN
rsADO("SAAMsgTRNR") = rsYSAAMSG0.SAAMsgTRNR
rsADO("SAAMsgMt") = rsYSAAMSG0.SAAMsgMt
rsADO("SAAMsgDev") = rsYSAAMSG0.SAAMsgDev

rsADO("SAAMsgDVal") = rsYSAAMSG0.SAAMsgDVal
rsADO("SAAMsgDTrt") = rsYSAAMSG0.SAAMSGDTRT
rsADO("SAAMsgId0") = rsYSAAMSG0.SAAMSGID0
Exit Function

Error_Handler:

rsYSAAMSG0_PutBuffer = Error
End Function

'---------------------------------------------------------
Public Function rsYSAAMSG0_GetBuffer(rsADO As ADODB.Recordset, rsYSAAMSG0 As typeYSAAMSG0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsYSAAMSG0_GetBuffer = Null

rsYSAAMSG0.SAAMSGID = rsADO("SAAMsgId")
rsYSAAMSG0.SAAMsgBICS = rsADO("SAAMsgBICS")
rsYSAAMSG0.SAAMsgBICR = rsADO("SAAMsgBICR")
rsYSAAMSG0.SAAMSGTYPE = rsADO("SAAMsgType")
rsYSAAMSG0.SAAMsgTRN = rsADO("SAAMsgTRN")
rsYSAAMSG0.SAAMsgTRNR = rsADO("SAAMsgTRNR")
rsYSAAMSG0.SAAMsgMt = rsADO("SAAMsgMt")
rsYSAAMSG0.SAAMsgDev = rsADO("SAAMsgDev")

rsYSAAMSG0.SAAMsgDVal = rsADO("SAAMsgDVal")
rsYSAAMSG0.SAAMSGDTRT = rsADO("SAAMsgDTrt")
rsYSAAMSG0.SAAMSGID0 = rsADO("SAAMsgId0")
Exit Function

Error_Handler:

rsYSAAMSG0_GetBuffer = Error
End Function


'---------------------------------------------------------
Public Function rsYSAAMSG1_PutBuffer(rsADO As ADODB.Recordset, rsYSAAMSG1 As typeYSAAMSG1)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsYSAAMSG1_PutBuffer = Null

rsADO("SAAMsgId") = rsYSAAMSG1.SAAMSGID
rsADO("SAAMsgSeq") = rsYSAAMSG1.SAAMsgSeq
rsADO("SAAMsgFld") = rsYSAAMSG1.SAAMsgFld
rsADO("SAAMsgFldX") = rsYSAAMSG1.SAAMsgFldX
rsADO("SAAMsgTxt") = rsYSAAMSG1.SAAMSGTXT
Exit Function

Error_Handler:

rsYSAAMSG1_PutBuffer = Error
End Function


'---------------------------------------------------------
Public Function rsYSAAMSG1_GetBuffer(rsADO As ADODB.Recordset, rsYSAAMSG1 As typeYSAAMSG1)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsYSAAMSG1_GetBuffer = Null

rsYSAAMSG1.SAAMSGID = rsADO("SAAMsgId")
rsYSAAMSG1.SAAMsgSeq = rsADO("SAAMsgSeq")
rsYSAAMSG1.SAAMsgFld = rsADO("SAAMsgFld")
rsYSAAMSG1.SAAMsgFldX = rsADO("SAAMsgFldX")
rsYSAAMSG1.SAAMSGTXT = rsADO("SAAMsgTxt")
Exit Function

Error_Handler:

rsYSAAMSG1_GetBuffer = Error
End Function


'---------------------------------------------------------
Public Function adoYSAAMSG1_AddNew(rsADO As ADODB.Recordset, rsYSAAMSG1 As typeYSAAMSG1)
'---------------------------------------------------------

On Error GoTo Error_Handler

adoYSAAMSG1_AddNew = Null
rsADO.AddNew
adoYSAAMSG1_AddNew = rsYSAAMSG1_PutBuffer(rsADO, rsYSAAMSG1)
rsADO.Update

Exit Function

Error_Handler:

adoYSAAMSG1_AddNew = Error


End Function


'---------------------------------------------------------
Public Function adoYSAAMSG0_AddNew(rsADO As ADODB.Recordset, rsYSAAMSG0 As typeYSAAMSG0)
'---------------------------------------------------------

On Error GoTo Error_Handler

adoYSAAMSG0_AddNew = Null
rsADO.AddNew
adoYSAAMSG0_AddNew = rsYSAAMSG0_PutBuffer(rsADO, rsYSAAMSG0)
rsADO.Update

Exit Function

Error_Handler:

adoYSAAMSG0_AddNew = Error


End Function

Public Function adoYSAAMSG0_Delete(rsADO As ADODB.Recordset, rsYSAAMSG0 As typeYSAAMSG0)
'---------------------------------------------------------
Dim xSql As String

On Error GoTo Error_Handler
adoYSAAMSG0_Delete = Null

xSql = "delete * from YSAAMSG0 where SAAMsgId = " & rsYSAAMSG0.SAAMSGID
Call FEU_ROUGE
Set rsADO = cnSAB073Y.Execute(xSql)
Call FEU_VERT
Exit Function

Error_Handler:

adoYSAAMSG0_Delete = Error

End Function


Public Sub rsYSAAMSG0_Init(rsYSAAMSG0 As typeYSAAMSG0)
rsYSAAMSG0.SAAMSGID = 0
rsYSAAMSG0.SAAMsgBICS = ""
rsYSAAMSG0.SAAMsgBICR = ""
rsYSAAMSG0.SAAMSGTYPE = ""
rsYSAAMSG0.SAAMsgTRN = ""
rsYSAAMSG0.SAAMsgTRNR = ""
rsYSAAMSG0.SAAMsgMt = 0
rsYSAAMSG0.SAAMsgDev = ""
rsYSAAMSG0.SAAMsgDVal = "00000000"
rsYSAAMSG0.SAAMSGDTRT = "00000000"
rsYSAAMSG0.SAAMSGID0 = 0
End Sub

Public Sub rsYSAAMSG1_Init(rsYSAAMSG1 As typeYSAAMSG1)
rsYSAAMSG1.SAAMSGID = 0
rsYSAAMSG1.SAAMsgFld = ""
rsYSAAMSG1.SAAMsgFldX = ""
rsYSAAMSG1.SAAMSGTXT = ""
End Sub

Public Sub cnSAB073Y_Open()
On Error GoTo Error_Handler
Dim X As String

cnSAB073Y.Open paramODBC_DSN_SAB073Y

Exit Sub

Error_Handler:

End Sub

Public Sub YSAAMSG_Import(lFile)
Dim V
Dim xIn As String, K As Integer, lenX As Integer
Dim kIn As Integer, Seq As Integer
On Error GoTo Error_Handle
Dim X As String
Dim mSeq As Integer
Dim K1 As Integer, I1 As Integer, I As Integer
Dim blnOk As Boolean, blnPrint As Boolean, blnSwift As Boolean, kPrint As Integer

Dim wYSAAMSG0 As typeYSAAMSG0

cnSAB073Y_Open


rsYSAAMSG0.Open "select * from YSAAMSG0", cnSAB073Y, , adLockOptimistic
rsYSAAMSG1.Open "select * from YSAAMSG1", cnSAB073Y, , adLockOptimistic

Open lFile For Input As #1
Seq = 0:
Call lstErr_AddItem(frmSAA.lstErr, frmSAA.cmdContext, "Import : " & lFile)
blnOk = False: blnPrint = False
Call lstErr_AddItem(frmSAA.lstErr, frmSAA.cmdContext, "Import : " & lFile)
Do Until EOF(1)
    DoEvents
    Line Input #1, xIn
    If Trim(xIn) <> "" Then
        If Mid$(xIn, 1, 13) = "U-UMID      =" Then
            Seq = Seq + 1
            rsYSAAMSG0_Init wYSAAMSG0
            blnOk = True
        End If
        If blnOk Then
                 'If Mid$(xIn, 1, 19) = "Identifier   = fin." Then wYSAAMSG0.SAAMsgType = Mid$(xIn, 20, 3)
           Select Case Mid$(xIn, 1, 14)
                Case "Identifier   ="
                            wYSAAMSG0.SAAMSGTYPE = Mid$(xIn, 20, 3)
                Case "Sender       ="
                            Line Input #1, xIn
                            wYSAAMSG0.SAAMsgBICS = Mid$(xIn, 6, 8)
                Case "Receiver     =":
                            Line Input #1, xIn
                            wYSAAMSG0.SAAMsgBICR = Mid$(xIn, 6, 8)
                Case "Transaction re"
                            wYSAAMSG0.SAAMsgTRN = Mid$(xIn, 20, 16)
                            K = InStr(40, xIn, "=")
                            wYSAAMSG0.SAAMsgTRNR = Mid$(xIn, K + 2, 16)
                Case "Amount      = ":
                            Call YSAAMSG_Import_Amount(xIn, wYSAAMSG0)
                Case "Date/Time   = ":
                            wYSAAMSG0.SAAMSGDTRT = "20" & Mid$(xIn, 21, 2) & Mid$(xIn, 18, 2) & Mid$(xIn, 15, 2)
                Case "Text          ":
                        YSAAMSG_Import_text
                    
                Case "Block 5:": blnPrint = False
                Case "Message Histor":
                    Line Input #1, xIn
                    Line Input #1, xIn
                    Line Input #1, xIn
                    Line Input #1, xIn
                    K = InStr(15, xIn, "Sequence Nr") + 12
                    If K = 12 Then Line Input #1, xIn: K = InStr(15, xIn, "Sequence Nr") + 12
                    If K = 12 Then Line Input #1, xIn: K = InStr(15, xIn, "Sequence Nr") + 12
                    lenX = Len(xIn)
                    wYSAAMSG0.SAAMSGID = Mid$(xIn, K, lenX - K)
                    blnOk = False
                    Call lstErr_ChangeLastItem(frmSAA.lstErr, frmSAA.cmdContext, "Sequence Nr : " & wYSAAMSG0.SAAMSGID)
                    V = adoYSAAMSG0_AddNew(rsYSAAMSG0, wYSAAMSG0)
                    If Not IsNull(V) Then MsgBox "erreur : YSAAMSG_Import_AddNew " & Seq, vbCritical, Error
                    For K = 1 To arrYSAAMSG1_Nb
                        arrYSAAMSG1(K).SAAMSGID = wYSAAMSG0.SAAMSGID
                        arrYSAAMSG1(K).SAAMsgSeq = K
                        V = adoYSAAMSG1_AddNew(rsYSAAMSG1, arrYSAAMSG1(K))
                        If Not IsNull(V) Then MsgBox "erreur : YSAAMSG_Import_AddNew YSAAMSG1" & Seq & " / " & K, vbCritical, Error
                    Next K
            End Select
        End If
   End If
Loop

Close
Call lstErr_AddItem(frmSAA.lstErr, frmSAA.cmdContext, "Nb : " & Seq)
cnSAB073Y_Close
Exit Sub

Error_Handle:
 MsgBox "erreur : YSAAMSG_Import" & xIn, vbCritical, Error
Close
cnSAB073Y_Open

End Sub


Public Sub cnSAB073Y_Close()
On Error Resume Next

cnSAB073Y.Close
Set cnSAB073Y = Nothing


End Sub


Public Sub YSAAMSG_Import_Amount(lIn As String, lYSAAMSG0 As typeYSAAMSG0)
Dim K As Integer
Dim X As String, X1 As String
Dim blnOk As Boolean, blnPoint As Boolean

If Trim(Mid$(lIn, 14, 27)) = "" Then Exit Sub

blnPoint = False
X = ""
K = 14

Do
    X1 = Mid$(lIn, K, 1)
    Select Case X1
        Case " ":   If blnPoint Then Exit Do
        Case ","
        Case ".":   blnPoint = True: X = X & ","
        Case Else:
                    X = X & X1
    End Select
    K = K + 1
Loop

lYSAAMSG0.SAAMsgMt = CCur(X)
lYSAAMSG0.SAAMsgDev = Mid$(lIn, K + 1, 3)

K = InStr(K + 4, lIn, "=")
X = Mid$(lIn, K + 2, 6)

If IsNumeric(X) Then lYSAAMSG0.SAAMsgDVal = "20" & X
        
       

End Sub

Public Sub YSAAMSG_Import_text()
Dim xIn As String
Dim K As Integer, lenX As Integer

arrYSAAMSG1_Nb = 0

Do Until EOF(1)
    DoEvents
    Line Input #1, xIn
    xIn = Trim(xIn)
    If xIn <> "" Then
        If Mid$(xIn, 1, 8) = "Block 5:" Then Exit Sub
        lenX = Len(xIn)
        If Mid$(xIn, 1, 1) = ":" Then
            arrYSAAMSG1_Nb = arrYSAAMSG1_Nb + 1
            arrYSAAMSG1(arrYSAAMSG1_Nb).SAAMsgFld = Mid$(xIn, 2, 2)
            If Mid$(xIn, 4, 1) = ":" Then
                arrYSAAMSG1(arrYSAAMSG1_Nb).SAAMsgFldX = ""
                arrYSAAMSG1(arrYSAAMSG1_Nb).SAAMSGTXT = Mid$(xIn, 5, lenX - 4)
            Else
                arrYSAAMSG1(arrYSAAMSG1_Nb).SAAMsgFldX = Mid$(xIn, 4, 1)
                arrYSAAMSG1(arrYSAAMSG1_Nb).SAAMSGTXT = Mid$(xIn, 6, lenX - 5)
            End If
        Else
            arrYSAAMSG1(arrYSAAMSG1_Nb).SAAMSGTXT = arrYSAAMSG1(arrYSAAMSG1_Nb).SAAMSGTXT & "_" & Mid$(xIn, 1, lenX)
        End If
   End If
Loop

End Sub
