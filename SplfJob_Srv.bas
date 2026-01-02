Attribute VB_Name = "srvSplfJob"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recSplfJobLen = 159 ' 34 + 125
Public Const recSplfJob_Block = 100

Type typeSplfJob
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    SJQAMJ                  As String * 8
    SJQID                   As Long
    SJQSEQ                  As Long
    SJQFILE                 As String * 10
    SJQUSR                  As String * 10
    SJQREF                  As String * 10
    SJQSTA                  As String * 3
    SJQPAGENB               As Long
    SJQEXNB                 As Long
    SJQHMS                  As String * 6
    SJQNAME                 As String * 10
    SJQOUTQ                 As String * 10
    SJQXAMJ                 As String * 8
    SJQXHMS                 As String * 6
    SJQXOUTQ                As String * 10
    SJQXSTA                 As String * 3
    SJQXEVTID               As Long
End Type
    
    
Public arrSplfJob() As typeSplfJob
Public arrSplfJob_NB As Integer
Public arrSplfJob_NBMax As Integer
Public arrSplfJob_Index As Integer
Public arrSplfJob_Suite As Boolean

Public xSplfJob As typeSplfJob

Public Const recSplfftpLen = 257 ' 34 + 223
Type typeSplffTP
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    SPFSEQ                  As Long
    SPFEVTID                As Long
    SPFSAUT                 As String * 4
    SPFTXT                  As String * 198
End Type

'-----------------------------------------------------
Function srvSplfJob_Update(recSplfJob As typeSplfJob)
'-----------------------------------------------------

srvSplfJob_Update = "?"

MsgTxtLen = 0
Call srvSplfJob_PutBuffer(recSplfJob)

If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvSplfJob_GetBuffer(recSplfJob)) Then
        Call srvSplfJob_Error(recSplfJob)
        srvSplfJob_Update = recSplfJob.Err
        Exit Function
    Else
        srvSplfJob_Update = Null
    End If
Else
    recSplfJob.Err = "srv"
End If


'=====================================================
End Function



'-----------------------------------------------------
Public Function srvSplfJob_Monitor(recSplfJob As typeSplfJob)
'-----------------------------------------------------

arrSplfJob_Suite = False
Select Case mId$(Trim(recSplfJob.Method), 1, 4)
    Case "Snap"
              srvSplfJob_Monitor = srvSplfJob_Snap(recSplfJob)
    Case Else
            srvSplfJob_Monitor = srvSplfJob_Seek(recSplfJob)
End Select

End Function

'-----------------------------------------------------
Sub srvSplfJob_Error(recSplfJob As typeSplfJob)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "SplfJob" & Chr$(10) & Chr$(13)

Select Case mId$(recSplfJob.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recSplfJob.Err
        I = vbCritical
End Select

MsgBox Msg & " : " & recSplfJob.SJQAMJ & " : " & recSplfJob.SJQID & " : " & recSplfJob.SJQSEQ _
        , I, "module : SplfJobs.bas  ( " & Trim(recSplfJob.obj) & " : " & Trim(recSplfJob.Method) & " )"

End Sub



'---------------------------------------------------------
Public Function srvSplfJob_GetBuffer(recSplfJob As typeSplfJob)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvSplfJob_GetBuffer = Null
recSplfJob.obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recSplfJob.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recSplfJob.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recSplfJob.Err = Space$(10) Then
    recSplfJob.SJQAMJ = Format$(Val(mId$(MsgTxt, K + 1, 8)), "00000000")
    recSplfJob.SJQID = CLng(Val(mId$(MsgTxt, K + 9, 6)))
    recSplfJob.SJQSEQ = CLng(Val(mId$(MsgTxt, K + 15, 5)))
    recSplfJob.SJQFILE = mId$(MsgTxt, K + 20, 10)
    recSplfJob.SJQUSR = mId$(MsgTxt, K + 30, 10)
    recSplfJob.SJQREF = mId$(MsgTxt, K + 40, 10)
    recSplfJob.SJQSTA = mId$(MsgTxt, K + 50, 3)
    recSplfJob.SJQPAGENB = CLng(Val(mId$(MsgTxt, K + 53, 5)))
    recSplfJob.SJQEXNB = CLng(Val(mId$(MsgTxt, K + 58, 3)))
    recSplfJob.SJQHMS = Format$(Val(mId$(MsgTxt, K + 61, 6)), "000000")
    recSplfJob.SJQNAME = mId$(MsgTxt, K + 67, 10)
    recSplfJob.SJQOUTQ = mId$(MsgTxt, K + 77, 10)
    recSplfJob.SJQXAMJ = Format$(Val(mId$(MsgTxt, K + 87, 8)), "00000000")
    recSplfJob.SJQXHMS = Format$(Val(mId$(MsgTxt, K + 95, 6)), "000000")
    recSplfJob.SJQXOUTQ = mId$(MsgTxt, K + 101, 10)
    recSplfJob.SJQXSTA = mId$(MsgTxt, K + 111, 3)
    recSplfJob.SJQXEVTID = CLng(Val(mId$(MsgTxt, K + 114, 12)))

Else
    srvSplfJob_GetBuffer = recSplfJob.Err
End If

MsgTxtIndex = MsgTxtIndex + recSplfJobLen

End Function

'---------------------------------------------------------
Public Function srvSplfFtp_GetBuffer(recSplfFtp As typeSplffTP)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvSplfFtp_GetBuffer = Null
recSplfFtp.obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recSplfFtp.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recSplfFtp.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recSplfFtp.Err = Space$(10) Then
    recSplfFtp.SPFSEQ = CLng(Val(mId$(MsgTxt, K + 9, 6)))
    recSplfFtp.SPFEVTID = CLng(Val(mId$(MsgTxt, K + 114, 12)))
    recSplfFtp.SPFSAUT = mId$(MsgTxt, K + 20, 10)
    recSplfFtp.SPFTXT = mId$(MsgTxt, K + 30, 10)

Else
    srvSplfFtp_GetBuffer = recSplfFtp.Err
End If

MsgTxtIndex = MsgTxtIndex + recSplfftpLen

End Function


'---------------------------------------------------------
Private Sub srvSplfJob_PutBuffer(recSplfJob As typeSplfJob)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recSplfJob.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recSplfJob.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34

    Mid$(MsgTxt, K + 1, 8) = Format$(recSplfJob.SJQAMJ, "00000000")
    Mid$(MsgTxt, K + 9, 6) = Format$(recSplfJob.SJQID, "000000")
    Mid$(MsgTxt, K + 15, 5) = Format$(recSplfJob.SJQSEQ, "00000")
    Mid$(MsgTxt, K + 20, 10) = recSplfJob.SJQFILE
    Mid$(MsgTxt, K + 30, 10) = recSplfJob.SJQUSR
    Mid$(MsgTxt, K + 40, 10) = recSplfJob.SJQREF
    Mid$(MsgTxt, K + 50, 3) = recSplfJob.SJQSTA
    Mid$(MsgTxt, K + 53, 5) = Format$(recSplfJob.SJQPAGENB, "00000")
    Mid$(MsgTxt, K + 58, 3) = Format$(recSplfJob.SJQEXNB, "000")
    Mid$(MsgTxt, K + 61, 6) = Format$(recSplfJob.SJQHMS, "000000")
    Mid$(MsgTxt, K + 67, 10) = recSplfJob.SJQNAME
    Mid$(MsgTxt, K + 77, 10) = recSplfJob.SJQOUTQ
    Mid$(MsgTxt, K + 87, 8) = Format$(recSplfJob.SJQXAMJ, "00000000")
    Mid$(MsgTxt, K + 101, 10) = recSplfJob.SJQXOUTQ
    Mid$(MsgTxt, K + 95, 6) = Format$(recSplfJob.SJQXHMS, "000000")
    Mid$(MsgTxt, K + 111, 3) = recSplfJob.SJQXSTA
    Mid$(MsgTxt, K + 114, 12) = Format$(recSplfJob.SJQXEVTID, "000000000000")
    

MsgTxtLen = MsgTxtLen + recSplfJobLen
End Sub



'---------------------------------------------------------
Private Function srvSplfJob_Seek(recSplfJob As typeSplfJob)
'---------------------------------------------------------

srvSplfJob_Seek = "?"
MsgTxtLen = 0
Call srvSplfJob_PutBuffer(recSplfJob)

'For I = 1 To 159
'    Debug.Print I; mId$(MsgTxt, I, 1); Asc(mId$(MsgTxt, I, 1))
'Next I

If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If IsNull(srvSplfJob_GetBuffer(recSplfJob)) Then
        srvSplfJob_Seek = Null
    Else
        Call srvSplfJob_Error(recSplfJob)
    End If
End If

End Function

'---------------------------------------------------------
Private Function srvSplfJob_Snap(recSplfJob As typeSplfJob)
'---------------------------------------------------------
srvSplfJob_Snap = "?"
MsgTxtLen = 0
Call srvSplfJob_PutBuffer(recSplfJob)
Call srvSplfJob_PutBuffer(arrSplfJob(0))
If IsNull(SndRcv()) Then
    srvSplfJob_Snap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(srvSplfJob_GetBuffer(recSplfJob)) Then
            Call arrSplfJob_AddItem(recSplfJob)
            arrSplfJob_Suite = True
        Else
            arrSplfJob_Suite = False
            Exit Do
        End If
    Loop
End If

End Function
'---------------------------------------------------------
Public Sub arrSplfJob_AddItem(recSplfJob As typeSplfJob)
'---------------------------------------------------------
          
arrSplfJob_NB = arrSplfJob_NB + 1
    
If arrSplfJob_NB > arrSplfJob_NBMax Then
    arrSplfJob_NBMax = arrSplfJob_NBMax + recSplfJob_Block
    ReDim Preserve arrSplfJob(arrSplfJob_NBMax)
End If
            
arrSplfJob(arrSplfJob_NB) = recSplfJob
End Sub



'---------------------------------------------------------
Public Sub recSplfJob_Init(recSplfJob As typeSplfJob)
'---------------------------------------------------------
MsgTxt = Space$(recSplfJobLen)
MsgTxtIndex = 0
Call srvSplfJob_GetBuffer(recSplfJob)
recSplfJob.obj = "SPLFJOB_S"

End Sub

'---------------------------------------------------------
Public Sub recSplfFtp_Init(recSplfFtp As typeSplffTP)
'---------------------------------------------------------
recSplfFtp.obj = "SPLFFTP"
recSplfFtp.Method = ""
recSplfFtp.Err = ""
recSplfFtp.SPFSEQ = 0
recSplfFtp.SPFEVTID = 0
recSplfFtp.SPFSAUT = ""
recSplfFtp.SPFTXT = ""
End Sub

