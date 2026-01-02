Attribute VB_Name = "srvSABUSRP0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const memoSABUSRP0Len = 95
Public Const recSABUSRP0Len = 129 ' 34 + 95
Public Const recSABUSRP0_Block = 200
Type typeSABUSRP0
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    SABUSRCUT        As Long                           ' code SAB UTILISATEUR
    SABUSRUTIP       As String * 10                    ' UTILISATEUR PRODUCTION
    SABUSRNOMP       As String * 30                    ' NOM PRODUCTION
    SABUSRUTIT       As String * 10                    ' UTILISATEUR TEST
    SABUSRNOMT       As String * 30                    '  NOM test
    SABUSRSAAU       As String * 10                    ' SAA UNIT
End Type
    
Public Sub srvSABUSRP0_Load(lSABUSRP0() As typeSABUSRP0, lSABUSRP0_Nb As Integer)
Dim mMethod As String, blnSABUSRP0_Suite
Dim wNbMax As Integer
Dim wSABUSRP0 As typeSABUSRP0

mMethod = Trim(lSABUSRP0(0).Method) & "+"
blnSABUSRP0_Suite = True: lSABUSRP0_Nb = 0
wNbMax = recSABUSRP0_Block + 2: ReDim Preserve lSABUSRP0(wNbMax)

wSABUSRP0 = lSABUSRP0(1)
Do Until Not blnSABUSRP0_Suite
    MsgTxtLen = 0
    Call srvSABUSRP0_PutBuffer(wSABUSRP0)
    Call srvSABUSRP0_PutBuffer(lSABUSRP0(0))
    If IsNull(SndRcv()) Then
        MsgTxtIndex = 0
        Do While MsgTxtIndex < MsgTxtLen
            If IsNull(srvSABUSRP0_GetBuffer(wSABUSRP0)) Then
            
                lSABUSRP0_Nb = lSABUSRP0_Nb + 1
                If lSABUSRP0_Nb > wNbMax Then
                    wNbMax = wNbMax + recSABUSRP0_Block
                    ReDim Preserve lSABUSRP0(wNbMax)
                End If
            
                lSABUSRP0(lSABUSRP0_Nb) = wSABUSRP0
                blnSABUSRP0_Suite = True
            Else
                blnSABUSRP0_Suite = False
                Exit Do
            End If
        Loop
    End If

    lSABUSRP0(0).Method = mMethod
Loop

End Sub

'-----------------------------------------------------
Public Function srvSABUSRP0_Monitor(recSABUSRP0 As typeSABUSRP0)
'-----------------------------------------------------

Select Case mId$(Trim(recSABUSRP0.Method), 1, 4)
    Case "Seek"
                srvSABUSRP0_Monitor = srvSABUSRP0_Seek(recSABUSRP0)
    Case Else
                recSABUSRP0.Err = recSABUSRP0.Method
                Call srvSABUSRP0_Error(recSABUSRP0)
                srvSABUSRP0_Monitor = recSABUSRP0.Err
End Select

End Function

'-----------------------------------------------------
Sub srvSABUSRP0_Error(recSABUSRP0 As typeSABUSRP0)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "SABUSRP0" & Chr$(10) & Chr$(13)

Select Case mId$(recSABUSRP0.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recSABUSRP0.Err
        I = vbCritical
End Select

MsgBox Msg, I, "module : SABUSRP0s.bas  ( " _
                & Trim(recSABUSRP0.obj) & " : " & Trim(recSABUSRP0.Method) & " )"

End Sub



'---------------------------------------------------------
Public Function srvSABUSRP0_GetBuffer(recSABUSRP0 As typeSABUSRP0)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvSABUSRP0_GetBuffer = Null
recSABUSRP0.obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recSABUSRP0.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recSABUSRP0.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recSABUSRP0.Err = Space$(10) Then
    recSABUSRP0.SABUSRCUT = CLng(Val(mId$(MsgTxt, K + 1, 5)))
    recSABUSRP0.SABUSRUTIP = mId$(MsgTxt, K + 6, 10)
    recSABUSRP0.SABUSRNOMP = mId$(MsgTxt, K + 16, 30)
    recSABUSRP0.SABUSRUTIT = mId$(MsgTxt, K + 46, 10)
    recSABUSRP0.SABUSRNOMT = mId$(MsgTxt, K + 56, 30)
    recSABUSRP0.SABUSRSAAU = mId$(MsgTxt, K + 86, 10)
Else
    srvSABUSRP0_GetBuffer = recSABUSRP0.Err
End If

MsgTxtIndex = MsgTxtIndex + recSABUSRP0Len

End Function

'---------------------------------------------------------
Public Sub srvSABUSRP0_PutBuffer(recSABUSRP0 As typeSABUSRP0)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, recSABUSRP0Len) = Space$(recSABUSRP0Len)

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recSABUSRP0.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recSABUSRP0.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34
    Mid$(MsgTxt, K + 1, 5) = Format$(recSABUSRP0.SABUSRCUT, "00000")
    Mid$(MsgTxt, K + 6, 10) = recSABUSRP0.SABUSRUTIP
    Mid$(MsgTxt, K + 16, 30) = recSABUSRP0.SABUSRNOMP
    Mid$(MsgTxt, K + 46, 10) = recSABUSRP0.SABUSRUTIT
    Mid$(MsgTxt, K + 56, 30) = recSABUSRP0.SABUSRNOMT
    Mid$(MsgTxt, K + 86, 10) = recSABUSRP0.SABUSRSAAU

MsgTxtLen = MsgTxtLen + recSABUSRP0Len
End Sub


'---------------------------------------------------------
Private Function srvSABUSRP0_Seek(recSABUSRP0 As typeSABUSRP0)
'---------------------------------------------------------

srvSABUSRP0_Seek = "?"
MsgTxtLen = 0
Call srvSABUSRP0_PutBuffer(recSABUSRP0)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If MsgTxtLen > 0 Then
        If IsNull(srvSABUSRP0_GetBuffer(recSABUSRP0)) Then
            srvSABUSRP0_Seek = Null
        Else
            Call srvSABUSRP0_Error(recSABUSRP0)
        End If
    End If
End If

End Function

'-----------------------------------------------------
Function srvSABUSRP0_Update(recSABUSRP0 As typeSABUSRP0)
'-----------------------------------------------------

srvSABUSRP0_Update = "?"

MsgTxtLen = 0
Call srvSABUSRP0_PutBuffer(recSABUSRP0)

If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvSABUSRP0_GetBuffer(recSABUSRP0)) Then
        Call srvSABUSRP0_Error(recSABUSRP0)
        srvSABUSRP0_Update = recSABUSRP0.Err
        Exit Function
    Else
        srvSABUSRP0_Update = Null
    End If
Else
    recSABUSRP0.Err = "srv"
End If


'=====================================================
End Function



'---------------------------------------------------------
Public Sub recSABUSRP0_Init(recSABUSRP0 As typeSABUSRP0)
'---------------------------------------------------------
MsgTxt = Space$(recSABUSRP0Len)
MsgTxtIndex = 0
Call srvSABUSRP0_GetBuffer(recSABUSRP0)
recSABUSRP0.obj = "SABUSRP0_S"

End Sub







