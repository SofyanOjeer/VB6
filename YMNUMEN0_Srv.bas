Attribute VB_Name = "srvYMNUMEN0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recYMNUMEN0Len = 77 ' 34 + 43
Public Const recYMNUMEN0_Block = 200
Type typeYMNUMEN0
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    MNUMENETB       As Integer                        ' ETABLISSEMENT
    MNUMENCGR       As Integer                        ' CODE UTIL. GROUPE
    MNUMENPRE       As Long                           ' CODE OPTION PRECED
    MNUMENORD       As Long                           ' ORDRE DANS MENU
    MNUMENCOD       As Long                           ' CODE OPTION
    MNUMENOIA       As String * 1                     ' INTER-AGENCE
    MNUMENJOQ       As String * 10                    ' FILE ATTENT.BATCH
    
    MNUMENCGR_0       As String * 5                       ' CODE UTIL. GROUPE
    MNUMENPRE_0       As String * 7                        ' CODE OPTION PRECED
    MNUMENORD_0       As String * 5                        ' ORDRE DANS MENU
    MNUMENCOD_0       As String * 7                        ' CODE OPTION

End Type
    
Public Sub srvYMNUMEN0_Load(lYMNUMEN0() As typeYMNUMEN0, lYMNUMEN0_Nb As Integer)
Dim mMethod As String, blnYMNUMEN0_Suite
Dim wNbMax As Integer
Dim wYMNUMEN0 As typeYMNUMEN0

mMethod = Trim(lYMNUMEN0(0).Method) & "+"
blnYMNUMEN0_Suite = True: lYMNUMEN0_Nb = 0
wNbMax = recYMNUMEN0_Block + 2: ReDim Preserve lYMNUMEN0(wNbMax)

wYMNUMEN0 = lYMNUMEN0(1)
Do Until Not blnYMNUMEN0_Suite
    MsgTxtLen = 0
    Call srvYMNUMEN0_PutBuffer(wYMNUMEN0)
    Call srvYMNUMEN0_PutBuffer(lYMNUMEN0(0))
    If IsNull(SndRcv()) Then
        MsgTxtIndex = 0
        Do While MsgTxtIndex < MsgTxtLen
            If IsNull(srvYMNUMEN0_GetBuffer(wYMNUMEN0)) Then
            
                lYMNUMEN0_Nb = lYMNUMEN0_Nb + 1
                If lYMNUMEN0_Nb > wNbMax Then
                    wNbMax = wNbMax + recYMNUMEN0_Block
                    ReDim Preserve lYMNUMEN0(wNbMax)
                End If
            
                lYMNUMEN0(lYMNUMEN0_Nb) = wYMNUMEN0
                blnYMNUMEN0_Suite = True
            Else
                blnYMNUMEN0_Suite = False
                Exit Do
            End If
        Loop
    End If

    lYMNUMEN0(0).Method = mMethod
Loop

End Sub

'-----------------------------------------------------
Public Function srvYMNUMEN0_Monitor(recYMNUMEN0 As typeYMNUMEN0)
'-----------------------------------------------------

Select Case mId$(Trim(recYMNUMEN0.Method), 1, 4)
    Case "Seek", "YMNU", "Dele"
                srvYMNUMEN0_Monitor = srvYMNUMEN0_Seek(recYMNUMEN0)
    Case Else
                recYMNUMEN0.Err = recYMNUMEN0.Method
                Call srvYMNUMEN0_Error(recYMNUMEN0)
                srvYMNUMEN0_Monitor = recYMNUMEN0.Err
End Select

End Function

'-----------------------------------------------------
Sub srvYMNUMEN0_Error(recYMNUMEN0 As typeYMNUMEN0)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "YMNUMEN0" & Chr$(10) & Chr$(13)

Select Case mId$(recYMNUMEN0.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recYMNUMEN0.Err
        I = vbCritical
End Select

MsgBox Msg, I, "module : YMNUMEN0s.bas  ( " _
                & Trim(recYMNUMEN0.obj) & " : " & Trim(recYMNUMEN0.Method) & " )"

End Sub



'---------------------------------------------------------
Public Function srvYMNUMEN0_GetBuffer(recYMNUMEN0 As typeYMNUMEN0)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvYMNUMEN0_GetBuffer = Null
recYMNUMEN0.obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recYMNUMEN0.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recYMNUMEN0.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recYMNUMEN0.Err = Space$(10) Then
    recYMNUMEN0.MNUMENETB = CInt(Val(mId$(MsgTxt, K + 1, 5)))
    recYMNUMEN0.MNUMENCGR = CInt(Val(mId$(MsgTxt, K + 6, 5)))
    recYMNUMEN0.MNUMENPRE = CLng(Val(mId$(MsgTxt, K + 11, 8)))
    recYMNUMEN0.MNUMENORD = CLng(Val(mId$(MsgTxt, K + 19, 6)))
    recYMNUMEN0.MNUMENCOD = CLng(Val(mId$(MsgTxt, K + 25, 8)))
    recYMNUMEN0.MNUMENOIA = mId$(MsgTxt, K + 33, 1)
    recYMNUMEN0.MNUMENJOQ = mId$(MsgTxt, K + 34, 10)
    
   
    recYMNUMEN0.MNUMENCGR_0 = Format(recYMNUMEN0.MNUMENCGR, "00000")
    recYMNUMEN0.MNUMENPRE_0 = Format(recYMNUMEN0.MNUMENPRE, "0000000")
    recYMNUMEN0.MNUMENORD_0 = Format(recYMNUMEN0.MNUMENORD, "00000")
    recYMNUMEN0.MNUMENCOD_0 = Format(recYMNUMEN0.MNUMENCOD, "0000000")
Else
    srvYMNUMEN0_GetBuffer = recYMNUMEN0.Err
End If

MsgTxtIndex = MsgTxtIndex + recYMNUMEN0Len

End Function

'---------------------------------------------------------
Public Sub srvYMNUMEN0_PutBuffer(recYMNUMEN0 As typeYMNUMEN0)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, recYMNUMEN0Len) = Space$(recYMNUMEN0Len)

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recYMNUMEN0.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recYMNUMEN0.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34

    Mid$(MsgTxt, K + 1, 5) = Format$(recYMNUMEN0.MNUMENETB, "0000 ")
    Mid$(MsgTxt, K + 6, 5) = Format$(recYMNUMEN0.MNUMENCGR, "0000 ")
    Mid$(MsgTxt, K + 11, 8) = Format$(recYMNUMEN0.MNUMENPRE, "0000000 ")
    Mid$(MsgTxt, K + 19, 6) = Format$(recYMNUMEN0.MNUMENORD, "00000 ")
    Mid$(MsgTxt, K + 25, 8) = Format$(recYMNUMEN0.MNUMENCOD, "0000000 ")
    Mid$(MsgTxt, K + 33, 1) = recYMNUMEN0.MNUMENOIA
    Mid$(MsgTxt, K + 34, 10) = recYMNUMEN0.MNUMENJOQ

MsgTxtLen = MsgTxtLen + recYMNUMEN0Len
End Sub

'---------------------------------------------------------
Private Function srvYMNUMEN0_Seek(recYMNUMEN0 As typeYMNUMEN0)
'---------------------------------------------------------

srvYMNUMEN0_Seek = "?"
MsgTxtLen = 0
Call srvYMNUMEN0_PutBuffer(recYMNUMEN0)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If MsgTxtLen > 0 Then
        If IsNull(srvYMNUMEN0_GetBuffer(recYMNUMEN0)) Then
            srvYMNUMEN0_Seek = Null
        Else
            Call srvYMNUMEN0_Error(recYMNUMEN0)
        End If
    End If
End If

End Function

Public Sub srvYMNUMEN0_ElpDisplay(recYMNUMEN0 As typeYMNUMEN0)
frmElpDisplay.fgData.Rows = 8
frmElpDisplay.fgData.Row = 1
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MNUMENETB    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ETABLISSEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYMNUMEN0.MNUMENETB
frmElpDisplay.fgData.Row = 2
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MNUMENCGR    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE UTIL. GROUPE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYMNUMEN0.MNUMENCGR
frmElpDisplay.fgData.Row = 3
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MNUMENPRE    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE OPTION PRECED"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYMNUMEN0.MNUMENPRE
frmElpDisplay.fgData.Row = 4
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MNUMENORD    5P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ORDRE DANS MENU"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYMNUMEN0.MNUMENORD
frmElpDisplay.fgData.Row = 5
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MNUMENCOD    7P"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE OPTION"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYMNUMEN0.MNUMENCOD
frmElpDisplay.fgData.Row = 6
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MNUMENOIA    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "INTER-AGENCE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYMNUMEN0.MNUMENOIA
frmElpDisplay.fgData.Row = 7
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MNUMENJOQ   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "FILE ATTENT.BATCH"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYMNUMEN0.MNUMENJOQ
frmElpDisplay.Show vbModal
End Sub

'-----------------------------------------------------
Function srvYMNUMEN0_Update(recYMNUMEN0 As typeYMNUMEN0)
'-----------------------------------------------------

srvYMNUMEN0_Update = "?"

MsgTxtLen = 0
Call srvYMNUMEN0_PutBuffer(recYMNUMEN0)

If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvYMNUMEN0_GetBuffer(recYMNUMEN0)) Then
        Call srvYMNUMEN0_Error(recYMNUMEN0)
        srvYMNUMEN0_Update = recYMNUMEN0.Err
        Exit Function
    Else
        srvYMNUMEN0_Update = Null
    End If
Else
    recYMNUMEN0.Err = "srv"
End If


'=====================================================
End Function



'---------------------------------------------------------
Public Sub recYMNUMEN0_Init(recYMNUMEN0 As typeYMNUMEN0)
'---------------------------------------------------------
MsgTxt = Space$(recYMNUMEN0Len)
MsgTxtIndex = 0
Call srvYMNUMEN0_GetBuffer(recYMNUMEN0)
recYMNUMEN0.obj = "ZMNUMEN0_S"
recYMNUMEN0.MNUMENETB = 1

End Sub




