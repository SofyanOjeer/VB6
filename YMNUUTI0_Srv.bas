Attribute VB_Name = "srvYMNUUTI0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const memoYMNUUTI0Len = 37
Public Const recYMNUUTI0Len = 71 ' 34 + 37
Public Const recYMNUUTI0_Block = 200
Type typeYMNUUTI0
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    MNUUTIETB       As Integer                        ' ETABLISSEMENT
    MNUUTICUT       As Integer                        ' CODE UTILISATEUR
    MNUUTICGR       As Integer                        ' CODE GROUPE
    MNUUTIDRG       As String * 1                     ' DROITS GROUPE
    MNUUTIOUT       As String * 10                    ' FILE ATTENTE
    MNUUTILAN       As String * 1                     ' LANGUE
    MNUUTIMSE       As String * 1                     ' MENU SERVICE
    MNUUTIAGE       As Integer                        ' AGENCE DEFAUT
    MNUUTISER       As String * 2                     ' SERVICE DEFAUT
    MNUUTISRV       As String * 2                     ' SOUS-SERV. DEFAUT
  
End Type
    
Public Sub srvYMNUUTI0_Load(lYMNUUTI0() As typeYMNUUTI0, lYMNUUTI0_Nb As Integer)
Dim mMethod As String, blnYMNUUTI0_Suite
Dim wNbMax As Integer
Dim wYMNUUTI0 As typeYMNUUTI0

mMethod = Trim(lYMNUUTI0(0).Method) & "+"
blnYMNUUTI0_Suite = True: lYMNUUTI0_Nb = 0
wNbMax = recYMNUUTI0_Block + 2: ReDim Preserve lYMNUUTI0(wNbMax)

wYMNUUTI0 = lYMNUUTI0(1)
Do Until Not blnYMNUUTI0_Suite
    MsgTxtLen = 0
    Call srvYMNUUTI0_PutBuffer(wYMNUUTI0)
    Call srvYMNUUTI0_PutBuffer(lYMNUUTI0(0))
    If IsNull(SndRcv()) Then
        MsgTxtIndex = 0
        Do While MsgTxtIndex < MsgTxtLen
            If IsNull(srvYMNUUTI0_GetBuffer(wYMNUUTI0)) Then
            
                lYMNUUTI0_Nb = lYMNUUTI0_Nb + 1
                If lYMNUUTI0_Nb > wNbMax Then
                    wNbMax = wNbMax + recYMNUUTI0_Block
                    ReDim Preserve lYMNUUTI0(wNbMax)
                End If
            
                lYMNUUTI0(lYMNUUTI0_Nb) = wYMNUUTI0
                blnYMNUUTI0_Suite = True
            Else
                blnYMNUUTI0_Suite = False
                Exit Do
            End If
        Loop
    End If

    lYMNUUTI0(0).Method = mMethod
Loop

End Sub

'-----------------------------------------------------
Public Function srvYMNUUTI0_Monitor(recYMNUUTI0 As typeYMNUUTI0)
'-----------------------------------------------------

Select Case mId$(Trim(recYMNUUTI0.Method), 1, 4)
    Case "Seek"
                srvYMNUUTI0_Monitor = srvYMNUUTI0_Seek(recYMNUUTI0)
    Case Else
                recYMNUUTI0.Err = recYMNUUTI0.Method
                Call srvYMNUUTI0_Error(recYMNUUTI0)
                srvYMNUUTI0_Monitor = recYMNUUTI0.Err
End Select

End Function

'-----------------------------------------------------
Sub srvYMNUUTI0_Error(recYMNUUTI0 As typeYMNUUTI0)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "YMNUUTI0" & Chr$(10) & Chr$(13)

Select Case mId$(recYMNUUTI0.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recYMNUUTI0.Err
        I = vbCritical
End Select

MsgBox Msg, I, "module : YMNUUTI0s.bas  ( " _
                & Trim(recYMNUUTI0.obj) & " : " & Trim(recYMNUUTI0.Method) & " )"

End Sub



'---------------------------------------------------------
Public Function srvYMNUUTI0_GetBuffer(recYMNUUTI0 As typeYMNUUTI0)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvYMNUUTI0_GetBuffer = Null
recYMNUUTI0.obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recYMNUUTI0.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recYMNUUTI0.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recYMNUUTI0.Err = Space$(10) Then
    recYMNUUTI0.MNUUTIETB = CInt(Val(mId$(MsgTxt, K + 1, 5)))
    recYMNUUTI0.MNUUTICUT = CInt(Val(mId$(MsgTxt, K + 6, 5)))
    recYMNUUTI0.MNUUTICGR = CInt(Val(mId$(MsgTxt, K + 11, 5)))
    recYMNUUTI0.MNUUTIDRG = mId$(MsgTxt, K + 16, 1)
    recYMNUUTI0.MNUUTIOUT = mId$(MsgTxt, K + 17, 10)
    recYMNUUTI0.MNUUTILAN = mId$(MsgTxt, K + 27, 1)
    recYMNUUTI0.MNUUTIMSE = mId$(MsgTxt, K + 28, 1)
    recYMNUUTI0.MNUUTIAGE = CInt(Val(mId$(MsgTxt, K + 29, 5)))
    recYMNUUTI0.MNUUTISER = mId$(MsgTxt, K + 34, 2)
    recYMNUUTI0.MNUUTISRV = mId$(MsgTxt, K + 36, 2)
Else
    srvYMNUUTI0_GetBuffer = recYMNUUTI0.Err
End If

MsgTxtIndex = MsgTxtIndex + recYMNUUTI0Len

End Function

'---------------------------------------------------------
Private Sub srvYMNUUTI0_PutBuffer(recYMNUUTI0 As typeYMNUUTI0)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, recYMNUUTI0Len) = Space$(recYMNUUTI0Len)

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recYMNUUTI0.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recYMNUUTI0.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34

    Mid$(MsgTxt, K + 1, 5) = Format$(recYMNUUTI0.MNUUTIETB, "0000 ")
    Mid$(MsgTxt, K + 6, 5) = Format$(recYMNUUTI0.MNUUTICUT, "0000 ")
    Mid$(MsgTxt, K + 11, 5) = Format$(recYMNUUTI0.MNUUTICGR, "0000 ")
    Mid$(MsgTxt, K + 16, 1) = recYMNUUTI0.MNUUTIDRG
    Mid$(MsgTxt, K + 17, 10) = recYMNUUTI0.MNUUTIOUT
    Mid$(MsgTxt, K + 27, 1) = recYMNUUTI0.MNUUTILAN
    Mid$(MsgTxt, K + 28, 1) = recYMNUUTI0.MNUUTIMSE
    Mid$(MsgTxt, K + 29, 5) = Format$(recYMNUUTI0.MNUUTIAGE, "0000 ")
    Mid$(MsgTxt, K + 34, 2) = recYMNUUTI0.MNUUTISER
    Mid$(MsgTxt, K + 36, 2) = recYMNUUTI0.MNUUTISRV

MsgTxtLen = MsgTxtLen + recYMNUUTI0Len
End Sub
Public Sub srvYMNUUTI0_ElpDisplay(recYMNUUTI0 As typeYMNUUTI0)
frmElpDisplay.fgData.Rows = 11
frmElpDisplay.fgData.Row = 1
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MNUUTIETB    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "ETABLISSEMENT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYMNUUTI0.MNUUTIETB
frmElpDisplay.fgData.Row = 2
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MNUUTICUT    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE UTILISATEUR"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYMNUUTI0.MNUUTICUT
frmElpDisplay.fgData.Row = 3
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MNUUTICGR    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "CODE GROUPE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYMNUUTI0.MNUUTICGR
frmElpDisplay.fgData.Row = 4
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MNUUTIDRG    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "DROITS GROUPE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYMNUUTI0.MNUUTIDRG
frmElpDisplay.fgData.Row = 5
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MNUUTIOUT   10A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "FILE ATTENTE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYMNUUTI0.MNUUTIOUT
frmElpDisplay.fgData.Row = 6
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MNUUTILAN    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "LANGUE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYMNUUTI0.MNUUTILAN
frmElpDisplay.fgData.Row = 7
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MNUUTIMSE    1A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "MENU SERVICE"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYMNUUTI0.MNUUTIMSE
frmElpDisplay.fgData.Row = 8
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MNUUTIAGE    4B"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "AGENCE DEFAUT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYMNUUTI0.MNUUTIAGE
frmElpDisplay.fgData.Row = 9
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MNUUTISER    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SERVICE DEFAUT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYMNUUTI0.MNUUTISER
frmElpDisplay.fgData.Row = 10
frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "MNUUTISRV    2A"
frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = "SOUS-SERV. DEFAUT"
frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recYMNUUTI0.MNUUTISRV
frmElpDisplay.Show vbModal
End Sub

'---------------------------------------------------------
Private Function srvYMNUUTI0_Seek(recYMNUUTI0 As typeYMNUUTI0)
'---------------------------------------------------------

srvYMNUUTI0_Seek = "?"
MsgTxtLen = 0
Call srvYMNUUTI0_PutBuffer(recYMNUUTI0)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If MsgTxtLen > 0 Then
        If IsNull(srvYMNUUTI0_GetBuffer(recYMNUUTI0)) Then
            srvYMNUUTI0_Seek = Null
        Else
            Call srvYMNUUTI0_Error(recYMNUUTI0)
        End If
    End If
End If

End Function

'-----------------------------------------------------
Function srvYMNUUTI0_Update(recYMNUUTI0 As typeYMNUUTI0)
'-----------------------------------------------------

srvYMNUUTI0_Update = "?"

MsgTxtLen = 0
Call srvYMNUUTI0_PutBuffer(recYMNUUTI0)

If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvYMNUUTI0_GetBuffer(recYMNUUTI0)) Then
        Call srvYMNUUTI0_Error(recYMNUUTI0)
        srvYMNUUTI0_Update = recYMNUUTI0.Err
        Exit Function
    Else
        srvYMNUUTI0_Update = Null
    End If
Else
    recYMNUUTI0.Err = "srv"
End If


'=====================================================
End Function



'---------------------------------------------------------
Public Sub recYMNUUTI0_Init(recYMNUUTI0 As typeYMNUUTI0)
'---------------------------------------------------------
MsgTxt = Space$(recYMNUUTI0Len)
MsgTxtIndex = 0
Call srvYMNUUTI0_GetBuffer(recYMNUUTI0)
recYMNUUTI0.obj = "ZMNUUTI0_S"
recYMNUUTI0.MNUUTIETB = 1

End Sub






