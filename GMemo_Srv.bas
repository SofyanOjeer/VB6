Attribute VB_Name = "srvGMemo"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recgMemoLen = 384 ' 34 + 350
Public Const recGMemo_Block = 20

Type typegMemo
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    IdRéférence             As Long
    MemoSéquence            As Long
    MemoSéquencePlus        As Integer
    
    Application             As String * 5
    FluxSéquence            As Long
    EchSéquence             As Long
    
    MemoLien1               As Long
    MemoLien2               As Long
    MemoNature              As String * 10
    MemoText                As String * 256
   
    Statut                  As String * 1
    StatutPlus              As String * 2
    Flag1                   As String * 1
    Flag2                   As String * 1
    Flag3                   As String * 1
    
    ElpId                   As Long
    ElpUpdate               As Integer
    ElpControl              As String * 10
    
End Type
    
Public arrgMemo() As typegMemo
Public arrgMemo_NB As Integer
Public arrgMemo_NBMax As Integer
Public arrgMemo_Index As Integer
Public arrgMemo_Suite As Boolean

Public xGMemo As typegMemo

'-----------------------------------------------------
Function srvGMemo_Update(recGMemo As typegMemo)
'-----------------------------------------------------

If blnMsgTxt_Concat_Transaction Then
    Call srvGMemo_PutBuffer(recGMemo)
    srvGMemo_Update = Null
    Exit Function
End If

srvGMemo_Update = "?"

MsgTxtLen = 0
Call srvGMemo_PutBuffer(recGMemo)

If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvGMemo_GetBuffer(recGMemo)) Then
        Call srvGMemo_Error(recGMemo)
        srvGMemo_Update = recGMemo.Err
        Exit Function
    Else
        srvGMemo_Update = Null
    End If
Else
    recGMemo.Err = "srv"
End If


'=====================================================
End Function



Public Sub srvGMemo_Load(recGMemoMin As typegMemo, recGMemoMax As typegMemo)
Dim mMethod As String

mMethod = Trim(recGMemoMin.Method) & "+"
arrgMemo_NBMax = 0
arrgMemo_Suite = True: arrgMemo_NB = 0
arrgMemo_NBMax = recGMemo_Block: ReDim arrgMemo(arrgMemo_NBMax)

arrgMemo(0) = recGMemoMax
arrgMemo_Suite = True
Do Until Not arrgMemo_Suite
    srvGMemo_Monitor recGMemoMin
    recGMemoMin = arrgMemo(arrgMemo_NB)
    recGMemoMin.Method = mMethod
Loop

End Sub

'-----------------------------------------------------
Function srvGMemo_Dtaq_Put(lFct As String, recGMemo As typegMemo)
'-----------------------------------------------------

srvGMemo_Dtaq_Put = Null
Select Case lFct
    Case "Init": MsgTxtLen = 0
    Case "Add": Call srvGMemo_PutBuffer(recGMemo)
                If MsgTxtLen + recgMemoLen >= recGMemo_Block * recgMemoLen Then
                    Call srvGMemo_Dtaq_Snd(recGMemo): MsgTxtLen = 0
                End If
    Case "Snd": If MsgTxtLen > 0 Then Call srvGMemo_Dtaq_Snd(recGMemo)
    Case Else: srvGMemo_Dtaq_Put = lFct
End Select
'=====================================================
End Function


'-----------------------------------------------------
Function srvGMemo_Dtaq_Snd(recGMemo As typegMemo)
'-----------------------------------------------------

srvGMemo_Dtaq_Snd = "?"

If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvGMemo_GetBuffer(recGMemo)) Then
        Call srvGMemo_Error(recGMemo)
        srvGMemo_Dtaq_Snd = recGMemo.Err
        Exit Function
    Else
        srvGMemo_Dtaq_Snd = Null
    End If
Else
    recGMemo.Err = "Snd"
End If


'=====================================================
End Function



'-----------------------------------------------------
Public Function srvGMemo_Monitor(recGMemo As typegMemo)
'-----------------------------------------------------
blnFR_Convert = False

arrgMemo_Suite = False
Select Case mId$(Trim(recGMemo.Method), 1, 4)
    Case "Seek", "Comp", "NUML"
                srvGMemo_Monitor = srvGMemo_Seek(recGMemo)
    Case "Snap"
              srvGMemo_Monitor = srvGMemo_Snap(recGMemo)
    Case Else
                recGMemo.Err = recGMemo.Method
                Call srvGMemo_Error(recGMemo)
                srvGMemo_Monitor = recGMemo.Err
End Select

End Function

'-----------------------------------------------------
Sub srvGMemo_Error(recGMemo As typegMemo)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "GMemo" & Chr$(10) & Chr$(13)

Select Case mId$(recGMemo.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recGMemo.Err
        I = vbCritical
End Select

MsgBox Msg & " : " & recGMemo.IdRéférence & " : " & recGMemo.MemoSéquence _
        , I, "module : GMemos.bas  ( " & Trim(recGMemo.obj) & " : " & Trim(recGMemo.Method) & " )"

End Sub



'---------------------------------------------------------
Public Function srvGMemo_GetBuffer(recGMemo As typegMemo)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvGMemo_GetBuffer = Null
recGMemo.obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recGMemo.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recGMemo.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recGMemo.Err = Space$(10) Then
    recGMemo.IdRéférence = CLng(Val(mId$(MsgTxt, K + 1, 12)))
    recGMemo.MemoSéquence = CLng(Val(mId$(MsgTxt, K + 13, 5)))
    recGMemo.MemoSéquencePlus = CInt(Val(mId$(MsgTxt, K + 18, 1)))
    
    recGMemo.Application = mId$(MsgTxt, K + 19, 5)
    recGMemo.FluxSéquence = CLng(Val(mId$(MsgTxt, K + 24, 5)))
    recGMemo.EchSéquence = CLng(Val(mId$(MsgTxt, K + 29, 5)))
    
    recGMemo.MemoLien1 = CLng(Val(mId$(MsgTxt, K + 34, 10)))
    recGMemo.MemoLien2 = CLng(Val(mId$(MsgTxt, K + 45, 9)))  'CLng(Val(mId$(MsgTxt, K + 44, 10)))
    recGMemo.MemoNature = mId$(MsgTxt, K + 54, 10)
    recGMemo.MemoText = mId$(MsgTxt, K + 64, 256)
  
    recGMemo.Statut = mId$(MsgTxt, K + 320, 1)
    recGMemo.StatutPlus = mId$(MsgTxt, K + 321, 2)
    recGMemo.Flag1 = mId$(MsgTxt, K + 323, 1)
    recGMemo.Flag2 = mId$(MsgTxt, K + 324, 1)
    recGMemo.Flag3 = mId$(MsgTxt, K + 325, 1)
    recGMemo.ElpId = CLng(Val(mId$(MsgTxt, K + 326, 12)))
    recGMemo.ElpUpdate = CInt(Val(mId$(MsgTxt, K + 338, 3)))
    recGMemo.ElpControl = mId$(MsgTxt, K + 341, 10)

Else
    srvGMemo_GetBuffer = recGMemo.Err
End If

MsgTxtIndex = MsgTxtIndex + recgMemoLen

End Function

'---------------------------------------------------------
Private Sub srvGMemo_PutBuffer(recGMemo As typegMemo)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recGMemo.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recGMemo.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34

Mid$(MsgTxt, K + 1, 12) = Format$(recGMemo.IdRéférence, "000000000000")
Mid$(MsgTxt, K + 13, 5) = Format$(recGMemo.MemoSéquence, "00000")
Mid$(MsgTxt, K + 18, 1) = Format$(recGMemo.MemoSéquencePlus, "00000")

Mid$(MsgTxt, K + 19, 5) = recGMemo.Application

Mid$(MsgTxt, K + 24, 5) = Format$(recGMemo.FluxSéquence, "00000")
Mid$(MsgTxt, K + 29, 5) = Format$(recGMemo.EchSéquence, "00000")

Mid$(MsgTxt, K + 34, 10) = Format$(recGMemo.MemoLien1, "0000000000")
Mid$(MsgTxt, K + 44, 10) = Format$(recGMemo.MemoLien2, "0000000000")
Mid$(MsgTxt, K + 54, 10) = recGMemo.MemoNature
Mid$(MsgTxt, K + 64, 256) = recGMemo.MemoText

Mid$(MsgTxt, K + 320, 1) = recGMemo.Statut
Mid$(MsgTxt, K + 321, 2) = recGMemo.StatutPlus
Mid$(MsgTxt, K + 323, 1) = recGMemo.Flag1
Mid$(MsgTxt, K + 324, 1) = recGMemo.Flag2
Mid$(MsgTxt, K + 325, 1) = recGMemo.Flag3
Mid$(MsgTxt, K + 326, 12) = Format$(recGMemo.ElpId, "000000000000")
Mid$(MsgTxt, K + 338, 3) = Format$(recGMemo.ElpUpdate, "000")
Mid$(MsgTxt, K + 341, 10) = recGMemo.ElpControl

MsgTxtLen = MsgTxtLen + recgMemoLen


  
End Sub



'---------------------------------------------------------
Private Function srvGMemo_Seek(recGMemo As typegMemo)
'---------------------------------------------------------

srvGMemo_Seek = "?"
MsgTxtLen = 0
Call srvGMemo_PutBuffer(recGMemo)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If IsNull(srvGMemo_GetBuffer(recGMemo)) Then
        srvGMemo_Seek = Null
    Else
        Call srvGMemo_Error(recGMemo)
    End If
End If

End Function

'---------------------------------------------------------
Private Function srvGMemo_Snap(recGMemo As typegMemo)
'---------------------------------------------------------
srvGMemo_Snap = "?"
MsgTxtLen = 0
Call srvGMemo_PutBuffer(recGMemo)
Call srvGMemo_PutBuffer(arrgMemo(0))
If IsNull(SndRcv()) Then
    srvGMemo_Snap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(srvGMemo_GetBuffer(recGMemo)) Then
            Call arrGMemo_AddItem(recGMemo)
            arrgMemo_Suite = True
        Else
            arrgMemo_Suite = False
            Exit Do
        End If
    Loop
End If

End Function

'---------------------------------------------------------
Public Sub recGMemo_Init(recGMemo As typegMemo)
'---------------------------------------------------------
MsgTxt = Space$(recgMemoLen)
MsgTxtIndex = 0
Call srvGMemo_GetBuffer(recGMemo)
recGMemo.obj = "SRVGMEMO    "

End Sub

'---------------------------------------------------------
Public Sub arrGMemo_AddItem(recGMemo As typegMemo)
'---------------------------------------------------------
          
arrgMemo_NB = arrgMemo_NB + 1
    
If arrgMemo_NB > arrgMemo_NBMax Then
    arrgMemo_NBMax = arrgMemo_NBMax + recGMemo_Block
    ReDim Preserve arrgMemo(arrgMemo_NBMax)
End If
            
arrgMemo(arrgMemo_NB) = recGMemo
End Sub


