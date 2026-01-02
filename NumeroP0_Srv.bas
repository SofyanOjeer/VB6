Attribute VB_Name = "srvNumeroP0"

'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recNumeroP0Len = 50 ' 34 + 16
Public Const recNumeroP0_Block = 50

Type typeNumeroP0
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    COSOC             As String * 3
    COAGC             As String * 3
    
    NOCTR             As String * 3
    CTEUR            As Long
    
End Type
    
Public arrNumeroP0() As typeNumeroP0
Public arrNumeroP0_NB As Integer
Public arrNumeroP0_NBMax As Integer
Public arrNumeroP0_Index As Integer
Public arrNumeroP0_Suite As Boolean

Public xNumeroP0 As typeNumeroP0

'-----------------------------------------------------
Function srvNumeroP0_Update(recNumeroP0 As typeNumeroP0)
'-----------------------------------------------------

srvNumeroP0_Update = "?"

MsgTxtLen = 0
Call srvNumeroP0_PutBuffer(recNumeroP0)

If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvNumeroP0_GetBuffer(recNumeroP0)) Then
        Call srvNumeroP0_Error(recNumeroP0)
        srvNumeroP0_Update = recNumeroP0.Err
        Exit Function
    Else
        srvNumeroP0_Update = Null
    End If
Else
    recNumeroP0.Err = "srv"
End If


'=====================================================
End Function



'-----------------------------------------------------
Public Function srvNumeroP0_Monitor(recNumeroP0 As typeNumeroP0)
'-----------------------------------------------------
blnFR_Convert = False

arrNumeroP0_Suite = False
Select Case mId$(Trim(recNumeroP0.Method), 1, 4)
    Case "Seek", "Add "
                srvNumeroP0_Monitor = srvNumeroP0_Seek(recNumeroP0)
    Case "Snap"
              srvNumeroP0_Monitor = srvNumeroP0_Snap(recNumeroP0)
    Case Else
                recNumeroP0.Err = recNumeroP0.Method
                Call srvNumeroP0_Error(recNumeroP0)
                srvNumeroP0_Monitor = recNumeroP0.Err
End Select

End Function

'-----------------------------------------------------
Sub srvNumeroP0_Error(recNumeroP0 As typeNumeroP0)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "NumeroP0" & Chr$(10) & Chr$(13)

Select Case mId$(recNumeroP0.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recNumeroP0.Err
        I = vbCritical
End Select

MsgBox Msg & " : " & recNumeroP0.NOCTR & " : " & recNumeroP0.CTEUR _
        , I, "module : NumeroP0s.bas  ( " & Trim(recNumeroP0.obj) & " : " & Trim(recNumeroP0.Method) & " )"

End Sub



'---------------------------------------------------------
Public Function srvNumeroP0_GetBuffer(recNumeroP0 As typeNumeroP0)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvNumeroP0_GetBuffer = Null
recNumeroP0.obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recNumeroP0.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recNumeroP0.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recNumeroP0.Err = Space$(10) Then
    recNumeroP0.COSOC = mId$(MsgTxt, K + 1, 3)
    recNumeroP0.COAGC = mId$(MsgTxt, K + 4, 3)
    
    recNumeroP0.NOCTR = mId$(MsgTxt, K + 7, 3)
    recNumeroP0.CTEUR = CLng(Val(mId$(MsgTxt, K + 10, 7)))

Else
    srvNumeroP0_GetBuffer = recNumeroP0.Err
End If

MsgTxtIndex = MsgTxtIndex + recNumeroP0Len

End Function

'---------------------------------------------------------
Private Sub srvNumeroP0_PutBuffer(recNumeroP0 As typeNumeroP0)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recNumeroP0.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recNumeroP0.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34
    Mid$(MsgTxt, K + 1, 3) = recNumeroP0.COSOC
    Mid$(MsgTxt, K + 4, 3) = recNumeroP0.COAGC
    
    Mid$(MsgTxt, K + 7, 3) = Format$(recNumeroP0.NOCTR, "000")
    Mid$(MsgTxt, K + 10, 7) = Format$(recNumeroP0.CTEUR, "0000000")


MsgTxtLen = MsgTxtLen + recNumeroP0Len


  
End Sub



'---------------------------------------------------------
Private Function srvNumeroP0_Seek(recNumeroP0 As typeNumeroP0)
'---------------------------------------------------------

srvNumeroP0_Seek = "?"
MsgTxtLen = 0
Call srvNumeroP0_PutBuffer(recNumeroP0)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If IsNull(srvNumeroP0_GetBuffer(recNumeroP0)) Then
        srvNumeroP0_Seek = Null
    Else
       '' Call srvNumeroP0_Error(recNumeroP0)
    End If
End If

End Function

'---------------------------------------------------------
Private Function srvNumeroP0_Snap(recNumeroP0 As typeNumeroP0)
'---------------------------------------------------------
srvNumeroP0_Snap = "?"
MsgTxtLen = 0
Call srvNumeroP0_PutBuffer(recNumeroP0)
Call srvNumeroP0_PutBuffer(arrNumeroP0(0))
If IsNull(SndRcv()) Then
    srvNumeroP0_Snap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(srvNumeroP0_GetBuffer(recNumeroP0)) Then
            Call arrNumeroP0_AddItem(recNumeroP0)
            arrNumeroP0_Suite = True
        Else
            arrNumeroP0_Suite = False
            Exit Do
        End If
    Loop
End If

End Function

'---------------------------------------------------------
Public Sub recNumeroP0_Init(recNumeroP0 As typeNumeroP0)
'---------------------------------------------------------
MsgTxt = Space$(recNumeroP0Len)
MsgTxtIndex = 0
Call srvNumeroP0_GetBuffer(recNumeroP0)
recNumeroP0.obj = "SRVNUMERO"
recNumeroP0.COSOC = SocId$
recNumeroP0.COAGC = SocAgence$
recNumeroP0.NOCTR = "000"
End Sub

'---------------------------------------------------------
Public Sub arrNumeroP0_AddItem(recNumeroP0 As typeNumeroP0)
'---------------------------------------------------------
          
arrNumeroP0_NB = arrNumeroP0_NB + 1
    
If arrNumeroP0_NB > arrNumeroP0_NBMax Then
    arrNumeroP0_NBMax = arrNumeroP0_NBMax + recNumeroP0_Block
    ReDim Preserve arrNumeroP0(arrNumeroP0_NBMax)
End If
            
arrNumeroP0(arrNumeroP0_NB) = recNumeroP0
End Sub




