Attribute VB_Name = "srvYSWIALI0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recYSWIALI0Len = 578 ' 34 +544
Public Const recYSWIALI0_Block = 20
Public Const memoYSWIALI0Len = 544
Public Const constYSWIALI0 = "YSWIALI0  "

Type typeYSWIALI0
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    SWIALIETA       As Integer                        ' ETABLISSEMENT
    SWIALIAGE       As Integer                        ' AGENCE
    SWIALISER       As String * 2                     ' SERVICE
    SWIALISSE       As String * 2                     ' SERVICE
    SWIALIMES       As String * 3                     ' TYPE MESSAGE
    SWIALINUM       As Long                           ' NUMERO INTERNE
    SWIALINEN       As String * 1                     ' NUMER ENVOI
    SWIALINLI       As Long                           ' NUMERO DE LIGNE
    SWIALIDON       As String * 512                   ' DONNE MESSAGE
    SWIALIOK        As String * 1                     ' PASSE SI OK
End Type
    
    
Public arrYSWIALI0() As typeYSWIALI0
Public arrYSWIALI0_NB As Integer
Public arrYSWIALI0_NBMax As Integer
Public arrYSWIALI0_Index As Integer
Public arrYSWIALI0_Suite As Boolean

'-----------------------------------------------------
Function srvYSWIALI0_Update(recYSWIALI0 As typeYSWIALI0)
'-----------------------------------------------------

srvYSWIALI0_Update = "?"

MsgTxtLen = 0
Call srvYSWIALI0_PutBuffer(recYSWIALI0)

If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvYSWIALI0_GetBuffer(recYSWIALI0)) Then
        Call srvYSWIALI0_Error(recYSWIALI0)
        srvYSWIALI0_Update = recYSWIALI0.Err
        Exit Function
    Else
        srvYSWIALI0_Update = Null
    End If
Else
    recYSWIALI0.Err = "srv"
End If
End Function

'-----------------------------------------------------
Sub srvYSWIALI0_Error(recYSWIALI0 As typeYSWIALI0)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "YSWIALI0" & Chr$(10) & Chr$(13)

Select Case mId$(recYSWIALI0.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recYSWIALI0.Err
        I = vbCritical
End Select

MsgBox Msg & " : " _
        , I, "module : YSWIALI0s.bas  ( " & Trim(recYSWIALI0.obj) & " : " & Trim(recYSWIALI0.Method) & " )"

End Sub


'=====================================================



'-----------------------------------------------------
Public Function srvYSWIALI0_Monitor(recYSWIALI0 As typeYSWIALI0)
'-----------------------------------------------------

arrYSWIALI0_Suite = False
Select Case mId$(Trim(recYSWIALI0.Method), 1, 4)
    Case "Snap"
              srvYSWIALI0_Monitor = srvYSWIALI0_Snap(recYSWIALI0)
    Case Else
            srvYSWIALI0_Monitor = srvYSWIALI0_Seek(recYSWIALI0)
End Select

End Function

'---------------------------------------------------------
Public Function srvYSWIALI0_GetBuffer(recYSWIALI0 As typeYSWIALI0)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvYSWIALI0_GetBuffer = Null
recYSWIALI0.obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recYSWIALI0.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recYSWIALI0.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recYSWIALI0.Err = Space$(10) Then
    recYSWIALI0.SWIALIETA = CInt(Val(mId$(MsgTxt, K + 1, 5)))
    recYSWIALI0.SWIALIAGE = CInt(Val(mId$(MsgTxt, K + 6, 5)))
    recYSWIALI0.SWIALISER = mId$(MsgTxt, K + 11, 2)
    recYSWIALI0.SWIALISSE = mId$(MsgTxt, K + 13, 2)
    recYSWIALI0.SWIALIMES = mId$(MsgTxt, K + 15, 3)
    recYSWIALI0.SWIALINUM = CLng(Val(mId$(MsgTxt, K + 18, 9)))
    recYSWIALI0.SWIALINEN = mId$(MsgTxt, K + 27, 1)
    recYSWIALI0.SWIALINLI = CLng(Val(mId$(MsgTxt, K + 28, 3)))
    recYSWIALI0.SWIALIDON = mId$(MsgTxt, K + 31, 512)
    recYSWIALI0.SWIALIOK = mId$(MsgTxt, K + 543, 1)
Else
    srvYSWIALI0_GetBuffer = recYSWIALI0.Err
End If

MsgTxtIndex = MsgTxtIndex + recYSWIALI0Len

End Function

'---------------------------------------------------------
Public Sub srvYSWIALI0_PutBuffer(recYSWIALI0 As typeYSWIALI0)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recYSWIALI0.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recYSWIALI0.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34

    Mid$(MsgTxt, K + 1, 5) = Format$(recYSWIALI0.SWIALIETA, "0000 ")
    Mid$(MsgTxt, K + 6, 5) = Format$(recYSWIALI0.SWIALIAGE, "0000 ")
    Mid$(MsgTxt, K + 11, 2) = recYSWIALI0.SWIALISER
    Mid$(MsgTxt, K + 13, 2) = recYSWIALI0.SWIALISSE
    Mid$(MsgTxt, K + 15, 3) = recYSWIALI0.SWIALIMES
    Mid$(MsgTxt, K + 18, 9) = Format$(recYSWIALI0.SWIALINUM, "00000000 ")
    Mid$(MsgTxt, K + 27, 1) = recYSWIALI0.SWIALINEN
    Mid$(MsgTxt, K + 28, 3) = Format$(recYSWIALI0.SWIALINLI, "00 ")
    Mid$(MsgTxt, K + 31, 512) = recYSWIALI0.SWIALIDON
    Mid$(MsgTxt, K + 543, 1) = recYSWIALI0.SWIALIOK
MsgTxtLen = MsgTxtLen + recYSWIALI0Len
End Sub



'---------------------------------------------------------
Private Function srvYSWIALI0_Seek(recYSWIALI0 As typeYSWIALI0)
'---------------------------------------------------------

srvYSWIALI0_Seek = "?"
MsgTxtLen = 0
Call srvYSWIALI0_PutBuffer(recYSWIALI0)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If IsNull(srvYSWIALI0_GetBuffer(recYSWIALI0)) Then
        srvYSWIALI0_Seek = Null
    Else
        Call srvYSWIALI0_Error(recYSWIALI0)
    End If
End If

End Function

'---------------------------------------------------------
Private Function srvYSWIALI0_Snap(recYSWIALI0 As typeYSWIALI0)
'---------------------------------------------------------
srvYSWIALI0_Snap = "?"
MsgTxtLen = 0
Call srvYSWIALI0_PutBuffer(recYSWIALI0)
Call srvYSWIALI0_PutBuffer(arrYSWIALI0(0))
If IsNull(SndRcv()) Then
    srvYSWIALI0_Snap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(srvYSWIALI0_GetBuffer(recYSWIALI0)) Then
            Call arrYSWIALI0_AddItem(recYSWIALI0)
            arrYSWIALI0_Suite = True
        Else
            arrYSWIALI0_Suite = False
            Exit Do
        End If
    Loop
End If

End Function
'---------------------------------------------------------
Public Sub arrYSWIALI0_AddItem(recYSWIALI0 As typeYSWIALI0)
'---------------------------------------------------------
          
arrYSWIALI0_NB = arrYSWIALI0_NB + 1
    
If arrYSWIALI0_NB > arrYSWIALI0_NBMax Then
    arrYSWIALI0_NBMax = arrYSWIALI0_NBMax + recYSWIALI0_Block
    ReDim Preserve arrYSWIALI0(arrYSWIALI0_NBMax)
End If
            
arrYSWIALI0(arrYSWIALI0_NB) = recYSWIALI0
End Sub



'---------------------------------------------------------
Public Sub recYSWIALI0_Init(recYSWIALI0 As typeYSWIALI0)
'---------------------------------------------------------
recYSWIALI0.obj = "ZSWIALI0_S"
recYSWIALI0.Method = ""
recYSWIALI0.Err = ""

End Sub










