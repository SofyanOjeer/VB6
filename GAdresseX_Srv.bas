Attribute VB_Name = "srvGAdresseX"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recGAdresseXLen = 277 ' 34 + 243

Type typeGAdresseX
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    Compte                  As String * 11
    Devise                  As String * 2
    Nature                  As String * 3
    Séquence                As Integer
    L0                      As String * 38
    L1                      As String * 38
    L2                      As String * 38
    L3                      As String * 38
    L4                      As String * 32
    CodePostal              As String * 5
    Pays                    As String * 2
    PaysLibellé             As String * 32
    Routage                 As String * 1
     
End Type
    
Public arrGAdresseX() As typeGAdresseX
Public arrGAdresseX_NB As Integer
Public arrGAdresseX_NBMax As Integer
Public arrGAdresseX_Index As Integer
Public arrGAdresseX_Suite As Boolean
'-----------------------------------------------------
Public Function srvGAdresseX_Monitor(recGAdresseX As typeGAdresseX)
'-----------------------------------------------------

arrGAdresseX_Suite = False
Select Case mId$(Trim(recGAdresseX.Method), 1, 4)
    Case "Seek"
                srvGAdresseX_Monitor = srvGAdresseX_Seek(recGAdresseX)
    Case "Snap"
              srvGAdresseX_Monitor = srvGAdresseX_Snap(recGAdresseX)
    Case Else
                recGAdresseX.Err = recGAdresseX.Method
                Call srvGAdresseX_Error(recGAdresseX)
                srvGAdresseX_Monitor = recGAdresseX.Err
End Select

End Function

'-----------------------------------------------------
Sub srvGAdresseX_Error(recGAdresseX As typeGAdresseX)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "GAdresseX" & Chr$(10) & Chr$(13)

Select Case mId$(recGAdresseX.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recGAdresseX.Err
        I = vbCritical
End Select

MsgBox Msg, I, "module : GAdresseX.bas  ( " _
                & Trim(recGAdresseX.obj) & " : " & Trim(recGAdresseX.Method) & " )"

End Sub



'---------------------------------------------------------
Public Function srvGAdresseX_GetBuffer(recGAdresseX As typeGAdresseX)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvGAdresseX_GetBuffer = Null
recGAdresseX.obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recGAdresseX.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recGAdresseX.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recGAdresseX.Err = Space$(10) Then
    recGAdresseX.Compte = mId$(MsgTxt, K + 1, 11)
    recGAdresseX.Devise = mId$(MsgTxt, K + 12, 3)
    recGAdresseX.Nature = mId$(MsgTxt, K + 15, 3)
    recGAdresseX.Séquence = CInt(Val(mId$(MsgTxt, K + 18, 2)))
    recGAdresseX.L0 = mId$(MsgTxt, K + 20, 38)
    recGAdresseX.L1 = mId$(MsgTxt, K + 58, 38)
    recGAdresseX.L2 = mId$(MsgTxt, K + 96, 38)
    recGAdresseX.L3 = mId$(MsgTxt, K + 134, 38)
    recGAdresseX.L4 = mId$(MsgTxt, K + 172, 32)
    recGAdresseX.CodePostal = mId$(MsgTxt, K + 204, 5)
    recGAdresseX.Pays = mId$(MsgTxt, K + 209, 2)
    
    recGAdresseX.PaysLibellé = mId$(MsgTxt, K + 211, 32)
    recGAdresseX.Routage = mId$(MsgTxt, K + 243, 1)

Else
    srvGAdresseX_GetBuffer = recGAdresseX.Err
End If

MsgTxtIndex = MsgTxtIndex + recGAdresseXLen

End Function

'---------------------------------------------------------
Private Sub srvGAdresseX_PutBuffer(recGAdresseX As typeGAdresseX)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recGAdresseX.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recGAdresseX.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34

Mid$(MsgTxt, K + 1, 11) = recGAdresseX.Compte
Mid$(MsgTxt, K + 12, 3) = recGAdresseX.Devise
Mid$(MsgTxt, K + 15, 3) = recGAdresseX.Nature
Mid$(MsgTxt, K + 18, 2) = Format$(recGAdresseX.Séquence, "00")

MsgTxtLen = MsgTxtLen + recGAdresseXLen
End Sub



'---------------------------------------------------------
Private Function srvGAdresseX_Seek(recGAdresseX As typeGAdresseX)
'---------------------------------------------------------

srvGAdresseX_Seek = "?"
MsgTxtLen = 0
Call srvGAdresseX_PutBuffer(recGAdresseX)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If MsgTxtLen > 0 Then
        If IsNull(srvGAdresseX_GetBuffer(recGAdresseX)) Then
            srvGAdresseX_Seek = Null
        Else
            Call srvGAdresseX_Error(recGAdresseX)
        End If
    End If
End If

End Function

'---------------------------------------------------------
Private Function srvGAdresseX_Snap(recGAdresseX As typeGAdresseX)
'---------------------------------------------------------
srvGAdresseX_Snap = "?"
MsgTxtLen = 0
Call srvGAdresseX_PutBuffer(recGAdresseX)
Call srvGAdresseX_PutBuffer(arrGAdresseX(0))
If IsNull(SndRcv()) Then
    srvGAdresseX_Snap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(srvGAdresseX_GetBuffer(recGAdresseX)) Then
            Call arrGAdresseX_AddItem(recGAdresseX)
            arrGAdresseX_Suite = True
        Else
            arrGAdresseX_Suite = False
            Exit Do
        End If
    Loop
End If

End Function

'---------------------------------------------------------
Public Sub recGAdresseX_Init(recGAdresseX As typeGAdresseX)
'---------------------------------------------------------
MsgTxt = Space$(recGAdresseXLen)
MsgTxtIndex = 0
Call srvGAdresseX_GetBuffer(recGAdresseX)
recGAdresseX.obj = "SRVGADRESS"
End Sub

'---------------------------------------------------------
Public Sub arrGAdresseX_AddItem(recGAdresseX As typeGAdresseX)
'---------------------------------------------------------
          
arrGAdresseX_NB = arrGAdresseX_NB + 1
    
If arrGAdresseX_NB > arrGAdresseX_NBMax Then
    arrGAdresseX_NBMax = arrGAdresseX_NBMax + 10
    ReDim Preserve arrGAdresseX(arrGAdresseX_NBMax)
End If
            
arrGAdresseX(arrGAdresseX_NB) = recGAdresseX
End Sub



