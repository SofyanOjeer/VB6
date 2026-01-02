Attribute VB_Name = "srvYSWIBIC0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recYSWIBIC0Len = 255 ' 34 +221
Public Const recYSWIBIC0_Block = 100
Public Const memoYSWIBIC0Len = 221
Public Const constYSWIBIC0 = "YSWIBIC0  "

Type typeYSWIBIC0
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    SWIBICBIC       As String * 11                    ' CODE BIC
    SWIBICINT       As String * 105                   ' INTITULE
    SWIBICVIL       As String * 35                    ' VILLE
    SWIBICCOM       As String * 70                    ' COMMENTAIRE
End Type
    
    
Public arrYSWIBIC0() As typeYSWIBIC0
Public arrYSWIBIC0_NB As Integer
Public arrYSWIBIC0_NBMax As Integer
Public arrYSWIBIC0_Index As Integer
Public arrYSWIBIC0_Suite As Boolean

'-----------------------------------------------------
Function srvYSWIBIC0_Update(recYSWIBIC0 As typeYSWIBIC0)
'-----------------------------------------------------

srvYSWIBIC0_Update = "?"

MsgTxtLen = 0
Call srvYSWIBIC0_PutBuffer(recYSWIBIC0)

If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If Not IsNull(srvYSWIBIC0_GetBuffer(recYSWIBIC0)) Then
        Call srvYSWIBIC0_Error(recYSWIBIC0)
        srvYSWIBIC0_Update = recYSWIBIC0.Err
        Exit Function
    Else
        srvYSWIBIC0_Update = Null
    End If
Else
    recYSWIBIC0.Err = "srv"
End If
End Function

'-----------------------------------------------------
Sub srvYSWIBIC0_Error(recYSWIBIC0 As typeYSWIBIC0)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "YSWIBIC0" & Chr$(10) & Chr$(13)

Select Case mId$(recYSWIBIC0.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recYSWIBIC0.Err
        I = vbCritical
End Select

MsgBox Msg & " : " & recYSWIBIC0.SWIBICBIC _
        , I, "module : YSWIBIC0s.bas  ( " & Trim(recYSWIBIC0.obj) & " : " & Trim(recYSWIBIC0.Method) & " )"

End Sub


'=====================================================



'-----------------------------------------------------
Public Function srvYSWIBIC0_Monitor(recYSWIBIC0 As typeYSWIBIC0)
'-----------------------------------------------------

arrYSWIBIC0_Suite = False
Select Case mId$(Trim(recYSWIBIC0.Method), 1, 4)
    Case "Snap"
              srvYSWIBIC0_Monitor = srvYSWIBIC0_Snap(recYSWIBIC0)
    Case Else
            srvYSWIBIC0_Monitor = srvYSWIBIC0_Seek(recYSWIBIC0)
End Select

End Function

'---------------------------------------------------------
Public Function srvYSWIBIC0_GetBuffer(recYSWIBIC0 As typeYSWIBIC0)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvYSWIBIC0_GetBuffer = Null
recYSWIBIC0.obj = mId$(MsgTxt, MsgTxtIndex + 1, 12)
recYSWIBIC0.Method = mId$(MsgTxt, MsgTxtIndex + 13, 12)
recYSWIBIC0.Err = mId$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recYSWIBIC0.Err = Space$(10) Then
    recYSWIBIC0.SWIBICBIC = mId$(MsgTxt, K + 1, 11)
    recYSWIBIC0.SWIBICINT = mId$(MsgTxt, K + 12, 105)
    recYSWIBIC0.SWIBICVIL = mId$(MsgTxt, K + 117, 35)
    recYSWIBIC0.SWIBICCOM = mId$(MsgTxt, K + 152, 70)

Else
    srvYSWIBIC0_GetBuffer = recYSWIBIC0.Err
    recYSWIBIC0.SWIBICINT = "??? " & recYSWIBIC0.SWIBICBIC
    recYSWIBIC0.SWIBICVIL = ""
    recYSWIBIC0.SWIBICCOM = ""

End If

MsgTxtIndex = MsgTxtIndex + recYSWIBIC0Len

End Function

'---------------------------------------------------------
Public Function srvYSWIBIC0_GetBuffer_ODBC(rsADO As ADODB.Recordset, recYSWIBIC0 As typeYSWIBIC0)
'---------------------------------------------------------
On Error GoTo Error_Handler
srvYSWIBIC0_GetBuffer_ODBC = Null

    recYSWIBIC0.SWIBICBIC = rsADO("SWIBICBIC")
    recYSWIBIC0.SWIBICINT = rsADO("SWIBICINT")
    recYSWIBIC0.SWIBICVIL = rsADO("SWIBICVIL")
    recYSWIBIC0.SWIBICCOM = rsADO("SWIBICCOM")

Exit Function

Error_Handler:
srvYSWIBIC0_GetBuffer_ODBC = Error

End Function


'---------------------------------------------------------
Public Sub srvYSWIBIC0_PutBuffer(recYSWIBIC0 As typeYSWIBIC0)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recYSWIBIC0.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recYSWIBIC0.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34

    Mid$(MsgTxt, K + 1, 11) = recYSWIBIC0.SWIBICBIC
    Mid$(MsgTxt, K + 12, 105) = recYSWIBIC0.SWIBICINT
    Mid$(MsgTxt, K + 117, 35) = recYSWIBIC0.SWIBICVIL
    Mid$(MsgTxt, K + 152, 70) = recYSWIBIC0.SWIBICCOM

MsgTxtLen = MsgTxtLen + recYSWIBIC0Len
End Sub



'---------------------------------------------------------
Private Function srvYSWIBIC0_Seek(recYSWIBIC0 As typeYSWIBIC0)
'---------------------------------------------------------

srvYSWIBIC0_Seek = "?"
MsgTxtLen = 0
Call srvYSWIBIC0_PutBuffer(recYSWIBIC0)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If IsNull(srvYSWIBIC0_GetBuffer(recYSWIBIC0)) Then
        srvYSWIBIC0_Seek = Null
    Else
        ''Call srvYSWIBIC0_Error(recYSWIBIC0)
    End If
End If

End Function

'---------------------------------------------------------
Private Function srvYSWIBIC0_Snap(recYSWIBIC0 As typeYSWIBIC0)
'---------------------------------------------------------
srvYSWIBIC0_Snap = "?"
MsgTxtLen = 0
Call srvYSWIBIC0_PutBuffer(recYSWIBIC0)
Call srvYSWIBIC0_PutBuffer(arrYSWIBIC0(0))
If IsNull(SndRcv()) Then
    srvYSWIBIC0_Snap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(srvYSWIBIC0_GetBuffer(recYSWIBIC0)) Then
            Call arrYSWIBIC0_AddItem(recYSWIBIC0)
            arrYSWIBIC0_Suite = True
        Else
            arrYSWIBIC0_Suite = False
            Exit Do
        End If
    Loop
End If

End Function
'---------------------------------------------------------
Public Sub arrYSWIBIC0_AddItem(recYSWIBIC0 As typeYSWIBIC0)
'---------------------------------------------------------
          
arrYSWIBIC0_NB = arrYSWIBIC0_NB + 1
    
If arrYSWIBIC0_NB > arrYSWIBIC0_NBMax Then
    arrYSWIBIC0_NBMax = arrYSWIBIC0_NBMax + recYSWIBIC0_Block
    ReDim Preserve arrYSWIBIC0(arrYSWIBIC0_NBMax)
End If
            
arrYSWIBIC0(arrYSWIBIC0_NB) = recYSWIBIC0
End Sub



'---------------------------------------------------------
Public Sub recYSWIBIC0_Init(recYSWIBIC0 As typeYSWIBIC0)
'---------------------------------------------------------
recYSWIBIC0.obj = "ZSWIBIC0_S"
recYSWIBIC0.Method = ""
recYSWIBIC0.Err = ""
recYSWIBIC0.SWIBICBIC = ""
recYSWIBIC0.SWIBICINT = ""
recYSWIBIC0.SWIBICVIL = ""
recYSWIBIC0.SWIBICCOM = ""
End Sub





