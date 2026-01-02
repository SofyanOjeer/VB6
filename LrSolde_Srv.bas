Attribute VB_Name = "srvLrSolde"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recLrSoldeLen = 166 ' 34 + 132

Type typeLrSolde
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    Text                    As String * 132
    
  End Type
    
Public arrLrSoldeSuite As Boolean
Public arrLrSoldeNb As Integer
'-----------------------------------------------------
Public Function Monitor(recLrSolde As typeLrSolde)
'-----------------------------------------------------

arrLrSoldeSuite = False
Select Case Mid$(Trim(recLrSolde.Method), 1, 4)
    Case "Snap"
              Monitor = Snap(recLrSolde)
    Case Else
                recLrSolde.Err = recLrSolde.Method
                Call ErrorX(recLrSolde)
                Monitor = recLrSolde.Err
End Select

End Function

'-----------------------------------------------------
Sub ErrorX(recLrSolde As typeLrSolde)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "LrSolde: "

Select Case Mid$(recLrSolde.Err, 9, 2)
    Case "22"
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recLrSolde.Err
        I = vbCritical
End Select

MsgBox Msg, I, "module : srvLrSolde.bas  ( " _
                & Trim(recLrSolde.obj) & " : " & Trim(recLrSolde.Method) & " )"

End Sub



'---------------------------------------------------------
Public Function GetBuffer(recLrSolde As typeLrSolde)
'---------------------------------------------------------
Dim K As Integer, I As Integer
GetBuffer = Null
recLrSolde.obj = Mid$(MsgTxt, MsgTxtIndex + 1, 12)
recLrSolde.Method = Mid$(MsgTxt, MsgTxtIndex + 13, 12)
recLrSolde.Err = Mid$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recLrSolde.Err = Space$(10) Then
    recLrSolde.Text = Mid$(MsgTxt, K + 1, 132)
Else
    GetBuffer = recLrSolde.Err
End If

MsgTxtIndex = MsgTxtIndex + recLrSoldeLen

End Function

'---------------------------------------------------------
Private Sub PutBuffer(recLrSolde As typeLrSolde)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recLrSolde.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recLrSolde.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34

Mid$(MsgTxt, K + 1, 132) = recLrSolde.Text
MsgTxtLen = MsgTxtLen + recLrSoldeLen
End Sub



'---------------------------------------------------------
Private Function Snap(recLrSolde As typeLrSolde)
'---------------------------------------------------------
Dim I As Integer
Snap = "?"
MsgTxtLen = 0
Call PutBuffer(recLrSolde)
'Call PutBuffer(arrLrSolde(0))
If IsNull(SndRcv()) Then
    Snap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(GetBuffer(recLrSolde)) Then
            arrLrSoldeNb = arrLrSoldeNb + 1
            Print #1, recLrSolde.Text
            arrLrSoldeSuite = True
        Else
            arrLrSoldeSuite = False
            Exit Do
        End If
    Loop
End If

End Function

'---------------------------------------------------------
Public Sub Init(recLrSolde As typeLrSolde)
'---------------------------------------------------------
MsgTxt = Space$(recLrSoldeLen)
MsgTxtIndex = 0
Call GetBuffer(recLrSolde)
recLrSolde.obj = "SRVLRSOLDE"
End Sub



