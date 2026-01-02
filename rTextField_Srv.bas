Attribute VB_Name = "srvrTextField"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------

Public Const constrTextField = "rTextField"

Type typerTextField
    
    Aid                            As Integer                  ' SAA identification
    text_s_umidl                   As Long                     ' Left part of the internal Swift message reference message)
    text_s_umidh                   As Long                     ' Right part of the internal Swift message reference
    
    field_cnt                      As Long                     '
    field_code                     As Integer                  '
    field_code_id                  As Integer                  '
    field_option                   As String '* 1              '
    value                          As String '* 1750           '
    value_memo                     As String '* 16             '
    sequence_id                    As String '* 1              '
    group_idx                      As Integer                  '
    
End Type


'-------------------------------------------------------------
Public Sub srvrTextField_Init(recrTextField As typerTextField)
'-------------------------------------------------------------

recrTextField.Aid = 0
recrTextField.text_s_umidl = 0
recrTextField.text_s_umidh = 0

recrTextField.field_cnt = 0
recrTextField.field_code = 0
recrTextField.field_code_id = 0
recrTextField.field_option = ""
recrTextField.value = ""
recrTextField.value_memo = ""
recrTextField.sequence_id = ""
recrTextField.group_idx = 0

End Sub


'-------------------------------------------------------------------
Public Sub srvrTextField_ElpDisplay(recrTextField As typerTextField)
'-------------------------------------------------------------------
'frmElpDisplay.fgData.Rows = 12

'frmElpDisplay.fgData.Row = 1
'frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "Aid"
'frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = ""
'frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recrTextField.Aid

'frmElpDisplay.fgData.Row = 2
'frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "text_s_umidl"
'frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = ""
'frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recrTextField.text_s_umidl

'frmElpDisplay.fgData.Row = 3
'frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "text_s_umidh"
'frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = ""
'frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recrTextField.text_s_umidh

'frmElpDisplay.fgData.Row = 4
'frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "field_cnt"
'frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = ""
'frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recrTextField.field_cnt

'frmElpDisplay.fgData.Row = 5
'frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "field_code"
'frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = ""
'frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recrTextField.field_code

'frmElpDisplay.fgData.Row = 6
'frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "field_code_id"
'frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = ""
'frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recrTextField.field_code_id

'frmElpDisplay.fgData.Row = 7
'frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "field_option"
'frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = ""
'frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recrTextField.field_option

'frmElpDisplay.fgData.Row = 8
'frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "value"
'frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = ""
'frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recrTextField.value
'frmElpDisplay.fgData.Row = 9
'frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "value_memo"
'frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = ""
'frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recrTextField.value_memo

'frmElpDisplay.fgData.Row = 10
'frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "sequence_id"
'frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = ""
'frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recrTextField.sequence_id

'frmElpDisplay.fgData.Row = 11
'frmElpDisplay.fgData.Col = 0: frmElpDisplay.fgData = "group_idx"
'frmElpDisplay.fgData.Col = 1: frmElpDisplay.fgData = ""
'frmElpDisplay.fgData.Col = 2: frmElpDisplay.fgData = recrTextField.group_idx

MsgBox "à faire : frmElpDisplay.Show ", vbCritical '  vbModal

End Sub

'------------------------------------------------------------------------------------------------------
Public Function srvrTextField_GetBuffer_ODBC(rsado As ADODB.Recordset, recrTextField As typerTextField)
'------------------------------------------------------------------------------------------------------
On Error Resume Next  'Error_Handler
srvrTextField_GetBuffer_ODBC = Null

recrTextField.Aid = rsado("Aid")
recrTextField.text_s_umidl = rsado("text_s_umidl")
recrTextField.text_s_umidh = rsado("text_s_umidh")

If IsNull(rsado("field_cnt")) Then
    recrTextField.field_cnt = 0
Else
    recrTextField.field_cnt = rsado("field_cnt")
End If
If IsNull(rsado("field_code")) Then
    recrTextField.field_code = 0
Else
    recrTextField.field_code = rsado("field_code")
End If
If IsNull(rsado("field_code_id")) Then
    recrTextField.field_code_id = 0
Else
    recrTextField.field_code_id = rsado("field_code_id")
End If
If IsNull(rsado("field_option")) Then
    recrTextField.field_option = ""
Else
    recrTextField.field_option = rsado("field_option")
End If
If IsNull(rsado("value")) Then
    recrTextField.value = ""
Else
    recrTextField.value = rsado("value")
End If
If IsNull(rsado("value_memo")) Then
    recrTextField.value_memo = ""
Else
    recrTextField.value_memo = rsado("value_memo")
End If
If IsNull(rsado("sequence_id")) Then
    recrTextField.sequence_id = ""
Else
    recrTextField.sequence_id = rsado("sequence_id")
End If
If IsNull(rsado("group_idx")) Then
    recrTextField.group_idx = 0
Else
    recrTextField.group_idx = rsado("group_idx")
End If

Exit Function

Error_Handler:
srvrTextField_GetBuffer_ODBC = Error

End Function




