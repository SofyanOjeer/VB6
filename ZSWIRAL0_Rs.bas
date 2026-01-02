Attribute VB_Name = "rsZSWIRAL0"
Option Explicit

Type typeZSWIRAL0
    SWIRALDON       As String * 512                   ' DONNE MESSAGE
    SWIRALETA       As Integer                        '
    SWIRALMES       As String * 3                     '
End Type

'---------------------------------------------------------
Public Function rsZSWIRAL0_GetBuffer(rsado As ADODB.Recordset, rsZSWIRAL0 As typeZSWIRAL0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsZSWIRAL0_GetBuffer = Null

rsZSWIRAL0.SWIRALDON = rsado("SWIALLDON")
rsZSWIRAL0.SWIRALETA = rsado("SWIALLETA")               '
rsZSWIRAL0.SWIRALMES = rsado("SWIALLMES")         '
Exit Function

Error_Handler:

rsZSWIRAL0_GetBuffer = Error

End Function


'---------------------------------------------------------
Public Sub rsZSWIRAL0_Init(rsZSWIRAL0 As typeZSWIRAL0)
'---------------------------------------------------------
rsZSWIRAL0.SWIRALDON = ""
rsZSWIRAL0.SWIRALETA = 0                  '
rsZSWIRAL0.SWIRALMES = ""               '
End Sub


