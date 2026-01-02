Attribute VB_Name = "rsZWIALL0"
Option Explicit

Type typeZSWIALL0
    
    SWIALLDON       As String * 512                   ' DONNE MESSAGE
End Type

'---------------------------------------------------------
Public Function rsZSWIALL0_GetBuffer(rsado As ADODB.Recordset, rsZSWIALL0 As typeZSWIALL0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsZSWIALL0_GetBuffer = Null

rsZSWIALL0.SWIALLDON = rsado("SWIALLDON")
Exit Function

Error_Handler:

rsZSWIALL0_GetBuffer = Error

End Function


'---------------------------------------------------------
Public Sub rsZSWIALL0_Init(rsZSWIALL0 As typeZSWIALL0)
'---------------------------------------------------------
rsZSWIALL0.SWIALLDON = ""
End Sub

