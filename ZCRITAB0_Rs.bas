Attribute VB_Name = "rsZCRITAB0"
'---------------------------------------------------------
Option Explicit
Type typeZCRITAB0
    CRITABETA       As Integer
    CRITABNUM       As Integer
    CRITABARG       As String * 15
    CRITABDON       As String * 80
    
    
End Type

'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsZCRITAB0_GetBuffer(rsado As ADODB.Recordset, rsZCRITAB0 As typeZCRITAB0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsZCRITAB0_GetBuffer = Null

rsZCRITAB0.CRITABETA = rsado("CRITABETA")
rsZCRITAB0.CRITABNUM = rsado("CRITABNUM")
rsZCRITAB0.CRITABARG = rsado("CRITABARG")
rsZCRITAB0.CRITABDON = rsado("CRITABDON")

Exit Function

Error_Handler:

rsZCRITAB0_GetBuffer = Error

End Function


'







