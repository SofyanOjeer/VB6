Attribute VB_Name = "rsZCLINEX0"
'---------------------------------------------------------
Option Explicit
Type typeZCLINEX0
      CLINEXETB     As Long
      CLINEXCLI     As String * 7
      CLINEXORG     As String * 7
      CLINEXDNO     As Long
      CLINEXDCR     As Long
      CLINEXDRE     As Long
      CLINEXNO1     As String * 6
      CLINEXNO2     As String * 6
      CLINEXDSA     As Long
      CLINEXUSR     As Long
      
End Type


Public Sub rsZCLINEX0_Init(rsZCLINEX0 As typeZCLINEX0)
rsZCLINEX0.CLINEXETB = 0
rsZCLINEX0.CLINEXCLI = ""
rsZCLINEX0.CLINEXORG = ""
rsZCLINEX0.CLINEXDNO = ""
rsZCLINEX0.CLINEXDCR = ""
rsZCLINEX0.CLINEXDRE = ""
rsZCLINEX0.CLINEXNO1 = ""
rsZCLINEX0.CLINEXNO2 = 0
rsZCLINEX0.CLINEXDSA = 0
rsZCLINEX0.CLINEXUSR = 0

End Sub

'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsZCLINEX0_GetBuffer(rsADO As ADODB.Recordset, rsZCLINEX0 As typeZCLINEX0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsZCLINEX0_GetBuffer = Null

rsZCLINEX0.CLINEXETB = rsADO("CLINEXETB")
rsZCLINEX0.CLINEXCLI = rsADO("CLINEXCLI")
rsZCLINEX0.CLINEXORG = rsADO("CLINEXORG")
rsZCLINEX0.CLINEXDNO = rsADO("CLINEXDNO")
rsZCLINEX0.CLINEXDCR = rsADO("CLINEXDCR")
rsZCLINEX0.CLINEXDRE = rsADO("CLINEXDRE")
rsZCLINEX0.CLINEXNO1 = rsADO("CLINEXNO1")
rsZCLINEX0.CLINEXNO2 = rsADO("CLINEXNO2")
rsZCLINEX0.CLINEXDSA = rsADO("CLINEXDSA")
rsZCLINEX0.CLINEXUSR = rsADO("CLINEXUSR")

Exit Function

Error_Handler:

rsZCLINEX0_GetBuffer = Error

End Function











