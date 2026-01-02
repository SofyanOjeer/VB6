Attribute VB_Name = "rsZCLIGRP0"
'---------------------------------------------------------
Option Explicit
Type typeZCLIGRP0
      CLIGRPETB     As Long
      CLIGRPCLI     As String * 7
      CLIGRPREG     As String * 7
      CLIGRPREL     As String * 3
      CLIGRPCOM     As String * 28
      CLIGRPAUT     As String * 1
      CLIGRPRAT     As String * 1
      CLIGRPTAU     As Double
      CLIGRPPAR     As Long
      
      CLIGRPCLI_RA1 As String
      CLIGRPREG_RA1 As String


End Type

Public Sub rsZCLIGRP0_Init(rsZCLIGRP0 As typeZCLIGRP0)
rsZCLIGRP0.CLIGRPETB = 0
rsZCLIGRP0.CLIGRPCLI = ""
rsZCLIGRP0.CLIGRPREG = ""
rsZCLIGRP0.CLIGRPREL = ""
rsZCLIGRP0.CLIGRPCOM = ""
rsZCLIGRP0.CLIGRPAUT = ""
rsZCLIGRP0.CLIGRPRAT = ""
rsZCLIGRP0.CLIGRPTAU = 0
rsZCLIGRP0.CLIGRPPAR = 0

rsZCLIGRP0.CLIGRPREG_RA1 = ""
rsZCLIGRP0.CLIGRPCLI_RA1 = ""
End Sub

'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsZCLIGRP0_GetBuffer(rsADO As ADODB.Recordset, rsZCLIGRP0 As typeZCLIGRP0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsZCLIGRP0_GetBuffer = Null

rsZCLIGRP0.CLIGRPETB = rsADO("CLIGRPETB")
rsZCLIGRP0.CLIGRPCLI = rsADO("CLIGRPCLI")
rsZCLIGRP0.CLIGRPREG = rsADO("CLIGRPREG")
rsZCLIGRP0.CLIGRPREL = rsADO("CLIGRPREL")
rsZCLIGRP0.CLIGRPCOM = rsADO("CLIGRPCOM")
rsZCLIGRP0.CLIGRPAUT = rsADO("CLIGRPAUT")
rsZCLIGRP0.CLIGRPRAT = rsADO("CLIGRPRAT")
rsZCLIGRP0.CLIGRPTAU = rsADO("CLIGRPTAU")
rsZCLIGRP0.CLIGRPPAR = rsADO("CLIGRPPAR")

Exit Function

Error_Handler:

rsZCLIGRP0_GetBuffer = Error

End Function










