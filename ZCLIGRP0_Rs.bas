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
      CLIGRPCLI_RA2 As String
      CLIGRPREG_RA1 As String
      CLIGRPREG_9X As Boolean
      CLIGRPCLI_RES As String
      CLIGRPCLI_NAT As String
End Type

Public Function libCLIGRPREL(lCLIGRPREL As String) As String
Dim X As String
X = lCLIGRPREL
Select Case lCLIGRPREL
    Case "ADM": X = "Administrateurs"
    Case "DIR": X = "Dirigeants"
    Case "FIL": X = "Filiales"
    Case "GGR": X = "Groupes"
 End Select
libCLIGRPREL = X

End Function

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
rsZCLIGRP0.CLIGRPCLI_RA2 = ""
rsZCLIGRP0.CLIGRPCLI_RES = ""
rsZCLIGRP0.CLIGRPREG_9X = False
End Sub

'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsZCLIGRP0_GetBuffer(rsAdo As ADODB.Recordset, rsZCLIGRP0 As typeZCLIGRP0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsZCLIGRP0_GetBuffer = Null

rsZCLIGRP0.CLIGRPETB = rsAdo("CLIGRPETB")
rsZCLIGRP0.CLIGRPCLI = rsAdo("CLIGRPCLI")
rsZCLIGRP0.CLIGRPREG = rsAdo("CLIGRPREG")
rsZCLIGRP0.CLIGRPREL = rsAdo("CLIGRPREL")
rsZCLIGRP0.CLIGRPCOM = rsAdo("CLIGRPCOM")
rsZCLIGRP0.CLIGRPAUT = rsAdo("CLIGRPAUT")
rsZCLIGRP0.CLIGRPRAT = rsAdo("CLIGRPRAT")
rsZCLIGRP0.CLIGRPTAU = rsAdo("CLIGRPTAU")
rsZCLIGRP0.CLIGRPPAR = rsAdo("CLIGRPPAR")

Exit Function

Error_Handler:

rsZCLIGRP0_GetBuffer = Error

End Function










