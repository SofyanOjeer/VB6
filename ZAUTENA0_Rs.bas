Attribute VB_Name = "rsZAUTENA0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim rsADO As ADODB.Recordset
 
Type typeZAUTENA0
    AUTENACLI   As String
    AUTENAAUT   As String
    AUTENADEV   As String
    AUTENAENC   As Currency
    AUTENAOPE   As String
    AUTENADOS   As Long
    
    DOSSLDPCI  As String
    DOSSLDSTA   As String
    DOSSLDMSD   As Currency
End Type
'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsZAUTENA0_GetBuffer(rsADO As ADODB.Recordset, rsZAUTENA0 As typeZAUTENA0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsZAUTENA0_GetBuffer = Null

rsZAUTENA0.AUTENACLI = rsADO("AUTENACLI")
rsZAUTENA0.AUTENAAUT = Trim(rsADO("AUTENAAUT"))
rsZAUTENA0.AUTENADEV = rsADO("AUTENADEV")
rsZAUTENA0.AUTENAENC = rsADO("AUTENAENC")

rsZAUTENA0.AUTENAOPE = rsADO("AUTENAOPE")
rsZAUTENA0.AUTENADOS = rsADO("AUTENADOS")

Exit Function

Error_Handler:

rsZAUTENA0_GetBuffer = Error

End Function









Public Sub rsZAUTENA0_Init(lZAUTENA0 As typeZAUTENA0)
lZAUTENA0.AUTENACLI = ""
lZAUTENA0.AUTENAAUT = ""
lZAUTENA0.AUTENADEV = ""
lZAUTENA0.AUTENAENC = 0
lZAUTENA0.AUTENAOPE = ""
lZAUTENA0.AUTENADOS = 0

lZAUTENA0.DOSSLDSTA = ""
lZAUTENA0.DOSSLDPCI = ""
lZAUTENA0.DOSSLDMSD = 0
End Sub



