Attribute VB_Name = "rsZBAGFAC0"
Option Explicit

Public Const constZBAGFAC0 = "ZBAGFAC0"

Type typeZBAGFAC0
    BAGFACDOP   As Long
    BAGFACDVA   As Long
    BAGFACMCO   As Currency
    BAGFACNAT   As String * 3
    BAGFACLI1   As String * 30
    BAGFACLI2   As String * 30
    BAGFACCPT   As String * 20
End Type

Public Function rsZBAGFAC0_GetBuffer(rsAdo As ADODB.Recordset, rsZBAGFAC0 As typeZBAGFAC0)
    
    On Error GoTo Error_Handler
    rsZBAGFAC0_GetBuffer = Null
    rsZBAGFAC0.BAGFACDOP = Val(rsAdo("BAGFACDOP"))
    rsZBAGFAC0.BAGFACDVA = Val(rsAdo("BAGFACDVA"))
    rsZBAGFAC0.BAGFACMCO = Val(rsAdo("BAGFACMCO"))
    rsZBAGFAC0.BAGFACNAT = rsAdo("BAGFACNAT")
    rsZBAGFAC0.BAGFACLI1 = rsAdo("BAGFACLI1")
    rsZBAGFAC0.BAGFACLI2 = rsAdo("BAGFACLI2")
    rsZBAGFAC0.BAGFACCPT = rsAdo("BAGFACCPT")

Exit Function

Error_Handler:

    rsZBAGFAC0_GetBuffer = Error

End Function

Public Sub rsZBAGFAC0_Init(rsZBAGFAC0 As typeZBAGFAC0)

    rsZBAGFAC0.BAGFACDOP = 0
    rsZBAGFAC0.BAGFACDVA = 0
    rsZBAGFAC0.BAGFACMCO = 0
    rsZBAGFAC0.BAGFACNAT = ""
    rsZBAGFAC0.BAGFACLI1 = ""
    rsZBAGFAC0.BAGFACLI2 = ""
    rsZBAGFAC0.BAGFACCPT = ""

End Sub

