Attribute VB_Name = "srvYETAFI0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Type typeYETAFI0
    Obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    ETAFICOM       As String * 20
    ETAFIOBL       As String * 10
    ETAFIINT       As String * 32
    ETAFISD0X      As Currency
    ETAFIDBX       As Currency
    ETAFICRX       As Currency
    ETAFISD1X      As Currency
    ETAFISD0       As Currency
    ETAFIDB        As Currency
    ETAFICR        As Currency
    ETAFISD1       As Currency
    ETAFIDBNB      As Long
    ETAFICRNB      As Long
    ETAFIDEV       As String * 3
    ETAFISTA       As String * 3

End Type

'---------------------------------------------------------
Public Function srvYETAFI0_GetBuffer_ODBC(rsADO As ADODB.Recordset, recYETAFI0 As typeYETAFI0)
'---------------------------------------------------------
On Error Resume Next 'GoTo Error_Handler
srvYETAFI0_GetBuffer_ODBC = Null

recYETAFI0.ETAFICOM = rsADO("ETAFICOM")
recYETAFI0.ETAFIOBL = rsADO("ETAFIOBL")
recYETAFI0.ETAFIINT = rsADO("ETAFIINT")
recYETAFI0.ETAFISD0X = rsADO("ETAFISD0X")
recYETAFI0.ETAFIDBX = rsADO("ETAFIDBX")
recYETAFI0.ETAFICRX = rsADO("ETAFICRX")
recYETAFI0.ETAFISD1X = rsADO("ETAFISD1X")
recYETAFI0.ETAFISD0 = rsADO("ETAFISD0")
recYETAFI0.ETAFIDB = rsADO("ETAFIDB")
recYETAFI0.ETAFICR = rsADO("ETAFICR")
recYETAFI0.ETAFISD1 = rsADO("ETAFISD1")
recYETAFI0.ETAFIDBNB = rsADO("ETAFIDBNB")
recYETAFI0.ETAFICRNB = rsADO("ETAFICRNB")
recYETAFI0.ETAFIDEV = rsADO("ETAFIDEV")
recYETAFI0.ETAFISTA = rsADO("ETAFISTA")

Exit Function

Error_Handler:
srvYETAFI0_GetBuffer_ODBC = Error

End Function



