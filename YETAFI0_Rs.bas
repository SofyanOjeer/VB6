Attribute VB_Name = "rsYETAFI0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Type typeYETAFI0
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
Public Function rsYETAFI0_GetBuffer(rsAdo As ADODB.Recordset, rsYETAFI0 As typeYETAFI0)
'---------------------------------------------------------
On Error Resume Next 'GoTo Error_Handler
rsYETAFI0_GetBuffer = Null

rsYETAFI0.ETAFICOM = rsAdo("ETAFICOM")
rsYETAFI0.ETAFIOBL = rsAdo("ETAFIOBL")
rsYETAFI0.ETAFIINT = rsAdo("ETAFIINT")
rsYETAFI0.ETAFISD0X = rsAdo("ETAFISD0X")
rsYETAFI0.ETAFIDBX = rsAdo("ETAFIDBX")
rsYETAFI0.ETAFICRX = rsAdo("ETAFICRX")
rsYETAFI0.ETAFISD1X = rsAdo("ETAFISD1X")
rsYETAFI0.ETAFISD0 = rsAdo("ETAFISD0")
rsYETAFI0.ETAFIDB = rsAdo("ETAFIDB")
rsYETAFI0.ETAFICR = rsAdo("ETAFICR")
rsYETAFI0.ETAFISD1 = rsAdo("ETAFISD1")
rsYETAFI0.ETAFIDBNB = rsAdo("ETAFIDBNB")
rsYETAFI0.ETAFICRNB = rsAdo("ETAFICRNB")
rsYETAFI0.ETAFIDEV = rsAdo("ETAFIDEV")
rsYETAFI0.ETAFISTA = rsAdo("ETAFISTA")

Exit Function

Error_Handler:
rsYETAFI0_GetBuffer = Error

End Function



