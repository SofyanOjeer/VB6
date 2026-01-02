Attribute VB_Name = "srvYFLUTP20"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim rsADO As ADODB.Recordset
 
Type typeYFLUTP20
    FLUTP2ID     As Long
    FLUTP2CCB    As Long
    FLUTP2ORIG   As String
    FLUTP2STA  As String
    
    FLUTP2ETB    As Integer
    FLUTP2AGE    As Integer
    FLUTP2SER    As String
    FLUTP2SSE    As String
    FLUTP2OPE   As String
    FLUTP2NAT   As String
    FLUTP2DOS   As Long
    FLUTP2DOSQ  As Long
    FLUTP2EVE    As String
    FLUTP2ECH   As Long
    FLUTP2MTD    As Currency
    FLUTP2DEV    As String
    
    FLUTP2NEG   As Long
    FLUTP2MAD   As Long
    FLUTP2NBJ   As Long
    FLUTP2ECHK   As Integer
    FLUTP2TX    As Double
    FLUTP2TXCB  As Double
    FLUTP2TXK  As String
End Type

Type typeSINFO_LIQU
    MTD     As Currency
    NB      As Long
    Durée   As Double
    Ecart   As Double
    Durée_S   As Double
    Ecart_S   As Double
    NBJ   As Double
End Type

'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsYFLUTP20_GetBuffer(rsADO As ADODB.Recordset, rsYFLUTP20 As typeYFLUTP20)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsYFLUTP20_GetBuffer = Null


rsYFLUTP20.FLUTP2ID = rsADO("FLUTP2ID")
rsYFLUTP20.FLUTP2CCB = rsADO("FLUTP2CCB")
rsYFLUTP20.FLUTP2ORIG = rsADO("FLUTP2ORIG")
rsYFLUTP20.FLUTP2STA = rsADO("FLUTP2STA")
rsYFLUTP20.FLUTP2ETB = rsADO("FLUTP2ETB")
rsYFLUTP20.FLUTP2AGE = rsADO("FLUTP2AGE")
rsYFLUTP20.FLUTP2SER = rsADO("FLUTP2SER")
rsYFLUTP20.FLUTP2SSE = rsADO("FLUTP2SSE")
rsYFLUTP20.FLUTP2OPE = rsADO("FLUTP2OPE")
rsYFLUTP20.FLUTP2NAT = rsADO("FLUTP2NAT")
rsYFLUTP20.FLUTP2DOS = rsADO("FLUTP2DOS")
rsYFLUTP20.FLUTP2DOSQ = rsADO("FLUTP2DOSQ")
rsYFLUTP20.FLUTP2EVE = rsADO("FLUTP2EVE")
rsYFLUTP20.FLUTP2ECH = rsADO("FLUTP2ECH")
rsYFLUTP20.FLUTP2MTD = rsADO("FLUTP2MTD")
rsYFLUTP20.FLUTP2DEV = rsADO("FLUTP2DEV")

rsYFLUTP20.FLUTP2NEG = rsADO("FLUTP2NEG")
rsYFLUTP20.FLUTP2MAD = rsADO("FLUTP2MAD")
rsYFLUTP20.FLUTP2NBJ = rsADO("FLUTP2NBJ")
rsYFLUTP20.FLUTP2ECHK = rsADO("FLUTP2ECHK")
rsYFLUTP20.FLUTP2TX = rsADO("FLUTP2TX")
rsYFLUTP20.FLUTP2TXCB = rsADO("FLUTP2TXCB")
rsYFLUTP20.FLUTP2TXK = rsADO("FLUTP2TXK")

Exit Function

Error_Handler:

rsYFLUTP20_GetBuffer = Error

End Function










Public Sub rsYFLUTP20_Init(lYFLUTP20 As typeYFLUTP20)
lYFLUTP20.FLUTP2ID = 0
lYFLUTP20.FLUTP2CCB = 0
lYFLUTP20.FLUTP2ORIG = ""
lYFLUTP20.FLUTP2STA = ""

lYFLUTP20.FLUTP2ETB = 1
lYFLUTP20.FLUTP2AGE = 1
lYFLUTP20.FLUTP2SER = ""
lYFLUTP20.FLUTP2SSE = ""
lYFLUTP20.FLUTP2OPE = ""
lYFLUTP20.FLUTP2NAT = ""
lYFLUTP20.FLUTP2DOS = 0
lYFLUTP20.FLUTP2DOSQ = 0
lYFLUTP20.FLUTP2EVE = ""
lYFLUTP20.FLUTP2ECH = 0
lYFLUTP20.FLUTP2MTD = 0
lYFLUTP20.FLUTP2DEV = ""

lYFLUTP20.FLUTP2NEG = 0
lYFLUTP20.FLUTP2MAD = 0
lYFLUTP20.FLUTP2NBJ = 0
lYFLUTP20.FLUTP2ECHK = 0
lYFLUTP20.FLUTP2TX = 0
lYFLUTP20.FLUTP2TXCB = 0
lYFLUTP20.FLUTP2TXK = ""

End Sub


Public Sub SINFO_LIQU_Init(lX As typeSINFO_LIQU)
lX.MTD = 0
lX.Durée_S = 0
lX.Ecart_S = 0
lX.Durée = 0
lX.Ecart = 0
lX.NB = 0
lX.NBJ = 0
End Sub

