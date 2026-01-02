Attribute VB_Name = "srvYTP7OPH0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim rsADO As ADODB.Recordset
 
Type typeYTP7OPH0
    TP7OPHDEV   As String * 3  ' devise
    TP7OPHCOM   As String * 20   ' compte
    TP7OPHDTR   As Long        ' N° séquence (info)
    TP7OPHOPE   As String      ' utilisateur maj
    TP7OPHDBD   As Currency    ' total débit
    TP7OPHDBN   As Long        ' nb débit
    TP7OPHCRD   As Currency    ' total crédit
    TP7OPHCRN   As Long        ' nb crédit
    
End Type
'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsYTP7OPH0_GetBuffer(rsADO As ADODB.Recordset, rsYTP7OPH0 As typeYTP7OPH0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsYTP7OPH0_GetBuffer = Null

rsYTP7OPH0.TP7OPHDEV = rsADO("TP7OPHDEV")
rsYTP7OPH0.TP7OPHCOM = rsADO("TP7OPHCOM")
rsYTP7OPH0.TP7OPHDTR = rsADO("TP7OPHDTR")
rsYTP7OPH0.TP7OPHOPE = rsADO("TP7OPHOPE")
rsYTP7OPH0.TP7OPHCRD = rsADO("TP7OPHCRD")
rsYTP7OPH0.TP7OPHCRN = rsADO("TP7OPHCRN")
rsYTP7OPH0.TP7OPHDBD = rsADO("TP7OPHDBD")
rsYTP7OPH0.TP7OPHDBN = rsADO("TP7OPHDBN")

Exit Function

Error_Handler:

rsYTP7OPH0_GetBuffer = Error

End Function









Public Sub rsYTP7OPH0_Init(lYTP7OPH0 As typeYTP7OPH0)
lYTP7OPH0.TP7OPHCOM = ""
lYTP7OPH0.TP7OPHDEV = ""
lYTP7OPH0.TP7OPHDTR = 0
lYTP7OPH0.TP7OPHOPE = ""
lYTP7OPH0.TP7OPHCRD = 0
lYTP7OPH0.TP7OPHCRN = 0
lYTP7OPH0.TP7OPHDBD = 0
lYTP7OPH0.TP7OPHDBN = 0

End Sub


