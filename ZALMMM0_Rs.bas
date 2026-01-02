Attribute VB_Name = "rsZALMMM0"
'---------------------------------------------------------
Option Explicit
Type typeZALMMM0

    ALMMMREC       As String * 2                      '
    ALMMMDAT       As String * 224                    '
    ALMMMNBR       As Long                            '


End Type

'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsZALMMM0_GetBuffer(rsSab As ADODB.Recordset, rsZALMMM0 As typeZALMMM0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsZALMMM0_GetBuffer = Null

rsZALMMM0.ALMMMREC = rsSab("ALMMMREC")
rsZALMMM0.ALMMMDAT = rsSab("ALMMMDAT")
rsZALMMM0.ALMMMNBR = rsSab("ALMMMNBR")

Exit Function

Error_Handler:

rsZALMMM0_GetBuffer = Error

End Function

'---------------------------------------------------------
Public Sub rsZALMMM0_Init(rsZALMMM0 As typeZALMMM0)
'---------------------------------------------------------
rsZALMMM0.ALMMMREC = ""
rsZALMMM0.ALMMMDAT = ""
rsZALMMM0.ALMMMNBR = 0

End Sub


'









