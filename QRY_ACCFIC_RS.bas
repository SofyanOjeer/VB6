Attribute VB_Name = "rsQRY_ACCFIC"
'---------------------------------------------------------
Option Explicit
Type typeQRY_ACCFIC
    COD_UTI       As String * 10
    D_UTIPRE        As Long
    CDOUTICOP       As String * 3
    NO_UTIDOS       As Long
    NO_UTIUTI       As Long
    D_UTIDRE        As Long
    CDOUTITMO       As String * 1
    MNT_UTI         As Currency
    COD_DEV       As String * 3
    D_DOSVAL        As Long
    NO_BQUE         As String * 7
   
End Type


'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsQRY_ACCFIC_GetBuffer(rsSab As ADODB.Recordset, rsQRY_ACCFIC As typeQRY_ACCFIC)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsQRY_ACCFIC_GetBuffer = Null

rsQRY_ACCFIC.COD_UTI = rsSab("COD_UTI")
rsQRY_ACCFIC.D_UTIPRE = rsSab("D_UTIPRE")
rsQRY_ACCFIC.CDOUTICOP = rsSab("CDOUTICOP")
rsQRY_ACCFIC.NO_UTIDOS = rsSab("NO_UTIDOS")
rsQRY_ACCFIC.NO_UTIUTI = rsSab("NO_UTIUTI")
rsQRY_ACCFIC.D_UTIDRE = rsSab("D_UTIDRE")
rsQRY_ACCFIC.CDOUTITMO = rsSab("CDOUTITMO")
rsQRY_ACCFIC.MNT_UTI = rsSab("MNT_UTI")
rsQRY_ACCFIC.COD_DEV = rsSab("COD_DEV")
rsQRY_ACCFIC.D_DOSVAL = rsSab("D_DOSVAL")
rsQRY_ACCFIC.NO_BQUE = rsSab("NO_BQUE")

Exit Function

Error_Handler:

rsQRY_ACCFIC_GetBuffer = Error

End Function



'











