Attribute VB_Name = "adoQRY_ACCFIC"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsQRY_ACCFIC_PutBuffer(rsado As ADODB.Recordset, rsQRY_ACCFIC As typeQRY_ACCFIC)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsQRY_ACCFIC_PutBuffer = Null
rsado("COD_UTI") = rsQRY_ACCFIC.COD_UTI
rsado("D_UTIPRE") = rsQRY_ACCFIC.D_UTIPRE
rsado("CDOUTICOP") = rsQRY_ACCFIC.CDOUTICOP
rsado("NO_UTIDOS") = rsQRY_ACCFIC.NO_UTIDOS
rsado("NO_UTIUTI") = rsQRY_ACCFIC.NO_UTIUTI
rsado("D_UTIDRE") = rsQRY_ACCFIC.D_UTIDRE
rsado("CDOUTITMO") = rsQRY_ACCFIC.CDOUTITMO
rsado("MNT_UTI") = rsQRY_ACCFIC.MNT_UTI
rsado("COD_DEV") = rsQRY_ACCFIC.COD_DEV
rsado("D_DOSVAL") = rsQRY_ACCFIC.D_DOSVAL
rsado("NO_BQUE") = rsQRY_ACCFIC.NO_BQUE
    
Exit Function

Error_Handler:

rsQRY_ACCFIC_PutBuffer = Error

End Function


'---------------------------------------------------------
Public Function adoQRY_ACCFIC_AddNew(rsado As ADODB.Recordset, rsQRY_ACCFIC As typeQRY_ACCFIC)
'---------------------------------------------------------
On Error GoTo Error_Handler

adoQRY_ACCFIC_AddNew = Null
rsado.AddNew
adoQRY_ACCFIC_AddNew = rsQRY_ACCFIC_PutBuffer(rsado, rsQRY_ACCFIC)
rsado.Update

Exit Function

Error_Handler:

adoQRY_ACCFIC_AddNew = Error

End Function



