Attribute VB_Name = "srvDRTAGRP"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim rsADO As ADODB.Recordset
 
Type typeDRTAGRP
 
      DRGRSTA     As String * 1    ' STATUT
      DRGRVER     As Integer       ' No VERSION
      DRGRPER     As Long          ' PERIODE TRAITEMENT
      DRGRCRTA    As Long          ' CODE RENTA
      DRGRCGRP    As Long          ' CODE RENTA REGROUPEMENT
      DRGRLIB     As String * 50   ' LIBELLE

End Type
Public xDRTAGRP As typeDRTAGRP
Public Function srvDRTAGRP_GetBuffer_ODBC(rsADO As ADODB.Recordset, lDRTAGRP As typeDRTAGRP)

On Error GoTo Error_Handler

srvDRTAGRP_GetBuffer_ODBC = Null

lDRTAGRP.DRGRSTA = rsADO("DRGRSTA")
lDRTAGRP.DRGRVER = rsADO("DRGRVER")
lDRTAGRP.DRGRPER = rsADO("DRGRPER")
lDRTAGRP.DRGRCRTA = rsADO("DRGRCRTA")
lDRTAGRP.DRGRCGRP = rsADO("DRGRCGRP")
lDRTAGRP.DRGRLIB = rsADO("DRGRLIB")

Exit Function
Error_Handler:
srvDRTAGRP_GetBuffer_ODBC = Error

End Function

