Attribute VB_Name = "srvDRLCRTA"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim rsADO As ADODB.Recordset
 
Type typeDRLCRTA
      
      DRRTACRTA    As Long          ' CODE RENTA
      DRRTALIB     As String * 50   ' LIBELLE
      DRRTANAT     As String * 1    ' NATURE
      DRRTAPCRT    As Long          ' DATE CREATION

End Type
Public xDRLCRTA As typeDRLCRTA
Public Function srvDRLCRTA_GetBuffer_ODBC(rsADO As ADODB.Recordset, lDRLCRTA As typeDRLCRTA)

On Error GoTo Error_Handler

srvDRLCRTA_GetBuffer_ODBC = Null

lDRLCRTA.DRRTACRTA = rsADO("DRRTACRTA")
lDRLCRTA.DRRTALIB = rsADO("DRRTALIB")
lDRLCRTA.DRRTANAT = rsADO("DRRTANAT")
lDRLCRTA.DRRTAPCRT = rsADO("DRRTAPCRT")

Exit Function
Error_Handler:
srvDRLCRTA_GetBuffer_ODBC = Error

End Function

