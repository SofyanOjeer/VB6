Attribute VB_Name = "srvDDEVISE"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim rsADO As ADODB.Recordset
 
Type typeDDEVISE
 
      DDEVDEV     As String * 3    ' CODE DEVISE ALPHA
      DDEVDEN     As String * 3    ' CODE DEVISE NUMERIQUE
      DDEVLIB     As String * 12   ' LIBELLE
      DDEVDEC     As String * 1    ' NOMBRE DECIMAL
      DDEVPCRT    As Long          ' DATE CREATION

End Type
Public xDDEVISE As typeDDEVISE
Public Function srvDDEVISE_GetBuffer_ODBC(rsADO As ADODB.Recordset, lDDEVISE As typeDDEVISE)

On Error GoTo Error_Handler

srvDDEVISE_GetBuffer_ODBC = Null

lDDEVISE.DDEVDEV = rsADO("DDEVDEV")
lDDEVISE.DDEVDEN = rsADO("DDEVDEN")
lDDEVISE.DDEVLIB = rsADO("DDEVLIB")
lDDEVISE.DDEVDEC = rsADO("DDEVDEC")
lDDEVISE.DDEVPCRT = rsADO("DDEVPCRT")

Exit Function
Error_Handler:
srvDDEVISE_GetBuffer_ODBC = Error

End Function

