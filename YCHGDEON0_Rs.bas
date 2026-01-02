Attribute VB_Name = "rsYCHGDEON0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim rsAdo As ADODB.Recordset
 
Type typeYCHQDEON0
 
      CHQDEONNUM   As String * 7  'numéro
      CHQDEONZIB   As String * 12  'interbancaire
      CHQDEONZIN   As String * 12  'compte
      CHQDEONAMJ   As Long         'date maj

End Type
Public xYCHQDEON0 As typeYCHQDEON0
Public Function rsYCHQDEON0_GetBuffer_ODBC(rsAdo As ADODB.Recordset, lYCHQDEON0 As typeYCHQDEON0)
On Error GoTo Error_Handler
rsYCHQDEON0_GetBuffer_ODBC = Null
lYCHQDEON0.CHQDEONNUM = rsAdo("CHQDEONNUM")
lYCHQDEON0.CHQDEONZIB = rsAdo("CHQDEONZIB")
lYCHQDEON0.CHQDEONZIN = rsAdo("CHQDEONZIN")
lYCHQDEON0.CHQDEONAMJ = rsAdo("CHQDEONAMJ")

Exit Function
Error_Handler:
rsYCHQDEON0_GetBuffer_ODBC = Error


End Function

Public Function rsYCHQDEON0_Init(lYCHQDEON0 As typeYCHQDEON0)
lYCHQDEON0.CHQDEONAMJ = 0
lYCHQDEON0.CHQDEONNUM = ""
lYCHQDEON0.CHQDEONZIB = ""
lYCHQDEON0.CHQDEONZIN = ""

End Function


