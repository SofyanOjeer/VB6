Attribute VB_Name = "srvYSWIHIB0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------

Type typeJSWIHIB0
 
        SWIHIBETA               As Integer                        ' ETABLISSEMENT
        SWIHIBNUM               As Long                           ' NUMERO INTERNE
        SWIHIBNEN               As Long                           ' NUMERO ENVOI
        SWIHIBNLI               As Long                           ' NUMERO LIGNE
        SWIHIBDET               As String * 70                    ' DETAIL

End Type
Public xJSWIHIB0 As typeJSWIHIB0

Public Function srvJSWIHIB0_GetBuffer_ODBC(rsADO As ADODB.Recordset, lJSWIHIB0 As typeJSWIHIB0)
On Error GoTo Error_Handler
srvJSWIHIB0_GetBuffer_ODBC = Null
lJSWIHIB0.SWIHIBETA = rsADO("SWIHIBETA")
lJSWIHIB0.SWIHIBNUM = rsADO("SWIHIBNUM")
lJSWIHIB0.SWIHIBNEN = rsADO("SWIHIBNEN")
lJSWIHIB0.SWIHIBNLI = rsADO("SWIHIBNLI")
lJSWIHIB0.SWIHIBDET = rsADO("SWIHIBDET")
Exit Function
Error_Handler:
srvJSWIHIB0_GetBuffer_ODBC = Error
End Function

 Public Function srvJSWIHIB0_Sql_Insert(lJSWIHIB0 As typeJSWIHIB0, lSql As String)
On Error GoTo Error_Handler
srvJSWIHIB0_Sql_Insert = Null
lSql = "Insert into " & paramIBM_Library_SAB & ".ZSWIHIB0 " _
        & "(SWIHIBETA,SWIHIBNUM,SWIHIBNEN,SWIHIBNLI,SWIHIBDET) value( " _
        & lJSWIHIB0.SWIHIBETA _
        & "," & lJSWIHIB0.SWIHIBNUM _
        & "," & lJSWIHIB0.SWIHIBNEN _
        & "," & lJSWIHIB0.SWIHIBNLI _
        & ",'" & lJSWIHIB0.SWIHIBETA & "')"
        
Exit Function
Error_Handler:
srvJSWIHIB0_Sql_Insert = Error
End Function



