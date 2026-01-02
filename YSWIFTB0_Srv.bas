Attribute VB_Name = "srvYSWIFTB0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------

Type typeJSWIFTB0
 
        SWIFTBETA               As Integer                        ' ETABLISSEMENT
        SWIFTBNUM               As Long                           ' NUMERO INTERNE
        SWIFTBNEN               As Long                           ' NUMERO ENVOI
        SWIFTBNLI               As Long                           ' NUMERO LIGNE
        SWIFTBDET               As String * 70                    ' DETAIL

End Type
Public xJSWIFTB0 As typeJSWIFTB0

Public Function srJSWIFTB0_GetBuffer_ODBC(rsADO As ADODB.Recordset, lJSWIFTB0 As typeJSWIFTB0)
On Error GoTo Error_Handler
srvYSWIFTB0_GetBuffer_ODBC = Null
lJSWIFTB0.SWIFTBETA = rsADO("SWIFTBETA")
lJSWIFTB0.SWIFTBNUM = rsADO("SWIFTBNUM")
lJSWIFTB0.SWIFTBNEN = rsADO("SWIFTBNEN")
lJSWIFTB0.SWIFTBNLI = rsADO("SWIFTBNLI")
lJSWIFTB0.SWIFTBDET = rsADO("SWIFTBDET")
Exit Function
Error_Handler:
srvYSWIFTB0_GetBuffer_ODBC = Error '
End Function

 Public Function srvYSWIFTB0_Sql_Insert(lJSWIFTB0 As typeJSWIFTB0, lSql As String)
On Error GoTo Error_Handler
srvYSWIFTB0_Sql_Insert = Null
lSql = "Insert into " & paramIBM_Library_SAB & ".ZSWIFTB0 " _
        & "(SWIFTBETA,SWIFTBNUM,SWIFTBNEN,SWIFTBNLI,SWIFTBDET) value( " _
        & lJSWIFTB0.SWIFTBETA _
        & "," & lJSWIFTB0.SWIFTBNUM _
        & "," & lJSWIFTB0.SWIFTBNEN _
        & "," & lJSWIFTB0.SWIFTBNLI _
        & ",'" & lJSWIFTB0.SWIFTBETA & "')"
        
Exit Function
Error_Handler:
srvYSWIFTB0_Sql_Insert = Error
End Function





