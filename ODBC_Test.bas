Attribute VB_Name = "ODBC_Test"
Option Explicit
Dim rsADO As New ADODB.Recordset

Public tableODBC As Recordset

Type typetableODBC
    
    BIATABID        As String * 12
    BIATABK1        As String * 12
    BIATABK2        As String * 12
    BIATABTEXT      As String * 128
End Type


Public Sub ODBC_Open()
Dim X As String
Dim xSQL As String
On Error GoTo fin
xSQL = "select * from ZCDODOS0" ' where BASFUTCLI='0012375'"
rsADO.Open xSQL, "DSN=BIADWH"
Do While Not rsADO.EOF
   ' X = Space$(43)
   ' Mid$(X, 1, 12) = rsADO("CDODOSDOS")
   ' Mid$(X, 13, 24) = rsADO("CDODOSMON")
   Debug.Print rsADO("CDODOSDOS"); rsADO("CDODOSMON")
    rsADO.MoveNext
Loop
Close
Exit Sub
fin:
MsgBox Error

End Sub
