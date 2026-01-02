Attribute VB_Name = "rsYBIARELV"
'---------------------------------------------------------
Option Explicit
Type typeYBIARELV
    BIARELCOM       As String * 20                    ' NUMERO COMPTE
    BIARELREL       As String * 1                     '
    BIARELID        As Long                           '
    BIARELNUM       As Long                           '
    BIARELSD0       As Currency                   '
    BIARELD0        As String * 8                    '
    BIARELSD1       As Currency                       '
    BIARELD1        As String * 8                    '
    BIAOLDCOM       As String * 11                    '
    BIAOLDDEV       As String * 3                    '
End Type

'---------------------------------------------------------
Public Sub rsYBIARELV_Init(rsYBIARELV As typeYBIARELV)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsYBIARELV.BIARELCOM = "" '      As String * 20                    ' NUMERO COMPTE
rsYBIARELV.BIARELREL = "" '       As String * 1                     '
rsYBIARELV.BIARELID = 0   '      As Long                           '
rsYBIARELV.BIARELNUM = 0  '       As Long                           '
rsYBIARELV.BIARELSD0 = 0   '      As Currency                   '
rsYBIARELV.BIARELD0 = ""  '       As String * 8                    '
rsYBIARELV.BIARELSD1 = 0  '       As Currency                       '
rsYBIARELV.BIARELD1 = ""  '       As String * 8                    '
rsYBIARELV.BIAOLDCOM = "" '       As String * 11                    '
rsYBIARELV.BIAOLDDEV = "" '       As String * 3                    '

   
Exit Sub

Error_Handler:


End Sub

'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsYBIARELV_GetBuffer(rsAdo As ADODB.Recordset, rsYBIARELV As typeYBIARELV)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsYBIARELV_GetBuffer = Null

rsYBIARELV.BIARELCOM = rsAdo("BIARELCOM")
rsYBIARELV.BIARELREL = rsAdo("BIARELREL")
rsYBIARELV.BIARELID = rsAdo("BIARELID")
rsYBIARELV.BIARELNUM = rsAdo("BIARELNUM")
rsYBIARELV.BIARELSD0 = rsAdo("BIARELSD0")
rsYBIARELV.BIARELD0 = rsAdo("BIARELD0")
rsYBIARELV.BIARELSD1 = rsAdo("BIARELSD1")
rsYBIARELV.BIARELD1 = rsAdo("BIARELD1")
rsYBIARELV.BIAOLDCOM = rsAdo("BIAOLDCOM")
rsYBIARELV.BIAOLDDEV = rsAdo("BIAOLDDEV")

Exit Function

Error_Handler:

rsYBIARELV_GetBuffer = Error

End Function


'








