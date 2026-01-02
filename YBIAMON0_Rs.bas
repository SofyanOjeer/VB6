Attribute VB_Name = "rsYBIAMON0"
'---------------------------------------------------------
Option Explicit
Type typeYBIAMON0
    MONAPP       As String * 10
    MONFLUX      As String * 10
    MONSTATUS    As String * 10
    MONNUM       As Long
    MONJOB       As String * 10
    MONPGM       As String * 10
    MONUSR       As String * 10
    MONAMJ       As Long
    MONHMS       As Long
    MONFILE      As String * 10
    MONUPDS  As Long         'Sequence mise à jour
End Type

'---------------------------------------------------------
Public Sub rsYBIAMON0_Init(rsYBIAMON0 As typeYBIAMON0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsYBIAMON0.MONAPP = ""
rsYBIAMON0.MONFLUX = ""
rsYBIAMON0.MONSTATUS = ""
rsYBIAMON0.MONNUM = 0
rsYBIAMON0.MONJOB = ""
rsYBIAMON0.MONPGM = ""
rsYBIAMON0.MONUSR = ""
rsYBIAMON0.MONAMJ = 0
rsYBIAMON0.MONHMS = 0
rsYBIAMON0.MONFILE = ""
rsYBIAMON0.MONUPDS = 0

   
Exit Sub

Error_Handler:


End Sub
'---------------------------------------------------------
Public Function rsYBIAMON0_Read(lYBIAMON0 As typeYBIAMON0)
'---------------------------------------------------------
Dim xSQL As String, V
On Error GoTo Error_Handler

rsYBIAMON0_Read = Null

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIAMON7" _
    & " where MONAPP = '" & lYBIAMON0.MONAPP & "'" _
    & " and MONFLUX = '" & lYBIAMON0.MONFLUX & "'"
    
Set rsSab = cnsab.Execute(xSQL)
If Not rsSab.EOF Then
    V = rsYBIAMON0_GetBuffer(rsSab, lYBIAMON0)
    If Not IsNull(V) Then rsYBIAMON0_Read = Trim(V)
Else
    rsYBIAMON0_Read = "? rsYBIAMON0_Read : " & lYBIAMON0.MONAPP & "_" & lYBIAMON0.MONFLUX
End If
Exit Function

Error_Handler:
'-------------
    rsYBIAMON0_Read = " rsYBIAMON0_Read : " & Error
End Function



'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsYBIAMON0_GetBuffer(rsAdo As ADODB.Recordset, rsYBIAMON0 As typeYBIAMON0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsYBIAMON0_GetBuffer = Null

rsYBIAMON0.MONAPP = rsAdo("MONAPP")
rsYBIAMON0.MONFLUX = rsAdo("MONFLUX")
rsYBIAMON0.MONSTATUS = rsAdo("MONSTATUS")
rsYBIAMON0.MONNUM = rsAdo("MONNUM")
rsYBIAMON0.MONJOB = rsAdo("MONJOB")
rsYBIAMON0.MONPGM = rsAdo("MONPGM")
rsYBIAMON0.MONUSR = rsAdo("MONUSR")
rsYBIAMON0.MONAMJ = rsAdo("MONAMJ")
rsYBIAMON0.MONHMS = rsAdo("MONHMS")
rsYBIAMON0.MONFILE = rsAdo("MONFILE")
rsYBIAMON0.MONUPDS = rsAdo("MONUPDS")

Exit Function

Error_Handler:

rsYBIAMON0_GetBuffer = Error

End Function


'







