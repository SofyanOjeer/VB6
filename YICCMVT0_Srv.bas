Attribute VB_Name = "srvYICCMVT0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim rsAdo As ADODB.Recordset
 
Type typeYICCMVT0

    ICCMVTETA   As Integer       ' établissement
    ICCMVTAGE   As Integer       ' agence
    ICCMVTCOM   As String * 20   ' compte
    ICCMVTSER   As String * 2    ' service
    ICCMVTSSE   As String * 2    ' service
    ICCMVTOPE   As String * 3  ' opération
    ICCMVTDOS   As Long         ' groupe
    ICCMVTEVE   As String * 3  ' évenement
    ICCMVTAMJ   As Long         ' date situation
    ICCMVTNAT   As String * 6   ' nature
    ICCMVTEVEG  As String * 3  ' évenement
    ICCMVTRBT   As Currency    ' provision
    ICCMVTPRO   As Currency    ' provision
    ICCMVTTDB   As Currency     ' cumul DB
    ICCMVTTCR   As Currency     ' cumul CR
    
End Type
Public Function sqlYICCMVT0_Insert(newY As typeYICCMVT0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlYICCMVT0_Insert = Null

xSet = " (ICCMVTETA"
xValues = " values(" & newY.ICCMVTETA

' Insertion :
'===================================================================================
If newY.ICCMVTAGE <> 0 Then xSet = xSet & ",ICCMVTAGE": xValues = xValues & ", " & newY.ICCMVTAGE
If newY.ICCMVTAMJ <> 0 Then xSet = xSet & ",ICCMVTAMJ": xValues = xValues & ", " & newY.ICCMVTAMJ
If newY.ICCMVTTDB <> 0 Then xSet = xSet & ",ICCMVTTDB": xValues = xValues & ", " & cur_P(newY.ICCMVTTDB)
If newY.ICCMVTTCR <> 0 Then xSet = xSet & ",ICCMVTTCR": xValues = xValues & ", " & cur_P(newY.ICCMVTTCR)

'===================================================================================

If Trim(newY.ICCMVTCOM) <> "" Then xSet = xSet & ",ICCMVTCOM": xValues = xValues & ", '" & Replace(Trim(newY.ICCMVTCOM), "'", "''") & "'"
Call FEU_ROUGE
xSQL = "Insert into " & paramIBM_Library_SABSPE & ".YICCMVT0" & xSet & ")" & xValues & ")"

Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYICCMVT0_Insert = "Erreur màj : " & newY.ICCMVTCOM
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYICCMVT0_Insert = Error
End Function

'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsYICCMVT0_GetBuffer(rsAdo As ADODB.Recordset, rsYICCMVT0 As typeYICCMVT0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsYICCMVT0_GetBuffer = Null

rsYICCMVT0.ICCMVTETA = rsAdo("ICCMVTETA")
rsYICCMVT0.ICCMVTAGE = rsAdo("ICCMVTAGE")
rsYICCMVT0.ICCMVTCOM = rsAdo("ICCMVTCOM")
rsYICCMVT0.ICCMVTSER = rsAdo("ICCMVTSER")
rsYICCMVT0.ICCMVTSSE = rsAdo("ICCMVTSSE")
rsYICCMVT0.ICCMVTDOS = rsAdo("ICCMVTDOS")
rsYICCMVT0.ICCMVTOPE = rsAdo("ICCMVTOPE")
rsYICCMVT0.ICCMVTEVE = rsAdo("ICCMVTEVE")
rsYICCMVT0.ICCMVTAMJ = rsAdo("ICCMVTAMJ")

rsYICCMVT0.ICCMVTNAT = rsAdo("ICCMVTNAT")
rsYICCMVT0.ICCMVTEVEG = rsAdo("ICCMVTEVEG")
rsYICCMVT0.ICCMVTRBT = rsAdo("ICCMVTRBT")
rsYICCMVT0.ICCMVTPRO = rsAdo("ICCMVTPRO")
rsYICCMVT0.ICCMVTTDB = rsAdo("ICCMVTTDB")
rsYICCMVT0.ICCMVTTCR = rsAdo("ICCMVTTCR")

Exit Function

Error_Handler:

rsYICCMVT0_GetBuffer = Error


End Function









Public Sub rsYICCMVT0_Init(lYICCMVT0 As typeYICCMVT0)
lYICCMVT0.ICCMVTETA = 0                  ' chèque
lYICCMVT0.ICCMVTAGE = 0
lYICCMVT0.ICCMVTCOM = ""
lYICCMVT0.ICCMVTSER = ""
lYICCMVT0.ICCMVTSSE = ""
lYICCMVT0.ICCMVTOPE = ""
lYICCMVT0.ICCMVTEVE = ""
lYICCMVT0.ICCMVTDOS = 0
lYICCMVT0.ICCMVTAMJ = 0

lYICCMVT0.ICCMVTNAT = ""
lYICCMVT0.ICCMVTEVEG = ""
lYICCMVT0.ICCMVTRBT = 0
lYICCMVT0.ICCMVTPRO = 0
lYICCMVT0.ICCMVTTDB = 0
lYICCMVT0.ICCMVTTCR = 0

End Sub




