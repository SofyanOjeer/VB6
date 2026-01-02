Attribute VB_Name = "srvYFECLOG0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim rsADO As ADODB.Recordset
 
Type typeYFECLOG0
 
      FECLOGAMJ   As Long
      FECLOGHMS   As Long
      FECLOGSEQ   As Integer
      FECLOGUSR   As String
      FECLOGK     As String
      FECLOGSTA   As String
      FECLOGAA    As Long
      FECLOGNB    As Long
      FECLOGTXT   As String
    

'____________________________________________________
End Type
Public xYFECMVT0 As typeYFECMVT0

Type typeYFECMVT0
 
      FECMVTAA    As Long
      FECMVTSEQ   As Long
      
      FECMVTPIE   As Long
      FECMVTECR   As Long
      FECMVTCOM   As String
      FECMVTMTD   As Currency
   

'____________________________________________________
End Type

Type typeYFEC0
 
      FEC_DEV   As String
      FEC_PCI(10)   As String
      FEC_SD0(10)   As Currency
      FEC_DB(10)    As Currency
      FEC_CR(10)    As Currency
      FEC_SD1(10)   As Currency
    

'____________________________________________________ Journalisation
End Type

Public Function sqlYFECLOG0_Insert(newY As typeYFECLOG0)
Dim V
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlYFECLOG0_Insert = Null
newY.FECLOGUSR = usrName_UCase
newY.FECLOGAMJ = DSys
newY.FECLOGHMS = time_Hms * 100
xSet = " (FECLOGAMJ"
xValues = " values(" & newY.FECLOGAMJ


' Détecter les modifications
'===================================================================================
If newY.FECLOGSEQ <> 0 Then xSet = xSet & ",FECLOGSEQ": xValues = xValues & " ," & newY.FECLOGSEQ
If newY.FECLOGNB <> 0 Then xSet = xSet & ",FECLOGNB": xValues = xValues & " ," & newY.FECLOGNB
If newY.FECLOGHMS <> 0 Then xSet = xSet & ",FECLOGHMS": xValues = xValues & " ," & newY.FECLOGHMS
If Trim(newY.FECLOGAA) <> "" Then xSet = xSet & ",FECLOGAA": xValues = xValues & " ," & newY.FECLOGAA


If Trim(newY.FECLOGUSR) <> "" Then xSet = xSet & ",FECLOGUSR": xValues = xValues & " ,'" & Replace(Trim(newY.FECLOGUSR), "'", "''") & "'"
If Trim(newY.FECLOGK) <> "" Then xSet = xSet & ",FECLOGK": xValues = xValues & " ,'" & newY.FECLOGK & "'"
If Trim(newY.FECLOGSTA) <> "" Then xSet = xSet & ",FECLOGSTA": xValues = xValues & " ,'" & Replace(Trim(newY.FECLOGSTA), "'", "''") & "'"
If Trim(newY.FECLOGTXT) <> "" Then xSet = xSet & ",FECLOGTXT": xValues = xValues & " ,'" & Replace(Trim(newY.FECLOGTXT), "'", "''") & "'"


xSQL = "Insert into " & paramIBM_Library_SABSPE_XXX & ".YFECLOG0" & xSet & ")" & xValues & ")"
Call FEU_ROUGE
Set rsADO = cnSab_Update.Execute(xSQL, Nb)
Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYFECLOG0_Insert = "Erreur màj : " & newY.FECLOGSEQ
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYFECLOG0_Insert = Error
End Function

Public Function rsYFECLOG0_GetBuffer(rsADO As ADODB.Recordset, lYFECLOG0 As typeYFECLOG0)
On Error GoTo Error_Handler
rsYFECLOG0_GetBuffer = Null

lYFECLOG0.FECLOGSEQ = rsADO("FECLOGSEQ")
lYFECLOG0.FECLOGUSR = rsADO("FECLOGUSR")
lYFECLOG0.FECLOGK = rsADO("FECLOGK")
lYFECLOG0.FECLOGSTA = rsADO("FECLOGSTA")
lYFECLOG0.FECLOGTXT = rsADO("FECLOGTXT")

lYFECLOG0.FECLOGAA = rsADO("FECLOGAA")
lYFECLOG0.FECLOGAMJ = rsADO("FECLOGAMJ")
lYFECLOG0.FECLOGHMS = rsADO("FECLOGHMS")
lYFECLOG0.FECLOGNB = rsADO("FECLOGNB")

Exit Function
Error_Handler:
rsYFECLOG0_GetBuffer = Error


End Function
Public Function rsYFECLOG0_Init(lYFECLOG0 As typeYFECLOG0)

lYFECLOG0.FECLOGSEQ = 0
lYFECLOG0.FECLOGUSR = ""
lYFECLOG0.FECLOGK = ""
lYFECLOG0.FECLOGSTA = ""
lYFECLOG0.FECLOGTXT = ""
    
lYFECLOG0.FECLOGAA = 0
lYFECLOG0.FECLOGAMJ = 0
lYFECLOG0.FECLOGHMS = 0
lYFECLOG0.FECLOGNB = 0
End Function








