Attribute VB_Name = "rsYCHQMON0"
'---------------------------------------------------------
Option Explicit
Type typeYCHQMON0

      CHQRC1ETA     As Long
      CHQRC1AGE     As Long
      CHQRC1SER     As String * 2
      CHQRC1SSE     As String * 2
      CHQRC1OPE     As String * 3
      CHQRC1DOS     As Long
      CHQRC1DCR     As Long
      CHQDATE       As Long
      CHQCOMPTE     As String * 20
      CHQCREM       As String * 8
      CHQDEVISE     As String * 3
      CHQMONTANT    As Currency
      CHQNB         As Long
      CHQMONSTA     As String * 1   '
      CHQMONUPDS    As Long

End Type

Public Function rsYCHQMON0_Read(oldY As typeYCHQMON0)
Dim xWhere As String, xSQL As String
Dim Nb As Long
rsYCHQMON0_Read = "?"

xWhere = " where CHQRC1ETA = " & oldY.CHQRC1ETA _
& " and CHQRC1AGE = " & oldY.CHQRC1AGE _
& " and CHQRC1SER = '" & oldY.CHQRC1SER & "'" _
& " and CHQRC1SSE = '" & oldY.CHQRC1SSE & "'" _
& " and CHQRC1OPE = '" & oldY.CHQRC1OPE & "'" _
& " and CHQRC1DOS = " & oldY.CHQRC1DOS '_
'& " and CHQMONUPDS = " & oldY.CHQMONUPDS

xSQL = "Select * from " & paramIBM_Library_SABSPE & ".YCHQMON0" & xWhere

Set rsSab = cnsab.Execute(xSQL, Nb)

'===================================================================================

If Not rsSab.EOF Then
    rsYCHQMON0_Read = rsYCHQMON0_GetBuffer(rsSab, oldY)
Else
    rsYCHQMON0_Read = "?Inconnu"
End If

End Function

'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsYCHQMON0_GetBuffer(rsSab As ADODB.Recordset, rsYCHQMON0 As typeYCHQMON0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsYCHQMON0_GetBuffer = Null

rsYCHQMON0.CHQRC1ETA = rsSab("CHQRC1ETA")
rsYCHQMON0.CHQRC1AGE = rsSab("CHQRC1AGE")
rsYCHQMON0.CHQRC1SER = rsSab("CHQRC1SER")
rsYCHQMON0.CHQRC1SSE = rsSab("CHQRC1SSE")
rsYCHQMON0.CHQRC1OPE = rsSab("CHQRC1OPE")
rsYCHQMON0.CHQRC1DOS = rsSab("CHQRC1DOS")
rsYCHQMON0.CHQRC1DCR = rsSab("CHQRC1DCR")
rsYCHQMON0.CHQDATE = rsSab("CHQDATE")
rsYCHQMON0.CHQCOMPTE = rsSab("CHQCOMPTE")
rsYCHQMON0.CHQCREM = rsSab("CHQCREM")
rsYCHQMON0.CHQDEVISE = rsSab("CHQDEVISE")
rsYCHQMON0.CHQMONTANT = rsSab("CHQMONTANT")
rsYCHQMON0.CHQNB = rsSab("CHQNB")
rsYCHQMON0.CHQMONSTA = rsSab("CHQMONSTA")
rsYCHQMON0.CHQMONUPDS = rsSab("CHQMONUPDS")

Exit Function

Error_Handler:

rsYCHQMON0_GetBuffer = Error

End Function

'---------------------------------------------------------
Public Sub rsYCHQMON0_Init(rsYCHQMON0 As typeYCHQMON0)
rsYCHQMON0.CHQRC1ETA = 0
rsYCHQMON0.CHQRC1AGE = 0
rsYCHQMON0.CHQRC1SER = ""
rsYCHQMON0.CHQRC1SSE = ""
rsYCHQMON0.CHQRC1OPE = ""
rsYCHQMON0.CHQRC1DOS = 0
rsYCHQMON0.CHQRC1DCR = 0
rsYCHQMON0.CHQDATE = 0
rsYCHQMON0.CHQCOMPTE = ""
rsYCHQMON0.CHQCREM = ""
rsYCHQMON0.CHQDEVISE = ""
rsYCHQMON0.CHQMONTANT = 0
rsYCHQMON0.CHQNB = 0
rsYCHQMON0.CHQMONSTA = ""
rsYCHQMON0.CHQMONUPDS = 0
'---------------------------------------------------------

End Sub


'









