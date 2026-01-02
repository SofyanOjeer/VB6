Attribute VB_Name = "srvYPDCLOG0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public paramPDCLOG_Path As String

Dim rsAdo As ADODB.Recordset
 
Type typeYPDCLOG0
 
      PDCLOGDTR     As String * 8   'date comptable
      PDCLOGUAMJ    As String * 8   'date màj
      PDCLOGUHMS    As String * 6   'heure màj
      PDCLOGUSEQ    As Long         'séq màj
      PDCLOGPIE     As Long         'N° pièce
      PDCLOGECR     As Long         'N° écriture
      PDCLOGNAT     As String * 3   'nature de l'enregistrement
      PDCLOGTXT     As String * 64  'texte
      PDCLOGSTA     As String * 1   'statut
      PDCLOGUUSR    As String * 12  'utilisateur màj
      PDCLOGUPDS    As Long         'séq màj
End Type
Public xYPDCLOG0 As typeYPDCLOG0
Public Function sqlYPDCLOG0_DeleteW(lWhere As String)
Dim X As String, xSQL As String, Nb As Long

On Error GoTo Error_Handler
sqlYPDCLOG0_DeleteW = Null
    
xSQL = "delete from " & paramIBM_Library_SABSPE_XXX & ".YPDCLOG0 " & lWhere
Call FEU_ROUGE
Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
 Call FEU_VERT
Exit Function
Error_Handler:
    sqlYPDCLOG0_DeleteW = Error
End Function

Public Function sqlYPDCLOG0_Insert(newY As typeYPDCLOG0)
Dim V
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlYPDCLOG0_Insert = Null
xSet = " (PDCLOGDTR"
xValues = " values('" & newY.PDCLOGDTR & "'"


' Détecter les modifications
'===================================================================================
If newY.PDCLOGPIE <> 0 Then xSet = xSet & ",PDCLOGPIE": xValues = xValues & " ," & newY.PDCLOGPIE
If newY.PDCLOGECR <> 0 Then xSet = xSet & ",PDCLOGECR": xValues = xValues & " ," & newY.PDCLOGECR
If newY.PDCLOGUSEQ <> 0 Then xSet = xSet & ",PDCLOGUSEQ": xValues = xValues & " ," & Replace(newY.PDCLOGUSEQ, ",", ".")

If Trim(newY.PDCLOGUAMJ) <> "" Then xSet = xSet & ",PDCLOGUAMJ": xValues = xValues & " ,'" & newY.PDCLOGUAMJ & "'"
If Trim(newY.PDCLOGUHMS) <> "" Then xSet = xSet & ",PDCLOGUHMS": xValues = xValues & " ,'" & newY.PDCLOGUHMS & "'"
If Trim(newY.PDCLOGUUSR) <> "" Then xSet = xSet & ",PDCLOGUUSR": xValues = xValues & " ,'" & newY.PDCLOGUUSR & "'"
If Trim(newY.PDCLOGNAT) <> "" Then xSet = xSet & ",PDCLOGNAT": xValues = xValues & " ,'" & newY.PDCLOGNAT & "'"
If Trim(newY.PDCLOGTXT) <> "" Then xSet = xSet & ",PDCLOGTXT": xValues = xValues & " ,'" & Replace(newY.PDCLOGTXT, "'", "''") & "'"
If Trim(newY.PDCLOGSTA) <> "" Then xSet = xSet & ",PDCLOGSTA": xValues = xValues & " ,'" & newY.PDCLOGSTA & "'"

Call FEU_ROUGE
xSQL = "Insert into " & paramIBM_Library_SABSPE_XXX & ".YPDCLOG0" & xSet & ")" & xValues & ")"
Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYPDCLOG0_Insert = "Erreur màj : " & newY.PDCLOGDTR & " " & newY.PDCLOGPIE & " " & newY.PDCLOGECR
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYPDCLOG0_Insert = Error
End Function



Public Function rsYPDCLOG0_GetBuffer(rsAdo As ADODB.Recordset, lYPDCLOG0 As typeYPDCLOG0)
On Error GoTo Error_Handler
rsYPDCLOG0_GetBuffer = Null

lYPDCLOG0.PDCLOGDTR = rsAdo("PDCLOGDTR")
lYPDCLOG0.PDCLOGPIE = rsAdo("PDCLOGPIE")
lYPDCLOG0.PDCLOGECR = rsAdo("PDCLOGECR")
lYPDCLOG0.PDCLOGUAMJ = rsAdo("PDCLOGUAMJ")
lYPDCLOG0.PDCLOGUHMS = rsAdo("PDCLOGUHMS")
lYPDCLOG0.PDCLOGUSEQ = rsAdo("PDCLOGUSEQ")
lYPDCLOG0.PDCLOGUUSR = rsAdo("PDCLOGUUSR")
lYPDCLOG0.PDCLOGNAT = rsAdo("PDCLOGNAT")
lYPDCLOG0.PDCLOGTXT = rsAdo("PDCLOGTXT")
lYPDCLOG0.PDCLOGSTA = rsAdo("PDCLOGSTA")
lYPDCLOG0.PDCLOGUPDS = rsAdo("PDCLOGUPDS")

Exit Function
Error_Handler:
rsYPDCLOG0_GetBuffer = Error


End Function

Public Function rsYPDCLOG0_Init(lYPDCLOG0 As typeYPDCLOG0)

lYPDCLOG0.PDCLOGDTR = ""
lYPDCLOG0.PDCLOGPIE = 0
lYPDCLOG0.PDCLOGECR = 0
lYPDCLOG0.PDCLOGUAMJ = ""
lYPDCLOG0.PDCLOGUHMS = ""
lYPDCLOG0.PDCLOGUSEQ = 0
lYPDCLOG0.PDCLOGUUSR = ""
lYPDCLOG0.PDCLOGNAT = ""
lYPDCLOG0.PDCLOGTXT = ""
lYPDCLOG0.PDCLOGSTA = ""
lYPDCLOG0.PDCLOGUPDS = 0

End Function






