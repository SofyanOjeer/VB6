Attribute VB_Name = "srvYSWIRAM0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
'Dim rsSabX As New ADODB.Recordset
Dim rsADO As ADODB.Recordset
Public mYSWIRAM0_SWISABSWID As Long, mYSWIRAM0_Col As Integer
Public arrYSWIRAM0_Fields_K1(100) As String, arrYSWIRAM0_Fields_V1(100) As String, arrYSWIRAM0_Fields_X1(100) As String, arrYSWIRAM0_Fields_Nb1 As Integer
Public arrYSWIRAM0_Fields_X2(100) As String
Public oldYSWISAB0_1 As typeYSWISAB0, oldYSWISAB0_2 As typeYSWISAB0

Type typeYSWIRAM0
 
      SWIRAMXID   As Long
      SWIRAMXOPE    As String
      
      SWIRAMXSEQ   As Long
       
      SWIRAMXREF   As String
      SWIRAMXBIC   As String
      SWIRAMXMTK   As String
      SWIRAMXES    As String
      
      SWIRAMX22    As String
      SWIRAMSTA    As String
      SWIRAMYAMJ   As Long
      SWIRAMYHMS   As Long
      SWIRAMYUSR    As String
      SWIRAMYUPD    As String
     
   
'____________________________________________________ Journalisation
    JORCV                   As Long
    JOSEQN                  As Long
    JRNBIATRN               As Long
    
    JOENTT          As String * 2
    JODATE          As String * 6

'____________________________________________________ Journalisation
End Type


Public Function sqlYSWIRAM0_Delete(oldY As typeYSWIRAM0)
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

On Error GoTo Error_Handler
sqlYSWIRAM0_Delete = Null
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where SWIRAMXID = " & oldY.SWIRAMXID


'===================================================================================

    
    xSql = "delete from " & paramIBM_Library_SABSPE_XXX & ".YSWIRAM0" & xWhere
    'Call FEU_ROUGE
    Set rsADO = cnSab_Update.Execute(xSql, Nb)
    'Call FEU_VERT
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYSWIRAM0_Delete = "Erreur màj : " & xWhere
        Exit Function
    End If
    


Exit Function
Error_Handler:
    sqlYSWIRAM0_Delete = Error
End Function

Public Function sqlYSWIRAM0_Update(newY As typeYSWIRAM0, oldY As typeYSWIRAM0)
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

On Error GoTo Error_Handler
sqlYSWIRAM0_Update = Null

'===================================================================================

xWhere = " where SWIRAMXID = " & oldY.SWIRAMXID
xSet = " set"
blnUpdate = False


' Détecter les modifications
'===================================================================================

If newY.SWIRAMXSEQ <> oldY.SWIRAMXSEQ Then blnUpdate = True:  xSet = xSet & " , SWIRAMXSEQ = " & newY.SWIRAMXSEQ
If newY.SWIRAMYAMJ <> oldY.SWIRAMYAMJ Then blnUpdate = True:  xSet = xSet & " , SWIRAMYAMJ = " & newY.SWIRAMYAMJ
If newY.SWIRAMYHMS <> oldY.SWIRAMYHMS Then blnUpdate = True:  xSet = xSet & " , SWIRAMYHMS = " & newY.SWIRAMYHMS

If newY.SWIRAMXOPE <> oldY.SWIRAMXOPE Then blnUpdate = True:  xSet = xSet & " , SWIRAMXOPE = '" & newY.SWIRAMXOPE & "'"
If newY.SWIRAMXES <> oldY.SWIRAMXES Then blnUpdate = True:  xSet = xSet & " , SWIRAMXES = '" & newY.SWIRAMXES & "'"
If newY.SWIRAMXMTK <> oldY.SWIRAMXMTK Then blnUpdate = True:  xSet = xSet & " , SWIRAMXMTK = '" & newY.SWIRAMXMTK & "'"
If newY.SWIRAMXBIC <> oldY.SWIRAMXBIC Then blnUpdate = True:  xSet = xSet & " , SWIRAMXBIC = '" & newY.SWIRAMXBIC & "'"
If newY.SWIRAMXREF <> oldY.SWIRAMXREF Then blnUpdate = True:  xSet = xSet & " , SWIRAMXREF = '" & Replace(newY.SWIRAMXREF, "'", "''") & "'"
If newY.SWIRAMX22 <> oldY.SWIRAMX22 Then blnUpdate = True:  xSet = xSet & " , SWIRAMX22 = '" & newY.SWIRAMX22 & "'"
If newY.SWIRAMSTA <> oldY.SWIRAMSTA Then blnUpdate = True:  xSet = xSet & " , SWIRAMSTA = '" & newY.SWIRAMSTA & "'"
If newY.SWIRAMYUSR <> oldY.SWIRAMYUSR Then blnUpdate = True:  xSet = xSet & " , SWIRAMYUSR = '" & newY.SWIRAMYUSR & "'"
If newY.SWIRAMYUPD <> oldY.SWIRAMYUPD Then blnUpdate = True:  xSet = xSet & " , SWIRAMYUPD = '" & newY.SWIRAMYUPD & "'"


If blnUpdate Then
    Mid$(xSet, 1, 6) = " set  "
    xSql = "update " & paramIBM_Library_SABSPE_XXX & ".YSWIRAM0" & xSet & xWhere
    'Call FEU_ROUGE
    Set rsADO = cnSab_Update.Execute(xSql, Nb)
    'Call FEU_VERT
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYSWIRAM0_Update = "Erreur màj : " & newY.SWIRAMXID
        Exit Function
    End If
    
End If

Exit Function
Error_Handler:
    sqlYSWIRAM0_Update = Error
End Function
Public Function sqlYSWIRAM0_Update_Field(oldY As typeYSWIRAM0, lSQL_Set As String)
Dim xSql As String, Nb As Long

On Error GoTo Error_Handler
sqlYSWIRAM0_Update_Field = Null



xSql = "update " & paramIBM_Library_SABSPE_XXX & ".YSWIRAM0 " & lSQL_Set & "" _
     & " where SWIRAMXID = " & oldY.SWIRAMXID

'Call FEU_ROUGE
Set rsADO = cnSab_Update.Execute(xSql, Nb)
'Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYSWIRAM0_Update_Field = "Erreur màj : " & oldY.SWIRAMXOPE & " - " & oldY.SWIRAMXID
    Exit Function
End If
    

Exit Function
Error_Handler:
    sqlYSWIRAM0_Update_Field = Error
End Function


Public Function sqlYSWIRAM0_Insert(newY As typeYSWIRAM0)
Dim V
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlYSWIRAM0_Insert = Null
xSet = " (SWIRAMXID "
xValues = " values(" & newY.SWIRAMXID

' Détecter les modifications
'===================================================================================

If newY.SWIRAMXSEQ <> 0 Then xSet = xSet & ",SWIRAMXSEQ": xValues = xValues & " ," & newY.SWIRAMXSEQ
If newY.SWIRAMYAMJ <> 0 Then xSet = xSet & ",SWIRAMYAMJ": xValues = xValues & " ," & newY.SWIRAMYAMJ
If newY.SWIRAMYHMS <> 0 Then xSet = xSet & ",SWIRAMYHMS": xValues = xValues & " ," & newY.SWIRAMYHMS

If Trim(newY.SWIRAMXOPE) <> "" Then xSet = xSet & ",SWIRAMXOPE": xValues = xValues & " ,'" & newY.SWIRAMXOPE & "'"
If Trim(newY.SWIRAMXES) <> "" Then xSet = xSet & ",SWIRAMXES": xValues = xValues & " ,'" & newY.SWIRAMXES & "'"
If Trim(newY.SWIRAMXMTK) <> "" Then xSet = xSet & ",SWIRAMXMTK": xValues = xValues & " ,'" & newY.SWIRAMXMTK & "'"
If Trim(newY.SWIRAMXBIC) <> "" Then xSet = xSet & ",SWIRAMXBIC": xValues = xValues & " ,'" & newY.SWIRAMXBIC & "'"
If Trim(newY.SWIRAMXREF) <> "" Then xSet = xSet & ",SWIRAMXREF": xValues = xValues & " ,'" & Replace(newY.SWIRAMXREF, "'", "''") & "'"
If Trim(newY.SWIRAMX22) <> "" Then xSet = xSet & ",SWIRAMX22": xValues = xValues & " ,'" & newY.SWIRAMX22 & "'"
If Trim(newY.SWIRAMSTA) <> "" Then xSet = xSet & ",SWIRAMSTA": xValues = xValues & " ,'" & newY.SWIRAMSTA & "'"
If Trim(newY.SWIRAMYUPD) <> "" Then xSet = xSet & ",SWIRAMYUPD": xValues = xValues & " ,'" & newY.SWIRAMYUPD & "'"
If Trim(newY.SWIRAMYUSR) <> "" Then xSet = xSet & ",SWIRAMYUSR": xValues = xValues & " ,'" & newY.SWIRAMYUSR & "'"

xSql = "Insert into " & paramIBM_Library_SABSPE_XXX & ".YSWIRAM0" & xSet & ")" & xValues & ")"
'Call FEU_ROUGE
Set rsADO = cnSab_Update.Execute(xSql, Nb)
'Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYSWIRAM0_Insert = "Erreur màj : " & newY.SWIRAMXOPE & " - " & newY.SWIRAMXID
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYSWIRAM0_Insert = Error
End Function

Public Function rsYSWIRAM0_GetBuffer(rsADO As ADODB.Recordset, lYSWIRAM0 As typeYSWIRAM0)
On Error GoTo Error_Handler
rsYSWIRAM0_GetBuffer = Null

lYSWIRAM0.JORCV = 0
lYSWIRAM0.JOSEQN = 0
lYSWIRAM0.JRNBIATRN = 0
lYSWIRAM0.JOENTT = ""
lYSWIRAM0.JODATE = ""

lYSWIRAM0.SWIRAMXOPE = rsADO("SWIRAMXOPE")

lYSWIRAM0.SWIRAMXID = rsADO("SWIRAMXID")
lYSWIRAM0.SWIRAMXSEQ = rsADO("SWIRAMXSEQ")

lYSWIRAM0.SWIRAMXREF = Trim(rsADO("SWIRAMXREF"))

lYSWIRAM0.SWIRAMXES = rsADO("SWIRAMXES")
lYSWIRAM0.SWIRAMXMTK = rsADO("SWIRAMXMTK")
lYSWIRAM0.SWIRAMXBIC = rsADO("SWIRAMXBIC")

lYSWIRAM0.SWIRAMX22 = rsADO("SWIRAMX22")

lYSWIRAM0.SWIRAMSTA = rsADO("SWIRAMSTA")
lYSWIRAM0.SWIRAMYAMJ = rsADO("SWIRAMYAMJ")
lYSWIRAM0.SWIRAMYHMS = rsADO("SWIRAMYHMS")
lYSWIRAM0.SWIRAMYUPD = rsADO("SWIRAMYUPD")
lYSWIRAM0.SWIRAMYUSR = rsADO("SWIRAMYUSR")




Exit Function
Error_Handler:
rsYSWIRAM0_GetBuffer = Error


End Function
'---------------------------------------------------------
Public Function rsJSWIRAM0_GetBuffer(rsADO As ADODB.Recordset, rsYSWIRAM0 As typeYSWIRAM0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsJSWIRAM0_GetBuffer = Null

rsJSWIRAM0_GetBuffer = rsYSWIRAM0_GetBuffer(rsADO, rsYSWIRAM0)
rsYSWIRAM0.JORCV = rsADO("JORCV")
rsYSWIRAM0.JOSEQN = rsADO("JOSEQN")
rsYSWIRAM0.JRNBIATRN = rsADO("JRNBIATRN")
rsYSWIRAM0.JOENTT = rsADO("JOENTT")
rsYSWIRAM0.JODATE = rsADO("JODATE")

Exit Function

Error_Handler:

rsJSWIRAM0_GetBuffer = Error

End Function


Public Function rsYSWIRAM0_Init(lYSWIRAM0 As typeYSWIRAM0)


lYSWIRAM0.SWIRAMXID = 0
lYSWIRAM0.SWIRAMXSEQ = 0
lYSWIRAM0.SWIRAMXOPE = ""
lYSWIRAM0.SWIRAMXREF = ""

lYSWIRAM0.SWIRAMXES = ""
lYSWIRAM0.SWIRAMXMTK = ""
lYSWIRAM0.SWIRAMXBIC = ""

lYSWIRAM0.SWIRAMYAMJ = 0
lYSWIRAM0.SWIRAMYHMS = 0
lYSWIRAM0.SWIRAMX22 = ""
lYSWIRAM0.SWIRAMSTA = ""
lYSWIRAM0.SWIRAMYUPD = ""
lYSWIRAM0.SWIRAMYUSR = ""


lYSWIRAM0.JORCV = 0
lYSWIRAM0.JOSEQN = 0
lYSWIRAM0.JRNBIATRN = 0
    
lYSWIRAM0.JOENTT = ""
lYSWIRAM0.JODATE = ""

End Function

















