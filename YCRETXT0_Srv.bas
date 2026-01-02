Attribute VB_Name = "srvYCRETXT0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Type typeYCRETXT0
    CRETXTETA       As Integer                        ' ETABLISSEMENT
    CRETXTCRE       As Long                       ' AGENCE OPERATRICE
    CRETXTYUSR   As String
    CRETXTYAMJ   As Long
    CRETXTYHMS   As Long
    CRETXTYVER   As Long
    CRETXTINFO       As String
    
End Type
Public Function sqlYCRETXT0_Insert(newY As typeYCRETXT0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlYCRETXT0_Insert = Null

xSet = " (CRETXTETA,CRETXTCRE"
xValues = " values(" & newY.CRETXTETA & "," & newY.CRETXTCRE

' Détecter les modifications
'===================================================================================
If Trim(newY.CRETXTINFO) <> "" Then xSet = xSet & ",CRETXTINFO": xValues = xValues & " ,'" & Replace(Trim(newY.CRETXTINFO), "'", "''") & "'"

If newY.CRETXTYVER <> 0 Then xSet = xSet & ",CRETXTYVER": xValues = xValues & " ," & newY.CRETXTYVER
If newY.CRETXTYAMJ <> 0 Then xSet = xSet & ",CRETXTYAMJ": xValues = xValues & " ," & newY.CRETXTYAMJ
If newY.CRETXTYHMS <> 0 Then xSet = xSet & ",CRETXTYHMS": xValues = xValues & " ," & newY.CRETXTYHMS

If Trim(newY.CRETXTYUSR) <> "" Then xSet = xSet & ",CRETXTYUSR": xValues = xValues & " ,'" & Replace(Trim(newY.CRETXTYUSR), "'", "''") & "'"

xSQL = "Insert into " & paramIBM_Library_SABSPE & ".YCRETXT0" & xSet & ")" & xValues & ")"

Set rsSab = cnsab.Execute(xSQL, Nb)

' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYCRETXT0_Insert = "Erreur màj : " & newY.CRETXTCRE
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYCRETXT0_Insert = "sqlYCRETXT0_Insert " & vbCrLf & Error
End Function

Public Function sqlYCRETXT0_Update(newY As typeYCRETXT0, oldY As typeYCRETXT0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

On Error GoTo Error_Handler
sqlYCRETXT0_Update = Null

' Contrôle  : Même clé d'accès old / new
'===================================================================================
If oldY.CRETXTCRE <> newY.CRETXTCRE Then
    sqlYCRETXT0_Update = "Erreur CRETXTCRE : " & newY.CRETXTCRE & " / " & oldY.CRETXTCRE
    Exit Function
End If
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where CRETXTETA = " & oldY.CRETXTETA _
       & " and CRETXTCRE = " & oldY.CRETXTCRE
       

      
newY.CRETXTYVER = newY.CRETXTYVER + 1
xSet = xSet & " set CRETXTYVER = " & newY.CRETXTYVER
blnUpdate = False

' Détecter les modifications
'===================================================================================
If newY.CRETXTYAMJ <> oldY.CRETXTYAMJ Then blnUpdate = True: xSet = xSet & " , CRETXTYAMJ = " & newY.CRETXTYAMJ
If newY.CRETXTYHMS <> oldY.CRETXTYHMS Then blnUpdate = True: xSet = xSet & " , CRETXTYHMS = " & newY.CRETXTYHMS

If newY.CRETXTINFO <> oldY.CRETXTINFO Then blnUpdate = True:  xSet = xSet & " , CRETXTINFO= '" & Replace(Trim(newY.CRETXTINFO), "'", "''") & "'"
If newY.CRETXTYUSR <> oldY.CRETXTYUSR Then blnUpdate = True:  xSet = xSet & " , CRETXTYUSR = '" & Replace(Trim(newY.CRETXTYUSR), "'", "''") & "'"
  
    xSQL = "update " & paramIBM_Library_SABSPE & ".YCRETXT0" & xSet & xWhere
    
    Set rsSab = cnsab.Execute(xSQL, Nb)
    
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYCRETXT0_Update = "Erreur màj : " & newY.CRETXTCRE
        Exit Function
    End If
    


Exit Function
Error_Handler:
    sqlYCRETXT0_Update = "sqlYCRETXT0_Update " & vbCrLf & Error
End Function


Public Sub rsYCRETXT0_Init(rsYCRETXT0 As typeYCRETXT0)
rsYCRETXT0.CRETXTETA = 0
rsYCRETXT0.CRETXTCRE = 0
rsYCRETXT0.CRETXTINFO = ""
rsYCRETXT0.CRETXTYAMJ = 0
rsYCRETXT0.CRETXTYHMS = 0
rsYCRETXT0.CRETXTYUSR = ""
rsYCRETXT0.CRETXTYVER = 0
End Sub
Public Function rsYCRETXT0_GetBuffer(rsAdo As ADODB.Recordset, lYCRETXT0 As typeYCRETXT0)
On Error GoTo Error_Handler
rsYCRETXT0_GetBuffer = Null
lYCRETXT0.CRETXTETA = rsAdo("CRETXTETA")
lYCRETXT0.CRETXTCRE = rsAdo("CRETXTCRE")
lYCRETXT0.CRETXTINFO = rsAdo("CRETXTINFO")
lYCRETXT0.CRETXTYUSR = Trim(rsAdo("CRETXTYUSR"))
lYCRETXT0.CRETXTYAMJ = rsAdo("CRETXTYAMJ")

lYCRETXT0.CRETXTYHMS = rsAdo("CRETXTYHMS")
lYCRETXT0.CRETXTYVER = rsAdo("CRETXTYVER")

Exit Function
Error_Handler:
rsYCRETXT0_GetBuffer = Error
End Function


