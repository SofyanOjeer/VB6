Attribute VB_Name = "srvYCLISCO0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
'Dim rsSabX As New ADODB.Recordset
Dim rsADO As ADODB.Recordset

Type typeYCLISCO0
 
      CLISCOID     As Long
      
      CLISCOPPEK   As String
      CLISCOPPEV   As Long
      CLISCODEBK   As String
      CLISCODEBV   As Long
      CLISCORESK   As String
      CLISCORESV   As Long
      CLISCONATK   As String
      CLISCONATV   As Long
      CLISCOACTK   As String
      CLISCOACTV   As Long
      CLISCOCOBK   As String
      CLISCOCOBV   As Long
      CLISCOFATK   As String
      CLISCOFATV   As Long
      CLISCOCRSK   As String
      CLISCOCRSV   As Long
      CLISCOSCOK   As String
      CLISCOSCOV   As Long
     
      CLISCOCLID   As String
      CLISCOCLIX   As String
      CLISCOCLIR   As String
    
      CLISCOSTA    As String

      
      CLISCOYAMJ   As Long
      CLISCOYHMS   As Long
      CLISCOYUSR   As String
      CLISCOYVER   As String
      CLISCOINFO   As String
     
   
End Type


Type typeSCOPAY
 
      Id       As String
      V        As Long
      
      GAFI_N   As String
      GAFI_G   As String
      Embargo  As String
      CRS      As String
      BIA_CA   As String
      BIA_1    As String
      BIA_2    As String
      YUSR     As String
      YAMJ     As Long
      YHMS     As Long
End Type

Public Function sqlYCLISCO0_Delete(oldY As typeYCLISCO0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

On Error GoTo Error_Handler
sqlYCLISCO0_Delete = Null
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where CLISCOID = " & oldY.CLISCOID & " and CLISCOYVER = " & oldY.CLISCOYVER


'===================================================================================

    
    xSQL = "delete from " & paramIBM_Library_SABSPE_XXX & ".YCLISCO0" & xWhere
    'Call FEU_ROUGE
    Set rsADO = cnSab_Update.Execute(xSQL, Nb)
    'Call FEU_VERT
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYCLISCO0_Delete = "Erreur màj : " & xWhere
        Exit Function
    End If
    


Exit Function
Error_Handler:
    sqlYCLISCO0_Delete = Error
End Function

Public Function sqlYCLISCO0_Update(newY As typeYCLISCO0, oldY As typeYCLISCO0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

On Error GoTo Error_Handler
sqlYCLISCO0_Update = Null

'===================================================================================

xWhere = " where CLISCOID = " & oldY.CLISCOID & " and CLISCOYVER = " & newY.CLISCOYVER
xSet = " set"
blnUpdate = False
newY.CLISCOYVER = newY.CLISCOYVER + 1

' Détecter les modifications
'===================================================================================
'If newY.CLISCOID <> oldY.CLISCOID Then blnUpdate = True:  xSet = xSet & " , CLISCOID = " & newY.CLISCOID
If newY.CLISCOPPEV <> oldY.CLISCOPPEV Then blnUpdate = True:  xSet = xSet & " , CLISCOPPEV = " & newY.CLISCOPPEV
If newY.CLISCOPPEK <> oldY.CLISCOPPEK Then blnUpdate = True:  xSet = xSet & " , CLISCOPPEK = '" & newY.CLISCOPPEK & "'"
If newY.CLISCODEBV <> oldY.CLISCODEBV Then blnUpdate = True:  xSet = xSet & " , CLISCODEBV = " & newY.CLISCODEBV
If newY.CLISCODEBK <> oldY.CLISCODEBK Then blnUpdate = True:  xSet = xSet & " , CLISCODEBK = '" & newY.CLISCODEBK & "'"
If newY.CLISCORESV <> oldY.CLISCORESV Then blnUpdate = True:  xSet = xSet & " , CLISCORESV = " & newY.CLISCORESV
If newY.CLISCORESK <> oldY.CLISCORESK Then blnUpdate = True:  xSet = xSet & " , CLISCORESK = '" & newY.CLISCORESK & "'"
If newY.CLISCONATV <> oldY.CLISCONATV Then blnUpdate = True:  xSet = xSet & " , CLISCONATV = " & newY.CLISCONATV
If newY.CLISCONATK <> oldY.CLISCONATK Then blnUpdate = True:  xSet = xSet & " , CLISCONATK = '" & newY.CLISCONATK & "'"
If newY.CLISCOACTV <> oldY.CLISCOACTV Then blnUpdate = True:  xSet = xSet & " , CLISCOACTV = " & newY.CLISCOACTV
If newY.CLISCOACTK <> oldY.CLISCOACTK Then blnUpdate = True:  xSet = xSet & " , CLISCOACTK = '" & newY.CLISCOACTK & "'"
If newY.CLISCOCOBV <> oldY.CLISCOCOBV Then blnUpdate = True:  xSet = xSet & " , CLISCOCOBV = " & newY.CLISCOCOBV
If newY.CLISCOCOBK <> oldY.CLISCOCOBK Then blnUpdate = True:  xSet = xSet & " , CLISCOCOBK = '" & newY.CLISCOCOBK & "'"
If newY.CLISCOFATV <> oldY.CLISCOFATV Then blnUpdate = True:  xSet = xSet & " , CLISCOFATV = " & newY.CLISCOFATV
If newY.CLISCOFATK <> oldY.CLISCOFATK Then blnUpdate = True:  xSet = xSet & " , CLISCOFATK = '" & newY.CLISCOFATK & "'"
If newY.CLISCOCRSV <> oldY.CLISCOCRSV Then blnUpdate = True:  xSet = xSet & " , CLISCOCRSV = " & newY.CLISCOCRSV
If newY.CLISCOCRSK <> oldY.CLISCOCRSK Then blnUpdate = True:  xSet = xSet & " , CLISCOCRSK = '" & newY.CLISCOCRSK & "'"

If newY.CLISCOSCOV <> oldY.CLISCOSCOV Then blnUpdate = True:  xSet = xSet & " , CLISCOSCOV = " & newY.CLISCOSCOV
If newY.CLISCOSCOK <> oldY.CLISCOSCOK Then blnUpdate = True:  xSet = xSet & " , CLISCOSCOK = '" & newY.CLISCOSCOK & "'"

If newY.CLISCOYAMJ <> oldY.CLISCOYAMJ Then blnUpdate = True:  xSet = xSet & " , CLISCOYAMJ = " & newY.CLISCOYAMJ
If newY.CLISCOYHMS <> oldY.CLISCOYHMS Then blnUpdate = True:  xSet = xSet & " , CLISCOYHMS = " & newY.CLISCOYHMS
If newY.CLISCOYVER <> oldY.CLISCOYVER Then blnUpdate = True:  xSet = xSet & " , CLISCOYVER = " & newY.CLISCOYVER

If newY.CLISCOCLID <> oldY.CLISCOCLID Then blnUpdate = True:  xSet = xSet & " , CLISCOCLID = '" & newY.CLISCOCLID & "'"
If newY.CLISCOCLIX <> oldY.CLISCOCLIX Then blnUpdate = True:  xSet = xSet & " , CLISCOCLIX = '" & Replace(newY.CLISCOCLIX, "'", "''") & "'"
If newY.CLISCOCLIR <> oldY.CLISCOCLIR Then blnUpdate = True:  xSet = xSet & " , CLISCOCLIR = '" & Replace(newY.CLISCOCLIR, "'", "''") & "'"
If newY.CLISCOSTA <> oldY.CLISCOSTA Then blnUpdate = True:  xSet = xSet & " , CLISCOSTA = '" & newY.CLISCOSTA & "'"
If newY.CLISCOYUSR <> oldY.CLISCOYUSR Then blnUpdate = True:  xSet = xSet & " , CLISCOYUSR = '" & newY.CLISCOYUSR & "'"

If newY.CLISCOINFO <> oldY.CLISCOINFO Then blnUpdate = True:  xSet = xSet & " , CLISCOINFO = '" & Replace(newY.CLISCOINFO, "'", "''") & "'"

If blnUpdate Then
    Mid$(xSet, 1, 6) = " set  "
    xSQL = "update " & paramIBM_Library_SABSPE_XXX & ".YCLISCO0" & xSet & xWhere
    'Call FEU_ROUGE
    Set rsADO = cnSab_Update.Execute(xSQL, Nb)
    'Call FEU_VERT
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYCLISCO0_Update = "Erreur màj : " & newY.CLISCOID
        Exit Function
    End If
    
End If

Exit Function
Error_Handler:
    sqlYCLISCO0_Update = Error
End Function
Public Function sqlYCLISCO0_Update_Field(oldY As typeYCLISCO0, lSQL_Set As String)
Dim xSQL As String, Nb As Long

On Error GoTo Error_Handler
sqlYCLISCO0_Update_Field = Null



xSQL = "update " & paramIBM_Library_SABSPE_XXX & ".YCLISCO0 " & lSQL_Set & "" _
     & " where CLISCOID = " & oldY.CLISCOID _
     & " and CLISCOYVER = " & oldY.CLISCOYVER
     
'Call FEU_ROUGE
Set rsADO = cnSab_Update.Execute(xSQL, Nb)
'Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYCLISCO0_Update_Field = "Erreur màj : " & oldY.CLISCOID
    Exit Function
End If
    

Exit Function
Error_Handler:
    sqlYCLISCO0_Update_Field = Error
End Function


Public Function sqlYCLISCO0_Insert(newY As typeYCLISCO0)
Dim V
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlYCLISCO0_Insert = Null
xSet = " (CLISCOID "
xValues = " values(" & newY.CLISCOID

' Détecter les modifications
'===================================================================================
If newY.CLISCOPPEV <> 0 Then xSet = xSet & ",CLISCOPPEV": xValues = xValues & " ," & newY.CLISCOPPEV
If Trim(newY.CLISCOPPEK) <> "" Then xSet = xSet & ",CLISCOPPEK": xValues = xValues & " ,'" & newY.CLISCOPPEK & "'"

If newY.CLISCODEBV <> 0 Then xSet = xSet & ",CLISCODEBV": xValues = xValues & " ," & newY.CLISCODEBV
If Trim(newY.CLISCODEBK) <> "" Then xSet = xSet & ",CLISCODEBK": xValues = xValues & " ,'" & newY.CLISCODEBK & "'"
If newY.CLISCORESV <> 0 Then xSet = xSet & ",CLISCORESV": xValues = xValues & " ," & newY.CLISCORESV
If Trim(newY.CLISCORESK) <> "" Then xSet = xSet & ",CLISCORESK": xValues = xValues & " ,'" & newY.CLISCORESK & "'"
If newY.CLISCONATV <> 0 Then xSet = xSet & ",CLISCONATV": xValues = xValues & " ," & newY.CLISCONATV
If Trim(newY.CLISCONATK) <> "" Then xSet = xSet & ",CLISCONATK": xValues = xValues & " ,'" & newY.CLISCONATK & "'"
If newY.CLISCOACTV <> 0 Then xSet = xSet & ",CLISCOACTV": xValues = xValues & " ," & newY.CLISCOACTV
If Trim(newY.CLISCOACTK) <> "" Then xSet = xSet & ",CLISCOACTK": xValues = xValues & " ,'" & newY.CLISCOACTK & "'"
If newY.CLISCOCOBV <> 0 Then xSet = xSet & ",CLISCOCOBV": xValues = xValues & " ," & newY.CLISCOCOBV
If Trim(newY.CLISCOCOBK) <> "" Then xSet = xSet & ",CLISCOCOBK": xValues = xValues & " ,'" & newY.CLISCOCOBK & "'"
If newY.CLISCOFATV <> 0 Then xSet = xSet & ",CLISCOFATV": xValues = xValues & " ," & newY.CLISCOFATV
If Trim(newY.CLISCOFATK) <> "" Then xSet = xSet & ",CLISCOFATK": xValues = xValues & " ,'" & newY.CLISCOFATK & "'"
If newY.CLISCOCRSV <> 0 Then xSet = xSet & ",CLISCOCRSV": xValues = xValues & " ," & newY.CLISCOCRSV
If Trim(newY.CLISCOCRSK) <> "" Then xSet = xSet & ",CLISCOCRSK": xValues = xValues & " ,'" & newY.CLISCOCRSK & "'"

If newY.CLISCOSCOV <> 0 Then xSet = xSet & ",CLISCOSCOV": xValues = xValues & " ," & newY.CLISCOSCOV
If Trim(newY.CLISCOSCOK) <> "" Then xSet = xSet & ",CLISCOSCOK": xValues = xValues & " ,'" & newY.CLISCOSCOK & "'"


If newY.CLISCOYAMJ <> 0 Then xSet = xSet & ",CLISCOYAMJ": xValues = xValues & " ," & newY.CLISCOYAMJ
If newY.CLISCOYHMS <> 0 Then xSet = xSet & ",CLISCOYHMS": xValues = xValues & " ," & newY.CLISCOYHMS
If newY.CLISCOYVER <> 0 Then xSet = xSet & ",CLISCOYVER": xValues = xValues & " ," & newY.CLISCOYVER

If Trim(newY.CLISCOCLID) <> "" Then xSet = xSet & ",CLISCOCLID": xValues = xValues & " ,'" & newY.CLISCOCLID & "'"
If Trim(newY.CLISCOCLIX) <> "" Then xSet = xSet & ",CLISCOCLIX": xValues = xValues & " ,'" & Replace(newY.CLISCOCLIX, "'", "''") & "'"
If Trim(newY.CLISCOCLIR) <> "" Then xSet = xSet & ",CLISCOCLIR": xValues = xValues & " ,'" & Replace(newY.CLISCOCLIR, "'", "''") & "'"
If Trim(newY.CLISCOSTA) <> "" Then xSet = xSet & ",CLISCOSTA": xValues = xValues & " ,'" & newY.CLISCOSTA & "'"

If Trim(newY.CLISCOYUSR) <> "" Then xSet = xSet & ",CLISCOYUSR": xValues = xValues & " ,'" & newY.CLISCOYUSR & "'"
      
If Trim(newY.CLISCOINFO) <> "" Then xSet = xSet & ",CLISCOINFO": xValues = xValues & " ,'" & Replace(newY.CLISCOINFO, "'", "''") & "'"
       
      

xSQL = "Insert into " & paramIBM_Library_SABSPE_XXX & ".YCLISCO0" & xSet & ")" & xValues & ")"
'Call FEU_ROUGE
Set rsADO = cnSab_Update.Execute(xSQL, Nb)
'Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYCLISCO0_Insert = "Erreur màj : " & newY.CLISCOID
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYCLISCO0_Insert = Error
End Function

Public Function rsYCLISCO0_GetBuffer(rsADO As ADODB.Recordset, lYCLISCO0 As typeYCLISCO0)
On Error GoTo Error_Handler
rsYCLISCO0_GetBuffer = Null


lYCLISCO0.CLISCOID = rsADO("CLISCOID")

lYCLISCO0.CLISCOPPEK = rsADO("CLISCOPPEK")
lYCLISCO0.CLISCOPPEV = rsADO("CLISCOPPEV")
lYCLISCO0.CLISCODEBK = rsADO("CLISCODEBK")
lYCLISCO0.CLISCODEBV = rsADO("CLISCODEBV")
lYCLISCO0.CLISCORESK = rsADO("CLISCORESK")
lYCLISCO0.CLISCORESV = rsADO("CLISCORESV")
lYCLISCO0.CLISCONATK = rsADO("CLISCONATK")
lYCLISCO0.CLISCONATV = rsADO("CLISCONATV")
lYCLISCO0.CLISCOACTK = rsADO("CLISCOACTK")
lYCLISCO0.CLISCOACTV = rsADO("CLISCOACTV")
lYCLISCO0.CLISCOCOBK = rsADO("CLISCOCOBK")
lYCLISCO0.CLISCOCOBV = rsADO("CLISCOCOBV")
lYCLISCO0.CLISCOFATK = rsADO("CLISCOFATK")
lYCLISCO0.CLISCOFATV = rsADO("CLISCOFATV")
lYCLISCO0.CLISCOCRSK = rsADO("CLISCOCRSK")
lYCLISCO0.CLISCOCRSV = rsADO("CLISCOCRSV")

lYCLISCO0.CLISCOSCOK = rsADO("CLISCOSCOK")
lYCLISCO0.CLISCOSCOV = rsADO("CLISCOSCOV")

lYCLISCO0.CLISCOCLID = rsADO("CLISCOCLID")
lYCLISCO0.CLISCOCLIX = rsADO("CLISCOCLIX")
lYCLISCO0.CLISCOCLIR = rsADO("CLISCOCLIR")
lYCLISCO0.CLISCOSTA = rsADO("CLISCOSTA")

lYCLISCO0.CLISCOYAMJ = rsADO("CLISCOYAMJ")
lYCLISCO0.CLISCOYHMS = rsADO("CLISCOYHMS")
lYCLISCO0.CLISCOYVER = rsADO("CLISCOYVER")
lYCLISCO0.CLISCOYUSR = rsADO("CLISCOYUSR")

lYCLISCO0.CLISCOINFO = rsADO("CLISCOINFO")

Exit Function
Error_Handler:
rsYCLISCO0_GetBuffer = Error


End Function
Public Function rsYCLISCO0_Init(lYCLISCO0 As typeYCLISCO0)


lYCLISCO0.CLISCOID = 0
lYCLISCO0.CLISCOPPEV = 0
lYCLISCO0.CLISCOPPEK = ""
lYCLISCO0.CLISCODEBV = 0
lYCLISCO0.CLISCODEBK = ""
lYCLISCO0.CLISCORESV = 0
lYCLISCO0.CLISCORESK = ""
lYCLISCO0.CLISCONATV = 0
lYCLISCO0.CLISCONATK = ""
lYCLISCO0.CLISCOACTV = 0
lYCLISCO0.CLISCOACTK = ""
lYCLISCO0.CLISCOCOBV = 0
lYCLISCO0.CLISCOCOBK = ""
lYCLISCO0.CLISCOFATV = 0
lYCLISCO0.CLISCOFATK = ""
lYCLISCO0.CLISCOCRSV = 0
lYCLISCO0.CLISCOCRSK = ""

lYCLISCO0.CLISCOSCOV = 0
lYCLISCO0.CLISCOSCOK = ""

lYCLISCO0.CLISCOCLID = ""
lYCLISCO0.CLISCOCLIX = ""
lYCLISCO0.CLISCOCLIR = ""

lYCLISCO0.CLISCOSTA = ""
lYCLISCO0.CLISCOYAMJ = 0
lYCLISCO0.CLISCOYHMS = 0
lYCLISCO0.CLISCOYVER = 0
lYCLISCO0.CLISCOYUSR = ""
lYCLISCO0.CLISCOINFO = ""

End Function



















