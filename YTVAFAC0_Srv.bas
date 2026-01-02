Attribute VB_Name = "srvYTVAFAC0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim rsAdo As ADODB.Recordset
 
Type typeYTVAFAC0
 
      TVAFACETA   As Long         'établissement
      TVAFACCLIC  As String * 1   'table client espace,G, D
      TVAFACCLI   As String * 7   'code client
      TVAFACCLIP  As String * 2   'pays de résidence
      TVAFACCLIT  As String * 18  'code TVA intracommunautaire
      TVAFACMTTC   As Currency     'montant en dev
      TVAFACMTVA  As Currency     'montant euro
      TVAFACMEXO  As Currency     'montant TVA issu du dossier
      TVAFACFACN  As Long         'n° facture
      TVAFACDTR   As Long         'date d'édition de la facture
      TVAFACSTA   As String * 1   'statut
      TVAFACUPDS  As Long         'SéQUENCE UPD
      TVAFACUSR   As String * 10   'user
      
      
End Type
Public xYTVAFAC0 As typeYTVAFAC0

Public Function sqlYTVAFAC0_Insert(newY As typeYTVAFAC0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlYTVAFAC0_Insert = Null

xSet = " (TVAFACETA,TVAFACFACN"
xValues = " values(" & newY.TVAFACETA & " ," & newY.TVAFACFACN

' Détecter les modifications
'===================================================================================
If newY.TVAFACMTTC <> 0 Then xSet = xSet & ",TVAFACMTTC": xValues = xValues & " ," & cur_P(newY.TVAFACMTTC)
If newY.TVAFACMTVA <> 0 Then xSet = xSet & ",TVAFACMTVA": xValues = xValues & " ," & cur_P(newY.TVAFACMTVA)
If newY.TVAFACMEXO <> 0 Then xSet = xSet & ",TVAFACMEXO": xValues = xValues & " ," & cur_P(newY.TVAFACMEXO)
If newY.TVAFACDTR <> 0 Then xSet = xSet & ",TVAFACDTR": xValues = xValues & " ," & newY.TVAFACDTR

If Trim(newY.TVAFACCLIC) <> "" Then xSet = xSet & ",TVAFACCLIC": xValues = xValues & " ,'" & newY.TVAFACCLIC & "'"
If Trim(newY.TVAFACCLI) <> "" Then xSet = xSet & ",TVAFACCLI": xValues = xValues & " ,'" & newY.TVAFACCLI & "'"
If Trim(newY.TVAFACCLIP) <> "" Then xSet = xSet & ",TVAFACCLIP": xValues = xValues & " ,'" & newY.TVAFACCLIP & "'"
If Trim(newY.TVAFACCLIT) <> "" Then xSet = xSet & ",TVAFACCLIT": xValues = xValues & " ,'" & Replace(Trim(newY.TVAFACCLIT), "'", "''") & "'"
If Trim(newY.TVAFACSTA) <> "" Then xSet = xSet & ",TVAFACSTA": xValues = xValues & " ,'" & newY.TVAFACSTA & "'"

newY.TVAFACUSR = usrName_UCase10
xSet = xSet & ",TVAFACUSR": xValues = xValues & " ,'" & usrName_UCase10 & "'"
Call FEU_ROUGE
xSQL = "Insert into " & paramIBM_Library_SABSPE & ".YTVAFAC0" & xSet & ")" & xValues & ")"

Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYTVAFAC0_Insert = "Erreur màj : " & newY.TVAFACFACN
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYTVAFAC0_Insert = Error
End Function

Public Function sqlYTVAFAC0_Update(newY As typeYTVAFAC0, oldY As typeYTVAFAC0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

On Error GoTo Error_Handler
sqlYTVAFAC0_Update = Null

' Contrôle  : Même clé d'accès old / new
'===================================================================================
If oldY.TVAFACETA <> newY.TVAFACETA _
Or oldY.TVAFACFACN <> newY.TVAFACFACN Then
    sqlYTVAFAC0_Update = "Erreur TVAFACFACN: " & newY.TVAFACFACN & "." & oldY.TVAFACFACN
    Exit Function
End If
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where TVAFACETA = " & oldY.TVAFACETA _
       & " and TVAFACFACN = " & oldY.TVAFACFACN _
       & " and TVAFACUPDS = " & oldY.TVAFACUPDS

newY.TVAFACUPDS = newY.TVAFACUPDS + 1
xSet = xSet & " set TVAFACUPDS = " & newY.TVAFACUPDS
blnUpdate = False

' Détecter les modifications
'===================================================================================
If newY.TVAFACMTTC <> oldY.TVAFACMTTC Then blnUpdate = True:  xSet = xSet & " , TVAFACMTTC = '" & cur_P(newY.TVAFACMTTC) & "'"
If newY.TVAFACMTVA <> oldY.TVAFACMTVA Then blnUpdate = True:  xSet = xSet & " , TVAFACMTVA = '" & cur_P(newY.TVAFACMTVA) & "'"
If newY.TVAFACMEXO <> oldY.TVAFACMEXO Then blnUpdate = True: xSet = xSet & " , TVAFACMEXO = " & newY.TVAFACMEXO
If newY.TVAFACDTR <> oldY.TVAFACDTR Then blnUpdate = True: xSet = xSet & " , TVAFACDTR = " & newY.TVAFACDTR

If newY.TVAFACCLIC <> oldY.TVAFACCLIC Then blnUpdate = True:  xSet = xSet & " , TVAFACCLIC = '" & Replace(Trim(newY.TVAFACCLIC), "'", "''") & "'"
If newY.TVAFACCLI <> oldY.TVAFACCLI Then blnUpdate = True:  xSet = xSet & " , TVAFACCLI = '" & Replace(Trim(newY.TVAFACCLI), "'", "''") & "'"
If newY.TVAFACCLIP <> oldY.TVAFACCLIP Then blnUpdate = True:  xSet = xSet & " , TVAFACCLIP = '" & Replace(Trim(newY.TVAFACCLIP), "'", "''") & "'"
If newY.TVAFACCLIT <> oldY.TVAFACCLIT Then blnUpdate = True:  xSet = xSet & " , TVAFACCLIT = '" & Replace(Trim(newY.TVAFACCLIT), "'", "''") & "'"
If newY.TVAFACSTA <> oldY.TVAFACSTA Then blnUpdate = True:  xSet = xSet & " , TVAFACSTA = '" & newY.TVAFACSTA & "'"

newY.TVAFACUSR = usrName_UCase10
xSet = xSet & " , TVAFACUSR = '" & usrName_UCase10 & "'"
If newY.TVAFACETA < 0 Then blnUpdate = True  ' records techniques

If blnUpdate Then
    
    xSQL = "update " & paramIBM_Library_SABSPE & ".YTVAFAC0" & xSet & xWhere
    Call FEU_ROUGE
    Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
    Call FEU_VERT
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYTVAFAC0_Update = "Erreur màj : " & newY.TVAFACFACN
        Exit Function
    End If
    
End If

Exit Function
Error_Handler:
    sqlYTVAFAC0_Update = Error
End Function



Public Function sqlYTVAFAC0_Init(newY As typeYTVAFAC0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String
Dim xxx As typeYTVAFAC0

On Error GoTo Error_Handler
sqlYTVAFAC0_Init = Null

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YTVAFAC0" & " where  TVAFACETA =  -1 and TVAFACFACN = 0"
Set rsAdo = cnSab_Update.Execute(xSQL, Nb)

xxx.TVAFACUPDS = rsAdo("TVAFACUPDS")
newY.TVAFACFACN = rsAdo("TVAFACDTR") + 1
newY.TVAFACUPDS = 0

' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where TVAFACETA = -1 and TVAFACFACN = 0" & " and TVAFACUPDS = " & xxx.TVAFACUPDS

xSet = " set TVAFACUPDS = " & xxx.TVAFACUPDS + 1 & " , TVAFACDTR = " & newY.TVAFACFACN


xSQL = "update " & paramIBM_Library_SABSPE & ".YTVAFAC0" & xSet & xWhere
Call FEU_ROUGE
Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYTVAFAC0_Init = "Erreur màj : " & newY.TVAFACETA
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYTVAFAC0_Init = Error
End Function

Public Function rsYTVAFAC0_GetBuffer(rsAdo As ADODB.Recordset, lYTVAFAC0 As typeYTVAFAC0)
On Error GoTo Error_Handler
rsYTVAFAC0_GetBuffer = Null

lYTVAFAC0.TVAFACETA = rsAdo("TVAFACETA")
lYTVAFAC0.TVAFACCLIC = rsAdo("TVAFACCLIC")
lYTVAFAC0.TVAFACCLI = rsAdo("TVAFACCLI")
lYTVAFAC0.TVAFACCLIP = rsAdo("TVAFACCLIP")
lYTVAFAC0.TVAFACCLIT = rsAdo("TVAFACCLIT")
lYTVAFAC0.TVAFACMTTC = rsAdo("TVAFACMTTC")
lYTVAFAC0.TVAFACMTVA = rsAdo("TVAFACMTVA")
lYTVAFAC0.TVAFACMEXO = rsAdo("TVAFACMEXO")
lYTVAFAC0.TVAFACFACN = rsAdo("TVAFACFACN")
lYTVAFAC0.TVAFACDTR = rsAdo("TVAFACDTR")
lYTVAFAC0.TVAFACSTA = rsAdo("TVAFACSTA")
lYTVAFAC0.TVAFACUPDS = rsAdo("TVAFACUPDS")
lYTVAFAC0.TVAFACUSR = rsAdo("TVAFACUSR")

Exit Function
Error_Handler:
rsYTVAFAC0_GetBuffer = Error


End Function

Public Function rsYTVAFAC0_Init(lYTVAFAC0 As typeYTVAFAC0)

lYTVAFAC0.TVAFACETA = 0      'établissement
lYTVAFAC0.TVAFACCLIC = ""     ' 1   'table client espace,G, D
lYTVAFAC0.TVAFACCLI = ""      ' 7   'code client
lYTVAFAC0.TVAFACCLIP = ""     ' 2   'pays de résidence
lYTVAFAC0.TVAFACCLIT = ""
lYTVAFAC0.TVAFACMTTC = 0      'montant en dev
lYTVAFAC0.TVAFACMTVA = 0     'montant euro
lYTVAFAC0.TVAFACMEXO = 0     'montant TVA issu du dossier
lYTVAFAC0.TVAFACFACN = 0         'n° facture
lYTVAFAC0.TVAFACDTR = 0      'date de traitement
lYTVAFAC0.TVAFACSTA = ""      ' 1   'statut
lYTVAFAC0.TVAFACUPDS = 0
End Function



