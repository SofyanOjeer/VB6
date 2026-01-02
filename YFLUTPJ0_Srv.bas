Attribute VB_Name = "srvYFLUTPJ0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim rsAdo As ADODB.Recordset
 
Type typeYFLUTPJ0
    FLUTPJID     As Long
    FLUTPJCCB    As Long
    FLUTPJORIG   As String
    FLUTPJSTA  As String
    
    FLUTPJETB    As Integer
    FLUTPJAGE    As Integer
    FLUTPJSER    As String
    FLUTPJSSE    As String
    FLUTPJOPE   As String
    FLUTPJNAT   As String
    FLUTPJDOS   As Long
    FLUTPJDOSQ  As Long
    FLUTPJEVE    As String
    FLUTPJECH   As Long
    FLUTPJMTD    As Currency
    FLUTPJDEV    As String
    
End Type

Type typeYFLUTPJ1
    FLUTPJDOS   As Long
    FLUTPJDOSQ  As Long
    FLUTPJCLI   As String
    FLUTPJTXT   As String
    
End Type

Public Function sqlYFLUTPJ0_Delete(oldY As typeYFLUTPJ0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

On Error GoTo Error_Handler
sqlYFLUTPJ0_Delete = Null
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where FLUTPJID = " & oldY.FLUTPJID

'===================================================================================

    
    xSQL = "delete from " & paramIBM_Library_SABSPE & ".YFLUTPJ0" & xWhere
    Call FEU_ROUGE
    Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
    Call FEU_VERT
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYFLUTPJ0_Delete = "Erreur màj : " & xWhere
        Exit Function
    End If
    


Exit Function
Error_Handler:
    sqlYFLUTPJ0_Delete = Error
End Function


Public Function sqlYFLUTPJ1_Delete(oldY As typeYFLUTPJ1)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

On Error GoTo Error_Handler
sqlYFLUTPJ1_Delete = Null
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where FLUTPJDOS = " & oldY.FLUTPJDOS & " and FLUTPJDOSQ = " & oldY.FLUTPJDOSQ

'===================================================================================

    
    xSQL = "delete from " & paramIBM_Library_SABSPE & ".YFLUTPJ1" & xWhere
    Call FEU_ROUGE
    Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
    Call FEU_VERT
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYFLUTPJ1_Delete = "Erreur màj : " & xWhere
        Exit Function
    End If
    


Exit Function
Error_Handler:
    sqlYFLUTPJ1_Delete = Error
End Function


Public Function sqlYFLUTPJ1_Delete_FLUTPJDOS(oldY As typeYFLUTPJ1)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

On Error GoTo Error_Handler
sqlYFLUTPJ1_Delete_FLUTPJDOS = Null
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where FLUTPJDOS = " & oldY.FLUTPJDOS
'===================================================================================

    
    xSQL = "delete from " & paramIBM_Library_SABSPE & ".YFLUTPJ1" & xWhere
    Call FEU_ROUGE
    Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
    Call FEU_VERT
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYFLUTPJ1_Delete_FLUTPJDOS = "Erreur màj : " & xWhere
        Exit Function
    End If
    


Exit Function
Error_Handler:
    sqlYFLUTPJ1_Delete_FLUTPJDOS = Error
End Function

'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsYFLUTPJ0_GetBuffer(rsAdo As ADODB.Recordset, rsYFLUTPJ0 As typeYFLUTPJ0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsYFLUTPJ0_GetBuffer = Null


rsYFLUTPJ0.FLUTPJID = rsAdo("FLUTPJID")
rsYFLUTPJ0.FLUTPJCCB = rsAdo("FLUTPJCCB")
rsYFLUTPJ0.FLUTPJORIG = rsAdo("FLUTPJORIG")
rsYFLUTPJ0.FLUTPJSTA = rsAdo("FLUTPJSTA")
rsYFLUTPJ0.FLUTPJETB = rsAdo("FLUTPJETB")
rsYFLUTPJ0.FLUTPJAGE = rsAdo("FLUTPJAGE")
rsYFLUTPJ0.FLUTPJSER = rsAdo("FLUTPJSER")
rsYFLUTPJ0.FLUTPJSSE = rsAdo("FLUTPJSSE")
rsYFLUTPJ0.FLUTPJOPE = rsAdo("FLUTPJOPE")
rsYFLUTPJ0.FLUTPJNAT = rsAdo("FLUTPJNAT")
rsYFLUTPJ0.FLUTPJDOS = rsAdo("FLUTPJDOS")
rsYFLUTPJ0.FLUTPJDOSQ = rsAdo("FLUTPJDOSQ")
rsYFLUTPJ0.FLUTPJEVE = rsAdo("FLUTPJEVE")
rsYFLUTPJ0.FLUTPJECH = rsAdo("FLUTPJECH")
rsYFLUTPJ0.FLUTPJMTD = rsAdo("FLUTPJMTD")
rsYFLUTPJ0.FLUTPJDEV = rsAdo("FLUTPJDEV")

Exit Function

Error_Handler:

rsYFLUTPJ0_GetBuffer = Error

End Function

'---------------------------------------------------------
Public Function rsYFLUTPJ1_GetBuffer(rsAdo As ADODB.Recordset, rsYFLUTPJ1 As typeYFLUTPJ1)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsYFLUTPJ1_GetBuffer = Null


rsYFLUTPJ1.FLUTPJDOS = rsAdo("FLUTPJDOS")
rsYFLUTPJ1.FLUTPJDOSQ = rsAdo("FLUTPJDOSQ")
rsYFLUTPJ1.FLUTPJCLI = rsAdo("FLUTPJCLI")
rsYFLUTPJ1.FLUTPJTXT = rsAdo("FLUTPJTXT")

Exit Function

Error_Handler:

rsYFLUTPJ1_GetBuffer = Error

End Function









Public Sub rsYFLUTPJ0_Init(lYFLUTPJ0 As typeYFLUTPJ0)
lYFLUTPJ0.FLUTPJID = 0
lYFLUTPJ0.FLUTPJCCB = 0
lYFLUTPJ0.FLUTPJORIG = ""
lYFLUTPJ0.FLUTPJSTA = ""

lYFLUTPJ0.FLUTPJETB = 1
lYFLUTPJ0.FLUTPJAGE = 1
lYFLUTPJ0.FLUTPJSER = ""
lYFLUTPJ0.FLUTPJSSE = ""
lYFLUTPJ0.FLUTPJOPE = ""
lYFLUTPJ0.FLUTPJNAT = ""
lYFLUTPJ0.FLUTPJDOS = 0
lYFLUTPJ0.FLUTPJDOSQ = 0
lYFLUTPJ0.FLUTPJEVE = ""
lYFLUTPJ0.FLUTPJECH = 0
lYFLUTPJ0.FLUTPJMTD = 0
lYFLUTPJ0.FLUTPJDEV = ""

End Sub

Public Sub rsYFLUTPJ1_Init(lYFLUTPJ1 As typeYFLUTPJ1)
lYFLUTPJ1.FLUTPJDOS = 0
lYFLUTPJ1.FLUTPJDOSQ = 0
lYFLUTPJ1.FLUTPJCLI = ""
lYFLUTPJ1.FLUTPJTXT = ""

End Sub


Public Function sqlYFLUTPJ0_Insert(newY As typeYFLUTPJ0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlYFLUTPJ0_Insert = Null

xSet = " (FLUTPJID"
xValues = " values(" & newY.FLUTPJID

' Insertion :
'===================================================================================
If newY.FLUTPJCCB <> 0 Then xSet = xSet & ",FLUTPJCCB": xValues = xValues & ", " & newY.FLUTPJCCB
If newY.FLUTPJETB <> 0 Then xSet = xSet & ",FLUTPJETB": xValues = xValues & ", " & newY.FLUTPJETB
If newY.FLUTPJAGE <> 0 Then xSet = xSet & ",FLUTPJAGE": xValues = xValues & ", " & newY.FLUTPJAGE
If newY.FLUTPJDOS <> 0 Then xSet = xSet & ",FLUTPJDOS": xValues = xValues & ", " & newY.FLUTPJDOS
If newY.FLUTPJDOSQ <> 0 Then xSet = xSet & ",FLUTPJDOSQ": xValues = xValues & ", " & newY.FLUTPJDOSQ
If newY.FLUTPJECH <> 0 Then xSet = xSet & ",FLUTPJECH": xValues = xValues & ", " & newY.FLUTPJECH
If newY.FLUTPJMTD <> 0 Then xSet = xSet & ",FLUTPJMTD": xValues = xValues & ", " & cur_P(newY.FLUTPJMTD)

'===================================================================================

If Trim(newY.FLUTPJORIG) <> "" Then xSet = xSet & ",FLUTPJORIG": xValues = xValues & ", '" & Replace(Trim(newY.FLUTPJORIG), "'", "''") & "'"
If Trim(newY.FLUTPJSTA) <> "" Then xSet = xSet & ",FLUTPJSTA": xValues = xValues & ", '" & Replace(Trim(newY.FLUTPJSTA), "'", "''") & "'"
If Trim(newY.FLUTPJSER) <> "" Then xSet = xSet & ",FLUTPJSER": xValues = xValues & ", '" & Replace(Trim(newY.FLUTPJSER), "'", "''") & "'"
If Trim(newY.FLUTPJSSE) <> "" Then xSet = xSet & ",FLUTPJSSE": xValues = xValues & ", '" & Replace(Trim(newY.FLUTPJSSE), "'", "''") & "'"
If Trim(newY.FLUTPJOPE) <> "" Then xSet = xSet & ",FLUTPJOPE": xValues = xValues & ", '" & Replace(Trim(newY.FLUTPJOPE), "'", "''") & "'"
If Trim(newY.FLUTPJNAT) <> "" Then xSet = xSet & ",FLUTPJNAT": xValues = xValues & ", '" & Replace(Trim(newY.FLUTPJNAT), "'", "''") & "'"
If Trim(newY.FLUTPJEVE) <> "" Then xSet = xSet & ",FLUTPJEVE": xValues = xValues & ", '" & Replace(Trim(newY.FLUTPJEVE), "'", "''") & "'"
If Trim(newY.FLUTPJDEV) <> "" Then xSet = xSet & ",FLUTPJDEV": xValues = xValues & ", '" & Replace(Trim(newY.FLUTPJDEV), "'", "''") & "'"

xSQL = "Insert into " & paramIBM_Library_SABSPE_XXX & ".YFLUTPJ0" & xSet & ")" & xValues & ")"
Call FEU_ROUGE
Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYFLUTPJ0_Insert = "Erreur màj : " & newY.FLUTPJID
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYFLUTPJ0_Insert = Error
End Function

Public Function sqlYFLUTPJ1_Insert(newY As typeYFLUTPJ1)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlYFLUTPJ1_Insert = Null

xSet = " (FLUTPJDOS , FLUTPJDOSQ"
xValues = " values(" & newY.FLUTPJDOS & ", " & newY.FLUTPJDOSQ

' Insertion :
'===================================================================================

If Trim(newY.FLUTPJCLI) <> "" Then xSet = xSet & ",FLUTPJCLI": xValues = xValues & ", '" & Replace(Trim(newY.FLUTPJCLI), "'", "''") & "'"
If Trim(newY.FLUTPJTXT) <> "" Then xSet = xSet & ",FLUTPJTXT": xValues = xValues & ", '" & Replace(Trim(newY.FLUTPJTXT), "'", "''") & "'"

xSQL = "Insert into " & paramIBM_Library_SABSPE_XXX & ".YFLUTPJ1" & xSet & ")" & xValues & ")"
Call FEU_ROUGE
Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYFLUTPJ1_Insert = "Erreur màj : " & newY.FLUTPJDOS
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYFLUTPJ1_Insert = Error
End Function


Public Function sqlYFLUTPJ0_Update(newY As typeYFLUTPJ0, oldY As typeYFLUTPJ0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean
Dim K As Integer

On Error GoTo Error_Handler
sqlYFLUTPJ0_Update = Null

' Contrôle  : Même clé d'accès old / new
'===================================================================================
If oldY.FLUTPJID <> newY.FLUTPJID Then
    sqlYFLUTPJ0_Update = "Erreur FLUTPJID : " & newY.FLUTPJID
    Exit Function
End If
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where FLUTPJID = " & oldY.FLUTPJID

xSet = " set "
blnUpdate = False


' Détecter les modifications
'===================================================================================
If newY.FLUTPJCCB <> oldY.FLUTPJCCB Then blnUpdate = True:  xSet = xSet & " , FLUTPJCCB = " & newY.FLUTPJCCB
If newY.FLUTPJORIG <> oldY.FLUTPJORIG Then blnUpdate = True:  xSet = xSet & " , FLUTPJORIG = '" & newY.FLUTPJORIG & "'"
If newY.FLUTPJSTA <> oldY.FLUTPJSTA Then blnUpdate = True:  xSet = xSet & " , FLUTPJSTA = '" & newY.FLUTPJSTA & "'"

If newY.FLUTPJETB <> oldY.FLUTPJETB Then blnUpdate = True:  xSet = xSet & " , FLUTPJETB = " & newY.FLUTPJETB
If newY.FLUTPJAGE <> oldY.FLUTPJAGE Then blnUpdate = True:  xSet = xSet & " , FLUTPJAGE = " & newY.FLUTPJAGE
If newY.FLUTPJSER <> oldY.FLUTPJSER Then blnUpdate = True:  xSet = xSet & " , FLUTPJSER = '" & newY.FLUTPJSER & "'"
If newY.FLUTPJSSE <> oldY.FLUTPJSSE Then blnUpdate = True:  xSet = xSet & " , FLUTPJSSE = '" & newY.FLUTPJSSE & "'"
If newY.FLUTPJOPE <> oldY.FLUTPJOPE Then blnUpdate = True:  xSet = xSet & " , FLUTPJOPE = '" & newY.FLUTPJOPE & "'"
If newY.FLUTPJNAT <> oldY.FLUTPJNAT Then blnUpdate = True:  xSet = xSet & " , FLUTPJNAT = '" & newY.FLUTPJNAT & "'"
If newY.FLUTPJDOS <> oldY.FLUTPJDOS Then blnUpdate = True:  xSet = xSet & " , FLUTPJDOS = " & newY.FLUTPJDOS
If newY.FLUTPJDOSQ <> oldY.FLUTPJDOSQ Then blnUpdate = True:  xSet = xSet & " , FLUTPJDOSQ = " & newY.FLUTPJDOSQ
If newY.FLUTPJEVE <> oldY.FLUTPJEVE Then blnUpdate = True:  xSet = xSet & " , FLUTPJEVE = '" & newY.FLUTPJEVE & "'"
If newY.FLUTPJECH <> oldY.FLUTPJECH Then blnUpdate = True:  xSet = xSet & " , FLUTPJECH = " & newY.FLUTPJECH
If newY.FLUTPJMTD <> oldY.FLUTPJMTD Then blnUpdate = True:  xSet = xSet & " , FLUTPJMTD = " & cur_P(newY.FLUTPJMTD)
If newY.FLUTPJDEV <> oldY.FLUTPJDEV Then blnUpdate = True:  xSet = xSet & " , FLUTPJDEV = '" & newY.FLUTPJDEV & "'"


If blnUpdate Then
    K = InStr(xSet, ",")
    If K > 0 Then Mid$(xSet, K, 1) = " "
    xSQL = "update " & paramIBM_Library_SABSPE_XXX & ".YFLUTPJ0" & xSet & xWhere
    Call FEU_ROUGE
    Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
    Call FEU_VERT
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYFLUTPJ0_Update = "Erreur màj : " & newY.FLUTPJID
        Exit Function
    End If
    
End If

Exit Function
Error_Handler:
    sqlYFLUTPJ0_Update = Error
End Function

Public Function sqlYFLUTPJ1_Update(newY As typeYFLUTPJ1, oldY As typeYFLUTPJ1)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean
Dim K As Integer

On Error GoTo Error_Handler
sqlYFLUTPJ1_Update = Null

' Contrôle  : Même clé d'accès old / new
'===================================================================================
If oldY.FLUTPJDOS <> newY.FLUTPJDOS Then
    sqlYFLUTPJ1_Update = "Erreur FLUTPJdos : " & newY.FLUTPJDOS
    Exit Function
End If
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where FLUTPJDOS = " & oldY.FLUTPJDOS & " and FLUTPJDOSQ = " & oldY.FLUTPJDOSQ

xSet = " set "
blnUpdate = False


' Détecter les modifications
'===================================================================================
If newY.FLUTPJCLI <> oldY.FLUTPJCLI Then blnUpdate = True:  xSet = xSet & " , FLUTPJCLI = '" & newY.FLUTPJCLI & "'"
If newY.FLUTPJTXT <> oldY.FLUTPJTXT Then blnUpdate = True:  xSet = xSet & " , FLUTPJTXT = '" & Replace(Trim(newY.FLUTPJTXT), "'", "''") & "'"


If blnUpdate Then
    K = InStr(xSet, ",")
    If K > 0 Then Mid$(xSet, K, 1) = " "
    xSQL = "update " & paramIBM_Library_SABSPE_XXX & ".YFLUTPJ1" & xSet & xWhere
    Call FEU_ROUGE
    Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
    Call FEU_VERT
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYFLUTPJ1_Update = "Erreur màj : " & newY.FLUTPJDOS & " -  " & newY.FLUTPJDOSQ
        Exit Function
    End If
    
End If

Exit Function
Error_Handler:
    sqlYFLUTPJ1_Update = Error
End Function







