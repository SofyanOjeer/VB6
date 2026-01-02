Attribute VB_Name = "srvYEICGCCLOG"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim rsAdo As ADODB.Recordset
 
Type typeYEICGCCLOG

    EICGCCLOGD   As Long        ' DATE maj
    EICGCCLOGH   As Long        ' heure maj
    EICGCCLOGU   As String * 10 ' utilisateur maj
    EICGCCLOGS   As Long        ' N° séquence (info)
    EICGCCLOGK   As String * 12 ' code action
    EICGCCLOGI   As Long        ' identifiant
    EICGCCLOGA   As String * 1  ' statut
    EICGCCLOGE   As Long        ' Echéance
    EICGCCLOGX   As String * 64 ' commentaire
    
    
'____________________________________________________ Journalisation
    JORCV                   As Long
    JOSEQN                  As Long
    JRNBIATRN               As Long
    JOENTT          As String * 2
    JODATE          As String * 6
'____________________________________________________ Journalisation
End Type
Public Function sqlYEICGCCLOG_Insert(newY As typeYEICGCCLOG)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlYEICGCCLOG_Insert = Null

xSet = " (EICGCCLOGS"
xValues = " values(" & newY.EICGCCLOGS

' Insertion :
'===================================================================================
If newY.EICGCCLOGD <> 0 Then xSet = xSet & ",EICGCCLOGD": xValues = xValues & ", " & newY.EICGCCLOGD
If newY.EICGCCLOGH <> 0 Then xSet = xSet & ",EICGCCLOGH": xValues = xValues & ", " & newY.EICGCCLOGH
If newY.EICGCCLOGI <> 0 Then xSet = xSet & ",EICGCCLOGI": xValues = xValues & ", " & newY.EICGCCLOGI
If newY.EICGCCLOGE <> 0 Then xSet = xSet & ",EICGCCLOGE": xValues = xValues & ", " & newY.EICGCCLOGE

'===================================================================================

If Trim(newY.EICGCCLOGK) <> "" Then xSet = xSet & ",EICGCCLOGK": xValues = xValues & ", '" & Replace(Trim(newY.EICGCCLOGK), "'", "''") & "'"
If Trim(newY.EICGCCLOGA) <> "" Then xSet = xSet & ",EICGCCLOGA": xValues = xValues & ", '" & Replace(Trim(newY.EICGCCLOGA), "'", "''") & "'"
If Trim(newY.EICGCCLOGX) <> "" Then xSet = xSet & ",EICGCCLOGX": xValues = xValues & ", '" & Replace(Trim(newY.EICGCCLOGX), "'", "''") & "'"
If newY.EICGCCLOGU <> "" Then xSet = xSet & ",EICGCCLOGU": xValues = xValues & ", '" & Replace(newY.EICGCCLOGU, "'", "''") & "'"
Call FEU_ROUGE
xSQL = "Insert into " & paramIBM_Library_SABSPE_XXX & ".YEICGCCLOG" & xSet & ")" & xValues & ")"

Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYEICGCCLOG_Insert = "Erreur màj : " & newY.EICGCCLOGX & " " & newY.EICGCCLOGS
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYEICGCCLOG_Insert = Error
End Function
Public Function sqlYEICGCCLOG_Update(newY As typeYEICGCCLOG, oldY As typeYEICGCCLOG)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean
Dim K As Integer

On Error GoTo Error_Handler
sqlYEICGCCLOG_Update = Null

' Contrôle  : Même clé d'accès old / new
'===================================================================================
If oldY.EICGCCLOGD <> newY.EICGCCLOGD _
Or oldY.EICGCCLOGH <> newY.EICGCCLOGH _
Or oldY.EICGCCLOGU <> newY.EICGCCLOGU _
Or oldY.EICGCCLOGS <> newY.EICGCCLOGS Then
    sqlYEICGCCLOG_Update = "Erreur EICGCCLOGD : " & newY.EICGCCLOGD & "." & oldY.EICGCCLOGS
    Exit Function
End If
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where EICGCCLOGD = " & oldY.EICGCCLOGD _
       & " and EICGCCLOGH = " & oldY.EICGCCLOGH _
       & " and EICGCCLOGU = '" & oldY.EICGCCLOGU & "'" _
       & " and EICGCCLOGS = " & oldY.EICGCCLOGS

xSet = " set "
blnUpdate = False


' Détecter les modifications
'===================================================================================
If newY.EICGCCLOGA <> oldY.EICGCCLOGA Then blnUpdate = True:  xSet = xSet & " , EICGCCLOGA = '" & newY.EICGCCLOGA & "'"
If newY.EICGCCLOGK <> oldY.EICGCCLOGK Then blnUpdate = True:  xSet = xSet & " , EICGCCLOGK = '" & newY.EICGCCLOGK & "'"
If newY.EICGCCLOGX <> oldY.EICGCCLOGX Then blnUpdate = True:  xSet = xSet & " , EICGCCLOGX = '" & Replace(Trim(newY.EICGCCLOGX), "'", "''") & "'"


If newY.EICGCCLOGI <> oldY.EICGCCLOGI Then blnUpdate = True:  xSet = xSet & " , EICGCCLOGI = " & newY.EICGCCLOGI
If newY.EICGCCLOGE <> oldY.EICGCCLOGE Then blnUpdate = True:  xSet = xSet & " , EICGCCLOGE = " & newY.EICGCCLOGE

If blnUpdate Then
    K = InStr(xSet, ",")
    If K > 0 Then Mid$(xSet, K, 1) = " "
    xSQL = "update " & paramIBM_Library_SABSPE_XXX & ".YEICGCCLOG" & xSet & xWhere
    Call FEU_ROUGE
    Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
    Call FEU_VERT
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYEICGCCLOG_Update = "Erreur màj : " & newY.EICGCCLOGD
        Exit Function
    End If
    
End If

Exit Function
Error_Handler:
    sqlYEICGCCLOG_Update = Error
End Function


'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsYEICGCCLOG_GetBuffer(rsAdo As ADODB.Recordset, rsYEICGCCLOG As typeYEICGCCLOG)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsYEICGCCLOG_GetBuffer = Null
rsYEICGCCLOG.JORCV = 0
rsYEICGCCLOG.JOSEQN = 0
rsYEICGCCLOG.JRNBIATRN = 0
rsYEICGCCLOG.JOENTT = ""
rsYEICGCCLOG.JODATE = ""

rsYEICGCCLOG.EICGCCLOGD = rsAdo("EICGCCLOGD")
rsYEICGCCLOG.EICGCCLOGH = rsAdo("EICGCCLOGH")
rsYEICGCCLOG.EICGCCLOGU = rsAdo("EICGCCLOGU")
rsYEICGCCLOG.EICGCCLOGS = rsAdo("EICGCCLOGS")
rsYEICGCCLOG.EICGCCLOGK = rsAdo("EICGCCLOGK")
rsYEICGCCLOG.EICGCCLOGI = rsAdo("EICGCCLOGI")
rsYEICGCCLOG.EICGCCLOGA = rsAdo("EICGCCLOGA")
rsYEICGCCLOG.EICGCCLOGE = rsAdo("EICGCCLOGE")
rsYEICGCCLOG.EICGCCLOGX = rsAdo("EICGCCLOGX")

Exit Function

Error_Handler:

rsYEICGCCLOG_GetBuffer = Error

End Function

'---------------------------------------------------------
Public Function rsJEICGCCLOG_GetBuffer(rsAdo As ADODB.Recordset, rsYEICGCCLOG As typeYEICGCCLOG)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsJEICGCCLOG_GetBuffer = Null

rsJEICGCCLOG_GetBuffer = rsYEICGCCLOG_GetBuffer(rsAdo, rsYEICGCCLOG)
rsYEICGCCLOG.JORCV = rsAdo("JORCV")
rsYEICGCCLOG.JOSEQN = rsAdo("JOSEQN")
rsYEICGCCLOG.JRNBIATRN = rsAdo("JRNBIATRN")
rsYEICGCCLOG.JOENTT = rsAdo("JOENTT")
rsYEICGCCLOG.JODATE = rsAdo("JODATE")

Exit Function

Error_Handler:

rsJEICGCCLOG_GetBuffer = Error

End Function










Public Sub rsYEICGCCLOG_Init(lYEICGCCLOG As typeYEICGCCLOG)
lYEICGCCLOG.EICGCCLOGD = DSys                ' DATE maj
lYEICGCCLOG.EICGCCLOGH = time_Hms            ' heure maj
lYEICGCCLOG.EICGCCLOGU = usrName_UCase       ' utilisateur maj
''''''''''''''lYEICGCCLOG.EICGCCLOGS = 0
lYEICGCCLOG.EICGCCLOGI = 0                  ' chèque
lYEICGCCLOG.EICGCCLOGE = 0
lYEICGCCLOG.EICGCCLOGA = ""                   ' statut
lYEICGCCLOG.EICGCCLOGX = ""                   ' commentaire

End Sub
