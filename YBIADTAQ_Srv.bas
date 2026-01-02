Attribute VB_Name = "srvYBIADTAQ"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim rsADO As ADODB.Recordset
 
Type typeYBIADTAQ
 
      BIADTAID    As Long
      
      BIADTASTA   As String
      BIADTAFCT   As String
      
      BIADTAIUSR   As String
      BIADTAIAMJ   As Long
      BIADTAIHMS   As Long
      
      BIADTAUAMJ   As Long
      BIADTAUHMS   As Long
      BIADTAUSEQ   As Long
      
      BIADTATXTE    As String
      BIADTATXTS    As String

    
'____________________________________________________ Journalisation
    JORCV                   As Long
    JOSEQN                  As Long
    JRNBIATRN               As Long
    
    JOENTT          As String * 2
    JODATE          As String * 6

'____________________________________________________ Journalisation
End Type
Public xYBIADTAQ As typeYBIADTAQ
Public Function sqlYBIADTAQ_Delete(oldY As typeYBIADTAQ)
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

On Error GoTo Error_Handler
sqlYBIADTAQ_Delete = Null
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where BIADTAID = " & oldY.BIADTAID _
       & " and BIADTAUSEQ = " & oldY.BIADTAUSEQ

'===================================================================================

    
    xSql = "delete from " & paramIBM_Library_SABSPE_XXX & ".YBIADTAQ" & xWhere
    Call FEU_ROUGE
    Set rsADO = cnSab_Update.Execute(xSql, Nb)
    Call FEU_VERT
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYBIADTAQ_Delete = "Erreur màj : " & xWhere
        Exit Function
    End If
    


Exit Function
Error_Handler:
    sqlYBIADTAQ_Delete = Error
End Function

Public Function sqlYBIADTAQ_Update(newY As typeYBIADTAQ, oldY As typeYBIADTAQ, blnUUSR As Boolean)
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

On Error GoTo Error_Handler
sqlYBIADTAQ_Update = Null

' Contrôle  : Même clé d'accès old / new
'===================================================================================
If oldY.BIADTAID <> newY.BIADTAID _
Or oldY.BIADTAUSEQ <> newY.BIADTAUSEQ Then
    sqlYBIADTAQ_Update = "Erreur BIADTAID : " & newY.BIADTAID
    Exit Function
End If
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where BIADTAID = " & oldY.BIADTAID _
       & " and BIADTAUSEQ = " & oldY.BIADTAUSEQ

newY.BIADTAUSEQ = newY.BIADTAUSEQ + 1
xSet = xSet & " set BIADTAUSEQ = " & newY.BIADTAUSEQ
blnUpdate = False

If blnUUSR Then
    newY.BIADTAUAMJ = DSys
    newY.BIADTAUHMS = time_Hms
End If

' Détecter les modifications
'===================================================================================
If newY.BIADTAIAMJ <> oldY.BIADTAIAMJ Then blnUpdate = True:  xSet = xSet & " , BIADTAIAMJ = " & newY.BIADTAIAMJ
If newY.BIADTAIHMS <> oldY.BIADTAIHMS Then blnUpdate = True:  xSet = xSet & " , BIADTAIHMS = " & newY.BIADTAIHMS
If newY.BIADTAUAMJ <> oldY.BIADTAUAMJ Then blnUpdate = True:  xSet = xSet & " , BIADTAUAMJ = " & newY.BIADTAUAMJ
If newY.BIADTAUHMS <> oldY.BIADTAUHMS Then blnUpdate = True:  xSet = xSet & " , BIADTAUHMS = " & newY.BIADTAUHMS

If newY.BIADTASTA <> oldY.BIADTASTA Then blnUpdate = True:  xSet = xSet & " , BIADTASTA = '" & newY.BIADTASTA & "'"
If newY.BIADTAIUSR <> oldY.BIADTAIUSR Then blnUpdate = True:  xSet = xSet & " , BIADTAIUSR = '" & newY.BIADTAIUSR & "'"
If newY.BIADTAFCT <> oldY.BIADTAFCT Then blnUpdate = True:  xSet = xSet & " , BIADTAFCT = '" & newY.BIADTAFCT & "'"
If newY.BIADTATXTE <> oldY.BIADTATXTE Then blnUpdate = True:  xSet = xSet & " , BIADTATXTE = '" & Replace(newY.BIADTATXTE, "'", "''") & "'"
If newY.BIADTATXTS <> oldY.BIADTATXTS Then blnUpdate = True:  xSet = xSet & " , BIADTATXTS = '" & Replace(newY.BIADTATXTS, "'", "''") & "'"


If blnUpdate Then
    
    xSql = "update " & paramIBM_Library_SABSPE_XXX & ".YBIADTAQ" & xSet & xWhere
    Call FEU_ROUGE
    Set rsADO = cnSab_Update.Execute(xSql, Nb)
    Call FEU_VERT
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYBIADTAQ_Update = "Erreur màj : " & newY.BIADTAID & "." & oldY.BIADTAIAMJ
        Exit Function
    End If
    
End If

Exit Function
Error_Handler:
    sqlYBIADTAQ_Update = Error
End Function

Public Function sqlYBIADTAQ_Insert(newY As typeYBIADTAQ)
Dim V
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

'On Error GoTo Error_Handler
sqlYBIADTAQ_Insert = Null
xSet = " (BIADTAID"
xValues = " values(" & newY.BIADTAID

newY.BIADTAIUSR = usrName_UCase
newY.BIADTAIAMJ = DSys
newY.BIADTAIHMS = time_Hms

' Détecter les modifications
'===================================================================================

If newY.BIADTAIAMJ <> 0 Then xSet = xSet & ",BIADTAIAMJ": xValues = xValues & " ," & newY.BIADTAIAMJ
If newY.BIADTAIHMS <> 0 Then xSet = xSet & ",BIADTAIHMS": xValues = xValues & " ," & newY.BIADTAIHMS
If newY.BIADTAUSEQ <> 0 Then xSet = xSet & ",BIADTAUSEQ": xValues = xValues & " ," & newY.BIADTAUSEQ
If newY.BIADTAUAMJ <> 0 Then xSet = xSet & ",BIADTAUAMJ": xValues = xValues & " ," & newY.BIADTAUAMJ
If newY.BIADTAUHMS <> 0 Then xSet = xSet & ",BIADTAUHMS": xValues = xValues & " ," & newY.BIADTAUHMS


If Trim(newY.BIADTASTA) <> "" Then xSet = xSet & ",BIADTASTA": xValues = xValues & " ,'" & newY.BIADTASTA & "'"

If Trim(newY.BIADTAIUSR) <> "" Then xSet = xSet & ",BIADTAIUSR": xValues = xValues & " ,'" & newY.BIADTAIUSR & "'"
If Trim(newY.BIADTAFCT) <> "" Then xSet = xSet & ",BIADTAFCT": xValues = xValues & " ,'" & newY.BIADTAFCT & "'"
If Trim(newY.BIADTATXTE) <> "" Then xSet = xSet & ",BIADTATXTE": xValues = xValues & " ,'" & Replace(newY.BIADTATXTE, "'", "''") & "'"
If Trim(newY.BIADTATXTS) <> "" Then xSet = xSet & ",BIADTATXTS": xValues = xValues & " ,'" & Replace(newY.BIADTATXTS, "'", "''") & "'"

xSql = "Insert into " & paramIBM_Library_SABSPE_XXX & ".YBIADTAQ" & xSet & ")" & xValues & ")"
Call FEU_ROUGE
Set rsADO = cnSab_Update.Execute(xSql, Nb)
Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYBIADTAQ_Insert = "Erreur màj : " & newY.BIADTAID
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYBIADTAQ_Insert = Error
End Function

Public Function rsYBIADTAQ_GetBuffer(rsADO As ADODB.Recordset, lYBIADTAQ As typeYBIADTAQ)
On Error GoTo Error_Handler
rsYBIADTAQ_GetBuffer = Null

lYBIADTAQ.JORCV = 0
lYBIADTAQ.JOSEQN = 0
lYBIADTAQ.JRNBIATRN = 0
lYBIADTAQ.JOENTT = ""
lYBIADTAQ.JODATE = ""

lYBIADTAQ.BIADTAID = rsADO("BIADTAID")
lYBIADTAQ.BIADTASTA = rsADO("BIADTASTA")
lYBIADTAQ.BIADTAFCT = rsADO("BIADTAFCT")

lYBIADTAQ.BIADTAIUSR = rsADO("BIADTAIUSR")
lYBIADTAQ.BIADTAIAMJ = rsADO("BIADTAIAMJ")
lYBIADTAQ.BIADTAIHMS = rsADO("BIADTAIHMS")
lYBIADTAQ.BIADTAUAMJ = rsADO("BIADTAUAMJ")
lYBIADTAQ.BIADTAUHMS = rsADO("BIADTAUHMS")
lYBIADTAQ.BIADTAUSEQ = rsADO("BIADTAUSEQ")

lYBIADTAQ.BIADTATXTE = rsADO("BIADTATXTE")
lYBIADTAQ.BIADTATXTS = rsADO("BIADTATXTS")

Exit Function
Error_Handler:
rsYBIADTAQ_GetBuffer = Error


End Function
'---------------------------------------------------------
Public Function rsJBIADTA0_GetBuffer(rsADO As ADODB.Recordset, rsYBIADTAQ As typeYBIADTAQ)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsJBIADTA0_GetBuffer = Null

rsJBIADTA0_GetBuffer = rsYBIADTAQ_GetBuffer(rsADO, rsYBIADTAQ)
rsYBIADTAQ.JORCV = rsADO("JORCV")
rsYBIADTAQ.JOSEQN = rsADO("JOSEQN")
rsYBIADTAQ.JRNBIATRN = rsADO("JRNBIATRN")
rsYBIADTAQ.JOENTT = rsADO("JOENTT")
rsYBIADTAQ.JODATE = rsADO("JODATE")

Exit Function

Error_Handler:

rsJBIADTA0_GetBuffer = Error

End Function


Public Function rsYBIADTAQ_Init(lYBIADTAQ As typeYBIADTAQ)


lYBIADTAQ.BIADTAID = 0
lYBIADTAQ.BIADTAIAMJ = 0
      
lYBIADTAQ.BIADTAIHMS = 0
lYBIADTAQ.BIADTASTA = ""
lYBIADTAQ.BIADTAIUSR = ""
lYBIADTAQ.BIADTAUAMJ = 0
lYBIADTAQ.BIADTAUHMS = 0
lYBIADTAQ.BIADTAUSEQ = 0
      
lYBIADTAQ.BIADTAFCT = ""
lYBIADTAQ.BIADTATXTE = ""
lYBIADTAQ.BIADTATXTS = ""




lYBIADTAQ.JORCV = 0
lYBIADTAQ.JOSEQN = 0
lYBIADTAQ.JRNBIATRN = 0
    
lYBIADTAQ.JOENTT = ""
lYBIADTAQ.JODATE = ""

End Function












Public Function sqlYBIADTAQ_BIADTAID(lYBIADTAQ As typeYBIADTAQ)
Static mBIADTAID As Long
Dim xSql As String, K As Long
Dim V, blnOk As Boolean
'On Error GoTo Error_Handler

sqlYBIADTAQ_BIADTAID = Null
blnOk = False
If mBIADTAID = 0 Then
    xSql = "select count(*) as Tally from " & paramIBM_Library_SABSPE & ".YBIADTAQ "
    Set rsADO = cnsab.Execute(xSql)
    mBIADTAID = rsADO("Tally")
End If

If mBIADTAID > 0 Then
    Do
        xSql = "select BIADTAID from " & paramIBM_Library_SABSPE & ".YBIADTAQ" _
             & " where BIADTAID >= " & mBIADTAID _
             & " order by BIADTAID desc"
        Set rsADO = cnsab.Execute(xSql)
        If Not rsADO.EOF Then
            mBIADTAID = rsADO("BIADTAID"): blnOk = True
        Else
            If mBIADTAID < 0 Then sqlYBIADTAQ_BIADTAID = "sqlYBIADTAQ_Request : erreur BIATABID": Exit Function
            mBIADTAID = mBIADTAID - 10
        End If
    Loop Until blnOk = True
End If

'==============================================================================

Call rsYBIADTAQ_Init(lYBIADTAQ)
mBIADTAID = mBIADTAID + 1
lYBIADTAQ.BIADTAID = mBIADTAID

'==============================================================================

Exit Function

Error_Handler:
    sqlYBIADTAQ_BIADTAID = Error

End Function
Public Function sqlYBIADTAQ_BIADTASTA(lYBIADTAQ As typeYBIADTAQ, lProgressBar1 As ProgressBar)
Dim xSql As String, K As Long
Dim V, blnExit As Boolean
On Error GoTo Error_Handler


sqlYBIADTAQ_BIADTASTA = Null
blnExit = False
Do
    xSql = "select * from " & paramIBM_Library_SABSPE & ".YBIADTAQ" _
         & " where BIADTAID = " & lYBIADTAQ.BIADTAID
    Set rsADO = cnsab.Execute(xSql)
    If rsADO.EOF Then
        sqlYBIADTAQ_BIADTASTA = lYBIADTAQ.BIADTAFCT & " " & lYBIADTAQ.BIADTAID & " inconnu dans YBIADTAQ"
        blnExit = True
    Else
        Select Case rsADO("BIADTASTA")
            Case "V": blnExit = True
            Case "E":
                    sqlYBIADTAQ_BIADTASTA = lYBIADTAQ.BIADTAFCT & " " & lYBIADTAQ.BIADTAID & " en ERREUR dans YBIADTAQ"
                    blnExit = True
            Case " ": K = K + 1
                lProgressBar1.value = K * (K - 1)
                If K > 5 Then
                    sqlYBIADTAQ_BIADTASTA = lYBIADTAQ.BIADTAFCT & " " & lYBIADTAQ.BIADTAID & " en ATTENTE dans YBIADTAQ"
                    blnExit = True
                Else
                    Call lstErr_AddItem(frmYGOSDOS0.lstErr, frmYGOSDOS0.cmdContext, "attente SYNCHRONISATION YBIADTAQ " & K): DoEvents

                    Wait_SS K
                End If
        End Select
    End If
Loop Until blnExit = True

'==============================================================================


'==============================================================================

Exit Function

Error_Handler:
    sqlYBIADTAQ_BIADTASTA = Error

End Function


