Attribute VB_Name = "srvYSWISAB0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim rsSabX As New ADODB.Recordset


Public SWISABWSTA_appe_date_time As String
Public SAA_Alerte_Amount As String, SAA_Alerte_Approval As String, SAA_Alerte_Routage As String
Public SAA_Alerte_cours_USD As Double, minSAA_Alerte_Approval As String

Public SAA_Alerte_Live_Entrant As typeUMID
Public SAA_Alerte_Live_Sortant As typeUMID
Public SAA_Alerte_Jrnl As typeUMID
Public SAA_Alerte_Origine_MT As Long
Public SAA_Alerte_SWIHIADEN As Long, SAA_Alerte_SWIHIADEN_SSS As Double
Public lastSWIHIADEN As Long
Public SAA_Alerte_Modification_MT As String

Public last_Jrnl_date_time_ES As Date, last_Jrnl_date_time_EVE As Date, last_mesg_crea_date_time As Date
Public last_Alerte_date_time_ES As Date, last_Alerte_date_time_EVE As Date

Public last_Alerte_Loop_ES As Long, last_Alerte_Loop_EVE As Long
Public blnAlerte_date_time_ES As Boolean, blnAlerte_date_time_EVE As Boolean


Public arrMT_Field() As String
Public arrMT_Field_Code_Nb As Integer, arrMT_Field_Code() As String, arrMT_Field_Lib() As String

Public arrMT_Nb As Integer, arrMT_Code() As String, arrMT_Lib() As String

Type typeUMID
      Aid     As Double
      Umidl   As Double
      Umidh   As Double
End Type

Dim rsADO As ADODB.Recordset
Type typeYSWISAB0
 
      SWISABSWID   As Long
      
      SWISABWID1   As Long
      SWISABWIDL   As Long
      SWISABWIDH   As Long
      
      SWISABWES    As String
      SWISABWMTK   As String
      SWISABWBIC   As String
      SWISABWDEV   As String
      SWISABWMTD   As Currency
      SWISABWN20   As String
      SWISABWL20   As String
      SWISABWSRV   As String
      SWISABWSTA   As String
      SWISABWAMJ   As Long
      SWISABWHMS   As Long
      
      SWISABSER    As String
      SWISABSSE    As String
      SWISABOPEC   As String
      SWISABOPEN   As Long
      SWISABZSWI   As Long
      SWISABSTA    As String
      
      SWISABXAMJ   As Long
      SWISABXHMS   As Long
      SWISABXGOS   As String
      SWISABXEVE   As String
      SWISABKPDE   As String
      SWISABK20    As String
      SWISABK999   As String
      SWISABKSRV   As String
    
'____________________________________________________ Journalisation
    JORCV                   As Long
    JOSEQN                  As Long
    JRNBIATRN               As Long
    
    JOENTT          As String * 2
    JODATE          As String * 6

'____________________________________________________ Journalisation
End Type
Public xYSWISAB0 As typeYSWISAB0
Public Sub arrMT_Load()
Dim xSql As String, X As String
'______________________________________________________________________________________________________
xSql = "select count(*) from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'SAA'" _
     & " and BIATABK1 = 'MT_Type'"
Set rsSabX = cnsab.Execute(xSql)

ReDim arrMT_Code(rsSabX(0) + 1), arrMT_Lib(rsSabX(0) + 1)

xSql = "select *from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'SAA'" _
     & " and BIATABK1 = 'MT_Type' order by BIATABK2"
Set rsSabX = cnsab.Execute(xSql)
arrMT_Nb = 0
Do While Not rsSabX.EOF
    arrMT_Nb = arrMT_Nb + 1
    arrMT_Code(arrMT_Nb) = Trim(rsSabX("BIATABK2"))
    arrMT_Lib(arrMT_Nb) = Trim(rsSabX("BIATABTXT"))
    
    rsSabX.MoveNext
Loop


End Sub
Public Function arrMT_Type_Scan(lType As String) As String
Static K As Integer
Dim X As String

If lType <> arrMT_Code(K) Then

    For K = 1 To arrMT_Nb
        If lType = arrMT_Code(K) Then
            X = arrMT_Lib(K)
            Call arrMT_Fields_Load(lType)

            Exit For
        End If
    Next K
End If

arrMT_Type_Scan = X

End Function

Public Function arrMT_Fields_Scan(lField As String) As String
Dim K As Integer, X As String
On Error GoTo Error_Handler

X = arrMT_Field(Val(Mid$(lField, 1, 2)))

If Len(lField) > 2 Then
    For K = 1 To arrMT_Field_Code_Nb
        If lField = arrMT_Field_Code(K) Then
            X = arrMT_Field_Lib(K)
            Exit For
        End If
    Next K
    
End If

arrMT_Fields_Scan = X
Exit Function

'------------------------------------------
Error_Handler:
    Dim V
    V = Error
Error_MsgBox:
    'If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : cmdYswisab0_Update"
    arrMT_Fields_Scan = ""
End Function

Public Sub arrMT_Fields_Load(lMT_Type As String)
Dim xSql As String, X As String, blnOk As Boolean

X = "MT_" & lMT_Type
'______________________________________________________________________________________________________
xSql = "select count(*) from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'SAA'" _
     & " and BIATABK1 = '" & X & "'"
Set rsSabX = cnsab.Execute(xSql)

If rsSabX(0) = 0 Then
    X = "MT_Fields"
    xSql = "select count(*) from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'SAA'" _
         & " and BIATABK1 = '" & X & "'"
    Set rsSabX = cnsab.Execute(xSql)
End If

ReDim arrMT_Field(100)
ReDim arrMT_Field_Code(rsSabX(0) + 1), arrMT_Field_Lib(rsSabX(0) + 1)

xSql = "select *from " & paramIBM_Library_SABSPE & ".YBIATAB0 where BIATABID = 'SAA'" _
     & " and BIATABK1 = '" & X & "' order by BIATABK2"
Set rsSabX = cnsab.Execute(xSql)
arrMT_Field_Code_Nb = 0
Do While Not rsSabX.EOF
    X = Trim(rsSabX("BIATABK2"))
    If IsNumeric(X) Then
        arrMT_Field(Val(X)) = Trim(rsSabX("BIATABTXT"))
    Else
        arrMT_Field_Code_Nb = arrMT_Field_Code_Nb + 1
        arrMT_Field_Code(arrMT_Field_Code_Nb) = X
        arrMT_Field_Lib(arrMT_Field_Code_Nb) = Trim(rsSabX("BIATABTXT"))
    End If
    
    rsSabX.MoveNext
Loop


End Sub


Public Function sqlYSWISAB0_Delete(oldY As typeYSWISAB0)
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

'On Error GoTo Error_Handler
sqlYSWISAB0_Delete = Null
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where SWISABSWID = " & oldY.SWISABSWID


'===================================================================================

    
    xSql = "delete from " & paramIBM_Library_SABSPE_XXX & ".YSWISAB0" & xWhere
    Call FEU_ROUGE
    Set rsADO = cnSab_Update.Execute(xSql, Nb)
    Call FEU_VERT
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYSWISAB0_Delete = "Erreur màj : " & xWhere
        Exit Function
    End If
    


Exit Function
Error_Handler:
    sqlYSWISAB0_Delete = Error
End Function

Public Function sqlYSWISAB0_Update(newY As typeYSWISAB0, oldY As typeYSWISAB0)
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

'On Error GoTo Error_Handler
sqlYSWISAB0_Update = Null

' Contrôle  : Même clé d'accès old / new
'===================================================================================
If oldY.SWISABSWID <> newY.SWISABSWID Then
    sqlYSWISAB0_Update = "Erreur SWISABSWID"
    Exit Function
End If
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where SWISABSWID = " & oldY.SWISABSWID

xSet = " set"
blnUpdate = False


' Détecter les modifications
'===================================================================================

If newY.SWISABWID1 <> oldY.SWISABWID1 Then blnUpdate = True:  xSet = xSet & " , SWISABWID1 = " & newY.SWISABWID1
If newY.SWISABWIDL <> oldY.SWISABWIDL Then blnUpdate = True:  xSet = xSet & " , SWISABWIDL = " & newY.SWISABWIDL
If newY.SWISABWIDH <> oldY.SWISABWIDH Then blnUpdate = True:  xSet = xSet & " , SWISABWIDH = " & newY.SWISABWIDH
If newY.SWISABWMTD <> oldY.SWISABWMTD Then blnUpdate = True:  xSet = xSet & " , SWISABWMTD = " & cur_P(newY.SWISABWMTD)

If newY.SWISABWAMJ <> oldY.SWISABWAMJ Then blnUpdate = True:  xSet = xSet & " , SWISABWAMJ = " & newY.SWISABWAMJ
If newY.SWISABWHMS <> oldY.SWISABWHMS Then blnUpdate = True:  xSet = xSet & " , SWISABWHMS = " & newY.SWISABWHMS
If newY.SWISABOPEN <> oldY.SWISABOPEN Then blnUpdate = True:  xSet = xSet & " , SWISABOPEN = " & newY.SWISABOPEN
If newY.SWISABZSWI <> oldY.SWISABZSWI Then blnUpdate = True:  xSet = xSet & " , SWISABZSWI = " & newY.SWISABZSWI
If newY.SWISABXAMJ <> oldY.SWISABXAMJ Then blnUpdate = True:  xSet = xSet & " , SWISABXAMJ = " & newY.SWISABXAMJ
If newY.SWISABXHMS <> oldY.SWISABXHMS Then blnUpdate = True:  xSet = xSet & " , SWISABXHMS = " & newY.SWISABXHMS


If newY.SWISABWES <> oldY.SWISABWES Then blnUpdate = True:  xSet = xSet & " , SWISABWES = '" & newY.SWISABWES & "'"
If newY.SWISABWMTK <> oldY.SWISABWMTK Then blnUpdate = True:  xSet = xSet & " , SWISABWMTK = '" & newY.SWISABWMTK & "'"
If newY.SWISABWBIC <> oldY.SWISABWBIC Then blnUpdate = True:  xSet = xSet & " , SWISABWBIC = '" & newY.SWISABWBIC & "'"
If newY.SWISABWDEV <> oldY.SWISABWDEV Then blnUpdate = True:  xSet = xSet & " , SWISABWDEV = '" & newY.SWISABWDEV & "'"
If newY.SWISABWN20 <> oldY.SWISABWN20 Then blnUpdate = True:  xSet = xSet & " , SWISABWN20 = '" & Replace(newY.SWISABWN20, "'", "''") & "'"
If newY.SWISABWL20 <> oldY.SWISABWL20 Then blnUpdate = True:  xSet = xSet & " , SWISABWL20 = '" & Replace(newY.SWISABWL20, "'", "''") & "'"
If newY.SWISABWSRV <> oldY.SWISABWSRV Then blnUpdate = True:  xSet = xSet & " , SWISABWSRV = '" & newY.SWISABWSRV & "'"
If newY.SWISABWSTA <> oldY.SWISABWSTA Then blnUpdate = True:  xSet = xSet & " , SWISABWSTA = '" & newY.SWISABWSTA & "'"

If newY.SWISABSER <> oldY.SWISABSER Then blnUpdate = True:  xSet = xSet & " , SWISABSER = '" & newY.SWISABSER & "'"
If newY.SWISABSSE <> oldY.SWISABSSE Then blnUpdate = True:  xSet = xSet & " , SWISABSSE = '" & newY.SWISABSSE & "'"
If newY.SWISABOPEC <> oldY.SWISABOPEC Then blnUpdate = True:  xSet = xSet & " , SWISABOPEC = '" & newY.SWISABOPEC & "'"
If newY.SWISABSTA <> oldY.SWISABSTA Then blnUpdate = True:  xSet = xSet & " , SWISABSTA = '" & newY.SWISABSTA & "'"

If newY.SWISABXGOS <> oldY.SWISABXGOS Then blnUpdate = True:  xSet = xSet & " , SWISABXGOS = '" & newY.SWISABXGOS & "'"
If newY.SWISABXEVE <> oldY.SWISABXEVE Then blnUpdate = True:  xSet = xSet & " , SWISABXEVE = '" & newY.SWISABXEVE & "'"
If newY.SWISABKPDE <> oldY.SWISABKPDE Then blnUpdate = True:  xSet = xSet & " , SWISABKPDE = '" & newY.SWISABKPDE & "'"
If newY.SWISABK20 <> oldY.SWISABK20 Then blnUpdate = True:  xSet = xSet & " , SWISABK20 = '" & newY.SWISABK20 & "'"
If newY.SWISABK999 <> oldY.SWISABK999 Then blnUpdate = True:  xSet = xSet & " , SWISABK999 = '" & newY.SWISABK999 & "'"
If newY.SWISABKSRV <> oldY.SWISABKSRV Then blnUpdate = True:  xSet = xSet & " , SWISABKSRV = '" & newY.SWISABKSRV & "'"


If blnUpdate Then
    Mid$(xSet, 1, 6) = " set  "
    xSql = "update " & paramIBM_Library_SABSPE_XXX & ".YSWISAB0" & xSet & xWhere
    Call FEU_ROUGE
    Set rsADO = cnSab_Update.Execute(xSql, Nb)
    Call FEU_VERT
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYSWISAB0_Update = "Erreur màj : " & newY.SWISABSWID
        Exit Function
    End If
    
End If

Exit Function
Error_Handler:
    sqlYSWISAB0_Update = Error
End Function
Public Function sqlYSWISAB0_Update_Field(lSWISABSWID As Long, lSQL_Set As String)
Dim xSql As String, Nb As Long

'On Error GoTo Error_Handler
sqlYSWISAB0_Update_Field = Null



xSql = "update " & paramIBM_Library_SABSPE_XXX & ".YSWISAB0 " & lSQL_Set & " where SWISABSWID = " & lSWISABSWID
Call FEU_ROUGE
Set rsADO = cnSab_Update.Execute(xSql, Nb)
Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYSWISAB0_Update_Field = "Erreur màj : " & lSWISABSWID
    Exit Function
End If
    

Exit Function
Error_Handler:
    sqlYSWISAB0_Update_Field = Error
End Function


Public Function sqlYSWISAB0_Insert(newY As typeYSWISAB0)
Dim V
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

'On Error GoTo Error_Handler
sqlYSWISAB0_Insert = Null
xSet = " (SWISABSWID "
xValues = " values(" & newY.SWISABSWID

' Détecter les modifications
'===================================================================================

If newY.SWISABWID1 <> 0 Then xSet = xSet & ",SWISABWID1": xValues = xValues & " ," & newY.SWISABWID1
If newY.SWISABWIDL <> 0 Then xSet = xSet & ",SWISABWIDL": xValues = xValues & " ," & newY.SWISABWIDL
If newY.SWISABWIDH <> 0 Then xSet = xSet & ",SWISABWIDH": xValues = xValues & " ," & newY.SWISABWIDH
If newY.SWISABWMTD <> 0 Then xSet = xSet & ",SWISABWMTD": xValues = xValues & " ," & cur_P(newY.SWISABWMTD)
If newY.SWISABWAMJ <> 0 Then xSet = xSet & ",SWISABWAMJ": xValues = xValues & " ," & newY.SWISABWAMJ
If newY.SWISABWHMS <> 0 Then xSet = xSet & ",SWISABWHMS": xValues = xValues & " ," & newY.SWISABWHMS
If newY.SWISABOPEN <> 0 Then xSet = xSet & ",SWISABOPEN": xValues = xValues & " ," & newY.SWISABOPEN
If newY.SWISABZSWI <> 0 Then xSet = xSet & ",SWISABZSWI": xValues = xValues & " ," & newY.SWISABZSWI
If newY.SWISABXAMJ <> 0 Then xSet = xSet & ",SWISABXAMJ": xValues = xValues & " ," & newY.SWISABXAMJ
If newY.SWISABXHMS <> 0 Then xSet = xSet & ",SWISABXHMS": xValues = xValues & " ," & newY.SWISABXHMS


If Trim(newY.SWISABWES) <> "" Then xSet = xSet & ",SWISABWES": xValues = xValues & " ,'" & newY.SWISABWES & "'"
If Trim(newY.SWISABWMTK) <> "" Then xSet = xSet & ",SWISABWMTK": xValues = xValues & " ,'" & newY.SWISABWMTK & "'"
If Trim(newY.SWISABWBIC) <> "" Then xSet = xSet & ",SWISABWBIC": xValues = xValues & " ,'" & newY.SWISABWBIC & "'"
If Trim(newY.SWISABWDEV) <> "" Then xSet = xSet & ",SWISABWDEV": xValues = xValues & " ,'" & newY.SWISABWDEV & "'"
If Trim(newY.SWISABWN20) <> "" Then xSet = xSet & ",SWISABWN20": xValues = xValues & " ,'" & Replace(newY.SWISABWN20, "'", "''") & "'"
If Trim(newY.SWISABWL20) <> "" Then xSet = xSet & ",SWISABWL20": xValues = xValues & " ,'" & Replace(newY.SWISABWL20, "'", "''") & "'"
If Trim(newY.SWISABWSRV) <> "" Then xSet = xSet & ",SWISABWSRV": xValues = xValues & " ,'" & newY.SWISABWSRV & "'"
If Trim(newY.SWISABWSTA) <> "" Then xSet = xSet & ",SWISABWSTA": xValues = xValues & " ,'" & newY.SWISABWSTA & "'"

If Trim(newY.SWISABSER) <> "" Then xSet = xSet & ",SWISABSER": xValues = xValues & " ,'" & newY.SWISABSER & "'"
If Trim(newY.SWISABSSE) <> "" Then xSet = xSet & ",SWISABSSE": xValues = xValues & " ,'" & newY.SWISABSSE & "'"
If Trim(newY.SWISABOPEC) <> "" Then xSet = xSet & ",SWISABOPEC": xValues = xValues & " ,'" & newY.SWISABOPEC & "'"
If Trim(newY.SWISABSTA) <> "" Then xSet = xSet & ",SWISABSTA": xValues = xValues & " ,'" & newY.SWISABSTA & "'"

If Trim(newY.SWISABXGOS) <> "" Then xSet = xSet & ",SWISABXGOS": xValues = xValues & " ,'" & newY.SWISABXGOS & "'"
If Trim(newY.SWISABXEVE) <> "" Then xSet = xSet & ",SWISABXEVE": xValues = xValues & " ,'" & newY.SWISABXEVE & "'"
If Trim(newY.SWISABKPDE) <> "" Then xSet = xSet & ",SWISABKPDE": xValues = xValues & " ,'" & newY.SWISABKPDE & "'"
If Trim(newY.SWISABK20) <> "" Then xSet = xSet & ",SWISABK20": xValues = xValues & " ,'" & newY.SWISABK20 & "'"
If Trim(newY.SWISABK999) <> "" Then xSet = xSet & ",SWISABK999": xValues = xValues & " ,'" & newY.SWISABK999 & "'"
If Trim(newY.SWISABKSRV) <> "" Then xSet = xSet & ",SWISABKSRV": xValues = xValues & " ,'" & newY.SWISABKSRV & "'"



xSql = "Insert into " & paramIBM_Library_SABSPE_XXX & ".YSWISAB0" & xSet & ")" & xValues & ")"
Call FEU_ROUGE
Set rsADO = cnSab_Update.Execute(xSql, Nb)
Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYSWISAB0_Insert = "Erreur màj : " & newY.SWISABSWID
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYSWISAB0_Insert = Error
End Function

Public Function rsYSWISAB0_GetBuffer(rsADO As ADODB.Recordset, lYSWISAB0 As typeYSWISAB0)
On Error GoTo Error_Handler
rsYSWISAB0_GetBuffer = Null

lYSWISAB0.JORCV = 0
lYSWISAB0.JOSEQN = 0
lYSWISAB0.JRNBIATRN = 0
lYSWISAB0.JOENTT = ""
lYSWISAB0.JODATE = ""

lYSWISAB0.SWISABSWID = rsADO("SWISABSWID")

lYSWISAB0.SWISABWID1 = rsADO("SWISABWID1")
lYSWISAB0.SWISABWIDL = rsADO("SWISABWIDL")
lYSWISAB0.SWISABWIDH = rsADO("SWISABWIDH")

lYSWISAB0.SWISABWES = rsADO("SWISABWES")
lYSWISAB0.SWISABWMTK = rsADO("SWISABWMTK")
lYSWISAB0.SWISABWBIC = rsADO("SWISABWBIC")
lYSWISAB0.SWISABWDEV = rsADO("SWISABWDEV")
lYSWISAB0.SWISABWMTD = rsADO("SWISABWMTD")
lYSWISAB0.SWISABWN20 = Trim(rsADO("SWISABWN20"))
lYSWISAB0.SWISABWL20 = Trim(rsADO("SWISABWL20"))
lYSWISAB0.SWISABWSRV = rsADO("SWISABWSRV")
lYSWISAB0.SWISABWSTA = rsADO("SWISABWSTA")

lYSWISAB0.SWISABWAMJ = rsADO("SWISABWAMJ")
lYSWISAB0.SWISABWHMS = rsADO("SWISABWHMS")

lYSWISAB0.SWISABSER = rsADO("SWISABSER")
lYSWISAB0.SWISABSSE = rsADO("SWISABSSE")
lYSWISAB0.SWISABOPEC = rsADO("SWISABOPEC")
lYSWISAB0.SWISABOPEN = rsADO("SWISABOPEN")
lYSWISAB0.SWISABZSWI = rsADO("SWISABZSWI")
lYSWISAB0.SWISABSTA = rsADO("SWISABSTA")

lYSWISAB0.SWISABXAMJ = rsADO("SWISABXAMJ")
lYSWISAB0.SWISABXHMS = rsADO("SWISABXHMS")
lYSWISAB0.SWISABXGOS = rsADO("SWISABXGOS")
lYSWISAB0.SWISABXEVE = rsADO("SWISABXEVE")
lYSWISAB0.SWISABKPDE = rsADO("SWISABKPDE")
lYSWISAB0.SWISABK20 = rsADO("SWISABK20")
lYSWISAB0.SWISABK999 = rsADO("SWISABK999")
lYSWISAB0.SWISABKSRV = rsADO("SWISABKSRV")



Exit Function
Error_Handler:
rsYSWISAB0_GetBuffer = Error


End Function
'---------------------------------------------------------
Public Function rsJSWISAB0_GetBuffer(rsADO As ADODB.Recordset, rsYSWISAB0 As typeYSWISAB0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsJSWISAB0_GetBuffer = Null

rsJSWISAB0_GetBuffer = rsYSWISAB0_GetBuffer(rsADO, rsYSWISAB0)
rsYSWISAB0.JORCV = rsADO("JORCV")
rsYSWISAB0.JOSEQN = rsADO("JOSEQN")
rsYSWISAB0.JRNBIATRN = rsADO("JRNBIATRN")
rsYSWISAB0.JOENTT = rsADO("JOENTT")
rsYSWISAB0.JODATE = rsADO("JODATE")

Exit Function

Error_Handler:

rsJSWISAB0_GetBuffer = Error

End Function


Public Function rsYSWISAB0_Init(lYSWISAB0 As typeYSWISAB0)

lYSWISAB0.SWISABSWID = 0

lYSWISAB0.SWISABWID1 = 0
lYSWISAB0.SWISABWIDL = 0
lYSWISAB0.SWISABWIDH = 0

lYSWISAB0.SWISABWES = ""
lYSWISAB0.SWISABWMTK = ""
lYSWISAB0.SWISABWBIC = ""
lYSWISAB0.SWISABWDEV = ""
lYSWISAB0.SWISABWMTD = 0
lYSWISAB0.SWISABWN20 = ""
lYSWISAB0.SWISABWL20 = ""
lYSWISAB0.SWISABWSRV = ""
lYSWISAB0.SWISABWSTA = ""

lYSWISAB0.SWISABWAMJ = 0
lYSWISAB0.SWISABWHMS = 0

lYSWISAB0.SWISABSER = ""
lYSWISAB0.SWISABSSE = ""
lYSWISAB0.SWISABOPEC = ""
lYSWISAB0.SWISABOPEN = 0
lYSWISAB0.SWISABZSWI = 0
lYSWISAB0.SWISABSTA = ""


      
lYSWISAB0.SWISABXAMJ = 0
lYSWISAB0.SWISABXHMS = 0
lYSWISAB0.SWISABXGOS = ""
lYSWISAB0.SWISABXEVE = ""
lYSWISAB0.SWISABKPDE = ""
lYSWISAB0.SWISABK20 = ""
lYSWISAB0.SWISABK999 = ""
lYSWISAB0.SWISABKSRV = ""
    


lYSWISAB0.JORCV = 0
lYSWISAB0.JOSEQN = 0
lYSWISAB0.JRNBIATRN = 0
    
lYSWISAB0.JOENTT = ""
lYSWISAB0.JODATE = ""

End Function














