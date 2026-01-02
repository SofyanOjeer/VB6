Attribute VB_Name = "srvYPCICPT0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim rsAdo As ADODB.Recordset
 
Type typeYPCICPT0
 
      PCICPTBASE    As String
      PCICPTLNK     As String
      PCICPTLEN     As Integer
      PCICPTMETA    As String
      PCICPTAUTO    As String
      PCICPTSUFX    As String
      PCICPTTXT     As String
    
      PCICPTUUSR   As String
      PCICPTUAMJ   As Long
      PCICPTUHMS   As Long
      PCICPTUSEQ   As Long
'____________________________________________________ Journalisation
    JORCV                   As Long
    JOSEQN                  As Long
    JRNBIATRN               As Long
    
    JOENTT          As String * 2
    JODATE          As String * 6

'____________________________________________________ Journalisation
End Type
Public xYPCICPT0 As typeYPCICPT0
Public Function sqlYPCICPT0_Delete(oldY As typeYPCICPT0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

On Error GoTo Error_Handler
sqlYPCICPT0_Delete = Null
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where PCICPTBASE = " & oldY.PCICPTBASE _
       & " and PCICPTUSEQ = " & oldY.PCICPTUSEQ

'===================================================================================

    
    xSQL = "delete from " & paramIBM_Library_SABSPE_XXX & ".YPCICPT0" & xWhere
    Call FEU_ROUGE
    Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
    Call FEU_VERT
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYPCICPT0_Delete = "Erreur màj : " & xWhere
        Exit Function
    End If
    


Exit Function
Error_Handler:
    sqlYPCICPT0_Delete = Error
End Function

Public Function sqlYPCICPT0_Update(newY As typeYPCICPT0, oldY As typeYPCICPT0, blnUUSR As Boolean)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

On Error GoTo Error_Handler
sqlYPCICPT0_Update = Null

' Contrôle  : Même clé d'accès old / new
'===================================================================================
If oldY.PCICPTBASE <> newY.PCICPTBASE _
Or oldY.PCICPTUSEQ <> newY.PCICPTUSEQ Then
    sqlYPCICPT0_Update = "Erreur PCICPTLEN : " & newY.PCICPTLEN & "." & oldY.PCICPTUSEQ
    Exit Function
End If
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where PCICPTBASE = '" & oldY.PCICPTBASE & "'" _
       & " and PCICPTUSEQ = " & oldY.PCICPTUSEQ

newY.PCICPTUSEQ = newY.PCICPTUSEQ + 1
xSet = xSet & " set PCICPTUSEQ = " & newY.PCICPTUSEQ
blnUpdate = False

If blnUUSR Then
    newY.PCICPTUUSR = usrName_UCase
    newY.PCICPTUAMJ = DSys
    newY.PCICPTUHMS = time_Hms
End If

' Détecter les modifications
'===================================================================================
If newY.PCICPTLNK <> oldY.PCICPTLNK Then blnUpdate = True:  xSet = xSet & " , PCICPTLNK = '" & newY.PCICPTLNK & "'"
If newY.PCICPTLEN <> oldY.PCICPTLEN Then blnUpdate = True:  xSet = xSet & " , PCICPTLEN = " & newY.PCICPTLEN

If newY.PCICPTMETA <> oldY.PCICPTMETA Then blnUpdate = True:  xSet = xSet & " , PCICPTMETA = '" & Replace(Trim(newY.PCICPTMETA), "'", "''") & "'"
If newY.PCICPTAUTO <> oldY.PCICPTAUTO Then blnUpdate = True:  xSet = xSet & " , PCICPTAUTO = '" & newY.PCICPTAUTO & "'"
If newY.PCICPTSUFX <> oldY.PCICPTSUFX Then blnUpdate = True:  xSet = xSet & " , PCICPTSUFX = '" & Replace(Trim(newY.PCICPTSUFX), "'", "''") & "'"
If newY.PCICPTTXT <> oldY.PCICPTTXT Then blnUpdate = True:  xSet = xSet & " , PCICPTTXT = '" & Replace(Trim(newY.PCICPTTXT), "'", "''") & "'"


If newY.PCICPTUUSR <> oldY.PCICPTUUSR Then xSet = xSet & " , PCICPTUUSR = '" & newY.PCICPTUUSR & "'"
If newY.PCICPTUAMJ <> oldY.PCICPTUAMJ Then xSet = xSet & " , PCICPTUAMJ = " & newY.PCICPTUAMJ
If newY.PCICPTUHMS <> oldY.PCICPTUHMS Then xSet = xSet & " , PCICPTUHMS = " & newY.PCICPTUHMS


If blnUpdate Then
    
    xSQL = "update " & paramIBM_Library_SABSPE_XXX & ".YPCICPT0" & xSet & xWhere
    Call FEU_ROUGE
    Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
    Call FEU_VERT
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYPCICPT0_Update = "Erreur màj : " & newY.PCICPTLEN
        Exit Function
    End If
    
End If

Exit Function
Error_Handler:
    sqlYPCICPT0_Update = Error
End Function

Public Function sqlYPCICPT0_Insert(newY As typeYPCICPT0)
Dim V
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlYPCICPT0_Insert = Null
xSet = " (PCICPTBASE"
xValues = " values('" & newY.PCICPTBASE & "'"

newY.PCICPTUUSR = usrName_UCase
newY.PCICPTUAMJ = DSys
newY.PCICPTUHMS = time_Hms

' Détecter les modifications
'===================================================================================
If newY.PCICPTLEN <> 0 Then xSet = xSet & ",PCICPTLEN": xValues = xValues & " ," & newY.PCICPTLEN
If newY.PCICPTUSEQ <> 0 Then xSet = xSet & ",PCICPTUSEQ": xValues = xValues & " ," & newY.PCICPTUSEQ
If newY.PCICPTUAMJ <> 0 Then xSet = xSet & ",PCICPTUAMJ": xValues = xValues & " ," & newY.PCICPTUAMJ
If newY.PCICPTUHMS <> 0 Then xSet = xSet & ",PCICPTUHMS": xValues = xValues & " ," & newY.PCICPTUHMS

If Trim(newY.PCICPTLNK) <> "" Then xSet = xSet & ",PCICPTLNK": xValues = xValues & " ,'" & newY.PCICPTLNK & "'"

If Trim(newY.PCICPTMETA) <> "" Then xSet = xSet & ",PCICPTMETA": xValues = xValues & " ,'" & Replace(Trim(newY.PCICPTMETA), "'", "''") & "'"
If Trim(newY.PCICPTAUTO) <> "" Then xSet = xSet & ",PCICPTAUTO": xValues = xValues & " ,'" & newY.PCICPTAUTO & "'"
If Trim(newY.PCICPTSUFX) <> "" Then xSet = xSet & ",PCICPTSUFX": xValues = xValues & " ,'" & Replace(Trim(newY.PCICPTSUFX), "'", "''") & "'"
If Trim(newY.PCICPTTXT) <> "" Then xSet = xSet & ",PCICPTTXT": xValues = xValues & " ,'" & Replace(Trim(newY.PCICPTTXT), "'", "''") & "'"

If Trim(newY.PCICPTUUSR) <> "" Then xSet = xSet & ",PCICPTUUSR": xValues = xValues & " ,'" & newY.PCICPTUUSR & "'"

xSQL = "Insert into " & paramIBM_Library_SABSPE_XXX & ".YPCICPT0" & xSet & ")" & xValues & ")"
Call FEU_ROUGE
Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYPCICPT0_Insert = "Erreur màj : " & newY.PCICPTLEN
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYPCICPT0_Insert = Error
End Function

Public Function rsYPCICPT0_GetBuffer(rsAdo As ADODB.Recordset, lYPCICPT0 As typeYPCICPT0)
On Error GoTo Error_Handler
rsYPCICPT0_GetBuffer = Null

lYPCICPT0.JORCV = 0
lYPCICPT0.JOSEQN = 0
lYPCICPT0.JRNBIATRN = 0
lYPCICPT0.JOENTT = ""
lYPCICPT0.JODATE = ""

lYPCICPT0.PCICPTBASE = rsAdo("PCICPTBASE")
lYPCICPT0.PCICPTLNK = rsAdo("PCICPTLNK")
lYPCICPT0.PCICPTLEN = rsAdo("PCICPTLEN")
lYPCICPT0.PCICPTMETA = rsAdo("PCICPTMETA")
lYPCICPT0.PCICPTAUTO = rsAdo("PCICPTAUTO")
lYPCICPT0.PCICPTSUFX = rsAdo("PCICPTSUFX")
lYPCICPT0.PCICPTTXT = rsAdo("PCICPTTXT")

lYPCICPT0.PCICPTUUSR = rsAdo("PCICPTUUSR")
lYPCICPT0.PCICPTUAMJ = rsAdo("PCICPTUAMJ")
lYPCICPT0.PCICPTUHMS = rsAdo("PCICPTUHMS")
lYPCICPT0.PCICPTUSEQ = rsAdo("PCICPTUSEQ")

Exit Function
Error_Handler:
rsYPCICPT0_GetBuffer = Error


End Function
'---------------------------------------------------------
Public Function rsJEICGCC0_GetBuffer(rsAdo As ADODB.Recordset, rsYPCICPT0 As typeYPCICPT0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsJEICGCC0_GetBuffer = Null

rsJEICGCC0_GetBuffer = rsYPCICPT0_GetBuffer(rsAdo, rsYPCICPT0)
rsYPCICPT0.JORCV = rsAdo("JORCV")
rsYPCICPT0.JOSEQN = rsAdo("JOSEQN")
rsYPCICPT0.JRNBIATRN = rsAdo("JRNBIATRN")
rsYPCICPT0.JOENTT = rsAdo("JOENTT")
rsYPCICPT0.JODATE = rsAdo("JODATE")

Exit Function

Error_Handler:

rsJEICGCC0_GetBuffer = Error

End Function


Public Function rsYPCICPT0_Init(lYPCICPT0 As typeYPCICPT0)

lYPCICPT0.PCICPTLEN = 0
lYPCICPT0.PCICPTBASE = ""
lYPCICPT0.PCICPTLNK = ""
lYPCICPT0.PCICPTMETA = ""
lYPCICPT0.PCICPTAUTO = ""
lYPCICPT0.PCICPTSUFX = ""
lYPCICPT0.PCICPTTXT = ""
    
lYPCICPT0.PCICPTUUSR = ""
lYPCICPT0.PCICPTUAMJ = 0
lYPCICPT0.PCICPTUHMS = 0
lYPCICPT0.PCICPTUSEQ = 0
End Function






