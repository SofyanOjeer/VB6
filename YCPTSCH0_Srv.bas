Attribute VB_Name = "srvYCPTSCH0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim rsAdo As ADODB.Recordset
 
Type typeYCPTSCH0
 
      SCHEMAFDT   As Long
      SCHEMAFUT   As Integer
      SCHEMAETA   As Integer
      SCHEMAOPE   As String * 3
      SCHEMAEVE   As String * 3
      SCHEMAPLA   As Integer     '
      SCHEMAARG   As String * 18
      CPTSCHUSR1  As String * 10
      CPTSCHAMJ1  As Long
      CPTSCHHMS1  As Long
      CPTSCHUSR2  As String * 10
      CPTSCHAMJ2  As Long
      CPTSCHHMS2  As Long
      CPTSCHTEXT  As String * 64   'motif
      CPTSCHSTA   As String * 1   'statut
      CPTSCHUPDS  As Long         'séquence
      CPTSCHUSR   As String * 10   'utilisateur
      End Type
Public xYCPTSCH0 As typeYCPTSCH0
Public Function sqlYCPTSCH0_Insert(newY As typeYCPTSCH0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlYCPTSCH0_Insert = Null

xSet = " (SCHEMAFDT"
xValues = " values(" & newY.SCHEMAFDT

' Détecter les modifications
'===================================================================================
If newY.SCHEMAFUT <> 0 Then xSet = xSet & ",SCHEMAFUT": xValues = xValues & " ," & newY.SCHEMAFUT
If newY.SCHEMAETA <> 0 Then xSet = xSet & ",SCHEMAETA": xValues = xValues & " ," & newY.SCHEMAETA
If newY.SCHEMAPLA <> 0 Then xSet = xSet & ",SCHEMAPLA": xValues = xValues & " ," & newY.SCHEMAPLA
If newY.CPTSCHAMJ1 <> 0 Then xSet = xSet & ",CPTSCHAMJ1": xValues = xValues & " ," & newY.CPTSCHAMJ1
If newY.CPTSCHHMS1 <> 0 Then xSet = xSet & ",CPTSCHHMS1": xValues = xValues & " ," & newY.CPTSCHHMS1
If newY.CPTSCHAMJ2 <> 0 Then xSet = xSet & ",CPTSCHAMJ2": xValues = xValues & " ," & newY.CPTSCHAMJ2
If newY.CPTSCHHMS2 <> 0 Then xSet = xSet & ",CPTSCHHMS2": xValues = xValues & " ," & newY.CPTSCHHMS2

If Trim(newY.SCHEMAOPE) <> "" Then xSet = xSet & ",SCHEMAOPE": xValues = xValues & " ,'" & newY.SCHEMAOPE & "'"
If Trim(newY.SCHEMAEVE) <> "" Then xSet = xSet & ",SCHEMAEVE": xValues = xValues & " ,'" & Trim(newY.SCHEMAEVE) & "'"
If Trim(newY.SCHEMAARG) <> "" Then xSet = xSet & ",SCHEMAARG": xValues = xValues & " ,'" & newY.SCHEMAARG & "'"
If Trim(newY.CPTSCHUSR1) <> "" Then xSet = xSet & ",CPTSCHUSR1": xValues = xValues & " ,'" & Replace(Trim(newY.CPTSCHUSR1), "'", "''") & "'"
If Trim(newY.CPTSCHUSR2) <> "" Then xSet = xSet & ",CPTSCHUSR2": xValues = xValues & " ,'" & Replace(Trim(newY.CPTSCHUSR2), "'", "''") & "'"
If Trim(newY.CPTSCHTEXT) <> "" Then xSet = xSet & ",CPTSCHTEXT": xValues = xValues & " ,'" & Replace(Trim(newY.CPTSCHTEXT), "'", "''") & "'"
If Trim(newY.CPTSCHSTA) <> "" Then xSet = xSet & ",CPTSCHSTA": xValues = xValues & " ,'" & newY.CPTSCHSTA & "'"

newY.CPTSCHUSR = usrName_UCase
xSet = xSet & ",CPTSCHUSR": xValues = xValues & " ,'" & usrName_UCase & "'"

xSQL = "Insert into " & paramIBM_Library_SABSPE & ".YCPTSCH0" & xSet & ")" & xValues & ")"
Call FEU_ROUGE
Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYCPTSCH0_Insert = "Erreur màj : " & newY.SCHEMAFDT
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYCPTSCH0_Insert = Error
End Function

Public Function sqlYCPTSCH0_Update(newY As typeYCPTSCH0, oldY As typeYCPTSCH0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

On Error GoTo Error_Handler
sqlYCPTSCH0_Update = Null

' Contrôle  : Même clé d'accès old / new
'===================================================================================
If oldY.SCHEMAFDT <> newY.SCHEMAFDT Then
    sqlYCPTSCH0_Update = "Erreur SCHEMAFDT : " & newY.SCHEMAFDT & " / " & oldY.SCHEMAFDT
    Exit Function
End If
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where SCHEMAFDT = " & oldY.SCHEMAFDT _
       & " and SCHEMAFUT = " & oldY.SCHEMAFUT _
       & " and SCHEMAETA = " & oldY.SCHEMAETA _
       & " and SCHEMAOPE = '" & oldY.SCHEMAOPE & "'" _
       & " and SCHEMAEVE = '" & oldY.SCHEMAEVE & "'" _
       & " and SCHEMAPLA = " & oldY.SCHEMAPLA _
       & " and SCHEMAARG = '" & oldY.SCHEMAARG & "'" _
       & " and CPTSCHUPDS = " & oldY.CPTSCHUPDS
       
newY.CPTSCHUPDS = newY.CPTSCHUPDS + 1
xSet = xSet & " set CPTSCHUPDS = " & newY.CPTSCHUPDS
blnUpdate = False

' Détecter les modifications
'===================================================================================
If newY.SCHEMAFDT <> oldY.SCHEMAFDT Then blnUpdate = True:  xSet = xSet & " , SCHEMAFDT = '" & cur_P(newY.SCHEMAFDT) & "'"
If newY.SCHEMAFUT <> oldY.SCHEMAFUT Then blnUpdate = True:  xSet = xSet & " , SCHEMAFUT = '" & cur_P(newY.SCHEMAFUT) & "'"
If newY.SCHEMAETA <> oldY.SCHEMAETA Then blnUpdate = True:  xSet = xSet & " , SCHEMAETA = '" & cur_P(newY.SCHEMAETA) & "'"
If newY.SCHEMAPLA <> oldY.SCHEMAPLA Then blnUpdate = True:  xSet = xSet & " , SCHEMAPLA = '" & cur_P(newY.SCHEMAPLA) & "'"
If newY.CPTSCHAMJ1 <> oldY.CPTSCHAMJ1 Then blnUpdate = True: xSet = xSet & " , CPTSCHAMJ1 = " & newY.CPTSCHAMJ1
If newY.CPTSCHHMS1 <> oldY.CPTSCHHMS1 Then blnUpdate = True: xSet = xSet & " , CPTSCHHMS1 = " & newY.CPTSCHHMS1
If newY.CPTSCHAMJ2 <> oldY.CPTSCHAMJ2 Then blnUpdate = True: xSet = xSet & " , CPTSCHAMJ2 = " & newY.CPTSCHAMJ2
If newY.CPTSCHHMS2 <> oldY.CPTSCHHMS2 Then blnUpdate = True: xSet = xSet & " , CPTSCHHMS2 = " & newY.CPTSCHHMS2

If newY.SCHEMAOPE <> oldY.SCHEMAOPE Then blnUpdate = True: xSet = xSet & " , SCHEMAOPE = " & newY.SCHEMAOPE
If newY.SCHEMAEVE <> oldY.SCHEMAEVE Then blnUpdate = True:  xSet = xSet & " , SCHEMAEVE = '" & newY.SCHEMAEVE & "'"
If newY.SCHEMAARG <> oldY.SCHEMAARG Then blnUpdate = True:  xSet = xSet & " , SCHEMAARG= '" & newY.SCHEMAARG & "'"
If newY.CPTSCHUSR1 <> oldY.CPTSCHUSR1 Then blnUpdate = True:  xSet = xSet & " , CPTSCHUSR1 = '" & Replace(Trim(newY.CPTSCHUSR1), "'", "''") & "'"
If newY.CPTSCHUSR2 <> oldY.CPTSCHUSR2 Then blnUpdate = True:  xSet = xSet & " , CPTSCHUSR2 = '" & Replace(Trim(newY.CPTSCHUSR2), "'", "''") & "'"
If newY.CPTSCHTEXT <> oldY.CPTSCHTEXT Then blnUpdate = True:  xSet = xSet & " , CPTSCHTEXT = '" & Replace(Trim(newY.CPTSCHTEXT), "'", "''") & "'"

newY.CPTSCHUSR = usrName_UCase
xSet = xSet & " , CPTSCHUSR = '" & usrName_UCase & "'"
If newY.SCHEMAFDT < 0 Then blnUpdate = True  ' records techniques

If blnUpdate Then
    
    xSQL = "update " & paramIBM_Library_SABSPE & ".YCPTSCH0" & xSet & xWhere
    Call FEU_ROUGE
    Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
    Call FEU_VERT
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYCPTSCH0_Update = "Erreur màj : " & newY.SCHEMAFDT
        Exit Function
    End If
    
End If

Exit Function
Error_Handler:
    sqlYCPTSCH0_Update = Error
End Function

Public Function rsYCPTSCH0_GetBuffer(rsAdo As ADODB.Recordset, lYCPTSCH0 As typeYCPTSCH0)
On Error GoTo Error_Handler
rsYCPTSCH0_GetBuffer = Null
lYCPTSCH0.SCHEMAFDT = rsAdo("SCHEMAFDT")
lYCPTSCH0.SCHEMAFUT = rsAdo("SCHEMAFUT")
lYCPTSCH0.SCHEMAETA = rsAdo("SCHEMAETA")
lYCPTSCH0.SCHEMAEVE = rsAdo("SCHEMAEVE")
lYCPTSCH0.SCHEMAOPE = rsAdo("SCHEMAOPE")
lYCPTSCH0.SCHEMAPLA = rsAdo("SCHEMAPLA")
lYCPTSCH0.SCHEMAARG = rsAdo("SCHEMAARG")
lYCPTSCH0.CPTSCHUSR1 = rsAdo("CPTSCHUSR1")
lYCPTSCH0.CPTSCHAMJ1 = rsAdo("CPTSCHAMJ1")
lYCPTSCH0.CPTSCHHMS1 = rsAdo("CPTSCHHMS1")
lYCPTSCH0.CPTSCHUSR2 = rsAdo("CPTSCHUSR2")
lYCPTSCH0.CPTSCHAMJ2 = rsAdo("CPTSCHAMJ2")
lYCPTSCH0.CPTSCHHMS2 = rsAdo("CPTSCHHMS2")
lYCPTSCH0.CPTSCHTEXT = rsAdo("CPTSCHTEXT")

lYCPTSCH0.CPTSCHSTA = rsAdo("CPTSCHSTA")
lYCPTSCH0.CPTSCHUPDS = rsAdo("CPTSCHUPDS")
lYCPTSCH0.CPTSCHUSR = rsAdo("CPTSCHUSR")

Exit Function
Error_Handler:
rsYCPTSCH0_GetBuffer = Error


End Function

Public Function rsYCPTSCH0_Init(lYCPTSCH0 As typeYCPTSCH0)
lYCPTSCH0.SCHEMAFDT = 0
lYCPTSCH0.SCHEMAFUT = 0
lYCPTSCH0.SCHEMAEVE = ""
lYCPTSCH0.SCHEMAOPE = ""
lYCPTSCH0.SCHEMAETA = 0
lYCPTSCH0.SCHEMAPLA = 0
lYCPTSCH0.SCHEMAARG = ""
lYCPTSCH0.CPTSCHUSR1 = ""
lYCPTSCH0.CPTSCHAMJ1 = 0
lYCPTSCH0.CPTSCHHMS1 = 0
lYCPTSCH0.CPTSCHUSR2 = ""
lYCPTSCH0.CPTSCHAMJ2 = 0
lYCPTSCH0.CPTSCHHMS2 = 0
lYCPTSCH0.CPTSCHTEXT = ""
lYCPTSCH0.CPTSCHSTA = ""
lYCPTSCH0.CPTSCHUPDS = 0
lYCPTSCH0.CPTSCHUSR = ""

End Function



