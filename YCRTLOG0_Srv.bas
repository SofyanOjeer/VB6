Attribute VB_Name = "srvYCRTLOG0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim rsAdo As ADODB.Recordset
 
Type typeYCRTLOG0
 
      CRTLOGID   As Long
      
      CRTLOGNAT   As String       'devise
      CRTLOGUUSR  As String       'code client
      CRTLOGUAMJ   As Long
      CRTLOGUHMS  As Long
      
      CRTLOGCPT   As String       'compte
      CRTLOGETA   As Long         'établissement
      CRTLOGPLA   As Long         'plan
      CRTLOGPIE   As Long         'pièce
      CRTLOGECR   As Long         'écriture
      
      CRTLOGTXT   As String
      

End Type
Public zYCRTLOG0 As typeYCRTLOG0



Public Function sqlYCRTLOG0_Update(newY As typeYCRTLOG0, oldY As typeYCRTLOG0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

On Error GoTo Error_Handler
sqlYCRTLOG0_Update = Null

' Contrôle  : Même clé d'accès old / new
'===================================================================================
If oldY.CRTLOGID <> newY.CRTLOGID Then
    sqlYCRTLOG0_Update = "Erreur CRTLOGIDE : " & newY.CRTLOGID
    Exit Function
End If
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

newY.CRTLOGUUSR = usrName_UCase10
newY.CRTLOGUAMJ = DSys
newY.CRTLOGUHMS = time_Hms

xWhere = " where CRTLOGID = " & oldY.CRTLOGID

xSet = xSet & " set CRTLOGTXT = '" & newY.CRTLOGTXT & "'"
blnUpdate = False
If newY.CRTLOGTXT <> oldY.CRTLOGTXT Then blnUpdate = True
' Détecter les modifications
'===================================================================================
If newY.CRTLOGETA <> oldY.CRTLOGETA Then blnUpdate = True: xSet = xSet & " , CRTLOGETA = " & newY.CRTLOGETA
If newY.CRTLOGPLA <> oldY.CRTLOGPLA Then blnUpdate = True: xSet = xSet & " , CRTLOGPLA = " & newY.CRTLOGPLA
If newY.CRTLOGPIE <> oldY.CRTLOGPIE Then blnUpdate = True: xSet = xSet & " , CRTLOGPIE = " & newY.CRTLOGPIE
If newY.CRTLOGECR <> oldY.CRTLOGECR Then blnUpdate = True: xSet = xSet & " , CRTLOGECR = " & newY.CRTLOGECR
If newY.CRTLOGUAMJ <> oldY.CRTLOGUAMJ Then blnUpdate = True: xSet = xSet & " , CRTLOGUAMJ = " & newY.CRTLOGUAMJ
If newY.CRTLOGUHMS <> oldY.CRTLOGUHMS Then blnUpdate = True: xSet = xSet & " , CRTLOGUHMS = " & newY.CRTLOGUHMS

If newY.CRTLOGCPT <> oldY.CRTLOGCPT Then blnUpdate = True:  xSet = xSet & " , CRTLOGCPT= '" & newY.CRTLOGCPT & "'"
If newY.CRTLOGNAT <> oldY.CRTLOGNAT Then blnUpdate = True:  xSet = xSet & " , CRTLOGNAT= '" & newY.CRTLOGNAT & "'"
If newY.CRTLOGUUSR <> oldY.CRTLOGUUSR Then blnUpdate = True:  xSet = xSet & " , CRTLOGUUSR = '" & newY.CRTLOGUUSR & "'"


If blnUpdate Then
    
    xSQL = "update " & paramIBM_Library_SABSPE & ".YCRTLOG0" & xSet & xWhere
    Call FEU_ROUGE
    Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
    Call FEU_VERT
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYCRTLOG0_Update = "Erreur màj : " & newY.CRTLOGID
        Exit Function
    End If
    
End If

Exit Function
Error_Handler:
    sqlYCRTLOG0_Update = Error
End Function

Public Function sqlYCRTLOG0_Insert(newY As typeYCRTLOG0)
Dim V
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler

CRTLOGID_Svt:
'=================
sqlYCRTLOG0_Insert = Null
If zYCRTLOG0.CRTLOGID = 0 Then
    xSQL = "select count(*) from " & paramIBM_Library_SABSPE & ".YCRTLOG0"
    Set rsAdo = cnSab_Update.Execute(xSQL)
    If Not rsAdo.EOF Then
        zYCRTLOG0.CRTLOGID = rsAdo(0)
        xSQL = "select CRTLOGID from " & paramIBM_Library_SABSPE & ".YCRTLOG0 where CRTLOGID >= " & zYCRTLOG0.CRTLOGID & " order by CRTLOGID desc"
        Set rsAdo = cnSab_Update.Execute(xSQL)
        If Not rsAdo.EOF Then zYCRTLOG0.CRTLOGID = rsAdo(0)
    End If
End If

zYCRTLOG0.CRTLOGID = zYCRTLOG0.CRTLOGID + 1
newY.CRTLOGID = zYCRTLOG0.CRTLOGID

xSet = " (CRTLOGID"
xValues = " values(" & newY.CRTLOGID

newY.CRTLOGUUSR = usrName_UCase10
newY.CRTLOGUAMJ = DSys
newY.CRTLOGUHMS = time_Hms


' Détecter les modifications
'===================================================================================
If newY.CRTLOGETA <> 0 Then xSet = xSet & ",CRTLOGETA": xValues = xValues & " ," & newY.CRTLOGETA
If newY.CRTLOGPLA <> 0 Then xSet = xSet & ",CRTLOGPLA": xValues = xValues & " ," & newY.CRTLOGPLA
If newY.CRTLOGPIE <> 0 Then xSet = xSet & ",CRTLOGPIE": xValues = xValues & " ," & newY.CRTLOGPIE
If newY.CRTLOGECR <> 0 Then xSet = xSet & ",CRTLOGECR": xValues = xValues & " ," & newY.CRTLOGECR
If newY.CRTLOGUAMJ <> 0 Then xSet = xSet & ",CRTLOGUAMJ": xValues = xValues & " ," & newY.CRTLOGUAMJ
If newY.CRTLOGUHMS <> 0 Then xSet = xSet & ",CRTLOGUHMS": xValues = xValues & ", " & newY.CRTLOGUHMS

If Trim(newY.CRTLOGNAT) <> "" Then xSet = xSet & ",CRTLOGNAT": xValues = xValues & " ,'" & newY.CRTLOGNAT & "'"
If Trim(newY.CRTLOGUUSR) <> "" Then xSet = xSet & ",CRTLOGUUSR": xValues = xValues & " ,'" & newY.CRTLOGUUSR & "'"
If Trim(newY.CRTLOGTXT) <> "" Then xSet = xSet & ",CRTLOGTXT": xValues = xValues & " ,'" & Replace(newY.CRTLOGTXT, "'", "''") & "'"
If Trim(newY.CRTLOGCPT) <> "" Then xSet = xSet & ",CRTLOGCPT": xValues = xValues & " ,'" & newY.CRTLOGCPT & "'"


xSQL = "Insert into " & paramIBM_Library_SABSPE & ".YCRTLOG0" & xSet & ")" & xValues & ")"
Call FEU_ROUGE
Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYCRTLOG0_Insert = "Erreur màj : " & newY.CRTLOGPIE & newY.CRTLOGECR
    Exit Function
End If
 
Exit Function
Error_Handler:
    '[IBM][Pilote ODBC iSeries Access][DB2 UDB]SQL0803 - La valeur indiquée est incorrecte car elle produirait une clé en double.
    X = Error
    If InStr(X, "SQL0803") > 0 Then GoTo CRTLOGID_Svt
    
    sqlYCRTLOG0_Insert = Error
End Function


Public Function rsYCRTLOG0_GetBuffer(rsAdo As ADODB.Recordset, lYCRTLOG0 As typeYCRTLOG0)
On Error GoTo Error_Handler
rsYCRTLOG0_GetBuffer = Null

lYCRTLOG0.CRTLOGID = rsAdo("CRTLOGID")

lYCRTLOG0.CRTLOGNAT = rsAdo("CRTLOGNAT")
lYCRTLOG0.CRTLOGUUSR = rsAdo("CRTLOGUUSR")
lYCRTLOG0.CRTLOGUAMJ = rsAdo("CRTLOGUAMJ")
lYCRTLOG0.CRTLOGUHMS = rsAdo("CRTLOGUHMS")

lYCRTLOG0.CRTLOGCPT = rsAdo("CRTLOGCPT")
lYCRTLOG0.CRTLOGETA = rsAdo("CRTLOGETA")
lYCRTLOG0.CRTLOGPLA = rsAdo("CRTLOGPLA")
lYCRTLOG0.CRTLOGPIE = rsAdo("CRTLOGPIE")
lYCRTLOG0.CRTLOGECR = rsAdo("CRTLOGECR")

lYCRTLOG0.CRTLOGTXT = rsAdo("CRTLOGTXT")


Exit Function
Error_Handler:
rsYCRTLOG0_GetBuffer = Error


End Function

Public Function rsYCRTLOG0_Init(lYCRTLOG0 As typeYCRTLOG0)

lYCRTLOG0.CRTLOGID = 0     'montant euro

lYCRTLOG0.CRTLOGNAT = ""
lYCRTLOG0.CRTLOGUUSR = ""
lYCRTLOG0.CRTLOGUAMJ = 0
lYCRTLOG0.CRTLOGUHMS = 0

lYCRTLOG0.CRTLOGETA = 0      'établissement
lYCRTLOG0.CRTLOGPLA = 0      'plan
lYCRTLOG0.CRTLOGPIE = 0      'pièce
lYCRTLOG0.CRTLOGECR = 0      'écriture
lYCRTLOG0.CRTLOGCPT = ""
lYCRTLOG0.CRTLOGTXT = ""



End Function






Public Function sqlYCRTLOG0_Insert_Transaction(newY As typeYCRTLOG0)
Dim V

V = cnSAB_Transaction("BeginTrans")
If IsNull(V) Then

    V = sqlYCRTLOG0_Insert(newY)
    
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
    Else
        V = cnSAB_Transaction("Commit")
    End If
End If
End Function
