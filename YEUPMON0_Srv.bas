Attribute VB_Name = "srvYEUPMON0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim rsADO As ADODB.Recordset
 
Type typeYEUPMON0
 
      EUPG2AOPE   As String * 3
      EUPG2ANUM   As Long
      EUPG2ACRE   As Long
      EUPG2ANEC   As Long
      
      EUPMONSEQ   As Long
      EUPMONLOT   As Long
      EUPMONID    As String * 36
      EUPMONSTA   As String * 1
      EUPMONDMO   As Long
      EUPMONHMO   As Long
      EUPMONDSW   As Long
      EUPMONHSW   As Long
      EUPMONTIC   As String * 20
      EUPMONDID   As String * 20
      EUPMONBIC   As String * 11
      EUPMONNOM   As String * 35
      EUPMONNO2   As String * 35
      EUPMONBFI   As String * 35
      EUPMONBF2   As String * 35
      EUPMONLIB   As String * 150
      EUPMONMON   As Currency
      EUPMONDEV   As String * 3
      EUPMONECH   As Long
      EUPMONPRI   As Integer

End Type
Public xYEUPMON0 As typeYEUPMON0
Public Function rsYEUPMON0_GetBuffer(rsADO As ADODB.Recordset, lYEUPMON0 As typeYEUPMON0)
On Error GoTo Error_Handler
rsYEUPMON0_GetBuffer = Null
lYEUPMON0.EUPG2AOPE = rsADO("EUPG2AOPE")
lYEUPMON0.EUPG2ANUM = rsADO("EUPG2ANUM")
lYEUPMON0.EUPG2ACRE = rsADO("EUPG2ACRE")
lYEUPMON0.EUPG2ANEC = rsADO("EUPG2ANEC")

lYEUPMON0.EUPMONID = rsADO("EUPMONID")
lYEUPMON0.EUPMONSTA = rsADO("EUPMONSTA")
lYEUPMON0.EUPMONDMO = rsADO("EUPMONDMO")
lYEUPMON0.EUPMONHMO = rsADO("EUPMONHMO")
lYEUPMON0.EUPMONDSW = rsADO("EUPMONDSW")
lYEUPMON0.EUPMONHSW = rsADO("EUPMONHSW")
lYEUPMON0.EUPMONTIC = rsADO("EUPMONTIC")
lYEUPMON0.EUPMONDID = rsADO("EUPMONDID")
lYEUPMON0.EUPMONBIC = rsADO("EUPMONBIC")
lYEUPMON0.EUPMONNOM = rsADO("EUPMONNOM")
lYEUPMON0.EUPMONNO2 = rsADO("EUPMONNO2")
lYEUPMON0.EUPMONLIB = rsADO("EUPMONLIB")
lYEUPMON0.EUPMONBFI = rsADO("EUPMONBFI")
lYEUPMON0.EUPMONBF2 = rsADO("EUPMONBF2")
lYEUPMON0.EUPMONMON = rsADO("EUPMONMON")
lYEUPMON0.EUPMONDEV = rsADO("EUPMONDEV")
lYEUPMON0.EUPMONECH = rsADO("EUPMONECH")
lYEUPMON0.EUPMONPRI = rsADO("EUPMONPRI")

Exit Function
Error_Handler:
rsYEUPMON0_GetBuffer = Error


End Function

Public Function rsYEUPMON0_Init(lYEUPMON0 As typeYEUPMON0)
lYEUPMON0.EUPMONID = ""
lYEUPMON0.EUPG2AOPE = ""
lYEUPMON0.EUPG2ANUM = 0
lYEUPMON0.EUPG2ACRE = 0
lYEUPMON0.EUPG2ANEC = 0
      
lYEUPMON0.EUPMONSEQ = 0
lYEUPMON0.EUPMONLOT = 0
lYEUPMON0.EUPMONID = ""
lYEUPMON0.EUPMONSTA = ""
lYEUPMON0.EUPMONDMO = 0
lYEUPMON0.EUPMONHMO = 0
lYEUPMON0.EUPMONDSW = 0
lYEUPMON0.EUPMONHSW = 0
lYEUPMON0.EUPMONTIC = ""
lYEUPMON0.EUPMONDID = ""
lYEUPMON0.EUPMONBIC = ""
lYEUPMON0.EUPMONNOM = ""
lYEUPMON0.EUPMONNO2 = ""
lYEUPMON0.EUPMONLIB = ""
lYEUPMON0.EUPMONBFI = ""
lYEUPMON0.EUPMONBF2 = ""
lYEUPMON0.EUPMONMON = 0
lYEUPMON0.EUPMONDEV = ""
lYEUPMON0.EUPMONECH = 0
lYEUPMON0.EUPMONPRI = 0

End Function

Public Function sqlYEUPMON0_Update(newY As typeYEUPMON0, oldY As typeYEUPMON0)
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

On Error GoTo Error_Handler
sqlYEUPMON0_Update = Null

' Contrôle  : Même clé d'accès old / new
'===================================================================================
If oldY.EUPMONID <> newY.EUPMONID Then
    sqlYEUPMON0_Update = "Erreur EUPMONID : " & newY.EUPMONID & " / " & oldY.EUPMONID
    Exit Function
End If
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where EUPMONID = '" & oldY.EUPMONID & "'" _
       & " and EUPMONSTA = '" & oldY.EUPMONSTA & "'"

newY.EUPMONDSW = DSys
newY.EUPMONHSW = time_Hms * 100

xSet = xSet & " set EUPMONHSW = " & newY.EUPMONHSW
blnUpdate = False

' Détecter les modifications
'===================================================================================
If newY.EUPMONDSW <> oldY.EUPMONDSW Then blnUpdate = True:  xSet = xSet & " , EUPMONDSW = " & newY.EUPMONDSW
If newY.EUPMONSTA <> oldY.EUPMONSTA Then blnUpdate = True:  xSet = xSet & " , EUPMONSTA = '" & newY.EUPMONSTA & "'"
If newY.EUPMONTIC <> oldY.EUPMONTIC Then blnUpdate = True:  xSet = xSet & " , EUPMONTIC = '" & newY.EUPMONTIC & "'"
If newY.EUPMONDID <> oldY.EUPMONDID Then blnUpdate = True:  xSet = xSet & " , EUPMONDID = '" & newY.EUPMONDID & "'"


    
    xSql = "update " & paramIBM_Library_SABSPE & ".YEUPMON0" & xSet & xWhere
    Call FEU_ROUGE
    Set rsADO = cnSab_Update.Execute(xSql, Nb)
    Call FEU_VERT
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYEUPMON0_Update = "Erreur màj : " & newY.EUPMONID
        Exit Function
    End If
    

Exit Function
Error_Handler:
    sqlYEUPMON0_Update = Error
End Function


