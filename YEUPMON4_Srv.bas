Attribute VB_Name = "srvYEUPMON4"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim rsAdo As ADODB.Recordset
 
Type typeYEUPMON4
 
      EUPMON2FIC   As String * 25
      EUPMON2STA   As String * 1
      EUPMON2DCR   As Long
      EUPMON2HCR   As Long
      EUPMON2DMO   As Long
      EUPMON2HMO   As Long
      EUPMON2DEN   As Long
      EUPMON2HEN   As Long
      EUPMON2DUP   As Long
      EUPMON2HUP   As Long

End Type
Public xYEUPMON4 As typeYEUPMON4
Public Function rsYEUPMON4_GetBuffer(rsAdo As ADODB.Recordset, lYEUPMON4 As typeYEUPMON4)
On Error GoTo Error_Handler
rsYEUPMON4_GetBuffer = Null

lYEUPMON4.EUPMON2FIC = rsAdo("EUPMON2FIC")
lYEUPMON4.EUPMON2STA = rsAdo("EUPMON2STA")
lYEUPMON4.EUPMON2DCR = rsAdo("EUPMON2DCR")
lYEUPMON4.EUPMON2HCR = rsAdo("EUPMON2HCR")
lYEUPMON4.EUPMON2DMO = rsAdo("EUPMON2DMO")
lYEUPMON4.EUPMON2HMO = rsAdo("EUPMON2HMO")
lYEUPMON4.EUPMON2DEN = rsAdo("EUPMON2DEN")
lYEUPMON4.EUPMON2HEN = rsAdo("EUPMON2HEN")
lYEUPMON4.EUPMON2DUP = rsAdo("EUPMON2DUP")
lYEUPMON4.EUPMON2HUP = rsAdo("EUPMON2HUP")

Exit Function
Error_Handler:
rsYEUPMON4_GetBuffer = Error


End Function

Public Function rsYEUPMON4_Init(lYEUPMON4 As typeYEUPMON4)
lYEUPMON4.EUPMON2FIC = ""
lYEUPMON4.EUPMON2STA = ""
lYEUPMON4.EUPMON2DCR = 0
lYEUPMON4.EUPMON2HCR = 0
lYEUPMON4.EUPMON2DMO = 0
lYEUPMON4.EUPMON2HMO = 0
lYEUPMON4.EUPMON2DEN = 0
lYEUPMON4.EUPMON2HEN = 0
lYEUPMON4.EUPMON2DUP = 0
lYEUPMON4.EUPMON2HUP = 0

End Function

Public Function sqlYEUPMON4_Update(newY As typeYEUPMON4, oldY As typeYEUPMON4)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

On Error GoTo Error_Handler
sqlYEUPMON4_Update = Null

' Contrôle  : Même clé d'accès old / new
'===================================================================================
If oldY.EUPMON2FIC <> newY.EUPMON2FIC Then
    sqlYEUPMON4_Update = "Erreur EUPMON2FIC : " & newY.EUPMON2FIC & " / " & oldY.EUPMON2FIC
    Exit Function
End If
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
'===================================================================================

xWhere = " where EUPMON2FIC = '" & oldY.EUPMON2FIC & "'" _
       & " and EUPMON2STA = '" & oldY.EUPMON2STA & "'"

newY.EUPMON2DUP = DSys
newY.EUPMON2HUP = time_Hms * 100

xSet = xSet & " set EUPMON2HUP = " & newY.EUPMON2HUP
blnUpdate = False

' Détecter les modifications
'===================================================================================
If newY.EUPMON2DUP <> oldY.EUPMON2DUP Then blnUpdate = True:  xSet = xSet & " , EUPMON2DUP = " & newY.EUPMON2DUP
If newY.EUPMON2STA <> oldY.EUPMON2STA Then blnUpdate = True:  xSet = xSet & " , EUPMON2STA = '" & newY.EUPMON2STA & "'"

If newY.EUPMON2DCR <> oldY.EUPMON2DCR Then blnUpdate = True:  xSet = xSet & " , EUPMON2DCR = " & newY.EUPMON2DCR
If newY.EUPMON2DCR <> oldY.EUPMON2DCR Then blnUpdate = True:  xSet = xSet & " , EUPMON2DCR = " & newY.EUPMON2DCR

If newY.EUPMON2DMO <> oldY.EUPMON2DMO Then blnUpdate = True:  xSet = xSet & " , EUPMON2DMO = " & newY.EUPMON2DMO
If newY.EUPMON2DMO <> oldY.EUPMON2DMO Then blnUpdate = True:  xSet = xSet & " , EUPMON2DMO = " & newY.EUPMON2DMO

If newY.EUPMON2DEN <> oldY.EUPMON2DEN Then blnUpdate = True:  xSet = xSet & " , EUPMON2DEN = " & newY.EUPMON2DEN
If newY.EUPMON2DEN <> oldY.EUPMON2DEN Then blnUpdate = True:  xSet = xSet & " , EUPMON2DEN = " & newY.EUPMON2DEN


    
    xSQL = "update " & paramIBM_Library_SABSPE & ".YEUPMON4" & xSet & xWhere
    Call FEU_ROUGE
    Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
    Call FEU_VERT
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYEUPMON4_Update = "Erreur màj : " & newY.EUPMON2FIC
        Exit Function
    End If
    

Exit Function
Error_Handler:
    sqlYEUPMON4_Update = Error
End Function





