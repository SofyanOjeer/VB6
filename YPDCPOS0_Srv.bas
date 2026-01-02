Attribute VB_Name = "srvYPDCPOS0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim rsAdo As ADODB.Recordset
 
Type typeYPDCPOS0
 
      PDCPOSDTR     As String * 8       'date comptable
      PDCPOSDEV     As String * 3       'devise
      PDCPOSPOSD    As Currency         'position devise
      PDCPOSPOSE    As Currency         'position euro
      PDCPOSPRIX    As Double           'prix de la position
      PDCPOSFIXT    As Double           'fixing
      PDCPOSFIXD    As String * 8       'date fixing
      PDCPOSPNL     As Currency         'pp
      PDCPOSRPC     As Currency         'RPC
      PDCPOSUPDS    As Long             'SéQUENCE UPD
      PDCPOSTERD    As Currency         'position devise
      PDCPOSTERE    As Currency         'position euro
      PDCPOSSWPD    As Currency         'position devise
      PDCPOSSWPE    As Currency         'position euro
     
      
End Type
Public xYPDCPOS0 As typeYPDCPOS0

Public Function sqlYPDCPOS0_DeleteW(lWhere As String, Nb As Long)
Dim X As String, xSQL As String
'Dim nb As Long

On Error GoTo Error_Handler
sqlYPDCPOS0_DeleteW = Null
    
xSQL = "delete from " & paramIBM_Library_SABSPE_XXX & ".YPDCPOS0 " & lWhere
Call FEU_ROUGE
Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
Call FEU_VERT
Exit Function
Error_Handler:
    sqlYPDCPOS0_DeleteW = Error
End Function


Public Function sqlYPDCPOS0_Insert(newY As typeYPDCPOS0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlYPDCPOS0_Insert = Null

xSet = " (PDCPOSDTR,PDCPOSDEV"
xValues = " values('" & newY.PDCPOSDTR & "' ,'" & newY.PDCPOSDEV & "'"

' Détecter les modifications
'===================================================================================
If newY.PDCPOSPOSD <> 0 Then xSet = xSet & ",PDCPOSPOSD": xValues = xValues & " ," & Replace(newY.PDCPOSPOSD, ",", ".")
If newY.PDCPOSPOSE <> 0 Then xSet = xSet & ",PDCPOSPOSE": xValues = xValues & " ," & Replace(newY.PDCPOSPOSE, ",", ".")
If newY.PDCPOSPRIX <> 0 Then xSet = xSet & ",PDCPOSPRIX": xValues = xValues & " ," & Replace(newY.PDCPOSPRIX, ",", ".")
If newY.PDCPOSFIXT <> 0 Then xSet = xSet & ",PDCPOSFIXT": xValues = xValues & " ," & Replace(newY.PDCPOSFIXT, ",", ".")
If newY.PDCPOSPNL <> 0 Then xSet = xSet & ",PDCPOSPNL": xValues = xValues & " ," & Replace(newY.PDCPOSPNL, ",", ".")
If newY.PDCPOSRPC <> 0 Then xSet = xSet & ",PDCPOSRPC": xValues = xValues & " ," & Replace(newY.PDCPOSRPC, ",", ".")
If newY.PDCPOSTERD <> 0 Then xSet = xSet & ",PDCPOSTERD": xValues = xValues & " ," & Replace(newY.PDCPOSTERD, ",", ".")
If newY.PDCPOSTERE <> 0 Then xSet = xSet & ",PDCPOSTERE": xValues = xValues & " ," & Replace(newY.PDCPOSTERE, ",", ".")
If newY.PDCPOSSWPD <> 0 Then xSet = xSet & ",PDCPOSSWPD": xValues = xValues & " ," & Replace(newY.PDCPOSSWPD, ",", ".")
If newY.PDCPOSSWPE <> 0 Then xSet = xSet & ",PDCPOSSWPE": xValues = xValues & " ," & Replace(newY.PDCPOSSWPE, ",", ".")

If Trim(newY.PDCPOSFIXD) <> "" Then xSet = xSet & ",PDCPOSFIXD": xValues = xValues & " ,'" & newY.PDCPOSFIXD & "'"
Call FEU_ROUGE
xSQL = "Insert into " & paramIBM_Library_SABSPE_XXX & ".YPDCPOS0" & xSet & ")" & xValues & ")"

Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYPDCPOS0_Insert = "Erreur màj : " & newY.PDCPOSDTR & newY.PDCPOSDEV
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYPDCPOS0_Insert = Error
End Function

Public Function sqlYPDCPOS0_Update(newY As typeYPDCPOS0, oldY As typeYPDCPOS0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

On Error GoTo Error_Handler
sqlYPDCPOS0_Update = Null

' Contrôle  : Même clé d'accès old / new
'===================================================================================
If oldY.PDCPOSDTR <> newY.PDCPOSDTR _
Or oldY.PDCPOSDEV <> newY.PDCPOSDEV Then
    sqlYPDCPOS0_Update = "Erreur PDCPOSDEV: " & newY.PDCPOSDTR & "." & oldY.PDCPOSDEV
    Exit Function
End If
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where PDCPOSDTR = '" & oldY.PDCPOSDTR & "'" _
       & " and PDCPOSDEV = '" & oldY.PDCPOSDEV & "'" _
       & " and PDCPOSUPDS = " & oldY.PDCPOSUPDS

newY.PDCPOSUPDS = newY.PDCPOSUPDS + 1
xSet = xSet & " set PDCPOSUPDS = " & newY.PDCPOSUPDS
blnUpdate = False

' Détecter les modifications
'===================================================================================
If newY.PDCPOSPOSD <> oldY.PDCPOSPOSD Then blnUpdate = True:  xSet = xSet & " , PDCPOSPOSD = " & cur_P(newY.PDCPOSPOSD)
If newY.PDCPOSPOSE <> oldY.PDCPOSPOSE Then blnUpdate = True:  xSet = xSet & " , PDCPOSPOSE = " & cur_P(newY.PDCPOSPOSE)
If newY.PDCPOSPRIX <> oldY.PDCPOSPRIX Then blnUpdate = True:  xSet = xSet & " , PDCPOSPRIX = " & cur_P(newY.PDCPOSPRIX)
If newY.PDCPOSFIXT <> oldY.PDCPOSFIXT Then blnUpdate = True:  xSet = xSet & " , PDCPOSFIXT = " & cur_P(newY.PDCPOSFIXT)
If newY.PDCPOSPNL <> oldY.PDCPOSPNL Then blnUpdate = True: xSet = xSet & " , PDCPOSPNL = " & cur_P(newY.PDCPOSPNL)
If newY.PDCPOSRPC <> oldY.PDCPOSRPC Then blnUpdate = True: xSet = xSet & " , PDCPOSRPC = " & cur_P(newY.PDCPOSRPC)
If newY.PDCPOSTERD <> oldY.PDCPOSTERD Then blnUpdate = True:  xSet = xSet & " , PDCPOSTERD = " & cur_P(newY.PDCPOSTERD)
If newY.PDCPOSTERE <> oldY.PDCPOSTERE Then blnUpdate = True:  xSet = xSet & " , PDCPOSTERE = " & cur_P(newY.PDCPOSTERE)
If newY.PDCPOSSWPD <> oldY.PDCPOSSWPD Then blnUpdate = True:  xSet = xSet & " , PDCPOSSWPD = " & cur_P(newY.PDCPOSSWPD)
If newY.PDCPOSSWPE <> oldY.PDCPOSSWPE Then blnUpdate = True:  xSet = xSet & " , PDCPOSSWPE = " & cur_P(newY.PDCPOSSWPE)

If newY.PDCPOSFIXD <> oldY.PDCPOSFIXD Then blnUpdate = True:  xSet = xSet & " , PDCPOSFIXD = '" & newY.PDCPOSFIXD & "'"


If blnUpdate Then
    
    xSQL = "update " & paramIBM_Library_SABSPE_XXX & ".YPDCPOS0" & xSet & xWhere
    Call FEU_ROUGE
    Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
    Call FEU_VERT
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYPDCPOS0_Update = "Erreur màj : " & newY.PDCPOSDEV
        Exit Function
    End If
    
End If

Exit Function
Error_Handler:
    sqlYPDCPOS0_Update = Error
End Function



Public Function rsYPDCPOS0_GetBuffer(rsAdo As ADODB.Recordset, lYPDCPOS0 As typeYPDCPOS0)
On Error GoTo Error_Handler
rsYPDCPOS0_GetBuffer = Null

lYPDCPOS0.PDCPOSDTR = rsAdo("PDCPOSDTR")
lYPDCPOS0.PDCPOSDEV = rsAdo("PDCPOSDEV")
lYPDCPOS0.PDCPOSPOSD = rsAdo("PDCPOSPOSD")
lYPDCPOS0.PDCPOSPOSE = rsAdo("PDCPOSPOSE")
lYPDCPOS0.PDCPOSPRIX = rsAdo("PDCPOSPRIX")
lYPDCPOS0.PDCPOSFIXT = rsAdo("PDCPOSFIXT")
lYPDCPOS0.PDCPOSFIXD = rsAdo("PDCPOSFIXD")
lYPDCPOS0.PDCPOSPNL = rsAdo("PDCPOSPNL")
lYPDCPOS0.PDCPOSRPC = rsAdo("PDCPOSRPC")
lYPDCPOS0.PDCPOSUPDS = rsAdo("PDCPOSUPDS")
lYPDCPOS0.PDCPOSTERD = rsAdo("PDCPOSTERD")
lYPDCPOS0.PDCPOSTERE = rsAdo("PDCPOSTERE")
lYPDCPOS0.PDCPOSSWPD = rsAdo("PDCPOSSWPD")
lYPDCPOS0.PDCPOSSWPE = rsAdo("PDCPOSSWPE")

Exit Function
Error_Handler:
rsYPDCPOS0_GetBuffer = Error


End Function

Public Function rsYPDCPOS0_Init(lYPDCPOS0 As typeYPDCPOS0)

lYPDCPOS0.PDCPOSDTR = ""
lYPDCPOS0.PDCPOSDEV = ""
lYPDCPOS0.PDCPOSPOSD = 0
lYPDCPOS0.PDCPOSPOSE = 0
lYPDCPOS0.PDCPOSPRIX = 0
lYPDCPOS0.PDCPOSFIXT = 0
lYPDCPOS0.PDCPOSFIXD = ""
lYPDCPOS0.PDCPOSPNL = 0
lYPDCPOS0.PDCPOSRPC = 0
lYPDCPOS0.PDCPOSUPDS = 0
lYPDCPOS0.PDCPOSTERD = 0
lYPDCPOS0.PDCPOSTERE = 0
lYPDCPOS0.PDCPOSSWPD = 0
lYPDCPOS0.PDCPOSSWPE = 0
End Function





