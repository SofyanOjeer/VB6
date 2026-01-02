Attribute VB_Name = "srvYGUIMAD0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim rsAdo As ADODB.Recordset
 
Type typeYGUIMAD0
 
      GUIMADID    As Long         'IDENTIFICATION
      GUIESPOPE   As String * 3   'code opération
      GUIESPNAT   As String * 3   'nature opération
      GUIESPDOS   As Long         'N° opération
      GUIESPMON   As Currency     'montant
      GUIESPDEV   As String * 3   'devise
      GUIESPCP1   As String * 20  'compte
      GUIESPCL1   As String * 7   'client
      GUIESPTI1   As String * 30  'bénéficiaire
      GUIESPDJO   As Long         'date création
      
      GUIMADMON   As Currency     'montant
      GUIMADTDO   As String * 30  'nom DO
      GUIMADTIN   As String * 30  'nom intermédiaire
      GUIMADMOT   As String * 30  'motif
      GUIMADLIEN  As Long         'nb document
      GUIMADSTA   As String * 1   'statut
      GUIMADUPDS  As Long         'Sequence mise à jour
      GUIMADUSR   As String * 10  'Utilisateur
End Type
Public xYGUIMAD0 As typeYGUIMAD0
Public Function sqlYGUIMAD0_Insert(newY As typeYGUIMAD0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlYGUIMAD0_Insert = Null

xSet = " (GUIMADID"
xValues = " values(" & newY.GUIMADID

' Détecter les modifications
'===================================================================================
If newY.GUIESPDOS <> 0 Then xSet = xSet & ",GUIESPDOS": xValues = xValues & " ," & newY.GUIESPDOS
If newY.GUIESPMON <> 0 Then xSet = xSet & ",GUIESPMON": xValues = xValues & " ," & cur_P(newY.GUIESPMON)
If newY.GUIESPDJO <> 0 Then xSet = xSet & ",GUIESPDJO": xValues = xValues & " ," & newY.GUIESPDJO
If newY.GUIMADMON <> 0 Then xSet = xSet & ",GUIMADMON": xValues = xValues & " ," & cur_P(newY.GUIMADMON)
If newY.GUIMADLIEN <> 0 Then xSet = xSet & ",GUIMADLIEN": xValues = xValues & " ," & newY.GUIMADLIEN

If Trim(newY.GUIESPOPE) <> "" Then xSet = xSet & ",GUIESPOPE": xValues = xValues & " ,'" & newY.GUIESPOPE & "'"
If Trim(newY.GUIESPNAT) <> "" Then xSet = xSet & ",GUIESPNAT": xValues = xValues & " ,'" & newY.GUIESPNAT & "'"
If Trim(newY.GUIESPDEV) <> "" Then xSet = xSet & ",GUIESPDEV": xValues = xValues & " ,'" & newY.GUIESPDEV & "'"
If Trim(newY.GUIESPCP1) <> "" Then xSet = xSet & ",GUIESPCP1": xValues = xValues & " ,'" & Trim(newY.GUIESPCP1) & "'"
If Trim(newY.GUIESPCL1) <> "" Then xSet = xSet & ",GUIESPCL1": xValues = xValues & " ,'" & Replace(Trim(newY.GUIESPCL1), "'", "''") & "'"
If Trim(newY.GUIESPTI1) <> "" Then xSet = xSet & ",GUIESPTI1": xValues = xValues & " ,'" & Replace(Trim(newY.GUIESPTI1), "'", "''") & "'"
If Trim(newY.GUIMADTDO) <> "" Then xSet = xSet & ",GUIMADTDO": xValues = xValues & " ,'" & Replace(Trim(newY.GUIMADTDO), "'", "''") & "'"
If Trim(newY.GUIMADTIN) <> "" Then xSet = xSet & ",GUIMADTIN": xValues = xValues & " ,'" & Replace(Trim(newY.GUIMADTIN), "'", "''") & "'"
If Trim(newY.GUIMADMOT) <> "" Then xSet = xSet & ",GUIMADMOT": xValues = xValues & " ,'" & Replace(Trim(newY.GUIMADMOT), "'", "''") & "'"
If Trim(newY.GUIMADSTA) <> "" Then xSet = xSet & ",GUIMADSTA": xValues = xValues & " ,'" & newY.GUIMADSTA & "'"

newY.GUIMADUSR = usrName_UCase
xSet = xSet & ",GUIMADUSR": xValues = xValues & " ,'" & usrName_UCase & "'"
Call FEU_ROUGE
xSQL = "Insert into " & paramIBM_Library_SABSPE & ".YGUIMAD0" & xSet & ")" & xValues & ")"

Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYGUIMAD0_Insert = "Erreur màj : " & newY.GUIMADID
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYGUIMAD0_Insert = Error
End Function

Public Function sqlYGUIMAD0_Update(newY As typeYGUIMAD0, oldY As typeYGUIMAD0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

On Error GoTo Error_Handler
sqlYGUIMAD0_Update = Null

' Contrôle  : Même clé d'accès old / new
'===================================================================================
If oldY.GUIMADID <> newY.GUIMADID Then
    sqlYGUIMAD0_Update = "Erreur GUIMADID : " & newY.GUIMADID & " / " & oldY.GUIMADID
    Exit Function
End If
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where GUIMADID = " & oldY.GUIMADID & " and GUIMADUPDS = " & oldY.GUIMADUPDS

newY.GUIMADUPDS = newY.GUIMADUPDS + 1
xSet = xSet & " set GUIMADUPDS = " & newY.GUIMADUPDS
blnUpdate = False

' Détecter les modifications
'===================================================================================
If newY.GUIESPDOS <> oldY.GUIESPDOS Then blnUpdate = True: xSet = xSet & " , GUIESPDOS = " & newY.GUIESPDOS
If newY.GUIESPMON <> oldY.GUIESPMON Then blnUpdate = True:  xSet = xSet & " , GUIESPMON = '" & cur_P(newY.GUIESPMON) & "'"
If newY.GUIESPDJO <> oldY.GUIESPDJO Then blnUpdate = True: xSet = xSet & " , GUIESPDJO = " & newY.GUIESPDJO
If newY.GUIMADMON <> oldY.GUIMADMON Then blnUpdate = True:  xSet = xSet & " , GUIMADMON = '" & cur_P(newY.GUIMADMON) & "'"
If newY.GUIMADLIEN <> oldY.GUIMADLIEN Then blnUpdate = True: xSet = xSet & " , GUIMADLIEN = " & newY.GUIMADLIEN


If newY.GUIESPOPE <> oldY.GUIESPOPE Then blnUpdate = True:  xSet = xSet & " , GUIESPOPE = '" & newY.GUIESPOPE & "'"
If newY.GUIESPNAT <> oldY.GUIESPNAT Then blnUpdate = True: xSet = xSet & " , GUIESPNAT = " & newY.GUIESPNAT
If newY.GUIESPDEV <> oldY.GUIESPDEV Then blnUpdate = True:  xSet = xSet & " , GUIESPDEV= '" & newY.GUIESPDEV & "'"
If newY.GUIESPCP1 <> oldY.GUIESPCP1 Then blnUpdate = True:  xSet = xSet & " , GUIESPCP1 = '" & Trim(newY.GUIESPCP1) & "'"
If newY.GUIESPCL1 <> oldY.GUIESPCL1 Then blnUpdate = True:  xSet = xSet & " , GUIESPCL1= '" & newY.GUIESPCL1 & "'"
If newY.GUIESPTI1 <> oldY.GUIESPTI1 Then blnUpdate = True:  xSet = xSet & " , GUIESPTI1 = '" & Replace(Trim(newY.GUIESPTI1), "'", "''") & "'"
If newY.GUIMADTDO <> oldY.GUIMADTDO Then blnUpdate = True:  xSet = xSet & " , GUIMADTDO = '" & Replace(Trim(newY.GUIMADTDO), "'", "''") & "'"
If newY.GUIMADTIN <> oldY.GUIMADTIN Then blnUpdate = True:  xSet = xSet & " , GUIMADTIN = '" & Replace(Trim(newY.GUIMADTIN), "'", "''") & "'"
If newY.GUIMADMOT <> oldY.GUIMADMOT Then blnUpdate = True:  xSet = xSet & " , GUIMADMOT = '" & Replace(Trim(newY.GUIMADMOT), "'", "''") & "'"
If newY.GUIMADSTA <> oldY.GUIMADSTA Then blnUpdate = True:  xSet = xSet & " , GUIMADSTA = '" & newY.GUIMADSTA & "'"

newY.GUIMADUSR = usrName_UCase
xSet = xSet & " , GUIMADUSR = '" & usrName_UCase & "'"
If newY.GUIMADID < 0 Then blnUpdate = True  ' records techniques

If blnUpdate Then
    
    xSQL = "update " & paramIBM_Library_SABSPE & ".YGUIMAD0" & xSet & xWhere
    Call FEU_ROUGE
    Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
    Call FEU_VERT
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYGUIMAD0_Update = "Erreur màj : " & newY.GUIMADID
        Exit Function
    End If
    
End If

Exit Function
Error_Handler:
    sqlYGUIMAD0_Update = Error
End Function

Public Function sqlYGUIMAD0_Init(newY As typeYGUIMAD0)
Dim X As String, xSQL As String, Nb As Long
Dim xWhere As String, xSet As String
Dim xxx As typeYGUIMAD0

On Error GoTo Error_Handler
sqlYGUIMAD0_Init = Null

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YGUIMAD0" & " where  GUIMADID =  -1"
Set rsAdo = cnSab_Update.Execute(xSQL, Nb)

xxx.GUIMADUPDS = rsAdo("GUIMADUPDS")
newY.GUIMADID = rsAdo("GUIESPDOS") + 1
newY.GUIMADUPDS = 0

' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where GUIMADID = -1" & " and GUIMADUPDS = " & xxx.GUIMADUPDS

xSet = " set GUIMADUPDS = " & xxx.GUIMADUPDS + 1 & " , GUIESPDOS = " & newY.GUIMADID


xSQL = "update " & paramIBM_Library_SABSPE & ".YGUIMAD0" & xSet & xWhere
Call FEU_ROUGE
Set rsAdo = cnSab_Update.Execute(xSQL, Nb)
Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYGUIMAD0_Init = "Erreur màj : " & newY.GUIMADID
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYGUIMAD0_Init = Error
End Function

Public Function rsYGUIMAD0_GetBuffer(rsAdo As ADODB.Recordset, lYGUIMAD0 As typeYGUIMAD0)
On Error GoTo Error_Handler
rsYGUIMAD0_GetBuffer = Null
lYGUIMAD0.GUIMADID = rsAdo("GUIMADID")
lYGUIMAD0.GUIESPOPE = rsAdo("GUIESPOPE")
lYGUIMAD0.GUIESPDOS = rsAdo("GUIESPDOS")
lYGUIMAD0.GUIESPNAT = rsAdo("GUIESPNAT")
lYGUIMAD0.GUIESPMON = rsAdo("GUIESPMON")
lYGUIMAD0.GUIESPDEV = rsAdo("GUIESPDEV")
lYGUIMAD0.GUIESPCP1 = rsAdo("GUIESPCP1")
lYGUIMAD0.GUIESPCL1 = rsAdo("GUIESPCL1")
lYGUIMAD0.GUIESPTI1 = rsAdo("GUIESPTI1")
lYGUIMAD0.GUIESPDJO = rsAdo("GUIESPDJO")

lYGUIMAD0.GUIMADMON = rsAdo("GUIMADMON")
lYGUIMAD0.GUIMADTDO = rsAdo("GUIMADTDO")
lYGUIMAD0.GUIMADTIN = rsAdo("GUIMADTIN")
lYGUIMAD0.GUIMADMOT = rsAdo("GUIMADMOT")
lYGUIMAD0.GUIMADLIEN = rsAdo("GUIMADLIEN")
lYGUIMAD0.GUIMADSTA = rsAdo("GUIMADSTA")

lYGUIMAD0.GUIMADUPDS = rsAdo("GUIMADUPDS")
lYGUIMAD0.GUIMADUSR = rsAdo("GUIMADUSR")

Exit Function
Error_Handler:
rsYGUIMAD0_GetBuffer = Error


End Function

Public Function rsYGUIMAD0_Init(lYGUIMAD0 As typeYGUIMAD0)
lYGUIMAD0.GUIMADID = 0
lYGUIMAD0.GUIESPDOS = 0
lYGUIMAD0.GUIESPNAT = 0
lYGUIMAD0.GUIESPOPE = ""
lYGUIMAD0.GUIESPMON = 0
lYGUIMAD0.GUIESPDEV = ""
lYGUIMAD0.GUIESPCP1 = ""
lYGUIMAD0.GUIESPCL1 = ""
lYGUIMAD0.GUIESPTI1 = ""
lYGUIMAD0.GUIESPDJO = 0
lYGUIMAD0.GUIMADMON = 0
lYGUIMAD0.GUIMADTDO = ""
lYGUIMAD0.GUIMADTIN = ""
lYGUIMAD0.GUIMADMOT = ""
lYGUIMAD0.GUIMADLIEN = 0
lYGUIMAD0.GUIMADSTA = ""
lYGUIMAD0.GUIMADUPDS = 0

End Function


