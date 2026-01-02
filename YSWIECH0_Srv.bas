Attribute VB_Name = "srvYSWIECH0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
'Dim rsSabX As New ADODB.Recordset
Dim rsADO As ADODB.Recordset
Public mYSWIECH0_SWISABSWID As Long, oldYSWIECH0_SWISABSWID As Long

Type typeYSWIECH0
 
      SWIECHSWID   As Long
      SWIECHSEQ0   As Long
      
      SWIECHSER    As String
      SWIECHSSE    As String
      SWIECHOPEC   As String
      SWIECHOPEN   As Long
      SWIECHWMTK   As String
      SWIECHWES    As String
      SWIECHWBIC   As String
      SWIECHDECH   As Long
      
      SWIECHSTA    As String
      SWIECHSTAK   As String
      SWIECHSWIX   As Long
      SWIECHSWIL   As Long
       
      SWIECHWDEV   As String
      SWIECHWMTD   As Currency
      SWIECHWN20   As String
      SWIECHWL20   As String
      SWIECHW22C   As String
      SWIECHW52A   As String
      SWIECHW57A   As String
      SWIECHW30V   As String
      SWIECHSENS   As String
      
      SWIECHYAMJ   As Long
      SWIECHYHMS   As Long
      SWIECHYUSR   As String
      SWIECHYVER   As String
     
   
End Type


Public Function sqlYSWIECH0_Delete(oldY As typeYSWIECH0)
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

On Error GoTo Error_Handler
sqlYSWIECH0_Delete = Null
' Vérifier si l'enregistrement n'a pas été modifié entre le 'Select ...... Update'
' incrémenter la séquence de maj
'===================================================================================

xWhere = " where SWIECHSWID = " & oldY.SWIECHSWID & " and SWIECHSEQ0 = " & oldY.SWIECHSEQ0


'===================================================================================

    
    xSql = "delete from " & paramIBM_Library_SABSPE_XXX & ".YSWIECH0" & xWhere
    'Call FEU_ROUGE
    Set rsADO = cnSab_Update.Execute(xSql, Nb)
    'Call FEU_VERT
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYSWIECH0_Delete = "Erreur màj : " & xWhere
        Exit Function
    End If
    


Exit Function
Error_Handler:
    sqlYSWIECH0_Delete = Error
End Function

Public Function sqlYSWIECH0_Update(newY As typeYSWIECH0, oldY As typeYSWIECH0)
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String
Dim blnUpdate As Boolean

On Error GoTo Error_Handler
sqlYSWIECH0_Update = Null

'===================================================================================

xWhere = " where SWIECHSWID = " & oldY.SWIECHSWID & " and SWIECHSEQ0 = " & newY.SWIECHSEQ0 & " and SWIECHYVER = " & newY.SWIECHYVER
xSet = " set"
blnUpdate = False
newY.SWIECHYVER = newY.SWIECHYVER + 1

' Détecter les modifications
'===================================================================================
'If newY.SWIECHSWID <> oldY.SWIECHSWID Then blnUpdate = True:  xSet = xSet & " , SWIECHSWID = " & newY.SWIECHSWID
'If newY.SWIECHSEQ0 <> oldY.SWIECHSEQ0 Then blnUpdate = True:  xSet = xSet & " , SWIECHSEQ0 = " & newY.SWIECHSEQ0

If newY.SWIECHOPEN <> oldY.SWIECHOPEN Then blnUpdate = True:  xSet = xSet & " , SWIECHOPEN = " & newY.SWIECHOPEN

If newY.SWIECHDECH <> oldY.SWIECHDECH Then blnUpdate = True:  xSet = xSet & " , SWIECHDECH = " & newY.SWIECHDECH
If newY.SWIECHSWIX <> oldY.SWIECHSWIX Then blnUpdate = True:  xSet = xSet & " , SWIECHSWIX = " & newY.SWIECHSWIX
If newY.SWIECHSWIL <> oldY.SWIECHSWIL Then blnUpdate = True:  xSet = xSet & " , SWIECHSWIL = " & newY.SWIECHSWIL
If newY.SWIECHYAMJ <> oldY.SWIECHYAMJ Then blnUpdate = True:  xSet = xSet & " , SWIECHYAMJ = " & newY.SWIECHYAMJ
If newY.SWIECHYHMS <> oldY.SWIECHYHMS Then blnUpdate = True:  xSet = xSet & " , SWIECHYHMS = " & newY.SWIECHYHMS
If newY.SWIECHYVER <> oldY.SWIECHYVER Then blnUpdate = True:  xSet = xSet & " , SWIECHYVER = " & newY.SWIECHYVER

If newY.SWIECHWMTD <> oldY.SWIECHWMTD Then blnUpdate = True:  xSet = xSet & " , SWIECHWMTD = " & cur_P(newY.SWIECHWMTD)

If newY.SWIECHSER <> oldY.SWIECHSER Then blnUpdate = True:  xSet = xSet & " , SWIECHSER = '" & newY.SWIECHSER & "'"
If newY.SWIECHSSE <> oldY.SWIECHSSE Then blnUpdate = True:  xSet = xSet & " , SWIECHSSE = '" & newY.SWIECHSSE & "'"
If newY.SWIECHOPEC <> oldY.SWIECHOPEC Then blnUpdate = True:  xSet = xSet & " , SWIECHOPEC = '" & newY.SWIECHOPEC & "'"
If newY.SWIECHWES <> oldY.SWIECHWES Then blnUpdate = True:  xSet = xSet & " , SWIECHWES = '" & newY.SWIECHWES & "'"
If newY.SWIECHWMTK <> oldY.SWIECHWMTK Then blnUpdate = True:  xSet = xSet & " , SWIECHWMTK = '" & newY.SWIECHWMTK & "'"
If newY.SWIECHWBIC <> oldY.SWIECHWBIC Then blnUpdate = True:  xSet = xSet & " , SWIECHWBIC = '" & newY.SWIECHWBIC & "'"
If newY.SWIECHWDEV <> oldY.SWIECHWDEV Then blnUpdate = True:  xSet = xSet & " , SWIECHWDEV = '" & Replace(newY.SWIECHWDEV, "'", "''") & "'"
If newY.SWIECHWN20 <> oldY.SWIECHWN20 Then blnUpdate = True:  xSet = xSet & " , SWIECHWN20 = '" & newY.SWIECHWN20 & "'"
If newY.SWIECHWL20 <> oldY.SWIECHWL20 Then blnUpdate = True:  xSet = xSet & " , SWIECHWL20 = '" & newY.SWIECHWL20 & "'"
If newY.SWIECHW52A <> oldY.SWIECHW52A Then blnUpdate = True:  xSet = xSet & " , SWIECHW52A = '" & newY.SWIECHW52A & "'"
If newY.SWIECHW57A <> oldY.SWIECHW57A Then blnUpdate = True:  xSet = xSet & " , SWIECHW57A = '" & newY.SWIECHW57A & "'"
If newY.SWIECHW30V <> oldY.SWIECHW30V Then blnUpdate = True:  xSet = xSet & " , SWIECHW30V = '" & newY.SWIECHW30V & "'"
If newY.SWIECHSENS <> oldY.SWIECHSENS Then blnUpdate = True:  xSet = xSet & " , SWIECHSENS = '" & newY.SWIECHSENS & "'"
If newY.SWIECHW22C <> oldY.SWIECHW22C Then blnUpdate = True:  xSet = xSet & " , SWIECHW22C = '" & newY.SWIECHW22C & "'"
If newY.SWIECHSTA <> oldY.SWIECHSTA Then blnUpdate = True:  xSet = xSet & " , SWIECHSTA = '" & newY.SWIECHSTA & "'"
If newY.SWIECHSTAK <> oldY.SWIECHSTAK Then blnUpdate = True:  xSet = xSet & " , SWIECHSTAK = '" & newY.SWIECHSTAK & "'"
If newY.SWIECHYUSR <> oldY.SWIECHYUSR Then blnUpdate = True:  xSet = xSet & " , SWIECHYUSR = '" & newY.SWIECHYUSR & "'"

If blnUpdate Then
    Mid$(xSet, 1, 6) = " set  "
    xSql = "update " & paramIBM_Library_SABSPE_XXX & ".YSWIECH0" & xSet & xWhere
    'Call FEU_ROUGE
    Set rsADO = cnSab_Update.Execute(xSql, Nb)
    'Call FEU_VERT
    ' Tester si la mise à jour a été effectuée
    '===================================================================================
    
    If Nb = 0 Then
        sqlYSWIECH0_Update = "Erreur màj : " & newY.SWIECHOPEN
        Exit Function
    End If
    
End If

Exit Function
Error_Handler:
    sqlYSWIECH0_Update = Error
End Function
Public Function sqlYSWIECH0_Update_Field(oldY As typeYSWIECH0, lSQL_Set As String)
Dim xSql As String, Nb As Long

On Error GoTo Error_Handler
sqlYSWIECH0_Update_Field = Null



xSql = "update " & paramIBM_Library_SABSPE_XXX & ".YSWIECH0 " & lSQL_Set & "" _
     & " where SWIECHSWID = " & oldY.SWIECHSWID _
     & " and SWIECHSEQ0 = " & oldY.SWIECHSEQ0 _
     & " and SWIECHYVER = " & oldY.SWIECHYVER
     
'Call FEU_ROUGE
Set rsADO = cnSab_Update.Execute(xSql, Nb)
'Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYSWIECH0_Update_Field = "Erreur màj : " & oldY.SWIECHSWID & " - " & oldY.SWIECHSEQ0
    Exit Function
End If
    

Exit Function
Error_Handler:
    sqlYSWIECH0_Update_Field = Error
End Function


Public Function sqlYSWIECH0_Insert(newY As typeYSWIECH0)
Dim V
Dim X As String, xSql As String, Nb As Long
Dim xWhere As String, xSet As String, xValues As String

On Error GoTo Error_Handler
sqlYSWIECH0_Insert = Null
xSet = " (SWIECHSWID "
xValues = " values(" & newY.SWIECHSWID

' Détecter les modifications
'===================================================================================
'If newY.SWIECHSWID <> 0 Then xSet = xSet & ",SWIECHSWID": xValues = xValues & " ," & newY.SWIECHSWID
If newY.SWIECHOPEN <> 0 Then xSet = xSet & ",SWIECHOPEN": xValues = xValues & " ," & newY.SWIECHOPEN
If newY.SWIECHSEQ0 <> 0 Then xSet = xSet & ",SWIECHSEQ0": xValues = xValues & " ," & newY.SWIECHSEQ0
If newY.SWIECHDECH <> 0 Then xSet = xSet & ",SWIECHDECH": xValues = xValues & " ," & newY.SWIECHDECH
If newY.SWIECHSWIX <> 0 Then xSet = xSet & ",SWIECHSWIX": xValues = xValues & " ," & newY.SWIECHSWIX
If newY.SWIECHSWIL <> 0 Then xSet = xSet & ",SWIECHSWIL": xValues = xValues & " ," & newY.SWIECHSWIL
If newY.SWIECHYAMJ <> 0 Then xSet = xSet & ",SWIECHYAMJ": xValues = xValues & " ," & newY.SWIECHYAMJ
If newY.SWIECHYHMS <> 0 Then xSet = xSet & ",SWIECHYHMS": xValues = xValues & " ," & newY.SWIECHYHMS
If newY.SWIECHYVER <> 0 Then xSet = xSet & ",SWIECHYVER": xValues = xValues & " ," & newY.SWIECHYVER
If newY.SWIECHWMTD <> 0 Then xSet = xSet & ",SWIECHWMTD": xValues = xValues & " ," & cur_P(newY.SWIECHWMTD)

If Trim(newY.SWIECHSER) <> "" Then xSet = xSet & ",SWIECHSER": xValues = xValues & " ,'" & newY.SWIECHSER & "'"
If Trim(newY.SWIECHSSE) <> "" Then xSet = xSet & ",SWIECHSSE": xValues = xValues & " ,'" & newY.SWIECHSSE & "'"
If Trim(newY.SWIECHOPEC) <> "" Then xSet = xSet & ",SWIECHOPEC": xValues = xValues & " ,'" & newY.SWIECHOPEC & "'"
If Trim(newY.SWIECHWES) <> "" Then xSet = xSet & ",SWIECHWES": xValues = xValues & " ,'" & newY.SWIECHWES & "'"
If Trim(newY.SWIECHWMTK) <> "" Then xSet = xSet & ",SWIECHWMTK": xValues = xValues & " ,'" & newY.SWIECHWMTK & "'"
If Trim(newY.SWIECHWBIC) <> "" Then xSet = xSet & ",SWIECHWBIC": xValues = xValues & " ,'" & newY.SWIECHWBIC & "'"
If Trim(newY.SWIECHWDEV) <> "" Then xSet = xSet & ",SWIECHWDEV": xValues = xValues & " ,'" & newY.SWIECHWDEV & "'"
If Trim(newY.SWIECHSTA) <> "" Then xSet = xSet & ",SWIECHSTA": xValues = xValues & " ,'" & newY.SWIECHSTA & "'"
If Trim(newY.SWIECHSTAK) <> "" Then xSet = xSet & ",SWIECHSTAK": xValues = xValues & " ,'" & newY.SWIECHSTAK & "'"

If Trim(newY.SWIECHWN20) <> "" Then xSet = xSet & ",SWIECHWN20": xValues = xValues & " ,'" & newY.SWIECHWN20 & "'"
If Trim(newY.SWIECHWL20) <> "" Then xSet = xSet & ",SWIECHWL20": xValues = xValues & " ,'" & newY.SWIECHWL20 & "'"
If Trim(newY.SWIECHW22C) <> "" Then xSet = xSet & ",SWIECHW22C": xValues = xValues & " ,'" & newY.SWIECHW22C & "'"
If Trim(newY.SWIECHW52A) <> "" Then xSet = xSet & ",SWIECHW52A": xValues = xValues & " ,'" & newY.SWIECHW52A & "'"
If Trim(newY.SWIECHW57A) <> "" Then xSet = xSet & ",SWIECHW57A": xValues = xValues & " ,'" & newY.SWIECHW57A & "'"
If Trim(newY.SWIECHW30V) <> "" Then xSet = xSet & ",SWIECHW30V": xValues = xValues & " ,'" & newY.SWIECHW30V & "'"
If Trim(newY.SWIECHSENS) <> "" Then xSet = xSet & ",SWIECHSENS": xValues = xValues & " ,'" & newY.SWIECHSENS & "'"

If Trim(newY.SWIECHYUSR) <> "" Then xSet = xSet & ",SWIECHYUSR": xValues = xValues & " ,'" & newY.SWIECHYUSR & "'"
      
       
      

xSql = "Insert into " & paramIBM_Library_SABSPE_XXX & ".YSWIECH0" & xSet & ")" & xValues & ")"
'Call FEU_ROUGE
Set rsADO = cnSab_Update.Execute(xSql, Nb)
'Call FEU_VERT
' Tester si la mise à jour a été effectuée
'===================================================================================

If Nb = 0 Then
    sqlYSWIECH0_Insert = "Erreur màj : " & newY.SWIECHOPEC & " - " & newY.SWIECHOPEN
    Exit Function
End If
 
Exit Function
Error_Handler:
    sqlYSWIECH0_Insert = Error
End Function

Public Function rsYSWIECH0_GetBuffer(rsADO As ADODB.Recordset, lYSWIECH0 As typeYSWIECH0)
On Error GoTo Error_Handler
rsYSWIECH0_GetBuffer = Null


lYSWIECH0.SWIECHSWID = rsADO("SWIECHSWID")
lYSWIECH0.SWIECHSEQ0 = rsADO("SWIECHSEQ0")

lYSWIECH0.SWIECHSER = rsADO("SWIECHSER")
lYSWIECH0.SWIECHSSE = rsADO("SWIECHSSE")
lYSWIECH0.SWIECHOPEC = rsADO("SWIECHOPEC")
lYSWIECH0.SWIECHOPEN = rsADO("SWIECHOPEN")
lYSWIECH0.SWIECHWMTK = rsADO("SWIECHWMTK")
lYSWIECH0.SWIECHWES = rsADO("SWIECHWES")
lYSWIECH0.SWIECHWBIC = rsADO("SWIECHWBIC")
lYSWIECH0.SWIECHDECH = rsADO("SWIECHDECH")
lYSWIECH0.SWIECHSTA = rsADO("SWIECHSTA")
lYSWIECH0.SWIECHSTAK = rsADO("SWIECHSTAK")
lYSWIECH0.SWIECHSWIX = rsADO("SWIECHSWIX")
lYSWIECH0.SWIECHSWIL = rsADO("SWIECHSWIL")

lYSWIECH0.SWIECHWDEV = rsADO("SWIECHWDEV")
lYSWIECH0.SWIECHWMTD = rsADO("SWIECHWMTD")
lYSWIECH0.SWIECHWN20 = rsADO("SWIECHWN20")
lYSWIECH0.SWIECHWL20 = rsADO("SWIECHWL20")
lYSWIECH0.SWIECHW22C = rsADO("SWIECHW22C")
lYSWIECH0.SWIECHW52A = rsADO("SWIECHW52A")
lYSWIECH0.SWIECHW57A = rsADO("SWIECHW57A")
lYSWIECH0.SWIECHW30V = rsADO("SWIECHW30V")
lYSWIECH0.SWIECHSENS = rsADO("SWIECHSENS")

lYSWIECH0.SWIECHYAMJ = rsADO("SWIECHYAMJ")
lYSWIECH0.SWIECHYHMS = rsADO("SWIECHYHMS")
lYSWIECH0.SWIECHYVER = rsADO("SWIECHYVER")
lYSWIECH0.SWIECHYUSR = rsADO("SWIECHYUSR")

Exit Function
Error_Handler:
rsYSWIECH0_GetBuffer = Error


End Function
Public Function rsYSWIECH0_Init(lYSWIECH0 As typeYSWIECH0)


lYSWIECH0.SWIECHSWID = 0
lYSWIECH0.SWIECHSEQ0 = 0

lYSWIECH0.SWIECHSER = ""
lYSWIECH0.SWIECHSSE = ""
lYSWIECH0.SWIECHOPEC = ""
lYSWIECH0.SWIECHOPEN = 0
lYSWIECH0.SWIECHWMTK = ""
lYSWIECH0.SWIECHWES = ""
lYSWIECH0.SWIECHWBIC = ""
lYSWIECH0.SWIECHDECH = 0

lYSWIECH0.SWIECHSTA = ""
lYSWIECH0.SWIECHSTAK = ""
lYSWIECH0.SWIECHSWIX = 0
lYSWIECH0.SWIECHSWIL = 0

lYSWIECH0.SWIECHWDEV = ""
lYSWIECH0.SWIECHWMTD = 0
lYSWIECH0.SWIECHWN20 = ""
lYSWIECH0.SWIECHWL20 = ""
lYSWIECH0.SWIECHW22C = ""
lYSWIECH0.SWIECHW52A = ""
lYSWIECH0.SWIECHW57A = ""
lYSWIECH0.SWIECHW30V = ""
lYSWIECH0.SWIECHSENS = ""

lYSWIECH0.SWIECHYAMJ = 0
lYSWIECH0.SWIECHYHMS = 0
lYSWIECH0.SWIECHYVER = 0
lYSWIECH0.SWIECHYUSR = ""

End Function



















