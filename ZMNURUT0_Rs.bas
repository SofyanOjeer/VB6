Attribute VB_Name = "rsZMNURUT0"
'---------------------------------------------------------
Option Explicit
Type typeZMNURUT0
    MNURUTUTI       As String * 10                    ' UTILISATEUR
    MNURUTNOM       As String * 30                    ' NOM
    MNURUTETB       As Integer                        ' ETAB. PAR DEFAUT
    MNURUTCUT       As Integer                        ' CODE INTERNE
    MNURUTLOG       As String * 1                     ' ENTREE LOGICIEL
    
    
End Type

Public Function MNURUTCUT_Get(lMNURUTUTI As String) As Integer
Dim xSQL As String
Dim Nb As Long
On Error Resume Next


xSQL = "select MNURUTCUT from " & paramIBM_Library_SAB & ".ZMNURUT0" _
     & " where MNURUTETB = " & currentZMNURUT0.MNURUTETB _
     & " where MNURUTUTI = " & lMNURUTUTI
Set rsSab = cnsab.Execute(xSQL)
If Not rsSab.EOF Then
    MNURUTCUT_Get = rsSab("MNURUTCUT")
Else
    MNURUTCUT_Get = 0
End If

End Function

Public Sub arrMNURUUTI_Load(larrMNURUTUTI() As String)
Dim xSQL As String
Dim Nb As Long
On Error Resume Next

xSQL = "select count(*) as Tally  from " & paramIBM_Library_SAB & ".ZMNURUT0"
Set rsSab = cnsab.Execute(xSQL)
Nb = rsSab("Tally")
ReDim larrMNURUTUTI(Nb + 10)

xSQL = "select MNURUTCUT, MNURUTUTI from " & paramIBM_Library_SAB & ".ZMNURUT0" _
     & " where MNURUTETB = " & currentZMNURUT0.MNURUTETB _
     & " order by MNURUTCUT"
     
Set rsSab = cnsab.Execute(xSQL)
Do While Not rsSab.EOF
    Nb = rsSab("MNURUTCUT")
    larrMNURUTUTI(Nb) = rsSab("MNURUTUTI")
    rsSab.MoveNext
Loop

End Sub

'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsZMNURUT0_GetBuffer(rsAdo As ADODB.Recordset, rsZMNURUT0 As typeZMNURUT0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsZMNURUT0_GetBuffer = Null

rsZMNURUT0.MNURUTUTI = rsAdo("MNURUTUTI")
rsZMNURUT0.MNURUTNOM = rsAdo("MNURUTNOM")
rsZMNURUT0.MNURUTETB = rsAdo("MNURUTETB")
rsZMNURUT0.MNURUTCUT = rsAdo("MNURUTCUT")
rsZMNURUT0.MNURUTLOG = rsAdo("MNURUTLOG")

Exit Function

Error_Handler:

rsZMNURUT0_GetBuffer = Error

End Function


'






