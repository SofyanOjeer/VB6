Attribute VB_Name = "rsZRELEVE0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Type typeZRELEVE0
    RELEVEETA       As Integer                        ' ETABLISSEMENT
    RELEVEPLA       As Long                           ' NUMERO PLAN
    RELEVECOM       As String * 20                    ' NUMERO COMPTE
    RELEVEREL       As String * 1                     ' TABLES BASE 019
    RELEVETYP       As String * 1                     ' 1 client , 2 compte
    RELEVENUM       As String * 20                    ' N° Client ou Compte
    RELEVEADR       As String * 2                     ' CODE ADRESSE
    RELEVEGES       As String * 1                     ' RELEVE GESTIONNAIRE
    RELEVEDER       As Long                           ' DATE DERNIER RELEVE
    RELEVEEXT       As Long                           ' NUMERO D'EXTRAIT

End Type
'---------------------------------------------------------
Public Function rsZRELEVE0_Read(lZRELEVE0 As typeZRELEVE0)
'---------------------------------------------------------
Dim xSQL As String, X As String, V
On Error GoTo Error_Handler

rsZRELEVE0_Read = Null
lZRELEVE0.RELEVENUM = lZRELEVE0.RELEVECOM
lZRELEVE0.RELEVETYP = "2"
lZRELEVE0.RELEVEADR = ""

'Lecture modalités extrait => adresse d'envoi
'=================================
Set rsSab = Nothing
xSQL = "select * from " & paramIBM_Library_SAB & ".ZRELEVE0" _
     & " where RELEVECOM = '" & lZRELEVE0.RELEVECOM & "'" _
     & " and RELEVEREL = '" & lZRELEVE0.RELEVEREL & "'" _
     & " and RELEVEETA = " & currentZMNURUT0.MNURUTETB

Set rsSab = cnsab.Execute(xSQL)

If Not rsSab.EOF Then
    V = rsZRELEVE0_GetBuffer(rsSab, lZRELEVE0)
Else

'<D,M> non trouvé : rechercher <*>
'<*>   non trouvé : rechercher <M>
'=================================
    If lZRELEVE0.RELEVEREL = "*" Then
        X = "M"
    Else
        X = "*"
    End If
    
    xSQL = "select * from " & paramIBM_Library_SAB & ".ZRELEVE0" _
         & " where RELEVECOM = '" & lZRELEVE0.RELEVECOM & "'" _
         & " and RELEVEREL = '" & X & "'" _
         & " and RELEVEETA = " & currentZMNURUT0.MNURUTETB
    
    Set rsSab = cnsab.Execute(xSQL)
    
    If Not rsSab.EOF Then V = rsZRELEVE0_GetBuffer(rsSab, lZRELEVE0)
End If

Exit Function

Error_Handler:
'-------------
    rsZRELEVE0_Read = " rsZRELEVE0_Read : " & Error
End Function

Public Sub rsZRELEVE0_Init(rsZRELEVE0 As typeZRELEVE0)
rsZRELEVE0.RELEVEETA = 0
rsZRELEVE0.RELEVEPLA = 0
rsZRELEVE0.RELEVECOM = ""
rsZRELEVE0.RELEVEREL = ""
rsZRELEVE0.RELEVETYP = ""
rsZRELEVE0.RELEVENUM = ""
rsZRELEVE0.RELEVEADR = ""
rsZRELEVE0.RELEVEGES = ""
rsZRELEVE0.RELEVEDER = 0
rsZRELEVE0.RELEVEEXT = 0
End Sub

'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsZRELEVE0_GetBuffer(rsAdo As ADODB.Recordset, rsZRELEVE0 As typeZRELEVE0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsZRELEVE0_GetBuffer = Null

rsZRELEVE0.RELEVEETA = rsAdo("RELEVEETA")
rsZRELEVE0.RELEVEPLA = rsAdo("RELEVEPLA")
rsZRELEVE0.RELEVECOM = rsAdo("RELEVECOM")
rsZRELEVE0.RELEVEREL = rsAdo("RELEVEREL")
rsZRELEVE0.RELEVETYP = rsAdo("RELEVETYP")
rsZRELEVE0.RELEVENUM = rsAdo("RELEVENUM")
rsZRELEVE0.RELEVEADR = rsAdo("RELEVEADR")
rsZRELEVE0.RELEVEGES = rsAdo("RELEVEGES")
rsZRELEVE0.RELEVEDER = rsAdo("RELEVEDER")
rsZRELEVE0.RELEVEEXT = rsAdo("RELEVEEXT")

Exit Function

Error_Handler:

rsZRELEVE0_GetBuffer = Error

End Function



'








