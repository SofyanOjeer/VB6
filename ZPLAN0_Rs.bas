Attribute VB_Name = "rsZPLAN0"
'---------------------------------------------------------
Option Explicit
Type typeZPLAN0
    
    PLANETABL       As Integer                        ' ETABLISSEMENT
    PLANPLAN        As Long                           ' NUMERO PLAN
    PLANCOOBL       As String * 10                    ' COMPTE OBLIGATOIRE
    PLANINTIT       As String * 32                    ' INTITULE
    PLANCOPRO       As String * 3                     ' TABLES BASE 014
    PLANCLASS       As Long                           ' CLASSE SECURITE
    PLANFONCT       As String * 1                     ' TABLES BASE 015
    PLANSESOL       As String * 1                     ' CODE SENS SOLDE D/C
    PLANGEDEP       As String * 1                     ' O/N
    PLANTIERS       As String * 1                     ' COMPTE TIERS O/N
    PLANFICOB       As String * 1                     ' O/N
    PLANCARAC       As Long                           ' 3 à 20
    PLANPESTO       As String * 1                     ' Mois, Trimestre, Année
    PLANNBPER       As Long                           ' 1 à 24
    PLANNBMOU       As Long                           ' NB MVT A CONSERVER
    PLANINEXT       As String * 32                    ' INTITUL EXTRAIT CPT
    PLANPROGR       As String * 8                     ' PROGRAMME DE CONTROL
    
End Type

'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsZPLAN0_GetBuffer(rsAdo As ADODB.Recordset, rsZPLAN0 As typeZPLAN0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsZPLAN0_GetBuffer = Null

'    rsZPLAN0.COMPTEETA = rsADO("COMPTEETA")
rsZPLAN0.PLANETABL = rsAdo("PLANETABL")
rsZPLAN0.PLANPLAN = rsAdo("PLANPLAN")
rsZPLAN0.PLANCOOBL = rsAdo("PLANCOOBL")
rsZPLAN0.PLANINTIT = rsAdo("PLANINTIT")
rsZPLAN0.PLANCOPRO = rsAdo("PLANCOPRO")
rsZPLAN0.PLANCLASS = rsAdo("PLANCLASS")
rsZPLAN0.PLANFONCT = rsAdo("PLANFONCT")
rsZPLAN0.PLANSESOL = rsAdo("PLANSESOL")
rsZPLAN0.PLANGEDEP = rsAdo("PLANGEDEP")
rsZPLAN0.PLANTIERS = rsAdo("PLANTIERS")
rsZPLAN0.PLANFICOB = rsAdo("PLANFICOB")
rsZPLAN0.PLANCARAC = rsAdo("PLANCARAC")
rsZPLAN0.PLANPESTO = rsAdo("PLANPESTO")
rsZPLAN0.PLANNBPER = rsAdo("PLANNBPER")
rsZPLAN0.PLANNBMOU = rsAdo("PLANNBMOU")
rsZPLAN0.PLANINEXT = rsAdo("PLANINEXT")
rsZPLAN0.PLANPROGR = rsAdo("PLANPROGR")

Exit Function

Error_Handler:

rsZPLAN0_GetBuffer = Error

End Function

'---------------------------------------------------------
Public Sub rsZPLAN0_Init(rsZPLAN0 As typeZPLAN0)
'---------------------------------------------------------

End Sub


Public Function rsZPLAN0_Read(lCOMPTEOBL As String, rsZPLAN0 As typeZPLAN0)
Dim xSQL As String
On Error GoTo Error_Handler

xSQL = "select * from " & paramIBM_Library_SAB & ".ZPLAN0  where PLANCOOBL = '" & lCOMPTEOBL & "'"
Set rsSab = cnsab.Execute(xSQL)

If Not rsSab.EOF Then
    rsZPLAN0_Read = rsZPLAN0_GetBuffer(rsSab, rsZPLAN0)
Else
    rsZPLAN0_Read = xSQL
End If
Exit Function

Error_Handler:

rsZPLAN0_Read = Error
End Function

Public Sub rsZPLAN0_cboPLANCOPRO(cboX As ComboBox)
Dim xSQL As String

cboX.Clear
xSQL = "select distinct PLANCOPRO from " & paramIBM_Library_SAB & ".ZPLAN0 order by PLANCOPRO"
     
Set rsSab = cnsab.Execute(xSQL)
Do While Not rsSab.EOF
    cboX.AddItem rsSab("PLANCOPRO")
    rsSab.MoveNext
Loop


End Sub


Public Sub rsZPLAN0_cboPLANCOOBL(cboX As ComboBox)
Dim xSQL As String

cboX.Clear
xSQL = "select PLANCOOBL from " & paramIBM_Library_SAB & ".ZPLAN0 order by PLANCOOBL"
     
Set rsSab = cnsab.Execute(xSQL)
Do While Not rsSab.EOF
    cboX.AddItem rsSab("PLANCOOBL")
    rsSab.MoveNext
Loop


End Sub


Public Sub rsZPLAN0_cboK2(cboX As Control)
Dim xSQL As String

cboX.Clear
xSQL = "select  PLANCOOBL,PLANCOPRO,PLANINTIT from " & paramIBM_Library_SAB & ".ZPLAN0 order by PLANCOOBL"
     
Set rsSab = cnsab.Execute(xSQL)
Do While Not rsSab.EOF
    cboX.AddItem rsSab("PLANCOOBL") & " : " & rsSab("PLANCOPRO") & "  " & rsSab("PLANINTIT")
    rsSab.MoveNext
Loop


End Sub

