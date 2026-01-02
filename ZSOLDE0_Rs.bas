Attribute VB_Name = "rsZSOLDE0"
'---------------------------------------------------------
Option Explicit
Type typeZSOLDE0
    
    SOLDEETA        As Integer                        ' ETABLISSEMENT
    SOLDEPLA        As Long                           ' NUMERO PLAN
    SOLDECOM        As String * 20                    ' NUMERO COMPTE
    SOLDEDMO        As Long                           ' DATE DERNIER MVT
    SOLDEDAN        As Long                           ' DATE ANTERIEUR
    SOLDECEN        As Currency                         ' SOLDE ENCOURS
    SOLDECAN        As Currency                         ' SOLDE ANTERIEUR
    SOLDEC01        As Currency                         ' SOLDE M
    SOLDEC02        As Currency                         ' SOLDE M -1
    SOLDEC03        As Currency                         ' SOLDE M -2
    SOLDEC04        As Currency                         ' SOLDE M -3
    SOLDEC05        As Currency                         ' SOLDE M -4
    SOLDEC06        As Currency                         ' SOLDE M -5
    SOLDEC07        As Currency                         ' SOLDE M -6
    SOLDEC08        As Currency                         ' SOLDE M -7
    SOLDEC09        As Currency                         ' SOLDE M -8
    SOLDEC10        As Currency                         ' SOLDE M -9
    SOLDEC11        As Currency                         ' SOLDE M -10
    SOLDEC12        As Currency                         ' SOLDE M -11
    SOLDEVEN        As Currency                         ' SOLDE VAL. ENCOURS
    SOLDEVAN        As Currency                         ' SOLDE VAL. ANTERIEUR
    SOLDEV01        As Currency                         ' SOLDE VAL. M
    SOLDEV02        As Currency                         ' SOLDE VAL. M -1
    SOLDEV03        As Currency                         ' SOLDE VAL. M -2
    SOLDEV04        As Currency                         ' SOLDE VAL. M -3
    SOLDEV05        As Currency                         ' SOLDE VAL. M -4
    SOLDEV06        As Currency                         ' SOLDE VAL. M -5
    SOLDEV07        As Currency                         ' SOLDE VAL. M -6
    SOLDEV08        As Currency                         ' SOLDE VAL. M -7
    SOLDEV09        As Currency                         ' SOLDE VAL. M -8
    SOLDEV10        As Currency                         ' SOLDE VAL. M -9
    SOLDEV11        As Currency                         ' SOLDE VAL. M -10
    SOLDEV12        As Currency                         ' SOLDE VAL. M -11
    
End Type

'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsZSOLDE0_GetBuffer(rsADO As ADODB.Recordset, rsZSOLDE0 As typeZSOLDE0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsZSOLDE0_GetBuffer = Null

rsZSOLDE0.SOLDEETA = rsADO("SOLDEETA")
rsZSOLDE0.SOLDEPLA = rsADO("SOLDEPLA")
rsZSOLDE0.SOLDECOM = rsADO("SOLDECOM")
rsZSOLDE0.SOLDEDMO = rsADO("SOLDEDMO")
rsZSOLDE0.SOLDEDAN = rsADO("SOLDEDAN")
rsZSOLDE0.SOLDECEN = rsADO("SOLDECEN")
rsZSOLDE0.SOLDECAN = rsADO("SOLDECAN")
rsZSOLDE0.SOLDEC01 = rsADO("SOLDEC01")
rsZSOLDE0.SOLDEC02 = rsADO("SOLDEC02")
rsZSOLDE0.SOLDEC03 = rsADO("SOLDEC03")
rsZSOLDE0.SOLDEC04 = rsADO("SOLDEC04")
rsZSOLDE0.SOLDEC05 = rsADO("SOLDEC05")
rsZSOLDE0.SOLDEC06 = rsADO("SOLDEC06")
rsZSOLDE0.SOLDEC07 = rsADO("SOLDEC07")
rsZSOLDE0.SOLDEC08 = rsADO("SOLDEC08")
rsZSOLDE0.SOLDEC09 = rsADO("SOLDEC09")
rsZSOLDE0.SOLDEC10 = rsADO("SOLDEC10")
rsZSOLDE0.SOLDEC11 = rsADO("SOLDEC11")
rsZSOLDE0.SOLDEC12 = rsADO("SOLDEC12")
rsZSOLDE0.SOLDEVEN = rsADO("SOLDEVEN")
rsZSOLDE0.SOLDEVAN = rsADO("SOLDEVAN")
rsZSOLDE0.SOLDEV01 = rsADO("SOLDEV01")
rsZSOLDE0.SOLDEV02 = rsADO("SOLDEV02")
rsZSOLDE0.SOLDEV03 = rsADO("SOLDEV03")
rsZSOLDE0.SOLDEV04 = rsADO("SOLDEV04")
rsZSOLDE0.SOLDEV05 = rsADO("SOLDEV05")
rsZSOLDE0.SOLDEV06 = rsADO("SOLDEV06")
rsZSOLDE0.SOLDEV07 = rsADO("SOLDEV07")
rsZSOLDE0.SOLDEV08 = rsADO("SOLDEV08")
rsZSOLDE0.SOLDEV09 = rsADO("SOLDEV09")
rsZSOLDE0.SOLDEV10 = rsADO("SOLDEV10")
rsZSOLDE0.SOLDEV11 = rsADO("SOLDEV11")
rsZSOLDE0.SOLDEV12 = rsADO("SOLDEV12")
Exit Function

Error_Handler:

rsZSOLDE0_GetBuffer = Error

End Function

'---------------------------------------------------------
Public Sub rsZSOLDE0_Init(rsZSOLDE0 As typeZSOLDE0)
'---------------------------------------------------------

End Sub









Public Function rsZSOLDE0_Read(lSOLDECOM As String, lK As Integer, lcurX As Currency)
Dim xSQL As String
rsZSOLDE0_Read = Null
xSQL = "select * from " & paramIBM_Library_SAB & ".ZSOLDE0  where SOLDECOM = '" & lSOLDECOM & "'"
Set rsSab = cnsab.Execute(xSQL)

If Not rsSab.EOF Then
    Call rsZSOLDE0_SOLDEC(rsSab, lK, lcurX)
Else
    lcurX = 0
    rsZSOLDE0_Read = "? SOLDECOM = " & lSOLDECOM
End If
End Function

Public Sub rsZSOLDE0_SOLDEC(rsADO As ADODB.Recordset, lK As Integer, lcurX As Currency)
Select Case lK
        Case 0: lcurX = rsADO("SOLDECEN")
        Case 1: lcurX = rsADO("SOLDEC01")
        Case 2: lcurX = rsADO("SOLDEC02")
        Case 3: lcurX = rsADO("SOLDEC03")
        Case 4: lcurX = rsADO("SOLDEC04")
        Case 5: lcurX = rsADO("SOLDEC05")
        Case 6: lcurX = rsADO("SOLDEC06")
        Case 7: lcurX = rsADO("SOLDEC07")
        Case 8: lcurX = rsADO("SOLDEC08")
        Case 9: lcurX = rsADO("SOLDEC09")
        Case 10: lcurX = rsADO("SOLDEC10")
        Case 11: lcurX = rsADO("SOLDEC11")
        Case 12: lcurX = rsADO("SOLDEC12")
End Select

End Sub
