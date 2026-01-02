Attribute VB_Name = "rsZMNUHLB0"
Option Explicit

Type typeZMNUHLB0
    
    MNUHLBETB       As Integer                        ' ETABLISSEMENT
    MNUHLBREF       As Long                           ' REFERENCE LOT
    MNUHLBCLA       As String * 1                     ' CLASSE
    MNUHLBNOM       As String * 10                    ' NOM GROUPE
    MNUHLBVAL       As String * 1                     ' VALIDE
    MNUHLBDBD       As Long                           ' DATE DEBUT
    MNUHLBDBH       As Long                           ' HEURE DEBUT
    MNUHLBFID       As Long                           ' DATE FIN
    MNUHLBFIH       As Long                           ' HEURE FIN
    MNUHLBSUS       As Integer                        ' USER SAISIE
    MNUHLBSDT       As Long                           ' DATE SAISIE
    MNUHLBSHE       As Long                           ' HEURE SAISIE
    
End Type

'---------------------------------------------------------
Public Function rsZMNUHLB0_GetBuffer(rsAdo As ADODB.Recordset, rsZMNUHLB0 As typeZMNUHLB0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsZMNUHLB0_GetBuffer = Null

rsZMNUHLB0.MNUHLBETB = rsAdo("MNUHLBETB")
rsZMNUHLB0.MNUHLBREF = rsAdo("MNUHLBREF")
rsZMNUHLB0.MNUHLBCLA = rsAdo("MNUHLBCLA")
rsZMNUHLB0.MNUHLBNOM = rsAdo("MNUHLBNOM")
rsZMNUHLB0.MNUHLBVAL = rsAdo("MNUHLBVAL")
rsZMNUHLB0.MNUHLBDBD = rsAdo("MNUHLBDBD")
rsZMNUHLB0.MNUHLBDBH = rsAdo("MNUHLBDBH")
rsZMNUHLB0.MNUHLBFID = rsAdo("MNUHLBFID")
rsZMNUHLB0.MNUHLBFIH = rsAdo("MNUHLBFIH")
rsZMNUHLB0.MNUHLBSUS = rsAdo("MNUHLBSUS")
rsZMNUHLB0.MNUHLBSDT = rsAdo("MNUHLBSDT")
rsZMNUHLB0.MNUHLBSHE = rsAdo("MNUHLBSHE")

Exit Function

Error_Handler:

rsZMNUHLB0_GetBuffer = Error

End Function
