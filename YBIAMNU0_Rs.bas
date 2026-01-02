Attribute VB_Name = "rsYBIAMNU0"
'---------------------------------------------------------
Option Explicit
Type typeYBIAMNU0
    Method                  As String * 12   'pour compatibilité version précédente cf SAB_MNU.frm
    
    Src                     As String * 3
    ID                      As String * 20
    Memo                    As Variant

End Type

Type typeYBIARUT0

    MNURUTUTI               As String * 10
    MNURUTNOM               As String * 30
    MNURUTETB               As String * 5 '3
    MNURUTCUT               As String * 5 '7
    MNURUTLOG               As String * 1

    MNUUTIETB               As String * 5 '3
    MNUUTICUT               As String * 5 '7
    MNUUTICGR               As String * 5 '7
    MNUUTIDRG               As String * 1
    MNUUTIOUT               As String * 10
    MNUUTILAN               As String * 1
    MNUUTIMSE               As String * 1
    MNUUTIAGE               As String * 5 '3
    MNUUTISER               As String * 2
    MNUUTISRV               As String * 2

End Type

'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsYBIAMNU0_GetBuffer(rsADO As ADODB.Recordset, rsYBIAMNU0 As typeYBIAMNU0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsYBIAMNU0_GetBuffer = Null

rsYBIAMNU0.Method = ""

rsYBIAMNU0.Src = rsADO("Src")
rsYBIAMNU0.ID = rsADO("ID")
rsYBIAMNU0.Memo = rsADO("Memo")

Exit Function

Error_Handler:

rsYBIAMNU0_GetBuffer = Error

End Function


'






