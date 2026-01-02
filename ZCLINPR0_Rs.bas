Attribute VB_Name = "rsZCLINPR0"
'---------------------------------------------------------
Option Explicit
Type typeZCLINPR0

    CLINPRETA       As Integer                        ' CODE ETABLISSEMENT
    CLINPRCLI       As String * 7                     ' NUMERO CLIENT
    CLINPRTYP       As String * 1                     ' professionnel
    CLINPRNUM       As String * 9                     ' du client

End Type

'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsZCLINPR0_GetBuffer(rsADO As ADODB.Recordset, rsZCLINPR0 As typeZCLINPR0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsZCLINPR0_GetBuffer = Null

rsZCLINPR0.CLINPRETA = rsADO("CLINPRETA")
rsZCLINPR0.CLINPRCLI = rsADO("CLINPRCLI")
rsZCLINPR0.CLINPRTYP = rsADO("CLINPRTYP")
rsZCLINPR0.CLINPRNUM = rsADO("CLINPRNUM")
Exit Function

Error_Handler:

rsZCLINPR0_GetBuffer = Error

End Function

Public Sub rsZCLINPR0_Init(rsZCLINPR0 As typeZCLINPR0)
rsZCLINPR0.CLINPRETA = 0
rsZCLINPR0.CLINPRCLI = ""
rsZCLINPR0.CLINPRTYP = ""
rsZCLINPR0.CLINPRNUM = ""
End Sub



