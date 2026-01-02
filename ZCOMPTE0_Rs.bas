Attribute VB_Name = "rsZCOMPTE0"
'---------------------------------------------------------
Option Explicit
Type typeZCOMPTE0
    COMPTEETA       As Integer                        ' ETABLISSEMENT
    COMPTEPLA       As Long                           ' NUMERO PLAN
    COMPTECOM       As String * 20                    ' NUMERO COMPTE
    COMPTEOBL       As String * 10                    ' COMPTE OBLIGATOIRE
    COMPTEINT       As String * 32                    ' INTITULE
    COMPTEAGE       As Integer                        ' AGENCE
    COMPTEDEV       As String * 3                     ' TABLES BASE 013
    COMPTEOUV       As Long                           ' DATE OUVERTURE
    COMPTECLO       As Long                           ' DATE CLOTURE
    COMPTELOR       As String * 1                     ' Lori/Nostri/AUTRE
    COMPTESUC       As String * 1                     ' O/N
    COMPTECLA       As Long                           ' CLASSE SECURITE
    COMPTEFON       As String * 1                     ' TABLES BASE 015
    COMPTEBLO       As Long                           ' DATE LIMITE BLOCAGE
    COMPTEMOT       As String * 32                    ' MOTIF BLOCAGE
    COMPTESEN       As String * 1                     ' CODE SENS SOLDE D/C
    COMPTEMOD       As Long                           ' DATE MODIFICATION
   
End Type

'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsZCOMPTE0_GetBuffer(rsAdo As ADODB.Recordset, rsZCOMPTE0 As typeZCOMPTE0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsZCOMPTE0_GetBuffer = Null

rsZCOMPTE0.COMPTEETA = rsAdo("COMPTEETA")
rsZCOMPTE0.COMPTEPLA = rsAdo("COMPTEPLA")
rsZCOMPTE0.COMPTECOM = rsAdo("COMPTECOM")
rsZCOMPTE0.COMPTEOBL = rsAdo("COMPTEOBL")
rsZCOMPTE0.COMPTEINT = rsAdo("COMPTEINT")
rsZCOMPTE0.COMPTEAGE = rsAdo("COMPTEAGE")
rsZCOMPTE0.COMPTEDEV = rsAdo("COMPTEDEV")
rsZCOMPTE0.COMPTEOUV = rsAdo("COMPTEOUV")
rsZCOMPTE0.COMPTECLO = rsAdo("COMPTECLO")
rsZCOMPTE0.COMPTELOR = rsAdo("COMPTELOR")
rsZCOMPTE0.COMPTESUC = rsAdo("COMPTESUC")
rsZCOMPTE0.COMPTECLA = rsAdo("COMPTECLA")
rsZCOMPTE0.COMPTEFON = rsAdo("COMPTEFON")
rsZCOMPTE0.COMPTEBLO = rsAdo("COMPTEBLO")
rsZCOMPTE0.COMPTEMOT = rsAdo("COMPTEMOT")
rsZCOMPTE0.COMPTESEN = rsAdo("COMPTESEN")
rsZCOMPTE0.COMPTEMOD = rsAdo("COMPTEMOD")
Exit Function

Error_Handler:

rsZCOMPTE0_GetBuffer = Error

End Function

'---------------------------------------------------------
Public Sub rsZCOMPTE0_Init(rsZCOMPTE0 As typeZCOMPTE0)
'---------------------------------------------------------

End Sub


'








