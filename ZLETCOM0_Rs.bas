Attribute VB_Name = "rsZLETCOM0"

'---------------------------------------------------------
Option Explicit
Type typeZLETCOM0
    LETCOMETA       As Integer                        ' ETABLISSEMENT
    LETCOMPLA       As Long                           ' NUMERO DU PLAN
    LETCOMCOM       As String * 20                    ' NUMERO DE COMPTE
    LETCOMAGR       As Integer                        ' AGC   RESPONSABLE
    LETCOMSER       As String * 2                     ' SRV   RESPONSABLE
    LETCOMSSR       As String * 2                     ' SSRV  RESPONSABLE
    LETCOMDDE       As Long                           ' DT DERNIER EXTACT
    LETCOMDDR       As Long                           ' DT DERNIER RAPPRO
    LETCOMDPR       As Long                           ' DT PROCHAIN RAPP.
    LETCOMPER       As String * 1                     ' PERIOD. TRT RAPP.
    LETCOMNBP       As Long                           ' NBR. DE PERIODES
    LETCOMDTR       As Long                           ' DERNIERE DATE TRT
    LETCOMPIE       As Long                           ' DERNIER NUM PIECE
    LETCOMECR       As Long                           ' DERNIER NUM ECR.
    LETCOMOUV       As Long                           ' DATE OUVERTURE
    LETCOMCLO       As Long                           ' DATE CLOTURE
    LETCOMDMC       As Long                           ' DATE MOD CRITERES
    LETCOMMON       As String * 1                     ' MONTANT
    LETCOMDVA       As String * 1                     ' DATE VALEUR
    LETCOMDOP       As String * 1                     ' DATE OPERATION
    LETCOMOPE       As String * 1                     ' CODE OPERATION
    LETCOMNU1       As Long                           ' NUM DE LIBELLE(1)
    LETCOMPO1       As Long                           ' POSITION(1)
    LETCOMLO1       As Long                           ' LONG DE CHAINE(1)
    LETCOMNU2       As Long                           ' NUM DE LIBELLE(2)
    LETCOMPO2       As Long                           ' POSITION(2)
    LETCOMLO2       As Long                           ' LONG DE CHAINE(2)
    LETCOMAGO       As String * 1                     ' AGENCE OPERATRICE
    LETCOMSEO       As String * 1                     ' SERVICE OPERATEUR
    LETCOMSSO       As String * 1                     ' SSERV.  OPERATEUR
    LETCOMCHE       As String * 1                     ' NUMERO DE CHEQUE
    LETCOMANA       As String * 1                     ' CODE ANALYTIQUE

End Type
'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsZLETCOM0_GetBuffer(rsado As ADODB.Recordset, rsZLETCOM0 As typeZLETCOM0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsZLETCOM0_GetBuffer = Null

rsZLETCOM0.LETCOMETA = rsado("LETCOMETA")
rsZLETCOM0.LETCOMPLA = rsado("LETCOMPLA")
rsZLETCOM0.LETCOMCOM = rsado("LETCOMCOM")
rsZLETCOM0.LETCOMAGR = rsado("LETCOMAGR")
rsZLETCOM0.LETCOMSER = rsado("LETCOMSER")
rsZLETCOM0.LETCOMSSR = rsado("LETCOMSSR")
rsZLETCOM0.LETCOMDDE = rsado("LETCOMDDE")
rsZLETCOM0.LETCOMDDR = rsado("LETCOMDDR")
rsZLETCOM0.LETCOMDPR = rsado("LETCOMDPR")
rsZLETCOM0.LETCOMPER = rsado("LETCOMPER")
rsZLETCOM0.LETCOMNBP = rsado("LETCOMNBP")
rsZLETCOM0.LETCOMDTR = rsado("LETCOMDTR")
rsZLETCOM0.LETCOMPIE = rsado("LETCOMPIE")
rsZLETCOM0.LETCOMECR = rsado("LETCOMECR")
rsZLETCOM0.LETCOMOUV = rsado("LETCOMOUV")
rsZLETCOM0.LETCOMCLO = rsado("LETCOMCLO")
rsZLETCOM0.LETCOMDMC = rsado("LETCOMDMC")
rsZLETCOM0.LETCOMMON = rsado("LETCOMMON")
rsZLETCOM0.LETCOMDVA = rsado("LETCOMDVA")
rsZLETCOM0.LETCOMDOP = rsado("LETCOMDOP")
rsZLETCOM0.LETCOMOPE = rsado("LETCOMOPE")
rsZLETCOM0.LETCOMNU1 = rsado("LETCOMNU1")
rsZLETCOM0.LETCOMPO1 = rsado("LETCOMPO1")
rsZLETCOM0.LETCOMLO1 = rsado("LETCOMLO1")
rsZLETCOM0.LETCOMNU2 = rsado("LETCOMNU2")
rsZLETCOM0.LETCOMPO2 = rsado("LETCOMPO2")
rsZLETCOM0.LETCOMLO2 = rsado("LETCOMLO2")
rsZLETCOM0.LETCOMAGO = rsado("LETCOMAGO")
rsZLETCOM0.LETCOMSEO = rsado("LETCOMSEO")
rsZLETCOM0.LETCOMSSO = rsado("LETCOMSSO")
rsZLETCOM0.LETCOMCHE = rsado("LETCOMCHE")
rsZLETCOM0.LETCOMANA = rsado("LETCOMANA")

Exit Function

Error_Handler:

rsZLETCOM0_GetBuffer = Error

End Function


'








