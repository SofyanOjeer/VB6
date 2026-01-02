Attribute VB_Name = "rsZDECMOU0"
Option Explicit
Type typeZDECMOU0
    DECMOUETA       As Integer                        ' Etablissment
    DECMOUCOM       As String * 20                    ' Compte
    DECMOUDTR       As Long                           ' Date traitement
    DECMOUAGE       As Integer                        ' Agence
    DECMOUSER       As String * 2                     ' Service
    DECMOUSSE       As String * 2                     ' Sous service
    DECMOUCOP       As String * 3                     ' Code opération
    DECMOUNOP       As Long                           ' Numéro opération
    DECMOUDRE       As Long                           ' Date de référence
    DECMOUDLR       As Long                           ' Date limite rejet
    DECMOUUIN       As Integer                        ' Utilis.initiateur
    DECMOUDCR       As Long                           ' Date création
    DECMOUDUT       As Long                           ' Date utilisation
    DECMOUUTI       As Integer                        ' Utilisateur
    DECMOUREA       As String * 1                     ' Rejet acceptation
    DECMOUNSQ       As Long                           ' N° Séquence
    DECMOUFUT       As String * 20                    ' ZONE FUTURE
    DECMOUORI       As String * 1                     ' surveillance
    DECMOUNAT       As String * 3                     ' nature opération
    DECMOUMRE       As String * 6                     ' code motif rejet
    DECMOUREQ       As Long                           ' code rejet equiv
    DECMOUAPS       As String * 3                     ' applica/mise en surv
    DECMOUMOS       As String * 6                     ' motif de surv.
    DECMOUFIL       As String * 100                   ' zone libre
                  
End Type
Public Function rsZDECMOU0_GetBuffer(rsADO As ADODB.Recordset, rsZDECMOU0 As typeZDECMOU0)
On Error GoTo Error_Handler
rsZDECMOU0_GetBuffer = Null
rsZDECMOU0.DECMOUETA = rsADO("DECMOUETA")
rsZDECMOU0.DECMOUCOM = rsADO("DECMOUCOM")
rsZDECMOU0.DECMOUDTR = rsADO("DECMOUDTR")
rsZDECMOU0.DECMOUAGE = rsADO("DECMOUAGE")
rsZDECMOU0.DECMOUSER = rsADO("DECMOUSER")
rsZDECMOU0.DECMOUSSE = rsADO("DECMOUSSE")
rsZDECMOU0.DECMOUCOP = rsADO("DECMOUCOP")
rsZDECMOU0.DECMOUNOP = rsADO("DECMOUNOP")
rsZDECMOU0.DECMOUDRE = rsADO("DECMOUDRE")
rsZDECMOU0.DECMOUDLR = rsADO("DECMOUDLR")
rsZDECMOU0.DECMOUUIN = rsADO("DECMOUUIN")
rsZDECMOU0.DECMOUDCR = rsADO("DECMOUDCR")
rsZDECMOU0.DECMOUDUT = rsADO("DECMOUDUT")
rsZDECMOU0.DECMOUUTI = rsADO("DECMOUUTI")
rsZDECMOU0.DECMOUREA = rsADO("DECMOUREA")
rsZDECMOU0.DECMOUNSQ = rsADO("DECMOUNSQ")
rsZDECMOU0.DECMOUFUT = rsADO("DECMOUFUT")

rsZDECMOU0.DECMOUORI = rsADO("DECMOUORI")
rsZDECMOU0.DECMOUNAT = rsADO("DECMOUNAT")
rsZDECMOU0.DECMOUMRE = rsADO("DECMOUMRE")
rsZDECMOU0.DECMOUREQ = rsADO("DECMOUREQ")
rsZDECMOU0.DECMOUAPS = rsADO("DECMOUAPS")
rsZDECMOU0.DECMOUMOS = rsADO("DECMOUMOS")
rsZDECMOU0.DECMOUFIL = rsADO("DECMOUFIL")
Exit Function
Error_Handler:
rsZDECMOU0_GetBuffer = Error
End Function

