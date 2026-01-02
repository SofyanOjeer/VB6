Attribute VB_Name = "rsZCDOCO20"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Type typeZCDOCO20
    CDOCO2ETB       As Integer                        ' Etablissement
    CDOCO2AGE       As Integer                        ' Agence
    CDOCO2SER       As String * 2                     ' Service
    CDOCO2SSE       As String * 2                     ' Sous service
    CDOCO2COP       As String * 3                     ' Code Opération
    CDOCO2DOS       As Long                           ' N° Dossier
    CDOCO2NUR       As Long                           ' N° Renouv
    CDOCO2UTI       As Long                           ' N° Utilisation
    CDOCO2EVE       As String * 2                     ' Evénement
    CDOCO2SEQ       As Long                           ' N° Séquence
    CDOCO2SPE       As Long                           ' N° Séq Pério
    CDOCO2TVA       As String * 1                     ' TVA O/N
    CDOCO2PER       As String * 1                     ' Périodicité
    CDOCO2CUM       As String * 1                     ' Cumulable (O/N)
    CDOCO2IND       As String * 1                     ' Indivisibilité
    CDOCO2AVE       As String * 1                     ' Avis à échéance
    CDOCO2TYA       As String * 2                     ' Type Assiette
    CDOCO2MTA       As Currency 'Long                           ' Mt Assiette  '20050613 JPL
    CDOCO2JRB       As String * 1                     ' Jours Reel/Banc
    CDOCO2ANN       As String * 1                     ' Type année
    CDOCO2NBJ       As Long                           ' Nb jours
    CDOCO2MIN       As Long                           ' Montant minimum
    CDOCO2MAX       As Long                           ' Montant maximum
    CDOCO2SEU       As Long                           ' Seuil exonérat°
    CDOCO2MT1       As Long                           ' Mt tranche 1
    CDOCO2MT2       As Long                           ' Mt tranche 2
    CDOCO2MT3       As Long                           ' Mt tranche 3
    CDOCO2MT4       As Long                           ' Mt tranche 4
    CDOCO2MT5       As Long                           ' Mt tranche 5
    CDOCO2MT6       As Long                           ' Mt tranche 6
    CDOCO2TX1       As Double                         ' Taux tranche 1
    CDOCO2TX2       As Double                         ' Taux tranche 2
    CDOCO2TX3       As Double                         ' Taux tranche 3
    CDOCO2TX4       As Double                         ' Taux tranche 4
    CDOCO2TX5       As Double                         ' Taux tranche 5
    CDOCO2TX6       As Double                         ' Taux tranche 6
    CDOCO2MON       As Long                           ' Montant Calculé
    CDOCO2MTV       As Long                           ' Montant TVA
    CDOCO2MTE       As Long                           ' Montant Av.Extr
    CDOCO2REP       As String * 1                     '
End Type
Public Sub rsZCDOCO20_Init(rsYCDOCO20 As typeZCDOCO20)
rsYCDOCO20.CDOCO2ETB = 0
rsYCDOCO20.CDOCO2AGE = 0
rsYCDOCO20.CDOCO2SER = ""
rsYCDOCO20.CDOCO2SSE = ""
rsYCDOCO20.CDOCO2COP = ""
rsYCDOCO20.CDOCO2DOS = 0
rsYCDOCO20.CDOCO2NUR = 0
rsYCDOCO20.CDOCO2UTI = 0
rsYCDOCO20.CDOCO2EVE = ""
rsYCDOCO20.CDOCO2SEQ = 0
rsYCDOCO20.CDOCO2SPE = 0
rsYCDOCO20.CDOCO2TVA = ""
rsYCDOCO20.CDOCO2PER = ""
rsYCDOCO20.CDOCO2CUM = ""
rsYCDOCO20.CDOCO2IND = ""
rsYCDOCO20.CDOCO2AVE = ""
rsYCDOCO20.CDOCO2TYA = ""
rsYCDOCO20.CDOCO2MTA = 0
rsYCDOCO20.CDOCO2JRB = ""
rsYCDOCO20.CDOCO2ANN = ""
rsYCDOCO20.CDOCO2NBJ = 0
rsYCDOCO20.CDOCO2MIN = 0
rsYCDOCO20.CDOCO2MAX = 0
rsYCDOCO20.CDOCO2SEU = 0
rsYCDOCO20.CDOCO2MT1 = 0
rsYCDOCO20.CDOCO2MT2 = 0
rsYCDOCO20.CDOCO2MT3 = 0
rsYCDOCO20.CDOCO2MT4 = 0
rsYCDOCO20.CDOCO2MT5 = 0
rsYCDOCO20.CDOCO2MT6 = 0
rsYCDOCO20.CDOCO2TX1 = 0
rsYCDOCO20.CDOCO2TX2 = 0
rsYCDOCO20.CDOCO2TX3 = 0
rsYCDOCO20.CDOCO2TX4 = 0
rsYCDOCO20.CDOCO2TX5 = 0
rsYCDOCO20.CDOCO2TX6 = 0
rsYCDOCO20.CDOCO2MON = 0
rsYCDOCO20.CDOCO2MTV = 0
rsYCDOCO20.CDOCO2MTE = 0
End Sub
Public Function rsZCDOCO20_GetBuffer(rsAdo As ADODB.Recordset, rsZCDOCO20 As typeZCDOCO20)
On Error GoTo Error_Handler
rsZCDOCO20_GetBuffer = Null
rsZCDOCO20.CDOCO2ETB = rsAdo("CDOCO2ETB")
rsZCDOCO20.CDOCO2AGE = rsAdo("CDOCO2AGE")
rsZCDOCO20.CDOCO2SER = rsAdo("CDOCO2SER")
rsZCDOCO20.CDOCO2SSE = rsAdo("CDOCO2SSE")
rsZCDOCO20.CDOCO2COP = rsAdo("CDOCO2COP")
rsZCDOCO20.CDOCO2DOS = rsAdo("CDOCO2DOS")
rsZCDOCO20.CDOCO2NUR = rsAdo("CDOCO2NUR")
rsZCDOCO20.CDOCO2UTI = rsAdo("CDOCO2UTI")
rsZCDOCO20.CDOCO2EVE = rsAdo("CDOCO2EVE")
rsZCDOCO20.CDOCO2SEQ = rsAdo("CDOCO2SEQ")
rsZCDOCO20.CDOCO2SPE = rsAdo("CDOCO2SPE")
rsZCDOCO20.CDOCO2TVA = rsAdo("CDOCO2TVA")
rsZCDOCO20.CDOCO2PER = rsAdo("CDOCO2PER")
rsZCDOCO20.CDOCO2CUM = rsAdo("CDOCO2CUM")
rsZCDOCO20.CDOCO2IND = rsAdo("CDOCO2IND")
rsZCDOCO20.CDOCO2AVE = rsAdo("CDOCO2AVE")
rsZCDOCO20.CDOCO2TYA = rsAdo("CDOCO2TYA")
rsZCDOCO20.CDOCO2MTA = rsAdo("CDOCO2MTA")
rsZCDOCO20.CDOCO2JRB = rsAdo("CDOCO2JRB")
rsZCDOCO20.CDOCO2ANN = rsAdo("CDOCO2ANN")
rsZCDOCO20.CDOCO2NBJ = rsAdo("CDOCO2NBJ")
rsZCDOCO20.CDOCO2MIN = rsAdo("CDOCO2MIN")
rsZCDOCO20.CDOCO2MAX = rsAdo("CDOCO2MAX")
rsZCDOCO20.CDOCO2SEU = rsAdo("CDOCO2SEU")
rsZCDOCO20.CDOCO2MT1 = rsAdo("CDOCO2MT1")
rsZCDOCO20.CDOCO2MT2 = rsAdo("CDOCO2MT2")
rsZCDOCO20.CDOCO2MT3 = rsAdo("CDOCO2MT3")
rsZCDOCO20.CDOCO2MT4 = rsAdo("CDOCO2MT4")
rsZCDOCO20.CDOCO2MT5 = rsAdo("CDOCO2MT5")
rsZCDOCO20.CDOCO2MT6 = rsAdo("CDOCO2MT6")
rsZCDOCO20.CDOCO2TX1 = rsAdo("CDOCO2TX1")
rsZCDOCO20.CDOCO2TX2 = rsAdo("CDOCO2TX2")
rsZCDOCO20.CDOCO2TX3 = rsAdo("CDOCO2TX3")
rsZCDOCO20.CDOCO2TX4 = rsAdo("CDOCO2TX4")
rsZCDOCO20.CDOCO2TX5 = rsAdo("CDOCO2TX5")
rsZCDOCO20.CDOCO2TX6 = rsAdo("CDOCO2TX6")
rsZCDOCO20.CDOCO2MON = rsAdo("CDOCO2MON")
rsZCDOCO20.CDOCO2MTV = rsAdo("CDOCO2MTV")
rsZCDOCO20.CDOCO2MTE = rsAdo("CDOCO2MTE")
rsZCDOCO20.CDOCO2REP = rsAdo("CDOCO2REP")
Exit Function
Error_Handler:
rsZCDOCO20_GetBuffer = Error
End Function
