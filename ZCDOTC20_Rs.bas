Attribute VB_Name = "rsZCDOTC20"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Type typeZCDOTC20
    CDOTC2ETB       As Integer                        ' CODE ETABLISSEMENT
    CDOTC2AGE       As Integer                        ' AGENCE
    CDOTC2SER       As String * 2                     ' SERVICE
    CDOTC2SSE       As String * 2                     ' SOUS-SERVICE
    CDOTC2COP       As String * 3                     ' CODE OPERATION
    CDOTC2DOS       As Long                           ' NUMERO DOSSIER
    CDOTC2NUR       As Long                           ' N° RENOUVELLEMENT
    CDOTC2UTI       As Long                           ' N° UTILILSAT°./MODIF
    CDOTC2EVE       As String * 2                     ' EVENEMENT
    CDOTC2SEQ       As Long                           ' N° SEQUENCE
    CDOTC2COM       As String * 6                     ' Code commission
    CDOTC2DEV       As String * 3                     ' Devise
    CDOTC2CAT       As String * 3                     ' Catégorie client
    CDOTC2CLI       As String * 7                     ' N° Client
    CDOTC2DEB       As Long                           ' Date début effet
    CDOTC2FIN       As Long                           ' Date fin effet
    CDOTC2TVA       As String * 1                     ' TVA (O/N)
    CDOTC2PER       As String * 1                     ' Périodicité
    CDOTC2CUM       As String * 1                     ' Cumulable (O/N)
    CDOTC2MTF       As Currency                       ' Montant fixe
    CDOTC2IND       As String * 1                     ' Indivisibilité (O/N)
    CDOTC2AVE       As String * 1                     ' Avis à échéance
    CDOTC2MT1       As Long                           ' Montant tranche 1
    CDOTC2MT2       As Long                           ' Montant tranche 2
    CDOTC2MT3       As Long                           ' Montant tranche 3
    CDOTC2MT4       As Long                           ' Montant tranche 4
    CDOTC2MT5       As Long                           ' Montant tranche 5
    CDOTC2MT6       As Long                           ' Montant tranche 6
    CDOTC2TX1       As Double                         ' Taux tranche 1
    CDOTC2TX2       As Double                         ' Taux tranche 2
    CDOTC2TX3       As Double                         ' Taux tranche 3
    CDOTC2TX4       As Double                         ' Taux tranche 4
    CDOTC2TX5       As Double                         ' Taux tranche 5
    CDOTC2TX6       As Double                         ' Taux tranche 6
    CDOTC2REP       As String * 1                     '
End Type
Public Sub rsZCDOTC20_Init(rsYCDOTC20 As typeZCDOTC20)
rsYCDOTC20.CDOTC2ETB = 0
rsYCDOTC20.CDOTC2AGE = 0
rsYCDOTC20.CDOTC2SER = ""
rsYCDOTC20.CDOTC2SSE = ""
rsYCDOTC20.CDOTC2COP = ""
rsYCDOTC20.CDOTC2DOS = 0
rsYCDOTC20.CDOTC2NUR = 0
rsYCDOTC20.CDOTC2UTI = 0
rsYCDOTC20.CDOTC2EVE = ""
rsYCDOTC20.CDOTC2SEQ = 0
rsYCDOTC20.CDOTC2COM = ""
rsYCDOTC20.CDOTC2DEV = ""
rsYCDOTC20.CDOTC2CAT = ""
rsYCDOTC20.CDOTC2CLI = ""
rsYCDOTC20.CDOTC2DEB = 0
rsYCDOTC20.CDOTC2FIN = 0
rsYCDOTC20.CDOTC2TVA = ""
rsYCDOTC20.CDOTC2PER = ""
rsYCDOTC20.CDOTC2CUM = ""
rsYCDOTC20.CDOTC2MTF = 0
rsYCDOTC20.CDOTC2IND = ""
rsYCDOTC20.CDOTC2AVE = ""
rsYCDOTC20.CDOTC2MT1 = 0
rsYCDOTC20.CDOTC2MT2 = 0
rsYCDOTC20.CDOTC2MT3 = 0
rsYCDOTC20.CDOTC2MT4 = 0
rsYCDOTC20.CDOTC2MT5 = 0
rsYCDOTC20.CDOTC2MT6 = 0
rsYCDOTC20.CDOTC2TX1 = 0
rsYCDOTC20.CDOTC2TX2 = 0
rsYCDOTC20.CDOTC2TX3 = 0
rsYCDOTC20.CDOTC2TX4 = 0
rsYCDOTC20.CDOTC2TX5 = 0
rsYCDOTC20.CDOTC2TX6 = 0
End Sub
Public Function rsZCDOTC20_GetBuffer(rsAdo As ADODB.Recordset, rsZCDOTC20 As typeZCDOTC20)
On Error GoTo Error_Handler
rsZCDOTC20_GetBuffer = Null
rsZCDOTC20.CDOTC2ETB = rsAdo("CDOTC2ETB")
rsZCDOTC20.CDOTC2AGE = rsAdo("CDOTC2AGE")
rsZCDOTC20.CDOTC2SER = rsAdo("CDOTC2SER")
rsZCDOTC20.CDOTC2SSE = rsAdo("CDOTC2SSE")
rsZCDOTC20.CDOTC2COP = rsAdo("CDOTC2COP")
rsZCDOTC20.CDOTC2DOS = rsAdo("CDOTC2DOS")
rsZCDOTC20.CDOTC2NUR = rsAdo("CDOTC2NUR")
rsZCDOTC20.CDOTC2UTI = rsAdo("CDOTC2UTI")
rsZCDOTC20.CDOTC2EVE = rsAdo("CDOTC2EVE")
rsZCDOTC20.CDOTC2SEQ = rsAdo("CDOTC2SEQ")
rsZCDOTC20.CDOTC2COM = rsAdo("CDOTC2COM")
rsZCDOTC20.CDOTC2DEV = rsAdo("CDOTC2DEV")
rsZCDOTC20.CDOTC2CAT = rsAdo("CDOTC2CAT")
rsZCDOTC20.CDOTC2CLI = rsAdo("CDOTC2CLI")
rsZCDOTC20.CDOTC2DEB = rsAdo("CDOTC2DEB")
rsZCDOTC20.CDOTC2FIN = rsAdo("CDOTC2FIN")
rsZCDOTC20.CDOTC2TVA = rsAdo("CDOTC2TVA")
rsZCDOTC20.CDOTC2PER = rsAdo("CDOTC2PER")
rsZCDOTC20.CDOTC2CUM = rsAdo("CDOTC2CUM")
rsZCDOTC20.CDOTC2MTF = rsAdo("CDOTC2MTF")
rsZCDOTC20.CDOTC2IND = rsAdo("CDOTC2IND")
rsZCDOTC20.CDOTC2AVE = rsAdo("CDOTC2AVE")
rsZCDOTC20.CDOTC2MT1 = rsAdo("CDOTC2MT1")
rsZCDOTC20.CDOTC2MT2 = rsAdo("CDOTC2MT2")
rsZCDOTC20.CDOTC2MT3 = rsAdo("CDOTC2MT3")
rsZCDOTC20.CDOTC2MT4 = rsAdo("CDOTC2MT4")
rsZCDOTC20.CDOTC2MT5 = rsAdo("CDOTC2MT5")
rsZCDOTC20.CDOTC2MT6 = rsAdo("CDOTC2MT6")
rsZCDOTC20.CDOTC2TX1 = rsAdo("CDOTC2TX1")
rsZCDOTC20.CDOTC2TX2 = rsAdo("CDOTC2TX2")
rsZCDOTC20.CDOTC2TX3 = rsAdo("CDOTC2TX3")
rsZCDOTC20.CDOTC2TX4 = rsAdo("CDOTC2TX4")
rsZCDOTC20.CDOTC2TX5 = rsAdo("CDOTC2TX5")
rsZCDOTC20.CDOTC2TX6 = rsAdo("CDOTC2TX6")
rsZCDOTC20.CDOTC2REP = rsAdo("CDOTC2REP")
Exit Function
Error_Handler:
rsZCDOTC20_GetBuffer = Error
End Function

