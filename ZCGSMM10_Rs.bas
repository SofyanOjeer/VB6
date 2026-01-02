Attribute VB_Name = "rsZCGSMM10"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Type typeZCGSMM10
    CGSMM1ETA       As Integer                        ' ETABLISSEMENT
    CGSMM1AGE       As Integer                        ' AGENCE
    CGSMM1SER       As String * 2                     ' SERVICE
    CGSMM1SES       As String * 2                     ' SOUS SERVICE
    CGSMM1OPE       As String * 6                     ' OPERATION
    CGSMM1NAT       As String * 6                     ' NATURE
    CGSMM1NUM       As Long                           ' NUMERO
    CGSMM1MON       As Currency                       ' NOMINAL
    CGSMM1NBR       As Long                           ' NOMBRE OPE.
    CGSMM1DEV       As String * 3                     ' DEVISE
    CGSMM1CLI       As String * 8                     ' TYPE CLI/CLIENT
    CGSMM1COM       As String * 20                    ' COMPTE
    CGSMM1ENG       As Long                           ' DATE ENGAGEMENT
    CGSMM1DEB       As Long                           ' DATE DEBUT
    CGSMM1FIN       As Long                           ' DATE FIN
    CGSMM1DUR       As Long                           ' DUREE PREAVIS
    CGSMM1TYP       As String * 1                     ' TYPE DE PREAVIS
    CGSMM1AUT       As String * 3                     ' CODE AUTORISAT.
    CGSMM1CVL       As Currency                       ' NOMINAL CONTREV.
    CGSMM1NLO       As Long                           ' NOMBRE DE LOT

End Type
'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsZCGSMM10_GetBuffer(rsado As ADODB.Recordset, rsZCGSMM10 As typeZCGSMM10)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsZCGSMM10_GetBuffer = Null
rsZCGSMM10.CGSMM1ETA = rsado("CGSMM1ETA")
rsZCGSMM10.CGSMM1AGE = rsado("CGSMM1AGE")
rsZCGSMM10.CGSMM1SER = rsado("CGSMM1SER")
rsZCGSMM10.CGSMM1SES = rsado("CGSMM1SES")
rsZCGSMM10.CGSMM1OPE = rsado("CGSMM1OPE")
rsZCGSMM10.CGSMM1NAT = rsado("CGSMM1NAT")
rsZCGSMM10.CGSMM1NUM = rsado("CGSMM1NUM")
rsZCGSMM10.CGSMM1MON = rsado("CGSMM1MON")
rsZCGSMM10.CGSMM1NBR = rsado("CGSMM1NBR")
rsZCGSMM10.CGSMM1DEV = rsado("CGSMM1DEV")
rsZCGSMM10.CGSMM1CLI = rsado("CGSMM1CLI")
rsZCGSMM10.CGSMM1COM = rsado("CGSMM1COM")
rsZCGSMM10.CGSMM1ENG = rsado("CGSMM1ENG")
rsZCGSMM10.CGSMM1DEB = rsado("CGSMM1DEB")
rsZCGSMM10.CGSMM1FIN = rsado("CGSMM1FIN")
rsZCGSMM10.CGSMM1DUR = rsado("CGSMM1DUR")
rsZCGSMM10.CGSMM1TYP = rsado("CGSMM1TYP")
rsZCGSMM10.CGSMM1AUT = rsado("CGSMM1AUT")
rsZCGSMM10.CGSMM1CVL = rsado("CGSMM1CVL")
rsZCGSMM10.CGSMM1NLO = rsado("CGSMM1NLO")
Exit Function
Error_Handler:
rsZCGSMM10_GetBuffer = Error

End Function

Public Sub rsZCGSMM10_Init(rsZCGSMM10 As typeZCGSMM10)
rsZCGSMM10.CGSMM1ETA = 0
rsZCGSMM10.CGSMM1AGE = 0
rsZCGSMM10.CGSMM1SER = ""
rsZCGSMM10.CGSMM1SES = ""
rsZCGSMM10.CGSMM1OPE = ""
rsZCGSMM10.CGSMM1NAT = ""
rsZCGSMM10.CGSMM1NUM = 0
rsZCGSMM10.CGSMM1MON = 0
rsZCGSMM10.CGSMM1NBR = 0
rsZCGSMM10.CGSMM1DEV = ""
rsZCGSMM10.CGSMM1CLI = ""
rsZCGSMM10.CGSMM1COM = ""
rsZCGSMM10.CGSMM1ENG = 0
rsZCGSMM10.CGSMM1DEB = 0
rsZCGSMM10.CGSMM1FIN = 0
rsZCGSMM10.CGSMM1DUR = 0
rsZCGSMM10.CGSMM1TYP = ""
rsZCGSMM10.CGSMM1AUT = ""
rsZCGSMM10.CGSMM1CVL = 0
rsZCGSMM10.CGSMM1NLO = 0
End Sub

