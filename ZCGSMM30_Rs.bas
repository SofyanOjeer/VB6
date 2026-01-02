Attribute VB_Name = "rsZCGSMM30"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Type typeZCGSMM30
    CGSMM3ETA       As Integer                        ' ETABLISSEMENT
    CGSMM3AGE       As Integer                        ' AGENCE
    CGSMM3SER       As String * 2                     ' SERVICE
    CGSMM3SES       As String * 2                     ' SOUS SERVICE
    CGSMM3OPE       As String * 6                     ' OPERATION
    CGSMM3NAT       As String * 6                     ' NATURE
    CGSMM3NUM       As Long                           ' NUMERO
    CGSMM3SEN       As String * 1                     ' SENS
    CGSMM3SEQ       As Long                           ' N° SEQUENCE
    CGSMM3DEV       As String * 3                     ' DEVISE
    CGSMM3REF       As String * 6                     ' CODE TAUX
    CGSMM3APP       As String * 1                     ' CODE APPLICAT°
    CGSMM3TAU       As Double                         ' TAUX FIXE
    CGSMM3MAR       As Double                         ' MARGE CLIENT
    CGSMM3MRC       As Double                         ' MARGE COMMERC.
    CGSMM3DVA       As Long                           ' DATE VAL CLIENT
    CGSMM3DTR       As Long                           ' DATE VAL TRESO
    CGSMM3DRG       As Long                           ' DATE REGLEMENT
    CGSMM3INT       As Currency                       ' INTERETS DS MOIS
    CGSMM3COU       As Currency                       ' INTERETS COURUS
    CGSMM3DEB       As Long                           ' DATE DEBUT PERIO
    CGSMM3FIN       As Long                           ' DATE FIN PERIODE
    CGSMM3ASS       As Currency                       ' MONTANT ASSIETTE
    CGSMM3NBJ       As Long                           ' NB JOUR OPE MOIS
    CGSMM3NBP       As Long                           ' NB JOUR PERIODE
    CGSMM3BAS       As Long                           ' BASE DEVISE
    CGSMM3MAC       As Currency                       ' MONT. MARGE COM.
    CGSMM3MIN       As Currency                       ' MONT. INTS.TRESO
    CGSMM3TXA       As Double                         ' TAUX D ANALYSE

End Type
'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsZCGSMM30_GetBuffer(rsado As ADODB.Recordset, rsZCGSMM30 As typeZCGSMM30)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsZCGSMM30_GetBuffer = Null
rsZCGSMM30.CGSMM3ETA = rsado("CGSMM3ETA")
rsZCGSMM30.CGSMM3AGE = rsado("CGSMM3AGE")
rsZCGSMM30.CGSMM3SER = rsado("CGSMM3SER")
rsZCGSMM30.CGSMM3SES = rsado("CGSMM3SES")
rsZCGSMM30.CGSMM3OPE = rsado("CGSMM3OPE")
rsZCGSMM30.CGSMM3NAT = rsado("CGSMM3NAT")
rsZCGSMM30.CGSMM3NUM = rsado("CGSMM3NUM")
rsZCGSMM30.CGSMM3SEN = rsado("CGSMM3SEN")
rsZCGSMM30.CGSMM3SEQ = rsado("CGSMM3SEQ")
rsZCGSMM30.CGSMM3DEV = rsado("CGSMM3DEV")
rsZCGSMM30.CGSMM3REF = rsado("CGSMM3REF")
rsZCGSMM30.CGSMM3APP = rsado("CGSMM3APP")
rsZCGSMM30.CGSMM3TAU = rsado("CGSMM3TAU")
rsZCGSMM30.CGSMM3MAR = rsado("CGSMM3MAR")
rsZCGSMM30.CGSMM3MRC = rsado("CGSMM3MRC")
rsZCGSMM30.CGSMM3DVA = rsado("CGSMM3DVA")
rsZCGSMM30.CGSMM3DTR = rsado("CGSMM3DTR")
rsZCGSMM30.CGSMM3DRG = rsado("CGSMM3DRG")
rsZCGSMM30.CGSMM3INT = rsado("CGSMM3INT")
rsZCGSMM30.CGSMM3COU = rsado("CGSMM3COU")
rsZCGSMM30.CGSMM3DEB = rsado("CGSMM3DEB")
rsZCGSMM30.CGSMM3FIN = rsado("CGSMM3FIN")
rsZCGSMM30.CGSMM3ASS = rsado("CGSMM3ASS")
rsZCGSMM30.CGSMM3NBJ = rsado("CGSMM3NBJ")
rsZCGSMM30.CGSMM3NBP = rsado("CGSMM3NBP")
rsZCGSMM30.CGSMM3BAS = rsado("CGSMM3BAS")
rsZCGSMM30.CGSMM3MAC = rsado("CGSMM3MAC")
rsZCGSMM30.CGSMM3MIN = rsado("CGSMM3MIN")
rsZCGSMM30.CGSMM3TXA = rsado("CGSMM3TXA")
Exit Function
Error_Handler:
rsZCGSMM30_GetBuffer = Error

End Function

Public Sub rsZCGSMM30_Init(rsZCGSMM30 As typeZCGSMM30)
rsZCGSMM30.CGSMM3ETA = 0
rsZCGSMM30.CGSMM3AGE = 0
rsZCGSMM30.CGSMM3SER = ""
rsZCGSMM30.CGSMM3SES = ""
rsZCGSMM30.CGSMM3OPE = ""
rsZCGSMM30.CGSMM3NAT = ""
rsZCGSMM30.CGSMM3NUM = 0
rsZCGSMM30.CGSMM3SEN = ""
rsZCGSMM30.CGSMM3SEQ = 0
rsZCGSMM30.CGSMM3DEV = ""
rsZCGSMM30.CGSMM3REF = ""
rsZCGSMM30.CGSMM3APP = ""
rsZCGSMM30.CGSMM3TAU = 0
rsZCGSMM30.CGSMM3MAR = 0
rsZCGSMM30.CGSMM3MRC = 0
rsZCGSMM30.CGSMM3DVA = 0
rsZCGSMM30.CGSMM3DTR = 0
rsZCGSMM30.CGSMM3DRG = 0
rsZCGSMM30.CGSMM3INT = 0
rsZCGSMM30.CGSMM3COU = 0
rsZCGSMM30.CGSMM3DEB = 0
rsZCGSMM30.CGSMM3FIN = 0
rsZCGSMM30.CGSMM3ASS = 0
rsZCGSMM30.CGSMM3NBJ = 0
rsZCGSMM30.CGSMM3NBP = 0
rsZCGSMM30.CGSMM3BAS = 0
rsZCGSMM30.CGSMM3MAC = 0
rsZCGSMM30.CGSMM3MIN = 0
rsZCGSMM30.CGSMM3TXA = 0
End Sub
