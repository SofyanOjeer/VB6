Attribute VB_Name = "rsZCDOSWI0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Type typeZCDOSWI0
    CDOSWIETB       As Integer                        ' CODE ETABLISSEMENT
    CDOSWIAGE       As Integer                        ' AGENCE
    CDOSWISER       As String * 2                     ' SERVICE
    CDOSWISSE       As String * 2                     ' SOUS-SERVICE
    CDOSWICOP       As String * 3                     ' CODE OPERATION
    CDOSWIDOS       As Long                           ' NUMERO DOSSIER
    CDOSWINUR       As Long                           ' N° RENOUVELLEMENT
    CDOSWIUTI       As Long                           ' N° UTILISATION
    CDOSWIPAI       As Long                           ' N° PAIEMENT
    CDOSWIREG       As Long                           ' N° REGLEMENT/ENCAIS
    CDOSWIBER       As String * 1                     ' BENEFICIAIR CLI/TIE
    CDOSWIBEN       As String * 7                     ' BENEFICIAIRE EXPORT
    CDOSWIBAR       As String * 1                     ' BANQU.BENEF.CLI/TIE
    CDOSWIBAB       As String * 7                     ' BANQUE BENEF
    CDOSWIBDE       As String * 12                    ' BIC BQDES
    CDOSWIBIN       As String * 12                    ' BIC BQINT
    CDOSWIBBD       As String * 12                    ' BIC BQBAD
    CDOSWIBBE       As String * 12                    ' BIC BQBEN
    CDOSWIBBA       As String * 12                    ' BIC BQBAN
    CDOSWIDDR       As Long                           ' DT DEM RBT
    CDOSWIDAV       As Long                           ' DT AVIS PAIE
    CDOSWILI1       As String * 79                    ' LIBEL AVI
    CDOSWILI2       As String * 79                    ' LIBEL AVI
    CDOSWILI3       As String * 79                    ' LIBEL AVI
    CDOSWILI4       As String * 79                    ' LIBEL AVI
    CDOSWIIBD       As String * 34                    ' IBAN BQ EMETT/DESTI
    CDOSWIIBB       As String * 34                    ' IBAN BQ BENEF
    CDOSWICBE       As String * 1                     ' CODE IBAN BENEF
    CDOSWIIBE       As String * 34                    ' IBAN BENEF.
    CDOSWICHA       As String * 1                     ' CHARGES O/B/S

End Type
Public Sub rsZCDOSWI0_Init(rsYCDOSWI0 As typeZCDOSWI0)
rsYCDOSWI0.CDOSWIETB = 0
rsYCDOSWI0.CDOSWIAGE = 0
rsYCDOSWI0.CDOSWISER = ""
rsYCDOSWI0.CDOSWISSE = ""
rsYCDOSWI0.CDOSWICOP = ""
rsYCDOSWI0.CDOSWIDOS = 0
rsYCDOSWI0.CDOSWINUR = 0
rsYCDOSWI0.CDOSWIUTI = 0
rsYCDOSWI0.CDOSWIPAI = 0
rsYCDOSWI0.CDOSWIREG = 0
rsYCDOSWI0.CDOSWIBER = ""
rsYCDOSWI0.CDOSWIBEN = ""
rsYCDOSWI0.CDOSWIBAR = ""
rsYCDOSWI0.CDOSWIBAB = ""
rsYCDOSWI0.CDOSWIBDE = ""
rsYCDOSWI0.CDOSWIBIN = ""
rsYCDOSWI0.CDOSWIBBD = ""
rsYCDOSWI0.CDOSWIBBE = ""
rsYCDOSWI0.CDOSWIBBA = ""
rsYCDOSWI0.CDOSWIDDR = 0
rsYCDOSWI0.CDOSWIDAV = 0
rsYCDOSWI0.CDOSWILI1 = ""
rsYCDOSWI0.CDOSWILI2 = ""
rsYCDOSWI0.CDOSWILI3 = ""
rsYCDOSWI0.CDOSWILI4 = ""
rsYCDOSWI0.CDOSWIIBD = ""
rsYCDOSWI0.CDOSWIIBB = ""
rsYCDOSWI0.CDOSWICBE = ""
rsYCDOSWI0.CDOSWIIBE = ""
rsYCDOSWI0.CDOSWICHA = ""
End Sub
Public Function rsZCDOSWI0_GetBuffer(rsAdo As ADODB.Recordset, rsZCDOSWI0 As typeZCDOSWI0)
On Error GoTo Error_Handler
rsZCDOSWI0_GetBuffer = Null
rsZCDOSWI0.CDOSWIETB = rsAdo("CDOSWIETB")
rsZCDOSWI0.CDOSWIAGE = rsAdo("CDOSWIAGE")
rsZCDOSWI0.CDOSWISER = rsAdo("CDOSWISER")
rsZCDOSWI0.CDOSWISSE = rsAdo("CDOSWISSE")
rsZCDOSWI0.CDOSWICOP = rsAdo("CDOSWICOP")
rsZCDOSWI0.CDOSWIDOS = rsAdo("CDOSWIDOS")
rsZCDOSWI0.CDOSWINUR = rsAdo("CDOSWINUR")
rsZCDOSWI0.CDOSWIUTI = rsAdo("CDOSWIUTI")
rsZCDOSWI0.CDOSWIPAI = rsAdo("CDOSWIPAI")
rsZCDOSWI0.CDOSWIREG = rsAdo("CDOSWIREG")
rsZCDOSWI0.CDOSWIBER = rsAdo("CDOSWIBER")
rsZCDOSWI0.CDOSWIBEN = rsAdo("CDOSWIBEN")
rsZCDOSWI0.CDOSWIBAR = rsAdo("CDOSWIBAR")
rsZCDOSWI0.CDOSWIBAB = rsAdo("CDOSWIBAB")
rsZCDOSWI0.CDOSWIBDE = rsAdo("CDOSWIBDE")
rsZCDOSWI0.CDOSWIBIN = rsAdo("CDOSWIBIN")
rsZCDOSWI0.CDOSWIBBD = rsAdo("CDOSWIBBD")
rsZCDOSWI0.CDOSWIBBE = rsAdo("CDOSWIBBE")
rsZCDOSWI0.CDOSWIBBA = rsAdo("CDOSWIBBA")
rsZCDOSWI0.CDOSWIDDR = rsAdo("CDOSWIDDR")
rsZCDOSWI0.CDOSWIDAV = rsAdo("CDOSWIDAV")
rsZCDOSWI0.CDOSWILI1 = rsAdo("CDOSWILI1")
rsZCDOSWI0.CDOSWILI2 = rsAdo("CDOSWILI2")
rsZCDOSWI0.CDOSWILI3 = rsAdo("CDOSWILI3")
rsZCDOSWI0.CDOSWILI4 = rsAdo("CDOSWILI4")
rsZCDOSWI0.CDOSWIIBD = rsAdo("CDOSWIIBD")
rsZCDOSWI0.CDOSWIIBB = rsAdo("CDOSWIIBB")
rsZCDOSWI0.CDOSWICBE = rsAdo("CDOSWICBE")
rsZCDOSWI0.CDOSWIIBE = rsAdo("CDOSWIIBE")
rsZCDOSWI0.CDOSWICHA = rsAdo("CDOSWICHA")
Exit Function
Error_Handler:
rsZCDOSWI0_GetBuffer = Error
End Function

