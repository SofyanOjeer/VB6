Attribute VB_Name = "rsZMOUVEA0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Type typeZMOUVEA0
    MOUVEMETA       As Integer                        ' ETABLISSEMENT
    MOUVEMPLA       As Long                           ' NUMERO PLAN
    MOUVEMCOM       As String * 20                    ' NUMERO COMPTE
    MOUVEMMON       As Currency                        ' MONTANT
    MOUVEMDOP       As Long                           ' DATE D'OPERATION
    MOUVEMDVA       As Long                           ' DATE DE VALEUR
    MOUVEMDCO       As Long                           ' DATE COMPTABLE
    MOUVEMDTR       As Long                           ' DATE DE TRAITEMENT
    MOUVEMPIE       As Long                           ' NUMERO DE PIECE
    MOUVEMECR       As Long                           ' NUMERO D'ECRITURE
    MOUVEMOPE       As String * 3                     ' CODE OPERATION
    MOUVEMNUM       As Long                           ' NUMERO OPERATION
    MOUVEMSCH       As Integer                        ' CODE SCHEMA
    MOUVEMUTI       As Integer                        ' UTILISATEUR
    MOUVEMAGE       As Integer                        ' AGENCE OPERATRICE
    MOUVEMSER       As String * 2                     ' SERVICE OPERATEUR
    MOUVEMSSE       As String * 2                     ' S/SERVICE OPERATEUR
    MOUVEMEXO       As String * 1                     ' CODE EXONERATION
    MOUVEMANA       As String * 6                     ' CODE ANALYTIQUE
    MOUVEMBDF       As String * 3                     ' CODE BANQUE DE FR.
    MOUVEMANU       As String * 1                     ' CODE ANNULATION
    MOUVEMRET       As String * 1                     ' MOUVEMENT RETRO
    MOUVEMEVE       As String * 3                     ' EVENEMENT
    MOUVEMSAN       As String * 6                     ' STRUCT ANALY-CODE
    MOUVEMSAD       As String * 80                    ' STRUCT ANALY-DONNEES

End Type
Public Sub rsZMOUVEA0_Init(rsZMOUVEA0 As typeZMOUVEA0)
rsZMOUVEA0.MOUVEMETA = 0
rsZMOUVEA0.MOUVEMPLA = 0
rsZMOUVEA0.MOUVEMCOM = ""
rsZMOUVEA0.MOUVEMMON = 0
rsZMOUVEA0.MOUVEMDOP = 0
rsZMOUVEA0.MOUVEMDVA = 0
rsZMOUVEA0.MOUVEMDCO = 0
rsZMOUVEA0.MOUVEMDTR = 0
rsZMOUVEA0.MOUVEMPIE = 0
rsZMOUVEA0.MOUVEMECR = 0
rsZMOUVEA0.MOUVEMOPE = ""
rsZMOUVEA0.MOUVEMNUM = 0
rsZMOUVEA0.MOUVEMSCH = 0
rsZMOUVEA0.MOUVEMUTI = 0
rsZMOUVEA0.MOUVEMAGE = 0
rsZMOUVEA0.MOUVEMSER = ""
rsZMOUVEA0.MOUVEMSSE = ""
rsZMOUVEA0.MOUVEMEXO = ""
rsZMOUVEA0.MOUVEMANA = ""
rsZMOUVEA0.MOUVEMBDF = ""
rsZMOUVEA0.MOUVEMANU = ""
rsZMOUVEA0.MOUVEMRET = ""
rsZMOUVEA0.MOUVEMEVE = ""
rsZMOUVEA0.MOUVEMSAN = ""
rsZMOUVEA0.MOUVEMSAD = ""
End Sub
Public Function rsZMOUVEA0_GetBuffer(rsAdo As ADODB.Recordset, rsZMOUVEA0 As typeZMOUVEA0)
On Error GoTo Error_Handler
rsZMOUVEA0_GetBuffer = Null
rsZMOUVEA0.MOUVEMETA = rsAdo("MOUVEMETA")
rsZMOUVEA0.MOUVEMPLA = rsAdo("MOUVEMPLA")
rsZMOUVEA0.MOUVEMCOM = rsAdo("MOUVEMCOM")
rsZMOUVEA0.MOUVEMMON = rsAdo("MOUVEMMON")
rsZMOUVEA0.MOUVEMDOP = rsAdo("MOUVEMDOP")
rsZMOUVEA0.MOUVEMDVA = rsAdo("MOUVEMDVA")
rsZMOUVEA0.MOUVEMDCO = rsAdo("MOUVEMDCO")
rsZMOUVEA0.MOUVEMDTR = rsAdo("MOUVEMDTR")
rsZMOUVEA0.MOUVEMPIE = rsAdo("MOUVEMPIE")
rsZMOUVEA0.MOUVEMECR = rsAdo("MOUVEMECR")
rsZMOUVEA0.MOUVEMOPE = rsAdo("MOUVEMOPE")
rsZMOUVEA0.MOUVEMNUM = rsAdo("MOUVEMNUM")
rsZMOUVEA0.MOUVEMSCH = rsAdo("MOUVEMSCH")
rsZMOUVEA0.MOUVEMUTI = rsAdo("MOUVEMUTI")
rsZMOUVEA0.MOUVEMAGE = rsAdo("MOUVEMAGE")
rsZMOUVEA0.MOUVEMSER = rsAdo("MOUVEMSER")
rsZMOUVEA0.MOUVEMSSE = rsAdo("MOUVEMSSE")
rsZMOUVEA0.MOUVEMEXO = rsAdo("MOUVEMEXO")
rsZMOUVEA0.MOUVEMANA = rsAdo("MOUVEMANA")
rsZMOUVEA0.MOUVEMBDF = rsAdo("MOUVEMBDF")
rsZMOUVEA0.MOUVEMANU = rsAdo("MOUVEMANU")
rsZMOUVEA0.MOUVEMRET = rsAdo("MOUVEMRET")
rsZMOUVEA0.MOUVEMEVE = rsAdo("MOUVEMEVE")
rsZMOUVEA0.MOUVEMSAN = rsAdo("MOUVEMSAN")
rsZMOUVEA0.MOUVEMSAD = rsAdo("MOUVEMSAD")
Exit Function
Error_Handler:
rsZMOUVEA0_GetBuffer = Error
End Function
