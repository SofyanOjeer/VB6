Attribute VB_Name = "rsZBASFUT0"
'---------------------------------------------------------
Option Explicit
Type typeZBASFUT0
    BASFUTETA       As Integer                        ' ETABLISSEMENT
    BASFUTOPE       As String * 3                     ' OPERATION
    BASFUTAGE       As Integer                        ' AGENCE
    BASFUTSER       As String * 2                     ' SERVICE
    BASFUTSSE       As String * 2                     ' SOUS SERVICE
    BASFUTDOS       As Long                           ' DOSSIER
    BASFUTDTE       As Long                           ' DATE EVENEMENT
    BASFUTEVE       As String * 3                     ' EVENEMENT
    BASFUTNUM       As Long                           ' NUMERO EVENEMENT
    BASFUTTYP       As String * 1                     ' TYPE EVENEMENT
    BASFUTNAT       As String * 3                     ' NATURE OPERATION
    BASFUTDVA       As Long                           ' DATE DE VALEUR
    BASFUTMON       As Currency                       ' MONTANT
    BASFUTSEN       As String * 1                     ' SENS OPERATION
    BASFUTDEV       As String * 3                     ' DEVISE
    BASFUTCPT       As String * 20                    ' COMPTE
    BASFUTTCL       As String * 1                     ' CLIENT TIERS
    BASFUTCLI       As String * 7                     ' CONTREPARTIE
    BASFUTTAU       As String * 1                     ' TAUX VARIABLE
    BASFUTNAG       As Integer                        ' AGENCE NETTING
    BASFUTNSE       As String * 2                     ' SERVICE NETTING
    BASFUTNSS       As String * 2                     ' S SERVICE NETTING
    BASFUTNDO       As Long                           ' DOSSIER NETTING
    BASFUTLIB       As String * 30                    ' LIBELLE
End Type

'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsZBASFUT0_GetBuffer(rsAdo As ADODB.Recordset, rsZBASFUT0 As typeZBASFUT0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsZBASFUT0_GetBuffer = Null

rsZBASFUT0.BASFUTETA = rsAdo("BASFUTETA")
rsZBASFUT0.BASFUTOPE = rsAdo("BASFUTOPE")
rsZBASFUT0.BASFUTAGE = rsAdo("BASFUTAGE")
rsZBASFUT0.BASFUTSER = rsAdo("BASFUTSER")
rsZBASFUT0.BASFUTSSE = rsAdo("BASFUTSSE")
rsZBASFUT0.BASFUTDOS = rsAdo("BASFUTDOS")
rsZBASFUT0.BASFUTDTE = rsAdo("BASFUTDTE")
rsZBASFUT0.BASFUTEVE = rsAdo("BASFUTEVE")
rsZBASFUT0.BASFUTNUM = rsAdo("BASFUTNUM")
rsZBASFUT0.BASFUTTYP = rsAdo("BASFUTTYP")
rsZBASFUT0.BASFUTNAT = rsAdo("BASFUTNAT")
rsZBASFUT0.BASFUTDVA = rsAdo("BASFUTDVA")
rsZBASFUT0.BASFUTMON = rsAdo("BASFUTMON")
rsZBASFUT0.BASFUTSEN = rsAdo("BASFUTSEN")
rsZBASFUT0.BASFUTDEV = rsAdo("BASFUTDEV")
rsZBASFUT0.BASFUTCPT = rsAdo("BASFUTCPT")
rsZBASFUT0.BASFUTTCL = rsAdo("BASFUTTCL")
rsZBASFUT0.BASFUTCLI = rsAdo("BASFUTCLI")
rsZBASFUT0.BASFUTTAU = rsAdo("BASFUTTAU")
rsZBASFUT0.BASFUTNAG = rsAdo("BASFUTNAG")
rsZBASFUT0.BASFUTNSE = rsAdo("BASFUTNSE")
rsZBASFUT0.BASFUTNSS = rsAdo("BASFUTNSS")
rsZBASFUT0.BASFUTNDO = rsAdo("BASFUTNDO")
rsZBASFUT0.BASFUTLIB = rsAdo("BASFUTLIB")
Exit Function

Error_Handler:

rsZBASFUT0_GetBuffer = Error

End Function

'---------------------------------------------------------
Public Sub rsZBASFUT0_Init(rsZBASFUT0 As typeZBASFUT0)
'---------------------------------------------------------
rsZBASFUT0.BASFUTETA = 0
rsZBASFUT0.BASFUTOPE = ""
rsZBASFUT0.BASFUTAGE = 0
rsZBASFUT0.BASFUTSER = ""
rsZBASFUT0.BASFUTSSE = ""
rsZBASFUT0.BASFUTDOS = 0
rsZBASFUT0.BASFUTDTE = 0
rsZBASFUT0.BASFUTEVE = ""
rsZBASFUT0.BASFUTNUM = 0
rsZBASFUT0.BASFUTTYP = ""
rsZBASFUT0.BASFUTNAT = ""
rsZBASFUT0.BASFUTDVA = 0
rsZBASFUT0.BASFUTMON = 0
rsZBASFUT0.BASFUTSEN = ""
rsZBASFUT0.BASFUTDEV = ""
rsZBASFUT0.BASFUTCPT = ""
rsZBASFUT0.BASFUTTCL = ""
rsZBASFUT0.BASFUTCLI = ""
rsZBASFUT0.BASFUTTAU = ""
rsZBASFUT0.BASFUTNAG = 0
rsZBASFUT0.BASFUTNSE = ""
rsZBASFUT0.BASFUTNSS = ""
rsZBASFUT0.BASFUTNDO = 0
rsZBASFUT0.BASFUTLIB = ""
End Sub


'








