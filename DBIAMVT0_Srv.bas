Attribute VB_Name = "srvDBIAMVT0"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Type typeDBIAMVT0
    Obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    MOUVEMETA       As String * 2                      ' ETABLISSEMENT
    MOUVEMPLA       As Long                           ' NUMERO PLAN
    MOUVEMCOM       As String * 20                    ' NUMERO COMPTE
    MOUVEMMON       As Currency                       ' MONTANT
    MOUVEMDOP       As Long                           ' DATE D'OPERATION
    MOUVEMDVA       As Long                           ' DATE DE VALEUR
    MOUVEMDCO       As Long                           ' DATE COMPTABLE
    MOUVEMDTR       As Long                           ' DATE DE TRAITEMENT
    MOUVEMPIE       As Long                           ' NUMERO DE PIECE
    MOUVEMECR       As Long                           ' NUMERO D'ECRITURE
    MOUVEMOPE       As String * 3                     ' CODE OPERATION
    MOUVEMNUM       As Long                           ' NUMERO OPERATION
    MOUVEMSER       As String * 2                     ' SERVICE OPERATEUR
    MOUVEMSSE       As String * 2                     ' S/SERVICE OPERATEUR
    MOUVEMEXO       As String * 1                     ' CODE EXONERATION
    MOUVEMANA       As String * 6                     ' CODE ANALYTIQUE
    MOUVEMBDF       As String * 3                     ' CODE BANQUE DE FR.
    MOUVEMANU       As String * 1                     ' CODE ANNULATION
    MOUVEMRET       As String * 1                     ' MOUVEMENT RETRO
    MOUVEMEVE       As String * 3                     ' EVENEMENT
    
    LIBELLIB1       As String * 30                    ' Libellé 1
    LIBELLIB2       As String * 30                    ' Libellé 2
    LIBELLIB3       As String * 30                    ' Libellé 3
    LIBELLIB4       As String * 30                    ' Libellé 4
    
    COMPTEDEV       As String * 3                     ' TABLES BASE 013
    COMPTECLA       As Long                           ' CLASSE SECURITE
    SCHDOSNAT       As String * 3                     '
    SCHDOSNUM       As Long                           '
    SCHDOSSEQ       As Long                           '
    SCHPRENAT       As String * 3                     '
    
    BIAMVTSD0       As Currency                       ' solde

End Type

'---------------------------------------------------------
Public Function srvDBIAMVT0_GetBuffer_ODBC(rsADO As ADODB.Recordset, recDBIAMVT0 As typeDBIAMVT0)
'---------------------------------------------------------
On Error Resume Next 'GoTo Error_Handler
srvDBIAMVT0_GetBuffer_ODBC = Null

recDBIAMVT0.MOUVEMETA = rsADO("MOUVEMETA")
recDBIAMVT0.MOUVEMPLA = rsADO("MOUVEMPLA")
recDBIAMVT0.MOUVEMCOM = rsADO("MOUVEMCOM")
recDBIAMVT0.MOUVEMMON = rsADO("MOUVEMMON")
recDBIAMVT0.MOUVEMDOP = rsADO("MOUVEMDOP")
recDBIAMVT0.MOUVEMDVA = rsADO("MOUVEMDVA")
recDBIAMVT0.MOUVEMDCO = rsADO("MOUVEMDCO")
recDBIAMVT0.MOUVEMDTR = rsADO("MOUVEMDTR")
recDBIAMVT0.MOUVEMPIE = rsADO("MOUVEMPIE")
recDBIAMVT0.MOUVEMECR = rsADO("MOUVEMECR")
recDBIAMVT0.MOUVEMOPE = rsADO("MOUVEMOPE")
recDBIAMVT0.MOUVEMNUM = rsADO("MOUVEMNUM")
recDBIAMVT0.MOUVEMSER = rsADO("MOUVEMSER")
recDBIAMVT0.MOUVEMSSE = rsADO("MOUVEMSSE")
recDBIAMVT0.MOUVEMEXO = rsADO("MOUVEMEXO")
recDBIAMVT0.MOUVEMANA = rsADO("MOUVEMANA")
recDBIAMVT0.MOUVEMBDF = rsADO("MOUVEMBDF")
recDBIAMVT0.MOUVEMANU = rsADO("MOUVEMANU")
recDBIAMVT0.MOUVEMRET = rsADO("MOUVEMRET")
recDBIAMVT0.MOUVEMEVE = rsADO("MOUVEMEVE")
recDBIAMVT0.LIBELLIB1 = rsADO("LIBELLIB1")
recDBIAMVT0.LIBELLIB2 = rsADO("LIBELLIB2")
recDBIAMVT0.LIBELLIB3 = rsADO("LIBELLIB3")
recDBIAMVT0.LIBELLIB4 = rsADO("LIBELLIB4")
recDBIAMVT0.COMPTEDEV = rsADO("COMPTEDEV")
recDBIAMVT0.COMPTECLA = rsADO("COMPTECLA")
recDBIAMVT0.SCHDOSNAT = rsADO("SCHDOSNAT")
recDBIAMVT0.SCHDOSNUM = rsADO("SCHDOSNUM")
recDBIAMVT0.SCHDOSSEQ = rsADO("SCHDOSSEQ")
recDBIAMVT0.SCHPRENAT = rsADO("SCHPRENAT")
recDBIAMVT0.MOUVEMETA = rsADO("MOUVEMETA")
recDBIAMVT0.BIAMVTSD0 = rsADO("BIAMVTSD0")

Exit Function

Error_Handler:
srvDBIAMVT0_GetBuffer_ODBC = Error

End Function


