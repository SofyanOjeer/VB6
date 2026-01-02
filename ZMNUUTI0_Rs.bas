Attribute VB_Name = "rsZMNUUTI0"
'---------------------------------------------------------
Option Explicit
Type typeZMNUUTI0
    MNUUTIETB       As Integer                        ' ETABLISSEMENT
    MNUUTIREF       As Long                           ' reference lot
    MNUUTICUT       As Integer                        ' CODE UTILISATEUR
    MNUUTIGR2       As String * 10                    ' CODE GROUPE MENU
    MNUUTIGR3       As String * 10                    ' CODE GROUPE DROITS
    MNUUTIGR4       As String * 10                    ' CODE GROUPE METIER
    MNUUTIOUT       As String * 10                    ' FILE ATTENTE
    MNUUTILAN       As String * 1                     ' LANGUE
    MNUUTIMSE       As String * 1                     ' MENU SERVICE
    MNUUTIAGE       As Integer                        ' AGENCE DEFAUT
    MNUUTISER       As String * 2                     ' SERVICE DEFAUT
    MNUUTISRV       As String * 2                     ' SOUS-SERV. DEFAUT
    MNUUTIGRS       As String * 10                    ' GROUPE MENU SERVICE
    MNUUTIGEN       As Integer                        ' CODE GENERIQUE
    MNUUTIPOS       As String * 10                    ' POSTE TRAVAIL
    MNUUTIMAI       As String                     ' ADRESSE MAIL
    
End Type

'---------------------------------------------------------
'---------------------------------------------------------
Public Function rsZMNUUTI0_GetBuffer(rsAdo As ADODB.Recordset, rsZMNUUTI0 As typeZMNUUTI0)
'---------------------------------------------------------
On Error GoTo Error_Handler
rsZMNUUTI0_GetBuffer = Null
rsZMNUUTI0.MNUUTIETB = rsAdo("MNUUTIETB")
rsZMNUUTI0.MNUUTIREF = rsAdo("MNUUTIREF")
rsZMNUUTI0.MNUUTICUT = rsAdo("MNUUTICUT")
rsZMNUUTI0.MNUUTIGR2 = rsAdo("MNUUTIGR2")
rsZMNUUTI0.MNUUTIGR3 = rsAdo("MNUUTIGR3")
rsZMNUUTI0.MNUUTIGR4 = rsAdo("MNUUTIGR4")
rsZMNUUTI0.MNUUTIOUT = rsAdo("MNUUTIOUT")
rsZMNUUTI0.MNUUTILAN = rsAdo("MNUUTILAN")
rsZMNUUTI0.MNUUTIMSE = rsAdo("MNUUTIMSE")
rsZMNUUTI0.MNUUTIAGE = rsAdo("MNUUTIAGE")
rsZMNUUTI0.MNUUTISER = rsAdo("MNUUTISER")
rsZMNUUTI0.MNUUTISRV = rsAdo("MNUUTISRV")
rsZMNUUTI0.MNUUTIGRS = rsAdo("MNUUTIGRS")
rsZMNUUTI0.MNUUTIGEN = rsAdo("MNUUTIGEN")
rsZMNUUTI0.MNUUTIPOS = rsAdo("MNUUTIPOS")
rsZMNUUTI0.MNUUTIMAI = Trim(rsAdo("MNUUTIMAI"))

Exit Function

Error_Handler:

rsZMNUUTI0_GetBuffer = Error

End Function


'






