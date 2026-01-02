Attribute VB_Name = "rsZDORCPT0"
Option Explicit
Type typeZDORCPT0
    DORCPTETA       As Integer                        ' ETABLISSEMENT
    DORCPTPLA       As Long                           ' NUMERO DU PLAN
    DORCPTCOM       As String * 20                    ' COMPTE
    DORCPTDOR       As String * 1                     ' COMPTE DORMANT
    DORCPTDDO       As Long                           ' DATE DORMANT
    DORCPTDMV       As Long                           ' DATE DERNIER MVT
    DORCPTDDE       As Long                           ' DT DERNIERE FACTU
    DORCPTDPR       As Long                           ' D.PROCHAINE FACTU
    DORCPTCOD       As Integer                        ' CODE UTILISATEUR
    DORCPTDMO       As Long                           ' DATE CHANGEM.ETAT
    DORCPTDRE       As Long                           ' DATE REPORT
    DORCPTMAJ       As Long                           ' DATE MAJ DORMANT

End Type
Public Function rsZDORCPT0_GetBuffer(rsAdo As ADODB.Recordset, rsZDORCPT0 As typeZDORCPT0)
On Error GoTo Error_Handler
rsZDORCPT0_GetBuffer = Null
rsZDORCPT0.DORCPTETA = rsAdo("DORCPTETA")
rsZDORCPT0.DORCPTPLA = rsAdo("DORCPTPLA")
rsZDORCPT0.DORCPTCOM = rsAdo("DORCPTCOM")
rsZDORCPT0.DORCPTDOR = rsAdo("DORCPTDOR")
rsZDORCPT0.DORCPTDDO = rsAdo("DORCPTDDO")
rsZDORCPT0.DORCPTDMV = rsAdo("DORCPTDMV")
rsZDORCPT0.DORCPTDDE = rsAdo("DORCPTDDE")
rsZDORCPT0.DORCPTDPR = rsAdo("DORCPTDPR")
rsZDORCPT0.DORCPTCOD = rsAdo("DORCPTCOD")
rsZDORCPT0.DORCPTDMO = rsAdo("DORCPTDMO")
rsZDORCPT0.DORCPTDRE = rsAdo("DORCPTDRE")
rsZDORCPT0.DORCPTMAJ = rsAdo("DORCPTMAJ")
Exit Function
Error_Handler:
rsZDORCPT0_GetBuffer = Error
End Function
