Attribute VB_Name = "rsZCHGPAS0"
Option Explicit

'---------------------------------------------------------
Type typeZCHGPAS0
    CHGPASET       As Integer                        ' CODE ETABLISSEMENT
    CHGPASAG       As Integer                   '
    CHGPASNU       As Long                   '
    CHGPASAB       As String * 12                    ' SIGLE USUEL
    CHGPASN1       As String * 32                    ' NOM OU DESIGNATION
    CHGPASN2       As String * 32                    ' PRENOM/DESIGNATION
    CHGPASA1       As String * 32                    ' ADRESSE 1
    CHGPASA2       As String * 32                    ' ADRESSE 2
    CHGPASC1       As String * 6                     ' CODE POSTAL
    CHGPASVI       As String * 25                    ' BUREAU DISTRIBUTEUR
    CHGPASPA       As String * 25                    ' PAYS
    CHGPASRE       As String * 3                     ' CDE PAYS DE RESIDENC
    CHGPASLG       As String * 1                     ' LANGUE MESSAGERIE
    CHGPASCP       As String * 34                    ' COMPTE
    
    CHGPASUTC       As Long                    '
    CHGPASDTC       As Long                   '
    CHGPASUTM       As Long               '
    CHGPASDTM       As Long                  '

End Type
Public Sub rsZCHGPAS0_Init(rsYCHGPAS0 As typeZCHGPAS0)
rsYCHGPAS0.CHGPASET = 0
rsYCHGPAS0.CHGPASAG = 0
rsYCHGPAS0.CHGPASNU = ""
rsYCHGPAS0.CHGPASAB = ""
rsYCHGPAS0.CHGPASN1 = ""
rsYCHGPAS0.CHGPASN2 = ""
rsYCHGPAS0.CHGPASA1 = ""
rsYCHGPAS0.CHGPASA2 = ""
rsYCHGPAS0.CHGPASC1 = ""
rsYCHGPAS0.CHGPASVI = ""
rsYCHGPAS0.CHGPASPA = ""
rsYCHGPAS0.CHGPASRE = ""
rsYCHGPAS0.CHGPASLG = ""
rsYCHGPAS0.CHGPASCP = ""
rsYCHGPAS0.CHGPASUTC = 0
rsYCHGPAS0.CHGPASDTC = 0
rsYCHGPAS0.CHGPASUTM = 0
rsYCHGPAS0.CHGPASDTM = 0
End Sub
Public Function rsZCHGPAS0_GetBuffer(rsAdo As ADODB.Recordset, rsZCHGPAS0 As typeZCHGPAS0)
On Error GoTo Error_Handler
rsZCHGPAS0_GetBuffer = Null
rsZCHGPAS0.CHGPASET = rsAdo("CHGPASET")
rsZCHGPAS0.CHGPASAG = rsAdo("CHGPASAG")
rsZCHGPAS0.CHGPASNU = rsAdo("CHGPASNU")
rsZCHGPAS0.CHGPASAB = rsAdo("CHGPASAB")
rsZCHGPAS0.CHGPASN1 = rsAdo("CHGPASN1")
rsZCHGPAS0.CHGPASN2 = rsAdo("CHGPASN2")
rsZCHGPAS0.CHGPASA1 = rsAdo("CHGPASA1")
rsZCHGPAS0.CHGPASA2 = rsAdo("CHGPASA2")
rsZCHGPAS0.CHGPASC1 = rsAdo("CHGPASC1")
rsZCHGPAS0.CHGPASVI = rsAdo("CHGPASVI")
rsZCHGPAS0.CHGPASPA = rsAdo("CHGPASPA")
rsZCHGPAS0.CHGPASRE = rsAdo("CHGPASRE")
rsZCHGPAS0.CHGPASLG = rsAdo("CHGPASLG")
rsZCHGPAS0.CHGPASCP = rsAdo("CHGPASCP")
rsZCHGPAS0.CHGPASUTC = rsAdo("CHGPASUTC")
rsZCHGPAS0.CHGPASDTC = rsAdo("CHGPASDTC")
rsZCHGPAS0.CHGPASUTM = rsAdo("CHGPASUTM")
rsZCHGPAS0.CHGPASDTM = rsAdo("CHGPASDTM")

Exit Function
Error_Handler:
rsZCHGPAS0_GetBuffer = Error
End Function


