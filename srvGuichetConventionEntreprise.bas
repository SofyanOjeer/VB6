Attribute VB_Name = "srvGuichetConventionEntreprise"
Option Explicit

Type typeGuichetConventionEntreprise


 Numéro                      As Long
RaisonSociale               As String * 40
DénominationCommerciale     As String * 40
objetSocial                 As String * 40
SiègeSocial                 As String * 40
CréationDate                As String * 8
FormeJuridique              As String * 30
CapitalSocial               As String * 30
NuméroSiren                 As String * 20
CodeApe                     As String * 4
OriginePays                 As String * 40
Résidencepays               As String * 40
AdresseCourrier1            As String * 40
AdresseCourrier2            As String * 40
AdresseCourrier3            As String * 40
AdresseCourrier4            As String * 40
NuméroTéléphone             As String * 20
NuméroTélécopie             As String * 20
NuméroTélex                 As String * 20
MandataireNoms1             As String * 30
MandataireNoms2             As String * 30
MandataireNoms3             As String * 30
MandataireNoms4             As String * 30
MandataireNoms5             As String * 30
MandatairePrénoms1          As String * 30
MandatairePrénoms2          As String * 30
MandatairePrénoms3          As String * 30
MandatairePrénoms4          As String * 30
MandatairePrénoms5          As String * 30
MandataireQualité1          As String * 30
MandataireQualité2          As String * 30
MandataireQualité3          As String * 30
MandataireQualité4          As String * 30
MandataireQualité5          As String * 30
CompteNuméro                As String * 11
CompteCléRib                As String * 2


OptionDevise                As String * 1
optionRelevé                As String * 1
OptionChèquier              As String * 1

Procuration1NaissanceNom            As String * 40
Procuration1NaissanceDate           As String * 8
Procuration1NaissanceLieu           As String * 40
Procuration1NaissancePays           As String * 40
Procuration1NaissanceNationalité    As String * 30
Procuration1IdentitéPièce           As String * 40
Procuration1IdentitéDate            As String * 8
Procuration1IdentitéLieu            As String * 40
Procuration1IdentitéAutorité        As String * 40

Procuration2NaissanceNom            As String * 40
Procuration2NaissanceDate           As String * 8
Procuration2NaissanceLieu           As String * 40
Procuration2NaissancePays           As String * 40
Procuration2NaissanceNationalité    As String * 30
Procuration2IdentitéPièce           As String * 40
Procuration2IdentitéDate            As String * 8
Procuration2IdentitéLieu            As String * 40
Procuration2IdentitéAutorité        As String * 40
SignatureNom1                       As String * 30
SignatureNom2                       As String * 30
SignatureNom3                       As String * 30
SignatureNom4                       As String * 30
SignatureNom5                       As String * 30
SignaturePrénom1                    As String * 30
SignaturePrénom2                    As String * 30
SignaturePrénom3                    As String * 30
SignaturePrénom4                    As String * 30
SignaturePrénom5                    As String * 30
SignatureQualité1                   As String * 30
SignatureQualité2                   As String * 30
SignatureQualité3                   As String * 30
SignatureQualité4                   As String * 30
SignatureQualité5                   As String * 30
SignatureDate                       As String * 8
End Type
Public recGuichetConventionEntreprise As typeGuichetConventionEntreprise
Dim FileNumber As Integer
Dim RecLength As Long
Public Position As Long
Public Function Ecrire()
On Error GoTo Error_Handler
Put FileNumber, recGuichetConventionEntreprise.Numéro, recGuichetConventionEntreprise
Ecrire = Null
GoTo End_Function

Error_Handler:
     MsgBox "erreur " & Error
   Ecrire = Err
End_Function:
End Function
Public Function Lire()
On Error GoTo Error_Handler
Get FileNumber, recGuichetConventionEntreprise.Numéro, recGuichetConventionEntreprise
Lire = Null
GoTo End_Function

Error_Handler:
    MsgBox "erreur " & Error
    Lire = Err
End_Function:
End Function

Public Sub Ouverture()
RecLength = Len(recGuichetConventionEntreprise) + 40
FileNumber = FreeFile
Open "c:\BiaSrv\GuichetConventionEntreprise.dta" For Random Access Read Write As FileNumber Len = RecLength
End Sub

Public Sub Rec_Init()


recGuichetConventionEntreprise.RaisonSociale = ""
recGuichetConventionEntreprise.DénominationCommerciale = ""
recGuichetConventionEntreprise.objetSocial = ""
recGuichetConventionEntreprise.SiègeSocial = ""

recGuichetConventionEntreprise.CréationDate = "00000000"
recGuichetConventionEntreprise.FormeJuridique = ""

recGuichetConventionEntreprise.CapitalSocial = ""
recGuichetConventionEntreprise.NuméroSiren = ""
recGuichetConventionEntreprise.CodeApe = ""
recGuichetConventionEntreprise.OriginePays = ""
recGuichetConventionEntreprise.Résidencepays = ""

recGuichetConventionEntreprise.AdresseCourrier1 = ""
recGuichetConventionEntreprise.AdresseCourrier2 = ""
recGuichetConventionEntreprise.AdresseCourrier3 = ""
recGuichetConventionEntreprise.AdresseCourrier4 = ""

recGuichetConventionEntreprise.NuméroTéléphone = ""
recGuichetConventionEntreprise.NuméroTélécopie = ""
recGuichetConventionEntreprise.NuméroTélex = ""

recGuichetConventionEntreprise.MandataireNoms1 = ""
recGuichetConventionEntreprise.MandataireNoms2 = ""
recGuichetConventionEntreprise.MandataireNoms3 = ""
recGuichetConventionEntreprise.MandataireNoms4 = ""
recGuichetConventionEntreprise.MandataireNoms5 = ""

recGuichetConventionEntreprise.MandatairePrénoms1 = ""
recGuichetConventionEntreprise.MandatairePrénoms2 = ""
recGuichetConventionEntreprise.MandatairePrénoms3 = ""
recGuichetConventionEntreprise.MandatairePrénoms4 = ""
recGuichetConventionEntreprise.MandatairePrénoms5 = ""

recGuichetConventionEntreprise.MandataireQualité1 = ""
recGuichetConventionEntreprise.MandataireQualité2 = ""
recGuichetConventionEntreprise.MandataireQualité3 = ""
recGuichetConventionEntreprise.MandataireQualité4 = ""
recGuichetConventionEntreprise.MandataireQualité5 = ""

recGuichetConventionEntreprise.CompteNuméro = ""
recGuichetConventionEntreprise.CompteCléRib = ""

recGuichetConventionEntreprise.Procuration1NaissanceNom = ""
recGuichetConventionEntreprise.Procuration1NaissanceDate = "00000000"
recGuichetConventionEntreprise.Procuration1NaissanceLieu = ""
recGuichetConventionEntreprise.Procuration1NaissancePays = ""
recGuichetConventionEntreprise.Procuration1NaissanceNationalité = ""

recGuichetConventionEntreprise.Procuration1IdentitéPièce = ""
recGuichetConventionEntreprise.Procuration1IdentitéDate = "00000000"
recGuichetConventionEntreprise.Procuration1IdentitéLieu = ""
recGuichetConventionEntreprise.Procuration1IdentitéAutorité = ""

recGuichetConventionEntreprise.Procuration2NaissanceNom = ""
recGuichetConventionEntreprise.Procuration2NaissanceDate = "00000000"
recGuichetConventionEntreprise.Procuration2NaissanceLieu = ""
recGuichetConventionEntreprise.Procuration2NaissancePays = ""
recGuichetConventionEntreprise.Procuration2NaissanceNationalité = ""

recGuichetConventionEntreprise.Procuration2IdentitéPièce = ""
recGuichetConventionEntreprise.Procuration2IdentitéDate = "00000000"
recGuichetConventionEntreprise.Procuration2IdentitéLieu = ""
recGuichetConventionEntreprise.Procuration2IdentitéAutorité = ""

recGuichetConventionEntreprise.SignatureNom1 = ""
recGuichetConventionEntreprise.SignatureNom2 = ""
recGuichetConventionEntreprise.SignatureNom3 = ""
recGuichetConventionEntreprise.SignatureNom4 = ""
recGuichetConventionEntreprise.SignatureNom5 = ""

recGuichetConventionEntreprise.SignaturePrénom1 = ""
recGuichetConventionEntreprise.SignaturePrénom2 = ""
recGuichetConventionEntreprise.SignaturePrénom3 = ""
recGuichetConventionEntreprise.SignaturePrénom4 = ""
recGuichetConventionEntreprise.SignaturePrénom5 = ""

recGuichetConventionEntreprise.SignatureQualité1 = ""
recGuichetConventionEntreprise.SignatureQualité2 = ""
recGuichetConventionEntreprise.SignatureQualité3 = ""
recGuichetConventionEntreprise.SignatureQualité4 = ""
recGuichetConventionEntreprise.SignatureQualité5 = ""

recGuichetConventionEntreprise.SignatureDate = "00000000"

recGuichetConventionEntreprise.optionRelevé = "M"
recGuichetConventionEntreprise.OptionChèquier = "P"
recGuichetConventionEntreprise.OptionDevise = ""


End Sub



