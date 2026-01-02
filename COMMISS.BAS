Attribute VB_Name = "srvCommission"
Option Explicit

Public Const recBarèmeLen = 280 ' 34 + 246
Type typeTranche
    Seuil                  As Currency
    Taux                   As Double
End Type

Type typeBarème
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    Id                     As String * 10
    Télex                  As Currency
    Swift                  As Currency
    Minimum                As Currency
    TrancheNb              As Integer
    Tranche(10)            As typeTranche
End Type
    
Public arrBarème() As typeBarème
Public arrBarèmeNb As Integer
Public arrBarèmeNbMax As Integer
Public arrBarèmeIndex As Integer
Public arrBarèmeSuite As Boolean



Public Function Barème_Calc(ByVal Brut, recBarème As typeBarème) As Currency
Dim I As Integer
Barème_Calc = 0
For I = 1 To recBarème.TrancheNb
    If Brut <= recBarème.Tranche(I).Seuil Then
        Barème_Calc = Barème_Calc + (Brut - recBarème.Tranche(I - 1).Seuil) * recBarème.Tranche(I).Taux
        Exit For
    Else
        Barème_Calc = Barème_Calc + (recBarème.Tranche(I).Seuil - recBarème.Tranche(I - 1).Seuil) * recBarème.Tranche(I).Taux
    End If
Next I
If Barème_Calc <= recBarème.Minimum Then Barème_Calc = recBarème.Minimum

End Function

Public Sub srvBarème_Load(recBarème As typeBarème)
Select Case Trim(recBarème.Id)
    Case "OPTRFCORR"
        recBarème.Télex = 100 / tauxTTC
        recBarème.Swift = 50 / tauxTTC
        recBarème.Minimum = 85
        recBarème.TrancheNb = 2
        recBarème.Tranche(0).Seuil = 0: recBarème.Tranche(0).Taux = 0
        recBarème.Tranche(1).Seuil = 500000: recBarème.Tranche(1).Taux = 0.0005
        recBarème.Tranche(2).Seuil = 921999999999999#: recBarème.Tranche(2).Taux = 0.00025
    Case "OPTRFPM"
        recBarème.Télex = 125
        recBarème.Swift = 0
        recBarème.Minimum = 70  '!!!! frmCommission_Barème minimum =80 pour commission de change
        recBarème.TrancheNb = 4
        recBarème.Tranche(0).Seuil = 0: recBarème.Tranche(0).Taux = 0
        recBarème.Tranche(1).Seuil = 400000: recBarème.Tranche(1).Taux = 0.001
        recBarème.Tranche(2).Seuil = 800000: recBarème.Tranche(2).Taux = 0.0005
        recBarème.Tranche(3).Seuil = 3000000: recBarème.Tranche(3).Taux = 0.00025
        recBarème.Tranche(4).Seuil = 921999999999999#: recBarème.Tranche(4).Taux = 0.000125
    Case "OPTRFPP"
        recBarème.Télex = 125
        recBarème.Swift = 0
        recBarème.Minimum = 85 / tauxTTC
        recBarème.TrancheNb = 4
        recBarème.Tranche(0).Seuil = 0: recBarème.Tranche(0).Taux = 0
        recBarème.Tranche(1).Seuil = 400000: recBarème.Tranche(1).Taux = 0.0012
        recBarème.Tranche(2).Seuil = 800000: recBarème.Tranche(2).Taux = 0.0006
        recBarème.Tranche(3).Seuil = 3000000: recBarème.Tranche(3).Taux = 0.0003
        recBarème.Tranche(4).Seuil = 921999999999999#: recBarème.Tranche(4).Taux = 0.000125
    Case Else
        recBarème.Minimum = 0
        recBarème.TrancheNb = 0
        recBarème.Tranche(0).Seuil = 0: recBarème.Tranche(0).Taux = 0
End Select
End Sub
