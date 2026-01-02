Attribute VB_Name = "prtMontant"
Option Explicit

Public Function CentaineEnLettres(X3 As String, strCent As String) As String
Dim Unité As String * 1
Dim CentEnLettres As String, DizaineEnLettres As String
CentaineEnLettres = ""
Unité = Mid$(X3, 3, 1)
If Mid$(X3, 2, 2) = "00" Then
    CentEnLettres = strCent
Else
    CentEnLettres = "cent "
End If
'-----------------centaines-------------

Select Case Mid$(X3, 1, 1)
    Case "1": CentaineEnLettres = "cent "
    Case "2": CentaineEnLettres = "deux " & CentEnLettres
    Case "3": CentaineEnLettres = "trois " & CentEnLettres
    Case "4": CentaineEnLettres = "quatre " & CentEnLettres
    Case "5": CentaineEnLettres = "cinq " & CentEnLettres
    Case "6": CentaineEnLettres = "six " & CentEnLettres
    Case "7": CentaineEnLettres = "sept " & CentEnLettres
    Case "8": CentaineEnLettres = "huit " & CentEnLettres
    Case "9": CentaineEnLettres = "neuf " & CentEnLettres
End Select
'----------------dizaine----------------
Select Case Mid$(X3, 2, 1)
    Case "0": DizaineEnLettres = Unité0Enlettres(Unité)
    Case "1": DizaineEnLettres = Unité1EnLettres(Unité)
    Case "2": DizaineEnLettres = "vingt " & UnitéEnLettres(Unité)
    Case "3": DizaineEnLettres = "trente " & UnitéEnLettres(Unité)
    Case "4": DizaineEnLettres = "quarante " & UnitéEnLettres(Unité)
    Case "5": DizaineEnLettres = "cinquante " & UnitéEnLettres(Unité)
    Case "6": DizaineEnLettres = "soixante " & UnitéEnLettres(Unité)
    Case "7": DizaineEnLettres = Unité7EnLettres(Unité)
    Case "8": DizaineEnLettres = Unité8EnLettres(Unité)
    Case "9": DizaineEnLettres = "quatre-vingt " & Unité1EnLettres(Unité)
End Select
CentaineEnLettres = CentaineEnLettres & DizaineEnLettres

End Function

Public Function MontantEnLettres(Valeur As Currency, libDev As String) As String

Dim strValeur As String
strValeur = Format$(Valeur, "000000000000.00")
MontantEnLettres = ""
Select Case Mid$(strValeur, 1, 3)
    Case "000"
    Case "001": MontantEnLettres = " un milliard "
    Case Else: MontantEnLettres = CentaineEnLettres(Mid$(strValeur, 1, 3), "cent ") & "milliards "
End Select

'---------------------------------------
Select Case Mid$(strValeur, 4, 3)
       
          Case "000"
          Case "001": MontantEnLettres = MontantEnLettres & " un million "
         Case Else: MontantEnLettres = MontantEnLettres & CentaineEnLettres(Mid$(strValeur, 4, 3), "cent ") & "millions "
End Select

Select Case Mid$(strValeur, 7, 3)
          Case "000"
          Case "001": MontantEnLettres = MontantEnLettres & "  mille "
         Case Else: MontantEnLettres = MontantEnLettres & CentaineEnLettres(Mid$(strValeur, 7, 3), "cent ") & "mille "
End Select

Mid$(strValeur, 13, 1) = "0"
MontantEnLettres = MontantEnLettres & CentaineEnLettres(Mid$(strValeur, 10, 3), "cents ") & " " & Trim(libDev) & " " & CentaineEnLettres(Mid$(strValeur, 13, 3), "cent ")

End Function
Public Function AmountInLetters(Valeur As Currency, libDev As String) As String
Dim strValeur As String
strValeur = Format$(Valeur, "000000000000.00")
AmountInLetters = ""
Select Case Mid$(strValeur, 1, 3)
    Case "000"
    Case "001": AmountInLetters = " one milliard "
    Case Else: AmountInLetters = HundredInLetters(Mid$(strValeur, 1, 3)) & "milliards "
End Select

'---------------------------------------
Select Case Mid$(strValeur, 4, 3)
       
          Case "000"
          Case "001": AmountInLetters = AmountInLetters & " one million "
         Case Else: AmountInLetters = AmountInLetters & HundredInLetters(Mid$(strValeur, 4, 3)) & "millions "
End Select

Select Case Mid$(strValeur, 7, 3)
          Case "000"
          Case "001": AmountInLetters = AmountInLetters & "  thousand "
         Case Else: AmountInLetters = AmountInLetters & HundredInLetters(Mid$(strValeur, 7, 3)) & "thousands "
End Select

Mid$(strValeur, 13, 1) = "0"
AmountInLetters = AmountInLetters & HundredInLetters(Mid$(strValeur, 10, 3)) & " " & Trim(libDev) & " " & HundredInLetters(Mid$(strValeur, 13, 3))

End Function

Public Function HundredInLetters(X3 As String) As String
Dim Unité As String * 1, strTiret As String * 1
Dim CentEnLettres As String, TenInLetters As String
HundredInLetters = ""
Unité = Mid$(X3, 3, 1)
'-----------------centaines-------------

Select Case Mid$(X3, 1, 1)
    Case "1": HundredInLetters = "one hundred "
    Case "2": HundredInLetters = "two hundreds "
    Case "3": HundredInLetters = "three hundreds "
    Case "4": HundredInLetters = "four hundreds "
    Case "5": HundredInLetters = "five hundreds "
    Case "6": HundredInLetters = "six hundreds "
    Case "7": HundredInLetters = "seven hundreds "
    Case "8": HundredInLetters = "eight hundreds "
    Case "9": HundredInLetters = "nine hundreds "
End Select
'----------------dizaine----------------
If Mid$(X3, 3, 1) = "0" Then
    strTiret = " "
Else
    strTiret = "-"
End If
Select Case Mid$(X3, 2, 1)
    Case "0": TenInLetters = UnitInLetters(Unité)
    Case "1": TenInLetters = Unit1InLetters(Unité)
    Case "2": TenInLetters = "twenty" & strTiret & UnitInLetters(Unité)
    Case "3": TenInLetters = "thirty" & strTiret & UnitInLetters(Unité)
    Case "4": TenInLetters = "forty" & strTiret & UnitInLetters(Unité)
    Case "5": TenInLetters = "fifty" & strTiret & UnitInLetters(Unité)
    Case "6": TenInLetters = "sixty" & strTiret & UnitInLetters(Unité)
    Case "7": TenInLetters = "seventy" & strTiret & UnitInLetters(Unité)
    Case "8": TenInLetters = "eighty" & strTiret & UnitInLetters(Unité)
    Case "9": TenInLetters = "ninety" & strTiret & UnitInLetters(Unité)
End Select
HundredInLetters = HundredInLetters & TenInLetters

End Function
Public Function UnitInLetters(X1 As String) As String
Select Case X1
    Case "1": UnitInLetters = "one "
    Case "2": UnitInLetters = "two "
    Case "3": UnitInLetters = "three "
    Case "4": UnitInLetters = "four "
    Case "5": UnitInLetters = "five "
    Case "6": UnitInLetters = "six "
    Case "7": UnitInLetters = "seven "
    Case "8": UnitInLetters = "eight "
    Case "9": UnitInLetters = "nine "
   End Select

End Function
Public Function Unit1InLetters(X1 As String) As String
Select Case X1
    Case "0": Unit1InLetters = "ten "
    Case "1": Unit1InLetters = "eleven "
    Case "2": Unit1InLetters = "twelve "
    Case "3": Unit1InLetters = "thirteen "
    Case "4": Unit1InLetters = "forteen "
    Case "5": Unit1InLetters = "fifteen "
    Case "6": Unit1InLetters = "sixteen "
    Case "7": Unit1InLetters = "seventeen "
    Case "8": Unit1InLetters = "eighteen "
    Case "9": Unit1InLetters = "nineteen "
   End Select

End Function



Public Function Unité1EnLettres(X1 As String) As String
Select Case X1
    Case "0": Unité1EnLettres = "dix "
    Case "1": Unité1EnLettres = "onze "
    Case "2": Unité1EnLettres = "douze "
    Case "3": Unité1EnLettres = "treize "
    Case "4": Unité1EnLettres = "quatorze "
    Case "5": Unité1EnLettres = "quinze "
    Case "6": Unité1EnLettres = "seize "
    Case "7": Unité1EnLettres = "dix-sept "
    Case "8": Unité1EnLettres = "dix-huit "
    Case "9": Unité1EnLettres = "dix-neuf "
   End Select

End Function


Public Function Unité7EnLettres(X1 As String) As String
Select Case X1
    Case "0": Unité7EnLettres = "soixante dix "
    Case "1": Unité7EnLettres = "soixante et onze "
    Case "2": Unité7EnLettres = "soixante douze "
    Case "3": Unité7EnLettres = "soixante treize "
    Case "4": Unité7EnLettres = "soixante quatorze "
    Case "5": Unité7EnLettres = "soixante quinze "
    Case "6": Unité7EnLettres = "soixante seize "
    Case "7": Unité7EnLettres = "soixante dix-sept "
    Case "8": Unité7EnLettres = "soixante dix-huit "
    Case "9": Unité7EnLettres = "soixante dix-neuf "
   End Select

End Function
Public Function Unité8EnLettres(X1 As String) As String
Select Case X1
    Case "0": Unité8EnLettres = "quatre-vingts "
    Case "1": Unité8EnLettres = "quatre-vingt un "
    Case "2": Unité8EnLettres = "quatre-vingt deux "
    Case "3": Unité8EnLettres = "quatre-vingt trois "
    Case "4": Unité8EnLettres = "quatre-vingt quatre "
    Case "5": Unité8EnLettres = "quatre-vingt cinq "
    Case "6": Unité8EnLettres = "quatre-vingt six "
    Case "7": Unité8EnLettres = "quatre-vingt sept "
    Case "8": Unité8EnLettres = "quatre-vingt huit "
    Case "9": Unité8EnLettres = "quatre-vingt neuf "
   End Select

End Function

Public Function Unité0Enlettres(X1 As String) As String
Select Case X1
    Case "0"
    Case "1": Unité0Enlettres = "un "
    Case "2": Unité0Enlettres = "deux "
    Case "3": Unité0Enlettres = "trois "
    Case "4": Unité0Enlettres = "quatre "
    Case "5": Unité0Enlettres = "cinq "
    Case "6": Unité0Enlettres = "six "
    Case "7": Unité0Enlettres = "sept "
    Case "8": Unité0Enlettres = "huit "
    Case "9": Unité0Enlettres = "neuf "
   End Select

End Function

Public Function UnitéEnLettres(X1 As String) As String
Select Case X1
    Case "1": UnitéEnLettres = "et un "
    Case "2": UnitéEnLettres = "deux "
    Case "3": UnitéEnLettres = "trois "
    Case "4": UnitéEnLettres = "quatre "
    Case "5": UnitéEnLettres = "cinq "
    Case "6": UnitéEnLettres = "six "
    Case "7": UnitéEnLettres = "sept "
    Case "8": UnitéEnLettres = "huit "
    Case "9": UnitéEnLettres = "neuf "
   End Select

End Function



