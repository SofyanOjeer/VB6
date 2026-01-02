Attribute VB_Name = "Module2"
Public Function Text_KeyWord(lText As String, lK As Integer) As String
Dim Kmin As Integer, Kmax As Integer, lenText As Integer, xKeyWord As String, blnOK As Boolean
Dim X1 As String, blnKeyWord As Boolean

lenText = Len(lText)
blnKeyWord = False
Do
    Kmin = lK + 1
    xKeyWord = ""
    blnOK = False
    
    For Kmax = Kmin To lenText
        X1 = mId$(lText, Kmax, 1)
        Select Case X1
            Case ".", "-", "_":
            Case "a" To "z": xKeyWord = xKeyWord & X1: blnOK = True
            Case "0" To "9": xKeyWord = xKeyWord & X1: blnOK = True
            Case Else: If blnOK Then Exit For
        End Select
                
    Next Kmax
    
    If Kmax >= lenText Then blnKeyWord = True
    lK = Kmax
    
    Select Case xKeyWord
        Case "l", "le", "la", "les", "du", "de", "des", "a", "au", "et", "ou":
        Case Else: blnKeyWord = True
    End Select
Loop Until blnKeyWord

Text_KeyWord = xKeyWord
End Function

Public Function Text_KeyWord(lText As String, lK As Integer) As String
Dim Kmin As Integer, Kmax As Integer, lenText As Integer, xKeyWord As String, blnOK As Boolean
Dim X1 As String, blnKeyWord As Boolean

lenText = Len(lText)
blnKeyWord = False
Do
    Kmin = lK + 1
    xKeyWord = ""
    blnOK = False
    
    For Kmax = Kmin To lenText
        X1 = mId$(lText, Kmax, 1)
        Select Case X1
            Case ".", "-", "_":
            Case "a" To "z": xKeyWord = xKeyWord & X1: blnOK = True
            Case "0" To "9": xKeyWord = xKeyWord & X1: blnOK = True
            Case Else: If blnOK Then Exit For
        End Select
                
    Next Kmax
    
    If Kmax >= lenText Then blnKeyWord = True
    lK = Kmax
    
    Select Case xKeyWord
        Case "l", "le", "la", "les", "du", "de", "des", "a", "au", "et", "ou":
        Case Else: blnKeyWord = True
    End Select
Loop Until blnKeyWord

Text_KeyWord = xKeyWord
End Function

Public Function Text_LCase(lText As String) As String
Dim X As String, I As Integer

X = LCase(Trim(lText))

For I = 1 To Len(X)
    Select Case mId$(X, I, 1)
        Case "à", "â", "ä": Mid$(X, I, 1) = "a"
        Case "é", "è", "ê", "ë": Mid$(X, I, 1) = "e"
        Case "î", "ï": Mid$(X, I, 1) = "i"
        Case "ô", "ö": Mid$(X, I, 1) = "o"
        Case "ù", "û", "ü": Mid$(X, I, 1) = "u"
        Case "ç": Mid$(X, I, 1) = "c"

   End Select
Next I
Text_LCase = X
End Function

