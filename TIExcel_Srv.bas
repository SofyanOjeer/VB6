Attribute VB_Name = "srvTIExcel"
Option Explicit

Public Const MemoCDPosPfLen = 150
Public Const MemoCDDosPfLen = 355
Public Const MemoCDPayPfLen = 96
Public Const MemoCDComEncPfLen = 134
Dim paramTIDB2_Output As String
Public paramAS400IN As String

Dim paramExcel_Exe As String, paramExcel_Dossier As String, paramExcel_Dossier_Référence As String
Dim paramExcel_Master As String, paramExcel_Posting As String, paramExcel_Posting2 As String
Dim paramExcel_PayDiff As String
Dim paramExcel_ComEnc As String

Dim appExcel As Excel.Application
Dim wbExcel As Excel.Workbook
Dim shtExcel As Excel.Worksheet
Dim rngExcel As Excel.Range
Dim I As Long, IMax As Long
Dim K As Long

Dim IdShell
Dim X As String, xIn As String, xIn_K As Integer
Dim DateAMJ As String
Dim Mnt As Currency


Dim X1 As String, X2 As String, X3 As String, X4 As String, X5 As String, X6 As String, X7 As String, X8 As String, X9 As String
Dim X10 As String, X11 As String, X12 As String, X13 As String, X14 As String, X15 As String, X16 As String, X17 As String, X18 As String, X19 As String
Dim X20 As String, X21 As String, X22 As String, X23 As String, X24 As String, X25 As String, X26 As String, X27 As String, X28 As String, X29 As String
Dim X30 As String, X31 As String, X32 As String, X33 As String, X34 As String, X35 As String, X36 As String, X37 As String, X38 As String, X39 As String

Public Sub ElpWait(lWait As Integer)
Application.Wait TimeSerial(Hour(Now()), Minute(Now()), Second(Now()) + lWait)
End Sub



Public Sub appExcel_Master_Export(lstMsg As ListBox)
On Error GoTo Error_Handle
lstMsg.AddItem "srvTIExcel.appExcel_Master_Export"

Set shtExcel = wbExcel.Sheets("Master")
AppActivate IdShell
DoEvents

''Set rngExcel = shtExcel.Columns(1)    ' Colonne KEY97

'''IMax = rngExcel.Find("").Row - 1

paramTIDB2_Output = paramAS400IN & "CDDOSW0"
Open paramTIDB2_Output For Output As #2

For I = 2 To 65000
    If I Mod 1000 = 0 Then Debug.Print I
    Set rngExcel = shtExcel.Rows(I)
    
    If Trim(rngExcel.Cells(1)) = "" Then Exit For
    X = Space(MemoCDDosPfLen)
    ' Références DOSSIER
    Mid$(X, 1, 12) = Format$(rngExcel.Cells(1), "000000000000")
    Mid$(X, 13, 3) = Format$(rngExcel.Cells(2), "   ")
    Mid$(X, 16, 6) = Format$(rngExcel.Cells(3), "000000")
    Mid$(X, 22, 20) = Format$(rngExcel.Cells(4), "                    ")
    ' Dates
    dateJma08_Amj08 rngExcel.Cells(7), DateAMJ
    Mid$(X, 42, 8) = DateAMJ
    dateJma08_Amj08 rngExcel.Cells(8), DateAMJ
    Mid$(X, 50, 8) = DateAMJ
    ' Codes
    Mid$(X, 58, 4) = Format$(rngExcel.Cells(9), "    ")
    Mid$(X, 62, 1) = " "
    Mid$(X, 63, 1) = Format$(rngExcel.Cells(19), " ")
    Mid$(X, 64, 1) = Format$(rngExcel.Cells(20), " ")
    Mid$(X, 65, 2) = Format$(rngExcel.Cells(10), "  ")
    Mid$(X, 67, 3) = Format$(rngExcel.Cells(11), "   ")
    Mid$(X, 70, 3) = Format$(rngExcel.Cells(12), "   ")
    Mid$(X, 73, 3) = Format$(rngExcel.Cells(13), "000")
    ' Pourcentage et montant garantie ou provision
    If IsNull(rngExcel.Cells(21)) Then
        Mid$(X, 76, 7) = "0000000"
    Else
        Mnt = rngExcel.Cells(21)
        Mid$(X, 76, 7) = Format$(Abs(Mnt) * 100, "0000000")
    End If
    Mnt = rngExcel.Cells(22)
    If rngExcel.Cells(23) = "ITL" Or rngExcel.Cells(23) = "GRD" Or rngExcel.Cells(23) = "PTE" _
       Or rngExcel.Cells(23) = "ESP" Or rngExcel.Cells(23) = "BEF" Or rngExcel.Cells(23) = "LUF" _
       Or rngExcel.Cells(23) = "JPY" Or rngExcel.Cells(23) = "CFA" Or rngExcel.Cells(23) = "XAF" _
       Or rngExcel.Cells(23) = "XOF" Then
       Mnt = curMaxD(Mnt, 0)
    Else
       Mnt = curMaxD(Mnt, 2)
    End If
    Mid$(X, 83, 17) = Format$(Abs(Mnt) * 100, "00000000000000000")
    ' Devise et compte garantie
    Mid$(X, 100, 3) = Format$(rngExcel.Cells(23), "   ")
    If IsNull(rngExcel.Cells(24)) Then
        Mid$(X, 103, 4) = "    "
    Else
        Mid$(X, 103, 4) = Format$(rngExcel.Cells(24), "    ")
    End If
    If IsNull(rngExcel.Cells(25)) Then
        Mid$(X, 107, 6) = "      "
    Else
        Mid$(X, 107, 6) = Format$(rngExcel.Cells(25), "      ")
    End If
    If IsNull(rngExcel.Cells(26)) Then
        Mid$(X, 113, 3) = "   "
    Else
        Mid$(X, 113, 3) = Format$(rngExcel.Cells(26), "   ")
    End If
    Mid$(X, 116, 25) = "                         "
    ' Montant dossier
    Mnt = rngExcel.Cells(14)
    If rngExcel.Cells(15) = "ITL" Or rngExcel.Cells(15) = "GRD" Or rngExcel.Cells(15) = "PTE" _
       Or rngExcel.Cells(15) = "ESP" Or rngExcel.Cells(15) = "BEF" Or rngExcel.Cells(15) = "LUF" _
       Or rngExcel.Cells(15) = "JPY" Or rngExcel.Cells(15) = "CFA" Or rngExcel.Cells(15) = "XAF" _
       Or rngExcel.Cells(15) = "XOF" Then
       Mnt = curMaxD(Mnt, 0)
    Else
       Mnt = curMaxD(Mnt, 2)
    End If
    Mid$(X, 141, 17) = Format$(Abs(Mnt) * 100, "00000000000000000")
    Mid$(X, 158, 3) = Format$(rngExcel.Cells(15), "   ")

    ' Pourcentages
    If IsNull(rngExcel.Cells(27)) Then
        Mid$(X, 161, 7) = "0000000"
    Else
        Mnt = rngExcel.Cells(27)
        Mid$(X, 161, 7) = Format$(Abs(Mnt) * 100, "0000000")
    End If
    If IsNull(rngExcel.Cells(28)) Then
        Mid$(X, 168, 7) = "0000000"
    Else
        Mnt = rngExcel.Cells(28)
        Mid$(X, 168, 7) = Format$(Abs(Mnt) * 100, "0000000")
    End If
    If IsNull(rngExcel.Cells(29)) Then
        Mid$(X, 175, 7) = "0000000"
    Else
        Mnt = rngExcel.Cells(29)
        Mid$(X, 175, 7) = Format$(Abs(Mnt) * 100, "0000000")
    End If
    Mid$(X, 182, 1) = Format$(rngExcel.Cells(30), " ")
    ' Outstanding et Liability
    Mnt = rngExcel.Cells(16)
    If rngExcel.Cells(15) = "ITL" Or rngExcel.Cells(15) = "GRD" Or rngExcel.Cells(15) = "PTE" _
       Or rngExcel.Cells(15) = "ESP" Or rngExcel.Cells(15) = "BEF" Or rngExcel.Cells(15) = "LUF" _
       Or rngExcel.Cells(15) = "JPY" Or rngExcel.Cells(15) = "CFA" Or rngExcel.Cells(15) = "XAF" _
       Or rngExcel.Cells(15) = "XOF" Then
       Mnt = curMaxD(Mnt, 0)
    Else
       Mnt = curMaxD(Mnt, 2)
    End If
    Mid$(X, 183, 17) = Format$(Abs(Mnt) * 100, "00000000000000000")
    Mnt = rngExcel.Cells(17)
    If rngExcel.Cells(18) = "ITL" Or rngExcel.Cells(18) = "GRD" Or rngExcel.Cells(18) = "PTE" _
       Or rngExcel.Cells(18) = "ESP" Or rngExcel.Cells(18) = "BEF" Or rngExcel.Cells(18) = "LUF" _
       Or rngExcel.Cells(18) = "JPY" Or rngExcel.Cells(18) = "CFA" Or rngExcel.Cells(18) = "XAF" _
       Or rngExcel.Cells(18) = "XOF" Then
       Mnt = curMaxD(Mnt, 0)
    Else
       Mnt = curMaxD(Mnt, 2)
    End If
    Mid$(X, 200, 17) = Format$(Abs(Mnt) * 100, "00000000000000000")
    Mid$(X, 217, 3) = Format$(rngExcel.Cells(18), "   ")
    ' Les différents intervenants du DOSSIER
    If IsNull(rngExcel.Cells(31)) Then
        Mid$(X, 220, 12) = "000000000000"
    Else
        Mid$(X, 220, 12) = Format$(rngExcel.Cells(31), "000000000000")
    End If
    Mid$(X, 232, 6) = "000000"
    If IsNull(rngExcel.Cells(32)) Then
        Mid$(X, 238, 12) = "000000000000"
    Else
        Mid$(X, 238, 12) = Format$(rngExcel.Cells(32), "000000000000")
    End If
    Mid$(X, 250, 6) = "000000"
    If IsNull(rngExcel.Cells(33)) Then
        Mid$(X, 256, 12) = "000000000000"
    Else
        Mid$(X, 256, 12) = Format$(rngExcel.Cells(33), "000000000000")
    End If
    Mid$(X, 268, 6) = "000000"
    If IsNull(rngExcel.Cells(34)) Then
        Mid$(X, 274, 12) = "000000000000"
    Else
        Mid$(X, 274, 12) = Format$(rngExcel.Cells(34), "000000000000")
    End If
    Mid$(X, 286, 6) = "000000"
    ' Infos de mise à jour DOSSIER
    Mid$(X, 292, 4) = Format$(rngExcel.Cells(5), "    ")
    Mid$(X, 296, 4) = Format$(rngExcel.Cells(6), "    ")
    Mid$(X, 300, 20) = "                    "
    Mid$(X, 320, 8) = "00000000"
    Mid$(X, 328, 8) = "00000000"
    Mid$(X, 336, 20) = "                    "
    
    Print #2, X
Next I

Close #2
Exit Sub

Error_Handle:
     lstMsg.AddItem "srvTIExcel.Monitor_SQL : " & Error
    Close

End Sub

Public Sub CSV_Master_Export(lstMsg As ListBox)
On Error GoTo Error_Handle


lstMsg.AddItem "srvTIExcel.CSV_Master_Export"

Open paramExcel_Master For Input As #1
Line Input #1, xIn

paramTIDB2_Output = paramAS400IN & "CDDOSW0"
Open paramTIDB2_Output For Output As #2
I = 0

Do Until EOF(1)
    DoEvents
    Line Input #1, xIn
    xIn = Trim(xIn) & ";"
    
    I = I + 1
    If I Mod 1000 = 0 Then Debug.Print I
    X = Space(MemoCDDosPfLen)
    ' Références DOSSIER
    xIn_K = 1
    
    X1 = CSV_X(xIn, xIn_K)
    X2 = CSV_X(xIn, xIn_K)
    X3 = CSV_X(xIn, xIn_K)
    X4 = CSV_X(xIn, xIn_K)
    X5 = CSV_X(xIn, xIn_K)
    X6 = CSV_X(xIn, xIn_K)
    X7 = CSV_X(xIn, xIn_K)
    X8 = CSV_X(xIn, xIn_K)
    X9 = CSV_X(xIn, xIn_K)
    X10 = CSV_X(xIn, xIn_K)
    X11 = CSV_X(xIn, xIn_K)
    X12 = CSV_X(xIn, xIn_K)
    X13 = CSV_X(xIn, xIn_K)
    X14 = CSV_X(xIn, xIn_K)
    X15 = CSV_X(xIn, xIn_K)
    X16 = CSV_X(xIn, xIn_K)
    X17 = CSV_X(xIn, xIn_K)
    X18 = CSV_X(xIn, xIn_K)
    X19 = CSV_X(xIn, xIn_K)
    X20 = CSV_X(xIn, xIn_K)
    X21 = CSV_X(xIn, xIn_K)
    X22 = CSV_X(xIn, xIn_K)
    X23 = CSV_X(xIn, xIn_K)
    X24 = CSV_X(xIn, xIn_K)
    X25 = CSV_X(xIn, xIn_K)
    X26 = CSV_X(xIn, xIn_K)
    X27 = CSV_X(xIn, xIn_K)
    X28 = CSV_X(xIn, xIn_K)
    X29 = CSV_X(xIn, xIn_K)
    X30 = CSV_X(xIn, xIn_K)
    X31 = CSV_X(xIn, xIn_K)
    X32 = CSV_X(xIn, xIn_K)
    X33 = CSV_X(xIn, xIn_K)
    X34 = CSV_X(xIn, xIn_K)
    
    ' Références DOSSIER
    Mid$(X, 1, 12) = Format$(X1, "000000000000")             '1
    Mid$(X, 13, 3) = Format$(X2, "   ")                      '2
    Mid$(X, 16, 6) = Format$(X3, "000000")                   '3
    Mid$(X, 22, 20) = Format$(X4, "                    ")    '4
    ' Dates
    dateJma08_Amj08 X7, DateAMJ                              '7
    Mid$(X, 42, 8) = DateAMJ
    dateJma08_Amj08 X8, DateAMJ                              '8
    Mid$(X, 50, 8) = DateAMJ
    ' Codes
    Mid$(X, 58, 4) = Format$(X9, "    ")                     '9
    Mid$(X, 62, 1) = " "
    Mid$(X, 63, 1) = Format$(X19, " ")                       '19
    Mid$(X, 64, 1) = Format$(X20, " ")
    Mid$(X, 65, 2) = Format$(X10, "  ")
    Mid$(X, 67, 3) = Format$(X11, "   ")
    Mid$(X, 70, 3) = Format$(X12, "   ")
    Mid$(X, 73, 3) = Format$(X13, "000")
    ' Pourcentage et montant garantie ou provision
    If X21 = "" Then
        Mid$(X, 76, 7) = "0000000"
    Else
        Mnt = Val(X21)
        Mid$(X, 76, 7) = Format$(Abs(Mnt) * 100, "0000000")
    End If
    Mnt = Val(X22)
    If X23 = "ITL" Or X23 = "GRD" Or X23 = "PTE" _
       Or X23 = "ESP" Or X23 = "BEF" Or X23 = "LUF" _
       Or X23 = "JPY" Or X23 = "CFA" Or X23 = "XAF" _
       Or X23 = "XOF" Then
       Mnt = curMaxD(Mnt, 0)
    Else
       Mnt = curMaxD(Mnt, 2)
    End If
    Mid$(X, 83, 17) = Format$(Abs(Mnt) * 100, "00000000000000000")
    ' Devise et compte garantie
    Mid$(X, 100, 3) = Format$(X23, "   ")
    If X24 = "" Then
        Mid$(X, 103, 4) = "    "
    Else
        Mid$(X, 103, 4) = Format$(X24, "    ")
    End If
    If X25 = "" Then
        Mid$(X, 107, 6) = "      "
    Else
        Mid$(X, 107, 6) = Format$(X25, "      ")
    End If
    If X26 = "" Then
        Mid$(X, 113, 3) = "   "
    Else
        Mid$(X, 113, 3) = Format$(X26, "   ")
    End If
    Mid$(X, 116, 25) = "                         "
    ' Montant dossier
    Mnt = Val(X14)
    If X15 = "ITL" Or X15 = "GRD" Or X15 = "PTE" _
       Or X15 = "ESP" Or X15 = "BEF" Or X15 = "LUF" _
       Or X15 = "JPY" Or X15 = "CFA" Or X15 = "XAF" _
       Or X15 = "XOF" Then
       Mnt = curMaxD(Mnt, 0)
    Else
       Mnt = curMaxD(Mnt, 2)
    End If
    Mid$(X, 141, 17) = Format$(Abs(Mnt) * 100, "00000000000000000")
    Mid$(X, 158, 3) = Format$(X15, "   ")

    ' Pourcentages
    If X27 = "" Then
        Mid$(X, 161, 7) = "0000000"
    Else
        Mnt = Val(X27)
        Mid$(X, 161, 7) = Format$(Abs(Mnt) * 100, "0000000")
    End If
    If X28 = "" Then
        Mid$(X, 168, 7) = "0000000"
    Else
        Mnt = Val(X28)
        Mid$(X, 168, 7) = Format$(Abs(Mnt) * 100, "0000000")
    End If
    If X29 = "" Then
        Mid$(X, 175, 7) = "0000000"
    Else
        Mnt = Val(X29)
        Mid$(X, 175, 7) = Format$(Abs(Mnt) * 100, "0000000")
    End If
    Mid$(X, 182, 1) = Format$(X30, " ")
    ' Outstanding et Liability
    Mnt = Val(X16)
    If X15 = "ITL" Or X15 = "GRD" Or X15 = "PTE" _
       Or X15 = "ESP" Or X15 = "BEF" Or X15 = "LUF" _
       Or X15 = "JPY" Or X15 = "CFA" Or X15 = "XAF" _
       Or X15 = "XOF" Then
       Mnt = curMaxD(Mnt, 0)
    Else
       Mnt = curMaxD(Mnt, 2)
    End If
    Mid$(X, 183, 17) = Format$(Abs(Mnt) * 100, "00000000000000000")
    Mnt = Val(X17)
    If X18 = "ITL" Or X18 = "GRD" Or X18 = "PTE" _
       Or X18 = "ESP" Or X18 = "BEF" Or X18 = "LUF" _
       Or X18 = "JPY" Or X18 = "CFA" Or X18 = "XAF" _
       Or X18 = "XOF" Then
       Mnt = curMaxD(Mnt, 0)
    Else
       Mnt = curMaxD(Mnt, 2)
    End If
    Mid$(X, 200, 17) = Format$(Abs(Mnt) * 100, "00000000000000000")
    Mid$(X, 217, 3) = Format$(X18, "   ")
    ' Les différents intervenants du DOSSIER
    If X31 = "" Then
        Mid$(X, 220, 12) = "000000000000"
    Else
        Mid$(X, 220, 12) = Format$(X31, "000000000000")
    End If
    Mid$(X, 232, 6) = "000000"
    If X32 = "" Then
        Mid$(X, 238, 12) = "000000000000"
    Else
        Mid$(X, 238, 12) = Format$(X32, "000000000000")
    End If
    Mid$(X, 250, 6) = "000000"
    If X33 = "" Then
        Mid$(X, 256, 12) = "000000000000"
    Else
        Mid$(X, 256, 12) = Format$(X33, "000000000000")
    End If
    Mid$(X, 268, 6) = "000000"
    If X34 = "" Then
        Mid$(X, 274, 12) = "000000000000"
    Else
        Mid$(X, 274, 12) = Format$(X34, "000000000000")
    End If
    Mid$(X, 286, 6) = "000000"
    ' Infos de mise à jour DOSSIER
    Mid$(X, 292, 4) = Format$(X5, "    ")
    Mid$(X, 296, 4) = Format$(X6, "    ")
    Mid$(X, 300, 20) = "                    "
    Mid$(X, 320, 8) = "00000000"
    Mid$(X, 328, 8) = "00000000"
    Mid$(X, 336, 20) = "                    "
  
    
    Print #2, X
Loop
lstMsg.AddItem "srvTIExcel.CSV_Master_Export : " & I
Close
Exit Sub

Error_Handle:
     lstMsg.AddItem "srvTIExcel.CSV_Master : " & I & " : " & Error
    Close

End Sub

' Extraction des commissions encaissées en devises du Trade Innovation

Public Sub CSV_ComEnc_Export(lstMsg As ListBox)
On Error GoTo Error_Handle

lstMsg.AddItem "srvTIExcel.CSV_ComEnc_Export"

Open paramExcel_ComEnc For Input As #1
Line Input #1, xIn

paramTIDB2_Output = paramAS400IN & "CDCPEW0"
Open paramTIDB2_Output For Output As #2
I = 0

Do Until EOF(1)
    DoEvents
    Line Input #1, xIn
    xIn = Trim(xIn) & ";"
    
    I = I + 1
    If I Mod 1000 = 0 Then Debug.Print I
    X = Space(MemoCDComEncPfLen)
   
    xIn_K = 1
    
    X1 = CSV_X(xIn, xIn_K)
    X2 = CSV_X(xIn, xIn_K)
    X3 = CSV_X(xIn, xIn_K)
    X4 = CSV_X(xIn, xIn_K)
    X5 = CSV_X(xIn, xIn_K)
    X6 = CSV_X(xIn, xIn_K)
    X7 = CSV_X(xIn, xIn_K)
    X8 = CSV_X(xIn, xIn_K)
    X9 = CSV_X(xIn, xIn_K)
    X10 = CSV_X(xIn, xIn_K)
    X11 = CSV_X(xIn, xIn_K)
    
    Mid$(X, 1, 12) = Format$(X1, "000000000000")
    Mid$(X, 13, 12) = "000000000000"
    Mid$(X, 25, 3) = "   "
    Mid$(X, 28, 6) = "000000"
    Mid$(X, 34, 12) = "000000000000"
    Mid$(X, 46, 3) = "   "
    Mid$(X, 49, 6) = "000000"
   ' Dates
    dateJma08_Amj08 X2, DateAMJ
    Mid$(X, 55, 8) = DateAMJ
    
    Mid$(X, 63, 3) = Format$(X3, "010")
   
   ' Devise
    Mid$(X, 83, 3) = Format$(X5, "   ")
   ' Montant
    Mnt = Val(X4)
    If X5 = "ITL" Or X5 = "GRD" Or X5 = "PTE" _
       Or X5 = "ESP" Or X5 = "BEF" Or X5 = "LUF" _
       Or X5 = "JPY" Or X5 = "CFA" Or X5 = "XAF" _
       Or X5 = "XOF" Then
       Mnt = curMaxD(Mnt, 0)
    Else
       Mnt = curMaxD(Mnt, 2)
    End If
    Mid$(X, 66, 17) = Format$(Abs(Mnt) * 100, "00000000000000000")
    Mid$(X, 86, 12) = Format$(X6, "000000000000")
    Mid$(X, 98, 10) = Format$(X7, "          ")
    Mid$(X, 108, 12) = Format$(X8, "000000000000")
    Mid$(X, 120, 12) = Format$(X9, "000000000000")
    Mid$(X, 132, 2) = Format$(X10, "  ")
    Mid$(X, 134, 1) = Format$(X11, " ")

    Print #2, X
Loop
lstMsg.AddItem "srvTIExcel.CSV_ComEnc_Export : " & I
Close
Exit Sub

Error_Handle:
     lstMsg.AddItem "srvTIExcel.CSV_ComEnc : " & I & " : " & Error
    Close

End Sub

Public Sub CSV_PayDiff_Export(lstMsg As ListBox)
On Error GoTo Error_Handle

lstMsg.AddItem "srvTIExcel.CSV_PayDiff_Export"

Open paramExcel_PayDiff For Input As #1
Line Input #1, xIn

paramTIDB2_Output = paramAS400IN & "CDPAYW0"
Open paramTIDB2_Output For Output As #2
I = 0

Do Until EOF(1)
    DoEvents
    Line Input #1, xIn
    xIn = Trim(xIn) & ";"
    
    I = I + 1
    If I Mod 1000 = 0 Then Debug.Print I
    X = Space(MemoCDPayPfLen)
   
    xIn_K = 1
    
    X1 = CSV_X(xIn, xIn_K)
    X2 = CSV_X(xIn, xIn_K)
    X3 = CSV_X(xIn, xIn_K)
    X4 = CSV_X(xIn, xIn_K)
    X5 = CSV_X(xIn, xIn_K)
    X6 = CSV_X(xIn, xIn_K)
    X7 = CSV_X(xIn, xIn_K)
    X8 = CSV_X(xIn, xIn_K)
    
    Mid$(X, 1, 12) = Format$(X3, "000000000000")
    Mid$(X, 13, 12) = Format$(X4, "000000000000")
    Mid$(X, 25, 3) = "   "
    Mid$(X, 28, 6) = "000000"
    Mid$(X, 34, 12) = "000000000000"
    Mid$(X, 46, 3) = "   "
    Mid$(X, 49, 6) = "000000"
   ' Dates
    dateJma08_Amj08 X5, DateAMJ
    Mid$(X, 55, 8) = DateAMJ
    dateJma08_Amj08 X6, DateAMJ
    Mid$(X, 63, 8) = DateAMJ
    
    Mid$(X, 71, 5) = Format$(X7, "00000")
    Mid$(X, 76, 1) = Format$(X8, " ")
    Mnt = Val(X1)
    If X2 = "ITL" Or X2 = "GRD" Or X2 = "PTE" _
       Or X2 = "ESP" Or X2 = "BEF" Or X2 = "LUF" _
       Or X2 = "JPY" Or X2 = "CFA" Or X2 = "XAF" _
       Or X2 = "XOF" Then
       Mnt = curMaxD(Mnt, 0)
    Else
       Mnt = curMaxD(Mnt, 2)
    End If
    Mid$(X, 77, 17) = Format$(Abs(Mnt) * 100, "00000000000000000")
    Mid$(X, 94, 3) = Format$(X2, "   ")
    
    Print #2, X
Loop
lstMsg.AddItem "srvTIExcel.CSV_PayDiff_Export : " & I
Close
Exit Sub

Error_Handle:
     lstMsg.AddItem "srvTIExcel.CSV_PayDiff : " & I & " : " & Error
    Close

End Sub


Public Sub CSV_Posting_Export(lstMsg As ListBox)
On Error GoTo Error_Handle



paramTIDB2_Output = paramAS400IN & "CDPOSW0"
Open paramTIDB2_Output For Output As #2


Open paramExcel_Posting For Input As #1
Line Input #1, xIn

CSV_Posting_Export_Sheet lstMsg

Close #1
Open paramExcel_Posting2 For Input As #1
Line Input #1, xIn

CSV_Posting_Export_Sheet lstMsg

Close
lstMsg.AddItem "srvTIExcel.CSV_Posting_Export : terminé"

Exit Sub

Error_Handle:
     lstMsg.AddItem "srvTIExcel.CSV_Posting : " & Error
    Close

End Sub

Public Sub CSV_Posting_Export_Sheet(lstMsg As ListBox)
On Error GoTo Error_Handle


lstMsg.AddItem "srvTIExcel.CSV_Posting_Export"

I = 0

Do Until EOF(1)
    DoEvents
    Line Input #1, xIn
    xIn = Trim(xIn) & ";"
    
    I = I + 1
    If I Mod 1000 = 0 Then Debug.Print I
    X = Space(MemoCDPosPfLen)
    If I Mod 1000 = 0 Then Debug.Print I
    
    xIn_K = 1
    X1 = CSV_X(xIn, xIn_K)
    X2 = CSV_X(xIn, xIn_K)
    X3 = CSV_X(xIn, xIn_K)
    X4 = CSV_X(xIn, xIn_K)
    X5 = CSV_X(xIn, xIn_K)
    X6 = CSV_X(xIn, xIn_K)
    X7 = CSV_X(xIn, xIn_K)
    X8 = CSV_X(xIn, xIn_K)
    X9 = CSV_X(xIn, xIn_K)
    X10 = CSV_X(xIn, xIn_K)
    X11 = CSV_X(xIn, xIn_K)
    X12 = CSV_X(xIn, xIn_K)
    X13 = CSV_X(xIn, xIn_K)
    X14 = CSV_X(xIn, xIn_K)
    X15 = CSV_X(xIn, xIn_K)
    X16 = CSV_X(xIn, xIn_K)
    X17 = CSV_X(xIn, xIn_K)

    Mid$(X, 1, 12) = Format$(X16, "000000000000")
    Mid$(X, 13, 12) = "000000000000"
    Mid$(X, 25, 3) = "   "
    Mid$(X, 28, 6) = "000000"
    Mid$(X, 34, 12) = "000000000000"
    Mid$(X, 46, 3) = "   "
    Mid$(X, 49, 6) = Format(Trim(X15), "000000")
    Mid$(X, 55, 4) = X12
  'Date
    dateJma08_Amj08 X5, DateAMJ
    Mid$(X, 59, 8) = DateAMJ
 
    If X1 = "" Then
        Mid$(X, 67, 4) = "    "
    Else
        Mid$(X, 67, 4) = Format$(X1, "    ")
    End If
    If X2 = "" Then
        Mid$(X, 71, 6) = "      "
    Else
        Mid$(X, 71, 6) = Format$(X2, "      ")
    End If
    If X3 = "" Then
        Mid$(X, 77, 3) = "   "
    Else
        Mid$(X, 77, 3) = Format$(X3, "   ")
    End If
    Mid$(X, 80, 25) = "                         "
    Mid$(X, 105, 3) = X4
    Mid$(X, 108, 1) = X6
 ' Devise
    Mid$(X, 126, 3) = X8
 
 ' Montant
    Mnt = Val(X7)
    If X8 = "ITL" Or X8 = "GRD" Or X8 = "PTE" _
       Or X8 = "ESP" Or X8 = "BEF" Or X8 = "LUF" _
       Or X8 = "JPY" Or X8 = "CFA" Or X8 = "XAF" _
       Or X8 = "XOF" Then
       Mnt = curMaxD(Mnt, 0)
    Else
       Mnt = curMaxD(Mnt, 2)
    End If
    Mid$(X, 109, 17) = Format$(Abs(Mnt) * 100, "00000000000000000")

    If X9 = "" Then
        Mid$(X, 129, 2) = "  "
    Else
        Mid$(X, 129, 2) = Format$(X9, "  ")
    End If
    If X10 = "" Then
        Mid$(X, 131, 6) = "      "
    Else
        Mid$(X, 131, 6) = Format$(X10, "      ")
    End If
    If X11 = "" Then
        Mid$(X, 137, 2) = "  "
    Else
        Mid$(X, 137, 2) = Format$(X11, "  ")
    End If
    If X17 = "" Then
        Mid$(X, 139, 12) = "000000000000"
    Else
        Mid$(X, 139, 12) = Format$(X17, "000000000000")
    End If

    Print #2, X
Loop
lstMsg.AddItem "srvTIExcel.CSV_Posting_Export : " & I

Exit Sub

Error_Handle:
     lstMsg.AddItem "srvTIExcel.CSV_Posting : " & I & " : " & Error
    Close

End Sub



Public Sub appExcel_Master_SQL(lstMsg As ListBox)
On Error GoTo Error_Handle
lstMsg.AddItem "srvTIExcel.appExcel_Monitor_Master_SQL"

Set shtExcel = wbExcel.Sheets("Master")

AppActivate IdShell
DoEvents
''SendKeys "^{PGUP}^{PGUP}^{PGUP}^{PGUP}", True        ' page précédente

SendKeys "%D+{F10}A", True     ' Lancement de la requête SQL

ElpWait 60
'
AppActivate IdShell
DoEvents
SendKeys "%F+{F10}g", True      ' Sauvegarde

Exit Sub

Error_Handle:
    lstMsg.AddItem "srvTIExcel.appExcel_Master_SQL : " & Error

End Sub


Public Sub appExcel_Posting_SQL(lstMsg As ListBox)
On Error GoTo Error_Handle
lstMsg.AddItem "srvTIExcel.appExcel_Monitor_Posting_SQL"

Set shtExcel = wbExcel.Sheets("Posting")

AppActivate IdShell
DoEvents
SendKeys "^{PGDN}", True        ' page suivante
SendKeys "%D+{F10}A", True     ' Lancement de la requête SQL

ElpWait 120
'
AppActivate IdShell
DoEvents
SendKeys "%F+{F10}g", True      ' Sauvegarde

Exit Sub

Error_Handle:
     lstMsg.AddItem "srvTIExcel.appExcel_Posting_SQL : " & Error

End Sub

Public Sub appExcel_Posting2_SQL(lstMsg As ListBox)
On Error GoTo Error_Handle
lstMsg.AddItem "srvTIExcel.appExcel_Monitor_Posting2_SQL"

Set shtExcel = wbExcel.Sheets("Posting2")

AppActivate IdShell
DoEvents
SendKeys "^{PGDN}", True        ' page suivante
SendKeys "%D+{F10}A", True     ' Lancement de la requête SQL

ElpWait 120
'
AppActivate IdShell
DoEvents
SendKeys "^{PGUP}^{PGUP}", True        ' page précédente
SendKeys "%F+{F10}g", True      ' Sauvegarde

Exit Sub

Error_Handle:
     lstMsg.AddItem "srvTIExcel.appExcel_Posting2_SQL : " & Error

End Sub


Public Sub appExcel_Posting_Export(lstMsg As ListBox)

On Error GoTo Error_Handle

paramTIDB2_Output = paramAS400IN & "CDPOSW0"
Open paramTIDB2_Output For Output As #2

lstMsg.AddItem "srvTIExcel.appExcel_Posting_Export 1/2"
Set shtExcel = wbExcel.Sheets("Posting")

AppActivate IdShell
DoEvents
SendKeys "^{PGDN}", True        ' page suivante

srvTIExcel.appExcel_Posting_Export_Sheet lstMsg

lstMsg.AddItem "srvTIExcel.appExcel_Posting_Export 2/2"
Set shtExcel = wbExcel.Sheets("Posting2")

AppActivate IdShell
DoEvents
SendKeys "^{PGDN}", True        ' page suivante
srvTIExcel.appExcel_Posting_Export_Sheet lstMsg

Close #2
Exit Sub

Error_Handle:
     lstMsg.AddItem "srvTIExcel.appExcel_Posting_Export : " & I & " : " & Error
    Close

End Sub


Public Sub appExcel_Posting_Export_Sheet(lstMsg As ListBox)

On Error GoTo Error_Handle
For I = 2 To 65000
    Set rngExcel = shtExcel.Rows(I)
    
        If Trim(rngExcel.Cells(16)) = "" Then Exit For

    X = Space(MemoCDPosPfLen)
    If I Mod 1000 = 0 Then Debug.Print I
    
    Mid$(X, 1, 12) = Format$(rngExcel.Cells(16), "000000000000")
    Mid$(X, 13, 12) = "000000000000"
    Mid$(X, 25, 3) = "   "
    Mid$(X, 28, 6) = "000000"
    Mid$(X, 34, 12) = "000000000000"
    Mid$(X, 46, 3) = "   "
    Mid$(X, 49, 6) = Format(Trim(rngExcel.Cells(15)), "000000")
    Mid$(X, 55, 4) = rngExcel.Cells(12)
  'Date
    dateJma08_Amj08 rngExcel.Cells(5), DateAMJ
    Mid$(X, 59, 8) = DateAMJ
 
    If IsNull(rngExcel.Cells(1)) Then
        Mid$(X, 67, 4) = "    "
    Else
        Mid$(X, 67, 4) = Format$(rngExcel.Cells(1), "    ")
    End If
    If IsNull(rngExcel.Cells(2)) Then
        Mid$(X, 71, 6) = "      "
    Else
        Mid$(X, 71, 6) = Format$(rngExcel.Cells(2), "      ")
    End If
    If IsNull(rngExcel.Cells(3)) Then
        Mid$(X, 77, 3) = "   "
    Else
        Mid$(X, 77, 3) = Format$(rngExcel.Cells(3), "   ")
    End If
    Mid$(X, 80, 25) = "                         "
    Mid$(X, 105, 3) = rngExcel.Cells(4)
    Mid$(X, 108, 1) = rngExcel.Cells(6)
 ' Devise
    Mid$(X, 126, 3) = rngExcel.Cells(8)
 
 ' Montant
    Mnt = rngExcel.Cells(7)
    If rngExcel.Cells(8) = "ITL" Or rngExcel.Cells(8) = "GRD" Or rngExcel.Cells(8) = "PTE" _
       Or rngExcel.Cells(8) = "ESP" Or rngExcel.Cells(8) = "BEF" Or rngExcel.Cells(8) = "LUF" _
       Or rngExcel.Cells(8) = "JPY" Or rngExcel.Cells(8) = "CFA" Or rngExcel.Cells(8) = "XAF" _
       Or rngExcel.Cells(8) = "XOF" Then
       Mnt = curMaxD(Mnt, 0)
    Else
       Mnt = curMaxD(Mnt, 2)
    End If
    Mid$(X, 109, 17) = Format$(Abs(Mnt) * 100, "00000000000000000")

    If IsNull(rngExcel.Cells(9)) Then
        Mid$(X, 129, 2) = "  "
    Else
        Mid$(X, 129, 2) = Format$(rngExcel.Cells(9), "  ")
    End If
    If IsNull(rngExcel.Cells(10)) Then
        Mid$(X, 131, 6) = "      "
    Else
        Mid$(X, 131, 6) = Format$(rngExcel.Cells(10), "      ")
    End If
    If IsNull(rngExcel.Cells(11)) Then
        Mid$(X, 137, 2) = "  "
    Else
        Mid$(X, 137, 2) = Format$(rngExcel.Cells(11), "  ")
    End If
    If IsNull(rngExcel.Cells(17)) Then
        Mid$(X, 139, 12) = "000000000000"
    Else
        Mid$(X, 139, 12) = Format$(rngExcel.Cells(17), "000000000000")
    End If

    Print #2, X
Next I

Exit Sub

Error_Handle:
     lstMsg.AddItem "srvTIExcel.appExcel_Posting_Export_Sheet : " & I & " : " & Error
    Close

End Sub

Public Sub param_Init(lstMsg As ListBox)

Set appExcel = Nothing
Set wbExcel = Nothing

paramExcel_Exe = "D:\Program Files\Microsoft Office\Office\Excel.exe "
paramExcel_Dossier = "D:\Temp\TI.xls"
paramExcel_Dossier_Référence = "\\FR11024427\.BiaSrc\Dta\TI_Sql.xls"
paramExcel_Master = "D:\Temp\TI_Master.csv"
paramExcel_Posting = "D:\Temp\TI_Posting.csv"
paramExcel_Posting2 = "D:\Temp\TI_Posting2.csv"
paramExcel_PayDiff = "D:\Temp\TI_PayDiff.csv"
paramExcel_ComEnc = "D:\Temp\TI_ComEnc.csv"

lstMsg.AddItem "D:\Program Files\Microsoft Office\Office\Excel.exe "
lstMsg.AddItem "\\FR11024427\.BiaSrc\Dta\TI_Sql.xls"
lstMsg.AddItem "D:\Temp\TI.xls"

End Sub

Public Sub appExcel_Exe(lWindow As Integer)

X = paramExcel_Exe & " " & Chr$(34) & paramExcel_Dossier & Chr$(34)
IdShell = Shell(X, lWindow)
AppActivate IdShell
DoEvents

End Sub

Public Sub appExcel_Monitor(lstMsg As ListBox)
lstMsg.AddItem "Master : début": DoEvents

On Error GoTo Err_Msg
param_Init lstMsg

appExcel_Exe 1

ElpWait 5

Set appExcel = GetObject(, "Excel.application")
Set wbExcel = appExcel.Workbooks.Open(paramExcel_Dossier)


appExcel_Master_Export lstMsg


appExcel_Posting_Export lstMsg


AppActivate IdShell
DoEvents
appExcel.Application.DisplayAlerts = False
appExcel.Workbooks.Close

ElpWait 2

SendKeys "%{F4}", True
lstMsg.AddItem "srvTIExcel.Monitor_OK"
Exit Sub

Err_Msg:
    On Error Resume Next
     lstMsg.AddItem "srvTIExcel.appExcel_Monitor : " & Error

End Sub

Public Sub CSV_Monitor(lstMsg As ListBox)
lstMsg.AddItem "srvTIExcel.CSV_Monitor : début": DoEvents

On Error GoTo Err_Msg
param_Init lstMsg

CSV_Master_Export lstMsg
CSV_Posting_Export lstMsg
CSV_PayDiff_Export lstMsg
CSV_ComEnc_Export lstMsg

lstMsg.AddItem "srvTIExcel.CSV_Monitor : OK"
Exit Sub

Err_Msg:
    On Error Resume Next
     lstMsg.AddItem "srvTIExcel.CSV_Monitor : " & Error

End Sub

Public Sub appExcel_Monitor_SQL(lstMsg As ListBox)

lstMsg.AddItem "srvTIExcel.appExcel_Monitor_SQL": DoEvents

On Error GoTo Err_Msg

param_Init lstMsg

X = Dir(paramExcel_Dossier)
If X <> "" Then Kill paramExcel_Dossier
msFileSystem.CopyFile paramExcel_Dossier_Référence, paramExcel_Dossier
Set msFile = msFileSystem.GetFile(paramExcel_Dossier)
msFile.Attributes = 0

appExcel_Exe 1

ElpWait 5

Set appExcel = GetObject(, "Excel.application")
Set wbExcel = appExcel.Workbooks.Open(paramExcel_Dossier)


appExcel_Master_SQL lstMsg


appExcel_Posting_SQL lstMsg
appExcel_Posting2_SQL lstMsg


AppActivate IdShell
DoEvents
appExcel.Application.DisplayAlerts = False
appExcel.Workbooks.Close

ElpWait 2

SendKeys "%{F4}", True

Exit Sub

Err_Msg:
    lstMsg.AddItem "srvTIExcel.appExcel_Monitor_SQL : " & Error
'    appExcel.Application.Quit

End Sub



Public Function CSV_X(xIn As String, xIn_K As Integer) As String
Dim K As Integer
K = InStr(xIn_K, xIn, ";")
If K <= 0 Then
    CSV_X = ""
Else
    CSV_X = mId$(xIn, xIn_K, K - xIn_K)
    xIn_K = K + 1
End If

End Function
