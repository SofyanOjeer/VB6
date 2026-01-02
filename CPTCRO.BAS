Attribute VB_Name = "srvCptCRO"
Option Explicit

Type typeCommission
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    Devise                  As String * 3
    Brut                    As String * 16
    optImputation           As String * 1
    optBarème               As String * 1
    chk1                    As String * 1
    chk2                    As String * 1
    chk3                    As String * 1
    chk4                    As String * 1
    chk5                    As String * 1
    filler1                 As String * 9
    MontantDevise           As String * 12
    MontantFRF              As String * 12
    MontantTVA              As String * 8
    filler2                 As String * 3
    Montant1                As String * 8
    Montant3                As String * 8
    Montant4                As String * 8
    Montant5                As String * 8
    filler3                 As String * 3
End Type

    


'---------------------------------------------------------
Public Sub recCommission_Init(recCommission As typeCommission)
'---------------------------------------------------------
 MsgTxt = String$(139, "0")
 MsgTxtIndex = 0
 Call srvCommission_GetBuffer(recCommission)
 recCommission.obj = "SRVCOMMISSIO"
End Sub

'---------------------------------------------------------
Public Sub srvCommission_GetBuffer(recCommission As typeCommission)
'---------------------------------------------------------
Dim K As Integer
MsgTxtIndex = 0
recCommission.obj = Mid$(MsgTxt, MsgTxtIndex + 1, 12)
recCommission.Method = Mid$(MsgTxt, MsgTxtIndex + 13, 12)
recCommission.Err = Mid$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

recCommission.Devise = Mid$(MsgTxt, K + 1, 3)
recCommission.Brut = Mid$(MsgTxt, K + 4, 16)
recCommission.optImputation = Mid$(MsgTxt, K + 20, 1)
Mid$(MsgTxt, K + 21, 1) = recCommission.optBarème
recCommission.chk1 = Mid$(MsgTxt, K + 22, 1)
recCommission.chk2 = Mid$(MsgTxt, K + 23, 1)
recCommission.chk3 = Mid$(MsgTxt, K + 24, 1)
recCommission.chk4 = Mid$(MsgTxt, K + 25, 1)
recCommission.chk5 = Mid$(MsgTxt, K + 26, 1)
'recCommission.filler1=Mid$(MsgTxt, K + 27, 9)
recCommission.MontantDevise = Mid$(MsgTxt, K + 36, 12)
recCommission.MontantFRF = Mid$(MsgTxt, K + 48, 12)
recCommission.MontantTVA = Mid$(MsgTxt, K + 60, 8)
'recCommission.filler2=Mid$(MsgTxt, K + 68, 3)
recCommission.Montant1 = Mid$(MsgTxt, K + 71, 8)
recCommission.Montant3 = Mid$(MsgTxt, K + 79, 8)
recCommission.Montant4 = Mid$(MsgTxt, K + 87, 8)
recCommission.Montant5 = Mid$(MsgTxt, K + 95, 8)
'recCommission.filler3=Mid$(MsgTxt, K + 103, 3)

End Sub

'---------------------------------------------------------
Public Sub srvCommission_PutBuffer(recCommission As typeCommission)
'---------------------------------------------------------
Dim K As Integer
MsgTxtLen = 0
Mid$(MsgTxt, MsgTxtLen + 1, 12) = recCommission.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recCommission.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34

Mid$(MsgTxt, K + 1, 3) = recCommission.Devise
Mid$(MsgTxt, K + 4, 16) = recCommission.Brut
Mid$(MsgTxt, K + 20, 1) = recCommission.optImputation
Mid$(MsgTxt, K + 21, 1) = recCommission.optBarème
Mid$(MsgTxt, K + 22, 1) = recCommission.chk1
Mid$(MsgTxt, K + 23, 1) = recCommission.chk2
Mid$(MsgTxt, K + 24, 1) = recCommission.chk3
Mid$(MsgTxt, K + 25, 1) = recCommission.chk4
Mid$(MsgTxt, K + 26, 1) = recCommission.chk5
'Mid$(MsgTxt, K + 27, 9) = recCommission.filler1
Mid$(MsgTxt, K + 36, 12) = recCommission.MontantDevise
Mid$(MsgTxt, K + 48, 12) = recCommission.MontantFRF
Mid$(MsgTxt, K + 60, 8) = recCommission.MontantTVA
'Mid$(MsgTxt, K + 68, 3) = recCommission.filler2
Mid$(MsgTxt, K + 71, 8) = recCommission.Montant1
Mid$(MsgTxt, K + 79, 8) = recCommission.Montant3
Mid$(MsgTxt, K + 87, 8) = recCommission.Montant4
Mid$(MsgTxt, K + 95, 8) = recCommission.Montant5
'Mid$(MsgTxt, K + 103, 3) = recCommission.filler3
End Sub



