Attribute VB_Name = "srvOpCpt"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public Const recOpCptLen = 285 ' 34 + 251

Type typeOpCpt
    obj                     As String * 12
    Method                  As String * 12
    Err                     As String * 10
    
    Référence              As String * 6
    Séquence               As Integer
    CodeOpération          As String * 10
    Société                As String * 3
    Agence                 As String * 3
    Devise                 As String * 3
    Compte                 As String * 11
    Brut                   As Currency
    Sens                   As String * 1
    AmjOpération           As String * 8
    AmjValeur              As String * 8
    Libellé                As String * 50
    chkCompte              As String * 1
    chkSolde               As String * 1
    chkAmjOpération        As String * 1
    chkAmjValeur           As String * 1
    optAvis                As String * 1
    optVirement            As String * 1
    optSwift               As String * 1
    
    optComShare            As String * 1
    optComImputation       As String * 1
    optComBarème           As String * 1
    ComPays                As String * 3
    CoursDevise            As Double
    ComMontantDevise       As Currency
    ComMontantFRF          As Currency
    chkCom1               As String * 1
    ComMontant1           As Currency
    chkCom2               As String * 1
    ComMontant2           As Currency
    chkCom3               As String * 1
    ComMontant3           As Currency
    chkCom4               As String * 1
    ComMontant4           As Currency
    chkCom5               As String * 1
    ComMontant5           As Currency
    chkCom6               As String * 1
    ComMontant6           As Currency
    optAvisLangue         As String * 1
End Type
    
Public arrOpCpt() As typeOpCpt
Public arrOpCptNb As Integer
Public arrOpCptNbMax As Integer
Public arrOpCptIndex As Integer
Public arrOpCptSuite As Boolean

Public Sub recOpCpt_Display(XpicBox As PictureBox, recOpCpt As typeOpCpt)
Dim X As String
XpicBox.ForeColor = warnUsrColor
XpicBox.CurrentX = 50
XpicBox.Print recOpCpt.Référence;
XpicBox.ForeColor = libUsr.ForeColor
XpicBox.CurrentX = 800
XpicBox.Print recOpCpt.Devise & ".";
XpicBox.Print Compte_Imp(recOpCpt.Compte);
XpicBox.ForeColor = IIf(recOpCpt.Sens = "D", errUsr.ForeColor, libUsr.ForeColor)
X = num_Display(recOpCpt.Brut, 15, 2, Lx, X, "0")
XpicBox.CurrentX = 3500 - XpicBox.TextWidth(X)
XpicBox.Print X;
XpicBox.Print " " & recOpCpt.Sens;
XpicBox.ForeColor = warnUsrColor
XpicBox.Print "   " & dateImp(recOpCpt.AmjOpération);
XpicBox.ForeColor = libUsr.ForeColor
XpicBox.Print "   " & recOpCpt.Libellé;

End Sub
  

'-----------------------------------------------------
Public Function srvOpCpt_Find(recOpCpt As typeOpCpt)
'-----------------------------------------------------
Dim I As Integer, X As Variant

X = "?"

For I = 1 To arrOpCptNb
    If recOpCpt.Référence = arrOpCpt(I).Référence Then
        X = Null
        Exit For
'!!!!!!!========
    End If
Next I

If Not IsNull(X) Then
    I = arrOpCptNb + 1
    recOpCpt.Method = "SnapP0      "
    arrOpCpt(0) = recOpCpt
    X = srvOpCpt_Snap(recOpCpt)
    If I > arrOpCptNb Then X = "inconnu"
End If

If IsNull(X) Then
    arrOpCptIndex = I
    recOpCpt = arrOpCpt(I)
End If

srvOpCpt_Find = X

End Function

'-----------------------------------------------------
Function srvOpCpt_Update(recOpCpt As typeOpCpt)
'-----------------------------------------------------

srvOpCpt_Update = "?"

MsgTxtLen = 0
Call srvOpCpt_PutBuffer(recOpCpt)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    recOpCpt.Err = Mid$(MsgTxt, MsgTxtIndex + 25, 10)
    If Trim(recOpCpt.Err) <> "" Then
        Call srvOpCpt_Error(recOpCpt)
        srvOpCpt_Update = recOpCpt.Err
        Exit Function
    Else
        srvOpCpt_Update = Null
    End If
Else
    recOpCpt.Err = "srv"
End If


'=====================================================
End Function

'-----------------------------------------------------
Public Function srvOpCpt_Monitor(recOpCpt As typeOpCpt)
'-----------------------------------------------------

arrOpCptSuite = False
Select Case Mid$(Trim(recOpCpt.Method), 1, 4)
    Case "Seek", "NUME"
                srvOpCpt_Monitor = srvOpCpt_Seek(recOpCpt)
    Case "Snap", "Prev"
              srvOpCpt_Monitor = srvOpCpt_Snap(recOpCpt)
    Case Else
                recOpCpt.Err = recOpCpt.Method
                Call srvOpCpt_Error(recOpCpt)
                srvOpCpt_Monitor = recOpCpt.Err
End Select

End Function

'-----------------------------------------------------
Sub srvOpCpt_Error(recOpCpt As typeOpCpt)
'-----------------------------------------------------

Dim Msg, Title
Dim I As Integer

Msg = "Opération de transfert (comptabilité): " ' & Chr$(10) & Chr$(13)

Select Case Mid$(recOpCpt.Err, 9, 2)
    Case "22"   '9922
        Msg = Msg & "Existe déjà"
        I = vbExclamation
    Case "23"   '9923
        Msg = Msg & "N'existe pas"
        I = vbExclamation
    Case Else
        Msg = Msg & "Error Code : " & recOpCpt.Err
        I = vbCritical
End Select

MsgBox Msg, I, "module : srvOpCpt_.bas  ( " _
                & Trim(recOpCpt.obj) & " : " & Trim(recOpCpt.Method) & " )"

End Sub



'---------------------------------------------------------
Public Function srvOpCpt_GetBuffer(recOpCpt As typeOpCpt)
'---------------------------------------------------------
Dim K As Integer, I As Integer
srvOpCpt_GetBuffer = Null
recOpCpt.obj = Mid$(MsgTxt, MsgTxtIndex + 1, 12)
recOpCpt.Method = Mid$(MsgTxt, MsgTxtIndex + 13, 12)
recOpCpt.Err = Mid$(MsgTxt, MsgTxtIndex + 25, 10)
K = MsgTxtIndex + 34

If recOpCpt.Err = Space$(10) Then
    recOpCpt.Référence = Mid$(MsgTxt, K + 1, 6)
    recOpCpt.Séquence = CInt(Val(Mid$(MsgTxt, K + 7, 3)))
    recOpCpt.CodeOpération = Mid$(MsgTxt, K + 10, 10)
    recOpCpt.Société = Mid$(MsgTxt, K + 20, 3)
    recOpCpt.Agence = Mid$(MsgTxt, K + 23, 3)
    recOpCpt.Devise = Mid$(MsgTxt, K + 26, 3)
    recOpCpt.Compte = Mid$(MsgTxt, K + 29, 11)
    recOpCpt.Brut = CCur(Val(Mid$(MsgTxt, K + 40, 17)) / 100)
    recOpCpt.Sens = Mid$(MsgTxt, K + 57, 1)
    recOpCpt.AmjOpération = Mid$(MsgTxt, K + 58, 8)
    recOpCpt.AmjValeur = Mid$(MsgTxt, K + 66, 8)
    recOpCpt.Libellé = Mid$(MsgTxt, K + 74, 50)
    recOpCpt.chkCompte = Mid$(MsgTxt, K + 124, 1)
    recOpCpt.chkSolde = Mid$(MsgTxt, K + 125, 1)
    recOpCpt.chkAmjOpération = Mid$(MsgTxt, K + 126, 1)
    recOpCpt.chkAmjValeur = Mid$(MsgTxt, K + 127, 1)
    recOpCpt.optAvis = Mid$(MsgTxt, K + 128, 1)
    recOpCpt.optVirement = Mid$(MsgTxt, K + 129, 1)
    recOpCpt.optSwift = Mid$(MsgTxt, K + 130, 1)
    recOpCpt.optComShare = Mid$(MsgTxt, K + 131, 1)
    recOpCpt.optComImputation = Mid$(MsgTxt, K + 132, 1)
    recOpCpt.optComBarème = Mid$(MsgTxt, K + 133, 1)
    recOpCpt.ComPays = Mid$(MsgTxt, K + 134, 3)
    recOpCpt.CoursDevise = CDbl(Val(Mid$(MsgTxt, K + 137, 12)) / 10000000)
    recOpCpt.ComMontantDevise = CCur(Val(Mid$(MsgTxt, K + 149, 12)) / 100)
    recOpCpt.ComMontantFRF = CCur(Val(Mid$(MsgTxt, K + 161, 12)) / 100)
    recOpCpt.chkCom1 = Mid$(MsgTxt, K + 173, 1)
    recOpCpt.ComMontant1 = CCur(Val(Mid$(MsgTxt, K + 174, 12)) / 100)
    recOpCpt.chkCom2 = Mid$(MsgTxt, K + 186, 1)
    recOpCpt.ComMontant2 = CCur(Val(Mid$(MsgTxt, K + 187, 12)) / 100)
    recOpCpt.chkCom3 = Mid$(MsgTxt, K + 199, 1)
    recOpCpt.ComMontant3 = CCur(Val(Mid$(MsgTxt, K + 200, 12)) / 100)
    recOpCpt.chkCom4 = Mid$(MsgTxt, K + 212, 1)
    recOpCpt.ComMontant4 = CCur(Val(Mid$(MsgTxt, K + 213, 12)) / 100)
    recOpCpt.chkCom5 = Mid$(MsgTxt, K + 225, 1)
    recOpCpt.ComMontant5 = CCur(Val(Mid$(MsgTxt, K + 226, 12)) / 100)
    recOpCpt.chkCom6 = Mid$(MsgTxt, K + 238, 1)
    recOpCpt.ComMontant6 = CCur(Val(Mid$(MsgTxt, K + 239, 12)) / 100)
    recOpCpt.optAvisLangue = Mid$(MsgTxt, K + 250, 1)
Else
    srvOpCpt_GetBuffer = recOpCpt.Err
End If

MsgTxtIndex = MsgTxtIndex + recOpCptLen

End Function

'---------------------------------------------------------
Public Sub srvOpCpt_PutBuffer(recOpCpt As typeOpCpt)
'---------------------------------------------------------
Dim K As Integer, I As Integer

Mid$(MsgTxt, MsgTxtLen + 1, 12) = recOpCpt.obj
Mid$(MsgTxt, MsgTxtLen + 13, 12) = recOpCpt.Method
Mid$(MsgTxt, MsgTxtLen + 25, 10) = Space$(10)
K = MsgTxtLen + 34

Mid$(MsgTxt, K + 1, 6) = recOpCpt.Référence
Mid$(MsgTxt, K + 7, 3) = Format$(recOpCpt.Séquence, "000")
Mid$(MsgTxt, K + 10, 10) = recOpCpt.CodeOpération
Mid$(MsgTxt, K + 20, 3) = recOpCpt.Société
Mid$(MsgTxt, K + 23, 3) = recOpCpt.Agence
Mid$(MsgTxt, K + 26, 3) = recOpCpt.Devise
Mid$(MsgTxt, K + 29, 11) = recOpCpt.Compte
Mid$(MsgTxt, K + 40, 17) = Format$(recOpCpt.Brut * 100, "00000000000000000")
Mid$(MsgTxt, K + 57, 1) = recOpCpt.Sens
Mid$(MsgTxt, K + 58, 8) = recOpCpt.AmjOpération
Mid$(MsgTxt, K + 66, 8) = recOpCpt.AmjValeur
Mid$(MsgTxt, K + 74, 50) = recOpCpt.Libellé
Mid$(MsgTxt, K + 124, 1) = recOpCpt.chkCompte
Mid$(MsgTxt, K + 125, 1) = recOpCpt.chkSolde
Mid$(MsgTxt, K + 126, 1) = recOpCpt.chkAmjOpération
Mid$(MsgTxt, K + 127, 1) = recOpCpt.chkAmjValeur
Mid$(MsgTxt, K + 128, 1) = recOpCpt.optAvis
Mid$(MsgTxt, K + 129, 1) = recOpCpt.optVirement
Mid$(MsgTxt, K + 130, 1) = recOpCpt.optSwift
Mid$(MsgTxt, K + 131, 1) = recOpCpt.optComShare
Mid$(MsgTxt, K + 132, 1) = recOpCpt.optComImputation
Mid$(MsgTxt, K + 133, 1) = recOpCpt.optComBarème
Mid$(MsgTxt, K + 134, 3) = recOpCpt.ComPays
Mid$(MsgTxt, K + 137, 12) = Format$(recOpCpt.CoursDevise * 10000000, "000000000000")
Mid$(MsgTxt, K + 149, 12) = Format$(recOpCpt.ComMontantDevise * 100, "000000000000")
Mid$(MsgTxt, K + 161, 12) = Format$(recOpCpt.ComMontantFRF * 100, "000000000000")
Mid$(MsgTxt, K + 173, 1) = recOpCpt.chkCom1
Mid$(MsgTxt, K + 174, 12) = Format$(recOpCpt.ComMontant1 * 100, "000000000000")
Mid$(MsgTxt, K + 186, 1) = recOpCpt.chkCom2
Mid$(MsgTxt, K + 187, 12) = Format$(recOpCpt.ComMontant2 * 100, "000000000000")
Mid$(MsgTxt, K + 199, 1) = recOpCpt.chkCom3
Mid$(MsgTxt, K + 200, 12) = Format$(recOpCpt.ComMontant3 * 100, "000000000000")
Mid$(MsgTxt, K + 212, 1) = recOpCpt.chkCom4
Mid$(MsgTxt, K + 213, 12) = Format$(recOpCpt.ComMontant4 * 100, "000000000000")
Mid$(MsgTxt, K + 225, 1) = recOpCpt.chkCom5
Mid$(MsgTxt, K + 226, 12) = Format$(recOpCpt.ComMontant5 * 100, "000000000000")
Mid$(MsgTxt, K + 238, 1) = recOpCpt.chkCom6
Mid$(MsgTxt, K + 239, 12) = Format$(recOpCpt.ComMontant6 * 100, "000000000000")
Mid$(MsgTxt, K + 250, 1) = recOpCpt.optAvisLangue

MsgTxtLen = MsgTxtLen + recOpCptLen
End Sub



'---------------------------------------------------------
Private Function srvOpCpt_Seek(recOpCpt As typeOpCpt)
'---------------------------------------------------------

srvOpCpt_Seek = "?"
MsgTxtLen = 0
Call srvOpCpt_PutBuffer(recOpCpt)
If IsNull(SndRcv()) Then
    MsgTxtIndex = 0
    If IsNull(srvOpCpt_GetBuffer(recOpCpt)) Then
        srvOpCpt_Seek = Null
    Else
        Call srvOpCpt_Error(recOpCpt)
    End If
End If

End Function

'---------------------------------------------------------
Private Function srvOpCpt_Snap(recOpCpt As typeOpCpt)
'---------------------------------------------------------
Dim I As Integer
srvOpCpt_Snap = "?"
MsgTxtLen = 0
Call srvOpCpt_PutBuffer(recOpCpt)
Call srvOpCpt_PutBuffer(arrOpCpt(0))
If IsNull(SndRcv()) Then
    srvOpCpt_Snap = Null
    MsgTxtIndex = 0
    Do While MsgTxtIndex < MsgTxtLen
        If IsNull(srvOpCpt_GetBuffer(recOpCpt)) Then
            Call arrOpCpt_AddItem(recOpCpt)
            arrOpCptSuite = True
        Else
            arrOpCptSuite = False
            Exit Do
        End If
    Loop
End If

End Function

'---------------------------------------------------------
Public Sub recOpCpt_Init(recOpCpt As typeOpCpt)
'---------------------------------------------------------
 MsgTxt = String$(recOpCptLen, "0")
 MsgTxtIndex = 0: Mid$(MsgTxt, MsgTxtIndex + 25, 10) = Space$(10)
 Call srvOpCpt_GetBuffer(recOpCpt)
 recOpCpt.Libellé = ""
 recOpCpt.obj = "SRVOPCPT"
End Sub

'---------------------------------------------------------
Public Sub arrOpCpt_AddItem(recOpCpt As typeOpCpt)
'---------------------------------------------------------
          
arrOpCptNb = arrOpCptNb + 1
    
If arrOpCptNb > arrOpCptNbMax Then
    arrOpCptNbMax = arrOpCptNbMax + 10
    ReDim Preserve arrOpCpt(arrOpCptNbMax)
End If
            
arrOpCpt(arrOpCptNb) = recOpCpt
End Sub




Public Sub recOpCpt_CommissionReset(recOpCpt As typeOpCpt)
recOpCpt.optComShare = "0"
recOpCpt.optComImputation = "0"
recOpCpt.optComBarème = "0"
recOpCpt_CommissionMontantZ recOpCpt


End Sub

Public Sub recOpCpt_CommissionMontantZ(recOpCpt As typeOpCpt)
recOpCpt.ComMontantDevise = 0
recOpCpt.ComMontantFRF = 0
recOpCpt.chkCom1 = "1"
recOpCpt.ComMontant1 = 0
recOpCpt.chkCom2 = "0"
recOpCpt.ComMontant2 = 0
recOpCpt.chkCom3 = "0"
recOpCpt.ComMontant3 = 0
recOpCpt.chkCom4 = "0"
recOpCpt.ComMontant4 = 0
recOpCpt.chkCom5 = "0"
recOpCpt.ComMontant5 = 0
recOpCpt.chkCom6 = "0"
recOpCpt.ComMontant6 = 0

End Sub


