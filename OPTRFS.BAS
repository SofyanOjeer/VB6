Attribute VB_Name = "srvOptrf"
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Public MopSwift(2) As typeOpSwift
Public XopSwift() As typeOpSwift
Public XopSwiftNb As Integer, XopSwiftNbMax As Integer, XopSwiftIndex As Integer
Public XopEch() As typeOpEch
Public XopEchNb As Integer, XopEchNbMax As Integer, XopEchIndex As Integer
Public mOpCpt(2) As typeOpCpt
Public XopCpt() As typeOpCpt
Public XopCptNb As Integer, XopCptNbMax As Integer, XopCptIndex As Integer

Public XopDevise(2) As typeDevise
Public XopBicSender As typeBic
Public XopBicIbanSender As typeBicIban
Public XopCompteSender As typeCompte
Public XopBicReceiver As typeBic
Public XopBicIbanReceiver As typeBicIban
Public XopBicIbanNostro As typeBicIban, XopBicIbanNostroLoro As typeBicIban
Public XopCompteNostro As typeCompte

Public XopAut As typeAuthorization
Public recOpCpt As typeOpCpt
Public recOpECh As typeOpEch
Public recCompte As typeCompte
Public recBdfG As typeBdfG

Public blnOpTrf_Update_on As Boolean

Public Sub XopBic_Load()
Dim V As Variant
If Trim(mId$(XopSwift(1).FieldText, 1, 11)) <> "" Then
    XopBicSender.Method = "SeekP0"
    XopBicSender.BicId = mId$(XopSwift(1).FieldText, 1, 11)
    V = srvBicFind(XopBicSender)
    recBicIbanInit XopBicIbanNostro
    XopBicIbanNostro.Method = "SeekP0"
    XopBicIbanNostro.BicId = SocBicIdNostro   'jpl 2000.03.21  SocBicId
    XopBicIbanNostro.IbanDevise = XopCpt(2).Devise
    XopBicIbanNostro.IbanBicId = mId$(XopSwift(1).FieldText, 13, 11)
        V = srvBicIbanFind(XopBicIbanNostro)
End If
If Val(XopCpt(1).Compte) <> 0 Then
    recCompteInit XopCompteSender
    XopCompteSender.Method = "SeekL1"
    XopCompteSender.Société = SocId$
    XopCompteSender.Agence = SocAgence$
    XopCompteSender.Devise = XopCpt(1).Devise
    XopCompteSender.Numéro = XopCpt(1).Compte
    V = srvCompteFind(XopCompteSender)
End If
If Val(XopCpt(2).Compte) <> 0 Then
    recCompteInit XopCompteNostro
    XopCompteNostro.Method = "SeekL1"
    XopCompteNostro.Société = SocId$
    XopCompteNostro.Agence = SocAgence$
    XopCompteNostro.Devise = XopCpt(2).Devise
    XopCompteNostro.Numéro = XopCpt(2).Compte
    V = srvCompteFind(XopCompteNostro)
End If
If Trim(mId$(XopSwift(2).FieldText, 1, 11)) <> "" Then
    XopBicReceiver.Method = "SeekP0"
    XopBicReceiver.BicId = mId$(XopSwift(2).FieldText, 1, 11)
     V = srvBicFind(XopBicReceiver)
    If Trim(mId$(XopSwift(2).FieldText, 13, 11)) <> "" Then
        recBicIbanInit XopBicIbanReceiver
        XopBicIbanReceiver.Method = "SeekP0"
        XopBicIbanReceiver.BicId = XopBicReceiver.BicId
        XopBicIbanReceiver.IbanDevise = XopCpt(2).Devise
        XopBicIbanReceiver.IbanBicId = mId$(XopSwift(2).FieldText, 13, 11)
        V = srvBicIbanFind(XopBicIbanReceiver)
    End If
End If
DevX XopCpt(1).Devise
XopDevise(1) = XDevise
DevX XopCpt(2).Devise
XopDevise(2) = XDevise


End Sub

'---------------------------------------------------------
Public Function srvOpTrf_dbAddNew() As String
'---------------------------------------------------------
Dim I As Integer
srvOpTrf_dbAddNew = ""
For I = 1 To XopCptNb
    XopCpt(I).Method = "AddNew"
    If Not IsNull(srvOpCpt_Update(XopCpt(I))) Then Exit Function
Next I

For I = 1 To XopEchNb
    XopEch(I).Method = "AddNew"
    If Not IsNull(srvOpEch_Update(XopEch(I))) Then Exit Function
Next I

For I = 1 To XopSwiftNb
    If Trim(XopSwift(I).FieldText) <> "" Then
        XopSwift(I).Method = "AddNew"
        If Not IsNull(srvOpSwift_Update(XopSwift(I))) Then Exit Function
    End If
Next I
srvOpTrf_dbAddNew = "Création effectuée"
End Function
'---------------------------------------------------------
Public Function srvOpTrf_dbDelete() As String
'---------------------------------------------------------
Dim I As Integer
srvOpTrf_dbDelete = ""
For I = 1 To XopCptNb
    XopCpt(I).Method = "Delete"
    If Not IsNull(srvOpCpt_Update(XopCpt(I))) Then Exit Function
Next I

For I = 1 To XopEchNb
    XopEch(I).Method = "Delete"
    If Not IsNull(srvOpEch_Update(XopEch(I))) Then Exit Function
Next I


For I = 1 To XopSwiftNb
    If Trim(XopSwift(I).FieldText) <> "" Then
        XopSwift(I).Method = "Delete"
        If Not IsNull(srvOpSwift_Update(XopSwift(I))) Then Exit Function
    End If
Next I
srvOpTrf_dbDelete = "Suppression effectuée"
End Function


Public Function srvOpTrf_dbRead(valRéférence As String)
srvOpTrf_dbReadSwift valRéférence
srvOpTrf_dbReadCpt valRéférence
srvOpTrf_dbReadEch valRéférence
End Function

Public Function srvOpTrf_dbReadSwift(valRéférence As String)
recOpSwift_Init XopSwift(0)
XopSwift(0).Référence = valRéférence
XopSwift(0).Method = "SnapP0"
arrOpSwift(0) = XopSwift(0)
arrOpSwift(0).FieldNature = "999"
arrOpSwift(0).FieldLine = "9"
arrOpSwiftNb = 0
srvOpSwift_Monitor XopSwift(0)
End Function

Public Function srvOpTrf_dbReadCpt(valRéférence As String)
recOpCpt_Init XopCpt(0)
XopCpt(0).Référence = valRéférence
XopCpt(0).Séquence = 0
XopCpt(0).Method = "SnapP0"
arrOpCpt(0) = XopCpt(0)
arrOpCpt(0).Séquence = 999
arrOpCptNb = 0
srvOpCpt_Monitor XopCpt(0)
End Function
Public Function srvOpTrf_dbReadEch(valRéférence As String)
recOpEch_Init XopEch(0)
XopEch(0).Référence = valRéférence
XopEch(0).Method = "SnapP0"
arrOpEch(0) = XopEch(0)
arrOpEch(0).EchéanceAmj = "99999999"
arrOpEch(0).EchéanceHms = "999999"
arrOpEchNb = 0
srvOpEch_Monitor XopEch(0)
End Function



'---------------------------------------------------------
Public Function srvOpTrf_dbUpdate() As String
'---------------------------------------------------------
srvOpTrf_dbUpdate = ""
srvOpTrf_dbRead XopSwift(1).Référence
srvOpTrf_dbUpdateOpCpt
srvOpTrf_dbUpdateOpEch
srvOpTrf_dbUpdateOpSwift
srvOpTrf_dbUpdate = "Mise à jour effectuée"
End Function

Public Function srvOpTrf_dbUpdateOpCpt()
Dim K As Integer, I As Integer
For I = 1 To XopCptNb
    If Trim(XopCpt(I).Référence) <> "" Then
        XopCpt(I).Method = "AddNew"
    Else
        XopCpt(I).Method = ""
    End If
Next I

For K = 1 To arrOpCptNb
    arrOpCpt(K).Method = "Delete"
    For I = 1 To XopCptNb
        If XopCpt(I).Référence = arrOpCpt(K).Référence _
        And XopCpt(I).Séquence = arrOpCpt(K).Séquence Then
            arrOpCpt(K).Method = ""
' à faire   : meme enregistrment ?
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'            If XopCpt(I).CptéanceNature = arrOpCpt(K).CptéanceNature _
'             XopCpt(I).Method = ""
'            Else
                XopCpt(I).Method = "Update"
'            End If
            Exit For
        End If
    Next I
Next K

For I = 1 To XopCptNb
    If Trim(XopCpt(I).Method) <> "" Then
        If Not IsNull(srvOpCpt_Update(XopCpt(I))) Then Exit For
    End If
Next I

For K = 1 To arrOpCptNb
    If Trim(arrOpCpt(K).Method) = "Delete" Then
        If Not IsNull(srvOpCpt_Update(arrOpCpt(K))) Then Exit For
    End If
Next K

End Function

Public Function srvOpTrf_dbUpdateOpEch()
Dim K As Integer, I As Integer
For I = 1 To XopEchNb
    If Trim(XopEch(I).Référence) <> "" Then
        XopEch(I).Method = "AddNew"
    Else
        XopEch(I).Method = ""
    End If
Next I

For K = 1 To arrOpEchNb
    arrOpEch(K).Method = "Delete"
    For I = 1 To XopEchNb
        If XopEch(I).Référence = arrOpEch(K).Référence _
        And XopEch(I).EchéanceAmj = arrOpEch(K).EchéanceAmj _
        And XopEch(I).EchéanceHms = arrOpEch(K).EchéanceHms Then
            arrOpEch(K).Method = ""
            
            If XopEch(I).EchéanceNature = arrOpEch(K).EchéanceNature _
            And XopEch(I).ActionAMJ = arrOpEch(K).ActionAMJ _
            And XopEch(I).ActionHMS = arrOpEch(K).ActionHMS _
            And XopEch(I).ActionNature = arrOpEch(K).ActionNature _
            And XopEch(I).ActionOpérateur = arrOpEch(K).ActionOpérateur Then
                XopEch(I).Method = ""
            Else
                XopEch(I).Method = "Update"
            End If
            Exit For
        End If
    Next I
Next K

For I = 1 To XopEchNb
    If Trim(XopEch(I).Method) <> "" Then
        If Not IsNull(srvOpEch_Update(XopEch(I))) Then Exit For
    End If
Next I

For K = 1 To arrOpEchNb
    If Trim(arrOpEch(K).Method) = "Delete" Then
        If Not IsNull(srvOpEch_Update(arrOpEch(K))) Then Exit For
    End If
Next K

End Function

Public Function srvOpTrf_dbUpdateOpSwift()
Dim K As Integer, I As Integer

For I = 1 To XopSwiftNb
    If Trim(XopSwift(I).FieldText) <> "" Then
        XopSwift(I).Method = "AddNew"
    Else
        XopSwift(I).Method = ""
    End If
    
Next I
For K = 1 To arrOpSwiftNb
    arrOpSwift(K).Method = "Delete"
    For I = 1 To XopSwiftNb
        If XopSwift(I).FieldNature = arrOpSwift(K).FieldNature _
        And XopSwift(I).FieldLine = arrOpSwift(K).FieldLine Then
            If XopSwift(I).FieldText = arrOpSwift(K).FieldText Then
                XopSwift(I).Method = ""
                arrOpSwift(K).Method = ""
            Else
                If Trim(XopSwift(I).FieldText) <> "" Then
                    XopSwift(I).Method = "Update"
                    arrOpSwift(K).Method = ""
                End If
            End If
            Exit For
        End If
    Next I
Next K

For I = 1 To XopSwiftNb
    If Trim(XopSwift(I).Method) <> "" Then
        If Not IsNull(srvOpSwift_Update(XopSwift(I))) Then 'Exit For
            Call MsgBox(XopSwift(I).Method, vbCritical, XopSwift(I).FieldNature & " " & XopSwift(I).FieldNature)
        End If
    End If
Next I
For K = 1 To arrOpSwiftNb
    If Trim(arrOpSwift(K).Method) = "Delete" Then
        
        If Not IsNull(srvOpSwift_Update(arrOpSwift(K))) Then ' Exit For
            Call MsgBox(XopSwift(I).Method, vbCritical, XopSwift(I).FieldNature & " " & XopSwift(I).FieldNature)
        End If
    End If
Next K

End Function




'---------------------------------------------------------
Public Sub XopSwift_Display(picBoxFieldText As PictureBox)
'---------------------------------------------------------
Dim I As Integer, K As Integer, XX As String, X2 As String

picBoxFieldText.Visible = True
picBoxFieldText.Cls
picBoxFieldText.FontBold = False
picBoxFieldText.CurrentY = -lstLineHeight
X2 = "": K = 0
For I = 1 To XopSwiftNb
    XX = mId$(XopSwift(I).FieldNature, 1, 2)
    If XX <> "01" Then
        picBoxFieldText.CurrentY = picBoxFieldText.CurrentY + lstLineHeight
        If XX <> X2 Then
            K = 0
            X2 = mId$(XopSwift(I).FieldNature, 1, 2)
            picBoxFieldText.ForeColor = warnUsrColor 'frmUsr.ForeColor
            picBoxFieldText.CurrentX = 50: picBoxFieldText.Print XopSwift(I).FieldNature;
        End If
        K = K + 1: XopSwift(I).FieldLine = K
        picBoxFieldText.ForeColor = libUsr.ForeColor
        picBoxFieldText.CurrentX = 600
        picBoxFieldText.Print XopSwift(I).FieldText;
    End If
Next I
End Sub
Public Sub xOpSwift_arrDisplay(XpicBox As PictureBox, IX0 As Integer)
Dim I As Integer, K As Integer, X As String, X2 As String, ix0b As Integer
X2 = "": XpicBox.CurrentY = 0
XpicBox.FontBold = False
XpicBox.FontSize = 6
For I = 1 To arrOpSwiftNb
    X = mId$(arrOpSwift(I).FieldNature, 1, 2)
    If X > "01" Then
       XpicBox.CurrentY = XpicBox.CurrentY + lstLineHeight
       XpicBox.FontBold = False
        If X <> X2 Then
            K = 0
            X2 = mId$(arrOpSwift(I).FieldNature, 1, 2)
           XpicBox.ForeColor = lblUsr.ForeColor
            XpicBox.CurrentX = IX0
            If IX0 = 0 Then
                XpicBox.Print mId$(lblOpTrf("Swift", X2), 1, 20);
                ix0b = IX0 + 1900
            Else
                XpicBox.Print X2;
                ix0b = IX0 + 300
            End If
            If X2 = "32" Or X2 = "57" Or X2 = "58" Or X2 = "71" Then XpicBox.FontBold = True
           XpicBox.ForeColor = libUsr.ForeColor
        End If
        K = K + 1: arrOpSwift(I).FieldLine = K
        XpicBox.CurrentX = ix0b
        XpicBox.Print Trim(arrOpSwift(I).FieldText);
    End If
Next I

End Sub


Public Sub XopSwift_Load()
Dim I As Integer
If arrOpSwiftNb > XopSwiftNbMax Then ReDim XopSwift(arrOpSwiftNb)
For I = 1 To arrOpSwiftNb
    XopSwift(I) = arrOpSwift(I)
Next I
XopSwiftNb = arrOpSwiftNb
XopSwiftNbMax = arrOpSwiftNb

End Sub

Public Sub Xopswift_MT100R()
Dim X As String, XX As String
Dim I As Integer, K As Integer, X2 As String * 2
ReDim arrOpSwift(XopSwiftNb + 3): arrOpSwiftNbMax = XopSwiftNb + 3: arrOpSwiftNb = 0
For I = 1 To XopSwiftNb
    X2 = mId$(XopSwift(I).FieldNature, 1, 2)
    If X2 <> "00" Then
        arrOpSwiftNb = arrOpSwiftNb + 1
        arrOpSwift(arrOpSwiftNb) = XopSwift(I)
        Select Case X2
            Case "20": arrOpSwift(arrOpSwiftNb).FieldText = XopCpt(2).CodeOpération & " " & XopCpt(2).Référence
            Case "32"
                arrOpSwift(arrOpSwiftNb).FieldText = dateImp(XopCpt(2).AmjValeur) & " " & XopDevise(2).DevX & " " _
                & num_Display(XopCpt(2).Brut - XopCpt(2).ComMontantDevise, 15, XopDevise(2).maxD, Lx, X, "0")
                If XopCpt(2).optSwift = 4 Then XopSwift_MT100CRI
            Case "57"
                If arrOpSwift(arrOpSwiftNb).FieldLine = 1 Then
                    arrOpSwift(arrOpSwiftNb).FieldNature = "54"
                    If XopCpt(2).Devise = "001" Then
                        If Trim(mId$(XopSwift(3).FieldText, 1, 27)) <> "" Then
                            arrOpSwift(arrOpSwiftNb).FieldText = "/" & mId$(XopSwift(3).FieldText, 1, 27)
                            arrOpSwiftNb = arrOpSwiftNb + 1
                            arrOpSwift(arrOpSwiftNb) = arrOpSwift(arrOpSwiftNb - 1)
                            arrOpSwift(arrOpSwiftNb).FieldLine = 2
                        End If
                    End If
                    arrOpSwift(arrOpSwiftNb).FieldText = mId$(XopSwift(2).FieldText, 13, 11)
                    arrOpSwiftNb = arrOpSwiftNb + 1
                    arrOpSwift(arrOpSwiftNb) = XopSwift(I)
                End If
            Case "71"
                If XopCpt(2).ComMontantFRF = 0 Then
                    arrOpSwiftNb = arrOpSwiftNb - 1
                Else
                    arrOpSwift(arrOpSwiftNb).FieldNature = "70 "
                    arrOpSwift(arrOpSwiftNb).FieldLine = 6
                    X = num_Display(XopCpt(2).Brut, 15, XopDevise(2).maxD, Lx, XX, "0")
                    arrOpSwift(arrOpSwiftNb).FieldText = "/" & XopDevise(2).DevX & " " & Trim(X) & " less o/charges"
                End If
            Case "72"
                If XopCpt(2).Devise = "400" Then
                    arrOpSwift(arrOpSwiftNb).FieldText = "/ACC/in case you contact our"
                    arrOpSwiftNb = arrOpSwiftNb + 1
                    arrOpSwift(arrOpSwiftNb) = XopSwift(I)
                    arrOpSwift(arrOpSwiftNb).FieldLine = 2
                    arrOpSwift(arrOpSwiftNb).FieldText = "/correspondant in USA,please don't"
                    arrOpSwiftNb = arrOpSwiftNb + 1
                    arrOpSwift(arrOpSwiftNb) = XopSwift(I)
                    arrOpSwift(arrOpSwiftNb).FieldLine = 2
                    arrOpSwift(arrOpSwiftNb).FieldText = "/mention the name of our bank."
                End If
            End Select
    End If
Next I
End Sub
Public Sub Xopswift_MT100S()
Dim I As Integer, K As Integer
ReDim arrOpSwift(XopSwiftNb): arrOpSwiftNbMax = XopSwiftNb: arrOpSwiftNb = 0
For I = 1 To XopSwiftNb
    If mId$(XopSwift(I).FieldNature, 1, 2) <> "01" Then
        arrOpSwiftNb = arrOpSwiftNb + 1
        arrOpSwift(arrOpSwiftNb) = XopSwift(I)
    End If
Next I

End Sub

Public Sub XopSwift_MT202()
Dim I As Integer, K As Integer, X As String
ReDim arrOpSwift(XopSwiftNb): arrOpSwiftNbMax = XopSwiftNb: arrOpSwiftNb = 0
arrOpSwift(1) = XopSwift(3): arrOpSwift(1).FieldText = XopCpt(2).CodeOpération & " " & XopCpt(2).Référence
arrOpSwift(2) = XopSwift(3): arrOpSwift(2).FieldNature = "21"
arrOpSwift(3) = XopSwift(4): arrOpSwift(3).FieldNature = "32"
                             arrOpSwift(3).FieldText = dateImp(XopCpt(2).AmjValeur) & " " & XopDevise(2).DevX & " " _
                             & num_Display(XopCpt(2).Brut - XopCpt(2).ComMontantDevise, 15, XopDevise(2).maxD, Lx, X, "0")
If mId$(XopSwift(1).FieldText, 12, 1) = "*" Then
    arrOpSwiftNb = 3
    XopSwift_MT202Bank
Else
    arrOpSwiftNb = 5
    arrOpSwift(4) = XopSwift(1): arrOpSwift(4).FieldNature = "57"
                                 arrOpSwift(4).FieldText = mId$(XopSwift(2).FieldText, 13, 11)
    arrOpSwift(5) = XopSwift(1): arrOpSwift(5).FieldNature = "58"
                                 arrOpSwift(5).FieldText = mId$(XopSwift(2).FieldText, 1, 11)
End If
End Sub



Public Function XopEch_Load() As String
Dim I As Integer
If arrOpEchNb > XopEchNbMax Then ReDim XopEch(arrOpEchNb)
XopEch_Load = ""
For I = 1 To arrOpEchNb
    XopEch(I) = arrOpEch(I)
    If Trim(XopEch(I).ActionNature) = "" Then
        XopEch_Load = XopEch(I).EchéanceNature
    End If
Next I
XopEchNb = arrOpEchNb
XopEchNbMax = arrOpEchNb
End Function

Public Sub XopEch_ActionValider()
TSys = time_Hms
recOpECh.EchéanceHms = TSys
recOpECh.EchéanceNature = constàCompta: XopEch_AddItem recOpECh

'$$$$$$$$ optrf: pas de virement automatique
'recOpECh.EchéanceHms =tsys
'Select Case XopCpt(2).optVirement
'    Case "1": recOpECh.EchéanceNature = constàVirXX: XopEch_AddItem recOpECh
'    Case "2": recOpECh.EchéanceNature = constàVirAvis: XopEch_AddItem recOpECh
'    Case "3": recOpECh.EchéanceNature = constàVirCRI: XopEch_AddItem recOpECh
'End Select
'$$$$$$$$ optrf: pas de SWIFT automatique
'If XopCpt(2).optSwift > "0" Then
'    recOpECh.EchéanceNature = constàMT100: XopEch_AddItem recOpECh
'    If XopCpt(2).optSwift = "2" Then
'        recOpECh.EchéanceNature = constàMT202: XopEch_AddItem recOpECh
'    End If
'End If
If XopCpt(1).optAvis <> "0" Then
    If XopCpt(1).optAvis = "2" Then
        recOpECh.EchéanceAmj = dateElp("Decade", 0, DSys)
        recOpECh.EchéanceHms = "170000"
    Else
        recOpECh.EchéanceAmj = DSys
        recOpECh.EchéanceHms = TSys
    End If
    recOpECh.EchéanceNature = constàAvisDb: XopEch_AddItem recOpECh
End If


If XopCpt(2).optAvis <> "0" Then
    If XopCpt(2).optAvis = "2" Then
        recOpECh.EchéanceAmj = dateElp("Decade", 0, DSys)
        recOpECh.EchéanceHms = "170000"
    Else
        recOpECh.EchéanceAmj = DSys
        recOpECh.EchéanceHms = TSys
    End If
    recOpECh.EchéanceNature = constàAvisCr: XopEch_AddItem recOpECh
End If

'$$$$$$$$ optrf: pas de rapprochement manuel
'recOpECh.EchéanceAmj = DateElp("Decade",0,Dsys)
'recOpECh.EchéanceHms = "170000"
'recOpECh.EchéanceNature = constàRapprocher: XopEch_AddItem recOpECh

'$$$$Call prtOpTrfX("Validation")
Call prtOpTrfX("Validation", "Service Comptabilité")

End Sub

Public Sub XopEch_ActionSaisie(ByVal valRéférence As String)
recOpEch_Init recOpECh
recOpECh.Référence = valRéférence
recOpECh.EchéanceAmj = DSys
recOpECh.EchéanceHms = time_Hms
recOpECh.EchéanceNature = constSaisie
recOpECh.ActionOpérateur = usrId
XopEch_AddItem recOpECh
End Sub


Public Sub XopCpt_ActionSaisie(ByVal valRéférence As String)
recOpCpt_Init recOpCpt
recOpCpt.Société = SocId$
recOpCpt.Agence = SocAgence$
recOpCpt.Référence = valRéférence
recOpCpt.AmjOpération = DSys
recOpCpt.CodeOpération = "OPTRF"
recOpCpt.Séquence = "001"
recOpCpt.Sens = "D"
recOpCpt.chkCom1 = "1"
recOpCpt.optComImputation = 4
XopCpt_AddItem recOpCpt
recOpCpt.Séquence = "002"
recOpCpt.Sens = "C"
XopCpt_AddItem recOpCpt
XopCpt(2).AmjValeur = "00000000" 'DSys
End Sub


Public Sub XopEch_Action(currentAction As String)
Dim I As Integer
For I = 1 To XopEchNb
    If Trim(XopEch(I).EchéanceNature) = Trim(currentAction) _
    And Trim(XopEch(I).ActionNature) = "" Then
        XOpEch_ActionOk XopEch(I)
        If Trim(currentAction) = constSaisie Then XopEch(I).ActionNature = constàValider
        recOpECh = XopEch(I)
        recOpECh.EchéanceAmj = recOpECh.ActionAMJ
        recOpECh.EchéanceHms = recOpECh.ActionHMS
        recOpECh.ActionAMJ = "00000000"
        recOpECh.ActionHMS = "000000"
        recOpECh.ActionNature = ""
        recOpECh.ActionOpérateur = ""
        Select Case Trim(currentAction)
            Case constSaisie
                recOpECh.EchéanceNature = constàValider: XopEch_AddItem recOpECh
                Call prtOpTrfX(constDemandeDeValidation) ' !!!!! champ " ...." implique l'impression d'une boîte Visa (prtOpTrf.bas\prtOpTrfX
            Case constàValider: XopEch_ActionValider
        End Select
        Exit For
    End If
Next I
End Sub

Public Sub XopSwift_Init(ByVal valRéférence As String)
recOpSwift_Init arrOpSwift(0)
arrOpSwift(0).Référence = valRéférence
arrOpSwift(0).FieldLine = 1
arrOpSwift(0).FieldNature = "00": Call XopSwift_Insert(arrOpSwift(0))
arrOpSwift(0).FieldNature = "01": Call XopSwift_Insert(arrOpSwift(0))
arrOpSwift(0).FieldNature = "02": Call XopSwift_Insert(arrOpSwift(0))
arrOpSwift(0).FieldNature = "20": Call XopSwift_Insert(arrOpSwift(0))
arrOpSwift(0).FieldNature = "50": Call XopSwift_Insert(arrOpSwift(0))
arrOpSwift(0).FieldNature = "52": Call XopSwift_Insert(arrOpSwift(0))
arrOpSwift(0).FieldNature = "57": Call XopSwift_Insert(arrOpSwift(0))
arrOpSwift(0).FieldNature = "59": Call XopSwift_Insert(arrOpSwift(0))
arrOpSwift(0).FieldNature = "70": Call XopSwift_Insert(arrOpSwift(0))
arrOpSwift(0).FieldNature = "72": Call XopSwift_Insert(arrOpSwift(0))
arrOpSwift(0).FieldText = "FREE"
arrOpSwift(0).FieldNature = "71": Call XopSwift_Insert(arrOpSwift(0))
arrOpSwift(0).FieldText = dateImp(DSys) & "   000.00"
arrOpSwift(0).FieldNature = "32A": Call XopSwift_Insert(arrOpSwift(0))

End Sub





Public Sub XopSwift_Insert(recOpSwift As typeOpSwift)
Dim I As Integer, K As Integer, X2 As String
X2 = mId$(recOpSwift.FieldNature, 1, 2)
For I = 1 To XopSwiftNb
    If X2 = mId$(XopSwift(I).FieldNature, 1, 2) Then Exit Sub
    If X2 < mId$(XopSwift(I).FieldNature, 1, 2) Then
        XopSwift_AddItem recOpSwift
        For K = XopSwiftNb - 1 To I Step -1
            XopSwift(K + 1) = XopSwift(K)
        Next K
        XopSwift(I) = recOpSwift
        Exit Sub
    End If
Next I
XopSwift_AddItem recOpSwift

End Sub


'---------------------------------------------------------
Public Sub XopSwift_AddItem(recOpSwift As typeOpSwift)
'---------------------------------------------------------
          
XopSwiftNb = XopSwiftNb + 1
    
If XopSwiftNb > XopSwiftNbMax Then
    XopSwiftNbMax = XopSwiftNbMax + 10
    ReDim Preserve XopSwift(XopSwiftNbMax)
End If
            
XopSwift(XopSwiftNb) = recOpSwift
End Sub

'---------------------------------------------------------
Public Sub XopCpt_AddItem(recOpCpt As typeOpCpt)
'---------------------------------------------------------
          
XopCptNb = XopCptNb + 1
    
If XopCptNb > XopCptNbMax Then
    XopCptNbMax = XopCptNbMax + 10
    ReDim Preserve XopCpt(XopCptNbMax)
End If
            
XopCpt(XopCptNb) = recOpCpt
End Sub

'---------------------------------------------------------
Public Sub XopCpt_Display(XpicBox As PictureBox)
'---------------------------------------------------------
Dim I As Integer, K As Integer, X As String, X2 As String, xDev As String
Dim IX0 As Integer
Dim curMt As Currency
XpicBox.Visible = True
XpicBox.Cls
XpicBox.CurrentY = -lstLineHeight
X2 = "": K = 0
XpicBox.FontBold = False
For I = 1 To XopCptNb
    XpicBox.ForeColor = libUsr.ForeColor
    XpicBox.FontBold = True
    XpicBox.FontSize = 8
    IX0 = 0 'IIf(I = 1, 0, 8100) + prtMinX
    XpicBox.CurrentY = XpicBox.CurrentY + lstLineHeight * 2
    XpicBox.CurrentX = IX0
    XpicBox.Print XopCpt(I).Devise & ".";
    XpicBox.Print Compte_Imp(XopCpt(I).Compte);
    XpicBox.CurrentX = IX0 + 2300
    If I = 1 Then
        XpicBox.Print XopCompteSender.Intitulé;
        If Trim(XopCompteSender.Intitulé2) <> "" Then
            XpicBox.CurrentX = IX0 + 2300
            XpicBox.CurrentY = XpicBox.CurrentY + lstLineHeight
            XpicBox.Print XopCompteSender.Intitulé2;
        End If
    Else
        XpicBox.Print XopCompteNostro.Intitulé;
        If Trim(XopCompteNostro.Intitulé2) <> "" Then
            XpicBox.CurrentX = IX0 + 2300
            XpicBox.CurrentY = XpicBox.CurrentY + lstLineHeight
            XpicBox.Print XopCompteNostro.Intitulé2;
        End If
  End If
    XpicBox.ForeColor = IIf(XopCpt(I).Sens = "D", errUsr.ForeColor, libUsr.ForeColor)
    XpicBox.FontUnderline = True
    If XopCpt(I).Sens = "C" Then
        curMt = XopCpt(I).Brut - XopCpt(I).ComMontantDevise
    Else
        curMt = XopCpt(I).Brut + XopCpt(I).ComMontantDevise
    End If
    X = num_Display(curMt, 15, XopDevise(I).maxD, Lx, X, "0")
    XpicBox.CurrentX = IX0 + 8650 - XpicBox.TextWidth(X)
    XpicBox.Print X;
    XpicBox.FontUnderline = False
    
    XpicBox.CurrentX = IX0 + 8700
    XpicBox.Print XopCpt(I).Sens;
    XpicBox.ForeColor = warnUsrColor
    XpicBox.Print "    " & XopDevise(I).DevX;
    XpicBox.ForeColor = libUsr.ForeColor
    
    XpicBox.CurrentY = XpicBox.CurrentY + lstLineHeight
    XpicBox.FontBold = False
    XpicBox.CurrentX = IX0 + 1000
    XpicBox.Print dateImp(XopCpt(I).AmjValeur);
    XpicBox.CurrentX = IX0 + 2300
    XpicBox.Print XopCpt(I).Libellé;
    XpicBox.ForeColor = IIf(XopCpt(I).Sens = "D", errUsr.ForeColor, libUsr.ForeColor)
    X = num_Display(XopCpt(I).Brut, 15, XopDevise(I).maxD, Lx, X, "0")
    XpicBox.CurrentX = IX0 + 8650 - XpicBox.TextWidth(X)
    XpicBox.Print X;
    
    XpicBox.ForeColor = libUsr.ForeColor
    XpicBox.FontSize = 7
    XpicBox.CurrentX = IX0
    XpicBox.Print "D.valeur: ";

    XpicBox.CurrentY = XpicBox.CurrentY + lstLineHeight
    If XopCpt(I).AmjOpération <> XopCpt(I).AmjValeur Then
        XpicBox.CurrentX = IX0
        XpicBox.Print "D.opération: ";
        XpicBox.CurrentX = IX0 + 1000
        XpicBox.Print dateImp(XopCpt(I).AmjOpération);
    End If
    XpicBox.CurrentX = IX0 + 2300:    XpicBox.Print "Commission:" & " (" & XopCpt(I).ComPays & ") ";
    XpicBox.Print lblOpTrf("optComImputation", XopCpt(I).optComImputation) & " ";
    If XopCpt(I).ComMontantFRF <> 0 Then
        XpicBox.Print lblOpTrf("optComShare", XopCpt(I).optComShare) & " barème ";
        XpicBox.Print lblOpTrf("optComBarème", XopCpt(I).optComBarème) & " ";
        If XopCpt(I).Devise <> "001" Then XpicBox.Print num_Display(XopCpt(I).CoursDevise, 2, 7, Lx, X, "0");
        XpicBox.FontSize = 8
        XpicBox.ForeColor = errUsr.ForeColor
        X = num_Display(XopCpt(I).ComMontantDevise, 15, XopDevise(I).maxD, Lx, X, "0")
        XpicBox.CurrentX = IX0 + 8650 - XpicBox.TextWidth(X)
        XpicBox.Print X;
 '       XpicBox.CurrentX = IX0 + 8700
 '       XpicBox.Print XopCpt(I).Sens;
        XpicBox.ForeColor = libUsr.ForeColor
        XpicBox.FontSize = 7
        If XopCpt(I).chkCom1 <> "0" Then
            XpicBox.CurrentY = XpicBox.CurrentY + lstLineHeight
            XpicBox.CurrentX = IX0 + 2300
            XpicBox.Print lblOpTrf("chkCom", "1");
            X = num_Display(XopCpt(I).ComMontant1, 12, 2, Lx, X, "0")
            XpicBox.CurrentX = IX0 + 5000 - XpicBox.TextWidth(X)
            XpicBox.Print X;
        End If
        
        If XopCpt(I).chkCom2 <> "0" Then
            XpicBox.CurrentY = XpicBox.CurrentY + lstLineHeight
            XpicBox.CurrentX = IX0 + 2300
            XpicBox.Print lblOpTrf("chkCom", "2");
            X = num_Display(XopCpt(I).ComMontant2, 12, 2, Lx, X, "0")
            XpicBox.CurrentX = IX0 + 5000 - XpicBox.TextWidth(X)
            XpicBox.Print X;
        End If
        
        If XopCpt(I).chkCom3 <> "0" Then
            XpicBox.CurrentY = XpicBox.CurrentY + lstLineHeight
            XpicBox.CurrentX = IX0 + 2300
            XpicBox.Print lblOpTrf("chkCom", "3");
            X = num_Display(XopCpt(I).ComMontant3, 12, 2, Lx, X, "0")
            XpicBox.CurrentX = IX0 + 5000 - XpicBox.TextWidth(X)
            XpicBox.Print X;
        End If
        
        If XopCpt(I).chkCom4 <> "0" Then
            XpicBox.CurrentY = XpicBox.CurrentY + lstLineHeight
            XpicBox.CurrentX = IX0 + 2300
            XpicBox.Print lblOpTrf("chkCom", "4");
            X = num_Display(XopCpt(I).ComMontant4, 12, 2, Lx, X, "0")
            XpicBox.CurrentX = IX0 + 5000 - XpicBox.TextWidth(X)
            XpicBox.Print X;
        End If
        
        If XopCpt(I).chkCom5 <> "0" Then
            XpicBox.CurrentY = XpicBox.CurrentY + lstLineHeight
            XpicBox.CurrentX = IX0 + 2300
            XpicBox.Print lblOpTrf("chkCom", "5");
            X = num_Display(XopCpt(I).ComMontant5, 12, 2, Lx, X, "0")
            XpicBox.CurrentX = IX0 + 5000 - XpicBox.TextWidth(X)
            XpicBox.Print X;
        End If
    End If
Next I
End Sub


Public Sub XopCpt_Load()
Dim I As Integer
If arrOpCptNb > XopCptNbMax Then ReDim XopCpt(arrOpCptNb)
For I = 1 To arrOpCptNb
    XopCpt(I) = arrOpCpt(I)
Next I
XopCptNb = arrOpCptNb
XopCptNbMax = arrOpCptNb
End Sub

'---------------------------------------------------------
Public Sub XopEch_Display(XpicBox As PictureBox)
'---------------------------------------------------------
Dim I As Integer, K As Integer, X As String, X2 As String, xDev As String

XpicBox.Visible = True
XpicBox.Cls
XpicBox.FontBold = False
XpicBox.CurrentY = -lstLineHeight
X2 = "": K = 0
For I = 1 To XopEchNb
    XpicBox.CurrentY = XpicBox.CurrentY + lstLineHeight
    Call recOpEch_Display(XpicBox, XopEch(I))
Next I
End Sub

'---------------------------------------------------------
Public Sub XopEch_AddItem(recOpECh As typeOpEch)
'---------------------------------------------------------
          
XopEchNb = XopEchNb + 1
    
If XopEchNb > XopEchNbMax Then
    XopEchNbMax = XopEchNbMax + 10
    ReDim Preserve XopEch(XopEchNbMax)
End If
            
XopEch(XopEchNb) = recOpECh
End Sub


Public Sub XopEch_ActionSuppress(currentAction As String)
Dim I As Integer
For I = 1 To XopEchNb
    If XopEch(I).EchéanceNature = currentAction Then
        XopEch(I).Référence = ""
    End If
    If XopEch(I).ActionNature = currentAction Then
        XopEch(I).ActionAMJ = "00000000"
        XopEch(I).ActionHMS = "000000"
        XopEch(I).ActionNature = ""
    End If
Next I
End Sub




Public Sub XOpEch_ActionOk(okOpEch As typeOpEch)
okOpEch.ActionAMJ = DSys
okOpEch.ActionHMS = time_Hms
okOpEch.ActionNature = mId$(okOpEch.EchéanceNature, 3, 10)
okOpEch.ActionOpérateur = usrId
End Sub

Public Sub srvOpTrf_dbUpdateOpEch_ActionOk(recOpECh As typeOpEch)
If Trim(recOpECh.ActionNature) = "" Then
    XOpEch_ActionOk recOpECh
    recOpECh.Method = "Update"
    srvOpEch_Update recOpECh
End If

End Sub

Public Sub XopSwift_MT202Bank()
Dim X As String, XX As String
Dim I As Integer, K As Integer, X2 As String * 2
For I = 1 To XopSwiftNb
    X2 = mId$(XopSwift(I).FieldNature, 1, 2)
    Select Case X2
        Case "50"
                arrOpSwiftNb = arrOpSwiftNb + 1
                arrOpSwift(arrOpSwiftNb) = XopSwift(I)
                arrOpSwift(arrOpSwiftNb).FieldNature = "52 "
        Case "57", "72"
                arrOpSwiftNb = arrOpSwiftNb + 1
                arrOpSwift(arrOpSwiftNb) = XopSwift(I)
        Case "59"
                arrOpSwiftNb = arrOpSwiftNb + 1
                arrOpSwift(arrOpSwiftNb) = XopSwift(I)
                arrOpSwift(arrOpSwiftNb).FieldNature = "58 "
    End Select
Next I

End Sub

Public Sub XopSwift_MT100CRI()
XopBicIbanNostroLoro = XopBicIbanNostro
XopBicIbanNostroLoro.Method = "SeekL1"
XopBicIbanNostroLoro.BicId = SocBicId
If IsNull(srvBicIbanFind(XopBicIbanNostroLoro)) Then
    arrOpSwiftNb = arrOpSwiftNb + 1
    arrOpSwift(arrOpSwiftNb).FieldNature = "53 "
    arrOpSwift(arrOpSwiftNb).FieldLine = 1
    arrOpSwift(arrOpSwiftNb).FieldText = "/D/" & XopBicIbanNostroLoro.IbanCompte
End If

End Sub

Public Sub XopTrf_Display(picBoxHeader As PictureBox)
Dim curMt As Currency

picBoxHeader.Cls
 
picBoxHeader.ForeColor = warnUsrColor
picBoxHeader.CurrentX = 0: picBoxHeader.CurrentY = 0
picBoxHeader.Print XopCpt(1).Devise & "." & Compte_Imp(XopCpt(1).Compte);
picBoxHeader.ForeColor = libUsr.ForeColor
picBoxHeader.Print " " & mId$(XopCompteSender.Intitulé, 1, 30);
picBoxHeader.ForeColor = vbRed
picBoxHeader.CurrentX = 4500
If XopCpt(1).Sens = "C" Then
    curMt = XopCpt(1).Brut - XopCpt(1).ComMontantDevise
Else
    curMt = XopCpt(1).Brut + XopCpt(1).ComMontantDevise
End If
picBoxHeader.Print Format$(curMt, "### ### ### ##0.00");

picBoxHeader.Print " " & Trim(XopDevise(1).DevLib);
picBoxHeader.ForeColor = libUsr.ForeColor

picBoxHeader.ForeColor = warnUsrColor
picBoxHeader.CurrentX = 0: picBoxHeader.CurrentY = 250
picBoxHeader.Print XopCpt(2).Devise & "." & Compte_Imp(XopCpt(2).Compte);
picBoxHeader.ForeColor = libUsr.ForeColor
picBoxHeader.Print " " & mId$(XopCompteNostro.Intitulé, 1, 30);
picBoxHeader.ForeColor = vbBlue
picBoxHeader.CurrentX = 4500
If XopCpt(2).Sens = "C" Then
    curMt = XopCpt(2).Brut - XopCpt(2).ComMontantDevise
Else
    curMt = XopCpt(2).Brut + XopCpt(2).ComMontantDevise
End If
picBoxHeader.Print Format$(curMt, "### ### ### ##0.00");

picBoxHeader.ForeColor = libUsr.ForeColor
picBoxHeader.Print " " & Trim(XopDevise(2).DevLib);

picBoxHeader.CurrentX = 7000: picBoxHeader.CurrentY = 0
picBoxHeader.Print XopSwift(1).FieldText;
picBoxHeader.CurrentX = 7000: picBoxHeader.CurrentY = 250
picBoxHeader.Print XopSwift(2).FieldText;
End Sub

Public Sub AnnulerValidation(valRéférence As String)
Dim I As Integer, X As String

X = MsgBox("Confirmez-vous cette transaction ?", vbYesNo + vbQuestion + vbDefaultButton2, "Demande d'annulation de validation")
If X <> vbYes Then Exit Sub

srvOpTrf_dbReadEch valRéférence

For I = 1 To arrOpEchNb
    Select Case Trim(arrOpEch(I).EchéanceNature)
        Case constSaisie
            arrOpEch(I).Method = constUpdate
            arrOpEch(I).ActionAMJ = ""
            arrOpEch(I).ActionHMS = ""
            arrOpEch(I).ActionNature = ""
            arrOpEch(I).ActionOpérateur = ""
        Case constàValider
                arrOpEch(I).Method = constDelete
        Case Else
            If Trim(arrOpEch(I).ActionNature) = "" Then
                arrOpEch(I).Method = constDelete
            Else
                MsgBox "Demande d'annulation de validation refusée", vbCritical, arrOpEch(I).ActionNature & " : déjà effectuée"
                Exit Sub
            End If
    End Select
Next I

For I = 1 To arrOpEchNb
    If Not IsNull(srvOpEch_Update(arrOpEch(I))) Then Exit For
Next I


End Sub
