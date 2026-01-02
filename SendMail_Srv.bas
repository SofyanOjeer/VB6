Attribute VB_Name = "srvSendMail"
Option Explicit

Dim rsSab_Local As New ADODB.Recordset

Private poSendMail As vbSendMail.clsSendMail
Type typeSendMail

    From                    As String         ' Required the fist time, optional thereafter
    FromDisplayName         As String         ' Optional, saved after first use
    Recipient               As String         ' Required, separate multiple entries with delimiter character
    RecipientDisplayName    As String         ' Optional, separate multiple entries with delimiter character
    CcRecipient             As String         ' Optional, separate multiple entries with delimiter character
    CcDisplayName           As String         ' Optional, separate multiple entries with delimiter character
    Subject                 As String         ' Optional
    Message                 As String         ' Optional
    Attachment              As String         ' Optional, separate multiple entries with delimiter character
    AsHTML                  As Boolean        ' Optional, default = FALSE, send mail as html or plain text

End Type

Public Function RGB_Html_Color(lColor As Long) As String
Dim xColor As String, X As String
xColor = Hex(lColor)
Select Case Len(xColor)
    Case 6:
    Case 2: xColor = "0000" & xColor
    Case 4: xColor = "00" & xColor
    Case 1: xColor = "00000" & xColor
    Case 3: xColor = "000" & xColor
    Case 5: xColor = "0" & xColor
End Select

RGB_Html_Color = " #" & Mid$(xColor, 5, 2) & Mid$(xColor, 3, 2) & Mid$(xColor, 1, 2)
End Function
Public Sub Route(lSendMail As typeSendMail)
Dim blnDir As Boolean
Dim K As Integer, K1 As Integer
Dim wSéparateur As String
Dim X As String
Dim mAttachment As String

Set poSendMail = New clsSendMail

On Error Resume Next
If lSendMail.Attachment <> "" Then
    wSéparateur = ""
    K1 = 1
    mAttachment = lSendMail.Attachment & ";"
    lSendMail.Attachment = ""
    Do
        K = InStr(K1, mAttachment, ";")
        If K > 0 Then
            X = Mid$(mAttachment, K1, K - K1)
            If X <> "" Then
                If Trim(Dir(X)) <> "" Then lSendMail.Attachment = lSendMail.Attachment & wSéparateur & X
            End If
            wSéparateur = ";"
            K1 = K + 1
        End If
    Loop While K > 0
End If

If blnOff_Line Then Exit Sub  'lSendMail.Recipient = lSendMail.From

    With poSendMail

        ' **************************************************************************
        ' Optional properties for sending email, but these should be set first
        ' if you are going to use them
        ' **************************************************************************

        .SMTPHostValidation = VALIDATE_NONE         ' Optional, default = VALIDATE_HOST_DNS
        .EmailAddressValidation = VALIDATE_SYNTAX   ' Optional, default = VALIDATE_SYNTAX
        .Delimiter = ";"                            ' Optional, default = ";" (semicolon)

        ' **************************************************************************
        ' Basic properties for sending email
        ' **************************************************************************
        .SMTPHost = paramSendMail_SMTPHost                                ' Required the fist time, optional thereafter
        .From = lSendMail.From                                   ' Required the fist time, optional thereafter
        .FromDisplayName = lSendMail.FromDisplayName             ' Optional, saved after first use
        
        .Recipient = "ojeer.s@bia-paris.fr"
        'lSendMail.Recipient                          ' Required, separate multiple entries with delimiter character
        .RecipientDisplayName = lSendMail.RecipientDisplayName    ' Optional, separate multiple entries with delimiter character
        .CcRecipient = lSendMail.CcRecipient                      ' Optional, separate multiple entries with delimiter character
        .CcDisplayName = lSendMail.CcDisplayName                    ' Optional, separate multiple entries with delimiter character
        .BccRecipient = lSendMail.From  'txtBcc                                     ' Optional, separate multiple entries with delimiter character
        '.ReplyToAddress = txtFrom.Text                             ' Optional, used when different than 'From' address
        .Subject = lSendMail.Subject                                ' Optional
        .Message = lSendMail.Message                                ' Optional
        .Attachment = lSendMail.Attachment                          ' Optional, separate multiple entries with delimiter character

        ' **************************************************************************
        ' Additional Optional properties, use as required by your application / environment
        ' **************************************************************************
        .AsHTML = lSendMail.AsHTML                   ' Optional, default = FALSE, send mail as html or plain text
       ' .ContentBase = ""                           ' Optional, default = Null String, reference base for embedded links
       ' .EncodeType = MyEncodeType                  ' Optional, default = MIME_ENCODE
       ' .Priority = etPriority                      ' Optional, default = PRIORITY_NORMAL
        .Receipt = True    'bReceipt                         ' Optional, default = FALSE
       ' .UseAuthentication = bAuthLogin             ' Optional, default = FALSE
       ' .UsePopAuthentication = bPopLogin           ' Optional, default = FALSE
       ' .UserName = txtUserName                     ' Optional, default = Null String
       ' .Password = txtPassword                     ' Optional, default = Null String, value is NOT saved
       ' .POP3Host = txtPopServer
       ' .MaxRecipients = 100                        ' Optional, default = 100, recipient count before error is raised
        
        ' **************************************************************************
        ' Advanced Properties, change only if you have a good reason to do so.
        ' **************************************************************************
         .ConnectTimeout = 10                      ' Optional, default = 10
         .ConnectRetry = 5                         ' Optional, default = 5
         .MessageTimeout = 60                      ' Optional, default = 60
         .PersistentSettings = True                ' Optional, default = TRUE
         .SMTPPort = 25                            ' Optional, default = 25

        ' **************************************************************************
        ' OK, all of the properties are set, send the email...
        ' **************************************************************************
        ' .Connect                                  ' Optional, use when sending bulk mail
        
        .Send                                       ' Required
        
        ' .Disconnect                               ' Optional, use when sending bulk mail
    '    txtServer.Text = .SMTPHost                  ' Optional, re-populate the Host in case
                                                    ' MX look up was used to find a host    End With
    End With
Set poSendMail = Nothing


End Sub

Public Sub Monitor(lSendMail As typeSendMail)
'-----------------------------------------------------------------

'-----------------------------------------------------------------
'émetteur par défaut

If lSendMail.From = "" Then lSendMail.From = paramSendMail_From
If lSendMail.FromDisplayName = "" Then lSendMail.FromDisplayName = frmElp_Caption

'-----------------------------------------------------------------
'Destinataire : par défaut application / service => adresses EMail des destinataires
If lSendMail.Recipient = "" Then lSendMail.Recipient = Exchange_Distribution(lSendMail.RecipientDisplayName, lSendMail.FromDisplayName)
'-----------------------------------------------------------------
'Copie : par défaut application / service => adresses EMail des destinataires
If lSendMail.CcRecipient = "" Then
    If lSendMail.CcDisplayName <> "" Then lSendMail.CcRecipient = Exchange_Distribution(lSendMail.CcDisplayName, lSendMail.FromDisplayName)
End If

'-----------------------------------------------------------------
'Debug.Print "srvendMail"; lSendMail.Recipient
If paramEnvironnement = constTest Then
    lSendMail.Recipient = Exchange_Distribution("TEST", "TEST")
    lSendMail.From = lSendMail.Recipient
    lSendMail.CcRecipient = Exchange_Distribution("TEST_CC", "")
End If

'-----------------------------------------------------------------
'$JPL 10-06-2013
Dim K As Integer
For K = 0 To arrMail_Nb
    If lSendMail.RecipientDisplayName = arrMail_K1(K) And lSendMail.FromDisplayName = arrMail_K2(K) Then
        lSendMail.Recipient = lSendMail.Recipient & ";" & arrMail_Memo(K)
    End If
Next K

'-----------------------------------------------------------------
Route lSendMail
'Call ecrit_Log(lSendMail)
'-----------------------------------------------------------------
'Reset les champs
lSendMail.From = ""
lSendMail.FromDisplayName = ""
lSendMail.Recipient = ""
lSendMail.RecipientDisplayName = ""
lSendMail.CcRecipient = ""
lSendMail.CcDisplayName = ""
lSendMail.Subject = ""
lSendMail.Message = ""
lSendMail.Attachment = ""
lSendMail.AsHTML = False

End Sub
Public Sub ecrit_Log(lSendMail As typeSendMail)
Dim FicSortie As Long
Dim ficName As String
Dim newDirectory As String
Dim xMemo As String

On Error Resume Next
    If InStr(lSendMail.Subject, "BLOOMBERG") > 0 And InStr(lSendMail.Subject, "négatif") > 0 Then
        Exit Sub
    End If
    xMemo = paramServer("\\LOGMAILS\")
    If IsNull(xMemo) Then
        Exit Sub
    End If
    If Dir(Trim(xMemo), vbDirectory) = "" Then
        MkDir Trim(xMemo)
    End If
    Select Case UCase(App.exeName)
        Case "BIA_SAB":
            ficName = Trim(xMemo) & "BIA_SAB_Mail_"
        Case "BIA_SYSTEM":
            ficName = Trim(xMemo) & "BIA_SYSTEM_Mail_"
        Case "BIA_SWIFT":
            ficName = Trim(xMemo) & "BIA_SWIFT_Mail_"
        Case "BIA_AUDIT":
            ficName = Trim(xMemo) & "BIA_AUDIT_Mail_"
        Case Else: Exit Sub
    End Select
    
    ficName = ficName & CStr(Year(Now)) & Mid(CStr(100 + Month(Now)), 2) & Mid(CStr(100 + Day(Now)), 2) & ".log"
    FicSortie = FreeFile
    If Dir(ficName, vbNormal) = "" Then
        Open ficName For Output As #FicSortie
        Close #FicSortie
    End If
    Open ficName For Append As #FicSortie
    Print #FicSortie, "Heure = " & Format(Time, "hh:nn:ss") & " Destinataires = " & lSendMail.Recipient & _
    " Objet = " & lSendMail.Subject & " Joints = " & lSendMail.Attachment
    Close #FicSortie

End Sub

Public Sub Email_Standard(lRecipient As String, lSubject As String, lMessage As String, blnAsHTML As Boolean, lAttachment As String)
Dim wSendMail As typeSendMail
Dim bgColor As String
Dim wPath As String
Dim xMontant As String

wSendMail.From = ""
wSendMail.FromDisplayName = currentUser.Id
wSendMail.Recipient = lRecipient

wSendMail.Subject = lSubject
wSendMail.Attachment = lAttachment
wSendMail.Message = lMessage

wSendMail.AsHTML = blnAsHTML

srvSendMail.Monitor wSendMail

End Sub

Public Sub Email_Alerte(lAppli As String, lUnit As String, lSubject As String, lMessage As String, blnAsHTML As Boolean, lAttachment As String)
Dim wSendMail As typeSendMail
Dim bgColor As String
Dim wPath As String
Dim xMontant As String

wSendMail.From = ""
wSendMail.FromDisplayName = lAppli
wSendMail.Recipient = ""
wSendMail.RecipientDisplayName = lUnit

wSendMail.Subject = lSubject
wSendMail.Attachment = lAttachment
wSendMail.Message = lMessage

wSendMail.AsHTML = blnAsHTML
srvSendMail.Monitor wSendMail

End Sub

Public Function Exchange_Distribution(lK1 As String, lK2 As String) As String
Dim X As String, K As Integer
If lK2 = "" Then
    X = Trim(lK1)
Else
    X = Trim(lK1) & "." & Trim(lK2)
End If

X = "select * from " & paramIBM_Library_SABSPE & ".YSSIMEL0" _
     & " where SSIMELNAT = '@' and SSIMELUIDX = '" & X & "'"
Set rsSab_Local = cnsab.Execute(X)

If Not rsSab_Local.EOF Then
    Exchange_Distribution = Trim(rsSab_Local("SSIMELINFO"))
Else
    Exchange_Distribution = ""
End If

End Function

