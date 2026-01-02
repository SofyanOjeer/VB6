Attribute VB_Name = "vbSendMail_Std"
Option Explicit

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

Public Sub vbSendMail_Test()
Dim wSendMail As typeSendMail
Dim bgColor As String
Dim iCol As Integer, K As Integer

wSendMail.From = "bia_info@bia-paris.fr"
wSendMail.Recipient = "loulergue.jp@bia-paris.fr"

bgColor = "MAGENTA"
wSendMail.Subject = "sujet = TEST"
wSendMail.Attachment = ""
wSendMail.Message = "<body bgcolor=" & Asc34 & bgColor & Asc34 & ">" _
                    & "<FONT face=" & Asc34 & prtFontName_Arial & Asc34 & ">" _
                    & htmlFontColor("BLUE") & "Ceci est un Test"

wSendMail.AsHTML = True

vbSendMail_Std.Monitor wSendMail

 

End Sub

Public Sub Monitor(lSendMail As typeSendMail)
'-----------------------------------------------------------------


'-----------------------------------------------------------------
Route lSendMail
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

Public Sub Route(lSendMail As typeSendMail)

Set poSendMail = New clsSendMail

On Error Resume Next
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
        
        .Recipient = lSendMail.Recipient                          ' Required, separate multiple entries with delimiter character
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

