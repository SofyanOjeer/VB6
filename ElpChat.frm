VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmElpChat 
   AutoRedraw      =   -1  'True
   Caption         =   "Test"
   ClientHeight    =   6372
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   9420
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6372
   ScaleWidth      =   9420
   Begin TabDlg.SSTab SSTab1 
      Height          =   5535
      Left            =   0
      TabIndex        =   13
      Top             =   600
      Width           =   9255
      _ExtentX        =   16320
      _ExtentY        =   9758
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "x"
      TabPicture(0)   =   "ElpChat.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "txtReceive"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtSend"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "chkAuto"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Options"
      TabPicture(1)   =   "ElpChat.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraOption"
      Tab(1).Control(1)=   "fraTimer"
      Tab(1).ControlCount=   2
      Begin VB.CheckBox chkAuto 
         Caption         =   "Auto "
         Height          =   375
         Left            =   2880
         TabIndex        =   24
         Top             =   120
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.Frame fraTimer 
         Height          =   1815
         Left            =   -74640
         TabIndex        =   20
         Top             =   600
         Width           =   8655
         Begin VB.TextBox txtFlash 
            Height          =   285
            Left            =   2520
            MaxLength       =   4
            TabIndex        =   23
            Top             =   1080
            Width           =   615
         End
         Begin VB.CheckBox chkFlash 
            Caption         =   "Flash (secondes)"
            Height          =   255
            Left            =   720
            TabIndex        =   22
            Top             =   1080
            Width           =   1575
         End
         Begin VB.TextBox txtWait 
            Height          =   285
            Left            =   2520
            MaxLength       =   4
            TabIndex        =   21
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame fraOption 
         Height          =   2895
         Left            =   -74640
         TabIndex        =   14
         Top             =   2520
         Width           =   8655
         Begin VB.CommandButton cmdOption 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Valider"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   500
            Left            =   4800
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   360
            Width           =   1200
         End
         Begin VB.TextBox txtKey 
            Height          =   375
            Left            =   2040
            TabIndex        =   8
            Top             =   2400
            Width           =   5295
         End
         Begin VB.TextBox txtPass2 
            Height          =   375
            Left            =   4800
            TabIndex        =   7
            Top             =   1680
            Width           =   1815
         End
         Begin VB.TextBox txtUser2 
            Height          =   375
            Left            =   2040
            TabIndex        =   6
            Top             =   1680
            Width           =   1815
         End
         Begin VB.TextBox txtK2 
            Height          =   375
            Left            =   2040
            TabIndex        =   3
            Top             =   360
            Width           =   1815
         End
         Begin VB.TextBox txtPass1 
            Height          =   375
            Left            =   4800
            TabIndex        =   5
            Top             =   960
            Width           =   1815
         End
         Begin VB.TextBox txtUser1 
            Height          =   375
            Left            =   2040
            TabIndex        =   4
            Top             =   960
            Width           =   1815
         End
         Begin VB.Label lblK2 
            Caption         =   "K2"
            Height          =   375
            Left            =   360
            TabIndex        =   18
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label lblKey 
            Caption         =   "Cipher"
            Height          =   375
            Left            =   360
            TabIndex        =   17
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label lblUser2 
            Caption         =   "Utilisateur 2"
            Height          =   375
            Left            =   360
            TabIndex        =   16
            Top             =   1800
            Width           =   1215
         End
         Begin VB.Label lblUser1 
            Caption         =   "Utilisateur 1"
            Height          =   375
            Left            =   360
            TabIndex        =   15
            Top             =   960
            Width           =   1215
         End
      End
      Begin VB.TextBox txtSend 
         Height          =   2220
         IMEMode         =   3  'DISABLE
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   0
         Top             =   3120
         Visible         =   0   'False
         Width           =   8400
      End
      Begin VB.TextBox txtReceive 
         Enabled         =   0   'False
         Height          =   2580
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   480
         Visible         =   0   'False
         Width           =   8400
      End
   End
   Begin VB.TextBox txtPassword 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   2400
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   0
      Width           =   1935
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00C0FFC0&
      Caption         =   "&Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   0
      Width           =   1200
   End
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   432
      Left            =   6720
      TabIndex        =   11
      Top             =   0
      Width           =   2145
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   500
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   0
      Width           =   500
   End
   Begin VB.CommandButton cmdContext 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Abandonner"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   0
      Width           =   1200
   End
End
Attribute VB_Name = "frmElpChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim arrTag() As Boolean, arrTagNb As Integer
Dim lastActiveControl_Name  As String, currentActiveControl_Name As String, currentAction As String
Dim MouseMoveActiveControl_Name  As String, MouseMoveActiveControl As typeUsrColor
Dim blnMsgBox_Quit As Boolean, blnAddNew As Boolean, blnGlobalControl As Boolean, blnControl As Boolean
Dim X As String, I As Integer, Msg As String, valX As String, X1 As String, V As Variant
Dim xFileName_Receive As String, xFileName_Send As String


Dim paramElpChat_Key As String
Dim xxxSend As String, xxxReceive As String
Dim mElpTable_Chat As typeElpTable, zElpTable_Chat As typeElpTable
Dim x30 As String * 30, X60 As String * 60
Dim wPass As String, w30 As String * 30, w60 As String * 60

''blnElpChat_Password As Boolean
Public Sub MouseMoveActiveControl_Reset()
For Each xobj In Me.Controls
    If MouseMoveActiveControl_Name = xobj.Name Then
        MouseMoveActiveControl_Name = ""
         If TypeOf xobj Is CommandButton Or TypeOf xobj Is ListBox Then
           xobj.BackColor = MouseMoveActiveControl.BackColor
        Else
            xobj.ForeColor = MouseMoveActiveControl.ForeColor
        End If
        Exit For
    End If
Next xobj
End Sub

Public Sub MouseMoveActiveControl_Set(C As Control)
If MouseMoveActiveControl_Name <> C.Name Then
    MouseMoveActiveControl_Reset
    If Not C.Enabled Then
        MouseMoveActiveControl_Name = ""
    Else
        MouseMoveActiveControl_Name = C.Name
        If TypeOf C Is CommandButton Or TypeOf C Is ListBox Then

            MouseMoveActiveControl.BackColor = C.BackColor
            C.BackColor = MouseMoveUsr.BackColor
        Else
            MouseMoveActiveControl.ForeColor = C.ForeColor
            C.ForeColor = MouseMoveUsr.ForeColor
        End If
    End If
End If

End Sub

Public Sub cmdContext_Quit()
blnControl = False
If blnMsgBox_Quit Then
    X = MsgBox("Voulez-vous réellement abandonner?", vbYesNo + vbQuestion + vbDefaultButton2, "Saisie non enregistrée")
 Else
    X = vbYes
 End If
 'If X = vbYes Then blnElpChat_Password = False: Unload Me

End Sub


Public Sub cmdContext_Return()
If Not blnElpChat_Password Then
    Form_Init
    If Not blnElpChat_Password Then Unload Me
End If

'SendKeys "{TAB}"

End Sub






'-------------------------------------------------------
Sub txt_GotFocus(C As Control)
'-------------------------------------------------------
'lstErr.Clear
currentActiveControl_Name = C.Name
C.ForeColor = txtUsr.ForeColor
C.BackColor = focusUsr.BackColor
End Sub
'-------------------------------------------------------
Sub txt_LostFocus(C As Control)
'-------------------------------------------------------
'lstErr.Clear
arrTag(Val(C.Tag)) = True
C.ForeColor = txtUsr.ForeColor
C.BackColor = txtUsr.BackColor
End Sub

Private Sub cmdContext_Click()
Select Case cmdContext.Caption
    Case Is = constcmdRechercher
    Case Is = constcmdAbandonner: cmdContext_Quit
End Select

End Sub

Private Sub cmdOk_Click()

X = Dir(xFileName_Receive)
If X <> "" Then Kill xFileName_Receive

X = Dir(xFileName_Send)
If X <> "" Then Kill xFileName_Send

X = Trim(txtSend)
If X <> "" Then
    Open xFileName_Send For Output As #1
    xxxSend = ElpCipher_C(X, paramElpChat_Key)

    Print #1, xxxSend
    Close #1
End If
cmdQuit_Click

End Sub

Private Sub cmdOption_Click()
Dim mK2 As Integer
mK2 = CInt(Val(txtK2))

xElpTable_Chat = zElpTable_Chat
xElpTable_Chat.Method = "AddNew"
xElpTable_Chat.Name = "chat_monitor"
xElpTable_Chat.K1 = "Monitor"
wPass = Trim(txtPassword)

xElpTable_Chat.K2 = Format$(mK2, "000000000000")

xElpTable_Chat.Memo = Space$(90)

w30 = Trim(txtUser1)
x30 = ElpCipher_C(w30, wPass)
Mid$(xElpTable_Chat.Memo, 1, 30) = x30

w60 = Trim(txtPass1)
X60 = ElpCipher_C(w60, wPass)
Mid$(xElpTable_Chat.Memo, 31, 60) = X60
V = dbElpTable_Update(xElpTable_Chat)
If Not IsNull(V) Then Exit Sub

xElpTable_Chat.K2 = Format$(mK2 + 1, "000000000000")

xElpTable_Chat.Memo = Space$(90)

w30 = Trim(txtUser2)
x30 = ElpCipher_C(w30, wPass)
Mid$(xElpTable_Chat.Memo, 1, 30) = x30

w60 = Trim(txtPass2)
X60 = ElpCipher_C(w60, wPass)
Mid$(xElpTable_Chat.Memo, 31, 60) = X60
V = dbElpTable_Update(xElpTable_Chat)
If Not IsNull(V) Then Exit Sub
'-------------------------------------------------

xElpTable_Chat = zElpTable_Chat
xElpTable_Chat.Method = "AddNew"
xElpTable_Chat.Name = "chat_key"
xElpTable_Chat.K2 = Format$(mK2, "000000000000")

wPass = Trim(txtPass1)
xElpTable_Chat.Memo = Space$(180) & ElpCipher_C(Trim(txtKey), wPass)

w30 = Trim(txtUser1)
x30 = ElpCipher_C(w30, wPass)
Mid$(xElpTable_Chat.Memo, 1, 30) = x30


w30 = Trim(txtUser2)
x30 = ElpCipher_C(w30, wPass)
Mid$(xElpTable_Chat.Memo, 31, 30) = x30

wPass = Trim(txtKey)
w60 = Trim(txtUser1) & Trim(txtUser2)
X60 = ElpCipher_C(w60, wPass)
Mid$(xElpTable_Chat.Memo, 61, 60) = X60
w60 = Trim(txtUser2) & Trim(txtUser1)
X60 = ElpCipher_C(w60, wPass)
Mid$(xElpTable_Chat.Memo, 121, 60) = X60

V = dbElpTable_Update(xElpTable_Chat)
If Not IsNull(V) Then Exit Sub

xElpTable_Chat.K2 = Format$(mK2 + 1, "000000000000")

wPass = Trim(txtPass2)
xElpTable_Chat.Memo = Space$(180) & ElpCipher_C(Trim(txtKey), wPass)

w30 = Trim(txtUser2)
x30 = ElpCipher_C(w30, wPass)
Mid$(xElpTable_Chat.Memo, 1, 30) = x30


w30 = Trim(txtUser1)
x30 = ElpCipher_C(w30, wPass)
Mid$(xElpTable_Chat.Memo, 31, 30) = x30

wPass = Trim(txtKey)
w60 = Trim(txtUser2) & Trim(txtUser1)
X60 = ElpCipher_C(w60, wPass)
Mid$(xElpTable_Chat.Memo, 61, 60) = X60
w60 = Trim(txtUser1) & Trim(txtUser2)
X60 = ElpCipher_C(w60, wPass)
Mid$(xElpTable_Chat.Memo, 121, 60) = X60

V = dbElpTable_Update(xElpTable_Chat)
If Not IsNull(V) Then Exit Sub

End Sub

'---------------------------------------------------------
Private Sub cmdQuit_Click()
'---------------------------------------------------------
blnElpChat_Password = False
Unload Me

End Sub




Private Sub cmdContext_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdContext
End Sub

Private Sub cmdPrint_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Set cmdPrint
End Sub


'---------------------------------------------------------
Private Sub Form_Activate()
'---------------------------------------------------------
Set XForm = Me
Me.Caption = "Documentation"
End Sub

'---------------------------------------------------------
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'---------------------------------------------------------
Select Case KeyCode
    Case Is = 13: KeyCode = 0: cmdContext_Return
    Case Is = 27: cmdContext_Quit
    Case Is = 44: KeyCode = 0: frmElpPrt.prtScreen

End Select

End Sub


'---------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------
Set XForm = Me
Call MeInit(arrTagNb)
ReDim arrTag(arrTagNb + 1)


cmdReset

End Sub
'---------------------------------------------------------
Public Sub cmdReset()
'---------------------------------------------------------
usrColor_Set
cmdContext.Caption = constcmdAbandonner: blnMsgBox_Quit = False
arrTag_Set False
lstErr.Visible = False
cmdOk.Visible = False
fraOption.Visible = False

If paramElpChat_Receive = "" Then ElpChat_Init: blnElpChat_Auto = False

ElpChat_Timer "Stop"

SSTab1.Visible = False
txtReceive = "": txtReceive.Visible = False
txtSend = "": txtSend.Visible = False
blnElpChat_Password = False
txtPassword = ""
If blnElpChat_Auto Then
    chkAuto = "1"
Else
    chkAuto = "0"
End If

recElpTable_Init zElpTable_Chat
zElpTable_Chat.ID = "ElpChat"
zElpTable_Chat.K1 = "Key"
xElpTable_Chat = zElpTable_Chat
End Sub



'---------------------------------------------------------
Public Sub arrTag_Set(ByVal B As Boolean)
'---------------------------------------------------------
For I = 0 To arrTagNb: arrTag(I) = B: Next I
End Sub



Public Sub Msg_Rcv(X As String)
'---------------------------------------------------------
cmdReset

End Sub


Public Sub Msg_Snd(ByVal X As String)
End Sub




Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMoveActiveControl_Reset
End Sub

Public Sub Form_Init()
If Not blnElpChat_Password Then
     param_Init
    If Not blnElpChat_Password Then Exit Sub
End If

X = paramElpChat_Folder & "\Test\" & DSys & "_" & time_Hms & "_" & mId$(usrId, 2, 1)

Open X For Output As #1
Print #1, X
Close #1


SSTab1.Visible = True
txtPassword.Visible = False
xxxReceive = ""
xxxSend = ""
cmdOk.Visible = True
txtReceive = "": txtReceive.Visible = True: txtReceive.Enabled = False
txtSend = "": txtSend.Visible = True
txtWait = Format$(paramElpChat_Wait / 1000, "####")
If blnElpChat_Auto Then
    chkAuto = "1"
Else
    chkAuto = "0"
End If
txtFlash = Format$(paramElpChat_Flash / 1000, "####")
If paramElpChat_Flash <> 0 Then
    chkFlash = "1"
Else
    chkFlash = "0"
End If

xFileName_Receive = paramElpChat_Receive
X = Dir(xFileName_Receive)
If X <> "" Then
    Open xFileName_Receive For Input As #1
    Line Input #1, X
    Close #1
    txtReceive = ElpCipher_D(X, paramElpChat_Key)
    Close #1
End If

xFileName_Send = paramElpChat_Send
X = Dir(xFileName_Send)
If X <> "" Then
    Open xFileName_Send For Input As #1
    Line Input #1, X
    Close #1
    txtSend = ElpCipher_D(X, paramElpChat_Key)
End If

If txtSend.Enabled And txtSend.Visible Then txtSend.SetFocus
End Sub
'-------------------------------------------------'


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If chkAuto = "1" Then
    blnElpChat_Auto = True: ElpChat_Timer "Start"
Else
    blnElpChat_Auto = False
End If
paramElpChat_Wait = Val(txtWait) * 1000
If chkFlash = "1" Then
    paramElpChat_Flash = Val(txtFlash) * 1000
Else
    paramElpChat_Flash = 0
End If
blnElpChat_Password = False
End Sub

Private Sub txtFlash_KeyPress(KeyAscii As Integer)
Call num_KeyAscii(KeyAscii)

End Sub


Private Sub txtK2_GotFocus()
txt_GotFocus txtK2
End Sub


Private Sub txtK2_KeyPress(KeyAscii As Integer)
Call num_KeyAscii(KeyAscii)
End Sub


Private Sub txtK2_LostFocus()
txt_LostFocus txtK2
End Sub


Private Sub txtKey_GotFocus()
txt_GotFocus txtKey
End Sub


Private Sub txtKey_LostFocus()
txt_LostFocus txtKey
End Sub


Private Sub txtPass1_GotFocus()
txt_GotFocus txtPass1
End Sub


Private Sub txtPass1_LostFocus()
txt_LostFocus txtPass1
End Sub


Private Sub txtPass2_GotFocus()
txt_GotFocus txtPass2
End Sub


Private Sub txtPass2_LostFocus()
txt_LostFocus txtPass2
End Sub


Private Sub txtPassword_GotFocus()
txt_GotFocus txtPassword
End Sub

Private Sub txtPassword_LostFocus()
txt_LostFocus txtPassword

End Sub


Private Sub txtReceive_GotFocus()
txt_GotFocus txtReceive
End Sub

Private Sub txtReceive_LostFocus()
txt_LostFocus txtReceive
End Sub


Private Sub txtSend_GotFocus()
txt_GotFocus txtSend

End Sub


Private Sub txtSend_LostFocus()
txt_LostFocus txtSend
End Sub


Private Sub txtUser1_GotFocus()
txt_GotFocus txtUser1
End Sub

Private Sub txtUser1_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)

End Sub


Private Sub txtUser1_LostFocus()
txt_LostFocus txtUser1
End Sub

Private Sub txtUser2_GotFocus()
txt_GotFocus txtUser2
End Sub

Private Sub txtUser2_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub


Private Sub txtUser2_LostFocus()
txt_LostFocus txtUser2
End Sub



Public Function param_Init()
param_Init = "?"
wPass = Trim(txtPassword)
If wPass = "" Then Exit Function
mElpTable_Chat = zElpTable_Chat
mElpTable_Chat.Method = "Seek>="
mElpTable_Chat.Err = 0
w30 = Trim(usrId)
x30 = ElpCipher_C(w30, wPass)
Do
    mElpTable_Chat.Err = tableElpTable_Read(mElpTable_Chat)
    If mElpTable_Chat.Err = 0 Then
        If mElpTable_Chat.K1 <> zElpTable_Chat.K1 _
        Or mElpTable_Chat.ID <> zElpTable_Chat.ID Then
            mElpTable_Chat.Err = 9996
        Else
            If mId$(mElpTable_Chat.Memo, 1, 30) = x30 Then
                txtUser1 = ElpCipher_D(mId$(mElpTable_Chat.Memo, 1, 30), wPass)
                txtUser2 = ElpCipher_D(mId$(mElpTable_Chat.Memo, 31, 30), wPass)
                paramElpChat_Key = ElpCipher_D(mId$(mElpTable_Chat.Memo, 181, Len(mElpTable_Chat.Memo) - 180), wPass)
                txtKey = paramElpChat_Key
                
                paramElpChat_Send = paramElpChat_Folder & mId$(mElpTable_Chat.Memo, 61, 60)
                paramElpChat_Receive = paramElpChat_Folder & mId$(mElpTable_Chat.Memo, 121, 60)
                blnElpChat_Password = True
                param_Init = Null
                fraOption.Visible = ElpChat_Monitor(wPass)

                Exit Function
            Else
                mElpTable_Chat.Method = "Seek>"
            End If
        End If
    End If
Loop While mElpTable_Chat.Err = 0
blnElpChat_Password = False

Exit Function

Table_Error:
param_Init = V
Exit Function

Memo_Error:
param_Init = "Memo"
MsgBox recElpTable.ID & " : " & recElpTable.K1 & " : " & recElpTable.K2 & " : Mémo absent", vbCritical, "TfluxEspèces_Compta_gen"
Exit Function

Num_Error:
param_Init = "Num"
MsgBox recElpTable.ID & " : " & recElpTable.K1 & " : " & recElpTable.K2 & " : " & recElpTable.Memo & " :Mémo non numérique", vbCritical, "TfluxEspèces_Param_Init"

End Function

Private Sub txtWait_KeyPress(KeyAscii As Integer)
Call num_KeyAscii(KeyAscii)

End Sub



Public Sub ElpX()
'Form_KeyUp      Case Is = 48: frmElpChat_Show
'Timer1_Timer    If blnElpChat_Auto Then ElpChat_Timer "Auto"
'Mainsoc         ElpChat_Init
'Elpchat.bas     wpass=""


End Sub
