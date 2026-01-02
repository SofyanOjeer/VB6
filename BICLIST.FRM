VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmBicList 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   Caption         =   "Correspondants :  Recherche"
   ClientHeight    =   4320
   ClientLeft      =   60
   ClientTop       =   645
   ClientWidth     =   9330
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4320
   ScaleWidth      =   9330
   Begin VB.ListBox lstX 
      Height          =   3180
      Left            =   2400
      Sorted          =   -1  'True
      TabIndex        =   18
      Top             =   840
      Visible         =   0   'False
      Width           =   4005
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   8880
      Picture         =   "BICLIST.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   0
      Width           =   400
   End
   Begin VB.PictureBox picList 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      Height          =   3255
      Left            =   2520
      ScaleHeight     =   3195
      ScaleWidth      =   6555
      TabIndex        =   5
      Top             =   840
      Width           =   6615
   End
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5760
      TabIndex        =   4
      Top             =   0
      Width           =   3105
   End
   Begin VB.TextBox txtSelect 
      Height          =   300
      Left            =   3600
      TabIndex        =   3
      Top             =   0
      Width           =   1815
   End
   Begin VB.CommandButton cmdPageNext 
      Caption         =   ">>"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2160
      TabIndex        =   2
      Top             =   0
      Width           =   960
   End
   Begin VB.CommandButton cmdPagePrior 
      Caption         =   "<<"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1200
      TabIndex        =   1
      Top             =   0
      Width           =   960
   End
   Begin VB.CommandButton cmdScan 
      Caption         =   "&Recherche"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1200
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   3375
      Left            =   0
      TabIndex        =   6
      Top             =   720
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   5953
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   16711680
      TabCaption(0)   =   "B I C "
      TabPicture(0)   =   "BICLIST.frx":0102
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraBic"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "I BA N"
      TabPicture(1)   =   "BICLIST.frx":011E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraIban"
      Tab(1).ControlCount=   1
      Begin VB.Frame fraIban 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   2805
         Left            =   -74880
         TabIndex        =   8
         Top             =   480
         Width           =   1695
         Begin VB.OptionButton optIbanDevise 
            Caption         =   "Devise"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   720
            Value           =   -1  'True
            Width           =   1400
         End
         Begin VB.OptionButton optIbanBicId 
            Caption         =   "Auprès de"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   1400
         End
      End
      Begin VB.Frame fraBic 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   2685
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   1575
         Begin VB.OptionButton optIntitulé 
            Caption         =   "intitulé"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   2280
            Width           =   1400
         End
         Begin VB.OptionButton optBicId 
            Caption         =   "Correspondant"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Value           =   -1  'True
            Width           =   1400
         End
         Begin VB.OptionButton optSwiftId 
            Caption         =   "code Swift"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   640
            Width           =   1400
         End
         Begin VB.OptionButton optTélexId 
            Caption         =   "code Télex"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   1080
            Width           =   1400
         End
         Begin VB.OptionButton optVilleId 
            Caption         =   "code Ville"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   1440
            Width           =   1400
         End
         Begin VB.OptionButton optPaysId 
            Caption         =   "code Pays"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   1840
            Width           =   1400
         End
      End
   End
   Begin VB.Menu txtBdfmnu 
      Caption         =   "txtBdfmnu"
      Visible         =   0   'False
      Begin VB.Menu txtBdfmnuRecherche 
         Caption         =   "Recherche Banque de France"
      End
   End
End
Attribute VB_Name = "frmBicList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim arrTag() As Boolean, arrTagNb As Integer, lstErrClear As Boolean
Dim X As String, I As Integer
Dim Msg As String
Dim mBicid As String

Dim recBic As typeBic
Dim recBicIban As typeBicIban
Dim oldBic As typeBic
Dim oldBicIban As typeBicIban

Private mMethod As String * 12

Private Const picListNb = 10
Private picListNbMin As Integer
Private picListNbMax As Integer
Private picListIndex As Integer

Private Sub lstX_Click()
Dim I As Integer
I = InStr(1, lstX.Text, " =:")
If I > 0 Then txtSelect = mId$(lstX.Text, I + 4, 2): txtSelect.SetFocus
lstX.Visible = False
frmBiclist_Scan
End Sub
Private Sub lstX_LostFocus()
lstX.Visible = False
End Sub



'---------------------------------------------------------
Public Sub lstX_Display()
'---------------------------------------------------------

lstX.Visible = True
Call LstDictio(19, lstX, "DicLib")
lstX.SetFocus

End Sub


'---------------------------------------------------------
Private Sub fraIban_Enabled()
'---------------------------------------------------------
Dim I As Integer

cmdPageNext.Caption = Format$(picListNbMax, "####") & "..>>"
If picListNbMax < arrBicIbanNb Or arrBicIbanNext Then
    cmdPageNext.Enabled = True
Else
    cmdPageNext.Enabled = False
End If

cmdPagePrior.Caption = "<<..." & Format$(picListNbMin, "####")
If picListNbMin > 1 Then
    cmdPagePrior.Enabled = True
Else
    cmdPagePrior.Enabled = False
End If
End Sub
'---------------------------------------------------------
Private Sub fraBic_Enabled()
'---------------------------------------------------------
Dim I As Integer

cmdPageNext.Caption = Format$(picListNbMax, "####") & "..>>"
If picListNbMax < arrBicNb Or arrBicNext Then
    cmdPageNext.Enabled = True
Else
    cmdPageNext.Enabled = False
End If

cmdPagePrior.Caption = "<<..." & Format$(picListNbMin, "####")
If picListNbMin > 1 Then
    cmdPagePrior.Enabled = True
Else
    cmdPagePrior.Enabled = False
End If
End Sub


Public Sub arrTag_Set(ByVal B As Boolean)
For I = 0 To arrTagNb: arrTag(I) = B: Next I
End Sub
Public Sub frmBicList_Clear()
txtSelect = ""
frmBiclist_Reset
End Sub

Public Sub frmBiclist_Scan()
picList.Cls
Select Case SSTab.Tab
    Case Is = 0: fraBic_Scan
    Case Is = 1: fraIban_Scan
End Select
End Sub
Public Sub fraBic_Scan()
Dim Xmax As String
arrBicNb = 0: arrBicIndex = 0

recBicInit recBic
arrBic(0) = recBic
X = Trim(txtSelect)
Xmax = X & "99999999999"

If optBicId Then
    mMethod = "SnapP0"
    recBic.BicId = Trim(txtSelect)
    arrBic(0).BicId = Xmax
End If

If optSwiftId Then
    mMethod = "SnapL1"
    recBic.SwiftId = Trim(txtSelect)
    arrBic(0).SwiftId = Xmax
End If

If optTélexId Then
    mMethod = "SnapL2"
    recBic.TélexId = Trim(txtSelect)
    arrBic(0).TélexId = Xmax
End If

If optVilleId Then
    mMethod = "SnapL4"
    recBic.VilleId = Trim(txtSelect)
    arrBic(0).VilleId = Xmax
    arrBic(0).PaysId = "99"
End If

If optPaysId Then
    mMethod = "SnapL3"
    recBic.PaysId = Trim(txtSelect)
    arrBic(0).PaysId = Xmax
    arrBic(0).BicId = "99999999999"
End If

If optIntitulé Then
    mMethod = "SnapL5"
    recBic.Intitulé = Trim(txtSelect)
    arrBic(0).Intitulé = Xmax
End If

recBic.Method = mMethod
arrBic(0).Method = mMethod

If IsNull(srvBicMon(recBic)) Then
    Call fraBic_Display
End If

End Sub
Public Sub fraIban_Scan()
Dim Xmax As String

arrBicIbanNb = 0
arrBicIbanIndex = 0

recBicIbanInit recBicIban
recBicIban.BicId = mBicid
arrBicIban(0) = recBicIban
X = Trim(txtSelect)
Xmax = X & "99999999999"

If optIbanBicId Then
    mMethod = "SnapL1"
    recBicIban.IbanBicId = Trim(txtSelect)
    arrBicIban(0).IbanBicId = Xmax
End If

If optIbanDevise Then
    mMethod = "SnapP0"
    recBicIban.IbanDevise = Format$(Val(txtSelect), "000")
    arrBicIban(0).IbanDevise = Xmax
    arrBicIban(0).IbanBicId = "99999999999"
End If

recBicIban.Method = mMethod
arrBicIban(0).Method = mMethod

If IsNull(srvBicIbanMon(recBicIban)) Then
    Call fraIban_Display
End If

End Sub

Private Sub cmdPageNext_Click()
Select Case SSTab.Tab
    Case Is = 0
        If arrBicIndex < arrBicNb Then
            fraBic_Display
        Else
            arrBicIndex = arrBicNb
            If arrBicNext Then fraBic_dbNext
        End If
    Case Is = 1
        If arrBicIbanIndex < arrBicIbanNb Then
            fraIban_Display
        Else
            arrBicIbanIndex = arrBicIbanNb
            If arrBicIbanNext Then fraIban_dbNext
        End If

End Select
End Sub

'---------------------------------------------------------
Private Sub fraIban_dbNext()
'---------------------------------------------------------
Mid$(recBicIban.Method, 7, 1) = "+"
If IsNull(srvBicIbanMon(recBicIban)) Then Call fraIban_Display
End Sub
'---------------------------------------------------------
Private Sub fraBic_dbNext()
'---------------------------------------------------------
Mid$(recBic.Method, 7, 1) = "+"
If IsNull(srvBicMon(recBic)) Then Call fraBic_Display
End Sub


Private Sub cmdPagePrior_Click()
Select Case SSTab.Tab
    Case Is = 0
        arrBicIndex = picListNbMin - picListNb - 1
        fraBic_Display
    Case Is = 1
        arrBicIbanIndex = picListNbMin - picListNb - 1
        fraIban_Display
End Select
End Sub

Private Sub cmdPrint_Click()
Dim X As String * 100
Select Case SSTab.Tab
    Case Is = 0: X = Format$(1, "000000") & Format$(arrBicNb, "000000")
    Case Is = 1: X = Format$(0, "000000") & Format$(0, "000000")
End Select
Mid$(X, 13, 12) = "frmBicList  "
prtBicX X
Form_Activate
End Sub
Private Sub cmdQuit_Click()
arrBic(0).Method = "X"
arrBicIban(0).Method = "X"
Unload Me
End Sub
Private Sub cmdScan_Click()
frmBiclist_Scan
End Sub
Public Sub frmBiclist_Reset()
picList.Cls
cmdPagePrior.Enabled = False
cmdPageNext.Enabled = False
arrTag_Set False
lstErrClear = True
Call lstErr_Clear(lstErr, txtSelect, "sélection")
usrColor_Set
picList.BackColor = &HE0E0E0
End Sub

Private Sub Form_Activate()
Set XForm = Me
MousePointer = 10
'MouseIcon = frmElp.imgBiaMouse
txtSelect.SetFocus
prtShow = False
recBic = oldBic
recBicIban = oldBicIban
mBicid = recBicIban.BicId
SSTab.Enabled = True
txtSelect.Enabled = True
Select Case Trim(recBic.Method)
    Case "Bic"
                SSTab.Tab = 0
                txtSelect = recBic.BicId
                If Trim(txtSelect) = "" Then
                    optPaysId.Value = "1"
                    lstX_Display
                Else
                    cmdScan_Click
                End If
    Case "IbanDevise"
                SSTab.Tab = 1
'                mBicid = recBicIban.BicId
                txtSelect = recBicIban.IbanDevise
                SSTab.Enabled = False
                txtSelect.Enabled = False
                cmdScan_Click
    Case Else:  SSTab.Tab = 1
                mBicid = recBicIban.BicId
'               txtSelect = recBicIban.IbanBicId
                cmdScan_Click
End Select
End Sub
Private Sub Form_Deactivate()
prtShow = True
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
 '  Case Is = 13: SendKeys "{TAB}"
    Case Is = 27: cmdQuit_Click
    Case Is = 34: cmdPageNext_Click
    Case Is = 33: cmdPagePrior_Click
    Case Is = 44: KeyCode = 0: frmElpPrt.prtScreen
End Select
End Sub
Private Sub Form_Load()
Set XForm = Me
Call MeInit(arrTagNb)
ReDim arrTag(arrTagNb + 1)
Call frmBicList_Clear
oldBic = arrBic(0)
oldBicIban = arrBicIban(0)
End Sub
Public Sub Msg_Rcv(X As String)
Msg = X
SSTab.Tab = mId$(X, 35, 1)
If SSTab.Tab = 0 Then
    txtSelect = mId$(X, 36, 11)
Else
    txtSelect = mId$(X, 47, 11)
    mBicid = mId$(X, 36, 11)
End If
cmdScan_Click
End Sub

Private Sub fraBic_Display()
Dim K As Integer, K1 As Integer, K2 As Integer, K3 As Integer

picList.Cls

If arrBicIndex < 0 Then
    arrBicIndex = 0
End If
If arrBicNb <= 0 Then
    X = "Néant ( " & Now & " )"
    Call lstErr_Clear(lstErr, txtSelect, X)
Else
    picListNbMin = arrBicIndex + 1
    Call lstErr_Clear(lstErr, txtSelect, "click = sélection")
End If
    
    K1 = 0
    K2 = picListNb - 1
    K3 = 1
For K = K1 To K2 Step K3

    If arrBicIndex >= arrBicNb Then
    Exit For
'       piclist.Visible = False
    Else
       arrBicIndex = arrBicIndex + 1
       picList.CurrentY = K * picLineHeight + 20
       picList.CurrentX = 50
       picList.ForeColor = libUsr.ForeColor
       picList.FontBold = False
       picList.Print arrBic(arrBicIndex).BicId;
       
       picList.CurrentX = 1500
       picList.Print arrBic(arrBicIndex).Intitulé;
        
    End If
Next K

picListNbMax = arrBicIndex
Call fraBic_Enabled
End Sub
Private Sub fraIban_Display()
Dim K As Integer, K1 As Integer, K2 As Integer, K3 As Integer

picList.Cls

If arrBicIbanIndex < 0 Then
    arrBicIbanIndex = 0
End If
If arrBicIbanNb <= 0 Then
    X = "Néant ( " & Now & " )"
    Call lstErr_Clear(lstErr, txtSelect, X)
Else
    picListNbMin = arrBicIbanIndex + 1
    Call lstErr_Clear(lstErr, txtSelect, "click = sélection")
End If
    
    K1 = 0
    K2 = picListNb - 1
    K3 = 1
For K = K1 To K2 Step K3

    If arrBicIbanIndex >= arrBicIbanNb Then
'       piclist.Visible = False
    Else
       arrBicIbanIndex = arrBicIbanIndex + 1
       picList.CurrentY = K * picLineHeight + 20
       picList.ForeColor = libUsr.ForeColor
       picList.FontBold = False
       picList.CurrentX = 50
       picList.Print arrBicIban(arrBicIbanIndex).BicId;
       picList.CurrentX = 1200
       picList.Print arrBicIban(arrBicIbanIndex).IbanDevise;
       picList.CurrentX = 1700
       picList.Print arrBicIban(arrBicIbanIndex).IbanBicId;
       picList.CurrentX = 3000
       picList.Print arrBicIban(arrBicIbanIndex).IbanPays & " " & arrBicIban(arrBicIbanIndex).IbanCompte;
        
    End If
Next K

picListNbMax = arrBicIbanIndex
Call fraIban_Enabled
End Sub

Private Sub optPaysId_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lstX_Display
End Sub

Private Sub picList_Click()
Select Case SSTab.Tab
    Case Is = 0
        If picListIndex > 0 And picListIndex <= arrBicNb Then
            arrBic(0) = arrBic(picListIndex)
            arrBic(0).Method = "Ok"
            arrBicIban(0).Method = "X"
            Unload Me
        Else
            Call lstErr_Clear(lstErr, txtSelect, "? Sélection")
            Exit Sub
        End If
    Case Is = 1
        If picListIndex > 0 And picListIndex <= arrBicIbanNb Then
            arrBicIban(0) = arrBicIban(picListIndex)
            arrBicIban(0).Method = "Ok"
            arrBic(0).Method = "X"
            Unload Me
        Else
            Call lstErr_Clear(lstErr, txtSelect, "? Sélection")
            Exit Sub
        End If
End Select
End Sub
Private Sub picList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
picListIndex = picListNbMin + (Y - 20) \ picLineHeight
End Sub
Private Sub SSTab_Click(PreviousTab As Integer)
txtSelect = ""
picList.Cls
End Sub

Private Sub txtSelect_GotFocus()
Call txt_GotFocus(txtSelect)
End Sub
Sub txt_GotFocus(C As Control)
'frmBiclist_Control
C.ForeColor = txtUsr.ForeColor
C.BackColor = focusUsr.BackColor
End Sub

Private Sub txtSelect_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub

