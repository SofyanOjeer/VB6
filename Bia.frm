VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmElp 
   AutoRedraw      =   -1  'True
   Caption         =   "BIA"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13875
   Icon            =   "Bia.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9495
   ScaleWidth      =   13875
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6960
      TabIndex        =   10
      Top             =   15
      Width           =   6825
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9090
      Left            =   60
      TabIndex        =   3
      Top             =   405
      Width           =   13830
      _ExtentX        =   24395
      _ExtentY        =   16034
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   14737632
      TabCaption(0)   =   "Applications"
      TabPicture(0)   =   "Bia.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fra0"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Services"
      TabPicture(1)   =   "Bia.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraServices"
      Tab(1).ControlCount=   1
      Begin VB.Frame fra0 
         BackColor       =   &H00E0E0E0&
         Height          =   8640
         Left            =   60
         TabIndex        =   9
         Top             =   360
         Width           =   13740
         Begin VB.ListBox lstMain 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2865
            Left            =   270
            Sorted          =   -1  'True
            TabIndex        =   16
            Top             =   1245
            Visible         =   0   'False
            Width           =   8325
         End
         Begin MSFlexGridLib.MSFlexGrid fgMain_App 
            Height          =   8100
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Visible         =   0   'False
            Width           =   6630
            _ExtentX        =   11695
            _ExtentY        =   14288
            _Version        =   393216
            Rows            =   1
            Cols            =   3
            RowHeightMin    =   350
            BackColor       =   16316664
            ForeColor       =   12582912
            BackColorFixed  =   14745568
            ForeColorFixed  =   16384
            BackColorSel    =   15794160
            BackColorBkg    =   16316664
            AllowBigSelection=   0   'False
            AllowUserResizing=   3
            FormatString    =   $"Bia.frx":047A
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSFlexGridLib.MSFlexGrid fgMain_App_X 
            Height          =   5715
            Left            =   7335
            TabIndex        =   15
            Top             =   270
            Visible         =   0   'False
            Width           =   6195
            _ExtentX        =   10927
            _ExtentY        =   10081
            _Version        =   393216
            Rows            =   1
            Cols            =   4
            RowHeightMin    =   350
            BackColor       =   16316664
            ForeColor       =   16384
            BackColorFixed  =   15399679
            ForeColorFixed  =   -2147483639
            BackColorSel    =   12648384
            BackColorBkg    =   16316664
            AllowBigSelection=   0   'False
            AllowUserResizing=   3
            FormatString    =   "<Application               |<doc   |<Libellé                                                           |<Programme        "
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Image imgFEU 
            DragMode        =   1  'Automatic
            Height          =   480
            Index           =   2
            Left            =   7740
            Picture         =   "Bia.frx":0526
            Top             =   7530
            Width           =   480
         End
         Begin VB.Image imgFEU 
            DragMode        =   1  'Automatic
            Height          =   480
            Index           =   1
            Left            =   7740
            Picture         =   "Bia.frx":0968
            Top             =   6810
            Width           =   480
         End
         Begin VB.Image imgFEU 
            DragMode        =   1  'Automatic
            Height          =   480
            Index           =   0
            Left            =   7740
            Picture         =   "Bia.frx":0DAA
            Top             =   6120
            Width           =   480
         End
         Begin VB.Label lblPrt_DeviceName 
            BackColor       =   &H00FFC0FF&
            Caption         =   "imprimante non définie"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Left            =   8415
            TabIndex        =   18
            Top             =   8070
            Width           =   4425
         End
         Begin VB.Image imgSocSignon 
            DragMode        =   1  'Automatic
            Height          =   1725
            Left            =   8400
            Stretch         =   -1  'True
            Top             =   6105
            Width           =   3795
         End
      End
      Begin VB.Frame fraServices 
         Height          =   8670
         Left            =   -74880
         TabIndex        =   4
         Top             =   360
         Width           =   13710
         Begin VB.ListBox lstPrinters 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   5685
            Left            =   9300
            Sorted          =   -1  'True
            TabIndex        =   17
            Top             =   1635
            Width           =   4215
         End
         Begin VB.FileListBox filDoc 
            Height          =   285
            Left            =   3600
            TabIndex        =   13
            Top             =   8040
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00FF00FF&
            Caption         =   "version A7 - Recette (voir param BIA_SAB.Main_Soc)"
            Height          =   1020
            Left            =   5520
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   120
            Visible         =   0   'False
            Width           =   3735
         End
         Begin VB.ListBox lstElp_Environnement 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   7020
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   9135
         End
         Begin VB.CommandButton cmdContext 
            BackColor       =   &H00C0C0C0&
            Caption         =   "&Recherche"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   7920
            Visible         =   0   'False
            Width           =   1200
         End
         Begin VB.CommandButton cmdMail 
            Height          =   615
            Left            =   1320
            Picture         =   "Bia.frx":11EC
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   8040
            Visible         =   0   'False
            Width           =   1200
         End
         Begin VB.ComboBox cboDataBase 
            Height          =   315
            Left            =   9360
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   135
            Width           =   4215
         End
         Begin VB.ListBox lstAnnuaire 
            BackColor       =   &H00E0E0E0&
            Height          =   255
            Left            =   9240
            Sorted          =   -1  'True
            TabIndex        =   5
            Top             =   8040
            Width           =   4215
         End
         Begin VB.Timer Timer1 
            Enabled         =   0   'False
            Interval        =   1000
            Left            =   2535
            Top             =   8160
         End
         Begin VB.Image Image1 
            Height          =   15
            Left            =   7920
            Top             =   2880
            Width           =   15
         End
      End
   End
   Begin VB.Image imgSocLogo 
      Height          =   435
      Left            =   5520
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1185
   End
   Begin VB.Label lblElpTimer 
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblElpTimer_Next 
      Caption         =   "ElpTimer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   1320
      TabIndex        =   1
      Top             =   -15
      Visible         =   0   'False
      Width           =   3825
   End
   Begin VB.Label lblMain 
      BackStyle       =   0  'Transparent
      Caption         =   "Initialisation  liaison AS400 ........."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   3480
      TabIndex        =   0
      Top             =   1800
      Width           =   1935
   End
End
Attribute VB_Name = "frmElp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim mHeight_0 As Integer, mWidth_0 As Integer, mHeight_2 As Integer, mWidth_2 As Integer, mWindowState As Integer
Dim arrTag() As Boolean, arrTagNb As Integer, lstErrClear As Boolean
Dim Elp_Aut As typeAuthorization
Dim blnAuto As Boolean, blnError As Boolean

Dim MouseMoveActiveControl_Name  As String, MouseMoveActiveControl As typeUsrColor

Dim blnPrinters As Boolean
Dim imgMail_Name As String, blnDatabase_Init As Boolean


Private Sub fgMain_App_X_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim exeName As String
If fgMain_App_X.Row >= 1 And fgMain_App_X.Row < fgMain_App_X.Rows Then
    If X > 1710 And X < 2200 Then
        If Trim(fgMain_App_X.Text) <> "" Then
            Me.MousePointer = 5  '= vbHourglass

            DS_Server_Open
            fgMain_App_X.Col = 1
            Call DS_Document_Load(Trim(fgMain_App_X.Text), paramDocuShare_Collection_SI_Doc)
            Me.MousePointer = 0
        End If
    Else

        fgMain_App_X.Col = 3
        exeName = Mid$(fgMain_App_X.Text, 1, 12)
        
         'If App.PrevInstance Then
         '   Call MsgBox("Il y a déjà une instance active", vbCritical, App_EXEName)
        'Else
            Call Shell_Exe("C:\BIASRV\" & exeName)
        'End If
    End If
fgMain_App_X.Col = 1
fgMain_App_X.CellForeColor = &HD0FFD0

End If

End Sub


Private Sub lstMain_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Static currentlstI As Integer
Dim lstI As Integer, strX As String
''''MouseMoveActiveControl_Set lstMain
lstI = lstMain.TopIndex + Int(y / paramList_Height + 0.5) 'Fix(Y / mLst_Height)
If lstI <> currentlstI Then
    If lstI >= 0 And lstI < lstMain.ListCount Then
    '    lstMain.ListIndex = lstI
    '    strX = lstMain.Text
    '    Mid$(strX, 13, 2) = "  "
        lstMain.ToolTipText = arrBiapgm(lstI)
    '    If y < 195 And lstMain.TopIndex > 0 Then lstMain.TopIndex = lstMain.TopIndex - 1
    '    If y > lstMain.Height - 195 Then lstMain.TopIndex = lstMain.TopIndex + 1
    End If
End If
End Sub


'---------------------------------------------------------
Private Sub lstMain_Click()
'---------------------------------------------------------
Call lstErr_Clear(lstErr, cmdContext, lstMain.Text)

End Sub


Private Sub lstMain_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim strX As String

strX = Space$(100)
strX = Trim(lstMain.Text)
If Mid$(lblMain.Caption, 1, 4) = "Menu" Then
    Msg_Monitor Mid$(strX, 1, 12)
Else
    Elp.usrId = Mid$(strX, 1, 12)
    usrService = Mid$(strX, 11, 3)
    usrGestionnaire = Mid$(strX, 14, 2)
    usrName = Mid$(strX, 17, 34)
    Me.Caption = frmElp_Caption

    Set XForm = Me
    Call MeInit(arrTagNb)
    ReDim arrTag(arrTagNb + 1)

    BIA_VB_APP '$JPL_HAB BiaPgm_Init
End If

End Sub


'
'-------------------------------------------------------
Sub txt_GotFocus(C As Control)
'-------------------------------------------------------
'frm_Control
'C.ForeColor = txtUsr.ForeColor
'C.BackColor = focusUsr.BackColor
'currentActiveControl_Name = C.Name
End Sub


'-------------------------------------------------------
Sub txt_LostFocus(C As Control)
'-------------------------------------------------------
'arrTag(Val(C.Tag)) = True
'C.ForeColor = txtUsr.ForeColor
'C.BackColor = txtUsr.BackColor
End Sub

'---------------------------------------------------------
Private Sub cmdQuit_Click()
'---------------------------------------------------------
If Not Form_QueryUnload_Msgbox Then mainEnd
End Sub


Private Sub cboDataBase_Click()
If cboDataBase.Text <> DataBase_Open Then
    'MDB_Close
    MDB_Open cboDataBase.Text, paramDataBase_Password
End If
End Sub


Private Sub cmdContext_Click()
frmElp.imgSocSignon.Picture = LoadPicture(strSocSignon)
'Elp_ResizeImg frmElp.imgSocSignon
cmdContext.Visible = False
cmdMail.Visible = True
'cmdMail.Visible = False
fgMain_App.Visible = True
fgMain_App_X.Visible = True
End Sub


Private Sub cmdMail_Click()
imgSocSignon.Picture = LoadPicture(imgMail_Name)
'Elp_ResizeImg frmElp.imgSocSignon
cmdContext.Visible = True
cmdMail.Visible = False
fgMain_App.Visible = False
fgMain_App_X.Visible = False
End Sub

Private Sub fgMain_App_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim strX As String

If fgMain_App.Row >= 1 And fgMain_App.Row < fgMain_App.Rows Then
    If X > 1710 And X < 2200 Then
        If Trim(fgMain_App.Text) <> "" Then
            Me.MousePointer = 5  '= vbHourglass

            DS_Server_Open
            fgMain_App.Col = 1
            Call DS_Document_Load(Trim(fgMain_App.Text), paramDocuShare_Collection_SI_Doc)
            
            Me.MousePointer = 0
        End If
    Else
        fgMain_App.Col = 0
        Msg_Monitor Mid$(fgMain_App.Text, 1, 12)
    End If
fgMain_App.Col = 1
fgMain_App.CellForeColor = &HD0FFD0
    
End If

End Sub

Private Sub lstAnnuaire_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Msg_Monitor "Annuaire    " & lstAnnuaire.Text

End Sub


Private Sub lstAnnuaire_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
MouseMoveActiveControl_Set lstAnnuaire

End Sub


Private Sub lstErr_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
MouseMoveActiveControl_Set lstErr

End Sub


Private Sub lstPrinters_Click()
Dim X As String, K As Integer
On Error GoTo Exit_sub

prtCollection_Index = Val(Mid$((lstPrinters.Text), 2, 2)) - 1
If prtCollection_Index >= 0 Then
    Set Printer = Printers(prtCollection_Index)
    frmElpPrt.prtColor_Check
    lstPrinters_Load Printer.Devicename
End If
Exit_sub:
End Sub


'---------------------------------------------------------
Private Sub Form_Activate()
'---------------------------------------------------------
Set XForm = Me
End Sub

Public Sub Msg_Rcv(Msg As String)
'---------------------------------------------------------
mWindowState = Me.WindowState
If Me.WindowState <> frmUsr_Windowstate Then Me.WindowState = frmUsr_Windowstate

Call BiaPgmAut_Init("ELP", Elp_Aut)

Form_Init


End Sub


'---------------------------------------------------------
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

'---------------------------------------------------------
On Error Resume Next
Select Case KeyCode
    Case Is = 27: cmdQuit_Click
'    Case Is = 34: cmdPageNext_Click
'    Case Is = 33: cmdPagePrior_Click
    Case Is = 44: KeyCode = 0: frmElpPrt.prtScreen
    Case Is = 13: SendKeys "{TAB}"
End Select


End Sub

'---------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------
mHeight_0 = Me.Height: mWidth_0 = Me.Width: mHeight_2 = 0: mWidth_2 = 0
SSTab1.Tab = 0
Set XForm = Me
Call MeInit(arrTagNb)
ReDim arrTag(arrTagNb + 1)
cmdContext.Caption = constcmdAbandonner
cmdContext.Visible = False
cboDataBase.Visible = False
lstAnnuaire.Visible = False
End Sub





Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
MouseMoveActiveControl_Reset

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then Cancel = Form_QueryUnload_Msgbox
End Sub

Private Sub Form_Resize()
'If mWindowState <> Me.WindowState Then
'    If Me.WindowState = 0 Or Me.WindowState = 2 Then
'        Elp_Form_Resize Me, mWindowState, mHeight_0, mWidth_0, mHeight_2, mWidth_2
'    End If
'End If
 

End Sub

Private Sub Form_Unload(Cancel As Integer)
If nomDuServeur <> paramServerSplf Then
    If blnTimer_Enabled And paramIMP_PDFCreator_Name = "PDF_BIA_SAB" Then
        Call KillProcess("PDFCreator.exe")
    End If
End If
If Not appExcelPublic Is Nothing Then
    appExcelPublic.Quit
    Set appExcelPublic = Nothing
End If
End
End Sub


Private Sub imgSocSignon_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
MouseMoveActiveControl_Reset
End Sub


Public Sub MouseMoveActiveControl_Set(C As Control)
If MouseMoveActiveControl_Name <> C.Name Then
    MouseMoveActiveControl_Reset
    If Not C.Enabled Then
        MouseMoveActiveControl_Name = ""
    Else
        MouseMoveActiveControl_Name = C.Name
        If TypeOf C Is CommandButton Then
            MouseMoveActiveControl.BackColor = C.BackColor
            C.BackColor = MouseMoveUsr.BackColor
        Else
            If TypeOf C Is ListBox Then
                Elp_ResizeControl C
            Else
                MouseMoveActiveControl.ForeColor = C.ForeColor
                C.ForeColor = MouseMoveUsr.ForeColor
            End If
        End If
    End If
End If

End Sub


Public Sub MouseMoveActiveControl_Reset()
For Each xobj In Me.Controls
    If MouseMoveActiveControl_Name = xobj.Name Then
        MouseMoveActiveControl_Name = ""
        If TypeOf xobj Is CommandButton Then
            xobj.BackColor = MouseMoveActiveControl.BackColor
        Else
            If TypeOf xobj Is ListBox Then
                xobj.Height = 200
            Else
                xobj.ForeColor = MouseMoveActiveControl.ForeColor
            End If
        End If
        Exit For
    End If
Next xobj
End Sub


Public Sub txt_X()
'Call txt_GotFocus(txt)
'KeyAscii = convUCase(KeyAscii)
'Call txt_LostFocus(txt)

'Call txt_GotFocus(txt)
'If XopDevise(2).maxD = 0 Then
'    Call num_KeyAscii(KeyAscii)
'Else
'    Call num_KeyAsciiD(KeyAscii, txt)
'End If
'Call txt_LostFocus(txt)

End Sub

Public Sub lstAnnuaire_Load()
Dim I As Integer
'arrAnnuaire_Load
'lstAnnuaire.AddItem " Annuaire"
'For I = 1 To arrAnnuaireNb
'    lstAnnuaire.AddItem Trim(arrAnnuaire(I).Nom) & ":" & Trim(arrAnnuaire(I).Prénoms) _
'                    & ":" & arrAnnuaire(I).Tél1 & ":" & arrAnnuaire(I).Tél2 & ":" & arrAnnuaire(I).Tél3
'Next I

End Sub

Public Sub lstPrinters_Load(Devicename As String)

On Error Resume Next
Dim mK As Integer, iLen As Integer, K As Integer, X As String
blnPrinters = True
lstPrinters.Clear
iLen = 0

If Trim(Devicename) = "" Then Devicename = XPrt.Devicename
Devicename = UCase$(Trim(Devicename))

If Devicename = paramIMP_PDFCreator_Name Then
    paramIMP_PDF_Path = paramIMP_PDF_Path_VBP
Else
    paramIMP_PDF_Path = paramIMP_PDF_Path_Temp
End If

For Each XPrt In Printers
    K = K + 1
    X = UCase$(Trim(XPrt.Devicename))
    If InStr(Devicename, X) > 0 Then
        If Len(X) > iLen Then iLen = Len(X): mK = K
    End If
    If nomDuServeur <> paramServerSplf Then
        If X = "PDFCREATOR" Then
            Set oPDF = CreateObject("PDFCreator.clsPDFCreator")
            oPDF.cStart ("/NoProcessingAtStartup")
            oPDF.cVisible = False
            oPDF.cPrinterStop = False
            oPDF.cOption("UseAutosave") = 1
            oPDF.cOption("UseAutosaveDirectory") = 1
            oPDF.cOption("AutosaveFormat") = 0
            oPDF.cStart
        End If
    End If
Next

K = 0
For Each XPrt In Printers
    K = K + 1
    If K = mK Then
         X = ">"
         prtCollection_Index = K - 1
          frmElpPrt.prtColor_Check
          lblPrt_DeviceName = Trim(XPrt.Devicename)
          lblPrt_DeviceName.ForeColor = vbMagenta
    Else
        X = "-"
    End If
    lstPrinters.AddItem X & Format$(K, "00 - ") & UCase$(Trim(XPrt.Devicename))
Next

End Sub




Private Sub lstPrinters_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
'MouseMoveActiveControl_Set lstPrinters

End Sub



Public Function Form_QueryUnload_Msgbox() As Boolean
Dim X As String
fgMain_App.Visible = False
fgMain_App_X.Visible = False
lstErr.Visible = False
lstPrinters.Visible = False

If paramEnvironnement = constProduction Then
    'X = MsgBox("Voulez-vous vraiment quitter l'application BIA ?", vbQuestion + vbYesNo, Me.Caption)
    X = vbYes
Else
    X = vbYes
End If
fgMain_App.Visible = True
fgMain_App_X.Visible = True
lstErr.Visible = True
lstPrinters.Visible = True
Form_QueryUnload_Msgbox = IIf(X = vbNo, True, False)

End Function

Private Sub Timer1_Timer()
If blnElpTimer_Auto Then ElpTimer_Monitor "Auto"
End Sub



Public Sub Form_Init()
blnDatabase_Init = True
If lstAnnuaire.ListCount <= 0 Then lstAnnuaire_Load
If Not blnPrinters Then
    Set XPrt = Printer
    lstPrinters_Load ""
    Sleep 100
    If prtCollection_Index >= 0 Then
        Set Printer = Printers(prtCollection_Index)
        frmElpPrt.prtColor_Check
    End If
    'lstPrinters.ListIndex = 0
End If
'''If Not blnDatabase_Init Then cboDataBase_Init

If Elp_Aut.Xspécial Then

    cboDataBase.Visible = True
    cboDataBase.Clear
    cboDataBase.AddItem DataBase_Local
    cboDataBase.AddItem DataBase_Master
    cboDataBase.AddItem "C:\Bia\BiaS820I_XDOC.mdb"   ''paramTemp_Folder & "\DataBase\BIAS820i.mdb"
    cboDataBase.ListIndex = 0
End If

End Sub

