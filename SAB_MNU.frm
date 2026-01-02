VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmSAB_MNU 
   AutoRedraw      =   -1  'True
   Caption         =   "SAB_MNU : gestion des menus"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13875
   Icon            =   "SAB_MNU.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9495
   ScaleWidth      =   13875
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   450
      Left            =   8280
      TabIndex        =   4
      Top             =   45
      Width           =   5055
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8895
      Left            =   0
      TabIndex        =   3
      Top             =   480
      Width           =   13815
      _ExtentX        =   24368
      _ExtentY        =   15690
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Groupes"
      TabPicture(0)   =   "SAB_MNU.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraUsr"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Import "
      TabPicture(1)   =   "SAB_MNU.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraImport"
      Tab(1).ControlCount=   1
      Begin VB.Frame fraImport 
         Caption         =   "Import / Export"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   8460
         Left            =   -74880
         TabIndex        =   7
         Top             =   360
         Width           =   13560
         Begin VB.CommandButton cmdImport_OK 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Importer ZMNUOPT0 => ElpKmInfo : ElpKmIndex"
            Height          =   1215
            Left            =   360
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   600
            Width           =   4275
         End
         Begin VB.ListBox lstW 
            Height          =   4350
            Left            =   5280
            Sorted          =   -1  'True
            TabIndex        =   8
            Top             =   840
            Visible         =   0   'False
            Width           =   2730
         End
      End
      Begin VB.Frame fraUsr 
         Height          =   8445
         Left            =   0
         TabIndex        =   5
         Top             =   360
         Width           =   13560
         Begin VB.Frame fralstUsr 
            Height          =   975
            Left            =   120
            TabIndex        =   27
            Top             =   7320
            Width           =   1935
            Begin VB.OptionButton optUsr_Source 
               Caption         =   "Groupe référence"
               Height          =   255
               Left            =   120
               TabIndex        =   29
               Top             =   600
               Width           =   1695
            End
            Begin VB.OptionButton optUsr_Cible 
               Caption         =   "Groupe à modifier"
               Height          =   195
               Left            =   120
               TabIndex        =   28
               Top             =   240
               Value           =   -1  'True
               Width           =   1575
            End
         End
         Begin VB.Frame fraSelect 
            Caption         =   "Groupe à modifier"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1695
            Left            =   2160
            TabIndex        =   19
            Top             =   6720
            Width           =   5415
            Begin VB.CommandButton cmdSelect_Scan 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Rechercher texte >"
               Height          =   375
               Left            =   240
               Style           =   1  'Graphical
               TabIndex        =   26
               Top             =   240
               Width           =   1695
            End
            Begin VB.TextBox txtSelect_Scan 
               Height          =   315
               Left            =   2040
               TabIndex        =   25
               Top             =   240
               Width           =   2295
            End
            Begin VB.Frame fraSelect_Niveau 
               Height          =   615
               Left            =   240
               TabIndex        =   20
               Top             =   960
               Width           =   3735
               Begin VB.OptionButton optSAB_MNU_L1 
                  Alignment       =   1  'Right Justify
                  Caption         =   "afficher niveau 1"
                  Height          =   240
                  Left            =   120
                  TabIndex        =   24
                  Top             =   240
                  Width           =   1575
               End
               Begin VB.OptionButton optSAB_MNU_L4 
                  Alignment       =   1  'Right Justify
                  Caption         =   "4"
                  Height          =   240
                  Left            =   3120
                  TabIndex        =   23
                  Top             =   240
                  Value           =   -1  'True
                  Width           =   435
               End
               Begin VB.OptionButton optSAB_MNU_L3 
                  Alignment       =   1  'Right Justify
                  Caption         =   " 3"
                  Height          =   240
                  Left            =   2400
                  TabIndex        =   22
                  Top             =   240
                  Width           =   555
               End
               Begin VB.OptionButton optSAB_MNU_L2 
                  Alignment       =   1  'Right Justify
                  Caption         =   " 2"
                  Height          =   240
                  Left            =   1800
                  TabIndex        =   21
                  Top             =   240
                  Width           =   495
               End
            End
         End
         Begin VB.Frame fraContextOptions 
            Caption         =   "Groupe référence"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1815
            Left            =   7800
            TabIndex        =   11
            Top             =   6600
            Width           =   5295
            Begin VB.CommandButton cmdlstSourceScan 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Rechercher texte >"
               Height          =   405
               Left            =   240
               Style           =   1  'Graphical
               TabIndex        =   18
               Top             =   240
               Width           =   1515
            End
            Begin VB.TextBox txtlstSourceScan 
               Height          =   285
               Left            =   2040
               TabIndex        =   17
               Top             =   240
               Width           =   2835
            End
            Begin VB.CheckBox chklstSource_Select_Inf 
               Caption         =   "sélectionner les options dépendantes"
               Height          =   270
               Left            =   2040
               TabIndex        =   15
               Top             =   1320
               Value           =   1  'Checked
               Width           =   2940
            End
            Begin VB.CommandButton cmdlstSource_Select_Clear 
               BackColor       =   &H00C0FFC0&
               Caption         =   "tout désectionner"
               Height          =   420
               Left            =   240
               Style           =   1  'Graphical
               TabIndex        =   14
               Top             =   1320
               Width           =   1635
            End
            Begin VB.CommandButton cmdlstSource_Select_All 
               BackColor       =   &H00C0C0FF&
               Caption         =   "tout sélectionner"
               Height          =   465
               Left            =   240
               Style           =   1  'Graphical
               TabIndex        =   13
               Top             =   720
               Width           =   1620
            End
            Begin VB.CheckBox chklstSource_Select 
               Caption         =   "n'afficher que les options sélectionnées"
               Height          =   270
               Left            =   2040
               TabIndex        =   12
               Top             =   960
               Width           =   3180
            End
         End
         Begin VB.ListBox lstSource 
            Height          =   6360
            ItemData        =   "SAB_MNU.frx":0044
            Left            =   7680
            List            =   "SAB_MNU.frx":004B
            Style           =   1  'Checkbox
            TabIndex        =   10
            Top             =   240
            Width           =   5760
         End
         Begin VB.ListBox lstUsr 
            Height          =   6885
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   6
            Top             =   240
            Width           =   1905
         End
         Begin MSFlexGridLib.MSFlexGrid fgZMNUMEN0 
            Height          =   6465
            Left            =   2160
            TabIndex        =   16
            Top             =   240
            Width           =   5400
            _ExtentX        =   9525
            _ExtentY        =   11404
            _Version        =   393216
            Rows            =   1
            Cols            =   6
            FixedCols       =   0
            RowHeightMin    =   250
            BackColor       =   15007711
            ForeColor       =   12582912
            ForeColorFixed  =   -2147483641
            BackColorSel    =   12648384
            BackColorBkg    =   14737632
            AllowBigSelection=   0   'False
            TextStyleFixed  =   4
            FocusRect       =   2
            HighLight       =   0
            GridLines       =   2
            AllowUserResizing=   3
            FormatString    =   $"SAB_MNU.frx":005A
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Sans Unicode"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   500
      Left            =   13320
      Picture         =   "SAB_MNU.frx":0151
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   500
   End
   Begin VB.CommandButton cmdContext 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Abandonner"
      Height          =   500
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   1200
   End
   Begin VB.Label libRéférenceInterne 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   2
      Top             =   0
      Width           =   5535
      WordWrap        =   -1  'True
   End
   Begin VB.Menu mnuContext 
      Caption         =   "mnuContext"
      Visible         =   0   'False
      Begin VB.Menu mnuContextOptions 
         Caption         =   "Options"
      End
      Begin VB.Menu mnuContext_x1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuContextAbandonner 
         Caption         =   "Abandonner"
      End
      Begin VB.Menu mnuContextQuitter 
         Caption         =   "Quitter"
      End
   End
   Begin VB.Menu mnuselect 
      Caption         =   "mnuSelect"
      Visible         =   0   'False
      Begin VB.Menu mnuSelect_Quit 
         Caption         =   "Abandonner"
      End
      Begin VB.Menu mnuselect1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelect_Delete 
         Caption         =   "Supprimer cette option et les options rattachées"
      End
      Begin VB.Menu mnuselecté 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelect_Insert 
         Caption         =   "Ajouter après cette option (même niveau)"
      End
      Begin VB.Menu mnuSelect_Insert_H 
         Caption         =   "Insérer SOUS cette option (hierarchie)"
      End
   End
End
Attribute VB_Name = "frmSAB_MNU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim mHeight_0 As Integer, mWidth_0 As Integer, mHeight_2 As Integer, mWidth_2 As Integer, mWindowState As Integer
Dim arrTag() As Boolean, arrTagNb As Integer
Dim lastActiveControl_Name  As String, currentActiveControl_Name As String, currentAction As String
Dim MouseMoveActiveControl_Name  As String, MouseMoveActiveControl As typeUsrColor
Dim blnMsgBox_Quit As Boolean, blnAddNew As Boolean, blnGlobalControl As Boolean, blnControl As Boolean
Dim X As String, I As Integer, Msg As String, valX As String, X1 As String, V As Variant, curX As Currency, dblX As Double
Dim intReturn As Integer
Dim SAB_MNUAut As typeAuthorization
Dim blnTransaction As Boolean

Dim fgZMNUMEN0_FormatString As String, fgZMNUMEN0_K As Integer
Dim fgZMNUMEN0_RowDisplay As Integer, fgZMNUMEN0_RowClick As Integer
Dim fgZMNUMEN0_ColorClick As Long, fgZMNUMEN0_ColorDisplay As Long
Dim fgZMNUMEN0_Sort1 As Integer, fgZMNUMEN0_Sort2 As Integer
Dim fgZMNUMEN0_SortAD As Integer, fgZMNUMEN0_Sort1_Old As Integer
Dim fgZMNUMEN0_arrIndex As Integer
Dim blnfgZMNUMEN0_DisplayLine As Boolean


Dim blncmdOk_Visible As Boolean, blnErr As Boolean, blncmdSave_Visible As Boolean
Dim blnSetfocus As Boolean

Dim lstSource_Lib As String, lstSource_CGR As String
Dim lstSource_TopIndex As Long
Dim meElpKMInfo As typeElpKmInfo, xElpKmInfo As typeElpKmInfo

Dim mNiveau_Display As Integer

Dim blnAuto As Boolean, blnAuto_Ok As Boolean

Dim SourceZMNUMEN0() As typeZMNUMEN0, SourceZMNUMEN0_Nb As Long
Dim SourceIndex() As Integer
Dim SourceZMNUOPT0() As typeZMNUOPT0, SourceZMNUOPT0_Nb As Long

Dim arrZMNUMEN0() As typeZMNUMEN0, arrZMNUMEN0_Nb As Long, meZMNUMEN0 As typeZMNUMEN0, xZMNUMEN0 As typeZMNUMEN0
Dim arrZMNUOPT0() As typeZMNUOPT0, arrZMNUOPT0_Nb As Long, meZMNUOPT0 As typeZMNUOPT0, xZMNUOPT0 As typeZMNUOPT0
Dim arrZMNUMEN0_Index As Long

Dim wMNUMENCOD As Long, wHierarchie As String, lenHierarchie As Integer

Dim newZMNUMEN0() As typeZMNUMEN0, newZMNUMEN0_Nb As Long

Dim meZMNUHLB0 As typeZMNUHLB0, xZMNUHLB0 As typeZMNUHLB0

Public Sub lstSource_Select(blnSelect As Boolean)
Dim I As Integer, K As Integer
Dim mIndex As Integer
Dim blnExit As Boolean

blnExit = False
mIndex = lstSource.ListIndex
'---------------------------------------rechercher
K = InStr(1, lstSource, "00")
K = InStr(K, lstSource, " ")
wMNUMENCOD = Val(mId$(lstSource, K - 7, 7))

wHierarchie = "*"
For I = 1 To SourceZMNUMEN0_Nb
    If SourceZMNUMEN0(I).MNUMENCOD = wMNUMENCOD Then
        wHierarchie = SourceZMNUMEN0(I).Hierarchie
        
        Exit For
    End If
Next I

lenHierarchie = Len(wHierarchie)
'--------------------------------------------------------------------------
For I = 1 To SourceZMNUMEN0_Nb
    If mId$(SourceZMNUMEN0(I).Hierarchie, 1, lenHierarchie) = wHierarchie Then
        If blnSelect Then
            If chklstSource_Select_Inf = "1" Then
                SourceZMNUMEN0(I).Method = constAddNew
            Else
                 SourceZMNUMEN0(I).Method = constAddNew: blnExit = True
           End If
            
        Else
            SourceZMNUMEN0(I).Method = ""
        End If
    End If
    If blnExit Then Exit For
Next I
'--------------------------------------------------------------------------
lstSource_Display
lstSource.TopIndex = lstSource_TopIndex
If mIndex < lstSource.ListCount Then
    lstSource.ListIndex = mIndex
Else
    lstSource.ListIndex = lstSource.ListCount - 1
End If

End Sub


'---------------------------------------------------------
Public Sub arrTag_Set(ByVal B As Boolean)
'---------------------------------------------------------
For I = 0 To arrTagNb: arrTag(I) = B: Next I
End Sub


Public Sub cmdContext_Quit()
'''If fraContextOptions.Visible Then fraContextOptions.Visible = False: Exit Sub
blnControl = False
If currentAction <> "" Then
    X = MsgBox("Voulez-vous réellement abandonner la mise à jour?", vbYesNo + vbQuestion + vbDefaultButton2, "Saisie non enregistrée")
    If X = vbYes Then
        currentAction = ""
    Else
        Exit Sub
    End If
End If


If SSTab1.Tab = 2 Then SSTab1.Tab = 1: Exit Sub

'If lstSelectUsr.Visible Then lstSelectUsr.Visible = False: Exit Sub

lstErr.Clear
If SSTab1.Tab > 0 Then
    SSTab1.Tab = SSTab1.Tab - 1
Else
End If
End Sub

Private Sub chklstSource_Select_Click()
fraContextOptions.Visible = True 'False
lstSource_Display

End Sub




Private Sub cmdContext_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
MouseMoveActiveControl_Set cmdContext

End Sub

Private Sub cmdImport_OK_Click()
Dim blnOk As Boolean
blnOk = True
Call lstErr_Clear(lstErr, cmdContext, "cmdImport_OK : Initialisation ")


If blnOk Then
    Me.Enabled = False: Me.MousePointer = vbHourglass

    cmdImport_ZMNUOPT0

    Me.Enabled = True: Me.MousePointer = 0
Call lstErr_AddItem(lstErr, cmdContext, "cmdImport_OK : terminé ")
End If

End Sub

Private Sub cmdPrint_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
MouseMoveActiveControl_Set cmdPrint

End Sub

Private Sub cmdImport_ZMNUOPT0()

Dim blnUpdate As Boolean
Dim kIn As Integer, seq As Long
On Error GoTo Error_Handler
Dim xIn As String, X As String
Dim lMNUOPTCOD As Long, xMNUOPTENS As String, xMNUOPTENT As String, xMNUOPTXXX As String
Dim xIn2 As String
Dim xSql As String
Dim V
Dim recElpKMInfo As typeElpKmInfo, recElpKMIndex As typeElpKmIndex
App_Debug = "cmdImport_ZMNUOPT0_Write : "
On Error GoTo Error_Handler
X = MsgBox("Voulez-vous vraiment mettre à jour [ElpKmInfo :ElpKmIndex] pour le fichier ZMNUOPT0   ?", vbQuestion + vbYesNo, Me.Caption)
If X = vbNo Then Exit Sub

xSql = "delete * from elpkmInfo where ElpKMSrc_Id = 1000"
Call FEU_ROUGE
Set rsMDB = cnMDB.Execute(xSql)

xSql = "delete * from ElpKMIndex where ElpKMSrc_Id = 1000"
Set rsMDB = cnMDB.Execute(xSql)
Call FEU_VERT
Call lstErr_AddItem(lstErr, cmdContext, "cmdImport_ZMNUOPT0_Write: début"): DoEvents


seq = 0

recElpKMInfo.ElpKMSrc_Id = 1000
recElpKMInfo.Pass = 1000
recElpKMInfo.Memo = ""

recElpKMIndex.ElpKMSrc_Id = 1000
recElpKMIndex.Classe = 1000
recElpKMInfo.Memo = ""

xSql = "select * from " & paramIBM_Library_SAB & ".ZMNUOPT0"
Set rsSab = cnsab.Execute(xSql)
    

Do While Not rsSab.EOF
    seq = seq + 1
    If seq Mod 100 = 0 Then Call lstErr_ChangeLastItem(frmSAB_MNU.lstErr, frmSAB_MNU.cmdContext, App_Debug & seq)
    DoEvents
    
    recElpKMInfo.Id = Format$(rsSab("MNUOPTCOD"), "000000000")
    recElpKMInfo.Description = rsSab("MNUOPTLIB")
    V = adoElpKmInfo_AddNew(rsMDB, recElpKMInfo)
    If Not IsNull(V) Then MsgBox V, vbCritical, frmElp_Caption & App_Debug & recElpKMInfo.Id

            
         recElpKMIndex.ElpKMSrc_Id = recElpKMInfo.ElpKMSrc_Id
         xIn2 = Text_LCase(recElpKMInfo.Description)
         blnUpdate = True
         kIn = 0
         Do
             X = Text_KeyWord(xIn2, kIn, False)
        
             If X <> "" Then
                 recElpKMIndex.Id = X
              '  If Trim(recElpKMInfo.ID) = "000012135" Then
              '      Debug.Print recElpKMInfo.ID; X, recElpKMInfo.Description
              '  End If
                xSql = "select * from ElpKMIndex where ID ='" & recElpKMIndex.Id & "' and ElpKMSrc_Id = 1000"
                Set rsMDB = cnMDB.Execute(xSql)
                 If rsMDB.EOF Then
                     recElpKMIndex.Memo = recElpKMInfo.Id
                     V = adoElpKmIndex_AddNew(rsMDB, recElpKMIndex)
                Else
                     recElpKMIndex.Memo = rsMDB("Memo") & recElpKMInfo.Id
                     V = adoElpKmIndex_Update(rsMDB, recElpKMIndex)
                 End If
                 
                If Not IsNull(V) Then MsgBox V, vbCritical, "FIN > " & frmElp_Caption & App_Debug & recElpKMInfo.Id & " " & recElpKMIndex.Id

             Else
                 blnUpdate = False
             End If
             
         Loop While blnUpdate
    
    rsSab.MoveNext
Loop

Call lstErr_ChangeLastItem(frmSAB_MNU.lstErr, frmSAB_MNU.cmdContext, App_Debug & seq)


Exit Sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug


End Sub




Private Sub cmdlstSource_Select_All_Click()
Dim I As Integer
fraContextOptions.Visible = True 'False
For I = 1 To SourceZMNUMEN0_Nb
    SourceZMNUMEN0(I).Method = constAddNew
Next I
lstSource_Display
End Sub

Private Sub cmdlstSource_Select_Clear_Click()
Dim I As Integer
fraContextOptions.Visible = True 'False
For I = 1 To SourceZMNUMEN0_Nb
    SourceZMNUMEN0(I).Method = constDelete
Next I
lstSource_Display

End Sub


Private Sub cmdlstSourceScan_Click()
Dim X As String, I As Integer
On Error Resume Next
lstErr.Clear
X = Text_LCase(txtlstSourceScan)
If X = "" Then
    Call lstErr_AddItem(lstErr, cmdContext, "?  Précisez la recherche ")
    txtlstSourceScan.SetFocus
Else
    lstSource.Visible = False
    Call lst_Scan_Text(X, lstSource)
    lstSource.Visible = True
    If lstSource.ListIndex = -1 Then lstSource.ListIndex = 0: Call lstErr_AddItem(lstErr, cmdContext, "? Option absente de la liste ")
    lstSource.TopIndex = lstSource.ListIndex - 5
End If


End Sub


Private Sub cmdSelect_Scan_Click()
Dim X As String, I As Integer, I0 As Integer, K As Integer
Dim wX As String
Dim xText As String

On Error Resume Next
lstErr.Clear
X = Text_LCase(txtSelect_Scan)
If X = "" Then
    Call lstErr_AddItem(lstErr, cmdContext, "?  Précisez la recherche ")
    txtSelect_Scan.SetFocus
Else
    fgZMNUMEN0.Visible = False
    I0 = fgZMNUMEN0.Row + 1
    For I = I0 To fgZMNUMEN0.Rows - 1
        fgZMNUMEN0.Row = I:
        fgZMNUMEN0.Col = 0: xText = fgZMNUMEN0.Text
        fgZMNUMEN0.Col = 1
        wX = Text_LCase(xText & " " & fgZMNUMEN0.Text)
        K = InStr(wX, X)
        If K > 0 Then
 '           Call fgZMNUMEN0_Color(fgZMNUMEN0_RowClick, MouseMoveUsr.BackColor, fgZMNUMEN0_ColorClick)
            Call fgZMNUMEN0_Color(fgZMNUMEN0_RowClick, vbRed, fgZMNUMEN0_ColorClick)
            fgZMNUMEN0.TopRow = I
            Exit For
        End If
    Next I
    fgZMNUMEN0.Visible = True
End If

End Sub


Private Sub lstUsr_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim xSql As String, K As Integer
K = 0
meZMNUHLB0.MNUHLBREF = Val(Space_Scan(lstUsr.Text, K))
meZMNUHLB0.MNUHLBNOM = Space_Scan(lstUsr.Text, K)
xSql = "select * from " & paramIBM_Library_SAB & ".ZMNUHLB0" _
     & " where MNUHLBNOM = '" & meZMNUHLB0.MNUHLBNOM & "'" _
     & " and   MNUHLBREF =" & meZMNUHLB0.MNUHLBREF _
     & " and   MNUHLBCLA = '2'" _
     & " and MNUHLBETB = " & currentZMNURUT0.MNURUTETB


Set rsSab = cnsab.Execute(xSql)
If Not rsSab.EOF Then
    Call rsZMNUHLB0_GetBuffer(rsSab, meZMNUHLB0)
    If optUsr_Cible Then
        fraSelect.Caption = meZMNUHLB0.MNUHLBCLA & " : " & meZMNUHLB0.MNUHLBNOM
        fgZMNUMEN0_Load
    Else
        fraContextOptions.Caption = meZMNUHLB0.MNUHLBCLA & " : " & meZMNUHLB0.MNUHLBNOM
        lstSource_Load
    End If
    optUsr_Cible = True
End If


End Sub

Private Sub lstSource_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
On Error Resume Next
lstSource_TopIndex = lstSource.TopIndex
lstSource_Select lstSource.Selected(lstSource.ListIndex)
txtlstSourceScan.SetFocus
End Sub



Private Sub mnuSelect_Delete_Click()
Dim Nb As Long
Dim V, xSql As String

On Error GoTo Error_Handler

'-------------------------------------------------------
App_Debug = "mnuSelect_Delete_Click"
'-------------------------------------------------------
Me.Enabled = False
Call lstErr_Clear(lstErr, cmdContext, "Suppression de l'option : " & meZMNUMEN0.MNUMENCOD)
Me.MousePointer = vbHourglass: DoEvents

wHierarchie = "*"
For I = 1 To arrZMNUMEN0_Nb
    If arrZMNUMEN0(I).MNUMENCOD = meZMNUMEN0.MNUMENCOD Then
        wHierarchie = arrZMNUMEN0(I).Hierarchie
        Exit For
    End If
Next I

lenHierarchie = Len(wHierarchie)
'--------------------------------------------------------------------------
Nb = -1
For I = 1 To arrZMNUMEN0_Nb
    If mId$(arrZMNUMEN0(I).Hierarchie, 1, lenHierarchie) = wHierarchie Then Nb = Nb + 1
Next I
'--------------------------------------------------------------------------
If Nb = 0 Then
    X = vbYes
Else
    X = MsgBox("Voulez réellement supprimer cette option et les " & Nb & " options associées ?", _
       vbYesNo + vbQuestion + vbDefaultButton2, _
       "Suppression de l'option : " & meZMNUMEN0.MNUMENCOD & " " & meZMNUOPT0.MNUOPTLIB)
End If
If X = vbYes Then
        cmdSelect_Delete
        fgZMNUMEN0_Load
    Else
        Call lstErr_AddItem(lstErr, cmdContext, "Abandon de la suppression de l'option : " & meZMNUMEN0.MNUMENCOD)

End If

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
Exit_sub:
Me.Enabled = True: Me.MousePointer = 0: DoEvents

End Sub

Private Sub mnuSelect_Insert_Click()
Dim V, xSql As String

On Error GoTo Error_Handler

'-------------------------------------------------------
App_Debug = "mnuSelect_Insert_Click"
'-------------------------------------------------------
Me.Enabled = False
Call lstErr_Clear(lstErr, cmdContext, "Ajout d'options : ")
Me.MousePointer = vbHourglass: DoEvents
cmdSelect_Insert
GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
Exit_sub:
Me.Enabled = True: Me.MousePointer = 0: DoEvents


End Sub


Private Sub mnuSelect_Insert_H_Click()
Dim V, xSql As String

On Error GoTo Error_Handler

'-------------------------------------------------------
App_Debug = "mnuSelect_Insert_H_Click"
'-------------------------------------------------------
Me.Enabled = False
Call lstErr_Clear(lstErr, cmdContext, "Ajout d'options Hiercharchie: ")
Me.MousePointer = vbHourglass: DoEvents
meZMNUMEN0.MNUMENPRE = meZMNUMEN0.MNUMENCOD
meZMNUMEN0.MNUMENORD = 0
cmdSelect_Insert
GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
Exit_sub:
Me.Enabled = True: Me.MousePointer = 0: DoEvents

End Sub


'---------------------------------------------------------
Public Sub cmdReset()
'---------------------------------------------------------
Dim I As Integer

blnControl = False
usrColor_Set
lstSource.BackColor = &HC0C0C0
lstUsr.BackColor = &HE0E0E0

cmdContext.Caption = constcmdRechercher: blnMsgBox_Quit = False
arrTag_Set False
lstErr.Visible = False
blncmdOk_Visible = False: blncmdSave_Visible = False
currentAction = ""

blnAuto = False
blnAuto_Ok = False

libRéférenceInterne = ""


fraImport.Visible = SAB_MNUAut.Xspécial
fraContextOptions.Visible = True 'False


optSAB_MNU_L4_Click

'chargement et affichage du MENU de référence G_ADMIN
'----------------------------------------------------------------------
optUsr_Source = True
'lstSource_Load 7

blnControl = True
End Sub
Public Sub lstSource_Display()
Dim X As String, X7 As String * 7
Dim I As Integer, K As Integer, wIndex As Integer
Dim V, blnOk As Boolean

On Error Resume Next
Me.Enabled = False
Call lstErr_Clear(lstErr, cmdContext, "Affichage menu Source> " & Time)
Me.MousePointer = vbHourglass: DoEvents
lstSource.Visible = False
lstSource.Clear
'----------------------------------------------------------------------

For I = 1 To SourceZMNUMEN0_Nb
     K = SourceIndex(I)
     blnOk = True
    If chklstSource_Select = "1" Then
        If SourceZMNUMEN0(K).Method <> constAddNew Then blnOk = False
    End If
    If blnOk Then
        X7 = Format$(SourceZMNUMEN0(K).MNUMENCOD, "0000000")
        Select Case SourceZMNUMEN0(K).Niveau
            Case "1": X = ""
            Case "2": X = "_" & vbTab
            Case "3": X = "_" & vbTab & "_" & vbTab
            Case Else: X = "_" & vbTab & "_" & vbTab & "_" & vbTab
        End Select
            wIndex = lstSource.ListCount
            lstSource.AddItem X & X7 & "  " & SourceZMNUOPT0(K).MNUOPTLIB
            If SourceZMNUMEN0(K).Method = constAddNew Then lstSource.Selected(wIndex) = True
    End If
Next I

lstSource.Visible = True
Call lstErr_Clear(lstErr, cmdContext, "Affichage menu : " & Time)
lstSource.ListIndex = 0
Me.Enabled = True: Me.MousePointer = 0: DoEvents

End Sub

Private Sub mnuContextOptions_Click()
fraContextOptions.Visible = True
End Sub


Public Sub lstUsr_Load()
Dim xSql As String

On Error Resume Next
lstUsr.Clear
xSql = "select MNUHLBREF, MNUHLBNOM from " & paramIBM_Library_SAB & ".ZMNUHLB0" _
     & " where MNUHLBCLA = '2' and MNUHLBVAL = '1' and MNUHLBFID = 0" _
     & " and MNUHLBETB = " & currentZMNURUT0.MNURUTETB
     
Set rsSab = cnsab.Execute(xSql)
Do While Not rsSab.EOF
    lstUsr.AddItem rsSab("MNUHLBREF") & " " & rsSab("MNUHLBNOM")
    rsSab.MoveNext
Loop

End Sub



Public Sub fgZMNUMEN0_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
mRow = fgZMNUMEN0.Row

If lRow > 0 And lRow < fgZMNUMEN0.Rows Then
    fgZMNUMEN0.Row = lRow
    For I = 0 To fgZMNUMEN0_arrIndex
        fgZMNUMEN0.Col = I: fgZMNUMEN0.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgZMNUMEN0.Row = mRow
    If fgZMNUMEN0.Row > 0 Then
        lRow = fgZMNUMEN0.Row
        lColor_Old = fgZMNUMEN0.CellBackColor
        For I = 0 To fgZMNUMEN0_arrIndex
          fgZMNUMEN0.Col = I: fgZMNUMEN0.CellBackColor = lColor
        Next I
        fgZMNUMEN0.Col = 0
    End If
End If

End Sub

Private Sub fgZMNUMEN0_Display()
Dim I As Long
Dim blnOk As Boolean
Dim kNiveau As Integer
On Error GoTo Error_Handler
SSTab1.Tab = 0
fgZMNUMEN0.Visible = False
fgZMNUMEN0_Reset
Call lstErr_Clear(lstErr, cmdContext, "Options : " & arrZMNUMEN0_Nb): DoEvents
blnOk = True
fgZMNUMEN0.Rows = 1
fgZMNUMEN0.FormatString = fgZMNUMEN0_FormatString
currentAction = "fgZMNUMEN0_Display"
kNiveau = 4
If optSAB_MNU_L3 Then kNiveau = 3
If optSAB_MNU_L2 Then kNiveau = 2
If optSAB_MNU_L1 Then kNiveau = 1
For I = 1 To arrZMNUMEN0_Nb
         
    xZMNUMEN0 = arrZMNUMEN0(I)
    xZMNUOPT0 = arrZMNUOPT0(I)
    If xZMNUMEN0.Niveau <= kNiveau Then
        fgZMNUMEN0.Rows = fgZMNUMEN0.Rows + 1
        fgZMNUMEN0.Row = fgZMNUMEN0.Rows - 1
        fgZMNUMEN0_DisplayLine I
        If xZMNUMEN0.Niveau = -1 Then blnOk = False
    End If
Next I

fgZMNUMEN0.Visible = True
If fgZMNUMEN0.Rows > 1 Then
    fgZMNUMEN0_Sort1 = 2: fgZMNUMEN0_Sort2 = 3: fgZMNUMEN0_Sort
    If Not blnOk And fgZMNUMEN0.Rows > 22 Then fgZMNUMEN0.TopRow = fgZMNUMEN0.Rows - 20
End If

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction



End Sub

Public Sub fgZMNUMEN0_DisplayLine(lIndex As Long)
Dim X As String
On Error Resume Next
fgZMNUMEN0.Col = 0: fgZMNUMEN0.Text = xZMNUMEN0.MNUMENCOD
Select Case xZMNUMEN0.Niveau
    Case "1": X = ""
    Case "2": X = "_  "
    Case "3": X = "_  _  "
    Case Else:  X = "_  _  _  "
End Select

fgZMNUMEN0.Col = 1: fgZMNUMEN0.Text = X & xZMNUOPT0.MNUOPTLIB
Select Case xZMNUMEN0.Niveau
    Case 1: fgZMNUMEN0.CellBackColor = &H80FF80
    Case 2: fgZMNUMEN0.CellBackColor = &HC0FFC0
    Case -1: fgZMNUMEN0.CellBackColor = vbMagenta
End Select
fgZMNUMEN0.Col = 2: fgZMNUMEN0.Text = xZMNUMEN0.Hierarchie
 
fgZMNUMEN0.Col = fgZMNUMEN0_arrIndex: fgZMNUMEN0.Text = lIndex

End Sub

Public Sub fgZMNUMEN0_Sort()
If fgZMNUMEN0.Rows > 1 Then
    fgZMNUMEN0.Row = 1
    fgZMNUMEN0.RowSel = fgZMNUMEN0.Rows - 1
    
    If fgZMNUMEN0_Sort1_Old = fgZMNUMEN0_Sort1 Then
        If fgZMNUMEN0_SortAD = 5 Then
            fgZMNUMEN0_SortAD = 6
        Else
            fgZMNUMEN0_SortAD = 5
        End If
    Else
        fgZMNUMEN0_SortAD = 5
    End If
    fgZMNUMEN0_Sort1_Old = fgZMNUMEN0_Sort1
    
    fgZMNUMEN0.Col = fgZMNUMEN0_Sort1
    fgZMNUMEN0.ColSel = fgZMNUMEN0_Sort2
    fgZMNUMEN0.Sort = fgZMNUMEN0_SortAD
End If

End Sub
Public Sub fgZMNUMEN0_SortX(lK As Integer)
Dim I As Integer, X As String
For I = 1 To fgZMNUMEN0.Rows - 1
    fgZMNUMEN0.Row = I
    fgZMNUMEN0.Col = fgZMNUMEN0_arrIndex
 '   meelpkmInfo_Index = Val(fgZMNUMEN0.Text)
 '   fgZMNUMEN0.Col = fgZMNUMEN0_arrIndex - 1
 '  x = meelpkmInfo(meelpkmInfo_Index).SCCOMPTE & meelpkmInfo(meelpkmInfo_Index).SCDEVISE
 '   Select Case lK
 '       Case 0: fgZMNUMEN0.Text = meelpkmInfo(meelpkmInfo_Index).SCSTATUS & X
 '       Case fgZMNUMEN0_arrIndex: fgZMNUMEN0.Text = Format$(meelpkmInfo_Index, "0000000000")
 '   End Select
Next I


fgZMNUMEN0_Sort1 = fgZMNUMEN0_arrIndex - 1: fgZMNUMEN0_Sort2 = fgZMNUMEN0_arrIndex - 1
fgZMNUMEN0_Sort
End Sub






Public Sub Form_Init()
Call lstErr_Clear(lstErr, cmdContext, "Initialisation ")
DoEvents

SSTab1.Tab = 0

blnControl = False
fgZMNUMEN0_FormatString = fgZMNUMEN0.FormatString

cmdReset


lstUsr_Load

End Sub

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

'-------------------------------------------------------
Sub txt_GotFocus(C As Control)
'-------------------------------------------------------
currentActiveControl_Name = C.Name
C.ForeColor = txtUsr.ForeColor
C.BackColor = focusUsr.BackColor
End Sub


'-------------------------------------------------------
Sub txt_LostFocus(C As Control)
'-------------------------------------------------------
arrTag(Val(C.Tag)) = True
C.ForeColor = txtUsr.ForeColor
C.BackColor = txtUsr.BackColor
End Sub


Private Sub cmdContext_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Select Case cmdContext.Caption
    Case Is = constcmdRechercher: Me.PopupMenu mnuContext, vbPopupMenuLeftButton
    Case Is = constcmdAbandonner: cmdContext_Quit
End Select

End Sub

Private Sub cmdPrint_Click()
Dim Msg As String
Dim I As Integer

Me.Enabled = False: Me.MousePointer = vbHourglass


Me.Enabled = True: Me.MousePointer = 0

End Sub

'---------------------------------------------------------
Private Sub Form_Activate()
'---------------------------------------------------------
Set XForm = Me
End Sub

'---------------------------------------------------------
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'---------------------------------------------------------
Select Case KeyCode
    Case Is = 13: KeyCode = 0: cmdContext_Return
    Case Is = 27: cmdContext_Quit
'   Case Is = 34: cmdPageNext_Click
'   Case Is = 33: cmdPagePrior_Click
    Case Is = 44: KeyCode = 0: frmElpPrt.prtScreen
End Select

End Sub

'---------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------
mHeight_0 = Me.Height: mWidth_0 = Me.Width: mHeight_2 = 0: mWidth_2 = 0: mWindowState = Me.WindowState
Set XForm = Me
Call MeInit(arrTagNb)
ReDim arrTag(arrTagNb + 1)
blnControl = False
fgZMNUMEN0.Clear: fgZMNUMEN0.Row = 0
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
MouseMoveActiveControl_Reset
End Sub


Private Sub fgZMNUMEN0_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If fgZMNUMEN0.Rows > 1 Then
'    Call fgZMNUMEN0_Color(fgZMNUMEN0_RowClick, MouseMoveUsr.BackColor, fgZMNUMEN0_ColorClick)
    Call fgZMNUMEN0_Color(fgZMNUMEN0_RowClick, vbRed, fgZMNUMEN0_ColorClick)
    fgZMNUMEN0.Col = fgZMNUMEN0_arrIndex:  arrZMNUMEN0_Index = CLng(fgZMNUMEN0.Text)
    meZMNUMEN0 = arrZMNUMEN0(arrZMNUMEN0_Index)
    meZMNUOPT0 = arrZMNUOPT0(arrZMNUMEN0_Index)
    fgZMNUMEN0.Col = 0: fgZMNUMEN0.LeftCol = 0
    If Trim(meZMNUOPT0.MNUOPTENT) = "" Then
        mnuSelect_Insert_H = True
    Else
        mnuSelect_Insert_H = False
    End If
    
    Me.PopupMenu mnuselect, vbPopupMenuLeftButton

End If

End Sub
Private Sub Form_Resize()
If mWindowState <> Me.WindowState Then
    If Me.WindowState = 0 Or Me.WindowState = 2 Then
        Elp_Form_Resize Me, mWindowState, mHeight_0, mWidth_0, mHeight_2, mWidth_2
    End If
End If

End Sub

Private Sub mnuContextAbandonner_Click()
cmdContext_Quit
End Sub

Private Sub mnuContextQuitter_Click()
Unload Me
End Sub

Public Sub Msg_Rcv(Msg As String)
'---------------------------------------------------------
mWindowState = Me.WindowState
If Me.WindowState <> frmUsr_Windowstate Then Me.WindowState = frmUsr_Windowstate

Call BiaPgmAut_Init(mId$(Msg, 1, 12), SAB_MNUAut)

blnSetfocus = True
Form_Init


End Sub


Public Sub cmdContext_Return()
If SSTab1.Tab = 0 Then
'    cmdlstSourceScan_Click
Else
    If currentAction = "" Then
        If SSTab1.Tab > 0 Then
            SSTab1.Tab = 0
        Else
           'SendKeys "{TAB}"
           ' cmdSelect_Click
        End If
    End If
End If
End Sub


Public Sub fgZMNUMEN0_Reset()
fgZMNUMEN0.Clear
fgZMNUMEN0.FormatString = fgZMNUMEN0_FormatString
fgZMNUMEN0_Sort1 = 0: fgZMNUMEN0_Sort2 = 0
fgZMNUMEN0_Sort1_Old = -1
fgZMNUMEN0_RowDisplay = 0: fgZMNUMEN0_RowClick = 0
fgZMNUMEN0_arrIndex = fgZMNUMEN0.Cols - 1
blnfgZMNUMEN0_DisplayLine = False
fgZMNUMEN0_SortAD = 6
fgZMNUMEN0.LeftCol = 0

End Sub






Private Sub optSAB_MNU_L1_Click()
fraContextOptions.Visible = True 'False
mNiveau_Display = 1
fgZMNUMEN0_Display
End Sub

Private Sub optSAB_MNU_L2_Click()
fraContextOptions.Visible = True 'False
mNiveau_Display = 2
fgZMNUMEN0_Display
End Sub


Private Sub optSAB_MNU_L3_Click()
fraContextOptions.Visible = True 'False
mNiveau_Display = 3
fgZMNUMEN0_Display
End Sub


Private Sub optSAB_MNU_L4_Click()
fraContextOptions.Visible = True 'False
mNiveau_Display = 4
fgZMNUMEN0_Display
End Sub




Private Sub txtlstSourceScan_GotFocus()
txt_GotFocus txtlstSourceScan

End Sub


Private Sub txtlstSourceScan_LostFocus()
txt_LostFocus txtlstSourceScan

End Sub





Public Sub cmdSelect_Delete()
Dim V, xSql As String
Dim Nb As Long
'-------------------------------------------------------
App_Debug = "cmdSelect_Delete"
'-------------------------------------------------------
V = cmdSelect_Transaction
If Not IsNull(V) Then GoTo Error_MsgBox

Call lstErr_AddItem(lstErr, cmdContext, "Suppression de l'option : " & meZMNUMEN0.MNUMENCOD): DoEvents

For I = 1 To arrZMNUMEN0_Nb
    If mId$(arrZMNUMEN0(I).Hierarchie, 1, lenHierarchie) = wHierarchie Then
        Call lstErr_ChangeLastItem(lstErr, cmdContext, "Suppression de l'option : " & arrZMNUMEN0(I).MNUMENCOD)

        '$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
        xSql = "delete  from " & paramIBM_Library_SAB & ".ZMNUMEN0 " _
            & " where MNUMENGRP = '" & arrZMNUMEN0(I).MNUMENGRP & "'" _
            & " and MNUMENREF =" & arrZMNUMEN0(I).MNUMENREF _
            & " and MNUMENETB =" & arrZMNUMEN0(I).MNUMENETB _
            & " and MNUMENPRE = " & arrZMNUMEN0(I).MNUMENPRE _
            & " and MNUMENCOD = " & arrZMNUMEN0(I).MNUMENCOD _
            & " and MNUMENORD  = " & arrZMNUMEN0(I).MNUMENORD
         Call FEU_ROUGE
        Set rsSab_Update = cnSab_Update.Execute(xSql, Nb)
        Call FEU_VERT

        If Nb <> 1 Then
            V = "Erreur suppression de l'option : " & arrZMNUMEN0(I).MNUMENCOD
            GoTo Error_MsgBox
        End If
    End If

Next I

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & App_Debug
Exit_sub:
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
    Else
        V = cnSAB_Transaction("Commit")
    End If
    
    '$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
End Sub


Public Sub fgZMNUMEN0_Load()
arrZMNUMEN0_Load meZMNUHLB0, arrZMNUMEN0, arrZMNUOPT0
arrZMNUMEN0_Nb = UBound(arrZMNUMEN0) - 1
fgZMNUMEN0_Display

End Sub

Public Sub lstSource_Load()
arrZMNUMEN0_Load meZMNUHLB0, arrZMNUMEN0, arrZMNUOPT0
arrZMNUMEN0_Nb = UBound(arrZMNUMEN0) - 1
fgZMNUMEN0_Display

'----------------------------------------------------------------------
SourceZMNUMEN0_Nb = arrZMNUMEN0_Nb
ReDim SourceZMNUMEN0(SourceZMNUMEN0_Nb), SourceZMNUOPT0(SourceZMNUMEN0_Nb)
ReDim SourceIndex(SourceZMNUMEN0_Nb)
For I = 1 To SourceZMNUMEN0_Nb
    fgZMNUMEN0.Row = I: fgZMNUMEN0.Col = fgZMNUMEN0_arrIndex
    SourceIndex(I) = Val(fgZMNUMEN0.Text)
    SourceZMNUMEN0(I) = arrZMNUMEN0(I)
    SourceZMNUOPT0(I) = arrZMNUOPT0(I)
Next I
lstSource_Display
'----------------------------------------------------------------------
fgZMNUMEN0_Reset

End Sub

Public Sub cmdSelect_Insert()
Dim I As Integer, K As Integer, X As String
Dim wMNUMENORD As Long, wHierarchie As String, lenHierarchie As Integer
Dim V, xSql As String
Dim blnTransaction As Boolean
On Error GoTo Error_Handler

blnTransaction = False
Set rsSab = Nothing
'-----------------------------------------------------------------------------------------------
App_Debug = "1-cmdSelect_Insert : chargement des options CIBLE de même niveau dans newZMNUNEN0"
'-----------------------------------------------------------------------------------------------
ReDim newZMNUMEN0(100)
newZMNUMEN0_Nb = 0
xSql = "select * from " & paramIBM_Library_SAB & ".ZMNUMEN0 " _
     & " where MNUMENGRP = '" & meZMNUMEN0.MNUMENGRP & "'" _
     & " and   MNUMENREF =" & meZMNUMEN0.MNUMENREF _
     & " and   MNUMENETB =" & meZMNUMEN0.MNUMENETB _
      & " and MNUMENPRE = " & meZMNUMEN0.MNUMENPRE _
     & " order by  MNUMENORD"
     
Set rsSab = cnsab.Execute(xSql)
Do While Not rsSab.EOF
    If newZMNUMEN0_Nb = UBound(newZMNUMEN0) Then ReDim Preserve newZMNUMEN0(newZMNUMEN0_Nb + 100)
    newZMNUMEN0_Nb = newZMNUMEN0_Nb + 1
    Call rsZMNUMEN0_GetBuffer(rsSab, newZMNUMEN0(newZMNUMEN0_Nb))
    rsSab.MoveNext
Loop
'-----------------------------------------------------------------------------------------------
App_Debug = "2-cmdSelect_Insert : ajout des options SOURCE dans newZMNUNEN0 "
'-----------------------------------------------------------------------------------------------
xZMNUMEN0.MNUMENCOD = -1
wHierarchie = "X": lenHierarchie = 1
For I = 1 To SourceZMNUMEN0_Nb
    K = SourceIndex(I)
    If SourceZMNUMEN0(K).Method = constAddNew Then
        If newZMNUMEN0_Nb = UBound(newZMNUMEN0) Then ReDim Preserve newZMNUMEN0(newZMNUMEN0_Nb + 100)
        newZMNUMEN0_Nb = newZMNUMEN0_Nb + 1
        newZMNUMEN0(newZMNUMEN0_Nb) = SourceZMNUMEN0(K)
        newZMNUMEN0(newZMNUMEN0_Nb).MNUMENETB = meZMNUMEN0.MNUMENETB
        newZMNUMEN0(newZMNUMEN0_Nb).MNUMENREF = meZMNUMEN0.MNUMENREF
        newZMNUMEN0(newZMNUMEN0_Nb).MNUMENGRP = meZMNUMEN0.MNUMENGRP
        
        If mId$(SourceZMNUMEN0(K).Hierarchie, 1, lenHierarchie) <> wHierarchie Then
            ' ce n'est pas un sous niveau : insertion après MNUMENORD
            newZMNUMEN0(newZMNUMEN0_Nb).MNUMENPRE = meZMNUMEN0.MNUMENPRE
            newZMNUMEN0(newZMNUMEN0_Nb).MNUMENORD = meZMNUMEN0.MNUMENORD
            xZMNUMEN0 = SourceZMNUMEN0(K)
            wHierarchie = SourceZMNUMEN0(K).Hierarchie: lenHierarchie = Len(wHierarchie)
        Else
            ' vérifier si ce sous-niveau existe déjà
            '========================================
            xSql = "select * from " & paramIBM_Library_SAB & ".ZMNUMEN0 " _
                & " where MNUMENGRP = '" & meZMNUMEN0.MNUMENGRP & "'" _
                & " and   MNUMENREF =" & meZMNUMEN0.MNUMENREF _
                & " and   MNUMENETB =" & meZMNUMEN0.MNUMENETB _
                & " and MNUMENPRE = " & SourceZMNUMEN0(K).MNUMENPRE _
                & " and MNUMENCOD = " & SourceZMNUMEN0(K).MNUMENCOD
                
            Set rsSab = cnsab.Execute(xSql)
            If Not rsSab.EOF Then
                V = "l'option " & SourceZMNUMEN0(K).MNUMENCOD & " existe déjà dans le menu " & SourceZMNUMEN0(K).MNUMENPRE
                GoTo Error_MsgBox
            End If
        End If
    End If
Next I
'-----------------------------------------------------------------------------------------------
App_Debug = "3-cmdSelect_Insert : gestion MNUMENORD : tri dans lstW, puis incrément ORD + 1000 "
'-----------------------------------------------------------------------------------------------
lstW.Clear
For I = 1 To newZMNUMEN0_Nb
    If newZMNUMEN0(I).MNUMENPRE = meZMNUMEN0.MNUMENPRE Then
        lstW.AddItem Format$(newZMNUMEN0(I).MNUMENORD, "0000000") & Format$(I, "0000000")
    End If
Next I

xZMNUMEN0.MNUMENORD = 0

For I = 0 To lstW.ListCount - 1
    lstW.ListIndex = I
    wMNUMENORD = Val(mId$(lstW.Text, 1, 7))
    If wMNUMENORD <= xZMNUMEN0.MNUMENORD Then
        K = Val(mId$(lstW.Text, 8, 7))
        newZMNUMEN0(K).Niveau = newZMNUMEN0(K).MNUMENORD  'ATTENTION : mémoriser MNUMENORD pour update SQL
        wMNUMENORD = xZMNUMEN0.MNUMENORD + 1000
        newZMNUMEN0(K).MNUMENORD = wMNUMENORD
        If newZMNUMEN0(K).Method <> constAddNew Then
            newZMNUMEN0(K).Method = constUpdate
        End If
    End If
    xZMNUMEN0.MNUMENORD = wMNUMENORD
Next I
'-----------------------------------------------------------------------------------------------
App_Debug = "4-cmdSelect_Insert : Mise à jour  "
'-----------------------------------------------------------------------------------------------
'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
V = cmdSelect_Transaction
If Not IsNull(V) Then GoTo Error_MsgBox
blnTransaction = True

Call lstErr_AddItem(lstErr, cmdContext, "Mise à jour des options : "): DoEvents

For I = newZMNUMEN0_Nb To 1 Step -1
   If newZMNUMEN0(I).Method = constUpdate Then
        Call lstErr_ChangeLastItem(lstErr, cmdContext, newZMNUMEN0(I).Method & " option : " & newZMNUMEN0(I).MNUMENCOD)
        xZMNUMEN0 = newZMNUMEN0(I)
        xZMNUMEN0.MNUMENORD = newZMNUMEN0(I).Niveau
        V = sqlZMNUMEN0_Update(newZMNUMEN0(I), xZMNUMEN0)
        If Not IsNull(V) Then GoTo Error_MsgBox
    End If

Next I
For I = 1 To newZMNUMEN0_Nb
   If newZMNUMEN0(I).Method = constAddNew Then
        Call lstErr_ChangeLastItem(lstErr, cmdContext, newZMNUMEN0(I).Method & " option : " & newZMNUMEN0(I).MNUMENCOD)
        V = sqlZMNUMEN0_Insert(newZMNUMEN0(I))
        If Not IsNull(V) Then GoTo Error_MsgBox
    End If

Next I

GoTo Exit_sub

'-----------------------------------------------------------------------------------------------
App_Debug = "9-cmdSelect_Insert : Exit"
'-----------------------------------------------------------------------------------------------

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & " : " & App_Debug
Exit_sub:
    If blnTransaction Then
    '$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        If Not IsNull(V) Then
            V = cnSAB_Transaction("Rollback")
        Else
            V = cnSAB_Transaction("Commit")
            cmdlstSource_Select_Clear_Click
            fgZMNUMEN0_Load
        End If
    End If

End Sub

Public Function cmdSelect_Transaction()
Dim V
V = cnSAB_Transaction("BeginTrans")
If IsNull(V) Then
    xZMNUHLB0 = meZMNUHLB0
    xZMNUHLB0.MNUHLBSUS = currentZMNURUT0.MNURUTCUT
    xZMNUHLB0.MNUHLBSDT = valDSys - 19000000
    xZMNUHLB0.MNUHLBSHE = Val(time_Hms)
    
    V = sqlZMNUHLB0_Update(xZMNUHLB0, meZMNUHLB0)
End If
cmdSelect_Transaction = V
End Function
