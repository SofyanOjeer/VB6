VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmSAB_Echelles 
   AutoRedraw      =   -1  'True
   Caption         =   "SAB_Echelles"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13875
   Icon            =   "SAB_Echelles.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9495
   ScaleWidth      =   13875
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   8280
      TabIndex        =   4
      Top             =   45
      Width           =   5055
   End
   Begin TabDlg.SSTab ssTab1 
      Height          =   8895
      Left            =   0
      TabIndex        =   3
      Top             =   500
      Width           =   13815
      _ExtentX        =   24368
      _ExtentY        =   15690
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Sélection"
      TabPicture(0)   =   "SAB_Echelles.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraSelect"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Statistiques"
      TabPicture(1)   =   "SAB_Echelles.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "filW2"
      Tab(1).Control(1)=   "filW"
      Tab(1).Control(2)=   "dirW"
      Tab(1).ControlCount=   3
      Begin VB.FileListBox filW2 
         Height          =   2625
         Left            =   -65640
         TabIndex        =   22
         Top             =   5400
         Visible         =   0   'False
         Width           =   2745
      End
      Begin VB.FileListBox filW 
         Height          =   2625
         Left            =   -65760
         TabIndex        =   10
         Top             =   2040
         Visible         =   0   'False
         Width           =   2745
      End
      Begin VB.DirListBox dirW 
         Height          =   1368
         Left            =   -68400
         TabIndex        =   11
         Top             =   1920
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.Frame fraSelect 
         Height          =   8445
         Left            =   0
         TabIndex        =   5
         Top             =   360
         Width           =   13560
         Begin VB.Frame fraSelect_Doc 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   7695
            Left            =   120
            TabIndex        =   7
            Top             =   720
            Width           =   13335
            Begin VB.Frame fraUpdate 
               Caption         =   "Sélection des relevés d'échelles à imprimer"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   3735
               Left            =   4320
               TabIndex        =   12
               Top             =   3600
               Width           =   8535
               Begin VB.CommandButton btnControle 
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "Contrôler"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   345
                  Left            =   6720
                  Style           =   1  'Graphical
                  TabIndex        =   23
                  Top             =   720
                  Width           =   1575
               End
               Begin VB.CheckBox chkUpdate_Nostro 
                  Caption         =   "Exclure les Nostros"
                  Height          =   375
                  Left            =   360
                  TabIndex        =   21
                  Top             =   1200
                  Value           =   1  'Checked
                  Width           =   4215
               End
               Begin VB.TextBox txtUpdate_Compte 
                  Height          =   285
                  Left            =   3480
                  TabIndex        =   19
                  Top             =   2880
                  Width           =   2775
               End
               Begin VB.CheckBox chkUpdate_Avis 
                  Caption         =   "Uniquement les comptes ayant un ticket d'agios"
                  Height          =   255
                  Left            =   360
                  TabIndex        =   18
                  Top             =   2160
                  Value           =   1  'Checked
                  Width           =   3975
               End
               Begin VB.CheckBox chkUpdate_SoldeZ 
                  Caption         =   "Exclure les comptes sans mouvements et solde nul"
                  Height          =   255
                  Left            =   360
                  TabIndex        =   17
                  Top             =   1680
                  Value           =   1  'Checked
                  Width           =   4335
               End
               Begin VB.CheckBox chkUpdate_Import 
                  Caption         =   "Importer AVI02P1 => YECHIMP0"
                  Height          =   255
                  Left            =   360
                  TabIndex        =   16
                  Top             =   360
                  Value           =   1  'Checked
                  Width           =   7695
               End
               Begin VB.CheckBox chkUpdate_Archivage 
                  Caption         =   "Archiver ECHEDI01P2 & AVI02P1"
                  Height          =   375
                  Left            =   360
                  TabIndex        =   15
                  Top             =   720
                  Value           =   1  'Checked
                  Width           =   4455
               End
               Begin VB.CommandButton cmdUpdate_Quit 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "Abandonner"
                  Height          =   525
                  Left            =   6720
                  Style           =   1  'Graphical
                  TabIndex        =   14
                  Top             =   2820
                  Width           =   1575
               End
               Begin VB.CommandButton cmdUpdate_Ok 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "Lancer l'impression"
                  Height          =   1005
                  Left            =   6720
                  Style           =   1  'Graphical
                  TabIndex        =   13
                  Top             =   1320
                  Width           =   1575
               End
               Begin VB.Label lblUpdate_Compte 
                  Caption         =   "uniquement les comptes commençant par les caractères"
                  Height          =   495
                  Left            =   720
                  TabIndex        =   20
                  Top             =   2760
                  Width           =   2295
               End
            End
            Begin VB.ListBox lstSelect 
               Height          =   7080
               Left            =   120
               TabIndex        =   9
               Top             =   240
               Width           =   3735
            End
            Begin MSFlexGridLib.MSFlexGrid fgSelect 
               Height          =   2505
               Left            =   4200
               TabIndex        =   8
               Top             =   240
               Width           =   8640
               _ExtentX        =   15240
               _ExtentY        =   4419
               _Version        =   393216
               Rows            =   1
               Cols            =   8
               FixedCols       =   0
               RowHeightMin    =   300
               BackColor       =   16777210
               ForeColor       =   8388608
               BackColorFixed  =   16776921
               ForeColorFixed  =   -2147483641
               BackColorSel    =   12648384
               BackColorBkg    =   16777210
               WordWrap        =   -1  'True
               AllowBigSelection=   0   'False
               TextStyleFixed  =   4
               FocusRect       =   2
               HighLight       =   0
               GridLines       =   3
               GridLinesFixed  =   1
               AllowUserResizing=   3
               FormatString    =   "<Utilisateur  |<Date       |<Heure       |>Etat          |>Job            |>Séq       ||"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
         End
         Begin VB.ComboBox cboSelect_SQL 
            Height          =   315
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   6
            Text            =   "cboSelect_SQL"
            Top             =   240
            Width           =   3615
         End
      End
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   500
      Left            =   13320
      Picture         =   "SAB_Echelles.frx":0044
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
   End
   Begin VB.Menu mnuPrint0 
      Caption         =   "mnuPrint0"
      Visible         =   0   'False
      Begin VB.Menu mnuPrint0_All 
         Caption         =   "Imprimer TOUS les courriers"
      End
   End
End
Attribute VB_Name = "frmSAB_Echelles"
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
Dim BIA_Echelles_Aut As typeAuthorization
Dim blnTransaction As Boolean
Dim blnAuto As Boolean, blnAuto_Ok As Boolean
Dim wAMJMin As String, WAMJMax As String, wHmsMin As Long, wHmsMax As Long
Dim wAmjMin7 As Long, wAmjMax7 As Long


Dim blncmdOk_Visible As Boolean, blnErr As Boolean, blncmdSave_Visible As Boolean
Dim blnSetfocus As Boolean

Dim fgSelect_FormatString As String, fgSelect_K As Integer
Dim fgSelect_RowDisplay As Integer, fgSelect_RowClick As Integer, fgSelect_ColClick As Integer
Dim fgSelect_ColorClick As Long, fgSelect_ColorDisplay As Long
Dim fgSelect_Sort1 As Integer, fgSelect_Sort2 As Integer
Dim fgSelect_SortAD As Integer, fgSelect_Sort1_Old As Integer
Dim fgSelect_arrIndex As Integer
Dim blnfgSelect_DisplayLine As Boolean

Dim cmdSelect_SQL_K As String
Dim cmdSelect_Ok_Caption As String
'______________________________________________________________________

Dim fgList_FormatString As String, fgList_K As Integer
Dim fgList_RowDisplay As Integer, fgList_RowClick As Integer, fgList_ColClick As Integer
Dim fgList_ColorClick As Long, fgList_ColorDisplay As Long
Dim fgList_Sort1 As Integer, fgList_Sort2 As Integer
Dim fgList_SortAD As Integer, fgList_Sort1_Old As Integer
Dim fgList_arrIndex As Integer
Dim blnfgList_DisplayLine As Boolean
'______________________________________________________________________

Dim xFileName_ECHEDI01P2 As String, xFileName_ECHAVI02P1 As String

'                                                                           '
Dim ctl_arrYECHREL0(365) As typeYECHREL0, ctl_arrYECHREL0_Nb As Integer
Dim ctl_mYECHIMP0 As typeYECHIMP0
Dim ctl_blnNewPage As Boolean, ctl_blnAvis As Boolean

Private Sub ctl_prtSAB_Echelles_ECHEDI01P2(lFile As String, lECHIMPJOBS As Integer, blnNostro As Boolean, blnSoldeZ As Boolean, blnCompteAvis As Boolean, selCompte As String, fic As Long)
Dim V, X As String, I As Integer
Dim xIn As String, xIn_Report As String
Dim idFile As Integer
Dim zYECHIMP0 As typeYECHIMP0, xYECHIMP0 As typeYECHIMP0
Dim blnHeader As Boolean, blnEnd As Boolean
Dim K As Integer, X8 As String, wAAAA As Long
Dim wMontant As Currency, wSens As String, wValeur As Long
Dim iAdresse As Integer
Dim blnOk As Boolean
Dim newName As String

On Error GoTo Error_Handle

rsYECHIMP0_Init zYECHIMP0
blnHeader = False: blnEnd = False
ctl_arrYECHREL0_Nb = 0

idFile = FreeFile
Open lFile For Input As #idFile
'____________________________________________________________________________________________
Line Input #idFile, xIn
If Mid$(xIn, 24, 10) <> "ECHEDI01P2" Then V = "Ce n'est pas un état ECHEDI01P2": GoTo Error_MsgBox

zYECHIMP0.ECHIMPDTRT = Val(Mid$(xIn, 5, 8))
zYECHIMP0.ECHIMPJOB = Val(Mid$(xIn, 13, 6))
zYECHIMP0.ECHIMPJOBS = lECHIMPJOBS '''Val(Mid$(xIn, 19, 5)) + 2  ''!!! à vérifier
'____________________________________________________________________________________________
Do Until EOF(idFile)
    Line Input #idFile, xIn
    If Mid$(xIn, 1, 3) = "064" Or Mid$(xIn, 13, 10) = "Nbre Débit" Then
                 ctl_arrYECHREL0_Nb = ctl_arrYECHREL0_Nb + 1
                 ctl_arrYECHREL0(ctl_arrYECHREL0_Nb).ECHRELSD = 0
                 ctl_arrYECHREL0(ctl_arrYECHREL0_Nb).ECHRELSDS = " "
                 ctl_arrYECHREL0(ctl_arrYECHREL0_Nb).ECHRELDVAL = 0
                 
                 X = Trim(Mid$(xIn, 26, 14))
                 If X = "" Then
                    ctl_arrYECHREL0(ctl_arrYECHREL0_Nb).ECHRELMDB = 0
                 Else
                    ctl_arrYECHREL0(ctl_arrYECHREL0_Nb).ECHRELMDB = CCur(X)     '!!!!!! nbr DB
                 End If
                 X = Trim(Mid$(xIn, 63, 14))
                 If X = "" Then
                    ctl_arrYECHREL0(ctl_arrYECHREL0_Nb).ECHRELMCR = 0
                 Else
                    ctl_arrYECHREL0(ctl_arrYECHREL0_Nb).ECHRELMCR = CCur(X)         '!!!!!! nbr CR
                 End If
                 
                Line Input #idFile, xIn
                Line Input #idFile, xIn
                ctl_arrYECHREL0(ctl_arrYECHREL0_Nb).ECHRELDVAL = dateX8_N8(Mid$(xIn, 46, 8))
                X = Trim(Mid$(xIn, 57, 14))
                If X = "" Then
                   ctl_arrYECHREL0(ctl_arrYECHREL0_Nb).ECHRELSD = 0
                Else
                   ctl_arrYECHREL0(ctl_arrYECHREL0_Nb).ECHRELSD = CCur(X)     '!!!!!! solde final
                End If
                ctl_arrYECHREL0(ctl_arrYECHREL0_Nb).ECHRELSDS = Mid$(xIn, 71, 1)

                If blnHeader Then
'______________________________________________________________________________________________
                    blnOk = True
                    If selCompte <> Mid$(xYECHIMP0.ECHIMPCPT, 1, Len(selCompte)) Then blnOk = False
                    If Mid$(xYECHIMP0.ECHIMPCPT, 1, 1) = "N" And Not blnNostro Then blnOk = False
                    If Not blnSoldeZ And ctl_arrYECHREL0_Nb = 2 And ctl_arrYECHREL0(0).ECHRELSD = 0 And ctl_arrYECHREL0(ctl_arrYECHREL0_Nb).ECHRELSD = 0 Then blnOk = False
                    If blnOk Then
                        Call ctl_prtSAB_Echelles_ECHEDI01P2_Relevé(xYECHIMP0, blnCompteAvis, fic)
                    End If
 '______________________________________________________________________________________________
                   blnHeader = False
                End If
                               
                Line Input #idFile, xIn
                If Mid$(xIn, 1, 3) = "$$ " Then Exit Do
                K = InStr(xIn, "ECHELLES")
                If K <= 0 Then Line Input #idFile, xIn
                If Mid$(xIn, 1, 3) = "$$ " Then Exit Do
    Else
        Select Case Mid$(xIn, 1, 3)
           Case "013"
                     If Not blnHeader Then
                       xYECHIMP0 = zYECHIMP0
                       xYECHIMP0.ECHIMPAD1 = Mid$(xIn, 47, 32)
                       iAdresse = 2
                       Do
                            Line Input #idFile, xIn
                            If Mid$(xIn, 1, 3) <> "022" Then
                                Select Case iAdresse
                                    Case 2: xYECHIMP0.ECHIMPAD2 = Mid$(xIn, 47, 32)
                                    Case 3: xYECHIMP0.ECHIMPAD3 = Mid$(xIn, 47, 32)
                                    Case 4: xYECHIMP0.ECHIMPAD4 = Mid$(xIn, 47, 32)
                                    Case 5: xYECHIMP0.ECHIMPAD5 = Mid$(xIn, 47, 32)
                                    Case 6: xYECHIMP0.ECHIMPAD6 = Mid$(xIn, 47, 32)
                                    Case 7: xYECHIMP0.ECHIMPAD7 = Mid$(xIn, 47, 32)
                                End Select
                                iAdresse = iAdresse + 1
                            Else
                            
           'Case "021"
                                Line Input #idFile, xIn
                                X = Trim(Mid$(xIn, 19, 20))
                                
                                 If blnHeader Then
                                    If X <> Trim(xYECHIMP0.ECHIMPCPT) Then V = "Erreur rupture page compte " & X: GoTo Error_MsgBox
                                 Else
                                   xYECHIMP0.ECHIMPCPT = X
                                   xYECHIMP0.ECHIMPDEV = Mid$(xIn, 13, 3)
                                End If
                                Exit Do
                            End If
                        Loop
                    End If
           Case "028"
                    Line Input #idFile, xIn
                    If Mid$(xIn, 13, 13) <> "Report valeur" Then
                        V = "Erreur Report valeur, compte " & xYECHIMP0.ECHIMPCPT
                        GoTo Error_MsgBox
                    End If
                    Line Input #idFile, xIn
                    If Not blnHeader Then
                        blnHeader = True
                        xYECHIMP0.ECHIMPDDEB = dateX6_N8(Mid$(xIn, 34, 6))
                        ctl_arrYECHREL0_Nb = 0
                        ctl_arrYECHREL0(0).ECHRELDVAL = xYECHIMP0.ECHIMPDDEB
                        ctl_arrYECHREL0(0).ECHRELSD = CCur(Trim(Mid$(xIn, 45, 14)))
                        ctl_arrYECHREL0(0).ECHRELSDS = Mid$(xIn, 59, 1)
                    End If
                    Line Input #idFile, xIn
           Case "062"
                    Line Input #idFile, xIn_Report
                    Do
                        Line Input #idFile, xIn
                    Loop Until Mid$(xIn, 1, 3) = "028"
                    Line Input #idFile, xIn
                    Line Input #idFile, xIn
                    Line Input #idFile, xIn
           Case Else
                    
                    If blnHeader Then
                        If Trim(Mid$(xIn, 5, 29)) <> "" Then        ' rupture mois dans certains cas
                             ctl_arrYECHREL0_Nb = ctl_arrYECHREL0_Nb + 1
                             X = Trim(Mid$(xIn, 13, 15))
                             If X = "" Then
                                ctl_arrYECHREL0(ctl_arrYECHREL0_Nb).ECHRELMDB = 0
                             Else
                                ctl_arrYECHREL0(ctl_arrYECHREL0_Nb).ECHRELMDB = CCur(X)
                             End If
                             
                            
                             Line Input #idFile, xIn
                             
                             X = Trim(Mid$(xIn, 17, 16))
                             If X = "" Then
                                ctl_arrYECHREL0(ctl_arrYECHREL0_Nb).ECHRELMCR = 0
                              Else
                                ctl_arrYECHREL0(ctl_arrYECHREL0_Nb).ECHRELMCR = CCur(X)
                             End If
                             X = Trim(Mid$(xIn, 45, 14))
                             If X = "" Then
                               ctl_arrYECHREL0(ctl_arrYECHREL0_Nb).ECHRELSD = 0
                              Else
                               ctl_arrYECHREL0(ctl_arrYECHREL0_Nb).ECHRELSD = CCur(X)
                             End If
                               
                             ctl_arrYECHREL0(ctl_arrYECHREL0_Nb).ECHRELSDS = Mid$(xIn, 59, 1)
                             ctl_arrYECHREL0(ctl_arrYECHREL0_Nb).ECHRELDVAL = dateX6_N8(Mid$(xIn, 34, 6))
                             
                             X = Trim(Mid$(xIn, 41, 3))
                             If X = "" Then
                                ctl_arrYECHREL0(ctl_arrYECHREL0_Nb).ECHRELNBJ = 0
                             Else
                                ctl_arrYECHREL0(ctl_arrYECHREL0_Nb).ECHRELNBJ = CInt(X)
                             End If
                             X = Trim(Mid$(xIn, 61, 14))
                             If X = "" Then
                                 ctl_arrYECHREL0(ctl_arrYECHREL0_Nb).ECHRELNBR = 0
                             Else
                                  ctl_arrYECHREL0(ctl_arrYECHREL0_Nb).ECHRELNBR = CCur(X)
                            End If
                             X = Trim(Mid$(xIn, 75, 10))
                            If X = "" Then
                                 ctl_arrYECHREL0(ctl_arrYECHREL0_Nb).ECHRELTAUX = 0
                             Else
                                If InStr(X, "%") > 0 Then
                                    X = Replace(Trim(Mid$(xIn, 75, 10)), "%", ",")
                                    ctl_arrYECHREL0(ctl_arrYECHREL0_Nb).ECHRELTAUX = CDbl(X)
                                Else
                                     ctl_arrYECHREL0(ctl_arrYECHREL0_Nb).ECHRELTAUX = CDbl(X) / 100000
                               End If
                            End If
                        Else
                             X = Trim(Mid$(xIn, 75, 10))
                             If ctl_arrYECHREL0(ctl_arrYECHREL0_Nb).ECHRELTAUX = 0 And X <> "" Then
                                If InStr(X, "%") > 0 Then
                                    X = Replace(Trim(Mid$(xIn, 75, 10)), "%", ",")
                                    ctl_arrYECHREL0(ctl_arrYECHREL0_Nb).ECHRELTAUX = CDbl(X)
                                Else
                                     ctl_arrYECHREL0(ctl_arrYECHREL0_Nb).ECHRELTAUX = CDbl(X) / 100000
                               End If
                            End If
                       
                        End If
                    End If
        End Select
    End If
Loop

Close idFile

MsgBox "Fin du traitement contrôle Echelles..."
Exit Sub

Error_Handle:
V = Error
Error_MsgBox:

MsgBox "prtSAB_Echelles_ECHAVI02P1" & Error, vbCritical, V
End Sub
Private Function ctl_prtSAB_Echelles_ECHEDI01P2_Relevé(lYECHIMP0 As typeYECHIMP0, blnCompteAvis As Boolean, fic As Long) As Boolean
Dim V, I As Integer
Dim X As String, xSQL As String, Nb As Long
Dim bufferSortie As String

On Error GoTo Error_Handle
    
    bufferSortie = ""
    ctl_prtSAB_Echelles_ECHEDI01P2_Relevé = True
    xSQL = "select * from " & paramIBM_Library_SABSPE & ".YECHIMP0" _
   & " where ECHIMPJOB = " & lYECHIMP0.ECHIMPJOB _
     & " and ECHIMPJOBS = " & lYECHIMP0.ECHIMPJOBS _
     & " and ECHIMPCPT = '" & Trim(lYECHIMP0.ECHIMPCPT) & "'" _
     & " and ECHIMPDDEB = " & lYECHIMP0.ECHIMPDDEB _
     & " and ECHIMPDTRT = " & lYECHIMP0.ECHIMPDTRT
    Set rsSab = cnsab.Execute(xSQL, Nb)
    If rsSab.EOF Then
        ctl_mYECHIMP0 = lYECHIMP0
        ctl_blnAvis = False
        If blnCompteAvis Then
            ctl_prtSAB_Echelles_ECHEDI01P2_Relevé = False
            Exit Function
        End If
    Else
        V = rsYECHIMP0_GetBuffer(rsSab, ctl_mYECHIMP0)
        If Not IsNull(V) Then GoTo Error_Handle
        ctl_blnAvis = True
    End If
    'If ctl_blnNewPage Then
        bufferSortie = Retourne_Num_Client(Trim(ctl_mYECHIMP0.ECHIMPCPT)) & ";"
        bufferSortie = bufferSortie & ctl_mYECHIMP0.ECHIMPCPT & ";"
        If isBanque(Trim(ctl_mYECHIMP0.ECHIMPCPT)) Then
            bufferSortie = bufferSortie & "x" & ";"
        End If
        Print #fic, bufferSortie
    'End If

    Exit Function

Error_Handle:
V = Error
Error_MsgBox:
MsgBox "prtSAB_Echelles_ECHEDI01P2_Relevé" & Error, vbCritical, V
End Function


Public Sub fgSelect_Color(lRow As Integer, lColor As Long, lColor_Old As Long)
Dim mRow As Integer, I As Integer
mRow = fgSelect.Row

If lRow > 0 And lRow < fgSelect.Rows Then
    fgSelect.Row = lRow
    For I = 0 To fgSelect_arrIndex
        fgSelect.Col = I: fgSelect.CellBackColor = lColor_Old
    Next I
End If
lRow = 0
If mRow > 0 Then
    fgSelect.Row = mRow
    If fgSelect.Row > 0 Then
        lRow = fgSelect.Row
        lColor_Old = fgSelect.CellBackColor
        For I = fgSelect_arrIndex To 0 Step -1
          fgSelect.Col = I: fgSelect.CellBackColor = lColor
        Next I
        fgSelect.LeftCol = 0
    End If
End If

End Sub
Private Sub fgSelect_Display_1()
Dim K As Long, X As String
Dim blnOk As Boolean, blnDisplay As Boolean
Dim wIndex As Long

On Error GoTo Error_Handler
ssTab1.Tab = 0
fgSelect.Visible = False
fgSelect_Reset
fgSelect.Rows = 1
fgSelect.FormatString = fgSelect_FormatString
cmdPrint.Enabled = False
currentAction = "fgselect_Display"

For K = 0 To filW.ListCount - 1
    fgSelect.Rows = fgSelect.Rows + 1
    fgSelect.Row = fgSelect.Rows - 1

    filW.ListIndex = K
    fgSelect_DisplayLine_1 K

Next K

Call lstErr_Clear(lstErr, cmdContext, "Nb enregistrements : " & fgSelect.Rows - 1): DoEvents
If fgSelect.Rows > 1 Then
'    fgSelect_Sort1 = 0: fgSelect_Sort2 = 2: fgSelect_Sort
    cmdPrint.Enabled = True
End If
fgSelect.Visible = True
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    ssTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction



End Sub
Public Sub fgSelect_DisplayLine_1(lIndex As Long)
Dim X As String, K As Integer, K1 As Integer

On Error Resume Next
    X = filW.FileName
    K = InStr(1, X, ".")
    fgSelect.Col = 0: fgSelect.Text = Mid$(X, 1, K - 1) ' user
    K1 = K + 1
    K = InStr(K1, X, "_")
    fgSelect.Col = 1: fgSelect.Text = Mid$(X, K1, K - K1) 'date
    K1 = K + 1
    K = InStr(K1, X, "_")
    fgSelect.Col = 2: fgSelect.Text = Mid$(X, K1, K - K1) 'heure
    K1 = K + 1
    K = InStr(K1, X, "_")
    fgSelect.Col = 3: fgSelect.Text = Mid$(X, K1, K - K1) 'etat
    K1 = K + 1
    K = InStr(K1, X, "_")
    fgSelect.Col = 4: fgSelect.Text = Mid$(X, K1, K - K1) 'job
    K1 = K + 1
    K = InStr(K1, X, ".")
    fgSelect.Col = 5: fgSelect.Text = Mid$(X, K1, K - K1) 'séquence
    
    fgSelect.Col = fgSelect_arrIndex: fgSelect.Text = lIndex
End Sub


Private Sub lstSelect_Load_1()
Dim K  As Long, K1 As Long

On Error GoTo Error_Handler
ssTab1.Tab = 0
fgSelect.Visible = False
cmdPrint.Enabled = False
currentAction = "lstSelect_Load_1"
cmdSelect_Ok_Caption = "Lancer la requête"

dirW.PATH = paramEditionSplf_Folder
lstSelect.Clear
K1 = Len(Trim(dirW.PATH))
For K = 0 To dirW.ListCount - 1
    X = dirW.List(K)
    lstSelect.AddItem Mid$(X, K1 + 2, Len(X) - K1)
Next K
blnControl = False
If lstSelect.ListCount > 0 Then Call lst_Scan(constProduction, lstSelect)
blnControl = True
Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    ssTab1.Tab = 0
    MsgBox V, vbCritical, Me.Name & " : " & currentAction



End Sub


Public Sub fgSelect_Reset()
fgSelect.Clear
fgSelect_Sort1 = 0: fgSelect_Sort2 = 0
fgSelect_Sort1_Old = -1
fgSelect_RowDisplay = 0: fgSelect_RowClick = 0
fgSelect_arrIndex = fgSelect.Cols - 1
blnfgSelect_DisplayLine = False
fgSelect_SortAD = 6
fgSelect.LeftCol = 0

End Sub



Public Sub fgSelect_Sort()
If fgSelect.Rows > 1 Then
    fgSelect.Row = 1
    fgSelect.RowSel = fgSelect.Rows - 1
    
    If fgSelect_Sort1_Old = fgSelect_Sort1 Then
        If fgSelect_SortAD = 5 Then
            fgSelect_SortAD = 6
        Else
            fgSelect_SortAD = 5
        End If
    Else
        fgSelect_SortAD = 5
    End If
    fgSelect_Sort1_Old = fgSelect_Sort1
    
    fgSelect.Col = fgSelect_Sort1
    fgSelect.ColSel = fgSelect_Sort2
    fgSelect.Sort = fgSelect_SortAD
End If

End Sub
Public Sub fgSelect_SortX(lK As Integer)
Dim I As Integer, X As String
Dim wIndex As Integer
For I = 1 To fgSelect.Rows - 1
    fgSelect.Row = I
    fgSelect.Col = fgSelect_arrIndex
    wIndex = Val(fgSelect.Text)
    Select Case lK
    End Select
    fgSelect.Col = fgSelect_arrIndex - 1
    fgSelect.Text = X
Next I


fgSelect_Sort1 = fgSelect_arrIndex - 1: fgSelect_Sort2 = fgSelect_arrIndex - 1
fgSelect_Sort
End Sub
'---------------------------------------------------------
Public Sub arrTag_Set(ByVal B As Boolean)
'---------------------------------------------------------
For I = 0 To arrTagNb: arrTag(I) = B: Next I
End Sub


Public Sub cmdContext_Quit()
Unload Me
End Sub




Private Sub ctl_prtSAB_Echelles_ECHAVI02P1(lFile As String, lECHIMPJOBS As Integer)
Dim V, X As String
Dim xIn As String
Dim idFile As Integer
Dim zYECHIMP0 As typeYECHIMP0, xYECHIMP0 As typeYECHIMP0
Dim blnInsert As Boolean, blnECHIMPCPT As Boolean
Dim K As Integer, X8 As String, wAAAA As Long
Dim wMontant As Currency, wSens As String, wValeur As Long
Dim kECHIMPAD As Integer
'On Error GoTo Error_Handle

rsYECHIMP0_Init zYECHIMP0
blnInsert = False

idFile = FreeFile
Open lFile For Input As #idFile

Do Until EOF(1)
    Line Input #idFile, xIn
    
    Select Case Mid$(xIn, 1, 3)
        Case "$  "
                    If Mid$(xIn, 24, 10) <> "ECHAVI02P1" Then V = "Ce n'est pas un état ECHAVI02P1": GoTo Error_MsgBox
                    
                    zYECHIMP0.ECHIMPDTRT = Val(Mid$(xIn, 5, 8))
                    zYECHIMP0.ECHIMPJOB = Val(Mid$(xIn, 13, 6))
                    zYECHIMP0.ECHIMPJOBS = Val(Mid$(xIn, 19, 5))
                    lECHIMPJOBS = zYECHIMP0.ECHIMPJOBS
                    xIn = "delete from " & paramIBM_Library_SABSPE & ".YECHIMP0 where echimpjob=" & zYECHIMP0.ECHIMPJOB & " and echimpjobs=" & zYECHIMP0.ECHIMPJOBS
                    Call FEU_ROUGE
                    Set rsSab = cnsab.Execute(xIn)
                    Call FEU_VERT
       Case "010"
                   If blnInsert Then
                        V = sqlYECHIMP0_Insert(xYECHIMP0)
                        If Not IsNull(V) Then GoTo Error_MsgBox
                    End If
                   
                   blnInsert = True
                   blnECHIMPCPT = False
                   zYECHIMP0.ECHIMPSEQ = zYECHIMP0.ECHIMPSEQ + 1
                   xYECHIMP0 = zYECHIMP0
        Case "$$ "
                   If blnInsert Then
                        V = sqlYECHIMP0_Insert(xYECHIMP0)
                        If Not IsNull(V) Then GoTo Error_MsgBox
                    End If
                        
        Case Else
            If Not blnECHIMPCPT Then
            '________________________________________________________________________________
                K = InStr(xIn, "N/REF")
                If K > 0 Then
                    xYECHIMP0.ECHIMPNREF = Mid$(xIn, 25, 6)
                    kECHIMPAD = 0
                Else
                    K = InStr(xIn, "Date d'opération :")
                    If K > 0 Then
                        xYECHIMP0.ECHIMPDOPE = dateX8_N8(Mid$(xIn, 26, 8))
                    Else
                        K = InStr(xIn, "Arrété du")
                        If K > 0 Then
                            xYECHIMP0.ECHIMPDDEB = dateX8_N8(Mid$(xIn, 17, 8))
                            wAAAA = Int(xYECHIMP0.ECHIMPDDEB / 10000) * 10000
                            xYECHIMP0.ECHIMPDFIN = dateX8_N8(Mid$(xIn, 29, 8))
                        Else
                            K = InStr(xIn, "COMPTE :")
                            If K > 0 Then
                                blnECHIMPCPT = True
                                xYECHIMP0.ECHIMPCPT = Mid$(xIn, 16, 20)
                                xYECHIMP0.ECHIMPDEV = Mid$(xIn, 37, 3)
                            End If
                        End If
                    End If
                End If
                kECHIMPAD = kECHIMPAD + 1
                Select Case kECHIMPAD
                    Case 1: xYECHIMP0.ECHIMPAD1 = Mid$(xIn, 52, 32)
                    Case 2: xYECHIMP0.ECHIMPAD2 = Mid$(xIn, 52, 32)
                    Case 3: xYECHIMP0.ECHIMPAD3 = Mid$(xIn, 52, 32)
                    Case 4: xYECHIMP0.ECHIMPAD4 = Mid$(xIn, 52, 32)
                    Case 5: xYECHIMP0.ECHIMPAD5 = Mid$(xIn, 52, 32)
                    Case 6: xYECHIMP0.ECHIMPAD6 = Mid$(xIn, 52, 32)
                    Case 7: xYECHIMP0.ECHIMPAD7 = Mid$(xIn, 52, 32)
                End Select
        Else
            '________________________________________________________________________________
            If Mid$(xIn, 7, 3) = "ECH" Then
                X = Replace(Trim(Mid$(xIn, 22, 20)), ".", "")
                wMontant = CCur(X)
                wSens = Mid$(xIn, 44, 1)
                wValeur = wAAAA + Val(Mid$(xIn, 50, 2)) * 100 + Val(Mid$(xIn, 47, 2))
                Select Case Trim(Mid$(xIn, 54, 32))
                    Case "Intérêts créditeurs"
                            xYECHIMP0.ECHIMPICRM = wMontant
                            xYECHIMP0.ECHIMPICRS = wSens
                            xYECHIMP0.ECHIMPICRV = wValeur
                            Line Input #idFile, xIn
                            xYECHIMP0.ECHIMPICRT = CDbl(Trim(Mid$(xIn, 66, 11)))
                    Case "Intérêts débiteurs"
                            xYECHIMP0.ECHIMPIDEM = wMontant
                            xYECHIMP0.ECHIMPIDES = wSens
                            xYECHIMP0.ECHIMPIDEV = wValeur
                            Line Input #idFile, xIn
                            xYECHIMP0.ECHIMPIDET = CDbl(Trim(Mid$(xIn, 59, 8)))
                    Case "Com.de mouvements", "Com. de mouvements"
                           xYECHIMP0.ECHIMPIDEV = wValeur
                           xYECHIMP0.ECHIMPCMVT = wMontant
                    Case "Com. de plus fort découvert"
                            xYECHIMP0.ECHIMPCPFD = wMontant
                            xYECHIMP0.ECHIMPIDEV = wValeur
                    Case "Com. de compte", "Com. de tenue de compte", "Frais de tenue de compte"
                            xYECHIMP0.ECHIMPCCPT = wMontant
                            xYECHIMP0.ECHIMPIDEV = wValeur
                            '===============================
                            'TODO Prélèvement libératoire
                            '===============================
                  Case Else
                        MsgBox "ECHAVI02P1 : non traité :" & Error, vbCritical, Mid$(xIn, 54, 32)
                      
               End Select
            Else
                If Mid$(xIn, 7, 6) = " Total" Then
                    X = Replace(Trim(Mid$(xIn, 22, 20)), ".", "")
                    xYECHIMP0.ECHIMPMON = CCur(X)
                    xYECHIMP0.ECHIMPMONS = Mid$(xIn, 44, 1)
                End If
             End If
       End If

            
   End Select
Loop

Close idFile

Exit Sub

Error_Handle:
V = Error
Error_MsgBox:
MsgBox "prtSAB_Echelles_ECHAVI02P1" & Error, vbCritical, V
Close
End Sub

Private Sub btnControle_Click()
Dim ret As Long
Dim ffic As Long
Dim wECHIMPJOBS As Integer
Dim blnNostro As Boolean, blnSoldeZ As Boolean, blnCompteAvis As Boolean, selCompte As String
Dim xFileName As String, xFileName_Avis As String

    ret = MsgBox("Le fichier CSV nommé echelles_aaaammjj.csv sera déposé dans le répertoire c:\temp" & vbCrLf & "Voulez-vous continuer ?", vbYesNo + vbQuestion + vbDefaultButton2, "Impression des ECHELLES")
    If ret = vbNo Then Exit Sub
    
    If chkUpdate_Import = "1" Then Call ctl_prtSAB_Echelles_ECHAVI02P1(filW2.PATH & "\" & xFileName_ECHAVI02P1, wECHIMPJOBS)
    
    ffic = FreeFile
    Open "c:\temp\echelles_" & Left(DSYS_Time, 4) & "_" & Mid(DSYS_Time, 5, 2) & "_" & Mid(DSYS_Time, 7, 2) & ".csv" For Output As #ffic
    Print #ffic, "Client; Compte; Banque"
    blnNostro = IIf(chkUpdate_Nostro = "1", False, True)
    blnSoldeZ = IIf(chkUpdate_SoldeZ = "1", False, True)
    blnCompteAvis = IIf(chkUpdate_Avis = "1", True, False)
    selCompte = Trim(txtUpdate_Compte)
    xFileName = filW2.PATH & "\" & xFileName_ECHEDI01P2
    Call ctl_prtSAB_Echelles_ECHEDI01P2(xFileName, wECHIMPJOBS, blnNostro, blnSoldeZ, blnCompteAvis, selCompte, ffic)
    
    Close #ffic
    
End Sub

Private Sub cboSelect_SQL_Click()
cmdSelect_SQL_K = Mid$(cboSelect_SQL, 1, 1)
If blnControl Then
    Me.Enabled = False: Me.MousePointer = vbHourglass
    Select Case cmdSelect_SQL_K
        Case "1": lstSelect_Load_1
    End Select
    Me.Enabled = True: Me.MousePointer = 0
End If
End Sub


Private Sub cmdContext_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
MouseMoveActiveControl_Set cmdContext

End Sub

Private Sub cmdPrint_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
MouseMoveActiveControl_Set cmdPrint

End Sub

'---------------------------------------------------------
Public Sub cmdReset()
'---------------------------------------------------------
Dim I As Integer

blnControl = False
usrColor_Set

cmdContext.Caption = constcmdRechercher: blnMsgBox_Quit = False
arrTag_Set False
lstErr.Visible = False
blncmdOk_Visible = False: blncmdSave_Visible = False
currentAction = ""
fraUpdate.Visible = False
chkUpdate_Archivage.Enabled = BIA_Echelles_Aut.Xspécial = True
chkUpdate_Import.Enabled = BIA_Echelles_Aut.Xspécial = True
blnAuto = False
blnAuto_Ok = False

libRéférenceInterne = ""
blnControl = True
cboSelect_SQL.ListIndex = 0

End Sub
Public Sub Form_Init()
Call lstErr_Clear(lstErr, cmdContext, "Initialisation ")
DoEvents

ssTab1.Tab = 0

blnControl = False

fgSelect.Visible = False
fgSelect_FormatString = fgSelect.FormatString
cboSelect_SQL.Clear
cboSelect_SQL.AddItem "1 - Impression Echelles"

cmdReset



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
    Select Case cmdSelect_SQL_K
    '    Case "2": cmdPrint_Ok_2
    '    Case "3": cmdPrint_Ok_3
    End Select

Me.Show

Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdUpdate_Ok_Click()
Dim V, X As String
Dim xFileName As String, xFileName_Avis As String, K As Integer
Dim xJob As String
Dim wECHIMPJOBS As Integer
Dim blnNostro As Boolean, blnSoldeZ As Boolean, blnCompteAvis As Boolean, selCompte As String
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> Début du traitement"): DoEvents
X = MsgBox("1 - Imprimante sans RECTO / VERSO" & vbCrLf & "2 - faire un archivage PDF ", vbYesNo + vbQuestion + vbDefaultButton2, "Impression des ECHELLES")
If X = vbNo Then GoTo Exit_sub

If chkUpdate_Import = "1" Then Call prtSAB_Echelles_ECHAVI02P1(filW2.PATH & "\" & xFileName_ECHAVI02P1, wECHIMPJOBS)

'$JPL 20141203 archivage automatique
X = UCase$(Trim(Printer.Devicename))
If InStr(1, X, "PDF") <= 0 Then
    prtPgmName = "prtSAB_Echelles"
Else
    'D'abord l'édition globale
    prtPgmName = paramServer("\\Facturation\") & "Echelles\" & DSYS_Time & "Relevé_Echelles.pdf"
    If paramEnvironnement = constTest Then prtPgmName = Replace(prtPgmName, "Facturation", "Test\Facturation")
End If

blnNostro = IIf(chkUpdate_Nostro = "1", False, True)
blnSoldeZ = IIf(chkUpdate_SoldeZ = "1", False, True)
blnCompteAvis = IIf(chkUpdate_Avis = "1", True, False)
selCompte = Trim(txtUpdate_Compte)
xFileName = filW2.PATH & "\" & xFileName_ECHEDI01P2
Call prtSAB_Echelles_ECHEDI01P2(xFileName, wECHIMPJOBS, blnNostro, blnSoldeZ, blnCompteAvis, selCompte)

If chkUpdate_Archivage = "1" Then
    msFileSystem.MoveFile filW2.PATH & "\" & xFileName_ECHEDI01P2, paramEditionSplf_Folder & "Archive\" & xFileName_ECHEDI01P2
    If xFileName_Avis <> "" Then msFileSystem.MoveFile filW2.PATH & "\" & xFileName_Avis, paramEditionSplf_Folder & "Archive\" & xFileName_Avis
    blnControl = False
    If lstSelect.ListCount > 0 Then Call lst_Scan(constArchive, lstSelect)
    blnControl = True
End If

Exit_sub:
    fraUpdate.Visible = False
    Call lstErr_AddItem(lstErr, cmdContext, "< Fin du Traitement"): DoEvents
    Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdUpdate_Quit_Click()
Unload Me
End Sub

Private Sub fgSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim K As Long, xJobSeq As String
Me.Enabled = False
On Error Resume Next
If y <= fgSelect.RowHeightMin Then
        Select Case fgSelect.Col
            Case 0: fgSelect_Sort1 = 0: fgSelect_Sort2 = 0: fgSelect_Sort
            Case 1:  fgSelect_Sort1 = 1: fgSelect_Sort2 = 1: fgSelect_Sort
            Case 2: fgSelect_Sort1 = 2: fgSelect_Sort2 = 2: fgSelect_Sort
           Case fgSelect_arrIndex:  fgSelect_SortX fgSelect_arrIndex
        End Select
Else
    If fgSelect.Rows > 1 Then
        Me.Enabled = False: Me.MousePointer = vbHourglass
        fgSelect.Col = fgSelect_arrIndex:  filW.ListIndex = CLng(fgSelect.Text)
        Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
        xFileName_ECHEDI01P2 = filW.FileName
        
        fgSelect.Col = 5
        xJobSeq = Format$(CInt(Val(Trim(fgSelect.Text)) + 2), "00000")
        fgSelect.Col = 4
        xFileName_ECHAVI02P1 = "ECHAVI02P1_" & Trim(fgSelect.Text) & "_" & xJobSeq & ".txt"
        chkUpdate_Import = "1"
        chkUpdate_Import.Enabled = True
        If Trim(Dir(xFileName_ECHAVI02P1)) = "" Then
            filW2.PATH = paramEditionSplf_Folder & Trim(lstSelect.Text) & "\"
            filW2.Pattern = "x.xxx"
            filW2.Pattern = "*" & Mid$(xFileName_ECHAVI02P1, 1, Len(xFileName_ECHAVI02P1) - 9) & "*"
            If filW2.ListCount = 0 Then
                chkUpdate_Import = "0"
                chkUpdate_Import.Enabled = False
                X = MsgBox("Voulez-vous continuer (il n'y a pas de fichier d'avis d'opération)?", vbYesNo + vbQuestion + vbDefaultButton2, "manque " & xFileName_ECHAVI02P1)
                If X = vbNo Then GoTo Exit_sub
            Else
                If filW2.ListCount = 1 Then
                    filW2.ListIndex = 0
                    xFileName_ECHAVI02P1 = filW2.FileName
                Else
                    MsgBox "Il y a plus d'un fichier avis  d'opération"
                    GoTo Exit_sub
                End If
                'xFileName_Avis = filW2.FileName
                'xFileName = filW2.PATH & "\" & filW2.FileName
                'Call lstErr_AddItem(lstErr, cmdContext, xFileName): DoEvents
            End If
        End If
        chkUpdate_Import.Caption = "Importer AVI02P1 => YECHIMP0" & xFileName_ECHAVI02P1
        fraUpdate.Visible = True
        Me.Enabled = True: Me.MousePointer = 0

   End If
End If
Exit_sub:
Me.Enabled = True
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
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
MouseMoveActiveControl_Reset
End Sub


Private Sub Form_Resize()
If mWindowState <> Me.WindowState Then
    If Me.WindowState = 0 Or Me.WindowState = 2 Then
        Elp_Form_Resize Me, mWindowState, mHeight_0, mWidth_0, mHeight_2, mWidth_2
    End If
End If

End Sub

Private Sub lstSelect_Click()
Dim xName As String, K As Integer
If blnControl Then
    Me.Enabled = False: Me.MousePointer = vbHourglass
        filW.PATH = paramEditionSplf_Folder & Trim(lstSelect.Text) & "\"
        filW.Pattern = "x.xxx"
        filW.Pattern = "*ECHEDI01P2*.txt"
        K = InStr(UCase$(filW.PATH), "ARCHIVE")
        If K > 0 Then
            chkUpdate_Archivage.Value = "0"
            chkUpdate_Archivage.Enabled = False
        Else
            chkUpdate_Archivage.Value = "1"
            chkUpdate_Archivage.Enabled = True
        End If
        
        fgSelect_Display_1
    Me.Enabled = True: Me.MousePointer = 0
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
Dim meUnit As typeUnit, X As String
mWindowState = Me.WindowState
If Me.WindowState <> frmUsr_Windowstate Then Me.WindowState = frmUsr_Windowstate

Call BiaPgmAut_Init(Mid$(Msg, 1, 12), BIA_Echelles_Aut)

blnSetfocus = True
Form_Init
blnAuto = False

Select Case UCase$(Trim(Mid$(Msg, 1, 12)))

    Case Else: blnAuto = False
                
                    
End Select


End Sub


Public Sub cmdContext_Return()
If ssTab1.Tab = 0 Then
'    If fraUpdate.Visible _
'   And fraUpdate_B.Enabled _
'    And cmdUpdate_Ok.Enabled Then cmdUpdate_Ok_Click: Exit Sub
Else
    If currentAction = "" Then
        If ssTab1.Tab > 0 Then
            ssTab1.Tab = 0
        Else
           'SendKeys "{TAB}"
           ' cmdSelect_Click
        End If
    End If
End If
End Sub









Private Sub mnuPrint0_All_Click()
Dim I As Long, K As Long
Me.Enabled = False: Me.MousePointer = vbHourglass
    

Me.Show

Me.Enabled = True: Me.MousePointer = 0



End Sub





Private Sub txtUpdate_Compte_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub


