VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmEUP_XCOM 
   AutoRedraw      =   -1  'True
   Caption         =   "EUP_LAB"
   ClientHeight    =   9150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13560
   Icon            =   "EUP_XCOM.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9150
   ScaleWidth      =   13560
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6720
      TabIndex        =   3
      Top             =   0
      Width           =   6252
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8655
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   13530
      _ExtentX        =   23865
      _ExtentY        =   15266
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "SEPA :EUP_XCOM"
      TabPicture(0)   =   "EUP_XCOM.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraTab0"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "....."
      TabPicture(1)   =   "EUP_XCOM.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraTab1"
      Tab(1).ControlCount=   1
      Begin VB.Frame fraTab1 
         Height          =   7932
         Left            =   -74880
         TabIndex        =   14
         Top             =   480
         Width           =   13212
         Begin VB.ListBox lstDétail 
            Height          =   5910
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   6372
         End
      End
      Begin VB.Frame fraTab0 
         Height          =   8205
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   13290
         Begin VB.Frame fraRetour 
            Height          =   4572
            Left            =   8280
            TabIndex        =   16
            Top             =   1200
            Visible         =   0   'False
            Width           =   4932
            Begin VB.Frame fraRetour_R2V 
               BackColor       =   &H00FFE0FF&
               Caption         =   "R2V - RETOUR REJET"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2052
               Left            =   120
               TabIndex        =   27
               Top             =   2400
               Width           =   4692
               Begin VB.TextBox txtRetour_R2V_MONSTATUS 
                  Height          =   288
                  Left            =   1080
                  TabIndex        =   31
                  Top             =   240
                  Width           =   1212
               End
               Begin VB.TextBox txtRetour_R2V_MONUSR 
                  Height          =   288
                  Left            =   1080
                  TabIndex        =   30
                  Top             =   600
                  Width           =   1212
               End
               Begin VB.TextBox txtRetour_R2V_MONFILE 
                  Height          =   288
                  Left            =   1080
                  TabIndex        =   29
                  Top             =   960
                  Width           =   1212
               End
               Begin VB.CommandButton cmdRetour_R2V 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "Demande de téléchargement"
                  Height          =   612
                  Left            =   840
                  Style           =   1  'Graphical
                  TabIndex        =   28
                  Top             =   1320
                  Width           =   3132
               End
               Begin VB.Label lblRetour_R2V_MONSTATUS 
                  BackColor       =   &H00FFE0FF&
                  Caption         =   "Statut"
                  Height          =   252
                  Left            =   120
                  TabIndex        =   36
                  Top             =   240
                  Width           =   732
               End
               Begin VB.Label lblRetour_R2V_MONAMJ 
                  BackColor       =   &H00FFE0FF&
                  Caption         =   "Update"
                  Height          =   252
                  Left            =   2520
                  TabIndex        =   35
                  Top             =   240
                  Width           =   1812
               End
               Begin VB.Label lblRetour_R2V_MONUSR 
                  BackColor       =   &H00FFE0FF&
                  Caption         =   "Utilisateur"
                  Height          =   252
                  Left            =   120
                  TabIndex        =   34
                  Top             =   600
                  Width           =   732
               End
               Begin VB.Label lblRetour_R2V_MONJOB 
                  BackColor       =   &H00FFE0FF&
                  Caption         =   "horodatage"
                  Height          =   252
                  Left            =   2520
                  TabIndex        =   33
                  Top             =   600
                  Width           =   1812
               End
               Begin VB.Label lblRetour_R2V_MONFILE 
                  BackColor       =   &H00FFE0FF&
                  Caption         =   "Fichier"
                  Height          =   252
                  Left            =   120
                  TabIndex        =   32
                  Top             =   960
                  Width           =   612
               End
            End
            Begin VB.Frame fraRetour_R0V 
               BackColor       =   &H00E0FFFF&
               Caption         =   "R0V - RETOUR"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2052
               Left            =   120
               TabIndex        =   17
               Top             =   240
               Width           =   4692
               Begin VB.TextBox txtRetour_R0V_MONSTATUS 
                  Height          =   288
                  Left            =   1080
                  TabIndex        =   21
                  Top             =   240
                  Width           =   1212
               End
               Begin VB.TextBox txtRetour_R0V_MONUSR 
                  Height          =   288
                  Left            =   1080
                  TabIndex        =   20
                  Top             =   600
                  Width           =   1212
               End
               Begin VB.TextBox txtRetour_R0V_MONFILE 
                  Height          =   288
                  Left            =   1080
                  TabIndex        =   19
                  Top             =   960
                  Width           =   1212
               End
               Begin VB.CommandButton cmdRetour_R0V 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "Demande de téléchargement"
                  Height          =   612
                  Left            =   840
                  Style           =   1  'Graphical
                  TabIndex        =   18
                  Top             =   1320
                  Width           =   3132
               End
               Begin VB.Label lblRetour_R0V_MONSTATUS 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "Statut"
                  Height          =   252
                  Left            =   120
                  TabIndex        =   26
                  Top             =   240
                  Width           =   732
               End
               Begin VB.Label lblRetour_R0V_MONAMJ 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "Update"
                  Height          =   252
                  Left            =   2520
                  TabIndex        =   25
                  Top             =   240
                  Width           =   1812
               End
               Begin VB.Label lblRetour_R0V_MONUSR 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "Utilisateur"
                  Height          =   252
                  Left            =   120
                  TabIndex        =   24
                  Top             =   600
                  Width           =   732
               End
               Begin VB.Label lblRetour_R0V_MONJOB 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "horodatage"
                  Height          =   252
                  Left            =   2520
                  TabIndex        =   23
                  Top             =   600
                  Width           =   1812
               End
               Begin VB.Label lblRetour_R0V_MONFILE 
                  BackColor       =   &H00E0FFFF&
                  Caption         =   "Fichier"
                  Height          =   252
                  Left            =   120
                  TabIndex        =   22
                  Top             =   960
                  Width           =   612
               End
            End
         End
         Begin VB.ListBox lstDoc_IFS 
            Height          =   1620
            Left            =   240
            TabIndex        =   13
            Top             =   6240
            Width           =   4092
         End
         Begin VB.FileListBox filDoc_XCOM 
            Height          =   1455
            Left            =   8640
            TabIndex        =   10
            Top             =   6240
            Width           =   4212
         End
         Begin VB.ComboBox cboSelect_SQL 
            Height          =   315
            Left            =   9480
            TabIndex        =   9
            Text            =   "cboSelect_SQL"
            Top             =   120
            Width           =   3615
         End
         Begin VB.Frame fraSelect_Options 
            Height          =   1005
            Left            =   120
            TabIndex        =   7
            Top             =   120
            Width           =   6075
         End
         Begin VB.CommandButton cmdSelect_Ok 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Exécuter la requête"
            Height          =   525
            Left            =   11040
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   600
            Width           =   1095
         End
         Begin MSFlexGridLib.MSFlexGrid fgSelect 
            Height          =   4428
            Left            =   120
            TabIndex        =   8
            Top             =   1200
            Width           =   7920
            _ExtentX        =   13970
            _ExtentY        =   7805
            _Version        =   393216
            Rows            =   1
            Cols            =   8
            FixedCols       =   0
            RowHeightMin    =   350
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
            FormatString    =   $"EUP_XCOM.frx":047A
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
         Begin VB.Label lbllstDoc_IFS 
            BackColor       =   &H00FFC0FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblstdoc_IFS"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   240
            TabIndex        =   12
            Top             =   5880
            Visible         =   0   'False
            Width           =   4092
         End
         Begin VB.Label lblflDoc_XCOM 
            BackColor       =   &H00FFC0FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblflDoc_XCOM"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   8640
            TabIndex        =   11
            Top             =   5880
            Visible         =   0   'False
            Width           =   4212
         End
      End
   End
   Begin VB.CommandButton cmdContext 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Abandonner"
      Height          =   500
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   1200
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   500
      Left            =   13080
      Picture         =   "EUP_XCOM.frx":0508
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   500
   End
   Begin VB.Label libSelect 
      BackColor       =   &H00FFFED9&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   4905
      WordWrap        =   -1  'True
   End
   Begin VB.Menu mnuContext 
      Caption         =   "mnuContext"
      Visible         =   0   'False
      Begin VB.Menu mnuContextAbandonner 
         Caption         =   "Abandonner"
      End
      Begin VB.Menu mnuContextQuitter 
         Caption         =   "Quitter"
      End
   End
   Begin VB.Menu mnuPrint0 
      Caption         =   "mnuPrint0"
      Visible         =   0   'False
      Begin VB.Menu mnuSelect_Print_Liste 
         Caption         =   "Imprimer liste"
      End
      Begin VB.Menu mnuSelect_Print_Détail 
         Caption         =   "Imprimer liste détaillée"
      End
   End
End
Attribute VB_Name = "frmEUP_XCOM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------
Option Explicit

Dim mHeight_0 As Integer, mWidth_0 As Integer, mHeight_2 As Integer, mWidth_2 As Integer, mWindowState As Integer
Dim arrTag() As Boolean, arrTagNb As Integer
Dim lastActiveControl_Name  As String, currentActiveControl_Name As String, currentAction As String, currentError As String
Dim MouseMoveActiveControl_Name  As String, MouseMoveActiveControl As typeUsrColor
Dim blnMsgBox_Quit As Boolean, blnAddNew As Boolean, blnGlobalControl As Boolean, blnControl As Boolean
Dim X As String, I As Integer, Msg As String, valX As String, X1 As String, V As Variant, curX As Currency, dblX As Double
Dim intReturn As Integer
Dim EUPXCOM_Aut As typeAuthorization
Dim curX1 As Currency, curX2 As Currency
Dim blnAuto As Boolean

Dim fgSelect_FormatString As String, fgSelect_K As Integer
Dim fgSelect_RowDisplay As Integer, fgSelect_RowClick As Integer, fgSelect_ColClick As Integer
Dim fgSelect_ColorClick As Long, fgSelect_ColorDisplay As Long
Dim fgSelect_Sort1 As Integer, fgSelect_Sort2 As Integer
Dim fgSelect_SortAD As Integer, fgSelect_Sort1_Old As Integer
Dim fgSelect_arrIndex As Integer
Dim blnfgSelect_DisplayLine As Boolean

Dim blnTransaction As Boolean
Dim cmdSelect_SQL_K As String
'______________________________________________________________________

Dim wAMJMin As String, WAMJMax As String, wHmsMin As Long, wHmsMax As Long

Dim xYEUPMON2 As typeYEUPMON2, newYEUPMON2 As typeYEUPMON2, oldYEUPMON2 As typeYEUPMON2
Dim arrYEUPMON2() As typeYEUPMON2, arrYEUPMON2_Nb As Long, arrYEUPMON2_Max As Long, arrYEUPMON2_Index As Long


Dim xYEUPMON4 As typeYEUPMON4, newYEUPMON4 As typeYEUPMON4, oldYEUPMON4 As typeYEUPMON4
Dim arrYEUPMON4() As typeYEUPMON4, arrYEUPMON4_Nb As Long, arrYEUPMON4_Max As Long, arrYEUPMON4_Index As Long

Dim arrIFS_File(100) As String, arrIFS_File_Nb As Integer
'______________________________________________________________________
Dim blnEUP_XCOMTEST As Boolean

Dim paramXCOM As String
Dim paramSchtasks_Create As String
Dim paramSchtasks_Delete As String
Dim paramSchtasks_Query As String
Dim paramSchtasks_Run As String
Dim paramSchtasks_TN As String
Dim paramSchtasks_Bat As String
Dim paramSchtasks_Log As String
Dim paramSchtasks_Nom As String
Dim paramSEPA_Aller_IFS As String, paramSEPA_Retour_IFS As String
Dim paramSEPA_Aller_XCOM As String, paramSEPA_Retour_XCOM As String
Dim paramSEPA_Retour_XCOM_File As String
Dim paramSEPA_Aller_log As String, paramSEPA_Retour_log As String
Dim paramSEPA_Aller_cmd As String, paramSEPA_Retour_cmd As String
Dim paramSEPA_Aller_Archive As String, paramSEPA_Retour_Archive As String
Dim paramToken As String
Dim paramSEPA_Aller_Path_log As String, paramSEPA_Retour_Path_log As String
Dim paramSEPA_Aller_Détail_log As String, paramSEPA_Retour_Détail_log As String
Dim paramSEPA_Aller_Fax As String

'______________________________________________________________________

Dim oldYBIAMON0_R0V As typeYBIAMON0
Dim oldYBIAMON0_R2V As typeYBIAMON0
Dim oldYBIAMON0 As typeYBIAMON0, newYBIAMON0 As typeYBIAMON0
Public Function cmdSelect_SQL_6_Control_A0C(lFile As String)
Dim wMsgId As String
Dim wFile As String
Dim blnAssgnmt As Boolean
Dim MsgId As String
Dim xIn As String
Dim Fic As Long
Dim K1 As Long
Dim K2 As Long

    On Error GoTo Error_Handler
    cmdSelect_SQL_6_Control_A0C = "?"
    Call lstErr_AddItem(lstErr, cmdContext, lFile): DoEvents
    wMsgId = Replace(lFile, ".xml", "")
    wFile = paramSEPA_Aller_XCOM & lFile
    blnAssgnmt = False
    MsgId = ""
    K1 = -1
    K2 = -1
    If Dir(wFile) <> "" Then
        Fic = FreeFile
        Open wFile For Input As #Fic
        Do Until EOF(Fic)
            Line Input #Fic, xIn
            If Not blnAssgnmt Then
                If InStr(xIn, "<Assgnmt>") > 0 Then blnAssgnmt = True
            End If
            If blnAssgnmt Then
                If InStr(xIn, "<MsgId>") > 0 Then K1 = InStr(xIn, "<MsgId>") + 7
                If InStr(xIn, "</MsgId>") > 0 Then K2 = InStr(xIn, "</MsgId>")
                If K1 > -1 And K2 = -1 Then
                    MsgId = MsgId & Trim(mId(xIn, K1))
                End If
                If K1 > -1 And K2 > -1 And MsgId = "" Then
                    MsgId = mId(xIn, K1, K2 - K1)
                    blnAssgnmt = False
                ElseIf K1 > -1 And K2 > -1 And MsgId <> "" Then
                    MsgId = MsgId & Trim(mId(xIn, 1, K2 - 1))
                    blnAssgnmt = False
                End If
            End If
        Loop
    End If
    Close #Fic
    If MsgId <> "" Then
        Print #501, Now & " | " & "<MsgId> :" & Trim(MsgId)
        Print #501, Now & " | " & "Préparation du Fax"
        Call cmdPrint_SEPA_Aller_Fax(lFile, 1, 0)
    End If
    cmdSelect_SQL_6_Control_A0C = Null
    Exit Function
Error_Handler:
    V = Error
Error_MsgBox:
    Error_Route V
    Print #501, Now & " | " & V
End Function

Public Function cmdSelect_SQL_6_Control_NEW(dFile As String)
Dim xml_articleDoc As New DOMDocument
Dim xml_sections As IXMLDOMNodeList
Dim xml_nodes As IXMLDOMNodeList
Dim articleDoc As New DOMDocument
Dim sections As IXMLDOMNodeList
Dim dCheminComplet As String
Dim dIbanCrediteur() As String
Dim dMontant_s() As Currency
Dim dIdentifiant_s() As String
Dim dIdentifiant As String
Dim dLibelle_s() As String
Dim dMontant As Currency
Dim dDevise_s() As String
Dim passDevise As String
Dim passIbanCrediteur As String
Dim passIdentifiant As String
Dim ddTemp() As String
Dim dDevise As String
Dim dNbDetail As Long
Dim indice As Long
Dim dTemp As String
Dim dRep As String
Dim dFic As Long
Dim ou As Long
Dim K1 As Long
Dim K2 As Long
Dim I As Long
Dim u As Long

Dim xmlParentParentIdentifiant As String
Dim xmlParentIdentifiant As String
Dim xmlIdentifiant As String
Dim xmlParentParentMontant As String
Dim xmlParentMontant As String
Dim xmlMontant As String
Dim xmlParentParentNbDetail As String
Dim xmlParentNbDetail As String
Dim xmlNbDetail As String
Dim xmlParentParentIdentifiant_s As String
Dim xmlParentIdentifiant_s As String
Dim xmlIdentifiant_s As String
Dim xmlParentParentIbanCrediteur_s As String
Dim xmlParentIbanCrediteur_s As String
Dim xmlIbanCrediteur_s As String
Dim xmlParentParentMontant_s As String
Dim xmlParentMontant_s As String
Dim xmlMontant_s As String
Dim xmlParentParentLibelle_s As String
Dim xmlParentLibelle_s As String
Dim xmlLibelle_s As String
   
    On Error GoTo dError_Handler
    
    
    'Acquisition des données xml de description                                                                         -
    xmlParentParentIdentifiant = ""
    xmlParentIdentifiant = ""
    xmlIdentifiant = ""
    xmlParentParentMontant = ""
    xmlParentMontant = ""
    xmlMontant = ""
    xmlParentParentNbDetail = ""
    xmlParentNbDetail = ""
    xmlNbDetail = ""
    xmlParentParentIdentifiant_s = ""
    xmlParentIdentifiant_s = ""
    xmlIdentifiant_s = ""
    xmlParentParentIbanCrediteur_s = ""
    xmlParentIbanCrediteur_s = ""
    xmlIbanCrediteur_s = ""
    xmlParentParentMontant_s = ""
    xmlParentMontant_s = ""
    xmlMontant_s = ""
    xmlParentParentLibelle_s = ""
    xmlParentLibelle_s = ""
    xmlLibelle_s = ""
    xml_articleDoc.Load paramFolder_Local & "\sepa_description.xml"
    Set xml_sections = xml_articleDoc.selectNodes("/Document/" & mId(dFile, 1, 3) & "/*")
    For I = 0 To xml_sections.Length - 1
        Set xml_nodes = xml_sections(I).selectNodes("/Document/" & mId(dFile, 1, 3) & "/" & xml_sections(I).nodeName & "/*")
        Select Case UCase(xml_sections(I).nodeName)
            Case "IDENTIFIANT":
                For u = 0 To xml_nodes.Length - 1
                    Select Case UCase(xml_nodes(u).nodeName)
                        Case "PARENTPARENT": xmlParentParentIdentifiant = xml_nodes(u).Text
                        Case "PARENT":       xmlParentIdentifiant = xml_nodes(u).Text
                        Case "CLE":          xmlIdentifiant = xml_nodes(u).Text
                    End Select
                Next u
            Case "MONTANT":
                For u = 0 To xml_nodes.Length - 1
                    Select Case UCase(xml_nodes(u).nodeName)
                        Case "PARENTPARENT": xmlParentParentMontant = xml_nodes(u).Text
                        Case "PARENT":       xmlParentMontant = xml_nodes(u).Text
                        Case "CLE":          xmlMontant = xml_nodes(u).Text
                    End Select
                Next u
            Case "NBDETAIL":
                For u = 0 To xml_nodes.Length - 1
                    Select Case UCase(xml_nodes(u).nodeName)
                        Case "PARENTPARENT": xmlParentParentNbDetail = xml_nodes(u).Text
                        Case "PARENT":       xmlParentNbDetail = xml_nodes(u).Text
                        Case "CLE":          xmlNbDetail = xml_nodes(u).Text
                    End Select
                Next u
            Case "IDENTIFIANT_S":
                For u = 0 To xml_nodes.Length - 1
                    Select Case UCase(xml_nodes(u).nodeName)
                        Case "PARENTPARENT": xmlParentParentIdentifiant_s = xml_nodes(u).Text
                        Case "PARENT":       xmlParentIdentifiant_s = xml_nodes(u).Text
                        Case "CLE":          xmlIdentifiant_s = xml_nodes(u).Text
                    End Select
                Next u
            Case "IBANCREDITEUR_S":
                For u = 0 To xml_nodes.Length - 1
                    Select Case UCase(xml_nodes(u).nodeName)
                        Case "PARENTPARENT": xmlParentParentIbanCrediteur_s = xml_nodes(u).Text
                        Case "PARENT":       xmlParentIbanCrediteur_s = xml_nodes(u).Text
                        Case "CLE":          xmlIbanCrediteur_s = xml_nodes(u).Text
                    End Select
                Next u
            Case "MONTANT_S":
                For u = 0 To xml_nodes.Length - 1
                    Select Case UCase(xml_nodes(u).nodeName)
                        Case "PARENTPARENT": xmlParentParentMontant_s = xml_nodes(u).Text
                        Case "PARENT":       xmlParentMontant_s = xml_nodes(u).Text
                        Case "CLE":          xmlMontant_s = xml_nodes(u).Text
                    End Select
                Next u
            Case "LIBELLE_S":
                For u = 0 To xml_nodes.Length - 1
                    Select Case UCase(xml_nodes(u).nodeName)
                        Case "PARENTPARENT": xmlParentParentLibelle_s = xml_nodes(u).Text
                        Case "PARENT":       xmlParentLibelle_s = xml_nodes(u).Text
                        Case "CLE":          xmlLibelle_s = xml_nodes(u).Text
                    End Select
                Next u
        End Select
        xml_nodes.Reset
        Set xml_nodes = Nothing
    Next I
    '                                                                                                                     -
    
    cmdSelect_SQL_6_Control_NEW = "?"
    If xml_sections.Length > 0 Then
        dIdentifiant = ""
        dMontant = 0
        dDevise = ""
        dNbDetail = 0
        paramSEPA_Aller_XCOM = "C:\TEMP\test\"
        If Right(paramSEPA_Aller_XCOM, 1) <> "\" Then
            dRep = "\"
        Else
            dRep = ""
        End If
        dCheminComplet = paramSEPA_Aller_XCOM & dRep & dFile
        dFile = Replace(dFile, ".xml", "")
        
        Close #501
        Open paramSEPA_Aller_XCOM & dRep & dFile & ".txt" For Output As #501
        Print #501, Now & "_6_NEW__________________________________________________________"
        
        articleDoc.Load dCheminComplet
        Set sections = articleDoc.selectNodes("//*")
        'Identifiant
        If xmlIdentifiant <> "" Then
            For I = 0 To sections.Length - 1
                If "<" & UCase(xmlIdentifiant) = "<" & UCase(sections(I).nodeName) Then
                    If "<" & UCase(xmlParentIdentifiant) = "<" & UCase(sections(I).parentNode.nodeName) Then
                        If "<" & UCase(xmlParentParentIdentifiant) = "<" & UCase(sections(I).parentNode.parentNode.nodeName) Then
                            dIdentifiant = sections(I).Text
                            Exit For
                        End If
                    End If
                End If
            Next I
        End If
        'montant
        If xmlMontant <> "" Then
            dTemp = ""
            For I = 0 To sections.Length - 1
                dTemp = "<" & UCase(sections(I).nodeName)
                If "<" & UCase(xmlMontant) = dTemp Then
                    If "<" & UCase(xmlParentMontant) = "<" & UCase(sections(I).parentNode.nodeName) Then
                        If "<" & UCase(xmlParentParentMontant) = "<" & UCase(sections(I).parentNode.parentNode.nodeName) Then
                            dMontant = dMontant + CCur(sections(I).Text)
                            'nombre d'opérations
                            dNbDetail = dNbDetail + 1
                            'Devise
                            ddTemp = Split(sections(I).XML, " ")
                            For u = 0 To UBound(ddTemp)
                                ddTemp(u) = Replace(ddTemp(u), " ", "")
                                K1 = 0
                                K2 = 0
                                ou = InStr(UCase(ddTemp(u)), "CCY=")
                                If ou > 0 Then
                                    K1 = InStr(1, ddTemp(u), Chr(34))
                                    If K1 > 0 Then
                                        K2 = InStr(K1 + 1, ddTemp(u), Chr(34))
                                        If K2 > 0 Then
                                            dDevise = mId(ddTemp(u), K1 + 1, K2 - K1 - 1)
                                        End If
                                    End If
                                End If
                            Next u
                        End If
                    End If
                End If
            Next I
        End If
        'nombre d'opérations
        If xmlNbDetail <> "" Then
            For I = 0 To sections.Length - 1
                If "<" & UCase(xmlNbDetail) = "<" & UCase(sections(I).nodeName) Then
                    If "<" & UCase(xmlParentNbDetail) = "<" & UCase(sections(I).parentNode.nodeName) Then
                        If "<" & UCase(xmlParentParentNbDetail) = "<" & UCase(sections(I).parentNode.parentNode.nodeName) Then
                            dNbDetail = Val(sections(I).Text)
                            Exit For
                        End If
                    End If
                End If
            Next I
        End If
        '                                                                                                      -
        'contrôle des lignes détail                                                                            -
        '                                                                                                      -
        If xmlIdentifiant_s <> "" Then
            indice = 0
            passIdentifiant = ""
            passIbanCrediteur = ""
            passDevise = ""
            ReDim dLibelle_s(0 To dNbDetail)
            ReDim dIdentifiant_s(0 To dNbDetail)
            ReDim dMontant_s(0 To dNbDetail)
            ReDim dDevise_s(0 To dNbDetail)
            ReDim dIbanCrediteur_s(0 To dNbDetail)
            For I = 0 To sections.Length - 1
                If "<" & UCase(xmlIdentifiant_s) = "<" & UCase(sections(I).nodeName) Then
                    If "<" & UCase(xmlParentIdentifiant_s) = "<" & UCase(sections(I).parentNode.nodeName) Then
                        If "<" & UCase(xmlParentParentIdentifiant_s) = "<" & UCase(sections(I).parentNode.parentNode.nodeName) Then
                            indice = indice + 1
                            dIdentifiant_s(indice) = sections(I).Text
                        End If
                    End If
                End If
                If xmlIbanCrediteur_s <> "" Then
                    If "<" & UCase(xmlIbanCrediteur_s) = "<" & UCase(sections(I).nodeName) Then
                        If "<" & UCase(xmlParentIbanCrediteur_s) = "<" & UCase(sections(I).parentNode.nodeName) Then
                            If "<" & UCase(xmlParentParentIbanCrediteur_s) = "<" & UCase(sections(I).parentNode.parentNode.nodeName) Then
                                 dIbanCrediteur_s(indice) = sections(I).Text
                            End If
                        End If
                    End If
                End If
                If xmlLibelle_s <> "" Then
                    If "<" & UCase(xmlLibelle_s) = "<" & UCase(sections(I).nodeName) Then
                        If "<" & UCase(xmlParentLibelle_s) = "<" & UCase(sections(I).parentNode.nodeName) Then
                            If "<" & UCase(xmlParentParentLibelle_s) = "<" & UCase(sections(I).parentNode.parentNode.nodeName) Then
                                dLibelle_s(indice) = sections(I).Text
                            End If
                        End If
                    End If
                End If
                If xmlMontant_s <> "" Then
                    dTemp = "<" & UCase(sections(I).nodeName)
                    If "<" & UCase(xmlMontant_s) = dTemp Then
                        If "<" & UCase(xmlParentMontant_s) = "<" & UCase(sections(I).parentNode.nodeName) Then
                            If "<" & UCase(xmlParentParentMontant_s) = "<" & UCase(sections(I).parentNode.parentNode.nodeName) Then
                                'pour rester compatible avec l'ancienne fonction
                                dMontant_s(indice) = CCur(sections(I).Text)
                                'Devise
                                ddTemp = Split(sections(I).XML, " ")
                                For u = 0 To UBound(ddTemp)
                                    ddTemp(u) = Replace(ddTemp(u), " ", "")
                                    K1 = 0
                                    K2 = 0
                                    ou = InStr(UCase(ddTemp(u)), "CCY=")
                                    If ou > 0 Then
                                        K1 = InStr(1, ddTemp(u), Chr(34))
                                        If K1 > 0 Then
                                            K2 = InStr(K1 + 1, ddTemp(u), Chr(34))
                                            If K2 > 0 Then
                                                dDevise_s(indice) = mId(ddTemp(u), K1 + 1, K2 - K1 - 1)
                                                If dDevise_s(indice) <> dDevise Then V = "Devise détail  : " & dDevise_s(indice) & " <> " & dDevise: GoTo dError_MsgBox
                                            End If
                                        End If
                                    End If
                                Next u
                            End If
                        End If
                    End If
                End If
                If indice > 0 Then
                    If passIdentifiant <> "" And passIbanCrediteur <> "" And passDevise <> "" Then
                        passIdentifiant = dIdentifiant_s(indice)
                        passIbanCrediteur = dIbanCrediteur_s(indice)
                        passDevise = dDevise_s(indice)
                        V = cmdSelect_SQL_6_Control_détail(passIdentifiant, passIbanCrediteur, passDevise, 0)
                        If Not IsNull(V) Then GoTo dError_MsgBox
                        passIdentifiant = ""
                        passIbanCrediteur = ""
                        passDevise = ""
                    End If
                End If
            Next I
        End If
        'écriture dans la log                                                                      -
        If xmlIdentifiant <> "" Then
            Print #501, Now & " | <" & xmlIdentifiant & "> : " & dIdentifiant
        End If
        If xmlNbDetail <> "" Then
            Print #501, Now & " | <" & xmlNbDetail & "> : " & dNbDetail
        End If
        If xmlMontant <> "" Then
            If dMontant <> 0 Then
                Print #501, Now & " | <" & xmlMontant & " Ccy = " & dDevise & " " & Format(dMontant, "### ### ###.00")
            Else
                Print #501, Now & " | <" & xmlMontant & " Ccy = " & dDevise
            End If
        End If
        For I = 1 To indice
            If dMontant_s(I) <> 0 Then
                Print #501, Now & " | " & dIdentifiant_s(I) & " : " & dDevise_s(I) & " " & Format(dMontant_s(I), "### ### ###.00") & " : " & dIbanCrediteur_s(I) & " : " & dLibelle_s(I)
            Else
                Print #501, Now & " | " & dIdentifiant_s(I) & " : " & dIbanCrediteur_s(I) & " : " & dLibelle_s(I)
            End If
        Next I
        cmdSelect_SQL_6_Control_NEW = Null
        Print #501, Now & " | " & "contrôle Nb et montant OK : préparation du Fax"
        Call cmdPrint_SEPA_Aller_Fax(dFile & ".xml", CInt(dNbDetail), dMontant)
    End If
    Close #501
    Exit Function

dError_Handler:
     
      V = Error
dError_MsgBox:
    Error_Route V
    Print #501, Now & " | " & V

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
        For I = 0 To fgSelect_arrIndex
          fgSelect.Col = I: fgSelect.CellBackColor = lColor
        Next I
        fgSelect.Col = 0
    End If
End If

End Sub
Private Sub fgSelect_Display_1()
Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
currentAction = "fgselect_Display"
SSTab1.Tab = 0
fgSelect.Visible = False
fgSelect_Reset

fgSelect.Rows = 1
fgSelect.FormatString = fgSelect_FormatString
    
For I = 1 To arrYEUPMON2_Nb
         
    xYEUPMON2 = arrYEUPMON2(I)
        fgSelect.Rows = fgSelect.Rows + 1
        fgSelect.Row = fgSelect.Rows - 1
        fgSelect_DisplayLine_1 I
Next I

fgSelect.Visible = True
Call lstErr_AddItem(lstErr, cmdContext, "Opérations : " & arrYEUPMON2_Nb): DoEvents
'If fgSelect.Rows > 1 Then
'    fgSelect_Sort1 = 0: fgSelect_Sort2 = 2: fgSelect_Sort
'End If

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction



End Sub

Private Sub fgSelect_Display_2()
Dim I As Long
Dim blnOk As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
currentAction = "fgselect_Display"
SSTab1.Tab = 0
fgSelect.Visible = False
fgSelect_Reset

fgSelect.Rows = 1
fgSelect.FormatString = fgSelect_FormatString
    
For I = 1 To arrYEUPMON4_Nb
         
    xYEUPMON4 = arrYEUPMON4(I)
        fgSelect.Rows = fgSelect.Rows + 1
        fgSelect.Row = fgSelect.Rows - 1
        fgSelect_DisplayLine_2 I
Next I

fgSelect.Visible = True
Call lstErr_AddItem(lstErr, cmdContext, "Opérations : " & arrYEUPMON4_Nb): DoEvents
'If fgSelect.Rows > 1 Then
'    fgSelect_Sort1 = 0: fgSelect_Sort2 = 2: fgSelect_Sort
'End If

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction



End Sub

Private Sub fgDétail_Display()
Dim I As Long, xIn As String
Dim blnOk As Boolean, blnDisplay As Boolean

On Error GoTo Error_Handler
SSTab1.Tab = 1

lstDétail.Clear
X = paramSEPA_Aller_IFS & "\" & xYEUPMON2.EUPMON2FIC
Open X For Input As #1
X = ""
Do Until EOF(1)
    DoEvents
    Line Input #1, xIn
    lstDétail.AddItem Trim(xIn)

    
Loop
Close #1

Call lstErr_AddItem(lstErr, cmdContext, "Opérations : " & arrYEUPMON2_Nb): DoEvents

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction



End Sub

Private Sub arrYEUPMON2_SQL(xWhere As String)
Dim V
Dim X As String, xSQL As String
On Error GoTo Error_Handler
ReDim arrYEUPMON2(101)
arrYEUPMON2_Max = 100: arrYEUPMON2_Nb = 0

Set rsSab = Nothing

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YEUPMON2 " & xWhere & " order by EUPMON2FIC"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    V = rsYEUPMON2_GetBuffer(rsSab, xYEUPMON2)

     If Not IsNull(V) Then
         If Not blnAuto Then MsgBox V, vbCritical, "frmEUP_XCOM.fgselect_Display"
        '' Exit Sub
     Else
         arrYEUPMON2_Nb = arrYEUPMON2_Nb + 1
         If arrYEUPMON2_Nb > arrYEUPMON2_Max Then
             arrYEUPMON2_Max = arrYEUPMON2_Max + 50
             ReDim Preserve arrYEUPMON2(arrYEUPMON2_Max)
         End If
         
         arrYEUPMON2(arrYEUPMON2_Nb) = xYEUPMON2
    End If
    rsSab.MoveNext

Loop

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Private Sub arrYEUPMON4_SQL(xWhere As String)
Dim V
Dim X As String, xSQL As String
On Error GoTo Error_Handler
ReDim arrYEUPMON4(101)
arrYEUPMON4_Max = 100: arrYEUPMON4_Nb = 0

Set rsSab = Nothing

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YEUPMON4 " & xWhere & " order by EUPMON2FIC"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    V = rsYEUPMON4_GetBuffer(rsSab, xYEUPMON4)

     If Not IsNull(V) Then
         If Not blnAuto Then MsgBox V, vbCritical, "arrYEUPMON4_SQL"
        '' Exit Sub
     Else
         arrYEUPMON4_Nb = arrYEUPMON4_Nb + 1
         If arrYEUPMON4_Nb > arrYEUPMON4_Max Then
             arrYEUPMON4_Max = arrYEUPMON4_Max + 50
             ReDim Preserve arrYEUPMON4(arrYEUPMON4_Max)
         End If
         
         arrYEUPMON4(arrYEUPMON4_Nb) = xYEUPMON4
    End If
    rsSab.MoveNext

Loop

Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction

End Sub

Public Sub fgSelect_DisplayLine_1(lIndex As Long)
On Error Resume Next
fgSelect.Col = 0: fgSelect.Text = xYEUPMON2.EUPMON2FIC
fgSelect.Col = 1: fgSelect.Text = xYEUPMON2.EUPMON2STA
If xYEUPMON2.EUPMON2DUP <> 0 Then fgSelect.Col = 2: fgSelect.Text = dateImp10(xYEUPMON2.EUPMON2DUP) & " " & timeImp8(xYEUPMON2.EUPMON2HUP)
If xYEUPMON2.EUPMON2DEN <> 0 Then fgSelect.Col = 3: fgSelect.Text = dateImp10(xYEUPMON2.EUPMON2DEN) & " " & timeImp8(xYEUPMON2.EUPMON2HEN)
If xYEUPMON2.EUPMON2DMO <> 0 Then fgSelect.Col = 4: fgSelect.Text = dateImp10(xYEUPMON2.EUPMON2DMO) & " " & timeImp8(xYEUPMON2.EUPMON2HMO)

fgSelect.Col = fgSelect_arrIndex: fgSelect.Text = lIndex



End Sub

Public Sub fgSelect_DisplayLine_2(lIndex As Long)
On Error Resume Next
fgSelect.Col = 0: fgSelect.Text = xYEUPMON4.EUPMON2FIC
fgSelect.Col = 1: fgSelect.Text = xYEUPMON4.EUPMON2STA
If xYEUPMON4.EUPMON2DUP <> 0 Then fgSelect.Col = 2: fgSelect.Text = dateImp10(xYEUPMON4.EUPMON2DUP) & " " & timeImp8(xYEUPMON4.EUPMON2HUP)
If xYEUPMON4.EUPMON2DEN <> 0 Then fgSelect.Col = 3: fgSelect.Text = dateImp10(xYEUPMON4.EUPMON2DEN) & " " & timeImp8(xYEUPMON4.EUPMON2HEN)
If xYEUPMON4.EUPMON2DMO <> 0 Then fgSelect.Col = 4: fgSelect.Text = dateImp10(xYEUPMON4.EUPMON2DMO) & " " & timeImp8(xYEUPMON4.EUPMON2HMO)

fgSelect.Col = fgSelect_arrIndex: fgSelect.Text = lIndex



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
For I = 1 To fgSelect.Rows - 1
    fgSelect.Row = I
    If lK = 2 Then
        fgSelect.Col = 2
        X = fgSelect.Text
    Else
        X = ""
    End If
    
    fgSelect.Col = 3
    X = X & Format$(Val(fgSelect.Text), "000000000000000.00")
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


Public Sub Msg_Rcv(Msg As String)
'---------------------------------------------------------
mWindowState = Me.WindowState
If Me.WindowState <> frmUsr_Windowstate Then Me.WindowState = frmUsr_Windowstate

Call BiaPgmAut_Init(mId$(Msg, 1, 12), EUPXCOM_Aut)
If UCase$(Trim(mId$(Msg, 1, 12))) = "EUP_XCOMTEST" Then blnEUP_XCOMTEST = True
Form_Init

Select Case UCase$(Trim(mId$(Msg, 1, 12)))
    Case "@EUP_XCOM": blnAuto = True
                    
                    cmdSelect_SQL_5
                    cmdSelect_SQL_6
                    cmdSelect_SQL_7
                    Unload Me
           
    Case Else: blnAuto = False
                    cmdSelect_SQL_1

End Select


End Sub


Public Sub cmdSendMail_SEPA_Alerte(lSubject As String, lMessage As String, lFile As String)
Dim wSendMail As typeSendMail
Dim bgColor As String
Dim X As String
wSendMail.FromDisplayName = "SEPA_ALERTE"
wSendMail.RecipientDisplayName = "XCOM"
bgColor = "<body bgcolor = #FF0000>"

wSendMail.Subject = lSubject
wSendMail.Attachment = lFile
wSendMail.Message = bgColor & lMessage
wSendMail.AsHTML = True

srvSendMail.Monitor wSendMail
End Sub

Public Sub cmdSendMail_Rejet(lK2 As String, lSubject As String, lMessage As String)
Dim wSendMail As typeSendMail
Dim bgColor As String
Dim X As String

wSendMail.FromDisplayName = lK2
wSendMail.RecipientDisplayName = "EUP_XCOM"
bgColor = "<body bgcolor = #FF80FF>"

wSendMail.Subject = lSubject
wSendMail.Attachment = ""
wSendMail.Message = bgColor & lMessage
wSendMail.AsHTML = True

srvSendMail.Monitor wSendMail

End Sub

Public Sub Form_Init()
Me.Enabled = False
Me.MousePointer = vbHourglass

Call lstErr_Clear(lstErr, cmdPrint, "Initialisation ")
DoEvents
blnControl = False

fgSelect_FormatString = fgSelect.FormatString

fgSelect.Visible = False
lstDoc_IFS.Visible = False
filDoc_XCOM.Visible = False

'_____________________________________________________________
If Not IsNull(param_Init) Then
    If Not blnAuto Then MsgBox "paramétrage inconsistant", vbCritical, "frmEUPXCOM.param_init"
    Unload Me
Else
 '   lstErr.Clear
End If

'_________________________________________________________________


cboSelect_SQL.Clear
cboSelect_SQL.AddItem "1 - ALLER  : Fichiers SEPA en attente"
cboSelect_SQL.AddItem "2 - RETOUR : Gestion SEPA"
If EUPXCOM_Aut.Xspécial Then
    param_Init_Explorer_IFS
    cboSelect_SQL.AddItem "5 - ALLER  : Move SAB\IFS => PeliSRV (V->W)"
    cboSelect_SQL.AddItem "6 - ALLER  : Contrôle du contenu (W->X|?)"
    cboSelect_SQL.AddItem "7 - ALLER  : Télétransmission (X->$|!)"
    cboSelect_SQL.AddItem "9 - RETOUR : Téléchargement / move "
    cboSelect_SQL.AddItem "! - VISTA : param_Init_Schtasks"
End If

cboSelect_SQL.ListIndex = 0


fgSelect.Enabled = True
cmdReset

Me.Enabled = True
Me.MousePointer = 0
End Sub


Private Sub cboSelect_SQL_Click()
cmdSelect_SQL_K = mId$(cboSelect_SQL, 1, 1)

fraSelect_Options.Enabled = True
cmdSelect_Ok_Click
End Sub



'---------------------------------------------------------
Public Sub cmdReset()
'---------------------------------------------------------
blnControl = False
usrColor_Set
cmdContext.Caption = constcmdRechercher: blnMsgBox_Quit = False
arrTag_Set False
'lstErr.Visible = False
currentAction = ""
fraSelect_Options.Enabled = True
'cmdSelect_Ok_Click



blnControl = True



End Sub


Public Function param_Init()
Dim xName As String, xMemo As String
param_Init = Null

Call lstErr_Clear(lstErr, cmdContext, ". EUP_XCOM_param_Init"): DoEvents

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
If blnEUP_XCOMTEST Then
    paramEnvironnement = "TEST"
    paramIBM_Library_SAB = "SAB073U"
    paramIBM_Library_SABSPE = "SAB073USPE"
    
    MsgBox "Environnement TEST SAB : " & paramIBM_Library_SAB & " - " & paramEnvironnement, vbExclamation, "SEPA"
End If
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'__________________________________________________________________
V = rsElpTable_Read("SEPA", "schtasks", "nom", xName, xMemo)
paramSchtasks_Nom = UCase$(Replace(UCase$(xMemo), "%WIN%", paramEnvironnement))
paramSchtasks_TN = "/tn " & paramSchtasks_Nom
Call lstErr_AddItem(lstErr, cmdContext, "paramSchtasks_Nom: " & paramSchtasks_Nom)
Call lstErr_AddItem(lstErr, cmdContext, "paramSchtasks_TN : " & paramSchtasks_TN)

V = rsElpTable_Read("SEPA", "schtasks", "bat", xName, xMemo)
paramSchtasks_Bat = Replace(xMemo, "%WIN%", paramEnvironnement)
Call lstErr_AddItem(lstErr, cmdContext, "paramSchtasks_TN : " & paramSchtasks_TN)

V = rsElpTable_Read("SEPA", "schtasks", "log", xName, xMemo)
paramSchtasks_Log = Replace(xMemo, "%WIN%", paramEnvironnement)
Call lstErr_AddItem(lstErr, cmdContext, "paramSchtasks_Log  : " & paramSchtasks_Log)

V = rsElpTable_Read("SEPA", "schtasks", "create", xName, xMemo)
X = Replace(xMemo, "%TN%", paramSchtasks_TN)
paramSchtasks_Create = Replace(X, "%TR%", "/tr " & paramSchtasks_Bat)
Call lstErr_AddItem(lstErr, cmdContext, "paramSchtasks_Create : " & paramSchtasks_Create)

V = rsElpTable_Read("SEPA", "schtasks", "delete", xName, xMemo)
paramSchtasks_Delete = Replace(xMemo, "%TN%", paramSchtasks_TN)
Call lstErr_AddItem(lstErr, cmdContext, "paramSchtasks_Delete  : " & paramSchtasks_Delete)

V = rsElpTable_Read("SEPA", "schtasks", "query", xName, xMemo)
paramSchtasks_Query = xMemo
Call lstErr_AddItem(lstErr, cmdContext, "paramSchtasks_Query : " & paramSchtasks_Query)

V = rsElpTable_Read("SEPA", "schtasks", "run", xName, xMemo)
paramSchtasks_Run = Replace(xMemo, "%TN%", paramSchtasks_TN)
Call lstErr_AddItem(lstErr, cmdContext, "paramSchtasks_Run  : " & paramSchtasks_Run)
'___________________________________________________________________


V = rsElpTable_Read("SEPA", "token", "", xName, xMemo)
X = Replace(xMemo, "%WIN%", paramEnvironnement)
paramToken = paramServer(X)
Call lstErr_AddItem(lstErr, cmdContext, "paramToken  : " & paramToken)


'___________________________________________________________________

V = rsElpTable_Read("SEPA", "Aller", "IFS", xName, xMemo)
X = Replace(xMemo, "%IBM%", paramIBM_AS400_ID)
paramSEPA_Aller_IFS = Replace(X, "%SAB%", paramIBM_Library_SAB)
Call lstErr_AddItem(lstErr, cmdContext, "paramSEPA_Aller_IFS  : " & paramSEPA_Aller_IFS)

V = rsElpTable_Read("SEPA", "Aller", "XCOM", xName, xMemo)
X = Replace(xMemo, "%WIN%", paramEnvironnement)
paramSEPA_Aller_XCOM = paramServer(X)
Call lstErr_AddItem(lstErr, cmdContext, "paramSEPA_Aller_XCOM  : " & paramSEPA_Aller_XCOM)

V = rsElpTable_Read("SEPA", "Aller", "Archive", xName, xMemo)
X = Replace(xMemo, "%WIN%", paramEnvironnement)
paramSEPA_Aller_Archive = paramServer(X)
Call lstErr_AddItem(lstErr, cmdContext, "paramSEPA_Aller_Archive  : " & paramSEPA_Aller_Archive)

V = rsElpTable_Read("SEPA", "Aller", "log", xName, xMemo)
X = Replace(xMemo, "%WIN%", paramEnvironnement)
paramSEPA_Aller_log = paramServer(X)
Call fileName_Split(paramSEPA_Aller_log, paramSEPA_Aller_Path_log, xName, xMemo)

Call lstErr_AddItem(lstErr, cmdContext, "paramSEPA_Aller_log : " & paramSEPA_Aller_log)


V = rsElpTable_Read("SEPA", "Aller", "Fax", xName, xMemo)
X = Replace(xMemo, "%WIN%", paramEnvironnement)
paramSEPA_Aller_Fax = paramServer(X)
Call lstErr_AddItem(lstErr, cmdContext, "paramSEPA_Aller_Fax  : " & paramSEPA_Aller_Fax)

paramSEPA_Aller_cmd = ""

'___________________________________________________________________

V = rsElpTable_Read("SEPA", "Retour", "IFS", xName, xMemo)
X = Replace(xMemo, "%IBM%", paramIBM_AS400_ID)
paramSEPA_Retour_IFS = Replace(X, "%SAB%", paramIBM_Library_SAB)
Call lstErr_AddItem(lstErr, cmdContext, "paramSEPA_Retour_IFS  : " & paramSEPA_Retour_IFS)

V = rsElpTable_Read("SEPA", "Retour", "XCOM", xName, xMemo)
X = Replace(xMemo, "%WIN%", paramEnvironnement)
paramSEPA_Retour_XCOM = paramServer(X)
Call lstErr_AddItem(lstErr, cmdContext, "paramSEPA_Retour_XCOM  : " & paramSEPA_Retour_XCOM)

V = rsElpTable_Read("SEPA", "Retour", "Archive", xName, xMemo)
X = Replace(xMemo, "%WIN%", paramEnvironnement)
paramSEPA_Retour_Archive = paramServer(X)
Call lstErr_AddItem(lstErr, cmdContext, "paramSEPA_Retour_Archive  : " & paramSEPA_Retour_Archive)

V = rsElpTable_Read("SEPA", "Retour", "log", xName, xMemo)
X = Replace(xMemo, "%WIN%", paramEnvironnement)
paramSEPA_Retour_log = paramServer(X)
Call fileName_Split(paramSEPA_Retour_log, paramSEPA_Retour_Path_log, xName, xMemo)

Call lstErr_AddItem(lstErr, cmdContext, "paramSEPA_Retour_log : " & paramSEPA_Retour_log)



paramSEPA_Retour_cmd = ""
'___________________________________________________________________


End Function





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
Me.Enabled = False: Me.MousePointer = vbHourglass
Select Case SSTab1.Tab
    Case 0:
            If fgSelect.Rows > 1 Then cmdPrint_Ok
                
    Case 1:
End Select
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub cmdSelect_SQL_1()
Dim V, fsoFile As File
Dim X As String
Dim xWhere As String, xAnd As String
Dim wAmj7 As Long
On Error GoTo Error_Handler


currentAction = "cmdSelect_SQL_1"
Call lstErr_Clear(lstErr, cmdContext, currentAction): DoEvents
fraRetour.Visible = False
fgSelect.Width = 12840

xWhere = "where EUPMON2STA not in ('$','A')"

Call arrYEUPMON2_SQL(xWhere)

fgSelect_Display_1

Call lstErr_AddItem(lstErr, cmdContext, "Répertoire : " & filDoc_XCOM.Path): DoEvents

lblflDoc_XCOM = paramSEPA_Aller_XCOM
filDoc_XCOM.Path = paramSEPA_Aller_XCOM
filDoc_XCOM.Pattern = "x.xxx"
filDoc_XCOM.Pattern = "*.*"
filDoc_XCOM.Visible = True
lblflDoc_XCOM.Visible = True


lbllstDoc_IFS.Caption = paramSEPA_Aller_IFS
lbllstDoc_IFS.Visible = True

If EUPXCOM_Aut.Xspécial Then
    Call lstErr_AddItem(lstErr, cmdContext, "Répertoire : " & paramSEPA_Aller_IFS): DoEvents
    lstDoc_IFS.Clear
    lstDoc_IFS.Visible = True
    param_Init_DIR paramSEPA_Aller_IFS
Else
    lstDoc_IFS.Visible = False
End If


Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub
Private Sub cmdSelect_SQL_2()
Dim V, fsoFile As File
Dim X As String
Dim xWhere As String, xAnd As String
Dim wAmj7 As Long
On Error GoTo Error_Handler

currentAction = "cmdSelect_SQL_2"
Call lstErr_Clear(lstErr, cmdContext, currentAction): DoEvents
fraRetour.Visible = True
fgSelect.Width = 7920

xWhere = "where EUPMON2STA not in ('V','R')"

Call arrYEUPMON4_SQL(xWhere)

fgSelect_Display_2

'______________________________________________________________
fraRetour_R0V_Display
fraRetour_R2V_Display
'_______________________________________________________________


Call lstErr_AddItem(lstErr, cmdContext, "Répertoire : " & filDoc_XCOM.Path): DoEvents

lblflDoc_XCOM = paramSEPA_Retour_XCOM
filDoc_XCOM.Path = paramSEPA_Retour_XCOM
filDoc_XCOM.Pattern = "x.xxx"
filDoc_XCOM.Pattern = "*.*"
filDoc_XCOM.Visible = True
lblflDoc_XCOM.Visible = True


lbllstDoc_IFS.Caption = paramSEPA_Retour_IFS
lbllstDoc_IFS.Visible = True

If EUPXCOM_Aut.Xspécial Then
    Call lstErr_AddItem(lstErr, cmdContext, "Répertoire : " & paramSEPA_Retour_IFS): DoEvents
    lstDoc_IFS.Clear
    lstDoc_IFS.Visible = True
    param_Init_DIR paramSEPA_Retour_IFS
Else
    lstDoc_IFS.Visible = False
End If


Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    If Not blnAuto Then MsgBox V, vbCritical, Me.Name & " : " & currentAction


End Sub

Private Sub cmdSelect_SQL_5()
Dim V
Dim X As String, Nb As Long
Dim xSQL As String
On Error GoTo Error_Handler

App_Debug = "cmdSelect_SQL_5"

currentAction = "cmdSelect_SQL_5"
Call lstErr_Clear(lstErr, cmdContext, currentAction): DoEvents

xSQL = "select count(*) as Tally  from " & paramIBM_Library_SABSPE & ".YEUPMON2 where EUPMON2STA = 'V'"
Set rsSab = cnsab.Execute(xSQL)
Nb = rsSab("Tally")

'déplacement des fichiers \\..\IFS\ ...xml vers le serveur \\....\SEPA_Aller

If Nb > 0 Then
    Close #1
    Open paramSchtasks_Bat For Output As #1
    Print #1, "MOVE /Y " & paramSEPA_Aller_IFS & "*.xml " & paramSEPA_Aller_XCOM & " >> " & paramSEPA_Aller_log; ""
    Close #1
    
    Call Shell_Exe(paramSchtasks_Run)
    Call lstErr_AddItem(lstErr, cmdContext, "@EUP_XCOM : 5 temporisation"): DoEvents
    Sleep (10000)
    
    lblflDoc_XCOM = paramSEPA_Aller_XCOM
    filDoc_XCOM.Path = paramSEPA_Aller_XCOM
    filDoc_XCOM.Pattern = "x.xxx"
    filDoc_XCOM.Pattern = "*.xml"
    For I = 0 To filDoc_XCOM.ListCount - 1
        filDoc_XCOM.ListIndex = I
        xSQL = "select *  from " & paramIBM_Library_SABSPE & ".YEUPMON2 where EUPMON2FIC = '" & filDoc_XCOM.FileName & "'"
        Set rsSab = cnsab.Execute(xSQL)
        If Not rsSab.EOF Then
            V = rsYEUPMON2_GetBuffer(rsSab, oldYEUPMON2)
            If oldYEUPMON2.EUPMON2STA = "V" Then
                newYEUPMON2 = oldYEUPMON2
                newYEUPMON2.EUPMON2STA = "W"
                cmdSelect_SQL_Transaction
            End If
        End If
    Next I

End If


Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    Error_Route V

End Sub

Public Function cmdSelect_SQL_Transaction()
Dim V, X As String, xSQL As String
Dim Nb As Long
Dim mMsgBox As String
On Error GoTo Error_Handler
'-------------------------------------------------------
App_Debug = "cmdSelect_SQL_Transaction"
'-------------------------------------------------------
cmdSelect_SQL_Transaction = Null

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________
V = sqlYEUPMON2_Update(newYEUPMON2, oldYEUPMON2)
If Not IsNull(V) Then GoTo Error_MsgBox

GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    Error_Route V
Exit_sub:
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
    Else
        V = cnSAB_Transaction("Commit")
    End If
    
    cmdSelect_SQL_Transaction = V
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


End Function

Public Sub cmdSelect_SQL_6()
Dim V, I As Integer
Dim X As String, Nb As Long
Dim xSQL As String
On Error GoTo Error_Handler

App_Debug = "cmdSelect_SQL_6"

currentAction = "cmdSelect_SQL_6"

Call lstErr_Clear(lstErr, cmdContext, currentAction): DoEvents
Close

xSQL = " where EUPMON2STA = 'W'"

arrYEUPMON2_SQL xSQL

For I = 1 To arrYEUPMON2_Nb
         
    oldYEUPMON2 = arrYEUPMON2(I)
    Call logAller_Open("6_NEW")
    If Dir(paramToken) = "" Then                          '? transfert en cours ?
         newYEUPMON2 = oldYEUPMON2
         V = cmdSelect_SQL_6_Control(oldYEUPMON2.EUPMON2FIC)
         If IsNull(V) Then
         
            V = cmdSelect_SQL_6_Emission(oldYEUPMON2.EUPMON2FIC) ' appel peltrans
            If IsNull(V) Then
                newYEUPMON2.EUPMON2STA = "X"
                 cmdSelect_SQL_Transaction
                 Print #501, Now & " | " & "Statut : 'X' " & paramIBM_Library_SABSPE & ".YEUPMON2 | " & oldYEUPMON2.EUPMON2FIC
            End If
            
         Else
             newYEUPMON2.EUPMON2STA = "?"
             cmdSelect_SQL_Transaction
             Print #501, Now & " | " & "Statut : '?' " & paramIBM_Library_SABSPE & ".YEUPMON2 | " & oldYEUPMON2.EUPMON2FIC
             cmdSendMail_SEPA_Alerte "Télétransmission : Anomalie du contenu du fichier : " & oldYEUPMON2.EUPMON2FIC, currentError, paramSEPA_Aller_XCOM & oldYEUPMON2.EUPMON2FIC

        End If
    Else
        Print #501, Now & " | " & "en attente autre transfert en cours > " & paramToken
        Call lstErr_AddItem(lstErr, cmdContext, paramToken): DoEvents

    End If
    
    Close #501

Next I


Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    Error_Route V

End Sub


Public Sub cmdSelect_SQL_9()
Dim V, I As Integer
Dim X As String, Nb As Long
Dim xSQL As String
On Error GoTo Error_Handler

App_Debug = "cmdSelect_SQL_9"

currentAction = "cmdSelect_SQL_9"

Call lstErr_Clear(lstErr, cmdContext, currentAction): DoEvents
Close

xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIAMON7  where MONAPP = 'SEPA' and MONSTATUS <> ''"
Set rsSab = cnsab.Execute(xSQL)

Do While Not rsSab.EOF
    V = rsYBIAMON0_GetBuffer(rsSab, oldYBIAMON0)
    If IsNull(V) Then
        newYBIAMON0 = oldYBIAMON0
        Select Case Trim(oldYBIAMON0.MONSTATUS)
            Case "Demande": cmdSelect_SQL_9_Demande
            Case "X_enCours": 'cmdSelect_SQL_9_X_enCours
            Case "W_Reçu": 'cmdSelect_SQL_9_W_Reçu
        End Select
    End If
    
    rsSab.MoveNext

Loop


Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    Error_Route V

End Sub

Public Sub cmdSelect_SQL_7()
Dim V
Dim X As String, Nb As Long
Dim xSQL As String
Dim dateEUPMON2 As Date
On Error GoTo Error_Handler

App_Debug = "cmdSelect_SQL_7"

currentAction = "cmdSelect_SQL_7"

Call lstErr_Clear(lstErr, cmdContext, currentAction): DoEvents
Close

xSQL = "select *  from " & paramIBM_Library_SABSPE & ".YEUPMON2 where EUPMON2STA = 'X'"
Set rsSab = cnsab.Execute(xSQL)

'Contrôle des fichiers  \\....\SEPA_Aller\...xml

Do While Not rsSab.EOF
    V = rsYEUPMON2_GetBuffer(rsSab, oldYEUPMON2)

     If Not IsNull(V) Then
         If Not blnAuto Then MsgBox V, vbCritical, "frmEUP_XCOM.cmdSelect_SQL_7"
        '' Exit Sub
     Else
        logAller_Open 7

         newYEUPMON2 = oldYEUPMON2
         dateEUPMON2 = Date_VB(oldYEUPMON2.EUPMON2DUP, oldYEUPMON2.EUPMON2HUP)
         If IsNull(cmdSelect_SQL_7_Emis(oldYEUPMON2.EUPMON2FIC, dateEUPMON2)) Then
             newYEUPMON2.EUPMON2STA = "$"
             cmdSelect_SQL_Transaction
             Print #501, Now & " | " & "Statut : '$' " & paramIBM_Library_SABSPE & ".YEUPMON2 | " & oldYEUPMON2.EUPMON2FIC
             Print #501, Now & " | " & "**************************************************************"
             Close #501
             cmdSendMail_SEPA_SG "Télétransmission terminée pour " & oldYEUPMON2.EUPMON2FIC, paramSEPA_Aller_Détail_log

          End If

        
        
    End If
    rsSab.MoveNext

Loop


Exit Sub

Error_Handler:
    V = Error
Error_MsgBox:
    SSTab1.Tab = 0
    Error_Route V


End Sub

Private Sub cmdRetour_R0V_Click()

Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> EUP_XCOM_cmdRetour_R0V ........"): DoEvents

oldYBIAMON0 = oldYBIAMON0_R0V
newYBIAMON0 = oldYBIAMON0_R0V
If Trim(newYBIAMON0.MONSTATUS) <> "" Then
    newYBIAMON0.MONSTATUS = ""
Else
    newYBIAMON0.MONSTATUS = "Demande"
End If
newYBIAMON0.MONFILE = ""
newYBIAMON0.MONJOB = Date
newYBIAMON0.MONPGM = Time

sqlYBIAMON0_Transaction

fraRetour_R0V_Display

Call lstErr_AddItem(lstErr, cmdContext, "< EUP_XCOM_cmdRetour_R0V"): DoEvents
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Function sqlYBIAMON0_Transaction()
On Error GoTo Error_Handler

sqlYBIAMON0_Transaction = Null
Call lstErr_Clear(lstErr, cmdContext, "> EUP_XCOM_sqlYBIAMON0_Transaction ........"): DoEvents

'$ TRANSACTION >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
V = cnSAB_Transaction("BeginTrans")
If Not IsNull(V) Then GoTo Error_MsgBox
'________________________________________________________________________________
V = sqlYBIAMON0_Update(newYBIAMON0, oldYBIAMON0, True)
If Not IsNull(V) Then GoTo Error_MsgBox



GoTo Exit_sub

'------------------------------------------
Error_Handler:
    V = Error
Error_MsgBox:
    Error_Route V
Exit_sub:
    If Not IsNull(V) Then
        V = cnSAB_Transaction("Rollback")
    Else
        V = cnSAB_Transaction("Commit")
    End If
    
'$ TRANSACTION <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
fraRetour_R0V_Display
Call lstErr_AddItem(lstErr, cmdContext, "< EUP_XCOM_sqlYBIAMON0_Transaction"): DoEvents
End Function

Private Sub cmdRetour_R2V_Click()

Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> EUP_XCOM_cmdRetour_R2V ........"): DoEvents

oldYBIAMON0 = oldYBIAMON0_R2V
newYBIAMON0 = oldYBIAMON0_R2V

If Trim(newYBIAMON0.MONSTATUS) <> "" Then
    newYBIAMON0.MONSTATUS = ""
Else
    newYBIAMON0.MONSTATUS = "Demande"
End If
newYBIAMON0.MONFILE = ""
newYBIAMON0.MONJOB = Date
newYBIAMON0.MONPGM = Time

sqlYBIAMON0_Transaction

fraRetour_R2V_Display

Call lstErr_AddItem(lstErr, cmdContext, "< EUP_XCOM_cmdRetour_R2V"): DoEvents
Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub cmdSelect_Ok_Click()
Dim blnOk As Boolean, Nb As Long

blnOk = fraSelect_Options.Enabled
Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> EUP_XCOM_cmdSelect_Ok ........"): DoEvents

fgSelect.Clear
filDoc_XCOM.Visible = False
lblflDoc_XCOM.Visible = False
lstDoc_IFS.Visible = False
lbllstDoc_IFS.Visible = False
If blnOk Then
    cmdSelect_Ok.Caption = "Options"
    cmdSelect_Ok.BackColor = &HFFFFFA   '&HC0FFFF
    fraSelect_Options.BackColor = &H8000000F
    Call usrColor_Container(fraSelect_Options, fraSelect_Options.BackColor)
    fraSelect_Options.Enabled = False
    Select Case cmdSelect_SQL_K
        Case "1":    cmdSelect_SQL_1
        Case "2":    cmdSelect_SQL_2
        Case "5":    cmdSelect_SQL_5
        Case "6":    cmdSelect_SQL_6
        Case "7":    cmdSelect_SQL_7
        Case "9":    cmdSelect_SQL_9
        Case "!":    param_Init_Schtasks

    End Select
    If Not blnAuto Then
        Select Case cmdSelect_SQL_K
            Case "5", "6", "7": cboSelect_SQL.ListIndex = 0: cmdSelect_SQL_1
            Case "9": cboSelect_SQL.ListIndex = 1: cmdSelect_SQL_2
        End Select
    End If
Else
    cmdSelect_Ok.Caption = constcmdRechercher
    cmdSelect_Ok.BackColor = &HC0FFC0
    fraSelect_Options.BackColor = &HFFFFFA    '&HC0FFFF
    Call usrColor_Container(fraSelect_Options, fraSelect_Options.BackColor)
    fraSelect_Options.Enabled = True
End If
Call lstErr_AddItem(lstErr, cmdContext, "< EUP_XCOM_cmdSelect_Ok"): DoEvents
Me.Enabled = True: Me.MousePointer = 0

Me.Enabled = False: Me.MousePointer = vbHourglass
Call lstErr_Clear(lstErr, cmdContext, "> EUP_XCOM_cmdSelect_Ok ........"): DoEvents
Call lstErr_AddItem(lstErr, cmdContext, "< EUP_XCOM_cmdSelect_Ok"): DoEvents
Me.Enabled = True: Me.MousePointer = 0

End Sub


Private Sub fgSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim K As Long
On Error Resume Next
If y <= fgSelect.RowHeightMin Then
    Select Case fgSelect.Col
        Case 0: fgSelect_Sort1 = 0: fgSelect_Sort2 = 2: fgSelect_Sort
        Case 1:  fgSelect_Sort1 = 1: fgSelect_Sort2 = 2: fgSelect_Sort
        Case 2: fgSelect_Sort1 = 2: fgSelect_Sort2 = 2: fgSelect_Sort
        Case 3: fgSelect_Sort1 = 3: fgSelect_Sort2 = 3: fgSelect_Sort
        Case 6: fgSelect_Sort1 = 6: fgSelect_Sort2 = 6: fgSelect_Sort
        Case 7: fgSelect_Sort1 = 7: fgSelect_Sort2 = 7: fgSelect_Sort
        Case 8: fgSelect_Sort1 = 8: fgSelect_Sort2 = 8: fgSelect_Sort
       Case fgSelect_arrIndex:  fgSelect_SortX fgSelect_arrIndex
    End Select
Else
    If fgSelect.Rows > 1 Then
        Call fgSelect_Color(fgSelect_RowClick, MouseMoveUsr.BackColor, fgSelect_ColorClick)
        fgSelect.Col = fgSelect_arrIndex:  arrYEUPMON2_Index = CLng(fgSelect.Text)
        fgSelect.LeftCol = 0
        xYEUPMON2 = arrYEUPMON2(arrYEUPMON2_Index)
        fgDétail_Display

   End If
End If
fgSelect.LeftCol = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
If blnEUP_XCOMTEST Then End

End Sub

Private Sub mnuContextAbandonner_Click()
cmdContext_Quit
End Sub


Private Sub mnuContextQuitter_Click()
Unload Me
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

Public Sub cmdContext_Quit()
blnControl = False
lstErr.Clear: lstErr.Height = 200
If SSTab1.Tab = 0 Then
        Unload Me
    Exit Sub
Else
    SSTab1.Tab = SSTab1.Tab - 1
End If

End Sub

Public Sub cmdContext_Return()
On Error Resume Next
If SSTab1.Tab = 0 Then
    fgSelect.Row = fgSelect.TopRow
    fgSelect.Col = fgSelect_arrIndex:
Else
    SendKeys "{TAB}"
End If
End Sub

'---------------------------------------------------------
Private Sub Form_Load()
'---------------------------------------------------------
Dim V
Dim xName  As String, xMemo As String
On Error GoTo Error_Handler

mHeight_0 = Me.Height: mWidth_0 = Me.Width: mHeight_2 = 0: mWidth_2 = 0: mWindowState = Me.WindowState
Set XForm = Me
Call MeInit(arrTagNb)
ReDim arrTag(arrTagNb + 1)
blnControl = False


Exit Sub

Error_Handler:

blnControl = False
If Not blnAuto Then MsgBox Error
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
'Call txt_GotFocus(txt)
'Call txt_LostFocus(txt)

End Sub


Private Sub mnuSelect_Print_Détail_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
cmdPrint_Ok '"D "
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub mnuSelect_Print_Liste_Click()
Me.Enabled = False: Me.MousePointer = vbHourglass
cmdPrint_Ok '"L "
Me.Enabled = True: Me.MousePointer = 0

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error Resume Next
If SSTab1.Tab = 0 Then cmdSelect_Ok.SetFocus

End Sub














Private Sub SSTab1_GotFocus()
Select Case SSTab1.Tab
    Case 0: fgSelect.LeftCol = 0
End Select
End Sub


Public Sub cmdPrint_Ok()
Dim iRow As Integer, K As Integer, I As Integer
Dim blnOk As Boolean

fgSelect.Visible = False
Call lstErr_Clear(Me.lstErr, Me.cmdContext, "Impression Etat : " & fgSelect.Rows - 1)

fgSelect.Visible = True
Me.Show
End Sub


Public Sub cmdPrint_SEPA_Aller_Fax(lFile As String, lNb As Integer, lMt As Currency)
Dim xIn As String
Dim xA0V_Nb As String, xA0V_Mt As String
Dim xA2V_Nb As String, xA2V_Mt As String
Dim xA2P_Nb As String, xA2P_Mt As String
Dim xA1P_Nb As String, xA1P_Mt As String
Dim xA3P_Nb As String, xA3P_Mt As String
Dim xA3V_Nb As String, xA3V_Mt As String
Dim xA4V_Nb As String, xA4V_Mt As String
Dim xA0C_Nb As String, xA0C_Mt As String

Dim indiceCaseACocher As Long
Dim indiceEnCours As Long
Dim indPosition As Long
Dim indice As Long

    Call lstErr_AddItem(Me.lstErr, Me.cmdContext, "Impression SEPA_Aller_Fax : ")
    xA0V_Nb = " "
    xA0V_Mt = " "
    xA2V_Nb = " "
    xA2V_Mt = " "
    xA2P_Nb = " "
    xA2P_Mt = " "
    xA1P_Nb = " "
    xA1P_Mt = " "
    xA3P_Nb = " "
    xA3P_Mt = " "
    xA3V_Nb = " "
    xA3V_Mt = " "
    xA4V_Nb = " "
    xA4V_Mt = " "
    xA0C_Nb = " "
    xA0C_Mt = " "
    If mId$(lFile, 1, 3) = "A0V" Then
        xA0V_Nb = Format$(lNb, "### ##0")
        xA0V_Mt = Format$(lMt, "### ### ##0.00") & " €"
        indiceCaseACocher = 1
    ElseIf mId$(lFile, 1, 3) = "A2V" Then
        xA2V_Nb = Format$(lNb, "### ##0")
        xA2V_Mt = Format$(lMt, "### ### ##0.00") & " €"
        indiceCaseACocher = 2
    ElseIf mId$(lFile, 1, 3) = "A2P" Then
        xA2P_Nb = Format$(lNb, "### ##0")
        xA2P_Mt = Format$(lMt, "### ### ##0.00") & " €"
        indiceCaseACocher = 3
    ElseIf mId$(lFile, 1, 3) = "A1P" Then
        xA1P_Nb = Format$(lNb, "### ##0")
        xA1P_Mt = Format$(lMt, "### ### ##0.00") & " €"
        indiceCaseACocher = 4
    ElseIf mId$(lFile, 1, 3) = "A3P" Then
        xA3P_Nb = Format$(lNb, "### ##0")
        xA3P_Mt = Format$(lMt, "### ### ##0.00") & " €"
        indiceCaseACocher = 5
    ElseIf mId$(lFile, 1, 3) = "A3V" Then
        xA3V_Nb = Format$(lNb, "### ###")
        xA3V_Mt = ""
        indiceCaseACocher = 6
    ElseIf mId$(lFile, 1, 3) = "A4V" Then
        xA4V_Nb = Format$(lNb, "### ##0")
        xA4V_Mt = Format$(lMt, "### ### ##0.00") & " €"
        indiceCaseACocher = 7
    ElseIf mId$(lFile, 1, 3) = "A0C" Then
        xA0C_Nb = Format$(lNb, "### ###")
        xA0C_Mt = ""
        indiceCaseACocher = 8
    End If

paramSEPA_Aller_Fax = "c:\temp\test\SEPA_Aller_Fax.rtf"
paramSEPA_Aller_Path_log = "c:\temp\test\"


Open paramSEPA_Aller_Fax For Input As #1
Open paramSEPA_Aller_Path_log & Replace(lFile, ".xml", "_Fax.doc") For Output As #2
indiceEnCours = 0
Do Until EOF(1)
    Line Input #1, xIn
    'Gestion de LA case à cocher 30/06/2011 DR
    For indice = 1 To 8
        If InStr(xIn, "CaseACocher" & indice) > 0 Then
            indiceEnCours = indice
        End If
    Next indice
    indPosition = InStr(xIn, "ffdefres")
    If indPosition > 0 Then
        If indiceEnCours = indiceCaseACocher Then
            Mid(xIn, indPosition, 9) = "ffdefres1"
        Else
            Mid(xIn, indPosition, 9) = "ffdefres0"
        End If
    End If
    'FIN Gestion de la case à cocher
    If InStr(xIn, "@DATE") > 0 Then xIn = Replace(xIn, "@DATE", dateImp(DSys))
    If InStr(xIn, "@A0V_NB") > 0 Then xIn = Replace(xIn, "@A0V_NB", xA0V_Nb)
    If InStr(xIn, "@A2V_NB") > 0 Then xIn = Replace(xIn, "@A2V_NB", xA2V_Nb)
    If InStr(xIn, "@A2P_NB") > 0 Then xIn = Replace(xIn, "@A2P_NB", xA2P_Nb)
    If InStr(xIn, "@A1P_NB") > 0 Then xIn = Replace(xIn, "@A1P_NB", xA1P_Nb)
    If InStr(xIn, "@A3P_NB") > 0 Then xIn = Replace(xIn, "@A3P_NB", xA3P_Nb)
    If InStr(xIn, "@A3V_NB") > 0 Then xIn = Replace(xIn, "@A3V_NB", xA3V_Nb)
    If InStr(xIn, "@A4V_NB") > 0 Then xIn = Replace(xIn, "@A4V_NB", xA4V_Nb)
    If InStr(xIn, "@A0C_NB") > 0 Then xIn = Replace(xIn, "@A0C_NB", xA0C_Nb)
    
    If InStr(xIn, "@A0V_MT") > 0 Then xIn = Replace(xIn, "@A0V_MT", xA0V_Mt)
    If InStr(xIn, "@A2V_MT") > 0 Then xIn = Replace(xIn, "@A2V_MT", xA2V_Mt)
    If InStr(xIn, "@A2P_MT") > 0 Then xIn = Replace(xIn, "@A2P_MT", xA2P_Mt)
    If InStr(xIn, "@A1P_MT") > 0 Then xIn = Replace(xIn, "@A1P_MT", xA1P_Mt)
    If InStr(xIn, "@A3P_MT") > 0 Then xIn = Replace(xIn, "@A3P_MT", xA3P_Mt)
    If InStr(xIn, "@A3V_MT") > 0 Then xIn = Replace(xIn, "@A3V_MT", xA3V_Mt)
    If InStr(xIn, "@A4V_MT") > 0 Then xIn = Replace(xIn, "@A4V_MT", xA4V_Mt)
    If InStr(xIn, "@A0C_MT") > 0 Then xIn = Replace(xIn, "@A0C_MT", xA0C_Mt)
    
    Print #2, xIn
Loop
Close #1
Close #2


End Sub







Public Sub Error_Route(V)

currentError = CStr(V) & "             ( " & Me.Name & " ~ " & App_Debug & " )"
If blnAuto Then
  '  Call cmdSendMail_Alerte(Me.Name & " ~ " & App_Debug, CStr(V))
Else
    MsgBox V, vbCritical, Me.Name & " ~ " & App_Debug
End If

End Sub

Public Function param_Init_Schtasks()
Dim xName As String, xMemo As String
Dim iWait As Integer, blnLog As Boolean, blnOk As Boolean
Dim X As String, xIn As String

On Error GoTo Error_Handler

param_Init_Schtasks = Null

Call lstErr_AddItem(lstErr, cmdContext, "param_Init_Schtasks : 1 paramètrage"): DoEvents


'_____________________________________________________________
' ? Existence de la tâche planifiée
'____________________________________________________________
Call lstErr_AddItem(lstErr, cmdContext, "param_Init_Schtasks : 2 planification"): DoEvents

Close #1
blnLog = False: blnOk = False
If Dir(paramSchtasks_Log) <> "" Then Kill paramSchtasks_Log
Open paramSchtasks_Bat For Output As #1
Print #1, paramSchtasks_Query & " > " & paramSchtasks_Log
Close #1

Call Shell_Exe(paramSchtasks_Bat)
Sleep (500)

For iWait = 1 To 10
    DoEvents
    Call lstErr_AddItem(lstErr, cmdContext, "param_Init_Schtasks : 2 planification - log :" & iWait): DoEvents
    If Dir(paramSchtasks_Log) <> "" Then
        Open paramSchtasks_Log For Input As #1
        blnLog = True
        Do Until EOF(1)
            Line Input #1, xIn
            xIn = UCase$(xIn)
           If InStr(xIn, paramSchtasks_Nom) > 0 Then blnOk = True: Exit Do
        Loop
        Close #1
        
        Exit For
    End If
    Sleep 1000
Next iWait
If Not blnLog Then
    Call MsgBox(paramSchtasks_Log & " :  non trouvé", vbCritical, "frmEUP_XCOM:param_Init_Schtasks : 2 planification")
    Unload Me
End If
If Not blnOk Then
    Call MsgBox(paramSchtasks_Nom & " :  création", vbInformation, "frmEUP_XCOM:param_Init_Schtasks : 2 planification")
    Call lstErr_AddItem(lstErr, cmdContext, "param_Init_Schtasks : 2 planification"): DoEvents
    Open paramSchtasks_Bat For Output As #1
    Print #1, paramSchtasks_Create & " > " & paramSchtasks_Log
    Close #1
    
    Call Shell_Exe(paramSchtasks_Bat)
    Sleep (500)
    Open paramSchtasks_Log For Input As #1
    Line Input #1, xIn
        Call MsgBox(Trim(xIn), vbInformation, "frmEUP_XCOM:param_Init_Schtasks : 2 planification")
    Close #1


End If
Exit Function

'------------------------------------------
Error_Handler:
    param_Init_Schtasks = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & " : " & App_Debug

End Function

Public Function param_Init_DIR(lIFS As String)

Dim xName As String, xMemo As String
Dim iWait As Integer, blnLog As Boolean, blnOk As Boolean, blnVide As Boolean
Dim X As String, xIn As String

On Error GoTo Error_Handler
param_Init_DIR = Null
Close #1
blnLog = False: blnOk = False

If Dir(paramSchtasks_Log) <> "" Then
    Call lstErr_AddItem(lstErr, cmdContext, "Kill : " & paramSchtasks_Log): DoEvents
    Kill paramSchtasks_Log
End If
Open paramSchtasks_Bat For Output As #1
Print #1, "DIR " & lIFS & " > " & paramSchtasks_Log; ""
Close #1

Call lstErr_AddItem(lstErr, cmdContext, "Shell_Exe : " & paramSchtasks_Run): DoEvents
Call Shell_Exe(paramSchtasks_Run)
Me.Show
Sleep (500)

For iWait = 1 To 10
    DoEvents
    Call lstErr_AddItem(lstErr, cmdContext, "param_Init_DIR :  - log :" & iWait): DoEvents
    If Dir(paramSchtasks_Log) <> "" Then
        Call lstErr_AddItem(lstErr, cmdContext, "Open : " & paramSchtasks_Log): DoEvents
        Open paramSchtasks_Log For Input As #1
        blnLog = True: blnVide = True
        Do Until EOF(1)
            Line Input #1, xIn
            If Not blnOk Then
                If mId$(xIn, 22, 17) = "<REP>          .." Then blnOk = True: blnVide = False
            Else
                If mId$(xIn, 18, 10) = "fichier(s)" Then
                    blnOk = False
                Else
                    lstDoc_IFS.AddItem mId$(xIn, 37, Len(xIn) - 36)
                End If
                
            End If
        Loop
        Close #1
        
        Exit For
    End If
    Sleep 500
Next iWait
If Not blnLog Then
    Call MsgBox(paramSchtasks_Log & " :  non trouvé", vbCritical, "frmEUP_XCOM:param_Init_DIR")
    Unload Me
End If
If blnVide Then
    Call MsgBox("Pas d'accès à " & lIFS, vbCritical, "frmEUP_XCOM:param_Init_DIR")
End If
Exit Function

'------------------------------------------
Error_Handler:
    param_Init_DIR = Error
Error_MsgBox:
    MsgBox V, vbCritical, frmElp_Caption & " : " & App_Debug


End Function

Public Function cmdSelect_SQL_6_Control(lFile As String)
Dim X As String, xIn As String, wFile As String, wMsgId As String
Dim K1 As Integer, K2 As Integer, xCur As Currency
Dim blnGrpHdr As Boolean, blnMsgId As Boolean, blnNbOfTxs As Boolean, blnAmt As Boolean
Dim mNbOfTxs As Integer, mCcy As String, mAmt As Currency, wTxId As String
Dim tNbOfTxs As Integer, tCcy As String, tAmt As Currency

Dim blnCdtrAcct As Boolean, blnUstrd As Boolean
Dim wCdtrAcct  As String, wUstrd  As String

Dim blnCdtTrfTxInf As Boolean

On Error GoTo Error_Handler
cmdSelect_SQL_6_Control = "?"

Call lstErr_AddItem(lstErr, cmdContext, lFile): DoEvents

'DR 18/12/2012 =================================
If mId$(lFile, 1, 3) = "A0C" Then
    Call cmdSelect_SQL_6_Control_A0C(lFile)
    cmdSelect_SQL_6_Control = Null
    Exit Function
End If
'FIN DR 18/12/2012 =============================

'DR 05/02/2013 =================================
If mId$(lFile, 1, 3) = "A3V" Then
    Call cmdSelect_SQL_6_Control_A3V(lFile)
    cmdSelect_SQL_6_Control = Null
    Exit Function
End If
'FIN DR 05/02/2013 =============================

Close #501
Open "c:\temp\test\A0V-000001290-1130424.txt" For Output As #501

wMsgId = Replace(lFile, ".xml", "")
paramSEPA_Aller_XCOM = "C:\TEMP\test\"
wFile = paramSEPA_Aller_XCOM & lFile
blnGrpHdr = False: blnMsgId = False: blnNbOfTxs = False: blnAmt = False
blnCdtTrfTxInf = False
mNbOfTxs = 0: mCcy = "": mAmt = 0
tNbOfTxs = 0: tCcy = "": tAmt = 0: wTxId = ""
blnCdtrAcct = False: blnUstrd = False
wCdtrAcct = "": wUstrd = ""
If Dir(wFile) <> "" Then
    'If Not blnAuto Then MsgBox "à faire : contrôle NB, Mt, IBAN / ZEUPG2A0+ZEUPG3A0", vbInformation
    Open wFile For Input As #1
    Do Until EOF(1)
        Line Input #1, xIn
        If Not blnGrpHdr Then
            If InStr(xIn, "<GrpHdr>") > 0 Then blnGrpHdr = True
        Else
        
            If InStr(xIn, "<MsgId>") > 0 Then
                If InStr(xIn, wMsgId) > 0 Then blnMsgId = True
            End If
            If blnMsgId Then
                If InStr(xIn, "<NbOfTxs>") > 0 Then
                    K1 = InStr(xIn, ">")
                    K2 = InStr(K1 + 1, xIn, "<")
                    mNbOfTxs = CInt(mId$(xIn, K1 + 1, K2 - K1 - 1))
                    Call lstErr_AddItem(lstErr, cmdContext, "mNbOfTxs :" & mNbOfTxs): DoEvents
                    Print #501, Now & " | " & "<NbOfTxs> :" & mNbOfTxs

                End If
                If InStr(xIn, "<TtlIntrBkSttlmAmt Ccy =") > 0 Or InStr(xIn, "<TtlRtrdIntrBkSttlmAmt Ccy =") > 0 Then
                    K1 = InStr(xIn, Asc34)
                    K2 = InStr(K1 + 1, xIn, Asc34)
                    mCcy = mId$(xIn, K1 + 1, K2 - K1 - 1)
                    K1 = InStr(K2 + 1, xIn, ">")
                    K2 = InStr(K1 + 1, xIn, "<")
                    mAmt = Val(mId$(xIn, K1 + 1, K2 - K1 - 1))
                    Call lstErr_AddItem(lstErr, cmdContext, mCcy & " :" & mAmt): DoEvents
                    Print #501, Now & " | " & "<TtlIntrBkSttlmAmt Ccy = " & mCcy & " " & Format(mAmt, "### ### ###.00")
              End If
            '______________________________________________________________________________________
                If Not blnCdtTrfTxInf Then
                    If InStr(xIn, "<CdtTrfTxInf>") > 0 Or InStr(xIn, "<TxInf>") > 0 Then
                        blnCdtTrfTxInf = True
                        blnCdtrAcct = False: blnUstrd = False
                        tCcy = "": xCur = 0: wCdtrAcct = "": wUstrd = ""
                    End If

                Else
                    If InStr(xIn, "<TxId>") > 0 Or InStr(xIn, "<RtrId>") > 0 Then
                        K1 = InStr(xIn, ">")
                        K2 = InStr(K1 + 1, xIn, "<")
                        wTxId = mId$(xIn, K1 + 1, K2 - K1 - 1)
                    End If
                    If InStr(xIn, "<IntrBkSttlmAmt Ccy =") > 0 Or InStr(xIn, "<RtrdIntrBkSttlmAmt Ccy =") > 0 Then
                        K1 = InStr(xIn, Asc34)
                        K2 = InStr(K1 + 1, xIn, Asc34)
                        tCcy = mId$(xIn, K1 + 1, K2 - K1 - 1)
                        K1 = InStr(K2 + 1, xIn, ">")
                        K2 = InStr(K1 + 1, xIn, "<")
                        xCur = Val(mId$(xIn, K1 + 1, K2 - K1 - 1))
                        
                    End If
                    If Not blnCdtrAcct Then
                        If InStr(xIn, "<CdtrAcct>") > 0 Or InStr(xIn, "<DbtrAcct>") > 0 Then blnCdtrAcct = True
                    Else
                        If InStr(xIn, "<IBAN>") > 0 Then
                            K1 = InStr(xIn, ">")
                            K2 = InStr(K1 + 1, xIn, "<")
                            wCdtrAcct = mId$(xIn, K1 + 1, K2 - K1 - 1)
                        End If
                        If InStr(xIn, "<Ustrd>") > 0 Then
                            blnCdtrAcct = False
                            K1 = InStr(xIn, ">")
                            K2 = InStr(K1 + 1, xIn, "<")
                            wUstrd = mId$(xIn, K1 + 1, K2 - K1 - 1)
                        End If

                    End If
                   

                    If InStr(xIn, "</CdtTrfTxInf>") > 0 Or InStr(xIn, "</TxInf>") > 0 Then
                        blnCdtTrfTxInf = False
                        Call lstErr_AddItem(lstErr, cmdContext, wTxId & " : " & tCcy & " :" & xCur): DoEvents
                        Print #501, Now & " | " & wTxId & " : " & tCcy & " " & Format(xCur, "### ### ###.00") & " : " & wCdtrAcct & " : " & wUstrd
                        If tCcy <> mCcy Then V = "Devise détail  : " & tCcy & " <> " & mCcy: GoTo Error_MsgBox
                        
                        If mId$(wTxId, 1, 3) = "A0V" Then
                            V = cmdSelect_SQL_6_Control_détail(wTxId, wCdtrAcct, tCcy, xCur)
                            If Not IsNull(V) Then GoTo Error_MsgBox
                        End If
                            
                        tNbOfTxs = tNbOfTxs + 1
                        tAmt = tAmt + xCur
                    End If
                End If

            End If
                
            
        End If
    Loop
    Close #1
End If

If blnMsgId Then
    If tNbOfTxs <> mNbOfTxs Then V = "NB détail  : " & tNbOfTxs & " <> " & mNbOfTxs: GoTo Error_MsgBox
    If tAmt <> mAmt Then V = "Total détail  : " & tAmt & " <> " & mAmt: GoTo Error_MsgBox

    cmdSelect_SQL_6_Control = Null
    Print #501, Now & " | " & "contrôle Nb et montant OK : préparation du Fax"
    
    'Call cmdPrint_SEPA_Aller_Fax(lFile, mNbOfTxs, mAmt)

End If

Close #501

Exit Function

Error_Handler:
    V = Error
Error_MsgBox:
    Error_Route V
    Print #501, Now & " | " & V


End Function

Public Function cmdSelect_SQL_6_Control_A3V(lFile As String)
Dim wMsgId As String
Dim wFile As String
Dim blnAssgnmt As Boolean
Dim MsgId As String
Dim xIn As String
Dim Fic As Long
Dim K1 As Long
Dim K2 As Long

    On Error GoTo Error_Handler
    cmdSelect_SQL_6_Control_A3V = "?"
    Call lstErr_AddItem(lstErr, cmdContext, lFile): DoEvents
    wMsgId = Replace(lFile, ".xml", "")
    wFile = paramSEPA_Aller_XCOM & lFile
    blnAssgnmt = False
    MsgId = ""
    K1 = -1
    K2 = -1
    If Dir(wFile) <> "" Then
        Fic = FreeFile
        Open wFile For Input As #Fic
        Do Until EOF(Fic)
            Line Input #Fic, xIn
            If Not blnAssgnmt Then
                If InStr(xIn, "<Assgnmt>") > 0 Then blnAssgnmt = True
            End If
            If blnAssgnmt Then
                If InStr(xIn, "<Id>") > 0 Then K1 = InStr(xIn, "<Id>") + 4
                If InStr(xIn, "</Id>") > 0 Then K2 = InStr(xIn, "</Id>")
                If K1 > -1 And K2 = -1 Then
                    MsgId = MsgId & Trim(mId(xIn, K1))
                End If
                If K1 > -1 And K2 > -1 And MsgId = "" Then
                    MsgId = mId(xIn, K1, K2 - K1)
                    blnAssgnmt = False
                ElseIf K1 > -1 And K2 > -1 And MsgId <> "" Then
                    MsgId = MsgId & Trim(mId(xIn, 1, K2 - 1))
                    blnAssgnmt = False
                End If
            End If
        Loop
    End If
    Close #Fic
    If MsgId <> "" Then
        Print #501, Now & " | " & "<Id> :" & Trim(MsgId)
        Print #501, Now & " | " & "Préparation du Fax"
        Call cmdPrint_SEPA_Aller_Fax(lFile, 1, 0)
    End If
    cmdSelect_SQL_6_Control_A3V = Null
    Exit Function
Error_Handler:
    V = Error
Error_MsgBox:
    Error_Route V
    Print #501, Now & " | " & V

End Function

Public Function cmdSelect_SQL_7_Emis(lFile As String, ldateEUPMON2 As Date)
Dim V, X As String, XCmd As String, XLog As String, xName As String, xMemo As String
Dim xIn As String, mIn As String, wFile As String
Dim blnFINOK As Boolean
On Error GoTo Error_Handler
cmdSelect_SQL_7_Emis = "?"
mIn = ""

Call lstErr_AddItem(lstErr, cmdContext, lFile): DoEvents

V = rsElpTable_Read("SEPA", "Aller", mId$(lFile, 1, 3) & ".log", xName, xMemo)
XLog = paramServer(Replace(xMemo, "%WIN%", paramEnvironnement))

blnFINOK = False
X = "FINOK: " & paramSEPA_Aller_XCOM & lFile

If Dir(XLog) <> "" Then
    Print #501, Now & " | " & "Contrôle : " & XLog
    Open XLog For Input As #1
    Do Until EOF(1)
        Line Input #1, xIn
        If mId$(xIn, 1, 1) = "$" Then mIn = xIn                 'horodatage
        If Trim(xIn) = X Then                                   ' télétransmission après maj YEUPMON2.EUPMON2STA = 'X'
             'obligé de tester l'égalité à partir du 08/11/2012 D.ROSILLETTE
             'If DateDiff("s", ldateEUPMON2, mId$(mIn, 2, 19)) > 0 Then
             If DateDiff("s", ldateEUPMON2, mId$(mIn, 2, 19)) >= 0 Then
                blnFINOK = True: Exit Do
            Else
                Print #501, Now & " | " & " !!! Télétransmission ancienne du " & mId$(mIn, 2, 19) & " ignorée !!!!"
            End If
            
        End If
    Loop
    Close #1
End If

If blnFINOK Then
    Print #501, Now & " | " & "Télétransmission OK : " & lFile
    cmdSelect_SQL_7_Emis = Null
Else
    Print #501, Now & " | " & "Télétransmission ?? :" & lFile
End If

Exit Function

Error_Handler:
    V = Error
Error_MsgBox:
    Error_Route V

End Function


Public Function cmdSelect_SQL_6_Emission(lFile As String)
'___________________________________________________________________
Dim V, X As String, XCmd As String, XLog As String, xName As String, xMemo As String
On Error GoTo Error_Handler
    
cmdSelect_SQL_6_Emission = Null

V = rsElpTable_Read("SEPA", "Aller", mId$(lFile, 1, 3) & ".cmd", xName, xMemo)
XCmd = paramServer(Replace(xMemo, "%WIN%", paramEnvironnement))
    
V = rsElpTable_Read("SEPA", "Aller", mId$(lFile, 1, 3) & ".log", xName, xMemo)
XLog = paramServer(Replace(xMemo, "%WIN%", paramEnvironnement))

X = XCmd & " " & paramSEPA_Aller_XCOM & lFile ''& " >> " & XLog
Call Shell_Exe(X)
Print #501, Now & " | " & "Télétransmission : " & X

Call lstErr_AddItem(lstErr, cmdContext, "@EUP_XCOM : 6 temporisation"): DoEvents
Sleep 10000

Exit Function

Error_Handler:
    V = Error
Error_MsgBox:
    Error_Route V
    Print #501, Now & " | " & V
    cmdSelect_SQL_6_Emission = V

End Function
Public Function cmdSelect_SQL_9_Réception(lFile As String)
'___________________________________________________________________
Dim V, X As String, XCmd As String, XLog As String, xName As String, xMemo As String
On Error GoTo Error_Handler
    
X = "cmdSelect_SQL_9_Réception"
    
cmdSelect_SQL_9_Réception = Null
V = rsElpTable_Read("SEPA", "Retour", mId$(lFile, 1, 3) & ".cmd", xName, xMemo)
XCmd = paramServer(Replace(xMemo, "%WIN%", paramEnvironnement))
    
V = rsElpTable_Read("SEPA", "Retour", mId$(lFile, 1, 3) & ".log", xName, xMemo)
XLog = paramServer(Replace(xMemo, "%WIN%", paramEnvironnement))

X = XCmd & " " & paramSEPA_Retour_XCOM & lFile & ".xml" ''& " >> " & XLog

Call Shell_Exe(X)
Print #501, Now & " | " & "Téléchargement : " & X

Call lstErr_AddItem(lstErr, cmdContext, "@EUP_XCOM : 9 temporisation"): DoEvents
Sleep 10000

Exit Function

Error_Handler:
    V = Error & " " & X
Error_MsgBox:
    Error_Route V
    Print #501, Now & " | " & V
    cmdSelect_SQL_9_Réception = V

End Function

Public Sub cmdSendMail_SEPA_SG(lSubject As String, lFile As String)
Dim wSendMail As typeSendMail
Dim bgColor As String
Dim xIn As String

    
wSendMail.FromDisplayName = "SEPA=>SG"
wSendMail.RecipientDisplayName = "XCOM"

wSendMail.Subject = lSubject
wSendMail.Attachment = Replace(lFile, ".log", "_Fax.doc")

Open lFile For Input As #501
Do Until EOF(501)
    Line Input #501, xIn
    xIn = Replace(xIn, "\\PELISRV\PELINT.dat", "\\...")
    wSendMail.Message = wSendMail.Message & Trim(xIn) & vbCrLf '"<HR>"
Loop
Close #501






wSendMail.AsHTML = False 'True

srvSendMail.Monitor wSendMail

End Sub

Public Sub logAller_Open(lX As String)
paramSEPA_Aller_Détail_log = paramSEPA_Aller_Path_log & Replace(oldYEUPMON2.EUPMON2FIC, ".xml", ".log")

Open paramSEPA_Aller_Détail_log For Append As #501

Print #501, Now & "_" & lX & "__________________________________________________________"

End Sub

Public Sub logRetour_Open(lX As String, lFile As String)

paramSEPA_Retour_Détail_log = paramSEPA_Retour_Path_log & lFile & ".log"

Open paramSEPA_Retour_Détail_log For Append As #501

Print #501, Now & "_" & lX & "______________________________________________________________"

End Sub

Public Function cmdSelect_SQL_6_Control_détail(lTxId As String, lCdtrAcct As String, lCcy As String, lCur As Currency)
On Error GoTo Error_Handler
Dim xSQL As String, wEUPG2A_IBAN As String, wEUPG2ADEV As String, wEUPG2AMON As Currency
cmdSelect_SQL_6_Control_détail = "? opération inconnue ZEUPG2A0/ZEUPG3A0"

xSQL = "select * from " & paramIBM_Library_SAB & ".ZEUPG2A0 " _
     & " where EUPG2AETB = " & currentZMNURUT0.MNURUTETB _
     & " AND   EUPG2AOPE = '" & mId$(lTxId, 1, 3) & "'" _
     & " AND   EUPG2ANUM = " & mId$(lTxId, 5, 9) _
     & " AND   EUPG2ACRE = " & mId$(lTxId, 15, 7)

Set rsSab = cnsab.Execute(xSQL)

If Not rsSab.EOF Then
    wEUPG2A_IBAN = rsSab("EUPG2APDT") & rsSab("EUPG2AIDS") & rsSab("EUPG2ABDT")
    wEUPG2ADEV = rsSab("EUPG2ADEV")
    wEUPG2AMON = rsSab("EUPG2AMON")
Else
    xSQL = "select * from " & paramIBM_Library_SAB & ".ZEUPG3A0 " _
         & " where EUPG3AETB = " & currentZMNURUT0.MNURUTETB _
         & " AND   EUPG3AOPE = '" & mId$(lTxId, 1, 3) & "'" _
         & " AND   EUPG3ADET = " & mId$(lTxId, 5, 9) _
         & " AND   EUPG3ACRE = " & mId$(lTxId, 15, 7)
    
    Set rsSab = cnsab.Execute(xSQL)
    
    If Not rsSab.EOF Then
        wEUPG2A_IBAN = rsSab("EUPG3APDT") & rsSab("EUPG3AIDS") & rsSab("EUPG3ABDT")
        wEUPG2ADEV = rsSab("EUPG3ADEV")
        wEUPG2AMON = rsSab("EUPG3AMON")
    Else
        Exit Function
    End If
End If

If Trim(wEUPG2A_IBAN) <> lCdtrAcct Then cmdSelect_SQL_6_Control_détail = "IBAN différent de l'IBAN d'origine (ZEUPG2A0)": Exit Function
If Trim(wEUPG2ADEV) <> lCcy Then cmdSelect_SQL_6_Control_détail = "Devise différente de la devise d'origine (ZEUPG2A0)": Exit Function

'$JPL 2009-07-24 pb avec SAB qui remet à zéro les montants des lots
'If Trim(wEUPG2AMON) <> lCur Then cmdSelect_SQL_6_Control_détail = "Montant différent du Montant d'origine (ZEUPG2A0)": Exit Function

cmdSelect_SQL_6_Control_détail = Null

Exit Function

Error_Handler:
    cmdSelect_SQL_6_Control_détail = Error


End Function

Public Sub param_Init_Explorer_IFS()
Dim X As String
If Not blnExplorer_IFS Then
    blnExplorer_IFS = True
    X = "CMD/C START/MIN explorer \\" & paramIBM_AS400_ID & "\IFS"
    
    Close #1
    Open paramSchtasks_Bat For Output As #1
    Print #1, "echo $%date% %time%"
    Print #1, X & " >> " & paramSEPA_Aller_log
    Close #1
    
    Call Shell_Exe(paramSchtasks_Run)
End If
End Sub

Public Sub fraRetour_R0V_Display()
Dim xSQL As String
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIAMON7  where MONAPP = 'SEPA' and MONFLUX = 'R0V'"
Set rsSab = cnsab.Execute(xSQL)

cmdRetour_R0V.Visible = False
If Not rsSab.EOF Then
    V = rsYBIAMON0_GetBuffer(rsSab, oldYBIAMON0_R0V)
Else
    Call rsYBIAMON0_Init(oldYBIAMON0_R0V)
    oldYBIAMON0_R0V.MONSTATUS = "? SEPA - R0V"
End If
txtRetour_R0V_MONSTATUS = oldYBIAMON0_R0V.MONSTATUS
lblRetour_R0V_MONAMJ = dateImp10_S(oldYBIAMON0_R0V.MONAMJ) & " - " & timeImp8(oldYBIAMON0_R0V.MONHMS)
txtRetour_R0V_MONUSR = oldYBIAMON0_R0V.MONUSR
lblRetour_R0V_MONJOB = oldYBIAMON0_R0V.MONJOB & " - " & oldYBIAMON0_R0V.MONPGM
txtRetour_R0V_MONFILE = oldYBIAMON0_R0V.MONFILE

Select Case Trim(oldYBIAMON0_R0V.MONSTATUS)
    Case "": cmdRetour_R0V.Caption = "Demande de téléchargement RETOUR":
             cmdRetour_R0V.BackColor = &HC0FFC0
             cmdRetour_R0V.Visible = True
    Case "Demande", "X_Erreur":
             cmdRetour_R0V.Caption = "Annuler le traitement RETOUR en cours"
             cmdRetour_R0V.BackColor = &HC0C0FF
             cmdRetour_R0V.Visible = True
End Select
        

End Sub
Public Sub fraRetour_R2V_Display()
Dim xSQL As String
xSQL = "select * from " & paramIBM_Library_SABSPE & ".YBIAMON7  where MONAPP = 'SEPA' and MONFLUX = 'R2V'"
Set rsSab = cnsab.Execute(xSQL)

cmdRetour_R2V.Visible = False

If Not rsSab.EOF Then
    V = rsYBIAMON0_GetBuffer(rsSab, oldYBIAMON0_R2V)
Else
    Call rsYBIAMON0_Init(oldYBIAMON0_R2V)
    oldYBIAMON0_R2V.MONSTATUS = "? SEPA - R2V"
End If
txtRetour_R2V_MONSTATUS = oldYBIAMON0_R2V.MONSTATUS
lblRetour_R2V_MONAMJ = dateImp10_S(oldYBIAMON0_R2V.MONAMJ) & " - " & timeImp8(oldYBIAMON0_R2V.MONHMS)
txtRetour_R2V_MONUSR = oldYBIAMON0_R2V.MONUSR
lblRetour_R2V_MONJOB = oldYBIAMON0_R2V.MONJOB & " - " & oldYBIAMON0_R2V.MONPGM
txtRetour_R2V_MONFILE = oldYBIAMON0_R2V.MONFILE


Select Case Trim(oldYBIAMON0_R2V.MONSTATUS)
    Case "": cmdRetour_R2V.Caption = "Demande de téléchargement RETOUR REJET":
             cmdRetour_R2V.BackColor = &HC0FFC0
             cmdRetour_R2V.Visible = True
    Case "Demande", "X_Erreur":
             cmdRetour_R2V.Caption = "Annuler le traitement RETOUR REJET en cours"
             cmdRetour_R2V.BackColor = &HC0C0FF
             cmdRetour_R2V.Visible = True
End Select

End Sub


Public Sub cmdSelect_SQL_9_Demande()

newYBIAMON0.MONNUM = newYBIAMON0.MONNUM + 1
newYBIAMON0.MONFILE = mId$(newYBIAMON0.MONFLUX, 1, 3) & "_" & Format(newYBIAMON0.MONNUM, "000000")
paramSEPA_Retour_XCOM_File = paramSEPA_Retour_XCOM & newYBIAMON0.MONFILE & ".xml"

Call logRetour_Open(9, newYBIAMON0.MONFILE)

If Dir(paramToken) <> "" Then                          '? transfert en cours ?
    Print #501, Now & " | " & "en attente autre transfert en cours > " & paramToken
    Call lstErr_AddItem(lstErr, cmdContext, paramToken): DoEvents
Else
    If Dir(paramSEPA_Retour_XCOM_File) <> "" Then           ' ?fichier déjà présent
        cmdSelect_SQL_9_Erreur "fichier déjà présent"
    Else
    
         V = cmdSelect_SQL_9_Réception(newYBIAMON0.MONFILE)   ' lancer la réception
         
         If IsNull(V) Then
             newYBIAMON0.MONSTATUS = "X_enCours"
             V = sqlYBIAMON0_Transaction
             If Not IsNull(V) Then GoTo Error_MsgBox
             Print #501, Now & " | " & "Statut : 'X_enCours' " & paramIBM_Library_SABSPE & ".YBIAMON7 | " & paramSEPA_Retour_XCOM_File
                        
         Else
            cmdSelect_SQL_9_Erreur CStr(V)
         End If
    End If

End If
    
Close #501
On Error GoTo Error_Handler
Exit Sub
Error_Handler:
    V = Error
Error_MsgBox:
    Error_Route V
    Print #501, Now & " | " & V

End Sub

Public Sub cmdSelect_SQL_9_Erreur(lMsg As String)

Print #501, Now & " | " & lMsg & " : " & paramSEPA_Retour_XCOM_File
Call lstErr_AddItem(lstErr, cmdContext, lMsg & " : " & paramSEPA_Retour_XCOM_File): DoEvents

newYBIAMON0 = oldYBIAMON0
newYBIAMON0.MONSTATUS = "X_Erreur"
sqlYBIAMON0_Transaction
'cmdSendMail_SEPA_Alerte "Télétransmission : Anomalie du contenu du fichier : " & paramSEPA_Retour_Détail_log, currentError, paramSEPA_Retour_Détail_log

End Sub
