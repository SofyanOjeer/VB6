VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmLrAttributDétail 
   Caption         =   "Lr Détail des attributs"
   ClientHeight    =   7080
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10335
   LinkTopic       =   "Form1"
   ScaleHeight     =   7080
   ScaleWidth      =   10335
   Begin VB.TextBox txtRéférence 
      Height          =   285
      Left            =   5760
      TabIndex        =   0
      Top             =   0
      Width           =   1455
   End
   Begin VB.CommandButton cmdContext 
      Caption         =   "&Abandonner"
      Height          =   300
      Left            =   0
      TabIndex        =   111
      Top             =   0
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00C0FFC0&
      Caption         =   "&Ok"
      Height          =   300
      Left            =   1100
      Style           =   1  'Graphical
      TabIndex        =   110
      TabStop         =   0   'False
      Top             =   0
      Width           =   1000
   End
   Begin VB.ListBox lstErr 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   7320
      TabIndex        =   109
      Top             =   0
      Width           =   2505
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   9840
      Picture         =   "LrAttributDétail.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   108
      Top             =   0
      Width           =   400
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6675
      Left            =   0
      TabIndex        =   105
      Top             =   360
      Width           =   10260
      _ExtentX        =   18098
      _ExtentY        =   11774
      _Version        =   393216
      Tabs            =   9
      TabsPerRow      =   7
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "AFFPU.Cdzon"
      TabPicture(0)   =   "LrAttributDétail.frx":0102
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblAFFPU"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblAGEMT"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblAGENT"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblAPPAR"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblAREFR"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblATTCF"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblAUTDV"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblBONIF"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblCAROB"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblCATET"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblCDRES"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblCDZON"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cboAFFPU"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cboAGEMT"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cboAGENT"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "cboCATET"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "cboCAROB"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "cboBONIF"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "cboAUTDV"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "cboATTCF"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "cboAREFR"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "cboAPPAR"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "cboCDRES"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "cboCDZON"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).ControlCount=   24
      TabCaption(1)   =   "CLCRC.Durom"
      TabPicture(1)   =   "LrAttributDétail.frx":011E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblCRDIV"
      Tab(1).Control(1)=   "lblCREIM"
      Tab(1).Control(2)=   "lblCREOR"
      Tab(1).Control(3)=   "lblCRETC"
      Tab(1).Control(4)=   "lblCRHYP"
      Tab(1).Control(5)=   "lblDCTOM"
      Tab(1).Control(6)=   "lblDRAC"
      Tab(1).Control(7)=   "lblCLCRC"
      Tab(1).Control(8)=   "lblCOTIT"
      Tab(1).Control(9)=   "lblCPEMS"
      Tab(1).Control(10)=   "lblDURIN"
      Tab(1).Control(11)=   "lblDUROM"
      Tab(1).Control(12)=   "cboDRAC"
      Tab(1).Control(13)=   "cboDCTOM"
      Tab(1).Control(14)=   "cboCRHYP"
      Tab(1).Control(15)=   "cboCRETC"
      Tab(1).Control(16)=   "cboCREOR"
      Tab(1).Control(17)=   "cboCREIM"
      Tab(1).Control(18)=   "cboCRDIV"
      Tab(1).Control(19)=   "cboCPEMS"
      Tab(1).Control(20)=   "cboCOTIT"
      Tab(1).Control(21)=   "cboCLCRC"
      Tab(1).Control(22)=   "cboDURIN"
      Tab(1).Control(23)=   "cboDUROM"
      Tab(1).ControlCount=   24
      TabCaption(2)   =   "DVOPR.Nacga"
      TabPicture(2)   =   "LrAttributDétail.frx":013A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblDVOPR"
      Tab(2).Control(1)=   "lblECART"
      Tab(2).Control(2)=   "lblECFIN"
      Tab(2).Control(3)=   "lblELIGB"
      Tab(2).Control(4)=   "lblFAMDV"
      Tab(2).Control(5)=   "lblFOPIF"
      Tab(2).Control(6)=   "lblFPRBG"
      Tab(2).Control(7)=   "lblGARCF"
      Tab(2).Control(8)=   "lblMLFCE"
      Tab(2).Control(9)=   "lblMONDV"
      Tab(2).Control(10)=   "lblMUTFG"
      Tab(2).Control(11)=   "lblNACGA"
      Tab(2).Control(12)=   "cboMONDV"
      Tab(2).Control(13)=   "cboMLFCE"
      Tab(2).Control(14)=   "cboGARCF"
      Tab(2).Control(15)=   "cboFPRBG"
      Tab(2).Control(16)=   "cboFOPIF"
      Tab(2).Control(17)=   "cboFAMDV"
      Tab(2).Control(18)=   "cboELIGB"
      Tab(2).Control(19)=   "cboECFIN"
      Tab(2).Control(20)=   "cboECART"
      Tab(2).Control(21)=   "cboDVOPR"
      Tab(2).Control(22)=   "cboMUTFG"
      Tab(2).Control(23)=   "cboNACGA"
      Tab(2).ControlCount=   24
      TabCaption(3)   =   "NACGR.Nater"
      TabPicture(3)   =   "LrAttributDétail.frx":0156
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cboNATCP"
      Tab(3).Control(1)=   "cboNACGR"
      Tab(3).Control(2)=   "cboNACPS"
      Tab(3).Control(3)=   "cboNAEGA"
      Tab(3).Control(4)=   "cboNAIMO"
      Tab(3).Control(5)=   "cboNAOCB"
      Tab(3).Control(6)=   "cboNAPRO"
      Tab(3).Control(7)=   "cboNARCP"
      Tab(3).Control(8)=   "cboNATCR"
      Tab(3).Control(9)=   "cboNATCS"
      Tab(3).Control(10)=   "cboNATDD"
      Tab(3).Control(11)=   "cboNATER"
      Tab(3).Control(12)=   "lblNATCP"
      Tab(3).Control(13)=   "lblNATER"
      Tab(3).Control(14)=   "lblNATDD"
      Tab(3).Control(15)=   "lblNATCS"
      Tab(3).Control(16)=   "lblNATCR"
      Tab(3).Control(17)=   "lblNARCP"
      Tab(3).Control(18)=   "lblNAPRO"
      Tab(3).Control(19)=   "lblNAOCB"
      Tab(3).Control(20)=   "lblNAIMO"
      Tab(3).Control(21)=   "lblNAEGA"
      Tab(3).Control(22)=   "lblNACPS"
      Tab(3).Control(23)=   "lblNACGR"
      Tab(3).ControlCount=   24
      TabCaption(4)   =   "NATIF.Paact"
      TabPicture(4)   =   "LrAttributDétail.frx":0172
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "lblNATIT"
      Tab(4).Control(1)=   "lblNATMA"
      Tab(4).Control(2)=   "lblNATOF"
      Tab(4).Control(3)=   "lblNATRS"
      Tab(4).Control(4)=   "lblNRAST"
      Tab(4).Control(5)=   "lblNREHB"
      Tab(4).Control(6)=   "lblOPCVM"
      Tab(4).Control(7)=   "lblOPEFC"
      Tab(4).Control(8)=   "lblOPFDH"
      Tab(4).Control(9)=   "lblOPREC"
      Tab(4).Control(10)=   "lblPAACT"
      Tab(4).Control(11)=   "lblNATIF"
      Tab(4).Control(12)=   "cboPAACT"
      Tab(4).Control(13)=   "cboOPREC"
      Tab(4).Control(14)=   "cboOPFDH"
      Tab(4).Control(15)=   "cboOPEFC"
      Tab(4).Control(16)=   "cboOPCVM"
      Tab(4).Control(17)=   "cboNREHB"
      Tab(4).Control(18)=   "cboNRAST"
      Tab(4).Control(19)=   "cboNATRS"
      Tab(4).Control(20)=   "cboNATOF"
      Tab(4).Control(21)=   "cboNATMA"
      Tab(4).Control(22)=   "cboNATIT"
      Tab(4).Control(23)=   "cboNATIF"
      Tab(4).ControlCount=   24
      TabCaption(5)   =   "PRIMP.Tycgr"
      TabPicture(5)   =   "LrAttributDétail.frx":018E
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "lblPRIMP"
      Tab(5).Control(1)=   "lblPROCB"
      Tab(5).Control(2)=   "lblREDES"
      Tab(5).Control(3)=   "lblREDHB"
      Tab(5).Control(4)=   "lblRESET"
      Tab(5).Control(5)=   "lblREZON"
      Tab(5).Control(6)=   "lblRISPA"
      Tab(5).Control(7)=   "lblSEMNT"
      Tab(5).Control(8)=   "lblSENOP"
      Tab(5).Control(9)=   "lblTCFPE"
      Tab(5).Control(10)=   "lblTOPIF"
      Tab(5).Control(11)=   "lblPERIO"
      Tab(5).Control(12)=   "cboTOPIF"
      Tab(5).Control(13)=   "cboTCFPE"
      Tab(5).Control(14)=   "cboSENOP"
      Tab(5).Control(15)=   "cboSEMNT"
      Tab(5).Control(16)=   "cboRISPA"
      Tab(5).Control(17)=   "cboREZON"
      Tab(5).Control(18)=   "cboRESET"
      Tab(5).Control(19)=   "cboREDHB"
      Tab(5).Control(20)=   "cboREDES"
      Tab(5).Control(21)=   "cboPROCB"
      Tab(5).Control(22)=   "cboPRIMP"
      Tab(5).Control(23)=   "cboPERIO"
      Tab(5).ControlCount=   24
      TabCaption(6)   =   "TYCOM.Zacti"
      TabPicture(6)   =   "LrAttributDétail.frx":01AA
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "lblZACTI"
      Tab(6).Control(1)=   "lblTYRES"
      Tab(6).Control(2)=   "lblTYPSU"
      Tab(6).Control(3)=   "lblTYPOR"
      Tab(6).Control(4)=   "lblTYETS"
      Tab(6).Control(5)=   "lblTYDSU"
      Tab(6).Control(6)=   "lblTYCOM"
      Tab(6).Control(7)=   "lblTYCGR"
      Tab(6).Control(8)=   "lblREESC1"
      Tab(6).Control(9)=   "lblREESC6"
      Tab(6).Control(10)=   "cboTYDSU"
      Tab(6).Control(11)=   "cboTYCOM"
      Tab(6).Control(12)=   "cboZACTI"
      Tab(6).Control(13)=   "cboTYRES"
      Tab(6).Control(14)=   "cboTYPSU"
      Tab(6).Control(15)=   "cboTYPOR"
      Tab(6).Control(16)=   "cboTYETS"
      Tab(6).Control(17)=   "cboTYCGR"
      Tab(6).Control(18)=   "cboREESC1"
      Tab(6).Control(19)=   "cboREESC6"
      Tab(6).ControlCount=   20
      TabCaption(7)   =   "Risques 1"
      TabPicture(7)   =   "LrAttributDétail.frx":01C6
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "cboCDCPFU"
      Tab(7).Control(1)=   "cboCDCRAC"
      Tab(7).Control(2)=   "cboCDCRTI"
      Tab(7).Control(3)=   "cboDUINIT"
      Tab(7).Control(4)=   "cboCDDURE"
      Tab(7).Control(5)=   "cboCRVENT"
      Tab(7).Control(6)=   "cboTYVENT"
      Tab(7).Control(7)=   "cboTYMTDV"
      Tab(7).Control(8)=   "cboCDREME"
      Tab(7).Control(9)=   "cboCDAGCO"
      Tab(7).Control(10)=   "cboCDCPJO"
      Tab(7).Control(11)=   "cboCDCPCO"
      Tab(7).Control(12)=   "lblCDCRAC"
      Tab(7).Control(13)=   "lblCDCRTI"
      Tab(7).Control(14)=   "lblDUINIT"
      Tab(7).Control(15)=   "lblCDDURE"
      Tab(7).Control(16)=   "lblCRVENT"
      Tab(7).Control(17)=   "lblTYVENT"
      Tab(7).Control(18)=   "lblTYMTDV"
      Tab(7).Control(19)=   "lblCDREME"
      Tab(7).Control(20)=   "lblCDAGCO"
      Tab(7).Control(21)=   "lblCDCPFU"
      Tab(7).Control(22)=   "lblCDCPJO"
      Tab(7).Control(23)=   "lblCDCPCO"
      Tab(7).ControlCount=   24
      TabCaption(8)   =   "Risques 2"
      TabPicture(8)   =   "LrAttributDétail.frx":01E2
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "cboCDSWAP"
      Tab(8).Control(1)=   "cboCDOPIM"
      Tab(8).Control(2)=   "cboCDOMPO"
      Tab(8).Control(3)=   "cboCDCRET"
      Tab(8).Control(4)=   "cboCDLODA"
      Tab(8).Control(5)=   "cboCDCREF"
      Tab(8).Control(6)=   "cboCDCRCO"
      Tab(8).Control(7)=   "cboCDCRIM"
      Tab(8).Control(8)=   "cboCDDEIN"
      Tab(8).Control(9)=   "cboCDBIOR"
      Tab(8).Control(10)=   "lblCDSWAP"
      Tab(8).Control(11)=   "lblCDOPIM"
      Tab(8).Control(12)=   "lblCDOMPO"
      Tab(8).Control(13)=   "lblCDCRET"
      Tab(8).Control(14)=   "lblCDLODA"
      Tab(8).Control(15)=   "lblCDCREF"
      Tab(8).Control(16)=   "lblCDCRCO"
      Tab(8).Control(17)=   "lblCDCRIM"
      Tab(8).Control(18)=   "lblCDDEIN"
      Tab(8).Control(19)=   "lblCDBIOR"
      Tab(8).ControlCount=   20
      Begin VB.ComboBox cboREESC6 
         ForeColor       =   &H00C00000&
         Height          =   315
         ItemData        =   "LrAttributDétail.frx":01FE
         Left            =   -70700
         List            =   "LrAttributDétail.frx":0200
         Sorted          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   82
         Top             =   5700
         Width           =   5700
      End
      Begin VB.ComboBox cboREESC1 
         ForeColor       =   &H00C00000&
         Height          =   315
         ItemData        =   "LrAttributDétail.frx":0202
         Left            =   -70700
         List            =   "LrAttributDétail.frx":0204
         Sorted          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   81
         Top             =   5200
         Width           =   5700
      End
      Begin VB.ComboBox cboNATCP 
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   -70680
         Style           =   2  'Dropdown List
         TabIndex        =   44
         Top             =   4200
         Width           =   5700
      End
      Begin VB.ComboBox cboNATIF 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70680
         Style           =   2  'Dropdown List
         TabIndex        =   49
         Top             =   720
         Width           =   5700
      End
      Begin VB.ComboBox cboCDCPFU 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70695
         Sorted          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   85
         Top             =   1695
         Width           =   5700
      End
      Begin VB.ComboBox cboCDSWAP 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70680
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   104
         Top             =   5200
         Width           =   5700
      End
      Begin VB.ComboBox cboCDOPIM 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   103
         Top             =   4700
         Width           =   5700
      End
      Begin VB.ComboBox cboCDOMPO 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   102
         Top             =   4200
         Width           =   5700
      End
      Begin VB.ComboBox cboCDCRET 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70680
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   101
         Top             =   3700
         Width           =   5700
      End
      Begin VB.ComboBox cboCDLODA 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   100
         Top             =   3200
         Width           =   5700
      End
      Begin VB.ComboBox cboCDCREF 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70680
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   99
         Top             =   2700
         Width           =   5700
      End
      Begin VB.ComboBox cboCDCRCO 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   98
         Top             =   2200
         Width           =   5700
      End
      Begin VB.ComboBox cboCDCRIM 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   97
         Top             =   1700
         Width           =   5700
      End
      Begin VB.ComboBox cboCDDEIN 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   96
         Top             =   1200
         Width           =   5700
      End
      Begin VB.ComboBox cboCDBIOR 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   95
         Top             =   700
         Width           =   5700
      End
      Begin VB.ComboBox cboCDCRAC 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   94
         Top             =   6200
         Width           =   5700
      End
      Begin VB.ComboBox cboCDCRTI 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   93
         Top             =   5700
         Width           =   5700
      End
      Begin VB.ComboBox cboDUINIT 
         ForeColor       =   &H00C00000&
         Height          =   315
         ItemData        =   "LrAttributDétail.frx":0206
         Left            =   -70695
         List            =   "LrAttributDétail.frx":0208
         Sorted          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   92
         Top             =   5205
         Width           =   5700
      End
      Begin VB.ComboBox cboCDDURE 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   91
         Top             =   4700
         Width           =   5700
      End
      Begin VB.ComboBox cboCRVENT 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70695
         Sorted          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   90
         Top             =   4200
         Width           =   5700
      End
      Begin VB.ComboBox cboTYVENT 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   89
         Top             =   3700
         Width           =   5700
      End
      Begin VB.ComboBox cboTYMTDV 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   88
         Top             =   3200
         Width           =   5700
      End
      Begin VB.ComboBox cboCDREME 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   87
         Top             =   2700
         Width           =   5700
      End
      Begin VB.ComboBox cboCDAGCO 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   86
         Top             =   2200
         Width           =   5700
      End
      Begin VB.ComboBox cboCDCPJO 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   84
         Top             =   1200
         Width           =   5700
      End
      Begin VB.ComboBox cboCDCPCO 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   83
         Top             =   700
         Width           =   5700
      End
      Begin VB.ComboBox cboPERIO 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Style           =   2  'Dropdown List
         TabIndex        =   61
         Top             =   700
         Width           =   5700
      End
      Begin VB.ComboBox cboTYCGR 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Style           =   2  'Dropdown List
         TabIndex        =   73
         Top             =   700
         Width           =   5700
      End
      Begin VB.ComboBox cboTYETS 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Style           =   2  'Dropdown List
         TabIndex        =   76
         Top             =   2200
         Width           =   5700
      End
      Begin VB.ComboBox cboTYPOR 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Style           =   2  'Dropdown List
         TabIndex        =   77
         Top             =   2700
         Width           =   5700
      End
      Begin VB.ComboBox cboTYPSU 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Style           =   2  'Dropdown List
         TabIndex        =   78
         Top             =   3200
         Width           =   5700
      End
      Begin VB.ComboBox cboTYRES 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Style           =   2  'Dropdown List
         TabIndex        =   79
         Top             =   3700
         Width           =   5700
      End
      Begin VB.ComboBox cboZACTI 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Style           =   2  'Dropdown List
         TabIndex        =   80
         Top             =   4200
         Width           =   5700
      End
      Begin VB.ComboBox cboTYCOM 
         ForeColor       =   &H00C00000&
         Height          =   315
         ItemData        =   "LrAttributDétail.frx":020A
         Left            =   -70700
         List            =   "LrAttributDétail.frx":020C
         Style           =   2  'Dropdown List
         TabIndex        =   74
         Top             =   1200
         Width           =   5700
      End
      Begin VB.ComboBox cboTYDSU 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Style           =   2  'Dropdown List
         TabIndex        =   75
         Top             =   1700
         Width           =   5700
      End
      Begin VB.ComboBox cboPRIMP 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Style           =   2  'Dropdown List
         TabIndex        =   62
         Top             =   1200
         Width           =   5700
      End
      Begin VB.ComboBox cboNACGA 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   6200
         Width           =   5700
      End
      Begin VB.ComboBox cboMUTFG 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   5700
         Width           =   5700
      End
      Begin VB.ComboBox cboDUROM 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   6200
         Width           =   5700
      End
      Begin VB.ComboBox cboDURIN 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   5700
         Width           =   5700
      End
      Begin VB.ComboBox cboCDZON 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   4300
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   6200
         Width           =   5700
      End
      Begin VB.ComboBox cboPROCB 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Style           =   2  'Dropdown List
         TabIndex        =   63
         Top             =   1700
         Width           =   5700
      End
      Begin VB.ComboBox cboREDES 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Style           =   2  'Dropdown List
         TabIndex        =   64
         Top             =   2200
         Width           =   5700
      End
      Begin VB.ComboBox cboREDHB 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Style           =   2  'Dropdown List
         TabIndex        =   65
         Top             =   2700
         Width           =   5700
      End
      Begin VB.ComboBox cboRESET 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Style           =   2  'Dropdown List
         TabIndex        =   66
         Top             =   3200
         Width           =   5700
      End
      Begin VB.ComboBox cboREZON 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Style           =   2  'Dropdown List
         TabIndex        =   67
         Top             =   3700
         Width           =   5700
      End
      Begin VB.ComboBox cboRISPA 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Style           =   2  'Dropdown List
         TabIndex        =   68
         Top             =   4200
         Width           =   5700
      End
      Begin VB.ComboBox cboSEMNT 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Style           =   2  'Dropdown List
         TabIndex        =   69
         Top             =   4680
         Width           =   5700
      End
      Begin VB.ComboBox cboSENOP 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Style           =   2  'Dropdown List
         TabIndex        =   70
         Top             =   5200
         Width           =   5700
      End
      Begin VB.ComboBox cboTCFPE 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Style           =   2  'Dropdown List
         TabIndex        =   71
         Top             =   5700
         Width           =   5700
      End
      Begin VB.ComboBox cboTOPIF 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Style           =   2  'Dropdown List
         TabIndex        =   72
         Top             =   6200
         Width           =   5700
      End
      Begin VB.ComboBox cboNATIT 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Style           =   2  'Dropdown List
         TabIndex        =   50
         Top             =   1200
         Width           =   5700
      End
      Begin VB.ComboBox cboNATMA 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Style           =   2  'Dropdown List
         TabIndex        =   51
         Top             =   1700
         Width           =   5700
      End
      Begin VB.ComboBox cboNATOF 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Style           =   2  'Dropdown List
         TabIndex        =   52
         Top             =   2200
         Width           =   5700
      End
      Begin VB.ComboBox cboNATRS 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Style           =   2  'Dropdown List
         TabIndex        =   53
         Top             =   2700
         Width           =   5700
      End
      Begin VB.ComboBox cboNRAST 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Style           =   2  'Dropdown List
         TabIndex        =   54
         Top             =   3200
         Width           =   5700
      End
      Begin VB.ComboBox cboNREHB 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Style           =   2  'Dropdown List
         TabIndex        =   55
         Top             =   3700
         Width           =   5700
      End
      Begin VB.ComboBox cboOPCVM 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Style           =   2  'Dropdown List
         TabIndex        =   56
         Top             =   4200
         Width           =   5700
      End
      Begin VB.ComboBox cboOPEFC 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Style           =   2  'Dropdown List
         TabIndex        =   57
         Top             =   4700
         Width           =   5700
      End
      Begin VB.ComboBox cboOPFDH 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Style           =   2  'Dropdown List
         TabIndex        =   58
         Top             =   5200
         Width           =   5700
      End
      Begin VB.ComboBox cboOPREC 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Style           =   2  'Dropdown List
         TabIndex        =   59
         Top             =   5700
         Width           =   5700
      End
      Begin VB.ComboBox cboPAACT 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Style           =   2  'Dropdown List
         TabIndex        =   60
         Top             =   6200
         Width           =   5700
      End
      Begin VB.ComboBox cboNACGR 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   700
         Width           =   5700
      End
      Begin VB.ComboBox cboNACPS 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   1200
         Width           =   5700
      End
      Begin VB.ComboBox cboNAEGA 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   1700
         Width           =   5700
      End
      Begin VB.ComboBox cboNAIMO 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   2160
         Width           =   5700
      End
      Begin VB.ComboBox cboNAOCB 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   2640
         Width           =   5700
      End
      Begin VB.ComboBox cboNAPRO 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   3200
         Width           =   5700
      End
      Begin VB.ComboBox cboNARCP 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   3700
         Width           =   5700
      End
      Begin VB.ComboBox cboNATCR 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70680
         Style           =   2  'Dropdown List
         TabIndex        =   45
         Top             =   4680
         Width           =   5700
      End
      Begin VB.ComboBox cboNATCS 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Style           =   2  'Dropdown List
         TabIndex        =   46
         Top             =   5200
         Width           =   5700
      End
      Begin VB.ComboBox cboNATDD 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Style           =   2  'Dropdown List
         TabIndex        =   47
         Top             =   5700
         Width           =   5700
      End
      Begin VB.ComboBox cboNATER 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Style           =   2  'Dropdown List
         TabIndex        =   48
         Top             =   6200
         Width           =   5700
      End
      Begin VB.ComboBox cboDVOPR 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   700
         Width           =   5700
      End
      Begin VB.ComboBox cboECART 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   1200
         Width           =   5700
      End
      Begin VB.ComboBox cboECFIN 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   1700
         Width           =   5700
      End
      Begin VB.ComboBox cboELIGB 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   2200
         Width           =   5700
      End
      Begin VB.ComboBox cboFAMDV 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   2700
         Width           =   5700
      End
      Begin VB.ComboBox cboFOPIF 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   3200
         Width           =   5700
      End
      Begin VB.ComboBox cboFPRBG 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   3700
         Width           =   5700
      End
      Begin VB.ComboBox cboGARCF 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   4200
         Width           =   5700
      End
      Begin VB.ComboBox cboMLFCE 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   4700
         Width           =   5700
      End
      Begin VB.ComboBox cboMONDV 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   5200
         Width           =   5700
      End
      Begin VB.ComboBox cboCLCRC 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   700
         Width           =   5700
      End
      Begin VB.ComboBox cboCOTIT 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1200
         Width           =   5700
      End
      Begin VB.ComboBox cboCPEMS 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1700
         Width           =   5700
      End
      Begin VB.ComboBox cboCRDIV 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   2200
         Width           =   5700
      End
      Begin VB.ComboBox cboCREIM 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   2700
         Width           =   5700
      End
      Begin VB.ComboBox cboCREOR 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   3200
         Width           =   5700
      End
      Begin VB.ComboBox cboCRETC 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   3700
         Width           =   5700
      End
      Begin VB.ComboBox cboCRHYP 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   4200
         Width           =   5700
      End
      Begin VB.ComboBox cboDCTOM 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   4700
         Width           =   5700
      End
      Begin VB.ComboBox cboDRAC 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -70700
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   5200
         Width           =   5700
      End
      Begin VB.ComboBox cboCDRES 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   4300
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   5700
         Width           =   5700
      End
      Begin VB.ComboBox cboAPPAR 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   4300
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2200
         Width           =   5700
      End
      Begin VB.ComboBox cboAREFR 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   4300
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2700
         Width           =   5700
      End
      Begin VB.ComboBox cboATTCF 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   4300
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   3200
         Width           =   5700
      End
      Begin VB.ComboBox cboAUTDV 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   4300
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   3700
         Width           =   5700
      End
      Begin VB.ComboBox cboBONIF 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   4300
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   4200
         Width           =   5700
      End
      Begin VB.ComboBox cboCAROB 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   4300
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   4700
         Width           =   5700
      End
      Begin VB.ComboBox cboCATET 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   4300
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   5200
         Width           =   5700
      End
      Begin VB.ComboBox cboAGENT 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   4300
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1700
         Width           =   5700
      End
      Begin VB.ComboBox cboAGEMT 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   4300
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1200
         Width           =   5700
      End
      Begin VB.ComboBox cboAFFPU 
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   4300
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   700
         Width           =   5700
      End
      Begin VB.Label lblREESC6 
         Caption         =   "PCI 6-7 créances && dettes rattachées"
         Height          =   255
         Left            =   -74880
         TabIndex        =   214
         Top             =   5700
         Width           =   4095
      End
      Begin VB.Label lblREESC1 
         Caption         =   "PCI 1-5 créances && dettes ratachées"
         Height          =   255
         Left            =   -74880
         TabIndex        =   213
         Top             =   5200
         Width           =   4095
      End
      Begin VB.Label lblNATCP 
         Caption         =   "NATCP Nature de la Contrepartie"
         Height          =   255
         Left            =   -74880
         TabIndex        =   212
         Top             =   4200
         Width           =   4095
      End
      Begin VB.Label lblNATIF 
         Caption         =   "NATIF nature des instruments financiers"
         Height          =   255
         Left            =   -74880
         TabIndex        =   211
         Top             =   700
         Width           =   4095
      End
      Begin VB.Label lblCDSWAP 
         Caption         =   "CDSWAP indicateur swap"
         Height          =   255
         Left            =   -74880
         TabIndex        =   210
         Top             =   5200
         Width           =   4215
      End
      Begin VB.Label lblCDOPIM 
         Caption         =   "CDOPIM encours financier opérations immobilières"
         Height          =   255
         Left            =   -74880
         TabIndex        =   209
         Top             =   4700
         Width           =   4215
      End
      Begin VB.Label lblCDOMPO 
         Caption         =   "CDOMPO encours financier opérations mobilières"
         Height          =   255
         Left            =   -74880
         TabIndex        =   208
         Top             =   4200
         Width           =   4215
      End
      Begin VB.Label lblCDCRET 
         Caption         =   "CDCRET mobilisation créance sur l'étranger"
         Height          =   255
         Left            =   -74880
         TabIndex        =   207
         Top             =   3700
         Width           =   4215
      End
      Begin VB.Label lblCDLODA 
         Caption         =   "CDLODA loi Dailly"
         Height          =   255
         Left            =   -74880
         TabIndex        =   206
         Top             =   3200
         Width           =   4215
      End
      Begin VB.Label lblCDCREF 
         Caption         =   "CDCREF code crédit refinançable(pour DOM TOM)"
         Height          =   255
         Left            =   -74880
         TabIndex        =   205
         Top             =   2700
         Width           =   4215
      End
      Begin VB.Label lblCDCRCO 
         Caption         =   "CDCRCO crédit lié créance commerciale"
         Height          =   255
         Left            =   -74880
         TabIndex        =   204
         Top             =   2200
         Width           =   4215
      End
      Begin VB.Label lblCDCRIM 
         Caption         =   "CDCRIM code importation Crédoc"
         Height          =   255
         Left            =   -74880
         TabIndex        =   203
         Top             =   1700
         Width           =   4215
      End
      Begin VB.Label lblCDDEIN 
         Caption         =   "CDDEIN code dépôt indisponible"
         Height          =   255
         Left            =   -74880
         TabIndex        =   202
         Top             =   1200
         Width           =   4215
      End
      Begin VB.Label lblCDBIOR 
         Caption         =   "CDBIOR code biller à ordre"
         Height          =   255
         Left            =   -74880
         TabIndex        =   201
         Top             =   700
         Width           =   4215
      End
      Begin VB.Label lblCDCRAC 
         Caption         =   "CDCRAC particulrité crédit acheteur"
         Height          =   255
         Left            =   -74880
         TabIndex        =   200
         Top             =   6200
         Width           =   4215
      End
      Begin VB.Label lblCDCRTI 
         Caption         =   "CDCRTI code créance titrisée"
         Height          =   255
         Left            =   -74880
         TabIndex        =   199
         Top             =   5700
         Width           =   4215
      End
      Begin VB.Label lblDUINIT 
         Caption         =   "DUINIT Durée initiale en mois"
         Height          =   255
         Left            =   -74880
         TabIndex        =   198
         Top             =   5200
         Width           =   4215
      End
      Begin VB.Label lblCDDURE 
         Caption         =   "CDDURE code durée C/T"
         Height          =   255
         Left            =   -74880
         TabIndex        =   197
         Top             =   4700
         Width           =   4215
      End
      Begin VB.Label lblCRVENT 
         Caption         =   "CRVENT critère de ventilation"
         Height          =   255
         Left            =   -74880
         TabIndex        =   196
         Top             =   4200
         Width           =   4215
      End
      Begin VB.Label lblTYVENT 
         Caption         =   "TYVENT type de ventilation"
         Height          =   255
         Left            =   -74880
         TabIndex        =   195
         Top             =   3700
         Width           =   4215
      End
      Begin VB.Label lblTYMTDV 
         Caption         =   "TYMTDV Type de montant"
         Height          =   255
         Left            =   -74880
         TabIndex        =   194
         Top             =   3200
         Width           =   4215
      End
      Begin VB.Label lblCDREME 
         Caption         =   "CDREME code résident si DOM/TOM"
         Height          =   255
         Left            =   -74880
         TabIndex        =   193
         Top             =   2700
         Width           =   4215
      End
      Begin VB.Label lblCDAGCO 
         Caption         =   "CDAGCO Code agent économique"
         Height          =   255
         Left            =   -74880
         TabIndex        =   192
         Top             =   2200
         Width           =   4215
      End
      Begin VB.Label lblCDCPFU 
         Caption         =   "CDCPFU critère de fusion"
         Height          =   255
         Left            =   -74880
         TabIndex        =   191
         Top             =   1700
         Width           =   4215
      End
      Begin VB.Label lblCDCPJO 
         Caption         =   "CDCPJO Indicateur compte joint"
         Height          =   255
         Left            =   -74880
         TabIndex        =   190
         Top             =   1200
         Width           =   4215
      End
      Begin VB.Label lblCDCPCO 
         Caption         =   "CDCPCO Compte individuel / collectif"
         Height          =   255
         Left            =   -74880
         TabIndex        =   189
         Top             =   700
         Width           =   4215
      End
      Begin VB.Label lblPERIO 
         Caption         =   "PERIO pèriode"
         Height          =   255
         Left            =   -74880
         TabIndex        =   187
         Top             =   700
         Width           =   4095
      End
      Begin VB.Label lblTYCGR 
         Caption         =   "TYCGR type de contregarantie"
         Height          =   255
         Left            =   -74880
         TabIndex        =   186
         Top             =   700
         Width           =   4095
      End
      Begin VB.Label lblTYCOM 
         Caption         =   "TYCOM type de commission sur engagement sur titres"
         Height          =   255
         Left            =   -74880
         TabIndex        =   185
         Top             =   1200
         Width           =   4095
      End
      Begin VB.Label lblTYDSU 
         Caption         =   "TYDSU type dette subordonnée à durée indéterminé"
         Height          =   255
         Left            =   -74880
         TabIndex        =   184
         Top             =   1700
         Width           =   4095
      End
      Begin VB.Label lblTYETS 
         Caption         =   "TYETS type d'établissement (consolidé)"
         Height          =   255
         Left            =   -74880
         TabIndex        =   183
         Top             =   2200
         Width           =   4095
      End
      Begin VB.Label lblTYPOR 
         Caption         =   "TYPOR type de portefeuille"
         Height          =   255
         Left            =   -74880
         TabIndex        =   182
         Top             =   2700
         Width           =   4095
      End
      Begin VB.Label lblTYPSU 
         Caption         =   "TYPSU type de prêt subordonné"
         Height          =   255
         Left            =   -74880
         TabIndex        =   181
         Top             =   3200
         Width           =   4095
      End
      Begin VB.Label lblTYRES 
         Caption         =   "TYRES moyenne ""maximum""de l'état 4006"
         Height          =   255
         Left            =   -74880
         TabIndex        =   180
         Top             =   3700
         Width           =   4095
      End
      Begin VB.Label lblZACTI 
         Caption         =   "ZACTI zone d'activité de la société"
         Height          =   255
         Left            =   -74880
         TabIndex        =   179
         Top             =   4200
         Width           =   4095
      End
      Begin VB.Label lblNACGA 
         Caption         =   "NACGA nature de la contregarantie reçue"
         Height          =   255
         Left            =   -74880
         TabIndex        =   178
         Top             =   6200
         Width           =   4215
      End
      Begin VB.Label lblMUTFG 
         Caption         =   "MUTFG mutualisation des fonds de garantie"
         Height          =   255
         Left            =   -74880
         TabIndex        =   177
         Top             =   5700
         Width           =   4215
      End
      Begin VB.Label lblDUROM 
         Caption         =   "DUROM durée initial (IEOM)"
         Height          =   255
         Left            =   -74880
         TabIndex        =   176
         Top             =   6200
         Width           =   4215
      End
      Begin VB.Label lblDURIN 
         Caption         =   "DURIN durée initial"
         Height          =   255
         Left            =   -74880
         TabIndex        =   175
         Top             =   5700
         Width           =   4215
      End
      Begin VB.Label lblCDZON 
         Caption         =   "CDZON code résidence de contrepartie"
         Height          =   255
         Left            =   120
         TabIndex        =   174
         Top             =   6200
         Width           =   3975
      End
      Begin VB.Label lblTOPIF 
         Caption         =   "TOPIF type d'opération d'engagement sur IF à terme"
         Height          =   255
         Left            =   -74880
         TabIndex        =   173
         Top             =   6200
         Width           =   4095
      End
      Begin VB.Label lblTCFPE 
         Caption         =   "TCFPE titres constituant des fonds propres d'un ETS"
         Height          =   255
         Left            =   -74880
         TabIndex        =   172
         Top             =   5700
         Width           =   4095
      End
      Begin VB.Label lblSENOP 
         Caption         =   "SENOP sens de l'opération"
         Height          =   255
         Left            =   -74880
         TabIndex        =   171
         Top             =   5200
         Width           =   4095
      End
      Begin VB.Label lblSEMNT 
         Caption         =   "SEMNT sens du montant"
         Height          =   255
         Left            =   -74880
         TabIndex        =   170
         Top             =   4700
         Width           =   4095
      End
      Begin VB.Label lblRISPA 
         Caption         =   "RISPA code risque pays"
         Height          =   255
         Left            =   -74880
         TabIndex        =   169
         Top             =   4200
         Width           =   4095
      End
      Begin VB.Label lblREZON 
         Caption         =   "REZON code résidence de l'émetteur des titres"
         Height          =   255
         Left            =   -74880
         TabIndex        =   168
         Top             =   3700
         Width           =   4095
      End
      Begin VB.Label lblRESET 
         Caption         =   "RESET code résident de l'émetteur des titres"
         Height          =   255
         Left            =   -74880
         TabIndex        =   167
         Top             =   3200
         Width           =   4095
      End
      Begin VB.Label lblREDHB 
         Caption         =   "REDHB répartition des engagements douteux de HB"
         Height          =   255
         Left            =   -74880
         TabIndex        =   166
         Top             =   2700
         Width           =   4095
      End
      Begin VB.Label lblREDES 
         Caption         =   "REDES dotations ou reprises de provisions"
         Height          =   255
         Left            =   -74880
         TabIndex        =   165
         Top             =   2200
         Width           =   4095
      End
      Begin VB.Label lblPROCB 
         Caption         =   "PROCB produit sur OP de crédit-bail ""assimilées"""
         Height          =   255
         Left            =   -74880
         TabIndex        =   164
         Top             =   1700
         Width           =   4095
      End
      Begin VB.Label lblPRIMP 
         Caption         =   "PRIMP provisions ayant ou non supporté l'impôt"
         Height          =   255
         Left            =   -74880
         TabIndex        =   163
         Top             =   1200
         Width           =   4095
      End
      Begin VB.Label lblPAACT 
         Caption         =   "PAACT pays de résidence (ISO 3166-1993)"
         Height          =   255
         Left            =   -74880
         TabIndex        =   162
         Top             =   6200
         Width           =   4095
      End
      Begin VB.Label lblOPREC 
         Caption         =   "OPREC opérations réciproques"
         Height          =   255
         Left            =   -74880
         TabIndex        =   161
         Top             =   5700
         Width           =   4095
      End
      Begin VB.Label lblOPFDH 
         Caption         =   "OPFDH opérations de financement"
         Height          =   255
         Left            =   -74880
         TabIndex        =   160
         Top             =   5200
         Width           =   4095
      End
      Begin VB.Label lblOPEFC 
         Caption         =   "OPEFC opérations fermes et conditionnelles"
         Height          =   255
         Left            =   -74880
         TabIndex        =   159
         Top             =   4700
         Width           =   4095
      End
      Begin VB.Label lblOPCVM 
         Caption         =   "OPCVM opcvm géré par l'établissement"
         Height          =   255
         Left            =   -74880
         TabIndex        =   158
         Top             =   4200
         Width           =   4095
      End
      Begin VB.Label lblNREHB 
         Caption         =   "NREHB niveau de risque attaché à un engagement de H"
         Height          =   255
         Left            =   -74880
         TabIndex        =   157
         Top             =   3700
         Width           =   4095
      End
      Begin VB.Label lblNRAST 
         Caption         =   "NRAST niveau de risque attaché à un stock"
         Height          =   255
         Left            =   -74880
         TabIndex        =   156
         Top             =   3200
         Width           =   4095
      End
      Begin VB.Label lblNATRS 
         Caption         =   "NATRS nature de compte à régime spécial"
         Height          =   255
         Left            =   -74880
         TabIndex        =   155
         Top             =   2700
         Width           =   4095
      End
      Begin VB.Label lblNATOF 
         Caption         =   "NATOF nature de l'objet financé"
         Height          =   255
         Left            =   -74880
         TabIndex        =   154
         Top             =   2200
         Width           =   4095
      End
      Begin VB.Label lblNATMA 
         Caption         =   "NATMA nature du marché"
         Height          =   255
         Left            =   -74880
         TabIndex        =   153
         Top             =   1700
         Width           =   4095
      End
      Begin VB.Label lblNATIT 
         Caption         =   "NATIT nature des titres en portefeuille ""EMIS"""
         Height          =   255
         Left            =   -74880
         TabIndex        =   152
         Top             =   1200
         Width           =   4095
      End
      Begin VB.Label lblNATER 
         Caption         =   "NATER nature emplois ressources pour CD, provisions"
         Height          =   255
         Left            =   -74880
         TabIndex        =   151
         Top             =   6200
         Width           =   4095
      End
      Begin VB.Label lblNATDD 
         Caption         =   "NATDD nature des débiteurs divers"
         Height          =   255
         Left            =   -74880
         TabIndex        =   150
         Top             =   5700
         Width           =   4095
      End
      Begin VB.Label lblNATCS 
         Caption         =   "NATCS nature de comptes de stock"
         Height          =   255
         Left            =   -74880
         TabIndex        =   149
         Top             =   5200
         Width           =   4095
      End
      Begin VB.Label lblNATCR 
         Caption         =   "NATCR nature des crédits"
         Height          =   255
         Left            =   -74880
         TabIndex        =   148
         Top             =   4680
         Width           =   4095
      End
      Begin VB.Label lblNARCP 
         Caption         =   "NARCP zone géographique de rattachement de la CP"
         Height          =   255
         Left            =   -74880
         TabIndex        =   147
         Top             =   3700
         Width           =   4095
      End
      Begin VB.Label lblNAPRO 
         Caption         =   "NAPRO nature de provision crédit-bail"
         Height          =   255
         Left            =   -74880
         TabIndex        =   146
         Top             =   3200
         Width           =   4095
      End
      Begin VB.Label lblNAOCB 
         Caption         =   "NAOCB nature des OP de crédit-bail "" assimilés"""
         Height          =   255
         Left            =   -74880
         TabIndex        =   145
         Top             =   2700
         Width           =   4095
      End
      Begin VB.Label lblNAIMO 
         Caption         =   "NAIMO nature des immobilisations"
         Height          =   255
         Left            =   -74880
         TabIndex        =   144
         Top             =   2200
         Width           =   4095
      End
      Begin VB.Label lblNAEGA 
         Caption         =   "NAEGA zone géographique de l'établissement garant"
         Height          =   255
         Left            =   -74880
         TabIndex        =   143
         Top             =   1700
         Width           =   4095
      End
      Begin VB.Label lblNACPS 
         Caption         =   "NACPS nature des agents de la contrepartie"
         Height          =   255
         Left            =   -74880
         TabIndex        =   142
         Top             =   1200
         Width           =   4095
      End
      Begin VB.Label lblNACGR 
         Caption         =   "NACGR VM garantie ou non, émise ou non par état CEE"
         Height          =   255
         Left            =   -74880
         TabIndex        =   141
         Top             =   700
         Width           =   4095
      End
      Begin VB.Label lblMONDV 
         Caption         =   "MONDV code devise - monnaie nationale"
         Height          =   255
         Left            =   -74880
         TabIndex        =   140
         Top             =   5200
         Width           =   4215
      End
      Begin VB.Label lblMLFCE 
         Caption         =   "MLFCE monnaie de la créance ou de l'engagement"
         Height          =   255
         Left            =   -74880
         TabIndex        =   139
         Top             =   4700
         Width           =   4215
      End
      Begin VB.Label lblGARCF 
         Caption         =   "GARCF garantie coface"
         Height          =   255
         Left            =   -74880
         TabIndex        =   138
         Top             =   4200
         Width           =   4095
      End
      Begin VB.Label lblFPRBG 
         Caption         =   "FPRBG provisions suceptibles d'être incluses ou non"
         Height          =   255
         Left            =   -74880
         TabIndex        =   137
         Top             =   3700
         Width           =   4215
      End
      Begin VB.Label lblFOPIF 
         Caption         =   "FOPIF finalité opérations instruments financiers"
         Height          =   255
         Left            =   -74880
         TabIndex        =   136
         Top             =   3200
         Width           =   4095
      End
      Begin VB.Label lblFAMDV 
         Caption         =   "FAMDV famille de devise"
         Height          =   255
         Left            =   -74880
         TabIndex        =   135
         Top             =   2700
         Width           =   4215
      End
      Begin VB.Label lblELIGB 
         Caption         =   "ELIGB éligibilité"
         Height          =   255
         Left            =   -74880
         TabIndex        =   134
         Top             =   2200
         Width           =   4215
      End
      Begin VB.Label lblECFIN 
         Caption         =   "ECFIN type d'encours financiers"
         Height          =   255
         Left            =   -74880
         TabIndex        =   133
         Top             =   1700
         Width           =   4215
      End
      Begin VB.Label lblECART 
         Caption         =   "ECART provisions et écarts"
         Height          =   255
         Left            =   -74880
         TabIndex        =   132
         Top             =   1200
         Width           =   4215
      End
      Begin VB.Label lblDVOPR 
         Caption         =   "DVOPR devise de l'objet donnant lieu à provision"
         Height          =   255
         Left            =   -74880
         TabIndex        =   131
         Top             =   700
         Width           =   4215
      End
      Begin VB.Label lblCPEMS 
         Caption         =   "CPEMS caractère des prêts et emprunts subordonnés"
         Height          =   255
         Left            =   -74880
         TabIndex        =   130
         Top             =   1700
         Width           =   3975
      End
      Begin VB.Label lblCOTIT 
         Caption         =   "COTIT admission à la côte d'une bourse de valeurs"
         Height          =   255
         Left            =   -74880
         TabIndex        =   129
         Top             =   1200
         Width           =   3975
      End
      Begin VB.Label lblCLCRC 
         Caption         =   "CLCRC crédit lié à des créances commerciales"
         Height          =   255
         Left            =   -74880
         TabIndex        =   128
         Top             =   700
         Width           =   4095
      End
      Begin VB.Label lblDRAC 
         Caption         =   "DRAC durée restant à courir"
         Height          =   255
         Left            =   -74880
         TabIndex        =   127
         Top             =   5200
         Width           =   3975
      End
      Begin VB.Label lblDCTOM 
         Caption         =   "DCTOM dépôt collecté (épargne) dans les tombées"
         Height          =   255
         Left            =   -74880
         TabIndex        =   126
         Top             =   4700
         Width           =   3975
      End
      Begin VB.Label lblCRHYP 
         Caption         =   "CRHYP crédit hypotécaire"
         Height          =   255
         Left            =   -74880
         TabIndex        =   125
         Top             =   4200
         Width           =   3975
      End
      Begin VB.Label lblCRETC 
         Caption         =   "CRETC créances/dettes à vue ou à terme"
         Height          =   255
         Left            =   -74880
         TabIndex        =   124
         Top             =   3700
         Width           =   3975
      End
      Begin VB.Label lblCREOR 
         Caption         =   "CREOR nature des créances origines"
         Height          =   255
         Left            =   -74880
         TabIndex        =   123
         Top             =   3200
         Width           =   3975
      End
      Begin VB.Label lblCREIM 
         Caption         =   "CREIM créances impayées"
         Height          =   255
         Left            =   -74880
         TabIndex        =   122
         Top             =   2700
         Width           =   3975
      End
      Begin VB.Label lblCRDIV 
         Caption         =   "CRDIV créditeurs divers (Famille)"
         Height          =   255
         Left            =   -74880
         TabIndex        =   121
         Top             =   2200
         Width           =   3855
      End
      Begin VB.Label lblCDRES 
         Caption         =   "CDRES code résident de contrepartie"
         Height          =   255
         Left            =   120
         TabIndex        =   120
         Top             =   5700
         Width           =   4095
      End
      Begin VB.Label lblCATET 
         Caption         =   "CATET catégorie d'établissement (organe central)"
         Height          =   255
         Left            =   120
         TabIndex        =   119
         Top             =   5200
         Width           =   4215
      End
      Begin VB.Label lblCAROB 
         Caption         =   "CAROB caractèristiques de l'émission des titres"
         Height          =   255
         Left            =   120
         TabIndex        =   118
         Top             =   4700
         Width           =   4215
      End
      Begin VB.Label lblBONIF 
         Caption         =   "BONIF bonification"
         Height          =   255
         Left            =   120
         TabIndex        =   117
         Top             =   4200
         Width           =   4215
      End
      Begin VB.Label lblAUTDV 
         Caption         =   "AUTDV gestion des devises (état 4006)"
         Height          =   255
         Left            =   120
         TabIndex        =   116
         Top             =   3700
         Width           =   4215
      End
      Begin VB.Label lblATTCF 
         Caption         =   "ATTCF comptabilité financière"
         Height          =   255
         Left            =   120
         TabIndex        =   115
         Top             =   3200
         Width           =   4215
      End
      Begin VB.Label lblAREFR 
         Caption         =   "AREFR accords de refinancement donnés ou reçus"
         Height          =   255
         Left            =   120
         TabIndex        =   114
         Top             =   2700
         Width           =   4215
      End
      Begin VB.Label lblAPPAR 
         Caption         =   "APPAR apparentement"
         Height          =   255
         Left            =   120
         TabIndex        =   113
         Top             =   2200
         Width           =   4215
      End
      Begin VB.Label lblAGENT 
         Caption         =   "AGENT agent économique de contre partie"
         Height          =   255
         Left            =   120
         TabIndex        =   112
         Top             =   1700
         Width           =   4095
      End
      Begin VB.Label lblAGEMT 
         Caption         =   "AGEMT agent économique émetteur des titres"
         Height          =   255
         Left            =   120
         TabIndex        =   107
         Top             =   1200
         Width           =   4215
      End
      Begin VB.Label lblAFFPU 
         Caption         =   "AFFPU affectation des fonds publics"
         Height          =   255
         Left            =   120
         TabIndex        =   106
         Top             =   700
         Width           =   4215
      End
   End
   Begin VB.Label libRéférence 
      Alignment       =   1  'Right Justify
      Caption         =   "Nature"
      Height          =   255
      Left            =   2280
      TabIndex        =   188
      Top             =   45
      Width           =   3255
   End
End
Attribute VB_Name = "frmLrAttributDétail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------
Option Explicit
'---------------------------------------------------------
Dim arrTag() As Boolean, arrTagNb As Integer, lstErrClear As Boolean
Dim lastActiveControl_Name  As String, currentActiveControl_Name As String
Dim blnMsgBox_Quit As Boolean, autLrAttribut As typeAuthorization

Dim recLrAttribut As typeLrAttribut
Dim currentMethod As String
Dim txtRéférence_Format As String

Dim recCompte As typeCompte


Private Sub cboaffpu_GotFocus()
lblAFFPU.ForeColor = warnUsrColor
End Sub

Private Sub cboaffpu_lostFocus()
lblAFFPU.ForeColor = lblUsr.ForeColor
End Sub


Private Sub cboAGEMT_GotFocus()
lblAGEMT.ForeColor = warnUsrColor
End Sub
Private Sub cboAGEMT_lostFocus()
lblAGEMT.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboAGENT_GotFocus()
lblAGENT.ForeColor = warnUsrColor
End Sub
Private Sub cboAGENT_lostFocus()
lblAGENT.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboAPPAR_GotFocus()
lblAPPAR.ForeColor = warnUsrColor
End Sub
Private Sub cboAPPAR_lostFocus()
lblAPPAR.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboAREFR_GotFocus()
lblAREFR.ForeColor = warnUsrColor
End Sub
Private Sub cboAREFR_lostFocus()
lblAREFR.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboATTCF_GotFocus()
lblATTCF.ForeColor = warnUsrColor
End Sub
Private Sub cboATTCF_lostFocus()
lblATTCF.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboAUTDV_GotFocus()
lblAUTDV.ForeColor = warnUsrColor
End Sub
Private Sub cboAUTDV_lostFocus()
lblAUTDV.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboBONIF_GotFocus()
lblBONIF.ForeColor = warnUsrColor
End Sub
Private Sub cboBONIF_lostFocus()
lblBONIF.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboCAROB_GotFocus()
lblCAROB.ForeColor = warnUsrColor
End Sub
Private Sub cboCAROB_lostFocus()
lblCAROB.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboCATET_GotFocus()
lblCATET.ForeColor = warnUsrColor
End Sub
Private Sub cboCATET_lostFocus()
lblCATET.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboCDRES_GotFocus()
lblCDRES.ForeColor = warnUsrColor
End Sub
Private Sub cboCDRES_lostFocus()
lblCDRES.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboCDZON_GotFocus()
lblCDZON.ForeColor = warnUsrColor
End Sub
Private Sub cboCDZON_lostFocus()
lblCDZON.ForeColor = lblUsr.ForeColor
SSTab1.Tab = SSTab1.Tab + 1
End Sub

Private Sub cboCLCRC_GotFocus()
lblCLCRC.ForeColor = warnUsrColor
End Sub
Private Sub cboCLCRC_lostFocus()
lblCLCRC.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboCOTIT_GotFocus()
lblCOTIT.ForeColor = warnUsrColor
End Sub
Private Sub cboCOTIT_lostFocus()
lblCOTIT.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboCPEMS_GotFocus()
lblCPEMS.ForeColor = warnUsrColor
End Sub
Private Sub cboCPEMS_lostFocus()
lblCPEMS.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboCRDIV_GotFocus()
lblCRDIV.ForeColor = warnUsrColor
End Sub
Private Sub cboCRDIV_lostFocus()
lblCRDIV.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboCREIM_GotFocus()
lblCREIM.ForeColor = warnUsrColor
End Sub
Private Sub cboCREIM_lostFocus()
lblCREIM.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboCREOR_GotFocus()
lblCREOR.ForeColor = warnUsrColor
End Sub
Private Sub cboCREOR_lostFocus()
lblCREOR.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboCRETC_GotFocus()
lblCRETC.ForeColor = warnUsrColor
End Sub
Private Sub cboCRETC_lostFocus()
lblCRETC.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboCRHYP_GotFocus()
lblCRHYP.ForeColor = warnUsrColor
End Sub
Private Sub cboCRHYP_lostFocus()
lblCRHYP.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboDCTOM_GotFocus()
lblDCTOM.ForeColor = warnUsrColor
End Sub

Private Sub cboDCTOM_lostFocus()
lblDCTOM.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboDRAC_GotFocus()
lblDRAC.ForeColor = warnUsrColor
End Sub
Private Sub cboDRAC_lostFocus()
lblDRAC.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboDURIN_GotFocus()
lblDURIN.ForeColor = warnUsrColor
End Sub
Private Sub cboDURIN_lostFocus()
lblDURIN.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboDUROM_GotFocus()
lblDUROM.ForeColor = warnUsrColor
End Sub
Private Sub cboDUROM_lostFocus()
lblDUROM.ForeColor = lblUsr.ForeColor
SSTab1.Tab = SSTab1.Tab + 1
End Sub

Private Sub cboDVOPR_GotFocus()
lblDVOPR.ForeColor = warnUsrColor
End Sub
Private Sub cboDVOPR_lostFocus()
lblDVOPR.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboECART_GotFocus()
lblECART.ForeColor = warnUsrColor
End Sub
Private Sub cboECART_lostFocus()
lblECART.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboECFIN_GotFocus()
lblECFIN.ForeColor = warnUsrColor
End Sub
Private Sub cboECFIN_lostFocus()
lblECFIN.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboELIGB_GotFocus()
lblELIGB.ForeColor = warnUsrColor
End Sub
Private Sub cboELIGB_lostFocus()
lblELIGB.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboFAMDV_GotFocus()
lblFAMDV.ForeColor = warnUsrColor
End Sub
Private Sub cboFAMDV_lostFocus()
lblFAMDV.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboFOPIF_GotFocus()
lblFOPIF.ForeColor = warnUsrColor
End Sub
Private Sub cboFOPIF_lostFocus()
lblFOPIF.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboFPRBG_GotFocus()
lblFPRBG.ForeColor = warnUsrColor
End Sub
Private Sub cboFPRBG_lostFocus()
lblFPRBG.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboGARCF_GotFocus()
lblGARCF.ForeColor = warnUsrColor
End Sub
Private Sub cboGARCF_lostFocus()
lblGARCF.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboMLFCE_GotFocus()
lblMLFCE.ForeColor = warnUsrColor
End Sub
Private Sub cboMLFCE_lostFocus()
lblMLFCE.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboMONDV_GotFocus()
lblMONDV.ForeColor = warnUsrColor
End Sub
Private Sub cboMONDV_lostFocus()
lblMONDV.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboMUTFG_GotFocus()
lblMUTFG.ForeColor = warnUsrColor
End Sub
Private Sub cboMUTFG_lostFocus()
lblMUTFG.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboNACGA_GotFocus()
lblNACGA.ForeColor = warnUsrColor
End Sub
Private Sub cboNACGA_lostFocus()
lblNACGA.ForeColor = lblUsr.ForeColor
SSTab1.Tab = SSTab1.Tab + 1
End Sub

Private Sub cboNACGR_GotFocus()
lblNACGR.ForeColor = warnUsrColor
End Sub
Private Sub cboNACGR_lostFocus()
lblNACGR.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboNACPS_GotFocus()
lblNACPS.ForeColor = warnUsrColor
End Sub
Private Sub cboNACPS_lostFocus()
lblNACPS.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboNAEGA_GotFocus()
lblNAEGA.ForeColor = warnUsrColor
End Sub
Private Sub cboNAEGA_lostFocus()
lblNAEGA.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboNAIMO_GotFocus()
lblNAIMO.ForeColor = warnUsrColor
End Sub
Private Sub cboNAIMO_lostFocus()
lblNAIMO.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboNAOCB_GotFocus()
lblNAOCB.ForeColor = warnUsrColor
End Sub
Private Sub cboNAOCB_lostFocus()
lblNAOCB.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboNAPRO_GotFocus()
lblNAPRO.ForeColor = warnUsrColor
End Sub
Private Sub cboNAPRO_lostFocus()
lblNAPRO.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboNARCP_GotFocus()
lblNARCP.ForeColor = warnUsrColor
End Sub
Private Sub cboNARCP_lostFocus()
lblNARCP.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboNATCP_GotFocus()
lblNATCP.ForeColor = warnUsrColor
End Sub
Private Sub cboNATCP_lostFocus()
lblNATCP.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboNATCR_GotFocus()
lblNATCR.ForeColor = warnUsrColor
End Sub
Private Sub cboNATCR_lostFocus()
lblNATCR.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboNATCS_GotFocus()
lblNATCS.ForeColor = warnUsrColor
End Sub
Private Sub cboNATCS_lostFocus()
lblNATCS.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboNATDD_GotFocus()
lblNATDD.ForeColor = warnUsrColor
End Sub
Private Sub cboNATDD_lostFocus()
lblNATDD.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboNATER_GotFocus()
lblNATER.ForeColor = warnUsrColor
End Sub
Private Sub cboNATER_lostFocus()
lblNATER.ForeColor = lblUsr.ForeColor
SSTab1.Tab = SSTab1.Tab + 1
End Sub

Private Sub cboNATIF_GotFocus()
lblNATIF.ForeColor = warnUsrColor
End Sub
Private Sub cboNATIF_lostFocus()
lblNATIF.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboNATIT_GotFocus()
lblNATIT.ForeColor = warnUsrColor
End Sub
Private Sub cboNATIT_lostFocus()
lblNATIT.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboNATMA_GotFocus()
lblNATMA.ForeColor = warnUsrColor
End Sub
Private Sub cboNATMA_lostFocus()
lblNATMA.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboNATOF_GotFocus()
lblNATOF.ForeColor = warnUsrColor
End Sub
Private Sub cboNATOF_lostFocus()
lblNATOF.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboNATRS_GotFocus()
lblNATRS.ForeColor = warnUsrColor
End Sub
Private Sub cboNATRS_lostFocus()
lblNATRS.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboNRAST_GotFocus()
lblNRAST.ForeColor = warnUsrColor
End Sub
Private Sub cboNRAST_lostFocus()
lblNRAST.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboNREHB_GotFocus()
lblNREHB.ForeColor = warnUsrColor
End Sub
Private Sub cboNREHB_lostFocus()
lblNREHB.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboOPCVM_GotFocus()
lblOPCVM.ForeColor = warnUsrColor
End Sub
Private Sub cboOPCVM_lostFocus()
lblOPCVM.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboOPEFC_GotFocus()
lblOPEFC.ForeColor = warnUsrColor
End Sub
Private Sub cboOPEFC_lostFocus()
lblOPEFC.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboOPFDH_GotFocus()
lblOPFDH.ForeColor = warnUsrColor
End Sub
Private Sub cboOPFDH_lostFocus()
lblOPFDH.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboOPREC_GotFocus()
lblOPREC.ForeColor = warnUsrColor
End Sub
Private Sub cboOPREC_lostFocus()
lblOPREC.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboPAACT_GotFocus()
lblPAACT.ForeColor = warnUsrColor
End Sub
Private Sub cboPAACT_lostFocus()
lblPAACT.ForeColor = lblUsr.ForeColor
SSTab1.Tab = SSTab1.Tab + 1
End Sub


Private Sub cboPERIO_GotFocus()
lblPERIO.ForeColor = warnUsrColor
End Sub
Private Sub cboPERIO_lostFocus()
lblPERIO.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboPRIMP_GotFocus()
lblPRIMP.ForeColor = warnUsrColor
End Sub
Private Sub cboPRIMP_lostFocus()
lblPRIMP.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboPROCB_GotFocus()
lblPROCB.ForeColor = warnUsrColor
End Sub
Private Sub cboPROCB_lostFocus()
lblPROCB.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboREDES_GotFocus()
lblREDES.ForeColor = warnUsrColor
End Sub
Private Sub cboREDES_lostFocus()
lblREDES.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboREDHB_GotFocus()
lblREDHB.ForeColor = warnUsrColor
End Sub
Private Sub cboREDHB_lostFocus()
lblREDHB.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboRESET_GotFocus()
lblRESET.ForeColor = warnUsrColor
End Sub
Private Sub cboRESET_lostFocus()
lblRESET.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboREZON_GotFocus()
lblREZON.ForeColor = warnUsrColor
End Sub
Private Sub cboREZON_lostFocus()
lblREZON.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboRISPA_GotFocus()
lblRISPA.ForeColor = warnUsrColor
End Sub
Private Sub cboRISPA_lostFocus()
lblRISPA.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboSEMNT_GotFocus()
lblSEMNT.ForeColor = warnUsrColor
End Sub
Private Sub cboSEMNT_lostFocus()
lblSEMNT.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboSENOP_GotFocus()
lblSENOP.ForeColor = warnUsrColor
End Sub
Private Sub cboSENOP_lostFocus()
lblSENOP.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboTCFPE_GotFocus()
lblTCFPE.ForeColor = warnUsrColor
End Sub
Private Sub cboTCFPE_lostFocus()
lblTCFPE.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboTOPIF_GotFocus()
lblTOPIF.ForeColor = warnUsrColor
End Sub
Private Sub cboTOPIF_lostFocus()
lblTOPIF.ForeColor = lblUsr.ForeColor
SSTab1.Tab = SSTab1.Tab + 1
End Sub

Private Sub cboTYCGR_GotFocus()
lblTYCGR.ForeColor = warnUsrColor
End Sub
Private Sub cboTYCGR_lostFocus()
lblTYCGR.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboTYCOM_GotFocus()
lblTYCOM.ForeColor = warnUsrColor
End Sub
Private Sub cboTYCOM_lostFocus()
lblTYCOM.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboTYDSU_GotFocus()
lblTYDSU.ForeColor = warnUsrColor
End Sub
Private Sub cboTYDSU_lostFocus()
lblTYDSU.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboTYETS_GotFocus()
lblTYETS.ForeColor = warnUsrColor
End Sub
Private Sub cboTYETS_lostFocus()
lblTYETS.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboTYPOR_GotFocus()
lblTYPOR.ForeColor = warnUsrColor
End Sub
Private Sub cboTYPOR_lostFocus()
lblTYPOR.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboTYPSU_GotFocus()
lblTYPSU.ForeColor = warnUsrColor
End Sub
Private Sub cboTYPSU_lostFocus()
lblTYPSU.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboTYRES_GotFocus()
lblTYRES.ForeColor = warnUsrColor
End Sub
Private Sub cboTYRES_lostFocus()
lblTYRES.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboZACTI_GotFocus()
lblZACTI.ForeColor = warnUsrColor
End Sub
Private Sub cboZACTI_lostFocus()
lblZACTI.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboCDCPCO_GotFocus()
lblCDCPCO.ForeColor = warnUsrColor
End Sub
Private Sub cboCDCPCO_lostFocus()
lblCDCPCO.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboCDCPJO_GotFocus()
lblCDCPJO.ForeColor = warnUsrColor
End Sub
Private Sub cboCDCPJO_lostFocus()
lblCDCPJO.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboCDCPFU_GotFocus()
lblCDCPFU.ForeColor = warnUsrColor
End Sub

Private Sub cboREESC1_GotFocus()
lblREESC1.ForeColor = warnUsrColor
End Sub

Private Sub cboREESC6_GotFocus()
lblREESC6.ForeColor = warnUsrColor
End Sub

Private Sub cboCDAGCO_GotFocus()
lblCDAGCO.ForeColor = warnUsrColor
End Sub
Private Sub cboCDAGCO_lostFocus()
lblCDAGCO.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboCDREME_GotFocus()
lblCDREME.ForeColor = warnUsrColor
End Sub
Private Sub cboCDREME_lostFocus()
lblCDREME.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboTYMTDV_GotFocus()
lblTYMTDV.ForeColor = warnUsrColor
End Sub
Private Sub cboTYMTDV_lostFocus()
lblTYMTDV.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboTYVENT_GotFocus()
lblTYVENT.ForeColor = warnUsrColor
End Sub
Private Sub cboTYVENT_lostFocus()
lblTYVENT.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboCRVENT_GotFocus()
lblCRVENT.ForeColor = warnUsrColor
End Sub
Private Sub cboCDDURE_GotFocus()
lblCDDURE.ForeColor = warnUsrColor
End Sub
Private Sub cboCDDURE_lostFocus()
lblCDDURE.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboDUINIT_GotFocus()
lblDUINIT.ForeColor = warnUsrColor
End Sub
Private Sub cboCDCRTI_GotFocus()
lblCDCRTI.ForeColor = warnUsrColor
End Sub
Private Sub cboCDCRTI_lostFocus()
lblCDCRTI.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboCDCRAC_GotFocus()
lblCDCRAC.ForeColor = warnUsrColor
End Sub
Private Sub cboCDCRAC_lostFocus()
lblCDCRAC.ForeColor = lblUsr.ForeColor
SSTab1.Tab = SSTab1.Tab + 1
End Sub

Private Sub cboCDBIOR_GotFocus()
lblCDBIOR.ForeColor = warnUsrColor
End Sub
Private Sub cboCDBIOR_lostFocus()
lblCDBIOR.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboCDDEIN_GotFocus()
lblCDDEIN.ForeColor = warnUsrColor
End Sub
Private Sub cboCDDEIN_lostFocus()
lblCDDEIN.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboCDCRIM_GotFocus()
lblCDCRIM.ForeColor = warnUsrColor
End Sub
Private Sub cboCDCRIM_lostFocus()
lblCDCRIM.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboCDCRCO_GotFocus()
lblCDCRCO.ForeColor = warnUsrColor
End Sub
Private Sub cboCDCRCO_lostFocus()
lblCDCRCO.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboCDCREF_GotFocus()
lblCDCREF.ForeColor = warnUsrColor
End Sub
Private Sub cboCDCREF_lostFocus()
lblCDCREF.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboCDLODA_GotFocus()
lblCDLODA.ForeColor = warnUsrColor
End Sub
Private Sub cboCDLODA_lostFocus()
lblCDLODA.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboCDCRET_GotFocus()
lblCDCRET.ForeColor = warnUsrColor
End Sub
Private Sub cboCDCRET_lostFocus()
lblCDCRET.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboCDOMPO_GotFocus()
lblCDOMPO.ForeColor = warnUsrColor
End Sub
Private Sub cboCDOMPO_lostFocus()
lblCDOMPO.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboCDOPIM_GotFocus()
lblCDOPIM.ForeColor = warnUsrColor
End Sub
Private Sub cboCDOPIM_lostFocus()
lblCDOPIM.ForeColor = lblUsr.ForeColor
End Sub

Private Sub cboCDSWAP_GotFocus()
lblCDSWAP.ForeColor = warnUsrColor
End Sub
Private Sub cboCDSWAP_lostFocus()
lblCDSWAP.ForeColor = lblUsr.ForeColor
SSTab1.Tab = 0
End Sub

Private Sub cboAFFPU_Click()
cbo_Value recLrAttribut.AFFPU, cboAFFPU
End Sub

Private Sub cboAGEMT_Click()
cbo_Value recLrAttribut.AGEMT, cboAGEMT
End Sub


Private Sub cboAGENT_Click()
cbo_Value recLrAttribut.AGENT, cboAGENT
End Sub


Private Sub cboAPPAR_Click()
cbo_Value recLrAttribut.APPAR, cboAPPAR
End Sub


Private Sub cboAREFR_Click()
cbo_Value recLrAttribut.AREFR, cboAREFR
End Sub


Private Sub cboATTCF_Click()
cbo_Value recLrAttribut.ATTCF, cboATTCF
End Sub


Private Sub cboAUTDV_Click()
cbo_Value recLrAttribut.AUTDV, cboAUTDV
End Sub


Private Sub cboBONIF_Click()
cbo_Value recLrAttribut.BONIF, cboBONIF
End Sub


Private Sub cboCAROB_Click()
cbo_Value recLrAttribut.CAROB, cboCAROB
End Sub


Private Sub cboCATET_Click()
cbo_Value recLrAttribut.CATET, cboCATET
End Sub


Private Sub cboCDCPCO_Click()
cbo_Value recLrAttribut.CDCPCO, cboCDCPCO
End Sub

Private Sub cboCDCPFU_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub

Private Sub cboREESC1_KeyPress(KeyAscii As Integer)
num_KeyAscii KeyAscii
End Sub

Private Sub cboREESC6_KeyPress(KeyAscii As Integer)
num_KeyAscii KeyAscii
End Sub

Private Sub cboCDCPFU_lostFocus()
lblCDCPFU.ForeColor = lblUsr.ForeColor
recLrAttribut.CDCPFU = Trim(cboCDCPFU)
End Sub

Private Sub cboREESC1_lostFocus()
Dim X As String
lblREESC1.ForeColor = lblUsr.ForeColor
recLrAttribut.REESC1 = Trim(cboREESC1)
If Trim(recLrAttribut.REESC1) <> "" And recLrAttribut.REESC1 <> X Then
    If Not IsNull(PCI_Check(recLrAttribut.REESC1)) Then cboREESC1.SetFocus
End If
End Sub

Private Sub cboREESC6_lostFocus()
Dim X As String
X = recLrAttribut.REESC6
lblREESC6.ForeColor = lblUsr.ForeColor
recLrAttribut.REESC6 = Trim(cboREESC6)
If Trim(recLrAttribut.REESC6) <> "" And recLrAttribut.REESC6 <> X Then
    If Not IsNull(PCI_Check(recLrAttribut.REESC6)) Then cboREESC6.SetFocus: Exit Sub
End If
SSTab1.Tab = SSTab1.Tab + 1
End Sub

Private Sub cboCDcpjo_Click()
cbo_Value recLrAttribut.CDCPJO, cboCDCPJO
End Sub
Private Sub cboCDcpfu_Click()
cbo_Value recLrAttribut.CDCPFU, cboCDCPFU
End Sub
Private Sub cboREESC1_Click()
cbo_Value recLrAttribut.REESC1, cboREESC1
End Sub
Private Sub cboREESC6_Click()
cbo_Value recLrAttribut.REESC6, cboREESC6
End Sub

Private Sub cboCDagco_Click()
cbo_Value recLrAttribut.CDAGCO, cboCDAGCO
End Sub

Private Sub cboCDreme_Click()
cbo_Value recLrAttribut.CDREME, cboCDREME
End Sub

Private Sub cboCRVENT_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub

Private Sub cboCRVENT_lostFocus()
lblCRVENT.ForeColor = lblUsr.ForeColor
recLrAttribut.CRVENT = Trim(cboCRVENT)
End Sub

Private Sub cboDUINIT_KeyPress(KeyAscii As Integer)
KeyAscii = convUCase(KeyAscii)
End Sub

Private Sub cboDUINIT_lostFocus()
lblDUINIT.ForeColor = lblUsr.ForeColor
recLrAttribut.DUINIT = Trim(cboDUINIT)
End Sub

Private Sub cbotymtdv_Click()
cbo_Value recLrAttribut.TYMTDV, cboTYMTDV
End Sub
Private Sub cbotyvent_Click()
cbo_Value recLrAttribut.TYVENT, cboTYVENT
End Sub
Private Sub cbocrvent_Click()
cbo_Value recLrAttribut.CRVENT, cboCRVENT
End Sub
Private Sub cboCDdure_Click()
cbo_Value recLrAttribut.CDDURE, cboCDDURE
End Sub
Private Sub cboduinit_Click()
cbo_Value recLrAttribut.DUINIT, cboDUINIT
End Sub
Private Sub cboCDcrti_Click()
cbo_Value recLrAttribut.CDCRTI, cboCDCRTI
End Sub
Private Sub cboCdcrac_Click()
cbo_Value recLrAttribut.CDCRAC, cboCDCRAC
End Sub
Private Sub cboCDbior_Click()
cbo_Value recLrAttribut.CDBIOR, cboCDBIOR
End Sub
Private Sub cboCDdein_Click()
cbo_Value recLrAttribut.CDDEIN, cboCDDEIN
End Sub
Private Sub cboCDcrim_Click()
cbo_Value recLrAttribut.CDCRIM, cboCDCRIM
End Sub
Private Sub cboCDcrco_Click()
cbo_Value recLrAttribut.CDCRCO, cboCDCRCO
End Sub
Private Sub cboCDcref_Click()
cbo_Value recLrAttribut.CDCREF, cboCDCREF
End Sub
Private Sub cboCDloda_Click()
cbo_Value recLrAttribut.CDLODA, cboCDLODA
End Sub
Private Sub cboCDcret_Click()
cbo_Value recLrAttribut.CDCRET, cboCDCRET
End Sub
Private Sub cboCDompo_Click()
cbo_Value recLrAttribut.CDOMPO, cboCDOMPO
End Sub
Private Sub cboCDopim_Click()
cbo_Value recLrAttribut.CDOPIM, cboCDOPIM
End Sub
Private Sub cboCDswap_Click()
cbo_Value recLrAttribut.CDSWAP, cboCDSWAP

End Sub


Private Sub cboCDRES_Click()
cbo_Value recLrAttribut.CDRES, cboCDRES
End Sub


Private Sub cboCDZON_Click()
cbo_Value recLrAttribut.CDZON, cboCDZON
End Sub


Private Sub cboCLCRC_Click()
cbo_Value recLrAttribut.CLCRC, cboCLCRC
End Sub


Private Sub cboCOTIT_Click()
cbo_Value recLrAttribut.COTIT, cboCOTIT
End Sub


Private Sub cboCPEMS_Click()
cbo_Value recLrAttribut.CPEMS, cboCPEMS
End Sub


Private Sub cboCRDIV_Click()
cbo_Value recLrAttribut.CRDIV, cboCRDIV
End Sub


Private Sub cboCREIM_Click()
cbo_Value recLrAttribut.CREIM, cboCREIM
End Sub


Private Sub cboCREOR_Click()
cbo_Value recLrAttribut.CREOR, cboCREOR
End Sub


Private Sub cboCRETC_Click()
cbo_Value recLrAttribut.CRETC, cboCRETC
End Sub


Private Sub cboCRHYP_Click()
cbo_Value recLrAttribut.CRHYP, cboCRHYP
End Sub


Private Sub cboDCTOM_Click()
cbo_Value recLrAttribut.DCTOM, cboDCTOM
End Sub


Private Sub cboDRAC_Click()
cbo_Value recLrAttribut.DRAC, cboDRAC
End Sub


Private Sub cboDURIN_Click()
cbo_Value recLrAttribut.DURIN, cboDURIN
End Sub


Private Sub cboDUROM_Click()
cbo_Value recLrAttribut.DUROM, cboDUROM
End Sub


Private Sub cboDVOPR_Click()
cbo_Value recLrAttribut.DVOPR, cboDVOPR
End Sub


Private Sub cboECART_Click()
cbo_Value recLrAttribut.ECART, cboECART
End Sub


Private Sub cboECFIN_Click()
cbo_Value recLrAttribut.ECFIN, cboECFIN
End Sub


Private Sub cboELIGB_Click()
cbo_Value recLrAttribut.ELIGB, cboELIGB
End Sub


Private Sub cboFAMDV_Click()
cbo_Value recLrAttribut.FAMDV, cboFAMDV
End Sub


Private Sub cboFOPIF_Click()
cbo_Value recLrAttribut.FOPIF, cboFOPIF
End Sub


Private Sub cboFPRBG_Click()
cbo_Value recLrAttribut.FPRBG, cboFPRBG
End Sub


Private Sub cboGARCF_Click()
cbo_Value recLrAttribut.GARCF, cboGARCF
End Sub


Private Sub cboMLFCE_Click()
cbo_Value recLrAttribut.MLFCE, cboMLFCE
End Sub


Private Sub cboMONDV_Click()
cbo_Value recLrAttribut.MONDV, cboMONDV
End Sub


Private Sub cboMUTFG_Click()
cbo_Value recLrAttribut.MUTFG, cboMUTFG
End Sub


Private Sub cboNACGA_Click()
cbo_Value recLrAttribut.NACGA, cboNACGA
End Sub


Private Sub cboNACGR_Click()
cbo_Value recLrAttribut.NACGR, cboNACGR
End Sub


Private Sub cboNACPS_Click()
cbo_Value recLrAttribut.NACPS, cboNACPS
End Sub


Private Sub cboNAEGA_Click()
cbo_Value recLrAttribut.NAEGA, cboNAEGA
End Sub


Private Sub cboNAIMO_Click()
cbo_Value recLrAttribut.NAIMO, cboNAIMO
End Sub


Private Sub cboNAOCB_Click()
cbo_Value recLrAttribut.NAOCB, cboNAOCB
End Sub


Private Sub cboNAPRO_Click()
cbo_Value recLrAttribut.NAPRO, cboNAPRO
End Sub


Private Sub cboNARCP_Click()
cbo_Value recLrAttribut.NARCP, cboNARCP
End Sub


Private Sub cboNATCP_Click()
cbo_Value recLrAttribut.NATCP, cboNATCP
End Sub


Private Sub cboNATCR_Click()
cbo_Value recLrAttribut.NATCR, cboNATCR
End Sub


Private Sub cboNATCS_Click()
cbo_Value recLrAttribut.NATCS, cboNATCS
End Sub


Private Sub cboNATDD_Click()
cbo_Value recLrAttribut.NATDD, cboNATDD
End Sub


Private Sub cboNATER_Click()
cbo_Value recLrAttribut.NATER, cboNATER
End Sub


Private Sub cboNATIF_Click()
cbo_Value recLrAttribut.NATIF, cboNATIF
End Sub


Private Sub cboNATIT_Click()
cbo_Value recLrAttribut.NATIT, cboNATIT
End Sub


Private Sub cboNATMA_Click()
cbo_Value recLrAttribut.NATMA, cboNATMA
End Sub


Private Sub cboNATOF_Click()
cbo_Value recLrAttribut.NATOF, cboNATOF
End Sub


Private Sub cboNATRS_Click()
cbo_Value recLrAttribut.NATRS, cboNATRS
End Sub


Private Sub cboNRAST_Click()
cbo_Value recLrAttribut.NRAST, cboNRAST
End Sub


Private Sub cboNREHB_Click()
cbo_Value recLrAttribut.NREHB, cboNREHB
End Sub


Private Sub cboOPCVM_Click()
cbo_Value recLrAttribut.OPCVM, cboOPCVM
End Sub


Private Sub cboOPEFC_Click()
cbo_Value recLrAttribut.OPEFC, cboOPEFC
End Sub


Private Sub cboOPFDH_Click()
cbo_Value recLrAttribut.OPFDH, cboOPFDH
End Sub


Private Sub cboOPREC_Click()
cbo_Value recLrAttribut.OPREC, cboOPREC
End Sub


Private Sub cboPAACT_Click()
cbo_Value recLrAttribut.PAACT, cboPAACT
End Sub


Private Sub cboPERIO_Click()
cbo_Value recLrAttribut.PERIO, cboPERIO
End Sub


Private Sub cboPRIMP_Click()
cbo_Value recLrAttribut.PRIMP, cboPRIMP
End Sub


Private Sub cboPROCB_Click()
cbo_Value recLrAttribut.PROCB, cboPROCB
End Sub


Private Sub cboREDES_Click()
cbo_Value recLrAttribut.REDES, cboREDES
End Sub


Private Sub cboREDHB_Click()
cbo_Value recLrAttribut.REDHB, cboREDHB
End Sub


Private Sub cboRESET_Click()
cbo_Value recLrAttribut.RESET, cboRESET
End Sub


Private Sub cboREZON_Click()
cbo_Value recLrAttribut.REZON, cboREZON
End Sub


Private Sub cboRISPA_Click()
cbo_Value recLrAttribut.RISPA, cboRISPA
End Sub


Private Sub cboSEMNT_Click()
'cbo_Value recLrAttribut.SEMNT, cboSEMNT
End Sub


Private Sub cboSENOP_Click()
cbo_Value recLrAttribut.SENOP, cboSENOP
End Sub


Private Sub cboTCFPE_Click()
cbo_Value recLrAttribut.TCFPE, cboTCFPE
End Sub


Private Sub cboTOPIF_Click()
cbo_Value recLrAttribut.TOPIF, cboTOPIF
End Sub


Private Sub cboTYCGR_Click()
cbo_Value recLrAttribut.TYCGR, cboTYCGR
End Sub


Private Sub cboTYCOM_Click()
cbo_Value recLrAttribut.TYCOM, cboTYCOM
End Sub


Private Sub cboTYDSU_Click()
cbo_Value recLrAttribut.TYDSU, cboTYDSU
End Sub


Private Sub cboTYETS_Click()
cbo_Value recLrAttribut.TYETS, cboTYETS
End Sub


Private Sub cboTYPOR_Click()
cbo_Value recLrAttribut.TYPOR, cboTYPOR
End Sub


Private Sub cboTYPSU_Click()
cbo_Value recLrAttribut.TYPSU, cboTYPSU

End Sub


Private Sub cboTYRES_Click()
cbo_Value recLrAttribut.TYRES, cboTYRES
End Sub


Private Sub cboZACTI_Click()
cbo_Value recLrAttribut.ZACTI, cboZACTI
End Sub


Private Sub cmdContext_Click()
Unload Me
End Sub

Private Sub cmdOk_Click()
Dim X As String

X = Trim(txtRéférence)
If X = "" Then Call lstErr_Clear(lstErr, txtRéférence, " Préciser la référence"): Exit Sub

recLrAttribut.Référence = Format(X, txtRéférence_Format)
    
If recLrAttribut.Method = constAddNew Then
    If srvLrAttribut.Scan(recLrAttribut) > 0 Then
        Call lstErr_Clear(lstErr, txtRéférence, "Cette référence existe déjà"): Exit Sub
    End If
End If

arrLrAttribut(0) = recLrAttribut
Unload Me

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case Is = 13: KeyCode = 0: cmdContext_Return
    Case Is = 27: cmdContext_Quit
    Case Is = 44: KeyCode = 0: frmElpPrt.prtScreen
End Select

End Sub

Public Sub cmdContext_Quit()
Unload Me
End Sub

Public Sub cmdContext_Return()
    SendKeys "{TAB}"
End Sub


Private Sub Form_Load()
Set XForm = Me
Call MeInit(arrTagNb)
ReDim arrTag(arrTagNb + 1)

'attributs Luca Report
cboAFFPU_init
cboAGEMT_Init
cboAGENT_Init
cboAPPAR_init
cboAREFR_init
cboATTCF_init
cboAUTDV_init
cboBONIF_init
cboCAROB_init
cboCATET_init
cboCDRES_init
cboCDZON_init
cboCLCRC_init
cboCOTIT_init
cboCPEMS_init
cboCRDIV_init
cboCREIM_init
cboCREOR_init
cboCRETC_init
cboCRHYP_init
cboDCTOM_init
cboDRAC_init
cboDURIN_init
cboDUROM_init
cboDVOPR_init
cboECART_init
cboECFIN_init
cboELIGB_init
cboFAMDV_init
cboFOPIF_init
cboFPRBG_init
cboGARCF_init
cboMLFCE_init
cboMONDV_init
cboMUTFG_init
cboNACGA_init
cboNACGR_init
cboNACPS_init
cboNAEGA_init
cboNAIMO_init
cboNAOCB_init
cboNAPRO_init
cboNARCP_init
cboNATCP_init
cboNATCR_init
cboNATCS_init
cboNATDD_init
cboNATER_init
cboNATIF_init
cboNATIT_init
cboNATMA_init
cboNATOF_init
cboNATRS_init
cboNRAST_init
cboNREHB_init
cboOPCVM_init
cboOPEFC_init
cboOPFDH_init
cboOPREC_init
cboPAACT_init
cboPERIO_init
cboPRIMP_init
cboPROCB_init
cboREDES_init
cboREDHB_init
cboRESET_init
cboREZON_init
cboRISPA_init
cboSEMNT_init
cboSENOP_init
cboTCFPE_init
cboTOPIF_init
cboTYCGR_init
cboTYCOM_init
cboTYDSU_init
cboTYETS_init
cboTYPOR_init
cboTYPSU_init
cboTYRES_init
cboZACTI_init

'attributs Luca Risques

cboCDCPCO_init
cboCDCPJO_init
cboCDCPFU_init
cboCDAGCO_init
cboCDREME_init
cboTYMTDV_init
cboTYVENT_init
cboCRVENT_init
cboCDDURE_init
cboDUINIT_init
cboCDCRTI_init
cboCDCRAC_init
cboCDBIOR_init
cboCDDEIN_init
cboCDCRIM_init
cboCDCRCO_init
cboCDCREF_init
cboCDLODA_init
cboCDCRET_init
cboCDOMPO_init
cboCDOPIM_init
cboCDSWAP_init

cboREESC1_init
cboREESC6_init

End Sub



Public Sub cboAFFPU_init()
cboAFFPU.AddItem " "
cboAFFPU.AddItem "1 Fonds publics affectés à la garantie"
cboAFFPU.AddItem "2 Fonds publics non affectés à la garantie"
cboAFFPU.AddItem "9 non significatif"

End Sub

Public Sub cboCDCPCO_init()
cboCDCPCO.AddItem " "
cboCDCPCO.AddItem "1 Individuel"
cboCDCPCO.AddItem "2 Collectif"
End Sub
Public Sub cboCDCPJO_init()
cboCDCPJO.AddItem " "
cboCDCPJO.AddItem "1 Oui"
cboCDCPJO.AddItem "2 Non"
End Sub
Public Sub cboCDCPFU_init()
cboCDCPFU.AddItem " "
End Sub
Public Sub cboREESC1_init()
cboREESC1.AddItem " "
End Sub

Public Sub cboREESC6_init()
cboREESC1.AddItem " "
End Sub

Public Sub cboCDAGCO_init()


' !!!!!!!!!!!!!!!!!!! longueur zone 5 caractères
cboCDAGCO.AddItem " "
cboCDAGCO.AddItem "101   ET banques centrales & instituts d'émission"
cboCDAGCO.AddItem "102   ET CCP"
cboCDAGCO.AddItem "103   ET banques & caisses de crédit municipal"
cboCDAGCO.AddItem "104   ET caisses d'épargne & de prévoyance"
cboCDAGCO.AddItem "105   ET caisse des dépôts & consignations"
cboCDAGCO.AddItem "106   ET trésor public"
cboCDAGCO.AddItem "107   ET sociétés financières"
cboCDAGCO.AddItem "108   ET institutions financières spécialisées"
cboCDAGCO.AddItem "109   ET caisse centrale des caisses d'épargne"
cboCDAGCO.AddItem "110   ET entr.eff.des op.de banque à l'étranger"
cboCDAGCO.AddItem "111   ET org.bancaires & financiers internat."
cboCDAGCO.AddItem "112   ET sièges à l'étranger"
cboCDAGCO.AddItem "113   ET succursales à l'étranger"
cboCDAGCO.AddItem "201   ET institutions financières"
cboCDAGCO.AddItem "202   ET OPCVM monètaires"
cboCDAGCO.AddItem "203   ET fonds communs de créances"
cboCDAGCO.AddItem "204   ET OPCVM non monètaires"
cboCDAGCO.AddItem "301   ET sociétés & quasi-sociétés non financières"
cboCDAGCO.AddItem "302   ET entrepreneurs individuels"
cboCDAGCO.AddItem "303   ET particuliers"
cboCDAGCO.AddItem "304   ET sociétés d'assurance   ET fonds de pension"
cboCDAGCO.AddItem "305   ET administrations centrales"
cboCDAGCO.AddItem "306   ET administrations publiques locales"
cboCDAGCO.AddItem "307   ET administrations de sécurité sociale"
cboCDAGCO.AddItem "308   ET administrations privées"
cboCDAGCO.AddItem "309   ET état"
cboCDAGCO.AddItem "310   ET administrations d'états fédérés"
cboCDAGCO.AddItem "999   valeur non significative"
End Sub
Public Sub cboCDREME_init()
cboCDREME.AddItem " "
cboCDREME.AddItem "1 Oui"
cboCDREME.AddItem "2 Non"
End Sub
Public Sub cboTYMTDV_init()
cboTYMTDV.AddItem " "
cboTYMTDV.AddItem "00 Encours ordinaire"
cboTYMTDV.AddItem "01 Impayé"
cboTYMTDV.AddItem "02 Autorisation non utilisée"
cboTYMTDV.AddItem "03 Garantie déclarable"
cboTYMTDV.AddItem "04 Garantie non déclarable"
cboTYMTDV.AddItem "05 Contre-garantie"
cboTYMTDV.AddItem "06 Encours non déclarable"
End Sub
Public Sub cboTYVENT_init()
cboTYVENT.AddItem " "
cboTYVENT.AddItem "C Catégorie de risques "
cboTYVENT.AddItem "P Numéro compte PCI "
cboTYVENT.AddItem "X à ignorer"
End Sub
Public Sub cboCRVENT_init()
cboCRVENT.AddItem " "
End Sub
Public Sub cboCDDURE_init()
cboCDDURE.AddItem " "
cboCDDURE.AddItem "C Court terme"
cboCDDURE.AddItem "T autre "
End Sub
Public Sub cboDUINIT_init()
cboDUINIT.AddItem " "
End Sub
Public Sub cboCDCRTI_init()
cboCDCRTI.AddItem " "
cboCDCRTI.AddItem "1 Oui"
cboCDCRTI.AddItem "2 Non"
End Sub
Public Sub cboCDCRAC_init()
cboCDCRAC.AddItem " "
cboCDCRAC.AddItem "A Crédit acheteur "
cboCDCRAC.AddItem "R Crédit relais de crédit acheteur "
cboCDCRAC.AddItem "M Créance mobilisée de crédit acheteur"
End Sub
Public Sub cboCDBIOR_init()
cboCDBIOR.AddItem " "
cboCDBIOR.AddItem "1 Oui"
cboCDBIOR.AddItem "2 Non"
End Sub
Public Sub cboCDDEIN_init()
cboCDDEIN.AddItem " "
cboCDDEIN.AddItem "1 Oui"
cboCDDEIN.AddItem "2 Non"
End Sub
Public Sub cboCDCRIM_init()
cboCDCRIM.AddItem " "
cboCDCRIM.AddItem "1 Oui"
cboCDCRIM.AddItem "2 Non"
End Sub
Public Sub cboCDCRCO_init()
cboCDCRCO.AddItem " "
cboCDCRCO.AddItem "1 Oui"
cboCDCRCO.AddItem "2 Non"
End Sub
Public Sub cboCDCREF_init()
cboCDCREF.AddItem " "
cboCDCREF.AddItem "1 Oui"
cboCDCREF.AddItem "2 Non"
End Sub
Public Sub cboCDLODA_init()
cboCDLODA.AddItem " "
cboCDLODA.AddItem "1 Oui"
cboCDLODA.AddItem "2 Non"
End Sub
Public Sub cboCDCRET_init()
cboCDCRET.AddItem " "
cboCDCRET.AddItem "1 Oui"
cboCDCRET.AddItem "2 Non"
End Sub
Public Sub cboCDOMPO_init()
cboCDOMPO.AddItem " "
cboCDOMPO.AddItem "1 Oui"
cboCDOMPO.AddItem "2 Non"
End Sub
Public Sub cboCDOPIM_init()
cboCDOPIM.AddItem " "
cboCDOPIM.AddItem "1 Oui"
cboCDOPIM.AddItem "2 Non"
End Sub
Public Sub cboCDSWAP_init()
cboCDSWAP.AddItem " "
cboCDSWAP.AddItem "1 Oui"
cboCDSWAP.AddItem "2 Non"

End Sub


Public Sub cboAGEMT_Init()
cboAGEMT.AddItem " "
cboAGEMT.AddItem "101 ET banques centrales & instituts d'émission"
cboAGEMT.AddItem "102 ET CCP"
cboAGEMT.AddItem "103 ET banques & caisses de crédit municipal"
cboAGEMT.AddItem "104 ET caisses d'épargne & de prévoyance"
cboAGEMT.AddItem "105 ET caisse des dépôts & consignations"
cboAGEMT.AddItem "106 ET trésor public"
cboAGEMT.AddItem "107 ET sociétés financières"
cboAGEMT.AddItem "108 ET institutions financières spécialisées"
cboAGEMT.AddItem "109 ET caisse centrale des caisses d'épargne"
cboAGEMT.AddItem "110 ET entr.eff.des op.de banque à l'étranger"
cboAGEMT.AddItem "111 ET org.bancaires & financiers internat."
cboAGEMT.AddItem "112 ET sièges à l'étranger"
cboAGEMT.AddItem "113 ET succursales à l'étranger"
cboAGEMT.AddItem "201 ET institutions financières"
cboAGEMT.AddItem "202 ET OPCVM monètaires"
cboAGEMT.AddItem "203 ET fonds communs de créances"
cboAGEMT.AddItem "204 ET OPCVM non monètaires"
cboAGEMT.AddItem "301 ET sociétés & quasi-sociétés non financières"
cboAGEMT.AddItem "302 ET entrepreneurs individuels"
cboAGEMT.AddItem "303 ET particuliers"
cboAGEMT.AddItem "304 ET sociétés d'assurance et fonds de pension"
cboAGEMT.AddItem "305 ET administrations centrales"
cboAGEMT.AddItem "306 ET administrations publiques locales"
cboAGEMT.AddItem "307 ET administrations de sécurité sociale"
cboAGEMT.AddItem "308 ET administrations privées"
cboAGEMT.AddItem "309 ET état"
cboAGEMT.AddItem "310 ET administrations d'états fédérés"
cboAGEMT.AddItem "999  valeur non significative"

End Sub
Public Sub cboAGENT_Init()
cboAGENT.AddItem " "
cboAGENT.AddItem "101 cp banques centrales & instituts d'émission"
cboAGENT.AddItem "102 cp CCP"
cboAGENT.AddItem "103 cp banques & caisses de crédit municipal"
cboAGENT.AddItem "104 cp caisses d'épargne & de prévoyance"
cboAGENT.AddItem "105 cp caisse des dépôts & consignation"
cboAGENT.AddItem "106 cp trésor public"
cboAGENT.AddItem "107 cp sociétés financières"
cboAGENT.AddItem "108 cp institutions financières spécialisées"
cboAGENT.AddItem "109 cp caisse centrale des caisses d'épargne"
cboAGENT.AddItem "110 cp entr.eff.des op.de banque à l'étranger"
cboAGENT.AddItem "111 cp org.bancaires & financiers internat."
cboAGENT.AddItem "112 cp sièges à l'étranger"
cboAGENT.AddItem "113 cp succursales à l'étranger"
cboAGENT.AddItem "201 cp institutions financières"
cboAGENT.AddItem "202 cp OPCVM monètaires"
cboAGENT.AddItem "203 cp fonds communs de créances"
cboAGENT.AddItem "204 cp OPCVM non monètaires"
cboAGENT.AddItem "301 cp sociétés & quasi-sociétés non financières"
cboAGENT.AddItem "302 cp entrepreneurs individuels"
cboAGENT.AddItem "303 cp particuliers"
cboAGENT.AddItem "304 cp sociétés d'assurance et fonds de pension"
cboAGENT.AddItem "305 cp administrations centrales"
cboAGENT.AddItem "306 cp administrations publiques locales"
cboAGENT.AddItem "307 cp administrations de sécurité sociale"
cboAGENT.AddItem "308 cp administrations privées"
cboAGENT.AddItem "309 cp état"
cboAGENT.AddItem "310 cp administrations d'états fédérés"
cboAGENT.AddItem "999  valeur non significative"

End Sub




Public Sub cboAPPAR_init()
cboAPPAR.AddItem " "
cboAPPAR.AddItem "1 Amont CRB 85-12"
cboAPPAR.AddItem "2 Aval CRB 85-12"
cboAPPAR.AddItem "3 Non Apparenté"
cboAPPAR.AddItem "4 Amont autres"
cboAPPAR.AddItem "5 Aval autres"
cboAPPAR.AddItem "9 Valeur non significative"


End Sub

Public Sub cboAREFR_init()
cboAREFR.AddItem " "
cboAREFR.AddItem "1 Repris ds le calcul du coeff du bénéficiaire"
cboAREFR.AddItem "2 N/Repris ds le calcul du coeff du bénéficiaire"
cboAREFR.AddItem "9 Valeur non significative"

End Sub

Public Sub cboATTCF_init()
cboATTCF.AddItem " "
cboATTCF.AddItem "1 élément lié à CPTA financière à ne pas reprendre"
cboATTCF.AddItem "2 élément lié à CPTA financière à reprendre"
cboATTCF.AddItem "3 élément non lié à CPTA financière"
cboATTCF.AddItem "9 Valeur non significative"

End Sub

Public Sub cboAUTDV_init()
cboAUTDV.AddItem " "
cboAUTDV.AddItem "0 Principales Devises"
cboAUTDV.AddItem "1 Première Devise Significative"
cboAUTDV.AddItem "2 Deuxième Devise Significative"
cboAUTDV.AddItem "3 Troisième Devise Significative"
cboAUTDV.AddItem "4 Quatrième Devise Significative"
cboAUTDV.AddItem "5 Cinquième Devise Significative"
cboAUTDV.AddItem "6 Autres Devises non Significatives"
cboAUTDV.AddItem "9 Valeur non significative"

End Sub

Public Sub cboBONIF_init()
cboBONIF.AddItem " "
cboBONIF.AddItem "1 Prêt bonifié par l'état"
cboBONIF.AddItem "3 Prêt non bonifié"
cboBONIF.AddItem "9 Valeur non significative"

End Sub

Public Sub cboCAROB_init()
cboCAROB.AddItem " "
cboCAROB.AddItem "1 Emission dans le cadre des CODEVI"
cboCAROB.AddItem "2 Autre cadre d'émission"
cboCAROB.AddItem "9 Valeur non significative"

End Sub

Public Sub cboCATET_init()
cboCATET.AddItem " "
cboCATET.AddItem "01 Etablissement de crédit"
cboCATET.AddItem "02 Fonds de garantie du réseau"
cboCATET.AddItem "03 Client fin contr par Etabl.de CRDT du Res."
cboCATET.AddItem "04 Client nonFin entrant dans Périm.de consol"
cboCATET.AddItem "05 Etablissement non affilié au réseau"
cboCATET.AddItem "99 Valeur non significative"

End Sub

Public Sub cboCDRES_init()
cboCDRES.AddItem " "
cboCDRES.AddItem "1 CP Résident"
cboCDRES.AddItem "2 CP non résident"
cboCDRES.AddItem "9 Valeur non significative"

End Sub

Public Sub cboCDZON_init()
cboCDZON.AddItem " "
cboCDZON.AddItem "1 CP états membres de l'union monètaire"
cboCDZON.AddItem "2 CP états  non membres de l'union monètaire"
cboCDZON.AddItem "9 Valeur non significative"

End Sub

Public Sub cboCLCRC_init()
cboCLCRC.AddItem " "
cboCLCRC.AddItem "1 Crédit lié à des créances commerciales"
cboCLCRC.AddItem "2 Crédit non lié à des créances commerciales"
cboCLCRC.AddItem "9 Valeur non significative"

End Sub

Public Sub cboCOTIT_init()
cboCOTIT.AddItem " "
cboCOTIT.AddItem "1 Titres côtés sur un marché règlementé"
cboCOTIT.AddItem "2 Titres non côtés sur un marché règlementé"
cboCOTIT.AddItem "9 Valeur non significative"

End Sub

Public Sub cboCPEMS_init()
cboCPEMS.AddItem " "
cboCPEMS.AddItem "1 Prêts et emprunts subordonnés (ART 4C)"
cboCPEMS.AddItem "2 Prêts et emprunts subordonnés (ART 4D)"
cboCPEMS.AddItem "3 Autres Prêts et emprunts subordonnés"
cboCPEMS.AddItem "9 Valeur non significative"

End Sub

Public Sub cboCRDIV_init()
cboCRDIV.AddItem " "
cboCRDIV.AddItem "1 Provision pour impôt différé"
cboCRDIV.AddItem "2 Autres créditeurs divers"
cboCRDIV.AddItem "9 Valeur non significative"

End Sub

Public Sub cboCREIM_init()
cboCREIM.AddItem " "
cboCREIM.AddItem "1 Créances avec des impayés"
cboCREIM.AddItem "2 Créances sans impayé"
cboCREIM.AddItem "3 Créances douteuses"
cboCREIM.AddItem "9 Valeur non significative"

End Sub

Public Sub cboCREOR_init()
cboCREOR.AddItem " "
cboCREOR.AddItem "20311 Ventes à tempérament"
cboCREOR.AddItem "20312 Prêts personnels"
cboCREOR.AddItem "20313 Différés de remboursement liés à l'us. de CB"
cboCREOR.AddItem "20314 Utilisation d'ouvertures de crédits permanents"
cboCREOR.AddItem "20318 Avances sur avoirs financiers"
cboCREOR.AddItem "20319 Autres crédits de trésorerie"
cboCREOR.AddItem "20320 Autres prêts règlementés"
cboCREOR.AddItem "20321 Prêts à taux zéro"
cboCREOR.AddItem "20322 Crédits sur fonds CODEVI (pbe)"
cboCREOR.AddItem "99999 Valeur non significative"

End Sub

Public Sub cboCRETC_init()
cboCRETC.AddItem " "
cboCRETC.AddItem "1 Créances ou dettes à vue"
cboCRETC.AddItem "2 Créances ou dettes à terme"
cboCRETC.AddItem "3 Prêts au jour le jour"
cboCRETC.AddItem "4 Emprunts au jour le jour"
cboCRETC.AddItem "9 Valeur non significative"

End Sub

Public Sub cboCRHYP_init()
cboCRHYP.AddItem " "
cboCRHYP.AddItem "1 Crédit hypotécaire"
cboCRHYP.AddItem "2 Crédit non hypotécaire"
cboCRHYP.AddItem "3 Autres crédits acquèreurs garanties par HYP."
cboCRHYP.AddItem "9 Valeur non significative"

End Sub

Public Sub cboDCTOM_init()
cboDCTOM.AddItem " "
cboDCTOM.AddItem "1 Pour le compte d'autres établissements"
cboDCTOM.AddItem "2 Pour le compte de l'établissements déclarant"
cboDCTOM.AddItem "9 Valeur non significative"

End Sub

Public Sub cboDRAC_init()
cboDRAC.AddItem " "
cboDRAC.AddItem "1 DRAC < = 1 MOIS"
cboDRAC.AddItem "2 1 MOIS < DRAC < = 3 MOIS"
cboDRAC.AddItem "3 3 MOIS < DRAC < = 6 MOIS"
cboDRAC.AddItem "4 6 MOIS < DRAC < = 1 AN"
cboDRAC.AddItem "5 1 AN < DRAC < = 5 ANS"
cboDRAC.AddItem "6 DRAC > 5 ANS"
cboDRAC.AddItem "9 Valeur non significative"

End Sub

Public Sub cboDURIN_init()
cboDURIN.AddItem " "
cboDURIN.AddItem "0 Di < = 9 jours"
cboDURIN.AddItem "1 10J < = Di > = 1 AN"
cboDURIN.AddItem "2 1 AN < Di < 2 ANS"
cboDURIN.AddItem "3 2 ANS < = Di < = 5 ANS"
cboDURIN.AddItem "4 Di > 5 ANS "
cboDURIN.AddItem "9 Valeur non significative"

End Sub

Public Sub cboDUROM_init()
cboDUROM.AddItem " "
cboDUROM.AddItem "1 Di < = 1 AN"
cboDUROM.AddItem "2 1 AN < Di < 7 ANS"
cboDUROM.AddItem "3 Di > = 7 ANS"
cboDUROM.AddItem "9 Valeur non significative"

End Sub

Public Sub cboDVOPR_init()
cboDVOPR.AddItem " "
cboDVOPR.AddItem "1 Monnaie nationale"
cboDVOPR.AddItem "2 Devises"
cboDVOPR.AddItem "9 Valeur non significative"

End Sub

Public Sub cboECART_init()
cboECART.AddItem " "
cboECART.AddItem "1 Provisions financières"
cboECART.AddItem "2 Ecarts positifs"
cboECART.AddItem "3 Ecarts négatifs"
cboECART.AddItem "9 Valeur non significative"

End Sub

Public Sub cboECFIN_init()
cboECFIN.AddItem " "
cboECFIN.AddItem "1 Operations de crédit-bail"
cboECFIN.AddItem "2 Opérations de location simple"
cboECFIN.AddItem "3 Opération de location avec option d'achat"
cboECFIN.AddItem "9 Valeur non significative"

End Sub

Public Sub cboELIGB_init()
cboELIGB.AddItem " "
cboELIGB.AddItem "1 Eligible BDF"
cboELIGB.AddItem "2 Eligible IEOM/NON MH"
cboELIGB.AddItem "3 Non Eligible BDF ou IEOM"
cboELIGB.AddItem "4 Eligible au marché hypotécaire"
cboELIGB.AddItem "5 Eligible IEOM/éligible MH"
cboELIGB.AddItem "6 Non Eligible MH"
cboELIGB.AddItem "9 Valeur non significative"

End Sub

Public Sub cboFAMDV_init()
cboFAMDV.AddItem " "
cboFAMDV.AddItem "01 EUROS"
cboFAMDV.AddItem "10 Autres devises Européennes"
cboFAMDV.AddItem "20 USD"
cboFAMDV.AddItem "30 JPY"
cboFAMDV.AddItem "40 CHF"
cboFAMDV.AddItem "50 Autres devises"
cboFAMDV.AddItem "99 Valeur non significative"

End Sub

Public Sub cboFOPIF_init()
cboFOPIF.AddItem " "
cboFOPIF.AddItem "10 Contrats gérés en micro couvertures"
cboFOPIF.AddItem "11 Contrats gérés en macro couvertures"
cboFOPIF.AddItem "20 Contrats d'échange (SWAP)"
cboFOPIF.AddItem "21 Spéculation (Autres opérations...)"
cboFOPIF.AddItem "99 Valeur non significative"

End Sub

Public Sub cboFPRBG_init()
cboFPRBG.AddItem " "
cboFPRBG.AddItem "1 Provisions susceptibles d'être incluses au FRBG"
cboFPRBG.AddItem "2 Provisions non susceptibles d'être incluses au FRBG"
cboFPRBG.AddItem "9 Valeur non significative"

End Sub

Public Sub cboGARCF_init()
cboGARCF.AddItem " "
cboGARCF.AddItem "1 Crédits garantis par la coface"
cboGARCF.AddItem "2 Crédits non garantis par la coface"
cboGARCF.AddItem "9 Valeur non significative"

End Sub

Public Sub cboMLFCE_init()
cboMLFCE.AddItem " "
cboMLFCE.AddItem "1 cr.libell & Financ ds la monnaie de l'emprunt"
cboMLFCE.AddItem "2 cr non libell ou non financ.ds la monnaie"
cboMLFCE.AddItem "9 Valeur non significative"

End Sub

Public Sub cboMONDV_init()
cboMONDV.AddItem " "
cboMONDV.AddItem "1 EUROS"
cboMONDV.AddItem "2 DEVISES"

End Sub

Public Sub cboMUTFG_init()
cboMUTFG.AddItem " "
cboMUTFG.AddItem "1 Dépôts intégralement mutualisés"
cboMUTFG.AddItem "2 Dépôts non intégralement mutualisés"
cboMUTFG.AddItem "9 Valeur non significative"

End Sub

Public Sub cboNACGA_init()
cboNACGA.AddItem " "
cboNACGA.AddItem "1 CGA états et banques centrales"
cboNACGA.AddItem "2 CGA institutions des communautés européennes"
cboNACGA.AddItem "3 CGA nantissement dépôts ART 4.2.1 CRB 91-05"
cboNACGA.AddItem "4 CGA établissements de crédits"
cboNACGA.AddItem "5 CGA Banques multilatérales de développement"
cboNACGA.AddItem "6 CGA Administrations locales"
cboNACGA.AddItem "7 CGA Nantissements dépôts ART.4.2.2 CRB 91-05"
cboNACGA.AddItem "9 Valeur non significative"

End Sub

Public Sub cboNACGR_init()
cboNACGR.AddItem " "
cboNACGR.AddItem "0 VM émises par états CEE"
cboNACGR.AddItem "1 VM non émises garanties par états CEE"
cboNACGR.AddItem "2 VM non émises non garanties par états CEE"
cboNACGR.AddItem "9 Valeur non significative"

End Sub

Public Sub cboNACPS_init()
cboNACPS.AddItem " "
cboNACPS.AddItem "1 Administrations centrales et banques centrales"
cboNACPS.AddItem "2 Institutions des communautés européennes"
cboNACPS.AddItem "3 Banques multilatérales de développement"
cboNACPS.AddItem "4 Administrations régionales ou locales"
cboNACPS.AddItem "5 Etablissements de crédit"
cboNACPS.AddItem "6 Clientèle ordinaire"
cboNACPS.AddItem "7 Etablissements financiers non EC"
cboNACPS.AddItem "9 Valeur non significative"

End Sub

Public Sub cboNAEGA_init()
cboNAEGA.AddItem " "
cboNAEGA.AddItem "1 Nationalité du garant zone A"
cboNAEGA.AddItem "2 Nationalité du garant zone B"
cboNAEGA.AddItem "9 Valeur non significative"

End Sub

Public Sub cboNAIMO_init()
cboNAIMO.AddItem " "
cboNAIMO.AddItem "43100 Immobilisations incorporelles en cours"
cboNAIMO.AddItem "43200 Immobilisations corporelles en cours"
cboNAIMO.AddItem "44111 Droit au bail"
cboNAIMO.AddItem "44119 Autres éléments du fond commercial"
cboNAIMO.AddItem "44120 Frais d'établissement"
cboNAIMO.AddItem "44190 Autres Immobilisations incorporelles"
cboNAIMO.AddItem "44200 Immobilisations corporelles d'exploitation"
cboNAIMO.AddItem "45100 Immobilisations incorporelles hors exploitation"
cboNAIMO.AddItem "45200 Immobilisations corporelles hors exploitation"
cboNAIMO.AddItem "99999 Valeur non significative"

End Sub

Public Sub cboNAOCB_init()
cboNAOCB.AddItem " "
cboNAOCB.AddItem "4611 Crédit-Bail mobilier"
cboNAOCB.AddItem "4612 Crédit-Bail immobilier"
cboNAOCB.AddItem "4613 Crédit-Bail sur actifs incorporels"
cboNAOCB.AddItem "4621 Crédit-Bail mobilier (immo en cours)"
cboNAOCB.AddItem "4622 Crédit-Bail immobilier (immo en cours)"
cboNAOCB.AddItem "4623 Crédit-Bail actifs incorp.(immo en cours)"
cboNAOCB.AddItem "4630 Crédit-Bail immobilisations non louées après résiliation"
cboNAOCB.AddItem "9999 Valeur non significative"

End Sub

Public Sub cboNAPRO_init()
cboNAPRO.AddItem " "
cboNAPRO.AddItem "1 Provision spéciale (ART 64 loi 1970"
cboNAPRO.AddItem "2 Autres provisions"
cboNAPRO.AddItem "9 Valeur non significative"

End Sub

Public Sub cboNARCP_init()
cboNARCP.AddItem " "
cboNARCP.AddItem "1 Nationalité de la contrepartie zone A"
cboNARCP.AddItem "2 Nationalité de la contrepartie zone B"
cboNARCP.AddItem "9 Valeur non significative"

End Sub

Public Sub cboNATCP_init()
cboNATCP.AddItem " "
cboNATCP.AddItem "1 Contrepartie = compte d'ordre"
cboNATCP.AddItem "2 Contrepartie différente de compte d'ordre"
cboNATCP.AddItem "9 Valeur non significative"

End Sub

Public Sub cboNATCR_init()
cboNATCR.AddItem " "
cboNATCR.AddItem "01 Avances garanties par BDC.BE.CAT"
cboNATCR.AddItem "02 Avances garanties par autres avoirs financiers"
cboNATCR.AddItem "10 Crédit sur fonds publics affectés pour l'état"
cboNATCR.AddItem "11 Autres Crédit sur fonds publics affectés"
cboNATCR.AddItem "20 Prêts immobilier conventionné"
cboNATCR.AddItem "21 Prêts immobilier conventionné non PIC"
cboNATCR.AddItem "30 Autres natures de crédits"
cboNATCR.AddItem "99 Valeur non significative"

End Sub

Public Sub cboNATCS_init()
cboNATCS.AddItem " "
cboNATCS.AddItem "371 Promotion immobilière"
cboNATCS.AddItem "372 Avoirs en or & métaux précieux"
cboNATCS.AddItem "373 Autres stocks & assimilés"
cboNATCS.AddItem "376 Autres emplois divers"
cboNATCS.AddItem "999 Valeur non significative"

End Sub

Public Sub cboNATDD_init()
cboNATDD.AddItem " "
cboNATDD.AddItem "1 dépôt de garanties"
cboNATDD.AddItem "2 CODEVI titres de développement industriel"
cboNATDD.AddItem "3 CODEVI autre titres"
cboNATDD.AddItem "4 Autres débiteurs divers"
cboNATDD.AddItem "9 Valeur non significative"






End Sub

Public Sub cboNATER_init()
cboNATER.AddItem " "
cboNATER.AddItem "010 Caisse banque centrale CCP"
cboNATER.AddItem "020 Effets publics, valeurs assimilées"
cboNATER.AddItem "030 Créances sur établissement de crédit"
cboNATER.AddItem "041 Créances commerciales sur clientèle"
cboNATER.AddItem "042 Créances sur clientèle autres concours"
cboNATER.AddItem "043 Créances sur clientèle comptes ordin. débit"
cboNATER.AddItem "050 Affacturage (financement des adhérents)"
cboNATER.AddItem "060 Obligations & autres titres à revenu fixe"
cboNATER.AddItem "070 Actions & autres titres à revenu variable"
cboNATER.AddItem "080 Promotion immobilière"
cboNATER.AddItem "090 Participations & activité portefeuille"
cboNATER.AddItem "100 Parts dans les entreprises liées"
cboNATER.AddItem "110 Crédit-Bail & location avec option d'achat"
cboNATER.AddItem "120 Location simple"
cboNATER.AddItem "130 Immobilisations incorporelles"
cboNATER.AddItem "140 Immobilisations corporelles"
cboNATER.AddItem "150 Capital souscrit non versé"
cboNATER.AddItem "160 Actions propres"
cboNATER.AddItem "170 Autre actif"
cboNATER.AddItem "180 Comptes de régularisation"
cboNATER.AddItem "190 Finances sur ressources publ.ou semi-publ."
cboNATER.AddItem "200 Ayant pour objet des opér.à caractère productif"
cboNATER.AddItem "210 Concours accordés par les instit.financières"
cboNATER.AddItem "999 Valeur non significative"


End Sub

Public Sub cboNATIF_init()
cboNATIF.AddItem " "
cboNATIF.AddItem "10 Instruments de taux d'intérêt"
cboNATIF.AddItem "20 Instruments de cours de change cambiste"
cboNATIF.AddItem "21 Inst. de cours de change long(échan.financier)"
cboNATIF.AddItem "30 Instruments sur actions & indice boursier"
cboNATIF.AddItem "40 Autres instruments à terme"
cboNATIF.AddItem "99 Valeur non significative"

End Sub

Public Sub cboNATIT_init()
cboNATIT.AddItem " "
cboNATIT.AddItem "110 Titres du marché interbancaire"
cboNATIT.AddItem "120 Certificats de dépôts"
cboNATIT.AddItem "121 Bons des instituts financ.spécial & Soc financ."
cboNATIT.AddItem "122 Billets de trésorerie"
cboNATIT.AddItem "123 Bons du trésor"
cboNATIT.AddItem "124 BMTN émis par des ets de CRDTS ou MAI de Titres"
cboNATIT.AddItem "125 BMTN émis par la clientèle"
cboNATIT.AddItem "130 Obligations"
cboNATIT.AddItem "140 Titres subordonnés à terme"
cboNATIT.AddItem "150 Parts de FCC ordinaire"
cboNATIT.AddItem "151 Parts de FCC spécifique"
cboNATIT.AddItem "160 Autres titres à revenu fixe"
cboNATIT.AddItem "210 Actions"
cboNATIT.AddItem "220 Parts d'OPCVM court terme"
cboNATIT.AddItem "221 Autres parts d'OPCVM"
cboNATIT.AddItem "230 Autres titres à revenu variable"
cboNATIT.AddItem "310 Billets d'affacturage"
cboNATIT.AddItem "999 Valeur non significative"


End Sub

Public Sub cboNATMA_init()
cboNATMA.AddItem " "
cboNATMA.AddItem "1 Marchés organisés"
cboNATMA.AddItem "2 Marchés de gré à gré"
cboNATMA.AddItem "3 Marchés assimilés à des marchés organisés"
cboNATMA.AddItem "9 Valeur non significative"

End Sub

Public Sub cboNATOF_init()
cboNATOF.AddItem " "
cboNATOF.AddItem "10 Véhicules automobiles neufs"
cboNATOF.AddItem "20 Véhicules automobiles d'occasion"
cboNATOF.AddItem "31 Matériel éléctronique grand public"
cboNATOF.AddItem "32 Appareil ménagers"
cboNATOF.AddItem "33 Meubles"
cboNATOF.AddItem "34 Divers"
cboNATOF.AddItem "40 Autres biens & services"
cboNATOF.AddItem "99 Valeur non significative"

End Sub

Public Sub cboNATRS_init()
cboNATRS.AddItem " "
cboNATRS.AddItem "01 Livrets A"
cboNATRS.AddItem "02 Livrets bleus"
cboNATRS.AddItem "03 Livrets jeunes"
cboNATRS.AddItem "10 Compte d'épargne à long terme"
cboNATRS.AddItem "11 PER et PEA"
cboNATRS.AddItem "12 Dépôts épargne sur livre des soc.de CDT diffé."
cboNATRS.AddItem "13 Autres comptes d'épargne à régime spécial"
cboNATRS.AddItem "99 Valeur non significative"

End Sub

Public Sub cboNRAST_init()
cboNRAST.AddItem " "
cboNRAST.AddItem "1 Stocks négociable ne comportant aucun risque"
cboNRAST.AddItem "2 Stocks Non négociable ou présentant un risque"
cboNRAST.AddItem "9 Valeur non significative"

End Sub

Public Sub cboNREHB_init()
cboNREHB.AddItem " "
cboNREHB.AddItem "1 Risque faible"
cboNREHB.AddItem "2 Risque modéré"
cboNREHB.AddItem "3 Risque moyen"
cboNREHB.AddItem "4 Risque élevé"
cboNREHB.AddItem "9 Valeur non significative"

End Sub

Public Sub cboOPCVM_init()
cboOPCVM.AddItem " "
cboOPCVM.AddItem "1 OPCVM dont l'établissement est gestionnaire"
cboOPCVM.AddItem "2 OPCVM dont l'établissement est dépositaire"
cboOPCVM.AddItem "3 OPCVM non géré par l'établissement"
cboOPCVM.AddItem "9 Valeur non significative"

End Sub

Public Sub cboOPEFC_init()
cboOPEFC.AddItem " "
cboOPEFC.AddItem "1 Opérations fermes"
cboOPEFC.AddItem "2 Opérations conditionnelles"
cboOPEFC.AddItem "9 Valeur non significative"

End Sub

Public Sub cboOPFDH_init()
cboOPFDH.AddItem " "
cboOPFDH.AddItem "1 Opération de financement dans le territoire ou dept"
cboOPFDH.AddItem "2 Opération de financement hors le territoire ou dept"
cboOPFDH.AddItem "9 Valeur non significative"

End Sub

Public Sub cboOPREC_init()
cboOPREC.AddItem " "
cboOPREC.AddItem "1 Opérations avec succursales (interzones)"
cboOPREC.AddItem "2 Opérations avec filiales"
cboOPREC.AddItem "3 Opérations non réciproques"
cboOPREC.AddItem "9 Valeur non significative"

End Sub

Public Sub cboPAACT_init()
cboPAACT.AddItem " "

End Sub

Public Sub cboPERIO_init()
cboPERIO.AddItem " "
cboPERIO.AddItem "0 Période en cours"
cboPERIO.AddItem "1 Période précédente"

End Sub

Public Sub cboPRIMP_init()
cboPRIMP.AddItem " "
cboPRIMP.AddItem "1 Provisions ayant supporté l'impôt"
cboPRIMP.AddItem "2 Provisions n'ayant pas supporté l'impôt"
cboPRIMP.AddItem "9 Valeur non significative"

End Sub

Public Sub cboPROCB_init()
cboPROCB.AddItem " "
cboPROCB.AddItem "1 Op. de Crédit-Bail indemnités de résiliation "
cboPROCB.AddItem "2 Op.de Crédit-Bail autres produits"
cboPROCB.AddItem "9 Valeur non significative"

End Sub

Public Sub cboREDES_init()
cboREDES.AddItem " "
cboREDES.AddItem "1 Charges ou produits d'expl.bancaire"
cboREDES.AddItem "2 Charges ou produits d'expl non bancaire"
cboREDES.AddItem "3 Correction de valeurs sur CR et HB"
cboREDES.AddItem "4 Dotations reprises des FRBG"
cboREDES.AddItem "5 Dot. reprises provisions sur titres de transaction"
cboREDES.AddItem "6 Dot. reprises provisions sur titres de placement"
cboREDES.AddItem "7 Dot. reprises provisions sur instrum.financiers"
cboREDES.AddItem "9 Valeur non significative"
cboREDES.AddItem "A Opérations avec les établissements de crédit"
cboREDES.AddItem "B Opérations avec la clientèle"
cboREDES.AddItem "C Obligations et autres titres à revenu fixe"
cboREDES.AddItem "D Autres intérêts"

End Sub

Public Sub cboREDHB_init()
cboREDHB.AddItem " "
cboREDHB.AddItem "0 Engagements en faveur d'EC"
cboREDHB.AddItem "1 Engagements en faveur de la clientèle"
cboREDHB.AddItem "2 Cautions,avals,autres garanties d'ordre d'EC"
cboREDHB.AddItem "3 Garanties d'ordre à la clientèle"
cboREDHB.AddItem "4 Titres acquis avec faculté de rachat/reprise"
cboREDHB.AddItem "5 Autres titres à livrer"
cboREDHB.AddItem "6 Engagements reçus d'établissement de crédit"
cboREDHB.AddItem "7 Cautions,avals,autres garanties reçus d'EC"
cboREDHB.AddItem "8 Titres vendus avec faculté de rachat/reprise"
cboREDHB.AddItem "9 Valeur non significative"
cboREDHB.AddItem "A Autres titres à recevoir"
cboREDHB.AddItem "B Intervention à l'émission/marché gris reçus"
cboREDHB.AddItem "C Intervention à l'émission/marché gris donnés"

End Sub

Public Sub cboRESET_init()
cboRESET.AddItem " "
cboRESET.AddItem "1 ET résident"
cboRESET.AddItem "2 ET non résident"
cboRESET.AddItem "9 Valeur non significative"

End Sub

Public Sub cboREZON_init()
cboREZON.AddItem " "
cboREZON.AddItem "1 ET états membres de l'union monétaire"
cboREZON.AddItem "2 ET états non membres de l'union monétaire"
cboREZON.AddItem "9 Valeur non significative"

End Sub

Public Sub cboRISPA_init()
cboRISPA.AddItem " "
cboRISPA.AddItem "1 Pays à risque au sens BDF"
cboRISPA.AddItem "2 Pays non risque"
cboRISPA.AddItem "9 Valeur non significative"

End Sub

Public Sub cboSEMNT_init()
cboSEMNT.AddItem " "
cboSEMNT.AddItem "C Crédit"
cboSEMNT.AddItem "D Débit"

End Sub

Public Sub cboSENOP_init()
cboSENOP.AddItem " "
cboSENOP.AddItem "0 Achat"
cboSENOP.AddItem "1 Vente"
cboSENOP.AddItem "9 Valeur non significative"

End Sub

Public Sub cboTCFPE_init()
cboTCFPE.AddItem " "
cboTCFPE.AddItem "1 Titres constituant des fonds propres > 10 %"
cboTCFPE.AddItem "2 Titres ne constituant pas des fonds propres "
cboTCFPE.AddItem "3 Titres constituant des fonds propres < = 10 %"
cboTCFPE.AddItem "9 Valeur non significative"

End Sub

Public Sub cboTOPIF_init()
cboTOPIF.AddItem " "
cboTOPIF.AddItem "1 Francs à recevoir contre Devises à livrer"
cboTOPIF.AddItem "2 Devises à recevoir contre Francs  à livrer"
cboTOPIF.AddItem "3 Devises à recevoir contre Devises à livrer"
cboTOPIF.AddItem "4 Devises à livrer contre Devises à recevoir"
cboTOPIF.AddItem "9 Valeur non significative"

End Sub

Public Sub cboTYCGR_init()
cboTYCGR.AddItem " "
cboTYCGR.AddItem "1 Contregarantie reçue affectée au bilan"
cboTYCGR.AddItem "2 Contregarantie reçue affectée au hors bilan"
cboTYCGR.AddItem "9 Valeur non significative"

End Sub

Public Sub cboTYCOM_init()
cboTYCOM.AddItem " "
cboTYCOM.AddItem "1 Commissions de garantie"
cboTYCOM.AddItem "2 Commissions de placement"
cboTYCOM.AddItem "3 Autres commissions"
cboTYCOM.AddItem "9 Valeur non significative"

End Sub

Public Sub cboTYDSU_init()
cboTYDSU.AddItem " "
cboTYDSU.AddItem "1 Avances d'équilibre reçues"
cboTYDSU.AddItem "2 Autres emprunts subordonnés"
cboTYDSU.AddItem "9 Valeur non significative"

End Sub

Public Sub cboTYETS_init()
cboTYETS.AddItem " "
cboTYETS.AddItem "1 Société mère"
cboTYETS.AddItem "2 Entreprises consolidées françaises"
cboTYETS.AddItem "3 Entreprises consolidées étrangères"
cboTYETS.AddItem "9 Valeur non significative"

End Sub

Public Sub cboTYPOR_init()
cboTYPOR.AddItem " "
cboTYPOR.AddItem "110 Parts dans SCI de promo (hors titres prêtés)"
cboTYPOR.AddItem "150 Parts dans SCI de promo (titres prêtés)"
cboTYPOR.AddItem "210 Autres parts entr.liées (hors titres prêtés)"
cboTYPOR.AddItem "211 Titres de participation (hors titres prêtés)"
cboTYPOR.AddItem "212 Titres activ.portefeuille (hors titres prêtés)"
cboTYPOR.AddItem "250 Autres parts entr.liées (titres prêtés)"
cboTYPOR.AddItem "251 Titres de participation (titres prêtés)"
cboTYPOR.AddItem "252 Titres activ.portefeuille (titres prêtés)"
cboTYPOR.AddItem "310 Titres de placement (hors titres prêtés)"
cboTYPOR.AddItem "350 Titres de placement (titres prêtés)"
cboTYPOR.AddItem "410 Titres d'investissement(hors titres prêtés)"
cboTYPOR.AddItem "450 Titres d'investissement(titres prêtés)"
cboTYPOR.AddItem "510 Titres subordonnés(hors titres prêtés)"
cboTYPOR.AddItem "550 Titres subordonnés(titres prêtés)"
cboTYPOR.AddItem "560 Appels de fonds"
cboTYPOR.AddItem "999 Valeur non significative"

End Sub

Public Sub cboTYPSU_init()
cboTYPSU.AddItem " "
cboTYPSU.AddItem "1 Avances d'équilibre données"
cboTYPSU.AddItem "2 Autres prêts subordonnés"
cboTYPSU.AddItem "3 Prêts participatifs"
cboTYPSU.AddItem "9 Valeur non significative"

End Sub

Public Sub cboTYRES_init()
cboTYRES.AddItem " "
cboTYRES.AddItem "0 Moyenne"
cboTYRES.AddItem "1 Maximum"
cboTYRES.AddItem "9 Valeur non significative"

End Sub

Public Sub cboZACTI_init()
cboZACTI.AddItem " "
cboZACTI.AddItem "0 Métropole"
cboZACTI.AddItem "1 DOM"
cboZACTI.AddItem "2 TOM"
cboZACTI.AddItem "3 Etranger"

End Sub


Public Sub Display()

recLrAttribut = arrLrAttribut(0)
arrLrAttribut(0).Method = ""
''''lblRéférence = recLrAttribut.Nature
txtRéférence = Trim(recLrAttribut.Référence)
txtRéférence.Enabled = IIf(recLrAttribut.Method = constAddNew, True, False)

If recLrAttribut.Nature = "T" Then
    txtRéférence.MaxLength = 3
    txtRéférence_Format = "000"
    libRéférence = Trim(DicLib(13, Trim(recLrAttribut.Référence)))
Else
    txtRéférence.MaxLength = 11
    txtRéférence_Format = "00000000000"
    libRéférence = Trim("Compte ........")
End If

If Trim(recLrAttribut.AFFPU) = "" Then
    cboAFFPU.ListIndex = -1
Else
    cbo_Scan recLrAttribut.AFFPU, cboAFFPU
End If


If Trim(recLrAttribut.AGEMT) = "" Then
    cboAGEMT.ListIndex = -1
Else
    cbo_Scan recLrAttribut.AGEMT, cboAGEMT
End If

If Trim(recLrAttribut.AGENT) = "" Then
    cboAGENT.ListIndex = -1
Else
    cbo_Scan recLrAttribut.AGENT, cboAGENT
End If

If Trim(recLrAttribut.APPAR) = "" Then
    cboAPPAR.ListIndex = -1
Else
    cbo_Scan recLrAttribut.APPAR, cboAPPAR
End If

If Trim(recLrAttribut.AREFR) = "" Then
    cboAREFR.ListIndex = -1
Else
    cbo_Scan recLrAttribut.AREFR, cboAREFR
End If

If Trim(recLrAttribut.ATTCF) = "" Then
    cboATTCF.ListIndex = -1
Else
    cbo_Scan recLrAttribut.ATTCF, cboATTCF
End If

If Trim(recLrAttribut.AUTDV) = "" Then
    cboAUTDV.ListIndex = -1
Else
    cbo_Scan recLrAttribut.AUTDV, cboAUTDV
End If

If Trim(recLrAttribut.BONIF) = "" Then
    cboBONIF.ListIndex = -1
Else
    cbo_Scan recLrAttribut.BONIF, cboBONIF
End If

If Trim(recLrAttribut.CAROB) = "" Then
    cboCAROB.ListIndex = -1
Else
    cbo_Scan recLrAttribut.CAROB, cboCAROB
End If

If Trim(recLrAttribut.CATET) = "" Then
    cboCATET.ListIndex = -1
Else
    cbo_Scan recLrAttribut.CATET, cboCATET
End If

If Trim(recLrAttribut.CDRES) = "" Then
    cboCDRES.ListIndex = -1
Else
    cbo_Scan recLrAttribut.CDRES, cboCDRES
End If

If Trim(recLrAttribut.CDZON) = "" Then
    cboCDZON.ListIndex = -1
Else
    cbo_Scan recLrAttribut.CDZON, cboCDZON
End If

If Trim(recLrAttribut.CLCRC) = "" Then
    cboCLCRC.ListIndex = -1
Else
    cbo_Scan recLrAttribut.CLCRC, cboCLCRC
End If

If Trim(recLrAttribut.COTIT) = "" Then
    cboCOTIT.ListIndex = -1
Else
    cbo_Scan recLrAttribut.COTIT, cboCOTIT
End If

If Trim(recLrAttribut.CPEMS) = "" Then
    cboCPEMS.ListIndex = -1
Else
    cbo_Scan recLrAttribut.CPEMS, cboCPEMS
End If

If Trim(recLrAttribut.CRDIV) = "" Then
    cboCRDIV.ListIndex = -1
Else
    cbo_Scan recLrAttribut.CRDIV, cboCRDIV
End If

If Trim(recLrAttribut.CREIM) = "" Then
    cboCREIM.ListIndex = -1
Else
    cbo_Scan recLrAttribut.CREIM, cboCREIM
End If

If Trim(recLrAttribut.CREOR) = "" Then
    cboCREOR.ListIndex = -1
Else
    cbo_Scan recLrAttribut.CREOR, cboCREOR
End If

If Trim(recLrAttribut.CRETC) = "" Then
    cboCRETC.ListIndex = -1
Else
    cbo_Scan recLrAttribut.CRETC, cboCRETC
End If

If Trim(recLrAttribut.CRHYP) = "" Then
    cboCRHYP.ListIndex = -1
Else
    cbo_Scan recLrAttribut.CRHYP, cboCRHYP
End If

If Trim(recLrAttribut.DCTOM) = "" Then
    cboDCTOM.ListIndex = -1
Else
    cbo_Scan recLrAttribut.DCTOM, cboDCTOM
End If

If Trim(recLrAttribut.DRAC) = "" Then
    cboDRAC.ListIndex = -1
Else
    cbo_Scan recLrAttribut.DRAC, cboDRAC
End If

If Trim(recLrAttribut.DURIN) = "" Then
    cboDURIN.ListIndex = -1
Else
    cbo_Scan recLrAttribut.DURIN, cboDURIN
End If

If Trim(recLrAttribut.DUROM) = "" Then
    cboDUROM.ListIndex = -1
Else
    cbo_Scan recLrAttribut.DUROM, cboDUROM
End If

If Trim(recLrAttribut.DVOPR) = "" Then
    cboDVOPR.ListIndex = -1
Else
    cbo_Scan recLrAttribut.DVOPR, cboDVOPR
End If

If Trim(recLrAttribut.ECART) = "" Then
    cboECART.ListIndex = -1
Else
    cbo_Scan recLrAttribut.ECART, cboECART
End If

If Trim(recLrAttribut.ECFIN) = "" Then
    cboECFIN.ListIndex = -1
Else
    cbo_Scan recLrAttribut.ECFIN, cboECFIN
End If

If Trim(recLrAttribut.ELIGB) = "" Then
    cboELIGB.ListIndex = -1
Else
    cbo_Scan recLrAttribut.ELIGB, cboELIGB
End If

If Trim(recLrAttribut.FAMDV) = "" Then
    cboFAMDV.ListIndex = -1
Else
    cbo_Scan recLrAttribut.FAMDV, cboFAMDV
End If

If Trim(recLrAttribut.FOPIF) = "" Then
    cboFOPIF.ListIndex = -1
Else
    cbo_Scan recLrAttribut.FOPIF, cboFOPIF
End If

If Trim(recLrAttribut.FPRBG) = "" Then
    cboFPRBG.ListIndex = -1
Else
    cbo_Scan recLrAttribut.FPRBG, cboFPRBG
End If

If Trim(recLrAttribut.GARCF) = "" Then
    cboGARCF.ListIndex = -1
Else
    cbo_Scan recLrAttribut.GARCF, cboGARCF
End If

If Trim(recLrAttribut.MLFCE) = "" Then
    cboMLFCE.ListIndex = -1
Else
    cbo_Scan recLrAttribut.MLFCE, cboMLFCE
End If

If Trim(recLrAttribut.MONDV) = "" Then
    cboMONDV.ListIndex = -1
Else
    cbo_Scan recLrAttribut.MONDV, cboMONDV
End If

If Trim(recLrAttribut.MUTFG) = "" Then
    cboMUTFG.ListIndex = -1
Else
    cbo_Scan recLrAttribut.MUTFG, cboMUTFG
End If

If Trim(recLrAttribut.NACGA) = "" Then
    cboNACGA.ListIndex = -1
Else
    cbo_Scan recLrAttribut.NACGA, cboNACGA
End If

If Trim(recLrAttribut.NACGR) = "" Then
    cboNACGR.ListIndex = -1
Else
    cbo_Scan recLrAttribut.NACGR, cboNACGR
End If

If Trim(recLrAttribut.NACPS) = "" Then
    cboNACPS.ListIndex = -1
Else
    cbo_Scan recLrAttribut.NACPS, cboNACPS
End If

If Trim(recLrAttribut.NAEGA) = "" Then
    cboNAEGA.ListIndex = -1
Else
    cbo_Scan recLrAttribut.NAEGA, cboNAEGA
End If

If Trim(recLrAttribut.NAIMO) = "" Then
    cboNAIMO.ListIndex = -1
Else
    cbo_Scan recLrAttribut.NAIMO, cboNAIMO
End If

If Trim(recLrAttribut.NAOCB) = "" Then
    cboNAOCB.ListIndex = -1
Else
    cbo_Scan recLrAttribut.NAOCB, cboNAOCB
End If

If Trim(recLrAttribut.NAPRO) = "" Then
    cboNAPRO.ListIndex = -1
Else
    cbo_Scan recLrAttribut.NAPRO, cboNAPRO
End If

If Trim(recLrAttribut.NARCP) = "" Then
    cboNARCP.ListIndex = -1
Else
    cbo_Scan recLrAttribut.NARCP, cboNARCP
End If

If Trim(recLrAttribut.NATCP) = "" Then
    cboNATCP.ListIndex = -1
Else
    cbo_Scan recLrAttribut.NATCP, cboNATCP
End If

If Trim(recLrAttribut.NATCR) = "" Then
    cboNATCR.ListIndex = -1
Else
    cbo_Scan recLrAttribut.NATCR, cboNATCR
End If

If Trim(recLrAttribut.NATCS) = "" Then
    cboNATCS.ListIndex = -1
Else
    cbo_Scan recLrAttribut.NATCS, cboNATCS
End If

If Trim(recLrAttribut.NATDD) = "" Then
    cboNATDD.ListIndex = -1
Else
    cbo_Scan recLrAttribut.NATDD, cboNATDD
End If

If Trim(recLrAttribut.NATER) = "" Then
    cboNATER.ListIndex = -1
Else
    cbo_Scan recLrAttribut.NATER, cboNATER
End If

If Trim(recLrAttribut.NATIF) = "" Then
    cboNATIF.ListIndex = -1
Else
    cbo_Scan recLrAttribut.NATIF, cboNATIF
End If

If Trim(recLrAttribut.NATIT) = "" Then
    cboNATIT.ListIndex = -1
Else
    cbo_Scan recLrAttribut.NATIT, cboNATIT
End If

If Trim(recLrAttribut.NATMA) = "" Then
    cboNATMA.ListIndex = -1
Else
    cbo_Scan recLrAttribut.NATMA, cboNATMA
End If

If Trim(recLrAttribut.NATOF) = "" Then
    cboNATOF.ListIndex = -1
Else
    cbo_Scan recLrAttribut.NATOF, cboNATOF
End If

If Trim(recLrAttribut.NATRS) = "" Then
    cboNATRS.ListIndex = -1
Else
    cbo_Scan recLrAttribut.NATRS, cboNATRS
End If

If Trim(recLrAttribut.NRAST) = "" Then
    cboNRAST.ListIndex = -1
Else
    cbo_Scan recLrAttribut.NRAST, cboNRAST
End If

If Trim(recLrAttribut.NREHB) = "" Then
    cboNREHB.ListIndex = -1
Else
    cbo_Scan recLrAttribut.NREHB, cboNREHB
End If

If Trim(recLrAttribut.OPCVM) = "" Then
    cboOPCVM.ListIndex = -1
Else
    cbo_Scan recLrAttribut.OPCVM, cboOPCVM
End If

If Trim(recLrAttribut.OPEFC) = "" Then
    cboOPEFC.ListIndex = -1
Else
    cbo_Scan recLrAttribut.OPEFC, cboOPEFC
End If

If Trim(recLrAttribut.OPFDH) = "" Then
    cboOPFDH.ListIndex = -1
Else
    cbo_Scan recLrAttribut.OPFDH, cboOPFDH
End If

If Trim(recLrAttribut.OPREC) = "" Then
    cboOPREC.ListIndex = -1
Else
    cbo_Scan recLrAttribut.OPREC, cboOPREC
End If

If Trim(recLrAttribut.PERIO) = "" Then
    cboPERIO.ListIndex = -1
Else
    cbo_Scan recLrAttribut.PERIO, cboPERIO
End If

If Trim(recLrAttribut.PRIMP) = "" Then
    cboPRIMP.ListIndex = -1
Else
    cbo_Scan recLrAttribut.PRIMP, cboPRIMP
End If

If Trim(recLrAttribut.PROCB) = "" Then
    cboPROCB.ListIndex = -1
Else
    cbo_Scan recLrAttribut.PROCB, cboPROCB
End If

If Trim(recLrAttribut.REDES) = "" Then
    cboREDES.ListIndex = -1
Else
    cbo_Scan recLrAttribut.REDES, cboREDES
End If

If Trim(recLrAttribut.REDHB) = "" Then
    cboREDHB.ListIndex = -1
Else
    cbo_Scan recLrAttribut.REDHB, cboREDHB
End If

If Trim(recLrAttribut.RESET) = "" Then
    cboRESET.ListIndex = -1
Else
    cbo_Scan recLrAttribut.RESET, cboRESET
End If

If Trim(recLrAttribut.REZON) = "" Then
    cboREZON.ListIndex = -1
Else
    cbo_Scan recLrAttribut.REZON, cboREZON
End If

If Trim(recLrAttribut.RISPA) = "" Then
    cboRISPA.ListIndex = -1
Else
    cbo_Scan recLrAttribut.RISPA, cboRISPA
End If

If Trim(recLrAttribut.SENOP) = "" Then
    cboSENOP.ListIndex = -1
Else
    cbo_Scan recLrAttribut.SENOP, cboSENOP
End If

If Trim(recLrAttribut.TCFPE) = "" Then
    cboTCFPE.ListIndex = -1
Else
    cbo_Scan recLrAttribut.TCFPE, cboTCFPE
End If

If Trim(recLrAttribut.TOPIF) = "" Then
    cboTOPIF.ListIndex = -1
Else
    cbo_Scan recLrAttribut.TOPIF, cboTOPIF
End If

If Trim(recLrAttribut.TYCGR) = "" Then
    cboTYCGR.ListIndex = -1
Else
    cbo_Scan recLrAttribut.TYCGR, cboTYCGR
End If

If Trim(recLrAttribut.TYCOM) = "" Then
    cboTYCOM.ListIndex = -1
Else
    cbo_Scan recLrAttribut.TYCOM, cboTYCOM
End If

If Trim(recLrAttribut.TYDSU) = "" Then
    cboTYDSU.ListIndex = -1
Else
    cbo_Scan recLrAttribut.TYDSU, cboTYDSU
End If

If Trim(recLrAttribut.TYETS) = "" Then
    cboTYETS.ListIndex = -1
Else
    cbo_Scan recLrAttribut.TYETS, cboTYETS
End If

If Trim(recLrAttribut.TYPOR) = "" Then
    cboTYPOR.ListIndex = -1
Else
    cbo_Scan recLrAttribut.TYPOR, cboTYPOR
End If

If Trim(recLrAttribut.TYPSU) = "" Then
    cboTYPSU.ListIndex = -1
Else
    cbo_Scan recLrAttribut.TYPSU, cboTYPSU
End If

If Trim(recLrAttribut.TYRES) = "" Then
    cboTYRES.ListIndex = -1
Else
    cbo_Scan recLrAttribut.TYRES, cboTYRES
End If

If Trim(recLrAttribut.ZACTI) = "" Then
    cboZACTI.ListIndex = -1
Else
    cbo_Scan recLrAttribut.ZACTI, cboZACTI
End If

If Trim(recLrAttribut.PAACT) = "" Then
    cboPAACT.ListIndex = -1
Else
    cbo_Scan recLrAttribut.PAACT, cboPAACT
End If

'attributs Luca Risques

If Trim(recLrAttribut.CDCPCO) = "" Then
    cboCDCPCO.ListIndex = -1
Else
    cbo_Scan recLrAttribut.CDCPCO, cboCDCPCO
End If

If Trim(recLrAttribut.CDCPJO) = "" Then
    cboCDCPJO.ListIndex = -1
Else
    cbo_Scan recLrAttribut.CDCPJO, cboCDCPJO
End If

If Trim(recLrAttribut.CDCPFU) = "" Then
    cboCDCPFU.ListIndex = -1
Else
    cboCDCPFU = recLrAttribut.CDCPFU
End If

If Trim(recLrAttribut.CDAGCO) = "" Then
    cboCDAGCO.ListIndex = -1
Else
    cbo_Scan recLrAttribut.CDAGCO, cboCDAGCO
End If

If Trim(recLrAttribut.CDREME) = "" Then
    cboCDREME.ListIndex = -1
Else
    cbo_Scan recLrAttribut.CDREME, cboCDREME
End If

If Trim(recLrAttribut.TYMTDV) = "" Then
    cboTYMTDV.ListIndex = -1
Else
    cbo_Scan recLrAttribut.TYMTDV, cboTYMTDV
End If

If Trim(recLrAttribut.TYVENT) = "" Then
    cboTYVENT.ListIndex = -1
Else
    cbo_Scan recLrAttribut.TYVENT, cboTYVENT
End If

If Trim(recLrAttribut.CRVENT) = "" Then
    cboCRVENT.ListIndex = -1
Else
    cboCRVENT = recLrAttribut.CRVENT
End If

If Trim(recLrAttribut.CDDURE) = "" Then
    cboCDDURE.ListIndex = -1
Else
    cbo_Scan recLrAttribut.CDDURE, cboCDDURE
End If

If Trim(recLrAttribut.DUINIT) = "" Then
    cboDUINIT.ListIndex = -1
Else
    cboDUINIT = recLrAttribut.DUINIT
End If

If Trim(recLrAttribut.CDCRTI) = "" Then
    cboCDCRTI.ListIndex = -1
Else
    cbo_Scan recLrAttribut.CDCRTI, cboCDCRTI
End If

If Trim(recLrAttribut.CDCRAC) = "" Then
    cboCDCRAC.ListIndex = -1
Else
    cbo_Scan recLrAttribut.CDCRAC, cboCDCRAC
End If

If Trim(recLrAttribut.CDBIOR) = "" Then
    cboCDBIOR.ListIndex = -1
Else
    cbo_Scan recLrAttribut.CDBIOR, cboCDBIOR
End If

If Trim(recLrAttribut.CDDEIN) = "" Then
    cboCDDEIN.ListIndex = -1
Else
    cbo_Scan recLrAttribut.CDDEIN, cboCDDEIN
End If

If Trim(recLrAttribut.CDCRIM) = "" Then
    cboCDCRIM.ListIndex = -1
Else
    cbo_Scan recLrAttribut.CDCRIM, cboCDCRIM
End If

If Trim(recLrAttribut.CDCRCO) = "" Then
    cboCDCRCO.ListIndex = -1
Else
    cbo_Scan recLrAttribut.CDCRCO, cboCDCRCO
End If

If Trim(recLrAttribut.CDCREF) = "" Then
    cboCDCREF.ListIndex = -1
Else
    cbo_Scan recLrAttribut.CDCREF, cboCDCREF
End If

If Trim(recLrAttribut.CDLODA) = "" Then
    cboCDLODA.ListIndex = -1
Else
    cbo_Scan recLrAttribut.CDLODA, cboCDLODA
End If

If Trim(recLrAttribut.CDCRET) = "" Then
    cboCDCRET.ListIndex = -1
Else
    cbo_Scan recLrAttribut.CDCRET, cboCDCRET
End If

If Trim(recLrAttribut.CDOMPO) = "" Then
    cboCDOMPO.ListIndex = -1
Else
    cbo_Scan recLrAttribut.CDOMPO, cboCDOMPO
End If

If Trim(recLrAttribut.CDOPIM) = "" Then
    cboCDOPIM.ListIndex = -1
Else
    cbo_Scan recLrAttribut.CDOPIM, cboCDOPIM
End If

If Trim(recLrAttribut.CDSWAP) = "" Then
    cboCDSWAP.ListIndex = -1
Else
    cbo_Scan recLrAttribut.CDSWAP, cboCDSWAP
End If


If Trim(recLrAttribut.REESC1) = "" Then
    cboREESC1.ListIndex = -1
Else
    cboREESC1 = recLrAttribut.REESC1
End If

If Trim(recLrAttribut.REESC6) = "" Then
    cboREESC6.ListIndex = -1
Else
    cboREESC6 = recLrAttribut.REESC6
End If

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
Select Case SSTab1.Tab
    Case 0: cboAFFPU.SetFocus
    Case 1: cboCLCRC.SetFocus
    Case 2: cboDVOPR.SetFocus
    Case 3: cboNACGR.SetFocus
    Case 4: cboNATIF.SetFocus
    Case 5: cboPERIO.SetFocus
    Case 6: cboTYCGR.SetFocus
    Case 7: cboCDCPCO.SetFocus
    Case 8: cboCDBIOR.SetFocus
End Select

End Sub

Private Sub txtRéférence_KeyPress(KeyAscii As Integer)
Call num_KeyAscii(KeyAscii)
End Sub



Public Function PCI_Check(Msg As String)
Dim X

srvCompte.recCompteInit recCompte
recCompte.Method = "SeekL1"
recCompte.Société = 999
recCompte.Agence = 999
recCompte.Devise = 9999
recCompte.Numéro = Format$(Val(Msg), "00000000000")
X = srvCompteFind(recCompte)
If Not IsNull(X) Then MsgBox "PCI inexistant : " & recCompte.Numéro, vbCritical, " LrAttribut"
PCI_Check = X
End Function
