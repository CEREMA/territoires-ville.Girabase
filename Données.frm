VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "ss32x25.ocx"
Begin VB.Form frmDonnées 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "Giratoire1"
   ClientHeight    =   7590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14670
   Icon            =   "Données.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7590
   ScaleWidth      =   14670
   Begin TabDlg.SSTab tabDonnées 
      Height          =   6375
      Left            =   3960
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   120
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   11245
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Description du site"
      TabPicture(0)   =   "Données.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraEnvironnement"
      Tab(0).Control(1)=   "txtNomGiratoire"
      Tab(0).Control(2)=   "txtLocalisation"
      Tab(0).Control(3)=   "fraCarBranches"
      Tab(0).Control(4)=   "lblNomGiratoire"
      Tab(0).Control(5)=   "lblLocalisation"
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Dimensionnement"
      TabPicture(1)   =   "Données.frx":0326
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "fraAnneau"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fraBranches"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "fraVariante"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Données de trafic"
      TabPicture(2)   =   "Données.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblPériode"
      Tab(2).Control(1)=   "fraTrafic(0)"
      Tab(2).Control(2)=   "fraTraficTout"
      Tab(2).Control(3)=   "fraTrafic(1)"
      Tab(2).Control(4)=   "cboPériode"
      Tab(2).Control(5)=   "fraQTE"
      Tab(2).Control(6)=   "cmdChangeMode"
      Tab(2).Control(7)=   "vgdTrafic(0)"
      Tab(2).ControlCount=   8
      Begin FPSpread.vaSpread vgdTrafic 
         Height          =   1215
         Index           =   0
         Left            =   -74760
         TabIndex        =   20
         Top             =   3480
         Width           =   2790
         _Version        =   131077
         _ExtentX        =   4921
         _ExtentY        =   2143
         _StockProps     =   64
         AutoSize        =   -1  'True
         ColHeaderDisplay=   1
         EditEnterAction =   5
         EditModePermanent=   -1  'True
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   4
         MaxRows         =   4
         ProcessTab      =   -1  'True
         ScrollBars      =   0
         SelectBlockOptions=   0
         SpreadDesigner  =   "Données.frx":035E
         UnitType        =   2
         UserResize      =   0
         VisibleCols     =   500
         VisibleRows     =   500
      End
      Begin VB.CommandButton cmdChangeMode 
         Caption         =   "Mode VL-PL-2R"
         Height          =   495
         Left            =   -70560
         TabIndex        =   22
         Top             =   2760
         Width           =   855
      End
      Begin VB.Frame fraQTE 
         Enabled         =   0   'False
         Height          =   1815
         Left            =   -71880
         TabIndex        =   47
         Top             =   3360
         Width           =   855
         Begin VB.TextBox txtQE 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   49
            Top             =   340
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox txtQT 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   48
            Top             =   1440
            Width           =   615
         End
         Begin VB.Label lblTQE 
            Caption         =   "TE"
            Height          =   255
            Left            =   240
            TabIndex        =   50
            Top             =   0
            Width           =   375
         End
      End
      Begin VB.ComboBox cboPériode 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         ItemData        =   "Données.frx":0525
         Left            =   -73560
         List            =   "Données.frx":0527
         TabIndex        =   11
         Text            =   "Période1"
         Top             =   840
         Width           =   3015
      End
      Begin VB.Frame fraVariante 
         Height          =   615
         Left            =   720
         TabIndex        =   44
         Top             =   480
         Width           =   5415
         Begin VB.TextBox txtVariante 
            Height          =   285
            Left            =   1080
            TabIndex        =   6
            Top             =   240
            Width           =   2655
         End
         Begin VB.Label lblDateModif 
            AutoSize        =   -1  'True
            Height          =   195
            Left            =   4080
            TabIndex        =   63
            Top             =   240
            Width           =   45
         End
         Begin VB.Label lblVariante 
            Caption         =   "Variante :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame fraBranches 
         Caption         =   "Branches"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3015
         Left            =   720
         TabIndex        =   36
         Top             =   3120
         Width           =   5415
         Begin FPSpread.vaSpread vgdLargBranche 
            Height          =   1935
            Left            =   360
            TabIndex        =   10
            Top             =   960
            Width           =   4680
            _Version        =   131077
            _ExtentX        =   8255
            _ExtentY        =   3413
            _StockProps     =   64
            AutoSize        =   -1  'True
            DisplayColHeaders=   0   'False
            EditEnterAction =   5
            EditModePermanent=   -1  'True
            EditModeReplace =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GrayAreaBackColor=   -2147483643
            MaxCols         =   5
            MaxRows         =   8
            ProcessTab      =   -1  'True
            ScrollBars      =   0
            SelectBlockOptions=   0
            ShadowDark      =   12632256
            ShadowText      =   4210752
            SpreadDesigner  =   "Données.frx":0529
            UserResize      =   0
            VisibleCols     =   500
            VisibleRows     =   500
         End
         Begin VB.Label lblLE15m 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "à 15 m"
            Height          =   255
            Left            =   1560
            TabIndex        =   43
            Top             =   720
            Width           =   735
         End
         Begin VB.Label lblLE4m 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "à 4 m"
            Height          =   255
            Left            =   840
            TabIndex        =   42
            Top             =   720
            Width           =   735
         End
         Begin VB.Label lblLargeurs 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Largeurs"
            Height          =   255
            Left            =   840
            TabIndex        =   41
            Top             =   240
            Width           =   3135
         End
         Begin VB.Label lblLS 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Sortie"
            Height          =   495
            Left            =   3240
            TabIndex        =   40
            Top             =   480
            Width           =   735
         End
         Begin VB.Label lblLI 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Ilôt"
            Height          =   495
            Left            =   2400
            TabIndex        =   39
            Top             =   480
            Width           =   735
         End
         Begin VB.Label lblEntrée 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Entrée"
            Height          =   255
            Left            =   840
            TabIndex        =   38
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label lblEntréeEvasée 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Entrée évasée"
            Height          =   495
            Left            =   4080
            TabIndex        =   37
            Top             =   480
            Width           =   855
         End
      End
      Begin VB.Frame fraEnvironnement 
         Caption         =   "Environnement"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   -74760
         TabIndex        =   18
         Top             =   2040
         Width           =   5895
         Begin VB.OptionButton optMilieu 
            Caption         =   "Urbain"
            Height          =   255
            Index           =   2
            Left            =   4320
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   360
            Width           =   975
         End
         Begin VB.OptionButton optMilieu 
            Caption         =   "Péri Urbain"
            Height          =   255
            Index           =   1
            Left            =   2280
            TabIndex        =   3
            Top             =   360
            Width           =   1455
         End
         Begin VB.OptionButton optMilieu 
            Caption         =   "Rase Campagne"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   2
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.TextBox txtNomGiratoire 
         Height          =   375
         Left            =   -73080
         TabIndex        =   0
         Top             =   720
         Width           =   4095
      End
      Begin VB.TextBox txtLocalisation 
         Height          =   495
         Left            =   -73080
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   1200
         Width           =   4095
      End
      Begin VB.Frame fraAnneau 
         Caption         =   "Anneau"
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
         Left            =   720
         TabIndex        =   24
         Top             =   1200
         Width           =   5415
         Begin VB.TextBox txtRg 
            BackColor       =   &H8000000B&
            Height          =   285
            Left            =   2880
            Locked          =   -1  'True
            TabIndex        =   64
            Top             =   1320
            Width           =   735
         End
         Begin VB.TextBox txtBf 
            Height          =   285
            Left            =   2880
            TabIndex        =   8
            Top             =   600
            Width           =   735
         End
         Begin VB.TextBox txtR 
            Height          =   285
            Left            =   2880
            MaxLength       =   8
            TabIndex        =   7
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtLA 
            Height          =   285
            Left            =   2880
            MaxLength       =   8
            TabIndex        =   9
            Top             =   960
            Width           =   735
         End
         Begin VB.Label lblMetres 
            Caption         =   "mètres"
            Height          =   255
            Index           =   3
            Left            =   3720
            TabIndex        =   32
            Top             =   1440
            Width           =   615
         End
         Begin VB.Label lblRg 
            Caption         =   "Rayon extérieur du giratoire :"
            Height          =   255
            Left            =   240
            TabIndex        =   31
            Top             =   1440
            Width           =   2295
         End
         Begin VB.Label lblMetres 
            Caption         =   "mètres"
            Height          =   255
            Index           =   1
            Left            =   3720
            TabIndex        =   30
            Top             =   720
            Width           =   615
         End
         Begin VB.Label lblBf 
            Caption         =   "Largeur de la bande franchissable : "
            Height          =   255
            Left            =   240
            TabIndex        =   29
            Top             =   720
            Width           =   2535
         End
         Begin VB.Label lblR 
            Caption         =   "Rayon de l'îlot infranchissable :"
            Height          =   255
            Left            =   240
            TabIndex        =   28
            Top             =   360
            Width           =   2295
         End
         Begin VB.Label lblLA 
            AutoSize        =   -1  'True
            Caption         =   "Largeur de l'anneau :"
            Height          =   195
            Left            =   240
            TabIndex        =   27
            Top             =   1080
            Width           =   2145
         End
         Begin VB.Label lblMetres 
            AutoSize        =   -1  'True
            Caption         =   "mètres"
            Height          =   195
            Index           =   0
            Left            =   3720
            TabIndex        =   26
            Top             =   360
            Width           =   465
         End
         Begin VB.Label lblMetres 
            AutoSize        =   -1  'True
            Caption         =   "mètres"
            Height          =   195
            Index           =   2
            Left            =   3720
            TabIndex        =   25
            Top             =   1080
            Width           =   465
         End
      End
      Begin VB.Frame fraCarBranches 
         Caption         =   "Caractéristiques des branches"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2895
         Left            =   -74760
         TabIndex        =   19
         Top             =   3240
         Width           =   6015
         Begin FPSpread.vaSpread vgdCarBranche 
            Height          =   1935
            Left            =   120
            TabIndex        =   5
            Top             =   840
            Width           =   5655
            _Version        =   131077
            _ExtentX        =   9975
            _ExtentY        =   3413
            _StockProps     =   64
            AutoSize        =   -1  'True
            BackColorStyle  =   3
            ColHeaderDisplay=   0
            DisplayColHeaders=   0   'False
            EditEnterAction =   5
            EditModePermanent=   -1  'True
            EditModeReplace =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   5
            MaxRows         =   8
            ProcessTab      =   -1  'True
            ScrollBars      =   0
            SelectBlockOptions=   0
            SpreadDesigner  =   "Données.frx":1692
            UserResize      =   0
            VisibleCols     =   500
            VisibleRows     =   500
         End
         Begin VB.Label lblNomBranche 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Nom"
            Height          =   495
            Left            =   600
            TabIndex        =   57
            Top             =   360
            Width           =   2295
         End
         Begin VB.Label lblTAD 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Tourne à droite"
            Height          =   495
            Left            =   5040
            TabIndex        =   56
            Top             =   360
            Width           =   735
         End
         Begin VB.Label lblRampe 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Rampe > 3%"
            Height          =   495
            Left            =   4320
            TabIndex        =   55
            Top             =   360
            Width           =   735
         End
         Begin VB.Label lblEcart 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Ecart"
            Height          =   495
            Left            =   3600
            TabIndex        =   54
            Top             =   360
            Width           =   735
         End
         Begin VB.Label lblAngleBranche 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Angle "
            Height          =   495
            Left            =   2880
            TabIndex        =   53
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.Frame fraTrafic 
         Caption         =   "Trafic Piétons"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   1
         Left            =   -74880
         TabIndex        =   51
         Top             =   1320
         Width           =   6255
         Begin FPSpread.vaSpread vgdTrafic 
            Height          =   495
            Index           =   1
            Left            =   600
            TabIndex        =   12
            Top             =   360
            Width           =   4935
            _Version        =   131077
            _ExtentX        =   8705
            _ExtentY        =   873
            _StockProps     =   64
            AutoSize        =   -1  'True
            ColHeaderDisplay=   1
            DisplayRowHeaders=   0   'False
            EditEnterAction =   5
            EditModePermanent=   -1  'True
            EditModeReplace =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   8
            MaxRows         =   1
            ProcessTab      =   -1  'True
            ScrollBars      =   0
            SelectBlockOptions=   0
            SpreadDesigner  =   "Données.frx":36C9
            UnitType        =   2
            UserResize      =   0
            VisibleCols     =   500
            VisibleRows     =   500
         End
      End
      Begin VB.Frame fraTraficTout 
         Caption         =   "Trafic"
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
         Left            =   -74640
         TabIndex        =   21
         Top             =   2640
         Visible         =   0   'False
         Width           =   3615
         Begin VB.OptionButton optTrafic 
            Caption         =   "VL"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   14
            Top             =   240
            Value           =   -1  'True
            Width           =   615
         End
         Begin VB.OptionButton optTrafic 
            Caption         =   "PL"
            Height          =   255
            Index           =   1
            Left            =   840
            TabIndex        =   17
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton optTrafic 
            Caption         =   "2R"
            Height          =   255
            Index           =   2
            Left            =   1440
            TabIndex        =   16
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton optTrafic 
            Caption         =   "Voir UVP"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   2280
            TabIndex        =   15
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame fraTrafic 
         Caption         =   "Trafic Véhicules"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3855
         Index           =   0
         Left            =   -74880
         TabIndex        =   52
         Top             =   2400
         Width           =   6255
         Begin VB.Frame fraQTS 
            Enabled         =   0   'False
            Height          =   495
            Left            =   120
            TabIndex        =   60
            Top             =   2280
            Width           =   2865
            Begin VB.TextBox txtQS 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   0
               Left            =   480
               Locked          =   -1  'True
               TabIndex        =   61
               Top             =   120
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label lblTQS 
               Caption         =   "TS"
               Height          =   255
               Left            =   120
               TabIndex        =   62
               Top             =   120
               Width           =   375
            End
         End
         Begin VB.Label lblTraficUVP 
            AutoSize        =   -1  'True
            Caption         =   "Trafic UVP"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   360
            TabIndex        =   13
            Top             =   240
            Width           =   2505
         End
      End
      Begin VB.Label lblNomGiratoire 
         AutoSize        =   -1  'True
         Caption         =   "Nom du Carrefour : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74760
         TabIndex        =   35
         Top             =   840
         Width           =   1770
      End
      Begin VB.Label lblLocalisation 
         AutoSize        =   -1  'True
         Caption         =   "Localisation :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74760
         TabIndex        =   34
         Top             =   1320
         Width           =   1155
      End
      Begin VB.Label lblPériode 
         AutoSize        =   -1  'True
         Caption         =   "Période :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74640
         TabIndex        =   33
         Top             =   840
         Width           =   780
      End
   End
   Begin VB.Frame fraInvite 
      Height          =   600
      Left            =   120
      TabIndex        =   58
      Top             =   6360
      Width           =   5000
      Begin VB.Label lblInvite 
         Height          =   1275
         Left            =   0
         TabIndex        =   59
         Top             =   0
         Width           =   11595
         WordWrap        =   -1  'True
      End
   End
   Begin MSComDlg.CommonDialog dlgFichier 
      Left            =   120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   ".gbs"
      Filter          =   "Giratoire (*.gbs)|*.Gbs"
   End
   Begin VB.Label lblNumBranche 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "B0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   2400
      TabIndex        =   65
      Top             =   240
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Line linVoieSortie 
      BorderColor     =   &H00C0C0C0&
      Index           =   0
      Visible         =   0   'False
      X1              =   0
      X2              =   1680
      Y1              =   120
      Y2              =   0
   End
   Begin VB.Line linVoieEntrée 
      BorderColor     =   &H00C0C0C0&
      Index           =   0
      Visible         =   0   'False
      X1              =   0
      X2              =   1560
      Y1              =   120
      Y2              =   0
   End
   Begin VB.Shape shpAnneauMil 
      Height          =   1095
      Left            =   720
      Shape           =   3  'Circle
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Shape shpAnneauInt 
      Height          =   855
      Left            =   840
      Shape           =   3  'Circle
      Top             =   885
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Shape shpAnneauExt 
      Height          =   1515
      Left            =   600
      Shape           =   3  'Circle
      Top             =   525
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.Line linBranche 
      Index           =   0
      Visible         =   0   'False
      X1              =   1440
      X2              =   3000
      Y1              =   2805
      Y2              =   2085
   End
   Begin VB.Shape shpPoignée 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   105
      Left            =   120
      Top             =   765
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label lblLibelléBranche 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Branche0"
      Height          =   195
      Index           =   0
      Left            =   1440
      TabIndex        =   46
      Top             =   240
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Line linBordVoieEntrée 
      BorderColor     =   &H00C0C0C0&
      Index           =   0
      Visible         =   0   'False
      X1              =   240
      X2              =   1800
      Y1              =   1725
      Y2              =   1605
   End
   Begin VB.Line linBordVoieSortie 
      BorderColor     =   &H00C0C0C0&
      Index           =   0
      Visible         =   0   'False
      X1              =   0
      X2              =   1680
      Y1              =   2325
      Y2              =   2205
   End
   Begin VB.Line linBordIlotGir 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   0
      Visible         =   0   'False
      X1              =   480
      X2              =   480
      Y1              =   3405
      Y2              =   2925
   End
   Begin VB.Line linBordIlotEntrée 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   0
      Visible         =   0   'False
      X1              =   480
      X2              =   1560
      Y1              =   2925
      Y2              =   3165
   End
   Begin VB.Line linBordIlotSortie 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Index           =   0
      Visible         =   0   'False
      X1              =   480
      X2              =   1560
      Y1              =   3405
      Y2              =   3165
   End
End
Attribute VB_Name = "frmDonnées"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
'*
'*          Projet GIRABASE - CERTU - CETE de l'Ouest
'*
'*          Module de feuille : DONNEES.FRM - frmDonnées
'*
'*          Feuille principale de GIRABASE
'*
'*          Feuille fille de la feuille MDIGirabase
'*
'*          Un objet GIRATOIRE est associé à chaque feuille frmDonnées
'*          La propriété Données de cet objet désigne cette feuille
'*
'******************************************************************************

Option Explicit


Public GiratoireProjet As GIRATOIRE
Public flagErreurFatale As Boolean

Private SauvHelpContextId As Integer
Private InviteGotFocus As String
Public FeuilleBranche As Form
Public AncienXc As Single
Public AncienYc As Single
Public DemiLargeur As Single
Public DemiHauteur As Single
Public FacteurZoom As Single
Public AutreOnglet As Boolean
Public ValidateObjet As Boolean
Public ChargementEnCours As Boolean ' Permet d'inhiber certains évènements lors du chargement (initialisations nécessaires)
Public Nouveau As Boolean           ' Indique si le projet est Nouveau ou lu (commande Ouvrir)
Public BrancheSélectée As Integer
Public DonnéeModifiée As Boolean    ' Indique si un champ a été modifié
Public FichierModifié As Boolean    ' Indique si une modification a été faite
                                    ' depuis le dernier enregistrement
Private DonnéeValide As Boolean     ' Indique si la donnée peut être validée
Private DiagramFlux As Boolean      ' Indique si le dessin du diagramme de flux a été demandé
Private NbBranches As Integer       ' Nombre de branches

'******************************************************************************
Private FlagPtDecimal As Boolean    ' Drapeau indiquant que l'utilisateur a frappé le point décimal
                                    'sur le pavé numérique
Private TraficActif As TRAFIC       'Objet Trafic
Private TraficModifié As Boolean  'Indicateur de modification de la matrice Trafic
Private IndicSaisie As Boolean      ' Indicateur interdisant la modification du nombre de branches dès qu'une donnée a été saisie
Public MessageEmis As Boolean       ' Indicateur signalant l'émission de messages d'erreur ou de recommandantion
' Indicateurs pour de gestion des controles SPREAD
Private Débordement As Boolean      ' Si Vrai, l'utilisateur a quitté le SPREAD avec TAB ou SHIFT TAB
Private DrapeauSuivant As Boolean   ' Si Vrai et Débordement, sortie par TAB (sinon SHIFT TAB)
Private controleEnCours As Boolean
Public TypeControleActif As String  ' Type de variable en cours de saisie
Public ControleActif As Control     ' Contrôle en cours de saisie
'Pour une matrice...
Public TypeMatriceActive  As Integer ' Type de matrice en cours de saisie
Public NuméroColonneActive          ' Numéro de la colonne active
Public NuméroLigneActive            ' Numéro de la ligne active
Private EvenementClick
' Sauvegarde des valeurs lors du GotFocus
Private sauvAngle() As Single
Private SauveValeur As String
Private SauveValeurSpread As String 'GS09
Private ChaineInvite, ChaineMessage As String
Private ChangementOnglet As Boolean


'******************************************************************************
' Choix d'une période dans la liste
'******************************************************************************
Public Sub cboPériode_Click()
Dim i, j As Integer
  If cboPériode.ListIndex = -1 Then
    Set TraficActif = Nothing
    AutorTrafic False
    cboPériode.SetFocus
    'Efface le diagramme de flux
    cLS
    'Efface la matrice de piétons
    With vgdTrafic(PIETON)
      .Row = 1
      For j = 1 To NbBranches
        .Col = j
        .Value = ""
      Next j
    End With
    'Efface la matrice de trafic
    txtQT = ""
    With vgdTrafic(VEHICULE)
      For i = 1 To NbBranches
        .Row = i
        For j = 1 To NbBranches
          .Col = j
          .Value = ""
        Next j
      Next i
    End With
    For j = 1 To NbBranches
      txtQS(j) = ""
    Next
    For i = 1 To NbBranches
      txtQE(i) = ""
    Next i
    Exit Sub
  End If
  
  On Error GoTo GestErr
  Set TraficActif = GiratoireProjet.colTrafics.Item(cboPériode.ListIndex + 1)
  If TraficActif.modeUVP Then
    If fraTraficTout.Visible Then
      'Bascule MODE VL-PL-2R --> UVP
      cmdChangeMode.Caption = IDl_ModeVLPL2R
      fraTraficTout.Visible = False
      lblTraficUVP.Visible = True
    End If
    vgdTrafic(VEHICULE).Enabled = True
    TraficActif.setVéhicule UVP
  Else
    If Not fraTraficTout.Visible Then
    'Bascule MODE  UVP  --> VL-PL-2R
      cmdChangeMode.Caption = IDl_ModeUVP
      fraTraficTout.Visible = True
      lblTraficUVP.Visible = False
      optTrafic(VL) = True
    End If
    TraficActif.setVéhicule Numopt(optTrafic)
  End If
  'Contrôle d'un trafic inexistant pour une voie d'entrée ou de sortie non nulle
  'rq0599 'Inversion des 2 tests
  ControleValeursTrafic
  ControleMatriceTrafic
  AutorTrafic True
  If DiagramFlux And Not TraficActif Is Nothing Then cLS: TraficActif.CalculDiagramFlux

Exit Sub
  
GestErr:
  If Err = 9 Then
    Exit Sub
  Else
    ErreurFatale
  End If
End Sub

Private Sub cboPériode_GotFocus()
  lblInvite = IDi_Période
  AutreOnglet = False
  Journal "GotFocus"
End Sub

Private Sub cboPériode_LostFocus()
  If ActiveControl.Name = "vgdTrafic" Then
'    With ActiveControl
'      .Action = 0
'    End With
  End If
End Sub

'******************************************************************************
' Validation du choix d'une période dans la liste
'******************************************************************************
Private Sub cboPériode_Validate(Cancel As Boolean)
Dim i As Integer
Dim inactTrafic As Boolean

If cboPériode = "" Then Cancel = True: Exit Sub
Journal "Validate"

' Mémorisation de l'état initial de l'activation de la saisie du trafic, avant autorisation
inactTrafic = cmdChangeMode.Enabled
AutorTrafic True

With ActiveControl
  If .ListCount = 0 Then
    GiratoireProjet.newPériode DrapeauMenu:=False  ' False Indique que l'appel ne vient pas du menu
    Exit Sub
  End If
  
  
  For i = 1 To .ListCount
    If StrComp(.List(i - 1), .Text, vbTextCompare) = 0 Then
      .ListIndex = i - 1
      Exit Sub
    End If
  Next
  
  If MsgBox(IDm_CréePériode & " " & .Text, vbYesNo + vbQuestion) = vbYes Then
    GiratoireProjet.newPériode DrapeauMenu:=False  ' False Indique que l'appel ne vient pas du menu
  Else
    Cancel = True
    AutorTrafic inactTrafic   ' Remise à l'état initial de l'activation de la saisie du trafic
  End If
End With

End Sub

'******************************************************************************
' Changement de mode de Saisie des matrices de Trafic
'******************************************************************************
Private Sub cmdChangeMode_Click()

  Journal "Click"
  
  If controlChangeMode() Then
    fraTraficTout.Visible = Not fraTraficTout.Visible
    lblTraficUVP.Visible = Not lblTraficUVP.Visible
    If fraTraficTout.Visible Then
      cmdChangeMode.Caption = IDl_ModeUVP
      optTrafic(VL) = True
      TraficActif.setVéhicule VL
    Else
      cmdChangeMode.Caption = IDl_ModeVLPL2R
      vgdTrafic(VEHICULE).Enabled = True
      TraficActif.setVéhicule UVP
    End If
    TraficActif.Dimensionner NbBranches, cmdChangeMode.Caption
    'Faire état de la modification
    DetectModif
  End If
  
End Sub
'******************************************************************************
' Active et désactive les matrices de saisie lorsque l'utilisateur
' change d'onglet
'  paramètre : onglet = numéro de l'onglet appelé
'******************************************************************************
Private Sub ChangeTabStop(ByVal Onglet As Integer)
  Select Case Onglet
    Case 0: vgdCarBranche.TabStop = True
            vgdLargBranche.TabStop = False
            vgdTrafic.Item(0).TabStop = False
            vgdTrafic.Item(1).TabStop = False
            'Lorsqu'on changé d'onglet, il faut se replacer sur la matrice de caractéristiques
            'des branches car la présentation de celle-ci a pu être affecté par le déplacement
            'graphique des branches faite à partir des autres onglets
            With vgdCarBranche
              .Col = 1
              .Row = 1
              .Action = 0
            End With
    Case 1: vgdCarBranche.TabStop = False
            vgdLargBranche.TabStop = True
            vgdTrafic.Item(0).TabStop = False
            vgdTrafic.Item(1).TabStop = False
            
    Case 2: vgdCarBranche.TabStop = False
            vgdLargBranche.TabStop = False
            vgdTrafic.Item(0).TabStop = True
            vgdTrafic.Item(1).TabStop = True
  End Select
End Sub

'******************************************************************************
' Demande de confirmation de changement de mode (à développer)
'******************************************************************************
Private Function controlChangeMode() As Boolean
  controlChangeMode = (MsgBox(IDm_ReinTrafic, vbYesNo + vbQuestion) = vbYes)
End Function

'******************************************************************************
' Activation de la feuille
'******************************************************************************
Public Sub Form_Activate()
Dim i As Integer

  GiratoireProjet.Activate
  GriserMenus True
  MDIGirabase.mnuTrafic(8).Checked = DiagramFlux
  For i = 1 To GiratoireProjet.NbBranches
    MDIGirabase.mnuBranche(i - 1).Caption = "&" & CStr(i) & " " & GiratoireProjet.colBranches.Item(i).nom
  Next
  
    ' Aide contextuelle
  MDIGirabase.HelpContextID = HelpContextID

  Journal "Activation"
  
End Sub

Private Sub Form_Click()
'  MDIGirabase.HelpContextID = IDhlp_Graphique
End Sub

Private Sub Form_DblClick()

  'Affichage des caractéristiques de la branche sélectionnée
  If monNumBrancheSelect > 0 Then
    VerifieDonnée
    frmCarBranche.Show vbModal
  End If
End Sub

Private Sub Form_Deactivate()

  ' Aide contextuelle
  HelpContextID = MDIGirabase.HelpContextID

  If TypeOf ActiveControl Is vaSpread Then
    ' Résoud un pb possible d'affichage lors qu'on passe de la fenêtre résultats à la fenêtre données
    With ActiveControl
      If .CellType = 10 Then  ' Pb en fait sur les check box
        .Col = 1
        .Action = 0
      End If
    End With
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
' Gestion du point décimal comme virgule
' Si l'utilisateur est ainsi configuré, on détecte la frappe du point décimal
' mais seule la fonction KeyPress semble en mesure de réafficher la virgule
Dim i As Integer

  If KeyCode = vbKeyDecimal And Shift = 0 Then
    FlagPtDecimal = True
  End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If FlagPtDecimal Then
    KeyAscii = gbPtDecimal
    FlagPtDecimal = False
  End If
  If Not DonnéeModifiée Then
    'affecte la couleur normale au controle au premier caractère frappé
    ActiveControl.ForeColor = vbWindowText
  End If
  DonnéeModifiée = True
End Sub

'******************************************************************************
' ConstruitChampTexte
' Reconstruit un champ de saisie à partir du contrôle texte
'******************************************************************************
Private Function ConstruitChampTexte(ByVal KeyAscii As Integer) As String
  Dim ChaineTraitée As String
  ConstruitChampTexte = ""
  If KeyAscii <> 8 And KeyAscii <> 0 And TypeOf ActiveControl Is TextBox Then
    With ActiveControl
      'Constitution de la nouvelle chaine
      If .SelLength = 0 Then
        'Aucun texte n'est sélectionné
        If .SelStart = 0 Then
          ChaineTraitée = Chr(KeyAscii) & .Text
        Else
          If .SelStart >= Len(.Text) Then
            ChaineTraitée = .Text & Chr(KeyAscii)
          Else
            ChaineTraitée = Left(.Text, .SelStart) & Chr(KeyAscii) & Mid(.Text, .SelStart + 1)
          End If
        End If
      Else
        'Texte sélectionné
        ChaineTraitée = ""
        If .SelStart >= 1 Then ChaineTraitée = Left(.Text, .SelStart)
        ChaineTraitée = ChaineTraitée & Chr(KeyAscii)
        If .SelStart + .SelLength < Len(.Text) Then
         ChaineTraitée = ChaineTraitée & Mid(.Text, .SelStart + .SelLength + 1)
        End If
      End If
     End With
     ConstruitChampTexte = ChaineTraitée
    End If
    
  End Function
    
Function LimiteNbDécimales(ByRef ChaineTraitée As String, ByVal NbDécimales As Integer) As Boolean
Dim i As Integer
Dim Décimale As Boolean
  LimiteNbDécimales = False
  'Recherche de la décimale
  i = 1
  Décimale = False
  Do While i <= Len(ChaineTraitée) And Not Décimale
    If Mid(ChaineTraitée, i, 1) = "," Then
      Décimale = True
    Else
      i = i + 1
    End If
  Loop
  'Validation ou non du caractère frappé
  If Décimale Then
    If Len(ChaineTraitée) > i + NbDécimales Then
      LimiteNbDécimales = True
      'Conserve NbDécimales
      ChaineTraitée = Left(ChaineTraitée, i + NbDécimales)
    End If
  End If
End Function

'******************************************************************************
' Demande de fermeture de la feuille
'******************************************************************************
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim flag As Integer
Dim reponse As Integer

If flagErreurFatale Then Exit Sub

flag = vbYesNoCancel + vbQuestion
If FichierModifié Then 'And Not VersionDemo Then
  reponse = MsgBox(IDm_Enregistrer & " " & Me.Caption, flag)
Else
  reponse = vbNo
End If

Select Case reponse
Case vbYes
  ' La ligne suivante a été rajoutée le 28/11/2000(AV : bug possible si Quitter alors qu'un autre giratoire que celui à sauvegarder est actif)
  Set gbProjetActif = GiratoireProjet
  gbProjetActif.Enregistrer Cancel
Case vbCancel
  Cancel = True
End Select

End Sub

'******************************************************************************
' Déchargement de la feuille
'******************************************************************************
Private Sub Form_Unload(Cancel As Integer)
Dim i As Integer

  If Not GiratoireProjet.Résultats Is Nothing Then Unload GiratoireProjet.Résultats
  With gbGiratoires
    For i = 1 To .count
      If .Item(i) Is GiratoireProjet Then .Remove i: Exit For
    Next
    ' Maintenance : 28/11/2000 : Déclanchement de Class_Terminate suite à l'activation de la protection
    Set GiratoireProjet = Nothing
    Set gbProjetActif = Nothing
    If gbGiratoires.count = 0 Then GriserMenus False
  End With
  
End Sub

Private Sub fraTrafic_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
 If Button = vbRightButton And tabDonnées.Tab = 2 Then PopupMenu MDIGirabase.mnuBarre(2), , , , MDIGirabase.mnuTrafic(0)
End Sub

'************************************************************************************************************************
' Procédures lblLibelléBranche_xxx : Déportées dans le module de dessin de DessinGiratoire.Bas, sous le nom LabelBranche_xxx
'************************************************************************************************************************
Private Sub lblLibelléBranche_DblClick(Index As Integer)
  LabelBranche_DblClick Index
End Sub

Private Sub lblLibelléBranche_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
  LabelBranche_MouseDown Index, Button, Shift, x, Y
End Sub

Private Sub lblLibelléBranche_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
  LabelBranche_MouseMove Index, Button, Shift, x, Y
End Sub

Private Sub lblLibelléBranche_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
  Dessin_MouseUp Button, Shift, x, Y
End Sub

'******************************************************************************
' Chargement de la feuille :  Fonction Nouveau ou Ouvrir
'   Création de l'objet GIRATOIRE
'******************************************************************************
Private Sub Form_Load()
Dim i As Integer

Set gbProjetActif = GiratoireProjet

SetDeviceIndependentWindow Me

  'Matrice en cours de saisie
  TypeMatriceActive = AUCUN
  'Modifie la couleur de fond des cellules verrouillées
  vgdLargBranche.LockBackColor = vbGrayText
  vgdCarBranche.LockBackColor = vbGrayText
  If Nouveau Then
    With tabDonnées
      'Désactive les onglets Dimensionnements et Trafics
      .TabEnabled(1) = False
      .TabEnabled(2) = False
    End With
  Else
    If Not GiratoireProjet.Lire Then Exit Sub
  End If
  
  'Onglet actif
  With tabDonnées
    .Visible = True
    .Tab = 0
  End With
  
  GiratoireProjet.Création
  
  ' Branches
  NbBranches = GiratoireProjet.NbBranches
  'Unités d'angle
  lblAngleBranche.Caption = IDl_Angle & " (" & libelAngle(GiratoireProjet.modeangle) & ")"

  HelpContextID = IDhlp_OngletSite
  
  With vgdCarBranche
    .Col = 2
    For i = 1 To 8
      .Row = i
      .TypeIntegerMax = angConv(2 * PI, False) - 1
    Next
'    .HelpContextID = IDhlp_OngletSite
  End With
  
'  vgdLargBranche.HelpContextID = IDhlp_OngletDimensionnement
'  vgdTrafic(PIETON).HelpContextID = IDhlp_OngletTrafic
'  vgdTrafic(VEHICULE).HelpContextID = IDhlp_OngletTrafic
  
  DoEvents
  
  cmdChangeMode.Caption = IDl_ModeVLPL2R
  
  fraCarBranches.Height = vgdCarBranche.Top + vgdCarBranche.Height + lblNomBranche.Top / 2
  fraBranches.Height = vgdLargBranche.Top + vgdLargBranche.Height + lblLargeurs.Top / 2
  fraTrafic(VEHICULE).Height = fraQTS.Top + fraQTS.Height + lblTraficUVP.Top
  
  DessinerGiratoire IsPremierDessin:=True    ' True : Premier Dessin du Giratoire
  
End Sub

'************************************************************************************************************************
' Procédures Form_Mousexxx : Déportées dans le module de dessin de DessinGiratoire.Bas, sous le nom Dessin_Mousexxx
'************************************************************************************************************************
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Not gbProjetActif.Données.VerifieDonnée Then Exit Sub
    'Cas du bouton gauche traité uniquement
    If Button = vbLeftButton Then
      
      Dessin_MouseDown Button, Shift, x, Y
    ElseIf Button = vbRightButton Then
      PopupMenu MDIGirabase.mnuBarre(1), , , , MDIGirabase.mnuSite(3)
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

If Button = 1 Then Dessin_MouseMove Button, Shift, x, Y
Exit Sub

If x > gbDemiLargeur * 2 And MDIGirabase.HelpContextID = IDhlp_Graphique Then
  MDIGirabase.HelpContextID = SauvHelpContextId
ElseIf x < gbDemiLargeur * 2 And MDIGirabase.HelpContextID <> IDhlp_Graphique Then
  SauvHelpContextId = MDIGirabase.HelpContextID
  MDIGirabase.HelpContextID = IDhlp_Graphique
End If

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
  Dessin_MouseUp Button, Shift, x, Y
End Sub

'******************************************************************************
' Redimensionnement de la feuille
'******************************************************************************
Private Sub Form_Resize()
  'On ne fait pas de traitements lors d'une mise en icone
  If WindowState <> vbMinimized Then
    If MDIGirabase.Cascade Then Set gbProjetActif = GiratoireProjet
    'Pour éviter le problème de fermeture si plusieurs fenêtres sont agrandies
    If Not gbProjetActif Is GiratoireProjet Then Exit Sub
    'Calage en bas et tout le long de la largeur de la Form GiraBase
    'de la zone de message
    With fraInvite
      .Left = 0
      .Top = ScaleHeight - .Height
      .Width = ScaleWidth
      lblInvite.Width = .Width
    End With
    
    'Calage à droite tout le long de la hauteur de la form GiraBase
    With tabDonnées
      .Top = 0
      .Left = ScaleWidth - .Width
      .Height = ScaleHeight - fraInvite.Height
    End With
    
    If Not ChargementEnCours Then Redess
    
    'Stockage du centre de la zone graphique pour le prochain resize
    AncienXc = DemiLargeur
    AncienYc = DemiHauteur
  
  End If

End Sub

'******************************************************************************
' Validation du champ optMilieu
'******************************************************************************
Private Sub optMilieu_Click(Index As Integer)
  Dim i As Integer
  
  Journal "Click"
  
  'Active les onglets Dimensionnement et trafic
  tabDonnées.TabEnabled(1) = True
  tabDonnées.TabEnabled(2) = True
  'Vérifie si le changement de milieu est autorisé
  ControleMilieu
  GriserMenus True
End Sub
Private Sub optMilieu_Validate(Index As Integer, Cancel As Boolean)
  'ControleMilieu
End Sub
'******************************************************************************
' Vérifier si le changement de Milieu est autorisé
'******************************************************************************
Private Sub ControleMilieu()
  optMilieu(0).ForeColor = vbWindowText
  'Validation 'Rq0599
'  If ValidationLA(Numopt(optMilieu), txtLA, "") Then
'    GiratoireProjet.Milieu = Numopt(optMilieu)
'    ControleRecommandations False
'  Else
'    optMilieu(GiratoireProjet.Milieu).SetFocus
'    MsgBox IDv_ModifMilieu, vbExclamation
'  End If
  TypeControleActif = TYPE_MILIEU
  If ValidationDonnées(Numopt(optMilieu)) Then
    GiratoireProjet.MajComplément Numopt(optMilieu)
    ControleRecommandations False
  End If
End Sub

'******************************************************************************
' Choix d'une matrice de trafic (VL - PL  - 2R -UVP)
'******************************************************************************
Private Sub optTrafic_Click(Index As Integer)

Journal "Click"

TraficActif.setVéhicule (Index)
' Trafic UVP : à voir seulement
vgdTrafic(VEHICULE).Enabled = (Index <> UVP)

End Sub

Function VerifieDonnée(Optional ByVal PreviousTab As Integer = -1) As Boolean
  Dim Cancel As Boolean
  Dim NewTab As Integer
  Dim RepositionnerFocus As Boolean
  VerifieDonnée = True
  If controleEnCours Then Exit Function
  AutreOnglet = True
  Dim Matrice As Integer
  'Une matrice était en cours de saisie lorsque l'on a voulu changer d'onglet
  'Dans certains cas, l'événement _LeaveCell n'a pu être appelé
  'On appelle l'événement _LeaveCell
  NewTab = tabDonnées.Tab
  If TypeMatriceActive <> AUCUN Then
    'AutreOnglet = True
    ChangementOnglet = True
   'S'il y avait un contrôle non validé au moment du changement
    'd'onglet, on lance la vérification de la donnée
    controleEnCours = True 'pour éviter les appels récursifs lors des changements d'onglet
    'On repasse dans l'onglet précédent
    'On passera dans l'onglet souhaité seulement si la donnée est valide
    'et si elle n'est sujette à aucun avertissement
    If PreviousTab >= 0 Then tabDonnées.Tab = PreviousTab
    controleEnCours = False
    Matrice = TypeMatriceActive
    Select Case TypeMatriceActive
      Case DIMENSION:
      'Si on clique dans la partie graphique alors qu'on était sur le spread
      'on place d'abord le focus en dehors
      'du spread avant de lancer la procédure de vérification
      'Lorsqu'il ya demande de changement d'onglet le focus est dejà en dehors du
      'spread
      txtVariante.SetFocus
      vgdLargBranche_LeaveCell NuméroColonneActive, NuméroLigneActive, -1, -1, Cancel
      Case BRANCHE:
      txtNomGiratoire.SetFocus
      vgdCarBranche_LeaveCell NuméroColonneActive, NuméroLigneActive, -1, -1, Cancel
      Case TRAFIC:
          Dim IndexEnCours As Integer
          If TypeControleActif = TYPE_QP Then
            IndexEnCours = PIETON
          Else
            IndexEnCours = VEHICULE
          End If
          cboPériode.SetFocus
          'Rq19/05 Point1
          'vgdTrafic_LeaveCell IndexEnCours, NuméroColonneActive, 1, -1, -1, Cancel
          vgdTrafic_LeaveCell IndexEnCours, NuméroColonneActive, NuméroLigneActive, -1, -1, Cancel
    End Select
    controleEnCours = True
    'If Not MessageEmis And DonnéeValide Then
    If ChangementOnglet And DonnéeValide Then
      'Changement d'onglet
      tabDonnées.Tab = NewTab
      lblInvite = ""
    Else
      VerifieDonnée = False
    End If
    controleEnCours = False
  End If
  If Not (ControleActif Is Nothing) Then
    'Si le champ qui a le focus a fait l'objet d'une modification
    RepositionnerFocus = False
    If DonnéeModifiée Then
      'S'il y avait un contrôle non validé au moment du changement
      'd'onglet, on lance la vérification de la donnée
      controleEnCours = True 'pour éviter les appels récursifs lors des changements d'onglet
      'On repasse dans l'onglet précédent
      'On passera dans l'onglet souhaité seulement si la donnée est valide
      If PreviousTab >= 0 Then tabDonnées.Tab = PreviousTab
      If ValidationDonnées(ControleActif.Text) Then
        'Récupère le nom du controle
        Dim ControleActif2 As Control
        Set ControleActif2 = ControleActif
        If ControleRecommandations(False, TYPE_AVANT) Then
         'S'il y a eu un message on reste dans l'onglet initial
          RepositionnerFocus = True
          Set ControleActif = ControleActif2
          Set ControleActif2 = Nothing
        Else
          tabDonnées.Tab = NewTab
        End If
      Else
        VerifieDonnée = False
        RepositionnerFocus = True
        'On affecte la valeur précédente
        ControleActif.Text = SauveValeur
      End If
      If RepositionnerFocus Then
        'Pour ne pas que le clic sur le dessin du giratoire
        ' soit pris en compte
        VerifieDonnée = False
        'On réinitialise l'invite qui sera recalculée dans l'evénement GotFocus
        lblInvite = ""
        'La donnée ne peut être validée,
        'on  positionne le focus à nouveau sur celle-ci
        ControleActif.SetFocus
        '0599
        InitControle True
        ControleRecommandations True
        'Pour éviter de redéclencher l'événement click...
        'Set ControleActif = Nothing
        'Retourne sur l'onglet précédemment sélectionné
        'tabDonnées.Tab = PreviousTab
      End If
    End If
    controleEnCours = False
  End If
  ChangementOnglet = False
End Function

'******************************************************************************
' Onglet principal
'******************************************************************************
Private Sub tabDonnées_Click(PreviousTab As Integer)
  'Vérifie les dernières données saisies avant la demande de changement d'onglet
  If controleEnCours Then Exit Sub
  
  
  VerifieDonnée PreviousTab
  If PreviousTab = tabDonnées.Tab Then Exit Sub 'Sortie si on n'a pas changé d'onglet
  
  Journal "Click"
  
  ChangeTabStop (tabDonnées.Tab)
  Select Case tabDonnées.Tab
    Case 0
    ' Aide contextuelle
      MDIGirabase.HelpContextID = IDhlp_OngletSite
      
    Case 1
    ' Aide contextuelle
      MDIGirabase.HelpContextID = IDhlp_OngletDimensionnement
      txtNomGiratoire.SetFocus
    Case 2
    ' Aide contextuelle
      MDIGirabase.HelpContextID = IDhlp_OngletTrafic
      If cboPériode.ListCount = 0 Or cboPériode = "" Then
        AutorTrafic False
        cboPériode.SetFocus
      ElseIf cboPériode.ListIndex = -1 And cboPériode.ListCount > 0 Then
        cboPériode.ListIndex = 0
      End If
      'Contrôle d'un trafic éventuel existant avec une largeur d'entrée ou
      'de sortie nulle
      If ControleTrafic Then
        tabDonnées.Tab = PreviousTab
      'Contrôle d 'un trafic inexistant pour une voie d'entrée ou de sortie
      'non nulle
      Else
        ControleMatriceTrafic
        ControleValeursTrafic
      End If
  End Select
  
  tabDonnées.HelpContextID = MDIGirabase.HelpContextID
  
End Sub

Private Sub tabDonnées_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
  If Button = vbRightButton And tabDonnées.Tab = 2 Then PopupMenu MDIGirabase.mnuBarre(2), , , , MDIGirabase.mnuTrafic(0)
End Sub

'******************************************************************************
' autorTrafic : il faut au moins une période de créée
'******************************************************************************
Private Sub AutorTrafic(ByVal actif As Boolean)
'  vgdTrafic(PIETON).Enabled = actif
'  vgdTrafic(VEHICULE).Enabled = actif
  fraTraficTout.Enabled = actif
  cmdChangeMode.Enabled = actif
End Sub

Private Sub txtBf_Change()
   DonnéeModifiée = True
End Sub

Private Sub txtBf_GotFocus()
  Journal "GotFocus"
  
  InitControle True
  ControleRecommandations True
End Sub

'******************************************************************************
' txtBf : doit être numérique
'******************************************************************************
Private Sub txtBf_KeyPress(KeyAscii As Integer)
  ' On alerte un caractère non numérique
  KeyAscii = ControleChampRéel(KeyAscii)
End Sub

Private Sub txtBf_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

'******************************************************************************
' Validation du champ txtBf
'******************************************************************************
Public Sub txtBf_Validate(Cancel As Boolean)
  If DonnéeModifiée Then
    txtBf = FormateRéel(txtBf.Text)
    If ValidationDonnées(txtBf) Then
      ControleRecommandations False
      Journal "Validate"
    Else
      Cancel = True
      InitControle False
      'Pour remette à jour le graphique
      calculRg True
    End If
  End If
End Sub

Private Sub txtLA_Change()
   DonnéeModifiée = True
End Sub

Private Sub txtLA_GotFocus()
  InitControle True
  ControleRecommandations True
  Journal "GotFocus"
End Sub

'******************************************************************************
' txtLA : doit être numérique
'******************************************************************************
Private Sub txtLA_KeyPress(KeyAscii As Integer)
  'On alerte un caractère non numérique
  KeyAscii = ControleChampRéel(KeyAscii)
End Sub

Private Sub txtLA_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

'******************************************************************************
' Validation du champ txtLA
'******************************************************************************
Public Sub txtLA_Validate(Cancel As Boolean)
  If DonnéeModifiée Then
    txtLA = FormateRéel(txtLA.Text)
    If ValidationDonnées(txtLA) Then
      ControleRecommandations False
      Journal "Validate"
    Else
      Cancel = True
      InitControle False
      'Pour remette à jour le graphique
      calculRg True
    End If
  End If
End Sub

Private Sub txtLocalisation_GotFocus()
  InitControle True
  ControleRecommandations True
  'Pour ne pas contrôler ce champ...
  Set ControleActif = Nothing
  Journal "GotFocus"
End Sub

Private Sub txtLocalisation_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

'******************************************************************************
' Validation du champ txtLocalisation
'******************************************************************************
Private Sub txtLocalisation_Validate(Cancel As Boolean)
  If DonnéeModifiée Then
    GiratoireProjet.Localisation = txtLocalisation
    DetectModif False
  End If
  Journal "Validate"
End Sub

Private Sub txtNomGiratoire_Change()
  DonnéeModifiée = True
End Sub

'******************************************************************************
' Focus sur le contrôle txtNomGiratoire
'******************************************************************************
Private Sub txtNomGiratoire_GotFocus()
  InitControle True
  'Déclenche les premières vérifications des données
  'Pour mettre en rouge les valeurs erronées
  'Le focus est positionné sur ce champ lors de l'appel à Girabase
  ControleRecommandations True
  'Pour ne pas contrôler ce champ...
  Set ControleActif = Nothing
  Journal "GotFocus"
End Sub

Private Sub txtNomGiratoire_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

'******************************************************************************
' Validation du champ txtNomGiratoire
'******************************************************************************
Private Sub txtNomGiratoire_Validate(Cancel As Boolean)
  If DonnéeModifiée Then
    GiratoireProjet.nom = txtNomGiratoire
    DetectModif False
  End If
  Journal "Validate"
End Sub


Private Sub txtR_Change()
  DonnéeModifiée = True
End Sub

'******************************************************************************
'
'******************************************************************************
Private Sub txtR_GotFocus()
  InitControle True
  ControleRecommandations True
  Journal "GotFocus"
End Sub

'******************************************************************************
' txtR : doit être numérique
'******************************************************************************
Private Sub txtR_KeyPress(KeyAscii As Integer)
  ' On alerte un caractère non numérique
  KeyAscii = ControleChampRéel(KeyAscii)
End Sub
Private Function ControleChampRéel(ByVal KeyAscii As Integer) As Integer
  Dim Chaine As String
  '1606 Filtre les caractères CTRL V et CTRL C et CtrlX
  If KeyAscii <> 3 And KeyAscii <> 22 And KeyAscii <> 24 Then
    ControleChampRéel = ControleRéel(KeyAscii)
    If KeyAscii <> 0 Then
      Chaine = ConstruitChampTexte(KeyAscii)
      If LimiteNbDécimales(Chaine, 1) Then
        Beep
        ControleChampRéel = 0
      End If
    End If
  Else
    ControleChampRéel = KeyAscii
  End If
End Function
'******************************************************************************
' FormateRéel
'  Fonction transformant un nombre réel en nombre réel formaté
'  avec suppression des 0 inutiles
'******************************************************************************
Private Function FormateRéel(ByVal Texte As String) As String
  If Texte = "" Then
    'Une absence de valeur est considérée comme une valeur nulle
    'Valeur
    Texte = "0"
    FormateRéel = "0"
  ElseIf IsNumeric(Texte) Then
    '1606...Conserve une seule décimale pour les valeurs réelles
    FormateRéel = CStr(Round(CSng(Texte), 1))
  Else
    '1606 Texte inchangé (Implémentation du Ctrl C, Ctrl V)
    FormateRéel = Texte
  End If
End Function

Private Sub txtR_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

'******************************************************************************
' Contrôle du champ txtR
'******************************************************************************
Public Sub txtR_Validate(Cancel As Boolean)
  If DonnéeModifiée Then
    txtR = FormateRéel(txtR.Text)
    If ValidationDonnées(txtR) Then
      ControleRecommandations False
      Journal "Validate"
    Else
      Cancel = True
      InitControle False
      'Pour remette à jour le graphique
      calculRg True
    End If
  End If
End Sub
'******************************************************************************
' Calcul du Rayon extérieur du  giratoire
'******************************************************************************
Public Sub calculRg(ByVal Dessiner As Boolean)
  txtRg = CDbl(txtR) + CDbl(txtLA)
  If txtBf <> "" Then txtRg = txtRg + CDbl(txtBf)
  GiratoireProjet.MajComplément
  
  If Dessiner Then DessinerGiratoire IsPremierDessin:=False        ' False : Giratoire déjà dessiné

End Sub

Private Sub txtVariante_Change()
  DonnéeModifiée = True
End Sub

Private Sub txtVariante_GotFocus()
  InitControle True
  Journal "GotFocus"
End Sub

Private Sub txtVariante_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtVariante_Validate(Cancel As Boolean)
  If DonnéeModifiée Then
    GiratoireProjet.NomVariante = txtVariante
    DetectModif False
  End If
  Journal "Validate"
End Sub

Private Sub vgdCarBranche_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
  'En lecture de fichier la variable NuméroLigneActive doit être alimentée
  'car il n'y a pas eu de passage par le GotFocus du spread
  NuméroLigneActive = Row
  DonnéeModifiée = True
  If Col = 4 Then
    GiratoireProjet.colBranches.Item(Row).Rampe = ButtonDown
  ElseIf Col = 5 Then
    GiratoireProjet.colBranches.Item(Row).TAD = ButtonDown
  End If
End Sub

Private Sub vgdCarBranche_Change(ByVal Col As Long, ByVal Row As Long)
' Ajout AV - 14.04.99 : si on fait bouger l'écart avec l'UpDown (incrément à double flèche), le drapeau n'est pas armé
  If Col = 3 Then DonnéeModifiée = True
  
End Sub

'******************************************************************************
' On entre dans la grille de saisie des branches
' Soit en amenant le focus à l'intérieur de la grille de saisie alors que
' celui-ci vient de l'extérieur de la grille
' soit lorsque l'on vient de modifier graphiquement une branche
'******************************************************************************
Private Sub vgdCarBranche_GotFocus()
Dim i As Integer
  If controleEnCours Then Exit Sub
  
  Journal "GotFocus"
  TypeMatriceActive = BRANCHE
  vgdCarBranche.OperationMode = OperationModeNormal
  ReDim sauvAngle(1 To NbBranches) As Single
  With GiratoireProjet.colBranches
    For i = 1 To NbBranches
      sauvAngle(i) = .Item(i).Angle
    Next
  End With
  
  AfficheSpreadNormal
  If BrancheSélectée > 0 Or AutreOnglet Then
    vgdCarBranchePrepare vgdCarBranche.ActiveRow, vgdCarBranche.ActiveCol
    If BrancheSélectée = 1 Then
      With vgdCarBranche
        'La première branche est sélectionnée graphiquement.
        'On positionne le focus sur le nom de la branche car l'angle
        'n'est pas modifiable
        .Col = 1
        .Action = 0
        .SetFocus
        'L'appel de la procédure suivante replace le focus sur la bonne cellule puis
        'affiche l'invite correspondant à la branche sélectée
        vgdDéplaceFocus 1, 2
      End With
    End If
    AutreOnglet = False
  Else
    PlaceFocus vgdCarBranche
    vgdCarBranchePrepare 1, 1
  End If
End Sub

'******************************************************************************
' vgdDéplaceFocus est appelée par SelectBranche dans DessinGiratoire.Bas
'******************************************************************************
Public Sub vgdDéplaceFocus(ByVal Row As Long, ByVal Col As Long)
    TypeMatriceActive = BRANCHE
    vgdCarBranchePrepare Row, Col
End Sub

'******************************************************************************
'
'******************************************************************************
Private Sub PlaceFocus(vgdSpread As vaSpread, Optional ByVal Col As Integer = 1, _
  Optional ByVal Row As Integer = 1)
  With vgdSpread
    .Row = Row
    .Col = Col
    .Action = 0
  End With
End Sub

'******************************************************************************
' Grille de saisie des caractéristiques des branches
' Prépare l 'invite en fonction de la nouvelle cellule active
'   paramètres : NewRow, NewCol : ligne et colonne de destination
'******************************************************************************
Public Sub vgdCarBranchePrepare(ByVal NewRow As Long, ByVal NewCol As Long)
  With vgdCarBranche
    'Sauvegarde l'ancienne valeur
    .Col = NewCol
    .Row = NewRow
    SauveValeurSpread = .Value
    Journal "***Prepare" & .Value
  End With
  NuméroLigneActive = NewRow
  NuméroColonneActive = NewCol
  Select Case NewCol
      Case 1: TypeControleActif = TYPE_COURANT
      Case 2: TypeControleActif = TYPE_ANGLE
      Case 3: TypeControleActif = TYPE_ANGLE
      Case 4: TypeControleActif = TYPE_COURANT
      Case 5: TypeControleActif = TYPE_COURANT
  End Select
  ControleRecommandations True, TYPE_MATRICE
End Sub

'******************************************************************************
' Validation d'une cellule de la grille Caractéristiques des branches
'******************************************************************************
Public Sub vgdCarBranche_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
Dim wValeur As Single, Ecart As Single, i As Integer
Dim wBranche As BRANCHE
  
  If controleEnCours Then Exit Sub
  
  
  TypeMatriceActive = BRANCHE
  With vgdCarBranche
    .Col = Col
    .Row = Row
  End With
  Journal "LeaveCell", NewCol, NewRow
  
  Set wBranche = GiratoireProjet.colBranches.Item(Row)
  DonnéeValide = True
  If DonnéeModifiée Then
    If Col = 2 Then
      DonnéeValide = ValidationDonnées(vgdCarBranche.Value, wBranche)
    ElseIf Col = 3 Then
      Dim Angle As Single
      Angle = GiratoireProjet.colBranches.Item(Row - 1).Angle + vgdCarBranche.Value
      DonnéeValide = ValidationDonnées(Angle, wBranche)
    End If
  End If
  If DonnéeValide Then
    'La donnée saisie a été validée
    If DonnéeModifiée Then
    Select Case Col
    Case 1    ' Nom
      wBranche.nom = vgdCarBranche.Value
      lblLibelléBranche(Row) = wBranche.nom
      ControleRecommandations False, TYPE_AUCUN 'Seulement pour l'invite
      DéplacerNomBranche lblLibelléBranche(Row), linBranche(Row), Cos(Angle), -Sin(Angle)        ' "-" pour le sinus : car l'axe des Y est vers le bas
      MDIGirabase.mnuBranche(Row - 1).Caption = "&" & CStr(Row) & " " & wBranche.nom
    Case 2    ' Angle
      With vgdCarBranche
        If .Value = "" Then
          MsgBox IDm_Obligatoire, vbOKOnly + vbExclamation
          Cancel = True
          Journal "Cancel"
          Exit Sub
        End If
      
        wValeur = CSng(.Value)
        wBranche.Angle = wValeur
        If Row > 1 Then
          Set wBranche = GiratoireProjet.colBranches.Item(Row - 1)
          Ecart = wValeur - wBranche.Angle
          .Col = 3
          If Ecart > 0 Then
            .Value = Ecart
          Else
            .Value = ""
          End If
          If Row < NbBranches Then
            Set wBranche = GiratoireProjet.colBranches.Item(Row + 1)
            .Row = Row + 1
            Ecart = wBranche.Angle - wValeur
            If Ecart > 0 Then
              .Value = Ecart
            Else
              .Value = ""
            End If
          End If
        End If
      End With
      
      monNumBrancheSelect = Row
      'Modifier la branche graphiquement sans spécifier l'invite
      ModifierBranche angConv(wValeur, True), False
      If DiagramFlux And Not TraficActif Is Nothing Then cLS: TraficActif.CalculDiagramFlux
          'Controle de l'écart avec l'angle précédent
      ControleRecommandations False, TYPE_ANGLE
    
    Case 3    ' Ecart
      If vgdCarBranche.Value <> "" Then
        Ecart = CSng(vgdCarBranche.Value)
        With vgdCarBranche
          .Col = 2
          .Value = GiratoireProjet.colBranches.Item(Row - 1).Angle + Ecart
          wBranche.Angle = .Value

          monNumBrancheSelect = Row
          ModifierBranche angConv(wBranche.Angle, True), False
          If DiagramFlux And Not TraficActif Is Nothing Then cLS: TraficActif.CalculDiagramFlux
          If Row < NbBranches Then
            wValeur = .Value
            Set wBranche = GiratoireProjet.colBranches.Item(Row + 1)
            .Row = Row + 1
            .Col = 3
            Ecart = wBranche.Angle - wValeur
            If Ecart > 0 Then
              .Value = Ecart
            Else
              .Value = ""
              .Col = 2
              .Action = 0
            End If
          End If
        'Controle de l'écart avec l'angle précédent
        ControleRecommandations False, TYPE_ANGLE
        End With
      End If
    
    Case 4      ' Rampe
      ControleRecommandations False, TYPE_AUCUN 'Seulement pour l'invite
  
    Case 5      ' Tourne a droite
      ControleRecommandations False, TYPE_AUCUN 'Seulement pour l'invite
      If DiagramFlux And Not TraficActif Is Nothing Then cLS: TraficActif.CalculDiagramFlux

    End Select
      
    End If
    Journal "***LeaveCell"
    If NewRow = -1 Then
      ' Traitement final : sortie de la grille de saisie
      'Rq19/09
      Journal "***LeaveCellNewRow-1"
      If MessageEmis And ChangementOnglet Then
        'Réactiver la cellule s'il y a eu une demande de changement d'onglet
        'et qu'un message d'avertissement doit être émis
        Cancel = True
        PlaceFocus vgdCarBranche, Col, Row
        vgdCarBranchePrepare Row, Col
        vgdCarBranche.SetFocus
        ChangementOnglet = False
        Journal "***LeaveCellChangeOnglet"
      Else
        'Interdire de modifier les branches, et déterminer le champ de saisie suivant
        BloqueNbBranches 0
        'La vérification est terminée
        TypeMatriceActive = AUCUN
        DonnéeValide = True
      End If
    Else
      'La valeur est valide
      'et l'utilisateur n'a pas tenté de sortir de la grille de saisie
      'Réactiver la cellule
      PlaceFocus vgdCarBranche, NewCol, NewRow
      vgdCarBranchePrepare NewRow, NewCol
      If FeuilleBranche Is Nothing Then
        vgdCarBranche.SetFocus
      End If
    End If
  Else
    'La valeur n'a pas été validée, le focus reste ou revient à sa position
    AutreOnglet = True
    'récupérer l'ancienne valeur
    vgdCarBranche.Value = SauveValeurSpread
    Journal "***Leavecell" & vgdCarBranche.Value
    'réactiver la cellule
    Cancel = True
    PlaceFocus vgdCarBranche, Col, Row
    vgdCarBranchePrepare Row, Col
    If FeuilleBranche Is Nothing Then
      vgdCarBranche.SetFocus
    End If
  End If
  
  If Cancel Then Journal "Cancel"
  
  DoEvents
End Sub

'******************************************************************************
' Bloquage du champ txtNbBranches  et détermination du champ de saisie suivant
'******************************************************************************
Private Sub BloqueNbBranches(ByVal numGrille As Integer)
Dim ControleSuivant As Control
  ' Si l'utilisateur est arrivé en bout de grille avec les touches de tabulation,
  ' activation du champ qui suit (DrapeauSuivant = True) ou qui précède
  If Débordement Then
    Débordement = False
    Select Case numGrille
    ' Onglet site
    Case 0
      If DrapeauSuivant Then
        Set ControleSuivant = txtNomGiratoire
      Else
        If GiratoireProjet.Milieu >= 0 Then
          Set ControleSuivant = optMilieu(GiratoireProjet.Milieu)
        Else
          Set ControleSuivant = txtLocalisation
        End If
      End If
    ' Onglet Dimensionnement
    Case 1
      If DrapeauSuivant Then
        Set ControleSuivant = txtVariante 'txtR
      Else
        Set ControleSuivant = txtLA
      End If
    ' Onglet Trafic
    Case 2
      Set ControleSuivant = vgdTrafic(PIETON)
    Case 3
      Set ControleSuivant = vgdTrafic(VEHICULE)
    End Select
    ControleSuivant.SetFocus
  End If
      
End Sub


'******************************************************************************
' Controle de validité des branches
' Paramètres
'   Numbranche  =0   => on contrôle toutes les branches
'               >0   => On contrôle seulement la branche n
' Retour
'    = 0 : pas d'erreur
'    > 0 : erreur entre la branche n et n-1
'         (on arrête sur la première erreur rencontrée)
'******************************************************************************
''Private Function controleChevauchementBranches(ByVal EcritMessage As Boolean, Optional ByVal NumBranche As Integer = 0) As Integer
''  Dim i, j As Integer
''  Dim fin As Integer
''  Dim Ecart, valeur As Single
''  controleChevauchementBranches = 0
''  With GiratoireProjet.colBranches
''    If NumBranche = 0 Then
''      i = 1
''      fin = .count
''    Else
''      i = NumBranche
''      fin = NumBranche
''    End If
''
''    Do While i <= fin And controleChevauchementBranches = 0
''      j = i Mod .count + 1
''      valeur = .Item(i).LI / 2 + .Item(i).LE4m + .Item(j).LS + .Item(j).LI / 2
''        Ecart = angConv(.Item(i).Angle - .Item(j).Angle, True)
''      If Ecart < 0 Then Ecart = Ecart + 2 * PI
''      If valeur > Ecart * txtRg Then
''        'Chevauchement pour la branche i
''        controleChevauchementBranches = i
''        If EcritMessage Then
''          Dim MessageAEcrire As String
''          MessageAEcrire = IDv_Chevauchement + .Item(i).nom + IDl_ET + .Item(j).nom & "."
''
''          Select Case TypeControleActif
''            Case TYPE_LE4M:
''              AfficheRecommandations TypeControleActif, MessageAEcrire, TYPE_LE4M
''            Case TYPE_LS:
''              AfficheRecommandations TypeControleActif, MessageAEcrire, TYPE_LS
''            Case TYPE_LI:
''              AfficheRecommandations TypeControleActif, MessageAEcrire, TYPE_LI
''          End Select
''        End If
''      End If
''      i = i + 1
''    Loop
''  End With
''End Function

'******************************************************************************
' Détection de la frappe de la touche TAB ou SHIFT TAB
'******************************************************************************
Private Sub vgdCarBranche_QueryAdvance(ByVal AdvanceNext As Boolean, Cancel As Boolean)
  Cancel = False
  Débordement = True
  DrapeauSuivant = AdvanceNext
End Sub

Public Sub vgdLargBrancheClic(ByVal Lig As Integer, ByVal ButtonDown As Boolean)
  NuméroLigneActive = Lig
  vgdLargBranche.Row = Lig
  vgdLargBranche.Col = 5
  TypeMatriceActive = DIMENSION
  vgdLargBranche.Value = ButtonDown
  vgdLargBrancheEntrée ButtonDown
End Sub

Private Sub vgdLargBranche_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
  vgdLargBrancheEntrée ButtonDown
  DonnéeModifiée = True
End Sub

'******************************************************************************
' Autorise ou interdit l'accès à la largeur d'entrée évasée
' suivant que la case à cocher Entrée Evasée est cochée ou non.
'******************************************************************************
Private Sub vgdLargBrancheEntrée(ByVal ButtonDown As Integer)
  vgdLargBranche.Col = 2
  If ButtonDown = 0 Then
    vgdLargBranche.Value = ""
    vgdLargBranche.Lock = True
    If TypeMatriceActive = DIMENSION Then
      GiratoireProjet.colBranches.Item(NuméroLigneActive).EntréeEvasée = False
      ControleRecommandations False, DIMENSION
    End If
  Else
    vgdLargBranche.Lock = False
  '0604
   ' vgdLargBranche.Value = GiratoireProjet.colBranches.Item(NuméroLigneActive).LE4m
    If TypeMatriceActive = DIMENSION Then
      With GiratoireProjet.colBranches.Item(NuméroLigneActive)
        '0604
        If Not .EntréeEvasée Then
          vgdLargBranche.Value = .LE4m
          .EntréeEvasée = True
          .LE15m = .LE4m
        End If
      End With
      ControleRecommandations False, DIMENSION
    End If
  End If
  lblInvite = Idi_Défaut
  If Not ChargementEnCours Then DonnéeModifiée = True
End Sub

'******************************************************************************
'
'******************************************************************************
Private Sub vgdLargBranche_GotFocus()
  Dim TypeVariableCourant As String
  If controleEnCours Then Exit Sub
  
  Journal "GotFocus"
  
  With vgdLargBranche
    .OperationMode = OperationModeNormal
    TypeMatriceActive = DIMENSION
    AfficheSpreadNormal 'retire la poignée
    'Prépare l'invite de la première cellule
    If AutreOnglet Then
      vgdLargBranchePrepare .ActiveRow, .ActiveCol
      AutreOnglet = False
    Else
      PlaceFocus vgdLargBranche
      vgdLargBranchePrepare 1, 1
    End If
  End With
End Sub

'******************************************************************************
' Prépare l'invite en fonction de la nouvelle cellule active
'******************************************************************************
Private Sub vgdLargBranchePrepare(ByVal NewRow As Long, ByVal NewCol As Long)
  With vgdLargBranche
    'Sauvegarde l'ancienne valeur
    .Col = NewCol
    .Row = NewRow
    SauveValeurSpread = .Value
    lblInvite = ""
    NuméroColonneActive = NewCol
    NuméroLigneActive = NewRow
    Select Case NewCol
      Case 1: TypeControleActif = TYPE_LE4M
      Case 2: TypeControleActif = TYPE_LE15M
      Case 3: TypeControleActif = TYPE_LI
      Case 4: TypeControleActif = TYPE_LS
      Case 5: TypeControleActif = TYPE_ENTREE
    End Select
    ControleRecommandations True, TYPE_MATRICE
  End With

End Sub

Private Sub vgdLargBranche_KeyPress(KeyAscii As Integer)
' Suppression AV 22/02/2000 - V4.0.18 suite à gestion du point décimal sur le Spread (GIRATOIRE.Création)
'  KeyAscii = ControleRéel(KeyAscii)
End Sub
'******************************************************************************
' ChangeLE4m
' La valeur de LE4m a changé
' D'autres cellules peuvent changer d''état
' Si la largeur d'entrée est nulle, il faut bloquer la largeur d'ilot
' l'entrée évasée et le tourne à droite
' Dans le cas contraire, il faut autoriser la saisie de contrôles qui pouvaient
' précédemment être verrouillés
'******************************************************************************
Public Sub ChangeLE4m(ByVal Row As Long, ByVal EntréeNulle As Boolean)
  With vgdLargBranche
    If EntréeNulle Then
      'Interdit la saisie de l'entrée évasée...
      .Row = Row
      .Col = 2
      .Value = ""
      .Lock = True
      'Interdit la largeur d'ilot LI...
      .Col = 3
      .Value = ""
      GiratoireProjet.colBranches.Item(Row).LI = 0
      .Lock = True
      'Interdit de cocher l'entrée évasée
      Dim TypeMatriceActiveA As Integer
      TypeMatriceActiveA = TypeMatriceActive
      TypeMatriceActive = AUCUN
      .Col = 5
      GiratoireProjet.colBranches.Item(Row).EntréeEvasée = False
      .Lock = True
      .Value = False
      TypeMatriceActive = TypeMatriceActiveA
      'Interdit le Tourne-A-Droite...
      GiratoireProjet.colBranches.Item(Row).TAD = False
      With vgdCarBranche
        .Row = Row
        .Col = 5
        .Value = False
        .Lock = True
      End With
    Else
      'Autorise l'entrée de la largeur d'ilot LI si la sortie n'est pas nulle
      .Row = Row
      .Col = 3
      If GiratoireProjet.colBranches.Item(Row).SortieNulle Then
        .Value = 0
      Else
        .Lock = False
        .Value = DEFAUT_LI 'remise à la valeur par défaut de LI...
        GiratoireProjet.colBranches.Item(Row).LI = DEFAUT_LI
      End If
      'Autorise de cocher l'entrée évasée
      .Col = 5
      .Lock = False
      'Autorise l'entrée du Tourne-A-Droite...
      If GiratoireProjet.R > 0 Then
        vgdCarBranche.Row = Row
        vgdCarBranche.Col = 5
        vgdCarBranche.Lock = False
      End If
    End If
    .Col = 1
  End With
End Sub

Public Sub ChangeLS(ByVal Row As Long, ByVal SortieNulle As Boolean)
  'Si la largeur de sortie est nulle, il faut bloquer la saisie
  'de la largeur d'ilot et lui imposer une valeur nulle
  With vgdLargBranche
    .Row = Row
    .Col = 3
    If SortieNulle Then
      .Lock = True
      .Value = ""
      GiratoireProjet.colBranches.Item(Row).LI = 0
    Else
      'Autorise l'entrée de LI si l'entrée n'est pas nulle
      If GiratoireProjet.colBranches.Item(Row).EntréeNulle Then
        .Value = ""
      Else
        .Lock = False
        .Value = DEFAUT_LI 'remise à la valeur par défaut de LI
        GiratoireProjet.colBranches.Item(Row).LI = DEFAUT_LI
      End If
    End If
    .Col = 4
  End With
End Sub

'******************************************************************************
' Validation d'une cellule de la grille Largeur des branches
'******************************************************************************
Public Sub vgdLargBranche_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
Dim wValeur As Single, Ecart As Single, i As Integer
Dim wBranche As BRANCHE
Dim Valeur As String
'Dim DonnéeValide As Boolean
Dim ESPrécédente As Boolean
  If controleEnCours Then Exit Sub
  
    
  TypeMatriceActive = DIMENSION
  With vgdLargBranche
    .Col = Col
    .Row = Row
  End With
  
  Journal "LeaveCell", NewCol, NewRow
  
  Set wBranche = GiratoireProjet.colBranches.Item(Row)
  Valeur = vgdLargBranche.Value
  RemplaceVirgule Valeur
  If DonnéeModifiée Then
    DonnéeValide = ValidationDonnées(Valeur, wBranche)
  Else
    DonnéeValide = True
  End If
  If DonnéeValide Then
    'Donnée valide
    If DonnéeModifiée Then
      'Affectation de la valeur à la variable associée à la cellule
      Select Case Col
      Case 1:
              wBranche.LE4m = Valeur
              ESPrécédente = wBranche.EntréeNulle
              wBranche.EntréeNulle = (CSng(Valeur) = 0#)
              If ESPrécédente <> wBranche.EntréeNulle Then
                ChangeLE4m Row, wBranche.EntréeNulle
                'Interdit le déplacement sur la même ligne du spread
                If NewRow = Row Then NewCol = 1
              End If
      Case 2: wBranche.LE15m = Valeur
      Case 3: wBranche.LI = Valeur
      Case 4: wBranche.LS = Valeur
              ESPrécédente = wBranche.SortieNulle
              wBranche.SortieNulle = (CSng(Valeur) = 0#)
              If ESPrécédente <> wBranche.SortieNulle Then
                ChangeLS Row, wBranche.SortieNulle
              End If
      Case 5: wBranche.EntréeEvasée = Valeur
      End Select
      
      ControleRecommandations False, TYPE_MATRICE
      DessinerBranche Row, angConv(wBranche.Angle, CVRADIAN)
      If DiagramFlux And Not TraficActif Is Nothing Then cLS: TraficActif.CalculDiagramFlux
    End If
    
    If NewRow = -1 Then
      If MessageEmis And ChangementOnglet Then
        'Réactiver la cellule s'il y a eu une demande de changement d'onglet
        'et qu'un message d'avertissement doit être émis
        Cancel = True
        PlaceFocus vgdLargBranche, Col, Row
        vgdLargBranchePrepare Row, Col
        vgdLargBranche.SetFocus
        ChangementOnglet = False
      Else
      '0599
        ''controleChevauchementBranches (True)
        vgdLargBranche.EditMode = True
        vgdLargBranche.Refresh
        'Fin de vérification
        TypeMatriceActive = AUCUN
        'Traitement final : sortie de la grille de saisie
        BloqueNbBranches 1
      End If
    Else
      'Repositionnement sur la nouvelle cellule
      vgdLargBranchePrepare NewRow, NewCol
      'Réactiver la cellule
      PlaceFocus vgdLargBranche, NewCol, NewRow
      If FeuilleBranche Is Nothing Then
        vgdLargBranche.SetFocus
      End If
    End If
  Else
    'Récupérer l'ancienne valeur
    AutreOnglet = True
    vgdLargBranche.Value = SauveValeurSpread
    'Réactiver la cellule
    Cancel = True
    PlaceFocus vgdLargBranche, Col, Row
    vgdLargBranchePrepare Row, Col
    If FeuilleBranche Is Nothing Then
      vgdLargBranche.SetFocus
    End If
  End If
  DoEvents
End Sub

' Modif AV : 09/03/99 : le Spread retourne toujours un point (.) dans Value pour le point décimal
' Il faut remettre ce qu'il faut pour que la fonction IsNumeric fonctionne correctement
Private Sub RemplaceVirgule(ByRef Chaine As String)
  Dim Position As Long
  Position = InStr(1, Chaine, ".", 1)
  If Position Then
    Mid(Chaine, Position) = Chr(gbPtDecimal)  ' = ","
  End If
End Sub

'******************************************************************************
' Détection de la frappe de la touche TAB ou SHIFT TAB
'******************************************************************************
Private Sub vgdLargBranche_QueryAdvance(ByVal AdvanceNext As Boolean, Cancel As Boolean)
  Cancel = False
  Débordement = True
  DrapeauSuivant = AdvanceNext
  
End Sub

Private Sub vgdTrafic_GotFocus(Index As Integer)
  'Si l'on n'est pas sorti de la matrice alors qu'un événement GotFocus
  'survient il faut quitter cette fonction
  'C'est le cas lorsque la dernière cellule d'une grille ne peut être validée
  'Un gotfocus est renvoyé mais il ne faut pas positionner le focus
  'sur la première cellule de la grille
  If controleEnCours Then Exit Sub
    
  Journal "GotFocus"
  
  With vgdTrafic(Index)
    .OperationMode = OperationModeNormal
     TypeMatriceActive = TRAFIC
    AfficheSpreadNormal 'Retire la poignée
    If AutreOnglet Then
      'Le focus était positionné sur un autre onglet et l'utilisateur a cliqué
      'sur le spread
      vgdTraficPrepare .ActiveRow, .ActiveCol, Index
      AutreOnglet = False
    Else
      'Le focus était sur un autre contrôle appartenant à l'onglet actif
      'L'utilisateur a cliqué dans une cellule du spread ou a frappé
      ' un touche (TAB...) qui mène à ce spread
      'Dans tous les cas on se positionne sur la première cellule accessible
      'Positionnement sur la première cellule si celle-ci est accessible
      'Le clic sera interprété par le spread qui positionnera le focus à
      'l'endroit voulu
      Dim Bloqué As Boolean
      .Col = 1
      .Row = 1
      Bloqué = .Lock
      'Positionnement sur la première cellule accessible
      If Bloqué Then
        Dim ColEnabled As Integer, RowEnabled As Integer
        Dim Continue As Boolean
        Continue = True
        RowEnabled = 1
        Do While Continue And RowEnabled <= NbBranches
          .Row = RowEnabled
          ColEnabled = 1
          Do While Continue And ColEnabled <= NbBranches
            .Col = ColEnabled
            If .Lock Then
              ColEnabled = ColEnabled + 1
            Else
              Continue = False
            End If
          Loop
          If Continue Then RowEnabled = RowEnabled + 1
        Loop
        PlaceFocus vgdTrafic(Index), ColEnabled, RowEnabled
        vgdTraficPrepare RowEnabled, ColEnabled, Index
        
        Dim Cancel As Boolean
        vgdTrafic_LeaveCell Index, 1, 1, ColEnabled, RowEnabled, Cancel
      Else
        'Prépare l'invite de la première cellule
        PlaceFocus vgdTrafic(Index), 1, 1
        vgdTraficPrepare 1, 1, Index
      End If
    End If
  End With
End Sub

Private Sub vgdTraficPrepare(ByVal NewRow As Long, ByVal NewCol As Long, ByVal Index As Integer)
  With vgdTrafic(Index)
    .Col = NewCol
    .Row = NewRow
    'Sauvegarde l'ancienne valeur
    SauveValeurSpread = .Value
    NuméroColonneActive = NewCol
    NuméroLigneActive = NewRow
    TypeControleActif = Index
    Select Case Index
      Case PIETON: TypeControleActif = TYPE_QP
      Case VEHICULE: TypeControleActif = TYPE_Q
    End Select
    'Affiche les messages de recommandations et affiche l'invite
    'pour la nouvelle cellule
    ControleRecommandations True, TYPE_MATRICE
  End With
End Sub

Private Sub vgdTrafic_KeyPress(Index As Integer, KeyAscii As Integer)
  'La donnée pouvait être affiché en rouge ;
  'toute nouvelle frappe repasse la valeur dans la couleur normale
  vgdTrafic.Item(Index).ForeColor = vbWindowText
End Sub

'******************************************************************************
' Validation d'une cellule de la grille Matrice de Trafics
'******************************************************************************
Private Sub vgdTrafic_LeaveCell(Index As Integer, ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
Dim wBranche As Integer
Dim Valeur As String
'Dim DonnéeValide As Boolean
  
  If TraficActif Is Nothing Then Exit Sub
  If controleEnCours Then Exit Sub
  
  TypeMatriceActive = TRAFIC
  With vgdTrafic(Index)
    .Col = Col
    .Row = Row
    Valeur = .Value
    
    Journal "LeaveCell", NewRow, NewCol
    
    If Valeur = "" Or Not DonnéeModifiée Then
      DonnéeValide = True
      Valeur = -1
    Else
      DonnéeValide = ValidationDonnées(Valeur)
    End If
    If DonnéeValide Then
      If DonnéeModifiée Then TraficModifié = True
      If DonnéeModifiée Then
      'Affectation de la valeur à la variable associée à la cellule
        Select Case Index
          Case VEHICULE:
            TraficActif.setQ Row, Col, .Value, True   ' True = pour faire afficher les totaux dans la grille
          Case PIETON:
            TraficActif.setQP Col, .Value
        End Select
        ControleRecommandations False, TYPE_MATRICE
        If DiagramFlux And Index = VEHICULE Then cLS: TraficActif.CalculDiagramFlux
        TraficActif.InvalideCalcul
      End If
      ' Traitement final : sortie de la grille de saisie
      If NewRow = -1 Then
        'Dans le cas d'émission d'un message et si on a tenté de positionner le focus
        'sur la matrice de Piéton/trafic ce dernier ne se repositionne pas correctement
        'On préfère faire disparaître le focus, ce qui permet d'éviter le problème
        If MessageEmis Then txtNomGiratoire.SetFocus
        If MessageEmis And ChangementOnglet Then
          'Réactive la cellule s'il y a eu une demande de changement d'onglet
          'et qu'un message d'avertissement a été émis
          Cancel = True
          PlaceFocus vgdTrafic(Index), Col, Row
          vgdTraficPrepare Row, Col, Index
          vgdTrafic(Index).SetFocus
          ChangementOnglet = False
        Else
          'Sortie de la matrice
          .EditMode = True
          TypeMatriceActive = AUCUN
          BloqueNbBranches 2 + Index
          If Index = VEHICULE And TraficModifié Then ControleMatriceTrafic
          TraficModifié = False
        End If
      Else
        'On passe à la cellule demandée
        vgdTraficPrepare NewRow, NewCol, Index
        '.Row = NewRow
        '.Col = NewCol
        '.Action = 0
        PlaceFocus vgdTrafic(Index), NewCol, NewRow
        .SetFocus
      End If
    Else
      'On reste sur la cellule en cours
      'Récupére l'ancienne valeur
      .Value = SauveValeurSpread
      'Réactive la cellule
      Cancel = True
      PlaceFocus vgdTrafic(Index), Col, Row
      '+19/05
      vgdTraficPrepare NewRow, NewCol, Index
      .SetFocus
      DoEvents
    End If
  End With
  
  If Cancel Then Journal "Cancel"
  
  DoEvents
End Sub

'******************************************************************************
' Détection de la frappe de la touche TAB ou SHIFT TAB
'******************************************************************************
Private Sub vgdTrafic_QueryAdvance(Index As Integer, ByVal AdvanceNext As Boolean, Cancel As Boolean)
  Cancel = False
  Débordement = True
  DrapeauSuivant = AdvanceNext
End Sub

'******************************************************************************
' AfficheRecommandations
' Prépare les messages d'invite ou la boite de message des anomalies
' Paramètres
'   TypeControleActif : Type du controle en cours de vérification
'   Message           : Message actuel
'   chaineMessage     : message résultant à écrire dans l'invite ou la boite de message
'    CP               : Contrôle principal concerné
'    C2 , C3, C4      : Autres contrôles optionnels concernés
' Appelé par :
'   ControleRecommandations
'******************************************************************************
Public Sub AfficheRecommandations(TypeControleActif As String, message As String, _
   Cp As String, Optional C2 As String = TYPE_AUCUN, _
   Optional C3 As String = TYPE_AUCUN, Optional C4 As String = TYPE_AUCUN)
    
    If TypeControleActif = Cp Then
      'Préparer l'invite
      ChaineInvite = ChaineInvite & message & " " ' & Chr(13)
    End If
    If TypeControleActif = Cp Or TypeControleActif = C2 Or _
    TypeControleActif = C3 Or TypeControleActif = C4 Then
      'Préparer le message (boite de dialogue) pour le controle actif
      ChaineMessage = ChaineMessage & message & " " & Chr(13)
    Else
      If Cp = TYPE_RG Then
      'Afficher un message d'avertissement sur la valeur du rayon extérieur
      'Ce message ne sera affiché que pour les boites de message
      'et invite relatifs aux variables R, Bf et LA
        If TypeControleActif = TYPE_R Or TypeControleActif = TYPE_LA Or TypeControleActif = TYPE_BF Then
          ChaineInvite = ChaineInvite & message & " " '& Chr(13)
          ChaineMessage = ChaineMessage & message & Chr(13)
        End If
      'Si une valeur LA, Bf, ou R est modifiée, elle peut entrainer une modification de RG
      'qui peut entrainer un avertissement sur une autre valeur entrant dans la composition de Rg
      ElseIf (C2 = TYPE_RG Or C3 = TYPE_RG Or C4 = TYPE_RG) And _
        (TypeControleActif = TYPE_R Or TypeControleActif = TYPE_LA Or TypeControleActif = TYPE_BF) Then
        ChaineMessage = ChaineMessage & message & Chr(13)
      End If
    End If
  
  'Colorer les avertissements
  ColoreRecommandations Cp
End Sub

'******************************************************************************
' ColoreRecommandations
' Affiche en rouge les valeurs qui sortent des intervalles recommandés
' Paramètres
'   Controle : controle en cours
' Appelé par : AfficheRecommandations
'******************************************************************************
Public Sub ColoreRecommandations(controle As String)
  Select Case controle
    Case TYPE_MILIEU: optMilieu(GiratoireProjet.Milieu).ForeColor = vbRed
    Case TYPE_R:  txtR.ForeColor = vbRed
    Case TYPE_LA: txtLA.ForeColor = vbRed
    Case TYPE_BF: txtBf.ForeColor = vbRed
    Case TYPE_RG: txtRg.ForeColor = vbRed
    Case TYPE_LE4M, TYPE_LE15M, TYPE_LI, TYPE_LS:
      vgdLargBranche.Row = NuméroLigneActive
      vgdLargBranche.Col = CalculeNoColonneLargeur(controle)
      vgdLargBranche.ForeColor = vbRed
    Case TYPE_ANGLE
      'Colore les colonnes Angle et écarts de la matrice
      'pour la ligne considérée
      vgdCarBranche.Row = NuméroLigneActive
      vgdCarBranche.Col = 2
      vgdCarBranche.ForeColor = vbRed
      vgdCarBranche.Col = 3
      vgdCarBranche.ForeColor = vbRed
    Case TYPE_QP
      With vgdTrafic(PIETON)
        .Col = NuméroColonneActive
        .ForeColor = vbRed
      End With
    Case TYPE_Q
      With vgdTrafic(VEHICULE)
        .Row = NuméroLigneActive
        .Col = NuméroColonneActive
        .ForeColor = vbRed
      End With
    End Select
    'Colore les contrôles erronées appartenant à la feuille des caractéristiques de branche
    If Not FeuilleBranche Is Nothing Then
      ColoreRecommandationsCarBranche (controle)
    End If
End Sub

Private Function ColoreRecommandationsCarBranche(controle As String)
  With FeuilleBranche
    Select Case controle
      Case TYPE_LE4M: .txtLE4m.ForeColor = vbRed
      Case TYPE_LE15M: .txtLE15m.ForeColor = vbRed
      Case TYPE_LI: .txtLE4m.ForeColor = vbRed
      Case TYPE_LS: .txtLS.ForeColor = vbRed
      Case TYPE_ANGLE:
        .txtAngleBranche.ForeColor = vbRed
        .txtEcart.ForeColor = vbRed
    End Select
  End With
End Function
Private Function CalculeNoColonneLargeur(NomVariable As String) As Integer
  Select Case NomVariable
    Case TYPE_LE4M: CalculeNoColonneLargeur = 1
    Case TYPE_LE15M: CalculeNoColonneLargeur = 2
    Case TYPE_LI: CalculeNoColonneLargeur = 3
    Case TYPE_LS: CalculeNoColonneLargeur = 4
    Case Else: CalculeNoColonneLargeur = 1
  End Select
End Function

'******************************************************************************
' ControleRecommandations
' Affiche en rouge les valeurs qui sortent des intervalles recommandés
' Cette procéfure est appelé si la valeur est valide
' Paramètres
'
' Appelé par :
'      GotFocus - Chaque fois que le focus est déplacé sur un champ,
'       cette procédure affiche l'invite relative au champ
'      Validate - Chaque fois qu'un champ est validé
'       par Tab ou par déplacement sur un autre champ,
'       la procédure affiche un message de recommandation si la valeur est en dehors
'       des valeurs recommandées, puis colore en rouge les valeurs non recommandées.
'      tabDonnées -
'        Lors de la validation d'un champ par clic sur un autre onglet,
'        l'événement Validate n'est pas appelé
'        L 'appel de cette procédure est fait à ce niveau
'******************************************************************************
Public Function ControleRecommandations(ByVal GotFocus As Boolean, Optional ByVal TypeControle As String = TYPE_COURANT)
 Dim TraficActif As TRAFIC
 Dim Control As Control
 Dim LE4max As Single
 Dim iMax As Integer
 Dim message As String
 Dim Erreur As Boolean
 
  ControleRecommandations = False
  If ChargementEnCours Then Exit Function
  Dim i As Integer
  Select Case TypeControle
    Case TYPE_COURANT
      Set ControleActif = ActiveControl
      TypeControleActif = Mid(ControleActif.Name, 4)
    Case TYPE_AVANT
      'On a cliqué sur un  contrôle qui ne gère pas l'événement Validate
      'On récupère s'il existe le contrôle précédant non vérifié
      TypeControleActif = Mid(ControleActif.Name, 4)
    Case Else
      'Le type de controle actif a déjà été affecté
  End Select

  ChaineInvite = ""
  ChaineMessage = ""
  If Not GotFocus Then
    'Remet les contrôles TextBox dans leur couleur initiale
    'lorsque l'ulilisateur a validé une donnée...
    For Each Control In Controls
      If TypeOf Control Is TextBox Then
        'Les valeurs cumulées des trafics ne sont pas modifiées
        If Left(Control, 5) <> Left(txtQE(1), 5) _
        And Left(Control, 5) <> Left(txtQS(1), 5) Then
          Control.ForeColor = vbWindowText
        End If
      End If
    Next
    'Remet les boutons radio de l'environnement dans sa couleur normale
    optMilieu(0).ForeColor = vbWindowText
  End If
  
  With GiratoireProjet
  'Site et Dimensionnement
  If .Milieu = rc And NbBranches > 6 Then
    'Changement du milieu
    AfficheRecommandations TypeControleActif, IDm_TropDeBranchesEnRC, TYPE_MILIEU, TYPE_NBBRANCHES
  End If
  If .R > 0 And .R < 3.5 Then
    AfficheRecommandations TypeControleActif, IDm_RTropGrandPourMiniG, TYPE_R
  End If
  If .R > 25 Then
    AfficheRecommandations TypeControleActif, IDm_RTropGrand, TYPE_R
  End If
   'Passage en validation Rq0599
   'If .R = 0 And .Milieu = RC Then
   ' AfficheRecommandations TypeControleActif, IDm_RNulEnRC, TYPE_R
  'End If
  'If .R = 0 And .Milieu = PU Then 'Rq0499
  If .R = 0 And .Milieu = PU And txtRg < 12 Then
    AfficheRecommandations TypeControleActif, IDm_RNulEnPU, TYPE_R, TYPE_MILIEU
  End If
  'Test de la largeur de l'anneau LA
  If .LA < 6 And .Milieu = rc Then
    AfficheRecommandations TypeControleActif, IDm_LATropEtroit, TYPE_LA
  ElseIf (.LA > 9 And .Milieu = rc) Or (.LA > 12) Then
    AfficheRecommandations TypeControleActif, IDm_LATropGrand, TYPE_LA
  End If
  'Test du rayon extérieur Rg
  If txtRg < 7.5 And .R = 0 Then
    AfficheRecommandations TypeControleActif, IDm_RgTropPetitPourMiniG, TYPE_RG
  End If
  If txtRg > 12 And .R = 0 Then
    AfficheRecommandations TypeControleActif, IDm_RgTropGrandPourMiniG, TYPE_RG
  End If
  If (txtRg > 12 And txtRg < 15) And .Milieu = rc Then
    AfficheRecommandations TypeControleActif, IDm_RgVoirGirationEnRC, TYPE_RG
  End If
  If (txtRg > 12 And txtRg < 15) And (.Milieu = PU Or .Milieu = CV) Then
    AfficheRecommandations TypeControleActif, IDm_RgVoirGiration, TYPE_RG
  End If
  message = ValidationRg()
  If message <> "" Then
    AfficheRecommandations TypeControleActif, message, TYPE_RG
  End If
  If .LA + .Bf < 7 And .Milieu <> rc Then
    AfficheRecommandations TypeControleActif, IDm_LATropEtroit, TYPE_LA, TYPE_BF
  End If
  'Rq0499
  message = ValidationBf(Recommandation:=True)
'  If .Bf < 1.5 And .R = 0 Then
'    AfficheRecommandations TypeControleActif, IDm_BfTropPetitPourMiniG, TYPE_BF, TYPE_R
'  End If
'  If (txtRg > 12 And txtRg < 15) And (.Bf < 1.5 Or .Bf > 2.5) Then
'    AfficheRecommandations TypeControleActif, IDm_Bf, TYPE_BF, TYPE_RG
'  End If
  'Controle des données de chaque ligne de branche
  Select Case TypeMatriceActive
    Case 0: 'Aucune matrice en cours de saisie
      'Controle de la matrice Dimensionnement (Pour modif R, LA, Bf, Angle et RC)
      DécoloreMatrice vgdLargBranche, 1, 1, NbBranches, 4
      'Contrôle rapport LE4/LE15 suivant RC, LE4m, LE15m
      'Largeur > dimension giratoire
      For NuméroLigneActive = 1 To NbBranches
        ControleDimensionnement1 TypeControleActif
        ControleDimensionnementN TypeControleActif
      Next
      DécoloreMatrice vgdCarBranche, 1, 2, NbBranches, 3
      For NuméroLigneActive = 1 To NbBranches
        'Pour un mini-giratoire seulement, contrôle des angles entre les branches
        If .R = 0 Then ControleCarBranches NuméroLigneActive
        'Erreur = controleChevauchementBranches(True, NuméroLigneActive) 'Contrôlé parce que repassé dans la couleur normale
      Next
      'Contrôle de la valeur maxi de Rg en fonction LE4Max en Rase campagne
      'Contrôle de la largeur d'anneau LA en fonction de LE4Max
      ControleLE4Max
    Case BRANCHE:
      
      'Pour un mini-giratoire seulement, contrôle des angles entre les branches
      If .R = 0 Then
         Dim j As Integer, Res As Integer
        j = NuméroLigneActive Mod NbBranches + 1
        If NuméroLigneActive < j Then
          DécoloreMatrice vgdCarBranche, NuméroLigneActive, 2, j, 3
        Else
          DécoloreMatrice vgdCarBranche, j, 2, j, 3
          DécoloreMatrice vgdCarBranche, NuméroLigneActive, 2, NuméroLigneActive, 3
        End If
        'Contrôle de l'écart avec la branche suivante
        Res = ControleCarBranches(j, False)
        'Si erreur, colore les cellules erronées et affiche le message
        'relatif à la branche suivante
        If Res <> 0 Then
          ColoreMatriceBranche (j)
          If Res < 0 Then
            ChaineMessage = IDm_AngleTropPetitPourMiniG
          Else
            ChaineMessage = IDm_AnglePourMiniG
          End If
          ChaineMessage = ChaineMessage & ".." & IDl_DE & "" _
          & GiratoireProjet.colBranches.Item(j).nom
        End If
        Res = ControleCarBranches(NuméroLigneActive)
      End If
    Case DIMENSION:
      'La matrice de saisie de caractéristiques ou de dimensionnement est active
      DécoloreMatrice vgdLargBranche, NuméroLigneActive, 1, NuméroLigneActive, 4
      ControleDimensionnement1 TypeControleActif
      ControleDimensionnementN TypeControleActif
      'Colore Angle ; Contrôle à partir de Angle, LE4m, LE15m?, LI et LS
      'Erreur = controleChevauchementBranches(True, NuméroLigneActive)
    Case TRAFIC
      'Tests sur les trafics
      Set TraficActif = GiratoireProjet.colTrafics.Item(cboPériode.ListIndex + 1)
      If TypeControleActif = TYPE_QP Then
        'Trafics piétons
        If TraficActif.getQP(NuméroColonneActive) > 999 Then
          AfficheRecommandations TypeControleActif, IDm_QPTropGrand, TYPE_QP
        End If
      Else
        'Trafics véhicule
        If TraficActif.getQ(NuméroLigneActive, NuméroColonneActive) > 1500 Then
          AfficheRecommandations TypeControleActif, IDm_QTropGrand, TYPE_Q
        End If
        'Tourne-à-droite non justifié
        If NuméroColonneActive = NuméroLigneActive Mod NbBranches + 1 Then
          If GiratoireProjet.colBranches.Item(NuméroLigneActive).TAD Then
            'Présence d'un TAD
            If TraficActif.getQ(NuméroLigneActive, NuméroColonneActive) < 100 And _
              TraficActif.getQ(NuméroLigneActive, NuméroColonneActive) <> DONNEE_INEXISTANTE Then
              AfficheRecommandations TypeControleActif, IDm_QTropPetitPourTAD, TYPE_Q
            Else
              'Mettre le TAD en surbrillance
              With vgdTrafic(VEHICULE)
                .Col = NuméroColonneActive
                .Row = NuméroLigneActive
                .ForeColor = vbGrayText
              End With
            End If
          End If
        End If
      End If
    End Select
  End With
  
  MessageEmis = False
  'Le controle en cours vient de perdre le focus
  If Not GotFocus Then
    lblInvite = ""
    If ChaineMessage <> "" Then
      controleEnCours = True
      MsgBox ChaineMessage, vbInformation
      controleEnCours = False
      ControleRecommandations = True
      MessageEmis = True
    End If
    'La valeur a été vérifiée et le focus est positionné sur une autre variable
    If TypeMatriceActive = 0 Then
      Set ControleActif = Nothing
    End If
    DetectModif
  End If
  'On vient d'avoir le focus ou le focus est resté sur le même objet
  If GotFocus Or ValidateObjet Then
    PrepareInvite ChaineInvite
    InviteGotFocus = ChaineInvite
    'Lignes ci-dessous à mettre si on veut mettre le focus en surbrillance
    If ValidateObjet And Not ActiveControl Is Nothing Then
      With ActiveControl
        .SelStart = 0
        .SelLength = Len(.Text)
      End With
    End If
  End If
  DonnéeModifiée = False ' la donnée n'est pas modifiée
End Function
Sub ColoreMatriceBranche(ByVal j As Integer)
 Dim Row, Col As Integer
  With vgdCarBranche
    Col = .ActiveCol
    Row = .ActiveRow
    .Col = 2
    .Row = j
    .ForeColor = vbRed
    .Col = 3
    .ForeColor = vbRed
    .Row = Row
    .Col = Col
    .Action = 0
  End With
End Sub

'******************************************************************************
' Procédure appelée si une modification a été faite et validée
' Passe FichierModifié à vrai
' Lorque l'indicateur CalculAFaire a la valeur VRAI, l'indicateur calculFait
' est mis à FAUX pour signaler que le calcul devra être refait
'******************************************************************************
Private Sub DetectModif(Optional ByVal CalculAFaire As Boolean = True)
  If Not FichierModifié Then
    FichierModifié = True
    GriserMenus True
  End If
  If CalculAFaire And tabDonnées.Tab < TRAFIC - 1 Then
    GiratoireProjet.CalculFait = False
  End If
End Sub

Private Sub ControleLE4Max()
  Dim i As Integer
  Dim LE4max As Single
  LE4max = 0
  NuméroLigneActive = 0
  For i = 1 To NbBranches
    If LE4max < GiratoireProjet.colBranches.Item(i).LE4m Then
        LE4max = GiratoireProjet.colBranches.Item(i).LE4m
        NuméroLigneActive = i
    End If
  Next i
  'If i <= NbBranches Then
    'Rq0499
    'If GiratoireProjet.colBranches.Item(i).LE4m >= 6 And txtRg < 20 And GiratoireProjet.Milieu = RC Then
    If LE4max < 8 And LE4max >= 6 And txtRg < 20 And GiratoireProjet.Milieu = rc Then
      AfficheRecommandations TypeControleActif, IDm_RgTropPetit, TYPE_RG, TYPE_LE4M
    End If
    'If GiratoireProjet.LA < LE4max * 1.2 Then 'Rq0499
    Dim LAU As Single
     LAU = GiratoireProjet.LA + 0.5 * GiratoireProjet.Bf
    If LAU < LE4max * 1.2 Then
      AfficheRecommandations TypeControleActif, _
      IDm_LATropEtroitPourEntrer & GiratoireProjet.colBranches.Item(NuméroLigneActive).nom & ".", _
      TYPE_LA, TYPE_LE4M
    End If
End Sub

Private Sub DécoloreMatrice(vgd As vaSpread, _
  ByVal i1 As Integer, ByVal j1 As Integer, ByVal i2 As Integer, ByVal j2 As Integer)
  With vgd
      .Row = i1
      .Col = j1
      .Row2 = i2
      .Col2 = j2
      .BlockMode = True
      .ForeColor = vbWindowText
      .BlockMode = False
    End With
End Sub

Private Sub ControleDimensionnement1(TypeControleActif As String)
  Dim BrancheActive As BRANCHE
  Set BrancheActive = GiratoireProjet.colBranches.Item(NuméroLigneActive)
  With BrancheActive
    'Contrôle la valeur LE4m de la matrice Dimensionnement
    If .LE4m = 0 Then
      AfficheRecommandations TypeControleActif, IDm_LENul, TYPE_LE4M
    End If
    If .LE4m > 0 And .LE4m < 1.5 Then
      AfficheRecommandations TypeControleActif, IDm_LETropPetit, TYPE_LE4M
    End If
    If .LE4m >= 1.5 And .LE4m < 2.5 Then
      AfficheRecommandations TypeControleActif, IDm_LE2Roues, TYPE_LE4M
    ElseIf .LE4m >= 2.5 And .LE4m < 3 Then
      AfficheRecommandations TypeControleActif, IDm_LEPetit, TYPE_LE4M
    ElseIf .LE4m > 8 And GiratoireProjet.Milieu = rc Then
      AfficheRecommandations TypeControleActif, IDm_LETropLargeEnRC, TYPE_LE4M
    End If
    'GS09 ne pas prendre en compte en Rase campagne
    If .LE4m >= 9 And GiratoireProjet.Milieu <> rc Then
      AfficheRecommandations TypeControleActif, IDm_LETropLargePourPiétons, TYPE_LE4M
    End If
    'Contrôle la valeur LS de la matrice Dimensionnement
    If .LS = 0 Then
      AfficheRecommandations TypeControleActif, IDm_LSNul, TYPE_LS
    End If
    If .LS > 0 And .LS < 1.5 Then
      AfficheRecommandations TypeControleActif, IDm_LSTropPetit, TYPE_LS
    ElseIf .LS >= 1.5 And .LS < 2.75 Then
      AfficheRecommandations TypeControleActif, IDm_LS2Roues, TYPE_LS
    ElseIf .LS >= 2.75 And .LS < 3.5 Then
      AfficheRecommandations TypeControleActif, IDm_LSPetit, TYPE_LS
    ElseIf .LS > 7 Then
      AfficheRecommandations TypeControleActif, IDm_LSTropLarge, TYPE_LS
    End If
    If GiratoireProjet.Milieu >= 0 Then
      'Pour réaliser ce contrôle le milieu doit être défini.
      If BrancheActive.LI > GiratoireProjet.LImax Then
        AfficheRecommandations TypeControleActif, IDm_LITropGrand, TYPE_LI, TYPE_LI
      End If
    End If
  End With
  Set BrancheActive = Nothing
End Sub

Private Sub ControleDimensionnementN(TypeControleActif As String)
  Dim rapport As Single
  Dim BrancheActive As BRANCHE
  Set BrancheActive = GiratoireProjet.colBranches.Item(NuméroLigneActive)
  With BrancheActive
    'Tests impliquant plusieurs valeurs
    If .EntréeEvasée Then
      If .LE15m = 0 Then
        rapport = 10 'Pour afficher un message
      Else
        rapport = .LE4m / .LE15m
      End If
      'condition de validité
      If rapport < 1 Or rapport > 2.5 Then
        AfficheRecommandations TypeControleActif, IDv_RapportLE, TYPE_LE4M
        AfficheRecommandations TypeControleActif, IDv_RapportLE, TYPE_LE15M
      End If
      If GiratoireProjet.Milieu = rc And rapport > 1 Then
        AfficheRecommandations TypeControleActif, IDm_EvasementEnRC, TYPE_LE4M
        AfficheRecommandations TypeControleActif, IDm_EvasementEnRC, TYPE_LE15M
      End If
      If GiratoireProjet.Milieu <> rc And rapport > 1.5 Then
        AfficheRecommandations TypeControleActif, IDm_EvasementTropPetit, TYPE_LE4M
        AfficheRecommandations TypeControleActif, IDm_EvasementTropPetit, TYPE_LE15M
      End If
    End If
    If .LI < 2 And Not .EntréeNulle And Not .SortieNulle And _
      GiratoireProjet.Milieu = CV And GiratoireProjet.R > 0 Then
      AfficheRecommandations TypeControleActif, IDm_LITropPetit, TYPE_LI, TYPE_LE4M, TYPE_LS, TYPE_BF
    End If
    'Condition de validité sur la largeur
    If .LE4m + .LI + .LS >= 2 * txtRg Then
      AfficheRecommandations TypeControleActif, IDv_LTropGrand, TYPE_LE4M, TYPE_LI, TYPE_LS, TYPE_RG
    End If
 
    If TypeMatriceActive = DIMENSION Then
      'Contrôle sur en saisie de la matrice, sinon contrôle global au niveau de Max(LE4m)
      If .LE4m < 8 And .LE4m >= 6 And txtRg < 20 And GiratoireProjet.Milieu = rc Then
        AfficheRecommandations TypeControleActif, IDm_RgTropPetit, TYPE_RG, TYPE_LE4M
      End If
      If GiratoireProjet.LA < .LE4m * 1.2 Then
        AfficheRecommandations TypeControleActif, _
        IDm_LATropEtroitPourEntrer & GiratoireProjet.colBranches.Item(NuméroLigneActive).nom & ".", _
        TYPE_LA, TYPE_LE4M
      End If
    End If
  End With
  Set BrancheActive = Nothing
End Sub

'******************************************************************************
' Controle de validité des branches
' Paramètres
'   Numbranche  : Numéro de branche à contrôler
' Retour
'    n = 0 : pas d'erreur
'    n > 0 : erreur entre la branche n et n-1
'******************************************************************************
Public Function ControleCarBranches(ByVal NumBranche As Integer, Optional EcrireMessage As Boolean = True) As Integer
  Dim EcartAngle, Valeur As Single
  ControleCarBranches = 0
  With GiratoireProjet.colBranches.Item(NumBranche)
    'On fait  les tests si la branche a une entrée
    If Not .EntréeNulle Then
        If NumBranche = 1 Then
          Valeur = GiratoireProjet.colBranches.Item(CInt(NbBranches)).Angle
          If gbProjetActif.modeangle = GRADE Then
            EcartAngle = (400 - Valeur) * 0.9
          Else
            EcartAngle = 360 - Valeur
          End If
          '1606
        Else
          EcartAngle = .Angle - GiratoireProjet.colBranches.Item(NumBranche - 1).Angle
          If gbProjetActif.modeangle = GRADE Then EcartAngle = EcartAngle * 0.9
          If EcartAngle < 0 Then EcartAngle = EcartAngle + 360
        End If
        If EcartAngle < 70 Then
          ControleCarBranches = -NumBranche
          If EcrireMessage Then
            AfficheRecommandations TypeControleActif, IDm_AngleTropPetitPourMiniG, TYPE_ANGLE
          End If
        ElseIf EcartAngle < 80 Then
          ControleCarBranches = NumBranche
          If EcrireMessage Then
            AfficheRecommandations TypeControleActif, IDm_AnglePourMiniG, TYPE_ANGLE
          End If
      End If
    End If
  End With
End Function

'******************************************************************************
' PrepareInvite
' Prepare l'invite pour une variable à saisir
'  en rajoutant au paramètre Message l'invite approprié
' Paramètres
'   Message : Message initial de l'invite
' Appelé par
'
'******************************************************************************
Private Sub PrepareInvite(ByVal message As String)
  lblInvite = message
  'L'invite est complété par l'invite par défaut
  If lblInvite <> "" Then lblInvite = lblInvite + vbCrLf
  Select Case TypeControleActif
    Case TYPE_BF
      lblInvite = lblInvite + IDi_BF
    Case TYPE_LE4M
      lblInvite = lblInvite + IDi_LE4M
    Case TYPE_LS
      lblInvite = lblInvite & IDi_LS
    Case TYPE_QP
      lblInvite = lblInvite & IDi_QP
    Case TYPE_Q
      lblInvite = lblInvite & IDl_DE & _
      GiratoireProjet.colBranches.Item(NuméroLigneActive).nom & IDl_VERS & _
      GiratoireProjet.colBranches.Item(NuméroColonneActive).nom & "..."
    Case Else
      'Invite général par défaut pour les contrôles qui n'ont pas leur propre invite
      lblInvite = lblInvite + Idi_Défaut
  End Select
End Sub

'******************************************************************************
'  ValideRayonEtBranches
'  Vérifie que la modification du rayon extérieur peut être validée
' Retourne le booléen TRUE si valide
'******************************************************************************
Private Function ValideRayonEtBranches(ByRef message As String) As Boolean
  Dim i As Integer
  Dim Continue As Boolean
  Dim uneBranche As BRANCHE
  i = 1
  Continue = True
  calculRg False
  Do While i < NbBranches And Continue
    Set uneBranche = gbProjetActif.colBranches.Item(i)
    Continue = VerifierAngleBranche(i, CSng(txtRg), uneBranche, message)
    i = i + 1
  Loop
  ValideRayonEtBranches = Continue
End Function

'******************************************************************************
' ValidationDonnées
' Vérifie si la donnée saisie peut être validée
' Retour
'   VRAI si oui
'   FAUX sinon. Un message signalant l'erreur est affichée
' Paramètres
'   Valeur = valeur de la variable à valider
' La variable TypeControleActif précise le type de donnée à vérifier
' Appel par
'   Validate des variables
'   tabDonnées_click lors d'un changement d'onglet
'   X_LeaveCell pour les matrices
'******************************************************************************
Public Function ValidationDonnées(ByVal Valeur As Variant, Optional wBranche As BRANCHE) As Boolean
  Dim message As String
  Dim x As Single
  
  message = ""
  ValidationDonnées = True
  If Valeur = "" Then Valeur = 0
  
  If IsNumeric(Valeur) Then
    x = CSng(Valeur)
    Select Case TypeControleActif
      
      'Contrôle de l'angle entre les branches
      Case TYPE_ANGLE
        If Valeur < 0 Then
          message = IDv_ValeurPositive
          ValidationDonnées = False
        Else
          'Contrôle de la compatibilité de l'angle avec les autres données
          ValidationDonnées = VerifierAngleBranche(NuméroLigneActive, x, wBranche, message)
        End If
      
      'Contrôle du rayon du giratoire
      Case TYPE_R
        If Valeur > 100# Then
          message = IDv_RayonInferieur100m
          ValidationDonnées = False
        ElseIf Valeur < 0 Then
          message = IDv_ValeurPositive
          ValidationDonnées = False
        'Rajout en validité Rq0599
        'Un mini-giratoire est interdit en rase campagne
        Else
          'Calcule le rayon extérieur Rg
          calculRg False
          If Valeur = 0 And GiratoireProjet.Milieu = rc Then
            message = IDm_RNulEnRC
            ValidationDonnées = False
            'Rajout en validité Rq0599
            'Un mini-giratoire est interdit en péri-urbain
            'Suppression car ce n'est pas une cause d'invalidité
'          ElseIf valeur = 0 And GiratoireProjet.Milieu = PU And txtRg <= 12 Then
'            message = IDm_RNulEnPU
'            ValidationDonnées = False
          Else
            'Compatibilité du rayon avec les branches
            ValidationDonnées = ValideRayonEtBranches(message)
            If ValidationDonnées Then
              'Si des TAD sont définis on ne peut passer en mini-giratoire sans effacer
              'ceux-ci
              ValidationDonnées = ValidationRetTAD(Valeur)
            End If
          End If
        End If
      
      'Contrôle de la bande franchissable Bf
      Case TYPE_BF
        If Valeur < 0 Or Valeur > 3 Then
          message = IDv_ControleBornes & str(0) & IDl_ET & str(3) & IDl_METRE & "."
          ValidationDonnées = False
        Else
          'Calcule le rayon extérieur Rg
          calculRg False
          'Compatibilité du rayon avec les branches
          ValidationDonnées = ValideRayonEtBranches(message)
        End If
      
      'Contrôle de la largeur d'anneau LA
      Case TYPE_LA
        ValidationDonnées = ValidationLA(GiratoireProjet.Milieu, Valeur, message)
        If ValidationDonnées Then
          calculRg False
          'Compatibilité de la largeur d'anneau avec les branches
          ValidationDonnées = ValideRayonEtBranches(message)
        End If
      
      'Contrôle de la largeur de l'îlot
      Case TYPE_LI
        If Valeur < 0 Then
          message = IDv_ValeurPositive
          ValidationDonnées = False
        Else
          ValidationDonnées = VerifierAngleBranche(NuméroLigneActive, x, wBranche, message)
        End If
      
      'Contrôle de la largeur d'entrée
      Case TYPE_LE4M
        If Valeur < 0 Or Valeur > 12 Then
          message = IDv_ControleBornes & str(0) & IDl_ET & str(12) & IDl_METRE & "."
          ValidationDonnées = False
        ElseIf Valeur = 0 And GiratoireProjet.colBranches.Item(NuméroLigneActive).TAD Then
          'Pas d'entrée nulle si TAD
          message = IDv_LE0etTAD
          ValidationDonnées = False
        ElseIf Valeur = 0 And GiratoireProjet.colBranches.Item(NuméroLigneActive).SortieNulle Then
          'Entrée et Sortie ne peuvent être toutes deux nulles
          message = IDv_LE0etLS0
          ValidationDonnées = False
        Else
          'Compatibilité evec les autres données
          ValidationDonnées = VerifierAngleBranche(NuméroLigneActive, x, wBranche, message)
        End If
      
      'Contrôle de la largeur de sortie LS
      Case TYPE_LS
        If Valeur < 0 Or Valeur > 10 Then
          message = IDv_ControleBornes & str(0) & IDl_ET & str(10) & IDl_METRE & "."
          ValidationDonnées = False
        ElseIf Valeur = 0 And GiratoireProjet.colBranches.Item(NuméroLigneActive).EntréeNulle Then
          'Entrée et Sortie ne peuvent être toutes deux nulles
          message = IDv_LE0etLS0
          ValidationDonnées = False
        '0699 a implémenter complètement dans prochaine version
        'en retirant inhibant le TAD de la branche si la sortie de i+1 est nulle
'        ElseIf valeur = 0 And _
'          GiratoireProjet.colBranches.Item(BranchePrécédent(NuméroLigneActive)).TAD Then
'          'sortie nulle et TAD sur la branche précédente
'          message = IDv_LS0etTAD
'          ValidationDonnées = False
        Else
          'Cohérence de l'ensemble
          ValidationDonnées = VerifierAngleBranche(NuméroLigneActive, x, wBranche, message)
        End If
      
      'Contrôle du trafic (en fait réalisé par le spread de trafic)
      Case TYPE_Q
        If Valeur < 0 Or Valeur > 2500 Then
          Const IDl_UVP = " uvp"
          message = IDv_ControleBornes & str(0) & IDl_ET & str(2500) & IDl_UVP & "."
          ValidationDonnées = False
        End If
      'Contrôle du trafic piéton (également réalisé par le spread de trafic)
      Case TYPE_QP
        'GS09 Correction test
        'If valeur < 0 Or valeur > 250 Then
        If Valeur < 0 Or Valeur > 2500 Then
          Const IDl_PIETON = " p"
          message = IDv_ControleBornes & str(0) & IDl_ET & str(2500) & IDl_PIETON & "."
          ValidationDonnées = False
        End If
        
      'Contrôle du type de site
      Case TYPE_MILIEU
        'Rajout en validité Rq0599
        If Valeur = rc And GiratoireProjet.R = 0 Then
          message = IDm_RNulEnRC
          ValidationDonnées = False
          '1606 Ce n'est pas une cause d'invalidité
'        ElseIf valeur = PU And GiratoireProjet.R = 0 And txtRg <= 12 Then
'          message = IDm_RNulEnPU
'          ValidationDonnées = False
        ElseIf Not ValidationLA(Valeur, GiratoireProjet.LA, message) Then
          message = IDv_ModifMilieu
          ValidationDonnées = False
        End If
        If Not ValidationDonnées And Not ChargementEnCours Then
          'Ne valide pas la modification et
          'Replace le focus à sa position antérieure
          optMilieu(GiratoireProjet.Milieu).SetFocus
        End If
    End Select
    
    If ValidationDonnées Then
      Select Case TypeControleActif
        'Contrôle de la valeur du rayon extérieur
        Case TYPE_R, TYPE_BF, TYPE_LA
          'essai 0599
           'suppression du test de validité sur Rg en saisie 1606
          'message = ValidationRg
          message = ""
          If message <> "" Then
            ValidationDonnées = False
          Else
            'Affectation des valeurs
            Select Case TypeControleActif
              Case TYPE_R: GiratoireProjet.R = Valeur
              Case TYPE_BF: GiratoireProjet.Bf = Valeur
              Case TYPE_LA: GiratoireProjet.LA = Valeur
            End Select
            calculRg True
          End If
      End Select
    End If
    
  Else
    'Valeur non numérique
    message = IDv_ValeurNumérique
    ValidationDonnées = False
  End If
  
  If message <> "" Then
    'Une erreur a été détectée
    'Impression du message
    controleEnCours = True
    MsgBox message, vbExclamation
    controleEnCours = False
    lblInvite = InviteGotFocus
  End If
End Function
'******************************************************************************
' ValidationRetTAD
' Vérifie si la donnée saisie R peut être validée
' Si un ou plusieurs tourne-à-droite existent,
' le rayon R ne peut être validé qu'après suppression de tous les TAD
' sur confirmation de l'utilisateur
' La fonction renvoie  VRAI si R peut être validé, FAUX dans le cas contraire
'******************************************************************************
Private Function ValidationRetTAD(ByVal Valeur As Variant) As Boolean
  Dim i As Integer
  Dim TAD As Boolean
  ValidationRetTAD = True
  TAD = False
  If Valeur = 0 Then
    'On transforme le giratoire en mini-giratoire
    'Des tournes à droite ont-ils été définis?
    For i = 1 To NbBranches
      If GiratoireProjet.colBranches.Item(i).TAD Then
        TAD = True
      End If
    Next i
    If TAD Then
      If MsgBox(IDv_RetTAD, vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
        ValidationRetTAD = False
      Else
        'Les tourne-à-droite sont supprimés
        For i = 1 To NbBranches
          GiratoireProjet.colBranches.Item(i).TAD = False
        Next i
      End If
    End If
    If ValidationRetTAD Then
      'Il faut interdire la possibilité de saisir les tournes à droite
      With vgdCarBranche
        .Col = 5
        For i = 1 To NbBranches
          .Row = i
          .Value = False
          .Lock = True
        Next i
      End With
    End If
  ElseIf Valeur <> 0 And GiratoireProjet.R = 0 Then
    'On transforme un mini-giratoire en giratoire normal
    'On autorise à nouveau l'accès aux tournes à droite
    'pour chaque branche dont l'entrée n'est pas nulle
    With vgdCarBranche
      .Col = 5
      'Parcours des branches
      For i = 1 To NbBranches
        .Row = i
        .Value = False
        'Autorisation de TAD
        If Not GiratoireProjet.colBranches.Item(i).EntréeNulle Then .Lock = False
      Next i
    End With
  End If
End Function


'******************************************************************************
' ValidationRg
' Vérifie si la donnée saisie Rg peut être validée
' Retour
'   Si pas d'erreur : la chaine message ne contient rien (chaine vide)
'   Si erreur : la chaine message signalant l'erreur
'
'******************************************************************************
Private Function ValidationRg() As String
  Dim message As String
  message = ""
  '??0699
  If txtR = 0 Then
    If txtRg < 7.5 Or txtRg > 12 Then
      message = IDv_ControleBornesRg & str(7.5) & IDl_ET & str(12) & IDl_METRE & "."
    End If
  Else
    If txtRg < 12 Then
      message = IDv_ValidationRgMinimal & str(12) & IDl_METRE & "."
    End If
  End If
  ValidationRg = message
End Function

Private Function ValidationBf(ByVal Recommandation As Boolean) As String
  Dim message As String
  message = ""
  With GiratoireProjet
    'If .Bf < 1.5 And .R = 0 Then
    If (.Bf < 1.5 Or .Bf > 2.5) And .R = 0 Then
      message = IDm_BfTropPetitPourMiniG
      If Recommandation Then AfficheRecommandations TypeControleActif, message, TYPE_BF, TYPE_R
'    ElseIf (txtRg > 12 And txtRg < 15) And (.Bf < 1.5 Or .Bf > 2.5) Then
'      If Recommandation Then AfficheRecommandations TypeControleActif, IDm_Bf, TYPE_BF, TYPE_RG
'    End If
  '1606 Le test Bf non compris entre 1.5 et 2.5 est seulement en recommandations
     ElseIf (txtRg > 12 And txtRg < 15) And (.Bf < 1.5 Or .Bf > 2#) And Recommandation Then
      AfficheRecommandations TypeControleActif, IDm_Bf, TYPE_BF, TYPE_RG
    End If
  End With
  ValidationBf = message
End Function

Private Function ValidationRgEtBranches(message As String) As Boolean
  Dim i As Integer
  ValidationRgEtBranches = True
  For i = 1 To NbBranches
    With GiratoireProjet.colBranches.Item(i)
      If txtRg < .LE4m + .LI Then
        ValidationRgEtBranches = False
      End If
    End With
  Next i
  If Not ValidationRgEtBranches Then message = IDv_RgOuBranchesIncorrect
End Function

Private Function ValidationRgEtUneBranche(ByVal TypeDonnée As String, ByVal Valeur As Variant, _
  message As String) As Boolean
  ValidationRgEtUneBranche = True
  If TypeDonnée = TYPE_LE4M Then
    If CSng(txtRg) < CSng(Valeur) + GiratoireProjet.colBranches.Item(NuméroLigneActive).LI Then
      ValidationRgEtUneBranche = False
    End If
  Else
    If CSng(txtRg) < GiratoireProjet.colBranches.Item(NuméroLigneActive).LE4m + CSng(Valeur) Then
      ValidationRgEtUneBranche = False
    End If
  End If
  If Not ValidationRgEtUneBranche Then message = IDv_RgOuUneBrancheIncorrect
End Function

'******************************************************************************
' ValidationLA
' Vérifie si la variable LA est validé
' Retour
'   VRAI si oui et on pourra déplacer le focus
'   FAUX sinon ; un message signalant l'erreur est affichée et,
'                on refusera le déplacement du focus
'
'******************************************************************************
Private Function ValidationLA(ByVal Milieu, ByVal LA, ByRef message As String)
  ValidationLA = True
  If Milieu = rc Then
    If LA < 4.5 Or LA > 12 Then
      message = IDv_ControleBornesLA & str(4.5) & IDl_ET & str(12) & IDl_METRE & "."
       ValidationLA = False
    End If
  Else
    If LA < 4.5 Or LA > 18 Then
      message = IDv_ControleBornesLA & str(4.5) & IDl_ET & str(18) & IDl_METRE & "."
      ValidationLA = False
    End If
  End If
End Function

'******************************************************************************
' InitControle
' Met le contrôle dans la couleur normale, (le passe en surbrillance)
' et sauvegarde sa valeur
'******************************************************************************
Public Sub InitControle(GotFocus As Boolean)
  If ActiveControl Is Nothing Then Exit Sub
  'Sortie par sécurité si ActiveControl est nothing ; ce qui ne devrait pas arrivé
  'puisque cette procédure est activée seulement lorsque le contrôle a le focus
  With ActiveControl
    If GotFocus Then
      'Désactive la poignée de sélection sur le dessin du giratoire
      shpPoignée.Visible = False
      lblInvite = ""
      .ForeColor = vbWindowText
      If TypeOf ActiveControl Is TextBox Then
        Journal "***InitCTRLSauveValeur=TEXT"
        SauveValeur = .Text
        'Lignes ci-dessous à mettre si on veut mettre le focus en surbrillance
        .SelStart = 0
        .SelLength = Len(.Text)
      End If
    Else
      Journal "***InitCtrlTEXT=SauveValeur"
      .Text = SauveValeur
      '0699
      DonnéeModifiée = False ' la donnée n'est pas modifiée
    End If
   End With
End Sub

'******************************************************************************
' Dessin du Giratoire
' IsPremierDessin : indique que  c'est la première fois qu'on dessine
' FacteurZoom=0 : indique qu'il faut recalculer l'échelle
'******************************************************************************
Private Sub DessinerGiratoire(IsPremierDessin As Boolean)
If FacteurZoom = 0 Then
  DemiHauteur = (ScaleHeight - fraInvite.Height) / 2
  DemiLargeur = (ScaleWidth - tabDonnées.Width) / 2
  DessinGiratoire.gbDemiHauteur = DemiHauteur
  DessinGiratoire.gbDemiLargeur = DemiLargeur
  DessinGiratoire.gbFacteurZoom = 0
End If

With GiratoireProjet
  DessinGiratoire.gbRayonInt = .R
  DessinGiratoire.gbRayonExt = .R + .LA + .Bf
  DessinGiratoire.gbBandeFranchissable = .Bf
End With

DessinGiratoire.DessinerTout IsPremierDessin
If DiagramFlux And Not TraficActif Is Nothing Then cLS: TraficActif.CalculDiagramFlux

End Sub

'******************************************************************************
' Dessin du Giratoire
' FacteurZoom=0 : indique qu'il faut recalculer l'échelle
'******************************************************************************
Public Sub Redess()
  FacteurZoom = 0
  DessinerGiratoire IsPremierDessin:=False
End Sub

Public Sub AfficheDiagramflux(ByVal Etat As Boolean)
Dim i As Integer
  cLS
  If Not Etat Then
    With GiratoireProjet.colBranches
      For i = 1 To NbBranches
        linBordIlotEntrée(i).Visible = (linBordIlotEntrée(i).Tag = "V")
        linBordIlotSortie(i).Visible = (linBordIlotSortie(i).Tag = "V")
        linBordIlotGir(i).Visible = (linBordIlotGir(i).Tag = "V")
        linBordIlotEntrée(i).Tag = ""
        linBordIlotSortie(i).Tag = ""
        linBordIlotGir(i).Tag = ""
      Next
    End With
  End If
  
  DiagramFlux = Etat
End Sub

'**********************************************************************************
' Réaffichage du Spread en mode normal suite à la sélection graphique d'une branche
'**********************************************************************************
Private Sub AfficheSpreadNormal()
  If TypeOf ActiveControl Is vaSpread Then
    With ActiveControl
      If .OperationMode = OperationModeRow Then .OperationMode = OperationModeNormal
    End With
    shpPoignée.Visible = False
  End If
End Sub

'**********************************************************************************
' Vérifie les données de la feuille de données avant de lancer le calcul
' Retourne
'   True : aucune erreur n'est apparue et le calcul peut être déclenché
'   False : des erreurs de dimensionnement ont été faites ;
'**********************************************************************************
Public Function ValiderFeuilleDonnées(Optional ByRef MessageImprim) As Boolean
  Dim message, Message2 As String
  Dim NoBrancheErronée As Integer
  Dim i, j As Integer
  Dim rapport As Single
  If ControleTrafic Then Exit Function
  
  'Vérification du rayon Rg
  message = ValidationRg
  If message <> "" Then message = message & vbCrLf
  'Vérification de la valeur Bf
  '1606Message2 = ValidationBf(Recommandation:=True)
  Message2 = ValidationBf(Recommandation:=False)
  If Message2 <> "" Then message = message & vbCrLf & Message2 & vbCrLf

  'Vérification de la validité des rapports LE4m/LE15m...
  With GiratoireProjet.colBranches
    For i = 1 To NbBranches
      If .Item(i).EntréeEvasée Then
        '0406 Eviter l'entrée nulle
        If .Item(i).LE15m = 0 Then
          rapport = 10 'Pour afficher un message
        Else
          rapport = .Item(i).LE4m / .Item(i).LE15m
        End If
        'condition de validité
        If rapport < 1 Or rapport > 2.5 Then message = message + IDl_Branche & .Item(i).nom & " : " & IDv_RapportLE & vbCrLf
      End If
    Next i
  End With
  If GiratoireProjet.colTrafics.count = 0 Then
    message = message + IDv_PasTrafic
  'ElseIf txtQT = "" Then
   'Vérication du trafic total
   ' message = message & vbCrLf & IDv_TraficTotalNul & vbCrLf
  End If
  
  If message = "" Then
    ValiderFeuilleDonnées = True
  Else
    If IsMissing(MessageImprim) Then
      message = message & vbCrLf & vbCrLf & IDv_NonValide
      MsgBox message, vbOKOnly + vbExclamation
    Else
      MessageImprim = message
    End If
  End If
End Function

'******************************************************************************
' Contrôle des données de trafic en fonction des largeurs d'entrée et de sortie
' Les lignes ou colonnes de la matrice de trafic seront oblitérées si les largeurs
' d'entrée ou sortie sont nulles
' Paramètre AfficheMessage :
'      Vrai  ->  on affiche les messages d'erreur
'                on ne verrouille pas les cellules dont les largeurs sont nulles
'      Faux  ->  on affiche pas les messages d'erreur
'                on verrouille les cellules dont les largeurs sont nulles
'******************************************************************************
Private Function ControleTrafic(Optional ByVal AfficheMessage As Boolean = True) As Boolean
Dim NumBranche, i, j, ip, ip2 As Integer
Dim message As String
Dim ErreurBranche As Boolean
Dim AfficherTraficCourant As Boolean
Const LIGNE = 0
Const COLONNE = 1

  'Pour l'ensemble des périodes
  'a-t-on des trafics sur des voies de largeur nulle (LE ou LS = 0)
  If cboPériode.ListIndex = -1 Then Exit Function
  ErreurBranche = False
  ip = 1
  Do While ip <= GiratoireProjet.nbPériodes  'Boucle sur les périodes
  With GiratoireProjet.colTrafics.Item(ip)
      For i = 1 To NbBranches
        If GiratoireProjet.colBranches.Item(i).LE4m = 0 And .getQE(i) > 0 Then
          If AfficheMessage Then
            ErreurBranche = True
          Else
            'Remettre à zéro les cellules de la ligne
            For ip2 = 1 To GiratoireProjet.nbPériodes
              AfficherTraficCourant = (ip2 = cboPériode.ListIndex + 1)
              GiratoireProjet.colTrafics.Item(ip2).AnnulerTrafic LIGNE, i, AfficherTraficCourant
            Next ip2
          End If
        End If
      Next i
      For i = 1 To NbBranches
        If GiratoireProjet.colBranches.Item(i).LS = 0 And .getQS(i) > 0 Then
          If AfficheMessage Then
            ErreurBranche = True
          Else
             'Remettre à zéro les cellules de la colonne
            For ip2 = 1 To GiratoireProjet.nbPériodes
              AfficherTraficCourant = (ip2 = cboPériode.ListIndex + 1)
              GiratoireProjet.colTrafics.Item(ip2).AnnulerTrafic COLONNE, i, AfficherTraficCourant
            Next ip2
          End If
        End If
      Next i
    End With
    ip = ip + 1
  Loop
  'Erreur rencontrée
  If ErreurBranche Then
    controleEnCours = True
    If MsgBox(IDv_TraficNonNul, vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
      'Revenir s'il y a lieu dans l'onglet Dimensionnement
      AfficheMessage = False
    Else
      'Passer s'il y a lieu dans l'onglet Trafic
      'Bloquer/débloquer les entrées/sorties du spread VEHICULE
      ControleTrafic False
      ErreurBranche = False
    End If
    controleEnCours = False
    
    ControleTrafic = ErreurBranche
  End If
  
  If AfficheMessage Then
    GiratoireProjet.colBranches.BlocageTrafic Me
    'Réaffecte les trafics affichés dans le spread courant
    'Fait par AnnulerTrafic dans TRAFIC.CLS (AV - 26.04.99)
  End If
End Function

'******************************************************************************
' Controle de la matrice de trafic
' Vérifie s'il y a un trafic entrant ou sortant sur chaque entrée ou sortie
' Ce contrôle ne peut se faire que sur la matrice UVP
' Il est déclenché
'     -en affichage de l'onglet Trafic
'     -en sortie de la matrice trafic
'     -au changement de période
' Remarque : le controle pendant la saisie n'est pas réalisable
'******************************************************************************
Private Sub ControleMatriceTrafic()
  Dim TraficActif As TRAFIC
  If ChargementEnCours Or cboPériode.ListIndex = -1 Then Exit Sub
  Set TraficActif = GiratoireProjet.colTrafics.Item(cboPériode.ListIndex + 1)
  '0599 Faire le contrôle même si la saisie du trafic est incomplète
  'If Not TraficActif.EstComplète Then Exit Sub
  Dim message, Message2 As String
  Dim i, j As Integer
  'Réinitialise les couleurs des cellules de trafic
  If ChargementEnCours Then Exit Sub
  For i = 1 To NbBranches
    txtQE(i).ForeColor = vbWindowText
    txtQS(i).ForeColor = vbWindowText
  Next i
  txtQT.ForeColor = vbWindowText
  'Tests des trafics cumulés d'entrée et de sortie
  For i = 1 To NbBranches
    With GiratoireProjet.colBranches.Item(i)
      If TraficActif.getQE(i) > 2500 Then
        message = message & IDl_Branche & .nom & " : " & IDm_QETropImportant & vbCrLf
        txtQE(i).ForeColor = vbRed
      End If
      'If TraficActif.getQE(i) = 0 And .LE4m <> 0 Then
      If txtQE(i) <> "" Then
        'des trafics ont été saisis sur la ligne i
        'Test sur le trafic UVP
        If TraficActif.getQE(i) = 0 And .LE4m <> 0 Then
          message = message & IDl_Branche & .nom & " : " & IDm_QENul & vbCrLf
          txtQE(i).ForeColor = vbRed
        End If
      End If

      If txtQS(i) <> "" Then
        'des trafics ont été saisis sur la colonne i
        '1606If txtQS(i) = 0 And .LS <> 0 Then
        'Test sur le trafic UVP
        If TraficActif.getQS(i) = 0 And .LS <> 0 Then
          message = message & IDl_Branche & .nom & " : " & IDm_QSNul & vbCrLf
          txtQS(i).ForeColor = vbRed
        End If
      End If
    End With
  Next i
  'Test du trafic total du giratoire
  Message2 = ControleTraficTotal
  If Message2 <> "" Then
    message = message & Message2
    txtQT.ForeColor = vbRed
  End If
  controleEnCours = True
  If message <> "" Then MsgBox (message)
  controleEnCours = False
End Sub


Public Function ControleTraficTotal(Optional ByVal NuméroTraficActif As Integer = 0) As String
  Dim message As String
  Dim QETotal As Integer, numéroBranche As Integer
  message = ""
  'Calcul du trafic total pour la période active
  If NuméroTraficActif = 0 Then
    If txtQT = "" Then
      QETotal = 0
    Else
      QETotal = CInt(txtQT)
    End If
  Else
    'Calcul du total pour la période demandée
    Dim i As Integer
    Dim TraficRésultat As TRAFIC
    QETotal = 0
    Set TraficRésultat = GiratoireProjet.colTrafics.Item(NuméroTraficActif)
    With TraficRésultat
      For i = 1 To NbBranches
        If .getQE(i) <> DONNEE_INEXISTANTE Then QETotal = QETotal + .getQE(i)
      Next i
    End With
  End If
  'Test sur le trafic total du giratoire
  '-------------------------------------
  If GiratoireProjet.R = 0 Then
    If QETotal > 1500 And QETotal <= 1800 Then
      message = IDm_QEGrandPourMiniG
    ElseIf QETotal > 1800 Then
      message = IDm_QETropGrandPourMiniG
    End If
  ElseIf QETotal > 5000 Then
    message = IDm_QETropGrand
  End If
  ControleTraficTotal = message
End Function

'******************************************************************************
' Controle des valeurs de la matrice de trafic courante
' Vérifie les valeurs de trafic et réaffecte les couleurs normales ou rouges
' à chaque cellule de la trafic active
' Ce contrôle est déclenché
'     -en affichage de l'onglet Trafic si la période est complète c-a-d si
'                   toutes les cellules ont été saisies.
'     -au changement de période
' Les mêmes tests sont également déclenchés dans ControleRecommandations
' lors de la saisie individuelle des valeurs de trafic
'******************************************************************************
Private Sub ControleValeursTrafic()
  Dim i, j As Integer
  
  If cboPériode.ListIndex = -1 Then Exit Sub
  
  Set TraficActif = GiratoireProjet.colTrafics.Item(cboPériode.ListIndex + 1)
  'Controle de la matrice des piétons
  With vgdTrafic(PIETON)
    For i = 1 To NbBranches
      .Row = i
      If TraficActif.getQP(i) > 999 Then
        'Trafic trop important
        .ForeColor = vbRed
      Else
        'Trafic normal
        .ForeColor = vbWindowText
      End If
    Next i
  End With
  'Controle de la matrice des véhicules
  With vgdTrafic(VEHICULE)
    For i = 1 To NbBranches
      .Row = i
      For j = 1 To NbBranches
        .Col = j
        If TraficActif.getQ(i, j) > 1500 Then
          'Trafic trop important
          .ForeColor = vbRed
        Else
          'Trafic normal
          If j = i Mod NbBranches + 1 And GiratoireProjet.colBranches.Item(i).TAD Then
            'Présence d'un TAD de i vers i+1
            'On fait ressortir celui-ci dans la cellule de trafic i->i+1
            .ForeColor = vbGrayText
          Else
            .ForeColor = vbWindowText
          End If
        End If
      Next j
    Next i
    'Tourne-à-droite non justifié sur la matrice véhicules
    For i = 1 To NbBranches
      j = i Mod NbBranches + 1
      If TraficActif.getQ(i, j) < 100 And _
        TraficActif.getQ(i, i) <> DONNEE_INEXISTANTE _
        And GiratoireProjet.colBranches.Item(i).TAD Then
          'Trafic ne nécessitant pas la présence d'un tourne à droite
          .Row = i
          .Col = j
          .ForeColor = vbRed
      End If
    Next i
  End With
End Sub

'******************************************************************************
' CalculBord
' Calcul les angles des bords extrèmes de chaussée relativement à l'axe central
' Les angles sont retournés dans les variables linBordVoieEntrée et
'  linBordVoieSortie
' Le calcul est fait pour la branche courante si NumBranche est égal à 0.
' Dans ce cas les valeurs géométriques de la branche sont transmis en paramètres.
' Si Nmbranche est différent de 0, les angles des bords de chaussée de la branche
' numbranche sont calculés. Les paramètres sont recherchés dans l'objet Branche
' La fonction renvoie TRUE si le calcul a pu s'effectuer normalement
' FAUX si la modification du giratoire entraine une impossibilité
'******************************************************************************

Private Function calculBord(ByVal NumBranche As Integer, ByVal LE4m As Single, _
 ByVal LI2 As Single, ByVal LS As Single, ByVal Rg As Single, _
  ByRef linBordVoieEntrée As Single, ByRef linBordVoieSortie As Single) As Boolean
   Dim unXLoc, unYLoc As Single
   If NumBranche > 0 Then
     With GiratoireProjet.colBranches.Item(NumBranche)
      LE4m = .LE4m
      LS = .LS
      LI2 = .LI / 2
     End With
   End If
    'Entrée ou sortie nulle
   If LE4m = 0 Or LS = 0 Then
    If LE4m = 0 Then
      'Entrée nulle
      unXLoc = Carré(Rg) - Carré(LS)
      If unXLoc >= 0 Then
        unXLoc = Sqr(unXLoc)
        unYLoc = LS
        linBordVoieEntrée = 0#
        linBordVoieSortie = Arccos(unXLoc / Rg)
      End If
    Else
      unXLoc = Carré(Rg) - Carré(LE4m)
      If unXLoc >= 0 Then
        unXLoc = Sqr(unXLoc)
        unYLoc = -LE4m
        linBordVoieEntrée = Arccos(unXLoc / Rg)
        linBordVoieSortie = 0#
      End If
    End If
  Else ' Cas général
    unXLoc = Carré(Rg) - Carré(LE4m + LI2)
    If unXLoc >= 0 Then
      unXLoc = Sqr(unXLoc)
      unYLoc = -(LE4m + LI2)
      linBordVoieEntrée = Arccos(unXLoc / Rg)
      unXLoc = Carré(Rg) - Carré(LS + LI2)
      If unXLoc >= 0 Then
        unXLoc = Sqr(unXLoc)
        unYLoc = LS + LI2
        linBordVoieSortie = Arccos(unXLoc / Rg)
      End If
    End If
  End If
  calculBord = (unXLoc >= 0)
 End Function
 
'******************************************************************************
' VerifierAngleBranche
' Cette fonction vérifie si les branches du giratoire ne se chevauchent pas.
' La fonction est appelée dès qu'une modification géométrique est réalisée
' sur les angles, les largeurs d'entrée, d'ilot et de sortie et sur le rayon
' extérieur.
' Les calculs sont analogues à ceux de la fonction VerifierAngleBranche du
' module DessinGiratoire déclenchée en cas de modification graphique de l'angle
' d'une branche. Les modifications graphiques interactives des rayons
' déclenchent l'appel de cette fonction ci-dessous car la validation de ces
' données est du ressort de FrmDonnées.
' La fonction returne Vrai si la modification du giratoire peut être validée.
' Sans le cas contraire, la chaine de caractères Message contenant le message
' d'erreur approprié est retournée en paramètre.
'******************************************************************************
  Private Function VerifierAngleBranche(ByVal NumBrancheSelect As Integer, x As Single, wBranche As BRANCHE, _
  ByRef message As String)
  Dim unAngleAval As Single, unAngleAmont As Single, unAngle As Single, unAnglePourTest As Single
  Dim Angle As Single, LE4m As Single, LI As Single, LS As Single, Rg As Single
  Dim unNumAval, unNumAmont As Integer
  Dim linBordVoieEntréeI As Single, linBordVoieSortieI As Single
  Dim linBordVoieEntréeAmont As Single, linBordVoieSortieAmont As Single
  Dim linBordVoieEntréeAval As Single, linBordVoieSortieAval As Single
  Dim LI2 As Single
 
  VerifierAngleBranche = True
  Rg = gbRayonExt
  Angle = wBranche.Angle
  LE4m = wBranche.LE4m
  LI = wBranche.LI
  LS = wBranche.LS
  Select Case TypeControleActif
    Case TYPE_ANGLE: Angle = x
    Case TYPE_LE4M
      If x > 0 And LE4m = 0 Then
        LI = DEFAUT_LI 'Remet la valeur par défaut de LI
      End If
      LE4m = x
    Case TYPE_LI: LI = x
    Case TYPE_LS
      If x > 0 And LS = 0 Then
        LI = DEFAUT_LI 'Remet la valeur par défaut de LI
      End If
      LS = x
    Case Else: Rg = x
  End Select

  message = ""
  LI2 = LI / 2# ' Moitié d'ilot réparti uniformément de chaque coté de l'axe
  'Calcul des angles de la branche concernée en fonction des valeurs
  ' LI, LS et LE
  'angle de la branche stockée dans les valeurs
  'linBordVoieEntrée, linBordVoieSortie
  'Calcul des numéros de branches amont et aval
  unNumAmont = (NumBrancheSelect - 1) Mod NbBranches
  If unNumAmont = 0 Then unNumAmont = NbBranches
  unNumAval = NumBrancheSelect Mod NbBranches + 1
  VerifierAngleBranche = calculBord(0, LE4m, LI2, LS, Rg, linBordVoieEntréeI, linBordVoieSortieI)
  If VerifierAngleBranche Then
    VerifierAngleBranche = calculBord(unNumAmont, LE4m, LI2, LS, Rg, linBordVoieEntréeAmont, linBordVoieSortieAmont)
  End If
  If VerifierAngleBranche Then
    VerifierAngleBranche = calculBord(unNumAval, LE4m, LI2, LS, Rg, linBordVoieEntréeAval, linBordVoieSortieAval)
  End If
  If Not VerifierAngleBranche Then
    'La modification du giratoire ou des branches ne peut être réalisée
    message = IDv_RgOuUneBrancheIncorrect
  Else
    '-------------------------------------------------------------
    'Vérification des angles des branches
    unAngle = angConv(Angle, True)

    '------------------------------------------------------------------------------------
    'Test si la branche dépasse sa branche amont ou aval
    ' On met des moins pour avoir des angles > 0 car le repère
    'écran est indirect (X > 0 vers la droite et Y > 0 vers le bas)
    unAngleAval = CSng(linBranche(unNumAval).Tag)
    unAngleAmont = CSng(linBranche(unNumAmont).Tag)
    unAngleAval = unAngleAval - linBordVoieSortieAval
    unAngleAmont = unAngleAmont + linBordVoieEntréeAmont
    If unAngleAval < unAngleAmont Then
      'Cas du déplacement d'une branche ayant un angle aval < à l'angle amont
      'car les angles stockés sont entre 0 et 2PI
      unAngleAval = unAngleAval + 2 * PI
      If unAngle < unAngleAmont Then
        unAnglePourTest = unAngle + 2 * PI
      Else
        unAnglePourTest = unAngle
      End If
    Else
      unAnglePourTest = unAngle
    End If
    
    Dim unAngleMini As Single
    Dim unAngleMaxi As Single
    'Calcule les angles de bords de chaussée de la branche à vérifier...
    unAngleMini = unAnglePourTest - linBordVoieSortieI
    unAngleMaxi = unAnglePourTest + linBordVoieEntréeI
    'retourne True si entre angle amont et angle aval, faux sinon
    VerifierAngleBranche = (unAngleMini > unAngleAmont And unAngleMaxi < unAngleAval)
    If Not VerifierAngleBranche Then
      If unAngleMini <= unAngleAmont Then
        message = IDv_Chevauchement + GiratoireProjet.colBranches.Item(unNumAmont).nom + _
          IDl_ET + GiratoireProjet.colBranches.Item(NumBrancheSelect).nom & "."
      Else
        message = IDv_Chevauchement + GiratoireProjet.colBranches.Item(NumBrancheSelect).nom + _
          IDl_ET + GiratoireProjet.colBranches.Item(unNumAval).nom & "."
      End If
    End If
  End If
End Function


