VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "ss32x25.ocx"
Begin VB.Form frmRésultats 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Visualisation des résultats"
   ClientHeight    =   8310
   ClientLeft      =   1530
   ClientTop       =   435
   ClientWidth     =   11460
   Icon            =   "Résultats.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8310
   ScaleWidth      =   11460
   Begin VB.TextBox txtPériode 
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   480
      TabIndex        =   27
      Top             =   960
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton cmdFermer 
      Caption         =   "Close"
      Height          =   375
      Left            =   600
      TabIndex        =   26
      Top             =   3960
      Width           =   1695
   End
   Begin VB.CommandButton cmdCourbes 
      Caption         =   "&Courbes"
      Height          =   375
      Left            =   600
      TabIndex        =   25
      Top             =   2760
      Width           =   1695
   End
   Begin VB.CommandButton cmdSaturer 
      Caption         =   "&Saturer la branche"
      Height          =   375
      Left            =   600
      TabIndex        =   24
      Top             =   3360
      Width           =   1695
   End
   Begin VB.PictureBox picRésultat 
      Align           =   2  'Align Bottom
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3495
      Left            =   0
      ScaleHeight     =   3495
      ScaleWidth      =   11460
      TabIndex        =   6
      Top             =   4815
      Width           =   11460
      Begin FPSpread.vaSpread vgdRecap 
         Height          =   1935
         Left            =   1320
         TabIndex        =   7
         Top             =   1320
         Width           =   7335
         _Version        =   131077
         _ExtentX        =   12938
         _ExtentY        =   3413
         _StockProps     =   64
         AutoSize        =   -1  'True
         DisplayColHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   7
         MaxRows         =   8
         MoveActiveOnFocus=   0   'False
         OperationMode   =   1
         ScrollBars      =   0
         SelectBlockOptions=   2
         SpreadDesigner  =   "Résultats.frx":030A
         UnitType        =   2
         UserResize      =   1
         VisibleCols     =   500
         VisibleRows     =   500
      End
      Begin VB.Frame fraRésultats 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   975
         Left            =   3000
         TabIndex        =   8
         Top             =   200
         Width           =   6735
         Begin VB.TextBox txtRC 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   14
            Text            =   "RC_UVP"
            Top             =   960
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.TextBox txtRCpc 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   1200
            TabIndex        =   13
            Text            =   "RC_%"
            Top             =   960
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.TextBox txtLK 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   2400
            TabIndex        =   12
            Text            =   "LK"
            Top             =   960
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.TextBox txtLKM 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   3480
            TabIndex        =   11
            Text            =   "LKM"
            Top             =   960
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.TextBox txtTMA 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   4680
            TabIndex        =   10
            Text            =   "TMA"
            Top             =   960
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.TextBox txtTTA 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   5760
            TabIndex        =   9
            Text            =   "TTA"
            Top             =   960
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label lblTA 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "average"
            Height          =   255
            Index           =   1
            Left            =   4440
            TabIndex        =   23
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label lblTA 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "total"
            Height          =   255
            Index           =   2
            Left            =   5520
            TabIndex        =   22
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label lblTA 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Queuing delay"
            Height          =   255
            Index           =   0
            Left            =   4440
            TabIndex        =   21
            Top             =   360
            Width           =   2175
         End
         Begin VB.Label lblLK 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "average"
            Height          =   255
            Index           =   1
            Left            =   2280
            TabIndex        =   20
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label lblLK 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "maximum"
            Height          =   255
            Index           =   2
            Left            =   3360
            TabIndex        =   19
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label lblLK 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Queue length"
            Height          =   255
            Index           =   0
            Left            =   2280
            TabIndex        =   18
            Top             =   360
            Width           =   2175
         End
         Begin VB.Label lblRC 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "in pcu/h"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   17
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label lblRC 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "in %"
            Height          =   255
            Index           =   2
            Left            =   1200
            TabIndex        =   16
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label lblRC 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Reserve Capacity"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   15
            Top             =   360
            Width           =   2175
         End
      End
   End
   Begin VB.PictureBox picCourbes 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   480
      ScaleHeight     =   675
      ScaleWidth      =   1755
      TabIndex        =   5
      Top             =   1560
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.ComboBox cboPériode 
      Height          =   315
      Left            =   480
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   480
      Width           =   2775
   End
   Begin TabDlg.SSTab tabRésultat 
      Height          =   4935
      Left            =   3600
      TabIndex        =   0
      Top             =   120
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   8705
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Opération"
      TabPicture(0)   =   "Résultats.frx":058E
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtRqFonctionnement"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Comments on period"
      TabPicture(1)   =   "Résultats.frx":05AA
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtRqTrafics"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Comments on conception"
      TabPicture(2)   =   "Résultats.frx":05C6
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "txtRqConceptions"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.TextBox txtRqConceptions 
         ForeColor       =   &H80000011&
         Height          =   4215
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   480
         Width           =   7095
      End
      Begin VB.TextBox txtRqTrafics 
         ForeColor       =   &H80000011&
         Height          =   4215
         Left            =   -74760
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   480
         Width           =   7095
      End
      Begin VB.TextBox txtRqFonctionnement 
         ForeColor       =   &H80000011&
         Height          =   4215
         Left            =   -74760
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   480
         Width           =   7095
      End
   End
   Begin VB.Menu mnuBarre 
      Caption         =   "&File"
      Index           =   0
      Begin VB.Menu mnuFichier 
         Caption         =   "&Close"
         Index           =   0
      End
      Begin VB.Menu mnuFichier 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuFichier 
         Caption         =   "Printer set&up"
         Index           =   2
      End
      Begin VB.Menu mnuFichier 
         Caption         =   "&Print..."
         Index           =   3
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFichier 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuFichier 
         Caption         =   "E&xit"
         Index           =   5
      End
   End
   Begin VB.Menu mnuBarre 
      Caption         =   "&Tools"
      Index           =   1
      Begin VB.Menu mnuTools 
         Caption         =   "&Saturer Branche"
         Index           =   0
      End
      Begin VB.Menu mnuTools 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuTools 
         Caption         =   "&Curves"
         Index           =   2
      End
      Begin VB.Menu mnuTools 
         Caption         =   "Curve &options"
         Index           =   3
      End
   End
   Begin VB.Menu mnuBarre 
      Caption         =   "Window"
      Index           =   2
      WindowList      =   -1  'True
      Begin VB.Menu mnuFenetre 
         Caption         =   "&Cascade"
         Index           =   0
      End
      Begin VB.Menu mnuFenetre 
         Caption         =   "Tile &horizontally"
         Index           =   1
      End
      Begin VB.Menu mnuFenetre 
         Caption         =   "Tile &vertically"
         Index           =   2
      End
   End
   Begin VB.Menu mnuBarre 
      Caption         =   "&?"
      Index           =   3
      Begin VB.Menu mnuAide 
         Caption         =   "&Help"
         Index           =   0
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuAide 
         Caption         =   "Help &on"
         Index           =   1
      End
      Begin VB.Menu mnuAide 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuAide 
         Caption         =   "About "
         Index           =   3
      End
   End
End
Attribute VB_Name = "frmRésultats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
'*
'*          Projet GIRABASE - CERTU - CETE de l'Ouest
'*
'*          Module de feuille : RESULTATS.FRM - frmRésultats
'*
'*          Feuille d'affichage de résultats de GIRABASE
'*
'*          Feuille fille de la feuille MDIGirabase
'*
'*          Un objet GIRATOIRE est associé à chaque feuille FrmRésultats
'*          La propriété Résultats de cet objet désigne cette feuille
'*
'*          La feuille Résultats peut être chargée temporairement par la fonction
'*          d'impression de la géométrie : appel de AfficheConception pour récupérer l'onglet Conception
'*
'**************************************************************************************************

Option Explicit

Public GiratoireProjet As GIRATOIRE
Public FeuilleLégende As Form
Private LégendeVue As Boolean
Private ChargementEnCours As Boolean

'Drapeau indiquant que c'est le module frmImprimer qui charge la feuille pour récupérer l'onglet Conception
Public Impressions As Boolean

Public NuméroBrancheActive As Integer
Private NuméroTraficActif As Integer
Private TraficActif As TRAFIC

Private NbBranches As Integer
Private Echelle As Single

'******************************************************************************
' Choix d'une période dans la liste
'******************************************************************************
Private Sub cboPériode_Click()
  If cboPériode.ListIndex = -1 Then Exit Sub
  
  NuméroTraficActif = cboPériode.ListIndex + 1
  Set TraficActif = GiratoireProjet.colTrafics.Item(NuméroTraficActif)
  
  With TraficActif
    ' Le calcul peut n'être pas fait pour une période saturée, ou encore pour une période de trafic modifiée depuis le dernier calcul
    If Not .CalculFait And .EstComplète Then
      GiratoireProjet.CalculCapacité TraficActif
    End If
  End With
  
  If TraficActif.BrancheSaturée = 0 Then
    cmdSaturer.Caption = IDl_SaturerBranche
  Else
    cmdSaturer.Caption = IDl_SupprimerPériode
  End If
  mnuTools(0).Enabled = (TraficActif.BrancheSaturée <> 0)
  mnuTools(0).Caption = cmdSaturer.Caption
  
  AfficheRésultats
  AfficheRemarques
  CourbeCapacité
  vgdRecap.SetFocus
  
  If Not TraficActif.EstComplète Then
    tabRésultat.Tab = 1
    txtRqTrafics.ForeColor = vbRed
    txtRqFonctionnement.ForeColor = vbRed
    txtPériode = cboPériode
    txtPériode.Visible = True
  Else
    txtRqTrafics.ForeColor = vbGrayText
    txtRqFonctionnement.ForeColor = vbGrayText
    txtPériode.Visible = False
  End If
  
End Sub

'******************************************************************************
' Bascule d'affichage des Courbes et des Remarques
'******************************************************************************
Private Sub cmdCourbes_Click()
  If picCourbes.Visible Then
    picCourbes.Visible = False
    tabRésultat.Visible = True
    cmdCourbes.Caption = IDl_Courbes
    mnuTools(2).Caption = IDl_Courbes
    mnuTools(2).HelpContextID = IDhlp_Courbes
    ChoixAideOnglet
  Else
    picCourbes.Visible = True
    tabRésultat.Visible = False
    cmdCourbes.Caption = IDl_Remarques
    mnuTools(2).Caption = IDl_Remarques
    mnuTools(2).HelpContextID = 1
    If Not EstChargée(FeuilleLégende) And Not LégendeVue Then mnuTools_Click 3: LégendeVue = True
    ChoixAideOnglet Courbes:=True
  End If
  
End Sub

Public Sub CourbeCapacité()
Dim I As Integer
Const MARGE = 100
Dim Espace As Integer
Dim h As Single
Dim Orx As Single
Dim Ory As Single
Dim X As Double
Dim Y As Double
Dim wTrafic As TRAFIC
 
  
  Espace = 250  ' Espace entre 2 graduations de 100 en 100
  Espace = picCourbes.Width / 31
  Echelle = Espace / 100

  h = picCourbes.Height - MARGE
  Orx = MARGE
  Ory = MARGE

  picCourbes.Cls
  picCourbes.ForeColor = QBColor(9)

  ' Traits d'axes
  picCourbes.Line (MARGE, h)-(trEchelX(2700, Orx), h), QBColor(1)
  picCourbes.CurrentX = picCourbes.CurrentX + Espace * 0.5
  picCourbes.CurrentY = h - Espace * 1.5
  picCourbes.CurrentY = h - Espace * 2
  picCourbes.Print IDl_Trafic
  picCourbes.CurrentX = trEchelX(2700, Orx) + Espace * 0.5
  picCourbes.Print IDl_Genant
  picCourbes.Line (MARGE, h)-(MARGE, trEchelY(1800, Ory)), QBColor(1)
  picCourbes.CurrentY = picCourbes.CurrentY - Espace
  picCourbes.CurrentX = MARGE
  picCourbes.Print IDl_Trafic & " " & IDl_Entrant

  ' Graduations
  For I = 1 To 27
    If I Mod 5 = 0 Then
      picCourbes.Line (MARGE + Espace * I, h)-(MARGE + Espace * I, h - 2 * MARGE)
      picCourbes.CurrentX = picCourbes.CurrentX - Espace * 0.5
      picCourbes.CurrentY = picCourbes.CurrentY - Espace * 0.5
      picCourbes.Print CStr(I * 100)
    Else
      picCourbes.Line (MARGE + Espace * I, h)-(MARGE + Espace * I, h - MARGE)
    End If
  Next

  For I = 1 To 18
    If I Mod 5 = 0 Then
      picCourbes.Line (MARGE, h - Espace * I)-(MARGE * 3, h - Espace * I)
      picCourbes.CurrentY = picCourbes.CurrentY - Espace * 0.5
      picCourbes.Print CStr(I * 100)
    Else
      picCourbes.Line (MARGE, h - Espace * I)-(MARGE * 2, h - Espace * I)
    End If
  Next

' S'il n'y a aucune période complète : pas de branche active
  If NuméroBrancheActive = 0 Then Exit Sub

    ' Dessiner la courbe
    Dim numPt As Integer
    numPt = 0
    Y = GiratoireProjet.colBranches.Item(NuméroBrancheActive).getCVH(numPt)
    If Y <= 1800 Then
      picCourbes.PSet (trEchelX(0, Orx), trEchelY(Y, Ory))
    Else
      While Y > 1800
        numPt = numPt + 10
        Y = GiratoireProjet.colBranches.Item(NuméroBrancheActive).getCVH(numPt)
      Wend
      picCourbes.PSet (trEchelX(numPt, Orx), trEchelY(Y, Ory))
    End If
    For I = numPt + 10 To 2700 Step 10
      Y = GiratoireProjet.colBranches.Item(NuméroBrancheActive).getCVH(I)
      picCourbes.Line -(trEchelX(I, Orx), trEchelY(Y, Ory))
    Next

    'Représenter les points (QG,QE) pour chaque période
    For I = 1 To cboPériode.ListCount
      Set wTrafic = GiratoireProjet.colTrafics.Item(I)
      With wTrafic
        If .EstComplète Then
          If Not .CalculFait Then GiratoireProjet.CalculCapacité wTrafic
          X = .getQG(NuméroBrancheActive)
          Y = .getQE(NuméroBrancheActive) / (1 - .getCp(NuméroBrancheActive))
          If GiratoireProjet.colBranches.Item(NuméroBrancheActive).TAD Then Y = Y - .getQ(NuméroBrancheActive, BrancheSuivant(NuméroBrancheActive))
          picCourbes.FillColor = .CouleurCourbe
          If Y < 2000 Then
            picCourbes.Circle (trEchelX(X, Orx), trEchelY(Y, Ory)), 50, .CouleurCourbe
          Else
            picCourbes.Circle (trEchelX(X, Orx), trEchelY(1850, Ory)), 50, .CouleurCourbe
            picCourbes.CurrentY = picCourbes.CurrentY - 50
            picCourbes.Line -(picCourbes.CurrentX, picCourbes.CurrentY - 150), QBColor(0)
            picCourbes.Line -(picCourbes.CurrentX - 50, picCourbes.CurrentY + 50), QBColor(0)
            picCourbes.Line -(picCourbes.CurrentX + 100, picCourbes.CurrentY), QBColor(0)
            picCourbes.Line -(picCourbes.CurrentX - 50, picCourbes.CurrentY - 50), QBColor(0)
          End If
        End If
      End With
    Next

End Sub

Private Function trEchelX(ByVal X As Single, Optional ByVal Decalage As Single = 0)
  trEchelX = Echelle * X + Decalage
End Function

Private Function trEchelY(ByVal Y As Single, Optional ByVal Decalage As Single = 0)
  trEchelY = picCourbes.Height - (Echelle * Y + Decalage)
End Function

Private Function arrondi(ByVal Valeur As Single, ByVal sens As Integer)
  arrondi = (CInt((Valeur - 50) / 100) + sens) * 100
End Function

Private Sub cmdFermer_Click()
  Unload Me
End Sub

'******************************************************************************
' Saturation d'une branche
'******************************************************************************
Private Sub cmdSaturer_Click()
Dim nomPériode As String
Dim I As Integer

  If cmdSaturer.Caption = IDl_SaturerBranche Then
    nomPériode = TraficActif.nom & IDl_AbrévSaturBranche & CStr(NuméroBrancheActive)
    For I = 0 To cboPériode.ListCount
      If nomPériode = cboPériode.List(I) Then
        cboPériode.ListIndex = I
        Exit Sub
      End If
    Next
    
    GiratoireProjet.colTrafics.add GiratoireProjet, nomPériode, NuméroBrancheActive    ' NuméroBrancheActive : pour indiquer la branche à saturation

    With frmTrafMult
      Set .TraficOrigine = TraficActif
      .Saturation = True
      .NumBranche = NuméroBrancheActive
      .cmdOK = True   ' Déclenche Form_Load de frmTrafMult sans l'afficher
      .Saturation = False
    End With
    If EstChargée(FeuilleLégende) Then FeuilleLégende.ChargerPériode GiratoireProjet.colTrafics.count, True
  Else
    ' Suppression de la période
    If MsgBox(IDm_SupprPériode & " " & TraficActif.nom, vbYesNo + vbDefaultButton2 + vbQuestion + vbMsgBoxHelpButton, , App.HelpFile, IDhlp_SaturerBranche) = vbYes Then
      GiratoireProjet.colTrafics.Remove GiratoireProjet, NuméroTraficActif
      If EstChargée(FeuilleLégende) Then FeuilleLégende.DéchargerPériode NuméroTraficActif
      With cboPériode
        .ListIndex = Min(NuméroTraficActif, .ListCount) - 1 ' déclenche l'évènement cboPériode_Click
      End With
    End If
End If

End Sub

'******************************************************************************
' Activation de la feuille
'******************************************************************************
Private Sub Form_Activate()
  
  If Not GiratoireProjet.Données.ValiderFeuilleDonnées Then Unload Me: Exit Sub

  ' Recalcul éventuel
  If Not GiratoireProjet.CalculFait Then GiratoireProjet.CalculCapacité
  
  Set gbProjetActif = GiratoireProjet
  
  ' Choix de la période : le même que celui de la feuille de données
  Dim numPériode As Integer
  numPériode = GiratoireProjet.Données.cboPériode.ListIndex
  cboPériode.ListIndex = numPériode
  
  If EstChargée(FeuilleLégende) Then
    With FeuilleLégende
      If .lblNom.UBound <> GiratoireProjet.colTrafics.count Then
      ' Il y a eu des périodes d'ajoutées ou d'enlevées : Rechargement de la fenêtre
        Dim Top As Single, Left As Single
        Top = .Top
        Left = .Left
        Unload FeuilleLégende
        .Show vbModeless, MDIGirabase
        .Move Left, Top
      Else
        .Visible = True
      End If
    End With
  End If
End Sub

'******************************************************************************
' Désactivation de la feuille
'******************************************************************************
Private Sub Form_Deactivate()
  If EstChargée(FeuilleLégende) Then FeuilleLégende.Visible = False
  cboPériode.ListIndex = -1
  ' Aide contextuelle
  'HelpContextID = MDIGirabase.HelpContextID
End Sub

'******************************************************************************
' Chargement de la feuille
'******************************************************************************
Private Sub Form_Load()
Dim I As Integer
Dim j As Integer
Dim k As Integer
Dim Pris As Boolean
Dim Couleur As Long

 ChargementEnCours = True
 
  Set GiratoireProjet = gbProjetActif
  With GiratoireProjet
    NbBranches = .NbBranches
    If Impressions Then Me.WindowState = vbMinimized: Exit Sub ' Appel temporaire par le module d'impression pour l'onglet Conception
    Caption = .Données.Caption & " - " & IDl_Résultats
    With .colTrafics
      For I = 1 To .count
        cboPériode.AddItem .Item(I).nom
        k = 10
        While .Item(I).CouleurCourbe = 0 And k < 16
          Couleur = QBColor(k)
          For j = 1 To I - 1
            If .Item(j).CouleurCourbe = Couleur Then Pris = True
          Next
          If Not Pris Then .Item(I).CouleurCourbe = Couleur
          k = k + 1
          Pris = False
        Wend
      Next
    End With
  End With
  
  ' Aide en ligne
  mnuFichier(3).HelpContextID = IDhlp_ImprimerRésultats
  vgdRecap.HelpContextID = IDhlp_Résultats
  mnuTools(0).HelpContextID = IDhlp_SaturerBranche
  mnuTools(3).HelpContextID = IDhlp_Courbes
  mnuAide(3).Caption = mnuAide(3).Caption & App.Title & "..."
    
  SetDeviceIndependentWindow Me
  
  ' Dimensionnement du tableau de résultats selon le nombre de branches
  Dimensionner
  DoEvents

  ' Positionner la partie haute sur la gauche de la feuille
  cboPériode.Left = 200
  txtPériode.Left = 200
  txtPériode.Top = cboPériode.Top
  tabRésultat.Left = cboPériode.Left + cboPériode.Width + 200

  ' Dimensionner le conteneur de résultats et y positionner le tableau en hauteur
  picRésultat.Height = vgdRecap.Height + 1400

  If Screen.TwipsPerPixelX > 12 And gbProjetActif.NbBranches > 6 Then
  ' 800*600 avec 7 ou 8 branches : on tasse un peut tout le bas
    With picRésultat
      .Height = vgdRecap.Height + fraRésultats.Height
    End With
    tabRésultat.Top = 20
    fraRésultats.Top = 0
    vgdRecap.Top = 1100
    vgdRecap.Top = fraRésultats.Height
  Else
    vgdRecap.Top = 1300
  End If

  Dim HOnglet As Long
  HOnglet = tabRésultat.Height + 2 * tabRésultat.Top '+ 120
  If HOnglet > picRésultat.Top Then
    ' La partie haute de la feuille 'mord' sur la partie basse
    Dim deltaH As Long
    Dim deltaW As Long
    Dim Coef As Single
    deltaH = tabRésultat.Height - txtRqConceptions.Height
    deltaW = tabRésultat.Width - txtRqConceptions.Width
    Coef = tabRésultat.Width / tabRésultat.Height
    ' diminuer la hauter de l'onglet pour qu'il 'tienne'
    tabRésultat.Height = picRésultat.Top - (HOnglet - tabRésultat.Height)
    tabRésultat.Width = tabRésultat.Height * Coef
    ' redimensionner les textbox en conséquence
    txtRqConceptions.Height = tabRésultat.Height - deltaH
    txtRqFonctionnement.Height = tabRésultat.Height - deltaH
    txtRqTrafics.Height = tabRésultat.Height - deltaH
    txtRqConceptions.Width = tabRésultat.Width - deltaW
    txtRqFonctionnement.Width = tabRésultat.Width - deltaW
    txtRqTrafics.Width = tabRésultat.Width - deltaW
  Else
    ' Ajuster la feuille en hauteur = Hauteur partie haute (HOnglet) + Hauteur partie basse
    Height = HOnglet + (Height - picRésultat.Top)
  End If
  
    ' Ajuster la feuille en largeur : partie haute et partie basse réunies
  Width = Max(tabRésultat.Left + tabRésultat.Width + 300, vgdRecap.Width + 400)
    ' Centrer le tableau en largeur
  vgdRecap.Left = tabRésultat.Left + tabRésultat.Width - vgdRecap.Width
    ' Cadrer le titre du tableau en face des colonnes
  fraRésultats.Left = vgdRecap.Left + vgdRecap.Width - fraRésultats.Width
  
' Superposition de l'image des courbes sur le groupe d'onglets des remarques
  With tabRésultat
    picCourbes.Move .Left, .Top, .Width, .Height
  End With

  cmdCourbes.Caption = IDl_Courbes
  
  Set FeuilleLégende = New frmOptCourbes
  tabRésultat.Tab = 2
  ChargementEnCours = False
  
End Sub

Private Sub Form_Resize()

  If ChargementEnCours Then Exit Sub
  
  'On ne fait pas de traitements lors d'une mise en icone
  If WindowState <> vbMinimized Then picRésultat.Height = Height - (tabRésultat.Top + tabRésultat.Height + 400)
  
End Sub

'******************************************************************************
' Déchargement de la feuille
'******************************************************************************
Private Sub Form_Unload(Cancel As Integer)

If Impressions Then Exit Sub

  With GiratoireProjet
    .delPériodeSaturées
    Set .Résultats = Nothing
    .CalculFait = False
    If EstChargée(FeuilleLégende) Then Unload FeuilleLégende
  End With
End Sub

'*************************************************************************************
          'Menu Aide
'*************************************************************************************
Private Sub mnuAide_Click(Index As Integer)
Dim retour%

  Select Case Index
  Case 0 'Aide
    With MDIGirabase.dlgFichier
      If HelpContextID = 0 Then
     ' HELP_FINDER : constante non fournie par le CommonDialog
        .HelpCommand = &HB
      Else                          ' Menu contextuel
        .HelpCommand = cdlHelpContext
        .HelpContext = HelpContextID
      End If
      .ShowHelp
    End With
            
  Case 1 'Aide sur...
    With MDIGirabase.dlgFichier
      .HelpCommand = cdlHelpPartialKey
      .ShowHelp
    End With
            
  Case 3 'A propos de Girabase
    frmApropos.Show 1
    
  End Select

End Sub

'*************************************************************************************
          'Menu Fenêtre
'*************************************************************************************
Private Sub mnuFenetre_Click(Index As Integer)
    MDIGirabase.Arrange Index
End Sub

'*************************************************************************************
          'Menu Fichier
'*************************************************************************************
Private Sub mnuFichier_Click(Index As Integer)
  Select Case Index
  Case 0  'Fermer
    Unload Me
  
  Case 2 'Configuration de l'impression
    ConfigImprimante
  Case 3 'Impression des résultats
    frmImprimer.Show vbModal
  Case 5 'Quitter
    'Quitte l'application
    Unload MDIGirabase

  End Select
  
End Sub

'*************************************************************************************
          'Menu Outils
'*************************************************************************************
Private Sub mnuTools_Click(Index As Integer)
  Select Case Index
  Case 0 ' Saturer une branche
    cmdSaturer = True
  Case 1 ' Suppression d'une période saturée
    cmdSaturer = True
  Case 2 ' Bascule Courbes/Remarques
    cmdCourbes = True
  Case 3 ' Options des courbes
    If Not FeuilleLégende.Visible Then
      FeuilleLégende.Show vbModeless, MDIGirabase
    End If
  End Select
End Sub

'*************************************************************************************
          'Clic droit sur la courbe : Menu contextuel
'*************************************************************************************
Private Sub picCourbes_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = vbRightButton Then PopupMenu mnuBarre(1), , , , mnuTools(3)
End Sub

'*************************************************************************************
          'Changement d'onglet
'*************************************************************************************
Private Sub tabRésultat_Click(PreviousTab As Integer)
  With tabRésultat
    If (PreviousTab = 0) Then .ZOrder
    If .Tab = 0 Then .ZOrder 1
    ChoixAideOnglet
  End With
End Sub

Private Sub ChoixAideOnglet(Optional ByVal Courbes As Boolean)

  If Courbes Then
    HelpContextID = IDhlp_Courbes
  Else
    Select Case tabRésultat.Tab
    ' Onglet fonctionnement
      Case 0
      HelpContextID = IDhlp_OngletFonctionnement
  
    ' Onglet période
      Case 1
      HelpContextID = IDhlp_OngletPériode
      ' Onglet conception
    Case 2
      HelpContextID = IDhlp_OngletConception
    End Select
    tabRésultat.HelpContextID = HelpContextID
  End If
  
  MDIGirabase.HelpContextID = HelpContextID

End Sub
'******************************************************************************
' Affichage des remarques dans les onglets dédiés
' Ne concerne que la période active
'******************************************************************************
Public Sub AfficheRemarques()
  'Affiche onglet Conception
  AfficheConception
  
  If TraficActif.EstComplète Then
    'Affiche onglet des périodes
    AffichePériode
    'Affiche l'onglet de fonctionnement
    AfficheFonctionnement
  Else
    'Les matrices de saisie relatives à la période en cours sont incomplètes,
    'on n'affiche aucun résultat (sauf la Conception : AV - 19.05.99)
    txtRqTrafics = IDc_TraficsIncomplets
    txtRqFonctionnement = IDc_TraficsIncomplets
  End If

End Sub

'******************************************************************************
' Affichage des remarques dans l'onglet de fonctionnement
' Ces remarques ne concernent que la branche et la période actives
'******************************************************************************

Private Sub AfficheFonctionnement()
Dim NumBranche As Integer
Dim QEntrant As Integer
Dim QSortant As Integer
Dim QTournant As Integer
Dim RCsurC As Single
Dim QImmédiat As Integer

  txtRqFonctionnement.Text = ""
  For NumBranche = 1 To NbBranches
    txtRqFonctionnement = txtRqFonctionnement & IDl_Branche & GiratoireProjet.colBranches.Item(NumBranche).nom & vbCrLf
    
    If TraficActif.BrancheSaturée = NumBranche Then
    'Branche issue d'une Matrice par saturation
    '------------------------------------------
      txtRqFonctionnement = txtRqFonctionnement & IDc_MatriceSaturation & vbCrLf
    ElseIf GiratoireProjet.colBranches.Item(NumBranche).EntréeNulle Then
      txtRqFonctionnement = txtRqFonctionnement & IDm_BrancheSortie & vbCrLf
  
    Else
      With TraficActif
        QImmédiat = .getQ(NumBranche, BrancheSuivant(NumBranche))
        QSortant = .getQS(NumBranche)
        QEntrant = .getQE(NumBranche)
        QTournant = .getQTournant(NumBranche)
      End With
    
      With GiratoireProjet.colBranches.Item(NumBranche)
      
        If .TAD Then
          QEntrant = QEntrant - QImmédiat
          QImmédiat = 0
        End If
        If gbProjetActif.colBranches.Item(BranchePrécédent(NumBranche)).TAD Then
          QSortant = QSortant - TraficActif.getQ(BranchePrécédent(NumBranche), NumBranche)
        End If
        
        If .SortieNulle Then txtRqFonctionnement = txtRqFonctionnement & IDm_BrancheEntrée & vbCrLf
        
        'Tests sur la largeur de l'îlot séparateur LI
        '--------------------------------------------
        '(Ce test apparaît aussi dans l'onglet "Conception")
        If .LI < 0.85 And Not .EntréeNulle And Not .SortieNulle And GiratoireProjet.R = 0 Then
          txtRqFonctionnement = txtRqFonctionnement & IDc_IlotASeparer & vbCrLf
        End If
            
        'Test sur la largeur de sortie LS
        '--------------------------------
        'Rq0499 prise en compte de LAU
        Dim LAU As Single
        LAU = GiratoireProjet.LAU
        If .LS < 6 Then 'Correction bogue dans le test non conforme au CdC
          If QSortant > 1200 Then
            txtRqFonctionnement = txtRqFonctionnement & IDc_LS2voiesN
            If GiratoireProjet.Milieu <> rc Then txtRqFonctionnement = txtRqFonctionnement & IDc_TraverséePiétons
            txtRqFonctionnement = txtRqFonctionnement & vbCrLf
          ElseIf (QSortant > 600 And LAU >= 10.5) Or _
            (QSortant > Max(3 * TraficActif.getQTournant(NumBranche), 900) And _
            LAU > 8 And GiratoireProjet.Milieu <> CV) Then
              txtRqFonctionnement = txtRqFonctionnement & IDc_LS2voiesP
              If GiratoireProjet.Milieu <> rc Then txtRqFonctionnement = txtRqFonctionnement & IDc_TraverséePiétons
              txtRqFonctionnement = txtRqFonctionnement & vbCrLf
          End If
        End If
        
        'LE4m Remplacé par LE : faire modification dans CDC
         Dim LE As Single          ' Largeur d'entrée
         If .EntréeEvasée Then
           LE = (.LE4m + .LE15m) / 2
         Else
           LE = .LE4m
         End If
        
        'Test sur réserve de capacité RC
        '-------------------------------
        
        RCsurC = 100# * TraficActif.getRC(NumBranche) / TraficActif.getC(NumBranche)
        
        If RCsurC < 15 Then
          'Première partie du test
          If RCsurC >= 0 Then
            txtRqFonctionnement = txtRqFonctionnement & IDc_RCfaible & vbCrLf
          Else
            txtRqFonctionnement = txtRqFonctionnement & IDc_RCnégative & vbCrLf
          End If
          Dim LEU As Single
          LEU = GiratoireProjet.LEU
          
          'Deuxième partie du test
          Dim Avertissement As Boolean
          Avertissement = False
          If QImmédiat > 0.33 * QEntrant And _
                          QEntrant > 600 Then
            txtRqFonctionnement = txtRqFonctionnement & IDc_RC1 & vbCrLf
            Avertissement = True
          End If
    
          If LE > 6 And LE < Min(LEU, 9) And _
            QEntrant > 0.5 * TraficActif.getQG(NumBranche) Then
            txtRqFonctionnement = txtRqFonctionnement & IDc_RC3
            If GiratoireProjet.Milieu <> rc Then txtRqFonctionnement = txtRqFonctionnement & IDc_RC3p
            txtRqFonctionnement = txtRqFonctionnement & vbCrLf
            Avertissement = True
          End If
          If LE < Min(LEU, 6) And _
            QEntrant > 0.5 * TraficActif.getQG(NumBranche) Then
            txtRqFonctionnement = txtRqFonctionnement & IDc_RC2
            If GiratoireProjet.Milieu <> rc Then txtRqFonctionnement = txtRqFonctionnement & IDc_RC2p
            txtRqFonctionnement = txtRqFonctionnement & vbCrLf
            Avertissement = True
          End If
          If LE > LEU Then
            txtRqFonctionnement = txtRqFonctionnement & IDc_RC4 & vbCrLf
            Avertissement = True
          End If
          If .KS > 0.5 And QSortant > 300 And _
            QSortant < QTournant Then
            txtRqFonctionnement = txtRqFonctionnement & IDc_RC5 & vbCrLf
            Avertissement = True
          End If
          If Not Avertissement Then
            txtRqFonctionnement = txtRqFonctionnement & IDc_RC6 & vbCrLf
          End If
        End If
        'Test RC suite et fin...
        Dim Q1000 As Boolean
        Q1000 = False
        Dim I As Integer
        For I = 1 To NbBranches
          If TraficActif.getQ(NumBranche, I) >= 1000 Then Q1000 = True
        Next I
        If Q1000 Then
          txtRqFonctionnement = txtRqFonctionnement & IDc_RC11 & vbCrLf
        End If
        'RCsurC testé par rapport à 50% (en non 100% comme stipulé dans CdC)
        If RCsurC > 50 And LE > 6 And LE < 9 _
          And QEntrant < 1000 Then
          If GiratoireProjet.Milieu <> CV Then
            txtRqFonctionnement = txtRqFonctionnement & IDc_RC12 & vbCrLf
          Else
            txtRqFonctionnement = txtRqFonctionnement & IDc_RC13 & vbCrLf
          End If
        End If
        If RCsurC > 50 And LE >= 9 Then
            txtRqFonctionnement = txtRqFonctionnement & IDc_RC14 & vbCrLf
        End If
        
        'Test du temps moyen d'attente TMA
        '---------------------------------
        Dim TMA As Single
        TMA = TraficActif.getTMA(NumBranche)
        If TMA >= 120 And GiratoireProjet.Milieu <> rc Then
          txtRqFonctionnement = txtRqFonctionnement & IDc_TMA2 & vbCrLf
        ElseIf TMA >= 60 And GiratoireProjet.Milieu <> rc Then
          txtRqFonctionnement = txtRqFonctionnement & IDc_TMA1 & vbCrLf
        ElseIf TMA >= 40 And GiratoireProjet.Milieu = rc Then
          txtRqFonctionnement = txtRqFonctionnement & IDc_TMA2 & vbCrLf
        ElseIf TMA >= 20 And GiratoireProjet.Milieu = rc Then
          txtRqFonctionnement = txtRqFonctionnement & IDc_TMA1 & vbCrLf
        End If
        'Test de la longueur de stockage LK
        '----------------------------------
        Dim LK, LKM As Single
        LK = TraficActif.getLK(NumBranche)
        LKM = TraficActif.getLKM(NumBranche)
        If LK >= 40 And GiratoireProjet.Milieu <> rc Then
          txtRqFonctionnement = txtRqFonctionnement & IDc_LK4 & vbCrLf
        ElseIf LK >= 20 And GiratoireProjet.Milieu <> rc Then
          txtRqFonctionnement = txtRqFonctionnement & IDc_LK3 & vbCrLf
        ElseIf LK >= 20 And GiratoireProjet.Milieu = rc Then
          txtRqFonctionnement = txtRqFonctionnement & IDc_LK2 & vbCrLf
        ElseIf LK >= 10 And GiratoireProjet.Milieu = rc Then
          txtRqFonctionnement = txtRqFonctionnement & IDc_LK1 & vbCrLf
        End If
        
        'Test LKM
        '--------
        If LKM >= 30 And LK < 20 And GiratoireProjet.Milieu <> rc Then
          txtRqFonctionnement = txtRqFonctionnement & IDc_LK6 & vbCrLf
        ElseIf LKM >= 25 And LK < 10 And GiratoireProjet.Milieu = rc Then
          txtRqFonctionnement = txtRqFonctionnement & IDc_LK5 & vbCrLf
        End If
      End With
      
    End If
    txtRqFonctionnement = txtRqFonctionnement & vbCrLf
  Next

End Sub

'******************************************************************************
' Affichage des remarques dans l'onglet "Remarque de conception"
' On affiche d'abord les remarques générales concernant le site, puis l'anneau
' puis les remarques concernant chacune des branches (AfficheConceptionBranches)
' txtRqConceptions est récupérable par la fonction d'impression (d'où : Public)
'******************************************************************************
Public Sub AfficheConception()
  Dim I, iMax As Integer
  Dim LE4mMax As Single
  
  txtRqConceptions.Text = ""
  
  With GiratoireProjet
    
    'Recherche du LE4m Maxi
    LE4mMax = 0
    For I = 1 To NbBranches
      If .colBranches.Item(I).LE4m > LE4mMax Then
        LE4mMax = .colBranches.Item(I).LE4m
        iMax = I
      End If
    Next I
    'Contrôle Type de site
    '---------------------
    If .Milieu = rc And NbBranches > 6 Then
      txtRqConceptions = IDm_TropDeBranchesEnRC & vbCrLf
    End If
    'Contrôle du rayon infranchissable du giratoire R
    '------------------------------------------------
    If .R > 25 Then
      txtRqConceptions = txtRqConceptions & _
        IDm_RTropGrand & " " & IDc_RTropGrand2 + vbCrLf
    End If
    'Contrôle de la largeur d'anneau
    '-------------------------------
    If .LA < 6 And .Milieu = rc Then
      txtRqConceptions = txtRqConceptions & IDm_LATropEtroit & vbCrLf
    ElseIf (.LA > 9 And .Milieu = rc) Or (.LA > 12) Then
      txtRqConceptions = txtRqConceptions & IDm_LATropGrand & vbCrLf
    End If
    'Contrôle de la largeur d'anneau LA comparée aux largeurs d'entrée LE
    '--------------------------------------------------------------------
    'If .LA < 1.2 * LE4mMax Then 'Rq0499
    If .LAU < 1.2 * LE4mMax Then
      txtRqConceptions = _
        txtRqConceptions & IDm_LATropEtroitPourEntrer & .colBranches.Item(iMax).nom & vbCrLf
    End If
    'RQ 17/5/99
    'If .LA < 6 And .Bf < 1.5 And (.Milieu = PU Or .Milieu = CV) Then
    If .LA + .Bf < 7 And .Milieu <> rc Then
      txtRqConceptions = txtRqConceptions & IDm_LATropEtroit & vbCrLf
    End If
    'Contrôle de la Bande franchissable
    Dim Rg As Single
    Rg = .R + .Bf + .LA
    If (Rg > 12 And Rg < 15) And (.Bf < 1.5 Or .Bf > 2.5) Then
      txtRqConceptions = txtRqConceptions & IDm_Bf & vbCrLf
    End If
    
    'Contrôle du rayon extérieur Rg
    '------------------------------
    If (Rg > 12 And Rg < 15) And .Milieu = rc Then
      txtRqConceptions = txtRqConceptions & IDm_RgVoirGirationEnRC & vbCrLf
    ElseIf (Rg > 12 And Rg < 15) Then
      txtRqConceptions = txtRqConceptions & IDm_RgVoirGiration & vbCrLf
    End If
    'Contrôle du rayon Rg et des largeurs d'entrée
    '---------------------------------------------
    'If LE4mMax >= 6 And Rg < 20 And .Milieu = RC Then
    '??Rq0499
    If LE4mMax >= 6 And LE4mMax < 8 And Rg < 20 And .Milieu = rc Then
      txtRqConceptions = txtRqConceptions & IDm_RgTropPetit & vbCrLf
    End If
  End With
  If txtRqConceptions <> "" Then txtRqConceptions = txtRqConceptions + vbCrLf
  AfficheConceptionBranches
  
End Sub

'******************************************************************************
' Affichage des remarques relatives à chacune des branches
' Onglet Remarque de conception
'******************************************************************************
Private Sub AfficheConceptionBranches()
    Dim NumeroBranche As Integer
    Dim numErreur As Integer
    Dim MessageBranche As String
    'Test par branche
    For NumeroBranche = 1 To NbBranches
      With GiratoireProjet.colBranches.Item(NumeroBranche)
        MessageBranche = ""
        
        'Contrôle de la largeur d'entrée à 4 mètres
        '------------------------------------------
        If Not .EntréeNulle Then
          If .LE4m < 3 Then
            MessageBranche = IDm_LEPetit & vbCrLf
          ElseIf .LE4m > 8 And GiratoireProjet.Milieu = rc Then
            MessageBranche = MessageBranche & IDm_LETropLargeEnRC & vbCrLf
          End If
          'Rq0499
          If .LE4m >= 9 Then
            MessageBranche = MessageBranche & IDm_LETropLargePourPiétons & vbCrLf
          End If
        End If
        'Contrôle de la largeur de sortie
        '--------------------------------
        If Not .SortieNulle Then
          If .LS < 3.5 Then
            MessageBranche = MessageBranche & IDm_LSPetit & vbCrLf
          ElseIf .LS > 7 Then
            MessageBranche = MessageBranche & IDm_LSTropLarge & vbCrLf
          End If
        End If
        'Contrôle du rapport LE4m / LE15m
        '--------------------------------
        If .EntréeEvasée Then
          Dim rapport As Single
          rapport = .LE4m / .LE15m
          If GiratoireProjet.Milieu = rc And rapport > 1 Then
            MessageBranche = MessageBranche & IDm_EvasementEnRC & vbCrLf
          ElseIf GiratoireProjet.Milieu <> rc And rapport > 1.5 Then
            MessageBranche = MessageBranche & IDm_EvasementTropPetit & vbCrLf
          End If
        End If
        'Contrôle de la largeur de l'îlot séparateur
        '-------------------------------------------
        'If .LI < 2 And .LE4m > 0 And .LS > 0 And
        If .LI < 2 And Not .EntréeNulle And Not .SortieNulle And _
          GiratoireProjet.Milieu = CV And GiratoireProjet.R > 0 Then
            MessageBranche = MessageBranche & IDm_LITropPetit & vbCrLf
        End If
        'Contrôle de la largeur de l'îlot séparateur par rapport à LImax
        If .LI > GiratoireProjet.LImax Then
          MessageBranche = MessageBranche & IDm_LITropGrand & vbCrLf
        End If
        'Contrôle de la largeur d'îlot LI à partir de R,LE4m et LS
        '---------------------------------------------------------
        'If Not .EntréeNulle And Not .SortieNulle Then
          'If .LI < 0.85 And .LE4m > 0 And .LS > 0 And GiratoireProjet.R = 0 Then
        If .LI < 0.85 And Not .EntréeNulle And Not .SortieNulle And GiratoireProjet.R = 0 Then
          MessageBranche = MessageBranche & IDc_IlotASeparer & vbCrLf
        End If
        'Contrôle des angles des branches pour un mini-giratoire
        '-------------------------------------------------------
        If GiratoireProjet.R = 0 Then
          numErreur = GiratoireProjet.Données.ControleCarBranches(NumeroBranche, False)
          If numErreur < 0 Then
            MessageBranche = MessageBranche & IDm_AngleTropPetitPourMiniG & vbCrLf
          ElseIf numErreur > 0 Then
            MessageBranche = MessageBranche & IDm_AnglePourMiniG & vbCrLf
          End If
        End If
        'Contrôle à partir des données de trafics
        'Comparaison de la largeur d'ilot avec le trafic piétons
        'Le message n'est affiché qu'une seule fois, même s'il y a plusieurs erreurs
        If Not .EntréeNulle And Not .SortieNulle Then
          'Rq0499
          'If .LI < TraficActif.getQP(NumeroBranche) / 300 Then

'          If .LI < TraficActif.getQP(NumeroBranche) / 300 And _
'            (GiratoireProjet.Milieu <> CV Or .LI >= 2) Then
'            MessageBranche = MessageBranche & IDc_IlotEtroit & vbCrLf
'          End If
          
          ' Retraitement pour utilisation dans le module d'impression (AV - 19.05.99)
          Dim wTrafic As TRAFIC
          Dim I As Integer
          For I = 1 To GiratoireProjet.nbPériodes
            Set wTrafic = GiratoireProjet.colTrafics.Item(I)
            If wTrafic.EstComplète And GiratoireProjet.Milieu <> rc Then
              If .LI < wTrafic.getQP(NumeroBranche) / 300 And _
              (GiratoireProjet.Milieu <> CV Or .LI >= 2) Then
              MessageBranche = MessageBranche & IDc_IlotEtroit & vbCrLf
              Exit For
              End If
            End If
          Next
        End If
        
        If MessageBranche <> "" Then
          txtRqConceptions = txtRqConceptions & IDl_Branche & .nom & vbCrLf & MessageBranche & vbCrLf
        End If
      End With
    Next NumeroBranche 'Fin du parcours des branches
End Sub

'******************************************************************************
' Affichage des remarques relatives aux trafics
' Onglet Remarques sur les trafics
' Affiche seulement des informations relatives à la période active
'******************************************************************************
Public Sub AffichePériode()
  Dim numéroBranche, j As Integer
  Dim MessageBranche As String
  'Test sur le trafic total du giratoire
  '-------------------------------------
  txtRqTrafics = GiratoireProjet.Données.ControleTraficTotal(NuméroTraficActif)
  If txtRqTrafics <> "" Then txtRqTrafics = txtRqTrafics & vbCrLf & vbCrLf
  'Parcours de l'ensemble des branches
  For numéroBranche = 1 To NbBranches
    With GiratoireProjet.colBranches.Item(numéroBranche)
      MessageBranche = ""
      'Il y a sur la branche un trafic qui ne justifie pas la présence d'un tourne-à-droite
      '------------------------------------------------------------------------------------
      j = numéroBranche Mod NbBranches + 1
      If TraficActif.getQ(numéroBranche, j) < 100 And .TAD Then
        MessageBranche = IDm_QTropPetitPourTAD & vbCrLf
      End If
      'Il y a sur la branche un trafic entrant nul alors que la largeur d'entrée est non nulle
      '---------------------------------------------------------------------------------------
      'If TraficActif.getQE(numéroBranche) <= 0 And .LE4m > 0 Then
      If TraficActif.getQE(numéroBranche) <= 0 And Not .EntréeNulle Then
         MessageBranche = MessageBranche & IDc_QEnul & vbCrLf
      End If
      'Il a sur la branche un trafic sortant nul alors que la largeur de sortie n'est pas nulle
      '----------------------------------------------------------------------------------------
      If TraficActif.getQS(numéroBranche) <= 0 And Not .SortieNulle Then
         MessageBranche = MessageBranche & IDc_QSnul & vbCrLf
      End If
      'Affichage des messages
      If MessageBranche <> "" Then
        txtRqTrafics = txtRqTrafics & IDl_Branche & .nom & vbCrLf & MessageBranche & vbCrLf
      End If
    End With
  Next numéroBranche
End Sub

'****************************************************************************************
'Affichage des résultats de capacité de toutes les branches pour la période de trafic active
'****************************************************************************************
Private Sub AfficheRésultats()
Dim Chaine As String
Dim nbSec As Long
Dim nbMin As Integer
Dim nbHeures As Long
Dim NumBranche As Integer
Dim rc As Double


With vgdRecap
  If TraficActif.EstComplète Then
    ' Affichage des résultats de chaque branche
    For NumBranche = 1 To NbBranches
      If GiratoireProjet.colBranches.Item(NumBranche).EntréeNulle Then
      Else
        .Row = NumBranche
        .Col = 2
        rc = TraficActif.getRC(NumBranche)
        If TraficActif.BrancheSaturée = NumBranche Then
          rc = 0
        Else
          rc = TraficActif.getRC(NumBranche)
        End If
        If rc < 0.5 Then .ForeColor = vbRed Else .ForeColor = vbBlack
        .Value = Format(rc, "####0")
        .Col = 3
        If rc <= 0.5 Then .ForeColor = vbRed Else .ForeColor = vbBlack
        .Value = Format(rc / TraficActif.getC(NumBranche), "####0%")
        If .Value = "%" Then .Value = ""
        .Col = 4
        .Value = Format(TraficActif.getLK(NumBranche), "###0") & IDl_VehiculeHeure
        .Col = 5
        .Value = Format(TraficActif.getLKM(NumBranche), "####0") & IDl_VehiculeHeure
        .Col = 6
        .Value = Format(TraficActif.getTMA(NumBranche), "#0") & "s"
        nbSec = TraficActif.getTTA(NumBranche)
        nbHeures = nbSec \ 3600
        If nbHeures > 0 Then Chaine = CStr(nbHeures) & "h"
        nbSec = nbSec - CLng(nbHeures) * 3600
        nbMin = nbSec \ 60
        Chaine = Chaine & CStr(nbMin) & "m"
        nbSec = nbSec - nbMin * 60
        Chaine = Chaine & CStr(nbSec) & "s"
        .Col = 7
        .Value = Chaine
        If nbSec > 30 Then nbMin = nbMin + 1
        If nbHeures > 100 Then
          .Value = Format(nbHeures + nbMin / 60, "###0h")
        Else
          .Value = Format(nbHeures + nbMin / 60, "###0.0h")
        End If
      End If
    Next
    ' Mode de sélection par ligne entière et sélection de la branche 1
    .SelectBlockOptions = 2
    SelectRow NuméroBrancheActive

  Else  ' Matrice de trafic incomplète (indicateur True)
    SelectRow NuméroBrancheActive, True
    .SelectBlockOptions = 0
    Dim I As Integer
    For NumBranche = 1 To NbBranches
      .Row = NumBranche
      For I = 2 To .MaxCols
        .Col = I
        .Value = ""
      Next
    Next
  End If

End With
  
End Sub

'****************************************************************************************
'Dimensionnement du tableau de synthèse
'****************************************************************************************
Private Sub Dimensionner()
Dim I As Integer

  With vgdRecap
    .OperationMode = OperationModeSingle
    .MaxRows = NbBranches
    .LockBackColor = vbGrayText
    .ColWidth(1) = cboPériode.Width * 0.85 * 1.6
    .ColWidth(1) = TextWidth(String(25, "E"))
    For I = 2 To .MaxCols
      .ColWidth(I) = txtTTA(0).Width
      .ColWidth(I) = fraRésultats.Width / 6.09
    Next
  End With
  For I = 1 To NbBranches
    vgdRecap.Row = I
    vgdRecap.Col = 1
    vgdRecap.Value = GiratoireProjet.colBranches.Item(I).nom
  Next
End Sub

Private Sub txtRqConceptions_GotFocus()
  HelpContextID = IDhlp_OngletConception
End Sub

Private Sub txtRqFonctionnement_GotFocus()
  HelpContextID = IDhlp_OngletFonctionnement
End Sub

Private Sub txtRqTrafics_GotFocus()
  HelpContextID = IDhlp_OngletPériode
End Sub

'****************************************************************************************
'Sélection d'une branche par l'utilisateur
'****************************************************************************************
Public Sub vgdRecap_Click(ByVal Col As Long, ByVal Row As Long)

  If NuméroBrancheActive = 0 Then Exit Sub
  If NuméroBrancheActive = Row Then Exit Sub
  
  If GiratoireProjet.colBranches.Item(Row).EntréeNulle Then
    Beep
    SelectRow NuméroBrancheActive
    Exit Sub
  End If
  
  SelectRow Row
'  AfficheFonctionnement
  CourbeCapacité
End Sub

'****************************************************************************************
' Mise en évidence de la Sélection de la branche
'****************************************************************************************
Public Sub SelectRow(ByVal Row As Long, Optional ByVal MatriceIncomplète)
  If Row = 0 Then Row = CalculPremiereBranche
  NuméroBrancheActive = Row
  With vgdRecap
    .Row = Row
    .Row2 = Row
    .Col = 0
    .Col2 = .MaxCols
    If IsMissing(MatriceIncomplète) Then
      .Action = 2 ' Mise en évidence
      .Col = 2
      If .Value = "" Then
        cmdSaturer.Enabled = False
      Else
        cmdSaturer.Enabled = (.Value < 0)
      End If
      cmdSaturer.Enabled = (cmdSaturer.Caption = IDl_SupprimerPériode) Or cmdSaturer.Enabled
      mnuTools(0).Enabled = cmdSaturer.Enabled
    Else
      .OperationMode = OperationModeNormal
      .Action = 14  ' Désélection de la ligne
      .OperationMode = OperationModeSingle
      NuméroBrancheActive = 0   ' Pour ignorerl'évènement vgdRecap
      cmdSaturer.Enabled = False
      mnuTools(0).Enabled = False
    End If
    .Visible = True
    DoEvents
  End With
  
End Sub

'****************************************************************************************
'Détermination de la première branche sélectionnable (Elimination des branches de sortie)
'****************************************************************************************
Private Function CalculPremiereBranche() As Long
  Dim num As Integer
  num = 1
  With GiratoireProjet.colBranches
    While .Item(num).EntréeNulle
      num = num + 1
    Wend
  End With
  CalculPremiereBranche = num
End Function


Private Sub vgdRecap_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
  If NewRow <> -1 And NewRow <> NuméroBrancheActive Then vgdRecap_Click 1, NewRow
End Sub
