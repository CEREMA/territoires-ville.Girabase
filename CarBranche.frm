VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmCarBranche 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Caractéristiques de la branche"
   ClientHeight    =   5370
   ClientLeft      =   1440
   ClientTop       =   3645
   ClientWidth     =   4950
   Icon            =   "CarBranche.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5370
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Aide"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   26
      Top             =   4680
      Width           =   1215
   End
   Begin VB.TextBox txtNumBranche 
      Height          =   285
      Left            =   4320
      MaxLength       =   1
      TabIndex        =   24
      Top             =   240
      Width           =   270
   End
   Begin VB.TextBox txtLE15m 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1200
      TabIndex        =   6
      Top             =   3720
      Width           =   495
   End
   Begin VB.CheckBox chkEntréeEvasée 
      Height          =   375
      Left            =   3960
      TabIndex        =   9
      Top             =   3720
      Width           =   375
   End
   Begin VB.TextBox txtEcart 
      Height          =   285
      Left            =   3360
      TabIndex        =   2
      Text            =   "8"
      Top             =   600
      Width           =   495
   End
   Begin VB.CheckBox chkRampe 
      Caption         =   "Rampe Supérieure à 3 %"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Value           =   1  'Checked
      Width           =   2415
   End
   Begin VB.CheckBox chkTAD 
      Caption         =   "Voie Tourne à Droite"
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
      TabIndex        =   4
      Top             =   1680
      Width           =   2415
   End
   Begin VB.TextBox txtNomBranche 
      Height          =   285
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
   Begin VB.TextBox txtLI 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2040
      MaxLength       =   8
      TabIndex        =   7
      Top             =   3720
      Width           =   495
   End
   Begin VB.TextBox txtLE4m 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   480
      MaxLength       =   8
      TabIndex        =   5
      Top             =   3720
      Width           =   495
   End
   Begin VB.TextBox txtLS 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2880
      MaxLength       =   8
      TabIndex        =   8
      Top             =   3720
      Width           =   495
   End
   Begin VB.TextBox txtAngleBranche 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   840
      MaxLength       =   7
      TabIndex        =   1
      Top             =   600
      Width           =   735
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Fermer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   10
      Top             =   4680
      Width           =   1215
   End
   Begin MSComCtl2.UpDown spnNumBranche 
      Height          =   285
      Left            =   4680
      TabIndex        =   23
      Top             =   240
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   503
      _Version        =   393216
      Value           =   1
      BuddyControl    =   "txtNumBranche"
      BuddyDispid     =   196610
      OrigLeft        =   4200
      OrigTop         =   240
      OrigRight       =   4440
      OrigBottom      =   615
      Max             =   8
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin VB.Label lblNuméro 
      Caption         =   "Numéro :"
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
      Left            =   3480
      TabIndex        =   25
      Top             =   240
      Width           =   855
   End
   Begin VB.Label lblEntréeEvasée 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Entrée Evasée"
      Height          =   495
      Left            =   3720
      TabIndex        =   22
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label lblUnitAngle 
      AutoSize        =   -1  'True
      Height          =   195
      Index           =   1
      Left            =   3960
      TabIndex        =   21
      Top             =   600
      Width           =   165
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   1815
      Left            =   120
      Top             =   2400
      Width           =   4575
   End
   Begin VB.Label lblEntrée 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Entrée"
      Height          =   255
      Left            =   360
      TabIndex        =   20
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label lblLI 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ilôt"
      Height          =   495
      Left            =   1920
      TabIndex        =   19
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label lblLS 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sortie"
      Height          =   495
      Left            =   2760
      TabIndex        =   18
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label lblLargeurs 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Largeurs"
      Height          =   255
      Left            =   360
      TabIndex        =   17
      Top             =   2760
      Width           =   3135
   End
   Begin VB.Label lblLE4m 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "à 4 m"
      Height          =   255
      Left            =   360
      TabIndex        =   16
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label lblLE15m 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "à 15 m"
      Height          =   255
      Left            =   1080
      TabIndex        =   15
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label lblEcart 
      Alignment       =   1  'Right Justify
      Caption         =   "Ecart : "
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
      Left            =   2300
      TabIndex        =   14
      Top             =   600
      Width           =   1000
   End
   Begin VB.Label lblNom 
      AutoSize        =   -1  'True
      Caption         =   "Nom :"
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
      Left            =   120
      TabIndex        =   13
      Top             =   240
      Width           =   615
   End
   Begin VB.Label lblUnitAngle 
      AutoSize        =   -1  'True
      Height          =   195
      Index           =   0
      Left            =   1680
      TabIndex        =   12
      Top             =   600
      Width           =   285
   End
   Begin VB.Label lblAngle 
      AutoSize        =   -1  'True
      Caption         =   "Angle :"
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
      Left            =   120
      TabIndex        =   11
      Top             =   600
      Width           =   615
   End
End
Attribute VB_Name = "frmCarBranche"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************************
'*
'*          Projet GIRABASE - CERTU - CETE de l'Ouest
'*
'*          Module de feuille : CARBRANCHE.FRM - frmCarBranche
'*
'*          Feuille de Saisie de l'ensemble des caractéristiques d'une branche (Angle et Dimensionnement)
'*
'* Les modifications faites dans cette feuille de saisie sont transmises dans les matrices associées.
'* La validation des données et la modification graphique du dessin du giratoire sont déclenchées
'* à partir des spreads.
'***********************************************************************************************************

Option Explicit
Private TitreInitial As String
Private NbBranches As Integer
Private SauveValeur As String
Private Valeur As String

Private FeuilleDonnées As Form
Private BrancheActive As BRANCHE
Private FeuilleChargée As Boolean


'******************************************************************************
' Détection de la frappe d'une touche
'******************************************************************************
Private Sub FrappeTouche()
  If Not FeuilleDonnées.DonnéeModifiée Then
    'Affecte la couleur normale au controle au premier caractère frappé
    ActiveControl.ForeColor = vbWindowText
    'Change la proprité de modification de la donnée
    FeuilleDonnées.DonnéeModifiée = True
  End If
End Sub

'***********************************************************************************************
' Validation de la largeur saisie : Mise à jour du tableau vgdLargBranche dans la feuille Données
'***********************************************************************************************
Private Function ValideLargeur(Valeur As Variant, Col As Integer) As Boolean
  Dim wBranche As BRANCHE
  ValideLargeur = True
  With FeuilleDonnées
    If .DonnéeModifiée Then
      'Si la largeur a été modifiée, on attribue la nouvelle valeur
      'à la cellule concernée du spread de dimensionnement
      'puis on simule la validation de la cellule par déclenchement de l'événement
      'Leave-Cell
      'La modification graphique sera opérée à partir du spread.
      .NuméroLigneActive = monNumBrancheSelect
      Set wBranche = gbProjetActif.colBranches.Item(monNumBrancheSelect)
      If .ValidationDonnées(Valeur, wBranche) Then
        .TypeMatriceActive = DIMENSION
        .vgdLargBranche.Col = Col
        .vgdLargBranche.Row = monNumBrancheSelect
        .vgdLargBranche.Value = Valeur
        .DonnéeModifiée = True
        'Déclenchement de LeaveCell
        'La valeur est validée et le giratoire est redessiné
        .vgdLargBranche_LeaveCell Col, monNumBrancheSelect, -1, -1, False
        If FeuilleDonnées.MessageEmis Then
          'Il y a eu un message d'avertissement,
          'La cellule doit être coloriée en rouge
          ActiveControl.ForeColor = vbRed
        End If
      Else
        ValideLargeur = False
        ActiveControl.Text = SauveValeur
      End If
      .DonnéeModifiée = False
    End If
  End With
End Function

'********************************************************************************************
' Validation de l'angle saisi : Mise à jour du tableau vgdCarBranche dans la feuille Données
'********************************************************************************************
Private Function ValideAngle(Valeur As Variant, Col As Integer) As Boolean
  Dim wBranche As BRANCHE
  ValideAngle = True
  With FeuilleDonnées
    If .DonnéeModifiée Then
      'Si l'angle a été modifiée, on attribue la nouvelle valeur
      'à la cellule concernée du spread de caractéristiques
      'puis on simule la validation de la cellule par déclenchement de l'événement
      'Leave-Cell
      'La modification graphique sera opérée à partir du spread.
      .TypeControleActif = TYPE_ANGLE
      Set wBranche = gbProjetActif.colBranches.Item(monNumBrancheSelect)
      If .ValidationDonnées(Valeur, wBranche) Then
        .TypeMatriceActive = BRANCHE
        .vgdCarBranche.Col = Col
        .vgdCarBranche.Row = monNumBrancheSelect
        .vgdCarBranche.Value = Valeur
        .DonnéeModifiée = True
        'Déclenchement de LeaveCell
        'La valeur est validée et le giratoire est redessiné
        .vgdCarBranche_LeaveCell Col, monNumBrancheSelect, -1, -1, False
        If FeuilleDonnées.MessageEmis Then
          'Il y a eu un message d'avertissement,
          'La cellule doit être coloriée en rouge
          ActiveControl.ForeColor = vbRed
        End If
        txtAngleBranche = CStr(BrancheActive.Angle)
        .vgdCarBranche.Col = 3
        .vgdCarBranche.Row = monNumBrancheSelect
        txtEcart = .vgdCarBranche.Value
      Else
        ValideAngle = False
        ActiveControl.Text = SauveValeur
      End If
      'Si la modification n'est pas validée, on repasse à l'état non modifié
      .DonnéeModifiée = False 'Si la modification n'est pas validée,
    End If
  End With
End Function

'******************************************************************************
' Case à cocher Entrée Evasée
'******************************************************************************
Private Sub chkEntréeEvasée_Click()
  txtLE15m.Enabled = (chkEntréeEvasée = vbChecked)
  If chkEntréeEvasée Then
    txtLE15m.BackColor = vbWhite
    txtLE15m = txtLE4m
  Else
    txtLE15m.BackColor = vbGrayText
    txtLE15m = ""
  End If
  FeuilleDonnées.TypeControleActif = TYPE_ENTREE
  'Déclenchement du clic dans la case à cocher du spread des branches du giratoire
  FeuilleDonnées.vgdLargBrancheClic monNumBrancheSelect, chkEntréeEvasée
End Sub

'******************************************************************************
' Case à cocher Rampe
' Validation de la présence ou l'absence de rampe dans le spread
' des caractéristiques des branches
'******************************************************************************
Private Sub chkRampe_Click()

  With FeuilleDonnées.vgdCarBranche
    .Col = 4
    .Row = monNumBrancheSelect
    .Value = (chkRampe = vbChecked)
    BrancheActive.Rampe = .Value
  End With
End Sub

'******************************************************************************
' Case à cocher  Tourne à droite
' Validation de la présence ou l'absence de TAD dans le spread
' des branches
'******************************************************************************
Private Sub chkTAD_Click()
  With FeuilleDonnées.vgdCarBranche
    .Col = 5
    .Row = monNumBrancheSelect
    .Value = (chkTAD = vbChecked)
    BrancheActive.TAD = .Value
  End With
End Sub

Private Sub cmdHelp_Click()

    SendKeys "{F1}", True

End Sub

'******************************************************************************
' Bouton Fermer
'******************************************************************************
Private Sub cmdOK_Click()
  'A la fermeture de la boite de dialogue, il faut préciser que le contrôle sur
  ' les spreads de caractéristiques des branches est terminé...
  FeuilleDonnées.TypeControleActif = TYPE_AUCUN
  FeuilleDonnées.TypeMatriceActive = AUCUN
  Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  FrappeTouche
End Sub

'******************************************************************************
' Chargement de la feuille
'******************************************************************************
Private Sub Form_Load()
Dim Ecart As Integer
Dim I As Integer
  
  'Icon = MDIGirabase.Icon
  
  ' Aide contextuelle
  HelpContextID = IDhlp_CarBranche
  
  TitreInitial = Caption
  FeuilleChargée = False
  With gbProjetActif
    For I = 0 To 1
      lblUnitAngle(I) = libelAngle(.modeangle)
    Next
    Set FeuilleDonnées = .Données
    NbBranches = .NbBranches
  End With
   
  With FeuilleDonnées
    Set .FeuilleBranche = Me
    With .tabDonnées
      Me.Move .Left + MDIGirabase.Left, .Top + MDIGirabase.Top + 200
    End With
    .tabDonnées.Visible = False
  End With
    
  spnNumBranche.Max = NbBranches
  txtNumBranche = monNumBrancheSelect
  
End Sub

'******************************************************************************
'******************************************************************************
' Déchargement de la feuille
'******************************************************************************
'******************************************************************************
Private Sub Form_Unload(Cancel As Integer)
  FeuilleDonnées.shpPoignée.Visible = False
  FeuilleDonnées.tabDonnées.Visible = True
  Set FeuilleDonnées.FeuilleBranche = Nothing
End Sub

Private Sub txtAngleBranche_GotFocus()
  FeuilleDonnées.TypeControleActif = TYPE_ANGLE
  InitControle FeuilleDonnées.TypeControleActif
End Sub

Private Sub txtAngleBranche_Validate(Cancel As Boolean)

  Cancel = Not ValideAngle(txtAngleBranche, 2)
End Sub



Private Sub txtEcart_GotFocus()
  FeuilleDonnées.TypeControleActif = TYPE_ANGLE
  InitControle FeuilleDonnées.TypeControleActif
End Sub

Private Sub txtEcart_Validate(Cancel As Boolean)
  Cancel = Not ValideAngle(txtEcart, 3)
End Sub

Private Sub txtLE15m_GotFocus()
  FeuilleDonnées.TypeControleActif = TYPE_LE15M
  InitControle FeuilleDonnées.TypeControleActif
End Sub


Private Sub txtLE15m_Validate(Cancel As Boolean)
  Cancel = Not ValideLargeur(txtLE15m, 2)
End Sub

Private Sub txtLE4m_GotFocus()
  FeuilleDonnées.TypeControleActif = TYPE_LE4M
  InitControle FeuilleDonnées.TypeControleActif
End Sub


Private Sub txtLE4m_Validate(Cancel As Boolean)
  Dim EntréePrécédente As Boolean
  EntréePrécédente = BrancheActive.EntréeNulle
  If ValideLargeur(txtLE4m, 1) Then
    Cancel = False
    ChangeLE4m EntréePrécédente, BrancheActive.EntréeNulle
    If Not EntréePrécédente And BrancheActive.EntréeNulle Then
      InitControle (ActiveControl)
    End If
  Else
    Cancel = True
  End If
End Sub
'******************************************************************************
' InitControle
' Met le contrôle dans la couleur normale, le passe en surbrillance
' et sauvegarde sa valeur
'******************************************************************************
Public Sub InitControle(ByVal controle As String)
  With ActiveControl
    If controle = TYPE_ANGLE Then
      FeuilleDonnées.TypeMatriceActive = BRANCHE
    Else
      FeuilleDonnées.TypeMatriceActive = DIMENSION
    End If
    FeuilleDonnées.ControleRecommandations True, controle
    .ForeColor = vbWindowText
    SauveValeur = .Text
   End With
End Sub

Private Sub txtLI_GotFocus()
  FeuilleDonnées.TypeControleActif = TYPE_LI
  InitControle FeuilleDonnées.TypeControleActif
End Sub

Private Sub txtLI_Validate(Cancel As Boolean)
  Cancel = Not ValideLargeur(txtLI, 3)
End Sub

Private Sub txtLS_GotFocus()
  FeuilleDonnées.TypeControleActif = TYPE_LS
  InitControle FeuilleDonnées.TypeControleActif
   
End Sub

Private Sub txtLS_Validate(Cancel As Boolean)
  Dim SortiePrécédente As Boolean
  SortiePrécédente = BrancheActive.SortieNulle
  If ValideLargeur(txtLS, 4) Then
    Cancel = False
    ChangeLS SortiePrécédente, BrancheActive.SortieNulle
  Else
    Cancel = True
  End If
End Sub

Private Sub txtNomBranche_GotFocus()
  FeuilleDonnées.TypeControleActif = TYPE_AUCUN
  InitControle FeuilleDonnées.TypeControleActif
End Sub

Private Sub txtNomBranche_KeyPress(KeyAscii As Integer)
  FrappeTouche
End Sub

Private Sub txtNomBranche_Validate(Cancel As Boolean)
Dim Angle As Single

  BrancheActive.nom = txtNomBranche
  With FeuilleDonnées.vgdCarBranche
    .Col = 1
    .Row = monNumBrancheSelect
    .Value = txtNomBranche
    FeuilleDonnées.lblLibelléBranche(.Row) = txtNomBranche
    MDIGirabase.mnuBranche(.Row - 1).Caption = "&" & CStr(.Row) & " " & txtNomBranche
  End With
  Angle = angConv(BrancheActive.Angle, CVRADIAN)
  With FeuilleDonnées
    DéplacerNomBranche .lblLibelléBranche(monNumBrancheSelect), .linBranche(monNumBrancheSelect), Cos(Angle), -Sin(Angle)        ' "-" pour le sinus : car l'axe des Y est vers le bas
  End With
End Sub

'******************************************************************************
' Changement de branche
'******************************************************************************
Private Sub txtNumBranche_Change()
Static Passage As Boolean
Dim Ecart As Single
  If Not IsNumeric(txtNumBranche) Then Passage = True: txtNumBranche = monNumBrancheSelect: Exit Sub
  If Passage Then Passage = False: Exit Sub
  
  monNumBrancheSelect = CInt(txtNumBranche)
    
  'Déplacement de la poignée de sélection : Emprunté au module DessinGiratoire : ModiferBranche
  Dim wLigne As Line
 
  Set wLigne = FeuilleDonnées.linBranche(monNumBrancheSelect)
  With FeuilleDonnées.shpPoignée
    .Left = wLigne.X2 - .Width / 2
    .Top = wLigne.Y2 - .Height / 2
  End With
  
  If monNumBrancheSelect = 1 Then Else
  With gbProjetActif.colBranches
    Set BrancheActive = .Item(monNumBrancheSelect)
    If monNumBrancheSelect > 1 Then
      Ecart = BrancheActive.Angle - .Item(monNumBrancheSelect - 1).Angle
      txtEcart.Enabled = True
      txtEcart.BackColor = vbWhite
      txtAngleBranche.Enabled = True
      txtAngleBranche.BackColor = vbWhite
    Else
      Ecart = 0
      txtEcart.Enabled = False
      txtEcart.BackColor = vbGrayText
      txtAngleBranche.Enabled = False
      txtAngleBranche.BackColor = vbGrayText
    End If
  End With
  
  With BrancheActive
    'Récupération et affichage des valeurs de la branche sélectionnée
    Caption = TitreInitial & " " & CStr(monNumBrancheSelect)
    txtNomBranche = .nom
    chkRampe = RetourneEntier(.Rampe)
    Dim Save As Single
    Save = .LE15m
    chkEntréeEvasée = RetourneEntier(.EntréeEvasée)
    .LE15m = Save 'Récupère la valeur de LE15m modifiée lors du clic
    chkTAD = RetourneEntier(.TAD)
    txtLE4m = CStr(.LE4m)
    ChangeLE4m .EntréeNulle, .EntréeNulle
    If .EntréeEvasée Then
      txtLE15m = CStr(.LE15m)
    Else
      txtLE15m.Enabled = False
      txtLE15m.BackColor = vbGrayText
    End If
    txtLI = CStr(.LI)
    txtLS = CStr(.LS)
    
    txtAngleBranche = CStr(.Angle)
    txtEcart = CStr(Ecart)
    VérifieChangeBranche
    
 End With

End Sub

'********************************************************************************
' Traitement final lors du changement de branche (Evènement txtNumBranche_Change)
'********************************************************************************
Private Sub VérifieChangeBranche()
  Dim Control As Control
  DoEvents
  With FeuilleDonnées
    'Remet les contrôles TextBox dans leur couleur initiale
    For Each Control In Controls
      If TypeOf Control Is TextBox Then
        Control.ForeColor = vbWindowText
      End If
    Next
    If FeuilleChargée Then
      txtNomBranche.SetFocus
    Else
      FeuilleChargée = True
    End If
    .TypeControleActif = TYPE_AUCUN
    .NuméroLigneActive = monNumBrancheSelect
    .TypeMatriceActive = BRANCHE
    .ControleRecommandations True, .TypeControleActif
    .TypeMatriceActive = DIMENSION
    .ControleRecommandations True, .TypeControleActif
  End With
End Sub

'******************************************************************************
' ChangeLE4m
' Si l'entrée est nulle, il faut refuser le TAD, l'entrée évasée
' et l'entrée de la largeur d'ilot LI
' Si l'entrée n'est pas nulle, on accepte le TAD si le giratoire
' n'est pas un mini-giratoire (R>0)
'******************************************************************************
Private Function ChangeLE4m(ByVal EntréePrécédente As Boolean, _
  ByVal EntréeNulle As Boolean) As Boolean
  'Cas de l'entrée nulle
  If EntréeNulle Then
    chkEntréeEvasée.Value = False
    chkEntréeEvasée.Enabled = False
    chkTAD.Value = False
    chkTAD.Enabled = False
    txtLI = ""
    txtLI.BackColor = vbGrayText
    txtLI.Enabled = False
    'Réaffecte les données concernées
    With BrancheActive
      .LI = 0#
      .TAD = False
      .EntréeEvasée = False
    End With
  Else
    'Si le rayon R est strictement positif on autorise le TAD
    chkTAD.Enabled = (gbProjetActif.R > 0)
    If EntréePrécédente Then
      'On autorise à nouveau l'entrée évasée et le TAD si le rayon du giratoire n'est pas nul
      chkEntréeEvasée.Enabled = True
      'On réaffecte la largeur d'ilot si la sortie n'est pas nulle
      If Not BrancheActive.SortieNulle Then
        txtLI.Enabled = True
        txtLI.BackColor = vbWhite
        txtLI = DEFAUT_LI 'valeur par défaut LI
      End If
    End If
  End If
  
End Function

'******************************************************************************
' ChangeLS
  'Si la largeur de sortie est nulle, il faut bloquer la saisie
  'de la largeur d'ilot et lui imposer une valeur nulle
'******************************************************************************
Public Sub ChangeLS(ByVal SortiePrécédente, ByVal SortieNulle As Boolean)
  If SortieNulle Then
    txtLI = ""
    txtLI.BackColor = vbGrayText
    txtLI.Enabled = False
    BrancheActive.LI = 0
  ElseIf SortiePrécédente Then
    'On réactive la valeur de LI
    txtLI.Enabled = True
    txtLI.BackColor = vbWhite
    txtLI = DEFAUT_LI 'valeur par défaut de LI
  End If
End Sub

Private Function RetourneEntier(Booléen As Boolean) As Integer
  If Booléen Then
    RetourneEntier = 1
  Else
    RetourneEntier = 0
  End If
End Function

Private Sub txtNumBranche_GotFocus()
  txtNumBranche.SelLength = 1
End Sub

Private Sub txtNumBranche_KeyPress(KeyAscii As Integer)
  If KeyAscii < 49 Or KeyAscii > 48 + NbBranches Then Beep: KeyAscii = 0
End Sub
