VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmOptCourbes 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Capacity curve options"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin MSComDlg.CommonDialog dlgCouleur 
      Left            =   360
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.PictureBox picCouleur 
      Height          =   255
      Index           =   0
      Left            =   4080
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblNom 
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label lblAbrégé 
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "frmOptCourbes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
'*
'*          Projet GIRABASE - CERTU - CETE de l'Ouest
'*
'*          Module de feuille : OPTCOURBES.FRM.FRM - frmOptCourbes
'*
'*          Feuille de choix d'options pour les courbes de capacité
'*
'*          Une feuille frmOptCourbes est associée à chaque feuille FrmRésultats
'*          La propriété FeuilleLégende de FrmRésultats désigne cette feuille
'*
'******************************************************************************

Option Explicit

'******************************************************************************
' Chargement de la feuille
'******************************************************************************
Private Sub Form_Load()
Dim I As Integer
Dim hLigne As Integer
  
  With gbProjetActif.Résultats
    
    Move .picRésultat.Left + .fraRésultats.Left, .picRésultat.Top + .vgdRecap.Top + 100
  End With
  
  hLigne = lblAbrégé(0).Height * 1.5

  For I = 1 To gbProjetActif.colTrafics.count
    ChargerPériode I
  Next
  
  Height = lblAbrégé(I - 1).Top + 2 * hLigne

  HelpContextID = IDhlp_Courbes
  
End Sub

'******************************************************************************
' Chargement d'une ligne correspondant à une période
'******************************************************************************
Public Sub ChargerPériode(ByVal I As Integer, Optional ByVal Unique As Boolean)
Dim hLigne As Integer
Dim wTrafic As TRAFIC

    hLigne = lblAbrégé(0).Height * 1.5
    
  ' Charger les controles
    Load lblAbrégé(I)
    Load lblNom(I)
    Load picCouleur(I)
  ' Positionner les controles
    lblAbrégé(I).Top = lblAbrégé(I).Top + hLigne * (I - 1)
    lblNom(I).Top = lblAbrégé(I).Top
    picCouleur(I).Top = lblAbrégé(I).Top
  ' Alimenter les controles
    lblAbrégé(I) = "P" & CStr(I)
    Set wTrafic = gbProjetActif.colTrafics.Item(I)
    lblNom(I) = wTrafic.nom
    picCouleur(I).BackColor = wTrafic.CouleurCourbe
    If Not wTrafic.EstComplète Then
      picCouleur(I).Enabled = False
      lblNom(I).Enabled = False
      lblAbrégé(I).Enabled = False
    End If
  ' Afficher les controles
    lblAbrégé(I).Visible = True
    lblNom(I).Visible = True
    picCouleur(I).Visible = True
    
    ' Appel depuis l'extérieur pour rajouter une seule ligne
    If Unique Then
      Height = Height + lblAbrégé(0).Height * 1.5
    End If

End Sub

'******************************************************************************
' Chargement d'une ligne correspondant à une période
'******************************************************************************
Public Sub DéchargerPériode(ByVal numPériode As Integer)

Dim hLigne As Integer
Dim Nb As Integer
Dim I As Integer

  Nb = lblNom.UBound
  For I = Nb To numPériode + 1 Step -1
    lblNom(I - 1) = lblNom(I)
    picCouleur(I - 1).BackColor = picCouleur(I).BackColor
  Next
  Unload lblAbrégé(Nb)
  Unload lblNom(Nb)
  Unload picCouleur(Nb)
  
  Height = Height - lblAbrégé(0).Height * 1.5

End Sub

'******************************************************************************
' Choix d'une couleur dans la boite de dialogue standard
'******************************************************************************
Private Sub picCouleur_Click(Index As Integer)

  With dlgCouleur
    .flags = cdlCCRGBInit Or cdlCCPreventFullOpen
    .Color = picCouleur(Index).BackColor
    On Error GoTo TraitementErreur
    .ShowColor
    picCouleur(Index).BackColor = .Color
    gbProjetActif.colTrafics.Item(Index).CouleurCourbe = .Color
    gbProjetActif.Résultats.CourbeCapacité
  End With
  Exit Sub
  
TraitementErreur:
  If Err <> cdlCancel Then ErreurFatale
  Exit Sub
End Sub
