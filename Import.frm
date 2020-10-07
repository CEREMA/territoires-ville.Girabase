VERSION 5.00
Begin VB.Form frmImport 
   Caption         =   "Import de trafics - "
   ClientHeight    =   2490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7740
   Icon            =   "Import.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2490
   ScaleWidth      =   7740
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Aide"
      Height          =   375
      Left            =   4680
      TabIndex        =   4
      Top             =   1680
      Width           =   1215
   End
   Begin VB.ComboBox cboPériode 
      Height          =   315
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   720
      Width           =   4815
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Annuler"
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label lblListe 
      Caption         =   "Liste des périodes"
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   840
      Width           =   1455
   End
End
Attribute VB_Name = "frmImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
'*
'*          Projet GIRABASE - CERTU - CETE de l'Ouest
'*
'*          Module de feuille : IMPORT.FRM - frmImport
'*
'*          Feuille d'import de matrices de trafic
'*
'******************************************************************************
Option Explicit

' Ces variables portent le même nom que dans frmDonnées pour généraliser les fonctions appelées dans Trafics et GIRATOIRE
Public GiratoireProjet As GIRATOIRE
Public FichierModifié As Boolean
Public DonnéeModifiée As Boolean

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdHelp_Click()
    SendKeys "{F1}", True
End Sub

Private Sub cmdOK_Click()
Dim I As Integer

  If gbProjetActif.newPériode(cboPériode) Then
    I = cboPériode.ListIndex + 1
    GiratoireProjet.colTrafics.Item(I).Importer
    GiratoireProjet.colTrafics.Remove GiratoireProjet, I
    If cboPériode.ListCount = 0 Then
      Unload Me
    Else
      cboPériode.ListIndex = 0
    End If
  End If
End Sub

Private Sub Form_Load()

  Icon = MDIGirabase.Icon

  ' Aide contextuelle
  HelpContextID = IDhlp_ImportMatrice

End Sub
