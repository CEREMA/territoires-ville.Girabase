VERSION 5.00
Begin VB.Form frmApropos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "A propos de "
   ClientHeight    =   3360
   ClientLeft      =   1815
   ClientTop       =   2505
   ClientWidth     =   4785
   LinkTopic       =   "frmApropos"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3360
   ScaleWidth      =   4785
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   3360
      TabIndex        =   1
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label lblMaitrise 
      Alignment       =   2  'Center
      Caption         =   "En collaboration"
      Height          =   255
      Index           =   2
      Left            =   3360
      TabIndex        =   4
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label lblMaitrise 
      Alignment       =   2  'Center
      Caption         =   "Maitrise d'oeuvre"
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   3
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label lblMaitrise 
      Alignment       =   2  'Center
      Caption         =   "Maitrise d'ouvrage"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
   Begin VB.Line linSeparation 
      BorderWidth     =   3
      X1              =   360
      X2              =   4440
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Image imgLogo 
      Height          =   975
      Index           =   2
      Left            =   3480
      Picture         =   "APropos.frx":0000
      Stretch         =   -1  'True
      Top             =   960
      Width           =   975
   End
   Begin VB.Image imgLogo 
      Height          =   975
      Index           =   0
      Left            =   360
      Picture         =   "APropos.frx":2BEF
      Stretch         =   -1  'True
      Top             =   960
      Width           =   975
   End
   Begin VB.Image imgLogo 
      Height          =   975
      Index           =   1
      Left            =   1920
      Picture         =   "APropos.frx":42801
      Stretch         =   -1  'True
      Top             =   960
      Width           =   975
   End
   Begin VB.Label lblVersion 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Girabase Version "
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   2520
      Width           =   2655
   End
End
Attribute VB_Name = "frmApropos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************************************
'     GIRATION v3 - CERTU/CETE de l'Ouest
'         Septembre 97

'   Réalisation : André VIGNAUD

'   Module de feuille : frmApropos    -   Fichier APROPOS.FRM

'**************************************************************************************
Option Explicit

Const IDl_LicenceNumero = "Licence n° "

Private Sub cmdOK_Click()
    
    Unload Me

End Sub

Private Sub Form_Load()

    Me.Icon = MDIGirabase.Icon
    'Affichage centré de la fenêtre
    Me.ScaleMode = vbTwips
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    Caption = Caption & App.Title
    
    lblVersion.Caption = App.Title & " " & IDl_Version & " " & App.Major & "." & App.Minor & "." & App.Revision
    lblVersion.Caption = lblVersion.Caption & vbCrLf & LBLICENCE & NumeroLicence

End Sub
