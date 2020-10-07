VERSION 5.00
Begin VB.Form frmParam 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Nouveau Giratoire"
   ClientHeight    =   2460
   ClientLeft      =   1920
   ClientTop       =   2880
   ClientWidth     =   3570
   Icon            =   "ParamGir.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2460
   ScaleWidth      =   3570
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Aide"
      Height          =   375
      Left            =   2520
      TabIndex        =   7
      Top             =   1920
      Width           =   855
   End
   Begin VB.Frame fraUnité 
      Caption         =   "Unités d'angle"
      Height          =   615
      Left            =   600
      TabIndex        =   6
      Top             =   840
      Width           =   2295
      Begin VB.OptionButton optUnité 
         Caption         =   "grades"
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optUnité 
         Caption         =   "degrés"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Annuler"
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox txtNbBranches 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2280
      MaxLength       =   1
      TabIndex        =   0
      Top             =   360
      Width           =   255
   End
   Begin VB.Label lblNbBranches 
      AutoSize        =   -1  'True
      Caption         =   "Nombre de branches :"
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   480
      Width           =   1575
   End
End
Attribute VB_Name = "frmParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
'*
'*          Projet GIRABASE - CERTU - CETE de l'Ouest
'*
'*          Module de feuille : PARAMGIR.FRM - frmParam
'*
'*          Feuille de création d'un nouveau giratoire
'*
'******************************************************************************

Option Explicit

'*** CONSTANTES DE CHAINE Susceptibles d'être traduites *****
Const IDm_NumericEntier = "Numérique entier obligatoirement"
Const IDm_BorneNbBranches = "Nombre de branches compris entre 3 et 8"
'*** DRAPEAU : Fin des CONSTANTES DE CHAINE Susceptibles d'être traduites *****

'******************************************************************************
' Abandon
'******************************************************************************
Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdHelp_Click()

  SendKeys "{F1}", True

End Sub

'******************************************************************************
' OK : si nbbranches ok , Création de l'objet Giratoire nouveau
'******************************************************************************
Private Sub cmdOK_Click()
  If controlNbBranches Then
    txtNbBranches.SetFocus
  Else
    ' gbCreFille décharge frmParam
    gbCreFille ""
  End If

End Sub

'******************************************************************************
' Chargement de la feuille
'******************************************************************************
Private Sub Form_Load()

'  Icon = MDIGirabase.Icon

  ' Aide contextuelle
  HelpContextID = IDhlp_Nouveau

  txtNbBranches = DEFAUTNBBRANCHES
  optUnité(DEGRE) = True
End Sub

'******************************************************************************
' controlNbBranches : controle de saisie du champ txtNbBranches
'******************************************************************************
Private Function controlNbBranches() As Boolean

  If txtNbBranches = "" Then
    MsgBox IDm_Obligatoire
    controlNbBranches = True
  ElseIf Not IsNumeric(txtNbBranches) Then
    MsgBox IDm_NumericEntier
    controlNbBranches = True
  ElseIf CInt(txtNbBranches) > 8 Or CInt(txtNbBranches) < 3 Then
    MsgBox (IDm_BorneNbBranches)
    controlNbBranches = True
  Else
    controlNbBranches = False
  End If
  
End Function

Private Sub txtNbBranches_Change()
  If txtNbBranches = "" Then Exit Sub
  txtNbBranches.SelLength = 1
  txtNbBranches.SelStart = 0
End Sub

Private Sub txtNbBranches_GotFocus()
    txtNbBranches.SelLength = 1
End Sub

Private Sub txtNbBranches_KeyPress(KeyAscii As Integer)
  If KeyAscii < 48 Or KeyAscii > 57 Then Beep: KeyAscii = 0
End Sub
