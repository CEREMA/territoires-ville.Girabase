VERSION 5.00
Begin VB.Form frmKey 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Licence register"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6165
   Icon            =   "frmKey.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleMode       =   0  'User
   ScaleWidth      =   6165
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   1320
      TabIndex        =   3
      Top             =   3360
      Width           =   1332
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2880
      TabIndex        =   4
      Top             =   3360
      Width           =   1332
   End
   Begin VB.TextBox TxtSerial 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Top             =   2640
      Width           =   2895
   End
   Begin VB.TextBox TxtLicence 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      TabIndex        =   0
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Image imgLogo 
      Height          =   1215
      Index           =   0
      Left            =   4560
      Picture         =   "frmKey.frx":0442
      Stretch         =   -1  'True
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label LblTitre 
      Caption         =   "Please, register your licence"
      Height          =   735
      Left            =   360
      TabIndex        =   6
      Top             =   600
      Width           =   3495
   End
   Begin VB.Label LblSerial 
      Alignment       =   1  'Right Justify
      Caption         =   "Serial :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Left            =   360
      TabIndex        =   5
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label LblLicence 
      Alignment       =   1  'Right Justify
      Caption         =   "License :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Left            =   360
      TabIndex        =   2
      Top             =   2040
      Width           =   1455
   End
End
Attribute VB_Name = "frmKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    'intialisation
    Me.Caption = Titre
    Me.LblTitre.Caption = Msg
    Me.LblLicence.Caption = LBLICENCE
    Me.LblSerial.Caption = LBSERIAL
    Me.cmdOK.Caption = BtnOK
    Me.cmdCancel.Caption = btnCancel
    
    'Modification de l'apparence et du contenu de lblserial selon le type de protection
    If TYPPROTECTION = CPM Then
        Me.TxtSerial.Text = LireTxt(NomFichierSerial)
        Me.TxtSerial.Visible = False
        Me.LblSerial.Visible = False
    Else
        Me.TxtSerial.Visible = True
        Me.LblSerial.Visible = True
    End If

End Sub

'l'utilisateur clique sur annuler
Private Sub cmdCancel_Click()
    Unload Me
End Sub

'l'utilisateur clique sur OK
Private Sub cmdOK_Click()

    'appel de la méthode
    If VerifLicence("rien", "rien", TxtLicence.Text, TxtSerial.Text) Then
        MsgBox MSGPWDVALID
    Else
        MsgBox MSGPWDINVALID 's'affiche aussi si la licence a expiré
    End If
End Sub

'gestion de l'activation du bouton OK
Function ActivercmdOK() As Boolean
    If TYPPROTECTION = CPM Then
        If Trim(Me.TxtLicence.Text) <> "" Then
            Me.cmdOK.Enabled = True
        Else
            Me.cmdOK.Enabled = False
        End If
    Else
        If Trim(Me.TxtLicence.Text) <> "" And Trim(Me.TxtSerial) <> "" Then
            Me.cmdOK.Enabled = True
        Else
            Me.cmdOK.Enabled = False
        End If
    End If
End Function

Private Sub TxtLicence_Change()
    ActivercmdOK
End Sub

Private Sub TxtSerial_Change()
    ActivercmdOK
End Sub
