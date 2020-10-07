VERSION 5.00
Begin VB.Form frmTrafMult 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4005
   ClientLeft      =   2085
   ClientTop       =   2700
   ClientWidth     =   6645
   Icon            =   "TrafMult.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4005
   ScaleWidth      =   6645
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picBoutons 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   6645
      TabIndex        =   13
      Top             =   3030
      Width           =   6645
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   375
         Left            =   600
         TabIndex        =   7
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Annuler"
         Height          =   375
         Left            =   2520
         TabIndex        =   8
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "Aide"
         Height          =   375
         Left            =   4560
         TabIndex        =   9
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame fraBranches 
      Height          =   975
      Left            =   2640
      TabIndex        =   4
      Top             =   720
      Visible         =   0   'False
      Width           =   3735
      Begin VB.TextBox txtCoefBranche 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   2760
         TabIndex        =   6
         Text            =   "1"
         Top             =   360
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CheckBox chkBranche 
         Caption         =   "Nom de la branche"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Visible         =   0   'False
         Width           =   2535
      End
   End
   Begin VB.TextBox txtCoefGen 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3840
      TabIndex        =   11
      Text            =   "1"
      Top             =   720
      Width           =   735
   End
   Begin VB.Frame fraTrafics 
      Caption         =   "Appliquer aux trafics  ... "
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   2415
      Begin VB.OptionButton optTrafic 
         Caption         =   "Sortants"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   1815
      End
      Begin VB.OptionButton optTrafic 
         Caption         =   "Entrants"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1935
      End
      Begin VB.OptionButton optTrafic 
         Caption         =   "Tous"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   1935
      End
   End
   Begin VB.Label lblPériode 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   240
      TabIndex        =   12
      Top             =   240
      Width           =   45
   End
   Begin VB.Label lblCoefGen 
      Caption         =   "Coefficient : "
      Height          =   255
      Left            =   2760
      TabIndex        =   10
      Top             =   720
      Width           =   975
   End
End
Attribute VB_Name = "frmTrafMult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
'*
'*          Projet GIRABASE - CERTU - CETE de l'Ouest
'*
'*          Module de feuille : TRAFMULT.FRM - frmTrafMult
'*
'*          Feuille de multiplication de matrice de trafic
'*
'******************************************************************************

Option Explicit

'*** CONSTANTES DE CHAINE Susceptibles d'être traduites *****
' Multiplication de matrice
Const IDl_CoefEntrée = "Coefficients en entrée"
Const IDl_CoefSortie = "Coefficients en sortie"
Const IDm_BorneCoefMult = "Le coefficient doit être <= "
Const IDm_TropGrand = "Valeurs de trafic trop grandes"
'*** DRAPEAU : Fin des CONSTANTES DE CHAINE Susceptibles d'être traduites *****

Const TOTAL = 0
Const ENTRANT = 1
Const SORTANT = 2
Const COEFMAX = 20

'Dim UneSaisie As Integer
Private sauvCoefBranche() As Single
Private sauvCoefGen As Single
Private FlagPtDecimal As Boolean

' Période de trafic à multiplier
Public TraficOrigine As TRAFIC
Private wBranches As Branches

' Utilisation de la feuille à seule fin de calculer une matrice par saturation d'une branche (alimentées par frmRésultats)
Public Saturation As Boolean
Public NumBranche As Integer


Private Sub cmdHelp_Click()
    
    SendKeys "{F1}", True

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
' Gestion du point décimal comme virgule
' Si l'utilisateur est ainsi configuré, on détecte la frappe du point décimal
' mais seule la fonction KeyPress semble en mesure de réafficher la virgule

If KeyCode = vbKeyDecimal And Shift = 0 Then
  FlagPtDecimal = True
End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

  If FlagPtDecimal Then KeyAscii = gbPtDecimal: FlagPtDecimal = False

End Sub

'******************************************************************************
' Abandon
'******************************************************************************
Private Sub cmdCancel_Click()
  With gbProjetActif.colTrafics
    .Remove gbProjetActif, gbProjetActif.colTrafics.count
  End With
  Unload Me
End Sub

'******************************************************************************
' OK : Exécution de la multiplication
'******************************************************************************
Private Sub cmdOK_Click()
Dim I As Integer

  If Numopt(optTrafic) <> 0 Then
  'On remet à 1 les éventuels coefficients modifiés alors que la case correspondante n'est pas cochée
    For I = 1 To gbProjetActif.NbBranches
      If chkBranche(I) = vbUnchecked Then txtCoefBranche(I) = 1
    Next
  End If

On Error GoTo GestErr
  TraficOrigine.Multiplier
  Unload Me
  
  Exit Sub
  
GestErr:
  If Err = 6 Then
    MsgBox IDm_TropGrand
    gbProjetActif.colTrafics.Remove gbProjetActif, gbProjetActif.nbPériodes
    Resume Next
  Else
    ErreurFatale
  End If
    
End Sub

'******************************************************************************
' Chargement de la feuille
'******************************************************************************
Private Sub Form_Load()
Dim I As Integer
Dim NbBranches As Integer

'Icon = MDIGirabase.Icon

With gbProjetActif
  NbBranches = .NbBranches
  Set wBranches = .colBranches
End With

If Saturation Then
' Utilisation détournée de la feuille : celle-ci ne s'affiche pas
  optTrafic(ENTRANT) = True
  For I = 1 To NbBranches
    Load chkBranche(I)
    Load txtCoefBranche(I)
  Next
  chkBranche(NumBranche) = vbChecked
  txtCoefBranche(NumBranche) = TraficOrigine.getC(NumBranche) / TraficOrigine.getQE(NumBranche)
  Exit Sub
End If

' Aide contextuelle
HelpContextID = IDhlp_MultPériode

frmTrafMult.Caption = IDl_Multiplication & IDl_DeLaPériode & " " & TraficOrigine.nom
lblPériode.Caption = IDl_Période & " " & gbProjetActif.Données.cboPériode
ReDim sauvCoefBranche(1 To NbBranches)

'création du tableau contenant le nom des branches et les coefficients
fraBranches.Height = fraBranches.Height + 300 * (NbBranches - 1)
For I = 1 To NbBranches
  Load chkBranche(I)
  With chkBranche(I)
    .Top = .Top + 300 * (I - 1)
    .Visible = True
  End With
  Load txtCoefBranche(I)
  With txtCoefBranche(I)
    .Top = .Top + 300 * (I - 1)
    .Visible = True
  End With
  chkBranche(I).Caption = gbProjetActif.colBranches.Item(I).nom
  chkBranche(I) = 1
Next I

Me.Height = fraBranches.Top + fraBranches.Height + picBoutons.Height + 400

'initialisation des branches
optTrafic(0).Value = True

End Sub

Private Sub chkBranche_Click(Index As Integer)
  txtCoefBranche(Index).Enabled = chkBranche(Index)
End Sub

'******************************************************************************
' Option : Trafic entrant, sortant ou ensemble
'******************************************************************************
Private Sub optTrafic_Click(Index As Integer)
  Dim Ensemble As Boolean
  Dim I As Integer
  
  If Saturation Then Exit Sub
  
  Ensemble = (Index = TOTAL)
  lblCoefGen.Visible = Ensemble
  txtCoefGen.Visible = Ensemble
  fraBranches.Visible = Not Ensemble
  
  If Not Ensemble Then
    If Index = ENTRANT Then
      fraBranches = IDl_CoefEntrée
      For I = 1 To wBranches.count
        If wBranches.Item(I).EntréeNulle Then
          chkBranche(I).Enabled = False
          txtCoefBranche(I).Enabled = False
        Else
          chkBranche(I).Enabled = True
          txtCoefBranche(I).Enabled = True
        End If
      Next
    Else
      fraBranches = IDl_CoefSortie
      For I = 1 To wBranches.count
        If wBranches.Item(I).SortieNulle Then
          chkBranche(I).Enabled = False
          txtCoefBranche(I).Enabled = False
        Else
          chkBranche(I).Enabled = True
          txtCoefBranche(I).Enabled = True
        End If
      Next
    End If
  End If
  
End Sub

Private Sub txtCoefBranche_GotFocus(Index As Integer)
  sauvCoefBranche(Index) = txtCoefBranche(Index)
End Sub

Private Sub txtCoefBranche_KeyPress(Index As Integer, KeyAscii As Integer)
  KeyAscii = ControleRéel(KeyAscii)
End Sub

Private Sub txtCoefBranche_Validate(Index As Integer, Cancel As Boolean)
  If MonCtrlNumeric(txtCoefBranche(Index), Obligatoire:=True, Positif:=True) Then txtCoefBranche(Index) = sauvCoefBranche(Index)
  If txtCoefBranche(Index) > COEFMAX Then MsgBox IDm_BorneCoefMult & CStr(COEFMAX): txtCoefBranche(Index) = sauvCoefBranche(Index)
End Sub

Private Sub txtCoefGen_GotFocus()
  sauvCoefGen = txtCoefGen
End Sub

Private Sub txtCoefGen_KeyPress(KeyAscii As Integer)
  KeyAscii = ControleRéel(KeyAscii)
End Sub

Private Sub txtCoefGen_Validate(Cancel As Boolean)
  If MonCtrlNumeric(txtCoefGen, Obligatoire:=True, Positif:=True) Then txtCoefGen = sauvCoefGen
  If txtCoefGen > COEFMAX Then MsgBox IDm_BorneCoefMult & CStr(COEFMAX): txtCoefGen = sauvCoefGen
End Sub

Public Function coeff(ByVal I As Integer, ByVal j As Integer) As Single
    
  Select Case Numopt(optTrafic)
  Case TOTAL
    coeff = txtCoefGen
  Case ENTRANT
    coeff = txtCoefBranche(I)
  Case SORTANT
    coeff = txtCoefBranche(j)
  End Select
  
End Function

