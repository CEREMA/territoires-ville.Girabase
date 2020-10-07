VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Giratoires"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'******************************************************************************
'*
'*          Projet GIRABASE - CERTU - CETE de l'Ouest
'*
'*          Module de classe : GIRATOIRES.BAS
'*
'*          Définition de la collection Giratoires
'*          (peut-être superflu : correspond à la collection de feuilles filles MDI)
'*
'******************************************************************************

Option Explicit

Private colGiratoires As New Collection

'******************************************************************************
' Ajout d'un objet GIRATOIRE
'******************************************************************************
Public Function add(ByVal NbBranches As Integer, ByVal modeangle As Integer) As GIRATOIRE
Dim giratoireNew As New GIRATOIRE
  
  Initialiser giratoireNew, NbBranches, modeangle
  colGiratoires.add giratoireNew
  Set add = giratoireNew

End Function

Public Function count() As Long
  count = colGiratoires.count
End Function

Public Function Item(ByVal Index As Variant) As GIRATOIRE
  Set Item = colGiratoires(Index)
End Function

Public Sub Remove(ByVal Index As Variant)
  colGiratoires.Remove Index
End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = colGiratoires.[_NewEnum]
End Function
'******************************************************************************
' Initilisation d'un objet GIRATOIRE avec des valeurs par défaut
'******************************************************************************
Public Sub Initialiser(ByRef wGiratoire As GIRATOIRE, ByVal NbBranches As Integer, ByVal modeangle As Integer)

  With wGiratoire
  ' Debug (AV - 25.02.99) : cette valeur ne doit pas être réinitialisée
    .modeangle = modeangle
    .NbBranches = NbBranches
    ' Initialisation avec Valeurs par défaut : CAHIER DES CHARGES §2.2 et 2.3
    If NbBranches > 5 Then
      .R = 9
    Else
      .R = 6
    End If
    .DateModif = Date
    .LA = 7
    '.Milieu = PU ' non spécifié au CC
    .Milieu = -1  'Aucun milieu spécifié
    .Bf = 2
    
    ' Création des branches
    Dim i As Integer
    Dim Ecart As Single

    Ecart = 2 * eqvPI(.modeangle) / NbBranches

    For i = 1 To NbBranches
      .colBranches.add i, Ecart
    Next
    
  End With
    
  'Set Reinit = wGiratoire

End Sub
