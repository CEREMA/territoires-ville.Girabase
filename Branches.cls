VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Branches"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'******************************************************************************
'*
'*          Projet GIRABASE - CERTU - CETE de l'Ouest
'*
'*          Module de classe : BRANCHES.CLS
'*
'*          Définition de la collection de branches d'un projet
'*
'******************************************************************************

Option Explicit


Private collectBranches As New Collection

'******************************************************************************
' Ajout d'un objet BRANCHE
'******************************************************************************
Public Function add(ByVal NumBranche As Integer, ByVal Ecart As Single) As BRANCHE

  Dim brancheNew As New BRANCHE
  Dim i As Integer
  
  With brancheNew
    ' Initialisation avec Valeurs par défaut : CAHIER DES CHARGES §2.2 et 2.3
    .nom = IDl_Branche & CStr(NumBranche)
    .Angle = (NumBranche - 1) * Ecart
    .Rampe = False
    .TAD = False
    .LE15m = 3.5
    .LE4m = 3.5
    .LI = 3
    .LS = 4
    .EntréeEvasée = False
    .Ecart = Ecart
  End With

  collectBranches.add brancheNew
  Set add = brancheNew

End Function

Public Function count() As Long
  count = collectBranches.count
End Function

Public Function Item(ByVal Index As Variant) As BRANCHE
  Set Item = collectBranches(Index)
End Function

Public Function NewEnum() As IUnknown
  Set NewEnum = collectBranches.[_NewEnum]
End Function

Public Sub Remove(ByVal Index As Variant)
  collectBranches.Remove Index
End Sub

'******************************************************************************
' Vérifie que le nom de la branche à renommer n'existe pas déjà
'******************************************************************************
Public Function EnDouble(ByVal nomBranche As String) As Boolean
Dim i As Integer

  For i = 1 To count() ' Recherche d'un doublon
    If Item(i).nom = nomBranche Then
      MsgBox IDm_DoublonBranche, vbQuestion
      EnDouble = True
      Exit Function
    End If
  Next
  
End Function
'******************************************************************************
' Affichage des données branches
'******************************************************************************
Public Sub Affiche(ByVal FeuilleDonnées As Form)
Dim i As Integer

  For i = 1 To count()
    collectBranches(i).Affiche FeuilleDonnées, i
  Next

End Sub

'******************************************************************************
' Verrouillage des données trafic
'******************************************************************************
' Il s'agit d'interdire l'accès à une cellule de la matrice de trafic
' si la largeur d'entrée ou la largeur de sortie correpondante est égale à 0 (trafic forcément nul)

Public Sub BlocageTrafic(ByVal FeuilleDonnées As Form)
Dim i As Integer
Dim j As Integer

  With FeuilleDonnées
  
    With .vgdTrafic(VEHICULE)
      For i = 1 To count()
        .Row = i
        For j = 1 To count()
          .Col = j
'          If Item(i).LE4m = 0 Or Item(j).LS = 0 Then
          If Item(i).EntréeNulle Or Item(j).SortieNulle Then
            .Lock = True
            .BackColor = vbGrayText   'vbInactiveCaptionText
          Else
            .Lock = False
            .BackColor = vbWhite
          End If
        Next
      Next
    End With
    
    For i = 1 To count()
'      If Item(i).LE4m = 0 Then
      If Item(i).EntréeNulle Then
        .txtQE(i).BackColor = vbInactiveCaptionText
      Else
        .txtQE(i).BackColor = vbWhite
      End If
'      If Item(i).LS = 0 Then
      If Item(i).SortieNulle Then
        .txtQS(i).BackColor = vbGrayText   'vbInactiveCaptionText
      Else
        .txtQS(i).BackColor = vbWhite
      End If
    Next
    
  End With

End Sub
