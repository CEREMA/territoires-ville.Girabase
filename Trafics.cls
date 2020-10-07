VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Trafics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'******************************************************************************
'*
'*          Projet GIRABASE - CERTU - CETE de l'Ouest
'*
'*          Module de classe : TRAFICS.CLS
'*
'*          Définition de la collection Trafics d'un projet
'*
'******************************************************************************

Option Explicit


Private collectTrafics As New Collection

'******************************************************************************
' Ajout d'un objet TRAFIC
'******************************************************************************
Public Function add(ByVal wGiratoire As GIRATOIRE, ByVal nom As String, Optional ByVal BrancheSaturée As Integer) As TRAFIC
Dim traficNew As New TRAFIC
Dim NbBranches As Integer
Dim Position As Integer

  With wGiratoire
    NbBranches = .NbBranches
    If BrancheSaturée <> 0 Then
      Position = count
    Else    ' On insère la nouvelle période réeelle avant les périodes saturées fictives
      Position = .nbPériodes
      .Données.cboPériode.AddItem nom
      .Données.FichierModifié = True
    End If
    
    If Not .Résultats Is Nothing Then
      With .Résultats
        .cboPériode.AddItem nom, Position
      End With
    End If
  End With
    
  With traficNew
    .nom = nom
    .VéhiculeCourant = UVP
    .modeUVP = True
    .Dimensionner NbBranches
    If BrancheSaturée <> 0 Then
      gbProjetActif.nbPériodesSaturées = gbProjetActif.nbPériodesSaturées + 1
      .BrancheSaturée = BrancheSaturée
    End If
  End With
  
  If Position = 0 Then
    collectTrafics.add traficNew
  Else
    collectTrafics.add traficNew, after:=Position
  End If
  
  Set add = traficNew
  
End Function

Public Function count() As Long
  count = collectTrafics.count
End Function

Public Function Item(ByVal Index As Variant) As TRAFIC
  Set Item = collectTrafics(Index)
End Function

Public Sub Remove(wGiratoire As GIRATOIRE, ByVal Index As Variant)
Dim i As Integer

  With wGiratoire
    If Item(Index).BrancheSaturée <> 0 Then
      Item(Index).TraficPère.SupprimerFils Item(Index).BrancheSaturée
      gbProjetActif.nbPériodesSaturées = gbProjetActif.nbPériodesSaturées - 1
      .Résultats.cboPériode.RemoveItem Index - 1
    Else
      Item(Index).SupprimerFils
      .Données.cboPériode.RemoveItem Index - 1
      .Données.FichierModifié = True
      If Not .Résultats Is Nothing Then
        .Résultats.cboPériode.RemoveItem Index - 1
      End If
    End If
  End With

  collectTrafics.Remove Index
End Sub

Public Function RechIndex(ByVal nom As String) As Integer
Dim i As Integer

  For i = 1 To count
    If Item(i).nom = nom Then
      RechIndex = i
      Exit Function
    End If
  Next
End Function

'******************************************************************************
' Vérifie que le nom de la période à créer ou à renommer n'existe pas déjà
'******************************************************************************
Public Function EnDouble(ByVal nomPériode As String) As Boolean
Dim i As Integer

' Modif AV - 15.06.99 - La Combo ne sait pas différencier la casse --> on n'autorise pas 2 périodes de même nom à la casse près
  For i = 1 To count() ' Recherche d'un doublon
    If UCase(Item(i).nom) = UCase(nomPériode) Then
      MsgBox IDm_DoublonPériode, vbQuestion
      EnDouble = True
      Exit Function
    End If
  Next
  
End Function

'****************************************************************************************
' Recherche du premier numéro disponible à proposer par défaut pour une nouvelle période
'****************************************************************************************
Public Function defautNew() As String
Dim i As Integer
Dim j As Integer
Dim nomPériode As String
Dim Trouvé As Boolean
Dim Libre As Boolean

  j = 1
  Do
    nomPériode = IDl_Période & CStr(j)
    For i = 1 To count()
      If Item(i).nom = nomPériode Then
        Trouvé = True
        Exit For
      End If
    Next
    Libre = Not Trouvé
    Trouvé = False
    j = j + 1
  Loop Until Libre
  
  defautNew = nomPériode
  
End Function

'****************************************************************************************
' Détermine si au moins une période de trafic est complète pour effectuer le calcul
'****************************************************************************************
Public Function Uncomplet()
Dim wTrafic As TRAFIC
  For Each wTrafic In collectTrafics
    If wTrafic.EstComplète Then
      Uncomplet = True
      Exit Function
    End If
  Next
  MsgBox IDm_IncompletPériode
End Function
