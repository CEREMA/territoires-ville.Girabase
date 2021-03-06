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
'*          D�finition de la collection Trafics d'un projet
'*
'******************************************************************************

Option Explicit


Private collectTrafics As New Collection

'******************************************************************************
' Ajout d'un objet TRAFIC
'******************************************************************************
Public Function add(ByVal wGiratoire As GIRATOIRE, ByVal nom As String, Optional ByVal BrancheSatur�e As Integer) As TRAFIC
Dim traficNew As New TRAFIC
Dim NbBranches As Integer
Dim Position As Integer

  With wGiratoire
    NbBranches = .NbBranches
    If BrancheSatur�e <> 0 Then
      Position = count
    Else    ' On ins�re la nouvelle p�riode r�eelle avant les p�riodes satur�es fictives
      Position = .nbP�riodes
      .Donn�es.cboP�riode.AddItem nom
      .Donn�es.FichierModifi� = True
    End If
    
    If Not .R�sultats Is Nothing Then
      With .R�sultats
        .cboP�riode.AddItem nom, Position
      End With
    End If
  End With
    
  With traficNew
    .nom = nom
    .V�hiculeCourant = UVP
    .modeUVP = True
    .Dimensionner NbBranches
    If BrancheSatur�e <> 0 Then
      gbProjetActif.nbP�riodesSatur�es = gbProjetActif.nbP�riodesSatur�es + 1
      .BrancheSatur�e = BrancheSatur�e
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
    If Item(Index).BrancheSatur�e <> 0 Then
      Item(Index).TraficP�re.SupprimerFils Item(Index).BrancheSatur�e
      gbProjetActif.nbP�riodesSatur�es = gbProjetActif.nbP�riodesSatur�es - 1
      .R�sultats.cboP�riode.RemoveItem Index - 1
    Else
      Item(Index).SupprimerFils
      .Donn�es.cboP�riode.RemoveItem Index - 1
      .Donn�es.FichierModifi� = True
      If Not .R�sultats Is Nothing Then
        .R�sultats.cboP�riode.RemoveItem Index - 1
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
' V�rifie que le nom de la p�riode � cr�er ou � renommer n'existe pas d�j�
'******************************************************************************
Public Function EnDouble(ByVal nomP�riode As String) As Boolean
Dim i As Integer

' Modif AV - 15.06.99 - La Combo ne sait pas diff�rencier la casse --> on n'autorise pas 2 p�riodes de m�me nom � la casse pr�s
  For i = 1 To count() ' Recherche d'un doublon
    If UCase(Item(i).nom) = UCase(nomP�riode) Then
      MsgBox IDm_DoublonP�riode, vbQuestion
      EnDouble = True
      Exit Function
    End If
  Next
  
End Function

'****************************************************************************************
' Recherche du premier num�ro disponible � proposer par d�faut pour une nouvelle p�riode
'****************************************************************************************
Public Function defautNew() As String
Dim i As Integer
Dim j As Integer
Dim nomP�riode As String
Dim Trouv� As Boolean
Dim Libre As Boolean

  j = 1
  Do
    nomP�riode = IDl_P�riode & CStr(j)
    For i = 1 To count()
      If Item(i).nom = nomP�riode Then
        Trouv� = True
        Exit For
      End If
    Next
    Libre = Not Trouv�
    Trouv� = False
    j = j + 1
  Loop Until Libre
  
  defautNew = nomP�riode
  
End Function

'****************************************************************************************
' D�termine si au moins une p�riode de trafic est compl�te pour effectuer le calcul
'****************************************************************************************
Public Function Uncomplet()
Dim wTrafic As TRAFIC
  For Each wTrafic In collectTrafics
    If wTrafic.EstCompl�te Then
      Uncomplet = True
      Exit Function
    End If
  Next
  MsgBox IDm_IncompletP�riode
End Function
