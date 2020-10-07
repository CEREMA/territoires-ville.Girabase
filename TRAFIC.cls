VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TRAFIC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'******************************************************************************
'*
'*          Projet GIRABASE - CERTU - CETE de l'Ouest
'*
'*          Module de classe : TRAFIC.BAS
'*
'*          Définition de l'objet TRAFIC
'*
'******************************************************************************

Option Explicit

' Un objet TRAFIC correspond à une matrice de trafics pour une période donnée
' Les tableaux sont dimensionnés à la création des périodes

' Note : une valeur de -1 indique une valeur non saisie - doit être évaluée à 0 dans les calculs
' Dans les fichiers, on inscrit la chaine correspondante avec le mot-clé "VIDE"
Private Const VIDE = -1
Private Const ID_Vide = "VIDE"

Private Const LGTRAFIC = 4      ' nombre maxi de chiffres d'un trafic

Private Const INVMATRICE = 0
Private Const MULTMATRICE = 1
Private Const IMPMATRICE = 2

Public nom As String
Private QP() As Integer            ' Trafic piétons par branche
Public modeUVP As Boolean          ' Indicateur : saisie en UVP ou en PL/VL/2 Roues
Private QVéhicule() As Integer     ' Matrice de trafics (VL, PL, 2R, UVP)

'Tableaux de données pour les calculs de capacité
'Données en entrée
Private QE() As Integer         ' Trafic entrant
Private QS() As Integer         ' Trafic sortant
Private Q() As Integer          ' Trafic UVP corrigé
Private QEntrant() As Integer   ' Trafic entrant corrigé
'Donnée intermédiaire
Private QTournant() As Integer  ' Trafic tournant (utile pour certains messages)
'Donnes en sortie
Private QG() As Integer        ' Trafic gênant
Private C() As Double          ' Capacité définitive
Private Cp() As Double         ' Perte de capacité due aux piétons
Private rc() As Double       ' Réserve de capacité (en uvp/h)
Private TMA() As Double        ' Temps moyen d'attente
Private TTA() As Double        ' Temps total d'attente
Private LK() As Single         ' Longueur moyenne de stockage
Private LKM() As Single        ' Longueur totale de stockage
  
Public VéhiculeCourant As Integer ' VL = 0 - PL = 1 - 2R = 2 - UVP = 3

Public BrancheSaturée As Integer
Private TraficFils() As Boolean
Public TraficPère As TRAFIC
Public CalculFait As Boolean
Public CouleurCourbe As Long

Public Sub InvalideCalcul()
  SupprimerFils
  CalculFait = False
End Sub

'****************************************************************************************************
' Suppression du Trafic Fils obtenu par saturation de la branche NumBranche
'****************************************************************************************************
Public Sub SupprimerFils(Optional ByVal NumBranche)
Dim i As Integer
  
  If IsMissing(NumBranche) Then
    'Suppression de tous les trafics fils
    For i = 1 To gbProjetActif.NbBranches
      With gbProjetActif.colTrafics
        If TraficFils(i) Then .Remove gbProjetActif, .RechIndex(nom & IDl_AbrévSaturBranche & CStr(i))
      End With
    Next
  Else
    ' Le trafic fils est déjà supprimé : positiont de l'indicateur
    TraficFils(NumBranche) = False
  End If
End Sub

Public Function EstSaturéeBranche(ByVal NumBranche As Integer)
  EstSaturéeBranche = TraficFils(NumBranche)
End Function

' (1) Ubound(QP) est utilisé ici fréquemment pour récupérer le nombre de branches du projet

'****************************************************************************************************
' Dimensionnement des tableaux de l'objet TRAFIC
'   Lors de sa création : Méthode add dans Trafics.cls  (mode="")
'   Lors du changement de mode      (mode = "Mode UVP" ou "Mode VL-PL-2R")
'****************************************************************************************************
Public Sub Dimensionner(ByVal NbBranches As Integer, Optional ByVal Mode As String = "")
Dim i As Integer
Dim j As Integer
Dim k As Integer

ReDim Preserve QP(1 To NbBranches)
ReDim Preserve QVéhicule(0 To 3, 1 To NbBranches, 1 To NbBranches)

Select Case Mode
' première Initialisation
Case ""
  ReDim TraficFils(1 To NbBranches)
  ReDim QE(1 To NbBranches)         ' Trafic entrant
  ReDim QS(1 To NbBranches)         ' Trafic sortant
  For i = 1 To NbBranches
    QP(i) = 10
    For j = 1 To NbBranches
      For k = VL To UVP
        QVéhicule(k, j, i) = VIDE
      Next
    Next
  Next

Case IDl_ModeUVP
' Bascule UVP-->PL-VL-2R : on conserve le trafic PIETON - on réinitialise le reste
  For i = 1 To NbBranches
    For j = 1 To NbBranches
      For k = VL To UVP
        QVéhicule(k, j, i) = VIDE
      Next
    Next
  Next
  ReDim QE(1 To NbBranches)         ' Trafic entrant
  ReDim QS(1 To NbBranches)         ' Trafic sortant
Case Else
' Bascule PL-VL-2R-->UVP : on conserve le trafic PIETON et le trafic UVP
  For i = 1 To NbBranches
    For j = 1 To NbBranches
      For k = VL To DEUXR
        QVéhicule(k, j, i) = VIDE
      Next
    Next
  Next
  ReDim Preserve QE(1 To NbBranches)         ' Trafic entrant
  ReDim Preserve QS(1 To NbBranches)         ' Trafic sortant
End Select

If Mode <> "" Then  ' Bascule
  modeUVP = Not modeUVP
  With gbProjetActif.Données
'    modeUVP = .lblTraficUVP.Visible
    If modeUVP Then
      VéhiculeCourant = UVP
    Else
      VéhiculeCourant = Numopt(.optTrafic)
    End If
  End With
End If
 
End Sub

'******************************************************************************
' Lecture d'un trafic Piéton
'******************************************************************************
Private Function getQPiéton() As Variant
  getQPiéton = QP
End Function


'******************************************************************************
' Lecture d'une ligne de trafic véhicule
'******************************************************************************
Private Function getQVéhicule(ByVal numBrancheEntrée As Integer, ByVal véhicule As Integer) As Variant
Dim i%
Dim QV(1 To 8) As Integer
  For i = 1 To gbProjetActif.NbBranches
    QV(i) = QVéhicule(véhicule, numBrancheEntrée, i)
  Next
  getQVéhicule = QV
End Function

'******************************************************************************
' Lecture d'un élément de la matrice UVP
'*******************************************************************************
Public Function getQ(ByVal numBrancheEntrée As Integer, ByVal numBrancheSortie As Integer, Optional ByVal véhicule As Integer = UVP) As Integer
  getQ = QVéhicule(véhicule, numBrancheEntrée, numBrancheSortie)
End Function

'******************************************************************************
' Lecture d'un élément de la matrice Piéton
'*******************************************************************************
Public Function getQP(ByVal NumBranche As Integer) As Integer
  getQP = QP(NumBranche)
End Function

'******************************************************************************
' Affectation d'un trafic Piéton
'******************************************************************************
Public Sub setQP(ByVal NumBranche As Integer, ByVal Valeur As String)
  QP(NumBranche) = convChaineEntier(Valeur)
End Sub

'******************************************************************************
' Affectation d'un trafic VL,PL, 2R ou UVP
' Si VéhiculeCourant <> UVP, celui-ci est systématiquement calculé
'*******************************************************************************
Public Sub setQ(ByVal numBrancheEntrée As Integer, ByVal numBrancheSortie As Integer, ByVal Valeur As String, Optional ByVal SaisieEnCours As Boolean)
Dim i As Integer
Dim newValeur As Integer
Dim exValeur As Integer

  ' En lecture de fichier, la valeur non renseignée est ID_Vide ("VIDE")
  If Valeur = ID_Vide Then Exit Sub
  
  exValeur = equivalTraficNum(QVéhicule(VéhiculeCourant, numBrancheEntrée, numBrancheSortie))
  newValeur = convChaineEntier(Valeur)
  QVéhicule(VéhiculeCourant, numBrancheEntrée, numBrancheSortie) = newValeur
 
    ' Calcul du Trafic UVP
  Dim exValeurUVP As Integer
  Dim newValeurUVP As Integer
  
  If VéhiculeCourant = UVP Then
    exValeurUVP = exValeur
    newValeurUVP = equivalTraficNum(newValeur)
  Else
    exValeurUVP = equivalTraficNum(QVéhicule(UVP, numBrancheEntrée, numBrancheSortie))
    newValeur = equivalTraficNum(newValeur)
    newValeurUVP = exValeurUVP + Int((newValeur - exValeur) * gbCoefVehicule(VéhiculeCourant))
    If VéhiculeCourant = DEUXR Then
    ' Dans le cas des 2Roues et si les 2 valeurs n'ont pas même parité, la fonction Int ci-dessus
    ' arrondit (par exemple) 2.5 en 2
    ' On rajoute 1 une fois sur 2 : lorsqu'on passe d'une valeur impaire à une valeur paire
      If Impaire(exValeur) And Not Impaire(newValeur) Then newValeurUVP = newValeurUVP + 1
    End If
    If newValeurUVP >= 0 Then
      QVéhicule(UVP, numBrancheEntrée, numBrancheSortie) = newValeurUVP
    Else
      QVéhicule(UVP, numBrancheEntrée, numBrancheSortie) = VIDE
    End If
  End If
  
    ' Calculer et afficher les totaux  correspondants
    QE(numBrancheEntrée) = QE(numBrancheEntrée) + newValeurUVP - exValeurUVP
    QS(numBrancheSortie) = QS(numBrancheSortie) + newValeurUVP - exValeurUVP
  If SaisieEnCours Then
    AfficheTraficTotal numBrancheEntrée, numBrancheSortie, equivalTraficNum(newValeur) - exValeur
  End If
  
End Sub

'************************************************************************************
' Afichage des Trafics Total Entrant, Total Sortant et Total dans la grille de saisie
'************************************************************************************
Private Sub AfficheTraficTotal(ByVal numBrancheEntrée As Integer, ByVal numBrancheSortie As Integer, ByVal Différence As Integer)
Dim i As Integer
  With gbProjetActif.Données
    .txtQE(numBrancheEntrée) = equivalTraficNum(convChaineEntier(.txtQE(numBrancheEntrée))) + Différence
    .txtQS(numBrancheSortie) = equivalTraficNum(convChaineEntier(.txtQS(numBrancheSortie))) + Différence
    .txtQT = equivalTraficNum(convChaineEntier(.txtQT)) + Différence
    
    If .txtQE(numBrancheEntrée) = 0 Then
      .txtQE(numBrancheEntrée) = ""
      For i = 1 To gbProjetActif.NbBranches
        If QVéhicule(VéhiculeCourant, numBrancheEntrée, i) = 0 Then
          .txtQE(numBrancheEntrée) = 0
          Exit For
        End If
      Next
    End If
    
    If .txtQS(numBrancheSortie) = 0 Then
      .txtQS(numBrancheSortie) = ""
      For i = 1 To gbProjetActif.NbBranches
        If QVéhicule(VéhiculeCourant, i, numBrancheSortie) = 0 Then
          .txtQS(numBrancheSortie) = 0
          Exit For
        End If
      Next
    End If
    
    If .txtQT = 0 Then .txtQT = ""
  End With
  
End Sub

Public Sub AnnulerTrafic(ByVal LigneColonne As Integer, ByVal num As Integer, ByVal SaisieEnCours As Boolean)
Const LIGNE = 0
Const COLONNE = 1
Dim i As Integer
Dim Mode As Integer
Dim SauvVéhiculeCourant As Integer

If modeUVP Then
  For i = 1 To gbProjetActif.NbBranches
    If LigneColonne = LIGNE Then
      setQ num, i, "", SaisieEnCours
    Else
      setQ i, num, "", SaisieEnCours
    End If
  Next
  
Else
  SauvVéhiculeCourant = VéhiculeCourant
  For Mode = VL To DEUXR
    VéhiculeCourant = Mode
    For i = 1 To gbProjetActif.NbBranches
      If LigneColonne = LIGNE Then
        setQ num, i, "", SaisieEnCours And SauvVéhiculeCourant = VéhiculeCourant
      Else
        setQ i, num, "", SaisieEnCours And SauvVéhiculeCourant = VéhiculeCourant
      End If
    Next
  Next
  VéhiculeCourant = SauvVéhiculeCourant
End If

With gbProjetActif
  With .Données.vgdTrafic(VEHICULE)
    If LigneColonne = LIGNE Then .Row = num Else .Col = num
    For i = 1 To gbProjetActif.NbBranches
      If LigneColonne = LIGNE Then .Col = i Else .Row = i
      .Value = ""
    Next
  End With
End With

End Sub


'******************************************************************************
' Changement du type de matrices : VL - PL - 2R - UVP
' Affichage des valeurs dans la grille
'*******************************************************************************
Public Sub setVéhicule(ByVal véhicule As Integer)
Dim NbBranches As Integer
Dim FeuilleDonnées As Form
Dim i As Integer
Dim j As Integer
Dim newValeur As Integer

  Set FeuilleDonnées = gbProjetActif.Données
  NbBranches = gbProjetActif.NbBranches
  VéhiculeCourant = véhicule

  With FeuilleDonnées.vgdTrafic(VEHICULE)
    .ReDraw = False
    FeuilleDonnées.txtQT = ""
    For j = 1 To NbBranches
      FeuilleDonnées.txtQS(j) = ""
    Next
    For i = 1 To NbBranches
      FeuilleDonnées.txtQE(i) = ""
      .Row = i
      For j = 1 To NbBranches
        .Col = j
        .Value = convEntierChaine(QVéhicule(VéhiculeCourant, i, j))
        newValeur = QVéhicule(VéhiculeCourant, i, j)
        AfficheTraficTotal i, j, equivalTraficNum(newValeur)
      Next
    Next
    .ReDraw = True
  End With
  
  With FeuilleDonnées.vgdTrafic(PIETON)
    .ReDraw = False
    .Row = 1
    For i = 1 To NbBranches
      .Col = i
      .Value = convEntierChaine(QP(i))
    Next
    .ReDraw = True
  End With

End Sub

'******************************************************************************
' Formatage du trafic Piéton : Chaine à écrire dans le fichier
'******************************************************************************
Public Function FormaterQP() As String
  FormaterQP = Formater(getQPiéton())
End Function

'******************************************************************************
' Formatage du trafic Véhicule : Chaine à écrire dans le fichier
'******************************************************************************
Public Function FormaterQV(ByVal véhicule As Integer, ByVal numBrancheEntrée As Integer) As String
  FormaterQV = Formater(getQVéhicule(numBrancheEntrée, véhicule))
End Function

'******************************************************************************
' Formate une chaine de manière à écrire une ligne de matrice sur une ligne du fichier
'*******************************************************************************
Private Function Formater(ByVal Valeurs As Variant) As String
Dim chaine As String
Dim i As Integer
  For i = 1 To UBound(QP) '(1)
    chaine = chaine + equivalTraficStr(Valeurs(i))
  Next
  Formater = Left(chaine, Len(chaine) - 1)

End Function

'******************************************************************************
' Calcul l'équivalence de la valeur d'un trafic selon qu'il est renseigné ou non
' Module Ecriture dans un fichier
'*******************************************************************************
Private Function equivalTraficStr(ByVal Valeur As Integer) As String
  If Valeur = VIDE Then
    equivalTraficStr = ID_Vide & " "
  Else
    equivalTraficStr = Format(Valeur, String(LGTRAFIC, "0") & " ")
  End If
  
End Function

'******************************************************************************
' Lecture des données d'une période de trafic
'******************************************************************************
Public Function Lire(ByVal numFich As Integer, ByRef numLigne As Integer) As Boolean
Dim chaine As String
Dim i As Integer
Dim véhicule As Integer
Dim NbBranches As Integer
  
  On Error GoTo TraitementErreur

  NbBranches = UBound(QP) '(1)
  Input #numFich, chaine
  numLigne = numLigne + 1
  DécryptTraficP chaine
  If modeUVP Then
    For i = 1 To NbBranches
      Input #numFich, chaine
      numLigne = numLigne + 1
      DécryptTrafic UVP, i, chaine
    Next
  Else
    For véhicule = VL To DEUXR
      For i = 1 To NbBranches
        Input #numFich, chaine
        numLigne = numLigne + 1
        DécryptTrafic véhicule, i, chaine
      Next
    Next
  End If
  
  Lire = True
  Exit Function
  
TraitementErreur:
  ErreurFatale
'  If Err <> 13 Then ErreurFatale
' L'erreur 13 'Type incompatible' est survenue lors de Cint dans EquivalTrafic
End Function

'******************************************************************************
' Affectation du trafic Piéton : décryptage d'une chaine lue dans le fichier
'******************************************************************************
Private Sub DécryptTraficP(ByVal chaine As String)
Dim i As Integer
  For i = 1 To UBound(QP) '(1)
    setQP i, Mid(chaine, 1 + (LGTRAFIC + 1) * (i - 1), LGTRAFIC)
'    setQP i, EquivalTrafic(Mid(chaine, 1 + (LGTRAFIC + 1) * (i - 1), LGTRAFIC))
  Next
End Sub

'******************************************************************************
' Affectation du trafic Véhicule : décryptage d'une chaine lue dans le fichier
'******************************************************************************
Private Sub DécryptTrafic(ByVal véhicule As Integer, ByVal numBrancheEntrée, ByVal chaine As String)
Dim i As Integer
  
  VéhiculeCourant = véhicule
  For i = 1 To UBound(QP) '(1)
    setQ numBrancheEntrée, i, Mid(chaine, 1 + (LGTRAFIC + 1) * (i - 1), LGTRAFIC)
'    setQ numBrancheEntrée, i, EquivalTrafic(Mid(chaine, 1 + (LGTRAFIC + 1) * (i - 1), LGTRAFIC))
   Next
  
End Sub

'******************************************************************************
' Calcul l'équivalence de la valeur d'un trafic selon qu'il est renseigné ou non
' Module Lecture d'un fichier
'*******************************************************************************
''Private Function EquivalTrafic(ByVal chaine As String) As Integer
''  If chaine = ID_Vide Then
''    EquivalTrafic = VIDE
''  Else
''    EquivalTrafic = cint(chaine)
''  End If
''End Function
''
Private Function convChaineEntier(ByVal Valeur As String) As Integer
  If Valeur = "" Then
    convChaineEntier = VIDE
  Else
    convChaineEntier = CInt(Valeur)
  End If
End Function

Private Function convEntierChaine(ByVal Valeur As Integer) As String
   If Valeur = VIDE Then
    convEntierChaine = ""
   Else
    convEntierChaine = Valeur
  End If
End Function

'******************************************************************************
' Calcul l'équivalence de la valeur d'un trafic selon qu'il est renseigné ou non
' Fonctions setQ et setVéhicule
'*******************************************************************************
Private Function equivalTraficNum(ByVal Valeur As Integer) As Integer
  If Valeur = VIDE Then
    equivalTraficNum = 0
  Else
    equivalTraficNum = Valeur
  End If
  
End Function


'******************************************************************************
' Calcul des diagrammes de flux
'*******************************************************************************
Public Sub CalculDiagramFlux(Optional ByVal Periph)

Dim i As Integer
Dim j As Integer
Dim NbBranches As Integer
Dim AngleArcBranche(0 To 1) As Single
Dim AngleDep As Single
Dim AngleFin As Single
Dim SauvFillColor As Long
Dim Rayon As Single, RayonUnitésDessin As Single, RayonUtile As Single
Dim ObjetDessin As Object
Dim DrawWidthMax As Single
Dim AngleFinPrec As Single
Dim dAlpha As Single
Dim XCentre As Single, YCentre As Single, Echelle As Single
Dim QPortion As Integer
Dim QMax As Integer
Dim QEnCours As Integer
Dim CoefEpaisseur As Single

Dim X1 As Single
Dim Y1 As Single
Dim X2 As Single
Dim Y2 As Single
Dim TraficSuperflu As Integer

NbBranches = gbProjetActif.NbBranches
ReDim QEReel(1 To NbBranches)
ReDim QSReel(1 To NbBranches)

If IsMissing(Periph) Then
  Set Periph = Screen
  Set ObjetDessin = gbProjetActif.Données
Else
  Set ObjetDessin = Periph  ' Printer
End If

ObjetDessin.FillStyle = vbFSSolid
SauvFillColor = ObjetDessin.FillColor
On Error Resume Next
ObjetDessin.FillColor = ObjetDessin.BackColor
On Error GoTo 0


If ObjetDessin Is Printer Then
  XCentre = frmImprimer.XCentre
  YCentre = frmImprimer.YCentre
  Echelle = frmImprimer.Echelle
Else
  XCentre = gbDemiLargeur
  YCentre = gbDemiHauteur
  Echelle = 1 / gbFacteurZoom
End If

' Pas d'épaisseur de trait supérieur à 1 cm (2cm sur l'imprimante)- environ : 1cm=567 twips
' On conserve également au moins l'équivalent d'1m à l'intérieur de l'anneau
If ObjetDessin Is Printer Then
  DrawWidthMax = Min(567 * 2, (gbProjetActif.LA - 2) / Echelle * 567) / Periph.TwipsPerPixelX
Else
  DrawWidthMax = Min(567, (gbProjetActif.LA - 2) / Echelle) / Periph.TwipsPerPixelX
End If

Rayon = gbProjetActif.R + gbProjetActif.Bf + 1
RayonUnitésDessin = Rayon / Echelle

'Calcul du trafic sur la portion d'anneau précédent la première branche (B-nbbranches à B-1)
For i = 1 To NbBranches
  For j = 1 To i  ' Ne pas prendre en compte les trafics sortant d'une branche de + gd numéro que l'entrante(sortant avant la Branche 1)
    QPortion = QPortion + equivalTraficNum(QVéhicule(UVP, i, j))
  Next
Next
If gbProjetActif.colBranches.Item(NbBranches).TAD Then QPortion = QPortion - equivalTraficNum(QVéhicule(UVP, NbBranches, 1))
QMax = QPortion

For i = 1 To NbBranches
  If gbProjetActif.colBranches.Item(i).TAD Then
    TraficSuperflu = equivalTraficNum(QVéhicule(UVP, i, BrancheSuivant(i)))
  Else
    TraficSuperflu = 0
  End If
  QEReel(i) = QE(i) - TraficSuperflu
  QSReel(BrancheSuivant(i)) = QS(BrancheSuivant(i)) - TraficSuperflu
Next

'Calcul des autres trafics : Trafic d'une portion = Trafic de la portion suivante diminué du trafic entrant augmenté du trafic sortant
For i = NbBranches To 2 Step -1
  QPortion = QPortion - QEReel(i) + QSReel(i)
  QMax = Max(QMax, QPortion)
Next

CoefEpaisseur = QMax / Min(DrawWidthMax, QMax / 25)

' QPortion est le trafic sur la portion d'anneau B1--> B2
QEnCours = QPortion

' Affichage du diagramme sur l'anneau
With gbProjetActif.colBranches
  
  For i = 1 To NbBranches
    If Periph Is Screen Then MasquerIlot i
    
  ' Trafic sur l'anneau
  
  ' Angle de l'axe du flux d'entrée : AngleArcBranche(1)
    With .Item(i)
      AngleArcBranche(1) = (.LE4m / 2) / Rayon
      ' AngleDep = Angle de la branche + angle formé par son extrémité et celle fictive d'une demi-largeur d'entrée
      AngleDep = angConv(.Angle, CVRADIAN) + AngleArcBranche(1)
    End With  ' Branche i
    
    ' Dessin de l'arc au droit de la branche précédente
    If i > 1 Then ' Pour la 1è branche, le calcul du dernier petit arc correspondant à la sortie sera effectué en fin
      AngleFinPrec = AngleDep - AngleArcBranche(0) - AngleArcBranche(1)
      CalculRayonFlux RayonUnitésDessin, QEnCours, CoefEpaisseur, ObjetDessin, Periph, RayonUtile
      ObjetDessin.Circle (XCentre, YCentre), RayonUtile, QBColor(9), AngleFinPrec, AngleDep
      ' QEnCours : Trafic sur la portion d'anneau Bi--> Bi+1
      QEnCours = QEnCours + QEReel(i)
    End If
    
  ' Angle de l'axe du flux de sortie : AngleArcBranche(0)
    With .Item(BrancheSuivant(i))
      AngleArcBranche(0) = (.LS / 2) / Rayon
      AngleFin = angConv(.Angle, CVRADIAN) - AngleArcBranche(0)
      If AngleFin <= 0 Then AngleFin = AngleFin + 2 * PI   ' Possible en fait Si BrancheSuivant(i)=1 (Angle=0)
    End With  ' BrancheSuivant(i)

    ' Portion  principale d'anneau (AngleDep et AngleFin excluant la demi-entrée de départ et la demi-sortie d'arrivée
    If QEnCours >= 1 Then
      CalculRayonFlux RayonUnitésDessin, QEnCours, CoefEpaisseur, ObjetDessin, Periph, RayonUtile, dAlpha
      ObjetDessin.Circle (XCentre, YCentre), RayonUtile, QBColor(9), AngleDep + dAlpha, AngleFin - dAlpha
    End If

    If ObjetDessin Is Printer Then
      ' Bouchage des trous
    End If

      ' QEnCours : Trafic sur la portion d'anneau au droit de la branche Bi+1 (entre la sortie et l'entrée)
    QEnCours = QEnCours - QSReel(BrancheSuivant(i))

    ' Trafic sur les Branches

  ' Trafic entrant
    If QEReel(i) >= 1 Then
      CalculRayonFlux RayonUnitésDessin, QEReel(i), CoefEpaisseur, ObjetDessin, Periph, RayonUtile
      X1 = XCentre + RayonUtile * Cos(AngleDep)
      Y1 = YCentre + RayonUtile * Sin(-AngleDep)
      X2 = X1 + maLongueurBranche * Cos(angConv(.Item(i).Angle, CVRADIAN)) / Echelle
      Y2 = Y1 + maLongueurBranche * Sin(-angConv(.Item(i).Angle, CVRADIAN)) / Echelle
      
      ObjetDessin.Line (X1, Y1)-(X2, Y2), QBColor(9)
      If ObjetDessin Is Printer Then
        EffacerArrondiImprimante .Item(i), XCentre, YCentre, Echelle, ObjetDessin, Periph, Entrée:=True
      Else
        EffacerArrondiBranche AngleDep, X2, Y2, ObjetDessin, Periph
      End If
      
      ObjetDessin.DrawWidth = 1
    End If
  
  ' Trafic sortant
    j = BrancheSuivant(i)
    If QSReel(j) >= 1 Then
      CalculRayonFlux RayonUnitésDessin, QSReel(j), CoefEpaisseur, ObjetDessin, Periph, RayonUtile
      X1 = XCentre + RayonUtile * Cos(AngleFin)
      Y1 = YCentre + RayonUtile * Sin(-AngleFin)
      X2 = X1 + maLongueurBranche * Cos(angConv(.Item(j).Angle, CVRADIAN)) / Echelle
      Y2 = Y1 + maLongueurBranche * Sin(-angConv(.Item(j).Angle, CVRADIAN)) / Echelle
      
      ObjetDessin.Line (X1, Y1)-(X2, Y2), QBColor(9)
      If ObjetDessin Is Printer Then
        EffacerArrondiImprimante .Item(j), XCentre, YCentre, Echelle, ObjetDessin, Periph, Entrée:=False
      Else
        EffacerArrondiBranche AngleFin, X2, Y2, ObjetDessin, Periph
      End If
      
      ObjetDessin.DrawWidth = 1
    End If
  
  Next

    ' Dessin du dernier arc, correspondant au droit de la première branche
    With .Item(1)
      AngleArcBranche(1) = (.LE4m / 2) / Rayon
      CalculRayonFlux RayonUnitésDessin, QEnCours, CoefEpaisseur, ObjetDessin, Periph, RayonUtile
      ObjetDessin.Circle (XCentre, YCentre), RayonUtile, QBColor(9), 2 * PI - AngleArcBranche(0), AngleArcBranche(1)
    End With  ' Branche 1

End With  ' colBranches

ObjetDessin.FillStyle = vbFSTransparent
On Error Resume Next
ObjetDessin.FillColor = SauvFillColor
ObjetDessin.DrawWidth = 1

End Sub

Private Sub MasquerIlot(ByVal NumBranche As Integer, Optional ByVal wLigne As Line)

  If wLigne Is Nothing Then
    With gbProjetActif.Données
      MasquerIlot 0, .linBordIlotEntrée(NumBranche)
      MasquerIlot 0, .linBordIlotSortie(NumBranche)
      MasquerIlot 0, .linBordIlotGir(NumBranche)
    End With
  Else
    If wLigne.Visible Then
      wLigne.Tag = "V"
      wLigne.Visible = False
    End If
  End If

End Sub
Private Sub CalculRayonFlux(ByVal R As Single, ByVal Q As Integer, ByVal CoefEpaisseur As Single, ByVal ObjetDessin As Object, ByVal Periph, ByRef RayonUtile As Single, Optional ByRef dAlpha)
Dim DW As Single
Dim DeltaR As Single

  DW = Max(1, (Q / CoefEpaisseur))
  ObjetDessin.DrawWidth = DW
  DeltaR = TransformUniteDessin((DW + 1) / 2 * Periph.TwipsPerPixelX, ObjetDessin)
  RayonUtile = R + DeltaR
  If Not IsMissing(dAlpha) Then
    dAlpha = 0.5 * DeltaR / R
  End If
End Sub

Private Function TransformUniteDessin(ByVal Valeur As Single, ByVal ObjetDessin As Object)
  If ObjetDessin Is Printer Then
    TransformUniteDessin = Valeur / 567   ' Transformation twips-->cm
  Else
    TransformUniteDessin = Valeur
  End If
End Function

Private Sub EffacerArrondiBranche(ByVal Angle As Single, ByVal X As Single, ByVal Y As Single, ByVal ObjetDessin As Object, ByVal Periph)
Dim Depart As Single, Final As Single
  Depart = Angle - PI / 2
  If Depart < 0 Then Depart = Depart + 2 * PI
  If Depart = 0 Then Depart = 0.0001 ' Sinon le tracé de secteur ne se fait pas (cf Aide en ligne de VB : CIRCLE argument start)
  If Depart > 2 * PI Then Depart = Depart - 2 * PI
  Final = Depart + PI
  If Final < 0 Then Final = Final + 2 * PI
  If Final > 2 * PI Then Final = Final - 2 * PI
  ObjetDessin.Circle (X, Y), (ObjetDessin.DrawWidth + 1) / 2 * Periph.TwipsPerPixelX, ObjetDessin.BackColor, -Depart, -Final

End Sub

Private Sub EffacerArrondiImprimante(ByVal wBranche As BRANCHE, ByVal XCentre As Single, YCentre As Single, ByVal Echelle As Single, ByVal ObjetDessin As Object, ByVal Periph, ByVal Entrée As Boolean)
Dim Xmin As Single, Ydeb As Single, Yfin As Single, AngleBranche As Single, X1 As Single, Y1 As Single, X2 As Single, Y2 As Single
Dim i As Integer, DW As Single, Largeur As Single

  If Entrée Then Largeur = -wBranche.LE4m / 2 Else Largeur = wBranche.LS / 2
  
  DW = ObjetDessin.DrawWidth
  ObjetDessin.DrawWidth = 1
  AngleBranche = angConv(wBranche.Angle, CVRADIAN)
  
 ' If Entrée Then
    Ydeb = (Largeur + DW * 0.5 * Periph.TwipsPerPixelX / 567 * Echelle)
    Yfin = (Largeur - DW * 0.5 * Periph.TwipsPerPixelX / 567 * Echelle)
    Ydeb = Ydeb / Echelle
    Yfin = Yfin / Echelle
    Ydeb = DW * 0.5 * Periph.TwipsPerPixelX / 567
    Yfin = -DW * 0.5 * Periph.TwipsPerPixelX / 567
    X1 = Printer.CurrentX + Ydeb * Sin(AngleBranche)
    Y1 = Printer.CurrentY + Ydeb * Cos(AngleBranche)
    X2 = Printer.CurrentX + Yfin * Sin(AngleBranche)
    Y2 = Printer.CurrentY + Yfin * Cos(AngleBranche)
    For i = 0 To CInt((DW + 1) / 2)
      ObjetDessin.Line (X1, Y1)-(X2, Y2), QBColor(9)
      X1 = X1 + 0.01 * Cos(AngleBranche)
      Y1 = Y1 + 0.01 * Sin(-AngleBranche)
      X2 = X2 + 0.01 * Cos(AngleBranche)
      Y2 = Y2 + 0.01 * Sin(-AngleBranche)
    Next

End Sub

''Private Function EquivalentImprim(ByVal Coord As Single, ByVal DemiDimension As Single, ByVal CoordCentre As Single, ByVal Echelle As Single) As Single
''    ' Emprunté à DessinerControleLigne dans frmImprimer
''    EquivalentImprim = (Coord - DemiDimension) / gbFacteurZoom / Echelle + CoordCentre

''End Function

'******************************************************************************
' Détermine si une matrice de trafic est complète
'******************************************************************************
Public Function EstComplète(Optional ByVal véhicule As Integer = 4) As Boolean

If véhicule = 4 Then
  If VéhiculeCourant = UVP And modeUVP Then
    EstComplète = EstComplète(UVP)
  Else
    EstComplète = EstComplète(VL) And EstComplète(PL) And EstComplète(DEUXR)
  End If
  Exit Function
End If

  Dim wBranches As Branches
  Dim i As Integer
  Dim j As Integer
  Set wBranches = gbProjetActif.colBranches
  For i = 1 To UBound(QP) '(1)
    For j = 1 To UBound(QP)
      If QVéhicule(véhicule, i, j) = VIDE Then
'        If wBranches.Item(i).LE4m <> 0 And wBranches.Item(j).LS <> 0 Then Exit Function
        If Not (wBranches.Item(i).EntréeNulle Or wBranches.Item(j).SortieNulle) Then Exit Function
      End If
    Next
  Next
  EstComplète = True
  
End Function

'******************************************************************************
' Duplication de matrice  : par import, inversion ou multiplication
' TypeDupli = INVMATRICE,MULTMATRICE, ou IMPMATRICE
'******************************************************************************
Public Sub Inverser()
  Dupliquer INVMATRICE
End Sub
Public Sub Multiplier()
  Dupliquer MULTMATRICE
End Sub
Public Sub Importer()
  Dupliquer IMPMATRICE
End Sub

Private Sub Dupliquer(ByVal TypeDupli As Integer)
  Dim i As Integer
  Dim wTrafic As TRAFIC
  Dim cboPériode As ComboBox
  
  With gbProjetActif.colTrafics
    If TypeDupli = MULTMATRICE And frmTrafMult.Saturation Then
  ' Le trafic obtenu par duplication (et donc par saturation) est le dernier créé
      Set wTrafic = .Item(gbProjetActif.colTrafics.count)
      TraficFils(wTrafic.BrancheSaturée) = True
      Set wTrafic.TraficPère = Me
    Else
      Set wTrafic = .Item(gbProjetActif.nbPériodes)
    End If
  End With
  
  With wTrafic
    ' Pour la saturation, on ne s'intéresse qu'au mode UVP (valeur par défaut dans Trafics.add)
    If .BrancheSaturée = 0 Then .modeUVP = modeUVP
    ' On duplique tel quel le trafic Piéton
    For i = 1 To UBound(QP) '(1)
      .setQP i, convEntierChaine(QP(i))
    Next
    Duplication wTrafic, TypeDupli
    
    If .BrancheSaturée <> 0 Then
    ' Activation de la période saturée créée comme période courante
      With gbProjetActif.Résultats.cboPériode
        .ListIndex = .ListCount - 1
      End With
    Else
    ' Activation de la période créée comme période courante
      With gbProjetActif.Données.cboPériode
        .ListIndex = .ListCount - 1
      End With
    End If
  End With
  
End Sub

Private Sub Duplication(ByVal wTrafic As TRAFIC, ByVal TypeDupli As Integer, Optional ByVal véhicule As Integer = 4)
  Dim i As Integer
  Dim j As Integer
  
  If véhicule = 4 Then
  ' Rechercher s'il faut dupliquer une ou 4 matrices
'    If VéhiculeCourant = UVP Then
    If modeUVP Then
      Duplication wTrafic, TypeDupli, UVP
    Else
      For i = VL To DEUXR
        Duplication wTrafic, TypeDupli, i
      Next
    End If
    Exit Sub
  End If
    
  Dim wBranches As Branches
  Set wBranches = gbProjetActif.colBranches
  
  With wTrafic
    .VéhiculeCourant = véhicule
    For i = 1 To UBound(QP) '(1)
      For j = 1 To UBound(QP) '(1)
'        If wBranches.Item(i).LE4m = 0 Or wBranches.Item(j).LS = 0 Then
        If wBranches.Item(i).EntréeNulle Or wBranches.Item(j).SortieNulle Then
          ' Debug --> AV - 17.03.99 - Reste à gérer les messages d'avertissement : Trafic Indéterminé
          .setQ i, j, ""
        Else
          If TypeDupli = IMPMATRICE Then        ' Recopie d'import
            .setQ i, j, convEntierChaine(QVéhicule(véhicule, i, j))
          ElseIf TypeDupli = INVMATRICE Then    'Inversion
            .setQ i, j, convEntierChaine(QVéhicule(véhicule, j, i))
          Else                                  ' Multiplication
            .setQ i, j, convEntierChaine(QVéhicule(véhicule, i, j)) * frmTrafMult.coeff(i, j)
          End If
        End If
      Next
    Next
  End With
  
End Sub

'******************************************************************************
' Calcul du Trafic Entrant
'******************************************************************************
Public Sub CalculTraficEntrant(ByVal wBranches As Branches)
Dim i As Integer
Dim j As Integer
Dim QV As Integer
Dim NbBranches As Integer

  NbBranches = wBranches.count
  ReDim Q(1 To NbBranches, 1 To NbBranches)
  ReDim QEntrant(1 To NbBranches)  ' Trafic tournant
  ReDim QTournant(1 To NbBranches)  ' Trafic tournant
  ReDim QG(1 To NbBranches)         ' Trafic gênant
  ReDim C(1 To NbBranches)          ' Capacité définitive
  ReDim Cp(1 To NbBranches)         ' Perte de capacité due aux piétons
  ReDim rc(1 To NbBranches)         ' Réserve de capacité (en uvp/h)
  ReDim TMA(1 To NbBranches)        ' Temps moyen d'attente
  ReDim TTA(1 To NbBranches)        ' Temps total d'attente
  ReDim LK(1 To NbBranches)         ' Longueur moyenne de stockage
  ReDim LKM(1 To NbBranches)        ' Longueur totale de stockage
  
  For i = 1 To NbBranches
    QEntrant(i) = QE(i)
    For j = 1 To NbBranches
      QV = QVéhicule(UVP, i, j)
    ' Recherche des valeurs à prendre en compte
    ' Ref NOTE de CALCUL §1.3
      If wBranches.Item(i).TAD And j = BrancheSuivant(i) Then ' BrancheSuivant est définie dans DessinGiratoire.BAS
        Q(i, j) = 0
        QEntrant(i) = QEntrant(i) - QV
      ElseIf QV = VIDE Then
        Q(i, j) = 0
      Else
        Q(i, j) = QV
      End If
    Next
  Next
 
End Sub

'******************************************************************************
' Calcul de capacité d'une branche
'******************************************************************************
Public Sub CalculCapacité(NumBranche As Integer, wBranches As Branches)
'___________________________________________________________________________________
Dim wBranche As BRANCHE

'Set BrancheActive = BranchesActive.Item(numBranche)
Set wBranche = wBranches.Item(NumBranche)

If wBranche.EntréeNulle Then Exit Sub

TraficGenant NumBranche, wBranches
RéserveCapacité NumBranche, wBranche
TempsAttente NumBranche, wBranche
LongueurStockage NumBranche, wBranche

End Sub

'******************************************************************************
' Calcul du trafic gênant
' Ref : Doc Interne pour la renumérotation des branches
'******************************************************************************
Private Sub TraficGenant(ByVal NumBranche As Integer, ByVal wBranches As Branches)

' Calcul de QG : Trafic gênant au droit de l'entrée
'___________________________________________________
  Dim Angle As Integer
  Dim eQvi As Integer, eQvj As Integer
  Dim i As Integer, j As Integer
  Dim wQTournant          ' Trafic tournant en provenance d'une branche
  Dim wQTE As Integer     ' Trafic tournant à l'extérieur, en provenance d'une branche
  Dim QTE As Integer      ' Trafic tournant à l'extérieur
  Dim QTI As Integer      ' Trafic tournant à l'intérieur
  Dim QT As Integer       ' Trafic tournant total
  Dim NbBranches As Integer
  Dim AngleFin As Integer
    
On Error GoTo GestErr

    NbBranches = UBound(QP)
    Angle = wBranches.Item(NumBranche).Angle
    
    ' Trafic tournant
    For i = 1 To NbBranches - 1
      eQvi = numReelBranche(i, NumBranche)
      For j = 1 To NbBranches - 1
        If j <= i Then
          eQvj = numReelBranche(j, NumBranche)
          ' Trafic tournant à prendre en compte au niveau de la branche
          wQTournant = Q(eQvi, eQvj)
          ' Trafic tournant à l'extérieur de l'anneau
          AngleFin = wBranches.Item(eQvj).Angle
          wQTE = KAE(Angle, AngleFin, j) * wQTournant
          QTE = QTE + wQTE
          ' Trafic tournant à l'intérieur de l'anneau
          QTI = QTI + wQTournant - wQTE
          QT = QT + wQTournant
        End If
      Next
    Next
  
    ' Trafic tournant gênant
    ' Ref NOTE de CALCUL §2.3.3
    Dim QTournantGenant As Integer
    With gbProjetActif
      QTournantGenant = .KE * QTE + .KI * QTI
    End With
    
    ' Ref NOTE de CALCUL §2.4 modifié suite à discussion avec B.GUICHET du 25.03.99
    ' Trafic sortant gênant
    Dim QSortant As Integer
    Dim QSortantGenant As Integer
    For i = 1 To NbBranches
      QSortant = QSortant + Q(i, NumBranche)
    Next
    
    'Test rajouté par AV (20/11/03) suite à un exemple bizarre d'un utilisateur
    If QT + QSortant > 0 Then
      QSortantGenant = wBranches.Item(NumBranche).KS * QSortant * (1 - QSortant / (QT + QSortant))
      ' Trafic gênant (tournant et sortant pondéré)
      QG(NumBranche) = QTournantGenant + QSortantGenant
    Else
    ' Il n'y aucun trafic sur la portion d'anneau qui précède l'entrée!!!!
      QG(NumBranche) = QTournantGenant
    End If
    
    'Mémorisation du trafic tournant
    QTournant(NumBranche) = QT
    
Exit Sub
GestErr:
    End Sub
  
Private Function numReelBranche(ByVal i As Integer, ByVal origine As Integer) As Integer
' Pour le calcul du trafic tournant, on renumérote les branches du Giratoire de 0 à n-1, la branche 0 étant toujours
' la branche en cours d'étude (branche origine)afin de pouvoir comparer + aisément l'ordre des branches :
' Un trafic tournant sera pris en compte uniquement s'il sort sur une branche de numéro inférieur ou égal à celui
' de la branche d'entrée (il est bien passé devant la branche 0)

' Pour plus de détails, voir la documentation interne

' Cette fonction permet de retrouver le numéro réel de la branche i

  numReelBranche = (i - 1 + origine) Mod gbProjetActif.NbBranches + 1
  
End Function
 
'******************************************************************************
' Calcul de KAE : Coeffficient d'affectation de trafic à l'extérieur de l'anneau
'******************************************************************************
Private Function KAE(ByVal AngleBrancheEnCours As Integer, ByVal AngleBrancheSortie As Integer, ByVal j As Integer) As Single
Dim Ecart As Integer

    ' Ref NOTE de CALCUL §2.3.2 + Doc Interne (Trafic tournant gênant)
  If j = 1 Or j = 2 Then
  'Trafic sortant sur une des 2 branches qui suivent l'entrée en cours d'étude (AngleBrancheEnCours)
    If AngleBrancheSortie < AngleBrancheEnCours Then AngleBrancheSortie = AngleBrancheSortie + 2 * eqvPI(gbProjetActif.modeangle)
    Ecart = Abs(AngleBrancheEnCours - AngleBrancheSortie)
    
  ' Transformation en grade, les formules de la note de calcul étant en degrés
    If gbProjetActif.modeangle = GRADE Then Ecart = Ecart * 0.9
  ' La formule qui suit était errnonée (AV : 16/08/2004)- corrigé en v 1.0.24
  '  If gbProjetActif.modeangle = GRADE Then Ecart = Ecart / 0.9
    
    If Ecart < 70 Then
      KAE = 1#
''      If gbFichierJournal Then Write #gbFichLog, "2.3.2.2 - KAE calculé avec angle < 70"
    ElseIf Ecart > 115 Then
      KAE = 0.1
''      If gbFichierJournal Then Write #gbFichLog, "2.3.2.2 - KAE calculé avec angle > 115"
    Else
      KAE = (120 - Ecart) / 50
''      If gbFichierJournal Then Write #gbFichLog, "2.3.2.2 - KAE calculé avec 70<angle<115"
    End If
  Else
    KAE = 0.1
''    If gbFichierJournal Then Write #gbFichLog, "2.3.2.1 - KAE = 0.1"
  End If
  
End Function
'******************************************************************************
' Calcul de capacité : C et RC
'******************************************************************************
Private Sub RéserveCapacité(ByVal NumBranche As Integer, ByVal wBranche As BRANCHE)
  Dim Cvh As Double   ' Capacité en véhicule/heure
'  Dim Cp As Double    ' Perte de capacité due aux piétons
  
  ' Ref NOTE de CALCUL §2.5.1  : Calcul de Ci
  Dim exposant As Double
  Cvh = wBranche.getCVH(QG(NumBranche), exposant)
  ' Ref NOTE de CALCUL §2.5.3  : Calcul de Cp
  Dim QPiéton As Integer
  QPiéton = QP(NumBranche)
  Cp(NumBranche) = wBranche.TTP / 10 * (1 - Exp(-QPiéton / 360)) * (1800 + QPiéton) / 2160 * Exp(exposant)
  ' Ref NOTE de CALCUL §2.5.4  : Calcul de C et RC
  C(NumBranche) = Cvh * (1 - Cp(NumBranche))
  rc(NumBranche) = C(NumBranche) - QEntrant(NumBranche)
  '   RCpourcent = RC(numBranche) / C(numbranche)      en pourcentage

End Sub

'**********************************************************************************************************
' Fonction caduque, transférée dans BRANCHE.CLS, mais conservée si on redemandait un jour plusieurs courbes (par milieu)
'**********************************************************************************************************
''Private Function getCVH(ByVal NumBranche As Integer, ByVal wBranche As BRANCHE, ByRef exposant As Double, Optional ByVal Milieu) As Double
''
''    Dim Ci As Double    ' Capacité initiale
''    Dim qgSeconde As Single
''    Dim Tf As Single
''    Dim Tg As Single
''    Dim Te As Single
''
''    If IsMissing(Milieu) Then
''      Te = gbProjetActif.Te
''      Tg = gbProjetActif.Tg
''      Tf = wBranche.Tf
''    Else
''      Te = gbTe(Milieu)
''      Tg = gbTg(Milieu)
''      Tf = wBranche.getTf(gbTf1(Milieu))
''    End If
''    qgSeconde = QG(NumBranche) / 3600
''
''    ' Ref NOTE de CALCUL §2.5.1  : Calcul de Ci
''    exposant = -qgSeconde * (Tg - Tf / 2)
''    Ci = 3600 / Tf * Exp(exposant)
''
''    ' Ref NOTE de CALCUL §2.5.2  : Calcul de Cvh
''    getCVH = Ci * (wBranche.LEGirabase / 3.5) ^ Te
''
''End Function

'******************************************************************************
' Calcul du temps moyen d'attente : TMA
'******************************************************************************
Private Sub TempsAttente(ByVal NumBranche As Integer, wBranche As BRANCHE)
    Dim CSA As Double
    Dim CSB As Double
    Dim SQ As Double    ' Trafic entrant par seconde
    
    ' Ref NOTE de CALCUL §2.6  : Calcul de TMA
    CSA = QEntrant(NumBranche) / C(NumBranche)
    SQ = QEntrant(NumBranche) / 3600
    
    If QEntrant(NumBranche) = 0 Then
    ' Pas de trafic entrant (ne devrait pas arriver, mais....)
      TMA(NumBranche) = 0
      TTA(NumBranche) = 0
      Exit Sub
    End If
    
    Dim exposant As Double
    exposant = -SQ * wBranche.Tf
    If CSA < 0.9 Then
      TMA(NumBranche) = (Exp(exposant) / (1 - CSA) - 1) / SQ
    ElseIf CSA > 1.1 Then
      TMA(NumBranche) = 1800 * (CSA - 1)
    Else
    ' Ref : formule corrigée par B.GUICHET - E-MAIL DU 25/03/99
      CSB = (10 * Exp(exposant) - 1) / SQ
      TMA(NumBranche) = CSB + 5 * (180 - CSB) * (CSA - 0.9)
    End If
    ' Ref : Plan d'action suite à sites pilotes(09/99) - p4
    TMA(NumBranche) = Max(0#, TMA(NumBranche))
' Ref : Modificatif 29/12/98
' Calcul du temps total d'attente : TMA
    TTA(NumBranche) = TMA(NumBranche) * QEntrant(NumBranche)
    
End Sub
    
'******************************************************************************
' Calcul des longueurs moyenne et maxi de stockage : LK et LKM
'******************************************************************************
Private Sub LongueurStockage(ByVal NumBranche As Integer, wBranche As BRANCHE)
    
    ' Ref NOTE de CALCUL §2.7.1  : Calcul de LK
    LK(NumBranche) = TMA(NumBranche) * Min(QEntrant(NumBranche), C(NumBranche)) / 3600
    ' Ref NOTE de CALCUL §2.7.1  : Calcul de LKM
    If QEntrant(NumBranche) < C(NumBranche) Then
      LKM(NumBranche) = 2 + 3 * LK(NumBranche)
    Else
    ' Ref : formule corrigée par B.GUICHET - E-MAIL DU 25/03/99
      LKM(NumBranche) = 2 + (3 + 2 * rc(NumBranche) / QEntrant(NumBranche)) * LK(NumBranche)
    End If
    
End Sub

Public Function getQE(ByVal NumBranche As Integer) As Integer         ' Trafic entrant
  getQE = QE(NumBranche)
End Function
Public Function getQS(ByVal NumBranche As Integer) As Integer         ' Trafic sortant
  getQS = QS(NumBranche)
End Function
Public Function getQTournant(ByVal NumBranche As Integer) As Integer  ' Trafic tournant (utile pour certains messages)
  getQTournant = QTournant(NumBranche)
End Function
Public Function getQG(ByVal NumBranche As Integer) As Integer        ' Trafic gênant
  getQG = QG(NumBranche)
End Function
Public Function getC(ByVal NumBranche As Integer) As Double          ' Capacité définitive
  getC = C(NumBranche)
End Function
Public Function getCp(ByVal NumBranche As Integer) As Double         ' Perte de capacité due aux
  getCp = Cp(NumBranche)
End Function
Public Function getRC(ByVal NumBranche As Integer) As Double       ' Réserve de capacité (en uvp/h)
  getRC = rc(NumBranche)
End Function
Public Function getTMA(ByVal NumBranche As Integer) As Double        ' Temps moyen d'attente
  getTMA = TMA(NumBranche)
End Function
Public Function getTTA(ByVal NumBranche As Integer) As Double        ' Temps total d'attente
  getTTA = TTA(NumBranche)
End Function
Public Function getLK(ByVal NumBranche As Integer) As Single         ' Longueur moyenne de stockage
  getLK = LK(NumBranche)
End Function
Public Function getLKM(ByVal NumBranche As Integer) As Single        ' Longueur totale de stockage
  getLKM = LKM(NumBranche)
End Function

Private Function Impaire(Valeur As Integer) As Boolean
  If CInt(Valeur / 2) <> Valeur / 2 Then Impaire = True
End Function
