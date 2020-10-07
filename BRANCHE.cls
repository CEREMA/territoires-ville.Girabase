VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BRANCHE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'******************************************************************************
'*
'*          Projet GIRABASE - CERTU - CETE de l'Ouest
'*
'*          Module de classe : BRANCHE.BAS
'*
'*          Définition de l'objet BRANCHE
'*
'******************************************************************************

Option Explicit

Public nom As String              ' Nom de la branche
Public Angle As Integer            ' Angle en degrés
Public Rampe As Boolean           ' Présence d'une rampe > 3%
Public EntréeEvasée As Boolean
Public TAD As Boolean             ' Tourne à droite

Public EntréeNulle As Boolean     'Branche de sortie uniquement
Public SortieNulle As Boolean     'Branche de sortie uniquement

Public LE4m As Single             ' Largeur d'entrée à 4m
Public LE15m  As Single           ' Largeur d'entrée à 15m
'Public LE as Single              ' Largeur d'entrée : dépend de TAD (NOTE DE CALCUL §1)--> calculé
Public LI As Single               ' Largeur de l'îlot
Public LS As Single               ' Largeur de sortie

Public Ecart As Single            ' A calculer ???
' Variables nécessaires au calcul de capacité
Public Tf As Single          ' Coefficient créneau complémentaire retenu
Public LEGirabase As Single  ' Largeur d'entrée retenue par GIRABASE
Public KS As Single          ' Coefficient de gêne trafic sortant
Public TTP As Single         ' Temps de traversée Piétons

'******************************************************************************
' Affichage des données d'une branche
'******************************************************************************
Public Sub Affiche(ByVal FeuilleDonnées As Form, ByVal NumBranche As Integer)
Dim i As Integer

  With FeuilleDonnées
    ' Afficher les caractéristiques des branches
    With .vgdCarBranche
      .Row = NumBranche
      .Col = 1
      .Value = nom
      .Col = 2
      .Value = CStr(Angle)
      If NumBranche = 1 Then .Lock = True
      ' Dans cette version , la branche numéro 1 est non modifiable(angle=0)
      .Col = 3
      If NumBranche = 1 Then
        .Lock = True
      Else
        .Value = CStr(Ecart)
      End If
      .Col = 4
      .Value = Rampe
      .Col = 5
      .Value = TAD
    End With
    
    ' Afficher les largeurs des branches

    With .vgdLargBranche
      .Row = NumBranche
      .Col = 1
      .Value = RemplacePoint(CStr(LE4m))
      .Col = 2
      .Value = RemplacePoint(CStr(LE15m))
      .Col = 3
      .Value = RemplacePoint(CStr(LI))
      .Col = 4
      .Value = RemplacePoint(CStr(LS))
      .Col = 5
      .Value = EntréeEvasée
      'Interdit la saisie de LI, EV et TAD si l'entrée LI est nulle
      If EntréeNulle Then gbProjetActif.Données.ChangeLE4m .Row, EntréeNulle
      If SortieNulle Then gbProjetActif.Données.ChangeLS .Row, SortieNulle
    End With
    
    ' Initialiser les trafics
''    With .vgdTrafic(VEHICULE)
''      .Row = NumBranche
''      For i = 1 To .MaxRows
''        .Col = i
''        .Value = ""
''      Next
''    End With
''    With .vgdTrafic(PIETON)
''      .Col = NumBranche
''      .Value = ""
''    End With
  End With
  
End Sub


' MODIF LCH 25 06 2020 : il faut rétablir le . comme séparateur décimal pour l'affichage
' sinon, à l'ouverture d'un projet, elles ne sont pas affichées dans le tableau des branches

Private Function RemplacePoint(ByVal Chaine As String) As String

  Dim Position As Long
  Position = InStr(1, Chaine, ",", 1)
  If Position Then
    Mid(Chaine, Position) = Chr(46)  ' = "."
  End If
  RemplacePoint = Chaine
End Function

'******************************************************************************
' Lecture des données d'une branche
'******************************************************************************
Public Function Lire(ByVal numFich As Integer, ByRef numLigne As Integer, ByVal modeangle As Integer) As Boolean

Dim variable As Variant
Dim Chaine As String

On Error GoTo TraitementErreur

  Input #numFich, nom, variable
  numLigne = numLigne + 1
  OkEntier variable, Angle, 0, 2 * eqvPI(modeangle) - 1
  
  Input #numFich, variable
  If VarType(variable) <> vbBoolean Then Err.Raise 100
  Rampe = variable
  
  Input #numFich, variable
  If VarType(variable) <> vbBoolean Then Err.Raise 100
  TAD = variable
  
  Input #numFich, variable
  numLigne = numLigne + 1
  If VarType(variable) <> vbBoolean Then Err.Raise 100
  EntréeEvasée = variable
  
  Input #numFich, Chaine
  OkFlottant Chaine, LE4m
  Input #numFich, Chaine
  OkFlottant Chaine, LE15m
  Input #numFich, Chaine
  OkFlottant Chaine, LI
  Input #numFich, Chaine
  OkFlottant Chaine, LS
  
  EntréeNulle = (LE4m = 0#)
  SortieNulle = (LS = 0#)

  Lire = True
  Exit Function
  
TraitementErreur:
  If Err <> 100 Then ErreurFatale
End Function

'******************************************************************************
' Paramètres de la branche : Tf, LEgirabase, KS et TTP
'******************************************************************************
Public Sub CalculParamBranche()

  With gbProjetActif
    Tf = getTf(.Tf1)
''      ' Ref §1.3
''    If Rampe Then
''      Tf = .Tf1 * 1.35
''    Else
''      Tf = .Tf1
''    End If

    ' Ref  §1.2
    Dim LE As Single          ' Largeur d'entrée
    If EntréeEvasée Then
      LE = (LE4m + LE15m) / 2
    Else
      LE = LE4m
    End If
    ' Ref §2.1.2
    LEGirabase = Min(.LEU, LE)
''    If gbFichierJournal Then
''      If LEGirabase = LE Then
''        Write #gbFichLog, "2.1.2 - LEGirabase = LE"
''      Else
''        Write #gbFichLog, "2.1.2 - LEGirabase = LEU"
''      End If
''    End If

    ' Détermination de coefficient de gêne KS (Trafic sortant au droit de l'entrée)
    ' Ref  §2.1.3
    Dim SLB As Single         ' Surlargeur de sortie
    SLB = Max(LS - 5#, 0#)
''    If gbFichierJournal Then
''      If SLB = 0# Then
''        Write #gbFichLog, "2.1.3 - SLB = 0"
''      Else
''        Write #gbFichLog, "2.1.3 - SLB = LS"
''      End If
''    End If
    
      ' Ref  §2.2.1
    KS = .RU / (.RU + .LAU) - (LI + 0.5 * SLB) / .LImax
      ' Ajout - suite aux observations faites au CERTU - AV - 08.03.99 (Répondu le 12.03.99)
    KS = Max(0, KS)
''    If gbFichierJournal Then
''      If KS = 0 Then
''        Write #gbFichLog, "2.2.1 - KS = 0"
''      Else
''        Write #gbFichLog, "2.2.1 - KS calculé"
''      End If
''    End If
    

    ' Temps de traversée piéton  (vitesse supposée = 1m/s)
    ' Ref §2.5.3
    ' Faire vérifier par le CERTU : S'agit-il bien de LE si entrée évasée (Répondu le 12.03.99)
'    TTP = Min(LE, 4)
    TTP = Min(LE4m, 4)

  End With

End Sub

Public Function getTf(ByVal Tf1 As Single) As Single
      ' Ref §1.3
    If Rampe Then
      getTf = Tf1 * 1.35
    Else
      getTf = Tf1
    End If
End Function

Public Function getCVH(ByVal QG As Integer, Optional ByRef exposant As Double) As Double
  
    Dim Ci As Double    ' Capacité initiale
    Dim qgSeconde As Single
    Dim Tg As Single
    Dim Te As Single
    
    Te = gbProjetActif.Te
    Tg = gbProjetActif.Tg
    qgSeconde = QG / 3600
    
    ' Ref NOTE de CALCUL §2.5.1  : Calcul de Ci
    exposant = -qgSeconde * (Tg - Tf / 2)
    Ci = 3600 / Tf * Exp(exposant)
    
    ' Ref NOTE de CALCUL §2.5.2  : Calcul de Cvh
    getCVH = Ci * (LEGirabase / 3.5) ^ Te
  
End Function

