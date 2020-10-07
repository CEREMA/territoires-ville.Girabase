Attribute VB_Name = "DessinGiratoire"
Option Explicit

Private Const IDl_Egal = " = "

Public Const maLongueurBranche As Single = 18 'Longueur réelle en mètres des branches pour le dessin (utilisé aussi par frmImprimer)
Private Const uneMarge As Single = 50 'Pour laisser de la place autour du dessin

Private Const monEpsilon = 100 'Pour tester les valeurs proches de zéro

' Constantes permettant de connaitre le type d'objet qui a été sélectionné
Private Const NoObjSelect = 0
Private Const ObjPoignee = 1
Private Const ObjBranche = 2
Private Const ObjAnneauInt = 3
Private Const ObjAnneauMil = 4
Private Const ObjAnneauExt = 5

Public monNumBrancheSelect As Integer 'Numéro de la branche sélectionnée, 0 sinon
Public gbRayonInt As Single
Public gbRayonExt As Single
Public gbBandeFranchissable As Single
Public gbDemiLargeur As Single
Public gbDemiHauteur As Single
Public gbFacteurZoom As Single

Private monAngle As String    ' Ajout - AV (05.02.99) : pour mise à jour du tableau

Private maLargeurAnneau As Single 'RVG : artifice pour maintenir l'anneau constant si modification interactive du rayon intérieur

'Drapeaux pour détecter si l'utilisateur a déjà cliqué et s'il est en train de faire glisser la souris
Private DebutClick As Boolean
Private Glisser As Boolean
 
Private monObjetSelect As Integer 'Objet graphique sélectionné


'***************************************************************************************
' MouseDown  : emprunté à la maquette du CERTU (GIRABASE.FRM)
'***************************************************************************************
Public Sub Dessin_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
Dim exNumBrancheSelect As Integer
    
    DebutClick = True
   
    'Déselection de la sélection précédente : indispensable pour que les MouseMove soient ignorées jusqu'à la fin de la procédure
    monObjetSelect = NoObjSelect
    exNumBrancheSelect = monNumBrancheSelect
     monNumBrancheSelect = 0
    
    'Test si sélection interactive d'un poignée
      With gbProjetActif.Données.shpPoignée
        If .Visible And Distance(X, .Left + .Width / 2, Y, .Top + .Height / 2) < monEpsilon Then
          'Le centre de la poignée et le pointeur souris sont proches
          'et si la poignée est sélectionnable ==> Sélection de la poignée
          If .Tag <> "" Then
            monObjetSelect = CInt(.Tag)
            If monObjetSelect = ObjBranche Then monNumBrancheSelect = exNumBrancheSelect
          Else
            monObjetSelect = ObjPoignee
          End If

          Exit Sub
          
        Else
        End If
      End With
    
    'Test si sélection interactive d'un branche (équation de droite : ux + vy +w = 0)
      Dim u As Single 'Coefficients d'une équation
      Dim v As Single 'de droite du type
      Dim w As Single 'uX + vY + w = 0
      Dim uneDistance As Single   ' Distance en Twips
      
      For i = 1 To gbProjetActif.NbBranches    ' On n'autorise pas de modifier la branche 1 (AV - 05.02.99)
        'Test si le pointeur souris est dans la boite englobante du segment de branche
        With gbProjetActif.Données.linBranche(i)
          If X > (Min(.X1, .X2) - monEpsilon) And X < (Max(.X1, .X2) + monEpsilon) And _
             Y > (Min(.Y1, .Y2) - monEpsilon) And Y < (Max(.Y1, .Y2) + monEpsilon) Then
            'Calcul du vecteur directeur de la branche i
            u = .Y1 - .Y2
            v = .X2 - .X1
            w = -u * .X1 - v * .Y1 'w = -ux-vy
              'Calcul de la distance entre le pointeur souris et la branche i
            uneDistance = Abs(u * X + v * Y + w) / Sqr(u * u + v * v)
            If (uneDistance < monEpsilon) Then
              'Cas où le click est proche de la branche i ==> Sélection de celle-ci
              SelectBranche i
              Exit Sub
            End If
          End If
        End With
      Next i
    
    'Test si sélection interactive d'un Cercle
      If SelectCercle(0, X, Y) Then Exit Sub
      
      ' Echec de la recherche
    With gbProjetActif.Données.shpPoignée
      .Tag = ""
      .Visible = False
    End With
    monObjetSelect = NoObjSelect
    monNumBrancheSelect = 0
    DebutClick = False
      
End Sub

Private Function SelectCercle(ByVal NumAnneau As Integer, ByVal X As Single, ByVal Y As Single) As Boolean
Dim unRayon(ObjAnneauInt To ObjAnneauExt) As Single   ' En twips
Dim nom As String

  unRayon(ObjAnneauInt) = gbRayonInt * gbFacteurZoom
  unRayon(ObjAnneauMil) = (gbRayonInt + gbBandeFranchissable) * gbFacteurZoom
  unRayon(ObjAnneauExt) = gbRayonExt * gbFacteurZoom

  'Calcul de la distance entre le point écran (X, Y) et le centre des cercles des anneaux = centre de la vue
    Dim DistanceAuCentreVue As Single  ' Distance en Twips
    Dim DistMin As Single
    Dim DistAnneau
    DistanceAuCentreVue = Distance(gbDemiLargeur, X, gbDemiHauteur, Y)
    DistMin = 2 * monEpsilon
  'Test si sélection interactive de l'anneau
    For NumAnneau = ObjAnneauInt To ObjAnneauExt
      DistAnneau = Abs(unRayon(NumAnneau) - DistanceAuCentreVue)
      If DistAnneau < monEpsilon Then
        If DistAnneau < DistMin Then
    '==> Sélection de l'anneau correspondant
          monObjetSelect = NumAnneau
          DistMin = DistAnneau
        End If
      End If
    Next
    
  If monObjetSelect = NoObjSelect Then Exit Function
  
  Select Case (monObjetSelect)
  Case ObjAnneauInt
    nom = "txtR"
  Case ObjAnneauMil
    nom = "txtBf"
  Case ObjAnneauExt
    nom = "txtLA"
  End Select

  'Cas où la distance entre le pointeur et le centre est proche du rayon
  '0699
  'Positionne le focus sur un contrôle quelconque
  gbProjetActif.Données.txtVariante.SetFocus
  'Repositionne le contrôle associé au cercle en cours
  gbProjetActif.Données.Controls(nom).SetFocus
  ' L'Instruction SetFocus génère un évènement GotFocus, qui appelle InitControle et rend la
  ' poignée invisible --> on la rend visible ci-dessous
  DoEvents
  
    'Stockage du type d'élément dont elle est la poignée
  gbProjetActif.Données.shpPoignée.Tag = CStr(monObjetSelect)
  With gbProjetActif.Données.shpPoignée
  'Coordonnées de la projection sur un cercle de la poignée de ce cercle
    Dim XProj As Single
    Dim YProj As Single
    'Affichage de la poignée sur l'anneau prés du pointeur souris
    XProj = X + (unRayon(monObjetSelect) - DistanceAuCentreVue) * Sgn(X - gbDemiLargeur)
    YProj = Y + (unRayon(monObjetSelect) - DistanceAuCentreVue) * Sgn(Y - gbDemiHauteur)
    .Left = XProj - .Width / 2
    .Top = YProj - .Height / 2
    .Visible = True
  End With
  
  maLargeurAnneau = gbRayonExt - gbRayonInt - gbBandeFranchissable
  
  SelectCercle = True

End Function

Private Function SelectCercleOK(ByVal NumAnneau As Integer, ByVal X As Single, ByVal Y As Single) As Boolean
Dim unRayon As Single   ' En twips
Dim nom As String


  Select Case NumAnneau
  Case ObjAnneauInt
    unRayon = gbRayonInt * gbFacteurZoom
    nom = "txtR"
  Case ObjAnneauMil
    unRayon = (gbRayonInt + gbBandeFranchissable) * gbFacteurZoom
    nom = "txtBf"
  Case ObjAnneauExt
    unRayon = gbRayonExt * gbFacteurZoom
    nom = "txtLA"
  End Select

  'Calcul de la distance entre le point écran (X, Y) et le centre des cercles des anneaux = centre de la vue
    Dim DistanceAuCentreVue As Single  ' Distance en Twips
    DistanceAuCentreVue = Distance(gbDemiLargeur, X, gbDemiHauteur, Y)
  
  'Test si sélection interactive de l'anneau
  If Abs(DistanceAuCentreVue - unRayon) < monEpsilon Then
    'Cas où la distance entre le pointeur et le centre est proche du rayon
    gbProjetActif.Données.Controls(nom).SetFocus
    ' L'Instruction SetFocus génère un évènement GotFocus, qui appelle InitControle et rend la
    ' poignée invisible --> on la rend visible plus loin
    DoEvents
    
    '==> Sélection de l'anneau correspondant
    monObjetSelect = NumAnneau
      'Stockage du type d'élément dont elle est la poignée
    gbProjetActif.Données.shpPoignée.Tag = CStr(NumAnneau)
    With gbProjetActif.Données.shpPoignée
    'Coordonnées de la projection sur un cercle de la poignée de ce cercle
      Dim XProj As Single
      Dim YProj As Single
      'Affichage de la poignée sur l'anneau prés du pointeur souris
      XProj = X + (unRayon - DistanceAuCentreVue) * Sgn(X - gbDemiLargeur)
      YProj = Y + (unRayon - DistanceAuCentreVue) * Sgn(Y - gbDemiHauteur)
      .Left = XProj - .Width / 2
      .Top = YProj - .Height / 2
      .Visible = True
    End With
    
    maLargeurAnneau = gbRayonExt - gbRayonInt - gbBandeFranchissable
    
    SelectCercleOK = True
  End If

End Function

Public Sub SelectBranche(num As Integer)
            
  With gbProjetActif.Données
    'Mise en surbrillance de la colonne de l'angle
    .BrancheSélectée = num
    
    With .vgdCarBranche
      .Row = num
      .Col = 2
      monAngle = .Value
      .Action = 0
      .SetFocus
      'L'événement GotFocus n'est plus déclenché si le focus est déjà sur le spread
      'L'appel de la procédure suivante replace le focus sur la bonne cellule puis
      'affiche l'invite correspondant à la branche sélectée
      gbProjetActif.Données.vgdDéplaceFocus num, 2
  ' L'Instruction SetFocus génère un évènement GotFocus, qui appelle AfficheSpreadNormal et rend la
  ' poignée invisible --> on la rend visible plus loin
      DoEvents
    End With

    .BrancheSélectée = 0
    'Sélection d'une branche par sélection de son numéro
    monNumBrancheSelect = num
    monObjetSelect = ObjBranche
    'Stockage du type d'élément dont elle est la poignée
    .shpPoignée.Tag = CStr(ObjBranche)
    'Affichage de la poignée à l'extrémité de la branche sélectionnée
    .shpPoignée.Left = .linBranche(monNumBrancheSelect).X2 - .shpPoignée.Width / 2
    .shpPoignée.Top = .linBranche(monNumBrancheSelect).Y2 - .shpPoignée.Height / 2
    .shpPoignée.Visible = True
  End With

End Sub

'***************************************************************************************
' MouseMove  : emprunté à la maquette du CERTU (GIRABASE.FRM)
'***************************************************************************************
Public Sub Dessin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim unRayonDyn As Single
  
  If monObjetSelect = NoObjSelect Then Exit Sub
  If monObjetSelect = ObjBranche And monNumBrancheSelect = 1 Then Exit Sub
  
  If DebutClick Then
    DebutClick = False
    Glisser = True
  End If
  
  If Not Glisser Then Exit Sub
  
  Dim lblInvite As Label
  Set lblInvite = gbProjetActif.Données.lblInvite
  
    'Cas du bouton gauche enfoncé traité uniquement
    If Button = 1 Then
      Dim shpAnneauInt As Shape
      Dim shpAnneauExt As Shape
      Dim shpAnneauMil As Shape
      With gbProjetActif.Données
        Set shpAnneauInt = .shpAnneauInt
        Set shpAnneauExt = .shpAnneauExt
        Set shpAnneauMil = .shpAnneauMil
      End With
      If monObjetSelect = NoObjSelect Then
          'Rien de sélectionné
          
      ElseIf monObjetSelect = ObjBranche Then
        'Modification dynamique de la branche
        ModifDynamicBranche X, Y
        
      ElseIf monObjetSelect = ObjAnneauInt Then
        'Modification du rayon intérieur : LA et Bf restent constantes--->
        '    Redimensionnement en conséquence des 3 anneaux
        'Calcul du nouveau rayon
        unRayonDyn = Distance(gbDemiLargeur, X, gbDemiHauteur, Y)
       'Déplacement de la poignée de sélection
       PoignéeMove X, Y
        'Modification de l'anneau intérieur
        DessinerAnneau shpAnneauInt, unRayonDyn
        gbRayonInt = trEchel(unRayonDyn, True)
        'Affichage dynamique de la valeur du rayon
        lblInvite.Caption = IDl_RayonIntérieur & IDl_Egal + Format(gbRayonInt, "0.0")
        
        ' Maintien de l'anneau constant
        gbRayonExt = gbRayonInt + maLargeurAnneau + gbBandeFranchissable
        unRayonDyn = trEchel(gbRayonExt, False)
        DessinerAnneau shpAnneauExt, unRayonDyn
        
        ' Maintien de la Bande franchissable constante
        unRayonDyn = trEchel(gbRayonInt + gbBandeFranchissable, False)
        DessinerAnneau shpAnneauMil, unRayonDyn
      
      ElseIf monObjetSelect = ObjAnneauExt Then
        'Modification du rayon extérieur : lui seul est modifié (LA est recalculé en conséquence)
        'Calcul du nouveau rayon
        unRayonDyn = Distance(gbDemiLargeur, X, gbDemiHauteur, Y)
        If unRayonDyn > trEchel(gbRayonInt + gbBandeFranchissable, False) Then
         'Diminution possible jusqu'au rayon intérieur
         'Déplacement de la poignée de sélection
         PoignéeMove X, Y
         'Modification de l'anneau intérieur
         DessinerAnneau shpAnneauExt, unRayonDyn
         gbRayonExt = trEchel(unRayonDyn, True)
         'Affichage dynamique de la valeur du rayon
         lblInvite.Caption = _
              IDl_LargeurAnneau & IDl_Egal & Format(gbRayonExt - gbRayonInt - gbBandeFranchissable, "0.0") & _
              "  (" & IDl_RayonExtérieur & IDl_Egal & Format(gbRayonExt, "0.0") & ")"
        Else
         'Cas où le rayon intérieur devient > au rayon extérieur
          lblInvite.Caption = IDm_LargeurAnneauNonNulle
        End If
      
      ElseIf monObjetSelect = ObjAnneauMil Then
        'Modification du rayon intermédiaire LA reste constante--->
        '    Redimensionnement en conséquence des 2 anneaux
        'Calcul du nouveau rayon
        unRayonDyn = Distance(gbDemiLargeur, X, gbDemiHauteur, Y)
        If unRayonDyn < trEchel(gbRayonInt, False) Then
          lblInvite.Caption = IDm_LargeurBandePositive
        ElseIf unRayonDyn >= trEchel(gbRayonExt, False) Then
          lblInvite.Caption = IDm_LargeurAnneauNonNulle
        Else
          PoignéeMove X, Y
         'Modification de l'anneau intermédiaire
         DessinerAnneau shpAnneauMil, unRayonDyn
         'Calcul de la nouvelle bande franchissable
         Dim unRayonMil As Single
         unRayonMil = trEchel(unRayonDyn, True)
         gbBandeFranchissable = unRayonMil - gbRayonInt
         'Modification de l'anneau extérieur
         gbRayonExt = gbRayonInt + gbBandeFranchissable + maLargeurAnneau
         unRayonDyn = trEchel(gbRayonExt, False)
         DessinerAnneau shpAnneauExt, unRayonDyn
         'Affichage dynamique de la valeur du rayon
         lblInvite.Caption = _
                      IDl_BandeFranchissable & IDl_Egal & Format(gbBandeFranchissable, "0.0")
        End If
      End If
      
    Else
      lblInvite.Caption = ""
    End If
    
End Sub

'***************************************************************************************
' MouseUp  : emprunté à la maquette du CERTU (GIRABASE.FRM)
'***************************************************************************************
Public Sub Dessin_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  'Stockage pour la modification interactive suivante
  
  DebutClick = False
                                                                          
  If Not Glisser Then Exit Sub
  
  With gbProjetActif.Données
    Dim unNbBranches As Integer
    unNbBranches = gbProjetActif.NbBranches
    'Mémorise la modification de donnée
    .DonnéeModifiée = True
    'RVG : anneau constant
    If (monObjetSelect = ObjAnneauExt Or _
       monObjetSelect = ObjAnneauMil Or _
       monObjetSelect = ObjAnneauInt) Then
     
      'Affichage de la poignée de sélection de l'anneau extérieur
      .shpPoignée.Visible = True
      .txtR = Format(gbRayonInt, "0.0")
      .txtLA = Format(gbRayonExt - gbRayonInt - gbBandeFranchissable, "0.0")
      .txtBf = Format(gbBandeFranchissable, "0.0")
      .ValidateObjet = True 'Valide l'objet
      Select Case monObjetSelect
        Case ObjAnneauInt
        .txtR_Validate False
        Case ObjAnneauMil
        .txtBf_Validate False
        Case ObjAnneauExt
        .txtLA_Validate False
      End Select
      .ValidateObjet = False
    ElseIf monObjetSelect = ObjBranche Then
      With .vgdCarBranche
        .Col = 2
        Dim Ecart, Valeur  As Integer
        Ecart = .Value - monAngle
        .Value = monAngle
        .Col = 3
        If .Value = "" Then
          .Value = -Ecart
        Else
          .Value = CInt(.Value) - Ecart
        End If
        If monNumBrancheSelect < unNbBranches Then
          .Row = monNumBrancheSelect + 1
          .Col = 2
          Valeur = .Value
          .Col = 3
          .Value = Valeur - monAngle
        End If
      End With
   
      'Valider l'angle et transférer la donnée
      .NuméroLigneActive = monNumBrancheSelect
      .vgdCarBranche_LeaveCell 2, monNumBrancheSelect, 2, monNumBrancheSelect, False
      '0699
      .vgdDéplaceFocus monNumBrancheSelect, 2
    End If
  End With
  
  Glisser = False

End Sub

'***************************************************************************************
' DessinerTout  : emprunté à la maquette du CERTU (GIRABASE.FRM)
'***************************************************************************************
Public Sub DessinerTout(ByVal IsPremierDessin As Boolean)
  Dim i As Integer
  Dim unMax As Single
  Dim unAngleBranche As Single
  Dim unCos As Single
  Dim unSin As Single
  Dim unXi As Single
  Dim unYi As Single
  Dim xDébutBranche As Single
  Dim yDébutBranche As Single
  Dim xFinBranche As Single
  Dim yFinBranche As Single
  Dim monFacteurZoomPrecedent As Single 'Facteur de zoom précédent
  Dim Poignée As Shape

  'Détermination du facteur de Zoom pour un cadrage maximun
  If gbFacteurZoom = 0 Then    ' Modif AV : 08.02.99 --> on ne rezoome pas à chaque redessin du giratoire (gérer quand même le MDI)
    If (gbDemiLargeur < gbDemiHauteur) Then
      gbFacteurZoom = (gbDemiLargeur - uneMarge * 5) / (gbRayonExt + maLongueurBranche)
    Else
      gbFacteurZoom = (gbDemiHauteur - uneMarge * 5) / (gbRayonExt + maLongueurBranche)
    End If
  End If
  
  'Correction si le facteur de zoom est négatif (cas aprés une mise en icone)
  If gbFacteurZoom <= 0 Then
      gbFacteurZoom = 1
  End If
  
  With gbProjetActif.Données
    .FacteurZoom = gbFacteurZoom
    Set Poignée = .shpPoignée
  
    'Positionnement dans le nouveau niveau de zoom
    If IsPremierDessin Then
      'Dimensionnement du tableau de branches lors du chargement
      'ou de la création d'un giratoire, donc lors du premier dessin
    Else
      'Cas du chargement d'un giratoire existant
      'ou du redessin à un nouveau niveau de zoom
      Dim Xc As Single
      Dim Yc As Single
      Xc = .AncienXc
      Yc = .AncienYc
      monFacteurZoomPrecedent = gbProjetActif.FacteurZoomPrecedent
      With Poignée
        .Left = gbDemiLargeur + (.Left - Xc) * (gbFacteurZoom / monFacteurZoomPrecedent)
        .Top = gbDemiHauteur + (.Top - Yc) * (gbFacteurZoom / monFacteurZoomPrecedent)
      End With
    End If
    
    'Dessin des anneaux
    ' Anneau intérieur
    Dim shpAnneau As Shape
    Set shpAnneau = .shpAnneauInt
    DessinerAnneau shpAnneau, gbFacteurZoom * gbRayonInt
    With shpAnneau
      .Visible = True
      ControlePoignée ObjAnneauInt, .Height / 2, Poignée
    End With
    
    ' Anneau intermédiaire
    Set shpAnneau = .shpAnneauMil
    With shpAnneau
      If gbBandeFranchissable = 0 Then
        .Visible = False
      Else
        DessinerAnneau shpAnneau, gbFacteurZoom * (gbRayonInt + gbBandeFranchissable)
        .Visible = True
        ControlePoignée ObjAnneauMil, .Height / 2, Poignée
      End If
    End With
    
    ' Anneau extérieur
    Set shpAnneau = .shpAnneauExt
    DessinerAnneau shpAnneau, gbFacteurZoom * gbRayonExt
    With shpAnneau
      .Visible = True
      ControlePoignée ObjAnneauExt, .Height / 2, Poignée
    End With
    
    'Dessin des branches avec leur numéro
    For i = 1 To gbProjetActif.NbBranches
      unAngleBranche = angConv(gbProjetActif.colBranches.Item(i).Angle, CVRADIAN)
      If IsPremierDessin Then
        'Création en mémoire des instances des tableaux de controles
        Load .linBranche(i)
        Load .linBordIlotEntrée(i)
        Load .linBordIlotSortie(i)
        Load .linBordIlotGir(i)
        Load .linBordVoieEntrée(i)
        Load .linVoieSortie(i)
        Load .linVoieEntrée(i)
        Load .linBordVoieSortie(i)
        Load .lblLibelléBranche(i)
        Load .lblNumBranche(i)
''        'On stocke l'angle de la branche créée dans
''        'son champ Tag  pour utilisation ultérieure
        .linBranche(i).Tag = CStr(unAngleBranche)
      End If
      
      unCos = Cos(unAngleBranche)
      unSin = -Sin(unAngleBranche)      ' "-" : car l'axe des Y est vers le bas
      'Creation de l'axe de la branche
      With .linBranche(i)
        TrRot gbRayonExt, 0, unXi, unYi, unCos, unSin
        .X1 = unXi
        .Y1 = unYi
        xDébutBranche = .X1
        yDébutBranche = .Y1
        TrRot gbRayonExt + maLongueurBranche, 0, unXi, unYi, unCos, unSin
        .X2 = unXi
        .Y2 = unYi
        xFinBranche = .X2
        yFinBranche = .Y2
        .Visible = True
      End With
      
      'Création du reste de la branche (Voies entrante, sortante et ilot)
      DessinerBranche i, unAngleBranche
      
      'Positionnement des noms de branches
      With .lblLibelléBranche(i)
'''        .Caption = Left(IDl_Branche, 1) & CStr(i)
       .Caption = "  " & gbProjetActif.colBranches.Item(i).nom & "  "
''        unMax = Max(.Width, .Height)+ uneMarge
        DecalXY xFinBranche, yFinBranche, unCos, unSin, .Caption
        .Left = xFinBranche ''+ (unMax * unCos - .Width) / 2
        .Top = yFinBranche ''+ (unMax * unSin - .Height) / 2
        .Visible = True
      End With
      
      With .lblNumBranche(i)
        .Caption = CStr(i)
        unMax = Max(.Width, .Height) + 3 * uneMarge     ' Coef 3 : Car on a substitué n à Bn
        .Left = xDébutBranche + (-unMax * unCos - .Width) / 2
        .Top = yDébutBranche + (-unMax * unSin - .Height) / 2
        .Visible = True
      End With
      
    Next i
    
  End With    ' gbProjetActif.Données
    
    'Stockage ancien facteur de zoom
  gbProjetActif.FacteurZoomPrecedent = gbFacteurZoom
  
End Sub

Private Sub ControlePoignée(ByVal NumAnneau As Integer, ByVal unRayon As Single, ByVal Poignée As Shape)
  If monObjetSelect = NumAnneau And Poignée.Visible Then
    Dim exRayon As Single
    With Poignée
      exRayon = Distance(.Left + .Width / 2, gbDemiLargeur, .Top + .Height / 2, gbDemiHauteur)
      If Abs(unRayon - exRayon) > monEpsilon Then
      ' La modification graphique a été refusée : on replace la poignée sur l'anneau
        .Left = gbDemiLargeur - .Width / 2 + (.Left + .Width / 2 - gbDemiLargeur) * (unRayon / exRayon)
        .Top = gbDemiHauteur - .Height / 2 + (.Top + .Height / 2 - gbDemiHauteur) * (unRayon / exRayon)
      End If
    End With
  End If
End Sub

'***************************************************************************************
' lblLibelléBranche_DblClick  : emprunté à la maquette du CERTU (GIRABASE.FRM)
'***************************************************************************************
Public Sub LabelBranche_DblClick(Index As Integer)
        'Affichage des caractéristiques de la branche sélectionnée
  frmCarBranche.Show vbModal

End Sub

'***************************************************************************************
' lblLibelléBranche_MouseDown  : emprunté à la maquette du CERTU (GIRABASE.FRM)
'***************************************************************************************
Public Sub LabelBranche_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  DebutClick = True
    
    'Déselection de la sélection précédente : indispensable pour que les MouseMove soient ignorées jusqu'à la fin de la procédure
    monObjetSelect = NoObjSelect
    monNumBrancheSelect = 0
    
  'Sélection d'une branche par sélection de son numéro
  SelectBranche Index
  
End Sub
'***************************************************************************************
' lblLibelléBranche_MouseMove  : emprunté à la maquette du CERTU (GIRABASE.FRM)
'***************************************************************************************
Public Sub LabelBranche_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
  If Index = 1 Then Exit Sub  ' On n'autorise pas de modifier la branche 1 (AV - 05.02.99)
  If monObjetSelect = NoObjSelect Then Exit Sub
  
  If DebutClick Then
    DebutClick = False
    Glisser = True
  End If
  
  If Not Glisser Then Exit Sub
  
  'Cas du bouton gauche enfoncé traité uniquement
  If Button = 1 Then
    With gbProjetActif.Données
      'Translation dans le repère absolu de la Form GiraBase
      X = .lblLibelléBranche(monNumBrancheSelect).Left + X
      Y = .lblLibelléBranche(monNumBrancheSelect).Top + Y
      'Modification dynamique de la branche
      ModifDynamicBranche X, Y
    End With
  End If
  
End Sub

'***************************************************************************************
' DessinerBranche  : emprunté à la maquette du CERTU (GIRABASE.FRM)
'***************************************************************************************
Public Sub DessinerBranche(ByVal unNumBranche As Integer, ByVal unAngle As Single)
  'Procédure dessinnant le reste de la branche :
  'Deux bords de Chaussée + ilôt séparateur
  
  Dim unXLoc As Single
  Dim unYLoc As Single
  Dim unCos As Single
  Dim unSin As Single
  Dim unXi As Single
  Dim unYi As Single
  Dim uneLargeurEntrée As Single
  Dim uneLargeurSortie As Single
  Dim unRayonIlot As Single
  Dim uneLongueurIlot As Single
  
  Dim wLigne As Line
  
  
'Récupération des largeurs d'entrée et de sortie de la branche
  With gbProjetActif.colBranches.Item(unNumBranche)
    uneLargeurEntrée = .LE4m
    uneLargeurSortie = .LS
    unRayonIlot = .LI / 2
    uneLongueurIlot = .LI
    uneLargeurEntrée = .LE4m
    uneLargeurSortie = .LS
    unCos = Cos(unAngle)
    unSin = -Sin(unAngle)      ' "-" : car l'axe des Y est vers le bas
  End With
  
  With gbProjetActif.Données
  
    If uneLargeurEntrée = 0 Or uneLargeurSortie = 0 Then
    ' Entrée ou sortie seulement (pas d'ilot)
      unRayonIlot = 0
      If uneLargeurEntrée = 0 Then
        unXLoc = Sqr(Carré(gbRayonExt) - Carré(uneLargeurSortie))
        unYLoc = uneLargeurSortie
        .linVoieEntrée(unNumBranche).Visible = False
        .linBordVoieEntrée(unNumBranche).Tag = 0
        .linBordVoieSortie(unNumBranche).Tag = Arccos(unXLoc / gbRayonExt)
        Set wLigne = .linVoieSortie(unNumBranche)
      Else
        unXLoc = Sqr(Carré(gbRayonExt) - Carré(uneLargeurEntrée))
        unYLoc = -uneLargeurEntrée
        .linVoieSortie(unNumBranche).Visible = False
        .linBordVoieSortie(unNumBranche).Tag = 0
        .linBordVoieEntrée(unNumBranche).Tag = Arccos(unXLoc / gbRayonExt)
        Set wLigne = .linVoieEntrée(unNumBranche)
      End If
      
      With wLigne
        TrRot unXLoc, unYLoc, unXi, unYi, unCos, unSin
        .X1 = unXi
        .Y1 = unYi
        unXLoc = gbRayonExt + maLongueurBranche
        TrRot unXLoc, unYLoc, unXi, unYi, unCos, unSin
        .X2 = unXi
        .Y2 = unYi
        .Visible = True
      End With
      .linBordVoieEntrée(unNumBranche).Visible = False
      .linBordVoieSortie(unNumBranche).Visible = False
    
    Else
      Set wLigne = .linBordVoieEntrée(unNumBranche)
      With wLigne
        .Visible = True
      'Calcul des coordonnées de l'intersection du bord
      'de la chaussée entrante avec l'anneau extérieur
      'dans le repère local de la branche (origine centre des anneaux)
        unXLoc = Sqr(Carré(gbRayonExt) - Carré(uneLargeurEntrée + unRayonIlot))
        unYLoc = -(uneLargeurEntrée + unRayonIlot)
        .Tag = Arccos(unXLoc / gbRayonExt)
      
      'Calcul des coordonnées absolues (translation centre vers origine (0,0)
      'de la feuile frmDonnées plus une rotation autour du centre des anneaux
      'et stockage dans linBordVoieEntrée
        TrRot unXLoc, unYLoc, unXi, unYi, unCos, unSin
        .X1 = unXi
        .Y1 = unYi
        TrRot gbRayonExt + uneLongueurIlot, -uneLargeurEntrée, unXi, unYi, unCos, unSin
        .X2 = unXi
        .Y2 = unYi
        Set wLigne = gbProjetActif.Données.linVoieEntrée(unNumBranche)
        With wLigne
          .Visible = True
          .X1 = unXi
          .Y1 = unYi
          TrRot gbRayonExt + Max(maLongueurBranche, uneLongueurIlot), -uneLargeurEntrée, unXi, unYi, unCos, unSin
          .X2 = unXi
          .Y2 = unYi
        End With
      End With
    
      Set wLigne = .linBordVoieSortie(unNumBranche)
      With wLigne
        .Visible = True
        'Calcul des coordonnées de l'intersection du bord
        'de la chaussée sortante avec l'anneau extérieur
        'dans le repère local de la branche (origine centre des anneaux)
        unXLoc = Sqr(Carré(gbRayonExt) - Carré(uneLargeurSortie + unRayonIlot))
        unYLoc = uneLargeurSortie + unRayonIlot
        .Tag = Arccos(unXLoc / gbRayonExt)
        
        'Calcul des coordonnées absolues
        'et stockage dans linBordVoieSortie
        TrRot unXLoc, unYLoc, unXi, unYi, unCos, unSin
        .X1 = unXi
        .Y1 = unYi
        TrRot gbRayonExt + uneLongueurIlot, uneLargeurSortie, unXi, unYi, unCos, unSin
        .X2 = unXi
        .Y2 = unYi
        Set wLigne = gbProjetActif.Données.linVoieSortie(unNumBranche)
        With wLigne
          .Visible = True
          .X1 = unXi
          .Y1 = unYi
          TrRot gbRayonExt + Max(maLongueurBranche, uneLongueurIlot), uneLargeurSortie, unXi, unYi, unCos, unSin
          .X2 = unXi
          .Y2 = unYi
        End With
      End With
    End If
    

    If unRayonIlot <> 0 Then
      'Calcul des trois points Mi (i valant 1,2 ou 3) du triangle formant l'ilot
      'C entre des anneaux et I intersection axe branche avec anneau extérieur
      '                  M1 <-|
      '                  |    |Rayon Ilot
      'C <--Rayon Ext--> I <--|-LongueurIlot---> M2
      '                  |    |Rayon Ilot
      '                  M3 <-|
      
      'Calcul de M1 qui est l'extrémité de linBordIlotEntrée et linBordIlotGir
      TrRot gbRayonExt, unRayonIlot, unXi, unYi, unCos, unSin
      .linBordIlotEntrée(unNumBranche).X1 = unXi
      .linBordIlotEntrée(unNumBranche).Y1 = unYi
      .linBordIlotGir(unNumBranche).X1 = unXi
      .linBordIlotGir(unNumBranche).Y1 = unYi
      'Calcul de M2 qui est l'extrémité de linBordIlotEntrée et linBordIlotSortie
      TrRot (gbRayonExt + uneLongueurIlot), 0, unXi, unYi, unCos, unSin
      .linBordIlotEntrée(unNumBranche).X2 = unXi
      .linBordIlotEntrée(unNumBranche).Y2 = unYi
      .linBordIlotSortie(unNumBranche).X2 = unXi
      .linBordIlotSortie(unNumBranche).Y2 = unYi
      'Calcul de M3 qui est l'extrémité de linBordIlotSortie et linBordIlotGir
      TrRot gbRayonExt, -unRayonIlot, unXi, unYi, unCos, unSin
      .linBordIlotSortie(unNumBranche).X1 = unXi
      .linBordIlotSortie(unNumBranche).Y1 = unYi
      .linBordIlotGir(unNumBranche).X2 = unXi
      .linBordIlotGir(unNumBranche).Y2 = unYi
      
      'Affichage des bords et de l'ilôt
      .linBordIlotEntrée(unNumBranche).Visible = True
      .linBordIlotSortie(unNumBranche).Visible = True
      .linBordIlotGir(unNumBranche).Visible = True
      .linBordVoieEntrée(unNumBranche).Visible = True
      .linBordVoieSortie(unNumBranche).Visible = True
    Else
      'Effacement des bords et de l'ilôt
      .linBordIlotEntrée(unNumBranche).Visible = False
      .linBordIlotSortie(unNumBranche).Visible = False
      .linBordIlotGir(unNumBranche).Visible = False
      If uneLargeurEntrée = 0 Or uneLargeurSortie = 0 Then
        .linBordVoieEntrée(unNumBranche).Visible = False
        .linBordVoieSortie(unNumBranche).Visible = False
      End If
    End If
 
  End With
  
End Sub

'***************************************************************************************
' Opère une Translation/Rotation d'un point connu dans un repère local
' L1, L2 : Coordonnées locales dans le repère de la branche
' unCos, unSin : cosinus et sinus de l'angle fait par l'axe des X du repère local et l'axe horizontal absolu
' X1, Y1 : Coordonnées du point (en twips) dans le repère de la feuille
'***************************************************************************************
Public Sub TrRot(ByVal L1 As Single, ByVal L2 As Single, ByRef X As Single, ByRef Y As Single, ByVal unCos As Single, ByVal unSin As Single)
  X = gbDemiLargeur + (L1 * unCos - L2 * unSin) * gbFacteurZoom
  Y = gbDemiHauteur + (L1 * unSin + L2 * unCos) * gbFacteurZoom
End Sub

'***************************************************************************************
' ModifDynamicBranche  : emprunté à la maquette du CERTU (GIRABASE.FRM)
'***************************************************************************************
Private Sub ModifDynamicBranche(X As Single, Y As Single)
  Dim unRayonDyn As Single
  Dim unDX As Single
  Dim unDY As Single
  Dim unCos As Single
  Dim unSin As Single
  Dim unAngle As Single
  Dim unNumAmont As Integer
  Dim unNumAval As Integer
  
  'Calcul de la distance au centre des anneaux du pointeur souris
  unDX = X - gbDemiLargeur
  unDY = Y - gbDemiHauteur
  unRayonDyn = Sqr(unDX * unDX + unDY * unDY)
  
  'Calcul du cosinus et sinus de l'angle du segment (centre anneaux, pointeur souris)
  unCos = unDX / unRayonDyn
  unSin = unDY / unRayonDyn
  
  'Calcul de l'angle de la branche
  unAngle = CalculerAngle(unCos, unSin)
  
  If VerifierAngleBranche(unAngle, unNumAmont, unNumAval) Then
    'Modification de la branche sélectionnée possible (on reste entre amont et aval)
    ModifierBranche unAngle
  Else
    'Cas où la branche dépasse sa branche aval ou amont
    gbProjetActif.Données.lblInvite.Caption = IDl_LaBranche & CStr(monNumBrancheSelect) & IDm_BorneBranche & CStr(unNumAmont) & IDl_ET & CStr(unNumAval)
  End If

End Sub

'***************************************************************************************
' CalculerAngle  : emprunté à la maquette du CERTU (GIRABASE.FRM)
'***************************************************************************************
Private Function CalculerAngle(ByVal unCos As Single, ByVal unSin As Single) As Single
  'Calcul d'un angle entre 0 et 2PI connaissant son Cosinus et son Sinus
  Dim unAngle As Single
  Dim unEpsilon As Single
  
  unEpsilon = 0.01
  
  'Pour avoir un repère direct ==> Y > 0 vers le haut
  unSin = -unSin
  
  If unCos = 0 Then
    If unSin < 1 + unEpsilon Then
      CalculerAngle = PI / 2
    Else
      CalculerAngle = 1.5 * PI
    End If
  Else
    unAngle = Atn(unSin / unCos) 'donne un résultat entre -PI/2 et +PI/2
    'Conversion pour avoir un résultat entre 0 et 2PI
    If unAngle >= 0 And unSin >= 0 Then
      CalculerAngle = unAngle
    ElseIf unAngle >= 0 And unSin < 0 Then
      CalculerAngle = PI + unAngle
    ElseIf unAngle < 0 And unSin < 0 Then
      CalculerAngle = 2 * PI + unAngle
    ElseIf unAngle < 0 And unSin > 0 Then
      CalculerAngle = PI + unAngle
    End If
  End If
  
End Function

'***************************************************************************************
' VerifierAngleBranche  : emprunté à la maquette du CERTU (GIRABASE.FRM)
'***************************************************************************************
Private Function VerifierAngleBranche(ByVal unAngle As Single, ByRef unNumAmont As Integer, ByRef unNumAval As Integer) As Boolean
  Dim unAngleAval As Single
  Dim unAngleAmont As Single
  Dim unAnglePourTest As Single
  
  With gbProjetActif.Données
    'Recherche du numéro de branche amont et aval
    'Les branches sont numérotées de 1 à NbBranches
    unNumAmont = BranchePrécédent(monNumBrancheSelect)
    unNumAval = BrancheSuivant(monNumBrancheSelect)
    
    'Test si la branche dépasse sa branche amont ou aval
    unAngleAval = CSng(.linBranche(unNumAval).Tag)
    unAngleAmont = CSng(.linBranche(unNumAmont).Tag)
  End With
  
  If unAngleAval < unAngleAmont Then
    'Cas du déplacement d'une branche ayant un angle aval < à l'angle amont
    'car les angles stockés sont entre 0 et 2PI
    unAngleAval = unAngleAval + 2 * PI
    If unAngle < unAngleAmont Then
      unAnglePourTest = unAngle + 2 * PI
    Else
      unAnglePourTest = unAngle
    End If
  Else
    unAnglePourTest = unAngle
  End If
  
  Dim EcartInf As Single
  Dim EcartSup As Single
  With gbProjetActif.Données
    EcartInf = angConv(ArrondiSup(angConv(CSng(.linBordVoieEntrée(unNumAmont).Tag) + CSng(.linBordVoieSortie(monNumBrancheSelect).Tag), False)), CVRADIAN)
    EcartSup = angConv(ArrondiSup(angConv(CSng(.linBordVoieSortie(unNumAval).Tag) + CSng(.linBordVoieEntrée(monNumBrancheSelect).Tag), False)), CVRADIAN)
  End With
  'retourne True si entre angle amont et angle aval, False sinon
  VerifierAngleBranche = unAnglePourTest > unAngleAmont + EcartInf And unAnglePourTest < unAngleAval - EcartSup
  
End Function

Private Function ArrondiSup(ByVal Valeur As Single) As Integer
  ArrondiSup = Round(Valeur + 0.5)
End Function

Public Function BrancheSuivant(ByVal i As Integer) As Integer
' retourne le numéro de branche qui suit immédiatement la branche numéro i
  BrancheSuivant = (i Mod gbProjetActif.NbBranches) + 1
End Function

Public Function BranchePrécédent(ByVal i As Integer) As Integer
' retourne le numéro de branche qui précède immédiatement la branche numéro i
  BranchePrécédent = i - 1
  If BranchePrécédent = 0 Then BranchePrécédent = gbProjetActif.NbBranches
End Function

'***************************************************************************************
' ModifierBranche  : emprunté à la maquette du CERTU (GIRABASE.FRM)
'***************************************************************************************
Public Sub ModifierBranche(unAngle As Single, Optional Invite As Boolean = True)
  Dim unCos As Single
  Dim unSin As Single
  Dim unMax As Single
  Dim wLigne As Line
  Dim xDébutBranche As Single
  Dim yDébutBranche As Single
  Dim xFinBranche As Single
  Dim yFinBranche As Single

  'Calcul du sinus et cosinus de l'angle
  unCos = Cos(unAngle)
  unSin = -Sin(unAngle)         ' "-" : car l'axe des Y est vers le bas
  
  With gbProjetActif.Données
    'Modification de l'axe de la branche
    Set wLigne = .linBranche(monNumBrancheSelect)
    With wLigne
      .X1 = gbDemiLargeur + gbRayonExt * gbFacteurZoom * unCos
      .Y1 = gbDemiHauteur + gbRayonExt * gbFacteurZoom * unSin
      .X2 = gbDemiLargeur + (gbRayonExt + maLongueurBranche) * gbFacteurZoom * unCos
      .Y2 = gbDemiHauteur + (gbRayonExt + maLongueurBranche) * gbFacteurZoom * unSin
    End With
    
    'Déplacement de la poignée de sélection
    With .shpPoignée
      .Left = wLigne.X2 - .Width / 2
      .Top = wLigne.Y2 - .Height / 2
    End With
    
    'Déplacement du nom de la branche
    DéplacerNomBranche .lblLibelléBranche(monNumBrancheSelect), wLigne, unCos, unSin
'''    With .lblLibelléBranche(monNumBrancheSelect)
'''''      unMax = Max(.Width, .Height) + uneMarge
'''      xFinBranche = wLigne.X2
'''      yFinBranche = wLigne.Y2
'''      DecalXY xFinBranche, yFinBranche, unCos, unSin, .Caption
'''      .Left = xFinBranche ''+ (unMax * unCos - .Width) / 2
'''      .Top = yFinBranche ''+ (unMax * unSin - .Height) / 2
'''    End With
    
    'Déplacement du numéro de la branche
    With .lblNumBranche(monNumBrancheSelect)
      xDébutBranche = wLigne.X1
      yDébutBranche = wLigne.Y1
      unMax = Max(.Width, .Height) + 3 * uneMarge   ' Coef 3 : Car on a substitué n à Bn
      .Left = xDébutBranche + (-unMax * unCos - .Width) / 2
      .Top = yDébutBranche + (-unMax * unSin - .Height) / 2
    End With
    
    'Déplacement du reste de la branche
    DessinerBranche monNumBrancheSelect, unAngle
    
    'Affichage dynamique de la valeur du rayon
    wLigne.Tag = CStr(unAngle)
    monAngle = Format(angConv(unAngle, False), "0")
    If Invite Then
      .lblInvite.Caption = IDl_Angle & IDl_DeLaBranche & IDl_Egal & monAngle & " " & libelAngle(gbProjetActif.modeangle)
    End If
    
  End With  ' gbProjetActif.Données
  
End Sub

Public Sub DéplacerNomBranche(ByVal lblBranche As Label, ByVal wLigne As Line, ByVal unCos As Single, ByVal unSin As Single)
Dim xFinBranche As Single, yFinBranche As Single
    'Déplacement du nom de la branche
    With lblBranche
      xFinBranche = wLigne.X2
      yFinBranche = wLigne.Y2
      DecalXY xFinBranche, yFinBranche, unCos, unSin, .Caption
      .Left = xFinBranche
      .Top = yFinBranche
    End With

End Sub

'******************************************************************************************
'Utilitaires empruntés à la maquette du CERTU (GIRABASE.FRM)
' DistanceAuCentreVue - gbDemiLargeur - gbDemiHauteur
' Obsolètes : gbDemiHauteur et gbDemiLargeur sont calculés une fois pour toutes dans Dessin_MouseDown
'             DistanceAuCentreVue ne sert que dans Dessin_MouseMove
'*******************************************************************************************

'---> A remettre dans OUTILS(?)

Public Function TransRot(ByVal p As PT, ByVal Trans As PT, ByVal alpha As Single, ByVal Echelle As Single) As PT
' Translation-Rotation d'un point p, point d'insertion de bloc avec ou sans facteur d'échelle
Dim p0 As New PT

  p0.X = p.X * Echelle
  p0.Y = p.Y * Echelle
  If alpha <> 0 Then
    Set p0 = Rotation(p0, angConv(alpha, CVRADIAN))
  End If
  ' translation
  p0.X = p0.X + Trans.X
  p0.Y = p0.Y + Trans.Y
  
  Set TransRot = p0
  
End Function

Public Function Rotation(ByVal p As PT, ByVal alpha As Single) As PT
Dim p0 As New PT

  p0.X = p.X * Cos(alpha) - p.Y * Sin(alpha)
  p0.Y = p.X * Sin(alpha) + p.Y * Cos(alpha)
  Set Rotation = p0
  
End Function

Public Function RotTrans(ByVal p As PT, ByVal Trans As PT, ByVal alpha As Single) As PT
Dim p0 As New PT
  
  Set p0 = Rotation(p, alpha)
  p0.X = p0.X + Trans.X
  p0.Y = p0.Y + Trans.Y
  Set RotTrans = p0
  
End Function

Public Function Distance(ByVal X As Double, ByVal X1 As Double, ByVal Y As Double, ByVal Y1 As Double) As Double
  Distance = Sqr(Carré(X1 - X) + Carré(Y1 - Y))
End Function

Public Function Carré(ByVal v As Double) As Variant
  Carré = v ^ 2
End Function

Private Function trEchel(ByVal L As Single, ByVal toReel As Boolean) As Single
'Transformée d'une longueur L : d'unités dessin en unités réelles ou réciproquement
' L est exprimé en unités dessin (toReel=False) ou réelles (toReel=True)
  If toReel Then trEchel = L / gbFacteurZoom Else trEchel = L * gbFacteurZoom
End Function

Private Function trEchelX(ByVal X As Single, ByVal toReel As Boolean) As Single
'Transformée de X : d'unités dessin en unités réelles ou réciproquement
' orx est exprimé en unités dessin
' mil.x est exprimé en unités réelles
' X est exprimé en unités dessin (toReel=False) ou réelles (toReel=True)
Dim mil As PT
Dim Orx As Single
  With gbProjetActif.Données
    If toReel Then
      trEchelX = mil.X + (X - Orx) / gbFacteurZoom
    Else
      trEchelX = (X - mil.X) * gbFacteurZoom + Orx
    End If
  End With
End Function

Private Function trEchelY(ByVal Y As Single, ByVal toReel As Boolean) As Single
'Transformée de Y : d'unités dessin en unités réelles ou réciproquement
' ory est exprimé en unités dessin
' mil.y est exprimé en unités réelles
' Y est exprimé en unités dessin (toReel=False) ou réeelles (toReel=True)
Dim mil As PT
Dim Ory As Single
  With gbProjetActif.Données
    If toReel Then
      trEchelY = mil.Y - (Y - Ory) / gbFacteurZoom
    Else
      trEchelY = -(Y - mil.Y) * gbFacteurZoom + Ory
    End If
  End With
End Function

Private Sub PoignéeMove(ByVal X As Single, ByVal Y As Single)
' Déplacement de la poignée où se trouve la souris
  With gbProjetActif.Données.shpPoignée
    .Left = X - .Width / 2
    .Top = Y - .Height / 2
  End With

End Sub
Private Sub DessinerAnneau(ByVal controle As Shape, ByVal Rayon As Single)
'Dessin de l'anneau - Rayon est en twips
  With controle
    .Width = Rayon * 2
    .Height = Rayon * 2
    .Left = gbDemiLargeur - Rayon
    .Top = gbDemiHauteur - Rayon
  End With
End Sub

Private Sub DecalXY(ByRef Left As Single, ByRef Top As Single, ByVal unCos As Single, ByVal unSin As Single, ByVal Texte As String)
  Dim LgTexte As Single
 
  LgTexte = gbProjetActif.Données.TextWidth(Texte)
  Top = Top - 100
  ''Left = Left + 100
  Left = Left + (unCos / 2 - 0.5) * LgTexte
  Top = Top + unSin * gbProjetActif.Données.TextHeight("")
  If Abs(unSin) < 0.15 Then Top = Top - gbProjetActif.Données.TextHeight("")
  If Abs(unCos) < 0.075 Then
    Top = Top - Sgn(unSin) * gbProjetActif.Données.TextHeight("") * 0.25
  End If
  Left = Max(0, Left)
  Left = Min(Left, gbDemiLargeur * 2 - LgTexte)

End Sub
