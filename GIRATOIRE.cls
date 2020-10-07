VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GIRATOIRE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'******************************************************************************
'*
'*          Projet GIRABASE - CERTU - CETE de l'Ouest
'*
'*          Module de classe : GIRATOIRE.CLS
'*
'*          Définition de l'objet GIRATOIRE
'*
'******************************************************************************

Option Explicit


Public NomVariante As String  '  Ces données peuvent être récupérées directement dans les controles
Public DateModif As Date      '     ""
Public NbBranches As Integer   '     ""

'Public nbPériodes As Integer  ' nombre de périodes de trafics
Public modeangle As Integer

Public nom As String
Public Localisation As String
Public Milieu As Integer
'Public Commentaire As String

Public R As Single            ' Rayon de l'ilôt infranchissable
Public LA As Single           ' Largeur de l'anneau
Public Bf As Single           ' Largeur de la bande franchissable
'Public Rg As Single          ' Rayon extérieur : calculé --> R+LA+Bf


Public colBranches As New Branches
Public colTrafics As New Trafics

' Feuilles filles MDI attachée au projet
Public Données As Form
Public Résultats As Form
 
Public FacteurZoomPrecedent As Single

' Variables nécessaires au calcul de capacité
Public Tg As Single          ' Coefficient Largeur d'entrée
Public Te As Single          ' Coefficient créneau critique
Public Tf1 As Single         ' Coefficient créneau complémentaire
Public LImax As Single       ' Largeur d'ilot maximum

Public RU As Single          ' Rayon utile de l'ilôt infranchissable
Public LAU As Single         ' Largeur utile de l'anneau
Public LEU As Single        ' Largeur d'entrée utile
Public KI As Single          ' Coefficient de gêne trafic tournant à l'intérieur de l'anneau
Public KE As Single          ' Coefficient de gêne trafic tournant à l'extérieur de l'anneau

Public CalculFait As Boolean
Public nbPériodesSaturées As Integer  ' nombre de périodes de trafics saturées

' Maintenance : 28/11/2000 : Activation de la protection pour empêcher que le fichier soit ouvert par qq d'autre
Private numFichProtégé As Integer


'******************************************************************************
' Création du Giratoire : Dimensionnement des tableaux et affichage des données
'*******************************************************************************
Public Sub Création() ' (ByVal DrapeauRein As Boolean)
  Dim DiffBranches As Integer
  Dim i As Integer
  Const LARGEURCELLULE = 570

DiffBranches = NbBranches - DEFAUTNBBRANCHES

' dimensionnement des grilles de saisie
With Données
  .vgdCarBranche.MaxRows = NbBranches
  For i = 1 To 4  ' Ajout AV 22/02/2000 - V4.0.18 gestion du point décimal sur le Spread
    .vgdLargBranche.Col = i
    .vgdLargBranche.TypeFloatDecimalChar = gbPtDecimal
  Next
  .vgdLargBranche.MaxRows = NbBranches
  .vgdTrafic(VEHICULE).MaxCols = NbBranches
  .vgdTrafic(VEHICULE).MaxRows = NbBranches
  .vgdTrafic(PIETON).MaxCols = NbBranches
  .vgdTrafic(VEHICULE).Col = -1
  .vgdTrafic(VEHICULE).ColWidth(-1) = WIDTHCOLTRAFIC
  .vgdTrafic(PIETON).Col = -1
  .vgdTrafic(PIETON).ColWidth(-1) = WIDTHCOLTRAFIC
  
'Repositionnement du cadre Trafic total entrant
  With .fraQTE
    .Move .Left + DiffBranches * LARGEURCELLULE, .Top, .Width, .Height + DiffBranches * 225
  End With
'Repositionnement du cadre Trafic total sortant
  With .fraQTS
    .Move .Left, .Top + DiffBranches * 225, .Width + DiffBranches * LARGEURCELLULE
  End With
'Repositionnement du Trafic Total dans le cadre Trafic total entrant
  With .txtQT
    .Move .Left, .Top + DiffBranches * 225
  End With
 'Chargement des trafics entrant et sortant
  For i = 1 To NbBranches
    Load .txtQS(i)
    Load .txtQE(i)
    With .txtQE(i)
      .Move .Left, .Top + (i - 1) * 225
      .Visible = True
    End With
    With .txtQS(i)
      .Move .Left + (i - 1) * LARGEURCELLULE
      .Visible = True
    End With
  Next
End With

Affiche
    
End Sub

Public Function nbPériodes() As Integer
  nbPériodes = colTrafics.count - nbPériodesSaturées
End Function
'******************************************************************************
' Affichage des données du Giratoire
'*******************************************************************************
Private Sub Affiche()
  Dim i As Integer
  
  With Données
    .lblDateModif = DateModif
    .txtNomGiratoire = nom
    .txtLocalisation = Localisation
    .txtVariante = NomVariante
    .txtR = R
    .txtBf = Bf
    .txtLA = LA
    .calculRg False ' False :ne pas redessiner le giratoire
    If Milieu <> -1 Then .optMilieu(Milieu) = True
  End With

  colBranches.Affiche Données
  
  If nbPériodes > 0 Then Données.cboPériode.ListIndex = 0 ' déclenche l'évènement cboPériode_Click

End Sub

'******************************************************************************
' Changement d'unités d'angle : Bascule DEGRE<-->GRADE
'*******************************************************************************
Public Sub ChangeUnitéAngle()
Dim CoefAngle As Single
Dim i As Integer
Dim AnglePrec As Integer

  If modeangle = DEGRE Then
    CoefAngle = 10 / 9
    modeangle = GRADE
  Else
    CoefAngle = 0.9
    modeangle = DEGRE
  End If
  
  'Libellés d'angle
  Données.lblAngleBranche.Caption = IDl_Angle & " (" & libelAngle(modeangle) & ")"
  'Borne supérieure de l'angle
  With Données.vgdCarBranche
    .Col = 2
    For i = 1 To 8
      .Row = i
      If modeangle = DEGRE Then
        .TypeIntegerMax = 359
      Else
        .TypeIntegerMax = 399
      End If
    Next
    'Borne de l'écart de l'angle
    .Col = 3
    For i = 1 To 8
      .Row = i
      If modeangle = DEGRE Then
        .TypeIntegerMax = 359
      Else
        .TypeIntegerMax = 399
      End If
    Next
  End With

  With colBranches
    For i = 1 To .count
    ' Recalcul du nouvel angle
      With .Item(i)
        .Angle = .Angle * CoefAngle
        If i > 1 Then .Ecart = .Angle - AnglePrec
        AnglePrec = .Angle
      End With
    Next
    
  'Affichage de la branche
    .Affiche Données
  End With
  
  GriserMenus True
End Sub

'******************************************************************************
' Activation du Giratoire : le giratoire devient le projet courant
'*******************************************************************************
Public Sub Activate()
  
  Set gbProjetActif = Me
  
    DessinGiratoire.gbDemiHauteur = Données.DemiHauteur
    DessinGiratoire.gbDemiLargeur = Données.DemiLargeur
    DessinGiratoire.gbFacteurZoom = Données.FacteurZoom
    DessinGiratoire.gbRayonInt = R
    DessinGiratoire.gbRayonExt = R + LA + Bf
    DessinGiratoire.gbBandeFranchissable = Bf

End Sub

'******************************************************************************
' Enregistrement du Giratoire
'*******************************************************************************
Public Sub Enregistrer(ByRef Cancel As Integer)
Dim NomFich As String
  
  ' Désactivation de la protection pour que le test qui suit ne concerne pas le programme en cours
  If numFichProtégé <> 0 Then Close numFichProtégé
  
  NomFich = Données.dlgFichier.FileName
  If NomFich <> "" And Not FichierProtégé(NomFich, , NomFich) Then
    Ecrire NomFich
  Else
    If numFichProtégé <> 0 Then ' Réactivation de la protection
      numFichProtégé = FreeFile
      Open NomFich For Random Lock Read As numFichProtégé
    End If
    EnregSous Cancel
  End If

End Sub

'******************************************************************************
' Enregistrement d'un Giratoire sous un autre nom
'*******************************************************************************
Public Sub EnregSous(ByRef Cancel As Integer)
  
  On Error GoTo TraitementErreur
  
  With Données.dlgFichier
    .flags = cdlOFNOverwritePrompt Or cdlOFNHideReadOnly Or cdlOFNNoChangeDir Or cdlOFNPathMustExist
    .InitDir = App.Path
    .ShowSave
    If Not Cancel Then
      If FichierProtégé(.FileName, , .FileName) Then
        Cancel = True
      Else
        Ecrire .FileName
        Données.Caption = suppExt(.FileTitle)
        If Not Résultats Is Nothing Then Résultats.Caption = IDl_Résultats & " - " & Données.Caption
        MRUmenu (.FileName)
      End If
    End If
  End With
  
  Exit Sub
  
TraitementErreur:   ' L'utilisateur a fait 'Annuler
  Cancel = True
  If Err = cdlCancel Then
    Resume Next
  Else
    ErreurFatale "EnregSous"
  End If

End Sub

'******************************************************************************
' Ecriture d'une Variante du Giratoire dans un fichier
'*******************************************************************************
Public Sub Ecrire(ByVal NomFich As String)
Dim numFich As Integer
Dim i As Integer
Dim j As Integer
Dim véhicule As Integer
Dim Chaine As String

' Maintenance : 28/11/2000 : Activation de la protection pour empêcher que le fichier soit ouvert par qq d'autre
  If numFichProtégé <> 0 Then Close numFichProtégé
  
  ' Ouverture du Fichier
  numFich = FreeFile
  Open NomFich For Output Lock Write As numFich
  
  ' Données générales
  DateModif = Date
  Données.lblDateModif = DateModif
  Write #numFich, App.Title, "Version " & App.Major
'  Write #numFich, App.Title, "Version provisoire T0.2"
  Write #numFich, NomVariante, DateModif
  Write #numFich, nom, Milieu  ', Commentaire
  Write #numFich, NbBranches, nbPériodes, modeangle
  'Anneau
  Write #numFich, R, Bf, LA
  'Localisation (commentaire multiligne)
  Write #numFich, Localisation
  
  'Branches
  Write #numFich, "BRANCHES"
  With colBranches
    For i = 1 To .count
      With .Item(i)
        Write #numFich, .nom, .Angle, .Rampe, .TAD
        Write #numFich, .EntréeEvasée, .LE4m, .LE15m, .LI, .LS
      End With
    Next
  End With
  
  'Trafics
  Write #numFich, "TRAFICS"
  With colTrafics
    For i = 1 To nbPériodes
      With .Item(i)
        Write #numFich, .nom, .modeUVP, .CouleurCourbe
        Write #numFich, .FormaterQP()
        If .modeUVP Then
          For j = 1 To NbBranches
            Write #numFich, .FormaterQV(UVP, j)
          Next
        Else
          For véhicule = VL To DEUXR
            For j = 1 To NbBranches
              Write #numFich, .FormaterQV(véhicule, j)
            Next
          Next
        End If
      End With
    Next
  End With
  
  'Fermeture
  Close numFich
  
  ' Verrouillage du fichier le temps de la session
' Maintenance : 28/11/2000
  numFichProtégé = FreeFile
  Open NomFich For Random Lock Read As numFichProtégé
  
  gbProjetActif.Données.FichierModifié = False
  GriserMenus True
     
End Sub

'******************************************************************************
' Lecture d'un Giratoire dans un fichier
'*******************************************************************************
Public Function Lire() As Boolean
Dim numFich As Integer
Dim numLigne As Integer
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim Chaine As String
Dim variable As Variant
Dim modeUVP As Boolean

Dim nomBranche As String       ' Nom de la branche
Dim Angle As Single            ' Angle en degrés
Dim Rampe As Boolean           ' Présence d'une rampe > 3%
Dim EntréeEvasée As Boolean
Dim TAD As Boolean             ' Tourne à droite

Dim LE4m As Single             ' Largeur d'entrée à 4m
Dim LE15m  As Single           ' Largeur d'entrée à 15m
Dim LI As Single               ' Largeur de l'îlot
Dim LS As Single               ' Largeur de sortie

Dim nbTrafics As Integer
Dim CouleurCourbe As Long

Dim NomFich As String

  NomFich = MDIGirabase.dlgFichier.FileName
  
On Error GoTo TraitementErreur
  
  If FichierProtégé(NomFich, , Titre:=NomFich, LectureSeuleAutorisée:=True) Then Err.Raise 101
  
  numFich = FreeFile
  Open NomFich For Input As numFich
  Input #numFich, Chaine
  numLigne = numLigne + 1
  If Chaine <> App.Title Then Err.Raise 100
  Input #numFich, Chaine
  If Chaine <> "Version " & App.Major Then Err.Raise 100
'  If chaine <> "Version provisoire T0.2" Then Err.Raise 100
  
  Input #numFich, NomVariante, variable
  numLigne = numLigne + 1
  If VarType(variable) <> vbDate Then Err.Raise 100
  DateModif = variable
  
  Input #numFich, nom, variable
  numLigne = numLigne + 1
  OkEntier variable, Milieu, -1, 2
  
  'Caractéristiques du projet
  Input #numFich, variable
  numLigne = numLigne + 1
  OkEntier variable, NbBranches, 3, 8
  Input #numFich, variable
  OkEntier variable, nbTrafics, 0
  Input #numFich, variable
  OkEntier variable, modeangle, DEGRE, GRADE
  
  'Caractéristiques de l'anneau
  Input #numFich, Chaine
  numLigne = numLigne + 1
  OkFlottant Chaine, R
  
  Input #numFich, Chaine
  OkFlottant Chaine, Bf
  
  Input #numFich, Chaine
  OkFlottant Chaine, LA

  Input #numFich, Chaine
  numLigne = numLigne + 1
  Do
    Localisation = Localisation & Chaine
    Input #numFich, Chaine
    numLigne = numLigne + 1
  Loop While Chaine <> "BRANCHES" And Not EOF(numFich)
  If Chaine <> "BRANCHES" Then Err.Raise 100
  
  ' Branches
  With colBranches
    For i = 1 To NbBranches
      If Not .Item(i).Lire(numFich, numLigne, modeangle) Then Err.Raise 100
      If i > 1 Then ' Calcul des écarts et vérification des angles croissants
        .Item(i).Ecart = .Item(i).Angle - .Item(i - 1).Angle
        If .Item(i).Ecart <= 0 Then numLigne = numLigne - 1: Err.Raise 100
      End If
    Next
    'Supppression des branches inutiles (par défaut le giratoire a été dimensionné à 8 (gbCreFille)
    For i = NbBranches + 1 To 8
      .Remove NbBranches + 1
    Next
  End With
  
  Input #numFich, Chaine
  numLigne = numLigne + 1
  If Chaine <> "TRAFICS" Then Err.Raise 100
  
  'Trafics
  With colTrafics
    For i = 1 To nbTrafics
      Input #numFich, Chaine, variable
      numLigne = numLigne + 1
      If colTrafics.EnDouble(Chaine) Then Err.Raise 100
      If VarType(variable) <> vbBoolean Then Err.Raise 100
      modeUVP = variable
      Input #numFich, variable
      OkLong variable, CouleurCourbe, 0, &HFFFFFF
      .add Me, Chaine
      With .Item(i)
        .modeUVP = modeUVP
        .CouleurCourbe = CouleurCourbe
        If Not .Lire(numFich, numLigne) Then Err.Raise 100
      End With
    Next
  End With

  'Fermeture
  Close numFich
  
' Maintenance : 28/11/2000
  ' Verrouillage du fichier le temps de la session
  numFichProtégé = FreeFile
  Open NomFich For Random Lock Read As numFichProtégé
    
  DésarmIndic
  Lire = True
  If gbFichierJournal Then Write #gbFichLog, MDIGirabase.dlgFichier.FileName
  Exit Function
  
TraitementErreur:
  gbErrLecFichier = True
  Select Case Err
  Case 100
    MsgBox IDm_ErrLectureFichier & " - " & IDm_ligne & " " & CStr(numLigne)
    Close numFich
    DésarmIndic
  Case 101
  Case Else
    ErreurFatale
  End Select
End Function

Private Sub DésarmIndic()
  ' Les indicateurs ont pu être armés pendant la lecture
  Données.DonnéeModifiée = False
  Données.FichierModifié = False
End Sub

'******************************************************************************
' Calcul de largeur d'ilot maximum
'*******************************************************************************
Public Sub MajComplément(Optional ByVal NouveauMilieu As Integer = -1)
Dim Tg As Single

  If NouveauMilieu <> -1 Then Milieu = NouveauMilieu
  If Milieu = -1 Then Exit Sub
  

  If R = 0 Then
  ' Mini-Giratoire  : e-mail du 6/9/99
    Tg = gbTg(rc)
    RU = 3.5
    LAU = LA + Bf - 3.5
  Else
    Tg = gbTg(Milieu)
  ' Ref : Modificatif 29/12/98
    RU = R + 0.5 * Bf
    LAU = LA + 0.5 * Bf
  End If
    
    ' Ref  §2.1.1

  ' Si R Et Bf=0 : invalide(le lancement du calcul sera empêché), mais acceptés provisoirt en saisie d'où On error...
  On Error Resume Next
  ' gbCoefLEU vaut 1.2 (1.1 en version belge)
  LEU = LAU / (gbCoefLEU * (1 + 1 / 2 / RU))

' Ref  §2.1.4
  LImax = Tg * Sqr(RU + LAU / 2)

End Sub

'********************************************************************************************************************
' Ajout d'une période de trafic
' Cette procédure peut être déclenchée de 3 manières
'   Saisie d'un nouveau nom dans la combo cboPériode                      --> DrapeauMenu = False
'   Appel du menu                                                         --> DrapeauMenu = True
'   Appel depuis une autre procédure (Inversion - Multiplication - Import)--> DrapeauMenu = Nom de la période à créer
' Dans ce dernier cas, le traitement final est différé (sera réalisé par TRAFIC.Dupliquer)
'********************************************************************************************************************
Public Function newPériode(ByVal DrapeauMenu As Variant) As Boolean
Dim nomPériode As String
Dim i As Integer

  Select Case DrapeauMenu
  Case False     ' Appel depuis la ComboBox
    nomPériode = Données.cboPériode.Text
  Case True      ' Appel depuis le Menu
    nomPériode = InputBox(IDl_NouvellePériode, App.Title, colTrafics.defautNew, , , App.HelpFile, IDhlp_NewPériode)
  Case Else      ' Appel avec passage du nouveau nom (Inversion - Multiplication - Import)
                 ' Pour inversion, on pourra ajouter un bouton d'aide si le CERTU le spécifie (code 204 : IDhlp_InversPériode)
    If InStr(DrapeauMenu, IDl_Inversion) <> 0 Then
      nomPériode = InputBox(IDl_NouvellePériode, App.Title, DrapeauMenu, , , App.HelpFile, IDhlp_InversPériode)
    ElseIf InStr(DrapeauMenu, IDl_Multiplication) <> 0 Then
      nomPériode = InputBox(IDl_NouvellePériode, App.Title, DrapeauMenu, , , App.HelpFile, IDhlp_MultPériode)
    Else
      nomPériode = InputBox(IDl_NouvellePériode, App.Title, DrapeauMenu)
    End If
  End Select

  If nomPériode <> "" And Not colTrafics.EnDouble(nomPériode) Then
    ' Ajout effectif
    colTrafics.add Me, nomPériode
    ' Traitement final : la période créée devient la période courante
    If VarType(DrapeauMenu) = vbBoolean Then
    ' Sinon, l'appel est fait depuis une autre fonction : traitement final différé
      With Données
        colTrafics.Item(colTrafics.count).modeUVP = .lblTraficUVP.Visible
        With .cboPériode
          .ListIndex = .ListCount - 1     ' déclenche l'évènement cboPériode_Click
        End With
        ' La ligne suivante doit en fait être appelée lors de l'activation de l'onglet (+ messages à porter lors des modifs dans vgdLargBranche)
        colBranches.BlocageTrafic Données
      End With
    End If
    
    newPériode = True
    GriserMenus True
  End If

End Function

'******************************************************************************
' Suppression d'une période de trafic
'*******************************************************************************
Public Sub delPériode(ByVal wTrafic As TRAFIC)
Dim nomPériode As String
Dim numPériode As Integer

  With wTrafic
    If MsgBox(IDm_SupprPériode & " " & .nom, vbYesNo + vbDefaultButton2 + vbQuestion + vbMsgBoxHelpButton, , App.HelpFile, IDhlp_DelPériode) = vbYes Then
      With Données.cboPériode
        numPériode = .ListIndex
        colTrafics.Remove Me, numPériode + 1
        .ListIndex = Min(numPériode, nbPériodes - 1)  ' déclenche l'évènement cboPériode_Click
      End With
      If nbPériodes = 0 Then
        Données.cboPériode_Click  ' l'évènement n'a pas été déclenché (ListIndex=-1)
        If Not Résultats Is Nothing Then Unload Résultats
      End If
      GriserMenus True
    End If
  End With
  
End Sub

'******************************************************************************
' Renommer une période de trafic
'*******************************************************************************
Public Sub renamePériode(ByVal wTrafic As TRAFIC)
Dim nomPériode As String
Dim numPériode As Integer

  With wTrafic
    nomPériode = InputBox(IDl_RenPériode, App.Title, .nom, , , App.HelpFile, IDhlp_RenamePériode)
    If nomPériode <> "" And Not colTrafics.EnDouble(nomPériode) Then
      ' Renommage effectif
      .nom = nomPériode
      With Données.cboPériode
        numPériode = .ListIndex
        .RemoveItem numPériode
        .AddItem nomPériode, numPériode
        .ListIndex = numPériode
      End With
      If Not Résultats Is Nothing Then
        With Résultats.cboPériode
          .RemoveItem numPériode
          .AddItem nomPériode, numPériode
        End With
      End If
      Données.FichierModifié = True
    End If
  End With
  
End Sub

'******************************************************************************
' Inversion des matrices d'une période de trafic
'*******************************************************************************
Public Sub inversPériode(ByVal wTrafic As TRAFIC)

  If wTrafic.EstComplète Then
    If newPériode(IDl_Inversion & IDl_DE & wTrafic.nom) Then
      wTrafic.Inverser
    End If
  Else
    MsgBox IDm_PériodeIncomplète, vbExclamation + vbMsgBoxHelpButton, App.Title, App.HelpFile, IDhlp_InversPériode
  End If
End Sub

'******************************************************************************
' Multiplication des matrices d'une période de trafic
'*******************************************************************************
Public Sub multPériode(ByVal wTrafic As TRAFIC)

  If wTrafic.EstComplète Then
    If newPériode(IDl_Multiplication & IDl_DE & wTrafic.nom) Then
      Set frmTrafMult.TraficOrigine = wTrafic
      frmTrafMult.Show vbModal
    End If
  Else
    MsgBox IDm_PériodeIncomplète, vbExclamation + vbMsgBoxHelpButton, App.Title, App.HelpFile, IDhlp_MultPériode
  End If
End Sub

'****************************************************************************************************
' Vérification de la compatibilité du projet importé avec le giratoire, en vue de l'import de matrice
'****************************************************************************************************
Public Function VerifCompatible(wGiratoire As GIRATOIRE) As Boolean
Dim i As Integer

  If wGiratoire.nbPériodes = 0 Then
    MsgBox IDm_PasDePériode, vbExclamation: Exit Function
  ElseIf wGiratoire.NbBranches <> NbBranches Then
    MsgBox IDm_NbBranchesDifférent, vbExclamation + vbMsgBoxHelpButton, , App.HelpFile, IDhlp_ImportMatrice: Exit Function
  Else
    For i = 1 To NbBranches
      If colBranches.Item(i).EntréeNulle <> wGiratoire.colBranches.Item(i).EntréeNulle Then
        MsgBox IDm_IncompatibleBrancheUnidirection, vbExclamation + vbMsgBoxHelpButton, , App.HelpFile, IDhlp_ImportMatrice: Exit Function
      End If
      If colBranches.Item(i).SortieNulle <> wGiratoire.colBranches.Item(i).SortieNulle Then
        MsgBox IDm_IncompatibleBrancheUnidirection, vbExclamation + vbMsgBoxHelpButton, , App.HelpFile, IDhlp_ImportMatrice: Exit Function
      End If
    Next
  End If
  
  VerifCompatible = True
    
End Function

'******************************************************************************
' Suppression des périodes saturées
'*******************************************************************************
Public Sub delPériodeSaturées()
Dim i As Integer
  With colTrafics
    For i = .count To 1 Step -1
      If .Item(i).BrancheSaturée <> 0 Then
        .Remove Me, i
      End If
    Next
  End With

End Sub

'******************************************************************************
' Calculs de capacité : Toutes les périodes complètes
'*******************************************************************************

Public Sub CalculCapacité(Optional ByVal wTrafic As TRAFIC = Nothing)
Dim i As Integer
Dim j As Integer

' Appel pourle calcul d'une seule période
  If Not wTrafic Is Nothing Then
    With wTrafic
      If .EstComplète Then
        .CalculTraficEntrant colBranches
        For j = 1 To NbBranches
          .CalculCapacité j, colBranches
        Next
        .CalculFait = True
      End If
    End With
    Exit Sub
  End If
  
' Appel pour le calcul de l'ensemble
  If Résultats Is Nothing Then
    ChargeFeuilleRésultat
  Else
    Résultats.SetFocus
    ' Suppression des périodes avec saturation de branche
    delPériodeSaturées
  End If
    
  CalculParamGiratoire
  For i = 1 To colBranches.count
    colBranches.Item(i).CalculParamBranche
  Next
  
  For i = 1 To colTrafics.count
    CalculCapacité colTrafics.Item(i)
  Next
  
  CalculFait = True
  GriserMenus True
 ' Affichage des résultats de la période active
  Résultats.cboPériode.ListIndex = Données.cboPériode.ListIndex
 
End Sub

'******************************************************************************
' Coefficients du Giratoire : Tg,Te,Tf1 - RU,RAU,LEU,LImax - KE, KI
'******************************************************************************
Private Sub CalculParamGiratoire()
   
  ' Détermination des coefficients pour le calcul des capacités : Te, Tg, Tf1
  
  If R = 0 Then
  ' Mini-giratoire  - cd e-mail du 6/9/99
    Te = gbTe(rc)
    Tg = gbTg(rc)
    Tf1 = gbTf1(rc)
  Else
    Te = gbTe(Milieu)
    Tg = gbTg(Milieu)
    Tf1 = gbTf1(Milieu)
  End If
   
  ' Détermination des coefficients de gêne KE (trafic tournant intérieur), KI (trafic tournant extérieur)
  
'  ' Ref : Modificatif 29/12/98
'  RU = R + 0.5 * Bf
'  LAU = LA + 0.5 * Bf
'
'  ' Ref  §2.1.1
'  LEU = LAU / (1.2 * (1 + 1 / 2 / RU))
'
'  ' Ref  §2.1.4
'  LImax = Tg * Sqr(RU + LAU / 2)
  
    ' Ref  §2.2.1 modifié par Réponse.Doc du 12/03/99
'  KI = 160 / (LAU * (RU + LAU))
  KI = 8 / LAU * Sqr(20 / (RU + LAU))
  If LAU > 8 Then
'    KE = 1 - (RU + LAU) ^ 2 * (LAU - 8) / LAU
    KE = 1 - (RU / (RU + LAU)) ^ 2 * (LAU - 8) / LAU
'    If gbFichierJournal Then Write #gbFichLog, "2.2.2 - KE calculé avec LAU >8"
 Else
    KE = 1
'    If gbFichierJournal Then Write #gbFichLog, "2.2.2 - KE =1 --> LAU <8"
  End If
  
  'Ref §2.3.3
  KI = Min(KE, KI)
  
End Sub

Private Sub ChargeFeuilleRésultat()
Dim i As Integer

  Set Résultats = New frmRésultats
  Set Résultats.GiratoireProjet = Me

End Sub

' Maintenance : 28/11/2000 : Activation de la protection pour empêcher que le fichier soit ouvert par qq d'autre
Private Sub Class_Terminate()
  If numFichProtégé <> 0 Then Close numFichProtégé
End Sub
