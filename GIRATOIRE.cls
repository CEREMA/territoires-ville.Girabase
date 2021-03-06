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
'*          D�finition de l'objet GIRATOIRE
'*
'******************************************************************************

Option Explicit


Public NomVariante As String  '  Ces donn�es peuvent �tre r�cup�r�es directement dans les controles
Public DateModif As Date      '     ""
Public NbBranches As Integer   '     ""

'Public nbP�riodes As Integer  ' nombre de p�riodes de trafics
Public modeangle As Integer

Public nom As String
Public Localisation As String
Public Milieu As Integer
'Public Commentaire As String

Public R As Single            ' Rayon de l'il�t infranchissable
Public LA As Single           ' Largeur de l'anneau
Public Bf As Single           ' Largeur de la bande franchissable
'Public Rg As Single          ' Rayon ext�rieur : calcul� --> R+LA+Bf


Public colBranches As New Branches
Public colTrafics As New Trafics

' Feuilles filles MDI attach�e au projet
Public Donn�es As Form
Public R�sultats As Form
 
Public FacteurZoomPrecedent As Single

' Variables n�cessaires au calcul de capacit�
Public Tg As Single          ' Coefficient Largeur d'entr�e
Public Te As Single          ' Coefficient cr�neau critique
Public Tf1 As Single         ' Coefficient cr�neau compl�mentaire
Public LImax As Single       ' Largeur d'ilot maximum

Public RU As Single          ' Rayon utile de l'il�t infranchissable
Public LAU As Single         ' Largeur utile de l'anneau
Public LEU As Single        ' Largeur d'entr�e utile
Public KI As Single          ' Coefficient de g�ne trafic tournant � l'int�rieur de l'anneau
Public KE As Single          ' Coefficient de g�ne trafic tournant � l'ext�rieur de l'anneau

Public CalculFait As Boolean
Public nbP�riodesSatur�es As Integer  ' nombre de p�riodes de trafics satur�es

' Maintenance : 28/11/2000 : Activation de la protection pour emp�cher que le fichier soit ouvert par qq d'autre
Private numFichProt�g� As Integer


'******************************************************************************
' Cr�ation du Giratoire : Dimensionnement des tableaux et affichage des donn�es
'*******************************************************************************
Public Sub Cr�ation() ' (ByVal DrapeauRein As Boolean)
  Dim DiffBranches As Integer
  Dim i As Integer
  Const LARGEURCELLULE = 570

DiffBranches = NbBranches - DEFAUTNBBRANCHES

' dimensionnement des grilles de saisie
With Donn�es
  .vgdCarBranche.MaxRows = NbBranches
  For i = 1 To 4  ' Ajout AV 22/02/2000 - V4.0.18 gestion du point d�cimal sur le Spread
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

Public Function nbP�riodes() As Integer
  nbP�riodes = colTrafics.count - nbP�riodesSatur�es
End Function
'******************************************************************************
' Affichage des donn�es du Giratoire
'*******************************************************************************
Private Sub Affiche()
  Dim i As Integer
  
  With Donn�es
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

  colBranches.Affiche Donn�es
  
  If nbP�riodes > 0 Then Donn�es.cboP�riode.ListIndex = 0 ' d�clenche l'�v�nement cboP�riode_Click

End Sub

'******************************************************************************
' Changement d'unit�s d'angle : Bascule DEGRE<-->GRADE
'*******************************************************************************
Public Sub ChangeUnit�Angle()
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
  
  'Libell�s d'angle
  Donn�es.lblAngleBranche.Caption = IDl_Angle & " (" & libelAngle(modeangle) & ")"
  'Borne sup�rieure de l'angle
  With Donn�es.vgdCarBranche
    .Col = 2
    For i = 1 To 8
      .Row = i
      If modeangle = DEGRE Then
        .TypeIntegerMax = 359
      Else
        .TypeIntegerMax = 399
      End If
    Next
    'Borne de l'�cart de l'angle
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
    .Affiche Donn�es
  End With
  
  GriserMenus True
End Sub

'******************************************************************************
' Activation du Giratoire : le giratoire devient le projet courant
'*******************************************************************************
Public Sub Activate()
  
  Set gbProjetActif = Me
  
    DessinGiratoire.gbDemiHauteur = Donn�es.DemiHauteur
    DessinGiratoire.gbDemiLargeur = Donn�es.DemiLargeur
    DessinGiratoire.gbFacteurZoom = Donn�es.FacteurZoom
    DessinGiratoire.gbRayonInt = R
    DessinGiratoire.gbRayonExt = R + LA + Bf
    DessinGiratoire.gbBandeFranchissable = Bf

End Sub

'******************************************************************************
' Enregistrement du Giratoire
'*******************************************************************************
Public Sub Enregistrer(ByRef Cancel As Integer)
Dim NomFich As String
  
  ' D�sactivation de la protection pour que le test qui suit ne concerne pas le programme en cours
  If numFichProt�g� <> 0 Then Close numFichProt�g�
  
  NomFich = Donn�es.dlgFichier.FileName
  If NomFich <> "" And Not FichierProt�g�(NomFich, , NomFich) Then
    Ecrire NomFich
  Else
    If numFichProt�g� <> 0 Then ' R�activation de la protection
      numFichProt�g� = FreeFile
      Open NomFich For Random Lock Read As numFichProt�g�
    End If
    EnregSous Cancel
  End If

End Sub

'******************************************************************************
' Enregistrement d'un Giratoire sous un autre nom
'*******************************************************************************
Public Sub EnregSous(ByRef Cancel As Integer)
  
  On Error GoTo TraitementErreur
  
  With Donn�es.dlgFichier
    .flags = cdlOFNOverwritePrompt Or cdlOFNHideReadOnly Or cdlOFNNoChangeDir Or cdlOFNPathMustExist
    .InitDir = App.Path
    .ShowSave
    If Not Cancel Then
      If FichierProt�g�(.FileName, , .FileName) Then
        Cancel = True
      Else
        Ecrire .FileName
        Donn�es.Caption = suppExt(.FileTitle)
        If Not R�sultats Is Nothing Then R�sultats.Caption = IDl_R�sultats & " - " & Donn�es.Caption
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
Dim v�hicule As Integer
Dim Chaine As String

' Maintenance : 28/11/2000 : Activation de la protection pour emp�cher que le fichier soit ouvert par qq d'autre
  If numFichProt�g� <> 0 Then Close numFichProt�g�
  
  ' Ouverture du Fichier
  numFich = FreeFile
  Open NomFich For Output Lock Write As numFich
  
  ' Donn�es g�n�rales
  DateModif = Date
  Donn�es.lblDateModif = DateModif
  Write #numFich, App.Title, "Version " & App.Major
'  Write #numFich, App.Title, "Version provisoire T0.2"
  Write #numFich, NomVariante, DateModif
  Write #numFich, nom, Milieu  ', Commentaire
  Write #numFich, NbBranches, nbP�riodes, modeangle
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
        Write #numFich, .Entr�eEvas�e, .LE4m, .LE15m, .LI, .LS
      End With
    Next
  End With
  
  'Trafics
  Write #numFich, "TRAFICS"
  With colTrafics
    For i = 1 To nbP�riodes
      With .Item(i)
        Write #numFich, .nom, .modeUVP, .CouleurCourbe
        Write #numFich, .FormaterQP()
        If .modeUVP Then
          For j = 1 To NbBranches
            Write #numFich, .FormaterQV(UVP, j)
          Next
        Else
          For v�hicule = VL To DEUXR
            For j = 1 To NbBranches
              Write #numFich, .FormaterQV(v�hicule, j)
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
  numFichProt�g� = FreeFile
  Open NomFich For Random Lock Read As numFichProt�g�
  
  gbProjetActif.Donn�es.FichierModifi� = False
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
Dim Angle As Single            ' Angle en degr�s
Dim Rampe As Boolean           ' Pr�sence d'une rampe > 3%
Dim Entr�eEvas�e As Boolean
Dim TAD As Boolean             ' Tourne � droite

Dim LE4m As Single             ' Largeur d'entr�e � 4m
Dim LE15m  As Single           ' Largeur d'entr�e � 15m
Dim LI As Single               ' Largeur de l'�lot
Dim LS As Single               ' Largeur de sortie

Dim nbTrafics As Integer
Dim CouleurCourbe As Long

Dim NomFich As String

  NomFich = MDIGirabase.dlgFichier.FileName
  
On Error GoTo TraitementErreur
  
  If FichierProt�g�(NomFich, , Titre:=NomFich, LectureSeuleAutoris�e:=True) Then Err.Raise 101
  
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
  
  'Caract�ristiques du projet
  Input #numFich, variable
  numLigne = numLigne + 1
  OkEntier variable, NbBranches, 3, 8
  Input #numFich, variable
  OkEntier variable, nbTrafics, 0
  Input #numFich, variable
  OkEntier variable, modeangle, DEGRE, GRADE
  
  'Caract�ristiques de l'anneau
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
      If i > 1 Then ' Calcul des �carts et v�rification des angles croissants
        .Item(i).Ecart = .Item(i).Angle - .Item(i - 1).Angle
        If .Item(i).Ecart <= 0 Then numLigne = numLigne - 1: Err.Raise 100
      End If
    Next
    'Supppression des branches inutiles (par d�faut le giratoire a �t� dimensionn� � 8 (gbCreFille)
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
  numFichProt�g� = FreeFile
  Open NomFich For Random Lock Read As numFichProt�g�
    
  D�sarmIndic
  Lire = True
  If gbFichierJournal Then Write #gbFichLog, MDIGirabase.dlgFichier.FileName
  Exit Function
  
TraitementErreur:
  gbErrLecFichier = True
  Select Case Err
  Case 100
    MsgBox IDm_ErrLectureFichier & " - " & IDm_ligne & " " & CStr(numLigne)
    Close numFich
    D�sarmIndic
  Case 101
  Case Else
    ErreurFatale
  End Select
End Function

Private Sub D�sarmIndic()
  ' Les indicateurs ont pu �tre arm�s pendant la lecture
  Donn�es.Donn�eModifi�e = False
  Donn�es.FichierModifi� = False
End Sub

'******************************************************************************
' Calcul de largeur d'ilot maximum
'*******************************************************************************
Public Sub MajCompl�ment(Optional ByVal NouveauMilieu As Integer = -1)
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
    
    ' Ref  �2.1.1

  ' Si R Et Bf=0 : invalide(le lancement du calcul sera emp�ch�), mais accept�s provisoirt en saisie d'o� On error...
  On Error Resume Next
  ' gbCoefLEU vaut 1.2 (1.1 en version belge)
  LEU = LAU / (gbCoefLEU * (1 + 1 / 2 / RU))

' Ref  �2.1.4
  LImax = Tg * Sqr(RU + LAU / 2)

End Sub

'********************************************************************************************************************
' Ajout d'une p�riode de trafic
' Cette proc�dure peut �tre d�clench�e de 3 mani�res
'   Saisie d'un nouveau nom dans la combo cboP�riode                      --> DrapeauMenu = False
'   Appel du menu                                                         --> DrapeauMenu = True
'   Appel depuis une autre proc�dure (Inversion - Multiplication - Import)--> DrapeauMenu = Nom de la p�riode � cr�er
' Dans ce dernier cas, le traitement final est diff�r� (sera r�alis� par TRAFIC.Dupliquer)
'********************************************************************************************************************
Public Function newP�riode(ByVal DrapeauMenu As Variant) As Boolean
Dim nomP�riode As String
Dim i As Integer

  Select Case DrapeauMenu
  Case False     ' Appel depuis la ComboBox
    nomP�riode = Donn�es.cboP�riode.Text
  Case True      ' Appel depuis le Menu
    nomP�riode = InputBox(IDl_NouvelleP�riode, App.Title, colTrafics.defautNew, , , App.HelpFile, IDhlp_NewP�riode)
  Case Else      ' Appel avec passage du nouveau nom (Inversion - Multiplication - Import)
                 ' Pour inversion, on pourra ajouter un bouton d'aide si le CERTU le sp�cifie (code 204 : IDhlp_InversP�riode)
    If InStr(DrapeauMenu, IDl_Inversion) <> 0 Then
      nomP�riode = InputBox(IDl_NouvelleP�riode, App.Title, DrapeauMenu, , , App.HelpFile, IDhlp_InversP�riode)
    ElseIf InStr(DrapeauMenu, IDl_Multiplication) <> 0 Then
      nomP�riode = InputBox(IDl_NouvelleP�riode, App.Title, DrapeauMenu, , , App.HelpFile, IDhlp_MultP�riode)
    Else
      nomP�riode = InputBox(IDl_NouvelleP�riode, App.Title, DrapeauMenu)
    End If
  End Select

  If nomP�riode <> "" And Not colTrafics.EnDouble(nomP�riode) Then
    ' Ajout effectif
    colTrafics.add Me, nomP�riode
    ' Traitement final : la p�riode cr��e devient la p�riode courante
    If VarType(DrapeauMenu) = vbBoolean Then
    ' Sinon, l'appel est fait depuis une autre fonction : traitement final diff�r�
      With Donn�es
        colTrafics.Item(colTrafics.count).modeUVP = .lblTraficUVP.Visible
        With .cboP�riode
          .ListIndex = .ListCount - 1     ' d�clenche l'�v�nement cboP�riode_Click
        End With
        ' La ligne suivante doit en fait �tre appel�e lors de l'activation de l'onglet (+ messages � porter lors des modifs dans vgdLargBranche)
        colBranches.BlocageTrafic Donn�es
      End With
    End If
    
    newP�riode = True
    GriserMenus True
  End If

End Function

'******************************************************************************
' Suppression d'une p�riode de trafic
'*******************************************************************************
Public Sub delP�riode(ByVal wTrafic As TRAFIC)
Dim nomP�riode As String
Dim numP�riode As Integer

  With wTrafic
    If MsgBox(IDm_SupprP�riode & " " & .nom, vbYesNo + vbDefaultButton2 + vbQuestion + vbMsgBoxHelpButton, , App.HelpFile, IDhlp_DelP�riode) = vbYes Then
      With Donn�es.cboP�riode
        numP�riode = .ListIndex
        colTrafics.Remove Me, numP�riode + 1
        .ListIndex = Min(numP�riode, nbP�riodes - 1)  ' d�clenche l'�v�nement cboP�riode_Click
      End With
      If nbP�riodes = 0 Then
        Donn�es.cboP�riode_Click  ' l'�v�nement n'a pas �t� d�clench� (ListIndex=-1)
        If Not R�sultats Is Nothing Then Unload R�sultats
      End If
      GriserMenus True
    End If
  End With
  
End Sub

'******************************************************************************
' Renommer une p�riode de trafic
'*******************************************************************************
Public Sub renameP�riode(ByVal wTrafic As TRAFIC)
Dim nomP�riode As String
Dim numP�riode As Integer

  With wTrafic
    nomP�riode = InputBox(IDl_RenP�riode, App.Title, .nom, , , App.HelpFile, IDhlp_RenameP�riode)
    If nomP�riode <> "" And Not colTrafics.EnDouble(nomP�riode) Then
      ' Renommage effectif
      .nom = nomP�riode
      With Donn�es.cboP�riode
        numP�riode = .ListIndex
        .RemoveItem numP�riode
        .AddItem nomP�riode, numP�riode
        .ListIndex = numP�riode
      End With
      If Not R�sultats Is Nothing Then
        With R�sultats.cboP�riode
          .RemoveItem numP�riode
          .AddItem nomP�riode, numP�riode
        End With
      End If
      Donn�es.FichierModifi� = True
    End If
  End With
  
End Sub

'******************************************************************************
' Inversion des matrices d'une p�riode de trafic
'*******************************************************************************
Public Sub inversP�riode(ByVal wTrafic As TRAFIC)

  If wTrafic.EstCompl�te Then
    If newP�riode(IDl_Inversion & IDl_DE & wTrafic.nom) Then
      wTrafic.Inverser
    End If
  Else
    MsgBox IDm_P�riodeIncompl�te, vbExclamation + vbMsgBoxHelpButton, App.Title, App.HelpFile, IDhlp_InversP�riode
  End If
End Sub

'******************************************************************************
' Multiplication des matrices d'une p�riode de trafic
'*******************************************************************************
Public Sub multP�riode(ByVal wTrafic As TRAFIC)

  If wTrafic.EstCompl�te Then
    If newP�riode(IDl_Multiplication & IDl_DE & wTrafic.nom) Then
      Set frmTrafMult.TraficOrigine = wTrafic
      frmTrafMult.Show vbModal
    End If
  Else
    MsgBox IDm_P�riodeIncompl�te, vbExclamation + vbMsgBoxHelpButton, App.Title, App.HelpFile, IDhlp_MultP�riode
  End If
End Sub

'****************************************************************************************************
' V�rification de la compatibilit� du projet import� avec le giratoire, en vue de l'import de matrice
'****************************************************************************************************
Public Function VerifCompatible(wGiratoire As GIRATOIRE) As Boolean
Dim i As Integer

  If wGiratoire.nbP�riodes = 0 Then
    MsgBox IDm_PasDeP�riode, vbExclamation: Exit Function
  ElseIf wGiratoire.NbBranches <> NbBranches Then
    MsgBox IDm_NbBranchesDiff�rent, vbExclamation + vbMsgBoxHelpButton, , App.HelpFile, IDhlp_ImportMatrice: Exit Function
  Else
    For i = 1 To NbBranches
      If colBranches.Item(i).Entr�eNulle <> wGiratoire.colBranches.Item(i).Entr�eNulle Then
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
' Suppression des p�riodes satur�es
'*******************************************************************************
Public Sub delP�riodeSatur�es()
Dim i As Integer
  With colTrafics
    For i = .count To 1 Step -1
      If .Item(i).BrancheSatur�e <> 0 Then
        .Remove Me, i
      End If
    Next
  End With

End Sub

'******************************************************************************
' Calculs de capacit� : Toutes les p�riodes compl�tes
'*******************************************************************************

Public Sub CalculCapacit�(Optional ByVal wTrafic As TRAFIC = Nothing)
Dim i As Integer
Dim j As Integer

' Appel pourle calcul d'une seule p�riode
  If Not wTrafic Is Nothing Then
    With wTrafic
      If .EstCompl�te Then
        .CalculTraficEntrant colBranches
        For j = 1 To NbBranches
          .CalculCapacit� j, colBranches
        Next
        .CalculFait = True
      End If
    End With
    Exit Sub
  End If
  
' Appel pour le calcul de l'ensemble
  If R�sultats Is Nothing Then
    ChargeFeuilleR�sultat
  Else
    R�sultats.SetFocus
    ' Suppression des p�riodes avec saturation de branche
    delP�riodeSatur�es
  End If
    
  CalculParamGiratoire
  For i = 1 To colBranches.count
    colBranches.Item(i).CalculParamBranche
  Next
  
  For i = 1 To colTrafics.count
    CalculCapacit� colTrafics.Item(i)
  Next
  
  CalculFait = True
  GriserMenus True
 ' Affichage des r�sultats de la p�riode active
  R�sultats.cboP�riode.ListIndex = Donn�es.cboP�riode.ListIndex
 
End Sub

'******************************************************************************
' Coefficients du Giratoire : Tg,Te,Tf1 - RU,RAU,LEU,LImax - KE, KI
'******************************************************************************
Private Sub CalculParamGiratoire()
   
  ' D�termination des coefficients pour le calcul des capacit�s : Te, Tg, Tf1
  
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
   
  ' D�termination des coefficients de g�ne KE (trafic tournant int�rieur), KI (trafic tournant ext�rieur)
  
'  ' Ref : Modificatif 29/12/98
'  RU = R + 0.5 * Bf
'  LAU = LA + 0.5 * Bf
'
'  ' Ref  �2.1.1
'  LEU = LAU / (1.2 * (1 + 1 / 2 / RU))
'
'  ' Ref  �2.1.4
'  LImax = Tg * Sqr(RU + LAU / 2)
  
    ' Ref  �2.2.1 modifi� par R�ponse.Doc du 12/03/99
'  KI = 160 / (LAU * (RU + LAU))
  KI = 8 / LAU * Sqr(20 / (RU + LAU))
  If LAU > 8 Then
'    KE = 1 - (RU + LAU) ^ 2 * (LAU - 8) / LAU
    KE = 1 - (RU / (RU + LAU)) ^ 2 * (LAU - 8) / LAU
'    If gbFichierJournal Then Write #gbFichLog, "2.2.2 - KE calcul� avec LAU >8"
 Else
    KE = 1
'    If gbFichierJournal Then Write #gbFichLog, "2.2.2 - KE =1 --> LAU <8"
  End If
  
  'Ref �2.3.3
  KI = Min(KE, KI)
  
End Sub

Private Sub ChargeFeuilleR�sultat()
Dim i As Integer

  Set R�sultats = New frmR�sultats
  Set R�sultats.GiratoireProjet = Me

End Sub

' Maintenance : 28/11/2000 : Activation de la protection pour emp�cher que le fichier soit ouvert par qq d'autre
Private Sub Class_Terminate()
  If numFichProt�g� <> 0 Then Close numFichProt�g�
End Sub
