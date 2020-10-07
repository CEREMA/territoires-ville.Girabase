Attribute VB_Name = "GirabaseMain"
'******************************************************************************
'*
'*          Projet GIRABASE - CERTU - CETE de l'Ouest
'*
'*          Module standard : GIRABASEMAIN.BAS
'*
'*          Définition des variables globales - Initialisation du programme
'*
'******************************************************************************

Option Explicit
'****************************************

  ' Version belge : Girawal
#Const GIRAWAL = True

'Variables littérales de Girabase


'******************************************************************************
  '--------------- Constantes numériques
'Code pour une donnée non définie
Public Const DONNEE_INEXISTANTE = -1

' Largeur des colonnes de la grille de trafic
'Public Const WIDTHCOLTRAFIC = 600
Public Const WIDTHCOLTRAFIC = 555
' Nombre de branches par défaut en initialisation
Public Const DEFAUTNBBRANCHES = 4
' Nombre de fichiers mémorisés
Public Const MAXFICH = 4
Public Const DEFAUT_LI = 3 'Valeur par défaut de la largeur d'ilot
' Angles
Public Const DEGRE = 0
Public Const GRADE = 1
Public Const CVRADIAN = True

  '   RC = Rase Campagne  -  PU = Périurbain  -  CV = Centre-ville
Public Const rc = 0
Public Const PU = 1
Public Const CV = 2
  
' Trafics
Public Const VEHICULE = 0
Public Const PIETON = 1

Public Const VL = 0
Public Const PL = 1
Public Const DEUXR = 2
Public Const UVP = 3

'Constante définissant les types de données
Public Const TYPE_AVANT = "X"    'Pour vérifier le dernier contrôle
Public Const TYPE_COURANT = ""   'Pour vérifier le contrôle en cours
Public Const TYPE_AUCUN = "0"    'Aucune donnée à contrôler
Public Const TYPE_MATRICE = "M"  'Matrice
Public Const TYPE_NBBRANCHES = "NbBranches"
Public Const TYPE_MILIEU = "Milieu"
Public Const TYPE_R = "R"
Public Const TYPE_BF = "Bf"
Public Const TYPE_LA = "LA"
Public Const TYPE_RG = "Rg"
Public Const TYPE_QP = "QP"
Public Const TYPE_Q = "Q"

Public Const TYPE_LE4M = "LE4m"
Public Const TYPE_LE15M = "LE15m"
Public Const TYPE_LI = "LI"
Public Const TYPE_LS = "LS"
Public Const TYPE_ENTREE = "EV"
Public Const TYPE_ANGLE = "Angle"

'Constantes pour identifier la matrice en cours de saisie (TypeMatriceActive)
Public Const AUCUN = 0
Public Const BRANCHE = 1
Public Const DIMENSION = 2
Public Const TRAFIC = 3

' Constantes d'aide en ligne
Public Const HELPNAME = "Girabase.chm"

Public Const IDhlp_Nouveau = 204 'ch02s04.html
Public Const IDhlp_Ouvrir = 203 'ch02s03.html
Public Const IDhlp_ImportMatrice = 224 'ch02s04s04s09
Public Const IDhlp_ConfigImprimante = 232
Public Const IDhlp_Imprimer = 232 'ch02s06
Public Const IDhlp_ImprimerDonnées = 233 'ch02s06s01
Public Const IDhlp_ImprimerRésultats = 236 'ch02s06s02

Public Const IDhlp_OngletSite = 205 'ch02s04s01         205
Public Const IDhlp_OngletDimensionnement = 210 ' ch02s04s02         210
Public Const IDhlp_OngletTrafic = 215 'ch02s04s04         215
Public Const IDhlp_CarBranche = 209 ' ch02s04s01s04      209

Public Const IDhlp_NewPériode = 219 'ch02s04s04s04      219
Public Const IDhlp_DelPériode = 220 'ch02s04s04s05      220
Public Const IDhlp_RenamePériode = 221 ' ch02s04s04s06      221
Public Const IDhlp_InversPériode = 222 'ch02s04s04s07      222
Public Const IDhlp_MultPériode = 223 'ch02s04s04s08      223
Public Const IDhlp_DiagramFlux = 225 ' ch02s04s04s10      225

Public Const IDhlp_OngletFonctionnement = 226 ' ch02s05            226
Public Const IDhlp_OngletPériode = 226 'ch02s05            226
Public Const IDhlp_OngletConception = 226 'ch02s05            226

Public Const IDhlp_Graphique = 0
  
Public Const IDhlp_Résultats = 226 ' ch02s05            226
Public Const IDhlp_SaturerBranche = 231 ' ch02s05s04s01      231
Public Const IDhlp_Courbes = 230 ' ch02s05s04         230
  
  ' Coefficients utilisés dans le calcul de capacité (Ref : NOTE DE CALCUL - §1.3)
  '   Te = Largeur d'entrée
Public Const gbTeRC As Single = 0.7
Public Const gbTePU As Single = 0.8
Public Const gbTeCV As Single = 0.85
  '   Tg = Créneau critique
#If GIRAWAL Then
  Public Const gbTgRC As Single = 4.9
  Public Const gbTgPU As Single = 4.1
  Public Const gbTgCV As Single = 4#
#Else
  Public Const gbTgRC As Single = 4.75
  Public Const gbTgPU As Single = 4.55
  Public Const gbTgCV As Single = 4.4
#End If

  '   Tf1 = Créneau complémentaire
#If GIRAWAL Then
  Public Const gbTf1RC As Single = 2.7
  Public Const gbTf1PU As Single = 2.5
  Public Const gbTf1CV As Single = 2.2
#Else
  Public Const gbTf1RC As Single = 2.25
  Public Const gbTf1PU As Single = 2.05
  Public Const gbTf1CV As Single = 1.8
#End If

#If GIRAWAL Then
  Public Const gbCoefLEU = 1.1
#Else
  Public Const gbCoefLEU = 1.2
#End If

'******************************************************************************
'Variables globales
'******************************************************************************
Public gbFichierJournal As Boolean
Public gbFichLog As Integer
Public gbFinPrématurée As Boolean
Public gbVersionDemo As Boolean
Public gbVersionDéveloppeur As Boolean
Public gbNomfich As String

' Suppression AV : 22/12/2000 - Inutilisé
'Public gbEpaisseurLigne
  
Public gbProjetActif As GIRATOIRE

Public gbGiratoires As New Giratoires

Public gbErrLecFichier As Boolean

' Derniers fichiers GIRATOIRE utilisés
'-----------------------------------------------------------
Public gbMRUFichiers() As String
Public gbNbMRUFichiers As Integer

' Unités d'angle
'--------------------
Public PI As Single
Public libelAngle(0 To 1) As String
Public eqvPI(0 To 1) As Integer ' équivalent pi selon le mode

' Point décimal selon le système utilisé
'---------------------------------------
Public gbPtDecimal As Integer

' Coefficients de trafic
'-----------------------
Public gbCoefVehicule(0 To 2) As Single

Public gbTe(0 To 2) As Single
Public gbTg(0 To 2) As Single
Public gbTf1(0 To 2) As Single

' Fonctions de  l'API Windows
'--------------------------------------
' Compilation conditionnelle Win32
#If Win32 Then
  Public Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
  End Type
  Public Declare Sub GetLocalTime Lib "Kernel32" (lpSystemTime As SYSTEMTIME)
#End If
 
Public Sub Main()


 ' Décryptage de la ligne de commande
  GetCommandLine
 
  ' Initialisation des variables globales
  gbIniVar
'  gbVersionDemo = True '   : peut-être n'y en aura-t-il pas pour GIRABASE
'  gbFichierJournal = InStr(UCase(Command), "/JOURNAL") <> 0 ' And Not gbVersionDemo
  
  If gbFichierJournal Then
    gbFichLog = FreeFile
    Open App.Path & "\girabase.log" For Output As gbFichLog
    Write #gbFichLog, "Chargemement de MDIGirabase"
  End If

  On Error GoTo TraitementErreur
  
'********************************
'test Protection
'********************************
  'Type de protection
        TYPPROTECTION = CPM
  ' Vérification de l'enregistrement
  If ProtectCheck("its00+-k") = "its00+-k" Then
    ' Affichage de la feuille principale
    MDIGirabase.Show
  Else 'la licence n'a pas été validée on ferme
     End
  End If
'********************************
  
  If gbFichierJournal Then
    DoEvents
    Write #gbFichLog, "Fin de chargemement de MDIGirabase"
    If Not gbVersionDéveloppeur Then Write #gbFichLog, "Appel de la protection"
  End If
  
  ' Initialisation de l'imprimante (NT ou non)
  Set gDlgPrint = MDIGirabase.dlgImprimer
  InitializePrinter
  
  If gbFichierJournal And Not gbVersionDéveloppeur Then
    Write #gbFichLog, "Retour de la protection"
  End If
    
  If ExistFich(gbNomfich) Then
    MDIGirabase.dlgFichier.FileName = gbNomfich
    gbCreFille gbNomfich
  End If
    
  Exit Sub
  
TraitementErreur:
  Resume Next
  
  
End Sub

'*************************************************************************************
' Initialisation des variables globales
'*************************************************************************************

Private Sub gbIniVar()

LireRegistry

  libelAngle(DEGRE) = IDl_Degrés
  libelAngle(GRADE) = IDl_Grades
  eqvPI(DEGRE) = 180
  eqvPI(GRADE) = 200
  PI = 4 * Atn(1)
  ' Détecte si le Point Décimal est une virgule ou un point, sur le système utilisateur
  'If IsNumeric("1.1") Then gbPtDecimal = Asc(".") Else gbPtDecimal = Asc(",")
  ReTrouvePtDecimal
  
  gbTe(rc) = gbTeRC
  gbTe(PU) = gbTePU
  gbTe(CV) = gbTeCV
  gbTg(rc) = gbTgRC
  gbTg(PU) = gbTgPU
  gbTg(CV) = gbTgCV
  gbTf1(rc) = gbTf1RC
  gbTf1(PU) = gbTf1PU
  gbTf1(CV) = gbTf1CV
  
  
  gbCoefVehicule(VL) = 1
  gbCoefVehicule(PL) = 2
  gbCoefVehicule(DEUXR) = 0.5
    
End Sub

'*************************************************************************************
'Mise en grisés des menus selon le contexte
'   Etat = Drapeau - si True, activation des options de menu
'*************************************************************************************
Public Sub GriserMenus(ByVal Etat As Boolean)
Dim i As Integer
    
  With MDIGirabase
    If Etat Then .mnuSite(5).Caption = IDm_MenuAngle & " " & libelAngle((gbProjetActif.modeangle + 1) Mod 2)
    
    For i = 2 To 4    ' Fermer - Enregistrer - Enregistrer sous
      .mnuFichier(i).Enabled = Etat
    Next
                      ' Importer une matrice
    .mnuFichier(6).Enabled = Etat
                      ' Imprimer
    .mnuFichier(9).Enabled = Etat
    
    If gbProjetActif Is Nothing Then Etat = False
    For i = 1 To 4
                      ' Site - Trafic - Capacité - Fenêtre
      .mnuBarre(i) = Etat
    Next
    
    If Etat Then
                    ' Enregistrer - Enregistrer sous
      .mnuFichier(3) = gbProjetActif.Données.FichierModifié
      .mnuFichier(4) = gbProjetActif.Milieu <> -1
                      ' Imprimer
      .mnuFichier(9).Enabled = gbProjetActif.Milieu <> -1 And Printers.count > 0
      
      For i = 1 To 2
                      ' Site - Trafic
        .mnuBarre(i) = gbProjetActif.Milieu <> -1
      Next

                    ' Editer une branche
      For i = 0 To 7
        .mnuBranche(i).Visible = (i < gbProjetActif.NbBranches)
      Next
      
      Etat = gbProjetActif.colTrafics.count > 0
      .mnuResult(0) = Etat
         
      On Error Resume Next
                    ' Options de trafic (sauf Nouveau)
      For i = 1 To 10
        .mnuTrafic(i) = Etat
      Next
                    
    End If
    
    .tbrFile.Buttons("btnSave").Enabled = .mnuFichier(3)
    .tbrFile.Buttons("btnPrint").Enabled = .mnuFichier(9)
    
  End With
    
    
End Sub

'*************************************************************************************
' Affichage dans le menu des derniers fichiers utilisés (MRU)
'*************************************************************************************
Public Sub MRUmenu(ByVal NomFich As String)
Dim i As Integer, j As Integer
  
  If Not ExistFich(NomFich) Then Exit Sub
  
  For i = 0 To gbNbMRUFichiers - 1
    If UCase(NomFich) = UCase(gbMRUFichiers(i)) Then
      ' Le fichier est déjà dans la liste, on va le remettre en tête
      For j = i To gbNbMRUFichiers - 2
      ' Resserrement du tableau en supprimant le fichier
        gbMRUFichiers(j) = gbMRUFichiers(j + 1)
      Next
      gbNbMRUFichiers = gbNbMRUFichiers - 1
    End If
  Next

  If gbNbMRUFichiers < MAXFICH Then    ' Debug: cette valeur pourra être paramétrée (n'apparait qu'ici), mais sans dépasser la taille de mnuSelect
    MDIGirabase.mnuSelect(gbNbMRUFichiers).Visible = True
    If gbNbMRUFichiers = 0 Then MDIGirabase.mnuSelect(4).Visible = True 'Création de la ligne de séparation
    ReDim Preserve gbMRUFichiers(0 To gbNbMRUFichiers)
    gbNbMRUFichiers = gbNbMRUFichiers + 1
  End If
  
  For i = gbNbMRUFichiers - 1 To 1 Step -1
    ' Décalage des fichiers pour insérer le nouveau en tête
    gbMRUFichiers(i) = gbMRUFichiers(i - 1)
    MDIGirabase.mnuSelect(i).Caption = "&" & CStr(i + 1) & " " & gbMRUFichiers(i)
  Next
  gbMRUFichiers(0) = NomFich
  MDIGirabase.mnuSelect(0).Caption = "&1 " & NomFich
  
  End Sub

'*************************************************************************************
' Configuration de l'imprimante
'*************************************************************************************
Public Function ConfigImprimante() As Boolean
    On Error GoTo ErrImpr
    
    With MDIGirabase.dlgImprimer
      .flags = cdlPDPrintSetup ' Pour afficher directement la fenêtre Configuration
      .Orientation = Printer.Orientation
      .ShowPrinter
      If Printer.Orientation <> .Orientation Then
        MsgBox "Le changement d'orientation ne peut être pris en compte depuis " & App.Title
        .Orientation = Printer.Orientation
      End If
    End With
    ConfigImprimante = True
    Exit Function
    
ErrImpr:
  If Err = cdlCancel Then
  ' L'utilisateur a fait 'Annuler
  Else
    ErreurFatale
  End If
  Exit Function

End Function

'*************************************************************************************
' Création d'une feuille fille frmDonnées
'*************************************************************************************
'Public Sub gbCreFille(ByVal Nom As String, ByVal Nouveau As Boolean)
Public Sub gbCreFille(ByVal nom As String)

' En création,  Nom="" - Nouveau = True
' En ouverture, Nom représente le nom du  fichier .GBS du giratoire - Nouveau = False
Dim GiratoireProjet As GIRATOIRE
Dim nbGiratoires As Integer
Dim NbBranches As Integer
Dim modeangle As Integer
Dim Nouveau As Boolean
Dim wGiratoire As GIRATOIRE
   
  Nouveau = (nom = "")
  If Nouveau Then
    NbBranches = frmParam.txtNbBranches
    modeangle = Numopt(frmParam.optUnité)
    Unload frmParam
  Else  'Avantlecture : dimensionnement du Giratoire au Maximum et unité d'angle par défaut
'    For Each wGiratoire In gbGiratoires
'      With wGiratoire.Données
'        If .dlgFichier.FileName = nom Then
'          .WindowState = Screen.ActiveForm.WindowState
'          .SetFocus
'          Exit Sub
'        End If
'      End With
'    Next
    NbBranches = 8
    modeangle = DEGRE
  End If
  
  With gbGiratoires
    Set GiratoireProjet = .add(NbBranches, modeangle)
    nbGiratoires = .count
    Set GiratoireProjet.Données = New frmDonnées
  End With

  With GiratoireProjet.Données
    Set .GiratoireProjet = GiratoireProjet
    .Nouveau = Nouveau
    .ChargementEnCours = True
    If Nouveau Then
      .Caption = IDl_Giratoire & CStr(nbGiratoires)    ' Déclenche Form_Load
    Else
      .Caption = nomCourt(nom)
      MDIGirabase.dlgFichier.FileName = ""  ' Réinitialisation pour l'ouverture suivante
      If gbErrLecFichier Then
        ' Echec de la lecture
        Unload GiratoireProjet.Données
        gbErrLecFichier = False
        Exit Sub
      End If
      .dlgFichier.FileName = nom
      MRUmenu nom     'MRUmenu est défini dans GIRABASEMAIN.BAS
    End If
  
    .Form_Activate
    .ChargementEnCours = False
  End With

End Sub

'*************************************************************************************
' Import de matrice
' Fonction dérivée de gbCreFille
'*************************************************************************************
Public Sub ImportMatrice(ByVal nom As String)

' Nom représente le nom du  fichier .GBS du giratoire
Dim GiratoireProjet As GIRATOIRE
Dim NbBranches As Integer
   
  NbBranches = 8
  
  ' Création d'un objet GIRATOIRE temporaire
  Set GiratoireProjet = gbGiratoires.add(NbBranches, DEGRE)
  
  Set GiratoireProjet.Données = frmImport
  Set frmImport.GiratoireProjet = GiratoireProjet
  
  If GiratoireProjet.Lire Then
    If gbProjetActif.VerifCompatible(GiratoireProjet) Then
      With GiratoireProjet.Données
        .Caption = .Caption & suppExt(nom)
        .cboPériode.ListIndex = 0
        .Show vbModal
      End With
    End If
  Else
    ' Echec de la lecture
  End If

  ' Suppression du l'objet GIRATOIRE temporaire
  gbGiratoires.Remove gbGiratoires.count
  Unload frmImport
  MDIGirabase.dlgFichier.FileName = ""  ' Réinitialisation pour l'ouverture suivante

End Sub

Private Sub LireRegistry()
  ' Doc Manuel Utilisateur : p351 et sq

' Variant devant contenir un tableau à deux dimensions renvoyé par la fonction GetAllSettings
Dim MySettings As Variant
' Nombre entier correspondant à la valeur du compteur.
Dim intSettings As Integer

Dim Setting As Variant

' Extrait les paramètres.
MySettings = GetAllSettings(Appname:=App.Title, Section:="Recent Files")
If Not IsEmpty(MySettings) Then
  For intSettings = UBound(MySettings, 1) To 0 Step -1     ' On les lit  à l'envers, car MRUmenu les ajoute par décalage
    MRUmenu MySettings(intSettings, 1)
  Next
End If

' Les lignes qui suivent ne sont pas utilisées dans la version actuelle mais çà réserve l'avenir
' Unités d'angle par défaut (si Nouveau)
Setting = GetSetting(Appname:=App.Title, Section:="Options", Key:="ModeAngle", Default:=DEGRE)
SaveSetting Appname:=App.Title, Section:="Options", Key:="ModeAngle", Setting:=Setting
' Nombre de branches par défaut (si Nouveau)
Setting = GetSetting(Appname:=App.Title, Section:="Options", Key:="NbBranches", Default:=DEFAUTNBBRANCHES)
SaveSetting Appname:=App.Title, Section:="Options", Key:="NbBranches", Setting:=Setting
Setting = GetSetting(Appname:=App.Title, Section:="Options", Key:="MaxMRUList", Default:=MAXFICH)
SaveSetting Appname:=App.Title, Section:="Options", Key:="MaxMRUList", Setting:=Setting
' Suppression AV : 22/12/2000 - Inutilisé
'Setting = GetSetting(Appname:=App.Title, Section:="Options", Key:="EpaisseurLigne", Default:=1)
'SaveSetting Appname:=App.Title, Section:="Options", Key:="EpaisseurLigne", Setting:=Setting
'gbEpaisseurLigne = CInt(Setting)

End Sub


Public Sub lireProtect()

'    If gbVersionDemo Then Exit Sub
'    Exit Sub
    
    'Do While MyCCToujours() <> 0
    '  Select Case MsgBox("Protection retirée", vbRetryCancel + vbCritical, "Gestion de la protection")
    '  Case vbCancel
    '    gbFinPrématurée = True
     '   Exit Do
    '  End Select
    'Loop
    
    If gbFinPrématurée Then Unload MDIGirabase

End Sub

'*************************************************************************************
' Ecriture d'un article dans le fichier journal
'*************************************************************************************
Public Sub Journal(ByVal MotClé As String, ParamArray Arguments())
Dim NomControle As String
Dim controle As Control

  If Not gbFichierJournal Then Exit Sub
  
  If MotClé = "Menu" Then
    Print #gbFichLog, "Menu : ", Arguments(UBound(Arguments))
    Exit Sub
  End If
  
  If MotClé = "Activation" Then
    Print #gbFichLog, "Activation feuille : ", Screen.ActiveForm.Caption
    Exit Sub
  End If
  
  Set controle = Screen.ActiveForm.ActiveControl
  If controle Is Nothing Then Exit Sub
  
  NomControle = Mid(controle.Name, 4)
  Print #gbFichLog, NomControle, MotClé
  
  If TypeOf controle Is TextBox Then
    Print #gbFichLog, controle
  ElseIf TypeOf controle Is ComboBox Then
    Print #gbFichLog, "Index : ", controle.ListIndex, controle
  ElseIf TypeOf controle Is ComboBox Then
    Print #gbFichLog, controle
  ElseIf TypeOf controle Is OptionButton Then
    Print #gbFichLog, "Option numéro : ", controle.Index, controle.Caption
  ElseIf TypeOf controle Is SSTab Then
    Print #gbFichLog, controle.Tab, controle.Caption
  ElseIf TypeOf controle Is vaSpread Then
    If MotClé = "LeaveCell" Then
      Print #gbFichLog, "Ligne : ", controle.Row, "Colonne : ", controle.Col, "Valeur : ", controle.Value
      Print #gbFichLog, "Nouvelle ligne : ", Arguments(0), "Nouvelle colonne : ", Arguments(1)
      ElseIf MotClé = "Cancel" Then
      Print #gbFichLog, "Valeur refusée"
    ElseIf NomControle = "Trafic" Then
      Write #gbFichLog, "Index : ", controle.Index
    End If
  End If

End Sub

'*************************************************************************************
' Décryptage de la ligne de commande
'*************************************************************************************
Private Function GetCommandLine(Optional MaxArgs As Variant)
' Déclare les variables.
  Dim C As String * 1
  Dim CmdLine As String
  Dim CmdLnLen As Integer
  Dim InArg As Boolean
  Dim i As Integer, NumArgs As Integer
  Dim ArgArray() As String
  Dim GuillemetOuvert As Boolean
 
 ' Vérifie si MaxArgs a été spécifié.
  If IsMissing(MaxArgs) Then MaxArgs = 10
' Définit un tableau au format approprié.
  ReDim ArgArray(1 To MaxArgs)
  NumArgs = 0: InArg = False
' Récupère les arguments de ligne de commande.
  CmdLine = Command()
  CmdLnLen = Len(CmdLine)
  
  If ExistFich(CmdLine) Then gbNomfich = CmdLine: Exit Function
  
' Analyse de la ligne de commande caractère par caractère.
  For i = 1 To CmdLnLen
    C = Mid(CmdLine, i, 1)
' Analyse de caractères d'espacement ou de tabulations.
    If (C <> " " And C <> vbTab) Or GuillemetOuvert Then
      ' Ni espace ni tabulation.
    ' Vérifie une éventuelle présence dans l'argument.
      If Not InArg Then
        ' Le nouvel argument commence.
        ' Vérifie si les arguments ne sont pas trop nombreux.
        If NumArgs = MaxArgs Then Exit For
        NumArgs = NumArgs + 1
        InArg = True
      End If
      ' Concatène un caractère à l'argument courant.
      If C <> Chr(34) Then
        ArgArray(NumArgs) = ArgArray(NumArgs) & C
      Else
        GuillemetOuvert = Not GuillemetOuvert
      End If
    Else
      ' Recherche un espace ou une tabulation.
      ' L'indicateur InArg prend la valeur False.
      InArg = False
      If C = Chr(34) Then GuillemetOuvert = Not GuillemetOuvert
    End If
  Next i

' Analyse logique des paramètres de la ligne de commande
  For i = 1 To NumArgs
    If UCase(ArgArray(i)) = "/JOURNAL" Then
      gbFichierJournal = True
    ElseIf UCase(ArgArray(i)) = "/D" Then
      If i < NumArgs And UCase(ArgArray(i + 1)) = "CETE44" Then gbVersionDéveloppeur = True
    ElseIf InStr(LCase(ArgArray(i)), ".gbs") <> 0 Then
      gbNomfich = ArgArray(i)
    End If
  Next

' Redimensionne le tableau pour qu'il puisse
' juste contenir les arguments.
'  ReDim Preserve ArgArray(1 To NumArgs)
' Renvoie le tableau dans le nom de fonction.
  GetCommandLine = ArgArray()
  
End Function
