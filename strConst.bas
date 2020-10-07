Attribute VB_Name = "strConst"
'**************************************************************************************
'          Projet GIRABASE - CERTU - CETE de l'Ouest

'     Mise à jour pour la version anglaise : Décembre 2000
'
'   Réalisation : André VIGNAUD

'   Module standard : strConst   -   STRCONST.BAS

'   Fonctions du module
'     Constantes de chaine à traduire

'**************************************************************************************
Option Explicit

' Constantes utilisées dans plusieurs modules
'--------------------------------------------
Public Const IDl_Version = "Version"

Public Const IDm_Obligatoire = "Saisie obligatoire"

Public Const IDl_ET = " et "
Public Const IDl_DE = " de "
Public Const IDl_VERS = " vers "

Public Const IDl_Giratoire = "Giratoire"
Public Const IDl_Branche = "Branche "
Public Const IDl_Période = "Période"
Public Const IDl_Angle = "Angle"
Public Const IDl_DeLaPériode = " de la " & IDl_Période ' Conserver les blancs en début et fin de chaine
Public Const IDl_LaBranche = "la " & IDl_Branche ' Conserver les blancs en début et fin de chaine
Public Const IDl_DeLaBranche = " de la " & IDl_Branche ' Conserver les blancs en début et fin de chaine
Public Const IDl_Multiplication = "Multiplication"    ' GIRATOIRE - TrafMult
  

'Titre de la fenêtre Résultats
Public Const IDl_Résultats = "Résultats"

Public Const IDm_SupprPériode = "Suppression de la période" 'GIRATOIRE - Résultats

Public Const IDl_Imprimante = "Imprimante"                  ' PrintAPI et frmImprimer
Public Const IDm_ErrImprim = "Erreur " & IDl_Imprimante     ' PrintAPI et frmImprimer

Public Const IDl_ModeVLPL2R = "Mode VL-PL-2R" ' Données - Imprimer
Public Const IDl_ModeUVP = "Mode UVP"          ' Données - Imprimer - TRAFIC

  '--------------- Autres Constantes chaines
Public Const IDl_METRE = " m"   ' Utilisées par frmDonnées et frmImprimer
Public Const IDl_AbrévSaturBranche = "/SBr"

' Module GirabaseMain
'--------------------
Public Const IDm_MenuAngle = "&Transformer les angles en"
Public Const IDl_Degrés = "degrés"
Public Const IDl_Grades = "grades"

'Module Giratoire.cls
'--------------------
' Constantes de libellés pour les périodes de trafic
Public Const IDl_NouvellePériode = "Nouvelle période"
Public Const IDl_RenPériode = "Renommer période"
Public Const IDl_Inversion = "Inversion"
Public Const IDm_PériodeIncomplète = "Période de trafic incomplète"
Public Const IDm_PasDePériode = "Pas de Périodes de trafic dans le projet lu"

' Import de matrice
Public Const IDm_NbBranchesDifférent = "Nombre de branches du projet importé différent du giratoire courant"
Public Const IDm_IncompatibleBrancheUnidirection = "Incompatibilité entre les deux projets (branche unidirectionnelle)"

' Constante de lecture de fichier
Public Const IDm_ErrLectureFichier = "Erreur en lecture du fichier"
Public Const IDm_ligne = "ligne"

' Module Outils
'--------------------
Public Const IDm_Numeric = "Numérique obligatoirement"
Public Const IDm_Positif = "Valeur strictement positive"
Public Const IDm_ErreurFatale = "Erreur fatale"
Public Const IDm_LectureSeule = "Fichier en lecture seule"
Public Const IDm_FichierUtilisé = "Fichier en cours d'utilisation"
Public Const IDm_FichierDéjaOuvert = "Fichier déjà ouvert"
Public Const IDm_EnregistrerSousDabord = "Enregistrez le d'abord sous un autre nom"

'Module Branches.cls
'--------------------
Public Const IDm_DoublonBranche = "Nom de branche déjà utilisé"

'Module Trafics.cls
'--------------------
Public Const IDm_DoublonPériode = "Nom de période déjà utilisé"
Public Const IDm_IncompletPériode = "Période(s) de trafic incomplètement saisie(s)"

'Module DessinGiratoire
'----------------------
' Constantes permettant d'afficher des messages dans l'invite
Public Const IDl_RayonIntérieur = "Rayon intérieur"
Public Const IDl_RayonExtérieur = "Rayon extérieur"
Public Const IDl_LargeurAnneau = "Largeur de l'anneau"
Public Const IDl_BandeFranchissable = "Bande franchissable"
Public Const IDm_LargeurAnneauNonNulle = "La largeur de l'anneau ne doit pas être nulle"
Public Const IDm_LargeurBandePositive = "La largeur de la bande franchissable doit être positive"
Public Const IDm_BorneBranche = " doit rester entre les branches "

'Module Imprimer
'---------------
Public Const IDl_ImprimanteEnCours = "Imprimante en cours"
'Const IDl_Imprimante = "Imprimante"
'Const IDm_ErrImprim = "Erreur " & IDl_Imprimante           ' MDIGirabase et frmImprimer

Public Const IDl_Date = "Date"
Public Const IDl_Page = "Page"
Public Const IDl_Suite = " (suite)"

Public Const IDl_EnMetre = " (en m)"
Public Const IDl_OUI = "OUI"
Public Const IDl_Neant = "Néant"

Public Const IDl_Branches = "Branches"
Public Const IDl_Conseils = "Conseils"

Public Const IDl_PériodesTrafic = "Périodes de trafic"
Public Const IDl_ToutesPériodes = "Toutes les périodes"
Public Const IDl_Entrant = "Entrant"
Public Const IDl_Sortant = "Sortant"
Public Const IDl_Total = "Total"
Public Const IDl_EnUVP = " en UVP"

Public Const IDl_FichierTexte = "Fichier texte"

Public Const IDm_SaisirFichier = "Saisir un nom de fichier"
Public Const IDm_ExistFichier = "existe déjà" & vbCrLf & "Voulez-vous le remplacer?"

Public Const IDm_SaturBranche = "En acceptant une saturation sur la branche"
'Const IDm_BrancheSortie = "Branche de sortie uniquement"
Public Const IDm_GiratoireNonConforme = "Giratoire non conforme"
Public Const IDm_DessinImpossible = "Giratoire non dessinable"
