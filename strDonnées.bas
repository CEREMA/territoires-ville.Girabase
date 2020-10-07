Attribute VB_Name = "strDonnées"

'*************************************************************************************
'*
'*          Projet GIRABASE - CERTU - CETE de l'Ouest
'*
'*          Module de base : DONNEES.BAS - strDonnées
'*
'*          CONSTANTES DE CHAINE Susceptibles d'être traduites du module frmDonnées
'*          Certaines constantes sont utilisées également utilisées par frmRésultats
'*
'*************************************************************************************

Option Explicit


'***********************************************
 
Public Const IDm_Enregistrer = "Enregistrer le giratoire"
Public Const IDm_CréePériode = "Créer la Période"
  
' Constantes se rapportant à un changement de contexte
'Public Const IDv_Chevauchement = "Chevauchement des branches entre "
Public Const IDv_Chevauchement = "Impossible de prendre en compte votre demande. Cela conduirait à faire chevaucher les branches "
Public Const IDv_ModifMilieu = _
"Le passage en rase campagne nécessite de réduire auparavant la largeur de l'anneau."
Public Const IDm_ReinTrafic = _
"Réinitialisation des Périodes de trafic"

' Constantes d'information
Public Const IDi_Période = _
  "Choisissez une période existante ou tapez le nom d'une nouvelle période..."
Public Const IDi_BF = "Le rayon extérieur doit être inférieur à 15 m, la bande franchissable doit être de pente inférieure à 6 % " _
  & "et être délimitée par des bordures de moins de 3 cm sans ligne continue." & _
  " Sinon, la bande franchissable est à associer à l'ilôt central infranchissable."
Public Const IDi_LE4M = _
 "Largeur conseillée entre marquages ou à défaut entre bordures, 3,5 à 4 m pour une voie, 6 à 7 m pour 2 voies, " _
 & "9 à 10 m pour 3 voies."
Public Const IDi_LS = _
 "Largeur conseillée entre marquages ou à défaut entre bordures, 4 à 5 m pour une voie, " _
 & "6 à 7 m pour 2 voies."
Public Const IDi_QP = "Trafic bidirectionnel..."
Public Const Idi_Défaut = "   "

' A Tests sur valeurs individuelles
  ' A1 -Site - Dimensionnement

    ' A1.1 Données et résultats
Public Const IDm_TropDeBranchesEnRC = _
 "Un giratoire à plus de 6 branches n'est pas recommandé en rase campagne."
Public Const IDm_RTropGrand = _
  "Un rayon d'îlot infranchissable supérieur à 25 m est très rarement justifié."
Public Const IDm_LATropGrand = _
  "Un anneau aussi large est inutile."
Public Const IDm_LEPetit = _
  "Si possible, une largeur d'entrée d'au moins 3 m est préférable."
Public Const IDm_LETropLargeEnRC = _
  "Une entrée à 3 voies n'est pas recommandée en rase campagne."
Public Const IDm_LETropLargePourPiétons = _
  "Les piétons auront des difficultés à traverser l'entrée."
Public Const IDm_LSPetit = _
  "Si possible, une largeur de sortie d'au moins 3,5 m est préférable."
Public Const IDm_LSTropLarge = _
  "Une sortie aussi large est rarement utile."
    
    ' A1.2 Données seules
Public Const IDv_RayonInferieur100m = "La valeur du rayon est limitée à 100 mètres."
Public Const IDm_RTropGrandPourMiniG = _
  "Pour un mini-giratoire, le rayon est pris égal à 0." _
& " Dans les autres cas, le rayon est obligatoirement supérieur à 3,5 m."
Public Const IDm_RNulEnRC = _
  "Les mini-giratoires ne sont pas autorisés en rase campagne."
Public Const IDm_RNulEnPU = _
  "Un mini-giratoire ne peut pas être réalisé en entrée d'agglomération " _
& "ou sur un itinéraire de contournement."
Public Const IDm_LENul = "Aucune entrée possible - Branche de sortie uniquement."
Public Const IDm_LETropPetit = "L'entrée est trop étroite."
Public Const IDm_LE2Roues = "Entrée spéciale 2 roues sinon l'entrée est trop étroite."
Public Const IDm_LSNul = "Aucune sortie possible - Branche d'entrée uniquement."
Public Const IDm_LSTropPetit = "La sortie est trop étroite."
Public Const IDm_LS2Roues = "Sortie spéciale 2 roues sinon la sortie est trop étroite."
  
  ' A2 Trafics
Public Const IDm_QPTropGrand = "Le trafic piéton est très important. Vérifiez vos données."
Public Const IDm_QTropGrand = "Le trafic est très important. Vérifiez vos données."
  
  ' A3 Tests sur valeurs calculées
    ' A3.1 Données et Résultats
Public Const IDm_RgVoirGiration = _
  "Vérifiez la giration des bus et poids-lourds."
Public Const IDm_RgVoirGirationEnRC = _
  "Cette taille de giratoire n'est acceptable que sur le réseau secondaire en rase campagne." _
& " Vérifiez la giration des bus et poids-lourds."
    ' A3.2 données seules
Public Const IDm_RgTropPetitPourMiniG = _
  "Le rayon extérieur est trop faible pour un mini-giratoire."
Public Const IDm_RgTropGrandPourMiniG = _
  "Un mini-giratoire n'est pas approprié." _
& "L'emprise disponible permet l'aménagement d'un giratoire semi-franchissable."
Public Const IDm_QETropImportant = _
  "Le trafic est très important pour une entrée. Vérifiez vos données"

' B Tests impliquant plusieurs valeurs
Public Const IDm_RgTropPetit = _
  "Avec une entrée à 2 voies, un rayon extérieur d'au moins 20 m est souhaitable en rase campagne."
Public Const IDm_EvasementEnRC = _
  "En rase campagne, l'évasement devrait être complet 35 m avant l'entrée."
Public Const IDm_EvasementTropPetit = _
  "La longueur d'évasement est courte."
Public Const IDm_LATropEtroit = "L'anneau est trop étroit."
Public Const IDm_LATropEtroitPourEntrer = _
  "L'anneau est trop étroit pour une circulation optimale de la voie d'entrée "
Public Const IDm_LITropPetit = _
  "La largeur d'îlot séparateur est insuffisante pour les piétons."
Public Const IDm_LITropGrand = _
  "Le trafic sortant n'a pas d'influence sur la capacité." _
& "Vous pouvez éventuellement réduire la largeur de l'îlot séparateur."
Public Const IDm_Bf = _
  "Pour un giratoire semi-franchissable," _
& " la largeur de bande franchissable doit être comprise entre 1,5 m et 2 m."
Public Const IDm_AngleTropPetitPourMiniG = _
  "Configuration dangereuse : risque de contournement permanent par la gauche pour le tourne à gauche."
Public Const IDm_AnglePourMiniG = _
  "Configuration dangereuse : risque de contournement par la gauche pour le tourne à gauche."
Public Const IDm_QTropPetitPourTAD = _
  "Le trafic ne justifie pas la présence de cette voie directe de tourne-à-droite."
  
Public Const IDv_RapportLE = _
  "Le rapport LE à 4m / LE à 15m doit être compris entre 1 et 2,5."
Public Const IDm_BfTropPetitPourMiniG = _
  "Le dôme central franchissable d'un mini-giratoire doit avoir un rayon compris entre 1,5 m et 2,5 m."
Public Const IDv_LTropGrand = _
  "La somme des largeurs d'entrée à 4m, d'ilot et de sortie doit être inférieure au diamètre extérieur de l'anneau."
Public Const IDm_QENul = _
  "Attention ! Incohérence entre la largeur d'entrée et le trafic."
Public Const IDm_QSNul = _
  "Attention ! Incohérence entre la largeur de sortie et le trafic."
Public Const IDm_QEGrandPourMiniG = _
  "Attention ! Trafic important, il existe un risque de dysfonctionnement."
Public Const IDm_QETropGrandPourMiniG = _
  "Le trafic est trop important pour un mini-giratoire."
Public Const IDm_QETropGrand = _
  "Trafic très important pour un giratoire."


' C Messages d'avertissement à la saisie

Public Const IDv_LE0etTAD = _
 "Une largeur d'entrée nulle n'est pas compatible avec la présence d'une voie de Tourne à Droite."
Public Const IDv_LE0etLS0 = _
 "La largeur d'entrée et la largeur de sortie ne peuvent être toutes deux nulles."
Public Const IDv_RgOuBranchesIncorrect = _
 "Le rayon extérieur n'est pas compatible avec la dimension des branches ; " & vbCrLf & _
 "il vous faudra d'abord rectifier les branches concernées avant de réduire le rayon extérieur."
Public Const IDv_RgOuUneBrancheIncorrect = _
  "Cette valeur n'est pas compatible avec les autres données géométriques du giratoire."
Public Const IDv_ValeurNumérique = _
 "La valeur n'est pas numérique."
Public Const IDv_ValeurPositive = _
 "La valeur doit être positive ou nulle."
Public Const IDv_ControleBornes = _
 "La valeur doit être comprise entre "
Public Const IDv_ControleBornesLA = _
 "La largeur d'anneau doit être comprise entre "
Public Const IDv_ValidationRgMinimal = _
 "La valeur du rayon extérieur doit être supérieure à "
Public Const IDv_ControleBornesRg = _
  "La valeur du rayon extérieur doit être comprise entre "
Public Const IDv_TraficTotalNul = _
  "Le trafic total de la période active est nul."

Public Const IDv_RetTAD = _
 "Les voies Tourne à Droite sont interdites sur un mini-giratoire." _
  & vbCrLf & "Voulez-vous supprimer les voies de tourne à droite introduites sur le giratoire?"
Public Const IDv_PasTrafic = _
  "Aucun trafic n'a été saisi."
Public Const IDv_NonValide = _
  "Des données non valides ou incomplètes ont été rencontrées lors de la phase de vérification." _
  & vbCrLf & "Vous devez corriger les erreurs pour obtenir les résultats."
Public Const IDv_TraficNonNul = _
  "Des Trafics non nuls ont été détectés sur des largeurs d'entrée ou de sortie nulles." _
  & vbCrLf & "Vous pouvez effacer les trafics ou corriger le dimensionnement ;" _
  & vbCrLf & "Voulez-vous effacer les trafics rencontrés sur des largeurs nulles?"
