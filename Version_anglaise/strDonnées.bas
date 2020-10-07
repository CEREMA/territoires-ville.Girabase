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
 
Public Const IDm_Enregistrer = "Save the roundabout"
Public Const IDm_CréePériode = "Create period"
  
' Constantes se rapportant à un changement de contexte
'Public Const IDv_Chevauchement = "Chevauchement des branches entre "
Public Const IDv_Chevauchement = "Impossible to consider your request. It would lead to overlapping of arms "
Public Const IDv_ModifMilieu = _
"Passage through a rural area necessitates the reduction beforehand of the circulatory carriageway width."
Public Const IDm_ReinTrafic = _
"Reinitialisation of traffic periods"

' Constantes d'information
Public Const IDi_Période = _
  "Choose a new period or enter the name if a new period…"
Public Const IDi_BF = "The external radius must be less than 15m,the crossing slip must be of a gradient less than 6% " _
& "and be demarcated by kerb less than 3 cm in height without a continuous line marking." _
& "If not, the crossing slip is considered as part of the non-traversable central island."
Public Const IDi_LE4M = _
 "Recommended width between road markings or, if absent, between kerbs, 3.5 to 4m for one lane, 6 to 7m for 2 lanes, " _
 & "9 to 10 m for 3 lanes."
Public Const IDi_LS = _
 "Recommended width between road markings or, if absent, between kerbs, 4 to 5m for one lane, " _
 & "6 to 7 m for 2 lanes."
Public Const IDi_QP = "Two-way traffic..."
Public Const Idi_Défaut = "   "

' A Tests sur valeurs individuelles
  ' A1 -Site - Dimensionnement

    ' A1.1 Données et résultats
Public Const IDm_TropDeBranchesEnRC = _
 "A roundabout of more than 6 arms is not recommended in a rural area."
Public Const IDm_RTropGrand = _
  "A central kerbed island radius larger than 25m is very rarely justified."
Public Const IDm_LATropGrand = _
  "Such a wide circulatory carriageway is unnecessary."
Public Const IDm_LEPetit = _
  "If possible, an entry width of at least 3m is preferable."
Public Const IDm_LETropLargeEnRC = _
  "A 3-lane entry is not recommended in a rural area."
Public Const IDm_LETropLargePourPiétons = _
  "Pedestrians will have difficulty crossing the entry."
Public Const IDm_LSPetit = _
  "If possible, an entry width of at least 3.5m is preferable."
Public Const IDm_LSTropLarge = _
  "Such a wide exit is rarely necessary."
    
    ' A1.2 Données seules
Public Const IDv_RayonInferieur100m = "The radius value is limited to 100m."
Public Const IDm_RTropGrandPourMiniG = _
  "For a mini-roundabout, the radius is taken as 0. " _
& "In other cases, the radius is mandatorily larger than 3.5m."
Public Const IDm_RNulEnRC = _
  "Mini-roundabouts are not authorised in rural areas."
Public Const IDm_RNulEnPU = _
  "A mini-roundabout cannot be placed at an entry to a built-up area or to a bypass route."
  
Public Const IDm_LENul = "No entry possible - exit arm only."
Public Const IDm_LETropPetit = "The entry is too narrow."
Public Const IDm_LE2Roues = "2 wheel special entry, if not the entry is too narrow."
Public Const IDm_LSNul = "No exit possible - entry arm only."
Public Const IDm_LSTropPetit = "Exit is too narrow."
Public Const IDm_LS2Roues = "2 wheel special entry, if not the entry is too narrow."
  
  ' A2 Trafics
Public Const IDm_QPTropGrand = "Pedestrian traffic is very high. Check your data."
Public Const IDm_QTropGrand = "Traffic is very high. Check your data."
  
  ' A3 Tests sur valeurs calculées
    ' A3.1 Données et Résultats
Public Const IDm_RgVoirGiration = _
  "Check the turning of buses and heavy vehicles."
Public Const IDm_RgVoirGirationEnRC = _
  "This size of roundabout is only acceptable on the secondary network in rural areas. " _
& "Check the turning of buses and heavy  vehicles."
    ' A3.2 données seules
Public Const IDm_RgTropPetitPourMiniG = _
  "The external radius is too small for a mini-roundabout."
Public Const IDm_RgTropGrandPourMiniG = _
  "A mini-roundabout is not appropriate. " _
& "The available development area is adequate for the development of a semi-traversable roundabout."
Public Const IDm_QETropImportant = _
  "Traffic is very high for an entry. Check your data."

' B Tests impliquant plusieurs valeurs
Public Const IDm_RgTropPetit = _
  "With a 2-lane entry, an external radius of at least 20 m is desirable in rural areas."
Public Const IDm_EvasementEnRC = _
  "In rural areas, flaring should be complete 35m before the entry."
Public Const IDm_EvasementTropPetit = _
  "Flare length is short."
Public Const IDm_LATropEtroit = "Circulatory carriageway is too narrow."
Public Const IDm_LATropEtroitPourEntrer = _
  "Circulatory carriageway is too narrow for an optimal circulation in the entry lane "
Public Const IDm_LITropPetit = _
  "The width of the traffic deflection island is insufficient for pedestrians."
Public Const IDm_LITropGrand = _
  "Exiting traffic does not have any influence on capacity. " _
& "You can possibly reduce the width of the traffic deflection island."
Public Const IDm_Bf = _
  "For a semi-traversable roundabout the width of the crossing slip must be between 1.5m and 2m."
  
Public Const IDm_AngleTropPetitPourMiniG = _
  "Dangerous configuration : risk of permanent bypass on the left for the 'turn-left'."
Public Const IDm_AnglePourMiniG = _
  "Dangerous configuration : risk of bypass on the left for the 'turn-left'."
Public Const IDm_QTropPetitPourTAD = _
  "Traffic does not justify the presence of this segregated right-turning lane."
  
Public Const IDv_RapportLE = _
  "The ratio EnW at 4m/EnW at 15m must range between 1 and 2.5."
Public Const IDm_BfTropPetitPourMiniG = _
  "The traversable central dome of a mini-roundabout must have a radius of between 1.5m and 2.5m."
Public Const IDv_LTropGrand = _
  "The sum of the entry width at 4m, the island and the exit widths must be less than the external diameter of the circulatory carriageway."
Public Const IDm_QENul = _
  "Warning! Inconsistency between the entry width and traffic."
Public Const IDm_QSNul = _
  "Warning! Inconsistency between the exit width and traffic."
Public Const IDm_QEGrandPourMiniG = _
  "Warning! Heavy traffic, there is a risk of malfunctioning."
Public Const IDm_QETropGrandPourMiniG = _
  "Traffic volume is too high for a mini-roundabout."
Public Const IDm_QETropGrand = _
  "Traffic volume very high for a roundabout."


' C Messages d'avertissement à la saisie

Public Const IDv_LE0etTAD = _
 "A zero entry width is not compatible with the presence of a turn-right lane."
Public Const IDv_LE0etLS0 = _
 "Both the entry width and the exit width cannot be zero."
Public Const IDv_RgOuBranchesIncorrect = _
 "The external radius is not compatible with the dimension of the arms ; " & vbCrLf & _
 "you must first rectify the arms before reducing the external radius."
Public Const IDv_RgOuUneBrancheIncorrect = _
  "This value is not compatible with the other geometric data of the roundabout."
Public Const IDv_ValeurNumérique = _
 "This is not a numerical value."
Public Const IDv_ValeurPositive = _
 "This value must be positive or zero."
Public Const IDv_ControleBornes = _
 "This value must range between "
Public Const IDv_ControleBornesLA = _
 "Circulatory carriageway width must range between "
Public Const IDv_ValidationRgMinimal = _
 "The value of the external radius must be larger than "
Public Const IDv_ControleBornesRg = _
  "The value of the external radius must range between "
Public Const IDv_TraficTotalNul = _
  "The total traffic of the active period is zero."

Public Const IDv_RetTAD = _
 "Turn-right lanes are prohibited on a mini-roundabout. " _
  & vbCrLf & "Do you want to delete the turn-right lanes appearing on the roundabout?"
Public Const IDv_PasTrafic = _
  "No traffic has been entered."
Public Const IDv_NonValide = _
  "Non-valid or incomplete data has been encountered during the checking phase." _
  & vbCrLf & "You have to correct the errors to obtain the results."
Public Const IDv_TraficNonNul = _
  "Non-zero traffic has been detected on zero entry or exit widths." _
  & vbCrLf & "You can delete the traffic or correct the sizing ;" _
  & vbCrLf & "Do you want to delete the traffic encountered on the zero widths?"
