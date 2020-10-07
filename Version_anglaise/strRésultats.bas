Attribute VB_Name = "strRésultats"
'*************************************************************************************
'*
'*          Projet GIRABASE - CERTU - CETE de l'Ouest
'*
'*          Module de base : RESULTATS.BAS - strRésultats
'*
'*          CONSTANTES DE CHAINE Susceptibles d'être traduites du module frmRésultats
'*
'*************************************************************************************

Option Explicit

'Constantes pour les libellés
Public Const IDl_Courbes = "&Curves"
Public Const IDl_Remarques = "&Comments"
Public Const IDl_SaturerBranche = "Saturate &arm"
Public Const IDl_SupprimerPériode = "Delete period"

' Constantes pour les courbes de capacité
Public Const IDl_Trafic = "Traffic"
Public Const IDl_Genant = "Opposing"

' Constantes pour l'affichage du tableau
Public Const IDl_VehiculeHeure = "vh"

'Constantes pour les conseils
Public Const IDm_BrancheSortie = "Exit arm only"
Public Const IDm_BrancheEntrée = "Entry arm only"

Public Const IDc_MatriceSaturation = "Arm with entering traffic limited to its capacity"
Public Const IDc_TraficsIncomplets = "Traffic for the period under study incomplete ; " & "recommendations relative to this period can not been displayed."
Public Const IDc_QEnul = "Since there is ever any traffic, the arm entry width should be zero."
Public Const IDc_QSnul = "Since there is ever any traffic, the arm exit width should be zero."
Public Const IDc_RTropGrand2 = "It could be reduced to the advantage of safety."
Public Const IDc_IlotEtroit = "A wider island would be preferable for pedestrians."
Public Const IDc_IlotASeparer = "Remember to separate the entry from the exit with a raised band, a paved zone or similar."
Public Const IDc_LS2voiesN = "A two lane exit is required. "
Public Const IDc_LS2voiesP = "A two lane exit can be considered. "
Public Const IDc_TraverséePiétons = "Beware of pedestrian crossings."

'Constantes pour les conseils de fonctionnement (Branche active)
Public Const IDc_RCnégative = "ENTRY SATURATED ; you can : "
Public Const IDc_RCfaible = "Caution, the reserve capacity is low ; you can : "
Public Const IDc_RC1 = " - consider a segregated right turning lane"
Public Const IDc_RC2 = " - widen the entry to 2 lanes"
Public Const IDc_RC2p = ", but pay attention to the processing of pedestrian crossings"
Public Const IDc_RC3 = " - widen the entry to 3 lanes"
Public Const IDc_RC3p = " if pedestrian traffic is very low"
Public Const IDc_RC4 = " - widen the circulatory carriageway and, if necessary, the entry"
Public Const IDc_RC5 = " - widen traffic deflection island"
Public Const IDc_RC6 = " - enlarge the roundabout"
Public Const IDc_RC11 = "Level of one of the movements is high enough to consider grade-separating the roundabout."
Public Const IDc_RC12 = "A one lane entry is probably sufficient."
Public Const IDc_RC13 = "A one lane entry is probably sufficient and would be more favorable for pedestrians"
Public Const IDc_RC14 = "A two lane entry is probably sufficient."
Public Const IDc_TMA1 = "The waiting time on the arm is long."
Public Const IDc_TMA2 = "The average waiting time on the arm is too long."
Public Const IDc_LK1 = "Queue on the arm is long. " & "Watch out for visibility loss due to vertical or horizontal transition curves."
Public Const IDc_LK2 = "Queue on the arm is very long. " & "Watch out for visibility loss due to vertical or horizontal transition curves."
Public Const IDc_LK3 = "Queue on the arm is long, bear in mind roundabout upstream."
Public Const IDc_LK4 = "Queue on the arm is very long, bear in mind roundabout upstream."
Public Const IDc_LK5 = "Queue on the arm could be long. " & "Watch out for visibility loss due to vertical or horizontal transition curves."
Public Const IDc_LK6 = "Queue on the arm could be long, bear in mind roundabout upstream."

