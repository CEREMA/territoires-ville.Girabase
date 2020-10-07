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
Public Const IDl_Courbes = "&Courbes"
Public Const IDl_Remarques = "&Remarques"
Public Const IDl_SaturerBranche = "Saturer la branche"
Public Const IDl_SupprimerPériode = "Supprimer la période"

' Constantes pour les courbes de capacité
Public Const IDl_Trafic = "Trafic"
Public Const IDl_Genant = "gênant"

' Constantes pour l'affichage du tableau
Public Const IDl_VehiculeHeure = "vh"

'Constantes pour les conseils
Public Const IDm_BrancheSortie = "Branche de sortie uniquement"
Public Const IDm_BrancheEntrée = "Branche d'entrée uniquement"

Public Const IDc_MatriceSaturation = "Branche avec un trafic en entrée limité à sa capacité"
Public Const IDc_TraficsIncomplets = "Les trafics de la période en cours sont incomplets ; " & "les conseils relatifs à cette période ne peuvent être édités."
Public Const IDc_QEnul = "Comme il n'y a jamais de trafic, la largeur d'entrée de la branche devrait être nulle."
Public Const IDc_QSnul = "Comme il n'y a jamais de trafic, la largeur de sortie la branche devrait être nulle."
Public Const IDc_RTropGrand2 = "Il peut être réduit au bénéfice de la sécurité."
Public Const IDc_IlotEtroit = "Un îlot plus large serait préférable pour les piétons."
Public Const IDc_IlotASeparer = "Penser à séparer l'entrée de la sortie par une bande en relief, une zone pavée ou autre."
Public Const IDc_LS2voiesN = "Une sortie à deux voies est nécessaire. "
Public Const IDc_LS2voiesP = "Une sortie à deux voies peut être envisagée. "
Public Const IDc_TraverséePiétons = "Attention aux traversées piétonnes."

'Constantes pour les conseils de fonctionnement (Branche active)
Public Const IDc_RCnégative = "ENTRÉE SATURÉE ; vous pouvez : "
Public Const IDc_RCfaible = "Attention, la réserve de capacité est faible ; vous pouvez : "
Public Const IDc_RC1 = " - envisager une voie directe de tourne-à-droite"
Public Const IDc_RC2 = " - élargir l'entrée à 2 voies"
Public Const IDc_RC2p = ", mais attention au traitement des traversées piétonnes"
Public Const IDc_RC3 = " - élargir l'entrée à 3 voies"
Public Const IDc_RC3p = " si le trafic piéton est très faible"
Public Const IDc_RC4 = " - élargir l'anneau et, si nécessaire, l'entrée"
Public Const IDc_RC5 = " - élargir l'îlot séparateur"
Public Const IDc_RC6 = " - agrandir le giratoire"
Public Const IDc_RC11 = "Un des mouvements est assez important pour envisager de déniveler le carrefour."
Public Const IDc_RC12 = "Une entrée à une voie suffit probablement."
Public Const IDc_RC13 = "Une entrée à une voie suffit probablement et serait plus favorable aux piétons"
Public Const IDc_RC14 = "Une entrée à 2 voies suffit probablement."
Public Const IDc_TMA1 = "Le temps d'attente sur la branche est important."
Public Const IDc_TMA2 = "Le temps moyen d'attente sur la branche est très important."
Public Const IDc_LK1 = "La file d'attente sur la branche est importante. " & "Attention aux pertes de visibilité en approche dues au profil en long ou au tracé."
Public Const IDc_LK2 = "La file d'attente sur la branche est très importante. " & "Attention aux pertes de visibilité en approche dues au profil en long ou au tracé."
Public Const IDc_LK3 = "La file d'attente sur la branche est importante, penser au carrefour en amont."
Public Const IDc_LK4 = "La file d'attente sur la branche est très importante, penser au carrefour en amont."
Public Const IDc_LK5 = "La file d'attente sur la branche peut être importante. " & "Attention aux pertes de visibilité en approche dues au profil en long ou au tracé."
Public Const IDc_LK6 = "La file d'attente sur la branche peut être importante, penser au carrefour en amont."

