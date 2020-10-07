Attribute VB_Name = "CopyControl"
'-----------------------------------------------------------------
'
'  CCTest.BAS
' Programme test de CopyControl pour GIRATION, pouvant servir de modèle pour d'autres produits
'
' Origine : CERTU - Mars 98
' Mise à jour : CETE de l'Ouest (A.VIGNAUD) - Avril 98
'
'-----------------------------------------------------------------

Option Explicit

' Variable globale pour A_Propos
Public SerialNumber As String

' Variable globale à placer dans le module principal
'Public VersionDemo As Boolean


' A paramétrer selon l'application
Private Const BonCode = "its99GB4.0"
Private Const BonneDLL = "GIRABASE.DLL"
Private LongueurCode As Integer
Private LongueurNomDLL As Integer

Private Const MinTime = 15 ' une vérif toutes les 15mn

' Conserve l'heure et le jour de la dernière modification pour MyCCToujours()
Private MyDerModif As Integer
Private MyNewDay As Integer

' Disque porteur de la protection
Private disquePorteur As Integer

' Mémorisation du disque et du répertoire courants
Private saveDrive As String
Private saveDir As String  '(utile uniqt ds l'envt de développement)

Const Long256 As Long = 256

'Déclaration de la structure
Type CCMB
  B1 As String * 1
  B2 As String * 1
  B3 As String * 1
  B4 As String * 1
  Func As String * 1
  Rcodelo As String * 1
  Rcodehi As String * 1
'  Rcode As Integer ' ADAPTATION 16bits pour faciliter Rcode = -1 (Integer = 2 octets)
  Drive As String * 1
  Dir As String * 4
  Vers1 As String * 1
  Vers2 As String * 1
  SN As String * 2
'  SN As Integer ' à la place de String * 2; pour faciliter la lecture (mais attention au +/-)
  Pcode As String * 9
  Pname As String * 13
  CCSN As String * 2
  Master As String * 1
  DrType As String * 1
  Copies1 As String * 1
  Copies2 As String * 1
  InitCopies As String * 2
  Useslo As String * 1
  Useshi As String * 1
  IUseslo As String * 1
  IUseshi As String * 1
  ExpD As String * 1
  ExpM As String * 1
  ExpYlo As String * 1
  ExpYhi As String * 1
  NotreDecalage As String * 4        '4 octets pris sur Remainder
  MsgSecurit As String * 256         '256 octets pris sur Remainder

  Remainder As String * 198   '458 - 260 // taille de la structure = 512o'
End Type

'Déclaration des variables'
Private myCC As CCMB

#If Win32 Then
'Déclaration de la DLL en 32 bits d'après <exemple>
  Declare Function ccdll Lib "Girabase.dll" Alias "CC32" (CC As CCMB) As Integer
#Else
'Déclaration de la DDL16
  Declare Function ccdll Lib "Girabase.dll" Alias "CCDLL" (CC As CCMB) As Integer
#End If

Private Sub ClearStruct(lpCC As CCMB)
'Initialisation de la structure d'après doc
' ... et obligatoirement remettre Dir à NULL pour A: en cas de modification sur C:)
  lpCC.B1 = "C"
  lpCC.B2 = "C"
  lpCC.B3 = "M"
  lpCC.B4 = "B"
  lpCC.Func = Chr$(0)
'  lpCC.Rcode = 0
  lpCC.Rcodelo = Chr$(255) ' l'ensemble des 2 octets (FF)
  lpCC.Rcodehi = Chr$(255) ' donne -1 (utile pour la version 16 bits)
  lpCC.Drive = Chr$(0)
  lpCC.Dir = String$(4, 0)
  lpCC.Vers1 = Chr$(0)
  lpCC.Vers2 = Chr$(0)
'  lpCC.SN = 0
  lpCC.SN = String$(2, 0)
  lpCC.Pcode = String$(9, 0)
  lpCC.Pname = String$(13, 0)
  lpCC.CCSN = String$(2, 0)
  lpCC.Master = Chr$(0)
  lpCC.DrType = Chr$(0)
  lpCC.Copies1 = Chr$(0)
  lpCC.Copies2 = Chr$(0)
  lpCC.InitCopies = String$(2, 0)
  lpCC.Useslo = Chr$(0)
  lpCC.Useshi = Chr$(0)
  lpCC.IUseslo = Chr$(0)
  lpCC.IUseshi = Chr$(0)
  lpCC.ExpD = Chr$(0)
  lpCC.ExpM = Chr$(0)
  lpCC.ExpYlo = Chr$(0)
  lpCC.ExpYhi = Chr$(0)
  lpCC.NotreDecalage = String$(4, 0)
  lpCC.MsgSecurit = String$(256, 0)
  lpCC.Remainder = String$(198, 0)
End Sub


'utilisation de la protection'
Public Sub ProtectCheck()
Dim comp_status As Integer
Dim comp_statusA As Integer
Dim CChaine As String
Dim NotreCode As String
Dim flag As Boolean

If gbVersionDemo Or gbVersionDéveloppeur Then Exit Sub

LongueurCode = Len(BonCode)
LongueurNomDLL = Len(BonneDLL)

'on commence par regarder sur le disque courant
ChangeRep   ' on rend courant le disque de l'application
disquePorteur = 0
comp_status = appelDLL(0, disquePorteur, flag)
RetrouveRep
If flag Then Unload MDIGirabase     ' DLL non trouvée

'MsgBox "1er appel " & CStr(comp_status)
disquePorteur = Asc(Left(App.Path, 1)) - 64   ' utilisé par les futurs appels de CCToujours

If (comp_status = -28) Then   'correspond au Msg : Transférez le jeton dans le répertoire courant !
'on regarde la disquette A
    disquePorteur = 1
    comp_statusA = appelDLL(0, disquePorteur, flag)
'    MsgBox "2è appel " & CStr(comp_statusA)
    Select Case comp_statusA
            Case 0
            comp_status = 0              'on valide
            Case -57                     'erreur apparaissant si protétégé en écriture
            comp_status = -5700
            Case -67, -26
            comp_status = comp_statusA
    End Select
End If



'-------- Récupération du Numéro de Série
'         ... et Test du nom de la DLL original, et du Code
If (comp_status = 0) Then   'myCC est soit celle du disque courant, soit celle de la disquette
'    If (myCC.SN >= 0) Then SerialNumber = Str$(myCC.SN) Else SerialNumber = Str$(65536 + myCC.SN)
    SerialNumber = Str$(Asc(myCC.SN) + Long256 * Asc(Mid(myCC.SN, 2)))
'    MsgBox "Num Série : " & SerialNumber
'    NotreCode = Mid(myCC.MsgSecurit, 2, 10)
    NotreCode = Mid(myCC.MsgSecurit, 2, LongueurCode)
'    BonCode = Chr(105) + Chr(116) + Chr(115) + Chr(57) + Chr(56) + "GIR3.0" '??? pourquoi pas simplement "its98GIR3.0" (AV - 20/04/98)
    'ATTENTION : à adapter au message de sécurité
    '       BonCode : its98GIR3.0
    '       Left  : nombre de caractères de la partie visible et caché + 1 (pour le caractère null ~Z)
    '       Right : nombre de caractères de la partie caché
'    If (Left$(myCC.Pname, 12) <> BonneDLL Or NotreCode <> BonCode) Then comp_status = -19000
    If (Left$(myCC.Pname, LongueurNomDLL) <> BonneDLL Or NotreCode <> BonCode) Then comp_status = -19000
End If


'------- Messages d'erreur et sortie
If (comp_status <> 0) Then
    Select Case comp_status
        Case -19
            CChaine = "Produit non installé !"
        Case -26
            CChaine = "Le numéro de licence ne correspond pas"
        Case -28
            CChaine = "Jeton introuvable"
        Case -35
            CChaine = "Vérification impossible : le disque est protégé en écriture !"
        Case -5700
            CChaine = "Vérification impossible : la disquette est protégée en écriture !"
        Case -67
            CChaine = "Veuillez recommencer plus tard" & Chr(13) & "Trop d'utilisateurs sont présents !"
        Case Else
            CChaine = "Erreur n° " & Str$(comp_status) & Chr(13) & App.Title & " n'a pas trouvé la protection"
    End Select
    
    
    MsgBox CChaine, vbCritical, "Gestion de la Protection"
    Unload MDIGirabase
'    Exit Sub
End If

MyDerModif = 60 * Hour(Now) + Minute(Now)
MyNewDay = Day(Now)

Exit Sub

End Sub


Public Function MyCCToujours() As Integer '------- Vérifie si la protection est toujours là ....
Dim comp_status As Integer         '------- Valeur de retour non nulle signifie erreur
Dim NotreCode As String
Dim Maintenant As Integer
Dim flag As Boolean

'Version de demo
'ou vérification effectuée avec succès depuis moins de 15 minutes (l'accès disquette est long !!!)
MyCCToujours = 0
If gbVersionDemo Or gbVersionDéveloppeur Then Exit Function

Maintenant = 60 * Hour(Now) + Minute(Now)
If Maintenant - MyDerModif < MinTime And MyNewDay = Day(Now) Then Exit Function
    'On pourrait passer à coté du controle si la fonction n'était appelée qu'une fois par mois....
                                        
'----- Vérification

'on regarde le disque qui a été vérifié au lancement
comp_status = appelDLL(2, disquePorteur, flag)
If flag Then MyCCToujours = -19500: Exit Function

'------- Test à nouveau du nom de la DLL original, et du Code
If (comp_status = 0) Then   'myCC est soit celle du disque courant, soit celle de la disquette
'    NotreCode = Right$(Left$(myCC.MsgSecurit, 11), 10)
    NotreCode = Mid(myCC.MsgSecurit, 2, LongueurCode)
'    If (Left$(myCC.Pname, 12) <> BonneDLL Or NotreCode <> BonCode) Then comp_status = -19000
    If (Left$(myCC.Pname, LongueurNomDLL) <> BonneDLL Or NotreCode <> BonCode) Then comp_status = -19000
End If

'------- en cas de succès : on note l'heure de la dernière vérification
If (comp_status = 0) Then
    MyDerModif = 60 * Hour(Now) + Minute(Now)
    MyNewDay = Day(Now)
End If

'------- Valeur de retour : non nulle signifie erreur
MyCCToujours = comp_status

End Function


Private Sub ChangeRep()
' Repositionnement éventuel du disque courant sur celui de l'application
' pour être sûr de trouver la protection

  If Mid(CurDir, 2, 1) = ":" Then
    If Left(CurDir, 1) <> Left(App.Path, 1) Then
      saveDrive = Left(CurDir, 1)
      ChDrive Left(App.Path, 1)
    End If
  End If
 
' Repositionnement éventuel du répertoire courant sur celui de l'application
' pour être sûr de trouver la DLL
' utile uniquement en environnement de développement
  If StrComp(CurDir, App.Path, 1) <> 0 Then
'    saveDir = CurDir
    ChDir App.Path
  End If

End Sub

Private Sub RetrouveRep()

  If saveDrive <> "" Then ChDrive saveDrive
  If saveDir <> "" Then ChDir saveDir

End Sub

Private Function appelDLL(ByVal fonction As Integer, ByVal disque As Integer, Absent As Boolean) As Integer
' Appel de la DLL de CopyControl pour vérification de la protection

' fonction = 0 - vérification + inscription en tant qu'utilisateur
'            2 - vérification seule

' disque   = 0 - disque courant : recherche sur CC_PATH, chemin programme (où se trouve la DLL),
'                                 Rép. de travail, Rép racine
'                                 sinon cf myCC.Dir
'            1 - disquette A

Dim souris As Integer   'sauvegarde de la forme souris
Dim Msg As String

  souris = Screen.MousePointer
  Screen.MousePointer = 11    ' sablier

  ClearStruct myCC       ' réinitialisation de la structure
'#If Win16 Then
  '  myCC.Rcode = -1     '16 bits : non demandé en 32bits
  '#End If
  myCC.Func = Chr$(fonction)
  myCC.Drive = Chr$(disque)

  On Error GoTo GestErr
  appelDLL = ccdll(myCC)   'appel de la DLL
  Screen.MousePointer = souris

Exit Function

'----gestion de l'absence de la DLL
GestErr:
  #If Win16 Then
'    MsgBox "Girabase.DLL non trouvée", vbCritical, "Gestion de la protection"
    If Err = 53 Then
      Msg = BonneDLL & " non trouvée"
    ElseIf Err = 48 Then
      Msg = "Anomalie dans l'appel de " & BonneDLL
    Else
      Msg = "Erreur en vérification de la protection " & CStr(Err)
    End If
  
  #Else
    Msg = Err.Description
  #End If
  
  MsgBox Msg, vbCritical, "Gestion de la protection"
  Absent = True
  Resume Next

End Function

