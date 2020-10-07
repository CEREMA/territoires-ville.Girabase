Attribute VB_Name = "Registry"
'**********************************************************************************************************
'
'         Projet Registry : Routines d'accès à la base de registres
'
'       Fonction principale : GetKeyValue
'
'         Septembre 1999
'
'         A. Vignaud  CETE de l'Ouest - DIOG/GTE-iTS
'**********************************************************************************************************
Option Explicit


' Options de sécurité des clés de registre...
' Constantes définies dans Winnt.h
Const KEY_ALL_ACCESS = &H2003F
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20

' Types de clés de registre ROOT...
' Constantes définies dans Winreg.h
Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_USERS = &H80000003
Const HKEY_PERFORMANCE_DATA = &H80000004    ' Win NT seulement
Const HKEY_CURRENT_CONFIG = &H80000005
Const HKEY_DYN_DATA = &H80000006            ' W95 et 98 seulement

Const ERROR_SUCCESS = 0
' Constantes définies dans Winnt.h
Const REG_SZ = 1                         ' Chaîne Unicode terminée par Null
Const REG_MULTI_SZ = 7                  ' Tableau de chaines Unicode
Const REG_DWORD = 4                      ' Nombre 32 bits


' Déclaration des Fonctions API
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" _
              (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, _
              ByVal samDesired As Long, ByRef phkResult As Long) _
        As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" _
              (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
              ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) _
        As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

'**********************************************************************************************************
' Exemple d'Appel de GetKeyValue : Recherche du séparateur décimal dans le panneau de config
' La valeur est stockée dans la chaine Info
'**********************************************************************************************************

Public Sub ReTrouvePtDecimal()
Dim Info As String
  
Const REGKEYINFO = "Control Panel\International"
Const REGVALINFO = "sDecimal"
  
  On Error GoTo InfoErr

    ' Lit dans le registre la donnée associée à la clé et à la valeur
    If GetKeyValue(HKEY_CURRENT_USER, REGKEYINFO, REGVALINFO, Info) Then
      gbPtDecimal = Asc(Info)
    ElseIf IsNumeric("1,1") Then
    ' Entrée de registre introuvable ou non renseignée...
      gbPtDecimal = Asc(",")
    ElseIf IsNumeric("1.1") Then
      gbPtDecimal = Asc(".")
    Else
      GoTo InfoErr
    End If
    
    Exit Sub
    
InfoErr:
    MsgBox "Problème sur le point décimal de Windows", vbOKOnly
End Sub

'**********************************************************************************************************
' GetKeyValue : Ouvre une clé de Registre et retourne la valeur associée à une sous-clé de la clé ouverte
' KeyRoot : Clé racine - ex HKEY_CLASSES_ROOT
' KeyName : Nom de la clé - ex "Excel.Sheet.5"
' SubKeyRef : Sous-Clé - ex "EditFlags" - Correspond à Nom (ou à Valeur dans Rechercher)
' KeyVal : Valeur associée à la sous-clé, toujours transformée en chaine, et retournée à l'appelant - Correspond à Données

' GetKeyValue retourne Faux en cas d'échec, Vrai autrement
'**********************************************************************************************************
Public Function GetKeyValue(ByVal KeyRoot As Long, ByVal KeyName As String, ByVal SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                  ' Compteur de boucle
    Dim rc As Long                 ' Code de retour
    Dim hKey As Long               ' Pointeur vers une clé du registre ouverte
    Dim hDepth As Long             '
    Dim KeyValType As Long         ' Type de données d'une valeur de la clé de registre
    Dim tmpVal As String           ' Stockage temporaire de la donnée associée à la valeur de clé de registre
    Dim KeyValSize As Long         ' Taille de la variable de clé de registre
    '------------------------------------------------------------
    ' Ouvre la clé de registre sous la clé racine {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
'    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Ouvre la clé de registre
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_QUERY_VALUE, hKey) ' Ouvre la clé de registre autorisant RegQueryValueEx
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Gère l'erreur...
    
    KeyValSize = 1024             ' Marque la taille de la variable
    tmpVal = String(KeyValSize, Chr(0))     ' Alloue l'espace pour la variable
    
    '------------------------------------------------------------
    ' Extrait la valeur de la clé de registre...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, KeyValType, tmpVal, KeyValSize)    ' Lit/Crée la valeur de clé
'    rc = RegQueryValueEx(hKey, "", 0, KeyValType, tmpVal, KeyValSize)    ' Lit/Crée la valeur  {Défaut} de clé
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError    ' Gérer les erreurs
   
    
    '------------------------------------------------------------
    ' Détermine le type de la valeur de la clé pour conversion...
    '------------------------------------------------------------
    Select Case KeyValType          ' Recherche les types de données...
    Case REG_SZ                     ' Type de données de clé de registre chaîne
      ' Tronquage de la chaine dès le 1er caractère '\0' fourni par le C
      KeyVal = Left(tmpVal, InStr(tmpVal, Chr(0)) - 1)
      
    Case REG_MULTI_SZ
      ' Multi chaines : les chaines sont séparées par 1 caractère '\0' - la fin de la liste en comporte au moins 2
      Dim pos As Integer
      ' Tronquage de la chaine dès qu'on trouve au moins 2 caractères '\0'
      KeyVal = Left(tmpVal, InStr(tmpVal, Chr(0) & Chr(0)))
      pos = InStr(KeyVal, Chr(0))
      While pos <> 0
      'Substitution de chaque séparateur de chaine par un retour-chariot (13)
        Mid(KeyVal, pos, 1) = vbCr
        pos = InStr(KeyVal, Chr(0))
      Wend

    Case REG_DWORD                  ' Type de données de clé de registre double mot
      tmpVal = Left(tmpVal, InStr(tmpVal, Chr(0)) - 1)
      For i = Len(tmpVal) To 1 Step -1        ' Convertit chaque bit
        KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Construit la valeur octet par octet.
      Next
      KeyVal = Format("&H" + KeyVal)   ' Convertit le double mot en chaîne
    End Select
  
    '------------------------------------------------------------
    ' Fin du traitement
    '------------------------------------------------------------
    GetKeyValue = True                    ' Retourne Réussi
    rc = RegCloseKey(hKey)                ' Ferme la clé de registre
    Exit Function                         ' Quitte
    
    '------------------------------------------------------------
    ' Traitement d'erreu
    '------------------------------------------------------------
GetKeyError:    ' Nettoyer suite à erreur...
    KeyVal = ""                           ' Affecte une chaîne vide à la valeur de retour
    GetKeyValue = False                   ' Retourne Échec
    rc = RegCloseKey(hKey)                ' Ferme la clé de registre
End Function


