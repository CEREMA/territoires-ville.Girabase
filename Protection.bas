Attribute VB_Name = "Protection"
Option Explicit

'Définition du type de protection
'QLM=Quick License Manager
'CPM=CopyMinder
Public TYPPROTECTION As Byte '= "CPM"
Public Const CPM = 0
Public Const QLM = 1

Public Const VER_ANGLAISE = 0
Public Const VER_FRANCAISE = 1

'Définition des paramètres de licence provenant de QLM
'Utile sssi TYPPROTECTION=QLM
Private Const IdProduit As Long = 1
Private Const Produit As String = "Girabase"
Private Const Guid As String = "{57BFA51B-E7B3-462D-AF90-8ACC822716A7}"
Private Const Vermaj As Long = 4
Private Const Vermin As Long = 0
Private Const message As String = "its-Girabase"

'Définition des paramètres des noms de fichiers utilisés avec Qlm
Private Const NomFichierLicence As String = "relic.ctu"
Public Const NomFichierSerial As String = "reser.ctu"

'Definition des messages
Public Const F_TITRE = "Enregistrement de la licence"
Public Const E_TITRE = "License register"
Public Const F_MSG = "Enregistrement de la licence SVP"
Public Const E_MSG = "Please register your license"
Public Const F_LBLICENCE = "Licence : "
Public Const E_LBLICENCE = "License : "
Public Const E_LBSERIAL = "Product-Key : "
Public Const F_LBSERIAL = "Clé-Produit : "
Public Const F_BTNOK = "Enregistrer"
Public Const E_BTNOK = "Register"
Public Const F_BTNCANCEL = "Annuler"
Public Const E_BTNCANCEL = "Cancel"
Public Const E_MSGPWDINVALID = "License Password invalid!"
Public Const F_MSGPWDINVALID = "licence invalide!"
Public Const E_MSGPWDEXPIRED = "Licence Password expired!"
Public Const F_MSGPWDEXPIRED = "licence expirée!"
Public Const E_MSGPWDVALID = "The value of the RegOptions passed via the serial number."
Public Const F_MSGPWDVALID = "Votre licence a été enregistrée avec succès."

Public Const REGKEYINFO = "SOFTWARE\CERTU\Girabase\4.0" 'ne sert pas
Public Const REGVALINFO = "US" 'ne sert pas

Public Const F_MSGREGERROR1 = "Erreur fatale. Version non compatible."
Public Const E_MSGREGERROR1 = "Fatal Error. Wrong version."
Public Const F_MSGREGERROR2 = "Fin d'exécution."
Public Const E_MSGREGERROR2 = "Execution failed."

Public Titre As String
Public Msg As String
Public LBLICENCE As String
Public LBSERIAL As String
Public BtnOK As String
Public btnCancel As String
Public MSGPWDINVALID As String
Public MSGPWDEXPIRED As String
Public MSGPWDVALID As String

Public MSGREGERROR1 As String
Public MSGREGERROR2 As String

'Le numéro de licence initialiser dans ce module permet de le visualier dans la
'fenêtre "A propos de"
Public NumeroLicence As String

'modification du titre de l'appli variables reprises dans le load de la fenêtre principale
Public GmodifTitreApplication As String
Public GvisibiliteMnuBarre As Boolean
Public GvisibiliteMnuLicence As Boolean

'LicenceStatus permet de recevoir le résultat soit de la saisie du code
'par le biais de la fenêtre soit du fichier serial.txt
Private LicenceStatus As Boolean

'le str est seulement là pour plus de sécurité si on souhaite
'mettre ce module dans une dll
'fonction appelée à partir de main
'la fonction renvoie le numéro de licence
Public Function ProtectCheck(str As String) As String

    Dim licenceOK As Boolean
    
'choix de la langue
    initlang (VER_FRANCAISE)

'initialisation
    licenceOK = False
    
On Error GoTo FIN_ERR

    'Appel de la fonction de validation du serial en passant en paramètre le serial
    'présent dans serial.txt. La licence permet de maintenir à jour la fenêtre "A propose de"
    If Not VerifLicence("rien", "rien", LireTxt(NomFichierLicence), LireTxt(NomFichierSerial)) Then
         'pas valide donc on donne une chance à l'utilsateur de saisir le bon serial
         'lancement de la fenêtre de validation du serial
         frmKey.Show 1
    Else
    End If

    licenceOK = LicenceStatus

    'message de controle
    If str = "its00+-k" Then
    Else
        licenceOK = False
    End If
    
    'retourne le résultat
    If licenceOK Then
        ProtectCheck = "its00+-k"
    End If

    Exit Function

FIN_ERR:
    MsgBox Err.Description
    MsgBox F_MSGREGERROR1 & vbCrLf & F_MSGREGERROR2, vbCritical
End Function

'fonction de validation de la licence QLM appelée soit à partir de protectchk soit
'à partir de frmkey
Public Function VerifLicence(txt1 As String, txt2 As String, strlic As String, strserial As String) As Boolean
        Dim bret As Boolean
        Dim license As IsLicense.IsLicenseMgr
        
        Dim errorMsg As String
        Dim nStatus As Integer
        Dim licenseKey As String
        
        bret = False
        
        Select Case TYPPROTECTION
            Case QLM
                Set license = New IsLicense.IsLicenseMgr
                
                license.DefineProduct IdProduit, Produit, Vermaj, Vermin, message, Guid
                'license.DefineProduct 2, "OndeV", 1, 0, "its-ondev", "{7E84410F-0BD7-458D-AAB8-4879F6CF09D7}"
        
                'Get the license key from your user interface or from your config file.
                'Note that QLM does not store this key. It is up to you to store it and retrieve it
        
                errorMsg = license.ValidateLicense(strserial)
        
                nStatus = license.GetStatus()
        
                If IsTrue(nStatus, IsLicense.ELicenseStatus.EKeyInvalid) Or _
                    IsTrue(nStatus, IsLicense.ELicenseStatus.EKeyProductInvalid) Or _
                    IsTrue(nStatus, IsLicense.ELicenseStatus.EKeyVersionInvalid) Or _
                    IsTrue(nStatus, IsLicense.ELicenseStatus.EKeyMachineInvalid) Or _
                    IsTrue(nStatus, IsLicense.ELicenseStatus.EKeyTampered) Then
                
                    ' the key is invalid
                    '(errorMsg)
                    bret = False
        
                ElseIf (IsTrue(nStatus, IsLicense.ELicenseStatus.EKeyDemo)) Then
        
                    If (IsTrue(nStatus, IsLicense.ELicenseStatus.EKeyExpired)) Then
                        ' the key has expired
                        'MsgBox (errorMsg)
                        MsgBox MSGPWDEXPIRED
                        bret = False
        
                    Else
        
                        ' the demo key is still valid
                        MsgBox (errorMsg)
                        
                        'on ferme la fenêtre de saisie de licence
                        Unload frmKey
        
                        'Modification apportées à la fenêtre principale
                        'cette modification sera fite lors du chargement de la fenêtre
                        'On ajoute version DEMO au titre
                        GmodifTitreApplication = " version DEMO"
                        'le menu de saisie de licence devient visisble
                        GvisibiliteMnuBarre = True
                        GvisibiliteMnuLicence = True
                        
                        'on écrit le numero de licence = Version Demo dans licence.txt
                        If EcrireTxt("Version Demo", NomFichierLicence) Then
                            bret = True
                        End If
                        
                        'on écrit le numero de série dans serial.txt
                        If EcrireTxt(strserial, NomFichierSerial) Then
                            bret = True
                        End If
                        
                        'initialisation du numéro de licence
                        If bret Then
                            NumeroLicence = LireTxt(NomFichierLicence)
                        End If
        
                    End If
                ElseIf (IsTrue(nStatus, IsLicense.ELicenseStatus.EKeyPermanent) And strlic <> "") Then 'la condition strlic permet de maintenir la fen^tre A propos de à jour
        
                    ' the key is OK
                    'si ok écriture du serial dans serial.txt dans le répertoire d'installation de l'application
                  
                    Unload frmKey
                    
                    'Modification apportées à la fenêtre principale
                    'cette modification sera fite lors du chargement de la fenêtre
                    'On ajoute version DEMO au titre
                    GmodifTitreApplication = ""
                    'le menu de saisie de licence devient visisble
                    GvisibiliteMnuBarre = False
                    GvisibiliteMnuLicence = False
                    
                    'on écrit le numero de licence dans licence.txt
                    If EcrireTxt(strlic, NomFichierLicence) Then
                        bret = True
                    End If
                    
                    'on écrit le numero de série dans serial.txt
                    If EcrireTxt(strserial, NomFichierSerial) Then
                        bret = True
                    End If
                    
                    'initialisation du numéro de licence
                    If bret Then
                        NumeroLicence = LireTxt(NomFichierLicence)
                    End If
                End If
                
            Case CPM
                'Mise à jour du numéro de série
                If (strlic = "") Then
                    'première exécution
                    'mise à jour du numéro de série
                    bret = False
                Else
                    'si ok écriture du serial dans serial.txt dans le répertoire d'installation de l'application
                  
                    Unload frmKey
                    
                    'Modification apportées à la fenêtre principale
                    'cette modification sera fite lors du chargement de la fenêtre
                    GmodifTitreApplication = ""
                    'le menu de saisie de licence devient visisble
                    GvisibiliteMnuBarre = False
                    GvisibiliteMnuLicence = False
                    
                    'on écrit le numero de licence dans licence.txt
                    If EcrireTxt(strlic, NomFichierLicence) Then
                        bret = True
                    End If
                    
                    'on écrit le numero de série dans serial.txt
                    If EcrireTxt(strserial, NomFichierSerial) Then
                        bret = True
                    End If
                    
                    'initialisation du numéro de licence
                    If bret Then
                        NumeroLicence = LireTxt(NomFichierLicence)
                    End If
                End If
        End Select
        
        'mise à jour de LicenceStatus
        VerifLicence = bret
        LicenceStatus = bret

     End Function


'fonction appelée par la fonction de validation de la licence
 Private Function IsTrue(ByVal nVal1 As Integer, ByVal nVal2 As Integer) As Boolean

    If (((nVal1 And nVal2) = nVal1) Or ((nVal1 And nVal2) = nVal2)) Then

        IsTrue = True
        Exit Function
    End If

    IsTrue = False
    
End Function

'fonction permettant de lire dans les fichiers txt
Public Function LireTxt(nomfic As String) As String

    Dim fso
    
    Dim filenumber As Integer
    Dim nomfichier, myString As String
    
    LireTxt = ""
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    'nomfichier = App.Path & "\" & nomfic '& ".ctu"
    nomfichier = MonCorrigerNomFichier(App.Path & "\" & nomfic)
    
    If (fso.FileExists(nomfichier)) Then
    ' Lit le numéro de fichier inutilisé.
        filenumber = FreeFile
    ' Crée le nom du fichier.
        Open nomfichier For Input As #filenumber
        Do While Not EOF(filenumber)   ' Effectue la boucle jusqu'à la fin du fichier.
            Input #filenumber, myString   ' Lit les données dans variables.
    ' débug
            LireTxt = Trim(myString)
        Loop
        Close #filenumber   ' Ferme le fichier.
    Else
    End If

End Function

'fonction permettant d'écrire le serial dans serial.txt
Private Function EcrireTxt(Chaine As String, nomfic As String) As Boolean
    Dim fso, f
    Dim filenumber As Integer
    Dim nomfichier, myString As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    'nomfichier = App.Path & "\" & nomfic '& ".ctu"
    nomfichier = MonCorrigerNomFichier(App.Path & "\" & nomfic)
    
' Met le fichier en mode normal s'il existe
    If (fso.FileExists(nomfichier)) Then
        Set f = fso.GetFile(nomfichier)
        f.Attributes = 0 '0=normal
    End If
    
' Lit le numéro de fichier inutilisé.
    filenumber = FreeFile

' Crée le nom du fichier.
    Open nomfichier For Output As #filenumber
        Write #filenumber, Chaine

    Close #filenumber   ' Ferme le fichier.

' Met le fichier en fichier caché
    Set f = fso.GetFile(nomfichier)
    f.Attributes = 2 '2=hidden

    EcrireTxt = True
End Function


'initialisation des messages en fonction de la langue
Public Function initlang(langue As Integer) As Boolean

    If langue = VER_ANGLAISE Then
        Titre = E_TITRE
        Msg = E_MSG
        LBLICENCE = E_LBLICENCE
        LBSERIAL = E_LBSERIAL
        BtnOK = E_BTNOK
        btnCancel = E_BTNCANCEL
        MSGPWDINVALID = E_MSGPWDINVALID
        MSGPWDEXPIRED = E_MSGPWDEXPIRED
        MSGPWDVALID = E_MSGPWDVALID
        MSGREGERROR1 = E_MSGREGERROR1
        MSGREGERROR2 = E_MSGREGERROR2
    ElseIf langue = VER_FRANCAISE Then
        Titre = F_TITRE
        Msg = F_MSG
        LBLICENCE = F_LBLICENCE
        LBSERIAL = F_LBSERIAL
        BtnOK = F_BTNOK
        btnCancel = F_BTNCANCEL
        MSGPWDINVALID = F_MSGPWDINVALID
        MSGPWDEXPIRED = F_MSGPWDEXPIRED
        MSGPWDVALID = F_MSGPWDVALID
        MSGREGERROR1 = F_MSGREGERROR1
        MSGREGERROR2 = F_MSGREGERROR2
    End If
    
End Function

'OF : copie de la fonction présente dans utilitaire.bas de certians logiciels
Public Function MonCorrigerNomFichier(unFileName As String) As String
    'Fonction retournant un nom de fichier corrigé
    'de double / par un seul
    Dim unePos As Integer, uneStringRes As String
    
    unePos = 1
    uneStringRes = unFileName
    
    Do
        unePos = InStr(1, uneStringRes, "\\")
        If unePos > 0 Then
            uneStringRes = Mid(uneStringRes, 1, unePos) + Mid(uneStringRes, unePos + 2)
        End If
    Loop While unePos > 0
    
    MonCorrigerNomFichier = uneStringRes
End Function
