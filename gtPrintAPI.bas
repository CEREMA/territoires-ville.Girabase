Attribute VB_Name = "PrintApi"
Option Explicit
'**********************************************************************
'        Tous Projets
'
'         Module : PrintAPI
'
'         Septembre 2000
'
'         A. Vignaud  CETE de l'Ouest - DIOG/Its
'**********************************************************************

Public gDlgPrint As MSComDlg.CommonDialog
Public gPrinter As String

' Constantes pour la Plateforme Système
Public Const VER_PLATFORM_WIN32_NT = 2
Public Const VER_PLATFORM_WIN32_WINDOWS = 1
Public Const VER_PLATFORM_WIN32s = 0

' Constantes pour la structure DEVMODE
Public Const CCHDEVICENAME = 32
Public Const CCHFORMNAME = 32

Public Const DM_ORIENTATION = &H1&
Public Const DM_PAPERSIZE = &H2&
Public Const DM_SCALE = &H10&
Public Const DM_COPIES = &H100&
Public Const DM_PRINTQUALITY = &H400&
Public Const DM_COLOR = &H800&
Public Const DM_DUPLEX = &H1000&

' Constante pour la structure DEVNAMES
Public Const DN_DEFAULTPRN = &H1

' Constantes pour la structure PRINTDLG
Public Const PD_ALLPAGES = &H0
Public Const PD_COLLATE = &H10
Public Const PD_DISABLEPRINTTOFILE = &H80000
Public Const PD_ENABLEPRINTHOOK = &H1000
Public Const PD_ENABLEPRINTTEMPLATE = &H4000
Public Const PD_ENABLEPRINTTEMPLATEHANDLE = &H10000
Public Const PD_ENABLESETUPHOOK = &H2000
Public Const PD_ENABLESETUPTEMPLATE = &H8000
Public Const PD_ENABLESETUPTEMPLATEHANDLE = &H20000
Public Const PD_HIDEPRINTTOFILE = &H100000
Public Const PD_NONETWORKBUTTON = &H200000
Public Const PD_NOPAGENUMS = &H8
Public Const PD_NOSELECTION = &H4
Public Const PD_NOWARNING = &H80
Public Const PD_PAGENUMS = &H2
Public Const PD_PRINTSETUP = &H40
Public Const PD_PRINTTOFILE = &H20
Public Const PD_RETURNDC = &H100
Public Const PD_RETURNDEFAULT = &H400
Public Const PD_RETURNIC = &H200
Public Const PD_SELECTION = &H1
Public Const PD_SHOWHELP = &H800
Public Const PD_USEDEVMODECOPIES = &H40000
Public Const PD_USEDEVMODECOPIESANDCOLLATE = &H40000

' Constantes pour la structure PAGESETUPDLG
Public Const PSD_DEFAULTMINMARGINS = &H0
Public Const PSD_DISABLEMARGINS = &H10
Public Const PSD_DISABLEORIENTATION = &H100
Public Const PSD_DISABLEPAGEPAINTING = &H80000
Public Const PSD_DISABLEPAPER = &H200
Public Const PSD_DISABLEPRINTER = &H20
Public Const PSD_ENABLEPAGEPAINTHOOK = &H40000
Public Const PSD_ENABLEPAGESETUPHOOK = &H2000
Public Const PSD_ENABLEPAGESETUPTEMPLATE = &H8000
Public Const PSD_ENABLEPAGESETUPTEMPLATEHANDLE = &H20000
Public Const PSD_INHUNDREDTHSOFMILLIMETERS = &H8
Public Const PSD_INTHOUSANDTHSOFINCHES = &H4
Public Const PSD_INWININIINTLMEASURE = &H0
Public Const PSD_MARGINS = &H2
Public Const PSD_MINMARGINS = &H1
Public Const PSD_NOWARNING = &H80
Public Const PSD_RETURNDEFAULT = &H400
Public Const PSD_SHOWHELP = &H800

' Constantes pour l'allocation mémoire
Public Const GMEM_FIXED = &H0
Public Const GMEM_MOVEABLE = &H2
Public Const GMEM_ZEROINIT = &H40

Public Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128      '  Maintenance string for PSS usage
End Type


Public Type POINTAPI
        x As Long
        Y As Long
End Type

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Type PAGESETUPDLG_TYPE
        lStructSize As Long
        hwndOwner As Long
        hDevMode As Long
        hDevNames As Long
        flags As Long
        ptPaperSize As POINTAPI
        rtMinMargin As RECT
        rtMargin As RECT
        hInstance As Long
        lCustData As Long
        lpfnPageSetupHook As Long
        lpfnPagePaintHook As Long
        lpPageSetupTemplateName As String
        hPageSetupTemplate As Long
End Type

'type definitions:
Type PRINTDLG_TYPE
        lStructSize As Long
        hwndOwner As Long
        hDevMode As Long
        hDevNames As Long
        hDC As Long
        flags As Long
        nFromPage As Integer
        nToPage As Integer
        nMinPage As Integer
        nMaxPage As Integer
        nCopies As Integer
        hInstance As Long
        lCustData As Long
        lpfnPrintHook As Long

        lpfnSetupHook As Long
        lpPrintTemplateName As String
        lpSetupTemplateName As String
        hPrintTemplate As Long
        hSetupTemplate As Long
End Type


Public Type DEVMODE_TYPE
        dmDeviceName As String * CCHDEVICENAME
        dmSpecVersion As Integer
        dmDriverVersion As Integer
        dmSize As Integer
        dmDriverExtra As Integer
        dmFields As Long
        dmOrientation As Integer
        dmPaperSize As Integer
        dmPaperLength As Integer
        dmPaperWidth As Integer
        dmScale As Integer
        dmCopies As Integer
        dmDefaultSource As Integer
        dmPrintQuality As Integer
        dmColor As Integer
        dmDuplex As Integer
        dmYResolution As Integer
        dmTTOption As Integer
        dmCollate As Integer
        dmFormName As String * CCHFORMNAME    ' utilisé que par NT
        ' les items suivants ne concernent pas les imprimantes
        dmUnusedPadding As Integer
        dmBitsPerPel As Long
        dmPelsWidth As Long
        dmPelsHeight As Long
        dmDisplayFlags As Long
        dmDisplayFrequency As Long
End Type

Public Type DEVNAMES_TYPE
        wDriverOffset As Integer
        wDeviceOffset As Integer
        wOutputOffset As Integer
        wDefault As Integer
        wInfo As String * 100
End Type

Public Type PRINTER_DEFAULTS
        pDatatype As String
        pDevMode As DEVMODE_TYPE
        DesiredAccess As Long
End Type

Public Declare Function GetVersionEx Lib "Kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Public Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

Public Declare Function GlobalLock Lib "Kernel32" (ByVal hMem As Long) As Long

Public Declare Function GlobalUnlock Lib "Kernel32" (ByVal hMem As Long) As Long

Public Declare Function GlobalAlloc Lib "Kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long

Public Declare Function GlobalFree Lib "Kernel32" (ByVal hMem As Long) As Long

Public Declare Function PageSetupDlg Lib "comdlg32.dll" Alias "PageSetupDlgA" (pPageSetupDlg As PAGESETUPDLG_TYPE) As Long

Public Declare Function dlgPrint Lib "comdlg32.dll" Alias "PrintDlgA" (pPrintdlg As PRINTDLG_TYPE) As Long
   
Public Declare Function PrinterProperties Lib "winspool.drv" (ByVal hwnd As Long, ByVal hPrinter As Long) As Long
   
Public Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, pDefault As PRINTER_DEFAULTS) As Long
Public Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Public Declare Function CreateDC Lib "GDI32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As DEVMODE_TYPE) As Long
   
Public Declare Function WritePrinter Lib "winspool.drv" (ByVal hPrinter As Long, pBuf As Any, ByVal cdBuf As Long, pcWritten As Long) As Long

Public Declare Function StartPage Lib "GDI32" (ByVal hDC As Long) As Long
Public Declare Function EndPage Lib "GDI32" (ByVal hDC As Long) As Long
'Public Declare Function StartDoc Lib "gdi32" Alias "StartDocA" (ByVal hDC As Long, lpdi As DOCINFO) As Long
Public Declare Function EndDoc Lib "GDI32" (ByVal hDC As Long) As Long

Public Declare Function GetLastError Lib "Kernel32" () As Long
   
''*******************************************************************
'' Détermine si Windows 95
''*******************************************************************
'Public Function SystemW95() As Boolean
'
'Dim VersionInfo As OSVERSIONINFO
'
'VersionInfo.dwOSVersionInfoSize = Len(VersionInfo)
'If GetVersionEx(VersionInfo) Then
'  SystemW95 = (VersionInfo.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS And VersionInfo.dwMinorVersion = 0)
'End If
'
'End Function

'*******************************************************************
' Détermine si la plateforme est NT
'*******************************************************************
Public Function PlateformeNT() As Boolean

Dim VersionInfo As OSVERSIONINFO

' 11/06/2002  : Même sous NT certains utilisateurs ont des pb : abandon définitif de la solution
Exit Function

VersionInfo.dwOSVersionInfoSize = Len(VersionInfo)
If GetVersionEx(VersionInfo) Then
  PlateformeNT = (VersionInfo.dwPlatformId = VER_PLATFORM_WIN32_NT)
End If

'PlateformeNT = True

End Function
   
'*****************************************************************************************************
' Initialisation de l'imprimante(sous NT) avec la base de registres, sous réserve de compatibilité
' 1 : L'imprimante est renseignée - 2 Elle existe (fait toujours partie de la liste des imprimantes du user)
'*****************************************************************************************************
Public Sub InitializePrinter()
    'Récup des paramètres de config d'imprimante pour ce logiciel
  Dim PrinterDeviceName As String
  Dim objPrinter As Printer
  Static DéjaAppelé As Boolean

  If DéjaAppelé Then Exit Sub
  If Printers.Count = 0 Then Exit Sub
  DéjaAppelé = True
  
  On Error GoTo GestErr
    
  gPrinter = Printer.DeviceName
  
  #If Not okW95 Then
'    If Not PlateformeNT Then Exit Sub
  #End If
  
  Screen.MousePointer = vbHourglass
  PrinterDeviceName = GetSetting(Appname:=App.Title, Section:="PrintOptions", Key:="DeviceName")
  
  If PrinterDeviceName = Printer.DeviceName Then
    LireRegistryPrn
  ElseIf PrinterDeviceName <> "" Then
    For Each objPrinter In Printers
      Debug.Print objPrinter.DeviceName
      If TronqueChaine(UCase(objPrinter.DeviceName), CCHDEVICENAME - 1) = TronqueChaine(UCase(PrinterDeviceName), CCHDEVICENAME - 1) Then
        Debug.Print "Je tente de changer d'imprimante courante" & vbCrLf & objPrinter.DeviceName & " - driver : " & objPrinter.DriverName
        Set Printer = objPrinter
        DoEvents
        LireRegistryPrn
        Exit For
      End If
    Next
  End If

  Debug.Print "Vérifier la fonte : " & Printer.Font.name
  If Printer.Font.Charset <> 0 Then
    Printer.Font.name = "Arial"
    Debug.Print "Passage en Arial réussi"
  End If
  
  Screen.MousePointer = vbDefault
  
Exit Sub

GestErr:
  
End Sub

Public Sub ReInitializePrinter()
Dim objPrinter As Printer

  ' Pas de pb de réinitialisation sous NT : l'imprimante par défaut n'a pas été changée
  If PlateformeNT Or gPrinter = "" Then Exit Sub

  If gPrinter <> Printer.DeviceName Then
    For Each objPrinter In Printers
      If objPrinter.DeviceName = gPrinter Then
        Set Printer = objPrinter
        DoEvents
        Exit For
      End If
    Next
  End If
  
End Sub

'*****************************************************************************************************
' Lecture (sous NT) dans la base de registres du nom de l'imprimante GIRATION et de ses propriétés
'*****************************************************************************************************
Public Sub LireRegistryPrn()
    
  If Not PlateformeNT Then Exit Sub
  ' Gestionnaire d'Erreur armé pour contourner une propriété non gérée par l'imprimante
  On Error Resume Next
  With Printer
    Debug.Print "Recherche de l'orientation"
    .Orientation = GetSetting(Appname:=App.Title, Section:="PrintOptions", Key:="Orientation", Default:=.Orientation)
    Debug.Print "Recherche de Copies"
    .Copies = GetSetting(Appname:=App.Title, Section:="PrintOptions", Key:="Copies", Default:=.Copies)
    Debug.Print "Recherche de Duplex"
    .Duplex = GetSetting(Appname:=App.Title, Section:="PrintOptions", Key:="Duplex", Default:=.Duplex)
    Debug.Print "Recherche de PaperSize"
    .PaperSize = GetSetting(Appname:=App.Title, Section:="PrintOptions", Key:="PaperSize", Default:=.PaperSize)
    Debug.Print "Recherche de Colormode"
    .ColorMode = GetSetting(Appname:=App.Title, Section:="PrintOptions", Key:="ColorMode", Default:=.ColorMode)
    Debug.Print "Recherche de Zoom"
    .Zoom = GetSetting(Appname:=App.Title, Section:="PrintOptions", Key:="Zoom", Default:=.Zoom)
    Debug.Print "Recherche de PrintQuality"
    .PrintQuality = GetSetting(Appname:=App.Title, Section:="PrintOptions", Key:="PrintQuality", Default:=.PrintQuality)
  End With

End Sub

'*******************************************************************
' Sauvegarde (sous NT) de registres du nom de l'imprimante GIRATION et de ses propriétés
'*******************************************************************
Private Sub EcrireRegistryPRN()
  SaveSetting Appname:=App.Title, Section:="PrintOptions", Key:="DeviceName", Setting:=Printer.DeviceName
  If Not PlateformeNT Then Exit Sub
    
    ' Gestionnaire d'Erreur armé pour contourner une propriété non gérée par l'imprimante
    On Error Resume Next
    With Printer
        SaveSetting Appname:=App.Title, Section:="PrintOptions", Key:="Orientation", Setting:=.Orientation
        SaveSetting Appname:=App.Title, Section:="PrintOptions", Key:="Copies", Setting:=.Copies
        SaveSetting Appname:=App.Title, Section:="PrintOptions", Key:="Duplex", Setting:=.Duplex
        SaveSetting Appname:=App.Title, Section:="PrintOptions", Key:="PaperSize", Setting:=.PaperSize
        SaveSetting Appname:=App.Title, Section:="PrintOptions", Key:="ColorMode", Setting:=.ColorMode
        SaveSetting Appname:=App.Title, Section:="PrintOptions", Key:="Zoom", Setting:=.Zoom
        SaveSetting Appname:=App.Title, Section:="PrintOptions", Key:="PrintQuality", Setting:=.PrintQuality
    End With
  

End Sub
   
'*******************************************************************
' Sous W95 ou 98 : Appel de la méthode du Commondialog Standard de VB (COMDLGG32.OCX)
' Sous NT : réécriture de la fonction par des appels API
' Dans tous les cas : Cancel=False si l'utilisateur fait Annuler ou ferme la boite avec le bouton Fermer
'*******************************************************************

Public Sub ShowPrinter(frmOwner As Form, Optional ByRef Cancel As Integer)

  Dim PrintDlg As PRINTDLG_TYPE
  Dim DEVMODE As DEVMODE_TYPE
  Dim Devname As DEVNAMES_TYPE
'    Dim PageSetupDialog As PAGESETUPDLG_TYPE
       
  Dim lpDevMode As Long, lpDevName As Long
  Dim bReturn As Integer
  Dim objPrinter As Printer, NewPrinterName As String

  InitializePrinter
  
'  If Not PlateformeNT Then
'    On Error GoTo ErrHandler
'    gDlgPrint.flags = cdlPDPrintSetup   ' Pour afficher directement la fenêtre Configuration
'    gDlgPrint.ShowPrinter
'    MsgBox Printer.DeviceName
'    EcrireRegistryPRN
'    Exit Sub
'
'ErrHandler:
'    MsgBox Err.Number
'
'    If Err <> cdlCancel And Err <> cdlNoDefaultPrn Then
''      MsgBox IDm_ErrImprim & " (" & Format(Err.Number, "000") & ")" & vbCrLf & Err.Description
'      MsgBox "erreur imprimante" & " (" & Format(Err.Number, "000") & ")" & vbCrLf & Err.Description
'    End If
'    Cancel = True
'   Exit Sub
'  End If


  ' Use dlgprint to get the handle to a memory
  ' block with a DevMode and DevName structures

  PrintDlg.lStructSize = Len(PrintDlg)
  PrintDlg.hwndOwner = frmOwner.hwnd
  PrintDlg.flags = PD_PRINTSETUP

'    PageSetupDialog.flags = PSD_DISABLEMARGINS Or PSD_MARGINS Or PSD_DISABLEPAGEPAINTING Or PSD_INHUNDREDTHSOFMILLIMETERS
'    With PageSetupDialog.rtMargin
'      .Bottom = 500
'      .Left = 500
'      .Right = 500
'      .Top = 500
'    End With

  'Set the current orientation, duplex, papersize, etc... setting
  DEVMODE.dmDeviceName = Printer.DeviceName
  DEVMODE.dmSize = Len(DEVMODE)
  
  'On initialize avec les valeurs du PRINTER par défaut
  DEVMODE.dmFields = DM_ORIENTATION Or DM_COPIES Or DM_DUPLEX Or DM_PAPERSIZE Or DM_COLOR Or DM_SCALE Or DM_PRINTQUALITY
  On Error Resume Next
  With Printer
    DEVMODE.dmOrientation = .Orientation
    DEVMODE.dmCopies = .Copies
    DEVMODE.dmDuplex = .Duplex
    DEVMODE.dmPaperSize = .PaperSize
    DEVMODE.dmColor = .ColorMode
    DEVMODE.dmScale = .Zoom
    DEVMODE.dmPrintQuality = .PrintQuality
  End With
  On Error GoTo 0
  
  'Allocate memory for the initialization hDevMode structure
  'and copy the settings gathered above into this memory
  PrintDlg.hDevMode = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(DEVMODE))
  lpDevMode = GlobalLock(PrintDlg.hDevMode)
  If lpDevMode > 0 Then
    CopyMemory ByVal lpDevMode, DEVMODE, Len(DEVMODE)
    bReturn = GlobalUnlock(PrintDlg.hDevMode)
  End If

  'Set the current driver, device, and port name strings
  With Devname
    .wDriverOffset = 8
    .wDeviceOffset = .wDriverOffset + 1 + Len(Printer.DriverName)
    .wOutputOffset = .wDeviceOffset + 1 + Len(Printer.Port)
    .wDefault = 0 ' 0 'DN_DEFAULTPRN(1)
  End With
  With Printer
    Devname.wInfo = .DriverName & Chr(0) & _
    .DeviceName & Chr(0) & .Port & Chr(0)
  End With

  'Allocate memory for the initial hDevName structure
  'and copy the settings gathered above into this memory
  PrintDlg.hDevNames = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(Devname))
  lpDevName = GlobalLock(PrintDlg.hDevNames)
  If lpDevName > 0 Then
    CopyMemory ByVal lpDevName, Devname, Len(Devname)
    bReturn = GlobalUnlock(lpDevName)
  End If

  Printer.ScaleMode = vbCentimeters
  
  'Call the print dialog up and let the user make changes
'    If PageSetupDlg(PageSetupDialog) Then
  If dlgPrint(PrintDlg) Then
     'First get the DevName structure.
     lpDevName = GlobalLock(PrintDlg.hDevNames)
     CopyMemory Devname, ByVal lpDevName, 45
     bReturn = GlobalUnlock(lpDevName)
     GlobalFree PrintDlg.hDevNames

     'Next get the DevMode structure and set the printer
     'properties appropriately
     lpDevMode = GlobalLock(PrintDlg.hDevMode)
     CopyMemory DEVMODE, ByVal lpDevMode, Len(DEVMODE)
     bReturn = GlobalUnlock(PrintDlg.hDevMode)
     GlobalFree PrintDlg.hDevMode
     NewPrinterName = UCase(suppCNull(DEVMODE.dmDeviceName))
     
     If TronqueChaine(UCase(Printer.DeviceName), CCHDEVICENAME - 1) <> NewPrinterName Then
        For Each objPrinter In Printers
          If TronqueChaine(UCase(objPrinter.DeviceName), CCHDEVICENAME - 1) = NewPrinterName Then
            Set Printer = objPrinter
            DoEvents
            Exit For
          End If
        Next
     End If
     
     On Error Resume Next

     'Set printer object properties according to selections made
     'by user
  '   DoEvents
     With Printer
         .Copies = DEVMODE.dmCopies
         .Duplex = DEVMODE.dmDuplex
         .Orientation = DEVMODE.dmOrientation
         .PaperSize = DEVMODE.dmPaperSize
         .ColorMode = DEVMODE.dmColor
         .Zoom = DEVMODE.dmScale
         .PrintQuality = DEVMODE.dmPrintQuality
     End With
     On Error GoTo 0
     EcrireRegistryPRN
  Else
    Cancel = True
  End If


End Sub

Public Function TronqueChaine(ByVal chaine As String, ByVal LgChaine As Integer)
  If Len(chaine) < LgChaine Then LgChaine = Len(chaine)
  TronqueChaine = Left(chaine, LgChaine)
End Function

