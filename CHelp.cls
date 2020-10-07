VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Enum HH_COMMAND
    HH_DISPLAY_TOPIC = &H0
    HH_HELP_FINDER = &H0
    HH_DISPLAY_TOC = &H1
    HH_DISPLAY_INDEX = &H2
    HH_DISPLAY_SEARCH = &H3
    HH_SET_WIN_TYPE = &H4
    HH_GET_WIN_TYPE = &H5
    HH_GET_WIN_HANDLE = &H6
    HH_GET_INFO_TYPES = &H7
    HH_SET_INFO_TYPES = &H8
    HH_SYNC = &H9
    HH_ADD_NAV_UI = &HA
    HH_ADD_BUTTON = &HB
    HH_GETBROWSER_APP = &HC
    HH_KEYWORD_LOOKUP = &HD
    HH_DISPLAY_TEXT_POPUP = &HE
    HH_HELP_CONTEXT = &HF
    HH_TP_HELP_CONTEXTMENU = &H10
    HH_TP_HELP_WM_HELP = &H11
    HH_CLOSE_ALL = &H12
    HH_ALINK_LOOKUP = &H13
End Enum

Private Type structHH_FTS_QUERY
    cbStruct As Long
    fUniCodeStrings As Long
    pszSearchQuery As String
    iProximity As Long
    fStemmedSearch As Long
    fTitleOnly As Long
    fExecute As Long
    pszWindow As String
End Type

Private Declare Function HtmlHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" _
        (ByVal hwndCaller As Long, _
        ByVal pszFile As String, _
        ByVal uCommand As HH_COMMAND, _
        dwData As Any) As Long

Public Sub Show(NewFile As String, Optional WindowPane As String, Optional ByRef ContextID)
    Dim Fichier As String
    Fichier = NewFile

    If Len(WindowPane) Then
        Fichier = Trim(Fichier) & ">" & Trim(WindowPane)
    End If

    If IsMissing(ContextID) Then
        Call HtmlHelp(0, Fichier, HH_DISPLAY_TOC, 0&)
    Else
        Call HtmlHelp(0, Fichier, HH_HELP_CONTEXT, ByVal CLng(ContextID))
    End If
End Sub

Public Sub ShowIndex(NewFile As String, Optional WindowPane As String)
    Dim Fichier As String
    Fichier = NewFile

    If Len(WindowPane) Then
        Fichier = Trim(Fichier) & ">" & Trim(WindowPane)
    End If

    Call HtmlHelp(0, Fichier, HH_DISPLAY_INDEX, ByVal 0&)
End Sub

Public Sub ShowSearch(NewFile As String, Optional WindowPane As String)

    Dim Fichier As String
    Dim HH_FTS_QUERY As structHH_FTS_QUERY
    With HH_FTS_QUERY
        .cbStruct = Len(HH_FTS_QUERY)
        .fUniCodeStrings = 0&
        .pszSearchQuery = "TEST"
        .iProximity = 0&
        .fStemmedSearch = 0&
        .fTitleOnly = 0&
        .fExecute = 1&
        .pszWindow = ""
    End With
    Fichier = NewFile

    If Len(WindowPane) Then
        Fichier = Trim(Fichier) & ">" & Trim(WindowPane)
    End If
    Call HtmlHelp(0, Fichier, HH_DISPLAY_SEARCH, HH_FTS_QUERY)

End Sub

