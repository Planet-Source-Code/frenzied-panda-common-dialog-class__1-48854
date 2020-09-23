VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetOpenFileName Lib "COMDLG32.DLL" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "COMDLG32.DLL" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private cdlg As OPENFILENAME
Private LastFileName As String
Private Type OPENFILENAME
  lStructSize       As Long
  hwndOwner         As Long
  hInstance         As Long
  lpstrFilter       As String
  lpstrCustomFilter As String
  nMaxCustFilter    As Long
  nFilterIndex      As Long
  lpstrFile         As String
  nMaxFile          As Long
  lpstrFileTitle    As String
  nMaxFileTitle     As Long
  lpstrInitialDir   As String
  lpstrTitle        As String
  Flags             As Long
  nFileOffset       As Integer
  nFileExtension    As Integer
  lpstrDefExt       As String
  lCustData         As Long
  lpfnHook          As Long
  lpTemplateName    As String
End Type

Private Const OFN_ALLOWMULTISELECT = &H200
Private Const OFN_CREATEPROMPT = &H2000
Private Const OFN_ENABLEHOOK = &H20
Private Const OFN_ENABLETEMPLATE = &H40
Private Const OFN_ENABLETEMPLATEHANDLE = &H80
Private Const OFN_EXPLORER = &H80000                         '  new look commdlg
Private Const OFN_EXTENSIONDIFFERENT = &H400
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_LONGNAMES = &H200000                       '  force long names for 3.x modules
Private Const OFN_NOCHANGEDIR = &H8
Private Const OFN_NODEREFERENCELINKS = &H100000
Private Const OFN_NOLONGNAMES = &H40000                      '  force no long names for 4.x modules
Private Const OFN_NONETWORKBUTTON = &H20000
Private Const OFN_NOREADONLYRETURN = &H8000
Private Const OFN_NOTESTFILECREATE = &H10000
Private Const OFN_NOVALIDATE = &H100
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_READONLY = &H1
Private Const OFN_SHAREAWARE = &H4000
Private Const OFN_SHAREFALLTHROUGH = 2
Private Const OFN_SHARENOWARN = 1
Private Const OFN_SHAREWARN = 0
Private Const OFN_SHOWHELP = &H10

Public Enum DialogFlags
  ALLOWMULTISELECT = OFN_ALLOWMULTISELECT
  CREATEPROMPT = OFN_CREATEPROMPT
  ENABLEHOOK = OFN_ENABLEHOOK
  ENABLETEMPLATE = OFN_ENABLETEMPLATE
  ENABLETEMPLATEHANDLE = OFN_ENABLETEMPLATEHANDLE
  EXPLORER = OFN_EXPLORER
  EXTENSIONDIFFERENT = OFN_EXTENSIONDIFFERENT
  FILEMUSTEXIST = OFN_FILEMUSTEXIST
  HIDEREADONLY = OFN_HIDEREADONLY
  LONGNAMES = OFN_LONGNAMES
  NOCHANGEDIR = OFN_NOCHANGEDIR
  NODEREFERENCELINKS = OFN_NODEREFERENCELINKS
  NOLONGNAMES = OFN_NOLONGNAMES
  NONETWORKBUTTON = OFN_NONETWORKBUTTON
  NOREADONLYRETURN = OFN_NOREADONLYRETURN
  NOTESTFILECREATE = OFN_NOTESTFILECREATE
  NOVALIDATE = OFN_NOVALIDATE
  OVERWRITEPROMPT = OFN_OVERWRITEPROMPT
  PATHMUSTEXIST = OFN_PATHMUSTEXIST
  ReadOnly = OFN_READONLY
  SHAREAWARE = OFN_SHAREAWARE
  SHAREFALLTHROUGH = OFN_SHAREFALLTHROUGH
  SHARENOWARN = OFN_SHARENOWARN
  SHAREWARN = OFN_SHAREWARN
  SHOWHELP = OFN_SHOWHELP
End Enum


Public Function ShowOpen(ByVal Form_hWnd As Long, ByVal Title As String, Optional ByVal InitDir As String = "", Optional ByVal Filter As String = "All Files (*.*)|*.*|", Optional Flags As DialogFlags = FILEMUSTEXIST Or PATHMUSTEXIST) As String
  Dim i As Integer
  Filter = Replace(Filter, "|", Chr(0))
  If Right(Filter, 1) <> Chr(0) Then Filter = Filter & Chr(0)
  If Len(InitDir) = 0 Then InitDir = LastFileName
  cdlg.lStructSize = Len(cdlg)
  cdlg.hwndOwner = Form_hWnd
  cdlg.hInstance = App.hInstance
  cdlg.lpstrFilter = Filter
  cdlg.lpstrFile = Space(254)
  cdlg.nMaxFile = 255
  cdlg.lpstrFileTitle = Space(254)
  cdlg.nMaxFileTitle = 255
  cdlg.lpstrInitialDir = InitDir
  cdlg.lpstrTitle = Title
  cdlg.Flags = Flags
  ShowOpen = IIf(GetOpenFileName(cdlg), Trim(cdlg.lpstrFile), "")
  If Len(ShowOpen) > 0 Then LastFileName = ShowOpen
End Function

Public Function ShowSave(ByVal Form_hWnd As Long, ByVal Title As String, Optional ByVal InitDir As String = "", Optional ByVal Filter As String = "All Files (*.*)|*.*|", Optional ByVal DefExt As String = "", Optional Flags As DialogFlags = OVERWRITEPROMPT) As String
  Dim i As Integer
  Filter = Replace(Filter, "|", Chr(0))
  If Right(Filter, 1) <> Chr(0) Then Filter = Filter & Chr(0)
  If Len(InitDir) = 0 Then InitDir = LastFileName
  cdlg.lStructSize = Len(cdlg)
  cdlg.lpstrTitle = Title
  cdlg.hwndOwner = Form_hWnd
  cdlg.hInstance = App.hInstance
  cdlg.lpstrFilter = Filter
  cdlg.lpstrDefExt = DefExt
  cdlg.lpstrFile = Space(255)
  cdlg.nMaxFile = 255
  cdlg.lpstrFileTitle = Space(254)
  cdlg.nMaxFileTitle = 255
  cdlg.lpstrInitialDir = InitDir
  cdlg.Flags = Flags
  ShowSave = IIf(GetSaveFileName(cdlg), Trim(cdlg.lpstrFile), "")
  If Len(ShowSave) > 0 Then LastFileName = ShowSave
End Function

Public Property Let FileName(szFileName)
  LastFileName = szFileName
End Property

Public Property Get FileName()
  FileName = LastFileName
End Property

