Attribute VB_Name = "modSystem"
Option Explicit

Global Const SIZE_KILO As Long = 1024
Global Const SIZE_MEGA As Long = SIZE_KILO * SIZE_KILO

Private Const OFN_ALLOWMULTISELECT As Long = &H200
Private Const OFN_CREATEPROMPT As Long = &H2000
Private Const OFN_ENABLEHOOK As Long = &H20
Private Const OFN_ENABLETEMPLATE As Long = &H40
Private Const OFN_ENABLETEMPLATEHANDLE As Long = &H80
Private Const OFN_EXPLORER As Long = &H80000
Private Const OFN_EXTENSIONDIFFERENT As Long = &H400
Private Const OFN_FILEMUSTEXIST As Long = &H1000
Private Const OFN_HIDEREADONLY As Long = &H4
Private Const OFN_LONGNAMES As Long = &H200000
Private Const OFN_NOCHANGEDIR As Long = &H8
Private Const OFN_NODEREFERENCELINKS As Long = &H100000
Private Const OFN_NOLONGNAMES As Long = &H40000
Private Const OFN_NONETWORKBUTTON As Long = &H20000
Private Const OFN_NOREADONLYRETURN As Long = &H8000& 'see comments
Private Const OFN_NOTESTFILECREATE As Long = &H10000
Private Const OFN_NOVALIDATE As Long = &H100
Private Const OFN_OVERWRITEPROMPT As Long = &H2
Private Const OFN_PATHMUSTEXIST As Long = &H800
Private Const OFN_READONLY As Long = &H1
Private Const OFN_SHAREAWARE As Long = &H4000
Private Const OFN_SHAREFALLTHROUGH As Long = 2
Private Const OFN_SHAREWARN As Long = 0
Private Const OFN_SHARENOWARN As Long = 1
Private Const OFN_SHOWHELP As Long = &H10
Private Const OFS_MAXPATHNAME As Long = 260

Private Const OFS_FILE_OPEN_FLAGS = OFN_EXPLORER _
             Or OFN_LONGNAMES _
             Or OFN_CREATEPROMPT _
             Or OFN_NODEREFERENCELINKS _
             Or OFN_FILEMUSTEXIST

Private Const OFS_FILE_SAVE_FLAGS = OFN_EXPLORER _
             Or OFN_LONGNAMES _
             Or OFN_OVERWRITEPROMPT _
             Or OFN_HIDEREADONLY

Global Const BIF_RETURNONLYFSDIRS = &H1
Global Const BIF_DONTGOBELOWDOMAIN = &H2
Global Const BIF_STATUSTEXT = &H4
Global Const BIF_RETURNFSANCESTORS = &H8
Global Const BIF_BROWSEFORCOMPUTER = &H1000
Global Const BIF_BROWSEFORPRINTER = &H2000
Global Const MAX_PATH As Long = 260

Private Const INVALID_HANDLE_VALUE = -1

Private Const FILESIZE_FIX = 4294967296#

Private Const HWND_NOTOPMOST  As Long = -2
Private Const HWND_TOPMOST  As Long = -1
Private Const SWP_NOMOVE  As Long = &H2
Private Const SWP_NOSIZE  As Long = &H1
Private Const SWP_SETWINDOWPOS  As Long = SWP_NOSIZE Or SWP_NOMOVE

Private Type OPENFILENAME
  nStructSize       As Long
  hwndOwner         As Long
  hInstance         As Long
  sFilter           As String
  sCustomFilter     As String
  nMaxCustFilter    As Long
  nFilterIndex      As Long
  sFile             As String
  nMaxFile          As Long
  sFileTitle        As String
  nMaxTitle         As Long
  sInitialDir       As String
  sDialogTitle      As String
  Flags             As Long
  nFileOffset       As Integer
  nFileExtension    As Integer
  sDefFileExt       As String
  nCustData         As Long
  fnHook            As Long
  sTemplateName     As String
End Type

Private Type tagINITCOMMONCONTROLSEX
  dwSize As Long
  dwICC As Long
End Type

Type FILETIME
   dwLowDateTime As Long
   dwHighDateTime As Long
End Type

Public Type SHFILEINFO
   hIcon As Long
   iIcon As Long
   dwAttributes As Long
   szDisplayName As String * MAX_PATH
   szTypeName As String * 80
End Type

Public Type WIN32_FIND_DATA
   dwFileAttributes As Long
   ftCreationTime As FILETIME
   ftLastAccessTime As FILETIME
   ftLastWriteTime As FILETIME
   nFileSizeHigh As Long
   nFileSizeLow As Long
   dwReserved0 As Long
   dwReserved1 As Long
   cFileName As String * MAX_PATH
   cAlternate As String * 14
End Type

Public Type OSVERSIONINFO
  dwOSVersionInfoSize As Long
  dwMajorVersion As Long
  dwMinorVersion As Long
  dwBuildNumber As Long
  dwPlatformId As Long
  szCSDVersion(0 To 127) As Byte
End Type

Private Declare Function GetOpenFileName Lib "comdlg32" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
Private Declare Sub InitCommonControls9x Lib "comctl32" Alias "InitCommonControls" ()
Private Declare Function InitCommonControlsEx Lib "comctl32" (lpInitCtrls As tagINITCOMMONCONTROLSEX) As Boolean

Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long

Private Declare Function RtlGetVersion Lib "ntdll.dll" (lpVersionInformation As OSVERSIONINFO) As Long

Public Function BrowseForFileA(Title As String, hwndOwner As Long) As String
  Dim ofn As OPENFILENAME
  Dim sFilter As String
  Dim pos As Long
  Dim buff As String
  Dim sLongName As String
  Dim sShortName As String
  sFilter = "All Files" & vbNullChar & "*.*" & vbNullChar & vbNullChar
  With ofn
    .nStructSize = Len(ofn)
    .hwndOwner = hwndOwner
    .sFilter = sFilter
    .nFilterIndex = 0
    .sFile = Space(1024) & vbNullChar & vbNullChar
    .nMaxFile = Len(.sFile)
    .sDefFileExt = "" & vbNullChar & vbNullChar
    .sFileTitle = vbNullChar & Space(512) & vbNullChar & vbNullChar
    .nMaxTitle = Len(.sFileTitle)
    .sInitialDir = vbNullChar & vbNullChar
    .sDialogTitle = Title
    .Flags = OFS_FILE_OPEN_FLAGS Or OFN_ALLOWMULTISELECT
  End With
  If GetOpenFileName(ofn) Then
    BrowseForFileA = Trim$(Left$(ofn.sFile, (Len(ofn.sFile) - 2)))
  End If
End Function

Public Function InitCommonControls(Optional ccFlags As COMMONCONTROLS_CLASSES = ccAll_Classes) As Boolean
  Dim icc As tagINITCOMMONCONTROLSEX
  On Error GoTo OldCC
  With icc
    .dwSize = Len(icc)
    .dwICC = ccFlags
  End With
  InitCommonControls = InitCommonControlsEx(icc)
ExitNow:
  On Error GoTo 0
  Exit Function
OldCC:
  InitCommonControls9x
  Resume ExitNow
End Function

Public Function FileExistsA(FileName As String) As Boolean
  Dim wfd As WIN32_FIND_DATA
  Dim hFile As Long
  hFile = FindFirstFile(FileName, wfd)
  FileExistsA = Not hFile = INVALID_HANDLE_VALUE
  Call FindClose(hFile)
End Function

Public Function GetFileSizeA(FileName As String) As Double
  Dim wfd As WIN32_FIND_DATA
  Dim hFile As Long
  Dim fS As Currency
  hFile = FindFirstFile(FileName, wfd)
  If Not hFile = INVALID_HANDLE_VALUE Then
    fS = wfd.nFileSizeLow
    If wfd.nFileSizeLow < 0 Then
      fS = (fS + FILESIZE_FIX)
    End If
    If wfd.nFileSizeHigh > 0 Then
      fS = (fS + (wfd.nFileSizeHigh * FILESIZE_FIX))
    End If
  End If
  Call FindClose(hFile)
  GetFileSizeA = CDbl(fS)
End Function

Public Sub WindowOnTop(hWnd As Long, OnTop As Boolean)
  Dim wFlags As Long
  If OnTop Then
    wFlags = HWND_TOPMOST
  Else
    wFlags = HWND_NOTOPMOST
  End If
  SetWindowPos hWnd, wFlags, 0&, 0&, 0&, 0&, SWP_SETWINDOWPOS
End Sub

Public Function GetRandomNumber(Min As Double, Max As Double) As Double
  Dim r As Double
  r = ((Rnd * Max) + Min)
  If r < Min Then r = Min Else If r > Max Then r = Max
  GetRandomNumber = Int(r)
End Function

Public Function IsWindowsVistaOrHigher() As Boolean
  Dim vInfo As OSVERSIONINFO
  vInfo.dwOSVersionInfoSize = LenB(vInfo)
  Call RtlGetVersion(vInfo)
  IsWindowsVistaOrHigher = IIf(vInfo.dwMajorVersion >= 6, True, False)
End Function
