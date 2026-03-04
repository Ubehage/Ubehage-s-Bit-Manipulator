Attribute VB_Name = "modMsgBox"
Option Explicit

Private Const TD_WARNING_ICON As Long = -1 'exclamation point in a yellow 'yield' triangle (same image as IDI_EXCLAMATION)
Private Const TD_ERROR_ICON As Long = -2 'round red circle containg 'X' (same as IDI_HAND)
Private Const TD_INFORMATION_ICON As Long = -3 'round blue circle containing 'i' (same image as IDI_ASTERISK)
Private Const TD_SHIELD_ICON As Long = -4 'security shield
Private Const IDI_APPLICATION = 32512& 'miniature picture of an application window
Private Const IDI_QUESTION = 32514& 'round blue circle containing '?'
Private Const TD_SHIELD_GRADIENT_ICON As Long = -5 'same image as TD_SHIELD_ICON; main message text on gradient blue background
Private Const TD_SHIELD_WARNING_ICON As Long = -6 'exclamation point in yellow Shield shape; main message text on gradient orange background
Private Const TD_SHIELD_ERROR_ICON As Long = -7 'X contained within Shield shape; main message text on gradient red background
Private Const TD_SHIELD_OK_ICON As Long = -8 'Shield shape containing green checkmark; main message text on gradient green background
Private Const TD_SHIELD_GRAY_ICON As Long = -9 'same image as TD_SHIELD_ICON; main message text on medium gray background
Private Const TD_NO_ICON As Long = 0 'no icon; text on white background

Private Const TDCBF_OK_BUTTON As Long = &H1&      'return value 1 (IDOK)
Private Const TDCBF_YES_BUTTON As Long = &H2&     'return value 6 (IDYES)
Private Const TDCBF_NO_BUTTON As Long = &H4&      'return value 7 (IDNO)
Private Const TDCBF_CANCEL_BUTTON As Long = &H8&  'return value 2 (IDCANCEL)
Private Const TDCBF_RETRY_BUTTON As Long = &H10&  'return value 4 (IDRETRY)
Private Const TDCBF_CLOSE_BUTTON As Long = &H20&  'return value 8 (IDCLOSE)

Private Const IDOK As Long = 1
Private Const IDCANCEL = 2
Private Const IDRETRY = 4
Private Const IDYES As Long = 6
Private Const IDNO As Long = 7
Private Const IDCLOSE = 8

'TaskDialog return codes
Private Const S_OK As Long = &H0 'Success
Private Const E_OUTOFMEMORY As Long = &H8007000E 'Out of memory
Private Const E_INVALIDARG As Long = &H80070057 'Invalid arguments
Private Const E_FAIL As Long = &H80004005 'Unspecified failure

Public Enum MessageBoxStyle
  mbsWarning = TD_WARNING_ICON
  mbsError = TD_ERROR_ICON
  mbsInformation = TD_INFORMATION_ICON
  mbsShield = TD_SHIELD_ICON
  mbsShieldGradient = TD_SHIELD_GRADIENT_ICON
  mbsShieldWarning = TD_SHIELD_WARNING_ICON
  mbsShieldError = TD_SHIELD_ERROR_ICON
  mbsShieldOK = TD_SHIELD_OK_ICON
  mbsShieldGray = TD_SHIELD_GRAY_ICON
  mbsNoStyle = TD_NO_ICON
  mbsApplicationIcon = IDI_APPLICATION
  mbsQuestion = IDI_QUESTION
End Enum

Public Enum MessageBoxButtons
  mbbOK = TDCBF_OK_BUTTON
  mbbYes = TDCBF_YES_BUTTON
  mbbNo = TDCBF_NO_BUTTON
  mbbCancel = TDCBF_CANCEL_BUTTON
  mbbRetry = TDCBF_RETRY_BUTTON
  mbbClose = TDCBF_CLOSE_BUTTON
End Enum

Private Declare Function TaskDialog Lib "comctl32.dll" (ByVal hwndParent As Long, ByVal hInstance As Long, ByVal pszWindowTitle As Long, ByVal pszMainInstruction As Long, ByVal pszContent As Long, ByVal dwCommonButtons As Long, ByVal pszIcon As Long, pnButton As Long) As Long

Public Function ShowMessageBox(MsgTitle As String, MsgMainMessage As String, MsgDescription As String, MsgStyle As MessageBoxStyle, MsgButtons As MessageBoxButtons) As MessageBoxButtons
  If IsWindowsVistaOrHigher() = True Then
    Dim r As Long, s As Long
    s = TaskDialog(frmMain.hWnd, 0&, StrPtr(MsgTitle), StrPtr(MsgMainMessage), StrPtr(MsgDescription), MsgButtons, MakeIntResource(MsgStyle), r)
    If s = S_OK Then
      ShowMessageBox = GetMessageBoxButtonFromReturnID(r)
    End If
  Else
    ShowMessageBox = ShowClassicMsgBox(MsgTitle, MsgMainMessage, MsgDescription, MsgStyle, MsgButtons)
  End If
End Function

Private Function MakeIntResource(ByVal dVal As Long) As Long
  MakeIntResource = &HFFFF& And dVal
End Function

Private Function ShowClassicMsgBox(MsgTitle As String, MsgMainMessage As String, MsgDescription As String, MsgStyle As MessageBoxStyle, MsgButtons As MessageBoxButtons) As MessageBoxButtons
  Dim mButtons As VbMsgBoxStyle, mText As String, r As VbMsgBoxResult
  If (MsgButtons And mbbYes) And (MsgButtons And mbbNo) Then
    If (MsgButtons And mbbCancel) Then mButtons = vbYesNoCancel Else mButtons = vbYesNo
  ElseIf (MsgButtons And mbbRetry) And (MsgButtons And mbbCancel) Then
    mButtons = vbRetryCancel
  ElseIf (MsgButtons And mbbOK) Then
    If (MsgButtons And mbbCancel) Then mButtons = vbOKCancel Else mButtons = vbOKOnly
  Else
    mButtons = vbOKOnly
  End If
  If (MsgStyle And mbsError) Or (MsgStyle And mbsShieldError) Then
    mButtons = mButtons Or vbExclamation
  ElseIf (MsgStyle And mbsInformation) Then
    mButtons = mButtons Or vbInformation
  ElseIf (MsgStyle And mbsQuestion) Then
    mButtons = mButtons Or vbQuestion
  ElseIf (MsgStyle And mbsWarning) Or (MsgStyle And mbsShieldWarning) Then
    mButtons = mButtons Or vbCritical
  End If
  mText = MsgMainMessage
  If MsgDescription <> "" Then mText = mText & vbCrLf & vbCrLf & MsgDescription
  r = MsgBox(mText, mButtons, MsgTitle)
  Select Case r
    Case VbMsgBoxResult.vbOK
      ShowClassicMsgBox = mbbOK
    Case VbMsgBoxResult.vbYes
      ShowClassicMsgBox = mbbYes
    Case VbMsgBoxResult.vbNo
      ShowClassicMsgBox = mbbNo
    Case VbMsgBoxResult.vbCancel
      ShowClassicMsgBox = mbbCancel
    Case VbMsgBoxResult.vbRetry
      ShowClassicMsgBox = mbbRetry
    Case Else
      ShowClassicMsgBox = mbbClose
  End Select
End Function

Private Function GetMessageBoxButtonFromReturnID(ReturnID As Long) As MessageBoxButtons
  Select Case ReturnID
    Case IDOK
      GetMessageBoxButtonFromReturnID = mbbOK
    Case IDCANCEL
      GetMessageBoxButtonFromReturnID = mbbCancel
    Case IDRETRY
      GetMessageBoxButtonFromReturnID = mbbRetry
    Case IDYES
      GetMessageBoxButtonFromReturnID = mbbYes
    Case IDNO
      GetMessageBoxButtonFromReturnID = mbbNo
    Case IDCLOSE
      GetMessageBoxButtonFromReturnID = mbbClose
  End Select
End Function
