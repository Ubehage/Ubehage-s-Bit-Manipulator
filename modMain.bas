Attribute VB_Name = "modMain"
Option Explicit

Global Const APP_NAME = "Ubehage's Bit Manipulator"

Global ChangedByCode As Boolean

Sub Main()
  Call InitCommonControls
  Randomize Timer
  LoadForm
End Sub

Private Sub LoadForm()
  Load frmMain
  frmMain.SetForm
End Sub

Public Function ShowErrorMessage(ErrorOperation As String, ErrorMessage As String, ErrorCode As Long, Optional Path As String, Optional MsgButtons As MessageBoxButtons = mbbRetry Or mbbCancel) As MessageBoxButtons
  Dim mMain As String, mDescription As String
  mMain = ErrorOperation & " failed."
  mDescription = "Error Description: " & ErrorMessage & vbCrLf & _
                  "Error Code: " & CStr(ErrorCode)
  If Path <> "" Then mDescription = mDescription & vbCrLf & "Path: " & Path
  ShowErrorMessage = ShowMessageBox(APP_NAME & " encountered a critical error!", mMain, mDescription, mbsShieldError, MsgButtons)
End Function
