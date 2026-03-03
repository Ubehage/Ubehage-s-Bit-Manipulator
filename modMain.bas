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
