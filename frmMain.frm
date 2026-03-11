VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00202020&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Ube's Bit Manipulator"
   ClientHeight    =   6270
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8895
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   8895
   StartUpPosition =   3  'Windows Default
   Begin BitManipulator.Button cmdNewFile 
      Height          =   555
      Left            =   4665
      TabIndex        =   14
      Top             =   4725
      Width           =   4140
      _ExtentX        =   7303
      _ExtentY        =   979
      Caption         =   "Modify & Write to New File..."
      BackColor       =   5023791
      HoverColor      =   6339651
      PressedColor    =   4033061
      ForeColor       =   15463654
      DisabledBackColor=   6455130
      DisabledTextColor=   13355947
      ButtonStyle     =   2
      FontName        =   "Consolas"
      FontSize        =   11.25
      FontBold        =   -1  'True
   End
   Begin BitManipulator.Button cmdTarget 
      Height          =   555
      Left            =   255
      TabIndex        =   13
      Top             =   4710
      Width           =   4140
      _ExtentX        =   7303
      _ExtentY        =   979
      Caption         =   "Modify & Overwrite Source File"
      BackColor       =   4539862
      HoverColor      =   6513642
      PressedColor    =   3223992
      ForeColor       =   15395579
      DisabledBackColor=   6776730
      DisabledTextColor=   10592542
      ButtonStyle     =   1
      FontName        =   "Consolas"
      FontSize        =   11.25
      FontBold        =   -1  'True
   End
   Begin BitManipulator.SeparatorLine sepLineCopyright 
      Height          =   30
      Left            =   420
      TabIndex        =   9
      Top             =   5520
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   53
   End
   Begin BitManipulator.Frame frmOptions 
      Height          =   2475
      Left            =   255
      TabIndex        =   1
      Top             =   1950
      Width           =   8355
      _ExtentX        =   14737
      _ExtentY        =   4366
      Caption         =   "BitFlip Options"
      Begin BitManipulator.CheckBox chkOnTop 
         Height          =   300
         Left            =   4425
         TabIndex        =   18
         Top             =   540
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   529
         Caption         =   "Keep this window on top"
      End
      Begin BitManipulator.SeparatorLine sepLineOptions 
         Height          =   960
         Left            =   3750
         TabIndex        =   17
         Top             =   555
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   1693
         Orientation     =   1
      End
      Begin BitManipulator.Button cmdRandom 
         Height          =   300
         Left            =   315
         TabIndex        =   12
         Top             =   405
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   529
         Caption         =   "Autoselect a random byte"
         BackColor       =   13582186
         HoverColor      =   16735364
         PressedColor    =   10038350
         ForeColor       =   16768230
         DisabledBackColor=   7883352
         DisabledTextColor=   12492970
         ButtonStyle     =   4
         FontName        =   "Consolas"
         FontSize        =   11.25
         FontBold        =   -1  'True
      End
      Begin BitManipulator.CheckBox chkRemoveBit 
         Height          =   300
         Left            =   285
         TabIndex        =   8
         Top             =   1980
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   529
         Caption         =   "Remove bit instead of flipping"
      End
      Begin VB.TextBox txtBit 
         BackColor       =   &H002A2A2A&
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   360
         Left            =   1830
         TabIndex        =   6
         Top             =   1500
         Width           =   1425
      End
      Begin VB.TextBox txtBytePos 
         BackColor       =   &H002A2A2A&
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   360
         Left            =   1725
         TabIndex        =   4
         Top             =   930
         Width           =   1425
      End
      Begin BitManipulator.CheckBox chkBrowse 
         Height          =   300
         Left            =   4290
         TabIndex        =   15
         Top             =   1500
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   529
         Caption         =   "Let me select the target file"
      End
      Begin VB.Label lblBrowse 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Default new name is ""[filename] Fucked by Ube"""
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00787878&
         Height          =   195
         Left            =   4350
         TabIndex        =   16
         Top             =   1935
         Width           =   4140
      End
      Begin VB.Label lblBit 
         AutoSize        =   -1  'True
         BackColor       =   &H002A2A2A&
         Caption         =   "Bit to Flip (1-8):"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   255
         TabIndex        =   5
         Top             =   1440
         Width           =   1470
      End
      Begin VB.Label lblBytePos 
         AutoSize        =   -1  'True
         BackColor       =   &H002A2A2A&
         Caption         =   "Byte Position:"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   990
         Width           =   1290
      End
   End
   Begin BitManipulator.Frame frmFile 
      Height          =   1380
      Left            =   360
      TabIndex        =   0
      Top             =   375
      Width           =   7260
      _ExtentX        =   12806
      _ExtentY        =   2434
      Caption         =   "Source File"
      Begin BitManipulator.Button cmdBrowse 
         Height          =   360
         Left            =   3360
         TabIndex        =   11
         Top             =   585
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   635
         Caption         =   "..."
         FontName        =   "Consolas"
         FontSize        =   9
      End
      Begin VB.TextBox txtFileName 
         BackColor       =   &H00202020&
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   360
         Left            =   525
         TabIndex        =   2
         Top             =   540
         Width           =   2430
      End
      Begin VB.Label lblSize 
         AutoSize        =   -1  'True
         BackColor       =   &H002A2A2A&
         Caption         =   "File Size: 0 bytes"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   225
         Left            =   240
         TabIndex        =   7
         Top             =   1020
         Width           =   1260
      End
   End
   Begin VB.Label lblCopyright 
      AutoSize        =   -1  'True
      BackColor       =   &H00202020&
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © Ubehage 2026"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   600
      TabIndex        =   10
      Top             =   5730
      Width           =   2100
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const SETTINGS_SETTINGS  As String = "Settings"
Private Const SETTINGS_ONTOP  As String = "OnTop"

Private Const BUTTON_SPACING As Long = 150

Private Const SIZE_CAPTION  As String = "Selected file size: %s bytes"
Private Const FILENAME_MASK As String = "%f% [fucked by Ube].%e%"

Dim TargetFileSize As Currency
Dim BytePosition As Currency
Dim BitIndex As Integer

Dim ResizedByCode As Boolean

Friend Sub SetForm()
  Me.Caption = APP_NAME & " [" & CStr(App.Major) & "," & CStr(App.Minor) & "." & CStr(App.Revision) & "]"
  Me.Show
  UpdateFileSize
End Sub

Private Sub MoveObjects()
  On Error GoTo ResizeError
  frmFile.Move 30, 30, (Me.ScaleWidth - 60)
  txtFileName.Move 90, (Screen.TwipsPerPixelY * 27)
  cmdBrowse.Move (frmFile.Width - (90 + cmdBrowse.Width)), txtFileName.Top
  txtFileName.Width = (cmdBrowse.Left - (txtFileName.Left + 90))
  lblSize.Move txtFileName.Left, ((txtFileName.Top + txtFileName.Height) + 60)
  frmFile.Height = ((lblSize.Top + lblSize.Height) + 90)
  frmOptions.Move frmFile.Left, ((frmFile.Top + frmFile.Height) + 90), frmFile.Width
  cmdRandom.Move 90, (Screen.TwipsPerPixelY * 27)
  txtBytePos.Top = ((cmdRandom.Top + cmdRandom.Height) + 120)
  lblBytePos.Move cmdRandom.Left, (txtBytePos.Top + ((txtBytePos.Height - lblBytePos.Height) \ 2))
  If lblBytePos.Width > lblBit.Width Then
    txtBytePos.Left = ((lblBytePos.Left + lblBytePos.Width) + 90)
  Else
    txtBytePos.Left = ((lblBytePos.Left + lblBit.Width) + 90)
  End If
  txtBit.Move txtBytePos.Left, ((txtBytePos.Top + txtBytePos.Height) + 90)
  lblBit.Move lblBytePos.Left, (txtBit.Top + ((txtBit.Height - lblBit.Height) \ 2))
  chkRemoveBit.Move lblBit.Left, ((txtBit.Top + txtBit.Height) + 90)
  frmOptions.Height = ((chkRemoveBit.Top + chkRemoveBit.Height) + 90)
  sepLineOptions.Move ((chkRemoveBit.Left + chkRemoveBit.Width) + 90), cmdRandom.Top, sepLineOptions.Width, ((chkRemoveBit.Top + chkRemoveBit.Height) - cmdRandom.Top)
  chkBrowse.Left = ((sepLineOptions.Left + sepLineOptions.Width) + 150)
  lblBrowse.Move (chkBrowse.Left + 90), ((chkRemoveBit.Top + chkRemoveBit.Height) - lblBrowse.Height)
  chkBrowse.Top = (lblBrowse.Top - (chkBrowse.Height + 30))
  chkOnTop.Move chkBrowse.Left, (chkBrowse.Top - (chkOnTop.Height + BUTTON_SPACING))
  Dim w As Long
  w = ((cmdTarget.Width + cmdNewFile.Width) + BUTTON_SPACING)
  cmdTarget.Move (Me.ScaleWidth - w) \ 2, ((frmOptions.Top + frmOptions.Height) + 150)
  cmdNewFile.Move ((cmdTarget.Left + cmdTarget.Width) + BUTTON_SPACING), cmdTarget.Top
  sepLineCopyright.Width = (Me.ScaleWidth * 0.6)
  sepLineCopyright.Move (Me.ScaleWidth - sepLineCopyright.Width) \ 2, ((cmdNewFile.Top + cmdNewFile.Height) + 60)
  lblCopyright.Move (Me.ScaleWidth - lblCopyright.Width) \ 2, ((sepLineCopyright.Top + sepLineCopyright.Height) + 30)
  ResizedByCode = True
  Me.Height = ((Me.Height - Me.ScaleHeight) + ((lblCopyright.Top + lblCopyright.Height) + 90))
  ResizedByCode = False
  On Error GoTo 0
  Exit Sub
ResizeError:
  'fuck errors. Let the user resize whatever they want
  Resume Next
End Sub

Private Sub UpdateFileSize()
  TargetFileSize = GetFileSizeA(txtFileName.Text)
  lblSize.Caption = Replace$(SIZE_CAPTION, "%s", GetTargetFileSizeString)
  cmdRandom.Enabled = (TargetFileSize > 0)
  lblBytePos.Enabled = cmdRandom.Enabled
  lblBit.Enabled = lblBytePos.Enabled
  txtBytePos.Enabled = lblBytePos.Enabled
  txtBit.Enabled = lblBit.Enabled
  chkRemoveBit.Enabled = txtBit.Enabled
  CheckIfReady
End Sub

Private Function GetTargetFileSizeString() As String
  GetTargetFileSizeString = IIf(TargetFileSize = 0, "0", Format$(TargetFileSize, "###,###,###,###,###"))
End Function

Private Sub CheckIfReady()
  Dim r As Boolean
  r = True
  If cmdRandom.Enabled = False Then
    r = False
  Else
    If (BytePosition <= 0 Or BytePosition > TargetFileSize) Then
      r = False
    Else
      If (BitIndex <= 0 Or BitIndex > 8) Then
        r = False
      End If
    End If
  End If
  cmdTarget.Enabled = r
  cmdNewFile.Enabled = r
End Sub

Private Sub SaveWindowSettings()
  SaveSetting APP_NAME, SETTINGS_SETTINGS, SETTINGS_ONTOP, CStr(chkOnTop.Value = vbChecked)
End Sub

Private Sub LoadWindowSettings()
  Dim v As String
  On Error GoTo SettingsError
  v = GetSetting(APP_NAME, SETTINGS_SETTINGS, SETTINGS_ONTOP, "False")
  chkOnTop.Value = IIf(CBool(v), vbChecked, vbUnchecked)
  On Error GoTo 0
  Exit Sub
SettingsError:
  Resume Next
End Sub

Private Function GetNewFileNameFromMask() As String
  Dim fN As String, fE As String, i As Long
  fN = txtFileName.Text
  i = InStrRev(fN, ".")
  If i > 0 Then
    fE = Right$(fN, (Len(fN) - i))
    fN = Left$(fN, (i - 1))
  End If
  GetNewFileNameFromMask = Replace$(Replace$(FILENAME_MASK, "%f%", fN), "%e%", fE)
End Function

Private Sub chkOnTop_Click()
  WindowOnTop Me.hWnd, (chkOnTop.Value = vbChecked)
End Sub

Private Sub cmdBrowse_Click()
  Dim fName As String
  fName = BrowseForFileA("Select file to modify...", Me.hWnd)
  If (fName <> "" Or FileExistsA(fName) = True) Then txtFileName.Text = fName
End Sub

Private Sub cmdNewFile_Click()
  Dim nFile As String
  If chkBrowse.Value = vbChecked Then
    nFile = BrowseForFileA("Select target file...", Me.hWnd)
  Else
    nFile = GetNewFileNameFromMask()
  End If
  If nFile <> "" Then
    If ManipulateBitToNewFile(txtFileName.Text, nFile, CCur(txtBytePos.Text), CInt(txtBit.Text), IIf(chkRemoveBit.Value = vbChecked, Bit_Manipulation_Method.bmRemove, Bit_Manipulation_Method.bmFlip)) = True Then
      Call ShowMessageBox("Done!", APP_NAME & " successfully saved the new file.", "Data was saved to """ & nFile & """." & vbCrLf & "Byte number " & txtBytePos.Text & " has been manipulated." & vbCrLf & _
                          "The bit at index " & txtBit.Text & " was " & IIf(chkRemoveBit.Value = vbChecked, "removed", "flipped") & ".", mbsShieldOK, mbbOK)
    Else
      Call ShowMessageBox("Failed!", APP_NAME & " failed to save the new file.", "", mbsShieldError, mbbOK)
    End If
  End If
End Sub

Private Sub cmdRandom_Click()
  txtBytePos.Text = CStr(GetRandomNumber(1, TargetFileSize))
  txtBit.Text = CStr(GetRandomNumber(1, 8))
End Sub

Private Sub cmdTarget_Click()
  If ManipulateBitInFile(txtFileName.Text, CCur(txtBytePos.Text), CInt(txtBit.Text), IIf(chkRemoveBit.Value = vbChecked, Bit_Manipulation_Method.bmRemove, Bit_Manipulation_Method.bmFlip)) = True Then
    Call ShowMessageBox("Done!", APP_NAME & " successfully changed the file.", "Byte number " & txtBytePos.Text & " has been manipulated." & vbCrLf & _
                        "The bit at index " & txtBit.Text & " was " & IIf(chkRemoveBit.Value = vbChecked, "removed", "flipped") & ".", mbsShieldOK, mbbOK)
  Else
    Call ShowMessageBox("Failed!", APP_NAME & " failed to complete the task.", "", mbsShieldError, mbbOK)
  End If
End Sub

Private Sub Form_Load()
  LoadWindowSettings
End Sub

Private Sub Form_Resize()
  If ResizedByCode = True Then Exit Sub
  MoveObjects
End Sub

Private Sub Form_Unload(Cancel As Integer)
  SaveWindowSettings
End Sub

Private Sub txtBit_Change()
  BitIndex = CInt(Val(txtBit.Text))
  CheckIfReady
End Sub

Private Sub txtBit_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
    Case vbKeyBack, vbKeyDelete, vbKeyHome, vbKeyEnd
      Exit Sub
    Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight
      Exit Sub
  End Select
  If (KeyAscii < 48 Or KeyAscii > 56) Then KeyAscii = 0
End Sub

Private Sub txtBytePos_Change()
  BytePosition = CCur(Val(txtBytePos.Text))
  CheckIfReady
End Sub

Private Sub txtBytePos_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
    Case vbKeyBack, vbKeyDelete, vbKeyHome, vbKeyEnd
      Exit Sub
    Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight
      Exit Sub
  End Select
  If (KeyAscii < 48 Or KeyAscii > 57) Then KeyAscii = 0
End Sub

Private Sub txtFileName_Change()
  UpdateFileSize
End Sub
