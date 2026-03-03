Attribute VB_Name = "modControls"
Option Explicit

Global Const FONT_MAIN As String = "Segoe UI"
Global Const FONT_SECONDARY As String = "Consolas"
Global Const FONTSIZE_MAIN As Integer = 11
Global Const FONTSIZE_SECONDARY As Integer = 9

Global Const COLOR_BACKGROUND As Long = 2105376
Global Const COLOR_CONTROLS As Long = 2763306
Global Const COLOR_BUTTON_HOVER As Long = 3684408
Global Const COLOR_BUTTON_PRESSED As Long = 3289650
Global Const COLOR_BACKGROUND_DISABLED As Long = 5263440

Global Const COLOR_TEXT As Long = 14737632
Global Const COLOR_TEXT_HOVER As Long = 15790320
Global Const COLOR_TEXT_DISABLED As Long = 7895160
Global Const COLOR_TEXT_ONGREEN As Long = 15463654
Global Const COLOR_TEXT_ONRED As Long = 15395579
Global Const COLOR_TEXT_DISABLED_ONGREEN As Long = 13355947
Global Const COLOR_TEXT_DISABLED_ONRED As Long = 10592542

Global Const COLOR_GREEN As Long = 5023791
Global Const COLOR_GREEN_HOVER As Long = 6339651
Global Const COLOR_GREEN_PRESSED As Long = 4033061
Global Const COLOR_GREEN_DISABLED As Long = 6455130
Global Const COLOR_YELLOW As Long = 4965861
Global Const COLOR_YELLOW_HOVER As Long = 6673645
Global Const COLOR_YELLOW_PRESSED As Long = 3710156
Global Const COLOR_YELLOW_DISABLED As Long = 7112080
Global Const COLOR_RED As Long = 4539862
Global Const COLOR_RED_HOVER As Long = 6513642
Global Const COLOR_RED_PRESSED As Long = 3223992
Global Const COLOR_RED_DISABLED As Long = 6776730
Global Const COLOR_OUTLINE As Long = 3815994
Global Const COLOR_OUTLINE_LIGHT As Long = 7368816

Private Const ICC_LISTVIEW_CLASSES  As Long = &H1
Private Const ICC_TREEVIEW_CLASSES  As Long = &H2
Private Const ICC_BAR_CLASSES  As Long = &H4
Private Const ICC_TAB_CLASSES  As Long = &H8
Private Const ICC_UPDOWN_CLASS  As Long = &H10
Private Const ICC_PROGRESS_CLASS  As Long = &H20
Private Const ICC_HOTKEY_CLASS  As Long = &H40
Private Const ICC_ANIMATE_CLASS  As Long = &H80
Private Const ICC_WIN95_CLASSES  As Long = &HFF
Private Const ICC_DATE_CLASSES  As Long = &H100
Private Const ICC_USEREX_CLASSES  As Long = &H200
Private Const ICC_COOL_CLASSES  As Long = &H400
Private Const ICC_INTERNET_CLASSES  As Long = &H800
Private Const ICC_PAGESCROLLER_CLASS  As Long = 1000
Private Const ICC_NATIVEFNTCTL_CLASS  As Long = 2000
Private Const ICC_STANDARD_CLASSES  As Long = 4000
Private Const ICC_LINK_CLASS  As Long = 8000

Public Enum COMMONCONTROLS_CLASSES
  ccListView_Classes = ICC_LISTVIEW_CLASSES
  ccTreeView_Classes = ICC_TREEVIEW_CLASSES
  ccToolBar_Classes = ICC_BAR_CLASSES
  ccTab_Classes = ICC_TAB_CLASSES
  ccUpDown_Classes = ICC_UPDOWN_CLASS
  ccProgress_Class = ICC_PROGRESS_CLASS
  ccHotkey_Class = ICC_HOTKEY_CLASS
  ccAnimate_Class = ICC_ANIMATE_CLASS
  ccWin95_Classes = ICC_WIN95_CLASSES
  ccCalendar_Classes = ICC_DATE_CLASSES
  ccComboEx_Classes = ICC_USEREX_CLASSES
  ccCoolBar_Classes = ICC_COOL_CLASSES
  ccInternet_Classes = ICC_INTERNET_CLASSES
  ccPageScroller_Class = ICC_PAGESCROLLER_CLASS
  ccNativeFont_Class = ICC_NATIVEFNTCTL_CLASS
  ccStandard_Classes = ICC_STANDARD_CLASSES
  ccLink_Class = ICC_LINK_CLASS
  ccAll_Classes = ccListView_Classes Or ccTreeView_Classes Or ccToolBar_Classes Or ccTab_Classes Or ccUpDown_Classes Or ccProgress_Class Or ccHotkey_Class Or ccAnimate_Class Or ccWin95_Classes Or ccCalendar_Classes Or ccComboEx_Classes Or ccCoolBar_Classes Or ccInternet_Classes Or ccPageScroller_Class Or ccNativeFont_Class Or ccStandard_Classes Or ccLink_Class
End Enum

Public Type POINTAPI
  X As Long
  Y As Long
End Type

Public Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Public Declare Sub GetWindowRect Lib "user32" (ByVal hWnd As Long, ByRef WindowRect As RECT)
Public Declare Function WindowFromPoint Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, ByRef lpPoint As POINTAPI) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long

Public Function IsPointInRect(pRect As RECT, pPoint As POINTAPI) As Boolean
  With pRect
    If (pPoint.X >= .Left And pPoint.X <= .Right) Then
      If (pPoint.Y >= .Top And pPoint.Y <= .Bottom) Then IsPointInRect = True
    End If
  End With
End Function

Public Function IsCursorOnWindow(hWnd As Long, Optional MouseIsDown As Boolean = False) As Boolean
  Dim mPos As POINTAPI, wRect As RECT, hTop As Long
  Call GetCursorPos(mPos)
  Call GetWindowRect(hWnd, wRect)
  If IsPointInRect(wRect, mPos) = False Then If MouseIsDown = False Then Exit Function
  hTop = WindowFromPoint(mPos.X, mPos.Y)
  If hTop <> hWnd Then If MouseIsDown = False Then Exit Function
  IsCursorOnWindow = True
End Function

