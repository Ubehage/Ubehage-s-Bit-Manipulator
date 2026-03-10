VERSION 5.00
Begin VB.UserControl SeparatorLine 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00202020&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "SeparatorLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const PROP_NAME_ORIENTATION = "Orientation"

Public Enum LINE_ORIENTATION
  loHorizontal = &H0
  loVertical = &H1
End Enum

Dim m_Orientation As LINE_ORIENTATION

Dim m_FixedSize As Long

Public Property Get Orientation() As LINE_ORIENTATION
  Orientation = m_Orientation
End Property
Public Property Let Orientation(New_Orientation As LINE_ORIENTATION)
  If m_Orientation = New_Orientation Then Exit Property
  m_Orientation = New_Orientation
  UserControl.PropertyChanged PROP_NAME_ORIENTATION
  UserControl_Resize
End Property

Public Sub Refresh()
  UserControl.Cls
  DrawLine
End Sub

Private Sub DrawLine()
  Dim LineRect As RECT
  With LineRect
    .Left = 0
    .Top = 0
    .Right = (UserControl.ScaleWidth - Screen.TwipsPerPixelX)
    .Bottom = (UserControl.ScaleHeight - Screen.TwipsPerPixelY)
    If m_Orientation = loHorizontal Then
      UserControl.Line (.Left, .Top)-(.Right, .Top), COLOR_OUTLINE_LIGHT
      .Top = (.Top + Screen.TwipsPerPixelY)
      UserControl.Line (.Left, .Top)-(.Right, .Top), COLOR_OUTLINE
    ElseIf m_Orientation = loVertical Then
      UserControl.Line (.Left, .Top)-(.Left, .Bottom), COLOR_OUTLINE_LIGHT
      .Left = (.Left + Screen.TwipsPerPixelX)
      UserControl.Line (.Left, .Top)-(.Left, .Bottom), COLOR_OUTLINE
    End If
  End With
End Sub

Private Sub UserControl_InitProperties()
  m_Orientation = loHorizontal
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  m_Orientation = PropBag.ReadProperty(PROP_NAME_ORIENTATION, LINE_ORIENTATION.loHorizontal)
End Sub

Private Sub UserControl_Resize()
  Dim s As Long
  If m_Orientation = loHorizontal Then
    s = ((UserControl.Height - UserControl.ScaleHeight) + (Screen.TwipsPerPixelY * 2))
    If UserControl.Height <> s Then UserControl.Height = s: Exit Sub
  Else
    s = ((UserControl.Width - UserControl.ScaleWidth) + (Screen.TwipsPerPixelX * 2))
    If UserControl.Width <> s Then UserControl.Width = s: Exit Sub
  End If
  Refresh
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  PropBag.WriteProperty PROP_NAME_ORIENTATION, m_Orientation, LINE_ORIENTATION.loHorizontal
End Sub
