VERSION 5.00
Begin VB.UserControl Frame3D 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   FillColor       =   &H00FFC19F&
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "Frame3D.ctx":0000
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   195
      TabIndex        =   0
      Top             =   0
      Width           =   45
   End
   Begin VB.Label lblCaptionShadow 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000014&
      Height          =   195
      Left            =   210
      TabIndex        =   1
      Top             =   15
      Width           =   45
   End
   Begin VB.Label lblCaptionBlank 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   45
   End
End
Attribute VB_Name = "Frame3D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//************************************************************************************************
'// Author: Morgan Haueisen (morganh@hartcom.net)
'// Copyright (c) 2004-2006
'// Version 1.0.2
'//
'// Legal:
'//
'//      Redistribution of this code, whole or in part, as source code or in binary form, alone or
'//      as part of a larger distribution or product, is forbidden for any commercial or for-profit
'//      use without the author's explicit written permission.
'//
'//      Redistribution of this code, as source code or in binary form, with or without
'//      modification, is permitted provided that the following conditions are met:
'//
'//      Redistributions of source code must include this list of conditions, and the following
'//      acknowledgment:
'//
'//      This code was developed by Morgan Haueisen.  <morganh@hartcom.net>
'//      Source code, written in Visual Basic, is freely available for non-commercial,
'//      non-profit use.
'//
'//      Redistributions in binary form, as part of a larger project, must include the above
'//      acknowledgment in the end-user documentation.  Alternatively, the above acknowledgment
'//      may appear in the software itself, if and wherever such third-party acknowledgments
'//      normally appear.
'//************************************************************************************************

Option Explicit

Private Type POINTAPI
   x As Long
   Y As Long
End Type

Private Type Rect
   left As Long
   top As Long
   right As Long
   bottom As Long
End Type

'// Collapsible Frame
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long) As Long
Private Declare Function SetRect Lib "user32" (ByRef lpRect As Rect, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObj As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function PtInRegion Lib "gdi32" (ByVal hRgn As Long, ByVal x As Long, ByVal Y As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal HWND As Long, ByRef lpPoint As POINTAPI) As Long

'// Used to draw the object's rounded border
Private Declare Function RoundRect Lib "gdi32" ( _
      ByVal hDC As Long, _
      ByVal left As Long, _
      ByVal top As Long, _
      ByVal right As Long, _
      ByVal bottom As Long, _
      ByVal EllipseWidth As Long, _
      ByVal EllipseHeight As Long) As Long

'// Used to make the rounded corners transparent
Private Declare Function SetWindowRgn Lib "user32" (ByVal HWND As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" ( _
      ByVal RectX1 As Long, _
      ByVal RectY1 As Long, _
      ByVal RectX2 As Long, _
      ByVal RectY2 As Long, _
      ByVal EllipseWidth As Long, _
      ByVal EllipseHeight As Long) As Long

'// The GetSysColor function retrieves the current color of the specified display element
'// Used to add gradient fill
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Public Enum enuBorderTypes
   [None Border] = 0
   [Frame Inserted] = 1
   [Frame Raised] = 2
   [Panel Flat Shadow] = 3
   [Panel Flat Highlight] = 4
   [Panel Raised] = 5
   [Panel Inserted] = 6
   [rFrame Inserted] = 7
   [rFrame Raised] = 8
   [rPanel Flat Shadow] = 9
   [rPanel Flat Highlight] = 10
   [rPanel Raised] = 11
   [rPanel Inserted] = 12
   [rNone Border] = 13
End Enum

Public Enum enuBevelInner
   [None Bevel] = 0
   [Inserted Bevel] = 1
   [Raised Bevel] = 2
   [Flat Shadow] = 3
   [Flat Highlight] = 4
End Enum

Public Enum enuCaption3D
   [Flat Caption] = 0
   [Inserted Caption] = 1
   [Raised Caption] = 2
End Enum

Public Enum enuCaptionLocation
   [Inside Frame] = 0
   [In Frame] = 1
End Enum

Public Enum enuChevronType
   [Round_Chev] = 0
   [Square_Chev] = 1
End Enum

Public Enum enuFloodType
   [Left To Right] = 0
   [Bottom To Top] = 1
End Enum

Public Enum enuCaptionAlignment
   [Top Left] = 0
   [Top Center] = 1
   [Top Right] = 2
   [Middle Left] = 3
   [Middle Center] = 4
   [Middle Right] = 5
   [Bottom Left] = 6
   [Bottom Center] = 7
   [Bottom Right] = 8
End Enum

Public Enum enuFillGradient
   [None Gradient] = 0
   [Fill Horizontal] = 1
   [Fill Vertical] = 2
End Enum

Public Event Click()
Public Event DblClick()
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event ChevronClick(Button As Integer, Shift As Integer, x As Single, Y As Single, Collapse As Boolean, Height As Long)

Private Const C_MINHEADERSIZE    As Long = 21

Private mudtBorderType           As enuBorderTypes
Private mudtBevelInner           As enuBevelInner
Private mudtCaption3D            As enuCaption3D
Private mudtCaptionAlignment     As enuCaptionAlignment
Private mudtCaptionLocation      As enuCaptionLocation
Private mudtFloodType            As enuFloodType
Private mudtFillGradient         As enuFillGradient

Private mlngBevelWidth           As Long
Private mlng3DHighlight          As OLE_COLOR
Private mlng3DShadow             As OLE_COLOR
Private mlngBackColor            As OLE_COLOR
Private mblnEnabled              As Boolean
Private msngTop                  As Single
Private msngLeft                 As Single
Private mlngFloodValue           As Long
Private mblnFloodShowPct         As Boolean
Private mlngFloodColor           As OLE_COLOR
Private msngInsideBorder         As Single
Private mlngCornerRadius         As Long
Private mlngCornerDia            As Long
Private mlngInsideHeight         As Long
Private mlngInsideWidth          As Long
Private mlngInsideLeft           As Long
Private mlngInsideTop            As Long
Private mlngChevronColor         As OLE_COLOR
Private mblnCollapsible          As Boolean
Private mlngFullHeight           As Long
Private mblnCollapse             As Boolean
Private mlngRegion               As Long
Private mbytChevronType          As enuChevronType

Public Property Get BackColor() As OLE_COLOR

   BackColor = mlngBackColor

End Property

Public Property Let BackColor(ByVal vNewValue As OLE_COLOR)

   mlngBackColor = vNewValue
   PropertyChanged "BackColor"
   Call UserControl_Resize

End Property

Public Property Let BevelInner(ByVal vNewValue As enuBevelInner)

   mudtBevelInner = vNewValue
   PropertyChanged "BevelInner"
   Call UserControl_Resize

End Property

Public Property Get BevelInner() As enuBevelInner

   BevelInner = mudtBevelInner

End Property

Public Property Get BevelWidth() As Long

   BevelWidth = mlngBevelWidth

End Property

Public Property Let BevelWidth(ByVal vNewValue As Long)

   mlngBevelWidth = vNewValue
   If mlngBevelWidth < 3 Then mlngBevelWidth = 3
   PropertyChanged "BevelWidth"
   Call UserControl_Resize

End Property

Public Property Let Border3DHighlight(ByVal vNewValue As OLE_COLOR)

   mlng3DHighlight = vNewValue
   lblCaptionShadow.ForeColor = mlng3DHighlight
   PropertyChanged "Border3DHighlight"
   Call UserControl_Resize

End Property

Public Property Get Border3DHighlight() As OLE_COLOR

   Border3DHighlight = mlng3DHighlight

End Property

Public Property Let Border3DShadow(ByVal vNewValue As OLE_COLOR)

   mlng3DShadow = vNewValue
   PropertyChanged "Border3DShadow"
   Call UserControl_Resize

End Property

Public Property Get Border3DShadow() As OLE_COLOR

   Border3DShadow = mlng3DShadow

End Property

Public Property Get BorderType() As enuBorderTypes

   BorderType = mudtBorderType

End Property

Public Property Let BorderType(ByVal vNewValue As enuBorderTypes)

   mudtBorderType = vNewValue
   PropertyChanged "BorderType"
   Call UserControl_Resize

End Property

Public Property Let Caption(ByVal vstrNewValue As String)

   lblCaption.Caption = vstrNewValue
   lblCaptionShadow.Caption = lblCaption.Caption
   lblCaption.Refresh
   lblCaptionShadow.Refresh
   PropertyChanged "Caption"
   Call UserControl_Resize

End Property

Public Property Get Caption() As String

   Caption = lblCaption.Caption

End Property

Public Property Let Caption3D(ByVal NewValue As enuCaption3D)

   mudtCaption3D = NewValue
   PropertyChanged "Caption3D"
   Call UserControl_Resize

End Property

Public Property Get Caption3D() As enuCaption3D

   Caption3D = mudtCaption3D

End Property

Public Property Get CaptionAlignment() As enuCaptionAlignment

   CaptionAlignment = mudtCaptionAlignment

End Property

Public Property Let CaptionAlignment(ByVal vNewValue As enuCaptionAlignment)

   mudtCaptionAlignment = vNewValue
   PropertyChanged "CaptionAlignment"
   Call UserControl_Resize

End Property

Public Property Let CaptionLocation(ByVal vNewValue As enuCaptionLocation)

   mudtCaptionLocation = vNewValue
   PropertyChanged "CaptionLocation"
   Call UserControl_Resize

End Property

Public Property Get CaptionLocation() As enuCaptionLocation

   CaptionLocation = mudtCaptionLocation

End Property

Public Property Let CaptionMAlignment(ByVal vNewValue As AlignmentConstants)

   lblCaption.Alignment = vNewValue
   lblCaptionShadow.Alignment = vNewValue
   PropertyChanged "CaptionMAlignment"
   Call UserControl_Resize

End Property

Public Property Get CaptionMAlignment() As AlignmentConstants

   CaptionMAlignment = lblCaption.Alignment

End Property

Public Property Let ChevronType(ByVal vNewValue As enuChevronType)

   mbytChevronType = vNewValue
   PropertyChanged "ChevronType"
   Call DrawControl

End Property

Public Property Get ChevronType() As enuChevronType

   ChevronType = mbytChevronType

End Property

Public Property Let Collapse(ByVal vNewValue As Boolean)

   mblnCollapse = vNewValue
   PropertyChanged "Collapse"

   On Error GoTo Err_Proc

   If mblnCollapse Then
      Call SetMinHeight
   Else
      UserControl.Height = mlngFullHeight
   End If

   Exit Property

Err_Proc:
   ErrHandler False, Err.Number, Err.Description, "Frame3D", "Collapse"
   Err.Clear
   Resume Next

End Property

Public Property Get Collapse() As Boolean

   Collapse = mblnCollapse

End Property

Public Property Get Collapsible() As Boolean

   Collapsible = mblnCollapsible

End Property

Public Property Let Collapsible(ByVal vNewValue As Boolean)

   mblnCollapsible = vNewValue
   PropertyChanged "Collapsible"
   Call DrawControl

End Property

Public Property Get CollapsibleColor() As OLE_COLOR

   CollapsibleColor = mlngChevronColor

End Property

Public Property Let CollapsibleColor(ByVal vNewValue As OLE_COLOR)

   mlngChevronColor = vNewValue
   PropertyChanged "ChevronColor"
   Call DrawControl

End Property

Public Property Get CornerRadius() As Long

   CornerRadius = mlngCornerRadius

End Property

Public Property Let CornerRadius(ByVal NewValue As Long)

   mlngCornerRadius = NewValue
   mlngCornerDia = mlngCornerRadius * 2
   PropertyChanged "CornerDiameter"
   Call UserControl_Resize

End Property

Private Sub DrawBevelInner()

  Dim sngBevelWidth As Single
  Dim lngCorner     As Long
  Dim intDrawWidth  As Integer

   On Error GoTo Err_Proc

   intDrawWidth = UserControl.DrawWidth
   If mudtBevelInner = 0 Then GoTo Exit_Proc

   UserControl.DrawWidth = 1

   sngBevelWidth = mlngBevelWidth + msngInsideBorder + (intDrawWidth / 2)
   lblCaptionBlank.Visible = False

   Select Case mudtBorderType
   Case 7 To 13  '// Rounded Corners
      lngCorner = mlngCornerRadius

   Case Else
      lngCorner = 0&
   End Select

   With UserControl

      Select Case mudtBevelInner
      Case 2 '// Raised

         '// Top
         UserControl.Line (sngBevelWidth + lngCorner, sngBevelWidth)-(.ScaleWidth - lngCorner - 1 - sngBevelWidth, sngBevelWidth), mlng3DHighlight
         '// Left
         UserControl.Line (sngBevelWidth, lngCorner + sngBevelWidth)-(sngBevelWidth, .ScaleHeight - sngBevelWidth - lngCorner - 1), mlng3DHighlight
         '// Right
         UserControl.Line (.ScaleWidth - sngBevelWidth - 1, lngCorner + sngBevelWidth)-(.ScaleWidth - sngBevelWidth - 1, .ScaleHeight - lngCorner - _
            sngBevelWidth - 1), mlng3DShadow
         '// Bottom
         UserControl.Line (lngCorner + sngBevelWidth, .ScaleHeight - sngBevelWidth - 1)-(.ScaleWidth - sngBevelWidth - lngCorner - 1, .ScaleHeight - _
            sngBevelWidth - 1), mlng3DShadow

         If lngCorner Then
            '// Top Left
            UserControl.Circle (lngCorner + sngBevelWidth, lngCorner + sngBevelWidth), lngCorner, mlng3DHighlight, 1.57, 3.14
            '// Bottom Left
            UserControl.Circle (lngCorner + sngBevelWidth, .ScaleHeight - sngBevelWidth - lngCorner - 1), lngCorner, mlng3DHighlight, 3.14, 4.71
            '// Top Right
            UserControl.Circle (.ScaleWidth - lngCorner - sngBevelWidth - 1, sngBevelWidth + lngCorner), lngCorner, mlng3DShadow, 6.28, 1.57
            '// Bottom Right
            UserControl.Circle (.ScaleWidth - lngCorner - sngBevelWidth - 1, .ScaleHeight - sngBevelWidth - lngCorner - 1), lngCorner, mlng3DShadow, _
               4.71
         End If

      Case 1 '// Inserted

         '// Top
         UserControl.Line (sngBevelWidth + lngCorner, sngBevelWidth)-(.ScaleWidth - lngCorner - 1 - sngBevelWidth, sngBevelWidth), mlng3DShadow
         '// Left
         UserControl.Line (sngBevelWidth, lngCorner + sngBevelWidth)-(sngBevelWidth, .ScaleHeight - sngBevelWidth - lngCorner - 1), mlng3DShadow
         '// Right
         UserControl.Line (.ScaleWidth - sngBevelWidth - 1, lngCorner + sngBevelWidth)-(.ScaleWidth - sngBevelWidth - 1, .ScaleHeight - lngCorner - _
            sngBevelWidth - 1), mlng3DHighlight
         '// Bottom
         UserControl.Line (lngCorner + sngBevelWidth, .ScaleHeight - sngBevelWidth - 1)-(.ScaleWidth - sngBevelWidth - lngCorner - 1, .ScaleHeight - _
            sngBevelWidth - 1), mlng3DHighlight

         If lngCorner Then
            '// Top Left
            UserControl.Circle (lngCorner + sngBevelWidth, lngCorner + sngBevelWidth), lngCorner, mlng3DShadow, 1.57, 3.14
            '// Bottom Left
            UserControl.Circle (lngCorner + sngBevelWidth, .ScaleHeight - sngBevelWidth - lngCorner - 1), lngCorner, mlng3DShadow, 3.14, 4.71
            '// Top Right
            UserControl.Circle (.ScaleWidth - lngCorner - sngBevelWidth - 1, sngBevelWidth + lngCorner), lngCorner, mlng3DHighlight, 6.28, 1.57
            '// Bottom Right
            UserControl.Circle (.ScaleWidth - lngCorner - sngBevelWidth - 1, .ScaleHeight - sngBevelWidth - lngCorner - 1), lngCorner, _
               mlng3DHighlight, 4.71
         End If

      Case 3 '// Flat Shadow

         With UserControl
            .ForeColor = mlng3DShadow
            Call RoundRect(.hDC, sngBevelWidth, sngBevelWidth, .ScaleWidth - sngBevelWidth, .ScaleHeight - sngBevelWidth, lngCorner, lngCorner)
         End With

      Case 4 '// Flat Highlight

         With UserControl
            .ForeColor = mlng3DHighlight
            Call RoundRect(.hDC, sngBevelWidth, sngBevelWidth, .ScaleWidth - sngBevelWidth, .ScaleHeight - sngBevelWidth, lngCorner, lngCorner)
         End With

      End Select
   End With

Exit_Proc:

   UserControl.DrawWidth = intDrawWidth
   intDrawWidth = intDrawWidth / 2
   '// Get inside workspace size in twips
   mlngInsideHeight = (UserControl.ScaleHeight - sngBevelWidth - sngBevelWidth - 2 - intDrawWidth) * Screen.TwipsPerPixelY
   mlngInsideWidth = (UserControl.ScaleWidth - sngBevelWidth - sngBevelWidth - 2 - intDrawWidth) * Screen.TwipsPerPixelX
   mlngInsideLeft = (sngBevelWidth + 1 - intDrawWidth) * Screen.TwipsPerPixelY
   mlngInsideTop = (sngBevelWidth + 1 - intDrawWidth) * Screen.TwipsPerPixelX

   If mlngFloodValue > 0 Then Call DrawFlood

   Exit Sub

Err_Proc:
   ErrHandler False, Err.Number, Err.Description, "Frame3D", "DrawBevelInner"
   Err.Clear
   Resume Next

End Sub

Private Sub DrawCaptionAlignment()

  Dim sngWidth   As Single
  Dim sngHeight  As Single
  Dim sngOffset  As Long
  Dim sngDrawWidth As Single

   On Error GoTo Err_Proc

   If LenB(lblCaption.Caption) Then

      sngDrawWidth = (UserControl.DrawWidth / 2) + 1

      Select Case mudtBorderType
      Case 7 To 13 '// Rounded Corners
         sngOffset = mlngCornerRadius
         If mudtCaptionLocation Then '// [In Frame]
            sngOffset = sngOffset + 2 + sngDrawWidth
         End If

      Case 1 To 6  '// Square Corners
         If mudtCaptionLocation Then '// [In Frame]
            sngOffset = 5 + sngDrawWidth
         Else '// [Inside]
            sngOffset = 2 + sngDrawWidth
         End If

      Case Else '// No border
         sngOffset = 1
      End Select

      sngWidth = UserControl.ScaleWidth
      sngHeight = UserControl.ScaleHeight

      If mudtBevelInner Then
         sngWidth = UserControl.ScaleWidth - (mlngBevelWidth * 2)
         sngHeight = UserControl.ScaleHeight - (mlngBevelWidth * 2)
      End If

      Select Case mudtCaptionAlignment
      Case 0 '// Top left
         msngLeft = sngOffset
         msngTop = sngDrawWidth

      Case 1 '// Top Center
         msngLeft = (sngWidth - lblCaption.Width) / 2
         msngTop = sngDrawWidth

      Case 2 '// Top right
         msngLeft = sngWidth - lblCaption.Width - sngOffset - 1
         msngTop = sngDrawWidth

         If mblnCollapsible Then
            msngLeft = msngLeft - 25
         End If

      Case 3 '// Mid Left
         msngLeft = 2 + sngDrawWidth
         msngTop = (sngHeight - lblCaption.Height + sngDrawWidth) / 2

      Case 4 '// Mid Center
         msngLeft = (sngWidth - lblCaption.Width + sngDrawWidth) / 2
         msngTop = (sngHeight - lblCaption.Height + sngDrawWidth) / 2

      Case 5 '// Mid Right
         msngLeft = sngWidth - lblCaption.Width - 2 - sngDrawWidth - 1
         msngTop = (sngHeight - lblCaption.Height - sngDrawWidth) / 2

      Case 6 '// Bot Left
         msngLeft = sngOffset
         msngTop = sngHeight - lblCaption.Height - 2 - sngDrawWidth

      Case 7 '// Bot Center
         msngLeft = (sngWidth - lblCaption.Width - 1) / 2
         msngTop = sngHeight - lblCaption.Height - 2 - sngDrawWidth

      Case 8 '// Bot Right
         msngLeft = sngWidth - lblCaption.Width - sngOffset - 1
         msngTop = sngHeight - lblCaption.Height - 2 - sngDrawWidth
      End Select

      If mudtBevelInner Then
         msngLeft = msngLeft + mlngBevelWidth
         msngTop = msngTop + mlngBevelWidth
      End If

      Call DrawCaptionStyle
      UserControl.Refresh

   End If '// LenB(lblCaption.Caption)>0

   Exit Sub

Err_Proc:
   ErrHandler False, Err.Number, Err.Description, "Frame3D", "DrawCaptionAlignment"
   Err.Clear
   Resume Next

End Sub

Private Sub DrawCaptionStyle()

   On Error GoTo Err_Proc

   lblCaption.Move msngLeft, msngTop

   Select Case mudtCaption3D
   Case 0 '// Flat
      lblCaptionShadow.Visible = False

   Case 1 '// Inserted
      lblCaptionShadow.Visible = True
      lblCaptionShadow.Move lblCaption.left + 1, lblCaption.top + 1

   Case 2 '// Raised
      lblCaptionShadow.Visible = True
      lblCaptionShadow.Move lblCaption.left - 1, lblCaption.top - 1
   End Select

   lblCaptionBlank.Move lblCaption.left - 1, lblCaption.top - 1, lblCaption.Width + 2, lblCaption.Height + 2

   Exit Sub

Err_Proc:
   ErrHandler False, Err.Number, Err.Description, "Frame3D", "DrawCaptionStyle"
   Err.Clear
   Resume Next

End Sub

Private Sub DrawChevron()

  Dim udtR         As Rect
  Dim lngHDC       As Long
  Dim sngOffset    As Single
  Dim sngWidth     As Single
  Dim sngTop       As Single
  Dim lngForeColor As Long
  Dim udtPoint     As POINTAPI
  Dim intDrawWidth As Integer

   On Error GoTo Err_Proc

   lngHDC = UserControl.hDC
   lngForeColor = UserControl.ForeColor
   UserControl.ForeColor = mlngChevronColor
   intDrawWidth = UserControl.DrawWidth
   UserControl.DrawWidth = 1

   '// Debug only
   '// UserControl.AutoRedraw = False

   '// locate Chevron

   Select Case mudtBorderType
   Case 7 To 13 '// Rounded Corners
      sngOffset = (mlngCornerRadius / 2)

   Case Else
      sngOffset = 0
   End Select

   sngWidth = UserControl.ScaleWidth - (intDrawWidth / 2)
   sngTop = (intDrawWidth / 2) + 2

   If mudtBevelInner Then
      sngWidth = UserControl.ScaleWidth - mlngBevelWidth - ((sngOffset + intDrawWidth) / 2)
      sngTop = sngTop + (mlngBevelWidth * 2) + (intDrawWidth / 2)
   End If

   Call SetRect(udtR, intDrawWidth, sngTop, sngWidth, C_MINHEADERSIZE + (intDrawWidth / 2))

   udtR.left = udtR.right - 20
   udtR.top = udtR.top + (udtR.bottom - udtR.top - (C_MINHEADERSIZE + intDrawWidth)) \ 2 + 2
   udtR.right = udtR.left + 16
   udtR.bottom = udtR.top + 16

   '// Required for PtInRegion - recreate in case border type/size changed
   If mlngRegion Then DeleteObject mlngRegion
   mlngRegion = CreateRectRgn(udtR.left + 1, udtR.top + 1, udtR.right - 2, udtR.top + udtR.bottom - 3)

   '// Draw border

   If mbytChevronType = Square_Chev Then
      '// Move the active point to bottom Left of box
      Call MoveToEx(lngHDC, udtR.left, udtR.bottom, udtPoint)
      '// Draw a line from the active point to the given point
      Call LineTo(lngHDC, udtR.left, udtR.top)          '// left side
      Call LineTo(lngHDC, udtR.right, udtR.top)         '// top
      Call LineTo(lngHDC, udtR.right, udtR.bottom)      '// right side
      Call LineTo(lngHDC, udtR.left, udtR.bottom)       '// bottom
   Else
      UserControl.Circle (udtR.left + 8, udtR.top + 8), 7
   End If

   '// Draw Chevron

   If Not mblnCollapse Then
      '// Top Left Half
      Call MoveToEx(lngHDC, udtR.left + 5, udtR.top + 7, udtPoint)
      Call LineTo(lngHDC, udtR.left + 8, udtR.top + 4)
      Call LineTo(lngHDC, udtR.left + 12, udtR.top + 8)
      '// Top Right Half
      Call MoveToEx(lngHDC, udtR.left + 6, udtR.top + 7, udtPoint)
      Call LineTo(lngHDC, udtR.left + 8, udtR.top + 5)
      Call LineTo(lngHDC, udtR.left + 11, udtR.top + 8)
      '// Bottom Left Half
      Call MoveToEx(lngHDC, udtR.left + 5, udtR.top + 11, udtPoint)
      Call LineTo(lngHDC, udtR.left + 8, udtR.top + 8)
      Call LineTo(lngHDC, udtR.left + 12, udtR.top + 12)
      '// Bottom Right Half
      Call MoveToEx(lngHDC, udtR.left + 6, udtR.top + 11, udtPoint)
      Call LineTo(lngHDC, udtR.left + 8, udtR.top + 9)
      Call LineTo(lngHDC, udtR.left + 11, udtR.top + 12)

   Else
      Call MoveToEx(lngHDC, udtR.left + 5, udtR.top + 5, udtPoint)
      Call LineTo(lngHDC, udtR.left + 8, udtR.top + 8)
      Call LineTo(lngHDC, udtR.left + 12, udtR.top + 4)
      Call MoveToEx(lngHDC, udtR.left + 6, udtR.top + 5, udtPoint)
      Call LineTo(lngHDC, udtR.left + 8, udtR.top + 7)
      Call LineTo(lngHDC, udtR.left + 11, udtR.top + 4)

      Call MoveToEx(lngHDC, udtR.left + 5, udtR.top + 9, udtPoint)
      Call LineTo(lngHDC, udtR.left + 8, udtR.top + 12)
      Call LineTo(lngHDC, udtR.left + 12, udtR.top + 8)
      Call MoveToEx(lngHDC, udtR.left + 6, udtR.top + 9, udtPoint)
      Call LineTo(lngHDC, udtR.left + 8, udtR.top + 11)
      Call LineTo(lngHDC, udtR.left + 11, udtR.top + 8)
   End If

   '// restore default
   UserControl.ForeColor = lngForeColor
   UserControl.DrawWidth = intDrawWidth
   Exit Sub

Err_Proc:
   ErrHandler False, Err.Number, Err.Description, "Frame3D", "DrawChevron"
   Err.Clear
   Resume Next

End Sub

Private Sub DrawControl()

  Dim sngBottom     As Single
  Dim sngTop        As Single
  Dim mlngCornerDia As Long

   On Error GoTo Err_Proc

   sngBottom = UserControl.ScaleHeight
   sngTop = 0!
   lblCaptionBlank.Visible = False

   If mudtBevelInner = 0 Then

      Select Case mudtCaptionAlignment
      Case 0, 1, 2 '// [Top Left], [Top Center], [Top Right]

         If mudtCaptionLocation = 1 And LenB(lblCaption.Caption) Then
            sngTop = lblCaption.Height / 2.25!
            lblCaptionBlank.Visible = True
         End If

      Case 6, 7, 8 '// [Bottom Left], [Bottom Center], [Bottom Right]

         If mudtCaptionLocation = 1 And LenB(lblCaption.Caption) Then
            sngBottom = UserControl.ScaleHeight - (lblCaption.Height / 2.25!)
            lblCaptionBlank.Visible = True
         End If

      End Select
   End If

   msngInsideBorder = sngTop
   mlngCornerDia = mlngCornerRadius * 2

   With UserControl
      .DrawMode = vbCopyPen
      .Cls
      .BackColor = mlngBackColor
      lblCaptionBlank.BackColor = mlngBackColor
      .Enabled = mblnEnabled
      If mudtFillGradient Then Call DrawGradiant

      '// Get inside workspace size in twips

      Select Case mudtBorderType
      Case 1, 2, 7, 8 '// Frame Inserted/Raised
         mlngInsideHeight = (.ScaleHeight - 4) * Screen.TwipsPerPixelY
         mlngInsideWidth = (.ScaleWidth - 4) * Screen.TwipsPerPixelX
         mlngInsideLeft = 2 * Screen.TwipsPerPixelX
         mlngInsideTop = 2 * Screen.TwipsPerPixelY

      Case Else
         mlngInsideHeight = (.ScaleHeight - 2) * Screen.TwipsPerPixelY
         mlngInsideWidth = (.ScaleWidth - 2) * Screen.TwipsPerPixelX
         mlngInsideLeft = Screen.TwipsPerPixelX
         mlngInsideTop = Screen.TwipsPerPixelY
      End Select

      '// Draw Border Type

      Select Case mudtBorderType
      Case 0 '// [None Border]

      Case 1 '// [Frame Inserted]
         .ForeColor = mlng3DShadow
         Call RoundRect(.hDC, 0&, sngTop, .ScaleWidth - 1, sngBottom - 1, 0&, 0&)
         .ForeColor = mlng3DHighlight
         Call RoundRect(.hDC, 1&, sngTop + 1&, .ScaleWidth, sngBottom, 0&, 0&)
         Call SetWindowRgn(.HWND, CreateRoundRectRgn(0&, 0&, .ScaleWidth + 1, .ScaleHeight + 1, 0&, 0&), True)

      Case 2 '// [Frame Raised]
         .ForeColor = mlng3DHighlight
         Call RoundRect(.hDC, 0&, sngTop, .ScaleWidth - 1, sngBottom - 1, 0&, 0&)
         .ForeColor = mlng3DShadow
         Call RoundRect(.hDC, 1&, sngTop + 1&, .ScaleWidth, sngBottom, 0&, 0&)
         Call SetWindowRgn(.HWND, CreateRoundRectRgn(0&, 0&, .ScaleWidth + 1, .ScaleHeight + 1, 0&, 0&), True)

      Case 3 '// [Panel Flat Shadow]
         .ForeColor = mlng3DShadow
         Call RoundRect(.hDC, 0&, sngTop, .ScaleWidth, sngBottom, 0&, 0&)
         Call SetWindowRgn(.HWND, CreateRoundRectRgn(0&, 0&, .ScaleWidth + 1, .ScaleHeight + 1, 0&, 0&), True)

      Case 4 '// [Panel Flat Highlight]
         .ForeColor = mlng3DHighlight
         Call RoundRect(.hDC, 0&, sngTop, .ScaleWidth, sngBottom, 0&, 0&)
         Call SetWindowRgn(.HWND, CreateRoundRectRgn(0&, 0&, .ScaleWidth + 1, .ScaleHeight + 1, 0&, 0&), True)

      Case 5 '// [Panel Raised]
         lblCaptionBlank.Visible = False
         '// Trick to allow FillStyle to work
         .ForeColor = .BackColor
         Call RoundRect(.hDC, 0&, 0&, .ScaleWidth, .ScaleHeight, 0&, 0&)
         UserControl.Line (.ScaleWidth - 1, .ScaleHeight - 1)-(.ScaleWidth - 1, 0), mlng3DShadow
         UserControl.Line (.ScaleWidth - 1, .ScaleHeight - 1)-(0, .ScaleHeight - 1), mlng3DShadow
         UserControl.Line (.ScaleWidth - 1, 0)-(0, 0), mlng3DHighlight
         UserControl.Line (0, .ScaleHeight - 1)-(0, 0), mlng3DHighlight
         Call SetWindowRgn(.HWND, CreateRoundRectRgn(0&, 0&, .ScaleWidth + 1, .ScaleHeight + 1, 0&, 0&), True)

      Case 6 '// [Inserted Panel Square Corners]
         lblCaptionBlank.Visible = False
         '// Trick to allow FillStyle to work
         .ForeColor = .BackColor
         Call RoundRect(.hDC, 0&, 0&, .ScaleWidth, .ScaleHeight, 0&, 0&)
         UserControl.Line (.ScaleWidth - 1, .ScaleHeight - 1)-(.ScaleWidth - 1, 0), mlng3DHighlight
         UserControl.Line (.ScaleWidth - 1, .ScaleHeight - 1)-(0, .ScaleHeight - 1), mlng3DHighlight
         UserControl.Line (.ScaleWidth - 2, 0)-(0, 0), mlng3DShadow
         UserControl.Line (0, .ScaleHeight - 2)-(0, 0), mlng3DShadow
         Call SetWindowRgn(.HWND, CreateRoundRectRgn(0&, 0&, .ScaleWidth + 1, .ScaleHeight + 1, 0&, 0&), True)

      Case 7 '// [rFrame Inserted]
         .ForeColor = mlng3DShadow
         Call RoundRect(.hDC, 0&, sngTop, .ScaleWidth - 1, sngBottom - 1, mlngCornerDia, mlngCornerDia)
         .ForeColor = mlng3DHighlight
         Call RoundRect(.hDC, 1&, sngTop + 1&, .ScaleWidth, sngBottom, mlngCornerDia, mlngCornerDia)
         '// Make corners transparent
         Call SetWindowRgn(.HWND, CreateRoundRectRgn(0&, 0&, .ScaleWidth + 1, .ScaleHeight + 1, mlngCornerDia, mlngCornerDia), True)

      Case 8 '// [rFrame Raised]
         .ForeColor = mlng3DHighlight
         Call RoundRect(.hDC, 0&, sngTop, .ScaleWidth - 1, sngBottom - 1, mlngCornerDia, mlngCornerDia)
         .ForeColor = mlng3DShadow
         Call RoundRect(.hDC, 1&, sngTop + 1&, .ScaleWidth, sngBottom, mlngCornerDia, mlngCornerDia)
         '// Make corners transparent
         Call SetWindowRgn(.HWND, CreateRoundRectRgn(0&, 0&, .ScaleWidth + 1, .ScaleHeight + 1, mlngCornerDia, mlngCornerDia), True)

      Case 9 '// [rPanel Flat Shadow]
         .ForeColor = mlng3DShadow
         Call RoundRect(.hDC, 0&, sngTop, .ScaleWidth, sngBottom, mlngCornerDia, mlngCornerDia)
         Call SetWindowRgn(.HWND, CreateRoundRectRgn(0&, 0&, .ScaleWidth + 1, .ScaleHeight + 1, mlngCornerDia, mlngCornerDia), True)

      Case 10 '// [rPanel Flat Highlight]
         .ForeColor = mlng3DHighlight
         Call RoundRect(.hDC, 0&, sngTop, .ScaleWidth, sngBottom, mlngCornerDia, mlngCornerDia)
         Call SetWindowRgn(.HWND, CreateRoundRectRgn(0&, 0&, .ScaleWidth + 1, .ScaleHeight + 1, mlngCornerDia, mlngCornerDia), True)

      Case 11 '// [rPanel Raised]
         lblCaptionBlank.Visible = False
         '// Trick to allow FillStyle to work
         .ForeColor = .BackColor
         Call RoundRect(.hDC, 0&, 0&, .ScaleWidth, .ScaleHeight, mlngCornerDia, mlngCornerDia)
         '// Top
         UserControl.Line (mlngCornerRadius, 0)-(.ScaleWidth - mlngCornerRadius - 1, 0), mlng3DHighlight
         '// Left
         UserControl.Line (0, mlngCornerRadius)-(0, .ScaleHeight - mlngCornerRadius - 1), mlng3DHighlight
         '// Right
         UserControl.Line (.ScaleWidth - 1, mlngCornerRadius)-(.ScaleWidth - 1, .ScaleHeight - mlngCornerRadius - 1), mlng3DShadow
         '// Bottom
         UserControl.Line (mlngCornerRadius, .ScaleHeight - 1)-(.ScaleWidth - mlngCornerRadius - 1, .ScaleHeight - 1), mlng3DShadow
         '// Top Left
         UserControl.Circle (mlngCornerRadius, mlngCornerRadius), mlngCornerRadius, mlng3DHighlight, 1.57, 3.14
         '// Bottom Left
         UserControl.Circle (mlngCornerRadius, .ScaleHeight - mlngCornerRadius - 1), mlngCornerRadius, mlng3DHighlight, 3.14, 4.71
         '// Top Right
         UserControl.Circle (.ScaleWidth - mlngCornerRadius - 1, mlngCornerRadius), mlngCornerRadius, mlng3DShadow, 6.28, 1.57
         '// Bottom Right
         UserControl.Circle (.ScaleWidth - mlngCornerRadius - 1, .ScaleHeight - mlngCornerRadius - 1), mlngCornerRadius, mlng3DShadow, 4.71
         Call SetWindowRgn(.HWND, CreateRoundRectRgn(0&, 0&, .ScaleWidth + 1, .ScaleHeight + 1, mlngCornerDia, mlngCornerDia), True)

      Case 12 '// [Inserted Panel Round Corners]
         lblCaptionBlank.Visible = False
         '// Trick to allow FillStyle to work
         .ForeColor = .BackColor
         Call RoundRect(.hDC, 0&, 0&, .ScaleWidth, .ScaleHeight, mlngCornerDia, mlngCornerDia)
         '// Top
         UserControl.Line (mlngCornerRadius, 0)-(.ScaleWidth - mlngCornerRadius - 1, 0), mlng3DShadow
         '// Left
         UserControl.Line (0, mlngCornerRadius)-(0, .ScaleHeight - mlngCornerRadius - 1), mlng3DShadow
         '// Right
         UserControl.Line (.ScaleWidth - 1, mlngCornerRadius)-(.ScaleWidth - 1, .ScaleHeight - mlngCornerRadius - 1), mlng3DHighlight
         '// Bottom
         UserControl.Line (mlngCornerRadius, .ScaleHeight - 1)-(.ScaleWidth - mlngCornerRadius - 1, .ScaleHeight - 1), mlng3DHighlight
         '// Top Left
         UserControl.Circle (mlngCornerRadius, mlngCornerRadius), mlngCornerRadius, mlng3DShadow, 1.57, 3.14
         '// Bottom Left
         UserControl.Circle (mlngCornerRadius, .ScaleHeight - mlngCornerRadius - 1), mlngCornerRadius, mlng3DShadow, 3.14, 4.71
         '// Top Right
         UserControl.Circle (.ScaleWidth - mlngCornerRadius - 1, mlngCornerRadius), mlngCornerRadius, mlng3DHighlight, 6.28, 1.57
         '// Bottom Right
         UserControl.Circle (.ScaleWidth - mlngCornerRadius - 1, .ScaleHeight - mlngCornerRadius - 1), mlngCornerRadius, mlng3DHighlight, 4.71
         Call SetWindowRgn(.HWND, CreateRoundRectRgn(0&, 0&, .ScaleWidth + 1, .ScaleHeight + 1, mlngCornerRadius, mlngCornerRadius), True)

      Case Else '// [None rBorder]
         '// Trick to allow FillStyle to work
         .ForeColor = .BackColor
         Call RoundRect(.hDC, 0&, 0&, .ScaleWidth, .ScaleHeight, 0&, 0&)
         Call SetWindowRgn(.HWND, CreateRoundRectRgn(1&, 1&, .ScaleWidth, .ScaleHeight, mlngCornerDia, mlngCornerDia), True)
      End Select

   End With

   Call DrawBevelInner
   Call DrawCaptionAlignment
   If mblnCollapsible Then Call DrawChevron

   UserControl.Refresh

   Exit Sub

Err_Proc:
   ErrHandler False, Err.Number, Err.Description, "Frame3D", "DrawControl"
   Err.Clear
   Resume Next

End Sub

Private Sub DrawFlood()

  Dim sngBevelWidth As Single
  Dim sngNewValue   As Single

   On Error GoTo Err_Proc

   '// Show percent complete?

   If mblnFloodShowPct Then
      lblCaption.Caption = CStr(mlngFloodValue) & "%"
      lblCaptionShadow.Caption = CStr(mlngFloodValue) & "%"
      Call DrawCaptionAlignment
   End If

   If mlngFloodValue Then

      '// Is there an inside border showing?

      If mudtBevelInner Then
         sngBevelWidth = mlngBevelWidth + msngInsideBorder + 1
      Else
         sngBevelWidth = msngInsideBorder + 1
      End If

      '// Flood Fill
      If mudtFloodType Then  '// [Bottom To Top]

         sngNewValue = UserControl.ScaleHeight - sngBevelWidth - sngBevelWidth - 1
         sngNewValue = sngNewValue - (sngNewValue * (mlngFloodValue / 100))

         UserControl.Line (sngBevelWidth, UserControl.ScaleHeight - sngBevelWidth - 1)-(UserControl.ScaleWidth - sngBevelWidth - 1, sngNewValue + _
            sngBevelWidth), mlngFloodColor, BF

      Else '/[Left To Right]

         sngNewValue = (UserControl.ScaleWidth - sngBevelWidth - sngBevelWidth) * mlngFloodValue / 100

         UserControl.Line (sngBevelWidth, sngBevelWidth)-(sngNewValue + sngBevelWidth - 1, UserControl.ScaleHeight - sngBevelWidth - 1), _
            mlngFloodColor, BF
      End If

   End If

Exit_Proc:
   Exit Sub

Err_Proc:
   ErrHandler False, Err.Number, Err.Description, "Frame3D", "DrawFlood"
   Err.Clear
   Resume Next

End Sub

Public Sub DrawGradiant()

  Dim lngI    As Long
  Dim lngStep As Long
  Dim sngRed1 As Single
  Dim sngGrn1 As Single
  Dim sngBlu1 As Single
  Dim sngRed2 As Single
  Dim sngGrn2 As Single
  Dim sngBlu2 As Single

   On Error Resume Next

   Call GetRGBColor(UserControl.FillColor, sngRed1, sngGrn1, sngBlu1)
   Call GetRGBColor(UserControl.BackColor, sngRed2, sngGrn2, sngBlu2)

   With UserControl

      Select Case mudtFillGradient
      Case 1 '// [Horizontal]
         lngStep = .ScaleWidth - mlngCornerRadius
         '// Get gradient color step
         sngRed2 = (sngRed2 - sngRed1) / lngStep
         sngGrn2 = (sngGrn2 - sngGrn1) / lngStep
         sngBlu2 = (sngBlu2 - sngBlu1) / lngStep

         '// Begin drawing horizontal gradient

         For lngI = 0 To lngStep
            UserControl.Line (lngI, 0)-(lngI, .ScaleHeight), RGB(CInt(sngRed1), CInt(sngGrn1), CInt(sngBlu1))
            sngRed1 = sngRed1 + sngRed2
            sngGrn1 = sngGrn1 + sngGrn2
            sngBlu1 = sngBlu1 + sngBlu2
         Next lngI

      Case 2 '// [Vertical]
         lngStep = .ScaleHeight - mlngCornerRadius
         '// Get gradient color step
         sngRed2 = (sngRed2 - sngRed1) / lngStep
         sngGrn2 = (sngGrn2 - sngGrn1) / lngStep
         sngBlu2 = (sngBlu2 - sngBlu1) / lngStep

         '// Begin drawing vertical gradient

         For lngI = 0 To lngStep
            UserControl.Line (0, lngI)-(.ScaleWidth, lngI), RGB(CInt(sngRed1), CInt(sngGrn1), CInt(sngBlu1))
            sngRed1 = sngRed1 + sngRed2
            sngGrn1 = sngGrn1 + sngGrn2
            sngBlu1 = sngBlu1 + sngBlu2
         Next lngI

      End Select
   End With

End Sub

Public Property Get DrawStyle() As DrawStyleConstants

   DrawStyle = UserControl.DrawStyle

End Property

Public Property Let DrawStyle(ByVal vNewValue As DrawStyleConstants)

   UserControl.DrawStyle = vNewValue
   PropertyChanged "DrawStyle"
   Call UserControl_Resize

End Property

Public Property Get DrawWidth() As Integer

   DrawWidth = UserControl.DrawWidth

End Property

Public Property Let DrawWidth(ByVal vNewValue As Integer)

   If vNewValue > 0 Then
      UserControl.DrawWidth = vNewValue
      PropertyChanged "DrawWidth"
      Call UserControl_Resize
   End If

End Property

Public Property Get Enabled() As Boolean

   Enabled = mblnEnabled

End Property

Public Property Let Enabled(ByVal vNewValue As Boolean)

   mblnEnabled = vNewValue
   PropertyChanged "Enabled"
   Call UserControl_Resize

End Property

Private Sub ErrHandler(Optional ByVal vblnDisplayError As Boolean = True, _
                       Optional ByVal vstrErrNumber As String = vbNullString, _
                       Optional ByVal vstrErrDescription As String = vbNullString, _
                       Optional ByVal vstrModuleName As String = vbNullString, _
                       Optional ByVal vstrProcName As String = vbNullString)

  Dim strTemp     As String
  Dim lngFN       As Long

   On Error Resume Next
   '// Purpose: Error handling - On Error

   '// Show Error Message

   If vblnDisplayError Then
      strTemp = "Error occured: "

      If LenB(vstrErrNumber) Then
         strTemp = strTemp & vstrErrNumber & vbNewLine
      Else
         strTemp = strTemp & vbNewLine
      End If

      If LenB(vstrErrDescription) Then strTemp = strTemp & "Description: " & vstrErrDescription & vbNewLine
      If LenB(vstrModuleName) Then strTemp = strTemp & "Module: " & vstrModuleName & vbNewLine
      If LenB(vstrProcName) Then strTemp = strTemp & "Function: " & vstrProcName
      MsgBox strTemp, vbCritical, App.Title & " - ERROR"
   End If

   '// Write error log
   lngFN = FreeFile
   Open App.Path & "\ErrorLog.txt" For Append As #lngFN
   Write #lngFN, Now, vstrErrNumber, vstrErrDescription, vstrModuleName, vstrProcName, App.Title & " v" & App.Major & "." & App.Minor & "." & _
      App.Revision, Environ("username"), Environ("computername")
   Close #lngFN

End Sub

Public Property Get FillColor() As OLE_COLOR

   FillColor = UserControl.FillColor

End Property

Public Property Let FillColor(ByVal vNewValue As OLE_COLOR)

   UserControl.FillColor = vNewValue
   PropertyChanged "FillColor"
   Call UserControl_Resize

End Property

Public Property Let FillGradient(ByVal vNewValue As enuFillGradient)

   mudtFillGradient = vNewValue
   PropertyChanged "FillGradient"
   Call UserControl_Resize

End Property

Public Property Get FillGradient() As enuFillGradient

   FillGradient = mudtFillGradient

End Property

Public Property Get FillStyle() As FillStyleConstants

   FillStyle = UserControl.FillStyle

End Property

Public Property Let FillStyle(ByVal vNewValue As FillStyleConstants)

   UserControl.FillStyle = vNewValue
   PropertyChanged "FillStyle"
   Call UserControl_Resize

End Property

Public Property Let FloodColor(ByVal vNewValue As OLE_COLOR)

   mlngFloodColor = vNewValue
   PropertyChanged "FloodColor"
   Call DrawFlood

End Property

Public Property Get FloodColor() As OLE_COLOR

   FloodColor = mlngFloodColor

End Property

Public Property Let FloodPercent(ByVal vNewValue As Long)

   '// Fix the value

   Select Case vNewValue
   Case Is > 100
      vNewValue = 100

   Case Is < 0
      vNewValue = 0
   End Select

   '// Clear old values if decreasing

   If vNewValue <= mlngFloodValue Then
      '// Save new property value
      mlngFloodValue = vNewValue
      PropertyChanged "FloodPercent"
      Call DrawFlood
      Call UserControl_Resize
   Else
      '// Save new property value
      mlngFloodValue = vNewValue
      PropertyChanged "FloodPercent"
      Call DrawFlood
   End If

End Property

Public Property Get FloodPercent() As Long

   FloodPercent = mlngFloodValue

End Property

Public Property Let FloodShowPct(ByVal vNewValue As Boolean)

   mblnFloodShowPct = vNewValue
   PropertyChanged "FloodShowPct"

   If mblnFloodShowPct Then
      lblCaption.Caption = CStr(mlngFloodValue) & "%"
      lblCaptionShadow.Caption = CStr(mlngFloodValue) & "%"
   Else
      lblCaption.Caption = vbNullString
      lblCaptionShadow.Caption = vbNullString
   End If

End Property

Public Property Get FloodShowPct() As Boolean

   FloodShowPct = mblnFloodShowPct

End Property

Public Property Get FloodType() As enuFloodType

   FloodType = mudtFloodType

End Property

Public Property Let FloodType(ByVal vNewValue As enuFloodType)

   mudtFloodType = vNewValue
   PropertyChanged "FloodType"
   Call UserControl_Resize

End Property

Public Property Get Font() As Font

   Set Font = lblCaption.Font

End Property

Public Property Set Font(ByRef vNewValue As Font)

   Set lblCaption.Font = vNewValue
   Set lblCaptionShadow.Font = vNewValue
   PropertyChanged "Font"
   Call UserControl_Resize

End Property

Public Property Get FontBold() As Boolean

   FontBold = lblCaption.FontBold

End Property

Public Property Let FontBold(ByVal vNewValue As Boolean)

   lblCaption.FontBold = vNewValue
   lblCaptionShadow.FontBold = vNewValue
   PropertyChanged "FontBold"
   Call UserControl_Resize

End Property

Public Property Get FontItalic() As Boolean

   FontItalic = lblCaption.FontItalic

End Property

Public Property Let FontItalic(ByVal vNewValue As Boolean)

   lblCaption.FontItalic = vNewValue
   lblCaptionShadow.FontItalic = vNewValue
   PropertyChanged "FontItalic"
   Call UserControl_Resize

End Property

Public Property Get FontName() As String

   FontName = lblCaption.FontName

End Property

Public Property Let FontName(ByVal vNewValue As String)

   lblCaption.FontName = vNewValue
   lblCaptionShadow.FontName = vNewValue
   PropertyChanged "FontName"
   Call UserControl_Resize

End Property

Public Property Get FontSize() As Long

   FontSize = lblCaption.FontSize

End Property

Public Property Let FontSize(ByVal vNewValue As Long)

   lblCaption.FontSize = vNewValue
   lblCaptionShadow.FontSize = vNewValue
   PropertyChanged "FontSize"
   Call UserControl_Resize

End Property

Public Property Let FontUnderline(ByVal vNewValue As Boolean)

   lblCaption.FontUnderline = vNewValue
   lblCaptionShadow.FontUnderline = vNewValue
   PropertyChanged "FontUnderline"

End Property

Public Property Get FontUnderline() As Boolean

   FontUnderline = lblCaption.FontUnderline

End Property

Public Property Get ForeColor() As OLE_COLOR

   ForeColor = lblCaption.ForeColor

End Property

Public Property Let ForeColor(ByVal vNewValue As OLE_COLOR)

   lblCaption.ForeColor = vNewValue
   PropertyChanged "ForeColor"

End Property

Private Sub GetRGBColor(ByVal vlngColor As Long, ByRef rsngRed As Single, ByRef rsngGrn As Single, ByRef rsngBlu As Single)

   On Error GoTo Err_Proc

   '// Is the color a VB color constant?

   If vlngColor < 0 Then
      '// Retrieves the current color of the specified display element
      vlngColor = GetSysColor(vlngColor And &HFF&)
   End If

   '// Separate the color into it's RGB values
   rsngRed = CSng((vlngColor And &HFF&))
   rsngGrn = CSng((vlngColor And &HFF00&) \ &H100&)
   rsngBlu = CSng((vlngColor And &HFF0000) \ &H10000)

   '// These passed values would normally be declared as Longs but
   '// the calling sub requires Singles

Exit_Proc:
   Exit Sub

Err_Proc:
   ErrHandler True, Err.Number, Err.Description, "Frame3D", "GetRGBColor"
   Err.Clear
   Resume Exit_Proc

End Sub

Public Property Get hDC() As Long

   hDC = UserControl.hDC

End Property

Public Property Get HWND() As Long

   HWND = UserControl.HWND

End Property

Public Property Get InsideHeight() As Long

   InsideHeight = mlngInsideHeight

End Property

Public Property Get InsideLeft() As Long

   InsideLeft = mlngInsideLeft

End Property

Public Property Get InsideTop() As Long

   InsideTop = mlngInsideTop

End Property

Public Property Get InsideWidth() As Long

   InsideWidth = mlngInsideWidth

End Property

Private Sub lblCaptionBlank_Click()

   RaiseEvent Click

End Sub

Private Sub lblCaptionBlank_DblClick()

   RaiseEvent DblClick

End Sub

Private Sub lblCaptionBlank_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

   RaiseEvent MouseDown(Button, Shift, x, Y)

End Sub

Private Sub lblCaptionBlank_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

   RaiseEvent MouseMove(Button, Shift, x, Y)

End Sub

Private Sub lblCaptionBlank_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

   RaiseEvent MouseUp(Button, Shift, x, Y)

End Sub

Private Sub lblCaptionShadow_Click()

   RaiseEvent Click

End Sub

Private Sub lblCaptionShadow_DblClick()

   RaiseEvent DblClick

End Sub

Private Sub lblCaptionShadow_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

   RaiseEvent MouseDown(Button, Shift, x, Y)

End Sub

Private Sub lblCaptionShadow_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

   RaiseEvent MouseMove(Button, Shift, x, Y)

End Sub

Private Sub lblCaptionShadow_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

   RaiseEvent MouseUp(Button, Shift, x, Y)

End Sub

Private Sub lblCaption_Click()

   RaiseEvent Click

End Sub

Private Sub lblCaption_DblClick()

   RaiseEvent DblClick

End Sub

Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

   RaiseEvent MouseDown(Button, Shift, x, Y)

End Sub

Private Sub lblCaption_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

   RaiseEvent MouseMove(Button, Shift, x, Y)

End Sub

Private Sub lblCaption_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

   RaiseEvent MouseUp(Button, Shift, x, Y)

End Sub

Public Property Get MouseIcon() As StdPicture

   Set MouseIcon = UserControl.MouseIcon

End Property

Public Property Set MouseIcon(ByVal vNewValue As StdPicture)

   On Local Error Resume Next
   Set UserControl.MouseIcon = vNewValue
   PropertyChanged "MouseIcon"
   On Local Error GoTo 0

End Property

Public Property Get MousePointer() As MousePointerConstants

   MousePointer = UserControl.MousePointer

End Property

Public Property Let MousePointer(ByVal vNewValue As MousePointerConstants)

   UserControl.MousePointer = vNewValue
   PropertyChanged "MousePointer"

End Property

Public Property Get Picture() As StdPicture

   Set Picture = UserControl.Picture

End Property

Public Property Set Picture(ByVal vNewValue As StdPicture)

   On Local Error Resume Next
   Set UserControl.Picture = vNewValue
   PropertyChanged "Picture"
   On Local Error GoTo 0

End Property

Private Sub SetMinHeight()

  Dim sngTop As Single

   On Error GoTo Err_Proc

   If mudtBevelInner Then
      sngTop = C_MINHEADERSIZE + (mlngBevelWidth * 2) + UserControl.DrawWidth + 1
   Else
      sngTop = C_MINHEADERSIZE + UserControl.DrawWidth + 1
   End If

   UserControl.Height = sngTop * Screen.TwipsPerPixelY

   Exit Sub

Err_Proc:
   ErrHandler False, Err.Number, Err.Description, "Frame3D", "SetMinHeight"
   Err.Clear
   Resume Next

End Sub

Public Property Let UseMnemonic(ByVal vNewValue As Boolean)

   lblCaption.UseMnemonic = vNewValue
   lblCaptionShadow.UseMnemonic = vNewValue
   PropertyChanged "UseMnemonic"
   Call UserControl_Resize

End Property

Public Property Get UseMnemonic() As Boolean

   UseMnemonic = lblCaption.UseMnemonic

End Property

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)

   Call UserControl_Click

End Sub

Private Sub UserControl_Click()

   RaiseEvent Click

End Sub

Private Sub UserControl_DblClick()

   RaiseEvent DblClick

End Sub

Private Sub UserControl_InitProperties()

   On Error GoTo Err_Proc

   mlngBevelWidth = 3&
   mudtBorderType = 1 '// [Frame Inserted]
   mlngBackColor = UserControl.Parent.BackColor
   mlng3DHighlight = vb3DHighlight
   mlng3DShadow = vb3DShadow
   mblnEnabled = True
   mlngFloodValue = 0&
   mblnFloodShowPct = False
   mudtFloodType = 0 '// [Left To Right]
   mlngFloodColor = UserControl.FillColor
   lblCaptionShadow.ForeColor = mlng3DHighlight
   lblCaption.UseMnemonic = False
   lblCaptionShadow.UseMnemonic = False
   mudtCaptionLocation = [Inside Frame]
   mlngCornerRadius = 7&
   mlngCornerDia = mlngCornerRadius * 2
   mblnCollapsible = False
   mblnCollapse = False
   mlngChevronColor = UserControl.Parent.ForeColor
   mlngFullHeight = UserControl.Height

Exit_Proc:
   Exit Sub

Err_Proc:
   ErrHandler True, Err.Number, Err.Description, "Frame3D", "UserControl_InitProperties"
   Err.Clear
   Resume Next

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

   RaiseEvent KeyDown(KeyCode, Shift)

End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)

   RaiseEvent KeyPress(KeyAscii)

End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)

   RaiseEvent KeyUp(KeyCode, Shift)

End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

   RaiseEvent MouseDown(Button, Shift, x, Y)

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

   RaiseEvent MouseMove(Button, Shift, x, Y)

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

   RaiseEvent MouseUp(Button, Shift, x, Y)

   '// Mouse click over Chevron?

   If mblnCollapsible Then
      If Button = 1 Then
         Dim PT As POINTAPI

         If GetCursorPos(PT) Then
            If ScreenToClient(UserControl.HWND, PT) Then
               If PtInRegion(mlngRegion, PT.x, PT.Y) Then

                  mblnCollapse = Not mblnCollapse

                  If mblnCollapse Then
                     Call SetMinHeight
                     RaiseEvent ChevronClick(Button, Shift, x, Y, True, UserControl.Height)
                  Else
                     UserControl.Height = mlngFullHeight
                     RaiseEvent ChevronClick(Button, Shift, x, Y, False, UserControl.Height)
                  End If

               End If
            End If
         End If
      End If
   End If

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

   On Error GoTo Err_Proc

   With PropBag
      mudtBorderType = .ReadProperty("BorderType", 1)
      mlngBevelWidth = .ReadProperty("BevelWidth", 3&)
      mudtBevelInner = .ReadProperty("BevelInner", 0)
      mudtCaption3D = .ReadProperty("Caption3D", 0)
      mudtCaptionAlignment = .ReadProperty("CaptionAlignment", 0)
      mudtCaptionLocation = .ReadProperty("CaptionLocation", 0)
      mlngBackColor = .ReadProperty("BackColor", UserControl.Parent.BackColor)
      mlng3DHighlight = .ReadProperty("Border3DHighlight", vb3DHighlight)
      mlng3DShadow = .ReadProperty("Border3DShadow", vb3DShadow)
      mblnEnabled = .ReadProperty("Enabled", True)
      mlngCornerRadius = .ReadProperty("CornerDiameter", 7&)
      mlngCornerDia = mlngCornerRadius * 2

      mlngFloodValue = .ReadProperty("FloodPercent", 0&)
      mblnFloodShowPct = .ReadProperty("FloodShowPct", 0)
      mudtFloodType = .ReadProperty("FloodType", 0)
      mlngFloodColor = .ReadProperty("FloodColor", UserControl.FillColor)
      mudtFillGradient = .ReadProperty("FillGradient", 0)

      mblnCollapsible = .ReadProperty("Collapsible", 0)
      mlngChevronColor = .ReadProperty("ChevronColor", UserControl.ForeColor)
      mblnCollapse = .ReadProperty("Collapse", 0)
      mlngFullHeight = .ReadProperty("FullHeight", UserControl.Height)
      mbytChevronType = .ReadProperty("ChevronType", 0)

      UserControl.FillColor = .ReadProperty("FillColor", UserControl.FillColor)
      UserControl.FillStyle = .ReadProperty("FillStyle", UserControl.FillStyle)
      UserControl.DrawStyle = .ReadProperty("DrawStyle", UserControl.DrawStyle)
      UserControl.MousePointer = .ReadProperty("MousePointer", UserControl.MousePointer)
      UserControl.MouseIcon = .ReadProperty("MouseIcon", UserControl.MouseIcon)
      UserControl.Picture = .ReadProperty("Picture", UserControl.Picture)
      UserControl.DrawWidth = .ReadProperty("DrawWidth", UserControl.DrawWidth)

      lblCaption.Alignment = .ReadProperty("CaptionMAlignment", lblCaption.Alignment)
      lblCaption.Font = .ReadProperty("Font", lblCaption.Font)
      lblCaption.FontBold = .ReadProperty("FontBold", lblCaption.FontBold)
      lblCaption.FontItalic = .ReadProperty("FontItalic", lblCaption.FontItalic)
      lblCaption.FontName = .ReadProperty("FontName", lblCaption.FontName)
      lblCaption.FontSize = .ReadProperty("FontSize", lblCaption.FontSize)
      lblCaption.FontStrikethru = .ReadProperty("FontStrikethru", lblCaption.FontStrikethru)
      lblCaption.FontUnderline = .ReadProperty("FontUnderline", lblCaption.FontUnderline)
      lblCaption.ForeColor = .ReadProperty("ForeColor", lblCaption.ForeColor)
      lblCaption.Caption = .ReadProperty("Caption", lblCaption.Caption)
      lblCaption.UseMnemonic = .ReadProperty("UseMnemonic", lblCaption.UseMnemonic)

      lblCaptionShadow.ForeColor = mlng3DHighlight

   End With

   If mblnFloodShowPct Then
      lblCaption.Caption = CStr(mlngFloodValue) & "%"
      lblCaptionShadow.Caption = CStr(mlngFloodValue) & "%"
   Else
      '// Trick to fix right justified text
      lblCaption.Caption = lblCaption.Caption & " "
      lblCaption.Caption = left$(lblCaption.Caption, Len(lblCaption.Caption) - 1)
   End If

   With lblCaptionShadow
      .Alignment = lblCaption.Alignment
      .Font = lblCaption.Font
      .FontBold = lblCaption.FontBold
      .FontItalic = lblCaption.FontItalic
      .FontName = lblCaption.FontName
      .FontSize = lblCaption.FontSize
      .FontStrikethru = lblCaption.FontStrikethru
      .FontUnderline = lblCaption.FontUnderline
      .Caption = lblCaption.Caption
      .UseMnemonic = lblCaption.UseMnemonic
   End With

   Call DrawControl
   
   Exit Sub

Err_Proc:
   ErrHandler True, Err.Number, Err.Description, "Frame3D", "UserControl_ReadProperties"
   Err.Clear
   Resume Next

End Sub

Private Sub UserControl_Resize()

   If Not mblnCollapse Then
      mlngFullHeight = UserControl.Height
      PropertyChanged "FullHeight"
   End If

   Call DrawControl

End Sub

Private Sub UserControl_Terminate()

   '// Clean-up
   If mlngRegion Then DeleteObject mlngRegion

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

   On Error GoTo Err_Proc

   With PropBag
      .WriteProperty "BorderType", mudtBorderType
      .WriteProperty "BevelWidth", mlngBevelWidth
      .WriteProperty "BevelInner", mudtBevelInner
      .WriteProperty "Caption3D", mudtCaption3D
      .WriteProperty "CaptionAlignment", mudtCaptionAlignment
      .WriteProperty "CaptionLocation", mudtCaptionLocation
      .WriteProperty "BackColor", mlngBackColor
      .WriteProperty "CornerDiameter", mlngCornerRadius

      .WriteProperty "FillColor", UserControl.FillColor
      .WriteProperty "FillStyle", UserControl.FillStyle
      .WriteProperty "DrawStyle", UserControl.DrawStyle
      .WriteProperty "DrawWidth", UserControl.DrawWidth

      .WriteProperty "FloodPercent", mlngFloodValue
      .WriteProperty "FloodShowPct", mblnFloodShowPct
      .WriteProperty "FloodType", mudtFloodType
      .WriteProperty "FloodColor", mlngFloodColor
      .WriteProperty "FillGradient", mudtFillGradient

      .WriteProperty "Collapsible", mblnCollapsible
      .WriteProperty "ChevronColor", mlngChevronColor
      .WriteProperty "Collapse", mblnCollapse
      .WriteProperty "FullHeight", mlngFullHeight
      .WriteProperty "ChevronType", mbytChevronType

      .WriteProperty "MousePointer", UserControl.MousePointer
      .WriteProperty "MouseIcon", UserControl.MouseIcon
      .WriteProperty "Picture", UserControl.Picture

      .WriteProperty "Border3DHighlight", mlng3DHighlight
      .WriteProperty "Border3DShadow", mlng3DShadow

      .WriteProperty "Enabled", mblnEnabled

      .WriteProperty "CaptionMAlignment", lblCaption.Alignment
      .WriteProperty "Font", lblCaption.Font
      .WriteProperty "FontBold", lblCaption.FontBold
      .WriteProperty "FontItalic", lblCaption.FontItalic
      .WriteProperty "FontName", lblCaption.FontName
      .WriteProperty "FontSize", lblCaption.FontSize
      .WriteProperty "FontStrikethru", lblCaption.FontStrikethru
      .WriteProperty "FontUnderline", lblCaption.FontUnderline
      .WriteProperty "ForeColor", lblCaption.ForeColor
      .WriteProperty "Caption", lblCaption.Caption
      .WriteProperty "UseMnemonic", lblCaption.UseMnemonic
   End With

   Exit Sub

Err_Proc:
   ErrHandler True, Err.Number, Err.Description, "Frame3D", "UserControl_WriteProperties"
   Err.Clear
   Resume Next

End Sub

