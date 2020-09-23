VERSION 5.00
Begin VB.UserControl mhButton 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   DefaultCancel   =   -1  'True
   PropertyPages   =   "MorgansButton.ctx":0000
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "MorgansButton.ctx":0035
   Begin VB.Timer OverTimer 
      Enabled         =   0   'False
      Interval        =   3
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "mhButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'// CHAMELEON BUTTON
'// Orignal By: gonchuki
'// Modified By: Morgan Haueisen

Option Explicit

Private Const C_COLOR_HIGHLIGHT As Long = 13&
Private Const C_COLOR_BTNFACE As Long = 15&
Private Const C_COLOR_BTNSHADOW As Long = 16&
Private Const C_COLOR_BTNTEXT As Long = 18&
Private Const C_COLOR_BTNHIGHLIGHT As Long = 20&
Private Const C_COLOR_BTNDKSHADOW As Long = 21&
Private Const C_COLOR_BTNLIGHT As Long = 22&

Private Const C_DT_CALCRECT As Long = &H400&
Private Const C_DT_WORDBREAK As Long = &H10&
Private Const C_DT_CENTER As Long = &H1& Or C_DT_WORDBREAK Or &H4&

Private Const C_PS_SOLID As Long = 0&
Private Const C_RGN_DIFF As Long = 4&
Private Const C_FXDEPTH As Long = &H28&

Private Type Rect
   left   As Long
   top    As Long
   right  As Long
   bottom As Long
End Type

Private Type POINTAPI
   x As Long
   Y As Long
End Type

Private Type BITMAPINFOHEADER
   biSize          As Long
   biWidth         As Long
   biHeight        As Long
   biPlanes        As Integer
   biBitCount      As Integer
   biCompression   As Long
   biSizeImage     As Long
   biXPelsPerMeter As Long
   biYPelsPerMeter As Long
   biClrUsed       As Long
   biClrImportant  As Long
End Type

Private Type RGBTRIPLE
   rgbBlue  As Byte
   rgbGreen As Byte
   rgbRed   As Byte
End Type

Private Type BITMAPINFO
   bmiHeader As BITMAPINFOHEADER
   bmiColors As RGBTRIPLE
End Type

Private Declare Function GetCurrentThemeName Lib "uxtheme.dll" ( _
      ByVal pszThemeFileName As Long, _
      ByVal dwMaxNameChars As Long, _
      ByVal pszColorBuff As Long, _
      ByVal cchMaxColorChars As Long, _
      ByVal pszSizeBuff As Long, _
      ByVal cchMaxSizeChars As Long) As Long
'Private Declare Function OpenThemeData Lib "uxtheme.dll" (ByVal hwnd As Long, ByVal pszClassList As Long) As Long
'Private Declare Function CloseThemeData Lib "uxtheme.dll" (ByVal hTheme As Long) As Long

Private Declare Function SetPixel Lib "gdi32" Alias "SetPixelV" ( _
      ByVal hDC As Long, _
      ByVal x As Long, _
      ByVal Y As Long, _
      ByVal crColor As Long) As Long
Private Declare Function GetNearestColor Lib "gdi32" ( _
      ByVal hDC As Long, _
      ByVal crColor As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" ( _
      ByVal lOleColor As Long, _
      ByVal lHPalette As Long, _
      lColorRef As Long) As Long
Private Declare Function GetBkColor Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetTextColor Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" ( _
      ByVal hDC As Long, _
      ByVal crColor As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" ( _
      ByVal hDC As Long, _
      ByVal lpStr As String, _
      ByVal nCount As Long, _
      ByRef lpRect As Rect, _
      ByVal wFormat As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" ( _
      ByVal hDC As Long, _
      ByRef lpRect As Rect, _
      ByVal hBrush As Long) As Long
Private Declare Function FrameRect Lib "user32" ( _
      ByVal hDC As Long, _
      ByRef lpRect As Rect, _
      ByVal hBrush As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" ( _
      ByVal hDC As Long, _
      ByRef lpRect As Rect) As Long
Private Declare Function Ellipse Lib "gdi32" ( _
      ByVal hDC As Long, _
      ByVal x1 As Long, _
      ByVal y1 As Long, _
      ByVal x2 As Long, _
      ByVal y2 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" ( _
      ByVal hDC As Long, _
      ByVal hObject As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" ( _
      ByVal hDC As Long, _
      ByVal x As Long, _
      ByVal Y As Long, _
      ByRef lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" ( _
      ByVal hDC As Long, _
      ByVal x As Long, _
      ByVal Y As Long) As Long
Private Declare Function CreatePen Lib "gdi32" ( _
      ByVal nPenStyle As Long, _
      ByVal nWidth As Long, _
      ByVal crColor As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" ( _
      ByVal x1 As Long, _
      ByVal y1 As Long, _
      ByVal x2 As Long, _
      ByVal y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" ( _
      ByVal hDestRgn As Long, _
      ByVal hSrcRgn1 As Long, _
      ByVal hSrcRgn2 As Long, _
      ByVal nCombineMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" ( _
      ByVal HWND As Long, _
      ByVal hRgn As Long, _
      ByVal bRedraw As Long) As Long
Private Declare Function GetClientRect Lib "user32" ( _
      ByVal HWND As Long, _
      ByRef lpRect As Rect) As Long
Private Declare Function InflateRect Lib "user32" ( _
      ByRef lpRect As Rect, _
      ByVal x As Long, _
      ByVal Y As Long) As Long
Private Declare Function OffsetRect Lib "user32" ( _
      ByRef lpRect As Rect, _
      ByVal x As Long, _
      ByVal Y As Long) As Long
Private Declare Function CopyRect Lib "user32" ( _
      ByRef lpDestRect As Rect, _
      ByRef lpSourceRect As Rect) As Long
Private Declare Function WindowFromPoint Lib "user32" ( _
      ByVal xPoint As Long, _
      ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (ByRef lpPoint As POINTAPI) As Long
Private Declare Function SetCapture Lib "user32" (ByVal HWND As Long) As Long
Private Declare Function ReleaseDC Lib "user32" ( _
      ByVal HWND As Long, _
      ByVal hDC As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal HWND As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal HWND As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" ( _
      ByVal aHDC As Long, _
      ByVal hBitmap As Long, _
      ByVal nStartScan As Long, _
      ByVal nNumScans As Long, _
      ByRef lpBits As Any, _
      ByRef lpbi As BITMAPINFO, _
      ByVal wUsage As Long) As Long
Private Declare Function SetDIBitsToDevice Lib "gdi32" ( _
      ByVal hDC As Long, _
      ByVal x As Long, _
      ByVal Y As Long, _
      ByVal dx As Long, _
      ByVal dy As Long, _
      ByVal SrcX As Long, _
      ByVal SrcY As Long, _
      ByVal Scan As Long, _
      ByVal NumScans As Long, _
      ByRef Bits As Any, _
      ByRef BitsInfo As BITMAPINFO, _
      ByVal wUsage As Long) As Long
Private Declare Function BitBlt Lib "gdi32" ( _
      ByVal hDestDC As Long, _
      ByVal x As Long, _
      ByVal Y As Long, _
      ByVal nWidth As Long, _
      ByVal nHeight As Long, _
      ByVal hSrcDC As Long, _
      ByVal xSrc As Long, _
      ByVal ySrc As Long, _
      ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" ( _
      ByVal hDC As Long, _
      ByVal nWidth As Long, _
      ByVal nHeight As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DrawIconEx Lib "user32" ( _
      ByVal hDC As Long, _
      ByVal xLeft As Long, _
      ByVal yTop As Long, _
      ByVal hIcon As Long, _
      ByVal cxWidth As Long, _
      ByVal cyWidth As Long, _
      ByVal istepIfAniCur As Long, _
      ByVal hbrFlickerFreeDraw As Long, _
      ByVal diFlags As Long) As Long
'// THEME BUTTON
Private Declare Function OpenThemeData Lib "uxtheme.dll" ( _
      ByVal HWND As Long, _
      ByVal pszClassList As Long) As Long
Private Declare Function CloseThemeData Lib "uxtheme.dll" (ByVal hTheme As Long) As Long
Private Declare Function DrawThemeBackground Lib "uxtheme.dll" ( _
      ByVal hTheme As Long, _
      ByVal lHDC As Long, _
      ByVal iPartId As Long, _
      ByVal iStateId As Long, _
      ByRef pRect As Rect, _
      ByRef pClipRect As Rect) As Long

Public Enum enuButtonTypes
   [Windows XP] = 3
   [Java metal] = 5
   [Office XP] = 9
   [Office 2003] = 10
   [KDE 2] = 14
End Enum

Public Enum enuColorTypes
   [Use Windows] = 1
   [Custom Colors] = 2
   [Force Standard] = 3
   [Use Container] = 4
End Enum

Public Enum enuPicPositions
   cbLeft = 0
   cbleftleft = 5
   cbRight = 1
   cbTop = 2
   cbBottom = 3
   cbBackground = 4
End Enum

Public Enum enuFX
   cbNone = 0
   cbEmbossed = 1
   cbEngraved = 2
   cbShadowed = 3
End Enum

'// events
Public Event Click()
Attribute Click.VB_UserMemId = -600
Attribute Click.VB_MemberFlags = "200"
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseOver()
Public Event MouseOut()

'// variables
Private mudtButtonType    As enuButtonTypes
Private mudtColorType     As enuColorTypes
Private mudtPicPosition   As enuPicPositions
Private mudtSFX           As enuFX '// font and picture effects

Private mlngHeight        As Long '// the height of the button
Private mlngWidth         As Long '// the width of the button

Private mlngBackC         As Long '// back color
Private mlngBackO         As Long '// back color when mouse is over
Private mlngForeC         As Long '// fore color
Private mlngForeO         As Long '// fore color when mouse is over
Private mlngMaskC         As Long '// mask color
Private mlngOXPb          As Long
Private mlngOXPf          As Long
Private mlngO3D           As Long
Private mlngO3H           As Long
Private mlngO3F           As Long

Private mblnUseMask       As Boolean
Private mblnUseGrey       As Boolean

Private mlngClrFace       As Long
Private mlngClrLight      As Long
Private mlngClrHighLight  As Long
Private mlngClrShadow     As Long
Private mlngClrDarkShadow As Long
Private mlngClrText       As Long
Private mlngClrTextO      As Long
Private mlngClrFaceO      As Long
Private mlngClrMask       As Long
Private mlngXPFace        As Long

Private mPicNormal        As StdPicture
Private mPicHover         As StdPicture

Private mstrCurText       As String '// current text

Private mudtRC            As Rect
Private mudtRC2           As Rect
Private mudtRC3           As Rect
Private mudtFC            As POINTAPI '// text and focus rect locations
Private mudtPicPt         As POINTAPI
Private mudtPicSz         As POINTAPI '// picture Position & Size
Private mlngRgnNorm       As Long

Private mbytLastButton    As Byte
Private mbytLastKeyDown   As Byte
Private mblnIsEnabled     As Boolean
Private mblnIsSoft        As Boolean
Private mblnHasFocus      As Boolean
Private mblnShowFocusR    As Boolean

Private mbytLastStat      As Byte
Private mstrTE            As String
Private mblnIsShown       As Boolean '// used to avoid unnecessary repaints
Private mblnIsOver        As Boolean

Private mlngCaptOpt       As Long
Private mblnIsCheckbox    As Boolean
Private mblncValue        As Boolean

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute BackColor.VB_UserMemId = -501
   
   BackColor = mlngBackC
   
End Property

Public Property Let BackColor(ByVal vNewValue As OLE_COLOR)
   
   mlngBackC = vNewValue
   If Not Ambient.UserMode Then mlngBackO = vNewValue
   Call SetColors
   Call Redraw(mbytLastStat, True)
   PropertyChanged "BCOL"
   
End Property

Public Property Get BackOver() As OLE_COLOR
Attribute BackOver.VB_ProcData.VB_Invoke_Property = ";Appearance"
   
   BackOver = mlngBackO
   
End Property

Public Property Let BackOver(ByVal vNewValue As OLE_COLOR)
   
   mlngBackO = vNewValue
   Call SetColors
   Call Redraw(mbytLastStat, True)
   PropertyChanged "BCOLO"
   
End Property

Public Property Get ButtonType() As enuButtonTypes
Attribute ButtonType.VB_ProcData.VB_Invoke_Property = ";Appearance"
   
   ButtonType = mudtButtonType
   
End Property

Public Property Let ButtonType(ByVal vNewValue As enuButtonTypes)
   
   mudtButtonType = vNewValue
   Call UserControl_Resize
   PropertyChanged "BTYPE"
   
End Property

Private Sub CalcPicSize()
   
   If Not mPicNormal Is Nothing Then
      mudtPicSz.x = UserControl.ScaleX(mPicNormal.Width, 8, UserControl.ScaleMode)
      mudtPicSz.Y = UserControl.ScaleY(mPicNormal.Height, 8, UserControl.ScaleMode)
    Else
      mudtPicSz.x = 0
      mudtPicSz.Y = 0
   End If
   
End Sub

Private Sub CalcTextRects()
   
   '// this sub will calculate the rects required to draw the text
   With mudtRC2
      Select Case mudtPicPosition
       Case cbLeft, cbleftleft
         .left = 1 + mudtPicSz.x
         .right = mlngWidth - 2
         .top = 1
         .bottom = mlngHeight - 2
       Case cbRight
         .left = 1
         .right = mlngWidth - 2 - mudtPicSz.x
         .top = 1
         .bottom = mlngHeight - 2
       Case cbTop
         .left = 1
         .right = mlngWidth - 2
         .top = 1 + mudtPicSz.Y
         .bottom = mlngHeight - 2
       Case cbBottom
         .left = 1
         .right = mlngWidth - 2
         .top = 1
         .bottom = mlngHeight - 2 - mudtPicSz.Y
       Case cbBackground
         .left = 1
         .right = mlngWidth - 2
         .top = 1
         .bottom = mlngHeight - 2
      End Select
   End With '// mudtRC2
   
   Call DrawText(UserControl.hDC, mstrCurText, Len(mstrCurText), mudtRC2, C_DT_CALCRECT Or C_DT_WORDBREAK)
   Call CopyRect(mudtRC, mudtRC2)
   
   mudtFC.x = mudtRC.right - mudtRC.left
   mudtFC.Y = mudtRC.bottom - mudtRC.top
   
   Select Case mudtPicPosition
    Case 0, 2 '//Left, Top
      Call OffsetRect(mudtRC, (mlngWidth - mudtRC.right) \ 2, (mlngHeight - mudtRC.bottom) \ 2)
    Case 1 '// Right
      Call OffsetRect(mudtRC, (mlngWidth - mudtRC.right - mudtPicSz.x - 4) \ 2, (mlngHeight - mudtRC.bottom) \ 2)
    Case 3 '// Bottom
      Call OffsetRect(mudtRC, (mlngWidth - mudtRC.right) \ 2, (mlngHeight - mudtRC.bottom - mudtPicSz.Y - 4) \ 2)
    Case 4 '// Background
      Call OffsetRect(mudtRC, (mlngWidth - mudtRC.right) \ 2, (mlngHeight - mudtRC.bottom) \ 2)
    Case 5
      Call OffsetRect(mudtRC, 10, (mlngHeight - mudtRC.bottom) \ 2)
   End Select
   
   Call CopyRect(mudtRC2, mudtRC)
   Call OffsetRect(mudtRC2, 1, 1)
   
   '// once we have the text position we are able to calculate the pic position
   '// exit if there's no picture
   If mPicNormal Is Nothing And mPicHover Is Nothing Then Exit Sub
   
   '// if there is no caption, or we have the picture as background
   '// then we put the picture at the center of the button
   If (Trim$(mstrCurText) <> vbNullString) And (mudtPicPosition <> 4) Then
      With mudtPicPt
         Select Case mudtPicPosition
          Case 0 '// left
            .x = mudtRC.left - mudtPicSz.x - 4
            .Y = (mlngHeight - mudtPicSz.Y) \ 2
          Case 1 '// right
            .x = mudtRC.right + 4
            .Y = (mlngHeight - mudtPicSz.Y) \ 2
          Case 2 '// top
            .x = (mlngWidth - mudtPicSz.x) \ 2
            .Y = mudtRC.top - mudtPicSz.Y - 2
          Case 3 '// bottom
            .x = (mlngWidth - mudtPicSz.x) \ 2
            .Y = mudtRC.bottom + 2
          Case 5
            .x = 5
            .Y = (mlngHeight - mudtPicSz.Y) \ 2
         End Select
      End With '// mudtPicPt
    
    Else '// center the picture
      mudtPicPt.x = (mlngWidth - mudtPicSz.x) \ 2
      mudtPicPt.Y = (mlngHeight - mudtPicSz.Y) \ 2
   End If
   
End Sub

Public Property Get Caption() As String
Attribute Caption.VB_ProcData.VB_Invoke_Property = ";Text"
Attribute Caption.VB_UserMemId = -518
   
   Caption = mstrCurText
   
End Property

Public Property Let Caption(ByVal vNewValue As String)
   
   mstrCurText = vNewValue
   Call SetAccessKeys
   Call CalcTextRects
   Call Redraw(0, True)
   PropertyChanged "TX"
   
End Property

Public Property Get CheckBoxBehaviour() As Boolean
   
   CheckBoxBehaviour = mblnIsCheckbox
   
End Property

Public Property Let CheckBoxBehaviour(ByVal vNewValue As Boolean)
   
   mblnIsCheckbox = vNewValue
   Call Redraw(mbytLastStat, True)
   PropertyChanged "CHECK"
   
End Property

Public Property Get ColorScheme() As enuColorTypes
Attribute ColorScheme.VB_ProcData.VB_Invoke_Property = ";Appearance"
   
   ColorScheme = mudtColorType
   
End Property

Public Property Let ColorScheme(ByVal vNewValue As enuColorTypes)
   
   mudtColorType = vNewValue
   Call SetColors
   Call Redraw(0, True)
   PropertyChanged "COLTYPE"
   
End Property

Private Function ConvertFromSystemColor(ByVal vColor As Long) As Long
   
   Call OleTranslateColor(vColor, 0, ConvertFromSystemColor)
   
End Function

Private Sub DoFX(ByVal vOffset As Long, ByVal vStdPic As StdPicture)
  
  Dim lngCurFace As Long
   
   If mudtSFX > cbNone Then
      
      If mudtButtonType = [Windows XP] Then
         lngCurFace = mlngXPFace
       Else
         If vOffset = -1 And mudtColorType <> [Custom Colors] Then
            lngCurFace = mlngOXPf
          Else
            lngCurFace = mlngClrFace
         End If
      End If
      
      Call TransBlt(UserControl.hDC, mudtPicPt.x + 1 + vOffset, mudtPicPt.Y + 1 + vOffset, mudtPicSz.x, mudtPicSz.Y, vStdPic, mlngClrMask, ShiftColor(lngCurFace, Abs(mudtSFX = cbEngraved) * C_FXDEPTH + (mudtSFX <> cbEngraved) * C_FXDEPTH))
      
      If mudtSFX < cbShadowed Then
         Call TransBlt(UserControl.hDC, mudtPicPt.x - 1 + vOffset, mudtPicPt.Y - 1 + vOffset, mudtPicSz.x, mudtPicSz.Y, vStdPic, mlngClrMask, ShiftColor(lngCurFace, Abs(mudtSFX <> cbEngraved) * C_FXDEPTH + (mudtSFX = cbEngraved) * C_FXDEPTH))
      End If
      
   End If
   
End Sub

Private Sub DrawCaption(ByVal vbytState As Byte)
      
   mlngCaptOpt = vbytState
   
   With UserControl
      '// in this select case, we only change the text color and draw only text that needs mudtRC2
      '// at the end, text that uses mudtRC will be drawn
      Select Case vbytState
       Case 0 '// normal caption
         Call txtFX(mudtRC)
         Call SetTextColor(.hDC, mlngClrText)
         
       Case 1 '// hover caption
         Call txtFX(mudtRC)
         Call SetTextColor(.hDC, mlngClrTextO)
         
       Case 2 '// down caption
         Call txtFX(mudtRC2)
         Call SetTextColor(.hDC, mlngClrTextO)
         Call DrawText(.hDC, mstrCurText, Len(mstrCurText), mudtRC2, C_DT_CENTER)
         
       Case 3 '// disabled embossed caption
         Call SetTextColor(.hDC, mlngClrHighLight)
         Call DrawText(.hDC, mstrCurText, Len(mstrCurText), mudtRC2, C_DT_CENTER)
         Call SetTextColor(.hDC, mlngClrShadow)
         
       Case 4 '// disabled grey caption
         Call SetTextColor(.hDC, mlngClrShadow)
         
       Case 5 '// WinXP disabled caption
         Call SetTextColor(.hDC, ShiftColor(mlngXPFace, -&H68, True))
         
       Case 6 '// KDE 2 disabled
         Call SetTextColor(.hDC, mlngClrHighLight)
         Call DrawText(.hDC, mstrCurText, Len(mstrCurText), mudtRC2, C_DT_CENTER)
         Call SetTextColor(.hDC, mlngClrFace)
         
       Case 7 '// KDE 2 down
         Call SetTextColor(.hDC, ShiftColor(mlngClrShadow, -&H32))
         Call DrawText(.hDC, mstrCurText, Len(mstrCurText), mudtRC2, C_DT_CENTER)
         Call SetTextColor(.hDC, mlngClrHighLight)
      End Select
      
      '// we now draw the text that is common in all the captions
      If vbytState <> 2 Then Call DrawText(.hDC, mstrCurText, Len(mstrCurText), mudtRC, C_DT_CENTER)
   End With
   
End Sub

Private Sub DrawFocusR()
   
   If mblnShowFocusR And mblnHasFocus Then
      Call SetTextColor(UserControl.hDC, mlngClrText)
      Call DrawFocusRect(UserControl.hDC, mudtRC3)
   End If
   
End Sub

Private Sub DrawLine(ByVal vX1 As Long, _
                     ByVal vY1 As Long, _
                     ByVal vX2 As Long, _
                     ByVal vY2 As Long, _
                     ByVal vColor As Long)
   
   '// a fast way to draw lines
   
  Dim udtPT     As POINTAPI
  Dim lngoldPen As Long
  Dim lngHPen   As Long
   
   With UserControl
      lngHPen = CreatePen(C_PS_SOLID, 1, vColor)
      lngoldPen = SelectObject(.hDC, lngHPen)
      
      MoveToEx .hDC, vX1, vY1, udtPT
      LineTo .hDC, vX2, vY2
      
      SelectObject .hDC, lngoldPen
      DeleteObject lngHPen
   End With
   
End Sub

Private Sub DrawPictures(ByVal vbytState As Byte)
   
   '// check if there is a main picture, if not then exit
   If mPicNormal Is Nothing Then Exit Sub
   
   With UserControl
      Select Case vbytState
       Case 0 '// normal & hover
         If Not mblnIsOver Then
            Call DoFX(0, mPicNormal)
            Call TransBlt(.hDC, mudtPicPt.x, mudtPicPt.Y, mudtPicSz.x, mudtPicSz.Y, mPicNormal, mlngClrMask, , , mblnUseGrey, (mudtButtonType = [Office XP]))
         
         Else
            If mudtButtonType = [Office XP] Or mudtButtonType = [Office 2003] Then
               Call DoFX(-1, mPicNormal)
               Call TransBlt(.hDC, mudtPicPt.x + 1, mudtPicPt.Y + 1, mudtPicSz.x, mudtPicSz.Y, mPicNormal, mlngClrMask, mlngClrShadow)
               Call TransBlt(.hDC, mudtPicPt.x - 1, mudtPicPt.Y - 1, mudtPicSz.x, mudtPicSz.Y, mPicNormal, mlngClrMask)
               
            Else
               If Not mPicHover Is Nothing Then
                  Call DoFX(0, mPicHover)
                  Call TransBlt(.hDC, mudtPicPt.x, mudtPicPt.Y, mudtPicSz.x, mudtPicSz.Y, mPicHover, mlngClrMask)
               Else
                  Call DoFX(0, mPicNormal)
                  Call TransBlt(.hDC, mudtPicPt.x, mudtPicPt.Y, mudtPicSz.x, mudtPicSz.Y, mPicNormal, mlngClrMask)
               End If
            End If
         End If
         
       Case 1 '// down
         If mPicHover Is Nothing Or mudtButtonType = [Office XP] Or mudtButtonType = [Office 2003] Then
            Select Case mudtButtonType
             Case [Java metal], [Office XP], [Office 2003]
               Call DoFX(0, mPicNormal)
               Call TransBlt(.hDC, mudtPicPt.x, mudtPicPt.Y, mudtPicSz.x, mudtPicSz.Y, mPicNormal, mlngClrMask)
             Case Else
               Call DoFX(1, mPicNormal)
               Call TransBlt(.hDC, mudtPicPt.x + 1, mudtPicPt.Y + 1, mudtPicSz.x, mudtPicSz.Y, mPicNormal, mlngClrMask)
            End Select
          Else
            Call TransBlt(.hDC, mudtPicPt.x + Abs(mudtButtonType <> [Java metal]), mudtPicPt.Y + Abs(mudtButtonType <> [Java metal]), mudtPicSz.x, mudtPicSz.Y, mPicHover, mlngClrMask)
         End If
         
       Case 2 '// disabled
         Select Case mudtButtonType
          Case [Java metal], [Office XP], [Office 2003]
            If mudtButtonType = [Office 2003] Then
               Call TransBlt(.hDC, mudtPicPt.x, mudtPicPt.Y, mudtPicSz.x, mudtPicSz.Y, mPicNormal, mlngClrMask, Abs(mudtButtonType = [Office 2003]) * ShiftColor(mlngClrShadow, &HD) + Abs(mudtButtonType <> [Office 2003]) * mlngClrShadow, True)
            Else
               Call TransBlt(.hDC, mudtPicPt.x, mudtPicPt.Y, mudtPicSz.x, mudtPicSz.Y, mPicNormal, mlngClrMask, Abs(mudtButtonType = [Office XP]) * ShiftColor(mlngClrShadow, &HD) + Abs(mudtButtonType <> [Office XP]) * mlngClrShadow, True)
            End If
            
          Case [Windows XP] '// for WinXP draw a greyscaled image
            Call TransBlt(.hDC, mudtPicPt.x + 1, mudtPicPt.Y + 1, mudtPicSz.x, mudtPicSz.Y, mPicNormal, mlngClrMask, , , True)
            
          Case Else '// draw classic embossed pictures
            Call TransBlt(.hDC, mudtPicPt.x + 1, mudtPicPt.Y + 1, mudtPicSz.x, mudtPicSz.Y, mPicNormal, mlngClrMask, mlngClrHighLight, True)
            Call TransBlt(.hDC, mudtPicPt.x, mudtPicPt.Y, mudtPicSz.x, mudtPicSz.Y, mPicNormal, mlngClrMask, mlngClrShadow, True)
            
         End Select
      End Select
      
   End With '// UserControl
   
   If mudtPicPosition = cbBackground Then Call DrawCaption(mlngCaptOpt)
   
End Sub

Private Sub DrawRectangle(ByVal vlngX As Long, _
                          ByVal vlngY As Long, _
                          ByVal vWidth As Long, _
                          ByVal vHeight As Long, _
                          ByVal vlngColor As Long, _
                          Optional ByVal vblnOnlyBorder As Boolean = False)
   
   '// this is my custom function to draw rectangles and frames
   '// it's faster and smoother than using the line method
   
  Dim udtRECT As Rect
  Dim lngBrush As Long
   
   udtRECT.left = vlngX
   udtRECT.top = vlngY
   udtRECT.right = vlngX + vWidth
   udtRECT.bottom = vlngY + vHeight
   
   lngBrush = CreateSolidBrush(vlngColor)
   
   If vblnOnlyBorder Then
      Call FrameRect(UserControl.hDC, udtRECT, lngBrush)
    Else
      Call FillRect(UserControl.hDC, udtRECT, lngBrush)
   End If
   
   Call DeleteObject(lngBrush)
   
End Sub

Private Function DrawXPThemeButton(ByVal vButtonState As Long) As Boolean
   
  Dim lngTheme   As Long
  Dim blnThemeOk As Boolean
  Dim btnRect    As Rect
   
   If Not mblnIsSoft And mudtColorType = [Use Windows] Then
      On Error Resume Next
      
      btnRect.left = 0
      btnRect.top = 0
      btnRect.right = UserControl.ScaleWidth
      btnRect.bottom = UserControl.ScaleHeight
      
      If vButtonState = 1 Then
         If mblnIsOver Then
            vButtonState = 2
          ElseIf ((mblnHasFocus Or Ambient.DisplayAsDefault) And mblnShowFocusR) Then
            vButtonState = 5
         End If
      End If
      
      lngTheme = OpenThemeData(UserControl.HWND, StrPtr("Button"))
      If lngTheme Then
         blnThemeOk = True
         Call DrawThemeBackground(lngTheme, UserControl.hDC, 1, vButtonState, btnRect, btnRect)
         Call CloseThemeData(lngTheme)
      End If
      
   End If
   
   DrawXPThemeButton = blnThemeOk
   
End Function

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_ProcData.VB_Invoke_Property = ";Behavior"
Attribute Enabled.VB_UserMemId = -514
   
   Enabled = mblnIsEnabled
   
End Property

Public Property Let Enabled(ByVal vNewValue As Boolean)
   
   mblnIsEnabled = vNewValue
   Call Redraw(0, True)
   UserControl.Enabled = mblnIsEnabled
   PropertyChanged "ENAB"
   
End Property

Public Property Get Font() As Font
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute Font.VB_UserMemId = -512
   
   Set Font = UserControl.Font
   
End Property

Public Property Set Font(ByRef rNewValue As Font)
   
   Set UserControl.Font = rNewValue
   Call CalcTextRects
   Call Redraw(0, True)
   PropertyChanged "FONT"
   
End Property

Public Property Get FontBold() As Boolean
Attribute FontBold.VB_MemberFlags = "400"
   
   FontBold = UserControl.FontBold
   
End Property

Public Property Let FontBold(ByVal vNewValue As Boolean)
   
   UserControl.FontBold = vNewValue
   Call CalcTextRects
   Call Redraw(0, True)
   
End Property

Public Property Get FontItalic() As Boolean
Attribute FontItalic.VB_MemberFlags = "400"
   
   FontItalic = UserControl.FontItalic
   
End Property

Public Property Let FontItalic(ByVal vNewValue As Boolean)
   
   UserControl.FontItalic = vNewValue
   Call CalcTextRects
   Call Redraw(0, True)
   
End Property

Public Property Get FontName() As String
Attribute FontName.VB_MemberFlags = "400"
   
   FontName = UserControl.FontName
   
End Property

Public Property Let FontName(ByVal vNewValue As String)
   
   UserControl.FontName = vNewValue
   Call CalcTextRects
   Call Redraw(0, True)
   
End Property

Public Property Get FontSize() As Integer
Attribute FontSize.VB_MemberFlags = "400"
   
   FontSize = UserControl.FontSize
   
End Property

Public Property Let FontSize(ByVal vNewValue As Integer)
   
   UserControl.FontSize = vNewValue
   Call CalcTextRects
   Call Redraw(0, True)
   
End Property

Public Property Get FontUnderline() As Boolean
Attribute FontUnderline.VB_MemberFlags = "400"
   
   FontUnderline = UserControl.FontUnderline
   
End Property

Public Property Let FontUnderline(ByVal vNewValue As Boolean)
   
   UserControl.FontUnderline = vNewValue
   Call CalcTextRects
   Call Redraw(0, True)
   
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute ForeColor.VB_UserMemId = -513
   
   ForeColor = mlngForeC
   
End Property

Public Property Let ForeColor(ByVal vNewValue As OLE_COLOR)
   
   mlngForeC = vNewValue
   If Not Ambient.UserMode Then mlngForeO = vNewValue
   Call SetColors
   Call Redraw(mbytLastStat, True)
   PropertyChanged "FCOL"
   
End Property

Public Property Get ForeOver() As OLE_COLOR
Attribute ForeOver.VB_ProcData.VB_Invoke_Property = ";Appearance"
   
   ForeOver = mlngForeO
   
End Property

Public Property Let ForeOver(ByVal vNewValue As OLE_COLOR)
   
   mlngForeO = vNewValue
   Call SetColors
   Call Redraw(mbytLastStat, True)
   PropertyChanged "FCOLO"
   
End Property

Public Property Get HWND() As Long
Attribute HWND.VB_UserMemId = -515
   
   HWND = UserControl.HWND
   
End Property

Private Function isMouseOver() As Boolean
   
  Dim udtPT As POINTAPI
   
   GetCursorPos udtPT
   isMouseOver = (WindowFromPoint(udtPT.x, udtPT.Y) = HWND)
   
End Function

Public Property Get MaskColor() As OLE_COLOR
Attribute MaskColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
   
   MaskColor = mlngMaskC
   
End Property

Public Property Let MaskColor(ByVal vNewValue As OLE_COLOR)
   
   mlngMaskC = vNewValue
   Call SetColors
   Call Redraw(mbytLastStat, True)
   PropertyChanged "MCOL"
   
End Property

Public Property Get MouseIcon() As StdPicture
Attribute MouseIcon.VB_ProcData.VB_Invoke_Property = ";Appearance"
   
   Set MouseIcon = UserControl.MouseIcon
   
End Property

Public Property Set MouseIcon(ByVal vNewValue As StdPicture)
   
   On Local Error Resume Next
   Set UserControl.MouseIcon = vNewValue
   PropertyChanged "MICON"
   On Error GoTo 0
   
End Property

Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_ProcData.VB_Invoke_Property = ";Appearance"
   
   MousePointer = UserControl.MousePointer
   
End Property

Public Property Let MousePointer(ByVal vNewValue As MousePointerConstants)
   
   UserControl.MousePointer = vNewValue
   PropertyChanged "MPTR"
   
End Property

Private Sub OverTimer_Timer()
   
   If Not isMouseOver Then
      OverTimer.Enabled = False
      mblnIsOver = False
      Call Redraw(0, True)
      RaiseEvent MouseOut
   End If
   
End Sub

Public Property Get PictureNormal() As StdPicture
Attribute PictureNormal.VB_ProcData.VB_Invoke_Property = ";Appearance"
   
   Set PictureNormal = mPicNormal
   
End Property

Public Property Set PictureNormal(ByVal vNewValue As StdPicture)
   
   Set mPicNormal = vNewValue
   Call CalcPicSize
   Call CalcTextRects
   Call Redraw(mbytLastStat, True)
   PropertyChanged "PICN"
   
End Property

Public Property Get PictureOver() As StdPicture
Attribute PictureOver.VB_ProcData.VB_Invoke_Property = ";Appearance"
   
   Set PictureOver = mPicHover
   
End Property

Public Property Set PictureOver(ByVal vNewValue As StdPicture)
   
   Set mPicHover = vNewValue
   '// only redraw if we need to see this picture immediately
   If mblnIsOver Then Call Redraw(mbytLastStat, True)
   PropertyChanged "PICO"
   
End Property

Public Property Get PicturePosition() As enuPicPositions
Attribute PicturePosition.VB_ProcData.VB_Invoke_Property = ";Position"
   
   PicturePosition = mudtPicPosition
   
End Property

Public Property Let PicturePosition(ByVal vNewValue As enuPicPositions)
   
   mudtPicPosition = vNewValue
   PropertyChanged "PICPOS"
   Call CalcTextRects
   Call Redraw(mbytLastStat, True)
   
End Property

Private Sub Redraw(ByVal vCurStat As Byte, ByVal vblnForce As Boolean)
   
  Dim lngI      As Long
  Dim sngStepXP As Single
  Dim lngXPFace As Long
  Dim lngColor  As Long
  
   '// here is the CORE of the button, everything is drawn here
   
   If mblnIsCheckbox And mblncValue Then vCurStat = 2
   
   If Not vblnForce Then  '// check drawing redundancy
      If (vCurStat = mbytLastStat) And (mstrTE = mstrCurText) Then GoTo Exit_Proc
   End If
   
   If mlngHeight = 0 Or Not mblnIsShown Then GoTo Exit_Proc '// we don't want errors
   
   mbytLastStat = vCurStat
   mstrTE = mstrCurText
   
   With UserControl
      .Cls
      If mblnIsOver And mudtColorType = [Custom Colors] Then
         lngColor = mlngBackC
         mlngBackC = mlngBackO
         Call SetColors
      End If
      
      Call DrawRectangle(0, 0, mlngWidth, mlngHeight, mlngClrFace)
      
      If mblnIsEnabled Then
         If vCurStat = 0 Then '// Button Normal State
            Select Case mudtButtonType
             Case [Java metal] '// Java
               Call DrawRectangle(1, 1, mlngWidth - 1, mlngHeight - 1, ShiftColor(mlngClrFace, &HC))
               If mblnIsOver Then Call DrawRectangle(1, 1, mlngWidth - 4, mlngHeight - 4, ShiftColor(mlngOXPf, -&HA))
               Call DrawRectangle(1, 1, mlngWidth - 1, mlngHeight - 1, mlngClrHighLight, True)
               Call DrawRectangle(0, 0, mlngWidth - 1, mlngHeight - 1, ShiftColor(mlngClrShadow, -&H1A), True)
               Call DrawCaption(Abs(mblnIsOver))
               If mblnHasFocus And mblnShowFocusR Then Call DrawFocusR
               
             Case [Office XP] '// Office XP
               If mblnIsOver Then Call DrawRectangle(1, 1, mlngWidth, mlngHeight, mlngOXPf)
               Call DrawCaption(Abs(mblnIsOver))
               If mblnIsOver Then Call DrawRectangle(0, 0, mlngWidth, mlngHeight, mlngOXPb, True)
               Call DrawFocusR
               
             Case [Office 2003] '// Office 2003
               If Not mblnIsOver Then
                  Call DrawGradiant(vbWhite, mlngO3F)
                Else
                  Call DrawGradiant(vbWhite, mlngO3H)
                  Call DrawRectangle(0, 0, mlngWidth, mlngHeight, vbBlack, True)
               End If
               Call DrawCaption(Abs(mblnIsOver))
               Call DrawFocusR
               
             Case [KDE 2] '// KDE 2
               If Not mblnIsOver Then
                  sngStepXP = 58 / mlngHeight
                  For lngI = 1 To mlngHeight
                     Call DrawLine(0, lngI, mlngWidth, lngI, ShiftColor(mlngClrHighLight, -sngStepXP * lngI))
                  Next lngI
                Else
                  Call DrawRectangle(0, 0, mlngWidth, mlngHeight, mlngClrLight)
               End If
               If Ambient.DisplayAsDefault Then
                  mblnIsShown = False
               End If
               Call DrawCaption(Abs(mblnIsOver))
               If Ambient.DisplayAsDefault Then
                  mblnIsShown = True
               End If
               Call DrawRectangle(0, 0, mlngWidth, mlngHeight, ShiftColor(mlngClrShadow, -&H32), True)
               Call DrawRectangle(1, 1, mlngWidth - 2, mlngHeight - 2, ShiftColor(mlngClrFace, -&H9), True)
               Call DrawRectangle(2, 2, mlngWidth - 4, 2, mlngClrHighLight)
               Call DrawRectangle(2, 4, 2, mlngHeight - 6, mlngClrHighLight)
               Call DrawFocusR
             
             Case Else '// XP Button
               If Not DrawXPThemeButton(1) Then
                  '// draw back color
                  sngStepXP = 25 / mlngHeight
                  For lngI = 1 To mlngHeight
                     Call DrawLine(0, lngI, mlngWidth, lngI, ShiftColor(mlngXPFace, -sngStepXP * lngI, True))
                  Next lngI
                  
                  Call DrawCaption(Abs(mblnIsOver))
                  
                  Call DrawRectangle(0, 0, mlngWidth, mlngHeight, &H733C00, True)
                  '// fill in the 4 corners
                  SetPixel .hDC, 1, 1, &H7B4D10
                  SetPixel .hDC, 1, mlngHeight - 2, &H7B4D10
                  SetPixel .hDC, mlngWidth - 2, 1, &H7B4D10
                  SetPixel .hDC, mlngWidth - 2, mlngHeight - 2, &H7B4D10
                  
                  If mblnIsOver Then
                     Call DrawRectangle(1, 2, mlngWidth - 2, mlngHeight - 4, &H31B2FF, True)
                     Call DrawLine(2, mlngHeight - 2, mlngWidth - 2, mlngHeight - 2, &H96E7&)
                     Call DrawLine(2, 1, mlngWidth - 2, 1, &HCEF3FF)
                     Call DrawLine(1, 2, mlngWidth - 1, 2, &H8CDBFF)
                     Call DrawLine(2, 3, 2, mlngHeight - 3, &H6BCBFF)
                     Call DrawLine(mlngWidth - 3, 3, mlngWidth - 3, mlngHeight - 3, &H6BCBFF)
                   ElseIf ((mblnHasFocus Or Ambient.DisplayAsDefault) And mblnShowFocusR) Then
                     Call DrawRectangle(1, 2, mlngWidth - 2, mlngHeight - 4, &HE7AE8C, True)
                     Call DrawLine(2, mlngHeight - 2, mlngWidth - 2, mlngHeight - 2, &HEF826B)
                     Call DrawLine(2, 1, mlngWidth - 2, 1, &HFFE7CE)
                     Call DrawLine(1, 2, mlngWidth - 1, 2, &HF7D7BD)
                     Call DrawLine(2, 3, 2, mlngHeight - 3, &HF0D1B5)
                     Call DrawLine(mlngWidth - 3, 3, mlngWidth - 3, mlngHeight - 3, &HF0D1B5)
                   Else '// we do not always draw the bevel because the above code would repaint over it
                     Call DrawLine(2, mlngHeight - 2, mlngWidth - 2, mlngHeight - 2, ShiftColor(mlngXPFace, -&H30, True))
                     Call DrawLine(1, mlngHeight - 3, mlngWidth - 2, mlngHeight - 3, ShiftColor(mlngXPFace, -&H20, True))
                     Call DrawLine(mlngWidth - 2, 2, mlngWidth - 2, mlngHeight - 2, ShiftColor(mlngXPFace, -&H24, True))
                     Call DrawLine(mlngWidth - 3, 3, mlngWidth - 3, mlngHeight - 3, ShiftColor(mlngXPFace, -&H18, True))
                     Call DrawLine(2, 1, mlngWidth - 2, 1, ShiftColor(mlngXPFace, &H10, True))
                     Call DrawLine(1, 2, mlngWidth - 2, 2, ShiftColor(mlngXPFace, &HA, True))
                     Call DrawLine(1, 2, 1, mlngHeight - 2, ShiftColor(mlngXPFace, -&H5, True))
                     Call DrawLine(2, 3, 2, mlngHeight - 3, ShiftColor(mlngXPFace, -&HA, True))
                  End If
                Else
                  Call DrawCaption(Abs(mblnIsOver))
               End If
            End Select
            
            Call DrawPictures(0)
            
          ElseIf vCurStat = 2 Then '// BUTTON IS DOWN
            Select Case mudtButtonType
             Case [Java metal] '// Java
               Call DrawRectangle(1, 1, mlngWidth - 2, mlngHeight - 2, ShiftColor(mlngClrShadow, &H10), False)
               Call DrawRectangle(0, 0, mlngWidth - 1, mlngHeight - 1, ShiftColor(mlngClrShadow, -&H1A), True)
               Call DrawLine(mlngWidth - 1, 1, mlngWidth - 1, mlngHeight, mlngClrHighLight)
               Call DrawLine(1, mlngHeight - 1, mlngWidth - 1, mlngHeight - 1, mlngClrHighLight)
               Call DrawCaption(2)
               If mblnHasFocus And mblnShowFocusR Then Call DrawFocusR
               
             Case [Office XP] '// Office XP
               If mblnIsOver Then
                  Call DrawRectangle(0, 0, mlngWidth, mlngHeight, Abs(mudtColorType = 2) * ShiftColor(mlngOXPf, -&H20) + Abs(mudtColorType <> 2) * ShiftColorOXP(mlngOXPb, &H80))
               End If
               Call DrawCaption(2)
               Call DrawRectangle(0, 0, mlngWidth, mlngHeight, mlngOXPb, True)
               Call DrawFocusR
               
             Case [Office 2003] '// Office 2003
               Call DrawGradiant(mlngO3D, mlngO3H)
               Call DrawCaption(2)
               Call DrawRectangle(0, 0, mlngWidth, mlngHeight, vbBlack, True)
               Call DrawFocusR
               
             Case [KDE 2] '// KDE 2
               Call DrawRectangle(1, 1, mlngWidth, mlngHeight, ShiftColor(mlngClrFace, -&H9))
               Call DrawRectangle(0, 0, mlngWidth, mlngHeight, ShiftColor(mlngClrShadow, -&H30), True)
               Call DrawLine(2, mlngHeight - 2, mlngWidth - 2, mlngHeight - 2, mlngClrHighLight)
               Call DrawLine(mlngWidth - 2, 2, mlngWidth - 2, mlngHeight - 1, mlngClrHighLight)
               Call DrawCaption(7)
               Call DrawFocusR
            
             Case Else '// XP Button
               If Not DrawXPThemeButton(3) Then
                  '// draw background
                  sngStepXP = 25 / mlngHeight
                  lngXPFace = ShiftColor(mlngXPFace, -32, True)
                  For lngI = 1 To mlngHeight
                     Call DrawLine(0, mlngHeight - lngI, mlngWidth, mlngHeight - lngI, ShiftColor(lngXPFace, -sngStepXP * lngI, True))
                  Next lngI
                  
                  Call DrawCaption(2)
                  
                  Call DrawRectangle(0, 0, mlngWidth, mlngHeight, &H733C00, True)
                  '// fill in the corners
                  SetPixel .hDC, 1, 1, &H7B4D10
                  SetPixel .hDC, 1, mlngHeight - 2, &H7B4D10
                  SetPixel .hDC, mlngWidth - 2, 1, &H7B4D10
                  SetPixel .hDC, mlngWidth - 2, mlngHeight - 2, &H7B4D10
                  
                  Call DrawLine(2, mlngHeight - 2, mlngWidth - 2, mlngHeight - 2, ShiftColor(lngXPFace, &H10, True))
                  Call DrawLine(1, mlngHeight - 3, mlngWidth - 2, mlngHeight - 3, ShiftColor(lngXPFace, &HA, True))
                  Call DrawLine(mlngWidth - 2, 2, mlngWidth - 2, mlngHeight - 2, ShiftColor(lngXPFace, &H5, True))
                  Call DrawLine(mlngWidth - 3, 3, mlngWidth - 3, mlngHeight - 3, mlngXPFace)
                  Call DrawLine(2, 1, mlngWidth - 2, 1, ShiftColor(lngXPFace, -&H20, True))
                  Call DrawLine(1, 2, mlngWidth - 2, 2, ShiftColor(lngXPFace, -&H18, True))
                  Call DrawLine(1, 2, 1, mlngHeight - 2, ShiftColor(lngXPFace, -&H20, True))
                  Call DrawLine(2, 2, 2, mlngHeight - 2, ShiftColor(lngXPFace, -&H16, True))
                Else
                  Call DrawCaption(2)
               End If
            End Select
            
            Call DrawPictures(1)
            
         End If
         
       Else '// Button Disabled
         Select Case mudtButtonType
          Case [Java metal] '// Java
            Call DrawCaption(4)
            Call DrawRectangle(0, 0, mlngWidth, mlngHeight, mlngClrShadow, True)
            
          Case [Office XP] '// Office XP
            Call DrawCaption(4)
            
          Case [Office 2003]
            Call DrawGradiant(vbWhite, mlngO3F)
            Call DrawCaption(4)
            
          Case [KDE 2] '// KDE 2
            sngStepXP = 58 / mlngHeight
            For lngI = 1 To mlngHeight
               Call DrawLine(0, lngI, mlngWidth, lngI, ShiftColor(mlngClrHighLight, -sngStepXP * lngI))
            Next lngI
            Call DrawRectangle(0, 0, mlngWidth, mlngHeight, ShiftColor(mlngClrShadow, -&H32), True)
            Call DrawRectangle(1, 1, mlngWidth - 2, mlngHeight - 2, ShiftColor(mlngClrFace, -&H9), True)
            Call DrawRectangle(2, 2, mlngWidth - 4, 2, mlngClrHighLight)
            Call DrawRectangle(2, 4, 2, mlngHeight - 6, mlngClrHighLight)
            Call DrawCaption(6)
         
          Case Else '// XP Button
            If Not DrawXPThemeButton(4) Then
               Call DrawRectangle(0, 0, mlngWidth, mlngHeight, ShiftColor(mlngXPFace, -&H18, True))
               Call DrawCaption(5)
               Call DrawRectangle(0, 0, mlngWidth, mlngHeight, ShiftColor(mlngXPFace, -&H54, True), True)
               SetPixel .hDC, 1, 1, ShiftColor(mlngXPFace, -&H48, True)
               SetPixel .hDC, 1, mlngHeight - 2, ShiftColor(mlngXPFace, -&H48, True)
               SetPixel .hDC, mlngWidth - 2, 1, ShiftColor(mlngXPFace, -&H48, True)
               SetPixel .hDC, mlngWidth - 2, mlngHeight - 2, ShiftColor(mlngXPFace, -&H48, True)
             Else
               Call DrawCaption(5)
            End If
          
         End Select
         Call DrawPictures(2)
      End If
      
   End With '// Usercontrol
   
   If mblnIsOver And mudtColorType = [Custom Colors] Then
      mlngBackC = lngColor
      Call SetColors
   End If

Exit_Proc:
   
End Sub

Private Sub SetAccessKeys()
   
   '// this is a TRUE access keys parser
   '// the basic rule is that if an ampersand is followed by another,
   '//   a single ampersand is drawn and this is not the access key.
   '//   So we continue searching for another possible access key.
   '//   I only do a second pass because no one writes text like "Me & them & everyone"
   '//   so the caption prop should be "Me && them && &everyone", this is rubbish and a
   '//   search like this would only waste time
   
  Dim lngAmpersandPos As Long
   
   '// we first clear the AccessKeys property, and will be filled if one is found
   UserControl.AccessKeys = vbNullString
   
   If Len(mstrCurText) > 1 Then
      lngAmpersandPos = InStr(1, mstrCurText, "&", vbTextCompare)
      If (lngAmpersandPos < Len(mstrCurText)) And (lngAmpersandPos > 0) Then
         '// if text is sonething like && then no access key should be assigned, so continue searching
         If Mid$(mstrCurText, lngAmpersandPos + 1, 1) <> "&" Then
            UserControl.AccessKeys = LCase$(Mid$(mstrCurText, lngAmpersandPos + 1, 1))
          Else '// do only a second pass to find another ampersand character
            lngAmpersandPos = InStr(lngAmpersandPos + 2, mstrCurText, "&", vbTextCompare)
            If Mid$(mstrCurText, lngAmpersandPos + 1, 1) <> "&" Then
               UserControl.AccessKeys = LCase$(Mid$(mstrCurText, lngAmpersandPos + 1, 1))
            End If
         End If
      End If
   End If
   
End Sub

Private Sub SetColors()
   
   '// this function sets the colors taken as a base to build
   '// all the other colors and styles.
   
   Select Case mudtColorType
    Case [Custom Colors]
      mlngClrFace = ConvertFromSystemColor(mlngBackC)
      mlngClrFaceO = ConvertFromSystemColor(mlngBackO)
      mlngClrText = ConvertFromSystemColor(mlngForeC)
      mlngClrTextO = ConvertFromSystemColor(mlngForeO)
      mlngClrShadow = ShiftColor(mlngClrFace, -&H40)
      mlngClrLight = ShiftColor(mlngClrFace, &H1F)
      mlngClrHighLight = ShiftColor(mlngClrFace, &H2F) '// it should be 3F but it looks too lighter
      mlngClrDarkShadow = ShiftColor(mlngClrFace, -&HC0)
      mlngOXPb = ShiftColor(mlngClrFace, -&H80)
      mlngOXPf = mlngClrFace
      mlngO3D = &H4E91FE 'Down
      mlngO3H = &H8BCFFF 'hover
      mlngO3F = mlngClrFace ' face
      
    Case [Force Standard]
      mlngClrFace = &HC0C0C0
      mlngClrFaceO = mlngClrFace
      mlngClrShadow = &H808080
      mlngClrLight = &HDFDFDF
      mlngClrDarkShadow = &H0
      mlngClrHighLight = &HFFFFFF
      mlngClrText = &H0
      mlngClrTextO = mlngClrText
      mlngOXPb = &H800000
      mlngOXPf = &HD1ADAD
      mlngO3D = &H4E91FE 'Down
      mlngO3H = &H8BCFFF 'hover
      mlngO3F = &HBA9EA0 ' face
      
    Case [Use Container]
      mlngClrFace = GetBkColor(GetDC(GetParent(HWND)))
      mlngClrFaceO = mlngClrFace
      mlngClrText = GetTextColor(GetDC(GetParent(HWND)))
      mlngClrTextO = mlngClrText
      mlngClrShadow = ShiftColor(mlngClrFace, -&H40)
      mlngClrLight = ShiftColor(mlngClrFace, &H1F)
      mlngClrHighLight = ShiftColor(mlngClrFace, &H2F)
      mlngClrDarkShadow = ShiftColor(mlngClrFace, -&HC0)
      mlngOXPb = GetSysColor(C_COLOR_HIGHLIGHT)
      mlngOXPf = ShiftColorOXP(mlngOXPb)
      mlngO3D = &H4E91FE 'Down
      mlngO3H = &H8BCFFF 'hover
      mlngO3F = &HBA9EA0 ' face
      
      
    Case Else '// [Use Windows]
      mlngClrFace = GetSysColor(C_COLOR_BTNFACE)
      mlngClrFaceO = mlngClrFace
      mlngClrShadow = GetSysColor(C_COLOR_BTNSHADOW)
      mlngClrLight = GetSysColor(C_COLOR_BTNLIGHT)
      mlngClrDarkShadow = GetSysColor(C_COLOR_BTNDKSHADOW)
      mlngClrHighLight = GetSysColor(C_COLOR_BTNHIGHLIGHT)
      mlngClrText = GetSysColor(C_COLOR_BTNTEXT)
      mlngClrTextO = mlngClrText
      mlngOXPb = GetSysColor(C_COLOR_HIGHLIGHT)
      mlngOXPf = ShiftColorOXP(mlngOXPb)
      mlngO3D = &H4E91FE 'Down
      mlngO3H = &H8BCFFF 'hover
      Call GetGradientColor
      mlngBackC = mlngO3F
      
   End Select
   
   mlngClrMask = ConvertFromSystemColor(mlngMaskC)
   mlngXPFace = ShiftColor(mlngClrFace, &H30, mudtButtonType = [Windows XP])
   
End Sub

Private Function ShiftColor(ByVal vlngColor As Long, _
                            ByVal vlngValue As Long, _
                            Optional ByVal vblnIsXP As Boolean = False) As Long
   
   '// this function will add or remove a certain Color
   '// quantity and return the result
   
  Dim lngRed   As Long
  Dim lngBlue  As Long
  Dim lngGreen As Long
  Const C_Max As Long = 255&
   
   '// this is just a tricky way to do it and will result in weird colors for WinXP and KDE2
   If mblnIsSoft Then vlngValue = vlngValue \ 2
   
   If Not vblnIsXP Then
      lngBlue = ((vlngColor \ &H10000) Mod &H100) + vlngValue
    Else
      lngBlue = ((vlngColor \ &H10000) Mod &H100)
      lngBlue = lngBlue + ((lngBlue * vlngValue) \ &HC0)
   End If
   lngGreen = ((vlngColor \ &H100) Mod &H100) + vlngValue
   lngRed = (vlngColor And &HFF) + vlngValue
   
   '// values will overflow a byte only in one direction
   '// eg: if we added 32 to our color, then only a > 255 overflow can occurr.
   If vlngValue > 0 Then
      If lngRed > C_Max Then lngRed = C_Max
      If lngGreen > C_Max Then lngGreen = C_Max
      If lngBlue > C_Max Then lngBlue = C_Max
    ElseIf vlngValue < 0 Then
      If lngRed < 0 Then lngRed = 0
      If lngGreen < 0 Then lngGreen = 0
      If lngBlue < 0 Then lngBlue = 0
   End If
   
   '// more optimization by replacing the RGB function by its correspondent calculation
   ShiftColor = lngRed + 256& * lngGreen + 65536 * lngBlue
   
End Function

Private Function ShiftColorOXP(ByVal vlngColor As Long, _
                               Optional ByVal vlngBase As Long = &HB0) As Long
   
  Dim lngRed As Long
  Dim lngBlue As Long
  Dim lngGreen As Long
  Dim lngDelta As Long
  Const C_Max As Long = 255&
   
   lngBlue = ((vlngColor \ &H10000) Mod &H100)
   lngGreen = ((vlngColor \ &H100) Mod &H100)
   lngRed = (vlngColor And &HFF)
   lngDelta = &HFF - vlngBase
   
   lngBlue = vlngBase + lngBlue * lngDelta \ &HFF
   lngGreen = vlngBase + lngGreen * lngDelta \ &HFF
   lngRed = vlngBase + lngRed * lngDelta \ &HFF
   
   If lngRed > C_Max Then lngRed = C_Max
   If lngGreen > C_Max Then lngGreen = C_Max
   If lngBlue > C_Max Then lngBlue = C_Max
   
   ShiftColorOXP = lngRed + 256& * lngGreen + 65536 * lngBlue
   
End Function

Public Property Get ShowFocusRect() As Boolean
Attribute ShowFocusRect.VB_ProcData.VB_Invoke_Property = ";Appearance"
   
   ShowFocusRect = mblnShowFocusR
   
End Property

Public Property Let ShowFocusRect(ByVal vNewValue As Boolean)
   
   mblnShowFocusR = vNewValue
   Call Redraw(mbytLastStat, True)
   PropertyChanged "FOCUSR"
   
End Property

Public Property Get SoftBevel() As Boolean
Attribute SoftBevel.VB_ProcData.VB_Invoke_Property = ";Appearance"
   
   SoftBevel = mblnIsSoft
   
End Property

Public Property Let SoftBevel(ByVal vNewValue As Boolean)
   
   mblnIsSoft = vNewValue
   Call SetColors
   Call Redraw(mbytLastStat, True)
   PropertyChanged "SOFT"
   
End Property

Public Property Get SpecialEffect() As enuFX
Attribute SpecialEffect.VB_ProcData.VB_Invoke_Property = ";Appearance"
   
   SpecialEffect = mudtSFX
   
End Property

Public Property Let SpecialEffect(ByVal vNewValue As enuFX)
   
   mudtSFX = vNewValue
   Call Redraw(mbytLastStat, True)
   PropertyChanged "FX"
   
End Property

Private Sub TransBlt(ByVal vDstDC As Long, _
                     ByVal vDstX As Long, _
                     ByVal vDstY As Long, _
                     ByVal vDstW As Long, _
                     ByVal vDstH As Long, _
                     ByVal vSrcPic As StdPicture, _
                     Optional ByVal vTransColor As Long = -1, _
                     Optional ByVal vBrushColor As Long = -1, _
                     Optional ByVal vMonoMask As Boolean = False, _
                     Optional ByVal vIsGreyscale As Boolean = False, _
                     Optional ByVal vXPBlend As Boolean = False)
   
  Dim lngB        As Long
  Dim lngH        As Long
  Dim lngF        As Long
  Dim lngI        As Long
  Dim lngWidth    As Long
  Dim lngTmpDC    As Long
  Dim lngTmpBmp   As Long
  Dim lngTmpObj   As Long
  Dim lngSr2DC    As Long
  Dim lngSr2Bmp   As Long
  Dim lngSr2Obj   As Long
  Dim udtData1()  As RGBTRIPLE
  Dim udtData2()  As RGBTRIPLE
  Dim udtInfo     As BITMAPINFO
  Dim udtBrushRGB As RGBTRIPLE
  Dim lnggCol     As Long
  Dim lngSrcDC    As Long
  Dim lngtObj     As Long
   
   If vDstW = 0 Or vDstH = 0 Then Exit Sub
   
   lngSrcDC = CreateCompatibleDC(hDC)
   
   If vDstW < 0 Then vDstW = UserControl.ScaleX(vSrcPic.Width, 8, UserControl.ScaleMode)
   If vDstH < 0 Then vDstH = UserControl.ScaleY(vSrcPic.Height, 8, UserControl.ScaleMode)
   
   If vSrcPic.Type = 1 Then '// check if it's an icon or a bitmap
      lngtObj = SelectObject(lngSrcDC, vSrcPic)
    Else
      Dim hBrush As Long
      lngtObj = SelectObject(lngSrcDC, CreateCompatibleBitmap(vDstDC, vDstW, vDstH))
      hBrush = CreateSolidBrush(MaskColor)
      DrawIconEx lngSrcDC, 0, 0, vSrcPic.Handle, 0, 0, 0, hBrush, &H1 Or &H2
      DeleteObject hBrush
   End If
   
   lngTmpDC = CreateCompatibleDC(lngSrcDC)
   lngSr2DC = CreateCompatibleDC(lngSrcDC)
   lngTmpBmp = CreateCompatibleBitmap(vDstDC, vDstW, vDstH)
   lngSr2Bmp = CreateCompatibleBitmap(vDstDC, vDstW, vDstH)
   lngTmpObj = SelectObject(lngTmpDC, lngTmpBmp)
   lngSr2Obj = SelectObject(lngSr2DC, lngSr2Bmp)
   ReDim udtData1(vDstW * vDstH * 3 - 1)
   ReDim udtData2(UBound(udtData1))
   With udtInfo.bmiHeader
      .biSize = Len(udtInfo.bmiHeader)
      .biWidth = vDstW
      .biHeight = vDstH
      .biPlanes = 1
      .biBitCount = 24
   End With
   
   BitBlt lngTmpDC, 0, 0, vDstW, vDstH, vDstDC, vDstX, vDstY, vbSrcCopy
   BitBlt lngSr2DC, 0, 0, vDstW, vDstH, lngSrcDC, 0, 0, vbSrcCopy
   GetDIBits lngTmpDC, lngTmpBmp, 0, vDstH, udtData1(0), udtInfo, 0
   GetDIBits lngSr2DC, lngSr2Bmp, 0, vDstH, udtData2(0), udtInfo, 0
   
   If vBrushColor > 0 Then
      udtBrushRGB.rgbBlue = (vBrushColor \ &H10000) Mod &H100
      udtBrushRGB.rgbGreen = (vBrushColor \ &H100) Mod &H100
      udtBrushRGB.rgbRed = vBrushColor And &HFF
   End If
   
   If Not mblnUseMask Then vTransColor = -1
   
   lngWidth = vDstW - 1
   
   For lngH = 0 To vDstH - 1
      lngF = lngH * vDstW
      For lngB = 0 To lngWidth
         lngI = lngF + lngB
         If GetNearestColor(hDC, CLng(udtData2(lngI).rgbRed) + 256& * udtData2(lngI).rgbGreen + 65536 * udtData2(lngI).rgbBlue) <> vTransColor Then
            With udtData1(lngI)
               If vBrushColor > -1 Then
                  If vMonoMask Then
                     If (CLng(udtData2(lngI).rgbRed) + udtData2(lngI).rgbGreen + udtData2(lngI).rgbBlue) <= 384 Then udtData1(lngI) = udtBrushRGB
                   Else
                     udtData1(lngI) = udtBrushRGB
                  End If
                Else
                  If vIsGreyscale Then
                     lnggCol = CLng(udtData2(lngI).rgbRed * 0.3) + udtData2(lngI).rgbGreen * 0.59 + udtData2(lngI).rgbBlue * 0.11
                     .rgbRed = lnggCol
                     .rgbGreen = lnggCol
                     .rgbBlue = lnggCol
                   Else
                     If vXPBlend Then
                        .rgbRed = (CLng(.rgbRed) + udtData2(lngI).rgbRed * 2) \ 3
                        .rgbGreen = (CLng(.rgbGreen) + udtData2(lngI).rgbGreen * 2) \ 3
                        .rgbBlue = (CLng(.rgbBlue) + udtData2(lngI).rgbBlue * 2) \ 3
                      Else
                        udtData1(lngI) = udtData2(lngI)
                     End If
                  End If
               End If
            End With
         End If
      Next lngB
   Next lngH
   
   SetDIBitsToDevice vDstDC, vDstX, vDstY, vDstW, vDstH, 0, 0, 0, vDstH, udtData1(0), udtInfo, 0
   
   Erase udtData1, udtData2
   DeleteObject SelectObject(lngTmpDC, lngTmpObj)
   DeleteObject SelectObject(lngSr2DC, lngSr2Obj)
   If vSrcPic.Type = 3 Then DeleteObject SelectObject(lngSrcDC, lngtObj)
   DeleteDC lngTmpDC
   DeleteDC lngSr2DC
   DeleteObject lngtObj
   DeleteDC lngSrcDC
   
End Sub

Private Sub txtFX(ByRef vRect As Rect)
   
   If mudtSFX > cbNone Then
      
      With UserControl
         Dim lngCurFace As Long
         Dim udtRECT As Rect
         
         CopyRect udtRECT, vRect
         OffsetRect udtRECT, 1, 1
         
         Select Case mudtButtonType
          Case [Windows XP], [KDE 2]
            lngCurFace = mlngXPFace
            
          Case Else
            If mbytLastStat = 0 And mblnIsOver And mudtColorType <> [Custom Colors] And mudtButtonType = [Office XP] Then
               lngCurFace = mlngOXPf
             Else
               lngCurFace = mlngClrFace
            End If
         End Select
         
         SetTextColor .hDC, ShiftColor(lngCurFace, Abs(mudtSFX = cbEngraved) * C_FXDEPTH + (mudtSFX <> cbEngraved) * C_FXDEPTH)
         DrawText .hDC, mstrCurText, Len(mstrCurText), udtRECT, C_DT_CENTER
         
         If mudtSFX < cbShadowed Then
            OffsetRect udtRECT, -2, -2
            SetTextColor .hDC, ShiftColor(lngCurFace, Abs(mudtSFX <> cbEngraved) * C_FXDEPTH + (mudtSFX = cbEngraved) * C_FXDEPTH)
            DrawText .hDC, mstrCurText, Len(mstrCurText), udtRECT, C_DT_CENTER
         End If
      End With '// UserControl
      
   End If
   
End Sub

Public Property Get UseGreyscale() As Boolean
Attribute UseGreyscale.VB_ProcData.VB_Invoke_Property = ";Appearance"
   
   UseGreyscale = mblnUseGrey
   
End Property

Public Property Let UseGreyscale(ByVal vNewValue As Boolean)
   
   mblnUseGrey = vNewValue
   If Not mPicNormal Is Nothing Then Call Redraw(mbytLastStat, True)
   PropertyChanged "NGREY"
   
End Property

Public Property Get UseMaskColor() As Boolean
Attribute UseMaskColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
   
   UseMaskColor = mblnUseMask
   
End Property

Public Property Let UseMaskColor(ByVal vNewValue As Boolean)
   
   mblnUseMask = vNewValue
   If Not mPicNormal Is Nothing Then Call Redraw(mbytLastStat, True)
   PropertyChanged "UMCOL"
   
End Property

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
   
   mbytLastButton = vbLeftButton
   Call UserControl_Click
   
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
   
   Call SetColors
   Call Redraw(mbytLastStat, True)
   
End Sub

Private Sub UserControl_Click()
   
   If mbytLastButton = vbLeftButton And mblnIsEnabled Then
      If mblnIsCheckbox Then mblncValue = Not mblncValue
      '// be sure that the normal status is drawn
      Call Redraw(0, True)
      UserControl.Refresh
      RaiseEvent Click
   End If
   
End Sub

Private Sub UserControl_DblClick()
   
   If mbytLastButton = vbLeftButton Then
      Call UserControl_MouseDown(1, 0, 0, 0)
      SetCapture HWND
   End If
   
End Sub

Private Sub UserControl_GotFocus()
   
   mblnHasFocus = True
   Call Redraw(mbytLastStat, True)
   
End Sub

Private Sub UserControl_Hide()
   
   mblnIsShown = False
   
End Sub

Private Sub UserControl_InitProperties()
   
   mblnIsEnabled = True
   mblnShowFocusR = True
   mblnUseMask = True
   mstrCurText = Ambient.DisplayName
   Set UserControl.Font = Ambient.Font
   mudtButtonType = [Windows XP]
   mudtColorType = [Use Windows]
   Call SetColors
   mlngBackC = mlngClrFace
   mlngBackO = mlngBackC
   mlngForeC = mlngClrText
   mlngForeO = mlngForeC
   mlngMaskC = &HC0C0C0
   Call CalcTextRects
   
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
   
   RaiseEvent KeyDown(KeyCode, Shift)
   
   mbytLastKeyDown = KeyCode
   Select Case KeyCode
    Case vbKeySpace '// spacebar pressed
      Call Redraw(2, False)
    Case vbKeyRight, vbKeyDown '// right and down arrows
      SendKeys "{Tab}"
    Case vbKeyLeft, vbKeyUp '// left and up arrows
      SendKeys "+{Tab}"
   End Select
   
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
   
   RaiseEvent KeyPress(KeyAscii)
   
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
   
   RaiseEvent KeyUp(KeyCode, Shift)
   
   '// spacebar pressed, and not cancelled by the user
   If (KeyCode = vbKeySpace) And (mbytLastKeyDown = vbKeySpace) Then
      If mblnIsCheckbox Then mblncValue = Not mblncValue
      Call Redraw(0, False)
      UserControl.Refresh
      RaiseEvent Click
   End If
   
End Sub

Private Sub UserControl_LostFocus()
   
   mblnHasFocus = False
   Call Redraw(mbytLastStat, True)
   
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
   
   RaiseEvent MouseDown(Button, Shift, x, Y)
   mbytLastButton = Button
   If Button <> vbRightButton Then Call Redraw(2, False)
   
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
   
   RaiseEvent MouseMove(Button, Shift, x, Y)
   If Button < vbRightButton Then
      If Not isMouseOver Then
         '// we are outside the button
         Call Redraw(0, False)
       Else
         '// we are inside the button
         If Button = 0 And Not mblnIsOver Then
            OverTimer.Enabled = True
            mblnIsOver = True
            Call Redraw(0, True)
            RaiseEvent MouseOver
          ElseIf Button = vbLeftButton Then
            mblnIsOver = True
            Call Redraw(2, False)
            mblnIsOver = False
         End If
      End If
   End If
   
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
   
   RaiseEvent MouseUp(Button, Shift, x, Y)
   If Button <> vbRightButton Then Call Redraw(0, False)
   
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   
   With PropBag
      mudtButtonType = .ReadProperty("BTYPE", [Windows XP])
      mstrCurText = .ReadProperty("TX", vbNullString)
      mblnIsEnabled = .ReadProperty("ENAB", True)
      Set UserControl.Font = .ReadProperty("FONT", UserControl.Font)
      mudtColorType = .ReadProperty("COLTYPE", 1)
      mblnShowFocusR = .ReadProperty("FOCUSR", True)
      mlngBackC = .ReadProperty("BCOL", GetSysColor(C_COLOR_BTNFACE))
      mlngBackO = .ReadProperty("BCOLO", mlngBackC)
      mlngForeC = .ReadProperty("FCOL", GetSysColor(C_COLOR_BTNTEXT))
      mlngForeO = .ReadProperty("FCOLO", mlngForeC)
      mlngMaskC = .ReadProperty("MCOL", &HC0C0C0)
      UserControl.MousePointer = .ReadProperty("MPTR", 0)
      Set UserControl.MouseIcon = .ReadProperty("MICON", Nothing)
      Set mPicNormal = .ReadProperty("PICN", Nothing)
      Set mPicHover = .ReadProperty("PICH", Nothing)
      mblnUseMask = .ReadProperty("UMCOL", True)
      mblnIsSoft = .ReadProperty("SOFT", False)
      mudtPicPosition = .ReadProperty("PICPOS", 0)
      mblnUseGrey = .ReadProperty("NGREY", False)
      mudtSFX = .ReadProperty("FX", 0)
      mblnIsCheckbox = .ReadProperty("CHECK", False)
      mblncValue = .ReadProperty("VALUE", False)
   End With
   
   UserControl.Enabled = mblnIsEnabled
   Call CalcPicSize
   Call CalcTextRects
   Call SetAccessKeys
   
End Sub

Private Sub UserControl_Resize()
   
  Dim lngRgn1 As Long
  Dim lngRgn2 As Long
   
   '// get button size
   GetClientRect UserControl.HWND, mudtRC3
   '// assign these values to mlngHeight and mlngWidth
   mlngHeight = mudtRC3.bottom
   mlngWidth = mudtRC3.right
   '// build the FocusRect size and position depending on the button type
   If mudtButtonType = [KDE 2] Then
      InflateRect mudtRC3, -5, -5
      OffsetRect mudtRC3, 1, 1
    Else
      InflateRect mudtRC3, -4, -4
   End If
   Call CalcTextRects
   
   If mlngRgnNorm Then DeleteObject mlngRgnNorm
   
   '// this creates the regions to "cut" the UserControl
   '// so it will be transparent in certain areas
   
   mlngRgnNorm = CreateRectRgn(0, 0, mlngWidth, mlngHeight)
   lngRgn2 = CreateRectRgn(0, 0, 0, 0)
   
   Select Case mudtButtonType
    Case [Java metal]
      lngRgn1 = CreateRectRgn(0, mlngHeight, 1, mlngHeight - 1)
      CombineRgn lngRgn2, mlngRgnNorm, lngRgn1, C_RGN_DIFF
      DeleteObject lngRgn1
      lngRgn1 = CreateRectRgn(mlngWidth, 0, mlngWidth - 1, 1)
      CombineRgn mlngRgnNorm, lngRgn2, lngRgn1, C_RGN_DIFF
      DeleteObject lngRgn1
      
    Case [Windows XP]
      lngRgn1 = CreateRectRgn(0, 0, 2, 1)
      CombineRgn lngRgn2, mlngRgnNorm, lngRgn1, C_RGN_DIFF
      DeleteObject lngRgn1
      lngRgn1 = CreateRectRgn(0, mlngHeight, 2, mlngHeight - 1)
      CombineRgn mlngRgnNorm, lngRgn2, lngRgn1, C_RGN_DIFF
      DeleteObject lngRgn1
      lngRgn1 = CreateRectRgn(mlngWidth, 0, mlngWidth - 2, 1)
      CombineRgn lngRgn2, mlngRgnNorm, lngRgn1, C_RGN_DIFF
      DeleteObject lngRgn1
      lngRgn1 = CreateRectRgn(mlngWidth, mlngHeight, mlngWidth - 2, mlngHeight - 1)
      CombineRgn mlngRgnNorm, lngRgn2, lngRgn1, C_RGN_DIFF
      DeleteObject lngRgn1
      lngRgn1 = CreateRectRgn(0, 1, 1, 2)
      CombineRgn lngRgn2, mlngRgnNorm, lngRgn1, C_RGN_DIFF
      DeleteObject lngRgn1
      lngRgn1 = CreateRectRgn(0, mlngHeight - 1, 1, mlngHeight - 2)
      CombineRgn mlngRgnNorm, lngRgn2, lngRgn1, C_RGN_DIFF
      DeleteObject lngRgn1
      lngRgn1 = CreateRectRgn(mlngWidth, 1, mlngWidth - 1, 2)
      CombineRgn lngRgn2, mlngRgnNorm, lngRgn1, C_RGN_DIFF
      DeleteObject lngRgn1
      lngRgn1 = CreateRectRgn(mlngWidth, mlngHeight - 1, mlngWidth - 1, mlngHeight - 2)
      CombineRgn mlngRgnNorm, lngRgn2, lngRgn1, C_RGN_DIFF
      DeleteObject lngRgn1
      
   End Select
   
   DeleteObject lngRgn2
   
   SetWindowRgn UserControl.HWND, mlngRgnNorm, True
   
   If mlngHeight Then Call Redraw(0, True)
   
End Sub

Private Sub UserControl_Show()
   
   mblnIsShown = True
   Call SetColors
   Call Redraw(0, True)
   
End Sub

Private Sub UserControl_Terminate()
   
   mblnIsShown = False
   DeleteObject mlngRgnNorm
   
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   
   With PropBag
      Call .WriteProperty("BTYPE", mudtButtonType)
      Call .WriteProperty("TX", mstrCurText)
      Call .WriteProperty("ENAB", mblnIsEnabled)
      Call .WriteProperty("FONT", UserControl.Font)
      Call .WriteProperty("COLTYPE", mudtColorType)
      Call .WriteProperty("FOCUSR", mblnShowFocusR)
      Call .WriteProperty("BCOL", mlngBackC)
      Call .WriteProperty("BCOLO", mlngBackO)
      Call .WriteProperty("FCOL", mlngForeC)
      Call .WriteProperty("FCOLO", mlngForeO)
      Call .WriteProperty("MCOL", mlngMaskC)
      Call .WriteProperty("MPTR", UserControl.MousePointer)
      Call .WriteProperty("MICON", UserControl.MouseIcon)
      Call .WriteProperty("PICN", mPicNormal)
      Call .WriteProperty("PICH", mPicHover)
      Call .WriteProperty("UMCOL", mblnUseMask)
      Call .WriteProperty("SOFT", mblnIsSoft)
      Call .WriteProperty("PICPOS", mudtPicPosition)
      Call .WriteProperty("NGREY", mblnUseGrey)
      Call .WriteProperty("FX", mudtSFX)
      Call .WriteProperty("CHECK", mblnIsCheckbox)
      Call .WriteProperty("VALUE", mblncValue)
   End With
   
End Sub

Public Property Get Value() As Boolean
   
   Value = mblncValue
   
End Property

Public Property Let Value(ByVal vNewValue As Boolean)
   
   mblncValue = vNewValue
   If mblnIsCheckbox Then Call Redraw(0, True)
   PropertyChanged "VALUE"
   
End Property

Private Sub DrawGradiant(ByVal vlngTopClr As Long, ByVal lngBottomColor As Long)

  Dim lngI    As Long
  Dim lngStep As Long
  Dim sngRed1 As Single
  Dim sngGrn1 As Single
  Dim sngBlu1 As Single
  Dim sngRed2 As Single
  Dim sngGrn2 As Single
  Dim sngBlu2 As Single

   On Error Resume Next
   'mlngWidth, mlngHeight

   Call GetRGBColor(vlngTopClr, sngRed1, sngGrn1, sngBlu1)
   Call GetRGBColor(lngBottomColor, sngRed2, sngGrn2, sngBlu2)

   lngStep = mlngHeight
   '// Get gradient color step
   sngRed2 = (sngRed2 - sngRed1) / lngStep
   sngGrn2 = (sngGrn2 - sngGrn1) / lngStep
   sngBlu2 = (sngBlu2 - sngBlu1) / lngStep
   
   '// Begin drawing vertical gradient
   For lngI = 0 To lngStep
      UserControl.Line (0, lngI)-(mlngWidth, lngI), RGB(CInt(sngRed1), CInt(sngGrn1), CInt(sngBlu1))
      sngRed1 = sngRed1 + sngRed2
      sngGrn1 = sngGrn1 + sngGrn2
      sngBlu1 = sngBlu1 + sngBlu2
   Next lngI

End Sub

Private Sub GetRGBColor(ByVal vlngColor As Long, _
                        ByRef rsngRed As Single, _
                        ByRef rsngGrn As Single, _
                        ByRef rsngBlu As Single)

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
   'ErrHandler True, Err.Number, Err.Description, "Frame3D", "GetRGBColor"
   Err.Clear
   Resume Exit_Proc

End Sub

Private Sub GetGradientColor()

   Select Case GetThemeName(UserControl.HWND)
   Case "NormalColor"
      mlngO3F = RGB(122, 150, 223)
      'mlngO3F = RGB(19, 97, 156)

   Case "Metallic"
      mlngO3F = RGB(200, 200, 200)

   Case "HomeStead"
      mlngO3F = RGB(150, 170, 115)
   
   Case Else
      mlngO3F = RGB(122, 150, 223)
   
   End Select

End Sub

Private Function GetThemeName(ByVal lngHWND As Long) As String

   'Returns The current Windows Theme Name

  Dim hTheme        As Long
  Dim sShellStyle   As String
  Dim sThemeFile    As String
  Dim lPtrThemeFile As Long
  Dim lPtrColorName As Long
  Dim iPos          As Long
  Const C_MaxChar As Long = 260

   On Error Resume Next
   hTheme = OpenThemeData(lngHWND, StrPtr("ExplorerBar"))

   If Not hTheme = 0 Then
      ReDim bThemeFile(0 To C_MaxChar * 2) As Byte
      lPtrThemeFile = VarPtr(bThemeFile(0))

      ReDim bColorName(0 To C_MaxChar * 2) As Byte
      lPtrColorName = VarPtr(bColorName(0))

      GetCurrentThemeName lPtrThemeFile, C_MaxChar, lPtrColorName, C_MaxChar, 0, 0
      sThemeFile = bThemeFile
      iPos = InStr(sThemeFile, vbNullChar)

      If iPos > 1 Then
         sThemeFile = left$(sThemeFile, iPos - 1)
      End If

      GetThemeName = bColorName
      iPos = InStr(GetThemeName, vbNullChar)

      If iPos > 1 Then
         GetThemeName = left$(GetThemeName, iPos - 1)
      End If

      sShellStyle = sThemeFile

      For iPos = Len(sThemeFile) To 1 Step -1

         If (Mid$(sThemeFile, iPos, 1) = "\") Then
            sShellStyle = left$(sThemeFile, iPos)
            Exit For
         End If

      Next iPos

      sShellStyle = sShellStyle & "Shell\" & GetThemeName & "\ShellStyle.dll"
      CloseThemeData hTheme

   Else 'hTheme=0
      GetThemeName = "Classic"
   End If

   On Error GoTo 0

End Function

