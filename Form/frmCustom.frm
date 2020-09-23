VERSION 5.00
Begin VB.Form frmCustom 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Convert It! - Custom Conversion"
   ClientHeight    =   2715
   ClientLeft      =   5190
   ClientTop       =   5070
   ClientWidth     =   5625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   5625
   Begin Convert_It.mhButton cmdAdd 
      Height          =   450
      Left            =   1785
      TabIndex        =   12
      Top             =   2100
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   794
      BTYPE           =   3
      TX              =   "&OK"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14653050
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "frmCustom.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Convert_It.Frame3D fraGroup 
      Height          =   1830
      Left            =   180
      Top             =   135
      Width           =   5250
      _ExtentX        =   9260
      _ExtentY        =   3228
      BorderType      =   9
      BevelWidth      =   3
      BevelInner      =   0
      Caption3D       =   0
      CaptionAlignment=   0
      CaptionLocation =   0
      BackColor       =   -2147483633
      CornerDiameter  =   7
      FillColor       =   16761247
      FillStyle       =   1
      DrawStyle       =   0
      DrawWidth       =   1
      FloodPercent    =   0
      FloodShowPct    =   0   'False
      FloodType       =   0
      FloodColor      =   16761247
      FillGradient    =   0
      Collapsible     =   0   'False
      ChevronColor    =   -2147483630
      Collapse        =   0   'False
      FullHeight      =   1830
      ChevronType     =   0
      MousePointer    =   0
      MouseIcon       =   "frmCustom.frx":001C
      Picture         =   "frmCustom.frx":0038
      Border3DHighlight=   -2147483628
      Border3DShadow  =   -2147483632
      Enabled         =   -1  'True
      CaptionMAlignment=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Verdana"
      FontSize        =   9.75
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   -2147483630
      Caption         =   "Group: Custom"
      UseMnemonic     =   0   'False
      Begin VB.VScrollBar VScroll1 
         Height          =   480
         Left            =   2025
         Max             =   19
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1170
         Width           =   270
      End
      Begin VB.ComboBox cboIndex 
         Height          =   315
         ItemData        =   "frmCustom.frx":0054
         Left            =   1275
         List            =   "frmCustom.frx":0094
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1245
         Width           =   780
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear &Entry"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2385
         TabIndex        =   8
         Top             =   1215
         Width           =   1110
      End
      Begin VB.TextBox txtUnit2 
         Height          =   285
         Left            =   3570
         MaxLength       =   30
         TabIndex        =   2
         Top             =   675
         Width           =   1335
      End
      Begin VB.TextBox txtFactor 
         Height          =   285
         Left            =   2115
         TabIndex        =   1
         Top             =   675
         Width           =   1335
      End
      Begin VB.TextBox txtUnit1 
         Height          =   285
         Left            =   465
         MaxLength       =   30
         TabIndex        =   0
         Top             =   675
         Width           =   1335
      End
      Begin VB.Label lblMisc 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   " Index "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   750
         TabIndex        =   10
         Top             =   1290
         Width           =   510
      End
      Begin VB.Label lblMisc 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   " Factor"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   2520
         TabIndex        =   7
         Top             =   435
         Width           =   510
      End
      Begin VB.Label lblMisc 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   " Unit "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   4020
         TabIndex        =   6
         Top             =   435
         Width           =   390
      End
      Begin VB.Label lblMisc 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   " Unit "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   825
         TabIndex        =   5
         Top             =   435
         Width           =   390
      End
      Begin VB.Label lblMisc 
         AutoSize        =   -1  'True
         Caption         =   " = "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   1830
         TabIndex        =   4
         Top             =   690
         Width           =   255
      End
      Begin VB.Label lblMisc 
         AutoSize        =   -1  'True
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   255
         TabIndex        =   3
         Top             =   690
         Width           =   105
      End
   End
   Begin Convert_It.mhButton cmdClose 
      Height          =   450
      Left            =   2970
      TabIndex        =   13
      Top             =   2100
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   794
      BTYPE           =   3
      TX              =   "&Cancel"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14653050
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "frmCustom.frx":00DF
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "frmCustom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnWorking   As Boolean

Private Sub cboIndex_Click()
         
  Dim intIndex As Integer
  
   intIndex = cboIndex.ListIndex + 1
  
   txtUnit1.Text = gcINI.GetSetting(C_Custom & Format$(intIndex, "00"), "D0")
   txtUnit2.Text = gcINI.GetSetting(C_Custom & Format$(intIndex, "00"), "D1")
   txtFactor.Text = gcINI.GetSetting(C_Custom & Format$(intIndex, "00"), "1")
   
   VScroll1.Value = cboIndex.ListIndex
   
End Sub

Private Sub cmdAdd_Click()

  Dim strIndex As String

   txtUnit1.Text = Trim$(txtUnit1.Text)
   txtUnit2.Text = Trim$(txtUnit2.Text)

   With frmMain
      strIndex = Format$(cboIndex.ListIndex + 1, "00")
      gcINI.SaveSetting C_Custom & strIndex, "D0", txtUnit1.Text
      gcINI.SaveSetting C_Custom & strIndex, "D1", txtUnit2.Text
      gcINI.SaveSetting C_Custom & strIndex, "0", 1
      gcINI.SaveSetting C_Custom & strIndex, "1", txtFactor.Text
   End With

   Unload Me

End Sub

Private Sub cmdClear_Click()
   
   txtUnit1.Text = vbNullString
   txtUnit2.Text = vbNullString
   txtFactor.Text = vbNullString
   
End Sub

Private Sub cmdClose_Click()

   Unload Me

End Sub

Private Sub Form_Load()

   Call CenterForm(Me)
   Me.Icon = frmMain.Icon
   cboIndex.ListIndex = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)

   Set frmCustom = Nothing

End Sub

Private Sub Numeric_All(ByRef rTextObj As TextBox)

  Dim lngI       As Long
  Dim lngN       As Long
  Dim strTemp    As String
  Dim bytChar    As Byte
  Dim blnNoMatch As Boolean

   On Error Resume Next

   If mblnWorking Then Exit Sub
   mblnWorking = True

   With rTextObj
      lngI = .SelStart

      For lngN = 1 To Len(.Text)
         bytChar = Asc(Mid$(.Text, lngN, 1))

         If ((bytChar >= 48 And bytChar <= 57) Or (bytChar = 45 And lngN = 1) Or (bytChar = 46 And blnNoMatch = False)) _
            Then
            If bytChar = 46 Then blnNoMatch = True
            strTemp = strTemp & Mid$(.Text, lngN, 1)
         Else
            lngI = lngI - 1
         End If

      Next lngN

      If lngI > Len(strTemp) Or lngI < 0 Then lngI = Len(strTemp)

      .Text = strTemp
      .SelStart = lngI

   End With

   mblnWorking = False

End Sub

Private Sub txtFactor_Change()
   
   Call Numeric_All(txtFactor)

End Sub

Private Sub VScroll1_Change()

   cboIndex.ListIndex = VScroll1.Value
   
End Sub
