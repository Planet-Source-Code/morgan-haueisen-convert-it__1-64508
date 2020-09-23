VERSION 5.00
Begin VB.Form frmAddNew 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Convert It! - Add New Conversion"
   ClientHeight    =   2715
   ClientLeft      =   4545
   ClientTop       =   4230
   ClientWidth     =   7605
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   7605
   Begin Convert_It.mhButton cmdAdd 
      Height          =   405
      Left            =   6090
      TabIndex        =   5
      Top             =   900
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   714
      BTYPE           =   3
      TX              =   "&Add"
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
      MICON           =   "frmAddNew.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Convert_It.Frame3D fraGroup 
      Height          =   2205
      Left            =   495
      Top             =   240
      Width           =   5250
      _ExtentX        =   9260
      _ExtentY        =   3889
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
      FullHeight      =   2205
      ChevronType     =   0
      MousePointer    =   0
      MouseIcon       =   "frmAddNew.frx":001C
      Picture         =   "frmAddNew.frx":0038
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
      Caption         =   "Group"
      UseMnemonic     =   0   'False
      Begin VB.TextBox txtUnitName 
         Height          =   285
         Left            =   1410
         MaxLength       =   30
         TabIndex        =   1
         Top             =   1365
         Width           =   3615
      End
      Begin VB.TextBox txtUnits 
         Height          =   285
         Left            =   1410
         TabIndex        =   0
         Text            =   "0"
         Top             =   945
         Width           =   3615
      End
      Begin VB.Label lblMisc 
         AutoSize        =   -1  'True
         Caption         =   " Unit Name: "
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
         Index           =   1
         Left            =   285
         TabIndex        =   4
         Top             =   1395
         Width           =   1080
      End
      Begin VB.Label lblMisc 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   " = "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   975
         TabIndex        =   3
         Top             =   915
         Width           =   390
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "Label1"
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
         Left            =   360
         TabIndex        =   2
         Top             =   525
         Width           =   555
      End
   End
   Begin Convert_It.mhButton cmdClose 
      Height          =   405
      Left            =   6090
      TabIndex        =   6
      Top             =   1335
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   714
      BTYPE           =   3
      TX              =   "&Close"
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
      MICON           =   "frmAddNew.frx":0054
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "frmAddNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnWorking   As Boolean
Private mblnItemAdded As Boolean

Private Sub cmdAdd_Click()

  Dim strIndex As String

   txtUnitName.Text = Trim$(txtUnitName.Text)

   If LenB(txtUnitName.Text) Then
      If Not (rVal(txtUnits.Text) = 0) Then

         With frmMain
            strIndex = CStr(.lstInput.ListCount)
            gcINI.SaveSetting .lstGroup.Text, "D" & strIndex, txtUnitName.Text
            gcINI.SaveSetting .lstGroup.Text, strIndex, txtUnits.Text
         End With

         mblnItemAdded = True
      End If

   End If

   Me.Hide

End Sub

Private Sub cmdClose_Click()

   Me.Hide

End Sub

Private Sub Form_Load()

   With frmMain
      fraGroup.Caption = "Group: " & .lstGroup.Text
      lblInfo.Caption = "1 " & gcINI.GetSetting(.lstGroup.Text, "D0")
   End With

   Call CenterForm(Me)
   Me.Icon = frmMain.Icon

End Sub

Private Sub Form_Unload(Cancel As Integer)

   Set frmAddNew = Nothing

End Sub

Public Property Get ItemAdded() As Boolean

   ItemAdded = mblnItemAdded

End Property

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

         If ((bytChar >= 48 And bytChar <= 57) Or (bytChar = 45 And lngN = 1) Or (bytChar = 46 And blnNoMatch = False)) Then
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

Private Sub txtUnits_Change()

   Call Numeric_All(txtUnits)

End Sub

