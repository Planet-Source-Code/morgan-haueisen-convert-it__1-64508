VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Convert It!"
   ClientHeight    =   5595
   ClientLeft      =   4395
   ClientTop       =   5085
   ClientWidth     =   7800
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   7800
   Begin Convert_It.Frame3D fraToolBar 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      Top             =   0
      Width           =   7800
      _ExtentX        =   13758
      _ExtentY        =   1058
      BorderType      =   0
      BevelWidth      =   3
      BevelInner      =   0
      Caption3D       =   0
      CaptionAlignment=   0
      CaptionLocation =   0
      BackColor       =   8421504
      CornerDiameter  =   7
      FillColor       =   -2147483633
      FillStyle       =   1
      DrawStyle       =   0
      DrawWidth       =   1
      FloodPercent    =   0
      FloodShowPct    =   0   'False
      FloodType       =   0
      FloodColor      =   16761247
      FillGradient    =   2
      Collapsible     =   0   'False
      ChevronColor    =   -2147483630
      Collapse        =   0   'False
      FullHeight      =   600
      ChevronType     =   0
      MousePointer    =   0
      MouseIcon       =   "frmMain.frx":1AFA
      Picture         =   "frmMain.frx":1B16
      Border3DHighlight=   -2147483628
      Border3DShadow  =   -2147483632
      Enabled         =   -1  'True
      CaptionMAlignment=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   -2147483630
      Caption         =   ""
      UseMnemonic     =   0   'False
      Begin Convert_It.mhButton cmdExit 
         Height          =   480
         Left            =   90
         TabIndex        =   9
         Top             =   60
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   847
         BTYPE           =   10
         TX              =   "&File"
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
         FOCUSR          =   0   'False
         BCOL            =   14653050
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "frmMain.frx":1B32
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Convert_It.mhButton cmdInvert 
         Height          =   480
         Left            =   765
         TabIndex        =   10
         Top             =   60
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   847
         BTYPE           =   10
         TX              =   "&Invert Selection"
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
         FOCUSR          =   0   'False
         BCOL            =   14653050
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "frmMain.frx":1B4E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Convert_It.mhButton cmdCopy 
         Height          =   480
         Left            =   2280
         TabIndex        =   11
         Top             =   60
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   847
         BTYPE           =   10
         TX              =   "&Copy to Clipboard"
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
         FOCUSR          =   0   'False
         BCOL            =   14653050
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "frmMain.frx":1B6A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Convert_It.mhButton cmdFixed 
         Height          =   480
         Left            =   3795
         TabIndex        =   12
         Top             =   60
         Width           =   2250
         _ExtentX        =   3969
         _ExtentY        =   847
         BTYPE           =   10
         TX              =   "&Dec-Point Precision = None"
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
         FOCUSR          =   0   'False
         BCOL            =   14653050
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "frmMain.frx":1B86
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Convert_It.mhButton cmdAbout 
         Height          =   480
         Left            =   6045
         TabIndex        =   13
         Top             =   60
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   847
         BTYPE           =   10
         TX              =   "&About"
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
         FOCUSR          =   0   'False
         BCOL            =   14653050
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "frmMain.frx":1BA2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin Convert_It.Frame3D Frame3D4 
      Height          =   960
      Left            =   2295
      Top             =   4590
      Width           =   5445
      _ExtentX        =   9604
      _ExtentY        =   1693
      BorderType      =   8
      BevelWidth      =   3
      BevelInner      =   0
      Caption3D       =   0
      CaptionAlignment=   0
      CaptionLocation =   0
      BackColor       =   -2147483633
      CornerDiameter  =   7
      FillColor       =   -2147483645
      FillStyle       =   1
      DrawStyle       =   0
      DrawWidth       =   1
      FloodPercent    =   0
      FloodShowPct    =   0   'False
      FloodType       =   0
      FloodColor      =   16761247
      FillGradient    =   1
      Collapsible     =   0   'False
      ChevronColor    =   -2147483630
      Collapse        =   0   'False
      FullHeight      =   960
      ChevronType     =   0
      MousePointer    =   0
      MouseIcon       =   "frmMain.frx":1BBE
      Picture         =   "frmMain.frx":1BDA
      Border3DHighlight=   -2147483628
      Border3DShadow  =   -2147483632
      Enabled         =   -1  'True
      CaptionMAlignment=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   -2147483630
      Caption         =   ""
      UseMnemonic     =   0   'False
      Begin VB.TextBox txtResult 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   735
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   495
         Width           =   2385
      End
      Begin VB.TextBox txtInput 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   735
         TabIndex        =   4
         Top             =   150
         Width           =   2385
      End
      Begin VB.Label lblOutputUnits 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3195
         TabIndex        =   8
         Top             =   525
         Width           =   2280
      End
      Begin VB.Label lblInputUnits 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3195
         TabIndex        =   7
         Top             =   180
         Width           =   2280
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Output: "
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
         Left            =   75
         TabIndex        =   5
         Top             =   525
         Width           =   660
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Input: "
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
         Left            =   195
         TabIndex        =   3
         Top             =   180
         Width           =   540
      End
   End
   Begin Convert_It.Frame3D fraInput 
      Height          =   3915
      Left            =   2295
      Top             =   645
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   6906
      BorderType      =   8
      BevelWidth      =   3
      BevelInner      =   0
      Caption3D       =   2
      CaptionAlignment=   0
      CaptionLocation =   0
      BackColor       =   -2147483646
      CornerDiameter  =   7
      FillColor       =   -2147483633
      FillStyle       =   1
      DrawStyle       =   0
      DrawWidth       =   1
      FloodPercent    =   0
      FloodShowPct    =   0   'False
      FloodType       =   0
      FloodColor      =   16761247
      FillGradient    =   2
      Collapsible     =   0   'False
      ChevronColor    =   -2147483630
      Collapse        =   0   'False
      FullHeight      =   3915
      ChevronType     =   0
      MousePointer    =   0
      MouseIcon       =   "frmMain.frx":1BF6
      Picture         =   "frmMain.frx":1C12
      Border3DHighlight=   -2147483628
      Border3DShadow  =   -2147483632
      Enabled         =   -1  'True
      CaptionMAlignment=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Verdana"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   -2147483630
      Caption         =   "Input"
      UseMnemonic     =   0   'False
      Begin VB.ListBox lstInput 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3555
         IntegralHeight  =   0   'False
         Left            =   105
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   2460
      End
   End
   Begin Convert_It.Frame3D fraGroup 
      Height          =   4920
      Left            =   60
      Top             =   645
      Width           =   2205
      _ExtentX        =   3889
      _ExtentY        =   8678
      BorderType      =   8
      BevelWidth      =   3
      BevelInner      =   0
      Caption3D       =   2
      CaptionAlignment=   0
      CaptionLocation =   0
      BackColor       =   -2147483646
      CornerDiameter  =   7
      FillColor       =   -2147483633
      FillStyle       =   1
      DrawStyle       =   0
      DrawWidth       =   1
      FloodPercent    =   0
      FloodShowPct    =   0   'False
      FloodType       =   0
      FloodColor      =   16761247
      FillGradient    =   2
      Collapsible     =   0   'False
      ChevronColor    =   -2147483630
      Collapse        =   0   'False
      FullHeight      =   4920
      ChevronType     =   0
      MousePointer    =   0
      MouseIcon       =   "frmMain.frx":1C2E
      Picture         =   "frmMain.frx":1C4A
      Border3DHighlight=   -2147483628
      Border3DShadow  =   -2147483632
      Enabled         =   -1  'True
      CaptionMAlignment=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Verdana"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   -2147483630
      Caption         =   "Group"
      UseMnemonic     =   -1  'True
      Begin VB.ListBox lstGroup 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4560
         IntegralHeight  =   0   'False
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   1935
      End
   End
   Begin Convert_It.Frame3D fraOutput 
      Height          =   3915
      Left            =   5025
      Top             =   645
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   6906
      BorderType      =   8
      BevelWidth      =   3
      BevelInner      =   0
      Caption3D       =   2
      CaptionAlignment=   0
      CaptionLocation =   0
      BackColor       =   -2147483646
      CornerDiameter  =   7
      FillColor       =   -2147483633
      FillStyle       =   1
      DrawStyle       =   0
      DrawWidth       =   1
      FloodPercent    =   0
      FloodShowPct    =   0   'False
      FloodType       =   0
      FloodColor      =   16761247
      FillGradient    =   2
      Collapsible     =   0   'False
      ChevronColor    =   -2147483630
      Collapse        =   0   'False
      FullHeight      =   3915
      ChevronType     =   0
      MousePointer    =   0
      MouseIcon       =   "frmMain.frx":1C66
      Picture         =   "frmMain.frx":1C82
      Border3DHighlight=   -2147483628
      Border3DShadow  =   -2147483632
      Enabled         =   -1  'True
      CaptionMAlignment=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Verdana"
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   -2147483630
      Caption         =   "Output"
      UseMnemonic     =   0   'False
      Begin VB.ListBox lstOutput 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3555
         IntegralHeight  =   0   'False
         Left            =   105
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   2460
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "mnuFile"
      Visible         =   0   'False
      Begin VB.Menu mnuAddNew 
         Caption         =   "&Add New"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Edit Existing"
      End
      Begin VB.Menu mnuCustom 
         Caption         =   "&Custom"
      End
      Begin VB.Menu sep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuFixDec 
      Caption         =   "mnuFixDec"
      Visible         =   0   'False
      Begin VB.Menu mnuD 
         Caption         =   "Decimal-Point Precision"
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu mnuD 
         Caption         =   "&1"
         Index           =   1
      End
      Begin VB.Menu mnuD 
         Caption         =   "&2"
         Index           =   2
      End
      Begin VB.Menu mnuD 
         Caption         =   "&3"
         Index           =   3
      End
      Begin VB.Menu mnuD 
         Caption         =   "&4"
         Index           =   4
      End
      Begin VB.Menu mnuD 
         Caption         =   "&5"
         Index           =   5
      End
      Begin VB.Menu mnuD 
         Caption         =   "&6"
         Index           =   6
      End
      Begin VB.Menu mnuD 
         Caption         =   "&7"
         Index           =   7
      End
      Begin VB.Menu mnuD 
         Caption         =   "&8"
         Index           =   8
      End
      Begin VB.Menu mnuD 
         Caption         =   "&9"
         Index           =   9
      End
      Begin VB.Menu mnuD 
         Caption         =   "1&0"
         Index           =   10
      End
      Begin VB.Menu mnuD 
         Caption         =   "&None"
         Index           =   11
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'//************************************************************************************************
'// Author: Morgan Haueisen (morganh@hartcom.net)
'// Copyright (c) 2006
'// Version 1.0.0
'// http://www.planetsourcecode.com/vb/scripts/BrowseCategoryOrSearchResults.asp?lngWId=1&blnAuthorSearch=TRUE&lngAuthorId=885253927&strAuthorName=Morgan%20Haueisen&txtMaxNumberOfEntriesPerPage=25
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

Private Const C_MenuTitle As String = "Dec-Point Precision = "

Private mdblInput  As Double
Private mdblOutput As Double
Private mintFixDec As Integer
Private mblnLoaded As Boolean
Private mlngLastG  As Long

Private Sub Calculate()

   On Error Resume Next

   If Not mblnLoaded Then Exit Sub
   
   Dim dblIn As Double
   dblIn = rVal(txtInput.Text)

   Select Case lstGroup.Text
   Case "Base"
      txtResult.Text = ConvertBase(txtInput.Text, mdblInput, mdblOutput)

   Case "Consumption"
      Select Case lstInput.Text
         Case "liter/100 km"
            dblIn = (1 / dblIn) * 100
         Case Else
            Call ShowResult(dblIn)
      End Select
      Select Case lstOutput.Text
         Case "liter/100 km"
            dblIn = (mdblOutput / mdblInput) * dblIn
            dblIn = (1 / dblIn) * 100
            If mintFixDec Then
               dblIn = Round(dblIn, mintFixDec)
            End If
            txtResult.Text = dblIn
         Case Else
            Call ShowResult(dblIn)
      End Select
   
   Case "Temperature"
      Select Case lstInput.Text
      Case "Fahrenheit"
         Select Case lstOutput.Text
         Case "Fahrenheit"
            txtResult.Text = dblIn

         Case "Celsius"
            txtResult.Text = (5 / 9) * (dblIn - 32)

         Case "Rankine"
            txtResult.Text = dblIn + 459.69

         Case "Kelvin"
            txtResult.Text = ((5 / 9) * (dblIn - 32)) + 273.15
         End Select

      Case "Celsius"
         Select Case lstOutput.Text
         Case "Fahrenheit"
            txtResult.Text = (9 / 5) * dblIn + 32

         Case "Celsius"
            txtResult.Text = txtInput.Text

         Case "Rankine"
            txtResult.Text = (9 / 5) * dblIn + 491.69

         Case "Kelvin"
            txtResult.Text = dblIn + 273.15
         End Select

      Case "Rankine"
         Select Case lstOutput.Text
         Case "Fahrenheit"
            txtResult.Text = dblIn - 459.69

         Case "Celsius"
            txtResult.Text = (5 / 9) * (dblIn - 491.69)

         Case "Rankine"
            txtResult.Text = txtInput.Text

         Case "Kelvin"
            dblIn = dblIn - 459.69          '// change to F
            dblIn = (5 / 9) * (dblIn - 32)  '// change F to C
            txtResult.Text = dblIn + 273.15 '// change C to K
         End Select

      Case "Kelvin"
         Select Case lstOutput.Text
         Case "Fahrenheit"
            txtResult.Text = (9 / 5) * (dblIn - 273.15) + 32

         Case "Celsius"
            txtResult.Text = dblIn - 273.15

         Case "Rankine"
            dblIn = dblIn - 273.15          '// change to C
            dblIn = (9 / 5) * dblIn + 32    '// change C to F
            txtResult.Text = dblIn + 459.69 '// change F to R

         Case "Kelvin"
            txtResult.Text = txtInput.Text
         End Select

      End Select

   Case Else
      Call ShowResult(dblIn)
   End Select

End Sub

Private Sub ShowResult(ByVal vInput As Double)
   If mintFixDec Then
      txtResult.Text = Round((mdblOutput / mdblInput) * vInput, mintFixDec)
   Else
      txtResult.Text = (mdblOutput / mdblInput) * vInput
   End If
End Sub

Private Sub cmdAbout_Click()

   frmAbout.Show , Me

End Sub

Private Sub cmdCopy_Click()

   '// Clear Clipboard.
   Clipboard.Clear
   '// Put text on Clipboard.
   Clipboard.SetText txtResult.Text

End Sub

Private Sub cmdExit_Click()

   PopupMenu mnuFile, , cmdExit.left, cmdExit.top + cmdExit.Height

End Sub

Private Sub cmdFixed_Click()

   PopupMenu mnuFixDec, , cmdFixed.left, cmdFixed.top + cmdFixed.Height

End Sub

Private Sub cmdInvert_Click()

  Dim lngI As Long

   lngI = lstInput.ListIndex
   lstInput.ListIndex = lstOutput.ListIndex
   lstOutput.ListIndex = lngI

End Sub

Private Sub Form_Load()

  Dim lngI    As Long
  Dim strTemp As String

   On Error Resume Next
   
   'Call ManifestWrite
   Call CenterForm(Me)

   Set gcINI = New clsGetPutINI

   Do
      strTemp = gcINI.GetSetting("GroupNames", CStr(lngI))

      If LenB(strTemp) Then
         lstGroup.AddItem strTemp
         lngI = lngI + 1
      Else
         Exit Do
      End If

   Loop Until LenB(strTemp) = 0

   txtInput.Text = GetSetting(App.Title, "LastUsed", "Value")
   mintFixDec = rVal(GetSetting(App.Title, "LastUsed", "FixDec", 0))
   lstGroup.ListIndex = rVal(GetSetting(App.Title, "LastUsed", "Group", 0))
   lstInput.ListIndex = rVal(GetSetting(App.Title, "LastUsed", "Input", 0))
   lstOutput.ListIndex = rVal(GetSetting(App.Title, "LastUsed", "Output", 0))
   mlngLastG = lstGroup.ListIndex

   If mintFixDec Then
      cmdFixed.Caption = C_MenuTitle & CStr(mintFixDec)
   Else
      cmdFixed.Caption = C_MenuTitle & "none"
   End If

   fraToolBar.BackColor = cmdExit.BackColor
   mblnLoaded = True
   Call Calculate

   Unload frmAbout

End Sub

Private Sub Form_Unload(Cancel As Integer)

   SaveSetting App.Title, "LastUsed", "Group", CStr(lstGroup.ListIndex)
   SaveSetting App.Title, "LastUsed", "Input", CStr(lstInput.ListIndex)
   SaveSetting App.Title, "LastUsed", "Output", CStr(lstOutput.ListIndex)
   SaveSetting App.Title, "LastUsed", "Value", txtInput.Text
   SaveSetting App.Title, "LastUsed", "FixDec", CStr(mintFixDec)
   
   SaveSetting App.Title, CStr(lstGroup.ListIndex), "Input", CStr(lstInput.ListIndex)
   SaveSetting App.Title, CStr(lstGroup.ListIndex), "Output", CStr(lstOutput.ListIndex)
   SaveSetting App.Title, CStr(lstGroup.ListIndex), "Value", txtInput.Text

   Set gcINI = Nothing
   Call EndApp(Me)
   Set frmMain = Nothing

End Sub

Private Sub lstGroup_Click()

  Dim lngI    As Long
  Dim strTemp As String

   On Error Resume Next

   SaveSetting App.Title, CStr(mlngLastG), "Input", CStr(lstInput.ListIndex)
   SaveSetting App.Title, CStr(mlngLastG), "Output", CStr(lstOutput.ListIndex)
   SaveSetting App.Title, CStr(mlngLastG), "Value", txtInput.Text
   mlngLastG = lstGroup.ListIndex
   
   lstInput.Clear
   lstOutput.Clear

   If Not (lstGroup.Text = C_Custom) Then
      
      Do
         strTemp = gcINI.GetSetting(lstGroup.Text, "D" & CStr(lngI))
   
         If LenB(strTemp) Then
            lstInput.AddItem strTemp
            lstInput.ItemData(lstInput.NewIndex) = lngI
   
            lstOutput.AddItem strTemp
            lstOutput.ItemData(lstOutput.NewIndex) = lngI
   
            lngI = lngI + 1
         Else
            Exit Do
         End If
   
      Loop Until LenB(strTemp) = 0
   
   Else
      
      For lngI = 1 To 20
         
         strTemp = gcINI.GetSetting(C_Custom & Format$(lngI, "00"), "D0")
         If LenB(strTemp) Then
            lstInput.AddItem strTemp
            lstInput.ItemData(lstInput.NewIndex) = CLng(CStr(lngI) & "0")
         End If
         
         strTemp = gcINI.GetSetting(C_Custom & Format$(lngI, "00"), "D1")
         If LenB(strTemp) Then
            lstInput.AddItem strTemp
            lstInput.ItemData(lstInput.NewIndex) = CLng(CStr(lngI) & "1")
         End If
      
      Next lngI
   
   End If
   
   lstInput.ListIndex = rVal(GetSetting(App.Title, CStr(lstGroup.ListIndex), "Input", 0))
   lstOutput.ListIndex = rVal(GetSetting(App.Title, CStr(lstGroup.ListIndex), "Output", 0))
   txtInput.Text = GetSetting(App.Title, CStr(lstGroup.ListIndex), "Value")
   
End Sub

Private Sub lstInput_Click()

  Dim strGIndex As String
  Dim strIndex  As String
  Dim strTemp   As String
  
   If lstInput.ListIndex >= 0 Then
      
      If lstGroup.Text = C_Custom Then
         cmdInvert.Enabled = False
         
         lstOutput.Clear
         strIndex = Format$(lstInput.ItemData(lstInput.ListIndex), "000")
         strGIndex = left$(strIndex, 2)
         strIndex = right$(strIndex, 1)
         
         strTemp = gcINI.GetSetting(C_Custom & strGIndex, "D0")
         lstOutput.AddItem strTemp
         lstOutput.ItemData(lstOutput.NewIndex) = CLng(strGIndex & "0")
      
         strTemp = gcINI.GetSetting(C_Custom & strGIndex, "D1")
         lstOutput.AddItem strTemp
         lstOutput.ItemData(lstOutput.NewIndex) = CLng(strGIndex & "1")
      
         lblInputUnits.Caption = gcINI.GetSetting(lstGroup.Text & strGIndex, "D" & CStr(strIndex))
         'lstInput.ToolTipText = lblInputUnits.Caption
         mdblInput = CDbl(gcINI.GetSetting(lstGroup.Text & strGIndex, CStr(strIndex)))
         
         lstOutput.ListIndex = 0
      
      Else
         cmdInvert.Enabled = True
         lblInputUnits.Caption = gcINI.GetSetting(lstGroup.Text, "D" & CStr(lstInput.ItemData(lstInput.ListIndex)))
         'lstInput.ToolTipText = lblInputUnits.Caption
         mdblInput = rVal(gcINI.GetSetting(lstGroup.Text, CStr(lstInput.ItemData(lstInput.ListIndex))))
      End If
      
      Call Calculate
      
   End If
   
End Sub

Private Sub lstOutput_Click()

  Dim strGIndex As String
  Dim strIndex  As String
  
   If lstOutput.ListIndex >= 0 Then
      
      If lstGroup.Text = C_Custom Then
         strIndex = Format$(lstOutput.ItemData(lstOutput.ListIndex), "000")
         strGIndex = left$(strIndex, 2)
         strIndex = right$(strIndex, 1)
         
         lblOutputUnits.Caption = gcINI.GetSetting(lstGroup.Text & strGIndex, "D" & CStr(strIndex))
         'lstOutput.ToolTipText = lblOutputUnits.Caption
         mdblOutput = CDbl(gcINI.GetSetting(lstGroup.Text & strGIndex, CStr(strIndex)))
      
      Else
         lblOutputUnits.Caption = gcINI.GetSetting(lstGroup.Text, "D" & CStr(lstOutput.ItemData(lstOutput.ListIndex)))
         'lstOutput.ToolTipText = lblOutputUnits.Caption
         mdblOutput = CDbl(gcINI.GetSetting(lstGroup.Text, CStr(lstOutput.ItemData(lstOutput.ListIndex))))
      End If
      
      Call Calculate
   End If
   
End Sub

Private Sub mnuAddNew_Click()
   
   If Not (lstGroup.Text = "Base") Then
      If Not (lstGroup.Text = "Temperature") Then
         If Not (lstGroup.Text = C_Custom) Then
            
            frmAddNew.Show vbModal, Me
            If frmAddNew.ItemAdded Then
               Call lstGroup_Click '// refresh list boxes
            End If
   
            Unload frmAddNew
         End If
      End If
   End If

End Sub

Private Sub mnuCustom_Click()

   frmCustom.Show vbModal, Me
   Call lstGroup_Click '// refresh list boxes

End Sub

Private Sub mnuD_Click(Index As Integer)

   mintFixDec = Index

   If mintFixDec < 11 Then
      cmdFixed.Caption = C_MenuTitle & CStr(mintFixDec)

   Else
      mintFixDec = 0
      cmdFixed.Caption = C_MenuTitle & "none"
   End If

   Call Calculate

End Sub

Private Sub mnuEdit_Click()

   If Not (lstGroup.Text = C_Custom) Then
      If lstInput.ItemData(lstInput.ListIndex) > 0 Then
         If Not (lstGroup.Text = "Base") Then
            If Not (lstGroup.Text = "Temperature") Then
            
               With frmEdit
                  .Show vbModal, Me
                  If .ItemDeleted Then
                     Call lstGroup_Click '// refresh list boxes
                  End If
               End With
   
               Unload frmEdit
            End If
         End If
      End If
   End If

End Sub

Private Sub mnuExit_Click()

   Unload Me

End Sub

Private Sub txtInput_Change()

   Call Calculate

End Sub

