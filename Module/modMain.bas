Attribute VB_Name = "modMain"
'//************************************************************************************************
'// Author: Morgan Haueisen (morganh@hartcom.net)
'// Copyright (c) 2006
'// Version 1.0.1
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

Private Type Rect
   left   As Long
   top    As Long
   right  As Long
   bottom As Long
End Type

Private Declare Function SystemParametersInfo Lib "user32" _
      Alias "SystemParametersInfoA" ( _
      ByVal uAction As Long, _
      ByVal uParam As Long, _
      ByRef lpvParam As Rect, _
      ByVal fuWinIni As Long) As Long

Public Const C_Custom As String = "{Custom}"

Public gcINI As clsGetPutINI

Public Sub CenterForm(xForm As Form)

  Dim Rc As Rect
  Dim T  As Long
  Dim b  As Long
  Dim L  As Long
  Dim r  As Long
  Dim mT As Long
  Dim mL As Long

   Call SystemParametersInfo(48&, 0&, Rc, 0&)

   T = Rc.top * Screen.TwipsPerPixelY
   b = Rc.bottom * Screen.TwipsPerPixelY
   L = Rc.left * Screen.TwipsPerPixelX
   r = Rc.right * Screen.TwipsPerPixelX

   mT = Abs((b / 2.2) - (xForm.Height / 2))
   mL = Abs((r / 2) - (xForm.Width / 2))

   If mT < T Then mT = T
   If mT > b - xForm.Height Then mT = b - xForm.Height
   If mL < L Then mL = L

   xForm.Move mL, mT

End Sub

Public Function ConvertBase(ByVal vNumIn As String, _
                            ByVal vBaseIn As Integer, _
                            ByVal vBaseOut As Integer) As String

   '// Orignal By: Aidan
   '// http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=1474&lngWId=1

   '// Converts a number from one base to another
   '// Binary = Base 2
   '// Octal = Base 8
   '// Decimal = Base 10
   '// Hexadecimal = Base 16

   '// vNumIn is the number which you wish to convert (A String including characters 0 - 9, A - Z)
   '// vBaseIn is the base of vNumIn (An integer value in Decimal between 1 & 36)
   '// vBaseOut is the base of the number the function returns (An Integer value in Decimal between 1 & 36)
   '// Returns a string in the desired base containing the characters 0 - 9, A - Z)
   '// Example: ConvertBase ("42", 8, 16) converts the octal number 42 into hexadecimal 22
   '// Returns the word "Error" if any of the input values are incorrect

  Dim lngI                As Long
  Dim strCurrentCharacter As String
  Dim lngCharacterValue   As Long
  Dim lngPlaceValue       As Long
  Dim dblRunningTotal     As Double
  Dim dblRemainder        As Double
  Dim dblBaseOutDouble    As Double
  Dim strNumInCaps        As String
  Dim strResult           As String

   '// Ensure input data is valid
   If LenB(vNumIn) = 0 Then
      ConvertBase = "Error"
      Exit Function
   Else
      
      '// Text to Text
      If vBaseIn = 0 And vBaseOut = 0 Then
         ConvertBase = vNumIn
         Exit Function
      End If
      
      '// Text to Dec or Hex
      If vBaseIn = 0 And (vBaseOut = 10 Or vBaseOut = 16) Then
         If vBaseOut = 10 Then
            ConvertBase = Asc(vNumIn)
         Else
            ConvertBase = Hex(Asc(vNumIn))
         End If
         Exit Function
      End If
         
      '// Dec or Hex to Text
      If rVal(vNumIn) > 0 Then
         If vBaseOut = 0 And (vBaseIn = 10 Or vBaseIn = 16) Then
            If vBaseIn = 10 Then
               ConvertBase = Chr$(CInt(vNumIn))
            Else
               ConvertBase = Chr$(CInt("&H" & vNumIn))
            End If
            Exit Function
         End If
      End If
      
   End If

   '// Ensure input data is valid
   If vBaseIn < 2 Or vBaseIn > 36 Or vBaseOut < 2 Or vBaseOut > 36 Then
      ConvertBase = "Error"
      Exit Function
   End If
   
   '// Ensure any letters in the input mumber are capitals
   strNumInCaps = UCase$(vNumIn)
   '// Convert strNumInCaps into Decimal
   lngPlaceValue = Len(strNumInCaps)

   For lngI = 1 To Len(strNumInCaps)
      lngPlaceValue = lngPlaceValue - 1
      strCurrentCharacter = Mid$(strNumInCaps, lngI, 1)
      lngCharacterValue = 0

      If Asc(strCurrentCharacter) > 64 And Asc(strCurrentCharacter) < 91 Then
         lngCharacterValue = Asc(strCurrentCharacter) - 55
      End If

      If lngCharacterValue = 0 Then
         
         '// Ensure vNumIn is correct
         If Asc(strCurrentCharacter) < 48 Or Asc(strCurrentCharacter) > 57 Then
            ConvertBase = "Error"
            Exit Function
         Else
            lngCharacterValue = rVal(strCurrentCharacter)
         End If

      End If

      If lngCharacterValue < 0 Or lngCharacterValue > vBaseIn - 1 Then
         '// Ensure vNumIn is correct
         ConvertBase = "Error"
         Exit Function
      End If

      dblRunningTotal = dblRunningTotal + lngCharacterValue * (vBaseIn ^ lngPlaceValue)
   Next lngI

   '// Convert Decimal Number into the desired base using Repeated Division
   Do
      dblBaseOutDouble = CDbl(vBaseOut)
      dblRemainder = dblRunningTotal - (Int(dblRunningTotal / dblBaseOutDouble) * dblBaseOutDouble)
      dblRunningTotal = (dblRunningTotal - dblRemainder) / vBaseOut

      If dblRemainder >= 10 Then
         strCurrentCharacter = Chr$(dblRemainder + 55)
      Else
         strCurrentCharacter = CStr(dblRemainder) ''right$(Str$(dblRemainder), Len(Str$(dblRemainder)) - 1)
      End If

      strResult = strCurrentCharacter & strResult
   Loop While dblRunningTotal > 0

   ConvertBase = strResult

End Function

Public Sub Main()

   Call IsAppRunning
   
   frmAbout.Show
   DoEvents

   If LenB(Dir$(App.Path & "\" & App.Title & ".dat")) = 0 Then
      MsgBox "ERROR! Unable to locate " & App.Title & ".dat", vbCritical
      Unload frmAbout
      End
   Else
      frmMain.Show
   End If
   
End Sub

Public Function rVal(ByVal vString As String) As Double

'// Returns the numbers contained in a string as a numeric value
   '// VB's Val function recognizes only the period (.) as a valid decimal separator.
   '// VB's CDbl errors on empty strings or values containing non-numeric values

  Dim lngI     As Long
  Dim lngS     As Long
  Dim bytAscV  As Byte
  Dim strTemp  As String
  
  On Error Resume Next

   vString = Trim$(UCase$(vString))
   If LenB(vString) Then
   
      Select Case left$(vString, 2)          '// Hex or Octal?
      Case Is = "&H", Is = "&O"
         lngS = 3
         strTemp = left$(vString, 2)
      Case Else
         lngS = 1
      End Select
      
      For lngI = lngS To Len(vString)
         bytAscV = Asc(Mid$(vString, lngI, 1))
         Select Case bytAscV
         Case 48 To 57 '// 1234567890
            strTemp = strTemp & Mid$(vString, lngI, 1)
         
         Case 44, 45, 46 '// , - .
            strTemp = strTemp & Mid$(vString, lngI, 1)
         
         Case 36, 163, 32 '// $
            '// Ignore
            
         Case Is > 57, Is < 44
            If left$(strTemp, 2) = "&H" Then '// Hex Values ?
               Select Case bytAscV
               Case 65 To 70 '// ABCDEF
                  strTemp = strTemp & Mid$(vString, lngI, 1)
               Case Else
                  Exit For
               End Select
            ElseIf bytAscV = 69 Then
               strTemp = strTemp & Mid$(vString, lngI, 1)
            Else
               Exit For
            End If
         End Select
      Next lngI
      
      If LenB(strTemp) Then
         rVal = CDbl(strTemp)
         If rVal = 0 Then
            strTemp = Replace$(strTemp, ".", ",")
            rVal = CDbl(strTemp)
         End If
      
      Else '// Check for boolean text (True or False)
         '// VB's CBool errors on empty or invalid strings (not True or False)
         '// Check for valid boolean
         rVal = CBool(vString)
      End If
   
   Else
      rVal = 0
   End If
   
Exit_Here:
   On Error GoTo 0
End Function

