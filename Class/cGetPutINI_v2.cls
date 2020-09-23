VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGetPutINI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'/**************************************/
'/*     Author: Morgan Haueisen        */
'/*             morganh@hartcom.net    */
'/*     Copyright (c) 1996-2006        */
'/*                                    */
'/* Write/Read from an INI file        */
'/**************************************/

Option Explicit

Private Declare Function GetPrivateProfileString Lib "kernel32" _
      Alias "GetPrivateProfileStringA" ( _
      ByVal lpAppName As String, _
      ByVal lpKeyName As String, _
      ByVal lpDefault As String, _
      ByVal lpReturnedString As String, _
      ByVal nSize As Long, _
      ByVal Filename As String) As Long

Private Declare Function WritePrivateProfileString Lib "kernel32.dll" _
      Alias "WritePrivateProfileStringA" ( _
      ByVal lpAppName As String, _
      ByVal lpKeyName As String, _
      ByVal lpDefault As String, _
      ByVal Filename As String) As Long

Private mstrINIFile As String

Private Sub Class_Initialize()

   mstrINIFile = App.Path & "\" & App.Title & ".dat"

End Sub

Public Function GetSetting(ByVal vSection As String, _
                           ByVal vKey As String, _
                           Optional ByVal vDefaultValue As String = vbNullString) As String

  Const C_BufferLen As Long = 255&   'Max Size of All Init
  Dim strBuffer     As String * C_BufferLen
  Dim lngR          As Long

   On Local Error Resume Next
   lngR = GetPrivateProfileString(vSection, vKey, vbNullString, strBuffer, C_BufferLen, mstrINIFile)

   If lngR Then
      GetSetting = Left$(strBuffer, lngR)
   Else
      GetSetting = vDefaultValue
   End If

   On Local Error GoTo 0

End Function

Public Property Let INI_FileName(ByVal vNewValue As String)

   mstrINIFile = vNewValue

End Property

Public Property Get INI_FileName() As String

   INI_FileName = mstrINIFile

End Property

Public Sub SaveSetting(ByVal vSection As String, ByVal vKey As String, ByVal vNewValue As String)

   On Local Error Resume Next
   Call WritePrivateProfileString(vSection, vKey, vNewValue, mstrINIFile)
   On Local Error GoTo 0

End Sub

