Attribute VB_Name = "modStartEnd"
'// Author: Morgan Haueisen (morganh@hartcom.net)
'// Copyright (c) 2005

Option Explicit

Private Declare Function SetErrorMode Lib "kernel32" (ByVal wMode As Long) As Long
Private Declare Function CreateMutex Lib "kernel32" _
      Alias "CreateMutexA" ( _
      ByRef lpMutexAttributes As Any, _
      ByVal bInitialOwner As Long, _
      ByVal lpName As String) As Long
Private Declare Function ReleaseMutex Lib "kernel32" (ByVal hMutex As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
    
Private mlngMutex As Long '// It will store a

Public Sub EndApp(Optional CallingForm As Form)

   '// Call from the closing Form's Form_Unload event
   '//
   '// Example:
   '//   Call EndApp(Me)

  Dim Frm As Form
  Const SEM_NOGPFAULTERRORBOX As Long = &H2&
  
   On Error Resume Next

   '// free memory
   If mlngMutex Then
      Call ReleaseMutex(mlngMutex)
      Call CloseHandle(mlngMutex)
   End If

   '// Close all open Forms
   For Each Frm In Forms

      If Frm.Name <> CallingForm.Name Then
         Unload Frm
         Set Frm = Nothing
      End If

   Next Frm

   '// Some versions of ComCtl32.DLL version 6.0 cause a crash at shutdown
   '// when you enable XP Visual Styles in an application that has a VB User Control.
   '// This instructs Windows to not display the UAE message box that invites you to send
   '// Microsoft information about the problem.
   If CBool(VB.App.LogMode()) Then '// Not running in IDE
      Call SetErrorMode(SEM_NOGPFAULTERRORBOX)
   End If

End Sub

Public Sub IsAppRunning()
    
  ''' Const ERROR_ALREADY_EXISTS = 183&
    
   If Not IsInIDE Then '// Ignore if running within IDE
      
      '// Is this application already open?
      '// (If it is open then end program)
      
      mlngMutex = CreateMutex(ByVal 0&, 1, App.Title)
      
      If (Err.LastDllError = 183&) Then
         '// free memory
         Call ReleaseMutex(mlngMutex)
         Call CloseHandle(mlngMutex)
         MsgBox App.Title & " is already running.", vbExclamation
         Call EndApp
         End
      End If
      
   End If
   
End Sub

Public Function IsInIDE() As Boolean

   '// Return whether we're running in the IDE.
   
   '// Assert invocations work only within the development environment and
   '// conditionally suspends execution (if set to False) at the line on which
   '// the method appears.
   '// When the module is compiled into an executable, the method calls on the
   '// Debug object are omitted.
  
   Debug.Assert zSetTrue(IsInIDE)

End Function

Private Function zSetTrue(ByRef bValue As Boolean) As Boolean

   '// Worker function for IsInIDE
   zSetTrue = True
   bValue = True

End Function
