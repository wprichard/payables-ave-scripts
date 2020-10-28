Attribute VB_Name = "modGlobals"
'*****************************************************************************************
' Module Name:  modMutex
' Description:  Provides declarations for using a Mutex
'
' Edit History:
' mm/dd/yyyy - Modified by name, company
'   Description of change
'
'*****************************************************************************************
'Public Declare Function OpenMutex Lib "kernel32" Alias "OpenMutexA" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal lpName As String) As Long
Public Declare Function CreateMutex Lib "kernel32" Alias "CreateMutexA" (lpMutexAttributes As Long, ByVal bInitialOwner As Long, ByVal lpName As String) As Long
Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Public Declare Function ReleaseMutex Lib "kernel32" (ByVal hMutex As Long) As Long
Public Declare Function GetLastError Lib "kernel32" () As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
 
'Public gobjLog As cLog

'From WINNT.H
'#define STATUS_WAIT_0                    ((DWORD   )0x00000000L)
Const STATUS_WAIT_0 = &H0

'#define STATUS_ABANDONED_WAIT_0          ((DWORD   )0x00000080L)
Const STATUS_ABANDONED_WAIT_0 = &H80

'#define STATUS_TIMEOUT                   ((DWORD   )0x00000102L)
Const STATUS_TIMEOUT = &H102

'Values returned by WaitForSingleObject
'From winbase.h
'#define WAIT_FAILED (DWORD)0xFFFFFFFF
Public Const WAIT_FAILED = &HFFFFFFFF

'#define WAIT_OBJECT_0       ((STATUS_WAIT_0 ) + 0 )
Public Const WAIT_OBJECT_0 = ((STATUS_WAIT_0) + 0)

'#define WAIT_ABANDONED         ((STATUS_ABANDONED_WAIT_0 ) + 0 )
Public Const WAIT_ABANDONED = ((STATUS_ABANDONED_WAIT_0) + 0)

'#define WAIT_TIMEOUT                        STATUS_TIMEOUT
Public Const WAIT_TIMEOUT = STATUS_TIMEOUT

'#define INFINITE            0xFFFFFFFF  // Infinite timeout


