Attribute VB_Name = "General"
' RegOCX.frm
'
' By Herman Liu
'
' To register or unregister OCX/DLL controls (1) With prior confirmation of the
' Product Version and (2) Without using Regsvr32.exe.  This code has the advantages
' of (a) user verification of the specific version being registered; (b) being free
' from the existence of a Regsvr32.exe file, speedier and a better error handling.
'
' Output: Entries entered into/removed from HKEY_CLASSES_ROOT in the registry.
'
Option Explicit

Public Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" _
    (ByVal lptstrFilename As String, lpdwHandle As Long) As Long

Public Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" _
    (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwlen As Long, lpdata As Any) As Long

Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" _
  (ByVal lpLibFileName As String) As Long
  
Public Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, _
    ByVal lpProcName As String) As Long

Public Declare Function CreateThread Lib "kernel32" (lpThreadAttributes As Any, _
   ByVal dwStackSize As Long, ByVal lpStartAddress As Long, ByVal lParameter As Long, _
   ByVal dwCreationFlags As Long, lpThreadID As Long) As Long
   
'Private Declare Function TerminateThread Lib "kernel32" (ByVal hThread As Long, _
   ByVal dwExitCode As Long) As Long
   
Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, _
    ByVal dwMilliseconds As Long) As Long
   
Public Declare Function GetExitCodeThread Lib "kernel32" (ByVal hThread As Long, _
    lpExitCode As Long) As Long

Public Declare Sub ExitThread Lib "kernel32" (ByVal dwExitCode As Long)

Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public ConName As String
Public mCompanyName As String
Public mProductVersion As String
Public RegFlag As Boolean
Public UnregFlag As Boolean
Public mresult
Public gcdg As Object
Public Dropped As Boolean
Public ReX As Integer
Public DoClear As Boolean
Public NotQueried As Boolean



Function IsFileThere(inFileSpec As String) As Boolean
    On Error Resume Next
    Dim i
    i = FreeFile
    Open inFileSpec For Input As i
    If Err Then
        IsFileThere = False
    Else
        Close i
        IsFileThere = True
    End If
End Function

