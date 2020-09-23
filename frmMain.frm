VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "System Registration Utility"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   3870
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkProcess 
      Caption         =   "Process when dropped"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3600
      Width           =   2295
   End
   Begin VB.TextBox Status 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   2880
      Width           =   6015
   End
   Begin VB.OptionButton optUnregister 
      Caption         =   "Unregister"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4800
      TabIndex        =   3
      Top             =   3600
      Width           =   1215
   End
   Begin VB.OptionButton optRegister 
      Caption         =   "Register"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   2
      Top             =   3600
      Width           =   1095
   End
   Begin VB.OptionButton optQuery 
      Caption         =   "Query"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   1
      Top             =   3600
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.ListBox FileList 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2790
      Left            =   0
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Top             =   0
      Width           =   6015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub chkProcess_Click()
If DoClear And NotQueried Then FileList.Clear

End Sub

Private Sub FileList_Click()
If FileList.ListIndex <> -1 Then
    ReX = FileList.ListIndex
    ConName = FileList.List(FileList.ListIndex)
    If Trim(FileList.List(FileList.ListIndex)) = "" Then
        Status.Text = "No path and filename yet"
        Exit Sub
    ElseIf IsFileThere(FileList.List(FileList.ListIndex)) = False Then
        Status.Text = "File not found"
        Exit Sub
    End If
    DispProdVersion FileList.List(FileList.ListIndex)

End If
End Sub

Private Sub FileList_DblClick()
If FileList.ListIndex <> -1 Then
    ReX = FileList.ListIndex
    ConName = FileList.List(FileList.ListIndex)
    If Trim(FileList.List(FileList.ListIndex)) = "" Then
        Status.Text = "No path and filename yet"
        Exit Sub
    ElseIf IsFileThere(FileList.List(FileList.ListIndex)) = False Then
        Status.Text = "File not found"
        Exit Sub
    End If
    If optQuery.Value Then
        DispProdVersion FileList.List(FileList.ListIndex)
    End If
    
    If optRegister.Value Then
        DispProdVersion FileList.List(FileList.ListIndex)
        RegUnReg FileList.List(FileList.ListIndex)
    End If
    If optUnregister.Value Then
        DispProdVersion FileList.List(FileList.ListIndex)
        RegUnReg FileList.List(FileList.ListIndex)
    End If

End If

End Sub

Private Sub FileList_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ReX As Integer
Dim XErr
Dropped = True
FileList.Clear
XErr = 0
For ReX = 1 To Data.Files.Count
    'FileList.AddItem Data.Files(ReX), 0
    If chkProcess.Value <> 0 Then
    ConName = Data.Files(ReX)
    If Trim(Data.Files(ReX)) = "" Then
        Status.Text = "No path and filename yet"
        Exit Sub
    ElseIf IsFileThere(Data.Files(ReX)) = False Then
        Status.Text = "File not found"
        Exit Sub
    End If
    If optQuery.Value Then
        FileList.AddItem Data.Files(ReX), 0
        DispProdVersion Data.Files(ReX)
    End If
    
    If optRegister.Value Then
        DispProdVersion Data.Files(ReX)
        RegUnReg Data.Files(ReX)
    End If
    If optUnregister.Value Then
        DispProdVersion Data.Files(ReX)
        RegUnReg Data.Files(ReX), "u"
    End If
    
    Else
        FileList.AddItem Data.Files(ReX), FileList.ListCount

    End If

Next ReX
Dropped = False
End Sub

Sub ErrMsgProc(mMsg As String)
    Status.Text = mMsg & vbCrLf & Err.Number & Space(5) & Err.Description
End Sub
Private Function GetFileInfo(inFileSpec As String) As Boolean
    On Error Resume Next
    Dim lInfoSize As Long
    Dim lpHandle As Long
    Dim strFileInfoString As String
    Dim i As Integer
    
    GetFileInfo = False                                ' Assume
    
     ' GetFileVersionInfoSize determines if system can obtain version info
     ' about the specified file.  If yes, it returns its size in bytes and
     ' a handle to the data.
    lpHandle = 0
    lInfoSize = GetFileVersionInfoSize(inFileSpec, lpHandle)
    If lInfoSize = 0 Then
        Exit Function
    End If

     ' We pass the file name, size(ignored), size of buffer and the buffer of
     ' version info to GetFileVersionInfo, which will fill the buffer with
     ' version info about the file. (Modified here).
    strFileInfoString = String(lInfoSize, 0)
    mresult = GetFileVersionInfo(ByVal inFileSpec, 0&, ByVal lInfoSize, _
          ByVal strFileInfoString)
    If mresult = 0 Then
        Exit Function
    End If

     ' We now have a block of version data, in an unreadable format though. If you
     ' wish, you may check the existence of "StringFileInfo" with InStr function.
     ' Normally we must call VerQueryValue to read selected pieces of data of the
     ' above, with arguments such as "\VarFileInfo\Translation" or "\StringFileInfo
     ' \lang-codepage\string-name" where lang-codepage is a code which has yet to be
     ' obtained from first 2 words(high-low) returned by "\VarFileInfo\Translation"
     ' from the strFileInfoString (and padded to fixed 8-digit), and string-name is
     ' one of predefined string names such as "CompanyName" & "FileDescription", etc.
     ' However, the following simple alternative is OK for our purpose.

     mCompanyName = ""
     mProductVersion = ""
     i = InStr(strFileInfoString, "CompanyName")
     If i > 0 Then
         i = i + 12
         mCompanyName = Mid$(strFileInfoString, i, 21)
     End If
     i = InStr(strFileInfoString, "FileDescription")
     If i > 0 Then
         i = i + 16
     End If
     i = InStr(strFileInfoString, "FileVersion")
     If i > 0 Then
         i = i + 12
     End If
     i = InStr(strFileInfoString, "InternalName")
     If i > 0 Then
         i = i + 16
     End If
     i = InStr(strFileInfoString, "LegalCopyright")
     If i > 0 Then
         i = i + 16
     End If
     i = InStr(strFileInfoString, "OriginalFilename")
     If i > 0 Then
         i = i + 20
     End If
     i = InStr(strFileInfoString, "ProductName")
     If i > 0 Then
         i = i + 12
     End If
     i = InStr(strFileInfoString, "ProductVersion")
     If i > 0 Then
         i = i + 16
         mProductVersion = Mid$(strFileInfoString, i)
     End If

     If Trim(mProductVersion) <> "" Then
         GetFileInfo = True
     End If
End Function
Private Sub RegUnReg(ByVal inFileSpec As String, Optional inHandle As String = "")
    On Error Resume Next
If optQuery.Value = False Then NotQueried = True
    Dim lLib As Long                 ' Store handle of the control library
    Dim lpDLLEntryPoint As Long      ' Store the address of function called
    Dim lpThreadID As Long           ' Pointer that receives the thread identifier
    Dim lpExitCode As Long           ' Exit code of GetExitCodeThread
    Dim mThread
    
      ' Load the control DLL, i. e. map the specified DLL file into the
      ' address space of the calling process
    lLib = LoadLibrary(inFileSpec)
    If lLib = 0 Then
         ' e.g. file not exists or not a valid DLL file
        If Dropped = True Then
            FileList.AddItem "Failed!: " & ConName, 0
        Else
            Status.Text = "Operation Failed!" & vbCrLf & ConName
        End If
        Exit Sub
    End If
    
      ' Find and store the DLL entry point, i.e. obtain the address of the
      ' “DllRegisterServer” or "DllUnregisterServer" function (to register
      ' or deregister the server’s components in the registry).
      '
    If inHandle = "" Then
        lpDLLEntryPoint = GetProcAddress(lLib, "DllRegisterServer")
    ElseIf inHandle = "U" Or inHandle = "u" Then
        lpDLLEntryPoint = GetProcAddress(lLib, "DllUnregisterServer")
    Else
        If Dropped = True Then
            FileList.AddItem "Unknown Handle!: " & ConName, 0
        Else
            Status.Text = "Unknown Command Handle!" & vbCrLf & ConName
        End If
        Exit Sub
    End If
    If lpDLLEntryPoint = vbNull Then
        GoTo earlyExit1
    End If
    
    Screen.MousePointer = vbHourglass
    
      ' Create a thread to execute within the virtual address space of the calling process
    mThread = CreateThread(ByVal 0, 0, ByVal lpDLLEntryPoint, ByVal 0, 0, lpThreadID)
    If mThread = 0 Then
        GoTo earlyExit1
    End If
    
      ' Use WaitForSingleObject to check the return state (i) when the specified object
      ' is in the signaled state or (ii) when the time-out interval elapses.  This
      ' function can be used to test Process and Thread.
    mresult = WaitForSingleObject(mThread, 10000)
    If mresult <> 0 Then
        GoTo earlyExit2
    End If
    
      ' We don't call the dangerous TerminateThread(); after the last handle
      ' to an object is closed, the object is removed from the system.
    CloseHandle mThread
    FreeLibrary lLib
    
    Screen.MousePointer = vbDefault
    Select Case LCase$(inHandle)
        Case ""
            If Dropped = True Then
                FileList.AddItem "Registered!: " & ConName, 0
            Else
                Status.Text = "Registered Successfully!" & vbCrLf & ConName
            End If
            
        Case "u"
            If Dropped = True Then
                FileList.AddItem "Unregistered!: " & ConName, 0
                DoClear = True
            Else
                Status.Text = "Unregistered Successfully!" & vbCrLf & ConName
                DoClear = False
            End If
    End Select
    Exit Sub
    
    
earlyExit1:
    Screen.MousePointer = vbDefault
    If Dropped = True Then
        FileList.AddItem "Failed! (Thread/EntryPoint): " & ConName, 0
    Else
        Status.Text = "Process failed in obtaining entry point or creating thread."
    End If

     ' Decrements the reference count of loaded DLL module before leaving
    FreeLibrary lLib
    Exit Sub
    
earlyExit2:
    Screen.MousePointer = vbDefault
    If Dropped = True Then
        FileList.AddItem "Failed! (State/Timeout): " & ConName, 0
    Else
        Status.Text = "Process failed in signaled state or time-out."
    End If
    FreeLibrary lLib
     ' Terminate the thread to free up resources that are used by the thread
     ' NB Calling ExitThread for an application's primary thread will cause
     ' the application to terminate
    lpExitCode = GetExitCodeThread(mThread, lpExitCode)
    ExitThread lpExitCode
End Sub

Private Sub DispProdVersion(inFile As String)
    If Not GetFileInfo(inFile) Then
        Status.Text = "(No Product Version available for this file)"
    Else
        Status.Text = "Company Name:  " & mCompanyName & vbCrLf & _
             "Product Version:  " & mProductVersion
    End If
End Sub

Private Sub optQuery_Click()
If DoClear And NotQueried Then FileList.Clear

End Sub

Private Sub optRegister_Click()
If DoClear And NotQueried Then FileList.Clear
End Sub

Private Sub optUnregister_Click()
If DoClear And NotQueried Then FileList.Clear

End Sub
