Attribute VB_Name = "MShell32"
' ****************************************************************
'  Shell32.Bas, Copyright ©1996-97 Karl E. Peterson
' ****************************************************************
'  You are free to use this code within your own applications, but you
'  are expressly forbidden FROM selling or otherwise distributing this
'  source code without prior written consent.
' ****************************************************************
'  Three methods to "Shell AND Wait" under Win32.
'  One deals with the infamous "Finished" behavior of Win95.
'  Fourth method that simply shells AND returns top-level hWnd.
' ****************************************************************
Option Explicit

Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long

Private Const STILL_ACTIVE = &H103
Private Const PROCESS_QUERY_INFORMATION = &H400
Private Const SYNCHRONIZE = &H100000

Public Const WAIT_FAILED = -1&        'Error on call
Public Const WAIT_OBJECT_0 = 0        'Normal completion
Public Const WAIT_ABANDONED = &H80&   '
Public Const WAIT_TIMEOUT = &H102&    'Timeout period elapsed
Public Const IGNORE = 0               'Ignore signal
Public Const INFINITE = -1&           'Infinite timeout


Private Const WM_CLOSE = &H10
Private Const GW_HWNDNEXT = 2
Private Const GW_OWNER = 4




Private Type SHELLEXECUTEINFO
    cbSize        As Long
    fMask         As Long
    hWnd          As Long
    lpVerb        As String
    lpFile        As String
    lpParameters  As String
    lpDirectory   As String
    nShow         As Long
    hInstApp      As Long
    lpIDList      As Long
    lpClass       As String
    hkeyClass     As Long
    dwHotKey      As Long
    hIcon         As Long
    hProcess      As Long
End Type

Private Const SEE_MASK_INVOKEIDLIST = &HC
Private Const SEE_MASK_NOCLOSEPROCESS = &H40
Private Const SEE_MASK_FLAG_NO_UI = &H400
Private Const FO_DELETE = &H3
Private Const FOF_ALLOWUNDO = &H40
Private Const MAX_PATH = 260
Private Declare Function ShellExecuteEx Lib "shell32.dll" (sei As SHELLEXECUTEINFO) As Long



Public Const STARTF_USESHOWWINDOW = &H1
Public Enum enSW
    SW_HIDE = 0
    SW_NORMAL = 1
    SW_SHOWNORMAL = 1
    SW_SHOWMINIMIZED = 2
    SW_SHOWMAXIMIZED = 3
    SW_MAXIMIZE = 3
    SW_SHOWNOACTIVATE = 4
    SW_SHOW = 5
    SW_MINIMIZE = 6
    SW_SHOWMINNOACTIVE = 7
    SW_SHOWNA = 8
    SW_RESTORE = 9
End Enum

Public Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
End Type
Public Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Byte
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type
Public Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type
Public Enum enPriority_Class
    NORMAL_PRIORITY_CLASS = &H20
    IDLE_PRIORITY_CLASS = &H40
    HIGH_PRIORITY_CLASS = &H80
End Enum
Public Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As SECURITY_ATTRIBUTES, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long


Public Function ShellEx(ByVal PathName As String, ByVal WorkDir As String, Milliseconds As Long, ByVal WindowStyle As enSW, ByVal Priority As enPriority_Class) As Boolean
    
    Dim pclass As Long
    Dim sinfo As STARTUPINFO
    Dim pinfo As PROCESS_INFORMATION
    Dim sec1 As SECURITY_ATTRIBUTES
    Dim sec2 As SECURITY_ATTRIBUTES
    
    'Set the structure size
    sec1.nLength = Len(sec1)
    sec2.nLength = Len(sec2)
    sinfo.cb = Len(sinfo)
    
    'Set the flags
    sinfo.dwFlags = STARTF_USESHOWWINDOW
    
    'Set the window's startup position
    sinfo.wShowWindow = WindowStyle
    
    'Set the priority class
    pclass = Priority
    
    'Start the program
    If CreateProcess(vbNullString, PathName, sec1, sec2, False, pclass, 0&, WorkDir, sinfo, pinfo) Then
        'Wait
        WaitForSingleObject pinfo.hProcess, Milliseconds
        ShellEx = True
    Else
        ShellEx = False
    End If
End Function

Public Function ShellAndWait(ByVal JobToDo As String, Optional ExecMode As VbAppWinStyle, Optional TimeOut) As Long
   '
   ' Shells a new process AND waits for it to complete.
   ' Calling application is totally non-responsive while
   ' new process executes.
   '
   Dim ProcessID As Long
   Dim hProcess As Long
   Dim nRet As Long
   Const fdwAccess = SYNCHRONIZE

   If IsMissing(ExecMode) Then
      ExecMode = vbMinimizedNoFocus
   Else
      If ExecMode < vbHide Or ExecMode > vbMinimizedNoFocus Then
         ExecMode = vbMinimizedNoFocus
      End If
   End If

   On Error Resume Next
      ProcessID = Shell(JobToDo, CLng(ExecMode))
      If Err Then
         ShellAndWait = vbObjectError + Err.Number
         Exit Function
      End If
   On Error GoTo 0

   If IsMissing(TimeOut) Then
      TimeOut = INFINITE
   End If

   hProcess = OpenProcess(fdwAccess, False, ProcessID)
   nRet = WaitForSingleObject(hProcess, CLng(TimeOut))
   Call CloseHandle(hProcess)

   Select Case nRet
      Case WAIT_TIMEOUT: Debug.Print "Timed out!"
      Case WAIT_OBJECT_0: Debug.Print "Normal completion."
      Case WAIT_ABANDONED: Debug.Print "Wait Abandoned!"
      Case WAIT_FAILED: Debug.Print "Wait Error:"; Err.LastDllError
   End Select
   ShellAndWait = nRet
End Function

Public Function ShellAndLoop(ByVal JobToDo As String, Optional ExecMode As VbAppWinStyle) As Long
   '
   ' Shells a new process AND waits for it to complete.
   ' Calling application is responsive while new process
   ' executes. It will react to new events, though execution
   ' of the current thread will not continue.
   '
   Dim ProcessID As Long
   Dim hProcess As Long
   Dim nRet As Long
   Const fdwAccess = PROCESS_QUERY_INFORMATION

   If IsMissing(ExecMode) Then
      ExecMode = vbMinimizedNoFocus
   Else
      If ExecMode < vbHide Or ExecMode > vbMinimizedNoFocus Then
         ExecMode = vbMinimizedNoFocus
      End If
   End If

   On Error Resume Next
      ProcessID = Shell(JobToDo, CLng(ExecMode))
      If Err Then
         ShellAndLoop = vbObjectError + Err.Number
         Exit Function
      End If
   On Error GoTo 0

   hProcess = OpenProcess(fdwAccess, False, ProcessID)
   Do
      GetExitCodeProcess hProcess, nRet
      DoEvents
      Sleep 100
   
   Loop While nRet = STILL_ACTIVE
   Call CloseHandle(hProcess)

   ShellAndLoop = nRet
End Function

Public Function ShellAndClose(ByVal JobToDo As String, Optional ExecMode As VbAppWinStyle) As Long
   '
   ' Shells a new process AND waits for it to complete.
   ' Calling application is responsive while new process
   ' executes. It will react to new events, though execution
   ' of the current thread will not continue.
   '
   ' Will close a DOS box when Win95 doesn't. More overhead
   ' than ShellAndLoop but useful when needed.
   '
   Dim ProcessID As Long
   Dim PID As Long
   Dim hProcess As Long
   Dim hWndJob As Long
   Dim nRet As Long
   Dim TitleTmp As String
   Const fdwAccess = PROCESS_QUERY_INFORMATION

   If IsMissing(ExecMode) Then
      ExecMode = vbMinimizedNoFocus
   Else
      If ExecMode < vbHide Or ExecMode > vbMinimizedNoFocus Then
         ExecMode = vbMinimizedNoFocus
      End If
   End If

   On Error Resume Next
      ProcessID = Shell(JobToDo, CLng(ExecMode))
      If Err Then
         ShellAndClose = vbObjectError + Err.Number
         Exit Function
      End If
   On Error GoTo 0

   hWndJob = FindWindow(vbNullString, vbNullString)
   Do Until hWndJob = 0
      If GetParent(hWndJob) = 0 Then
         Call GetWindowThreadProcessId(hWndJob, PID)
         If PID = ProcessID Then Exit Do
      End If
      hWndJob = GetWindow(hWndJob, GW_HWNDNEXT)
   Loop

   hProcess = OpenProcess(fdwAccess, False, ProcessID)
   Do
      TitleTmp = Space(256)
      nRet = GetWindowText(hWndJob, TitleTmp, Len(TitleTmp))
      If nRet Then
         TitleTmp = UCase(Left(TitleTmp, nRet))
         If InStr(TitleTmp, "FINISHED") = 1 Then
            Call SendMessage(hWndJob, WM_CLOSE, 0, 0)
         End If
      End If

      GetExitCodeProcess hProcess, nRet
      DoEvents
      Sleep 100
   Loop While nRet = STILL_ACTIVE
   Call CloseHandle(hProcess)

   ShellAndClose = nRet
End Function

Public Function hWndShell(ByVal JobToDo As String, Optional ExecMode As VbAppWinStyle) As Long
   '
   ' Shells a new process AND returns the hWnd
   ' of its main window.
   '
   Dim ProcessID As Long
   Dim PID As Long
   Dim hProcess As Long
   Dim hWndJob As Long

   If IsMissing(ExecMode) Then
      ExecMode = vbMinimizedNoFocus
   Else
      If ExecMode < vbHide Or ExecMode > vbMinimizedNoFocus Then
         ExecMode = vbMinimizedNoFocus
      End If
   End If

   On Error Resume Next
      ProcessID = Shell(JobToDo, CLng(ExecMode))
      If Err Then
         hWndShell = 0
         Exit Function
      End If
   On Error GoTo 0

   hWndJob = FindWindow(vbNullString, vbNullString)
   Do While hWndJob <> 0
      If GetParent(hWndJob) = 0 Then
         Call GetWindowThreadProcessId(hWndJob, PID)
         If PID = ProcessID Then
            hWndShell = hWndJob
            Exit Do
         End If
      End If
      hWndJob = GetWindow(hWndJob, GW_HWNDNEXT)
   Loop
End Function




Public Function ShowFileProperties(FileName As String, OwnerHWnd As Long) As Long
    Dim sei As SHELLEXECUTEINFO
    With sei
        .cbSize = Len(sei)
        .fMask = SEE_MASK_NOCLOSEPROCESS Or SEE_MASK_INVOKEIDLIST Or SEE_MASK_FLAG_NO_UI
        .hWnd = OwnerHWnd
        .lpVerb = "properties"
        .lpFile = FileName
        .lpParameters = vbNullChar
        .lpDirectory = vbNullChar
        .nShow = 0
        .hInstApp = 0
        .lpIDList = 0
    End With
    ShellExecuteEx sei
    ShowFileProperties = sei.hInstApp
End Function
