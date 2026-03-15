VERSION 5.00
Begin VB.Form FTSObject 
   AutoRedraw      =   -1  'True
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Processing..."
   ClientHeight    =   1950
   ClientLeft      =   4530
   ClientTop       =   3870
   ClientWidth     =   5175
   Icon            =   "FTSObject.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   300
      Top             =   960
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      Caption         =   "Initializing..."
      Height          =   195
      Left            =   1080
      TabIndex        =   0
      Top             =   300
      Width           =   810
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   300
      Picture         =   "FTSObject.frx":000C
      Stretch         =   -1  'True
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "FTSObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mDataFolder  As String
Private mUserID      As String
Private mPswd        As String
Private mLaunchFile  As String
Private mStatusLog   As String
Private mWaitForFile As String
Private mTitle As String

Private mStarted As Boolean
Private mFoundIT As Boolean
Private mFinished As Boolean
Private mDelayTime As Long


Public Sub Run(DataFolder As String, UserID As String, Pswd As String, LaunchFile As String, Optional WaitForFile As String, Optional Title As String)

    mDataFolder = DataFolder
    mUserID = UserID
    mPswd = Pswd
    mLaunchFile = LaunchFile
    mWaitForFile = WaitForFile
    mTitle = Title
    mDelayTime = 0
    
    If Not FileExists(mLaunchFile) Then Err.Raise 53, , "File not found: " & mLaunchFile
        
    If FileExt(mLaunchFile) = "mac" Then
        mStatusLog = ForceExt(LaunchFile, "log")
        mStarted = False
        mFoundIT = False
        mFinished = False
        Me.Show vbModal
    Else
        Call RunTsObject
    End If
    
End Sub

Private Sub Form_Load()
    Me.Caption = mTitle
    Call IniGetForm(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call IniPutForm(Me)
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
    Dim iFile As Integer
    Static StartTime As Single
    
    'launch macro
    If Not mStarted Then
        Screen.MousePointer = vbHourglass
        Call Kill(mStatusLog)
        Call Kill(mWaitForFile)
        Call RunTsObject
        mStarted = True
        mFoundIT = False
    End If
    
    Call ReadStatusFile
    
    If mFinished Then
        'wait a second or 2 so user can see that last step completed
        mDelayTime = mDelayTime + 1
        If mDelayTime > 100 Then
            Unload Me
            Screen.MousePointer = vbDefault
        End If
    Else
        If Not mFoundIT Then
            If Dir(mWaitForFile) <> "" Then
                iFile = FreeFile
                Open mWaitForFile For Input Lock Read Write As iFile
                mFoundIT = Err.Number <> 0
                Close iFile
            End If
        Else
            If Dir(mWaitForFile) <> "" Then
                iFile = FreeFile
                Open mWaitForFile For Input Lock Read Write As iFile
                mFinished = Err.Number = 0
                Close iFile
                StartTime = Timer
            End If
        End If
    End If
    
End Sub





Private Sub RunTsObject()
    Dim TSobjPath     As String
    Dim BatchFileName As String
    Dim BatchFile     As Integer

    
    TSobjPath = RegGetKey(HKEY_LOCAL_MACHINE, "SOFTWARE\Timberline\General", "System Directory", "C:\Program Files\Timberline Office\Accounting\")
    TSobjPath = PathAppend(TSobjPath, "tsobject.exe")
    BatchFileName = AppWorkingFolder & "\" & App.EXEName & ".bat"
    BatchFile = FreeFile
       
    
    Open BatchFileName For Output As #BatchFile
'    Print #BatchFile, FileDrive(mDataFolder)
'    Print #BatchFile, "cd " & vbQuote & mDataFolder & vbQuote
'    Print #BatchFile, vbQuote & TSobjPath & vbQuote & " " & vbQuote & mLaunchFile & vbQuote & " " & mUserID & " " & mPswd
'    Print #BatchFile, "cd \"
'    Print #BatchFile, "c:"
    
    
    Print #BatchFile, "pushd " & mDataFolder
    Print #BatchFile, vbQuote & TSobjPath & vbQuote & " " & vbQuote & mLaunchFile & vbQuote & " " & mUserID & " " & mPswd
    Print #BatchFile, "popd"
    
    Close BatchFile
    
    Call Shell(BatchFileName, vbHide)
    
End Sub


Private Sub ReadStatusFile()
On Error GoTo eh
    Dim i As Integer
    Dim s As String
    Dim f As String
    
    Dim sTime As String
    Dim sTask As String
    Dim sMsg  As String
    
    
    i = FreeFile
    Open mStatusLog For Input As i
    
    f = ""
    'skip 2 lines of headings
    Line Input #i, s
    Line Input #i, s
    'read each data line
    While Not EOF(i)
        Line Input #i, s
        
        sTime = Trim(Mid(s, 1, 19))
        sTask = Trim(Mid(s, 79, 36))
        sMsg = Trim(Mid(s, 115))
        
        If IsDate(sTime) Then
            sTime = Format(TimeValue(sTime), "h:mm:ss ampm")
        Else
            sTime = ""
        End If
        
        f = f & sTime & " " & sMsg & " " & sTask & "..." & vbCrLf
    Wend
    Close i
    
    lblStatus.Caption = f
    lblStatus.Refresh
    
eh: Exit Sub
End Sub
