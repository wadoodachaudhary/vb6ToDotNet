VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{19B73C7C-C76C-11D2-849E-840A94D4800A}#2.2#0"; "FTP.ocx"
Begin VB.Form FFTP 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Updating"
   ClientHeight    =   1425
   ClientLeft      =   5835
   ClientTop       =   2895
   ClientWidth     =   4380
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "FFTP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   4380
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2490
      Top             =   300
   End
   Begin DevPowerFTP.FTP FTP1 
      Left            =   2970
      Top             =   300
      _ExtentX        =   847
      _ExtentY        =   847
      _Type           =   "00007211263922370A1D04132E7464606E627F65662A07282635313B352035243722743339223F"
      Asynchronous    =   0   'False
   End
   Begin MSComCtl2.Animation Animation1 
      Height          =   915
      Left            =   180
      TabIndex        =   0
      Top             =   0
      Width           =   4125
      _ExtentX        =   7276
      _ExtentY        =   1614
      _Version        =   393216
      AutoPlay        =   -1  'True
      FullWidth       =   275
      FullHeight      =   61
   End
   Begin VB.Label lblFile 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Updating the Sales 1440 Website..."
      Height          =   240
      Left            =   210
      TabIndex        =   1
      Tag             =   "file123.zip from ftp://ftp.yoohoo.com"
      Top             =   1020
      Width           =   2625
   End
End
Attribute VB_Name = "FFTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private mMode As String
Private mLocalFileName As String
Private mRemoteFileName As String


Private mServer As String
Private mPort As Integer
Private mUID As String
Private mPWD As String
Private mRemoteFolder As String


Public Sub Upload(LocalFileName As String)
On Error Resume Next
    mMode = "upload"
    mLocalFileName = LocalFileName
    Me.Show vbModal
End Sub

Public Sub Download(RemoteFileName As String, LocalFileName As String)
On Error Resume Next
    mMode = "download"
    mLocalFileName = LocalFileName
    mRemoteFileName = RemoteFileName
    Me.Show vbModal
End Sub


Private Sub Form_Load()
    Call IniGetForm(Me)
    Animation1.Open App.Path & "\filecopy.avi"
    Me.Show
    DoEvents
    Call Connect
    If mMode = "upload" Then
        Call FTP1.PutFile(mLocalFileName, mRemoteFolder, FileTitle(mLocalFileName))
    Else
        Call FTP1.GetFile(mRemoteFileName, FilePath(mLocalFileName), mLocalFileName)
    End If
End Sub


Private Sub Connect()
    Dim i As Long
    Dim s As String
    
    'get remote path
    s = Trim(HFApp.Options(WebUploadPath))
    
    'remove ftp:// if present
    If left(s, 6) = "ftp://" Then s = Mid(s, 7)
    
    'get folder
    i = InStr(1, s, "/", vbTextCompare)
    If i > 0 Then mRemoteFolder = Mid(s, i)
    
    'get server and port
    mServer = Parse(s, 1, "/")
    mPort = Val(Parse(mServer, 2, ":"))
    If mPort = 0 Then mPort = 21
    mServer = Parse(mServer, 1, ":")
    
    'get usr/pwd
    mUID = HFApp.Options(WebUploadUser)
    mPWD = HFApp.Options(WebUploadPswd)
    
    'connect to server
    Call FTP1.Connect(mServer, mUID, mPWD, mPort)

End Sub


Private Sub Form_Unload(Cancel As Integer)
    Call IniPutForm(Me)
End Sub

Private Sub FTP1_TransferComplete(ByVal RemoteFileName As String, ByVal LocalFileName As String, ByVal Download As Boolean)
    Unload Me
End Sub

Private Sub FTP1_TransferProgress(ByVal RemoteFileName As String, ByVal LocalFileName As String, ByVal Download As Boolean, ByVal Bytes As Long, ByVal TotalBytes As Long, Cancel As Boolean)
    MsgBox "progress " & Bytes / TotalBytes * 100
End Sub

Private Sub Timer1_Timer()
    Me.Refresh
    DoEvents
End Sub
