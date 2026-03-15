VERSION 5.00
Begin VB.Form FAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   5655
   ClientLeft      =   5400
   ClientTop       =   4605
   ClientWidth     =   7515
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FAbout.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   377
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   501
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdNav 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   5955
      TabIndex        =   0
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"FAbout.frx":000C
      Height          =   960
      Left            =   60
      TabIndex        =   24
      Top             =   4800
      Width           =   5520
   End
   Begin VB.Label lblLockedRecords 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "locked records?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   6015
      TabIndex        =   23
      Top             =   4110
      Width           =   1230
   End
   Begin VB.Label lblLoggedInUsers 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "who is logged in?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   6015
      TabIndex        =   22
      Top             =   4320
      Width           =   1230
   End
   Begin VB.Label lblRevision 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Revision"
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   3930
      TabIndex        =   21
      Top             =   1500
      Width           =   615
   End
   Begin VB.Label lblServer 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Server"
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   3930
      TabIndex        =   20
      Top             =   1290
      Width           =   465
   End
   Begin VB.Label lblWeburl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "www.HyphenSolutions.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   3930
      TabIndex        =   19
      Top             =   2490
      Width           =   1950
   End
   Begin VB.Label lblInfoEmail 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "info@HyphenSolutions.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   3930
      TabIndex        =   18
      Top             =   2280
      Width           =   1965
   End
   Begin VB.Label lblCopyright 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "copyright"
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   3930
      TabIndex        =   17
      Top             =   2070
      Width           =   645
   End
   Begin VB.Label lblDevelopedBy 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hyphen Solutions"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   21
      Left            =   3930
      TabIndex        =   16
      Top             =   1860
      Width           =   1245
   End
   Begin VB.Label lblLicensedFor 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Licensed for"
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   3930
      TabIndex        =   15
      Top             =   3510
      Width           =   870
   End
   Begin VB.Label lblSerialNumber 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Serial Number"
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   3930
      TabIndex        =   14
      Top             =   3300
      Width           =   990
   End
   Begin VB.Label lblCompanyName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Company"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00504B4A&
      Height          =   195
      Left            =   3930
      TabIndex        =   13
      Top             =   3090
      Width           =   780
   End
   Begin VB.Label lblClientID 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Client ID"
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   3930
      TabIndex        =   12
      Top             =   2880
      Width           =   600
   End
   Begin VB.Label lblDatabase 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Database"
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   3930
      TabIndex        =   11
      Top             =   1080
      Width           =   690
   End
   Begin VB.Label lblSystemVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "System Version"
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   3930
      TabIndex        =   10
      Top             =   690
      Width           =   1080
   End
   Begin VB.Label lblModuleName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ModuleName"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00504B4A&
      Height          =   240
      Left            =   3930
      TabIndex        =   9
      Top             =   450
      Width           =   1395
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Developed by"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   5
      Left            =   2820
      TabIndex        =   8
      Top             =   1860
      Width           =   990
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Licensed for"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   9
      Left            =   2925
      TabIndex        =   7
      Top             =   3510
      Width           =   870
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Serial Number"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   8
      Left            =   2805
      TabIndex        =   6
      Top             =   3300
      Width           =   990
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Company"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   7
      Left            =   3135
      TabIndex        =   5
      Top             =   3090
      Width           =   660
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Client ID"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   6
      Left            =   3195
      TabIndex        =   4
      Top             =   2880
      Width           =   600
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Database"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   4
      Left            =   3105
      TabIndex        =   3
      Top             =   1080
      Width           =   690
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "System Version"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   2
      Left            =   2715
      TabIndex        =   2
      Top             =   690
      Width           =   1080
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Module"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   0
      Left            =   3270
      TabIndex        =   1
      Top             =   480
      Width           =   525
   End
   Begin VB.Image Image1 
      Height          =   5655
      Index           =   0
      Left            =   0
      Picture         =   "FAbout.frx":0133
      Top             =   0
      Visible         =   0   'False
      Width           =   7500
   End
   Begin VB.Image Image1 
      Height          =   5655
      Index           =   1
      Left            =   0
      Picture         =   "FAbout.frx":8A271
      Top             =   210
      Width           =   7500
   End
End
Attribute VB_Name = "FAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
    Dim s As String
    Dim rs As Recordset
    Dim userAccess As Long
    
    Call IniGetForm(Me)
    
    Image1(0).Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    Image1(1).Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    If JobSimplicity Then
        Image1(0).Visible = True
        Image1(1).Visible = False
    End If
    
    
    
    lblModuleName = mModuleName
    
      
    lblSystemVersion = mSystemVersion
    
    
    If HFApp.Databases(dbHomefront).State = adStateOpen Then
        lblDatabase = HFApp.Databases(dbHomefront).Properties("DBMS Name") & " (" & HFApp.Databases(dbHomefront).Properties("DBMS Version") & ")"
        lblServer = HFApp.Databases(dbHomefront).Properties("Current Catalog") & "@" & HFApp.Databases(dbHomefront).Properties("Server Name")
        lblRevision = "Revision: " & HFApp.SqlExec("select max(revision) from dbrevisions")(0)
    Else
        lblDatabase = ""
        lblServer = ""
        lblRevision = ""
    End If
    
    
    lblCopyright = Year(VBA.Date()) & " © Copyright Hyphen Solutions"
    
    
    
    
    lblClientID = HFApp.ClientID
    lblCompanyName = HFApp.LicensedOwner
    
    
    lblSerialNumber = HFApp.LicenseNumber
    lblLicensedFor = HFApp.LicensedUsers & " concurrent users"
    
    lblLoggedInUsers.Enabled = HFApp.Databases(dbHomefront).State = adStateOpen
    lblLockedRecords.Enabled = HFApp.DivisionID <> 0
    
    Call WindowOnTop(Me, True)

End Sub

Private Sub cmdNav_Click(Index As Integer)
    Unload Me
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Call IniPutForm(Me)
End Sub

Private Sub lblInfoEmail_Click()
    Call ShellFile(hwnd, "mailto:info@HomeFrontSoftware.com")
End Sub

Private Sub lblLockedRecords_Click()
    Call HFApp.RunTask("LockedRecords")
End Sub

Private Sub lblLoggedInUsers_Click()
    Call HFApp.RunTask("ShowUserMonitor")
End Sub



Private Sub lblWeburl_Click()
    Call ShellFile(hwnd, "www.HomeFrontSoftware.com")
End Sub
