VERSION 5.00
Object = "{55473EAC-7715-4257-B5EF-6E14EBD6A5DD}#1.0#0"; "vbalProgBar6.ocx"
Begin VB.Form FDbUpgrade 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Database Upgrade"
   ClientHeight    =   5280
   ClientLeft      =   7815
   ClientTop       =   6060
   ClientWidth     =   7470
   Icon            =   "FDbUpgrade.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   352
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   498
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picFrame 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   4710
      Index           =   1
      Left            =   3444
      ScaleHeight     =   314
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   493
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   5232
      Visible         =   0   'False
      Width           =   7395
      Begin vbalProgBarLib6.vbalProgressBar ProgressBar 
         Height          =   375
         Left            =   420
         TabIndex        =   15
         Top             =   4170
         Width           =   6570
         _ExtentX        =   11589
         _ExtentY        =   661
         Picture         =   "FDbUpgrade.frx":0E42
         ForeColor       =   0
         BarPicture      =   "FDbUpgrade.frx":0E5E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblStep 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Step 6 of 17 -- Add new feature to something  "
         Height          =   195
         Left            =   480
         TabIndex        =   14
         Top             =   690
         UseMnemonic     =   0   'False
         Width           =   3255
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Please wait..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   330
         TabIndex        =   13
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   1605
      End
   End
   Begin VB.PictureBox picFrame 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   4710
      Index           =   2
      Left            =   12012
      ScaleHeight     =   314
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   493
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   5364
      Visible         =   0   'False
      Width           =   7395
      Begin VB.TextBox txtErrorStatement 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   3945
         Left            =   1020
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   7
         Top             =   690
         Width           =   6315
      End
      Begin VB.Image Image4 
         Height          =   480
         Left            =   420
         Picture         =   "FDbUpgrade.frx":0E7A
         Top             =   840
         Width           =   480
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "The upgrade cannot be completed"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   4
         Left            =   330
         TabIndex        =   6
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   4125
      End
   End
   Begin VB.CommandButton cmdNav 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   6120
      TabIndex        =   1
      Top             =   4860
      Width           =   1155
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "&OK"
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   4875
      TabIndex        =   0
      Top             =   4860
      Width           =   1155
   End
   Begin VB.PictureBox picFrame 
      Align           =   1  'Align Top
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   4710
      Index           =   0
      Left            =   0
      Picture         =   "FDbUpgrade.frx":1744
      ScaleHeight     =   4710
      ScaleWidth      =   7470
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   7470
      Begin VB.OptionButton optBackup 
         BackColor       =   &H80000005&
         Caption         =   "Yes, I have backed up my data"
         Height          =   255
         Index           =   3
         Left            =   2910
         TabIndex        =   10
         Top             =   2700
         Width           =   3015
      End
      Begin VB.OptionButton optBackup 
         BackColor       =   &H80000005&
         Caption         =   "No, I have not backed up my data"
         Height          =   255
         Index           =   2
         Left            =   2910
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   2940
         Value           =   -1  'True
         Width           =   3075
      End
      Begin VB.Label lblDescription 
         BackStyle       =   0  'Transparent
         Caption         =   "When you are ready to upgrade your data, ensure all users are logged out and click OK to continue."
         Height          =   795
         Index           =   6
         Left            =   2670
         TabIndex        =   11
         Top             =   3510
         UseMnemonic     =   0   'False
         Width           =   4635
      End
      Begin VB.Label lblDescription 
         BackStyle       =   0  'Transparent
         Caption         =   $"FDbUpgrade.frx":272FE
         Height          =   1095
         Index           =   5
         Left            =   2670
         TabIndex        =   8
         Top             =   1590
         UseMnemonic     =   0   'False
         Width           =   4635
      End
      Begin VB.Label lblDescription 
         BackStyle       =   0  'Transparent
         Caption         =   $"FDbUpgrade.frx":27409
         Height          =   795
         Index           =   0
         Left            =   2670
         TabIndex        =   4
         Top             =   690
         UseMnemonic     =   0   'False
         Width           =   4635
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A database upgrade is required"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   2640
         TabIndex        =   3
         Top             =   210
         UseMnemonic     =   0   'False
         Width           =   3795
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   -200
      X2              =   9719
      Y1              =   315
      Y2              =   315
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   0
      X1              =   -12
      X2              =   9707
      Y1              =   314
      Y2              =   314
   End
End
Attribute VB_Name = "FDbUpgrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private mConnection   As Connection
Private mConnection2  As Connection
Private mDbRevision   As Long
Private mCompleted    As Boolean

Private Type StatementType
    Description  As String  'displayed in listbox
    SQLStatement As String  'query to run
    IgnoreErrors As Boolean
End Type
Private mStatements() As StatementType


Public Function DropIndexesAndKeys(TableName As String, UniqueOnly As Boolean) As String
    Dim s As String
    
    s = ""
    s = s & "declare @s nvarchar(max)=''" & vbCrLf
    s = s & "select @s=@s + case when i.is_primary_key=1 or i.is_unique_constraint=1" & vbCrLf
    s = s & "then 'alter table ' + t.name + ' drop constraint ' + i.name + ' '" & vbCrLf
    s = s & "else 'drop index ' + i.name + ' on ' + t.name + ' ' end " & vbCrLf
    s = s & "from sys.tables t" & vbCrLf
    s = s & "join sys.indexes i on t.object_id=i.object_id" & vbCrLf
    s = s & "where i.type_desc<>'HEAP'" & vbCrLf
    s = s & "and t.name=" & DbQuote(Str, TableName) & vbCrLf
    If UniqueOnly Then s = s & "and i.is_unique=1" & vbCrLf
    s = s & "exec sp_executeSQL @s" & vbCrLf
    
    DropIndexesAndKeys = s
End Function
Public Function Upgrade(ConnectionString As String) As Boolean
On Error GoTo eh
    Dim r      As Long
    Dim prev   As String
    Dim s      As String
    Dim majorversion As Long

    'open connection
    Set mConnection = New Connection
    Set mConnection2 = New Connection
    mConnection.Open ConnectionString
    mConnection2.Open ConnectionString
    mConnection.CommandTimeout = 3000
    mConnection2.CommandTimeout = 3000
    
    'reposition and initialize frames
    For r = 0 To 2
        picFrame(r).Move 0, 0
    Next
    
    
    
    
    'legacy steps removed - dont lose this number.
    ReDim mStatements(8620)
    Call LoadDBRevisions
    
    
    'get revision number
    On Error Resume Next
    mDbRevision = mConnection.Execute("SELECT MAX(Revision) FROM dbo.DBRevisions")(0)
    
    'show a warning if the db is at a higher revision level than the app
    If mDbRevision > UBound(mStatements) Then
    
        s = ""
        s = s & "A new version has been installed on the server. You should" & vbCrLf
        s = s & "upgrade your workstation software as soon as possible. If you" & vbCrLf
        s = s & "choose not to upgrade your software you may experience" & vbCrLf
        s = s & "problems." & vbCrLf & vbCrLf
        s = s & "Please contact your HomeFront administrator for instructions."
        Upgrade = vbOK = MsgBox(s, vbOKCancel + vbExclamation, App.ProductName)
    
    ElseIf mDbRevision <= 8620 Then
    
        s = ""
        s = s & "This database cannot be upgraded by this version of " & App.ProductName & vbCrLf
        s = s & "as it is too old. Please contact " & App.ProductName & " support for assistance." & vbCrLf
        Call MsgBox(s, vbOKOnly + vbExclamation, App.ProductName)
        Upgrade = False
    
    
    'show screen if there are 'un-run' statements
    ElseIf mDbRevision < UBound(mStatements) Then
        
        '-- do check sql version -----------------------------------
        'Microsoft SQL Server 2005           9.00.1399.06  -- Microsoft extended support for SQL Server 2005 ended on April 12, 2016
        'Microsoft SQL Server 2008 (SP1)     10.0.2573.0
        'Microsoft SQL Server 2008 R2 (RTM)  10.50.1600.1
        'Microsoft SQL Server 2012           11.0.2100.60
        'Microsoft SQL Server 2014           12.0.2254.0
        'Microsoft SQL Server 2016 (RTM)     13.0.1601.5
        s = mConnection.Execute("SELECT SERVERPROPERTY('ProductVersion')")(0)
        majorversion = Val(Parse(s, 1, "."))
        Select Case True
        Case majorversion <= 9:   s = "Microsoft ended extended support for SQL Server 2005 on April 12, 2016."
        'Case majorversion = 10
        'Case majorversion = 11
        'Case majorversion = 12
        'Case majorversion = 13
        Case Else: s = ""
        End Select
        If s <> "" Then
            
            s = "Unsupported SQL Server version detected. This database cannot be upgraded." & vbCrLf & _
                s & vbCrLf & vbCrLf & _
                "Please upgrade to a recent version of SQL Server."
            Call MsgBox(s, vbOKOnly + vbExclamation, App.ProductName)
            Upgrade = False
        
        Else
        
            '-- do upgrade -----------------------------------
            Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
            Screen.MousePointer = vbNormal
            Me.Show vbModal
            Upgrade = mCompleted
        
        End If
    Else
        Upgrade = True
    
    End If
    
    On Error Resume Next
    mConnection.Close
    mConnection2.Close
    Set mConnection = Nothing
    Set mConnection2 = Nothing
    Unload Me
    
Exit Function
eh: MsgBox Err.Description
End Function



Public Sub sql(Description As String, SQLStatement As String, Optional IgnoreErrors As Boolean)
    Dim i As Long
    i = UBound(mStatements) + 1

 


'If InIde And IsIn(i, 14342) Then Stop
    ReDim Preserve mStatements(i)
    mStatements(i).Description = Description
    mStatements(i).SQLStatement = SQLStatement
    mStatements(i).IgnoreErrors = IgnoreErrors
End Sub


Private Sub cmdNav_Click(Index As Integer)
    Select Case Index
        Case 0 'ok
            cmdNav(0).Enabled = False
            cmdNav(1).Enabled = False
            picFrame(1).Visible = True
            picFrame(0).Visible = False
            Call RunStatements
            
        Case 1 'cancel
            Unload Me
    End Select
End Sub




Private Function RunStatements() As Boolean
On Error Resume Next

    Dim i As Long
    Dim s As String
    Dim SQLStatement As String
    
    Screen.MousePointer = vbHourglass
    
    ProgressBar.Max = UBound(mStatements)
    ProgressBar.Min = mDbRevision
    
    Call DisableAudit
    
    For i = mDbRevision + 1 To UBound(mStatements)
    
        Err.Clear
            
        SQLStatement = mStatements(i).SQLStatement
            
        lblStep.Caption = (i - mDbRevision) & " of " & (UBound(mStatements) - mDbRevision + 1) & "  --  " & mStatements(i).Description
        ProgressBar.Value = i
        lblStep.Refresh
        DoEvents
        
        
        Select Case True
            Case SQLStatement = ""
            
            Case left(SQLStatement, 7) = "RUNPROC"
                Call RUNPROC_EncryptPswds
            
            Case Else
                Err.Clear
                mConnection.CommandTimeout = 3600000 '1 hr
                mConnection.Execute SQLStatement
        End Select
                    
        If Err.Number = 0 Or mStatements(i).IgnoreErrors Or CBool(InStr(1, Err.Description, "Warning:")) Then
            s = ""
            s = s & "INSERT INTO DBRevisions(Revision,Description,UStmp,TStmp,SqlStatement)" & vbCrLf
            s = s & "VALUES(" & i & vbCrLf
            s = s & "      ," & DbQuote(Str, Mid(mStatements(i).Description, 1, 50)) & vbCrLf
            s = s & "      ," & DbQuote(Str, HFApp.LoginID) & vbCrLf
            s = s & "      ,GETDATE()" & vbCrLf
            s = s & "      ," & DbQuote(Str, left(SQLStatement, 7599)) & ")"
            mConnection.Execute s
        Else
            Screen.MousePointer = vbDefault
            txtErrorStatement.Text = "Error: " & Trim(Parse(Err.Description, Parse(Err.Description, , "]"), "]")) & vbCrLf & vbCrLf & _
                                     "Step: " & i & vbCrLf & _
                                     String(80, "-") & vbCrLf & _
                                     SQLStatement
            picFrame(2).Visible = True
            picFrame(1).Visible = False
            cmdNav(1).Enabled = True
            mCompleted = False
            Exit Function
        End If
        
    Next
    
    mConnection.Execute "exec ZYB_CreateDefaults"
    mCompleted = True
    Unload Me
    
End Function


Private Sub Form_Activate()
    Call WindowOnTop(Me, True)
    If InIde() Then
        optBackup(3).Value = True
        SetCtrlFocus Me.cmdNav(0)
    End If
End Sub


Private Sub Form_Resize()
    picFrame(1).Move picFrame(0).left, picFrame(0).Top, picFrame(0).Width, picFrame(0).Height
    picFrame(2).Move picFrame(0).left, picFrame(0).Top, picFrame(0).Width, picFrame(0).Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call EnableAudit
End Sub

Private Sub optBackup_Click(Index As Integer)
    cmdNav(0).Enabled = optBackup(3).Value
End Sub


Private Sub EnableAudit()
On Error Resume Next
    
    mConnection.Execute "enable trigger audit_view_change on database"
    mConnection.Execute "enable trigger audit_table_change on database"
    mConnection.Execute "enable trigger audit_trigger_change on database"

End Sub
Private Sub DisableAudit()
On Error Resume Next
    
    mConnection.Execute "disable trigger audit_view_change on database"
    mConnection.Execute "disable trigger audit_table_change on database"
    mConnection.Execute "disable trigger audit_trigger_change on database"

End Sub




Private Sub RUNPROC_EncryptPswds()
    Dim hff As Object
    Dim pwd As String
    Dim s As String
    Dim rs As Recordset
    
    Set hff = CreateObject("ZYBFunctions.HomeFrontFunctions")
    
    s = "select user_id,user_password from user_manager where user_password<>'' and user_access in (-1,-2)"
    Set rs = mConnection.Execute(s)

    While Not rs.EOF
        pwd = "" & rs("user_password")
        pwd = hff.MyCrypt(pwd, "Fazlul")

        s = ""
        s = s & "update user_manager set" & vbCrLf
        s = s & " user_access=6" & vbCrLf
        s = s & ",user_password=" & DbQuote(Str, pwd) & vbCrLf
        s = s & "where user_id=" & DbQuote(Str, "" & rs("user_id")) & vbCrLf
        Call mConnection2.Execute(s)
        rs.MoveNext
    Wend

End Sub

