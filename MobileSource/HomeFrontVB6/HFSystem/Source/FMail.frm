VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.1#0"; "Codejock.Controls.v13.1.0.ocx"
Begin VB.Form FMail 
   BackColor       =   &H80000005&
   Caption         =   "New Message"
   ClientHeight    =   7230
   ClientLeft      =   2655
   ClientTop       =   5265
   ClientWidth     =   11745
   Icon            =   "FMail.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   11745
   Begin VB.Frame frmHeader 
      Height          =   2805
      Left            =   150
      TabIndex        =   10
      Top             =   855
      Width           =   10845
      Begin VB.PictureBox picCanvas 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   8460
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   11
         Top             =   390
         Visible         =   0   'False
         Width           =   480
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   9135
         Top             =   465
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComctlLib.ListView lstAttachments 
         Height          =   1290
         Left            =   105
         TabIndex        =   7
         Top             =   1380
         Width           =   4890
         _ExtentX        =   8625
         _ExtentY        =   2275
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         Icons           =   "SmallIcons"
         SmallIcons      =   "SmallIcons"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         NumItems        =   0
      End
      Begin XtremeSuiteControls.PushButton Command1 
         Height          =   1080
         Index           =   0
         Left            =   165
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   180
         Width           =   1080
         _Version        =   851969
         _ExtentX        =   1905
         _ExtentY        =   1905
         _StockProps     =   79
         Caption         =   "&Send"
         BackColor       =   -2147483633
         Appearance      =   3
      End
      Begin VB.TextBox txtSubject 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2100
         TabIndex        =   6
         Top             =   885
         Width           =   3885
      End
      Begin VB.TextBox txtCC 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2100
         TabIndex        =   4
         Top             =   540
         Width           =   3885
      End
      Begin VB.TextBox txtTo 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2100
         TabIndex        =   2
         Text            =   "michael@google.com"
         Top             =   195
         Width           =   3885
      End
      Begin XtremeSuiteControls.PushButton Command1 
         Height          =   315
         Index           =   1
         Left            =   1335
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   195
         Width           =   675
         _Version        =   851969
         _ExtentX        =   1191
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "&To"
         BackColor       =   -2147483633
         Appearance      =   3
      End
      Begin XtremeSuiteControls.PushButton Command1 
         Height          =   315
         Index           =   2
         Left            =   1335
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   555
         Width           =   675
         _Version        =   851969
         _ExtentX        =   1191
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "&Cc"
         BackColor       =   -2147483633
         Appearance      =   3
      End
      Begin MSComctlLib.ImageList SmallIcons 
         Left            =   9840
         Top             =   570
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "S&ubject"
         Height          =   240
         Left            =   1335
         TabIndex        =   5
         Top             =   945
         Width           =   675
      End
   End
   Begin VB.TextBox txtBody 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4845
      Left            =   30
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   8
      Top             =   2340
      Width           =   11685
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   9
      Top             =   0
      Width           =   11745
      _ExtentX        =   20717
      _ExtentY        =   1058
      ButtonWidth     =   1667
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Attach File"
            Key             =   "Attach"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cut"
            Key             =   "Cut"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Copy"
            Key             =   "Copy"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Paste"
            Key             =   "Paste"
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList LargeIcons 
         Left            =   11010
         Top             =   30
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMail.frx":000C
               Key             =   "Attach"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMail.frx":08E6
               Key             =   "Copy"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMail.frx":11C0
               Key             =   "Urgent"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMail.frx":1A9A
               Key             =   "Cut"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMail.frx":2374
               Key             =   "Paste"
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "FMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mInitialized As Boolean
Private mCancel As Boolean
Private mTORecipients As String
Private mCCRecipients As String
Private mSubject As String
Private mBody As String
Private mAttachments As String






Public Function Edit(TORecipients As String, CCRecipients As String, Subject As String, Body As String, Attachments As String) As Boolean
    On Error Resume Next
    Dim i As Long
    Dim files As String
    
    Load Me
    
    mTORecipients = TORecipients
    mCCRecipients = CCRecipients
    mSubject = Subject
    mBody = Body
    mAttachments = Attachments

    mInitialized = False
    mCancel = True
    
    Screen.MousePointer = vbDefault
    Me.Show vbModal
    If Not mCancel Then
        TORecipients = txtTo.Text
        CCRecipients = txtCC.Text
        Subject = txtSubject.Text
        Body = txtBody.Text
        
        files = ""
        For i = 1 To lstAttachments.ListItems.Count
            files = files & "; " & lstAttachments.ListItems(i).Key
        Next
        files = Mid(files, 3)
        
    End If
    Edit = Not mCancel
    Unload Me
    
End Function





Private Sub Command1_Click(Index As Integer)
    Dim s As String
    
    Select Case Index
    Case 0 'send
        'validate
        If Not ValidMessage Then Exit Sub
        mCancel = False
        Me.Hide
        
    Case 1 'to
        If Not FPickList.Choose(HFApp.Databases(dbHomefront), "Contact", "select * from dbo.EmailAddressBook where divisionid is null or divisionid=" & DbQuote(Num, HFApp.DivisionID), , , , , "divisionid") Then Exit Sub
        txtTo.Text = txtTo.Text & "; " & FPickList.SelectedItem("name") & " <" & FPickList.SelectedItem("email") & ">"
        While IsIn(left(txtTo.Text, 1), " ", ";")
            txtTo.Text = Mid(txtTo.Text, 2)
        Wend
        
    Case 2 'cc
        If Not FPickList.Choose(HFApp.Databases(dbHomefront), "Contact", "select * from dbo.EmailAddressBook where divisionid is null or divisionid=" & DbQuote(Num, HFApp.DivisionID), , , , , "divisionid") Then Exit Sub
        txtCC.Text = txtCC.Text & "; " & FPickList.SelectedItem("name") & " <" & FPickList.SelectedItem("email") & ">"
        While IsIn(left(txtCC.Text, 1), " ", ";")
            txtCC.Text = Mid(txtCC.Text, 2)
        Wend
        
        
    End Select
End Sub

Private Function ValidMessage() As Boolean
    Dim i As Long
    Dim s As String
    
    ValidMessage = False
    
    If Trim(txtTo.Text) = "" And Trim(txtCC.Text) = "" Then
        MsgBox "At least one recipient is required.", vbQuestion, "HomeFront"
        Exit Function
    End If
    
    
    For i = 1 To Parse(txtTo.Text, , ";")
        s = Parse(txtTo.Text, i, ";")
        s = Parse(s, 2, "<")
        s = Parse(s, 1, ">")
        If s = "" Then s = Parse(txtTo.Text, i, ";")
        
        
        If InStr(1, s, " ", vbTextCompare) <> 0 Or _
           InStr(1, s, "@", vbTextCompare) = 0 Or _
           InStr(1, s, ".", vbTextCompare) = 0 Then
            
            If vbNo = MsgBox("HomeFront does not recognize this address" & vbCrLf & vbCrLf & s & vbCrLf & vbCrLf & "Do you want to send the message anyway?", vbQuestion + vbYesNo, "HomeFront") Then Exit Function
        
        End If
        
    Next
        
    If Trim(txtSubject.Text) = "" Then
        If vbNo = MsgBox("There is no subject. Do you want to send the message anyway?", vbQuestion + vbYesNo, "HomeFront") Then Exit Function
    End If
    
    ValidMessage = True
    

End Function

Private Sub Form_Activate()
    Dim i As Long
    
    If mInitialized Then Exit Sub
    
    txtTo.Text = mTORecipients
    txtCC.Text = mCCRecipients
    txtSubject.Text = mSubject
    txtBody.Text = mBody
        
        
    For i = 1 To Parse(mAttachments, , ";")
        Call AddFile(Parse(mAttachments, i, ";"))
    Next
    
    txtTo.SetFocus

    mInitialized = True
End Sub

Private Sub Form_Load()
    Call SetToolbarIcons(Toolbar, LargeIcons)
    Call IniGetForm(Me)
    
    
End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim rc As Integer
    
    If mCancel Then
        rc = MsgBox("Do you want to send this message?", vbQuestion + vbYesNoCancel, "HomeFront")
        Select Case rc
            Case vbCancel: Cancel = True
            Case vbYes:    mCancel = False
            Case vbNo:     mCancel = True
        End Select
    End If
    
End Sub

Private Sub Form_Resize()
On Error Resume Next

    frmHeader.BorderStyle = 0
    frmHeader.Move 0, Toolbar.Top + Toolbar.Height, Me.ScaleWidth
    
    If lstAttachments.ListItems.Count > 0 Then
        lstAttachments.Visible = True
        frmHeader.Height = 2820
    Else
        lstAttachments.Visible = False
        frmHeader.Height = 1365
    End If
    
    lstAttachments.Width = frmHeader.Width - 2 * lstAttachments.left
    txtBody.Move 0, frmHeader.Top + frmHeader.Height, Me.ScaleWidth, Me.ScaleHeight - (frmHeader.Top + frmHeader.Height)
    txtTo.Width = frmHeader.Width - txtTo.left - 240
    txtCC.Width = txtTo.Width
    txtSubject.Width = txtTo.Width
End Sub


Private Sub lstAttachments_DblClick()
    Dim Filename As String
    Filename = lstAttachments.SelectedItem.Key
    If Filename <> "" Then Call ShellFile(Me.hwnd, Filename, , False)
End Sub

Private Sub lstAttachments_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo 0

    If KeyCode = vbKeyDelete Then
        Call lstAttachments.ListItems.Remove(lstAttachments.SelectedItem.Index)
        Call Form_Resize
        lstAttachments.Sorted = True
    End If
    
End Sub



Private Sub AddFile(Filename As String)
On Error Resume Next

    If Filename = "" Then Exit Sub
    
    Dim FInfo As New ClsFileInfo
    FInfo.FullPathName = Trim(Filename)
    
    'get icon
    Set picCanvas.Picture = New StdPicture
    'Call ImageList_Draw(FInfo.hSmlIList, FInfo.hSmlIcon, picCanvas.hDC, 0, 0, ILD_TRANSPARENT)
    Call ImageList_Draw(FInfo.hLrgIList, FInfo.hLrgIcon, picCanvas.hDC, 0, 0, ILD_TRANSPARENT)
    Me.SmallIcons.ListImages.Add , FInfo.FileExtension, picCanvas.Image
    
    'add file to list
    lstAttachments.ListItems.Add , FInfo.FullPathName, FInfo.Filename, FInfo.FileExtension
    Call Form_Resize
    lstAttachments.Sorted = True

End Sub


Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    
    Case "Attach"
        On Error Resume Next
        Dim file As String
        If Not VBGetOpenFileName(file, , , , , , , , , "Insert File", , Me.hwnd) Then Exit Sub
        Call AddFile(file)
    
    Case "Cut"
        txtBody.SelText = ""
    Case "Copy"
        Clipboard.Clear
        Clipboard.SetText txtBody.SelText
    Case "Paste"
        txtBody.SelText = Clipboard.GetText
    End Select
End Sub


Private Sub txtSubject_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
'txtSubject.Text = (Data.files(1))
End Sub

