VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{C9460280-3EED-11D0-A647-00A0C91EF7B9}#1.0#0"; "ImageViewer2.OCX"
Begin VB.Form FDocuments 
   Caption         =   "Documents"
   ClientHeight    =   6330
   ClientLeft      =   10995
   ClientTop       =   3345
   ClientWidth     =   8130
   Icon            =   "FDocuments.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6330
   ScaleWidth      =   8130
   Begin SCRIBBLELib.ImageViewer ImageViewer 
      Height          =   3915
      Left            =   2820
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   780
      Width           =   3315
      _Version        =   65536
      _ExtentX        =   5847
      _ExtentY        =   6906
      _StockProps     =   0
      Border          =   0   'False
      LicenseKey      =   "9912 single developer license"
   End
   Begin MSComctlLib.ListView lvThumbs 
      Height          =   5475
      Left            =   -120
      TabIndex        =   0
      Top             =   840
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   9657
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483636
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Height          =   780
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   1376
      ButtonWidth     =   1058
      ButtonHeight    =   1376
      Style           =   1
      ImageList       =   "ToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save"
            Key             =   "save"
            Object.ToolTipText     =   "Save As"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Print"
            Key             =   "print"
            Object.ToolTipText     =   "Print"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Email"
            Key             =   "mail"
            Object.ToolTipText     =   "Email"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Rotate"
            Key             =   "rotate"
            Object.ToolTipText     =   "Rotate"
         EndProperty
      EndProperty
      Begin VB.PictureBox picZoom 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   2760
         ScaleHeight     =   255
         ScaleWidth      =   1815
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   60
         Width           =   1815
         Begin HFPayables.VBCombo cboZoom 
            Height          =   240
            Left            =   540
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   0
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   423
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Zoom"
            Height          =   195
            Left            =   60
            TabIndex        =   4
            Top             =   30
            Width           =   405
         End
      End
      Begin MSComctlLib.ImageList ilThumbs 
         Left            =   9480
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   122
         ImageHeight     =   170
         MaskColor       =   12632256
         UseMaskColor    =   0   'False
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDocuments.frx":000C
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ToolbarIcons 
         Left            =   8880
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   6
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDocuments.frx":2B1E
               Key             =   "save"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDocuments.frx":33F8
               Key             =   "mail"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDocuments.frx":3CD2
               Key             =   "rotate"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDocuments.frx":45AC
               Key             =   "right"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDocuments.frx":4E86
               Key             =   "left"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDocuments.frx":5760
               Key             =   "print"
            EndProperty
         EndProperty
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   4800
         ScaleHeight     =   735
         ScaleWidth      =   26430
         TabIndex        =   6
         Top             =   0
         Width           =   26435
         Begin VB.Label lblDescription 
            AutoSize        =   -1  'True
            Caption         =   "Vendor: BRECKENR"
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   7
            Top             =   60
            Width           =   1485
         End
         Begin VB.Label lblDescription 
            AutoSize        =   -1  'True
            Caption         =   "Invoice: PR123-6"
            Height          =   195
            Index           =   1
            Left            =   0
            TabIndex        =   8
            Top             =   240
            Width           =   1245
         End
         Begin VB.Label lblPath 
            AutoSize        =   -1  'True
            Caption         =   "C:\document\path"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   0
            TabIndex        =   9
            Top             =   540
            Width           =   1320
         End
      End
   End
   Begin VB.Image imgCanvas 
      Height          =   1905
      Left            =   6420
      Stretch         =   -1  'True
      Top             =   840
      Visible         =   0   'False
      Width           =   1470
   End
End
Attribute VB_Name = "FDocuments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const SRCFILE = "FDocuments::"
Private mPath  As String


Public Property Let Description(Index As Integer, RHS As String)
    lblDescription(Index).Caption = RHS
End Property



Public Property Let path(RHS As String)
On Error Resume Next
    mPath = RHS
    lblPath.Caption = FilePath(RHS)
    Call CreatePath("Documents Path", FilePath(RHS))
End Property

Private Sub Form_Activate()
On Error Resume Next
    lvThumbs.SetFocus
End Sub

Private Sub Form_Load()
On Error GoTo eh

    Dim i     As Long
    Dim color As Long
    Dim path  As String
    Dim File  As String
    Call IniGetForm(Me)
    
    'set up toolbar
    Call SetToolbarIcons(Toolbar, ToolbarIcons)
    Toolbar.Visible = True
    cboZoom.AddItem "Full Page"
    cboZoom.AddItem "Page Width"
    cboZoom.AddItem "25%"
    cboZoom.AddItem "50%"
    cboZoom.AddItem "75%"
    cboZoom.AddItem "100%"
    cboZoom.AddItem "150%"
    cboZoom.AddItem "200%"
    cboZoom.AddItem "500%"
    cboZoom.ListIndex = 0
    
    'set up viewer
    Call OleTranslateColor(&H8000000C, 0, color)
    Call ImageViewer.SetBackgroundColor(color)
       
    'load files
    path = FilePath(mPath)
    File = Dir(mPath, , True)
    While File <> ""
        ImageViewer.FileName = ""
        ImageViewer.FileName = PathAppend(path, File)
        Set imgCanvas.Picture = New StdPicture
        imgCanvas.Picture = ImageViewer.Copy2PictureBox()
        ilThumbs.ListImages.Add , "K" & ImageViewer.FileName, imgCanvas.Picture
        File = Dir
    Wend
    
    Set lvThumbs.Icons = ilThumbs
    For i = 2 To ilThumbs.ListImages.count
        lvThumbs.ListItems.Add , ilThumbs.ListImages.Item(i).Key, i, i
    Next
    
    On Error Resume Next
    lvThumbs.ListItems(1).selected = True
    Call lvThumbs_ItemClick(lvThumbs.SelectedItem)

Exit Sub
eh: Call ErrHandler(SRCFILE & "Form_Load")
End Sub

Private Sub lblPath_Click()
On Error Resume Next
    Call ShellFile(Me.hwnd, lblPath.Caption)
End Sub

Private Sub lvThumbs_ItemClick(ByVal Item As MSComctlLib.ListItem)
'On Error Resume Next
    Me.Caption = Mid(Item.Key, 2)
    ImageViewer.Visible = False
    ImageViewer.FileName = Mid(Item.Key, 2)
    Call ReSizeImage
    ImageViewer.Visible = True
End Sub


Private Sub Form_Resize()
On Error Resume Next
    lvThumbs.Move 0, Toolbar.Height, lvThumbs.Width, Me.ScaleHeight - Toolbar.Height
    ImageViewer.Move lvThumbs.Width + 60, lvThumbs.Top, Me.ScaleWidth - 60 - lvThumbs.Width, lvThumbs.Height
    picZoom.Move Toolbar.Buttons(5).left + Toolbar.Buttons(5).Width + 60, (Toolbar.Height - picZoom.Height) / 2
End Sub

Private Sub cboZoom_Click()
    Call ReSizeImage
End Sub
Private Sub cboZoom_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call ReSizeImage
End Sub
Private Sub cboZoom_Validate(Cancel As Boolean)
    If cboZoom.ListIndex = -1 And Not IsNumeric(cboZoom.Text) Then
        Cancel = True
    Else
        Call ReSizeImage
    End If
End Sub
Private Sub ReSizeImage()
On Error Resume Next
    Select Case cboZoom.ListIndex
        Case 0 'full page
            ImageViewer.View = 9
        Case 1 'page width
            ImageViewer.View = 8
            ImageViewer.ViewSize = (ImageViewer.Width - 240) / ScreenTwipsPerPixelX / ImageViewer.FileWidth * 100
        Case -1 'user specified
            cboZoom.Text = Int(Val(cboZoom.Text)) & "%"
            ImageViewer.View = 8
            ImageViewer.ViewSize = Val(cboZoom.Text) / 100
        Case Else 'selected from numeric value from list
            ImageViewer.View = 8
            ImageViewer.ViewSize = Val(cboZoom.Text) '/ 100
    End Select
    cboZoom.SelStart = 0
    cboZoom.SelLength = Len(cboZoom.Text)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call IniPutForm(Me)
End Sub


Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim s As String
    Dim i As Long
    Dim FileName As String
    Dim Copies As Integer
    
    Select Case Button.Key
        Case "save"
            If lvThumbs.ListItems.count > 0 Then
                For i = 1 To lvThumbs.ListItems.count
                    lvThumbs.ListItems(i).selected = i = lvThumbs.SelectedItem.Index
                Next
                FileName = ImageViewer.FileName
                If VBGetSaveFileName(FileName, , , "Document Image (*.tif)|*.tif", 5, , , "tif", Me.hwnd) Then
                    Call ImageViewer.Save(StripExtension(FileName), "tif")
                End If
            End If
            
        Case "mail"
            s = ""
            For i = 1 To lvThumbs.ListItems.count
                If lvThumbs.ListItems(i).selected Then
                    s = s & ";" & Mid(lvThumbs.ListItems(i).Key, 2)
                End If
            Next
            s = Mid(s, 2)
            On Error Resume Next
            Call FMain.SendMail(True, " ", "", "", "", s)
            
        
        Case "print"
            For i = 1 To lvThumbs.ListItems.count
                lvThumbs.ListItems(i).selected = i = lvThumbs.SelectedItem.Index
            Next
            Call ImageViewer.PrintImage(True)
            
        Case "rotate"
            Call ImageViewer.Rotate90
            
    End Select

End Sub


