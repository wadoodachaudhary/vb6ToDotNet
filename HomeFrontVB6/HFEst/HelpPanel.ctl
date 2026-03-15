VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.UserControl HelpPanel 
   Alignable       =   -1  'True
   ClientHeight    =   7395
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3915
   KeyPreview      =   -1  'True
   ScaleHeight     =   7395
   ScaleWidth      =   3915
   Begin SHDocVwCtl.WebBrowser html 
      Height          =   1035
      Left            =   450
      TabIndex        =   7
      Top             =   690
      Width           =   1095
      ExtentX         =   1931
      ExtentY         =   1826
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      MousePointer    =   1  'Arrow
      ScaleHeight     =   375
      ScaleWidth      =   3915
      TabIndex        =   0
      Top             =   0
      Width           =   3915
      Begin VB.Image imgEmail 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   3300
         MousePointer    =   1  'Arrow
         Picture         =   "HelpPanel.ctx":0000
         ToolTipText     =   "Submit a Request"
         Top             =   60
         Width           =   270
      End
      Begin VB.Image imgIndex 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   2160
         MousePointer    =   1  'Arrow
         Picture         =   "HelpPanel.ctx":03C2
         ToolTipText     =   "Table of Contents"
         Top             =   60
         Width           =   270
      End
      Begin VB.Image imgClose 
         Height          =   240
         Left            =   3600
         MousePointer    =   1  'Arrow
         Picture         =   "HelpPanel.ctx":0784
         ToolTipText     =   "Close"
         Top             =   60
         Width           =   270
      End
      Begin VB.Image imgForward 
         Height          =   240
         Left            =   2760
         MousePointer    =   1  'Arrow
         Picture         =   "HelpPanel.ctx":0B46
         ToolTipText     =   "Forward"
         Top             =   60
         Width           =   270
      End
      Begin VB.Image imgBack 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   2460
         MousePointer    =   1  'Arrow
         Picture         =   "HelpPanel.ctx":0F08
         ToolTipText     =   "Back"
         Top             =   60
         Width           =   270
      End
      Begin VB.Image imgFind 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   3060
         MousePointer    =   1  'Arrow
         Picture         =   "HelpPanel.ctx":12CA
         ToolTipText     =   "Search"
         Top             =   60
         Width           =   270
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   1
         X1              =   0
         X2              =   30005
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   0
         X2              =   30005
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "How To..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         MousePointer    =   1  'Arrow
         TabIndex        =   1
         Top             =   90
         UseMnemonic     =   0   'False
         Width           =   855
      End
   End
   Begin VB.PictureBox picFind 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   6915
      Left            =   0
      ScaleHeight     =   6915
      ScaleWidth      =   3855
      TabIndex        =   2
      Top             =   1740
      Visible         =   0   'False
      Width           =   3855
      Begin VSFlex8Ctl.VSFlexGrid grdFind 
         Height          =   4095
         Left            =   300
         TabIndex        =   6
         Top             =   1140
         Width           =   3375
         _cx             =   5953
         _cy             =   7223
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483629
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   0
         HighLight       =   0
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   0
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   0
         Cols            =   2
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"HelpPanel.ctx":168C
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   2
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.TextBox txtFind 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   300
         TabIndex        =   4
         Top             =   420
         Width           =   3075
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Topics Found:"
         Height          =   195
         Left            =   180
         TabIndex        =   5
         Top             =   900
         UseMnemonic     =   0   'False
         Width           =   1020
      End
      Begin VB.Image imgGo 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   3420
         Picture         =   "HelpPanel.ctx":16C9
         ToolTipText     =   "Search"
         Top             =   405
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "What word or phrase would you like to search for?"
         Height          =   195
         Left            =   180
         TabIndex        =   3
         Top             =   180
         UseMnemonic     =   0   'False
         Width           =   3555
      End
   End
End
Attribute VB_Name = "HelpPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Private mSearchMode As Boolean
Private Searching As Boolean
Private CancelSearch As Boolean

Public Event Hide()

Private Sub grdFind_GotFocus()
On Error Resume Next
    grdFind.CellBackColor = vbHighlight
    grdFind.CellForeColor = vbHighlightText
End Sub
Private Sub grdFind_LostFocus()
On Error Resume Next
    grdFind.CellBackColor = vbButtonFace
    grdFind.CellForeColor = vbWindowText
End Sub
Private Sub grdFind_RowColChange()
On Error Resume Next
    grdFind.Cell(flexcpBackColor, 0, 0, grdFind.Rows - 1, 0) = vbWindowBackground
    grdFind.Cell(flexcpForeColor, 0, 0, grdFind.Rows - 1, 0) = vbWindowText
    grdFind.CellBackColor = vbHighlight
    grdFind.CellForeColor = vbHighlightText
End Sub

Private Sub html_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
On Error Resume Next
    Dim IndexFile  As String
    Dim FileName   As String
    Dim filenumber As Integer
    Dim htmlText   As String
    
    
    
    Static Callback  As Boolean
    'things get weird with index.db files which are needed so that back/forward
    'buttons operate properly.
    If Not Callback And FileTitle("" & URL) = "index.db" Then
        Callback = False
        URL = Mid("" & URL, 1, Len("" & URL) - 9)
    End If
    
    'if file is given
    If FileExt("" & URL) <> "" Then
        'show it
        Callback = False
        Cancel = False
    Else
        'if index exists
        If PathExists(PathAppend("" & URL, "index.htm")) Then
            'show it
            Cancel = True
            Call html.Navigate(PathAppend("" & URL, "index.htm"))
        Else
            
            'remove any existing index
            IndexFile = PathAppend("" & URL, "index.db")
            Call SetAttr(IndexFile, vbNormal)
            Call Kill(IndexFile)
            
            'create a directory listing
            htmlText = "<html><head><title>" & FileTitle("" & URL) & "</title></head><body>" & vbCrLf
            FileName = Dir(PathAppend("" & URL, "*.*"), vbDirectory, True)
            While FileName <> ""
                If FileName <> "." And FileName <> ".." Then
                    htmlText = htmlText & "<a href=""" & URL & "\" & FileName & """>" & StripExtension(FileName) & "</a><br>" & vbCrLf
                End If
                FileName = Dir()
            Wend
            htmlText = htmlText & "</body></html>"
            
            'write to temp file
            filenumber = FreeFile
            Open IndexFile For Output As filenumber
            Print #filenumber, htmlText
            Close #filenumber
            Call SetAttr(IndexFile, vbHidden)
            
            
            'load temp file
            Callback = True ' set callback so we know that we already deleted and recreated the file
            Cancel = True
            Call html.Navigate(IndexFile)
            
        End If
    
    End If
End Sub

Private Sub html_NavigateError(ByVal pDisp As Object, URL As Variant, Frame As Variant, StatusCode As Variant, Cancel As Boolean)
On Error Resume Next
Cancel = True
End Sub

Private Sub imgClose_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent Hide
End Sub

Private Sub imgEmail_Click()
On Error Resume Next
    Call ShellFile(UserControl.hwnd, "mailto:support@homefront-software.com?subject=Homefront (version " & App.Major & "." & App.Minor & ".0." & App.Revision & ")&body=Thank you for your email support request. Our support team will respond in the shortest time possible, but may be delayed due to the complexity of the problem or volume of requests for support. Please be as specific you can be with the details of your concern. This will assist our technicians in resolving your problem in a timely manner.", , False)
End Sub

Private Sub imgIndex_Click()
    Search = False
    Call html.Navigate(PathAppend(RegGetKey(HKEY_LOCAL_MACHINE, "SOFTWARE\Homefront", "HomeFrontLicensePath", App.path), "help\Estimating"))
End Sub

Private Sub txtFind_GotFocus()
    SelectAll txtFind
End Sub

Private Sub txtFind_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call imgGo_Click
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    Select Case True
        Case KeyCode = vbKeyEscape
            If Searching Then
                CancelSearch = True
            Else
                RaiseEvent Hide
            End If
    End Select
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
    html.Move -2 * Screen.TwipsPerPixelX, 23 * Screen.TwipsPerPixelY, UserControl.Width + (2 * Screen.TwipsPerPixelX), UserControl.Height + (0 * Screen.TwipsPerPixelY) - 23 * Screen.TwipsPerPixelY

    imgClose.left = UserControl.Width - 270 * 1
    imgEmail.left = UserControl.Width - 270 * 2
    imgFind.left = UserControl.Width - 270 * 3
    imgForward.left = UserControl.Width - 270 * 4
    imgBack.left = UserControl.Width - 270 * 5
    imgIndex.left = UserControl.Width - 270 * 6

    picFind.Move html.left, html.Top, html.Width, html.Height
    txtFind.Width = picFind.Width - 600 - imgGo.Width - 60
    imgGo.left = picFind.Width - 300 - imgGo.Width
    grdFind.Width = picFind.Width - 600
    grdFind.Height = picFind.Height - grdFind.Top - 300
End Sub

Private Sub grdFind_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call grdFind_DblClick
End Sub

Private Sub grdFind_DblClick()
On Error Resume Next
    picFind.Visible = False
    Call html.Navigate(grdFind.TextMatrix(grdFind.Row, 1))
End Sub

Private Sub UserControl_Initialize()
On Error Resume Next
    Call html.Navigate(PathAppend(RegGetKey(HKEY_LOCAL_MACHINE, "SOFTWARE\Homefront", "HomeFrontLicensePath", App.path), "help\estimating"))
End Sub


Private Sub imgGo_Click()
    If Trim(txtFind.Text) = "" Then Exit Sub
    
    grdFind.Rows = 0
    grdFind.ForeColor = vbWindowText
    grdFind.Enabled = True
    
    CancelSearch = False
    Searching = True
    Call FindFiles(PathAppend(RegGetKey(HKEY_LOCAL_MACHINE, "SOFTWARE\Homefront", "HomeFrontLicensePath", App.path), "help\estimating"), "*" & txtFind.Text & "*")
    Searching = False
    
    Call grdFind_LostFocus
    Call grdFind.AutoSize(0)
    
    If grdFind.Rows = 0 Then
        grdFind.Enabled = False
        grdFind.AddItem ""
        grdFind.AddItem "   No topics matching your search were found."
        grdFind.AddItem ""
        grdFind.AddItem "   Homefront suggests that you:"
        grdFind.AddItem "     - Check your spelling."
        grdFind.AddItem "     - Try a different or more general term."
        If Parse(txtFind.Text, , " ") > 1 Then
            grdFind.AddItem "     - Use fewer words to broaden your search."
        End If
        grdFind.AddItem "     - Use wildcards to broaden your search."
        grdFind.AddItem "         *  matches any number of characters."
        grdFind.AddItem "         ?  matches any single character."
        grdFind.Cell(flexcpBackColor, 0, 0, grdFind.Rows - 1, 0) = vbWindowBackground
        grdFind.ForeColor = vbHighlight
    Else
        grdFind.Enabled = True
    End If
    
End Sub

Function FindFiles(path As String, SearchStr As String)
On Error GoTo exitsub
    Dim FileName As String
    Dim DirName As String
    Dim dirNames() As String
    Dim i As Integer
    
    Screen.MousePointer = vbHourglass
    
    If Right(path, 1) <> "\" Then path = path & "\"
    
    'get list of sub directories
    ReDim dirNames(0)
    DirName = Dir(path, vbDirectory)
    While DirName <> ""
        If (DirName <> ".") And (DirName <> "..") Then
            If GetAttr(path & DirName) And vbDirectory Then
                dirNames(UBound(dirNames)) = DirName
                ReDim Preserve dirNames(UBound(dirNames) + 1)
            End If
        End If
        DirName = Dir()
        DoEvents
        If CancelSearch Then GoTo exitsub
    Wend
    
    'search this directory for matching files and directories
    FileName = Dir(path & SearchStr, vbDirectory)
    While Len(FileName) <> 0
        If (FileName <> ".") And (FileName <> "..") And FileName <> "index.db" Then
            grdFind.AddItem StripExtension(FileName) & vbTab & path & FileName
        End If
        FileName = Dir()
        DoEvents
        If CancelSearch Then GoTo exitsub
    Wend
    
    'recursively search each sub directory
    For i = 0 To UBound(dirNames) - 1
        Call FindFiles(path & dirNames(i) & "\", SearchStr)
        DoEvents
        If CancelSearch Then GoTo exitsub
    Next i
    
exitsub:
    Screen.MousePointer = vbDefault
End Function

Private Sub html_TitleChange(ByVal Text As String)
    lblTitle.Caption = Text
    lblTitle.tag = Text
End Sub
Private Sub imgBack_Click()
On Error Resume Next
    If Search Then
        Search = False
    Else
        html.GoBack
    End If
End Sub
Private Sub imgForward_Click()
On Error Resume Next
    If Search Then
        Search = False
    Else
        html.GoForward
    End If
End Sub
Private Sub imgFind_Click()
    Search = Not Search
End Sub
Public Property Get Search() As Boolean
    Search = mSearchMode
End Property
Public Property Let Search(RHS As Boolean)
    mSearchMode = RHS
    picFind.Visible = mSearchMode
    If picFind.Visible Then txtFind.SetFocus
    lblTitle.Caption = IIf(mSearchMode, "Find Help...", lblTitle.tag)
End Property


