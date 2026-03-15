VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{C46A8909-8107-11D8-8671-00C1261173F0}#2.3#0"; "TwainControlX.ocx"
Begin VB.Form FDocuments 
   Caption         =   "Documents"
   ClientHeight    =   5460
   ClientLeft      =   11400
   ClientTop       =   2055
   ClientWidth     =   8235
   FillColor       =   &H00FF0000&
   Icon            =   "FDocuments.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5460
   ScaleWidth      =   8235
   Begin VB.PictureBox picCanvas 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   495
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   8
      Top             =   4650
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picRoot 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   135
      Picture         =   "FDocuments.frx":058A
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   7
      Top             =   4635
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picFolderFull 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   510
      Picture         =   "FDocuments.frx":0B14
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   5
      Top             =   4335
      Visible         =   0   'False
      Width           =   240
      Begin VB.PictureBox Picture2 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   0
         Picture         =   "FDocuments.frx":109E
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   6
         Top             =   1620
         Visible         =   0   'False
         Width           =   240
      End
   End
   Begin VB.PictureBox picFolderEmpty 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   150
      Picture         =   "FDocuments.frx":1628
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   4
      Top             =   4350
      Visible         =   0   'False
      Width           =   240
   End
   Begin VSFlex8Ctl.VSFlexGrid gData 
      Height          =   3555
      Left            =   150
      TabIndex        =   0
      Top             =   690
      Width           =   7815
      _cx             =   13785
      _cy             =   6271
      Appearance      =   0
      BorderStyle     =   0
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
      BackColorSel    =   -2147483645
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   3
      GridLines       =   0
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   13
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FDocuments.frx":1BB2
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
      ExplorerBar     =   2
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   2
      Editable        =   2
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   1
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
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   8235
      _ExtentX        =   14526
      _ExtentY        =   1058
      ButtonWidth     =   1323
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Open"
            Key             =   "Open"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Print"
            Key             =   "Print"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save As"
            Key             =   "SaveAs"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Remove"
            Key             =   "Delete"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Find"
            Key             =   "Find"
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin TwainControlX.Twain Twain1 
         Height          =   480
         Left            =   7260
         TabIndex        =   3
         Top             =   0
         Visible         =   0   'False
         Width           =   480
         CurrentDevice   =   -1
         UseInterface    =   -1  'True
         WaitForAcquire  =   -1  'True
         DoubleBuffered  =   0   'False
         Enabled         =   -1  'True
         Object.Visible         =   -1  'True
         Cursor          =   0
         HelpType        =   0
         HelpKeyword     =   ""
         JPEGQuality     =   90
         PixelType       =   -1
         ShowProgress    =   0   'False
         Units           =   0
         Resolution      =   -1
         ImageLeft       =   0
         ImageTop        =   0
         ImageRight      =   0
         ImageBottom     =   0
         XResolution     =   -1
         YResolution     =   -1
         AppName         =   "Ciansoft TwainControlX"
         UseADF          =   0   'False
         MultiImage      =   0   'False
         ImagesToRead    =   0
         KeepImages      =   0   'False
         SelectedImage   =   1
         AutoDeskew      =   0   'False
         Contrast        =   -1
         Brightness      =   -1
         AutoBright      =   0   'False
         DuplexEnabled   =   0   'False
         ClearBeforeAcquire=   -1  'True
         Threshold       =   -1
         BlankTol        =   100
      End
      Begin MSComctlLib.ImageList LargeIcons 
         Left            =   6615
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   10
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDocuments.frx":1D86
               Key             =   "Find"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDocuments.frx":2660
               Key             =   "Open"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDocuments.frx":2F3A
               Key             =   "Preview"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDocuments.frx":3814
               Key             =   "Send"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDocuments.frx":40EE
               Key             =   "New"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDocuments.frx":49C8
               Key             =   "Save"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDocuments.frx":52A2
               Key             =   "SaveAs"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDocuments.frx":5B7C
               Key             =   "Delete"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDocuments.frx":6456
               Key             =   "Print"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FDocuments.frx":1CE18
               Key             =   "Add"
            EndProperty
         EndProperty
      End
      Begin HFSystem.VBCombo cboScanner 
         Height          =   240
         Left            =   3570
         TabIndex        =   2
         Top             =   150
         Width           =   2985
         _ExtentX        =   5265
         _ExtentY        =   423
         Style           =   2
         Text            =   "Combo1"
      End
   End
   Begin VB.Menu mnuGrid 
      Caption         =   "<Grid>"
      Visible         =   0   'False
      Begin VB.Menu mnuGridSub 
         Caption         =   "Remove this column"
         Index           =   0
      End
      Begin VB.Menu mnuColumns 
         Caption         =   "Insert a column"
         Begin VB.Menu mnuColumnsSub 
            Caption         =   "(none available)"
            Index           =   0
         End
      End
   End
   Begin VB.Menu mnuDocClass 
      Caption         =   "<DocClass>"
      Visible         =   0   'False
      Begin VB.Menu mnuDocClassSub 
         Caption         =   "Add File..."
         Index           =   0
      End
      Begin VB.Menu mnuDocClassSub 
         Caption         =   "Scan File..."
         Index           =   1
      End
   End
   Begin VB.Menu mnuRevisions 
      Caption         =   "<File Popup>"
      Visible         =   0   'False
      Begin VB.Menu mnuRevisionsSub 
         Caption         =   "Open"
         Index           =   0
      End
      Begin VB.Menu mnuRevisionsSub 
         Caption         =   "Print"
         Index           =   1
      End
      Begin VB.Menu mnuRevisionsSub 
         Caption         =   "Save As..."
         Index           =   2
      End
      Begin VB.Menu mnuRevisionsSub 
         Caption         =   "Remove"
         Index           =   3
      End
      Begin VB.Menu mnuRevisionsSub 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuRevisionsSub 
         Caption         =   "Add Revision..."
         Index           =   5
      End
   End
End
Attribute VB_Name = "FDocuments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Private Const SRCFILE = "FDocuments::"

Private MouseGrid As VSFlexGrid
Private MouseCol As Long

Private Const CLASS_ADDFILE = 0
Private Const CLASS_SCANFILE = 1

Private Const REVISION_OPEN = 0
Private Const REVISION_PRINT = 1
Private Const REVISION_SAVE = 2
Private Const REVISION_REMOVE = 3
Private Const REVISION_ADD = 5



Public Sub ShowForm(FormCaption As String, ParamArray ObjectIDsAndFolderNames())
        
    Dim i As Long
    Dim objid As String
    Dim Caption As String
    For i = 0 To UBound(ObjectIDsAndFolderNames) Step 2
        objid = Trim(ObjectIDsAndFolderNames(i))
        Caption = Trim(ObjectIDsAndFolderNames(i + 1))
        If objid <> "" And Caption <> "" Then
            Call LoadData(objid, Caption)
        End If
    Next
    Me.Show vbModal
End Sub

Public Sub LoadData(ObjectID As String, ObjectDescription As String)
On Error GoTo eh
With gData

    Dim s      As String
    Dim rs     As Recordset
    Dim Path   As String
    Dim Name   As String
    Dim RowType As String
    Dim relativepath As String
    Dim Folder As String
    Dim DocumentID As String
    Dim r      As Long
    Dim i      As Long
    Dim j      As Long
    
    
    'rowdata is folderkey
    'objectid\pathsegment1\pathsegment2\...\pathsegmentN\documentid
    
    
    
    'create root folder for objectid
    .AddItem ""
    r = .Rows - 1
    .IsSubtotal(r) = True
    .RowOutlineLevel(r) = 0
    .TextMatrix(r, .ColIndex("Name")) = ObjectDescription
    .TextMatrix(r, .ColIndex("RowType")) = "root"
    .TextMatrix(r, .ColIndex("ObjectID")) = ObjectID
    .Cell(flexcpPicture, r, .ColIndex("Name")) = picRoot.Image
    .Cell(flexcpFontBold, r, .ColIndex("Name")) = True
    relativepath = ObjectID
    .RowData(r) = relativepath
    
    
    
    'create display path folders --- LIMIT PATHS LOADED TO ONLY THOSE VISIBLE TO THE CURRENT USER....
    s = "select documentclass from dms_documentclasses order by 1"
    Set rs = HFApp.SqlExec(s, dbHomefront)
    While Not rs.EOF
        relativepath = ObjectID
        Path = "" & rs("DocumentClass")
        For i = 1 To Parse(Path, , "\")
            Folder = Parse(Path, i, "\")
            If Folder <> "" Then
                relativepath = relativepath & "\" & Folder
                j = .FindRow(relativepath)
                If j = -1 Then
                    .AddItem ""
                    r = .Rows - 1
                    .IsSubtotal(r) = True
                    .RowOutlineLevel(r) = i
                    .TextMatrix(r, .ColIndex("Name")) = Folder
                    .TextMatrix(r, .ColIndex("RowType")) = "folder"
                    .Cell(flexcpPicture, r, .ColIndex("Name")) = picFolderEmpty.Image
                    .Cell(flexcpFontBold, r, .ColIndex("Name")) = True
                    .RowData(r) = relativepath
                    .TextMatrix(r, .ColIndex("ObjectID")) = ObjectID
                End If
            End If
        Next
        .TextMatrix(r, .ColIndex("DocumentClass")) = "" & rs("DocumentClass")
        .TextMatrix(r, .ColIndex("RowType")) = "class"
        rs.MoveNext
    Wend
        
        
    'create document and revisions
    s = ""
    s = s & "select d.DocumentClass,d.DocumentID,d.Name,r.Revision,d.ObjectID,r.CreatedDate,r.CreatedBy,r.FileTitle,r.FileName,r.Notes,r.Embedded,r.AssetID,r.FileExt" & vbCrLf
    s = s & " " & vbCrLf
    s = s & "from dms_documents d " & vbCrLf
    s = s & "join dms_revisions r on d.documentid=r.documentid" & vbCrLf
    s = s & "where d.objectid=" & DbQuote(Str, ObjectID) & vbCrLf
    s = s & "order by d.Name,r.revision desc" & vbCrLf
    Set rs = HFApp.SqlExec(s, dbHomefront)
    While Not rs.EOF
        
        'find document/add revision
        Path = ObjectID & "\" & rs("DocumentClass") & "\" & rs("DocumentID")
        j = .FindRow(Path)
        If j <> -1 Then
            r = .GetNodeRow(j, flexNTLastChild)
            If r = -1 Then
                r = j + 1
            Else
                r = r + 1
            End If
            .AddItem "", r
            Name = "" & rs("FileTitle")
            RowType = "revision"
            Path = ObjectID & "\" & rs("DocumentClass") & "\" & rs("DocumentID") & "\" & rs("Revision")
            .Cell(flexcpForeColor, r, 0, r, .Cols - 1) = vbHighlight
            .GetNode(j).Expanded = False
        Else
            'find path/add document
            Path = ObjectID & "\" & rs("DocumentClass")
            j = .FindRow(Path)
            If j <> -1 Then
                r = .GetNodeRow(j, flexNTLastChild)
                If r = -1 Then
                    r = j + 1
                Else
                    r = r + 1
                End If
                .AddItem "", r
                RowType = "document"
                Name = "" & rs("Name")
                Path = Path & "\" & rs("DocumentID")
            End If
        End If
        
        .TextMatrix(r, .ColIndex("ObjectID")) = ObjectID
        .RowData(r) = Path
        .IsSubtotal(r) = True
        .RowOutlineLevel(r) = .RowOutlineLevel(j) + 1
        .TextMatrix(r, .ColIndex("Name")) = Name
        .TextMatrix(r, .ColIndex("RowType")) = RowType
        .TextMatrix(r, .ColIndex("DocumentID")) = "" & rs("DocumentID")
        .TextMatrix(r, .ColIndex("Revision")) = "" & rs("Revision")
        .TextMatrix(r, .ColIndex("DocumentClass")) = "" & rs("DocumentClass")
        .TextMatrix(r, .ColIndex("Notes")) = "" & rs("Notes")
        .TextMatrix(r, .ColIndex("Filename")) = "" & rs("Filename")
        .TextMatrix(r, .ColIndex("FileTitle")) = "" & rs("FileTitle")
        .TextMatrix(r, .ColIndex("CreatedDate")) = "" & rs("CreatedDate")
        .TextMatrix(r, .ColIndex("CreatedBy")) = "" & rs("CreatedBy")
        .Cell(flexcpPicture, r, .ColIndex("Name")) = FileIcon("" & rs("FileExt"))
        .TextMatrix(r, .ColIndex("Embedded")) = "" & rs("Embedded")
        .TextMatrix(r, .ColIndex("AssetID")) = "" & rs("AssetID")
        
        rs.MoveNext
    Wend
        
End With
Exit Sub
eh: Call errHandler(SRCFILE & "LoadData")
End Sub




Private Sub Form_Resize()
On Error Resume Next
    gData.Move 0, Toolbar.Height, Me.ScaleWidth, Me.ScaleHeight - Toolbar.Height
    cboScanner.Move Toolbar.Width - cboScanner.Width - 120, (Toolbar.Height - cboScanner.Height) / 2
End Sub


Private Sub Form_Load()
On Error GoTo eh
    
    Call SetToolbarIcons(Toolbar, LargeIcons)
    Call IniGetForm(Me)
    
    With gData
        .Rows = 1
        .OutlineBar = flexOutlineBarCompleteLeaf
        .OutlineCol = .ColIndex("Name")
    End With
'    Call IniGetGrid(Me, gData)

Exit Sub
eh: Call errHandler(SRCFILE & "Form_Load")
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Call IniPutForm(Me)
    Call IniPutGrid(Me, gData)
    Call IniPut(AppIni, Me.Name, "Scanner", cboScanner.Text)
    
On Error Resume Next: Call Kill(PathAppend(TempPath, "*.*"))
End Sub

Private Function FileIcon(ext As String) As IPictureDisp
    Dim s As String
    Dim FInfo As New ClsFileInfo
        
    s = PathAppend(AppWorkingFolder, "tmp." & ext)
    Call CreateFile(s)
    
    Set picCanvas.Picture = New StdPicture
    FInfo.FullPathName = s
    Call ImageList_Draw(FInfo.hSmlIList, FInfo.hSmlIcon, picCanvas.hDC, 0, 0, ILD_TRANSPARENT)
    
    Set FileIcon = picCanvas.Image
    
End Function


Private Sub CreateFile(Filename As String)
    Dim i As Integer
    i = FreeFile
    Open Filename For Output As #i
    Close #i
End Sub


Private Sub gData_AfterMoveColumn(ByVal Col As Long, Position As Long)
    gData.OutlineCol = gData.ColIndex("Name")
End Sub

Private Sub gData_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With gData
    .ComboList = ""
    If .ColKey(Col) = "Notes" Then .ComboList = "|..."
    Select Case .TextMatrix(Row, .ColIndex("RowType"))
        Case "Document":  Cancel = Not IsIn(.ColKey(Col), "Notes", "Name")
        Case "Revision":  Cancel = Not IsIn(.ColKey(Col), "Notes")
        Case Else:        Cancel = True
    End Select
    End With
End Sub

Private Sub gData_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim s As String
    s = gData.TextMatrix(Row, Col)
    If FComments.Edit(s, gData, , "Notes") Then
        gData.TextMatrix(Row, Col) = s
        Call gData_AfterEdit(Row, Col)
    End If
End Sub

Private Sub gData_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With gData
        If Button <> vbRightButton Or .MouseRow < 0 Then Exit Sub
        
        Select Case True
        Case .MouseRow = 0
            Call ShowColumnMenu(gData)
        
        Case .TextMatrix(.Row, .ColIndex("RowType")) = "class"
            PopupMenu mnuDocClass
            
        Case IsIn(.TextMatrix(.Row, .ColIndex("RowType")), "document", "revision")
            PopupMenu mnuRevisions
            
        
        End Select
    
    End With

End Sub

Private Sub gData_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With gData
    Select Case .ColKey(Col)
        Case "Name": Cancel = .EditText = ""
    End Select
    End With
End Sub

Private Sub gData_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim s As String
    With gData
    s = ""
    Select Case .ColKey(Col)
        Case "Name":  s = "update dms_documents set name =" & DbQuote(Str, .TextMatrix(Row, Col)) & " where documentid=" & DbQuote(Str, .TextMatrix(Row, .ColIndex("DocumentID")))
        Case "Notes": s = "update dms_revisions set notes=" & DbQuote(Str, .TextMatrix(Row, Col)) & " where documentid=" & DbQuote(Str, .TextMatrix(Row, .ColIndex("DocumentID"))) & " and revision=" & DbQuote(Num, .TextMatrix(Row, .ColIndex("Revision")))
    End Select
    If s <> "" Then Call HFApp.SqlExec(s)
    End With
End Sub


Private Sub ShowColumnMenu(Grid As VSFlexGrid, Optional Sortable As Boolean = True, Optional Hideable As Boolean = True)
On Error GoTo eh
    Dim i As Long
    Dim j As Long
    
    Set MouseGrid = Grid
    MouseCol = MouseGrid.MouseCol
   

    'load Grid column names
    mnuColumnsSub(0).Visible = True
    For i = mnuColumnsSub.UBound To 1 Step -1
        Unload mnuColumnsSub(i)
    Next
    For i = 0 To MouseGrid.Cols - 1
        If MouseGrid.ColHidden(i) And MouseGrid.TextMatrix(0, i) <> "" Then
            j = j + 1
            Load mnuColumnsSub(j)
            mnuColumnsSub(j).tag = MouseGrid.ColKey(i)
            mnuColumnsSub(j).Caption = MouseGrid.TextMatrix(0, i)
            mnuColumnsSub(j).Visible = True
            mnuColumnsSub(j).Enabled = True
        End If
    Next
    
    
    
    If j = 0 Then mnuColumnsSub(0).Caption = "(none available)"
    mnuColumnsSub(0).Visible = j = 0
    mnuColumnsSub(0).Enabled = False
    
    'show menu
    PopupMenu mnuGrid

    Exit Sub
eh: Call errHandler(SRCFILE & "ShowColumnMenu")
End Sub

Private Sub mnuColumnsSub_Click(Index As Integer)
On Error Resume Next
    MouseGrid.ColHidden(MouseGrid.ColIndex(mnuColumnsSub(Index).tag)) = False
    MouseGrid.ColPosition(MouseGrid.ColIndex(mnuColumnsSub(Index).tag)) = IIf(MouseCol < 0, MouseGrid.Cols - 1, MouseCol)
    gData.OutlineCol = gData.ColIndex("Name")
End Sub
Private Sub mnuGridSub_Click(Index As Integer)
On Error Resume Next
    Dim i As Long
    Select Case Index
        Case 0
            For i = 0 To MouseGrid.Cols - 1
                If MouseGrid.ColHidden(i) = False Then
                    MouseGrid.ColHidden(MouseCol) = True
                    Exit Sub
                End If
            Next
            gData.OutlineCol = gData.ColIndex("Name")
    End Select
End Sub

Private Sub mnuDocClassSub_Click(Index As Integer)
    With gData
    Select Case Index
        Case CLASS_ADDFILE:   Call AddDocument(gData.Row, "")
        Case CLASS_SCANFILE
    End Select
    End With
End Sub


Private Sub mnuRevisionsSub_Click(Index As Integer)
With gData
    Dim s As String
    Dim Filename As String
    
    If Not IsIn(.TextMatrix(.Row, .ColIndex("RowType")), "document", "revision") Then Exit Sub
    
    Filename = .TextMatrix(.Row, .ColIndex("Filename"))
    
    Select Case Index
    Case REVISION_OPEN:    Call ShellFile(Me.hwnd, Filename)
    Case REVISION_PRINT:   Call ShellFile(Me.hwnd, Filename, , , True)
    Case REVISION_REMOVE:  Call Delete(.Row)
    Case REVISION_ADD:     Call AddRevision(.Row, "")
    Case REVISION_SAVE:
        If VBGetSaveFileName(s, , , , , , "Save As", , Me.hwnd) Then
            s = ForceExt(s, FileExt(Filename))
            Call FileCopy(Filename, s)
        End If
    End Select

End With
End Sub


Private Sub AddDocument(Row As Long, Filename As String)
On Error GoTo eh
With gData

    Dim s As String
    Dim DocumentID As String
    Dim r As Long
    
    If Filename = "" Then Call VBGetOpenFileName(Filename, , , , , True, , , , "Attach File")
    If Filename = "" Then Exit Sub
    
    s = "exec DMS_AddDocument"  'objectid, filename, createdby, documentclass, notes
    s = s & " " & DbQuote(Str, .TextMatrix(Row, .ColIndex("ObjectID")))
    s = s & ", " & DbQuote(Str, Filename)
    s = s & ", " & DbQuote(Str, HFApp.LoginID)
    s = s & ", " & DbQuote(Str, .TextMatrix(Row, .ColIndex("DocumentClass")))
    s = s & ",''"
    DocumentID = HFApp.SqlExec(s, dbHomefront)(0)
    
    
    'add to grid
    r = .GetNodeRow(Row, flexNTLastChild)
    If r = -1 Then
        r = Row + 1
    Else
        r = r + 1
    End If
    .AddItem "", r

    .TextMatrix(r, .ColIndex("ObjectID")) = .TextMatrix(Row, .ColIndex("ObjectID"))
    .RowData(r) = .TextMatrix(Row, .ColIndex("ObjectID")) & "\" & .TextMatrix(Row, .ColIndex("DocumentClass")) & "\" & DocumentID
    .IsSubtotal(r) = True
    .RowOutlineLevel(r) = .RowOutlineLevel(Row) + 1
    .TextMatrix(r, .ColIndex("Name")) = FileTitle(Filename)
    .TextMatrix(r, .ColIndex("RowType")) = "document"
    .TextMatrix(r, .ColIndex("DocumentID")) = DocumentID
    .TextMatrix(r, .ColIndex("Revision")) = "0"
    .TextMatrix(r, .ColIndex("DocumentClass")) = .TextMatrix(Row, .ColIndex("DocumentClass"))
    .TextMatrix(r, .ColIndex("Notes")) = ""
    .TextMatrix(r, .ColIndex("Filename")) = Filename
    .TextMatrix(r, .ColIndex("FileTitle")) = FileTitle(Filename)
    .TextMatrix(r, .ColIndex("CreatedDate")) = Now()
    .TextMatrix(r, .ColIndex("CreatedBy")) = HFApp.LoginID
    .Cell(flexcpPicture, r, .ColIndex("Name")) = FileIcon(FileExt(Filename))
    .TextMatrix(r, .ColIndex("Embedded")) = "false"
    .TextMatrix(r, .ColIndex("AssetID")) = ""
    
End With
Exit Sub
eh: Call errHandler(SRCFILE & "AddDocument")
End Sub

Private Sub Delete(Row As Long)
With gData
    Dim s As String
    Dim r As Long
    If .TextMatrix(Row, .ColIndex("RowType")) = "document" Then
        s = ""
        s = s & "delete DMS_Documents where DocumentID=" & DbQuote(Str, .TextMatrix(Row, .ColIndex("DocumentID"))) & vbCrLf
        s = s & "delete DMS_Revisions where DocumentID=" & DbQuote(Str, .TextMatrix(Row, .ColIndex("DocumentID")))
        Call HFApp.SqlExec(s, dbHomefront)
        r = .GetNodeRow(Row, flexNTLastChild)
        While r <> -1
            .RemoveItem r
            r = .GetNodeRow(Row, flexNTLastChild)
        Wend
        .RemoveItem Row
    Else
        s = "delete DMS_Revisions where DocumentID=" & DbQuote(Str, .TextMatrix(Row, .ColIndex("DocumentID"))) & " and Revision=" & DbQuote(Num, .TextMatrix(Row, .ColIndex("Revision")))
        Call HFApp.SqlExec(s, dbHomefront)
        .RemoveItem Row
    End If
End With
End Sub



Private Sub AddRevision(Row As Long, Filename As String)
On Error GoTo eh
With gData

    Dim s As String
    Dim FInfo As ClsFileInfo
    Dim Revision As Long
    Dim r As Long
    Dim docName As String
    
    If Filename = "" Then Call VBGetOpenFileName(Filename, , , , , True, , , , "Attach File")
    If Filename = "" Then Exit Sub
    
    s = "exec DMS_AddRevision"
    s = s & " " & DbQuote(Str, .TextMatrix(Row, .ColIndex("DocumentID")))
    s = s & ", " & DbQuote(Str, Filename)
    s = s & ", " & DbQuote(Str, HFApp.LoginID)
    s = s & ",''"
    Revision = HFApp.SqlExec(s, dbHomefront)(0)
    
    
    'convert current row (document) to a revision node
    r = Row
    .Cell(flexcpForeColor, r, 0, r, .Cols - 1) = vbHighlight
    .RowOutlineLevel(r) = .RowOutlineLevel(r) + 1
    .TextMatrix(r, .ColIndex("RowType")) = "revision"
    docName = .TextMatrix(r, .ColIndex("Name"))
    .TextMatrix(r, .ColIndex("Name")) = .TextMatrix(r, .ColIndex("FileTitle"))

        
    'then insert this new revision as the new document node
    Row = r + 1
    .AddItem "", r
    .TextMatrix(r, .ColIndex("ObjectID")) = .TextMatrix(Row, .ColIndex("ObjectID"))
    .RowData(r) = .TextMatrix(Row, .ColIndex("ObjectID")) & "\" & .TextMatrix(Row, .ColIndex("DocumentClass")) & "\" & .TextMatrix(Row, .ColIndex("DocumentID")) & "\" & Revision
    .IsSubtotal(r) = True
    .RowOutlineLevel(r) = .RowOutlineLevel(Row) - 1
    .TextMatrix(r, .ColIndex("Name")) = docName
    .TextMatrix(r, .ColIndex("RowType")) = "document"
    .TextMatrix(r, .ColIndex("DocumentID")) = .TextMatrix(Row, .ColIndex("DocumentID"))
    .TextMatrix(r, .ColIndex("Revision")) = Revision
    .TextMatrix(r, .ColIndex("DocumentClass")) = .TextMatrix(Row, .ColIndex("DocumentClass"))
    .TextMatrix(r, .ColIndex("Notes")) = ""
    .TextMatrix(r, .ColIndex("Filename")) = Filename
    .TextMatrix(r, .ColIndex("FileTitle")) = FileTitle(Filename)
    .TextMatrix(r, .ColIndex("CreatedDate")) = Now()
    .TextMatrix(r, .ColIndex("CreatedBy")) = HFApp.LoginID
    .Cell(flexcpPicture, r, .ColIndex("Name")) = FileIcon(FileExt(Filename))
    .TextMatrix(r, .ColIndex("Embedded")) = "false"
    .TextMatrix(r, .ColIndex("AssetID")) = ""
    .Row = r
    
End With
Exit Sub
eh: Call errHandler(SRCFILE & "AddRevision")
End Sub


Private Sub OLEDragDrop(Data, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo eh
    Dim i As Long
    Dim f As Long
    Dim files() As String
    Dim dda As DragDropAttachment
    Dim abyData() As Byte
    
    
    If Data.GetFormat(vbCFFiles) Then
        'file dragged from desktop
        If TypeOf Data Is DataObject Or TypeOf Data Is MSComctlLib.DataObject Then
            'dropped on normal targets like form or frame
            ReDim files(Data.files.Count)
            For i = 1 To Data.files.Count
                files(i) = Data.files(i)
            Next
        ElseIf TypeOf Data Is VSFlex8Ctl.VSDataObject Then
            'dropped on flexgrid who cant follow the dang specs!
            ReDim files(Data.FileCount)
            For i = 1 To Data.FileCount
                files(i) = Data.files(i - 1)
            Next
        End If
        
    Else
        'file dragged from outlook
        Set dda = New DragDropAttachment
        Set dda.Source = Data
        ReDim files(dda.Count())
        For i = 1 To dda.Count()
            'get filename from outlook, add to list of files
            files(i) = PathAppend(TempPath(), dda.Filename(i - 1))
            'save file to disk
            On Error Resume Next
            Call CreatePath("", FilePath(files(i)))
            Kill files(i)
            On Error GoTo eh
            f = FreeFile()
            Open files(i) For Binary As #f
            
            'dont do this. must save to bytarray first otherwise it corrupts the file?? maybe.
            'Put #f, , dda.Attachment(i - 1)
            
            abyData = dda.Attachment(i - 1)
            Put #f, , abyData
            
            Close #f
                        
        Next
    
    End If
    
    
    'attach each file
    With gData
        .Row = .MouseRow
        If .TextMatrix(.Row, .ColIndex("RowType")) = "class" Then
            For i = 1 To UBound(files)
                Call AddDocument(.MouseRow, files(i))
            Next
        End If
        If .TextMatrix(.Row, .ColIndex("RowType")) = "document" Then
            For i = 1 To UBound(files)
                Call AddRevision(.MouseRow, files(i))
            Next
        End If
    End With
    
    Effect = vbDropEffectNone
    
Exit Sub
eh: Call errHandler(SRCFILE & "OLEDragDrop")
End Sub


Private Sub gData_OLEDragDrop(Data As VSFlex8Ctl.VSDataObject, Effect As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub

Private Sub gData_OLEDragOver(Data As VSFlex8Ctl.VSDataObject, Effect As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, State As Integer)
On Error Resume Next
    With gData
        Effect = IIf(IsIn(.TextMatrix(.MouseRow, .ColIndex("RowType")), "document", "class"), vbDropEffectCopy, vbDropEffectNone)
    End With
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Select Case UCase(Trim(Button.Key))
        Case "FIND":      Call FFind.ShowForm(gData, False)
    End Select
    
    With gData
        If .Row < 1 Then Exit Sub
        If .TextMatrix(.Row, .ColIndex("FileName")) = "" Then Exit Sub
    End With
    
    Select Case UCase(Trim(Button.Key))
        Case "OPEN":      Call mnuRevisionsSub_Click(REVISION_OPEN)
        Case "PRINT":     Call mnuRevisionsSub_Click(REVISION_PRINT)
        Case "SAVEAS":    Call mnuRevisionsSub_Click(REVISION_SAVE)
        Case "DELETE":    Call mnuRevisionsSub_Click(REVISION_REMOVE)
    End Select
End Sub

