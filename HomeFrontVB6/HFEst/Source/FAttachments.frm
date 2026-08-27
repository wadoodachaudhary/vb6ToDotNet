VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form FAttachments 
   Caption         =   "Attachments"
   ClientHeight    =   5370
   ClientLeft      =   3690
   ClientTop       =   2325
   ClientWidth     =   7275
   FillColor       =   &H00FF0000&
   Icon            =   "FAttachments.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5370
   ScaleWidth      =   7275
   Begin VB.PictureBox picEmbedded 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1350
      Picture         =   "FAttachments.frx":058A
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   3
      Top             =   3990
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picFolder 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1050
      Picture         =   "FAttachments.frx":0B14
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   2
      Top             =   3960
      Visible         =   0   'False
      Width           =   240
   End
   Begin VSFlex8Ctl.VSFlexGrid gData 
      Height          =   3555
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   6555
      _cx             =   11562
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   3
      GridLines       =   0
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FAttachments.frx":109E
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
   Begin VB.PictureBox picCanvas 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   660
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   1
      Top             =   3990
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Menu mnuColumns 
      Caption         =   "<Columns>"
      Visible         =   0   'False
      Begin VB.Menu mnuColumnsSub 
         Caption         =   "Name"
         Index           =   0
      End
      Begin VB.Menu mnuColumnsSub 
         Caption         =   "Classification"
         Index           =   1
      End
      Begin VB.Menu mnuColumnsSub 
         Caption         =   "Location"
         Index           =   3
      End
      Begin VB.Menu mnuColumnsSub 
         Caption         =   "Size"
         Index           =   4
      End
      Begin VB.Menu mnuColumnsSub 
         Caption         =   "Date Modified"
         Index           =   5
      End
      Begin VB.Menu mnuColumnsSub 
         Caption         =   "Date Created"
         Index           =   6
      End
      Begin VB.Menu mnuColumnsSub 
         Caption         =   "Date Attached"
         Index           =   7
      End
   End
   Begin VB.Menu mnuFolders 
      Caption         =   "<Folder Popup>"
      Visible         =   0   'False
      Begin VB.Menu mnuFoldersSub 
         Caption         =   "Insert File..."
         Index           =   0
      End
      Begin VB.Menu mnuFoldersSub 
         Caption         =   "Link to File..."
         Index           =   1
      End
   End
   Begin VB.Menu mnuFiles 
      Caption         =   "<File Popup>"
      Visible         =   0   'False
      Begin VB.Menu mnuFilesSub 
         Caption         =   "Open"
         Index           =   0
      End
      Begin VB.Menu mnuFilesSub 
         Caption         =   "Print"
         Index           =   1
      End
      Begin VB.Menu mnuFilesSub 
         Caption         =   "Save As..."
         Index           =   2
      End
      Begin VB.Menu mnuFilesSub 
         Caption         =   "Remove"
         Index           =   3
      End
      Begin VB.Menu mnuFilesSub 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuFilesSub 
         Caption         =   "Cut"
         Enabled         =   0   'False
         Index           =   5
      End
      Begin VB.Menu mnuFilesSub 
         Caption         =   "Copy"
         Index           =   6
      End
      Begin VB.Menu mnuFilesSub 
         Caption         =   "Paste"
         Enabled         =   0   'False
         Index           =   7
      End
      Begin VB.Menu mnuFilesSub 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuFilesSub 
         Caption         =   "Send to..."
         Index           =   9
      End
      Begin VB.Menu mnuFilesSub 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu mnuFilesSub 
         Caption         =   "Properties..."
         Index           =   11
      End
   End
End
Attribute VB_Name = "FAttachments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const SRCFILE = "FAttachments::"

Private Const mcCOL_NAME = 0
Private Const mcCOL_CLASS = 1
Private Const mcCOL_LOCATION = 3
Private Const mcCOL_SIZE = 4
Private Const mcCOL_MODIFIED = 5
Private Const mcCOL_CREATED = 6
Private Const mcCOL_ATTACHED = 7

Private Const mcFILE_OPEN = 0
Private Const mcFILE_PRINT = 1
Private Const mcFILE_SAVE = 2
Private Const mcFILE_REMOVE = 3
Private Const mcFILE_CUT = 5
Private Const mcFILE_COPY = 6
Private Const mcFILE_PASTE = 7
Private Const mcFILE_SENDTO = 9
Private Const mcFILE_PROPERTIES = 11

Public Function AddFile(ObjectID As String, Embedded As Boolean, Optional FileName As String)
    Dim s As String
    Dim FInfo As ClsFileInfo
    
    If FileName = "" Then
        If VBGetOpenFileName(FileName, , , , , True, , , , "Attach File") Then
            If PathIsLocalPath(FileName) And Not Embedded Then
                If MsgBox("This is a local file.  It may not be accessible to other users." & vbCrLf & "Are you sure this is what you want to do?", vbQuestion Or vbYesNo, App.ProductName) = vbNo Then
                    FileName = ""
                End If
            End If
        End If
    End If
    If FileName = "" Then Exit Function

    Set FInfo = New ClsFileInfo
    FInfo.FullPathName = FileName
    
    If Not FInfo.FileExists Then
        MsgBox "file not found"
    Else
        s = ""
        s = s & "insert into attachments(objectid,filename,modifieddate,createddate,attacheddate,embedded)" & vbCrLf
        s = s & "values(" & DbQuote(Str, ObjectID) & vbCrLf
        s = s & "      ," & DbQuote(Str, FileName) & vbCrLf
        s = s & "      ," & DbQuote(DateTime, FInfo.ModifyTime) & vbCrLf
        s = s & "      ," & DbQuote(DateTime, FInfo.CreationTime) & vbCrLf
        s = s & "      ,getdate()" & vbCrLf
        s = s & "      ," & DbQuote(Bit, Embedded) & ")" & vbCrLf
        Call HFApp.SqlExec(s, dbHomefront)
        
        If Embedded Then
            Call DBPutFile(HFApp.Databases(dbHomefront), "Attachments WHERE objectid=" & DbQuote(Str, ObjectID) & " and filename=" & DbQuote(Str, FileName), "FileImage", FileName)
        End If
        
        If IsFormLoaded("FAttachments") Then Call AddFileToGrid(ObjectID, Embedded, FileName)
    End If

End Function

Public Sub Showform(FormCaption As String, ParamArray ObjectIDsAndFolderNames())
    Dim i      As Long
    Dim s      As String
    Dim rs     As Recordset
    Dim Folder As String
    Dim FInfo As New ClsFileInfo
    
    s = ""
    
    For i = 0 To UBound(ObjectIDsAndFolderNames) Step 2
        s = s & "union select " & DbQuote(Str, ObjectIDsAndFolderNames(i)) & " ObjectID," & DbQuote(Str, ObjectIDsAndFolderNames(i + 1)) & " Folder,a.filename,a.documentclass,a.filesize,a.filetitle,a.path,a.modifieddate,a.createddate,a.attacheddate,a.embedded from system_setup s left outer join attachments a on(a.objectid=" & DbQuote(Str, ObjectIDsAndFolderNames(i)) & ")" & vbCrLf
    Next
    s = Mid(s, 7) & "order by 1,6,3" 'folder,filetitle,filename
    Set rs = HFApp.SqlExec(s, dbHomefront)
    With gData
        .Rows = 1
        .OutlineCol = .ColIndex("FileTitle")
        .ExplorerBar = flexExMove
        .OutlineBar = flexOutlineBarSimpleLeaf
        While Not rs.EOF
            
            If Folder = "" Or UCase(Folder) <> UCase("" & rs("Folder")) Then
                Folder = "" & rs("folder")
               .AddItem ""
               
               .RowData(.Rows - 1) = "" & rs("ObjectID")
                .TextMatrix(.Rows - 1, .ColIndex("FileTitle")) = Folder
                .Cell(flexcpPicture, .Rows - 1, .ColIndex("Filetitle")) = picFolder.Picture
                .IsSubtotal(.Rows - 1) = True
                .RowOutlineLevel(.Rows - 1) = 0
            End If
        
            
            If "" & rs("FileName") <> "" Then
                .AddItem ""
                .RowData(.Rows - 1) = "" & rs("ObjectID")
                If "" & rs("Embedded") = "True" Then
                    s = PathAppend(AppWorkingFolder, "tmp." & FileExt("" & rs("FileName")))
                    Call CreateFile(s)
                    FInfo.FullPathName = s
                    .TextMatrix(.Rows - 1, .ColIndex("FileName")) = "" & rs("FileName")
                    .TextMatrix(.Rows - 1, .ColIndex("FileSize")) = FInfo.FormatFileSize(Val("" & rs("FileSize")))
                    .TextMatrix(.Rows - 1, .ColIndex("ModifiedDate")) = format("" & rs("ModifiedDate"), "general date")
                    .TextMatrix(.Rows - 1, .ColIndex("CreatedDate")) = format("" & rs("CreatedDate"), "general date")
                    .Cell(flexcpPicture, .Rows - 1, .ColIndex("Path")) = picEmbedded.Picture
                    .Cell(flexcpForeColor, .Rows - 1, .ColIndex("FileTitle")) = &HC00000 '&H00FF0000&  '&H800000
                Else
                    FInfo.FullPathName = "" & rs("FileName")
                    .TextMatrix(.Rows - 1, .ColIndex("FileName")) = "" & rs("FileName")
                    .TextMatrix(.Rows - 1, .ColIndex("FileSize")) = FInfo.FormatFileSize(FInfo.FileSize)
                    .TextMatrix(.Rows - 1, .ColIndex("ModifiedDate")) = format(FInfo.ModifyTime, "general date")
                    .TextMatrix(.Rows - 1, .ColIndex("CreatedDate")) = format(FInfo.CreationTime, "general date")
                End If
                
                .TextMatrix(.Rows - 1, .ColIndex("AttachedDate")) = format("" & rs("AttachedDate"), "general date")
                .Cell(flexcpPicture, .Rows - 1, .ColIndex("Filetitle")) = FileIcon(FInfo.hSmlIList, FInfo.hSmlIcon)
                .TextMatrix(.Rows - 1, .ColIndex("FileTitle")) = "" & rs("FileTitle")
                .TextMatrix(.Rows - 1, .ColIndex("Path")) = "" & rs("Path")
                .TextMatrix(.Rows - 1, .ColIndex("ObjectID")) = "" & rs("ObjectID")
                .TextMatrix(.Rows - 1, .ColIndex("DocumentClass")) = "" & rs("DocumentClass")
                .TextMatrix(.Rows - 1, .ColIndex("Embedded")) = "" & rs("Embedded")
                .IsSubtotal(.Rows - 1) = True
                .RowOutlineLevel(.Rows - 1) = 1
            End If
            rs.MoveNext
        Wend
        
        
'        Call .Outline(0)
'        If .Rows > 1 Then
'            .GetNode(1).Expanded = True
'            Call .Select(1, .ColIndex("FileTitle"))
'        End If
        
    End With
    
    Me.Caption = FormCaption
    
    Me.Show vbModal
    
End Sub





Private Sub Form_Load()
    Call IniGetForm(Me)
    Call IniGetGrid(Me, gData)
End Sub

Private Sub Form_Resize()
On Error Resume Next
    gData.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call IniPutForm(Me)
    Call IniPutGrid(Me, gData)
    
On Error Resume Next: Call Kill(PathAppend(TempPath, "*.*"))
End Sub

Private Sub gData_AfterMoveColumn(ByVal Col As Long, Position As Long)
    With gData
        .OutlineCol = .ColIndex("FileTitle")
    End With
End Sub

Private Sub gData_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With gData
        .ComboList = ""
        .AutoSearch = flexSearchFromCursor
        Select Case .ColKey(.Col)
            Case "DocumentClass":
                If .TextMatrix(.Row, .ColIndex("FileName")) = "" Then
                    Cancel = True
                Else
                    .EditMaxLength = 50
                    .AutoSearch = flexSearchNone
                    .ComboList = "| |" & .BuildComboList(HFApp.SqlExec("select distinct documentclass from attachments where isnull(documentclass,'') <>''"), "documentclass")
                End If
            Case Else: Cancel = True
        End Select
    End With
End Sub

Private Sub gData_DblClick()
    If gData.MouseRow > 1 Then
        Call mnuFilesSub_Click(mcFILE_OPEN)
    End If
End Sub

Private Sub gData_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case True
        Case KeyCode = vbKeyDelete:                   Call mnuFilesSub_Click(mcFILE_REMOVE)
        Case KeyCode = vbKeyReturn:                   Call mnuFilesSub_Click(mcFILE_OPEN)
        Case KeyCode = vbKeyP And Shift = vbCtrlMask: Call mnuFilesSub_Click(mcFILE_PRINT)
        Case KeyCode = vbKeyS And Shift = vbCtrlMask: Call mnuFilesSub_Click(mcFILE_SAVE)
        Case KeyCode = vbKeyEscape:                   Unload Me
    End Select
End Sub

Private Sub gData_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With gData
    If Button = vbRightButton Then
        If .MouseRow = 0 Then
            mnuColumnsSub(mcCOL_NAME).checked = Not .ColHidden(.ColIndex("FileTitle"))
            mnuColumnsSub(mcCOL_CLASS).checked = Not .ColHidden(.ColIndex("DocumentClass"))
            mnuColumnsSub(mcCOL_LOCATION).checked = Not .ColHidden(.ColIndex("Path"))
            mnuColumnsSub(mcCOL_SIZE).checked = Not .ColHidden(.ColIndex("FileSize"))
            mnuColumnsSub(mcCOL_MODIFIED).checked = Not .ColHidden(.ColIndex("ModifiedDate"))
            mnuColumnsSub(mcCOL_CREATED).checked = Not .ColHidden(.ColIndex("CreatedDate"))
            mnuColumnsSub(mcCOL_ATTACHED).checked = Not .ColHidden(.ColIndex("AttachedDate"))
            PopupMenu Me.mnuColumns
        Else
            If .TextMatrix(.Row, .ColIndex("FileName")) <> "" Then
                mnuFilesSub(mcFILE_PROPERTIES).Enabled = .TextMatrix(.Row, .ColIndex("Embedded")) = "False"
                PopupMenu Me.mnuFiles
            Else
                PopupMenu Me.mnuFolders
            End If
        End If
    End If
    End With
End Sub


Private Sub gData_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With gData
        Select Case .ColKey(.Col)
            Case "DocumentClass":
                Call HFApp.SqlExec("update attachments set documentclass=" & DbQuote(Str, .EditText, , True) & " where objectid=" & DbQuote(Str, .TextMatrix(.Row, .ColIndex("ObjectID"))) & " and filename=" & DbQuote(Str, .TextMatrix(.Row, .ColIndex("FileName"))), dbHomefront)
        End Select
    End With
End Sub

Private Sub mnuColumnsSub_Click(Index As Integer)
    With gData
    Select Case Index
        Case mcCOL_NAME:      .ColHidden(.ColIndex("FileTitle")) = Not .ColHidden(.ColIndex("FileTitle"))
        Case mcCOL_CLASS:     .ColHidden(.ColIndex("DocumentClass")) = Not .ColHidden(.ColIndex("DocumentClass"))
        Case mcCOL_LOCATION:  .ColHidden(.ColIndex("Path")) = Not .ColHidden(.ColIndex("Path"))
        Case mcCOL_SIZE:      .ColHidden(.ColIndex("FileSize")) = Not .ColHidden(.ColIndex("FileSize"))
        Case mcCOL_MODIFIED:  .ColHidden(.ColIndex("ModifiedDate")) = Not .ColHidden(.ColIndex("ModifiedDate"))
        Case mcCOL_CREATED:   .ColHidden(.ColIndex("CreatedDate")) = Not .ColHidden(.ColIndex("CreatedDate"))
        Case mcCOL_ATTACHED:  .ColHidden(.ColIndex("AttachedDate")) = Not .ColHidden(.ColIndex("AttachedDate"))
    End Select
    End With
End Sub

Private Sub mnuFilesSub_Click(Index As Integer)
    Dim ObjectID As String
    Dim FileName As String
    Dim Embedded As Boolean
    Dim s        As String
    Dim r        As Long
    
    With gData
    
        ObjectID = .TextMatrix(.Row, .ColIndex("ObjectID"))
        FileName = .TextMatrix(.Row, .ColIndex("FileName"))
        Embedded = .TextMatrix(.Row, .ColIndex("Embedded")) = "True"
        
        If Embedded Then
            s = PathAppend(TempPath, FileTitle(FileName))
            On Error Resume Next
            Call DBGetFile(s, , HFApp.Databases(dbHomefront), "Attachments WHERE objectid=" & DbQuote(Str, ObjectID) & " and filename=" & DbQuote(Str, FileName), "FileImage")
            On Error GoTo 0
            FileName = s
        Else
            FileName = .TextMatrix(.Row, .ColIndex("FileName"))
        End If
        
        Select Case Index
            Case mcFILE_OPEN
                Call ShellFile(Me.hwnd, FileName)
            
            Case mcFILE_PRINT
                Call ShellFile(Me.hwnd, FileName, , , True)
                
            Case mcFILE_SAVE
                If VBGetSaveFileName(s, , , , , , "Save As", , Me.hwnd) Then Call FileCopy(FileName, s)
            
            Case mcFILE_REMOVE
                Call HFApp.SqlExec("DELETE FROM Attachments WHERE objectid=" & DbQuote(Str, ObjectID) & " and filename=" & DbQuote(Str, .TextMatrix(.Row, .ColIndex("FileName"))), dbHomefront)
                Call .RemoveItem
                
            Case mcFILE_COPY:
                Call ClipboardSetFiles(FileName & Chr(0))
                
            Case mcFILE_PROPERTIES:
                Call ShowFileProperties(FileName, Me.hwnd)
                
            Case mcFILE_SENDTO:
                Screen.MousePointer = vbHourglass
                On Error Resume Next
                Call HFApp.SendMail(True, " ", "", "Emailing: " & FileTitle(FileName), "The message is ready to be sent with the following file or link attachments:" & vbCrLf & vbCrLf & FileTitle(FileName) & vbCrLf & vbCrLf & "Note: To protect against computer viruses, e-mail programs may prevent sending or receiving certain types of file attachments.  Check your e-mail security settings to determine how attachments are handled.", FileName)
                Screen.MousePointer = vbDefault
        End Select
    
    End With
End Sub


Private Function FileIcon(hList As Long, hIcon As Long) As IPictureDisp
    Set picCanvas.Picture = New StdPicture
    Call ImageList_Draw(hList, hIcon, picCanvas.hdc, 0, 0, ILD_TRANSPARENT)
    Set FileIcon = picCanvas.Image
End Function

Private Sub CreateFile(FileName As String)
    Dim i As Integer
    i = FreeFile
    Open FileName For Output As #i
    Close #i
End Sub

Private Sub mnuFoldersSub_Click(Index As Integer)
    Call FAttachments.AddFile(gData.RowData(gData.Row), Index = 0)
End Sub

Private Sub AddFileToGrid(ObjectID As String, Embedded As Boolean, FileName As String)
    'this only gets called when you right click insert/linkto file.
    Dim r As Long
    Dim FInfo As New ClsFileInfo
    With gData
        FInfo.FullPathName = FileName
        
        r = .Row + 1
        Call .AddItem("", r)
        .IsSubtotal(r) = True
        .RowOutlineLevel(r) = 1
        
        .TextMatrix(r, .ColIndex("FileName")) = FileName
        .TextMatrix(r, .ColIndex("FileSize")) = FInfo.FormatFileSize(FInfo.FileSize)
        .TextMatrix(r, .ColIndex("ModifiedDate")) = format(FInfo.ModifyTime, "general date")
        .TextMatrix(r, .ColIndex("CreatedDate")) = format(FInfo.CreationTime, "general date")
        If Embedded Then
            .Cell(flexcpPicture, r, .ColIndex("Path")) = picEmbedded.Picture
            .Cell(flexcpForeColor, r, .ColIndex("FileTitle")) = &HC00000 '&H00FF0000&  '&H800000
        End If
        
        .TextMatrix(r, .ColIndex("AttachedDate")) = format(Now(), "general date")
        .Cell(flexcpPicture, r, .ColIndex("Filetitle")) = FileIcon(FInfo.hSmlIList, FInfo.hSmlIcon)
        .TextMatrix(r, .ColIndex("FileTitle")) = FileTitle(FileName)
        .TextMatrix(r, .ColIndex("Path")) = FilePath(FileName)
        .TextMatrix(r, .ColIndex("ObjectID")) = ObjectID
        .TextMatrix(r, .ColIndex("DocumentClass")) = ""
        .TextMatrix(r, .ColIndex("Embedded")) = Embedded
        
    End With
End Sub



Private Sub gData_OLEDragDrop(Data As VSFlex8Ctl.VSDataObject, Effect As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
On Error GoTo eh
    Dim i As Long
    Dim ObjectID As String
    
    gData.Row = gData.MouseRow
    ObjectID = gData.RowData(gData.MouseRow)
    
    Effect = vbDropEffectNone
    For i = 0 To Data.FileCount - 1
        Call AddFile(ObjectID, Shift = 0, Data.Files(i))
    Next
eh: Exit Sub
End Sub

Private Sub gData_OLEDragOver(Data As VSFlex8Ctl.VSDataObject, Effect As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, State As Integer)
    Effect = IIf(gData.MouseRow > 0 And Data.GetFormat(vbCFFiles), vbDropEffectCopy, vbDropEffectNone)
End Sub

