VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FAttachments 
   Caption         =   "Attachments"
   ClientHeight    =   5370
   ClientLeft      =   11520
   ClientTop       =   3450
   ClientWidth     =   7275
   FillColor       =   &H00FF0000&
   Icon            =   "FAttachments.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5370
   ScaleWidth      =   7275
   Begin VB.PictureBox picEmbedded 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1380
      Picture         =   "FAttachments.frx":058A
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   3
      Top             =   4590
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picFolder 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1080
      Picture         =   "FAttachments.frx":0B14
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   2
      Top             =   4560
      Visible         =   0   'False
      Width           =   240
   End
   Begin VSFlex8Ctl.VSFlexGrid gData 
      Height          =   3555
      Left            =   150
      TabIndex        =   0
      Top             =   690
      Width           =   6555
      _cx             =   1977036074
      _cy             =   1977030783
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
      Cols            =   14
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
   Begin VB.PictureBox picCanvas 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   690
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   1
      Top             =   4590
      Visible         =   0   'False
      Width           =   240
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   4
      Top             =   0
      Width           =   7275
      _ExtentX        =   12832
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
      Begin MSComctlLib.ImageList LargeIcons 
         Left            =   6660
         Top             =   -30
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   41
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FAttachments.frx":1308
               Key             =   "EditAssembly"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FAttachments.frx":1BE2
               Key             =   "Find"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FAttachments.frx":24BC
               Key             =   "web"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FAttachments.frx":2D96
               Key             =   "ExcelImport"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FAttachments.frx":3670
               Key             =   "ExcelExport"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FAttachments.frx":3F4A
               Key             =   "Publish"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FAttachments.frx":4824
               Key             =   "Forecast"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FAttachments.frx":50FE
               Key             =   "ViewPOs"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FAttachments.frx":59D8
               Key             =   "ViewBudgets"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FAttachments.frx":62B2
               Key             =   "Open"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FAttachments.frx":6B8C
               Key             =   "Preview"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FAttachments.frx":7466
               Key             =   "Send"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FAttachments.frx":7D40
               Key             =   "TakeoffOneTime"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FAttachments.frx":861A
               Key             =   "Estimate"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FAttachments.frx":8EF4
               Key             =   "TakeoffAssembly"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FAttachments.frx":97CE
               Key             =   "NewAssembly"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FAttachments.frx":A0A8
               Key             =   "New"
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FAttachments.frx":A982
               Key             =   "TakeoffItem"
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FAttachments.frx":B25C
               Key             =   "TakeoffCustom"
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FAttachments.frx":BB36
               Key             =   "Save"
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FAttachments.frx":C410
               Key             =   "SaveAs"
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FAttachments.frx":CCEA
               Key             =   "Delete"
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FAttachments.frx":D5C4
               Key             =   "Print"
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FAttachments.frx":23F86
               Key             =   "RePrice"
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FAttachments.frx":24860
               Key             =   "PricebookSearch"
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FAttachments.frx":2513A
               Key             =   "Pricebook"
            EndProperty
            BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FAttachments.frx":25A14
               Key             =   "PricebookEdit"
            EndProperty
            BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FAttachments.frx":262EE
               Key             =   "PricebookExport"
            EndProperty
            BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FAttachments.frx":26BC8
               Key             =   "PricebookImport"
            EndProperty
            BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FAttachments.frx":274A2
               Key             =   "PricebookNew"
            EndProperty
            BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FAttachments.frx":27D7C
               Key             =   "View"
            EndProperty
            BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FAttachments.frx":28656
               Key             =   "Vendor1"
            EndProperty
            BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FAttachments.frx":28F30
               Key             =   "Vendor"
            EndProperty
            BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FAttachments.frx":2980A
               Key             =   "Add"
            EndProperty
            BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FAttachments.frx":2A0E4
               Key             =   "Attachments"
            EndProperty
            BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FAttachments.frx":2A3FE
               Key             =   ""
            EndProperty
            BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FAttachments.frx":2ACD8
               Key             =   ""
            EndProperty
            BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FAttachments.frx":2B5B2
               Key             =   "Design Center Options"
            EndProperty
            BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FAttachments.frx":2BE8C
               Key             =   "Global Options"
            EndProperty
            BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FAttachments.frx":2C766
               Key             =   "Models and Options"
            EndProperty
            BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FAttachments.frx":2D040
               Key             =   "Generate"
            EndProperty
         EndProperty
      End
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
         Caption         =   "Revision"
         Index           =   2
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
         Caption         =   "Web Portals"
         Index           =   5
      End
      Begin VB.Menu mnuColumnsSub 
         Caption         =   "Web Location"
         Index           =   6
      End
      Begin VB.Menu mnuColumnsSub 
         Caption         =   "Date Modified"
         Index           =   7
      End
      Begin VB.Menu mnuColumnsSub 
         Caption         =   "Date Created"
         Index           =   8
      End
      Begin VB.Menu mnuColumnsSub 
         Caption         =   "Date Attached"
         Index           =   9
      End
   End
   Begin VB.Menu mnuFolders 
      Caption         =   "<Folder Popup>"
      Visible         =   0   'False
      Begin VB.Menu mnuFoldersSub 
         Caption         =   "Link to File..."
         Index           =   0
      End
      Begin VB.Menu mnuFoldersSub 
         Caption         =   "Insert File..."
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
         Caption         =   "Show in web folder..."
         Index           =   5
      End
      Begin VB.Menu mnuFilesSub 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuFilesSub 
         Caption         =   "Cut"
         Enabled         =   0   'False
         Index           =   7
      End
      Begin VB.Menu mnuFilesSub 
         Caption         =   "Copy"
         Index           =   8
      End
      Begin VB.Menu mnuFilesSub 
         Caption         =   "Paste"
         Enabled         =   0   'False
         Index           =   9
      End
      Begin VB.Menu mnuFilesSub 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu mnuFilesSub 
         Caption         =   "Send to..."
         Index           =   11
      End
      Begin VB.Menu mnuFilesSub 
         Caption         =   "-"
         Index           =   12
      End
      Begin VB.Menu mnuFilesSub 
         Caption         =   "Properties..."
         Index           =   13
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
Private Const mcCOL_REVISION = 2
Private Const mcCOL_LOCATION = 3
Private Const mcCOL_SIZE = 4
Private Const mcCOL_WEBPORTALS = 5
Private Const mcCOL_WEBLOCATION = 6
Private Const mcCOL_MODIFIED = 7
Private Const mcCOL_CREATED = 8
Private Const mcCOL_ATTACHED = 9

Private Const mcFOLDER_LINK = 0
Private Const mcFOLDER_EMBED = 1

Private Const mcFILE_OPEN = 0
Private Const mcFILE_PRINT = 1
Private Const mcFILE_SAVE = 2
Private Const mcFILE_REMOVE = 3
Private Const mcFILE_SHOWINWEBFOLDER = 5
Private Const mcFILE_CUT = 7
Private Const mcFILE_COPY = 8
Private Const mcFILE_PASTE = 9
Private Const mcFILE_SENDTO = 11
Private Const mcFILE_PROPERTIES = 13



Public Function AddFile(ObjectID As String, Folder As String, Embedded As Boolean, Optional Filename As String)
On Error GoTo herror
    Dim s As String
    Dim FInfo As ClsFileInfo
    Dim i As Long
    
    
    If Filename = "" Then
        Call VBGetOpenFileName(Filename, , , , , True, , , , "Attach File")
    End If
    If Filename = "" Then Exit Function
    
    
        
    If PathIsLocalPath(Filename) And Not Embedded Then
        If MsgBox("This is a local file.  It may not be accessible to other users." & vbCrLf & "Are you sure this is what you want to do?", vbQuestion Or vbYesNo, App.ProductName) = vbNo Then Exit Function
    End If
    
    
    Set FInfo = New ClsFileInfo
    FInfo.FullPathName = Filename
    
    
    If Not FInfo.FileExists Then
        MsgBox "file not found", vbExclamation, App.ProductName
    Else
        
        s = ""
        s = s & "insert into attachments(objectid,filename,modifieddate,createddate,attacheddate,embedded,webvendors,weblocation)" & vbCrLf
        s = s & "values(" & DbQuote(Str, ObjectID) & vbCrLf
        s = s & "      ," & DbQuote(Str, Filename) & vbCrLf
        s = s & "      ," & DbQuote(DateTime, FInfo.ModifyTime) & vbCrLf
        s = s & "      ," & DbQuote(DateTime, FInfo.CreationTime) & vbCrLf
        s = s & "      ,getdate()" & vbCrLf
        s = s & "      ," & DbQuote(Bit, Embedded) & vbCrLf
        s = s & "      ,1,'Documents')"
        Call HFApp.SqlExec(s, dbHomefront)
        
        If Embedded Then
            Call DBPutFile(HFApp.Databases(dbHomefront), "Attachments WHERE objectid=" & DbQuote(Str, ObjectID) & " and filename=" & DbQuote(Str, Filename), "FileImage", Filename)
        End If
                
        If IsFormLoaded("FAttachments") Then
            Call UpdateWebPortal(AddFileToGrid(ObjectID, Folder, Embedded, Filename))
        End If
        
    End If
   
Exit Function
herror:

If Err.Number <> 0 Then
    MsgBox Err.Description, vbOKOnly, "Add File Attachment"
End If
End Function

Public Sub ShowForm(FormCaption As String, ParamArray ObjectIDsAndFolderNames())
On Error Resume Next
    Dim i      As Long
    Dim s      As String
    Dim f      As String
    Dim rs     As Recordset
    Dim Folder As String
    Dim FInfo As New ClsFileInfo
    
    
    'look for files in customer/job folder
    For i = 0 To UBound(ObjectIDsAndFolderNames) Step 2
        Call CreateJobLinks(Trim(ObjectIDsAndFolderNames(i)))
    Next
    
    
    
    s = ""
    For i = 0 To UBound(ObjectIDsAndFolderNames) Step 2
        If Trim(ObjectIDsAndFolderNames(i)) <> "" And Trim(ObjectIDsAndFolderNames(i + 1)) <> "" Then
            s = s & "union select " & DbQuote(Str, ObjectIDsAndFolderNames(i)) & " ObjectID," & DbQuote(Str, ObjectIDsAndFolderNames(i + 1)) & " Folder,a.filename,a.documentclass,a.filesize,a.filetitle,a.path,a.modifieddate,a.createddate,a.attacheddate,a.embedded,a.webcustomers,a.webvendors,a.weblocation,a.Revision from system_setup s left outer join attachments a on(a.objectid=" & DbQuote(Str, ObjectIDsAndFolderNames(i)) & ") where s.id = " & HFApp.DivisionID & vbCrLf
        End If
    Next
    s = Mid(s, 7) & "order by 2,6" 'folder,filetitle
    Set rs = HFApp.SqlExec(s, dbHomefront)
    With gData
        .Rows = 1
        While Not rs.EOF
            
            .AddItem ""
            .TextMatrix(.Rows - 1, .ColIndex("Folder")) = "" & rs("folder")
            If "" & rs("Embedded") = "True" Then
                s = PathAppend(AppWorkingFolder, "tmp." & FileExt("" & rs("FileName")))
                Call CreateFile(s)
                FInfo.FullPathName = s
                        
                .TextMatrix(.Rows - 1, .ColIndex("FileSize")) = FInfo.FormatFileSize(Val("" & rs("FileSize")))
                .TextMatrix(.Rows - 1, .ColIndex("ModifiedDate")) = Format("" & rs("ModifiedDate"), "general date")
                .TextMatrix(.Rows - 1, .ColIndex("CreatedDate")) = Format("" & rs("CreatedDate"), "general date")
                .Cell(flexcpPicture, .Rows - 1, .ColIndex("Path")) = picEmbedded.Picture
                .Cell(flexcpForeColor, .Rows - 1, .ColIndex("FileTitle")) = &HC00000 '&H00FF0000&  '&H800000
            Else
                On Error Resume Next
                f = "" & rs("FileName")
                FInfo.FullPathName = f
                .TextMatrix(.Rows - 1, .ColIndex("FileName")) = f
                If FileExists(f) Then
                    .TextMatrix(.Rows - 1, .ColIndex("FileSize")) = FInfo.FormatFileSize(FInfo.FileSize)
                    .TextMatrix(.Rows - 1, .ColIndex("ModifiedDate")) = Format(FInfo.ModifyTime, "general date")
                    .TextMatrix(.Rows - 1, .ColIndex("CreatedDate")) = Format(FInfo.CreationTime, "general date")
                End If
                On Error GoTo 0
            End If
            .TextMatrix(.Rows - 1, .ColIndex("AttachedDate")) = Format("" & rs("AttachedDate"), "general date")
            
            .TextMatrix(.Rows - 1, .ColIndex("FileName")) = "" & rs("FileName")
            .Cell(flexcpPicture, .Rows - 1, .ColIndex("Filetitle")) = FileIcon(FInfo.hSmlIList, FInfo.hSmlIcon)
            
            .TextMatrix(.Rows - 1, .ColIndex("FileTitle")) = "" & rs("FileTitle")
            .TextMatrix(.Rows - 1, .ColIndex("Path")) = "" & rs("Path")
            .TextMatrix(.Rows - 1, .ColIndex("ObjectID")) = "" & rs("ObjectID")
            .TextMatrix(.Rows - 1, .ColIndex("Revision")) = "" & rs("Revision")
            .TextMatrix(.Rows - 1, .ColIndex("DocumentClass")) = "" & rs("DocumentClass")
            .TextMatrix(.Rows - 1, .ColIndex("Embedded")) = "" & rs("Embedded")
            
            Select Case True
                Case "" & rs("WebCustomers") = "True" And "" & rs("WebVendors") = "True":    .TextMatrix(.Rows - 1, .ColIndex("WebPortals")) = "All"
                Case "" & rs("WebCustomers") = "True":                                       .TextMatrix(.Rows - 1, .ColIndex("WebPortals")) = "Customer"
                Case "" & rs("WebVendors") = "True":                                         .TextMatrix(.Rows - 1, .ColIndex("WebPortals")) = "Vendor"
                Case Else:                                                                   .TextMatrix(.Rows - 1, .ColIndex("WebPortals")) = "None"
            End Select
            .TextMatrix(.Rows - 1, .ColIndex("WebLocation")) = "" & rs("WebLocation")
            If .TextMatrix(.Rows - 1, .ColIndex("WebLocation")) = "" Then .TextMatrix(.Rows - 1, .ColIndex("WebLocation")) = "Documents"
            
            
            'hide these if not a job or customer file
            If Not IsIn(left(.TextMatrix(.Rows - 1, .ColIndex("ObjectID")), 2), "C~", "J~") Then
                .TextMatrix(.Rows - 1, .ColIndex("WebLocation")) = ""
                .TextMatrix(.Rows - 1, .ColIndex("WebPortals")) = ""
            End If
            
            rs.MoveNext
        Wend
        
    End With
    
    Call ReGroup
    
    Me.Caption = Replace(Replace(FormCaption, Chr(10), ""), Chr(13), "")
    
    
    Me.Show vbModal
    
End Sub

Private Sub ReGroup()
    Dim r As Long
    With gData
    
    
        'clear old groups
        Call .Subtotal(flexSTClear)
        'move these to left
        .ColPosition(.ColIndex("ObjectID")) = 0
        .ColPosition(.ColIndex("Folder")) = 1
        .ColPosition(.ColIndex("FileTitle")) = 2
        'apply formatting
        .OutlineCol = 1
        .OutlineBar = flexOutlineBarSimpleLeaf
        .ExplorerBar = flexExSortShowAndMove
        .MergeCells = flexMergeOutline
        .TextMatrix(0, .ColIndex("Folder")) = .TextMatrix(0, .ColIndex("FileTitle"))
        .MergeCellsFixed = flexMergeFree
        .MergeRow(0) = True
        .ColWidth(1) = 240
        'sort first or subtotal method will make duplicates
        .Col = 0
        .ColSel = 1
        .Sort = flexSortStringAscending
        'make groups
        Call .Subtotal(flexSTNone, 1, , , , , True, , 0)
    
    
        'hide empty file rows and put icons on folder rows
        For r = 1 To .Rows - 1
            If .IsSubtotal(r) Then
                .Cell(flexcpPicture, r, .ColIndex("Folder")) = picFolder.Picture
            ElseIf .TextMatrix(r, .ColIndex("filename")) = "" Then
                .RowHidden(r) = True
            End If
        Next
    
    
    End With
End Sub



Private Sub Form_Load()
On Error GoTo eh
    Dim i As Long
    Dim s As String
    
    Call SetToolbarIcons(Toolbar, LargeIcons)
    Call IniGetForm(Me)
    
    mnuFoldersSub(mcFOLDER_EMBED).Enabled = HFApp.SQLEdition Like "*Express Edition*"
    
    
    
    Call IniGetGrid(Me, gData)

Exit Sub
eh: Call errHandler(SRCFILE & "Form_Load")
End Sub

Private Sub Form_Resize()
On Error Resume Next
    gData.Move 0, Toolbar.Height, Me.ScaleWidth, Me.ScaleHeight - Toolbar.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call IniPutForm(Me)
    Call IniPutGrid(Me, gData)
    
On Error Resume Next: Call Kill(PathAppend(TempPath, "*.*"))
End Sub


Private Sub gData_AfterCollapse(ByVal Row As Long, ByVal State As Integer)
    Dim r As Long
    With gData
        For r = 1 To .Rows - 1
            If .IsSubtotal(r) Then
               
            ElseIf .TextMatrix(r, .ColIndex("filename")) = "" Then
                .RowHidden(r) = True
            End If
        Next
    End With
End Sub

Private Sub gData_BeforeMoveColumn(ByVal Col As Long, Position As Long)
    If Col < 2 Then Position = Col
    If Col > 2 And Position < 3 Then Position = Col
End Sub

Private Sub gData_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    With gData
        If Row <> 0 Then
        If Col = .ColIndex("Folder") And Not .IsSubtotal(Row) Then Done = True
        End If
    End With
End Sub

Private Sub mnuColumnsSub_Click(Index As Integer)
    With gData
    Select Case Index
        Case mcCOL_NAME:        .ColHidden(.ColIndex("FileTitle")) = Not .ColHidden(.ColIndex("FileTitle"))
        Case mcCOL_CLASS:       .ColHidden(.ColIndex("DocumentClass")) = Not .ColHidden(.ColIndex("DocumentClass"))
        Case mcCOL_REVISION:    .ColHidden(.ColIndex("Revision")) = Not .ColHidden(.ColIndex("Revision"))
        Case mcCOL_LOCATION:    .ColHidden(.ColIndex("Path")) = Not .ColHidden(.ColIndex("Path"))
        Case mcCOL_SIZE:        .ColHidden(.ColIndex("FileSize")) = Not .ColHidden(.ColIndex("FileSize"))
        Case mcCOL_WEBPORTALS:  .ColHidden(.ColIndex("WebPortals")) = Not .ColHidden(.ColIndex("WebPortals"))
        Case mcCOL_WEBLOCATION: .ColHidden(.ColIndex("WebLocation")) = Not .ColHidden(.ColIndex("WebLocation"))
        Case mcCOL_MODIFIED:    .ColHidden(.ColIndex("ModifiedDate")) = Not .ColHidden(.ColIndex("ModifiedDate"))
        Case mcCOL_CREATED:     .ColHidden(.ColIndex("CreatedDate")) = Not .ColHidden(.ColIndex("CreatedDate"))
        Case mcCOL_ATTACHED:    .ColHidden(.ColIndex("AttachedDate")) = Not .ColHidden(.ColIndex("AttachedDate"))
        
        
    End Select
    End With
End Sub

Private Sub mnuFiles_Click()
    With gData
        mnuFilesSub(mcFILE_SHOWINWEBFOLDER).Enabled = IsIn(left(.TextMatrix(.Row, .ColIndex("ObjectID")), 2), "C~", "J~")
    End With
End Sub

Private Sub mnuFilesSub_Click(Index As Integer)
    Dim ObjectID As String
    Dim Filename As String
    Dim WebPath  As String
    Dim WebFileName As String
    Dim Embedded As Boolean
    Dim s        As String
    Dim r        As Long
    
    With gData
    
        ObjectID = .TextMatrix(.Row, .ColIndex("ObjectID"))
        Filename = .TextMatrix(.Row, .ColIndex("FileName"))
        Embedded = .TextMatrix(.Row, .ColIndex("Embedded")) = "True"
        Call GetDetails(.Row, "", "", "", "", WebPath, WebFileName)
               
        If Filename = "" Then Exit Sub
        If Not FileExists(Filename) Then Filename = WebFileName
        
        'if needed extract db file
        If Embedded And Index <> mcFILE_REMOVE Then
            s = PathAppend(TempPath, FileTitle(Filename))
            On Error Resume Next
            Call DBGetFile(s, , HFApp.Databases(dbHomefront), "Attachments WHERE objectid=" & DbQuote(Str, ObjectID) & " and filename=" & DbQuote(Str, Filename), "FileImage")
            On Error GoTo 0
            Filename = s
        End If
        
        'do requested action
        Select Case Index
            Case mcFILE_SHOWINWEBFOLDER
                If WebFileName <> "" Then
                    Call ShellFile(Me.hwnd, WebPath)
                End If
            
            Case mcFILE_OPEN
                If FileExists(Filename) Then Call ShellFile(Me.hwnd, Filename)
                
            Case mcFILE_PRINT
                Call ShellFile(Me.hwnd, Filename, , , True)
                
            Case mcFILE_SAVE
                If VBGetSaveFileName(s, , , , , , "Save As", , Me.hwnd) Then Call FileCopy(Filename, s)
            
            Case mcFILE_REMOVE
                .TextMatrix(.Row, .ColIndex("WebPortals")) = "None"
                Call UpdateWebPortal(.Row)
                Call HFApp.SqlExec("DELETE FROM Attachments WHERE objectid=" & DbQuote(Str, ObjectID) & " and filename=" & DbQuote(Str, .TextMatrix(.Row, .ColIndex("FileName"))), dbHomefront)
                Call .RemoveItem
                
            
            Case mcFILE_COPY:
                Call ClipboardSetFiles(Filename & Chr(0))
                
            Case mcFILE_PROPERTIES:
                Call ShowFileProperties(Filename, Me.hwnd)
                
            Case mcFILE_SENDTO:
                Screen.MousePointer = vbHourglass
                On Error Resume Next
                Call HFApp.SendMail(True, " ", "", "Emailing: " & FileTitle(Filename), "The message is ready to be sent with the following file or link attachments:" & vbCrLf & vbCrLf & FileTitle(Filename) & vbCrLf & vbCrLf & "Note: To protect against computer viruses, e-mail programs may prevent sending or receiving certain types of file attachments.  Check your e-mail security settings to determine how attachments are handled.", Filename)
                Screen.MousePointer = vbDefault
        
        End Select
    
    End With
End Sub


Private Function FileIcon(hList As Long, hIcon As Long) As IPictureDisp
    Set picCanvas.Picture = New StdPicture
    Call ImageList_Draw(hList, hIcon, picCanvas.hDC, 0, 0, ILD_TRANSPARENT)
    Set FileIcon = picCanvas.Image
End Function

Private Sub CreateFile(Filename As String)
    Dim i As Integer
    i = FreeFile
    Open Filename For Output As #i
    Close #i
End Sub

Private Sub mnuFoldersSub_Click(Index As Integer)
On Error GoTo eh

    If gData.Row < 0 Then Exit Sub
        
        
    Select Case Index
        Case mcFOLDER_LINK
            Call FAttachments.AddFile(gData.TextMatrix(gData.Row, gData.ColIndex("ObjectID")), gData.TextMatrix(gData.Row, gData.ColIndex("Folder")), False)
            
        Case mcFOLDER_EMBED
            Call FAttachments.AddFile(gData.TextMatrix(gData.Row, gData.ColIndex("ObjectID")), gData.TextMatrix(gData.Row, gData.ColIndex("Folder")), True)
            
                        
    End Select
    
    
Exit Sub
eh: If InStr(1, Err.Description, "duplicate", vbTextCompare) Then
        MsgBox "This file is already attached.", vbExclamation, App.ProductName
    Else
        Call errHandler(SRCFILE & "mnuFolderSub_Click")
    End If
End Sub

Private Function AddFileToGrid(ObjectID As String, Folder As String, Embedded As Boolean, Filename As String) As Long
    'this only gets called when you right click insert/linkto file.
    Dim r As Long
    Dim FInfo As New ClsFileInfo
    With gData
        FInfo.FullPathName = Filename
        Call .AddItem("")
        
        .TextMatrix(.Rows - 1, .ColIndex("Folder")) = Folder
        .TextMatrix(.Rows - 1, .ColIndex("ObjectID")) = ObjectID
        
        .TextMatrix(.Rows - 1, .ColIndex("FileName")) = Filename
        .TextMatrix(.Rows - 1, .ColIndex("FileSize")) = FInfo.FormatFileSize(FInfo.FileSize)
        .TextMatrix(.Rows - 1, .ColIndex("ModifiedDate")) = Format(FInfo.ModifyTime, "general date")
        .TextMatrix(.Rows - 1, .ColIndex("CreatedDate")) = Format(FInfo.CreationTime, "general date")
        If Embedded Then
            .Cell(flexcpPicture, .Rows - 1, .ColIndex("Path")) = picEmbedded.Picture
            .Cell(flexcpForeColor, .Rows - 1, .ColIndex("FileTitle")) = &HC00000 '&H00FF0000&  '&H800000
        End If
        .TextMatrix(.Rows - 1, .ColIndex("AttachedDate")) = Format(Now(), "general date")
        .Cell(flexcpPicture, .Rows - 1, .ColIndex("Filetitle")) = FileIcon(FInfo.hSmlIList, FInfo.hSmlIcon)
        .TextMatrix(.Rows - 1, .ColIndex("FileTitle")) = FileTitle(Filename)
        .TextMatrix(.Rows - 1, .ColIndex("Path")) = FilePath(Filename)
        .TextMatrix(.Rows - 1, .ColIndex("DocumentClass")) = ""
        .TextMatrix(.Rows - 1, .ColIndex("Revision")) = ""
        .TextMatrix(.Rows - 1, .ColIndex("Embedded")) = Embedded
        
        If IsIn(left(ObjectID, 2), "C~", "J~") Then
            .TextMatrix(.Rows - 1, .ColIndex("WebPortals")) = "Vendor"
            .TextMatrix(.Rows - 1, .ColIndex("WebLocation")) = "Documents"
        End If
        AddFileToGrid = .Rows - 1
    End With
    Call ReGroup
End Function



Private Sub gData_OLEDragDrop(Data As VSFlex8Ctl.VSDataObject, Effect As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
On Error GoTo eh
    Dim i As Long
    Dim ObjectID As String
    Dim Folder As String
    
    gData.Row = gData.MouseRow
    ObjectID = gData.TextMatrix(gData.Row, gData.ColIndex("ObjectID"))
    Folder = gData.TextMatrix(gData.Row, gData.ColIndex("Folder"))
    
    
    Effect = vbDropEffectNone
    For i = 0 To Data.FileCount - 1
        Call AddFile(ObjectID, Folder, False, Data.files(i))
    Next
eh: Exit Sub
End Sub

Private Sub gData_OLEDragOver(Data As VSFlex8Ctl.VSDataObject, Effect As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, State As Integer)
    Effect = IIf(gData.MouseRow > 0 And Data.GetFormat(vbCFFiles), vbDropEffectCopy, vbDropEffectNone)
End Sub


Private Sub CreateJobLinks(ObjectID As String)
On Error GoTo eh
    Dim s As String
    Dim file As String
    Dim Folder As String
    
    
    
    Dim fso As New Scripting.FileSystemObject
    Dim fsFolder As Scripting.Folder
    Dim fsSubfolder As Scripting.Folder
    Dim fsFile As Scripting.file

    
    Select Case left(ObjectID, 2)
        Case "J~":   file = Replace(ObjectID, "J~", "Job_No_")
        Case "C~":   file = Replace(ObjectID, "C~", "Customer_No_")
        Case Else:   Exit Sub
    End Select
    
    Folder = HFApp.Options(Document_Folder)
    If Folder = "" Then Folder = GetFolderPath(CSIDL_COMMON_APPDATA)
    
    
    'loop thru file tree inserting each one
    Set fsFolder = fso.GetFolder(PathAppend(Folder, file))
    For Each fsFile In fsFolder.files
        s = "Insert into Attachments (ObjectID,FileName,Embedded) Values(" & DbQuote(Str, ObjectID) & "," & DbQuote(Str, fsFile.Path) & ",0)"
        On Error Resume Next
        HFApp.SqlExec s
        On Error GoTo eh
    Next
    For Each fsSubfolder In fsFolder.SubFolders
        For Each fsFile In fsSubfolder.files
            s = "Insert into Attachments (ObjectID,FileName,Embedded)" & vbCrLf
            s = s & " Values(" & DbQuote(Str, ObjectID) & "," & DbQuote(Str, fsFile.Path) & ",0)"
            On Error Resume Next
            HFApp.SqlExec s
            On Error GoTo eh
        Next
    Next
  
Exit Sub
eh:
If Err.Number = 76 Then
    Exit Sub
Else
    Call errHandler("CreateJobLinks", s)
End If
End Sub

Private Sub GetDetails(Row As Long, Job As String, Customer As String, LocalFile As String, WebRoot As String, WebPath As String, WebFile As String)
    With gData
        'parse out all the variables
        WebRoot = HFApp.Options.ValueByName("WebAttachmentsFolder")
        LocalFile = .TextMatrix(Row, .ColIndex("FileName"))
        Customer = Parse(.TextMatrix(Row, .ColIndex("ObjectID")), 2, "C~")
        If Customer = "" Then
            Job = Parse(.TextMatrix(Row, .ColIndex("ObjectID")), 2, "J~")
            If Job <> "" Then
                WebPath = PathAppend(WebRoot, Job)
                WebFile = PathAppend(WebRoot, Job, .TextMatrix(Row, .ColIndex("Filetitle")))
            End If
        Else
            Job = "" & HFApp.SqlExec("select job_no from tblcustomers where customer_no=" & DbQuote(Str, Customer), dbHomefront)(0)
            If Job <> "" Then
                WebPath = PathAppend(WebRoot, Job, Customer)
                WebFile = PathAppend(WebRoot, Job, Customer, .TextMatrix(Row, .ColIndex("Filetitle")))
            End If
        End If
    End With
End Sub

Private Sub UpdateWebPortal(Row As Long)
On Error GoTo eh
    
    Dim s As String
    Dim LocalFile As String
    Dim Customer As String
    Dim Job As String
    Dim WebRoot As String
    Dim WebPath As String
    Dim WebFile As String
    Dim Embedded As Boolean
    Dim ObjectID As String
    
    With gData
        
        ObjectID = .TextMatrix(Row, .ColIndex("ObjectID"))
        Embedded = .TextMatrix(.Row, .ColIndex("Embedded")) = "True"

        Call GetDetails(Row, Job, Customer, LocalFile, WebRoot, WebPath, WebFile)
    
        If WebPath = "" Then Exit Sub
        
        LocalFile = .TextMatrix(Row, .ColIndex("FileName"))
        
        If .TextMatrix(Row, .ColIndex("WebPortals")) <> "None" Then
            
            If Embedded Then
                s = PathAppend(TempPath, FileTitle(LocalFile))
                On Error Resume Next
                Call DBGetFile(s, , HFApp.Databases(dbHomefront), "Attachments WHERE objectid=" & DbQuote(Str, ObjectID) & " and filename=" & DbQuote(Str, LocalFile), "FileImage")
                On Error GoTo 0
                LocalFile = s
            End If
    
            On Error Resume Next
            Call CreatePath("", FilePath(WebFile))
            If LocalFile <> WebFile Then Call FileCopy(LocalFile, WebFile)
            On Error GoTo eh
        
        End If
    
        'copy to, or delete file from web folder
        If .TextMatrix(Row, .ColIndex("WebPortals")) = "None" Then
            
            s = ""
            s = s & "delete from tblcustomerfiles" & vbCrLf
            s = s & "where isnull(customer_no,'')=" & DbQuote(Str, Customer) & vbCrLf
            s = s & "  and job_no=" & DbQuote(Str, Job) & vbCrLf
            s = s & "  and linkurl=" & DbQuote(Str, .TextMatrix(Row, .ColIndex("Filetitle"))) & vbCrLf
            Call HFApp.SqlExec(s, dbHomefront)
        
        Else
        
            
            s = "exec SP_SaveCustomerFile " & _
                DbQuote(Str, Job) & "," & _
                "'New File'," & _
                DbQuote(Str, .TextMatrix(Row, .ColIndex("WebLocation"))) & "," & _
                DbQuote(Str, Filename(.TextMatrix(Row, .ColIndex("FileTitle")))) & "," & _
                DbQuote(Str, .TextMatrix(Row, .ColIndex("FileTitle"))) & "," & _
                DbQuote(Str, .TextMatrix(Row, .ColIndex("DocumentClass"))) & "," & _
                "" & _
                DbQuote(Str, Customer) & "," & _
                "'N/A'," & _
                DbQuote(Str, .TextMatrix(Row, .ColIndex("Revision"))) & "," & _
                DbQuote(Bit, IsIn(.TextMatrix(Row, .ColIndex("WebPortals")), "Customer", "All")) & "," & _
                DbQuote(Bit, IsIn(.TextMatrix(Row, .ColIndex("WebPortals")), "Vendor", "All")) & "," & _
                DbQuote(DateTime, .TextMatrix(Row, .ColIndex("ModifiedDate")))
                
            Call HFApp.SqlExec(s, dbHomefront)
        End If
        
        
        
    End With

Exit Sub
eh: Call errHandler(SRCFILE & "UpdateWebPortal")
End Sub














Private Sub gData_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With gData
        Select Case True
            Case .ColKey(.Col) = "DocumentClass":
                'If HFApp.UserPermission("LimitAttachmentClasses") = False Then
                    Call HFApp.SqlExec("update attachments set documentclass=" & DbQuote(Str, .EditText, , True) & " where objectid=" & DbQuote(Str, .TextMatrix(.Row, .ColIndex("ObjectID"))) & " and filename=" & DbQuote(Str, .TextMatrix(.Row, .ColIndex("FileName"))), dbHomefront)
                'End If
                
             Case .ColKey(.Col) = "Revision":
                Call HFApp.SqlExec("update attachments set revision=" & DbQuote(Str, .EditText, , True) & " where objectid=" & DbQuote(Str, .TextMatrix(.Row, .ColIndex("ObjectID"))) & " and filename=" & DbQuote(Str, .TextMatrix(.Row, .ColIndex("FileName"))), dbHomefront)
                
            Case .ColKey(.Col) = "WebLocation"
                Call HFApp.SqlExec("update attachments set WebLocation=" & DbQuote(Str, .EditText, , True) & " where objectid=" & DbQuote(Str, .TextMatrix(.Row, .ColIndex("ObjectID"))) & " and filename=" & DbQuote(Str, .TextMatrix(.Row, .ColIndex("FileName"))), dbHomefront)
                
            Case .ColKey(.Col) = "WebPortals" And .EditText = "All"
                Call HFApp.SqlExec("update attachments set WebCustomers=1,WebVendors=1 where objectid=" & DbQuote(Str, .TextMatrix(.Row, .ColIndex("ObjectID"))) & " and filename=" & DbQuote(Str, .TextMatrix(.Row, .ColIndex("FileName"))), dbHomefront)
                
            Case .ColKey(.Col) = "WebPortals" And .EditText = "Customer"
                Call HFApp.SqlExec("update attachments set WebCustomers=1,WebVendors=0 where objectid=" & DbQuote(Str, .TextMatrix(.Row, .ColIndex("ObjectID"))) & " and filename=" & DbQuote(Str, .TextMatrix(.Row, .ColIndex("FileName"))), dbHomefront)
            
            Case .ColKey(.Col) = "WebPortals" And .EditText = "Vendor"
                Call HFApp.SqlExec("update attachments set WebCustomers=0,WebVendors=1 where objectid=" & DbQuote(Str, .TextMatrix(.Row, .ColIndex("ObjectID"))) & " and filename=" & DbQuote(Str, .TextMatrix(.Row, .ColIndex("FileName"))), dbHomefront)
            
            Case .ColKey(.Col) = "WebPortals"
                Call HFApp.SqlExec("update attachments set WebCustomers=0,WebVendors=0 where objectid=" & DbQuote(Str, .TextMatrix(.Row, .ColIndex("ObjectID"))) & " and filename=" & DbQuote(Str, .TextMatrix(.Row, .ColIndex("FileName"))), dbHomefront)
                
        End Select
        Call UpdateWebPortal(Row)
    End With
End Sub

Private Sub gData_AfterMoveColumn(ByVal Col As Long, Position As Long)
    With gData
        .OutlineCol = .ColIndex("Folder")
    End With
End Sub

Private Sub gData_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With gData
        .ComboList = ""
        .AutoSearch = flexSearchFromCursor
        If .TextMatrix(.Row, .ColIndex("FileName")) = "" Then
            Cancel = True
        Else
            Select Case .ColKey(.Col)
            
                Case "WebPortals", "WebLocation"
                    .AutoSearch = flexSearchNone
                    Cancel = Not IsIn(left(.TextMatrix(Row, .ColIndex("ObjectID")), 2), "C~", "J~")

                Case "Revision":
                    .EditMaxLength = 15
                    .AutoSearch = flexSearchNone

                Case "DocumentClass":
                    .EditMaxLength = 50
                    .AutoSearch = flexSearchNone
                    .ComboList = " |" & .BuildComboList(HFApp.SqlExec("select distinct documentclass from attachments where isnull(documentclass,'') <>'' order by documentclass"), "documentclass")
                    If HFApp.UserPermission("LimitAttachmentClasses") = False Then
                        .ComboList = "|" & .ComboList
                    End If
                    
                Case Else
                    Cancel = True
            End Select
        End If
    End With
End Sub


Private Sub gData_DblClick()
    If gData.MouseRow > 1 Then
        Call mnuFilesSub_Click(mcFILE_OPEN)
    End If
End Sub


Private Sub gData_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    With gData
    If Button = vbRightButton Then
        If .MouseRow = 0 Or .Row < 1 Then
            mnuColumnsSub(mcCOL_NAME).Checked = Not .ColHidden(.ColIndex("FileTitle"))
            mnuColumnsSub(mcCOL_CLASS).Checked = Not .ColHidden(.ColIndex("DocumentClass"))
            mnuColumnsSub(mcCOL_REVISION).Checked = Not .ColHidden(.ColIndex("Revision"))
            mnuColumnsSub(mcCOL_LOCATION).Checked = Not .ColHidden(.ColIndex("Path"))
            mnuColumnsSub(mcCOL_SIZE).Checked = Not .ColHidden(.ColIndex("FileSize"))
            mnuColumnsSub(mcCOL_MODIFIED).Checked = Not .ColHidden(.ColIndex("ModifiedDate"))
            mnuColumnsSub(mcCOL_CREATED).Checked = Not .ColHidden(.ColIndex("CreatedDate"))
            mnuColumnsSub(mcCOL_ATTACHED).Checked = Not .ColHidden(.ColIndex("AttachedDate"))
            
            mnuColumnsSub(mcCOL_WEBPORTALS).Checked = Not .ColHidden(.ColIndex("WebPortals"))
            mnuColumnsSub(mcCOL_WEBLOCATION).Checked = Not .ColHidden(.ColIndex("WebLocation"))
            
            
            PopupMenu Me.mnuColumns
        Else
            .Row = .MouseRow
            If .IsSubtotal(.Row) Then
                PopupMenu Me.mnuFolders
            Else
                mnuFilesSub(mcFILE_PROPERTIES).Enabled = .TextMatrix(.Row, .ColIndex("Embedded")) = "False"
                PopupMenu Me.mnuFiles
            End If
        End If
    End If
    End With
End Sub





Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case True
        Case KeyCode = vbKeyF And Shift = vbCtrlMask:  Call Toolbar_ButtonClick(Toolbar.Buttons("Find"))
        Case KeyCode = vbKeyReturn:                    Call mnuFilesSub_Click(mcFILE_OPEN)
        Case KeyCode = vbKeyP And Shift = vbCtrlMask:  Call mnuFilesSub_Click(mcFILE_PRINT)
        Case KeyCode = vbKeyS And Shift = vbCtrlMask:  Call mnuFilesSub_Click(mcFILE_SAVE)
        Case KeyCode = vbKeyDelete:                    Call mnuFilesSub_Click(mcFILE_REMOVE)
        Case KeyCode = vbKeyEscape:                    Unload Me
        Case Else
    End Select
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
        Case "OPEN":      Call mnuFilesSub_Click(mcFILE_OPEN)
        Case "PRINT":     Call mnuFilesSub_Click(mcFILE_PRINT)
        Case "SAVEAS":    Call mnuFilesSub_Click(mcFILE_SAVE)
        Case "DELETE":    Call mnuFilesSub_Click(mcFILE_REMOVE)
    End Select
End Sub

