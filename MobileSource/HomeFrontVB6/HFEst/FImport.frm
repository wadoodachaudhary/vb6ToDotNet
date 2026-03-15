VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{55473EAC-7715-4257-B5EF-6E14EBD6A5DD}#1.0#0"; "vbalProgBar6.ocx"
Begin VB.Form FDataImport 
   Caption         =   "Data Import Wizard"
   ClientHeight    =   6165
   ClientLeft      =   7185
   ClientTop       =   465
   ClientWidth     =   7305
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FImport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   7305
   Begin VB.PictureBox WizFoot 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   585
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   7305
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   5580
      Width           =   7305
      Begin VB.CommandButton cmdNav 
         Caption         =   "&Cancel"
         Height          =   375
         Index           =   0
         Left            =   1860
         TabIndex        =   1
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "< &Back"
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   3060
         TabIndex        =   2
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "&Next >"
         Height          =   375
         Index           =   2
         Left            =   4200
         TabIndex        =   3
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "&Finish"
         Enabled         =   0   'False
         Height          =   375
         Index           =   3
         Left            =   5400
         TabIndex        =   4
         Top             =   120
         Width           =   1095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   -60
         X2              =   26420
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   1
         X1              =   0
         X2              =   26480
         Y1              =   15
         Y2              =   15
      End
   End
   Begin HFEst.WizHead WizHead1 
      Align           =   1  'Align Top
      Height          =   900
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   7305
      _ExtentX        =   12885
      _ExtentY        =   1588
      Caption         =   "Import Vendors"
      Description     =   "The vendor import wizard will help you update your database"
      Icon            =   "FImport.frx":000C
   End
   Begin VB.Frame WizFrame 
      BorderStyle     =   0  'None
      Height          =   3495
      Index           =   0
      Left            =   0
      TabIndex        =   6
      Top             =   900
      Width           =   7515
      Begin vbalProgBarLib6.vbalProgressBar ProgressBar 
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   3060
         Visible         =   0   'False
         Width           =   7155
         _ExtentX        =   12621
         _ExtentY        =   556
         Picture         =   "FImport.frx":08E6
         ForeColor       =   0
         BarPicture      =   "FImport.frx":0902
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         XpStyle         =   -1  'True
      End
      Begin VSFlex8Ctl.VSFlexGrid gData 
         Height          =   2655
         Left            =   120
         TabIndex        =   0
         Top             =   720
         Width           =   7155
         _cx             =   12621
         _cy             =   4683
         Appearance      =   2
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
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FImport.frx":091E
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
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
      Begin VB.Label lblFileDate 
         AutoSize        =   -1  'True
         Caption         =   "Last Modified April 18, 2007"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   1965
      End
      Begin VB.Label lblFilename 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   45
      End
   End
End
Attribute VB_Name = "FDataImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Const srcFile = "FImport::"
Private mFilename  As String

Private mTitle           As String
Private mTableName       As String
Private mKeyColumns      As String
Private mRequiredColumns As String
Private mColumnMappings  As String

Private mCancel As Boolean

Public Function ImportData(TableName As String, KeyColumns As String, Title As String, Optional RequiredColumns As String, Optional ColumnMappings As String) As Boolean
On Error GoTo eh

    'ie: Call FDataImport.ImportData("Vendor", _
    '                                "tblVendors", _
    '                                "Vendor_ID,Vendor_Name", _
    '                                "VendorID=Vendor_ID,Vendor=Vendor_ID")


    Dim i As Long
    
    
    mTableName = TableName
    mKeyColumns = KeyColumns
    mTitle = Title
    mRequiredColumns = RequiredColumns
    mColumnMappings = ColumnMappings
    
    If mKeyColumns = "" Then Call Err.Raise(5, , "You must specify at least one column name in the ""KeyColumns"" parameter")
    
    i = IniGet(AppIni, "Options", mTableName & "ImportFileExt", 0)
    If VBGetOpenFileName(mFilename, , , , , True, "Excel Files (*.xls)|*.xls|Text Files (*.txt;*.csv)|*.txt;*.csv|All Files (*.*)|*.*", i, , , , FMain.hwnd) Then
        Call IniPut(AppIni, "Options", mTableName & "ImportFileExt", i)
        
        CurrentFrame = 0
        
        mCancel = True
        Me.Show vbModal
        ImportData = Not mCancel

    Else
        mCancel = True
        Unload Me
    End If
        
Exit Function
eh: Call errHandler(srcFile & "ShowForm")
End Function


Private Function TextMatrix(r As Long, ColKey As String) As String
On Error Resume Next
    TextMatrix = gData.TextMatrix(r, gData.ColIndex(ColKey))
End Function


Private Property Get CurrentFrame() As Long
    Dim i As Long
    For i = 0 To WizFrame.UBound
        If WizFrame(i).Visible = True Then
            CurrentFrame = i
            Exit Property
        End If
    Next
    CurrentFrame = WizFrame.LBound
End Property

Private Property Let CurrentFrame(RHS As Long)
On Error GoTo eh
    Dim i As Long
    Dim fileinfo As New ClsFileInfo
    Dim s As String
    Dim r As Long
    
    
    Screen.MousePointer = vbHourglass
    For i = 0 To WizFrame.UBound
        WizFrame(i).Visible = i = RHS
    Next
    
    cmdNav(1).Enabled = RHS > 0
    cmdNav(2).Enabled = RHS < WizFrame.UBound
    cmdNav(3).Enabled = RHS = WizFrame.UBound
        
        
    With gData
    Select Case RHS
    
        Case 0 ' show file
            If WizFrame(RHS).tag = "" Then
                WizFrame(RHS).tag = "LOADED"
                lblFilename.Caption = "File name: " & mFilename
                fileinfo.FullPathName = mFilename
                lblFileDate.Caption = "Last Modified: " & format(fileinfo.ModifyTime, "long date")
                
                Select Case FileExt(mFilename)
                    Case "xls"
                        mFilename = SaveToCSV(mFilename)
                        Call ReadText(mFilename)
                        On Error Resume Next
                        Kill mFilename
                        On Error GoTo eh
                    Case Else
                        Call ReadText(mFilename)
                End Select

            End If
                            
    End Select
    End With
    
    
    
    
    Screen.MousePointer = vbDefault
    Exit Property
eh: If InStr(1, Err.Description, "duplicate", vbTextCompare) Then
        Resume Next
    Else
        Call errHandler(srcFile & "CurrentFrame", s)
    End If
End Property

Private Function SaveToCSV(FileName As String) As String
On Error Resume Next
    Dim s As String
    Dim xlSheet As Object 'Excel.Worksheet
    Set xlSheet = GetObject(FileName).Sheets(1)
    s = TempFile("csv")
    Kill s
    Call xlSheet.SaveAs(s, 6, , , , , False)
    Set xlSheet = Nothing
    SaveToCSV = s
End Function

Private Sub ReadText(FileName As String)
On Error GoTo eh
    Dim i As Long
    Dim c As Long
    Dim s As String
    Dim rs As Recordset
    Dim c1 As String
    Dim c2 As String
    
    Dim rowkey As String
    Dim missingkey As Boolean
    
    
    Select Case FileExt(FileName)
        Case "csv": Call gData.LoadGrid(FileName, flexFileCommaText)
        Case Else:  Call gData.LoadGrid(FileName, flexFileTabText)
    End Select
    
    
    
    With gData
        
        'set column indexs
        For i = 1 To .cols - 1
            .ColKey(i) = Replace(.TextMatrix(1, i), " ", "")
        Next
        Call .RemoveItem(0)
        .FixedRows = 1
        
        'translate column headings
        On Error Resume Next
        For i = 1 To Parse(mColumnMappings)
            c1 = Trim(Parse(Parse(mColumnMappings, i), 1, "="))
            c2 = Trim(Parse(Parse(mColumnMappings, i), 2, "="))
            If c1 <> "" And c2 <> "" And .ColIndex(c1) <> -1 Then .ColKey(.ColIndex(c1)) = c2
        Next
        On Error GoTo eh
        For i = 1 To .cols - 1
            .TextMatrix(0, i) = .ColKey(i)
        Next
        
        
        'check that required cols are available
        s = ""
        For i = 1 To Parse(mKeyColumns)
            c1 = Parse(mKeyColumns, i)
            If c1 <> "" And .ColIndex(c1) = -1 Then
                s = s & ", " & c1
            Else
                .Cell(flexcpFontBold, 0, .ColIndex(c1)) = True
                .Cell(flexcpForeColor, 0, .ColIndex(c1)) = vbHighlight
            End If
        Next
        For i = 1 To Parse(mRequiredColumns)
            c1 = Parse(mRequiredColumns, i)
            If c1 <> "" And .ColIndex(c1) = -1 Then: s = s & ", " & c1
        Next
        If s <> "" Then Err.Raise 5, , "Required columns are missing: " & Mid(s, 3)
        
        
        
        'remove rows that have no keys
        For i = .Rows - 1 To 1 Step -1
            rowkey = ""
            missingkey = False
            For c = 1 To Parse(mKeyColumns)
                c1 = Parse(mKeyColumns, c)
                If .ColIndex(c1) < 0 Then Exit Sub
                rowkey = rowkey & Trim(.TextMatrix(i, .ColIndex(c1)))
                If Trim(.TextMatrix(i, .ColIndex(c1))) = "" Then missingkey = True
            Next
            
            If rowkey = "" Then
                Call .RemoveItem(i)
            Else
                If missingkey Then
                    Err.Raise 5, , "This file cannot be imported because key columns are empty. Please check your data and retry the import."
                End If
            End If
            
        Next
        
        
    End With
Exit Sub
eh: MsgBox Err.Description, vbExclamation, App.ProductName
    cmdNav(3).Enabled = False
End Sub


Private Sub cmdNav_Click(Index As Integer)
On Error Resume Next
    Select Case Index
        Case 0: Unload Me
        Case 1: CurrentFrame = CurrentFrame - 1
        Case 2: CurrentFrame = CurrentFrame + 1
        Case 3: If SaveData Then Unload Me
    End Select
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Const margin = 60
    Dim i As Long
    For i = 0 To WizFrame.UBound
        WizFrame(i).BorderStyle = 0
        WizFrame(i).Move 0, WizHead1.Height, Me.ScaleWidth, Me.ScaleHeight - WizHead1.Height - WizFoot.Height
    Next
    cmdNav(0).Move Me.ScaleWidth - (4 * (margin + cmdNav(0).Width)), 2 * margin
    cmdNav(1).Move Me.ScaleWidth - (3 * (margin + cmdNav(0).Width)), 2 * margin
    cmdNav(2).Move Me.ScaleWidth - (2 * (margin + cmdNav(0).Width)), 2 * margin
    cmdNav(3).Move Me.ScaleWidth - (1 * (margin + cmdNav(0).Width)), 2 * margin
    gData.Move margin, gData.Top, WizFrame(0).Width - 2 * margin, WizFrame(0).Height - gData.Top - margin
    ProgressBar.Move gData.left, gData.Top + gData.Height - ProgressBar.Height, gData.Width
End Sub

Private Sub Form_Load()
On Error GoTo eh
    
    WizHead1.Caption = "Import " & mTitle
    WizHead1.Description = "The " & mTitle & " import wizard will help you update your database"
    
    
    Call IniGetForm(Me, , mTableName)
Exit Sub
eh: Call errHandler(srcFile & "SaveData", s)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call IniPutForm(Me, , mTableName)
End Sub



Private Function SaveData() As Boolean
On Error GoTo eh
    Dim tabdef As Recordset
    
    Dim i As Long
    Dim c As Long
    Dim d As ADODB.Field
    Dim s As String
    Dim w As String
    
    Screen.MousePointer = vbHourglass
    ProgressBar.Visible = True
    
    Set tabdef = HFApp.SqlExec("select * from " & mTableName & " where 1=2")
    With gData
    For i = 1 To .Rows - 1
        ProgressBar.Value = (i / (.Rows - 1)) * 100
        s = ""
        w = ""
        For c = 1 To Parse(mKeyColumns)
            Set d = tabdef.fields(Parse(mKeyColumns, c))
            s = s & "," & DbQuote(DbQuoteType(d.Type), .TextMatrix(i, .ColIndex(Parse(mKeyColumns, c))), False, True, d.DefinedSize)
            w = w & " and " & d.Name & "=" & DbQuote(DbQuoteType(d.Type), .TextMatrix(i, .ColIndex(d.Name)), False, True, d.DefinedSize)
        Next
        For c = 1 To Parse(mRequiredColumns)
            Set d = tabdef.fields(Parse(mRequiredColumns, c))
            s = s & "," & DbQuote(DbQuoteType(d.Type), .TextMatrix(i, .ColIndex(Parse(mRequiredColumns, c))), False, True, d.DefinedSize)
        Next
        s = "INSERT INTO " & mTableName & "(" & mKeyColumns & IIf(mRequiredColumns = "", "", ",") & mRequiredColumns & ") VALUES (" & Mid(s, 2) & ")"
Debug.Print s
        Call HFApp.SqlExec(s, dbHomefront)
        
        s = ""
        For c = 1 To .cols - 1
            On Error Resume Next
            Set d = Nothing
            Set d = tabdef.fields(.ColKey(c))
            On Error GoTo eh
            If InStr(1, "," & mKeyColumns & ",", "," & .ColKey(c) & ",") = 0 And Not d Is Nothing Then
                s = s & "," & d.Name & "=" & DbQuote(DbQuoteType(d.Type), .TextMatrix(i, c), False, True, d.DefinedSize)
            End If
        Next
        s = "UPDATE " & mTableName & " SET " & Mid(s, 2) & " WHERE " & Mid(w, 6)
Debug.Print s
        Call HFApp.SqlExec(s, dbHomefront)
    
    Next
    End With
    
    mCancel = False
    ProgressBar.Visible = False
    Screen.MousePointer = vbDefault
    
    SaveData = True
Exit Function
eh: If InStr(1, Err.Description, "duplicate", vbTextCompare) Then
        Resume Next
    Else
        Call errHandler(srcFile & "SaveData", s)
    End If
End Function


Private Function MYTextMatrix(Row As Long, col As Long) As String
On Error Resume Next
    MYTextMatrix = gData.TextMatrix(Row, col)
End Function



