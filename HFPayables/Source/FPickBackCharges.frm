VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form FPickBackCharges 
   Caption         =   "Back Chargable PO's"
   ClientHeight    =   5640
   ClientLeft      =   3480
   ClientTop       =   4395
   ClientWidth     =   8340
   Icon            =   "FPickBackCharges.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   8340
   Begin VB.CommandButton cmdNav 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   2
      Left            =   4320
      TabIndex        =   3
      Top             =   5400
      Width           =   1215
   End
   Begin VSFlex8Ctl.VSFlexGrid gList 
      Height          =   4005
      Left            =   120
      TabIndex        =   1
      Top             =   1290
      Width           =   5415
      _cx             =   9551
      _cy             =   7064
      Appearance      =   1
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
      GridColorFixed  =   -2147483633
      TreeColor       =   -2147483632
      FloodColor      =   -2147483635
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FPickBackCharges.frx":000C
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   1
      AutoSearchDelay =   3
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   5
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
   Begin VB.CommandButton cmdNav 
      Caption         =   "&Open"
      Height          =   375
      Index           =   1
      Left            =   3060
      TabIndex        =   2
      Top             =   5400
      Width           =   1215
   End
   Begin TDBDate6Ctl.TDBDate dteAccounting 
      Height          =   225
      Left            =   2370
      TabIndex        =   0
      Top             =   900
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   397
      Calendar        =   "FPickBackCharges.frx":0048
      Caption         =   "FPickBackCharges.frx":013D
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FPickBackCharges.frx":01A2
      Keys            =   "FPickBackCharges.frx":01C0
      Spin            =   "FPickBackCharges.frx":022C
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   0
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "dd-mmm-yy"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   -1
      FirstMonth      =   4
      ForeColor       =   -2147483640
      Format          =   "dd-mmm-yy"
      HighlightText   =   2
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxDate         =   2958465
      MinDate         =   -657434
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "__-___-__"
      ValidateMode    =   0
      ValueVT         =   1
      Value           =   1.10541259469229E-317
      CenturyMode     =   0
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   270
      Picture         =   "FPickBackCharges.frx":0254
      Top             =   270
      Width           =   480
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"FPickBackCharges.frx":0B1E
      Height          =   735
      Index           =   0
      Left            =   930
      TabIndex        =   5
      Top             =   150
      Width           =   6825
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Accounting Date"
      Height          =   195
      Index           =   10
      Left            =   1125
      TabIndex        =   4
      Top             =   900
      Width           =   1200
   End
End
Attribute VB_Name = "FPickBackCharges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const SRCFILE = "FPickList::"

Private mRecordset    As Recordset
Private mcolumns As Collection
Private mdata() As String
Private mAccountingDate As Date

Private mQuery As String
Private mItemIcon     As StdPicture
Private mHideColumns  As String

Public Function Choose(Query As String, Optional ItemIcon As IPictureDisp, Optional HideColumns As String, Optional AccountingDate) As Boolean
On Error GoTo eh
    
    Screen.MousePointer = vbHourglass

    If IsMissing(AccountingDate) Or IsNull(AccountingDate) Then
        mAccountingDate = VBA.Date
    Else
        mAccountingDate = AccountingDate
    End If
    mQuery = Query
    Set mItemIcon = ItemIcon
    mHideColumns = HideColumns
    
    gList.SelectionMode = flexSelectionListBox
    gList.AllowSelection = True
    Screen.MousePointer = vbDefault
    
    Me.Show vbModal
    Choose = UBound(mdata, 1) > 0

Exit Function:
eh: If Err.Description = "Method 'FindRow' of object 'IVSFlexGrid' failed" Then
        Resume Next
    Else
        Call ErrHandler(SRCFILE & "Choose")
End If
End Function


Public Function SelectedItems() As Long
    SelectedItems = UBound(mdata, 1)
End Function

Public Function SelectedItem(Column, Optional Row As Long = 1) As String
On Error Resume Next
    If Column = "AccountingDate" Then
        SelectedItem = mAccountingDate
    Else
        SelectedItem = mdata(Row, mcolumns(Column))
    End If
End Function


Private Sub LoadData()
On Error GoTo eh
    Dim i As Long
    Dim s As String
    
    Screen.MousePointer = vbHourglass
    Set mRecordset = New Recordset
    mRecordset.Open mQuery, HFApp.Databases(AccountingDB)
    Set gList.DataSource = mRecordset
    
    
    On Error Resume Next
    Set gList.Cell(flexcpPicture, 1, 0, gList.Rows - 1, 0) = mItemIcon
    gList.AutoSize 0
    For i = 0 To gList.Cols - 1
        
        gList.ColHidden(i) = InStr(1, "," & mHideColumns & ",", "," & gList.TextMatrix(0, i) & ",", vbTextCompare)
    
    
        Select Case mRecordset.Fields(i).Type
            Case adDouble
                gList.ColFormat(i) = "#,###.00"
                gList.ColDataType(i) = flexDTCurrency
            Case adDate, adDBDate
                gList.ColDataType(i) = flexDTDate
        End Select
    Next
    
    
    If gList.Rows > 1 Then
        gList.SetFocus
        Call SetCtrlFocus(gList)
        gList.Select 1, 0
    End If
    
    cmdNav(1).Enabled = gList.Row > 0
    
    Screen.MousePointer = vbDefault
Exit Sub
eh: Call ErrHandler(SRCFILE & "cboView_Click")
End Sub

Private Sub cmdNav_Click(Index As Integer)
On Error Resume Next
    Dim i As Long
    Dim r As Long
    Dim c As Long
        
    With gList
        Set mcolumns = New Collection
        For c = 0 To .Cols - 1
            Call mcolumns.Add(c, Replace(.TextMatrix(0, c), " ", ""))
        Next
        
        Select Case Index
            
            Case 1 'OK
                ReDim mdata(.SelectedRows, .Cols - 1)
                For i = 0 To .SelectedRows - 1
                    r = .SelectedRow(i)
                    For c = 0 To gList.Cols - 1
                        mdata(i + 1, c) = LTrim(gList.TextMatrix(r, c))
                    Next
                 Next
                
                
            Case 2 'Cancel
                ReDim mdata(0, .Cols - 1)
                
        End Select
    End With
    Unload Me
Exit Sub
eh: Call ErrHandler(SRCFILE & "cmdNav_Click")
End Sub


Private Sub Form_Load()
    Call IniGetForm(Me)
    dteAccounting.Format = App.Options.Value(TimberlineDateFormat)
    dteAccounting.DisplayFormat = App.Options.Value(TimberlineDateFormat)
    dteAccounting.Value = mAccountingDate
    LoadData
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Me.Height = Max(Me.Height, 4035)
    Me.Width = Max(Me.Width, 4515)
    gList.Move gList.left, gList.Top, Me.ScaleWidth - 2 * gList.left, Me.ScaleHeight - gList.Top - cmdNav(1).Height - 2 * gList.left
    cmdNav(1).Move Me.ScaleWidth - 2 * cmdNav(1).Width - 2 * gList.left, Me.ScaleHeight - cmdNav(1).Height - gList.left
    cmdNav(2).Move Me.ScaleWidth - 1 * cmdNav(1).Width - 1 * gList.left, Me.ScaleHeight - cmdNav(1).Height - gList.left
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Call IniPutForm(Me)
End Sub

Private Sub gList_BeforeSort(ByVal Col As Long, Order As Integer)
On Error GoTo eh
    Dim i As Long
    Dim s As String
    
    s = Parse(mQuery, 1, "ORDER BY ")
    s = s & " order by " & Col + 1 & IIf(Order = 1, " asc ", " desc")
    Set mRecordset = HFApp.SqlExec(s, AccountingDB)
    
    
On Error Resume Next
    gList.Col = gList.MouseCol
Exit Sub
eh: Call ErrHandler(SRCFILE & "gList_BeforeSort")
End Sub

Private Sub gList_DblClick()
On Error GoTo eh
    If gList.Row > 0 Then
        Call cmdNav_Click(1)
    End If
Exit Sub
eh: Call ErrHandler(SRCFILE & "gList_DblClick")
End Sub

Private Sub gList_SelChange()
    cmdNav(1).Default = gList.SelectedRows > 0
End Sub
