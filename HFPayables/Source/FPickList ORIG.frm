VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form FPickList 
   Caption         =   "Select List"
   ClientHeight    =   5835
   ClientLeft      =   8715
   ClientTop       =   4755
   ClientWidth     =   5670
   Icon            =   "FPickList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   5670
   Begin VB.CommandButton cmdNav 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   2
      Left            =   4320
      TabIndex        =   9
      Top             =   5400
      Width           =   1215
   End
   Begin VSFlex8Ctl.VSFlexGrid gList 
      Height          =   4035
      Left            =   120
      TabIndex        =   0
      Top             =   1260
      Width           =   5415
      _cx             =   9551
      _cy             =   7117
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
      SheetBorder     =   -2147483632
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
      FormatString    =   $"FPickList.frx":000C
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
      Editable        =   2
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
      FrozenRows      =   1
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
   Begin VB.PictureBox ViewsBar 
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   5655
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   840
      Width           =   5655
      Begin VB.ComboBox cboViews 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   60
         Width           =   4695
      End
      Begin VB.Label Label1 
         Caption         =   "View:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   300
         TabIndex        =   8
         Top             =   120
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "&New"
      Height          =   375
      Index           =   0
      Left            =   1740
      TabIndex        =   1
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "&Open"
      Default         =   -1  'True
      Height          =   375
      Index           =   1
      Left            =   3060
      TabIndex        =   2
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Click on a column heading to sort the list on that column."
      Height          =   255
      Left            =   300
      TabIndex        =   6
      Top             =   600
      Width           =   5235
   End
   Begin VB.Label Label2 
      Caption         =   "Type directly into the list to search for an item."
      Height          =   255
      Left            =   300
      TabIndex        =   5
      Top             =   360
      Width           =   5235
   End
   Begin VB.Label lblCaption 
      Caption         =   "Select an item from the list:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   180
      TabIndex        =   4
      Top             =   120
      Width           =   5355
   End
End
Attribute VB_Name = "FPickList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const SRCFILE = "FPickList::"

Private mCancel       As Boolean
Private mConnection   As Connection
Private mRecordset    As Recordset
Private mDefSortAsc   As Boolean
Private mViews        As String
Private mcolumns As Collection
Private mdata()  As String
Private mview    As Long


Private mItemIcon     As StdPicture
Private mHideColumns  As String
Private mMultiSelect  As Boolean


Public Function Choose(Connection As Connection, ItemTitle As String, Queries As String, Optional DefaultItem As String, Optional SortAsc As Boolean = True, Optional ShowNew As Boolean = False, Optional ItemIcon As IPictureDisp, Optional HideColumns As String, Optional MultiSelect As Boolean = False, Optional DefaultView As Long) As Boolean
On Error GoTo eh
    '---------------------------------------------------------------
    'Queries is a delimited list of views that may be displayed
    ' example:
    '---------------------------------------------------------------
    '  Dim SQL as string
    '  Dim cID as long
    '
    '  SQL = "All Customers     " & chr(1) & "select ID,Name,City from customers               " & chr(1) & Col1Format & chr(0) & _
    '        "Active Customers  " & chr(1) & "select ID,Name,City from customers where active=1" & chr(0) & _
    '        "InActive Customers" & chr(1) & "select ID,Name,City from customers where active=0"
    '
    '  cID = FPickList.Choose(myDB, "Customer", SQL)
    '
    '---------------------------------------------------------------
    ' example:
    '
    '  cID = FPickList.Choose(myDB, "Customer", "select ID,Name,City from customers")
    '
    '---------------------------------------------------------------
    Dim i As Long


    Screen.MousePointer = vbHourglass
    Me.Caption = App.ProductName & " - " & ItemTitle & " List"
    Call IniGetForm(Me, , Me.Caption)
    If CBool(InStr(1, "aeiou", left(ItemTitle, 1), vbTextCompare)) Then
        lblCaption = "Choose an " & ItemTitle & "."
    Else
        lblCaption = "Choose a " & ItemTitle & "."
    End If

    'curList.ConnectionString = ConnectString
    Set mConnection = Connection
    Set mItemIcon = ItemIcon
    
    
    mDefSortAsc = SortAsc
    mHideColumns = HideColumns
    
    If MultiSelect Then
        gList.SelectionMode = flexSelectionListBox
        gList.AllowSelection = True
    End If
    
    mViews = Queries
    cboViews.Clear
    For i = 1 To Parse(mViews, , Chr(0))
        'add titles to combobox
        cboViews.AddItem Parse(Parse(mViews, i, Chr(0)), 1, Chr(1))
    Next
    
    'config form
    If cboViews.ListCount > 1 Then
        ViewsBar.Visible = True
        gList.Move gList.left, 1260, gList.Width, 4035
    Else
        mViews = Chr(1) & mViews
        ViewsBar.Visible = False
        gList.Move gList.left, 900, gList.Width, 4395
    End If
    
    cboViews.ListIndex = Max(0, Min(cboViews.ListCount - 1, DefaultView - 1))
    
    If DefaultItem <> "" Then
        i = gList.FindRow(DefaultItem, , 0, False, False)
        If i > 0 Then
            Call gList.Select(i, 0)
            Call gList.ShowCell(i, 0)
            gList.TopRow = i
        End If
    End If
    
    If gList.Rows = 1 And cboViews.Visible Then
        Call SetCtrlFocus(cboViews)
    End If
    
    cmdNav(0).Visible = ShowNew
    
    Screen.MousePointer = vbDefault
    
    Me.Show vbModal
    On Error Resume Next
    Choose = UBound(mdata, 2) > 0 And Not mCancel

Exit Function:
eh: If err.Description = "Method 'FindRow' of object 'IVSFlexGrid' failed" Then
        Resume Next
    Else
        Call ErrHandler(SRCFILE & "Choose")
    End If
End Function


Public Function SelectedItems() As Long
    SelectedItems = UBound(mdata, 2)
End Function

Public Function SelectedItem(Column, Optional Row As Long = 1) As String
On Error Resume Next
    SelectedItem = mdata(mcolumns(Column), Row)
End Function

Public Function SelectedView() As Long
On Error Resume Next
    SelectedView = mview
End Function

Private Sub cboViews_Click()
On Error GoTo eh
    Dim i As Long
    Dim s As String
    
    Screen.MousePointer = vbHourglass
    
    
    Set mRecordset = New ADODB.Recordset
    
    s = Parse(Parse(mViews, cboViews.ListIndex + 1, Chr(0)), 2, Chr(1))
    If InStr(1, s, "order by ", vbTextCompare) Then
    Else
        s = s & " order by 1 " & IIf(mDefSortAsc, "asc ", "desc ")
    End If
    
    
    
' USED TO BE JUST THIS BUT SOMETIMES THE GRID DOESNT FILL. IT ADDS THE ROWS BUT THEY ARE ALL BLANK
'    mRecordset.Open s, mConnection
'
' THEN WE TESTED ALL THESE. THE LAST ONE WORKED. EXCEPT YOU CANT ADD THE FILTER ROW TO A DYNAMIC
' QUERY THAT COMES FROM ACCESS. THIS CAUSED ERRORS WHEN SELECTING FROM ONCENTER SOFTWARES DB.
'    'mRecordset.Open s, mConnection, adOpenForwardOnly 'sometimes grid doesnt fill
'    'mRecordset.Open s, mConnection, adOpenKeyset      'cant add filter row
'    'mRecordset.Open s, mConnection, adOpenStatic      'cant add filter row
'    mRecordset.Open s, mConnection, adOpenDynamic      'this is the only one left. thankfully it works
        
' SURE HOPE THIS ONE WORKS...
'    If (mConnection.Properties("DBMS Name") = "MS Jet") Then
'        mRecordset.Open s, mConnection
'    Else
'        mRecordset.Open s, mConnection, adOpenDynamic
'    End If
'
' DARYL! YOU HAVE TO ADD COMMENTS... I GUESS IT STILL DIDNT WORK.
'    If HFApp.Options.ValueByName("ForwardOnly") = "True" Then
'        mRecordset.Open s, mConnection, adOpenForwardOnly     'this is the only one left. thankfully it works
'    Else
'        mRecordset.Open s, mConnection, adOpenDynamic      'this is the only one left. thankfully it works
'    End If
'
' JET DBS MUST USE FORWARDONLY, OTHERWISE THE FILTERBAR ROW CANT BE ADDED
    If (mConnection.Properties("DBMS Name") = "MS Jet") Then
        mRecordset.Open s, mConnection
    Else
        If HFApp.Options.ValueByName("ForwardOnly") = "True" Then
            mRecordset.Open s, mConnection, adOpenForwardOnly
        Else
            mRecordset.Open s, mConnection, adOpenDynamic
        End If
    End If
    
    
    
    
    
    gList.DataMode = flexDMBoundNoRowCount
    Set gList.DataSource = mRecordset
 
    
    Call gList.AddItem("", 1)
    gList.Cell(flexcpBackColor, 1, 0, 1, gList.Cols - 1) = vbInfoBackground
    gList.FrozenRows = 1
    
    On Error Resume Next
    Set gList.Cell(flexcpPicture, 2, 0, gList.Rows - 1, 0) = mItemIcon
    gList.AutoSize 0
    For i = 0 To gList.Cols - 1
        
        gList.ColHidden(i) = InStr(1, "," & mHideColumns & ",", "," & gList.TextMatrix(0, i) & ",", vbTextCompare)
    
    
        Select Case mRecordset.Fields(i).Type
            Case adDouble
       '         gList.ColFormat(i) = "#,###.00"
       '         gList.ColDataType(i) = flexDTCurrency
            Case adDate, adDBDate
'                gList.ColFormat(i) = hfapp.Options(FormatDate)
                gList.ColDataType(i) = flexDTDate
        End Select
    Next
    
    
    If gList.Rows > 1 Then
        gList.SetFocus
        Call SetCtrlFocus(gList)
        gList.Select 2, 0
    End If
    
    cmdNav(1).Enabled = gList.Row > 0
    
    
        
    Screen.MousePointer = vbDefault
Exit Sub
eh: Call ErrHandler(SRCFILE & "cboView_Click")
End Sub

Private Sub cmdNav_Click(Index As Integer)
On Error Resume Next
'On Error GoTo eh
    Dim i As Long
    Dim r As Long
    Dim c As Long
    Dim s As String
        
    With gList
        Set mcolumns = New Collection
        For c = 0 To .Cols - 1
            Call mcolumns.Add(c, Replace(.TextMatrix(0, c), " ", ""))
        Next
        
        mview = cboViews.ListIndex + 1
        
        Select Case Index
            Case 0 'New
                mCancel = False
                ReDim mdata(.Cols - 1, 1)
            
            Case 1 'OK
                If .Row = 1 Then Exit Sub
                mCancel = False
                If .SelectionMode = flexSelectionListBox Then
                    ReDim mdata(.Cols - 1, 0)
                    For i = 0 To .SelectedRows - 1
                        r = .SelectedRow(i)
                        If r > 1 And Not .RowHidden(r) Then
                            ReDim Preserve mdata(.Cols - 1, UBound(mdata, 2) + 1)
                            For c = 0 To gList.Cols - 1
                                mdata(c, UBound(mdata, 2)) = LTrim(gList.TextMatrix(r, c))
                            Next
                        End If
                     Next
                Else
                    If .Row <= 1 Then
                        ReDim mdata(.Cols - 1, 0)
                    Else
                        ReDim mdata(.Cols - 1, 1)
                        For c = 0 To gList.Cols - 1
                            mdata(c, 1) = LTrim(gList.TextMatrix(.Row, c))
                        Next
                    End If
                End If
                
                
            Case 2 'Cancel
                mCancel = True
                ReDim mdata(.Cols - 1, 0)
        
        End Select
    End With
    Unload Me
Exit Sub
End Sub


Private Sub Form_Resize()
On Error Resume Next

    Me.Height = Max(Me.Height, 4035)
    Me.Width = Max(Me.Width, 4515)

    ViewsBar.Width = Me.ScaleWidth
    cboViews.Width = ViewsBar.ScaleWidth - cboViews.left - gList.left
    gList.Move gList.left, gList.Top, Me.ScaleWidth - 2 * gList.left, Me.ScaleHeight - gList.Top - cmdNav(0).Height - 2 * gList.left
    cmdNav(3).Move Me.ScaleWidth - 3 * cmdNav(0).Width - 3 * gList.left, Me.ScaleHeight - cmdNav(0).Height - gList.left
    cmdNav(0).Move Me.ScaleWidth - 3 * cmdNav(0).Width - 3 * gList.left, Me.ScaleHeight - cmdNav(0).Height - gList.left
    cmdNav(1).Move Me.ScaleWidth - 2 * cmdNav(0).Width - 2 * gList.left, Me.ScaleHeight - cmdNav(0).Height - gList.left
    cmdNav(2).Move Me.ScaleWidth - 1 * cmdNav(0).Width - 1 * gList.left, Me.ScaleHeight - cmdNav(0).Height - gList.left
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Call IniPutForm(Me, , Me.Caption)
End Sub

Private Sub gList_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim r As Long
    Dim c As Long
    
    With gList
        If .Row <> 1 Then Exit Sub
        .Redraw = flexRDNone
        For r = 2 To .Rows - 1
            .RowHidden(r) = False
            If .EditText <> "" Then
                For c = 0 To .Cols - 1
                    If Not (UCase(.Cell(flexcpTextDisplay, r, c)) Like "*" & UCase(.Cell(flexcpTextDisplay, 1, c)) & "*") Then
                        .RowHidden(r) = True
                        Exit For
                    End If
                Next
            End If
        Next
        .Redraw = flexRDBuffered
    End With
End Sub

Private Sub gList_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With gList
        If Row = 1 Then
            Cancel = False
            .AutoSearch = flexSearchNone
        Else
            Cancel = True
            .AutoSearch = flexSearchFromCursor
        End If
    End With
End Sub

Private Sub gList_AfterSort(ByVal Col As Long, Order As Integer)
    gList.FixedRows = 1
End Sub

Private Sub gList_BeforeSort(ByVal Col As Long, Order As Integer)
    gList.FixedRows = 2
End Sub




Private Sub gList_DblClick()
On Error GoTo eh
    If gList.Row > 0 Then
        Call cmdNav_Click(1)
    End If
Exit Sub
eh: Call ErrHandler(SRCFILE & "gList_DblClick")
End Sub

Private Sub gList_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    With gList
        If .Row = 1 And .RowSel = 1 And KeyCode = vbKeyDelete Then
            .Text = ""
            For i = 2 To .Rows - 1
                .RowHidden(i) = False
            Next
        End If
    End With
End Sub




