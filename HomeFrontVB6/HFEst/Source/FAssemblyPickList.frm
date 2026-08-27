VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{E502CA28-473F-401A-943A-CF61F7BD61D0}#1.1#0"; "vbalARLB6.ocx"
Begin VB.Form FAssemblyPickList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Items"
   ClientHeight    =   5775
   ClientLeft      =   1695
   ClientTop       =   2655
   ClientWidth     =   8265
   Icon            =   "FAssemblyPickList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   8265
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboDataSource 
      Height          =   240
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   423
      Style           =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ExtendedUI      =   0   'False
      DropDownWidth   =   0
   End
   Begin vbalARListBarLib6.vbalARListBar ButtonBar 
      Height          =   4770
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   8414
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonWidth     =   84
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "&Open"
      Default         =   -1  'True
      Height          =   375
      Index           =   1
      Left            =   5700
      TabIndex        =   1
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton cmdNav 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   2
      Left            =   6960
      TabIndex        =   2
      Top             =   5280
      Width           =   1215
   End
   Begin VSFlex8Ctl.VSFlexGrid gData 
      Height          =   5055
      Left            =   1500
      TabIndex        =   0
      Top             =   120
      Width           =   6675
      _cx             =   11774
      _cy             =   8916
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
      GridColorFixed  =   -2147483633
      TreeColor       =   -2147483632
      FloodColor      =   -2147483635
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
      Cols            =   10
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FAssemblyPickList.frx":000C
      ScrollTrack     =   0   'False
      ScrollBars      =   2
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   1
      AutoSearchDelay =   3
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   4
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
   Begin VB.Label lblDebug 
      AutoSize        =   -1  'True
      Caption         =   "Debug"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   60
      TabIndex        =   5
      Top             =   5520
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "FAssemblyPickList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const SRCFILE = "FItemPickList::"

Private mStatus           As AssemblyStatuses
Private mGlobalOptions    As Boolean
Private mModelsAndOptions As Boolean
Private mDCOptions        As Boolean

Private Type ViewDefs
    name         As String
    KeyFlds      As String
    DisplayFlds  As String
    LevelIcons   As String
    FilterClause As String
    AssemblyType As AssemblyTypes
End Type
Private mViews()   As ViewDefs
Private mViewIndex As Long
Private mViewQuery As String

Private mSelectedDataSource As Long
Private mSelectedItems()     As String


Public Function Choose(Caption As String, MultiSelect As Boolean, GlobalOptions As Boolean, ModelsAndOptions As Boolean, DCOptions As Boolean, DataSource As Boolean, Status As AssemblyStatuses) As Boolean
On Error GoTo eh
    ReDim mSelectedItems(6, 0) As String
    cboDataSource.Enabled = DataSource
    cboDataSource.Visible = cboDataSource.Enabled
    Me.Caption = Caption
    
    mStatus = Status
    mGlobalOptions = GlobalOptions
    mModelsAndOptions = ModelsAndOptions
    mDCOptions = DCOptions
    
    gData.AllowSelection = MultiSelect
    Call LoadViews
    Me.Show vbModal
    Choose = UBound(mSelectedItems, 2) > 0
Exit Function:
eh: Call ErrHandler(SRCFILE & "Choose")
End Function


Private Sub cboDataSource_Click()
    Call ButtonBar_ItemClick(ButtonBar.SelectedIndex)
End Sub

Private Sub cmdNav_Click(Index As Integer)
On Error Resume Next
    Dim r As Long
    Dim i As Long
    Dim AssemblyType As AssemblyTypes

    With gData
        Select Case Index
            Case 1 'OK


                For r = 0 To .SelectedRows - 1
                    If .TextMatrix(.SelectedRow(r), .ColIndex("Assembly")) <> "" Then
                        i = i + 1
                        ReDim Preserve mSelectedItems(6, i) As String
                        
                        
                        mSelectedDataSource = cboDataSource.ListIndex
                        AssemblyType = .ValueMatrix(.SelectedRow(r), .ColIndex("AssemblyType"))
                        
                        'return community depending on option type and settings
                        Select Case True
                            Case AssemblyType = atModel And HFApp.Options(ModelByArea):           mSelectedItems(1, i) = .TextMatrix(.SelectedRow(r), .ColIndex("Community"))
                            Case AssemblyType = atOption And HFApp.Options(OptionByArea):         mSelectedItems(1, i) = .TextMatrix(.SelectedRow(r), .ColIndex("Community"))
                            Case AssemblyType = atGlobal And HFApp.Options(GlobalOptionByArea):   mSelectedItems(1, i) = .TextMatrix(.SelectedRow(r), .ColIndex("Community"))
                            Case AssemblyType = atDesignCenter And HFApp.Options(DCOptionByArea): mSelectedItems(1, i) = .TextMatrix(.SelectedRow(r), .ColIndex("Community"))
                            Case Else:                                                                                                       mSelectedItems(1, i) = ""
                        End Select
                        
                        'return phase depending on option type and settings
                        Select Case True
                            Case AssemblyType = atModel And HFApp.Options(ModelsByArea_Phase):         mSelectedItems(2, i) = .TextMatrix(.SelectedRow(r), .ColIndex("CommunityPhase"))
                            Case AssemblyType = atOption And HFApp.Options(OptionByAreaPhase):         mSelectedItems(2, i) = .TextMatrix(.SelectedRow(r), .ColIndex("CommunityPhase"))
                            Case AssemblyType = atGlobal And HFApp.Options(GlobalOptionByAreaPhase):   mSelectedItems(2, i) = .TextMatrix(.SelectedRow(r), .ColIndex("CommunityPhase"))
                            Case AssemblyType = atDesignCenter And HFApp.Options(DCOptionByAreaPhase): mSelectedItems(2, i) = .TextMatrix(.SelectedRow(r), .ColIndex("CommunityPhase"))
                            Case Else:                                                                                                            mSelectedItems(2, i) = ""
                        End Select

                        mSelectedItems(3, i) = .TextMatrix(.SelectedRow(r), .ColIndex("Assembly"))
                        mSelectedItems(4, i) = .TextMatrix(.SelectedRow(r), .ColIndex("Model"))
                        mSelectedItems(5, i) = .TextMatrix(.SelectedRow(r), .ColIndex("OptionID"))
                        mSelectedItems(6, i) = AssemblyType
                        
                    End If
                Next
                
                If i > 0 Then Unload Me

            Case 2 'Cancel
                Unload Me
        End Select
    End With
Exit Sub
eh: Call ErrHandler(SRCFILE & "cmdNav_Click")
End Sub


Private Sub Form_Load()
    lblDebug.Visible = InIde
    Call IniGetForm(Me)
End Sub




Public Function SelectedItems() As Long
    SelectedItems = UBound(mSelectedItems, 2)
End Function
Public Function SelectedCommunity(Index As Long) As String
    SelectedCommunity = mSelectedItems(1, Index)
End Function
Public Function SelectedCommunityPhase(Index As Long) As String
    SelectedCommunityPhase = mSelectedItems(2, Index)
End Function
Public Function SelectedDataSource() As Long
    SelectedDataSource = mSelectedDataSource
End Function

Public Function SelectedAssembly(Index As Long) As String
    SelectedAssembly = mSelectedItems(3, Index)
End Function
Public Function SelectedModel(Index As Long) As String
    SelectedModel = mSelectedItems(4, Index)
End Function
Public Function SelectedOptionID(Index As Long) As String
    SelectedOptionID = mSelectedItems(5, Index)
End Function
Public Function SelectedAssemblyType(Index As Long) As AssemblyTypes
    SelectedAssemblyType = mSelectedItems(6, Index)
End Function

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    If cboDataSource.Enabled Then Call IniPut(AppIni, "FItemPicklist", "DataSource", cboDataSource.ListIndex)
    Call IniPutForm(Me)
End Sub


Private Sub gData_DblClick()
    Call cmdNav_Click(1)
End Sub

Private Sub gData_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    With gData
        If .Row < 0 Then Exit Sub
        Select Case KeyCode
                
            Case vbKeyLeft
                If .GetNodeRow(.Row, flexNTFirstChild) <> -1 And .IsCollapsed(.Row) <> flexOutlineCollapsed Then
                    .IsCollapsed(.Row) = flexOutlineCollapsed
                Else
                    If .GetNodeRow(.Row, flexNTParent) <> -1 Then .Row = .GetNodeRow(.Row, flexNTParent)
                End If
                
            Case vbKeyReturn, vbKeyRight
                If .GetNodeRow(.Row, flexNTFirstChild) <> -1 Then
                    If .IsCollapsed(.Row) = flexOutlineCollapsed Then
                        .IsCollapsed(.Row) = flexOutlineExpanded
                    Else
                        .Row = .GetNodeRow(.Row, flexNTFirstChild)
                    End If
                End If
                
        End Select
    End With
End Sub

Private Sub gData_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    gData.ToolTipText = gData.ColKey(gData.MouseCol)
End Sub

Private Sub gData_RowColChange()
'Me.Caption = Me.gData.TextMatrix(gData.Row, gData.ColIndex("Assembly"))
'    gData.Col = 0
End Sub






Private Sub LoadViews()
    Dim i As Long
    ReDim mViews(2) As ViewDefs
    
    
    ButtonBar.ImageList = FMain.LargeIcons
    
    
    
    
    mViews(i).name = "Models and Options"
    If mModelsAndOptions Then Call ButtonBar.Add(, mViews(i).name, FMain.LargeIcons.ListImages.Item(mViews(i).name).Index - 1, , , , i)
    mViews(i).FilterClause = "AssemblyType IN(" & atModel & "," & atOption & ")"
    mViews(i).KeyFlds = Mid(IIf(HFApp.Options(ModelByArea), ",Community", "") & IIf(HFApp.Options(ModelByArea) And HFApp.Options(ModelsByArea_Phase), ",CommunityPhase", "") & ",Model,OptionID", 2)
    mViews(i).DisplayFlds = Mid(IIf(HFApp.Options(ModelByArea), ",Community + '  --  ' + CommunityDesc", "") & IIf(HFApp.Options(ModelByArea) And HFApp.Options(ModelsByArea_Phase), ",CommunityPhase", "") & ",Assembly + '  --  ' + AssemblyDesc,Assembly + '  --  ' + AssemblyDesc", 2)
    mViews(i).LevelIcons = Mid(IIf(HFApp.Options(ModelByArea), ",area", "") & IIf(HFApp.Options(ModelByArea) And HFApp.Options(ModelsByArea_Phase), ",category", "") & ",model,option", 2)
    mViews(i).AssemblyType = atOption 'models are handled special
    i = i + 1
    
    
    mViews(i).name = "Global Options"
    If mGlobalOptions Then Call ButtonBar.Add(, mViews(i).name, FMain.LargeIcons.ListImages.Item(mViews(i).name).Index - 1, , , , i)
    mViews(i).FilterClause = "AssemblyType=" & atGlobal
    mViews(i).KeyFlds = Mid(IIf(HFApp.Options(GlobalOptionByArea), ",Community", "") & IIf(HFApp.Options(GlobalOptionByArea) And HFApp.Options(GlobalOptionByAreaPhase), ",CommunityPhase", "") & IIf(HFApp.Options(UseMajorGroup), ",MajorGroup", "") & ",Category,OptionID", 2)
    mViews(i).DisplayFlds = Mid(IIf(HFApp.Options(GlobalOptionByArea), ",Community + '  --  ' + CommunityDesc", "") & IIf(HFApp.Options(GlobalOptionByArea) And HFApp.Options(GlobalOptionByAreaPhase), ",CommunityPhase", "") & IIf(HFApp.Options(UseMajorGroup), ",MajorGroupDesc", "") & ",CategoryDesc,Assembly + '  --  ' + AssemblyDesc", 2)
    mViews(i).LevelIcons = Mid(IIf(HFApp.Options(GlobalOptionByArea), ",area", "") & IIf(HFApp.Options(GlobalOptionByArea) And HFApp.Options(GlobalOptionByAreaPhase), ",category", "") & IIf(HFApp.Options(UseMajorGroup), ",groupphase", "") & ",phase,option", 2)
    mViews(i).AssemblyType = atGlobal
    i = i + 1
    
    
    mViews(i).name = "Design Center Options"
    If mDCOptions Then Call ButtonBar.Add(, mViews(i).name, FMain.LargeIcons.ListImages.Item(mViews(i).name).Index - 1, , , , i)
    mViews(i).FilterClause = "AssemblyType=" & atDesignCenter
    mViews(i).KeyFlds = Mid(IIf(HFApp.Options(DCOptionByArea), ",Community", "") & IIf(HFApp.Options(DCOptionByArea) And HFApp.Options(DCOptionByAreaPhase), ",CommunityPhase", "") & IIf(HFApp.Options(UseMajorGroup), ",MajorGroup", "") & ",Category,OptionID", 2)
    mViews(i).DisplayFlds = Mid(IIf(HFApp.Options(DCOptionByArea), ",Community + '  --  ' + CommunityDesc", "") & IIf(HFApp.Options(DCOptionByArea) And HFApp.Options(DCOptionByAreaPhase), ",CommunityPhase", "") & IIf(HFApp.Options(UseMajorGroup), ",MajorGroupDesc", "") & ",CategoryDesc,Assembly + '  --  ' + AssemblyDesc", 2)
    mViews(i).LevelIcons = Mid(IIf(HFApp.Options(DCOptionByArea), ",area", "") & IIf(HFApp.Options(DCOptionByArea) And HFApp.Options(DCOptionByAreaPhase), ",category", "") & IIf(HFApp.Options(UseMajorGroup), ",groupphase", "") & ",phase,option", 2)
    mViews(i).AssemblyType = atDesignCenter
    i = i + 1
    
    
    ButtonBar.ItemSelected(1) = True
    
    cboDataSource.AddItem "Estimating"
    cboDataSource.AddItem "Sales Items"
    
    
    'only remember this if you are allowed to choose it
    If cboDataSource.Enabled Then
        cboDataSource.ListIndex = Val(IniGet(AppIni, "FItemPicklist", "DataSource"))
    Else
        cboDataSource.ListIndex = 0
    End If
        
    
    
    
End Sub


Private Sub ButtonBar_ItemClick(ByVal lIndex As Long)
    mViewIndex = ButtonBar.ItemData(lIndex)
    Call LoadData(True)
End Sub
Private Sub gData_BeforeCollapse(ByVal Row As Long, ByVal State As Integer, Cancel As Boolean)
Static bInHere As Boolean
    With gData
        If State = flexOutlineExpanded Then
            .Row = Row
            Call LoadData(False)
        End If
    End With
End Sub

Private Sub LoadData(ClearTree As Boolean)
On Error Resume Next

    Dim rs As Recordset
    Dim r As Long
    Dim s       As String
    Dim i       As Long
    Dim levels  As Long
    Dim Level   As Long
    Dim n       As VSFlexNode
    Dim WhereClause As String
    Dim KeyFld     As String
    Dim DisplayFld As String
    Dim KeyValue   As String
    Dim DisplayValue   As String
    Dim Picture As StdPicture
    Dim ParentRow As Long

'PROPERTY                  CONTAINS
'.rowdata()                Where clause
'.cell(flexcpText, r, 0)   DisplayFld Value
'.cell(flexcpData, r, 0)   DisplayFld Field Name (also used as tooltip)
'.Cell(flexcpText, r, 1)   KeyFld Value
'.Cell(flexcpData, r, 1)   KeyFld Field Name
    
    With gData
        ParentRow = .Row
        '.Redraw = flexRDNone
        levels = Parse(mViews(mViewIndex).DisplayFlds)
        
        If ClearTree Then .Rows = 0
        
        If .Rows = 0 Then
            Level = -1
        Else
            Level = .RowOutlineLevel(.Row)
        End If
        
        If Level + 1 = levels Then
            'user has opened the lowest level so do nothing
        Else
            On Error Resume Next
            If .TextMatrix(.GetNodeRow(.Row, flexNTFirstChild), 0) = "dummy" Then
                .RemoveItem .GetNodeRow(.Row, flexNTFirstChild)
            Else
                .Redraw = flexRDBuffered
                Exit Sub
            End If
            On Error GoTo eh
            
            KeyFld = Parse(mViews(mViewIndex).KeyFlds, Level + 2)
            DisplayFld = Parse(mViews(mViewIndex).DisplayFlds, Level + 2)
            
            
            
            s = ""
            s = s & "SELECT DISTINCT " & KeyFld & "," & DisplayFld & vbCrLf
            s = s & "  FROM " & IIf(cboDataSource.ListIndex = 0, "AssemblyList", "SalesAssemblyList") & vbCrLf
            On Error Resume Next
            WhereClause = .RowData(.Row)
            If WhereClause = "" Then WhereClause = " AND " & mViews(mViewIndex).FilterClause
            On Error GoTo eh
            s = s & " WHERE " & Mid(WhereClause, 6) & vbCrLf
            'for model view only
            If mViews(mViewIndex).name = "Models and Options" Then
                If KeyFld = "Model" Then s = s & " AND AssemblyType=" & atModel & vbCrLf
                If KeyFld = "OptionID" Then s = s & " AND AssemblyType=" & atOption & vbCrLf
            End If

            s = s & " AND Status<=" & mStatus & vbCrLf
            
            s = s & "ORDER BY 2" '& DisplayFld
            
            mViewQuery = s
            
            Set rs = HFApp.SqlExec(s)
            Set Picture = FMain.SmallIcons.ListImages.Item(Parse(mViews(mViewIndex).LevelIcons, Level + 2)).Picture
            While Not rs.EOF
                KeyValue = "" & rs(0)
                DisplayValue = Trim("" & rs(1))
                
                If Level = -1 Then
                    r = .Rows
                    
                    Call .AddItem(DisplayValue, r)
                    
                    .TextMatrix(r, .ColIndex(KeyFld)) = KeyValue
                    If KeyFld = "Model" Or KeyFld = "OptionID" Then
                        .TextMatrix(r, .ColIndex("Assembly")) = Trim(Parse(DisplayValue, 1, "  --  "))
                    End If
                    
                    .Cell(flexcpPicture, r, 0) = Picture

                    .Cell(flexcpData, r, 0) = DisplayFld
                    .Cell(flexcpText, r, 1) = KeyValue
                    .Cell(flexcpData, r, 1) = KeyFld
                    
                    .IsSubtotal(r) = True
                    .RowOutlineLevel(r) = Level + 1
                    If Level + 2 <> levels Then
                        Call .GetNode(r).AddNode(flexNTFirstChild, "dummy") ' add dummy child row so grid knows this one can be opened
                    End If
                    Set n = .GetNode(r)
                    n.Expanded = False
                    
                    .RowData(r) = WhereClause & " AND " & KeyFld & " = " & DbQuote(Str, KeyValue)
                    
                Else
                    r = ParentRow
                    Set n = .GetNode(r).AddNode(flexNTLastChild, DisplayValue)
                    
                    'set all text values to same as parent row
                    For i = 1 To .Cols - 1
                        .TextMatrix(n.Row, i) = .TextMatrix(r, i)
                    Next
                    'now set new value
                    .TextMatrix(n.Row, .ColIndex(KeyFld)) = KeyValue
                    
                    If KeyFld = "Model" Then
                        .TextMatrix(n.Row, .ColIndex("Assembly")) = Trim(Parse(DisplayValue, 1, "  --  "))
                        .TextMatrix(n.Row, .ColIndex("AssemblyType")) = atModel
                    End If
                    If KeyFld = "OptionID" Then
                        .TextMatrix(n.Row, .ColIndex("Assembly")) = Trim(Parse(DisplayValue, 1, "  --  "))
                        .TextMatrix(n.Row, .ColIndex("AssemblyType")) = mViews(mViewIndex).AssemblyType
                    End If
                    
                    
                    
                    .Cell(flexcpPicture, n.Row, 0) = Picture
                    .Cell(flexcpData, n.Row, 0) = DisplayFld
                    .Cell(flexcpText, n.Row, 1) = KeyValue
                    .Cell(flexcpData, n.Row, 1) = KeyFld
                    If Level + 2 < levels Then
                        Call .GetNode(n.Row).AddNode(flexNTFirstChild, "dummy") ' add dummy child row so grid knows this one can be opened
                    End If
                    n.Expanded = False
                    .RowData(n.Row) = WhereClause & " AND " & KeyFld & " = " & DbQuote(Str, KeyValue)
                    
                End If
                
                rs.MoveNext
            Wend
        End If
        
        Call .AutoSize(0)
        
        
        .Redraw = flexRDBuffered
    End With
Exit Sub
eh: Call ErrHandler(SRCFILE & "LoadData")
End Sub

Private Sub lblDebug_Click()
    Dim i As Long
    lblDebug.ForeColor = vbRed
    For i = 0 To gData.Cols - 1
        gData.ColHidden(i) = False
    Next
    Call gData.AutoSize(0, gData.Cols - 1)
    gData.ScrollBars = flexScrollBarBoth
End Sub

