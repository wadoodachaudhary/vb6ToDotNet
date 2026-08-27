VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form FTakeoff 
   Caption         =   "Form1"
   ClientHeight    =   6870
   ClientLeft      =   3015
   ClientTop       =   1995
   ClientWidth     =   11640
   Icon            =   "FTakeoffNEW.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6870
   ScaleWidth      =   11640
   Begin VB.TextBox txtQty 
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   5370
      TabIndex        =   5
      Top             =   6240
      Width           =   555
   End
   Begin VB.TextBox txtHelp 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   3270
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   13
      Top             =   2370
      Width           =   1935
   End
   Begin VB.Frame frmOnScreen 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   315
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   10065
      Begin VB.CheckBox chkCreateAssemblies 
         Alignment       =   1  'Right Justify
         Caption         =   "Create new Assembly for each Takeoff?"
         Height          =   375
         Left            =   6480
         TabIndex        =   17
         Top             =   0
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.TextBox txtOnScreenProject 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   1320
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   9
         Top             =   45
         Width           =   4185
      End
      Begin VB.Image cmdBrowse 
         Height          =   240
         Left            =   5520
         Picture         =   "FTakeoffNEW.frx":000C
         Top             =   45
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "OnCenter Project"
         Height          =   195
         Index           =   2
         Left            =   30
         TabIndex        =   10
         Top             =   60
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdUndoPass 
      Enabled         =   0   'False
      Height          =   315
      Left            =   5220
      Picture         =   "FTakeoffNEW.frx":0156
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Add Pass"
      Top             =   2730
      Width           =   315
   End
   Begin HFEst.Slider Slider 
      Height          =   60
      Index           =   0
      Left            =   -90
      Top             =   3120
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   106
      Orientation     =   1
      Max             =   7335
   End
   Begin VB.CommandButton cmdAddPass 
      Enabled         =   0   'False
      Height          =   345
      Left            =   5220
      Picture         =   "FTakeoffNEW.frx":02A0
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Add Pass"
      Top             =   2400
      Width           =   315
   End
   Begin VB.CommandButton cmdNav 
      Cancel          =   -1  'True
      Caption         =   "C&lose"
      Height          =   315
      Index           =   1
      Left            =   7500
      TabIndex        =   7
      Top             =   6240
      Width           =   1035
   End
   Begin HFEst.Slider Slider 
      Height          =   2640
      Index           =   1
      Left            =   2940
      Top             =   420
      Width           =   60
      _ExtentX        =   106
      _ExtentY        =   4657
      Max             =   3600
   End
   Begin VSFlex8Ctl.VSFlexGrid gItems 
      Height          =   2715
      Left            =   150
      TabIndex        =   4
      Top             =   3300
      Width           =   6195
      _cx             =   10927
      _cy             =   4789
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
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   3
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   36
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FTakeoffNEW.frx":082A
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   6
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   5
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   1
      PicturesOver    =   0   'False
      FillStyle       =   1
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
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   -2147483643
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "&OK"
      Height          =   315
      Index           =   0
      Left            =   6420
      TabIndex        =   6
      Top             =   6240
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VSFlex8Ctl.VSFlexGrid gExcel 
      Height          =   2055
      Left            =   13200
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1110
      Visible         =   0   'False
      Width           =   2895
      _cx             =   5106
      _cy             =   3625
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
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
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
   Begin VSFlex8Ctl.VSFlexGrid gList 
      Height          =   2595
      Left            =   90
      TabIndex        =   1
      Top             =   360
      Width           =   2595
      _cx             =   4577
      _cy             =   4577
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
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   3
      GridLines       =   0
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   2
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FTakeoffNEW.frx":0CFA
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   7
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   2
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   5
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
   Begin VSFlex8Ctl.VSFlexGrid gVariables 
      Height          =   2145
      Left            =   8490
      TabIndex        =   12
      Top             =   720
      Width           =   3045
      _cx             =   5371
      _cy             =   3784
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
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483633
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483633
      FocusRect       =   1
      HighLight       =   0
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   0
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   12
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FTakeoffNEW.frx":0D4C
      ScrollTrack     =   0   'False
      ScrollBars      =   2
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   5
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   2
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
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   -2147483643
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin HFEst.Slider Slider 
      Height          =   60
      Index           =   2
      Left            =   3270
      Top             =   2100
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   106
      Orientation     =   1
      Max             =   7335
   End
   Begin VSFlex8Ctl.VSFlexGrid gHFVariables 
      Height          =   1605
      Left            =   3270
      TabIndex        =   11
      Top             =   360
      Width           =   2505
      _cx             =   4419
      _cy             =   2831
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
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483633
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483633
      FocusRect       =   1
      HighLight       =   0
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   0
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   11
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FTakeoffNEW.frx":0EC5
      ScrollTrack     =   0   'False
      ScrollBars      =   2
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   5
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   2
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   1
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
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   -2147483643
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.ComboBox cboCategory 
      Height          =   240
      Left            =   1770
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   6247
      Width           =   2775
   End
   Begin VB.Label lblCategory 
      Alignment       =   1  'Right Justify
      Caption         =   "Override JC Category"
      Height          =   255
      Left            =   90
      TabIndex        =   16
      Top             =   6240
      Width           =   1575
   End
   Begin VB.Label lblQty 
      AutoSize        =   -1  'True
      Caption         =   "Sale Qty"
      Height          =   195
      Left            =   4680
      TabIndex        =   15
      Top             =   6270
      Width           =   600
   End
End
Attribute VB_Name = "FTakeoff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Private Const SRCFILE = "FTakeoff::"

Private mAssemblyInitialized  As Boolean

Private mCommercial As Boolean

Private mSessionID           As String
Private mFilename            As String

Private mParentForm          As Form
Private mIsChangeOrder       As Boolean
Private mTakeoffMode         As String
Private mAssemblyType        As AssemblyTypes
Private mAssemblyDescription As String
Private mCommunity           As String
Private mCommunityPhase      As String
Private mAssembly            As String
Private mModel               As String
Private mOptionID            As String
Private mJob                 As String
Private mVendor              As String

Private mUseAltCostCodesForCO As Boolean

Private mOnScreenBidID       As Long
Private mOnScreenProjectName As String

Private Const KEYED_SET_DEFAULT_ITERATOR = 1
Private Const OPEN_FLAG_READ_ONLY = 2
Private Const ITERATOR_OK = 1
Dim PECOMServer       As Object 'New PECOMServer
Dim PEAccess          As Object 'PEAccess
Dim PEDatabase        As Object 'PEDatabase
Dim PEEstimate        As Object 'PEEstimate
Dim PEDBAssembly      As Object 'PEDBAssembly
Dim PEAssemblyTakeoff As Object 'PEAssemblyTakeoff

'stupid flags
Dim mSorting As Boolean ' see AddAssembly() and gVariables_RowColChange()

'items menu constants
Private Const mcITEM_COPY = 0
Private Const mcITEM_SUBSTITUE = 1
Private Const mcITEM_REMOVE = 2

Private Type ViewDefs
    Name         As String
    KeyFlds      As String
    SortFlds     As String
    DisplayFlds  As String
    FromWhere    As String
    Database     As Connections
End Type
Private mViews()   As ViewDefs
Private mViewIndex As Long

Private Sub ReplaceItemChart(r As Long)
    Dim s As String
    Dim rs As Recordset
    Dim i As Long
    
    With gItems
    
    
        s = ""
        s = s & "select *" & vbCrLf
        s = s & "  from itemchartdetails" & vbCrLf
        s = s & " where name=" & DbQuote(str, .TextMatrix(r, .ColIndex("itemchart"))) & vbCrLf
        s = s & " and dimension1=" & DbQuote(str, GetVariableValue(.TextMatrix(r, .ColIndex("variable1")))) & vbCrLf
        s = s & " and dimension2=" & DbQuote(str, GetVariableValue(.TextMatrix(r, .ColIndex("variable2")))) & vbCrLf
        s = s & " and dimension3=" & DbQuote(str, GetVariableValue(.TextMatrix(r, .ColIndex("variable3")))) & vbCrLf
        Set rs = HFApp.SqlExec(s, dbHomeFront)
        If rs.EOF Then
            txtHelp.Text = txtHelp.Text & vbCrLf & "Chart """ & .TextMatrix(r, .ColIndex("itemchart")) & """ found no match."
        Else
            i = FindItem("" & rs("phase"), "" & rs("item"), "" & rs("name"))
            If i = -1 Then
                Call AddItem("" & rs("phase"), "" & rs("item"))
                i = .Rows - 1
                .TextMatrix(i, .ColIndex("resolvedformula")) = .TextMatrix(r, .ColIndex("resolvedformula"))
                .TextMatrix(i, .ColIndex("rawformula")) = .TextMatrix(r, .ColIndex("rawformula"))
                .TextMatrix(i, .ColIndex("itemchart")) = .TextMatrix(r, .ColIndex("itemchart"))
            End If
            .TextMatrix(i, .ColIndex("flag")) = 1
        End If
    End With
End Sub

Private Function FindItem(Phase As String, Item As String, itemchart As String) As Long
    Dim i As Long
    With gItems
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("Phase")) = Phase And .TextMatrix(i, .ColIndex("Item")) = Item And .TextMatrix(i, .ColIndex("ItemChart")) = itemchart Then
                FindItem = i
                Exit Function
            End If
        Next
    End With
    FindItem = -1
End Function

Private Function GetVariableValue(Name As String) As String
    Dim i As Long
    With gHFVariables
    For i = 0 To .Rows - 1
        If .TextMatrix(i, .ColIndex("name")) = Name Then
            GetVariableValue = Trim(.TextMatrix(i, .ColIndex("value")))
            Exit Function
        End If
    Next
    End With

End Function
Private Sub cmdAddPass_Click()
On Error GoTo eh

    Dim Qty As Double
    Dim i As Long
    Dim s As String
    Dim f As String
    
    If gHFVariables.Visible Then
        
        If gHFVariables.Rows = 0 Then Exit Sub
        
        Call LookupCommunityStandards
        
        'do homefront pass
        Qty = gHFVariables.TextMatrix(0, gHFVariables.ColIndex("Value"))
        With gItems
        
            'resolve item charts
            txtHelp.Text = ""
            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("flag")) = 0
            Next
            For i = 1 To .Rows - 1
                If .TextMatrix(i, .ColIndex("Phase")) = "" And .TextMatrix(i, .ColIndex("ItemChart")) <> "" Then
                    Call ReplaceItemChart(i)
                End If
            Next
            
            'first pass - resolve and add "normal" items
            For i = 1 To .Rows - 1
                f = .TextMatrix(i, .ColIndex("resolvedformula"))
                If (Not FormulaHasLookups(f)) And Not (.TextMatrix(i, .ColIndex("ItemChart")) <> "" And .TextMatrix(i, .ColIndex("flag")) = 0) Then
    
                    .TextMatrix(i, .ColIndex("p5")) = .TextMatrix(i, .ColIndex("p4"))
                    .TextMatrix(i, .ColIndex("p4")) = .TextMatrix(i, .ColIndex("p3"))
                    .TextMatrix(i, .ColIndex("p3")) = .TextMatrix(i, .ColIndex("p2"))
                    .TextMatrix(i, .ColIndex("p2")) = .TextMatrix(i, .ColIndex("p1"))
                    .TextMatrix(i, .ColIndex("p1")) = .TextMatrix(i, .ColIndex("takeoffqty"))
    
                    On Error Resume Next
                    .TextMatrix(i, .ColIndex("PassTakeoffQty")) = (Qty * CalcFormula(gHFVariables, f, True, Me))
                    .TextMatrix(i, .ColIndex("TakeoffQty")) = Val(.TextMatrix(i, .ColIndex("TakeoffQty"))) + Val(.TextMatrix(i, .ColIndex("PassTakeoffQty")))
                    If Err.Number <> 0 Then
                        MsgBox "Error evaluating row " & i & " - " & .TextMatrix(i, .ColIndex("description")) & vbCrLf & vbCrLf & .TextMatrix(i, .ColIndex("rawformula")), vbExclamation, App.ProductName
                        Exit Sub
                    End If
                    .Row = i
                    Call gItems_AfterEdit(.Row, .ColIndex("takeoffqty"))
                    Call gItems_AfterEdit(.Row, .ColIndex("PassTakeoffQty"))
                    
                End If
            Next
            
            
            'second pass - resolve and add any items that use lookup functions in the formula
            For i = 1 To .Rows - 1
                f = .TextMatrix(i, .ColIndex("resolvedformula"))
                If FormulaHasLookups(f) Then
                    
                    .TextMatrix(i, .ColIndex("p5")) = .TextMatrix(i, .ColIndex("p4"))
                    .TextMatrix(i, .ColIndex("p4")) = .TextMatrix(i, .ColIndex("p3"))
                    .TextMatrix(i, .ColIndex("p3")) = .TextMatrix(i, .ColIndex("p2"))
                    .TextMatrix(i, .ColIndex("p2")) = .TextMatrix(i, .ColIndex("p1"))
                    .TextMatrix(i, .ColIndex("p1")) = .TextMatrix(i, .ColIndex("takeoffqty"))
    
                    On Error Resume Next
                    .TextMatrix(i, .ColIndex("PassTakeoffQty")) = (CalcFormula(gHFVariables, f, True, Me))
                    .TextMatrix(i, .ColIndex("TakeoffQty")) = Val(.TextMatrix(i, .ColIndex("TakeoffQty"))) + Val(.TextMatrix(i, .ColIndex("PassTakeoffQty")))
                    If Err.Number <> 0 Then
                        MsgBox "Error evaluating row " & i & " - " & .TextMatrix(i, .ColIndex("description")) & vbCrLf & vbCrLf & .TextMatrix(i, .ColIndex("rawformula")), vbExclamation, App.ProductName
                        Exit Sub
                    End If
                    .Row = i
                    Call gItems_AfterEdit(.Row, .ColIndex("TakeoffQty"))
                    Call gItems_AfterEdit(.Row, .ColIndex("PassTakeoffQty"))
                End If
            Next
            
            
        End With
    Else
        'do timberline pass
        Qty = gVariables.TextMatrix(0, Me.gVariables.ColIndex("Value"))
        On Error Resume Next
        PEAssemblyTakeoff.AssemblyFactor = 1
        On Error GoTo eh
        PEAssemblyTakeoff.Quantity = gVariables.ValueMatrix(0, 1)
        
        PEAssemblyTakeoff.AddPass
        Call ReadAssemblyItems
        gVariables.Row = 0
        gVariables.Col = 1
        gVariables.SetFocus
    End If

    txtQty.Text = Val(txtQty.Text) + Qty

Exit Sub
eh:
Select Case Err.Number
    Case -2147219163
        s = "Unable to add pass." & vbCrLf
        For i = 0 To PEAssemblyTakeoff.NewErrors.Count - 1
            s = s & PEAssemblyTakeoff.NewErrors.Item(i).Description & vbCrLf
        Next
        MsgBox s, vbExclamation
    Case -2147219444
        MsgBox Err.Description, vbExclamation, "Estimating"
        
    Case Else
        Call errHandler("cmdAddPass_Click")
        
End Select
End Sub


Private Sub cmdUndoPass_Click()
On Error Resume Next
    Dim i As Long
    If gHFVariables.Visible Then
        'do homefront undo
        With gItems
            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("takeoffqty")) = .TextMatrix(i, .ColIndex("prevtakeoffqty"))

                .TextMatrix(i, .ColIndex("takeoffqty")) = .TextMatrix(i, .ColIndex("p1"))
                .TextMatrix(i, .ColIndex("p1")) = .TextMatrix(i, .ColIndex("p2"))
                .TextMatrix(i, .ColIndex("p2")) = .TextMatrix(i, .ColIndex("p3"))
                .TextMatrix(i, .ColIndex("p3")) = .TextMatrix(i, .ColIndex("p4"))
                .TextMatrix(i, .ColIndex("p4")) = .TextMatrix(i, .ColIndex("p5"))

                .Row = i
                Call gItems_AfterEdit(.Row, .ColIndex("takeoffqty"))
            Next
        End With
    Else
        'do timberline undo
        Call PEAssemblyTakeoff.UndoPass
        Call ReadAssemblyItems
    End If

End Sub




Private Sub cmdBrowse_Click()
    Dim s As String

    s = ""
    s = s & "select b.UID,b.BidNo,b.JobName as ProjectName,b.JobID as JobNo,e.firstname + ' ' + e.lastname as Estimator" & vbCrLf
    s = s & "from bids b" & vbCrLf
    s = s & "left outer join Employees e on(b.estimatoruid=e.uid)" & vbCrLf
    If FPickList.Choose(OnScreenConnection, "Project", s, "" & mOnScreenBidID, , , , "UID") Then
        mOnScreenBidID = FPickList.SelectedItem("UID")
        mOnScreenProjectName = FPickList.SelectedItem("ProjectName")
        txtOnScreenProject.Text = mOnScreenProjectName
        Call GetVariableValues
    End If

End Sub

Private Sub cmdNav_Click(Index As Integer)
    Select Case cmdNav(Index).Caption

        Case "&Cancel"
            gItems.Rows = 1
            gVariables.Rows = 0
            gHFVariables.Rows = 0
            txtHelp.Text = ""
            txtQty.Text = ""
            cmdNav(0).Visible = False
            cmdAddPass.Enabled = False
            cmdUndoPass.Enabled = False
            cmdNav(1).Caption = "C&lose"
            gList.SetFocus

        Case "&OK"
            If SaveData Then
                gItems.Rows = 1
                gVariables.Rows = 0
                txtHelp.Text = ""
                txtQty.Text = ""
                gHFVariables.Rows = 0
                cmdNav(0).Visible = False
                cmdAddPass.Enabled = False
                cmdUndoPass.Enabled = False
                cmdNav(1).Caption = "C&lose"
                gList.SetFocus
            End If

        Case "C&lose"
            Unload Me

    End Select

End Sub



Private Sub Form_Load()
    On Error GoTo eh
    Dim s As String

    If mTakeoffMode = "Custom" Then Exit Sub

    LoadedForm = "FTakeoff"
    mUseAltCostCodesForCO = HFApp.Options.ValueByName("UseAltCostCodesForCO") = "True"
    Call LoadWBSDescriptions
    
    Call IniGetForm(Me)
    Call IniGetGrid(Me, gItems)
    Slider(0).Visible = True
    Slider(1).Visible = True
    Slider(2).Visible = True



    txtOnScreenProject.Text = mOnScreenProjectName
    frmOnScreen.Visible = OpenOnScreen()


    gVariables.Rows = 0
    gHFVariables.Rows = 0
    txtHelp.Text = ""
    gList.Rows = 0
    gItems.Rows = 1

    Call LoadCategories
    Call LoadViews


    Screen.MousePointer = vbDefault

Exit Sub
eh: Call PECOMErrHandler(SRCFILE & "Form_Load")
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    LoadedForm = ""
    Call CloseOnScreen
    Call HFApp.SqlExec("DELETE FROM ImportedItems WHERE SessionID=" & DbQuote(str, mSessionID), dbHomeFront)
    Call IniPutForm(Me)
    Call IniPutGrid(Me, gItems)

    If mTakeoffMode = "Item" And IsBetween(mViewIndex, 8, 11) Then
        Call IniPut(AppIni, "FTakeoff", "ItemView", mViewIndex)
    End If
    If mTakeoffMode <> "Item" And (IsBetween(mViewIndex, 0, 3) Or IsBetween(mViewIndex, 6, 7)) Then
        Call IniPut(AppIni, "FTakeoff", "AssemblyView", mViewIndex)
    End If
End Sub


Private Sub gHFVariables_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
On Error GoTo EXITSUB
Static inhere As Boolean

    If inhere Then Exit Sub
    If mSorting Then Exit Sub
        
    inhere = True
    

    Dim lastrow As Long
    lastrow = GridLastVisibleRow(gVariables)
    With gHFVariables


        If NewCol = 0 Then
            Cancel = True
        End If

        If NewCol = 2 Then
            Cancel = Not (NewRow > 0 And .TextMatrix(NewRow, .ColIndex("listofvalues")) = "" And (.TextMatrix(NewRow, NewCol) = "" Or .Cell(flexcpData, NewRow, .ColIndex("UOM")) = ""))
        End If

        .Col = 1

    End With
EXITSUB:
    inhere = False
End Sub

Private Sub gHFVariables_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    Dim s As String
    Dim rs As Recordset

    Dim ConditionName As String
    Dim ConditionType As String
    Dim ConditionUOM As String
    Dim VariableUOM As String


    With gHFVariables
        i = Row
        VariableUOM = .TextMatrix(i, .ColIndex("UOM"))
'        ConditionName = .TextMatrix(i, .ColIndex("ConditionName"))
'        ConditionType = .TextMatrix(i, .ColIndex("ConditionType"))
'        ConditionUOM = .TextMatrix(i, .ColIndex("ConditionUOM"))

        s = ""
        s = s & "select c.name" & vbCrLf
        s = s & "      ,q.description as type" & vbCrLf
        s = s & "      ,round(t.Quantity1 * u.conversion * x.factor,0) as qty" & vbCrLf
        s = s & "      ,x.tounit as uom" & vbCrLf
        s = s & "  from bids b" & vbCrLf
        s = s & "      ,bidconditions c" & vbCrLf
        s = s & "      ,bidtakeofftotals t " & vbCrLf
        s = s & "      ,HF_UnitCodes u" & vbCrLf
        s = s & "      ,HF_QtyCodes  q" & vbCrLf
        s = s & "      ,HF_Conversions x" & vbCrLf
        s = s & " where b.uid=c.biduid" & vbCrLf
        s = s & "   and b.uid=t.biduid" & vbCrLf
        s = s & "   and c.uid=t.bidconditionuid" & vbCrLf
        s = s & "   and c.uom1=u.code" & vbCrLf
        s = s & "   and c.quantity1=q.code" & vbCrLf
        s = s & "   and u.description=x.fromunit" & vbCrLf
        s = s & "   and b.uid=" & DbQuote(Num, mOnScreenBidID) & vbCrLf
'        s = s & "   and c.name=" & DbQuote(Str, ConditionName) & vbCrLf
'        s = s & "   and q.description=" & DbQuote(Str, ConditionType) & vbCrLf
        If VariableUOM <> "" Then s = s & "   and x.tounit=" & DbQuote(str, VariableUOM) & vbCrLf
        s = s & "UNION" & vbCrLf
        s = s & "select c.name" & vbCrLf
        s = s & "      ,q.description as type" & vbCrLf
        s = s & "      ,round(t.Quantity2 * u.conversion * x.factor,0) as qty" & vbCrLf
        s = s & "      ,x.tounit as uom" & vbCrLf
        s = s & "  from bids b" & vbCrLf
        s = s & "      ,bidconditions c" & vbCrLf
        s = s & "      ,bidtakeofftotals t " & vbCrLf
        s = s & "      ,HF_UnitCodes u" & vbCrLf
        s = s & "      ,HF_QtyCodes  q" & vbCrLf
        s = s & "      ,HF_Conversions x" & vbCrLf
        s = s & " where b.uid=c.biduid" & vbCrLf
        s = s & "   and b.uid=t.biduid" & vbCrLf
        s = s & "   and c.uid=t.bidconditionuid" & vbCrLf
        s = s & "   and c.uom2=u.code" & vbCrLf
        s = s & "   and c.quantity2=q.code" & vbCrLf
        s = s & "   and u.description=x.fromunit" & vbCrLf
        s = s & "   and b.uid=" & DbQuote(Num, mOnScreenBidID) & vbCrLf
'        s = s & "   and c.name=" & DbQuote(Str, ConditionName) & vbCrLf
'        s = s & "   and q.description=" & DbQuote(Str, ConditionType) & vbCrLf
        If VariableUOM <> "" Then s = s & "   and x.tounit=" & DbQuote(str, VariableUOM) & vbCrLf
        s = s & "UNION" & vbCrLf
        s = s & "select c.name" & vbCrLf
        s = s & "      ,q.description as type" & vbCrLf
        s = s & "      ,round(t.Quantity3 * u.conversion * x.factor,0) as qty" & vbCrLf
        s = s & "      ,x.tounit as uom" & vbCrLf
        s = s & "  from bids b" & vbCrLf
        s = s & "      ,bidconditions c" & vbCrLf
        s = s & "      ,bidtakeofftotals t " & vbCrLf
        s = s & "      ,HF_UnitCodes u" & vbCrLf
        s = s & "      ,HF_QtyCodes  q" & vbCrLf
        s = s & "      ,HF_Conversions x" & vbCrLf
        s = s & " where b.uid=c.biduid" & vbCrLf
        s = s & "   and b.uid=t.biduid" & vbCrLf
        s = s & "   and c.uid=t.bidconditionuid" & vbCrLf
        s = s & "   and c.uom3=u.code" & vbCrLf
        s = s & "   and c.quantity3=q.code" & vbCrLf
        s = s & "   and u.description=x.fromunit" & vbCrLf
        s = s & "   and b.uid=" & DbQuote(Num, mOnScreenBidID) & vbCrLf
'        s = s & "   and c.name=" & DbQuote(Str, ConditionName) & vbCrLf
'        s = s & "   and q.description=" & DbQuote(Str, ConditionType) & vbCrLf
        If VariableUOM <> "" Then s = s & "   and x.tounit=" & DbQuote(str, VariableUOM) & vbCrLf
        s = s & "UNION" & vbCrLf
        s = s & "select c.name" & vbCrLf
        s = s & "      ,'Height' as type" & vbCrLf
        s = s & "      ,round(c.height * x.factor / iif(b.measurebase=1,b.scalefactor1,1)) as qty" & vbCrLf
        s = s & "      ,x.tounit as uom" & vbCrLf
        s = s & "  from bids b" & vbCrLf
        s = s & "      ,bidconditions c" & vbCrLf
        s = s & "      ,HF_Conversions x" & vbCrLf
        s = s & " where b.uid=c.biduid" & vbCrLf
        s = s & "   and c.height/iif(b.measurebase=1,1,b.scalefactor2)<>0" & vbCrLf
        s = s & "   and iif(b.measurebase=0,'IN','mm')=x.fromunit" & vbCrLf
        s = s & "   and b.uid=" & DbQuote(Num, mOnScreenBidID) & vbCrLf
'        s = s & "   and c.name=" & DbQuote(Str, ConditionName) & vbCrLf
'        s = s & "   and 'Height'=" & DbQuote(Str, ConditionType) & vbCrLf
        If VariableUOM <> "" Then s = s & "   and x.tounit=" & DbQuote(str, VariableUOM) & vbCrLf
        s = s & "UNION" & vbCrLf
        s = s & "select c.name" & vbCrLf
        s = s & "      ,'Thickness' as type" & vbCrLf
        s = s & "      ,round(c.thickness * x.factor / iif(b.measurebase=1,b.scalefactor1,1)) as qty" & vbCrLf
        s = s & "      ,x.tounit as uom" & vbCrLf
        s = s & "  from bids b" & vbCrLf
        s = s & "      ,bidconditions c" & vbCrLf
        s = s & "      ,HF_Conversions x" & vbCrLf
        s = s & " where b.uid=c.biduid" & vbCrLf
        s = s & "   and round(c.thickness/iif(b.measurebase=1,b.scalefactor1,1))<>0" & vbCrLf
        s = s & "   and iif(b.measurebase=0,'IN','mm')=x.fromunit" & vbCrLf
        s = s & "   and b.uid=" & DbQuote(Num, mOnScreenBidID) & vbCrLf
'        s = s & "   and c.name=" & DbQuote(Str, ConditionName) & vbCrLf
'        s = s & "   and 'Thickness'=" & DbQuote(Str, ConditionType) & vbCrLf
        If VariableUOM <> "" Then s = s & "   and x.tounit=" & DbQuote(str, VariableUOM) & vbCrLf
        s = s & "order by 1,2,3" & vbCrLf
        If FPickList.Choose(OnScreenConnection, "Condition", s) Then
            .TextMatrix(Row, .ColIndex("value")) = FPickList.SelectedItem("qty")
            .TextMatrix(Row, .ColIndex("uom")) = FPickList.SelectedItem("uom")
        End If

    End With

End Sub

Private Sub gItems_AfterSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long)
    gItems.ColSel = gItems.Col
End Sub

Private Sub gItems_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    With gItems
    Select Case .ColKey(Col)
        Case "POIndex"
            If FPickList.Choose(HFApp.Databases(dbHomeFront), "POIndex", "select poindex from tblpoindex where DivisionID = " & HFApp.DivisionID) Then
                gItems.Text = FPickList.SelectedItem(1)
            End If
    End Select
    End With
    
End Sub

Private Sub gItems_GotFocus()
    If gItems.Row < 0 Then gItems.Row = 1
    If gItems.Col < 0 Then gItems.Col = 0
End Sub

Private Sub gList_BeforeCollapse(ByVal Row As Long, ByVal State As Integer, Cancel As Boolean)
    Dim r As Long
    Dim levels As Long
    Dim Level As Long

    With gList
        If State = flexOutlineExpanded Then
            .Row = Row
            Call LoadList(False)
        End If
    End With

End Sub
Private Sub gList_DblClick()
    If gList.GetNode.Expanded Then
        Call gList_KeyDown(vbKeyReturn, 0)
    Else
        gList.IsCollapsed(gList.Row) = flexOutlineExpanded
    End If
End Sub

Private Sub gList_GotFocus()
On Error Resume Next
    If gList.Row < 0 Then gList.Row = 0
    If gList.Col < 0 Then gList.Col = 0
End Sub

Private Sub gList_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo eh
    Dim i As Long
    Dim Level As Long
    Dim levels As Long
    Dim s As String

    With gList
        If .Row < 0 Then Exit Sub
        Select Case KeyCode

            Case vbKeyF4
                s = "select i.Phase,p.Description PhaseDesc, i.Item, POIndex, i.Description, OrderUOM from tblPhaseItem i join tblestphases p on (p.Divisionid = i.DivisionID and p.phase = i.phase) where i.DivisionID = " & HFApp.DivisionID & " order by i.Phase,i.item"
                If FPickList.Choose(HFApp.Databases(dbHomeFront), "Item Takeoff", s, , , , , , True, 1) Then
                    For i = 1 To FPickList.SelectedItems
                        Call AddItem(FPickList.SelectedItem("Phase", i), FPickList.SelectedItem("Item", i))
                    Next
                End If
            Case vbKeyF
                If Shift = vbCtrlMask Then
                    s = "select i.Phase,p.Description PhaseDesc, i.Item, POIndex, i.Description, OrderUOM from tblPhaseItem i join tblestphases p on (p.DivisionID = i.DivisionID and p.phase = i.phase) where i.DivisionID = " & HFApp.DivisionID & " order by i.Phase,i.item"
                    If FPickList.Choose(HFApp.Databases(dbHomeFront), "Item Takeoff", s, , , , , , True, 1) Then
                        For i = 1 To FPickList.SelectedItems
                            Call AddItem(FPickList.SelectedItem("Phase", i), FPickList.SelectedItem("Item", i))
                        Next
                    End If
                
                'Call FFind.ShowForm(gList, True)
                End If
            Case vbKeyReturn
                levels = Parse(mViews(mViewIndex).DisplayFlds, , "|")
                If .Rows > 0 Then
                    Level = .RowOutlineLevel(.Row) + 1
                End If
                
                'if not on lowest level then open/close node
                'except if you are on the jobs/quote view. that is the only view that you can dbl click on a group node.
                If Level <> levels And mViewIndex <> 5 Then
                    If gList.IsCollapsed(gList.Row) = flexOutlineCollapsed Then
                        gList.IsCollapsed(gList.Row) = flexOutlineExpanded
                    Else
                        gList.IsCollapsed(gList.Row) = flexOutlineCollapsed
                    End If
                Else

                    Select Case .Cell(flexcpData, .Row, 1)

                        'HomeFront asssembly
                        Case "m.model+'~'+m.assembly", "m.optionid+'~'+m.assembly"
                            txtQty.Text = ""
                            .Cell(flexcpFontBold, 0, 0, .Rows - 1, 0) = False
                            .Cell(flexcpFontBold, .Row, 0) = True
                            Call .AutoSize(0, .Cols - 1)
                            Call AddModelOrOption(.GetNode().Key)
                            SetCtrlFocus gHFVariables

                        'job/quote job
                        Case "job"
                            txtQty.Text = ""
                            .Cell(flexcpFontBold, 0, 0, .Rows - 1, 0) = False
                            .Cell(flexcpFontBold, .Row, 0) = True
                            Call .AutoSize(0, .Cols - 1)
                            Call AddQuote(Level, .Cell(flexcpText, .Row, 1), "", 0)

                        
                        'job/quote
                        Case "changeorder"
                            txtQty.Text = ""
                            .Cell(flexcpFontBold, 0, 0, .Rows - 1, 0) = False
                            .Cell(flexcpFontBold, .Row, 0) = True
                            Call .AutoSize(0, .Cols - 1)
                            Call AddQuote(Level, .Cell(flexcpText, .GetNodeRow(.Row, flexNTParent), 1), .Cell(flexcpText, .Row, 1), 0)

                        
                        'job/quote
                        Case "assembly"
                            txtQty.Text = ""
                            .Cell(flexcpFontBold, 0, 0, .Rows - 1, 0) = False
                            .Cell(flexcpFontBold, .Row, 0) = True
                            Call .AutoSize(0, .Cols - 1)
                            If mViewIndex = 13 Then
                                Call AddQuote(Level, "", "", Val(.Cell(flexcpText, .Row, 1)))
                            Else
                               Call AddQuote(Level, "", "", .Cell(flexcpText, .Row, 1))
                            End If




                        'Timberline Assembly
                        Case "assembly_id"
                            txtQty.Text = ""
                            .Cell(flexcpFontBold, 0, 0, .Rows - 1, 0) = False
                            .Cell(flexcpFontBold, .Row, 0) = True
                            Call .AutoSize(0, .Cols - 1)
                            
                            If Not mAssemblyInitialized Then
                                Call AddAssembly(.Cell(flexcpText, .Row, 1))
                                mAssemblyInitialized = True
                            End If
                            Call AddAssembly(.Cell(flexcpText, .Row, 1))
                            SetCtrlFocus gVariables


                        'MasterBuilder
                        Case "tkflin.recnum"
                            txtQty.Text = ""
                            .Cell(flexcpFontBold, 0, 0, .Rows - 1, 0) = False
                            .Cell(flexcpFontBold, .Row, 0) = True
                            Call .AutoSize(0, .Cols - 1)
                            Call AddMBTakeoff(.GetNode().Key)

                        'HomeFront item
                        Case Else
                            Call AddItem(Parse(.Cell(flexcpText, .Row, 1), 1, "/"), Parse(.Cell(flexcpText, .Row, 1), 2, "/"))
                            SetCtrlFocus gHFVariables

                    End Select


                End If
                KeyCode = 0

            Case vbKeyLeft
                If .GetNodeRow(.Row, flexNTFirstChild) <> -1 And .IsCollapsed(.Row) <> flexOutlineCollapsed Then
                    .IsCollapsed(.Row) = flexOutlineCollapsed
                Else
                    If .GetNodeRow(.Row, flexNTParent) <> -1 Then .Row = .GetNodeRow(.Row, flexNTParent)
                End If

            Case vbKeyRight
                If .GetNodeRow(.Row, flexNTFirstChild) <> -1 Then
                    If .IsCollapsed(.Row) = flexOutlineCollapsed Then
                        .IsCollapsed(.Row) = flexOutlineExpanded
                    Else
                        .Row = .GetNodeRow(.Row, flexNTFirstChild)
                    End If
                End If

        End Select
    End With
Exit Sub
eh: Call errHandler("gList_KeyDown")

End Sub



Private Sub gList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
    
        FMain.mnuTakeOffSub.Item(5).Enabled = HFApp.Databases(dbEstimating).State = adStateOpen
        FMain.mnuTakeOffSub.Item(7).Enabled = HFApp.Databases(dbAccounting).State = adStateOpen And HFApp.Options(AccountingSystem) = asMasterBuilder

        FMain.mnuTakeOffAssemblyView(7).Enabled = False
                
        PopupMenu FMain.mnuTakeoff
    End If
End Sub


Private Sub LoadCategories()
On Error GoTo eh

    Dim rs As Recordset
    With cboCategory
        Set rs = HFApp.SqlExec("SELECT ISNULL(Category,'') + ' - ' + ISNULL(Description,'') FROM StandardCategories where DivisionID = " & HFApp.DivisionID & " ORDER BY Category")
        .Clear
        .AddItem ""
        While Not rs.EOF
            .AddItem "" & rs(0)
            rs.MoveNext
        Wend
    End With

Exit Sub
eh: Call PECOMErrHandler(SRCFILE & "LoadCategories")
End Sub


Private Function SaveData() As Boolean
On Error GoTo eh
    Dim i As Long
    Dim r As Long
    Dim rc As Long
    Dim filenumber As Integer
    Dim bIgnoreZero As Boolean
    Dim s As String
    Dim rs As Recordset

    Dim Assembly As String
    Dim AssemblyDescription As String
    Dim Model As String
    Dim Phase As String
    Dim Item As String
    Dim Job As String
    Dim Description As String
    Dim OrderQty As Double
    Dim OrderUOM As String
    Dim TakeoffQty As Double
    Dim TakeoffUOM As String
    Dim ConversionFactor As Double
    Dim RoundTo As Double
    Dim RoundDir As Long
    Dim WastePercent As Long
    Dim JCExtra  As String
    Dim JCCostCode As String
    Dim JCCostCodeDesc As String
    Dim JCCategory As String
    Dim JCCategoryDesc As String
    Dim Vendor As String
    Dim VendorName As String
    Dim price As Double
    Dim TaxGroup As String
    Dim TaxGroupName As String
    Dim JCTaxRate As Double
    Dim NJCTaxRate As Double
    Dim POIndex As String
    Dim Comments As String
    Dim Formula As String
    Dim Location As String
    Dim LastAssembly As String
    Dim Customer As String, ChangeOrder As String, ChangeStatus As String
    
    Dim w As Long
    Dim WBS(40) As String

    Call LookupCommunityStandards
    
    With gItems
        .Redraw = flexRDNone
        Screen.MousePointer = vbHourglass
        'prompt to include items with qty=0
        For r = 1 To .Rows - 1
            If Not .RowHidden(r) Then
                If Val(.TextMatrix(r, .ColIndex("OrderQty"))) = 0 Then
    
                    Select Case HFApp.Options(ZeroQtyTakeoffs)
                        Case ztAccept
                              bIgnoreZero = False
                              Exit For
    
                        Case ztIgnore
                              bIgnoreZero = True
                              Exit For
    
                        Case ztPrompt
                            rc = MsgBox("You have items with an order quantity of zero." & vbCrLf & "Do you want to add these items anyway?", vbQuestion + vbYesNoCancel, App.ProductName)
                            If rc = vbCancel Then
                                SaveData = False
                                Exit Function
                            Else
                                bIgnoreZero = (rc = vbNo)
                                Exit For
                            End If
                    End Select
                End If
            End If
        Next


        'filenumber = FreeFile
       ' Open App.Path & "\TaxGroupLookup.txt" For Output As #filenumber

        
        For r = 1 To .Rows - 1
            If Not .RowHidden(r) Then
                'this is just to give visual that something is happening
                .Row = r
                Call .ShowCell(r, 0)
                .Refresh
    
                If Val(.TextMatrix(r, .ColIndex("OrderQty"))) <> 0 Or Not bIgnoreZero Then
                    
                    
                    
                    'take items from estimating
                    price = 0
                    If HFApp.Options(UsePricingFromEstimating) Then price = Val(.TextMatrix(r, .ColIndex("Price")))
    
                    Assembly = .TextMatrix(r, .ColIndex("Assembly"))
                    AssemblyDescription = .TextMatrix(r, .ColIndex("AssemblyDescription"))
                    Model = .TextMatrix(r, .ColIndex("Model"))
                    Phase = .TextMatrix(r, .ColIndex("Phase"))
                    Job = .TextMatrix(r, .ColIndex("Job"))
                    Item = .TextMatrix(r, .ColIndex("Item"))
                    Description = .TextMatrix(r, .ColIndex("Description"))
                    OrderQty = Val(.TextMatrix(r, .ColIndex("OrderQty")))
                    OrderUOM = .TextMatrix(r, .ColIndex("OrderUOM"))
                    TakeoffQty = Val(.TextMatrix(r, .ColIndex("TakeoffQty")))
                    TakeoffUOM = .TextMatrix(r, .ColIndex("TakeoffUOM"))
                    ConversionFactor = Val(.TextMatrix(r, .ColIndex("ConversionFactor")))
                    JCExtra = .TextMatrix(r, .ColIndex("JCExtra"))
                    JCCostCode = .TextMatrix(r, .ColIndex("JCCostCode"))
                    JCCategory = .TextMatrix(r, .ColIndex("JCCategory"))
                    POIndex = .TextMatrix(r, .ColIndex("POIndex"))
                    Comments = .TextMatrix(r, .ColIndex("Notes"))
                    Formula = .TextMatrix(r, .ColIndex("RawFormula"))
    
                    Location = .TextMatrix(r, .ColIndex("Location"))
                    
                    For w = 1 To 40
                    WBS(w) = .TextMatrix(r, .ColIndex("WBS" & format(w, "00")))
                    Next
    
                        
                        
                    'override category if selected
                    If cboCategory.ListIndex > 0 Then JCCategory = Parse(cboCategory.Text, 1, " - ")
    
                    'get costcode/category from poindex if not defined on item
                    On Error Resume Next
                    If JCCostCode = "" Then JCCostCode = HFApp.SqlExec("select jccostcode from tblpoindex where DivisionID = " & HFApp.DivisionID & " and poindex = " & DbQuote(str, POIndex))(0)
                    If JCCategory = "" Then JCCategory = HFApp.SqlExec("select JCCategory from tblpoindex where DivisionID = " & HFApp.DivisionID & " and poindex = " & DbQuote(str, POIndex))(0)
                    On Error GoTo eh
    
                    'get costcode/category descriptions
                    s = ""
                    s = s & "SELECT code.Description,cat.Description" & vbCrLf
                    s = s & "  FROM system_setup" & vbCrLf
                    s = s & "       LEFT OUTER JOIN StandardCostCodes code ON(system_setup.ID = code.DivisionID and code.CostCode=" & DbQuote(str, JCCostCode) & ")" & vbCrLf
                    s = s & "       LEFT OUTER JOIN StandardCategories cat ON(system_setup.ID = cat.DivisionID and cat.Category=" & DbQuote(str, JCCategory) & ")" & vbCrLf
                    s = s & " Where ID = " & HFApp.DivisionID
                    Set rs = HFApp.SqlExec(s)
                    If Not rs.EOF Then
                        JCCostCodeDesc = "" & rs(0)
                        JCCategoryDesc = "" & rs(1)
                    End If
    
                    If mParentForm.Name <> "FAssembly" Then
                    
                        If mVendor = "" Then
                            'get default vendor
                            s = ""
                            s = s & "SELECT Vendor_ID,Vendor_Name" & vbCrLf
                            s = s & "  FROM tblVendors" & vbCrLf
                            s = s & " WHERE DivisionID = " & HFApp.DivisionID & " and Vendor_id=dbo.Purch_GetCommunityVendor(" & DbQuote(str, mCommunity) & "," & DbQuote(str, POIndex) & "," & HFApp.DivisionID & ")" & vbCrLf
                            Set rs = HFApp.SqlExec(s)
                            If rs.EOF Then
                                Vendor = ""
                                VendorName = ""
                            Else
                                Vendor = "" & rs(0)
                                VendorName = "" & rs(1)
                            End If
                        Else
                            Vendor = mVendor
                        End If
                        
                        
                        'get vendor rate
                        s = ""
                        s = s & "SELECT dbo.Purch_GetItemRate(0,0,"
                        s = s & DbQuote(str, mCommunity) & ","
                        s = s & DbQuote(str, mCommunityPhase) & ","
                        s = s & DbQuote(str, Assembly) & ","
                        s = s & DbQuote(str, Model) & ","
                        s = s & "'',"
                        s = s & DbQuote(str, Phase) & ","
                        s = s & DbQuote(str, Item) & ","
                        s = s & "0,"
                        s = s & DbQuote(str, Vendor) & ","
                        s = s & "GETDATE()," & HFApp.DivisionID & ")"
                        Set rs = HFApp.SqlExec(s)
                        If Not rs.EOF Then
                            If Val("" & rs(0)) <> 0 Then price = Val("" & rs(0))
                        End If
    
                        'get tax rates
                        s = ""
                        s = s & "SELECT *"
                        s = s & " FROM TaxGroups "
                        s = s & " WHERE DivisionID = " & HFApp.DivisionID & " and TaxGroup=dbo.Purch_GetDefaultTaxGroup("
                        s = s & DbQuote(str, mJob) & ","
                        s = s & DbQuote(str, mCommunity) & ","
                        s = s & DbQuote(str, mCommunityPhase) & ","
                        s = s & DbQuote(str, mModel) & ","
                        s = s & DbQuote(str, mAssembly) & ","
                        s = s & DbQuote(str, Phase) & ","
                        s = s & DbQuote(str, Item) & ","
                        s = s & DbQuote(str, Vendor) & ","
                        s = s & DbQuote(str, JCCategory) & "," & HFApp.DivisionID & ")"
                        
                        Set rs = HFApp.SqlExec(s)
                        TaxGroup = ""
                        TaxGroupName = ""
                        JCTaxRate = 0
                        NJCTaxRate = 0
                        If Not rs.EOF Then
                            TaxGroup = "" & rs("TaxGroup")
                            TaxGroupName = "" & rs("Description")
                            JCTaxRate = "" & rs("JCRate")
                            NJCTaxRate = "" & rs("NJCRate")
                            
                            'Print #filenumber, s & "; TaxGroup = " & TaxGroup
                        End If
    
                    End If
                    Call mParentForm.AddItem(Assembly, AssemblyDescription, Model, Phase, Item, Description, OrderQty, OrderUOM, TakeoffQty, TakeoffUOM, ConversionFactor, RoundTo, RoundDir, WastePercent, JCExtra, JCCostCode, JCCostCodeDesc, JCCategory, JCCategoryDesc, Vendor, VendorName, price, TaxGroup, TaxGroupName, JCTaxRate, NJCTaxRate, POIndex, Comments, Formula, Val(txtQty.Text), Location, WBS, Job)
    
                End If
            End If
        Next
        
        'Close filenumber
        
        .SelectionMode = flexSelectionFree
        .HighLight = flexHighlightWithFocus
        .Redraw = flexRDDirect
        Screen.MousePointer = vbDefault
    End With
    SaveData = True



Exit Function
eh: Call PECOMErrHandler(SRCFILE & "SaveData", s)

End Function































Public Sub Takeoff(ParentForm As Form, _
                   IsChangeOrder As Boolean, _
                   TakeoffMode As String, _
                   OnScreenBidID As Long, _
                   OnScreenProjectName As String, _
                   AssemblyType As AssemblyTypes, _
                   AssemblyDescription As String, _
                   community As String, _
                   CommunityPhase As String, _
                   Model As String, _
                   OptionID As String, _
                   Assembly As String, _
                   Job As String, _
                   Optional Vendor As String)

On Error GoTo eh
'TakeoffMode = "Assembly","Item","Custom"

    Dim i As Integer
    Dim ExtensionIndex   As Long
    Dim EstPrefix As String

    mSessionID = MachineName & App.hInstance


    Set mParentForm = ParentForm

    mTakeoffMode = TakeoffMode
    mAssemblyType = AssemblyType
    mAssemblyDescription = AssemblyDescription
    mCommunity = community
    mCommunityPhase = CommunityPhase
    mAssembly = Assembly
    mModel = Model
    mVendor = Vendor
    mOptionID = OptionID
    mOnScreenBidID = OnScreenBidID
    mOnScreenProjectName = OnScreenProjectName
    mIsChangeOrder = IsChangeOrder
    
    If mJob <> Job Then
        Set TakeoffParameters = New Collection
        mJob = Job
    End If



    Select Case mTakeoffMode

        Case "OneTime"
            Call FTakeoffOneTime.Takeoff(ParentForm, AssemblyDescription, community, CommunityPhase, Model, Assembly, Job)

        Case "Custom"
            ExtensionIndex = Max(1, Val(IniGet(AppIni, "Options", "CustomTakeoffType", 1)))
            If Not VBGetOpenFileName(mFilename, , , , , True, "Assembly Files (*.pee;*.xls;*.txt;*.csv)|*.pee;*.xls;*.txt;*.csv|Timberline Estimates (*.pee)|*.pee|Excel Files (*.xls)|*.xls|Text Files (*.txt;*.csv)|*.txt;*.csv", ExtensionIndex, , "Custom Takeoff - " & mAssemblyDescription, "pee", Screen.ActiveForm.hWnd) Then Exit Sub
            Call IniPut(AppIni, "Options", "CustomTakeoffType", ExtensionIndex)
            Select Case FileExt(mFilename)
                Case "pee"
                    Call ReadEstimate(mFilename)
                Case "xls"
                    mFilename = SaveToCSV(mFilename)
                    Call ReadText(mFilename)
                    On Error Resume Next
                    Kill mFilename
                    On Error GoTo eh
                Case Else
                    Call ReadText(mFilename)
            End Select

        Case "Item"
            Screen.MousePointer = vbHourglass
            mViewIndex = Val(IniGet(AppIni, "FTakeoff", "ItemView", 0))
            If Not IsBetween(mViewIndex, 8, 11) Then mViewIndex = 8
            Me.Show vbModal
            
            
        Case "Assembly"
            Screen.MousePointer = vbHourglass
            mViewIndex = Val(IniGet(AppIni, "FTakeoff", "AssemblyView", 4))
            
            If Not (IsBetween(mViewIndex, 0, 3) Or IsBetween(mViewIndex, 6, 7)) Then mViewIndex = 0
            If IsBetween(mViewIndex, 6, 7) And HFApp.Databases(dbEstimating).State <> adStateOpen Then mViewIndex = 0
            
            
            Me.Show vbModal


    End Select
    Call HFApp.SqlExec("delete from importeditems where sessionid=" & DbQuote(str, mSessionID), dbHomeFront)

    Assembly = mAssembly
    AssemblyType = mAssemblyType
    AssemblyDescription = mAssemblyDescription
    Model = mModel
    OptionID = mOptionID
    OnScreenBidID = mOnScreenBidID
    OnScreenProjectName = mOnScreenProjectName


Exit Sub
eh: Call PECOMErrHandler(SRCFILE & "Takeoff", , mFilename)
End Sub







Private Sub ReadEstimate(FileName As String)
On Error GoTo eh
    Dim Connection As New ADODB.Connection
    Dim s As String
    Dim i As Long
    Dim Categories As String
    Dim c As String
    Dim rs As Recordset

    Screen.MousePointer = vbHourglass

    Call Connection.Open(TsConnectionString(FilePath(FileName), HFApp.Options(EstimatingUser), HFApp.Options(EstimatingPswd), StandardNames))


    s = ""
    s = s & "select i.phase_code               Phase" & vbCrLf
    s = s & "      ,i.item_number              ItemNumber" & vbCrLf
    s = s & "      ,i.category_codes           CostCategories" & vbCrLf
    s = s & "      ,i.item_desc                Description" & vbCrLf
    s = s & "      ,i.Takeoff_Quantity         Qty" & vbCrLf
    s = s & "      ,i.Takeoff_Unit             UOM" & vbCrLf
    s = s & "      ,i.item_note                Comments" & vbCrLf
    s = s & "      ,i.material_price           MPrice" & vbCrLf
    s = s & "      ,i.equipment_price          EPrice" & vbCrLf
    s = s & "      ,i.sub_price                SPrice" & vbCrLf
    s = s & "      ,i.other_price              OPrice" & vbCrLf
    s = s & "      ,i.labor_price              LPrice" & vbCrLf
    s = s & "      ,if(i.item_number='|',i.material_jc_phase,convert('',sql_varchar))        MCode" & vbCrLf
    s = s & "      ,if(i.item_number='|',i.material_jc_category,convert('',sql_varchar))     MCat" & vbCrLf
    s = s & "      ,if(i.item_number='|',i.labor_jc_phase,convert('',sql_varchar))           LCode" & vbCrLf
    s = s & "      ,if(i.item_number='|',i.labor_jc_category,convert('',sql_varchar))        LCat" & vbCrLf
    s = s & "      ,if(i.item_number='|',i.subcontract_jc_phase,convert('',sql_varchar))     SCode" & vbCrLf
    s = s & "      ,if(i.item_number='|',i.subcontract_jc_category,convert('',sql_varchar))  SCat" & vbCrLf
    s = s & "      ,if(i.item_number='|',i.equipment_jc_phase,convert('',sql_varchar))       ECode" & vbCrLf
    s = s & "      ,if(i.item_number='|',i.equipment_jc_category,convert('',sql_varchar))    ECat" & vbCrLf
    s = s & "      ,if(i.item_number='|',i.other_jc_phase,convert('',sql_varchar))           OCode" & vbCrLf
    s = s & "      ,if(i.item_number='|',i.other_jc_category,convert('',sql_varchar))        OCat" & vbCrLf
    s = s & "      ,h.assembly_id" & vbCrLf
    s = s & "  from " & vbQuote & "pee_" & StripExtension(FileTitle(FileName)) & "__estimate_item" & vbQuote & " i" & vbCrLf
    s = s & "       left outer join " & vbQuote & "pee_" & StripExtension(FileTitle(FileName)) & "__estimate_assembly_hdr" & vbQuote & " h on(i.assembly_header=h.header)" & vbCrLf
    s = s & "order by i.group_phase,i.Phase_Code,i.item_number,i.seq_number" & vbCrLf
    Set rs = Connection.Execute(s)
    While Not rs.EOF

        Categories = "" & rs("CostCategories")
        For i = 1 To Len(Categories)
            c = Mid(Categories, i, 1)

            s = ""
            s = s & "INSERT INTO ImportedItems(SessionID,Phase,Item,Description,UOM,Comments,Qty,Rate,assembly,JCCostCode,JCCategory) VALUES"
            s = s & "(" & DbQuote(str, mSessionID)
            s = s & "," & DbQuote(str, "" & rs("Phase"), , True)
            s = s & "," & DbQuote(str, "" & rs("ItemNumber") & c, , True)
            s = s & "," & DbQuote(str, "" & rs("Description"), , True)
            s = s & "," & DbQuote(str, "" & rs("UOM"), , True)
            s = s & "," & DbQuote(str, "" & rs("Comments"), , True)
            s = s & "," & DbQuote(Num, "" & rs("Qty"))
            s = s & "," & DbQuote(Num, "" & rs(c & "Price"))
            s = s & "," & DbQuote(str, "" & rs("Assembly_id"))
            s = s & "," & DbQuote(str, HFApp.FormatCostCode("" & rs(c & "Code")))
            s = s & "," & DbQuote(str, "" & rs(c & "Cat")) & ")"
            Call HFApp.SqlExec(s, dbHomeFront)

        Next
        rs.MoveNext
    Wend
    Connection.Close

    Call ReadImportedItemsBack(False)

    Screen.MousePointer = vbDefault


Exit Sub
eh: Screen.MousePointer = vbDefault
    MsgBox App.ProductName & " is unable to open this file:" & vbCrLf & _
           FileName & vbCrLf & vbCrLf & _
           "Several conditions could cause this." & vbCrLf & _
           "  1.) The file is in use by another user or application." & vbCrLf & _
           "  2.) The file is not a valid Timberline estimate file." & vbCrLf & vbCrLf & vbCrLf & Err.Description, vbExclamation, App.ProductName
End Sub






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
    Dim s As String
    Dim rs As Recordset
    Dim FileHasHeadings As Boolean

    Select Case FileExt(FileName)
        Case "csv":  Call gExcel.LoadGrid(FileName, flexFileCommaText)
        Case Else:   Call gExcel.LoadGrid(FileName, flexFileTabText)
    End Select



    With gExcel

    'set grid layout
    'check if file has headings
    FileHasHeadings = False
    For i = 1 To .Cols - 1
        If IsIn(LCase(.TextMatrix(1, i)), "phase", "item", "description", "qty", "uom") Then
            FileHasHeadings = True
            Exit For
        End If
    Next
    If FileHasHeadings Then
        'use file headings
        For i = 1 To .Cols - 1
            .ColKey(i) = Replace(.TextMatrix(1, i), " ", "")
        Next
        Call .RemoveItem(1)
    Else
        'use default layout
        If .Cols = 4 Then
            For i = 1 To Min(5, .Cols - 1)
                .ColKey(i) = Choose(i, "description", "qty", "uom")
            Next
        Else
            For i = 1 To Min(5, .Cols - 1)
                .ColKey(i) = Choose(i, "phase", "item", "description", "qty", "uom")
            Next
        End If
    End If

    'translate column headings
    If .ColIndex("Desc") <> -1 Then .ColKey(.ColIndex("Desc")) = "Description"
    If .ColIndex("Quantity") <> -1 Then .ColKey(.ColIndex("Quantity")) = "Qty"
    If .ColIndex("Units") <> -1 Then .ColKey(.ColIndex("Units")) = "Qty"
    If .ColIndex("Unit") <> -1 Then .ColKey(.ColIndex("Unit")) = "UOM"
    If .ColIndex("CostCode") <> -1 Then .ColKey(.ColIndex("CostCode")) = "JCCostCode"
    If .ColIndex("Category") <> -1 Then .ColKey(.ColIndex("Category")) = "JCCategory"
    If .ColIndex("Cat") <> -1 Then .ColKey(.ColIndex("Cat")) = "JCCategory"
    If .ColIndex("Price") <> -1 Then .ColKey(.ColIndex("Price")) = "Rate"
    If .ColIndex("Pretax") <> -1 Then .ColKey(.ColIndex("Pretax")) = "Rate"
    If .ColIndex("BOM_Phase") <> -1 Then .ColKey(.ColIndex("BOM_Phase")) = "POIndex"
    If .ColIndex("BOM_Class") <> -1 Then .ColKey(.ColIndex("BOM_Class")) = "POIndex"
    If .ColIndex("BOM") <> -1 Then .ColKey(.ColIndex("BOM")) = "POIndex"

    'check that required cols are available
    If .ColIndex("Description") = -1 Then: Err.Raise 5, , "Required column ""Description"" is missing"
    If .ColIndex("Qty") = -1 Then: Err.Raise 5, , "Required column ""Qty"" is missing"
    If .ColIndex("UOM") = -1 Then: Err.Raise 5, , "Required column ""UOM"" is missing"

    'ensure optional columns are available
    If .ColIndex("Phase") = -1 Then:        .Cols = .Cols + 1: .ColKey(.Cols - 1) = "Phase"
    If .ColIndex("Item") = -1 Then:         .Cols = .Cols + 1: .ColKey(.Cols - 1) = "Item"
    If .ColIndex("POIndex") = -1 Then:      .Cols = .Cols + 1: .ColKey(.Cols - 1) = "POIndex"
    If .ColIndex("JCExtra") = -1 Then:      .Cols = .Cols + 1: .ColKey(.Cols - 1) = "JCExtra"
    If .ColIndex("JCCostCode") = -1 Then:   .Cols = .Cols + 1: .ColKey(.Cols - 1) = "JCCostCode"
    If .ColIndex("JCCategory") = -1 Then:   .Cols = .Cols + 1: .ColKey(.Cols - 1) = "JCCategory"
    If .ColIndex("Rate") = -1 Then:         .Cols = .Cols + 1: .ColKey(.Cols - 1) = "Rate"
    If .ColIndex("Comments") = -1 Then:     .Cols = .Cols + 1: .ColKey(.Cols - 1) = "Comments"
    If .ColIndex("Location") = -1 Then:     .Cols = .Cols + 1: .ColKey(.Cols - 1) = "Location"


    'write data to temp table
    For i = 1 To .Rows - 1
        If Trim(.TextMatrix(i, .ColIndex("qty"))) <> "" Then
            s = ""
            s = s & "INSERT INTO ImportedItems(SessionID,Phase,Item,POIndex,Description,Location,Qty,UOM,JCExtra,JCCostCode,JCCategory,Rate,Comments) VALUES"
            s = s & "(" & DbQuote(str, mSessionID)
            s = s & "," & DbQuote(str, .TextMatrix(i, .ColIndex("Phase")), , True, 20)
            s = s & "," & DbQuote(str, .TextMatrix(i, .ColIndex("Item")), , True, 15)
            s = s & "," & DbQuote(str, .TextMatrix(i, .ColIndex("POIndex")), , True, 20)
            s = s & "," & DbQuote(str, .TextMatrix(i, .ColIndex("Description")), , True, 200)
            s = s & "," & DbQuote(str, .TextMatrix(i, .ColIndex("Location")), , True, 200)
            s = s & "," & DbQuote(Num, .TextMatrix(i, .ColIndex("Qty")))
            s = s & "," & DbQuote(str, .TextMatrix(i, .ColIndex("UOM")), , True, 10)
            s = s & "," & DbQuote(str, .TextMatrix(i, .ColIndex("JCExtra")), , True, 10)
            s = s & "," & DbQuote(str, .TextMatrix(i, .ColIndex("JCCostCode")), , True, 15)
            s = s & "," & DbQuote(str, .TextMatrix(i, .ColIndex("JCCategory")), , True, 3)
            s = s & "," & DbQuote(Num, .TextMatrix(i, .ColIndex("Rate")))
            s = s & "," & DbQuote(str, .TextMatrix(i, .ColIndex("Comments")), , True, 4000) & ")"
            Call HFApp.SqlExec(s, dbHomeFront)
        End If
    Next


    Call ReadImportedItemsBack



    End With
    Unload Me
Exit Sub
eh: MsgBox Err.Description, vbExclamation, App.ProductName
End Sub



Private Sub ReadImportedItemsBack(Optional UpdateDescription As Boolean = True)
    Dim s As String
    Dim rs As Recordset
    Dim bRePrice As Boolean
    Dim WBS(40)
    bRePrice = MsgBox("Do you want to update the item costs in this" & vbCrLf & "estimate using Precision Builders's pricing database?", vbQuestion + vbYesNo, App.ProductName) = vbYes


    'make community standard substitutions
    s = ""
    s = s & "UPDATE ImportedItems" & vbCrLf
    s = s & "SET Phase=s.Phase" & vbCrLf
    s = s & "   ,Item=s.Item" & vbCrLf
    s = s & "   ,comments=s.Notes" & vbCrLf
    s = s & "  FROM ImportedItems i" & vbCrLf
    s = s & "       JOIN CommunityStandards s ON(s.Community='AND' " & vbCrLf
    s = s & "                                AND s.CommunityPhase='1'" & vbCrLf
    s = s & "                                AND i.Phase=s.StdPhase" & vbCrLf
    s = s & "                                AND i.Item=s.StdItem)" & vbCrLf
    Call HFApp.SqlExec(s, dbHomeFront)


    'read data back
    s = ""
    s = s & "SELECT pi.Phase" & vbCrLf
    s = s & "      ,pi.Item" & vbCrLf
    If UpdateDescription Then
        s = s & "      ,ISNULL(pi.Description,ii.Description) Description" & vbCrLf
    Else
        s = s & "      ,ii.Description" & vbCrLf
    End If
    s = s & "      ,case when UPPER(ii.uom)=UPPER(pi.takeoffuom) then " & vbCrLf
    s = s & "            ii.Qty * ISNULL(NULLIF(pi.ConversionFactor,0),1) else " & vbCrLf
    s = s & "            ii.Qty end OrderQty" & vbCrLf
    s = s & "      ,ISNULL(pi.OrderUOM,ii.UOM) OrderUOM" & vbCrLf
    s = s & "      ,case when UPPER(ii.uom)=UPPER(pi.OrderUOM) then " & vbCrLf
    s = s & "            ii.Qty / ISNULL(NULLIF(pi.ConversionFactor,0),1) else " & vbCrLf
    s = s & "            ii.Qty end TakeoffQty" & vbCrLf
    s = s & "      ,ISNULL(pi.TakeoffUOM,ii.UOM) TakeoffUOM" & vbCrLf
    s = s & "      ,ISNULL(pi.ConversionFactor,1) ConversionFactor" & vbCrLf
    s = s & "      ,pi.RoundTo" & vbCrLf
    s = s & "      ,pi.RoundDir" & vbCrLf
    s = s & "      ,pi.WastePercent" & vbCrLf
    s = s & "      ,ii.JCExtra" & vbCrLf
    s = s & "      ,cc.CostCode JCCostCode" & vbCrLf
    s = s & "      ,cc.Description JCCostCodeDesc" & vbCrLf
    s = s & "      ,ct.Category JCCategory" & vbCrLf
    s = s & "      ,ct.Description JCCategoryDesc" & vbCrLf
    s = s & "      ,v.Vendor_ID Vendor" & vbCrLf
    s = s & "      ,v.Vendor_Name VendorName" & vbCrLf
    If bRePrice Then
        s = s & "      ,dbo.Purch_GetItemRate(0,0," & DbQuote(str, mCommunity) & "," & DbQuote(str, mCommunityPhase) & "," & DbQuote(str, mAssembly) & "," & DbQuote(str, mModel) & ",'',pi.Phase,pi.Item,0,v.Vendor_id,getdate()," & HFApp.DivisionID & ") price" & vbCrLf
    Else
        s = s & "      ,ii.Rate Price" & vbCrLf
    End If
    s = s & "      ,tg.TaxGroup" & vbCrLf
    s = s & "      ,tg.Description TaxGroupName" & vbCrLf
    s = s & "      ,tg.JCRate JCTaxRate" & vbCrLf
    s = s & "      ,tg.NJCRate NJCTaxRate" & vbCrLf
    s = s & "      ,ii.Comments" & vbCrLf
    s = s & "      ,po.POIndex" & vbCrLf
    s = s & "      ,pi.Formula" & vbCrLf
    s = s & "      ,ISNULL(ii.assembly," & DbQuote(str, mAssembly) & ") Assembly" & vbCrLf
    s = s & "      ,ISNULL(ii.assemblydesc," & DbQuote(str, mAssemblyDescription) & ") AssemblyDesc" & vbCrLf
    s = s & "      ,ii.Location" & vbCrLf
    s = s & "FROM ImportedItems ii" & vbCrLf
    'if import file does not contain a type designator (LMSEO) on end of item number then add an M
    s = s & "     LEFT OUTER JOIN tblPhaseItem pi ON(pi.Divisionid = " & HFApp.DivisionID & " and ii.Phase=pi.Phase AND ii.Item+case when right(rtrim(ii.Item),1) like('[^LMSEO]') then 'M' else '' end=pi.Item)" & vbCrLf
    s = s & "     LEFT OUTER JOIN tblPOIndex po ON(po.DivisionID = " & HFApp.DivisionID & " and po.poindex=isnull(nullif(ii.poindex,''),pi.poindex))" & vbCrLf
    If mVendor = "" Then
        s = s & "     LEFT OUTER JOIN tblVendors v ON(v.DivisionID = " & HFApp.DivisionID & " and v.Vendor_ID=dbo.Purch_GetCommunityVendor(" & DbQuote(str, mCommunity) & ",po.POIndex," & HFApp.DivisionID & "))" & vbCrLf
    Else
        s = s & "     LEFT OUTER JOIN tblVendors v ON(v.DivisionID = " & HFApp.DivisionID & " and v.Vendor_ID=" & DbQuote(str, mVendor) & ")" & vbCrLf
    End If
    s = s & "     LEFT OUTER JOIN StandardCostCodes  cc ON(cc.Divisionid = " & HFApp.DivisionID & " and REPLACE(cc.CostCode,'-','')=CASE WHEN ii.JCCostCode<>'' THEN REPLACE(ii.JCCostCode,'-','') WHEN " & DbQuote(Num, mAssemblyType) & "<>0 and isnull(pi.AltJCCostCode,'')<>'' THEN REPLACE(pi.AltJCCostCode,'-','') WHEN pi.JCCostCode<>'' THEN REPLACE(pi.JCCostCode,'-','') ELSE REPLACE(po.JCCostCode,'-','') END)" & vbCrLf
    s = s & "     LEFT OUTER JOIN StandardCategories ct ON(ct.Divisionid = " & HFApp.DivisionID & " and ct.Category=CASE WHEN isnull(ii.JCCategory,'')<>'' THEN ii.JCCategory WHEN " & DbQuote(Num, mAssemblyType) & "<>0 and isnull(pi.AltJCCategory,'')<>'' THEN REPLACE(pi.AltJCCategory,'-','') WHEN " & DbQuote(Num, mAssemblyType) & "<>0 THEN REPLACE(pi.JCCategory,'-','') WHEN isnull(pi.JCCategory,'')<>'' THEN pi.JCCategory ELSE po.JCCategory END)" & vbCrLf
    s = s & "     LEFT OUTER JOIN TaxGroups tg ON(tg.Divisionid = " & HFApp.DivisionID & " and tg.TaxGroup=dbo.Purch_GetDefaultTaxGroup(" & DbQuote(str, mJob) & "," & DbQuote(str, mCommunity) & "," & DbQuote(str, mCommunityPhase) & "," & DbQuote(str, mModel) & "," & DbQuote(str, mAssembly) & ",ii.Phase,ii.Item,v.Vendor_ID,ct.Category," & HFApp.DivisionID & "))" & vbCrLf
    s = s & "WHERE SessionID=" & DbQuote(str, mSessionID)
    Set rs = HFApp.SqlExec(s, dbHomeFront)
    While Not rs.EOF

        Call mParentForm.AddItem("" & rs("Assembly"), _
                                 "" & rs("AssemblyDesc"), _
                                 mModel, _
                                 "" & rs("Phase"), _
                                 "" & rs("Item"), _
                                 "" & rs("Description"), _
                                 Val("" & rs("OrderQty")), _
                                 "" & rs("OrderUOM"), _
                                 Val("" & "" & rs("TakeoffQty")), _
                                 "" & rs("TakeoffUOM"), _
                                 Val("" & "" & rs("ConversionFactor")), Val("" & "" & rs("RoundTo")), Val("" & "" & rs("RoundDir")), Val("" & "" & rs("WastePercent")), _
                                 "" & rs("JCExtra"), _
                                 "" & rs("JCCostCode"), _
                                 "" & rs("JCCostCodeDesc"), _
                                 "" & rs("JCCategory"), _
                                 "" & rs("JCCategoryDesc"), _
                                 "" & rs("Vendor"), _
                                 "" & rs("VendorName"), _
                                 Val("" & "" & rs("Price")), _
                                 "" & rs("TaxGroup"), _
                                 "" & rs("TaxGroupName"), _
                                 Val("" & "" & rs("JCTaxRate")), _
                                 Val("" & "" & rs("NJCTaxRate")), _
                                 "" & rs("POIndex"), _
                                 "" & rs("Comments"), "" & rs("Formula"), 1, "" & rs("Location"), WBS)
        rs.MoveNext
    Wend
Exit Sub

End Sub





Private Sub LoadViews()
On Error GoTo eh
    Dim i As Long
    ReDim mViews(13) As ViewDefs

    mCommercial = HFApp.Options.ValueByName("BuilderType") = "Commercial"

    '-------------------------------------------------------------------
    '  HF Assemblies
    '-------------------------------------------------------------------
    If Not mCommercial Then
        FMain.mnuTakeOffModelsView.Item(0).Caption = "Models"
        FMain.mnuTakeOffModelsView.Item(1).Caption = "Models Specific Options"
        FMain.mnuTakeOffModelsView.Item(2).Caption = "Global Options"
        FMain.mnuTakeOffModelsView.Item(3).Caption = "Design Center Options"
    End If
    mViews(i).Name = "Assemblies"
    mViews(i).KeyFlds = "m.community|m.model+'~'+m.assembly"
    mViews(i).DisplayFlds = "isnull(nullif(m.community+isnull(' - '+l.description,''),''),' -- Non-regional  -- ')|m.Assembly + isnull(' - '+m.description,'')"
    mViews(i).FromWhere = " from tbldbassemblymaster m left outer join tbllocality l on(m.community=l.area) where m.DivisionID = " & HFApp.DivisionID & " and isnull(m.inactive,0)=0 and m.assemblytype=0"
    mViews(i).Database = dbHomeFront
    i = i + 1
    
    mViews(i).Name = "Assemblies Specific Extras"
    mViews(i).KeyFlds = "m.community|m.model|m.optionid+'~'+m.assembly"
    mViews(i).DisplayFlds = "isnull(nullif(m.community+isnull(' - '+l.description,''),''),' -- Non-regional -- ')|m.Model + isnull(' - ' +em.description,'')|m.Assembly + isnull(' - '+m.description,'')"
    mViews(i).FromWhere = " from tbldbassemblymaster m left outer join tbllocality l on(m.community=l.area) left outer join estmodellist em on(m.model=em.model) where m.DivisionID = " & HFApp.DivisionID & " and isnull(m.inactive,0)=0 and m.assemblytype=2"
    mViews(i).Database = dbHomeFront
    i = i + 1

    mViews(i).Name = "Global Extras"
    mViews(i).KeyFlds = "m.community|m.category|m.optionid+'~'+m.assembly"
    mViews(i).DisplayFlds = "isnull(nullif(m.community+isnull(' - '+l.description,''),''),' -- Non-regional -- ')|m.category+isnull(' - '+c.description,'')|m.Assembly + isnull(' - '+m.description,'')"
    mViews(i).FromWhere = " from tbldbassemblymaster m left outer join tbllocality l on(m.community=l.area) left outer join tblcategories c on(m.category=c.category) where m.DivisionID = " & HFApp.DivisionID & " and isnull(m.inactive,0)=0 and m.assemblytype=3"
    mViews(i).Database = dbHomeFront
    i = i + 1

    mViews(i).Name = "Design Center Options"
    mViews(i).KeyFlds = "m.community|m.category|m.optionid+'~'+m.assembly"
    mViews(i).DisplayFlds = "isnull(nullif(m.community+isnull(' - '+l.description,''),''),' -- Non-regional Options -- ')|m.category+isnull(' - '+c.description,'')|m.Assembly + isnull(' - '+m.description,'')"
    mViews(i).FromWhere = " from tbldbassemblymaster m left outer join tbllocality l on(m.community=l.area) left outer join tblcategories c on(m.category=c.category) where m.DivisionID = " & HFApp.DivisionID & " and isnull(m.inactive,0)=0 and m.assemblytype=4"
    mViews(i).Database = dbHomeFront
    i = i + 1



    '-------------------------------------------------------------------
    '  Quotes and Jobs
    '-------------------------------------------------------------------

    mViews(i).Name = "Quotes" ' by number
    mViews(i).KeyFlds = "job|changeorder|assembly"
    mViews(i).DisplayFlds = "job + ' - ' + jobdesc|changeorderdesc|assemblydesc"
    mViews(i).SortFlds = "||assemblytype"
    mViews(i).FromWhere = " from OpenJobAndQuoteAssemblies where OpenJobAndQuoteAssemblies.DivisionID = " & HFApp.DivisionID & " and isquote=1"
    mViews(i).Database = dbHomeFront
    i = i + 1

    mViews(i).Name = "Open Jobs" ' by number
    mViews(i).KeyFlds = "job|changeorder|assembly"
    mViews(i).DisplayFlds = "job + ' - ' + jobdesc|changeorderdesc|assemblydesc"
    mViews(i).SortFlds = "||assemblytype"
    mViews(i).FromWhere = " from OpenJobAndQuoteAssemblies where OpenJobAndQuoteAssemblies.DivisionID = " & HFApp.DivisionID & " and isquote=0"
    mViews(i).Database = dbHomeFront
    i = i + 1






    '-------------------------------------------------------------------
    '  Estimating Assemblies
    '-------------------------------------------------------------------

    mViews(i).Name = "Estimating Assemblies"
    mViews(i).KeyFlds = "assembly_id|assembly_id"
    mViews(i).DisplayFlds = "assembly_id + '   ' + assembly_description|assembly_id + '   ' + assembly_description"
    mViews(i).FromWhere = "  from dat_pwa__db_assembly_header where assembly_level=1"
    mViews(i).Database = dbEstimating
    i = i + 1

    mViews(i).Name = "Estimating Assemblies"
    mViews(i).KeyFlds = "left(assembly_description,1)|assembly_id"
    mViews(i).DisplayFlds = "left(assembly_description,1)|assembly_description + '   ' + assembly_id"
    mViews(i).FromWhere = "  from dat_pwa__db_assembly_header where assembly_level=0"
    mViews(i).Database = dbEstimating
    i = i + 1


    '-------------------------------------------------------------------
    '  items
    '-------------------------------------------------------------------

    mViews(i).Name = "Items by Groups & Phases"
    mViews(i).KeyFlds = "GrpPhase|Phase|Phase+'/'+Item"
    mViews(i).SortFlds = "GrpSortOrder|PhaseSortOrder|ItemSortOrder"
    mViews(i).DisplayFlds = "GrpPhase + '   ' + GrpDesc|Phase + '   ' + PhaseDesc|Item + '   ' + ItemDesc"
    mViews(i).FromWhere = "FROM EstimatingItems WHERE DivisionID = " & HFApp.DivisionID & " and 1=1"
    i = i + 1

    mViews(i).Name = "Items by Purchase Orders"
    mViews(i).KeyFlds = "POIndex|Phase+'/'+item"
    mViews(i).DisplayFlds = "POIndex + '   ' + POIndexDescription|ItemDesc"
    mViews(i).FromWhere = "FROM EstimatingItems WHERE DivisionID = " & HFApp.DivisionID & " and 1=1"
    i = i + 1

    mViews(i).Name = "Items by Cost Codes"
    mViews(i).KeyFlds = "JCCostCode|Phase+'/'+item"
    mViews(i).DisplayFlds = "JCCostCode + '   ' + JCCostCodeDesc|ItemDesc"
    mViews(i).FromWhere = "FROM EstimatingItems WHERE DivisionID = " & HFApp.DivisionID & " and 1=1"
    i = i + 1

    mViews(i).Name = "Items"
    mViews(i).KeyFlds = "left(itemdesc,1)|Phase+'/'+item"
    mViews(i).DisplayFlds = "left(itemdesc,1)|ItemDesc"
    mViews(i).FromWhere = "FROM EstimatingItems WHERE DivisionID = " & HFApp.DivisionID & " and 1=1"
    i = i + 1


    '-------------------------------------------------------------------
    '  Master Builder Takeoffs
    '-------------------------------------------------------------------

    mViews(i).Name = "Takeoffs by Job"
    mViews(i).KeyFlds = "tkflin.recnum"
    mViews(i).DisplayFlds = "ltrim(str(tkflin.recnum)) + '   ' + actrec.jobnme"
    mViews(i).FromWhere = "  from tkflin,actrec where tkflin.recnum=actrec.recnum and actrec.status<>6"
    mViews(i).Database = dbAccounting
    i = i + 1
    
    mViews(i).Name = "Custom Options" ' by number
    mViews(i).KeyFlds = "job|changeorder|assembly"
    mViews(i).DisplayFlds = "job + ' - ' + jobdesc|changeorderdesc|assemblydesc"
    mViews(i).SortFlds = "||assemblytype"
    mViews(i).FromWhere = " from dbo.UnApprovedCustomOptions where isquote=0"
    mViews(i).Database = dbHomeFront
    
    Call LoadList(True)
Exit Sub
eh:
If Err.Number <> 0 Then
    Call errHandler("Load Views", "")
    
End If
End Sub

Public Sub mnuTakeOffSub_Click(Index As Integer)
    Select Case Index
        Case 0: Call gList_KeyDown(vbKeyF, vbCtrlMask)
    End Select
End Sub
Public Sub mnuTakeoffCustomOptions_Click()
    mViewIndex = 13
    Call LoadList(True)
End Sub


Private Sub LoadList(ClearTree As Boolean)
On Error GoTo eh
    Dim rs As Recordset
    Dim r As Long
    Dim s       As String
    Dim i       As Long
    Dim levels  As Long
    Dim Level   As Long
    Dim n       As VSFlexNode
    Dim WhereClause As String
    Dim KeyFld     As String
    Dim SortFld     As String
    Dim DisplayFld As String
    Dim KeyValue   As String
    Dim DisplayValue   As String
    Dim done As Boolean


Dim t As Single
t = Timer

'PROPERTY                  CONTAINS
'.rowdata()                Where clause
'.cell(flexcpText, r, 0)   DisplayFld Value
'.cell(flexcpData, r, 0)   DisplayFld Field Name (also used as tooltip)
'.Cell(flexcpText, r, 1)   KeyFld Value
'.Cell(flexcpData, r, 1)   KeyFld Field Name

'Reserved for commercial on the fly assembly creation.
'If mViewIndex < 7 Then
'    chkCreateAssemblies.Visible = True
'End If

If IsIn(mViewIndex, 6, 7) Then Call OpenTLEstimating


    Me.Caption = mViews(mViewIndex).Name
    gVariables.Rows = 0
    gHFVariables.Rows = 0
    txtHelp.Text = ""

    With gList
        .Redraw = flexRDNone
        levels = Parse(mViews(mViewIndex).DisplayFlds, , "|")
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

            KeyFld = Parse(mViews(mViewIndex).KeyFlds, Level + 2, "|")
            SortFld = Parse(mViews(mViewIndex).SortFlds, Level + 2, "|")
            If SortFld = "" Then SortFld = 2
            DisplayFld = Parse(mViews(mViewIndex).DisplayFlds, Level + 2, "|")

            s = ""
            s = s & "SELECT DISTINCT " & KeyFld & "," & DisplayFld & "," & SortFld & vbCrLf
            s = s & mViews(mViewIndex).FromWhere & vbCrLf
            On Error Resume Next
            WhereClause = .RowData(.Row)
            On Error GoTo eh

            If WhereClause <> "" Then s = s & WhereClause & vbCrLf

            
If (mViewIndex = 4 Or mViewIndex = 5) And Level = 1 Then
    s = s & " and rowtype='Assembly'"
End If
    
            If Left(WhereClause, 26) = " AND max(g.assembly_id) = " Then
                s = s & "HAVING " & Mid(WhereClause, 6) & vbCrLf
            End If
            s = s & " ORDER BY " & SortFld & IIf(SortFld = "2", "", ",2")
            s = Replace(s, "|", ",")

            If mViewIndex = 6 Then
            
                Call LoadTLList(Level)
            Else

'MsgBox s
                Set rs = HFApp.SqlExec(s, mViews(mViewIndex).Database)
                While Not rs.EOF
                    KeyValue = "" & rs(0)
                    DisplayValue = Trim("" & rs(1))

                    If Level = -1 Then
                        r = .Rows
                        Call .AddItem(DisplayValue, r)
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
                        If KeyFld = "tkflin.recnum" Then
                            .RowData(r) = WhereClause & " AND " & KeyFld & " = " & DbQuote(Num, KeyValue)
                        Else
                            .RowData(r) = WhereClause & " AND " & KeyFld & " = " & DbQuote(str, KeyValue)
                        End If
                    Else
                        r = .Row
                        Set n = .GetNode(r).AddNode(flexNTLastChild, DisplayValue)
                        .Cell(flexcpData, n.Row, 0) = DisplayFld
                        .Cell(flexcpText, n.Row, 1) = KeyValue
                        .Cell(flexcpData, n.Row, 1) = KeyFld
                        If Level + 2 < levels Then
                            Call .GetNode(n.Row).AddNode(flexNTFirstChild, "dummy") ' add dummy child row so grid knows this one can be opened
                        End If
                        .Row = r
                        n.Expanded = False
                        If KeyFld = "tkflin.recnum" Then
                            .RowData(n.Row) = WhereClause & " AND " & KeyFld & " = " & DbQuote(Num, KeyValue)
                        Else
                            .RowData(n.Row) = WhereClause & " AND " & KeyFld & " = " & DbQuote(str, KeyValue)
                        End If
                    End If
                    rs.MoveNext
                Wend
            End If
        End If
        Call .AutoSize(0)

        On Error Resume Next
        If ClearTree Then Call .Select(0, 0)

        .Redraw = flexRDBuffered
    End With
Exit Sub
eh: Call errHandler("LoadList", s)
End Sub


Private Sub LoadTLList(Level As Long)
    Dim GrpAssemblies As Object 'PEDBGroupAssemblyIterator
    Dim GrpAssembly   As Object 'PEDBGroupAssembly
    Dim Assemblies As Object 'PEDBAssemblyIterator
    Dim Assembly   As Object 'PEDBAssembly
    Dim n As VSFlexNode
    Dim KeyValue As String
    Dim DisplayValue As String
    Dim r As Long
    
Dim m As String
Dim t As Single
t = Timer

    With gList
        If Level = -1 Then
            Set GrpAssemblies = PEDatabase.DBGroupAssemblies.NewIterator(KEYED_SET_DEFAULT_ITERATOR)
m = m & "create group iterator:  " & format(Timer - t, "0.000") & vbCrLf
t = Timer
            While GrpAssemblies.Next = ITERATOR_OK
                Set GrpAssembly = GrpAssemblies.object
                
                KeyValue = GrpAssembly.code
                DisplayValue = Trim(GrpAssembly.code) & "   " & Trim(GrpAssembly.Description)
                
                .AddItem ""
                r = .Rows - 1
                .IsSubtotal(r) = True
                .RowOutlineLevel(r) = Level + 1
                .Cell(flexcpText, r, 0) = DisplayValue
                .Cell(flexcpData, r, 0) = "assembly_id + '   ' + assembly_description"
                .Cell(flexcpText, r, 1) = KeyValue
                .Cell(flexcpData, r, 1) = "assembly_id"
                .RowData(r) = "assembly_id = " & DbQuote(str, KeyValue)
            
                'add dummy child row so grid knows this one can be opened
                Call .GetNode(r).AddNode(flexNTFirstChild, "dummy")
                .GetNode(r).Expanded = False
            Wend
m = m & "iterate thru groups:  " & format(Timer - t, "0.000") & vbCrLf
t = Timer
        Else
            Set GrpAssemblies = PEDatabase.DBGroupAssemblies.NewIterator(KEYED_SET_DEFAULT_ITERATOR)
m = m & "create group iterator:  " & format(Timer - t, "0.000") & vbCrLf
t = Timer
            Set GrpAssembly = PEDatabase.DBGroupAssemblies.NewObject
            GrpAssembly.code = .Cell(flexcpText, .Row, 1)
            Call GrpAssemblies.EQ(GrpAssembly)
            Set GrpAssembly = GrpAssemblies.Buffer
m = m & "get selected group:  " & format(Timer - t, "0.000") & vbCrLf
t = Timer
            
            
            Set Assemblies = GrpAssembly.Assemblies.NewIterator(KEYED_SET_DEFAULT_ITERATOR)
m = m & "create assm iterator:  " & format(Timer - t, "0.000") & vbCrLf
t = Timer
            While Assemblies.Next = ITERATOR_OK
            
                Set Assembly = Assemblies.object
                
                KeyValue = Assembly.code
                DisplayValue = Trim(Assembly.code) & "   " & Trim(Assembly.Description)
                
                
                Set n = .GetNode(.Row).AddNode(flexNTLastChild, DisplayValue)
                
                .Cell(flexcpData, n.Row, 0) = "assembly_id + '   ' + assembly_description"
                .Cell(flexcpText, n.Row, 1) = KeyValue
                .Cell(flexcpData, n.Row, 1) = "assembly_id"
                .RowData(n.Row) = "assembly_id = " & DbQuote(str, KeyValue)
                
            Wend
m = m & "iterate thru assm:  " & format(Timer - t, "0.000") & vbCrLf
t = Timer
        End If
    End With
    
'MsgBox m
    
End Sub




Private Sub AddModelOrOption(Key As String)
    Dim s As String
    Dim f As String
    Dim rs As Recordset
    Dim r As Long
    Dim w As Long
    
    Dim variablelist As String


    gVariables.Visible = False
    gHFVariables.Visible = True
    gHFVariables.Rows = 0
    txtHelp.Text = ""
    variablelist = "Quantity"

    cmdAddPass.Enabled = True
    cmdUndoPass.Enabled = True
    cmdNav(0).Visible = True
    cmdNav(0).Enabled = True
    cmdNav(1).Caption = "&Cancel"

    s = ""
    s = s & "select m.assembly,m.description assemblydescription,m.assemblytype,m.model,m.optionid,i.phase,i.item,d.itemchart,i.description,d.formula,d.TakeoffQty,i.takeoffuom,i.conversionfactor,i.orderuom,isnull(nullif(i.JCCostCode,''),isnull(p.JCCostCode,'')) jccostcode,i.jccategory,i.altjccostcode,i.altjccategory,i.price,i.roundto,i.rounddir,isnull(nullif(d.POIndex,''),isnull(i.POIndex,'')) poindex,c.variable1,c.variable2,c.variable3,i.wastepercent" & vbCrLf
    s = s & ",isnull(nullif(d.notes,''),i.notes) Notes" & vbCrLf
    s = s & ",isnull(nullif(d.location,''),i.location) Location" & vbCrLf
    For w = 1 To 40
    s = s & ",isnull(nullif(d.WBS" & format(w, "00") & ",''),i.WBS" & format(w, "00") & ") WBS" & format(w, "00") & vbCrLf
    Next
    s = s & " from tbldbassemblymaster m " & vbCrLf
    s = s & " left outer join tbldbassemblydetails d on(m.DivisionID = d.DivisionID and m.community=d.community and m.model=d.model and m.optionid=d.optionid and m.assembly=d.assembly)" & vbCrLf
    s = s & " left outer join CommunityStandards s ON(s.Community=" & DbQuote(str, mCommunity) & " AND s.CommunityPhase=dbo.Purch_GetCommunityStandardPhase(" & DbQuote(str, mCommunity) & "," & DbQuote(str, mCommunityPhase) & ", d.phase , d.Item) AND s.StdPhase=d.Phase AND s.StdItem=d.Item)" & vbCrLf
    s = s & " left outer join tblphaseitem i on(i.DivisionID = d.DivisionID and i.phase=isnull(s.phase,d.phase) and i.item=isnull(s.item,d.item))" & vbCrLf
    s = s & " LEFT OUTER JOIN tblPOIndex p ON(p.DivisionID = d.DivisionID and isnull(nullif(d.POIndex,''),isnull(i.POIndex,''))=p.POIndex)" & vbCrLf
    s = s & " left outer join itemcharts c on(d.itemchart=c.name)" & vbCrLf
    s = s & "where " & Mid(Key, 6) & vbCrLf
    s = s & "and m.DivisionID = " & HFApp.DivisionID
    Set rs = HFApp.SqlExec(s, dbHomeFront)
    With gItems
    .Rows = 1
    While Not rs.EOF

        If Trim(mAssembly) = "" Then
            mAssemblyType = "" & rs("AssemblyType")
            mAssembly = "" & rs("Assembly")
            mAssemblyDescription = "" & rs("AssemblyDescription")
            mModel = "" & rs("Model")
            mOptionID = "" & rs("OptionID")
        End If
        
        
        .AddItem ""
        r = .Rows - 1
        .TextMatrix(r, .ColIndex("Assembly")) = "" & rs("Assembly")
        .TextMatrix(r, .ColIndex("AssemblyDescription")) = "" & rs("AssemblyDescription")
        .TextMatrix(r, .ColIndex("Model")) = "" & rs("Model")
        .TextMatrix(r, .ColIndex("Phase")) = "" & rs("Phase")
        .TextMatrix(r, .ColIndex("Item")) = "" & rs("Item")
        .TextMatrix(r, .ColIndex("Description")) = "" & rs("Description")
        .TextMatrix(r, .ColIndex("Job")) = "" & mJob
        .TextMatrix(r, .ColIndex("WastePercent")) = Val("" & rs("WastePercent"))

        If "" & rs("ItemChart") <> "" Then
            .TextMatrix(r, .ColIndex("itemchart")) = "" & rs("ItemChart")
            .TextMatrix(r, .ColIndex("Description")) = "" & rs("ItemChart")
            .TextMatrix(r, .ColIndex("variable1")) = "" & rs("variable1")
            .TextMatrix(r, .ColIndex("variable2")) = "" & rs("variable2")
            .TextMatrix(r, .ColIndex("variable3")) = "" & rs("variable3")
            If "" & rs("variable1") <> "" Then variablelist = variablelist & Chr(1) & rs("variable1")
            If "" & rs("variable2") <> "" Then variablelist = variablelist & Chr(1) & rs("variable2")
            If "" & rs("variable3") <> "" Then variablelist = variablelist & Chr(1) & rs("variable3")
            .RowHidden(r) = True
        End If

        .TextMatrix(r, .ColIndex("rawformula")) = "" & rs("formula")
        .TextMatrix(r, .ColIndex("resolvedformula")) = ResolveFormula(.TextMatrix(r, .ColIndex("rawformula")))
        
        variablelist = variablelist & Chr(1) & ParseVariables(.TextMatrix(r, .ColIndex("resolvedformula")))
        
        If .TextMatrix(r, .ColIndex("resolvedformula")) = "" Then
            .TextMatrix(r, .ColIndex("resolvedformula")) = "" & rs("TakeoffQty")
        End If


        .TextMatrix(r, .ColIndex("Location")) = "" & rs("Location")
        For w = 1 To 40
        .TextMatrix(r, .ColIndex("WBS" & format(w, "00"))) = "" & rs("WBS" & format(w, "00"))
        Next

        .TextMatrix(r, .ColIndex("TakeoffUOM")) = "" & rs("TakeoffUOM")
        .TextMatrix(r, .ColIndex("TakeoffQty")) = 0
        .TextMatrix(r, .ColIndex("p1")) = 0
        .TextMatrix(r, .ColIndex("p2")) = 0
        .TextMatrix(r, .ColIndex("p3")) = 0
        .TextMatrix(r, .ColIndex("p4")) = 0
        .TextMatrix(r, .ColIndex("p5")) = 0
        .TextMatrix(r, .ColIndex("OrderQty")) = 0
        .TextMatrix(r, .ColIndex("OrderUOM")) = "" & rs("OrderUOM")
        .TextMatrix(r, .ColIndex("ConversionFactor")) = "" & rs("ConversionFactor")
        If mAssemblyType = atModel Or (mUseAltCostCodesForCO And Not mIsChangeOrder) Then
            .TextMatrix(r, .ColIndex("JCCostCode")) = "" & rs("JCCostCode")
            .TextMatrix(r, .ColIndex("JCCategory")) = "" & rs("JCCategory")
        Else
            .TextMatrix(r, .ColIndex("JCCostCode")) = IIf("" & rs("AltJCCostCode") = "", "" & rs("JCCostCode"), "" & rs("AltJCCostCode"))
            .TextMatrix(r, .ColIndex("JCCategory")) = IIf("" & rs("AltJCCategory") = "", "" & rs("JCCategory"), "" & rs("AltJCCategory"))
        End If
        .TextMatrix(r, .ColIndex("Price")) = "" & rs("Price")
        .TextMatrix(r, .ColIndex("RoundTo")) = "" & rs("RoundTo")
        .TextMatrix(r, .ColIndex("RoundDir")) = "" & rs("RoundDir")
        .TextMatrix(r, .ColIndex("POIndex")) = "" & rs("POIndex")
        .TextMatrix(r, .ColIndex("Notes")) = "" & rs("Notes")

        rs.MoveNext
    Wend
    End With

    Call LoadVariables(True, gHFVariables, variablelist)
    Call GetVariableValues
    Call Form_Resize

    'set qty to 1
    gHFVariables.TextMatrix(0, gHFVariables.ColIndex("Value")) = 1


End Sub

Private Sub AddMBTakeoff(Key As String)
    Dim s As String
    Dim rs As Recordset
    Dim rs1 As Recordset
    Dim r As Long

    gVariables.Rows = 0

    cmdAddPass.Enabled = False
    cmdUndoPass.Enabled = False
    cmdNav(0).Visible = True
    cmdNav(0).Enabled = True
    cmdNav(1).Caption = "&Cancel"

    s = ""
    s = s & "select phsnum,prtnum,prtdsc,linqty,untdsc,linprc" & vbCrLf
    s = s & "  from tkflin " & vbCrLf
    s = s & " where linqty*linprc<>0 " & Key & vbCrLf
    Set rs = HFApp.SqlExec(s, dbAccounting)
    With gItems
    .Rows = 1
    While Not rs.EOF

        Set rs1 = HFApp.SqlExec("SELECT * FROM tblPhaseItem WHERE DivisionID = " & HFApp.DivisionID & " and PartNumber=" & DbQuote(str, rs("prtnum")))

        .AddItem ""
        r = .Rows - 1
        .TextMatrix(r, .ColIndex("Description")) = "" & rs("prtdsc")
        .TextMatrix(r, .ColIndex("TakeoffQty")) = "" & rs("linqty")
        .TextMatrix(r, .ColIndex("TakeoffUOM")) = "" & rs("untdsc")
        .TextMatrix(r, .ColIndex("Price")) = "" & rs("linprc")
        .TextMatrix(r, .ColIndex("JCExtra")) = IIf(Val("" & rs("phsnum")) = 0, "", Val("" & rs("phsnum")))

        On Error Resume Next
        .TextMatrix(r, .ColIndex("Phase")) = "" & rs1("Phase")
        .TextMatrix(r, .ColIndex("Item")) = "" & rs1("Item")
        .TextMatrix(r, .ColIndex("OrderQty")) = RoundTo(Val("" & rs("linqty")) * Val("" & rs1("ConversionFactor")), Val("" & rs1("RoundTo")), Val("" & rs1("RoundDir")))
        .TextMatrix(r, .ColIndex("OrderUOM")) = "" & rs1("OrderUOM")
        .TextMatrix(r, .ColIndex("ConversionFactor")) = IIf(Val("" & rs1("ConversionFactor")) = 0, 1, Val("" & rs1("ConversionFactor")))
        If mAssemblyType = atModel Or (mUseAltCostCodesForCO And Not mIsChangeOrder) Then
            .TextMatrix(r, .ColIndex("JCCostCode")) = "" & rs1("JCCostCode")
            .TextMatrix(r, .ColIndex("JCCategory")) = "" & rs1("JCCategory")
        Else
            .TextMatrix(r, .ColIndex("JCCostCode")) = IIf("" & rs1("AltJCCostCode") = "", "" & rs1("JCCostCode"), "" & rs1("AltJCCostCode"))
            .TextMatrix(r, .ColIndex("JCCategory")) = IIf("" & rs1("AltJCCategory") = "", "" & rs1("JCCategory"), "" & rs1("AltJCCategory"))
        End If
        .TextMatrix(r, .ColIndex("RoundTo")) = "" & rs1("RoundTo")
        .TextMatrix(r, .ColIndex("RoundDir")) = "" & rs1("RoundDir")
        .TextMatrix(r, .ColIndex("POIndex")) = "" & rs1("POIndex")
        .TextMatrix(r, .ColIndex("Notes")) = "" & rs1("Notes")
        On Error GoTo 0

        rs.MoveNext
    Wend
    End With

End Sub

Public Sub AddItem(Phase As String, Item As String)
    Dim s As String
    Dim rs As Recordset
    Dim r As Long
    Dim w As Long
    Dim variablelist As String

    gVariables.Visible = False
    gHFVariables.Visible = True
    txtHelp.Text = ""
    cmdAddPass.Enabled = True
    cmdUndoPass.Enabled = True
    cmdNav(0).Visible = True
    cmdNav(0).Enabled = True
    cmdNav(1).Caption = "&Cancel"


    s = ""
    s = s & "select *" & vbCrLf
    s = s & " from EstimatingItems" & vbCrLf
    s = s & "where DivisionID = " & HFApp.DivisionID & " and phase=" & DbQuote(str, Phase) & vbCrLf
    s = s & "  and item=" & DbQuote(str, Item) & vbCrLf
    Set rs = HFApp.SqlExec(s, dbHomeFront)
    If Not rs.EOF Then

        With gItems
            .AddItem ""
            r = .Rows - 1


            .TextMatrix(r, .ColIndex("Model")) = mModel
            .TextMatrix(r, .ColIndex("Assembly")) = mAssembly
            .TextMatrix(r, .ColIndex("AssemblyDescription")) = mAssemblyDescription

            .TextMatrix(r, .ColIndex("rawformula")) = "" & rs("formula")
            .TextMatrix(r, .ColIndex("resolvedformula")) = ResolveFormula(.TextMatrix(r, .ColIndex("rawformula")))
            variablelist = variablelist & Chr(1) & ParseVariables(.TextMatrix(r, .ColIndex("resolvedformula")))
            
            If .TextMatrix(r, .ColIndex("resolvedformula")) = "" Then
                .TextMatrix(r, .ColIndex("resolvedformula")) = "1"
            End If
            
            .TextMatrix(r, .ColIndex("orderQty")) = 0
            .TextMatrix(r, .ColIndex("TakeoffQty")) = 0
            .TextMatrix(r, .ColIndex("p1")) = 0
            .TextMatrix(r, .ColIndex("p2")) = 0
            .TextMatrix(r, .ColIndex("p3")) = 0
            .TextMatrix(r, .ColIndex("p4")) = 0
            .TextMatrix(r, .ColIndex("p5")) = 0

            .TextMatrix(r, .ColIndex("Phase")) = "" & rs("Phase")
            .TextMatrix(r, .ColIndex("Item")) = "" & rs("Item")
            .TextMatrix(r, .ColIndex("Description")) = "" & rs("ItemDesc")
            .TextMatrix(r, .ColIndex("TakeoffUOM")) = "" & rs("TakeoffUOM")
            .TextMatrix(r, .ColIndex("OrderUOM")) = "" & rs("OrderUOM")
            .TextMatrix(r, .ColIndex("ConversionFactor")) = "" & rs("ConversionFactor")
            
            If mAssemblyType = atModel Or (mUseAltCostCodesForCO And Not mIsChangeOrder) Then
                .TextMatrix(r, .ColIndex("JCCostCode")) = "" & rs("JCCostCode")
                .TextMatrix(r, .ColIndex("JCCategory")) = "" & rs("JCCategory")
            Else
                .TextMatrix(r, .ColIndex("JCCostCode")) = IIf("" & rs("AltJCCostCode") = "", "" & rs("JCCostCode"), "" & rs("AltJCCostCode"))
                .TextMatrix(r, .ColIndex("JCCategory")) = IIf("" & rs("AltJCCategory") = "", "" & rs("JCCategory"), "" & rs("AltJCCategory"))
            End If
            
            .TextMatrix(r, .ColIndex("Price")) = "" & rs("Price")
            .TextMatrix(r, .ColIndex("WastePercent")) = "" & rs("WastePercent")
            .TextMatrix(r, .ColIndex("RoundTo")) = "" & rs("RoundTo")
            .TextMatrix(r, .ColIndex("RoundDir")) = "" & rs("RoundDir")
            .TextMatrix(r, .ColIndex("POIndex")) = "" & rs("POIndex")
            .TextMatrix(r, .ColIndex("Notes")) = "" & rs("Notes")
            .TextMatrix(r, .ColIndex("Location")) = "" & rs("Location")
            For w = 1 To 40
                .TextMatrix(r, .ColIndex("WBS" & format(w, "00"))) = "" & rs("WBS" & format(w, "00"))
            Next
        
            variablelist = "Quantity"
            For r = 0 To .Rows - 1
                variablelist = variablelist & Chr(1) & ParseVariables(.TextMatrix(r, .ColIndex("resolvedformula")))
            Next

        
        
        End With

        Call LoadVariables(False, gHFVariables, variablelist)
        Call GetVariableValues
        Call Form_Resize
    
        'set qty to 1
        If mViewIndex <> 4 Then
            gHFVariables.TextMatrix(0, gHFVariables.ColIndex("Value")) = 1
        End If
        
        
    End If
End Sub

Private Sub AddQuote(Level As Long, Job As String, ChangeOrder As String, EstAssemblyID As Long)
    Dim s As String
    Dim rs As Recordset
    Dim r As Long
    Dim where_clause As String
    Dim variablelist As String


    gVariables.Visible = False
    gHFVariables.Visible = True
    gHFVariables.Rows = 0
    txtHelp.Text = ""
'    variablelist = "Quantity"

    cmdAddPass.Enabled = True
    cmdUndoPass.Enabled = True
    cmdNav(0).Visible = True
    cmdNav(0).Enabled = True
    cmdNav(1).Caption = "&Cancel"

    s = ""
    If mViewIndex = 13 Then
        s = s & "select i.*,a.Description AssemblyDescription,a.Area,c.Model,c.Series,'' Formula from tblcustomers c" & vbCrLf
        s = s & " join custompreestimateitems i on (c.customer_no = i.customerno)"
        s = s & " join alloptions a on a.customer_no = i.customerno and a.Location = case i.Location when 1 then 'ADDENDUM' when 2 then 'CHANGEORDER' when 3 then 'DESIGNCENTER' else '' end and a.seq = i.seq"
        s = s & " where c.purchased = 0 and c.lastestimateindex = 0" & vbCrLf
        where_clause = Replace(gList.GetNode().Key, "assembly", "cast(i.seq as varchar)")
        where_clause = Replace(where_clause, "job", "Customerno")
        where_clause = Replace(where_clause, "AND changeorder = ''", "")
    Else
        s = s & "select *" & vbCrLf
        s = s & " from estimateditems " & vbCrLf
    End If
    Select Case Level
        Case 1 'job
            s = s & "where (isnull(BudgetDeleted,0)=0 or isnull(PODeleted,0)=0) and job_no=" & DbQuote(str, Job) & vbCrLf
        Case 2 'changeorder
            s = s & "where (isnull(BudgetDeleted,0)=0 or isnull(PODeleted,0)=0) and job_no=" & DbQuote(str, Job) & vbCrLf
            s = s & "  and changeorder=" & DbQuote(str, ChangeOrder) & vbCrLf
        Case 3 'assembly
            If mViewIndex = 13 Then
                s = s & " " & where_clause & vbCrLf
            Else
                s = s & "where (isnull(BudgetDeleted,0)=0 or isnull(PODeleted,0)=0) and estassemblyid=" & DbQuote(Num, EstAssemblyID) & vbCrLf
            End If
    End Select
    Set rs = HFApp.SqlExec(s, dbHomeFront)
    
    With gItems
    .Rows = 1
    While Not rs.EOF

        If Trim(mAssembly) = "" And mViewIndex <> 13 Then
            mAssemblyType = "" & rs("AssemblyType")
            mAssembly = "" & rs("Assembly")
            mAssemblyDescription = "" & rs("AssemblyDescription")
            mModel = "" & rs("Model")
            mOptionID = "" & rs("OptionID")
        End If
        
        
        .AddItem ""
        r = .Rows - 1
        If mViewIndex = 13 Then
            .TextMatrix(r, .ColIndex("Assembly")) = "" & rs("Assembly")
            .TextMatrix(r, .ColIndex("AssemblyDescription")) = "" & rs("AssemblyDescription")
            .TextMatrix(r, .ColIndex("Model")) = "" & rs("Model")
            .TextMatrix(r, .ColIndex("Phase")) = "" & rs("Phase")
            .TextMatrix(r, .ColIndex("Item")) = "" & rs("Item")
            .TextMatrix(r, .ColIndex("Description")) = "" & rs("Description")
    
    
            .TextMatrix(r, .ColIndex("rawformula")) = "" & rs("formula")
            .TextMatrix(r, .ColIndex("resolvedformula")) = ResolveFormula(.TextMatrix(r, .ColIndex("rawformula")))
            variablelist = variablelist & Chr(1) & ParseVariables(.TextMatrix(r, .ColIndex("resolvedformula")))
            If .TextMatrix(r, .ColIndex("resolvedformula")) = "" Then
                .TextMatrix(r, .ColIndex("resolvedformula")) = "" & rs("TakeoffQty")
            End If
            .TextMatrix(r, .ColIndex("TakeoffUOM")) = "" & rs("TakeoffUOM")
            .TextMatrix(r, .ColIndex("TakeoffQty")) = "" & rs("TakeoffQty")
            .TextMatrix(r, .ColIndex("p1")) = 0
            .TextMatrix(r, .ColIndex("p2")) = 0
            .TextMatrix(r, .ColIndex("p3")) = 0
            .TextMatrix(r, .ColIndex("p4")) = 0
            .TextMatrix(r, .ColIndex("p5")) = 0
            .TextMatrix(r, .ColIndex("OrderQty")) = "" & rs("BudgetQty")
            .TextMatrix(r, .ColIndex("OrderUOM")) = "" & rs("OrderUOM")
            .TextMatrix(r, .ColIndex("ConversionFactor")) = "" & rs("ConversionFactor")
            .TextMatrix(r, .ColIndex("JCCostCode")) = "" & rs("JCCostCode")
            .TextMatrix(r, .ColIndex("JCCategory")) = "" & rs("JCCategory")
            .TextMatrix(r, .ColIndex("Price")) = "" & rs("BudgetRate")
            .TextMatrix(r, .ColIndex("Location")) = "" & rs("Area")
            '.TextMatrix(r, .ColIndex("RoundTo")) = "" & rs("RoundTo")
            '.TextMatrix(r, .ColIndex("RoundDir")) = "" & rs("RoundDir")
            .TextMatrix(r, .ColIndex("POIndex")) = "" & rs("POIndex")
        Else
            .TextMatrix(r, .ColIndex("Assembly")) = "" & rs("Assembly")
            .TextMatrix(r, .ColIndex("AssemblyDescription")) = "" & rs("AssemblyDescription")
            .TextMatrix(r, .ColIndex("Model")) = "" & rs("Model")
            .TextMatrix(r, .ColIndex("Phase")) = "" & rs("EstPhase")
            .TextMatrix(r, .ColIndex("Item")) = "" & rs("EstItem")
            .TextMatrix(r, .ColIndex("Description")) = "" & rs("ItemDesc")
    
    
            .TextMatrix(r, .ColIndex("rawformula")) = "" & rs("formula")
            .TextMatrix(r, .ColIndex("resolvedformula")) = ResolveFormula(.TextMatrix(r, .ColIndex("rawformula")))
            variablelist = variablelist & Chr(1) & ParseVariables(.TextMatrix(r, .ColIndex("resolvedformula")))
            If .TextMatrix(r, .ColIndex("resolvedformula")) = "" Then
                .TextMatrix(r, .ColIndex("resolvedformula")) = "" & rs("TakeoffQty")
            End If
            .TextMatrix(r, .ColIndex("TakeoffUOM")) = "" & rs("TakeoffUOM")
            .TextMatrix(r, .ColIndex("TakeoffQty")) = "" & rs("TakeoffQty")
            .TextMatrix(r, .ColIndex("p1")) = 0
            .TextMatrix(r, .ColIndex("p2")) = 0
            .TextMatrix(r, .ColIndex("p3")) = 0
            .TextMatrix(r, .ColIndex("p4")) = 0
            .TextMatrix(r, .ColIndex("p5")) = 0
            .TextMatrix(r, .ColIndex("OrderQty")) = "" & rs("POQty")
            .TextMatrix(r, .ColIndex("OrderUOM")) = "" & rs("OrderUOM")
            .TextMatrix(r, .ColIndex("ConversionFactor")) = "" & rs("ConversionFactor")
            .TextMatrix(r, .ColIndex("JCCostCode")) = "" & rs("JCCostCode")
            .TextMatrix(r, .ColIndex("JCCategory")) = "" & rs("JCCategory")
            .TextMatrix(r, .ColIndex("Price")) = "" & rs("PORate")
            .TextMatrix(r, .ColIndex("RoundTo")) = "" & rs("RoundTo")
            .TextMatrix(r, .ColIndex("RoundDir")) = "" & rs("RoundDir")
            .TextMatrix(r, .ColIndex("POIndex")) = "" & rs("POIndex")
            .TextMatrix(r, .ColIndex("Location")) = "" & rs("Location")
        End If
        rs.MoveNext
    Wend
    End With



    'Call LoadVariables(True, gHFVariables, variablelist)
    'Call GetVariableValues
    Call Form_Resize


    
End Sub

Private Sub AddAssembly(Assembly As String)
    Dim i As Long
    Dim s As String
    Dim defValue As String
    Dim PEDBAssemblies    As Object 'PEDBAssemblyIterator
    Dim PEVariableValue As Object 'PEVariableValue
    Dim PEVariableDefinition As Object 'PEVariableDefinition
    Dim rs As Recordset


    gItems.Rows = 1
    gVariables.Rows = 0
    gVariables.Visible = True
    gHFVariables.Visible = False


    'load assembly
    Set PEDBAssemblies = PEDatabase.DBAssemblies.NewIterator(KEYED_SET_DEFAULT_ITERATOR)
    Set PEDBAssembly = PEDatabase.DBAssemblies.NewObject
    PEDBAssembly.code = Assembly
    Call PEDBAssemblies.GE(PEDBAssembly)
    Set PEDBAssembly = PEDBAssemblies.Buffer

    'do takeoff
    Set PEAssemblyTakeoff = PEEstimate.Takeoff.AssemblyTakeoff
    s = PEDBAssembly.code
    Call PEAssemblyTakeoff.LoadAssemblyByCode(Assembly)

    'load grid with variables
    With gVariables
        .Redraw = flexRDBuffered
        .Rows = 0
        .Rows = PEAssemblyTakeoff.VariableValues.Count + 1
        .Cell(flexcpBackColor, 0, 0, .Rows - 1, 0) = vbButtonFace
        .Cell(flexcpBackColor, 0, 2, .Rows - 1, 2) = vbButtonFace
        .TextMatrix(0, 0) = "Quantity"
        .TextMatrix(0, 1) = 1
        .TextMatrix(0, 2) = "each" '& PEDBAssembly.Unit
        .TextMatrix(0, 3) = "-1"

        For i = 1 To PEAssemblyTakeoff.VariableValues.Count
            Set PEVariableValue = PEAssemblyTakeoff.VariableValues.Item(i)
            Set PEVariableDefinition = PEVariableValue.VariableDefinition
            With gVariables
            
                .RowData(i) = PEVariableValue
                .TextMatrix(i, 0) = "" & PEVariableDefinition.Name

                On Error Resume Next
                defValue = ""
                defValue = "" & PEVariableDefinition.DefaultValue
                defValue = TakeoffParameters.Item(PEVariableDefinition.Name)
                PEVariableValue.value = Val(defValue)
                On Error GoTo 0

                If PEVariableDefinition.IsYesNoType Then
                    
                    .TextMatrix(i, 1) = IIf(defValue, "Yes", "No")
                    PEVariableValue.value = IIf(defValue, 1, 0)
                
                Else
                    If defValue = "" Then
                        .TextMatrix(i, 1) = ""
                    Else
                        .TextMatrix(i, 1) = "" & PEVariableValue.value
                    End If
                End If

                .TextMatrix(i, 2) = "" & PEVariableDefinition.Unit
                On Error Resume Next
                .TextMatrix(i, 3) = 1
                .TextMatrix(i, 3) = Val("" & PEVariableValue.DisplayOrder)
                If .ValueMatrix(i, 3) = 0 Then
                    .TextMatrix(i, 3) = 999
                    .RowHidden(i) = True
                Else
                    .RowHidden(i) = False
                End If
                On Error GoTo 0

                .TextMatrix(i, 4) = "" & PEVariableDefinition.Min
                .TextMatrix(i, 5) = "" & PEVariableDefinition.Max
                .TextMatrix(i, 6) = "" & PEVariableDefinition.VariableHelp
                .TextMatrix(i, 7) = "" & PEVariableDefinition.IsYesNoType


                'now get onscreen values from HF db
                Set rs = HFApp.SqlExec("select * from variables where name=" & DbQuote(str, .TextMatrix(i, 0)), dbHomeFront)
                If Not rs.EOF Then
                    .TextMatrix(i, .ColIndex("ConditionName")) = "" & rs("ConditionName")
                    .TextMatrix(i, .ColIndex("ConditionType")) = "" & rs("ConditionType")
                    .TextMatrix(i, .ColIndex("ConditionUOM")) = "" & rs("ConditionUOM")
                    .TextMatrix(i, .ColIndex("ConditionTakeoffUOM")) = "" & rs("UOM")
                End If
                

            End With
        Next

        Call Form_Resize 'to sort & size columns
    End With

    cmdAddPass.Enabled = True
    cmdUndoPass.Enabled = True
    cmdNav(0).Visible = True
    cmdNav(0).Enabled = False
    cmdNav(1).Caption = "&Cancel"

    Call GetVariableValues

    gVariables.SetFocus
    Call gVariables.Select(0, 1)


End Sub

Private Sub ReadAssemblyItems()
    Dim s As String
    Dim r As Long
    Dim i As Long
    Dim rs As Recordset
    Dim PEItemArray           As Object 'PETakeoffItemArray
    Dim PETakeoffItem         As Object 'PETakeoffItem

    Dim Phase As String
    Dim Item As String
    Dim Qty As Double

    Set PEItemArray = PEAssemblyTakeoff.TakeoffItemArray
    gItems.Rows = 1
    For i = 1 To PEItemArray.Count

        Set PETakeoffItem = PEItemArray.Item(i)
        Phase = Trim(PETakeoffItem.PhaseCode)
        Item = Trim(PETakeoffItem.ItemCode)
        Qty = PETakeoffItem.TakeoffQuantity

        If Phase <> "" And Item <> "" Then

            'select at takeoff items seems to be listed twice, once with a takeoff qty=0
            If Trim(Phase) = HFApp.Options(SelectAtTakeoffPhase) And Trim(Item) = HFApp.Options(SelectAtTakeoffItem) And PETakeoffItem.TakeoffQuantity <> 0 Then
                If FPickList.Choose(HFApp.Databases(dbHomeFront), "Select at Takeoff Item", "SELECT Phase,Item,Description FROM tblphaseitem WHERE DivisionID = " & HFApp.DivisionID & " and phase<>" & DbQuote(str, HFApp.Options(SelectAtTakeoffPhase)) & " And item<>" & DbQuote(str, HFApp.Options(SelectAtTakeoffItem))) Then
                    Phase = FPickList.SelectedItem("Phase")
                    Item = FPickList.SelectedItem("Item")
                Else
                    MsgBox "Select at takeoff item will not be included in this pass.", vbInformation, App.ProductName
                End If
            End If


            s = ""
            s = s & "select phase,item,itemdesc,takeoffuom,conversionfactor,orderuom,jccostcode,jccategory,altjccostcode,altjccategory,price,roundto,rounddir,poindex,notes,wastepercent" & vbCrLf
            s = s & " from EstimatingItems" & vbCrLf
            s = s & "where DivisionID = " & HFApp.DivisionID & " and phase=" & DbQuote(str, Phase) & vbCrLf
            s = s & "  and itemnumber=" & DbQuote(str, Item) & vbCrLf
            Set rs = HFApp.SqlExec(s, dbHomeFront)
            If rs.EOF Then
                MsgBox "Precision Builder is not able to use this assembly." & vbCrLf & vbCrLf & "    " & vbBullet & " Item " & Phase & "/" & Item & " is not in Precision Builders's phase item database." & vbCrLf & vbCrLf & "Before using this assembly you must either add the item to HomeFronts" & vbCrLf & "database or remove this item from the assembly.", vbExclamation, App.ProductName
                Exit Sub
            Else
                With gItems
                While Not rs.EOF
                    .AddItem ""
                    r = .Rows - 1
                    .TextMatrix(r, .ColIndex("Model")) = mModel
                    .TextMatrix(r, .ColIndex("Assembly")) = mAssembly
                    .TextMatrix(r, .ColIndex("AssemblyDescription")) = PEDBAssembly.Description
                    .TextMatrix(r, .ColIndex("Phase")) = "" & rs("Phase")
                    .TextMatrix(r, .ColIndex("Item")) = "" & rs("Item")
                    .TextMatrix(r, .ColIndex("Description")) = "" & rs("ItemDesc")
                    .TextMatrix(r, .ColIndex("TakeoffQty")) = Qty
                    .TextMatrix(r, .ColIndex("TakeoffUOM")) = "" & rs("TakeoffUOM")
                    .TextMatrix(r, .ColIndex("WastePercent")) = Val("" & rs("WastePercent"))
                    .TextMatrix(r, .ColIndex("OrderQty")) = RoundTo(Qty * (100 + Val("" & rs("WastePercent"))) / 100 * Val("" & rs("ConversionFactor")), Val("" & rs("RoundTo")), Val("" & rs("RoundDir")))
                    .TextMatrix(r, .ColIndex("OrderUOM")) = "" & rs("OrderUOM")
                    .TextMatrix(r, .ColIndex("ConversionFactor")) = "" & rs("ConversionFactor")
                    If mAssemblyType = atModel Or (mUseAltCostCodesForCO And Not mIsChangeOrder) Then
                        .TextMatrix(r, .ColIndex("JCCostCode")) = "" & rs("JCCostCode")
                        .TextMatrix(r, .ColIndex("JCCategory")) = "" & rs("JCCategory")
                    Else
                        .TextMatrix(r, .ColIndex("JCCostCode")) = IIf("" & rs("AltJCCostCode") = "", "" & rs("JCCostCode"), "" & rs("AltJCCostCode"))
                        .TextMatrix(r, .ColIndex("JCCategory")) = IIf("" & rs("AltJCCategory") = "", "" & rs("JCCategory"), "" & rs("AltJCCategory"))
                    End If
                    .TextMatrix(r, .ColIndex("Price")) = "" & rs("Price")
                    .TextMatrix(r, .ColIndex("RoundTo")) = "" & rs("RoundTo")
                    .TextMatrix(r, .ColIndex("RoundDir")) = "" & rs("RoundDir")
                    .TextMatrix(r, .ColIndex("POIndex")) = "" & rs("POIndex")
                    .TextMatrix(r, .ColIndex("Notes")) = "" & rs("Notes")
                    rs.MoveNext
                Wend
                End With

            End If


        End If
    Next

    Call gItems.AutoSize(0, gItems.Cols - 1)
    cmdNav(0).Enabled = gItems.Rows > 1

End Sub

Private Sub gItems_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With gItems
        .ComboList = ""
        Select Case .ColKey(Col)
            Case "OrderQty", "TakeoffQty"
            Case "POIndex"
                .ComboList = "..."
            Case "Location"
                .EditMaxLength = 200
            Case Else
                Cancel = True
        End Select
    End With
End Sub
Private Sub gItems_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim r As Long
    Dim TakeoffQty As Double
    Dim OrderQty As Double
    Dim Conversion As Double
    Dim RoundDir As Long
    Dim RoundUnit As Double
    Dim WastePercent As Long

    With gItems
        For r = Min(.Row, .RowSel) To Max(.Row, .RowSel)
            Select Case .ColKey(Col)
            
                Case "PassTakeoffQty"
                    TakeoffQty = Val(.TextMatrix(r, .ColIndex("PassTakeoffQty")))
                    Conversion = Val(.TextMatrix(r, .ColIndex("ConversionFactor")))
                    If Conversion = 0 Then Conversion = 1
                    RoundDir = Val(.TextMatrix(r, .ColIndex("RoundDir")))
                    RoundUnit = Val(.TextMatrix(r, .ColIndex("RoundTo")))
                    WastePercent = Val(.TextMatrix(r, .ColIndex("WastePercent")))
                    OrderQty = RoundTo(TakeoffQty * (100 + WastePercent) / 100 * Conversion, RoundUnit, RoundDir)
                    .TextMatrix(r, .ColIndex("PassOrderQty")) = OrderQty
            
                Case "TakeoffQty"
                    TakeoffQty = Val(.TextMatrix(r, .ColIndex("TakeoffQty")))
                    Conversion = Val(.TextMatrix(r, .ColIndex("ConversionFactor")))
                    If Conversion = 0 Then Conversion = 1
                    RoundDir = Val(.TextMatrix(r, .ColIndex("RoundDir")))
                    RoundUnit = Val(.TextMatrix(r, .ColIndex("RoundTo")))
                    WastePercent = Val(.TextMatrix(r, .ColIndex("WastePercent")))
                    OrderQty = RoundTo(TakeoffQty * (100 + WastePercent) / 100 * Conversion, RoundUnit, RoundDir)
                    
                    .TextMatrix(r, .ColIndex("OrderQty")) = OrderQty

                Case "OrderQty"
                    OrderQty = Val(.TextMatrix(r, .ColIndex("OrderQty")))
                    Conversion = Val(.TextMatrix(r, .ColIndex("ConversionFactor")))
                    If Conversion = 0 Then Conversion = 1
                    WastePercent = Val(.TextMatrix(r, .ColIndex("WastePercent")))
                    TakeoffQty = Round(OrderQty / ((100 + WastePercent) / 100) / Conversion, 5)
                    .TextMatrix(r, .ColIndex("TakeoffQty")) = TakeoffQty

            End Select
        Next
    End With
End Sub

Private Sub gItems_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
    If Button <> vbRightButton Then Exit Sub
    With gItems
        If .MouseRow = 0 Then
            Cancel = True
            Call FMain.ShowColumnMenu(gItems)
        Else
            If .Row > 0 Then Call PopupMenu(FMain.mnuAssemblyItems)
        End If
    End With
End Sub

Private Sub gItems_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim r As Long
    Dim i As Long
    Dim s As String
    Select Case True
        Case KeyCode = vbKeyDelete And Shift = vbCtrlMask
            With gItems
                For r = Max(.RowSel, .Row) To Min(.RowSel, .Row) Step -1
                    Call .RemoveItem(r)
                Next
            End With

        Case KeyCode = vbKeyF And Shift = vbCtrlMask
            s = "select i.Phase,p.Description PhaseDesc, i.Item, POIndex, i.Description, OrderUOM from tblPhaseItem i join tblestphases p on (p.divisionid = i.divisionid and p.phase = i.phase) where i.DivisionID = " & HFApp.DivisionID & "  order by i.Phase,i.item"
            If FPickList.Choose(HFApp.Databases(dbHomeFront), "Item Takeoff", s, , , , , , True, 1) Then
                For i = 1 To FPickList.SelectedItems
                    Call AddItem(FPickList.SelectedItem("Phase", i), FPickList.SelectedItem("Item", i))
                Next
            End If
            'Call FFind.ShowForm(gItems)
    End Select
End Sub

Private Sub gItems_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case gItems.ColKey(Col)
        Case "TakeoffQty", "OrderQty"
            gItems.EditText = Val(gItems.EditText)
    End Select
End Sub

Public Sub mnuAssemblyItemsSub_Click(Index As Integer)
    Dim s      As String
    Dim c      As Long
    Dim r      As Long
    Dim NewRow As Long

    With gItems
    Select Case Index

        Case mcITEM_COPY
            For r = Max(.Row, .RowSel) To Min(.Row, .RowSel) Step -1
                NewRow = Max(.Row, .RowSel) + 1
                .AddItem "", NewRow
                For c = 0 To .Cols - 1
                    .TextMatrix(NewRow, c) = .TextMatrix(r, c)
                Next
            Next

        Case mcITEM_SUBSTITUE
            s = ""
            s = s & "SELECT i.Phase" & vbCrLf
            s = s & "      ,i.Item" & vbCrLf
            s = s & "      ,i.Description" & vbCrLf
            s = s & "      ,i.TakeoffUOM" & vbCrLf
            s = s & "      ,i.ConversionFactor" & vbCrLf
            s = s & "      ,i.OrderUOM" & vbCrLf
            s = s & "      ,i.POIndex" & vbCrLf
            s = s & "      ,i.Notes" & vbCrLf
            s = s & "      ,i.JCCostCode CostCode" & vbCrLf
            s = s & "      ,cod.description CostCodeDesc" & vbCrLf
            s = s & "      ,i.JCCategory Category" & vbCrLf
            s = s & "      ,cat.Description CategoryDesc" & vbCrLf
            s = s & "  FROM tblPhaseItem i" & vbCrLf
            s = s & "LEFT OUTER JOIN StandardCostCodes cod ON(i.DivisionID = cod.DivisionID and i.jccostcode=cod.costcode)" & vbCrLf
            s = s & "LEFT OUTER JOIN StandardCategories cat ON(i.DivisionID = cat.DivisionID and i.jccategory=cat.category)" & vbCrLf
            s = s & " WHERE i.DivisionID = " & HFApp.DivisionID
            .HighLight = flexHighlightAlways
            If FPickList.Choose(HFApp.Databases(dbHomeFront), "Item", s, , , , , "TakeoffUOM,ConversionFactor,OrderUOM,POIndex,Notes,CostCode,CostCodeDesc,Category,CategoryDesc") Then
                For r = Max(.Row, .RowSel) To Min(.Row, .RowSel) Step -1
                    .TextMatrix(r, .ColIndex("Assembly")) = ""
                    .TextMatrix(r, .ColIndex("AssemblyDescription")) = ""
                    .TextMatrix(r, .ColIndex("Phase")) = FPickList.SelectedItem("Phase")
                    .TextMatrix(r, .ColIndex("Item")) = FPickList.SelectedItem("Item")
                    .TextMatrix(r, .ColIndex("Description")) = FPickList.SelectedItem("Description")
                    .TextMatrix(r, .ColIndex("TakeoffUOM")) = FPickList.SelectedItem("TakeoffUOM")
                    .TextMatrix(r, .ColIndex("ConversionFactor")) = Val("" & FPickList.SelectedItem("ConversionFactor"))
                    .TextMatrix(r, .ColIndex("OrderUOM")) = FPickList.SelectedItem("OrderUOM")
                    .TextMatrix(r, .ColIndex("POIndex")) = FPickList.SelectedItem("POIndex")
                    .TextMatrix(r, .ColIndex("Notes")) = FPickList.SelectedItem("Notes")
                    .TextMatrix(r, .ColIndex("JCCostCode")) = FPickList.SelectedItem("CostCode")
                    .TextMatrix(r, .ColIndex("JCCostCodeDesc")) = FPickList.SelectedItem("CostCodeDesc")
                    .TextMatrix(r, .ColIndex("JCCategory")) = FPickList.SelectedItem("Category")
                    .TextMatrix(r, .ColIndex("JCCategoryDesc")) = FPickList.SelectedItem("CategoryDesc")
                    .TextMatrix(r, .ColIndex("OrderQty")) = .ValueMatrix(r, .ColIndex("TakeoffQty")) * .ValueMatrix(r, .ColIndex("ConversionFactor"))
                Next
            End If
            .HighLight = flexHighlightWithFocus


        Case mcITEM_REMOVE
            Call gItems_KeyDown(vbKeyDelete, vbCtrlMask)



    End Select
    End With
End Sub

Public Sub mnuTakeOffAssemblyView_Click(Index As Integer)
    mViewIndex = Index
    Call LoadList(True)
End Sub

Public Sub mnuTakeOffJobView_Click(Index As Integer)
    mViewIndex = Index
    Call LoadList(True)
End Sub

Public Sub mnuTakeOffItemView_Click(Index As Integer)
    mViewIndex = Index
    Call LoadList(True)
End Sub

Public Sub mnuTakeOffModelsView_Click(Index As Integer)
    mViewIndex = Index
    If Index < 4 Then
        Call LoadList(True)
    Else
        Call frmOptionListMultiSelect.ShowForm(mJob)
    End If
End Sub

Public Sub mnuTakeOffEstimateView_Click(Index As Integer)
    mViewIndex = Index
    Call LoadList(True)
End Sub


Private Sub gVariables_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo eh
    Dim s As String
    Dim i As Long
    Dim w As Long

    With gVariables

        gItems.SetFocus
        gVariables.SetFocus

        i = .ColWidth(0)
        Call .AutoSize(0)
        w = .ColWidth(0)
        .ColWidth(0) = i

        
        s = Trim(.TextMatrix(.Row, 6))
        s = Replace(s, vbCr, "")
        s = Replace(s, vbLf, vbCrLf)
        txtHelp.Text = s

        .Row = Min(.Row, GridLastVisibleRow(gVariables))
        cmdAddPass.Default = .Row = GridLastVisibleRow(gVariables)

    End With
eh: Exit Sub
End Sub

Private Sub gVariables_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With gVariables
        If .TextMatrix(Row, 7) = "True" Then
            .ComboList = "|Yes|No"
        Else
            .ComboList = ""
        End If
    End With
End Sub

Private Sub gVariables_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
On Error GoTo EXITSUB

    Dim lastrow As Long
    lastrow = GridLastVisibleRow(gVariables)
    With gVariables

        If mSorting Then Exit Sub

        If NewCol = 0 Then
            Cancel = True
            .Row = Max(0, OldRow - 1)
        End If

        If NewCol = 2 Then
            Cancel = True
            .Row = Min(OldRow + 1, lastrow)
        End If

        .Col = 1

    End With
EXITSUB:
End Sub

Private Sub gVariables_GotFocus()
On Error Resume Next
    If gVariables.Row < 0 Then gVariables.Row = 0
    gVariables.Col = 1
End Sub

Private Sub gVariables_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next

    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        gVariables.Row = gVariables.Row + 1
        Call gVariables.ShowCell(gVariables.Row, gVariables.Col)
    End If
End Sub

Private Sub gVariables_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
On Error Resume Next

    If KeyCode = vbKeyReturn And gVariables.Row <> gVariables.Rows - 1 Then
        KeyCode = 0
        gVariables.Row = gVariables.Row + 1
    End If
    Call gVariables.ShowCell(gVariables.Row, gVariables.Col)
End Sub

Private Sub gVariables_LostFocus()
    'picVariableHelp.Visible = False
End Sub

Private Sub gVariables_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim i As Long
    Dim d As Long
    Dim PEVariableValue As Object 'PEVariableValue

    On Error GoTo eh
    With gVariables

        If Row = 0 Then
            PEAssemblyTakeoff.Quantity = Val(.EditText)
            On Error Resume Next
            .Row = .Row + 1
            .Col = 1
            Exit Sub
        End If

        Set PEVariableValue = .RowData(Row)

        If .TextMatrix(Row, 7) = "True" Then
            If IsNumeric(.EditText) Then
                .EditText = IIf(Val(.EditText) <> 0, "Yes", "No")
            Else
                .EditText = IIf(.EditText = "Yes" Or .EditText = "Y", "Yes", "No")
            End If
            PEVariableValue.value = IIf(.EditText = "Yes", 1, 0)
        Else
            If Not IsNumeric(.EditText) Then
                Cancel = True
            Else
                If Val(.EditText) > .ValueMatrix(Row, 5) And .ValueMatrix(Row, 5) > 0 Then
                    MsgBox "The maximum value for " & .TextMatrix(Row, 0) & " is '" & .ValueMatrix(Row, 5) & "'. The number you have entered is too large.", vbInformation, App.ProductName
                    Cancel = True
                End If
                If Row > 0 And Val(.EditText) < .ValueMatrix(Row, 4) Then
                    MsgBox "The minimum value for " & .TextMatrix(Row, 0) & " is '" & .ValueMatrix(Row, 4) & "'. The number you have entered is too small.", vbInformation, App.ProductName
                    Cancel = True
                End If
            End If
            .EditText = Val(.EditText)
        End If

        If Not Cancel Then
        
            .Text = .EditText
        
        
            If PEVariableValue.VariableDefinition.IsYesNoType Then
                PEVariableValue.value = IIf(.EditText = "Yes", 1, 0)
            Else
                PEVariableValue.value = Val(.EditText)
            End If
            On Error Resume Next
            Call TakeoffParameters.Remove(.TextMatrix(Row, 0))
            Call TakeoffParameters.Add(PEVariableValue.value, .TextMatrix(Row, 0))

            For i = 1 To .Rows - 1
                Set PEVariableValue = .RowData(i)


                On Error Resume Next
                .TextMatrix(i, 3) = 1
                .TextMatrix(i, 3) = Val("" & PEVariableValue.DisplayOrder)
                If .ValueMatrix(i, 3) = 0 Then
                    .TextMatrix(i, 3) = 999
                    .RowHidden(i) = True
                Else
                    .RowHidden(i) = False
                End If
                On Error GoTo 0

            Next
            On Error GoTo 0
            Call Form_Resize

            On Error Resume Next
            .Col = 1
            .Row = Row


        End If
    End With
Exit Sub
eh:
If Err.Number <> 0 Then
    MsgBox Err.Description
    Resume Next
End If

End Sub

Private Sub gHFVariables_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo eh
    Dim s As String
    Dim i As Long
    Dim w As Long

    With gHFVariables
        txtHelp.Text = .TextMatrix(NewRow, .ColIndex("help"))
        cmdAddPass.Default = .Row = .Rows - 1
    End With
eh: Exit Sub
End Sub

Private Sub gHFVariables_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With gHFVariables
        .ComboList = ""

        Select Case .ColKey(Col)
            Case "value"
                If .TextMatrix(Row, .ColIndex("ListOfValues")) = "" Then
                    If frmOnScreen.Visible Then
                        .ComboList = "|..."
                    End If
                Else
                    .ComboList = .TextMatrix(Row, .ColIndex("ListOfValues"))
                End If


            Case "uom"
                If Row > 0 And (.TextMatrix(Row, Col) = "" Or .Cell(flexcpData, Row, .ColIndex("UOM")) = "") Then
                    .ComboList = "EA|IN|LF|LY|m|mm|IN²|SF|SQ|SY|m²|mm²|CF|CY|m³|mm³"
                Else
                    Cancel = True
                End If

            Case Else
                Cancel = True
        End Select
    End With
End Sub

Private Sub gHFVariables_GotFocus()
On Error Resume Next
    If gHFVariables.Row < 0 Then gHFVariables.Row = 0
    gHFVariables.Col = 1
End Sub

Private Sub gHFVariables_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
        KeyCode = 0
        Call gHFVariables.Select(gHFVariables.Row + 1, 1)
        Call gHFVariables.ShowCell(gHFVariables.Row, 1)
    End If
End Sub



Private Sub gHFVariables_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
On Error Resume Next
    If (KeyCode = vbKeyTab Or KeyCode = vbKeyReturn) And gHFVariables.Row <> gHFVariables.Rows - 1 Then
        KeyCode = 0
        Call gHFVariables.Select(gHFVariables.Row + 1, 1)
    End If
    Call gHFVariables.ShowCell(gHFVariables.Row, gHFVariables.Col)
End Sub

Private Sub gHFVariables_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim i As Long
    Dim d As Long
    Dim PEVariableValue As Object 'PEVariableValue
    Dim f As Double
    Dim s As String
    Dim rs As Recordset

    
    With gHFVariables
        Select Case .ColKey(Col)
            Case "uom"
                If Val(.TextMatrix(Row, .ColIndex("Value"))) <> 0 And .TextMatrix(Row, .ColIndex("uom")) <> "" Then
                    s = ""
                    s = s & "select factor" & vbCrLf
                    s = s & "  from unitconversions" & vbCrLf
                    s = s & "where fromunit=" & DbQuote(str, .TextMatrix(Row, .ColIndex("uom"))) & vbCrLf
                    s = s & "  and tounit=" & DbQuote(str, .EditText)
                    Set rs = HFApp.SqlExec(s, dbHomeFront)
                    If Not rs.EOF Then
                        f = Val("" & rs(0))
                        If f = 0 Then f = 1
                        .TextMatrix(Row, .ColIndex("Value")) = .TextMatrix(Row, .ColIndex("Value")) * f
                        Call Form_Resize
                    End If
                End If
                gHFVariables.SetFocus
                
            Case "value"
                If .TextMatrix(Row, .ColIndex("ListOfValues")) <> "" Then
                    If IsNumeric(.EditText) Then
                        .EditText = Parse(.TextMatrix(Row, .ColIndex("ListOfValues")), Val(.EditText), "|")
                    End If
                Else
                    If Not IsNumeric(.EditText) Then
                        Cancel = True
                    Else
                        If Val(.EditText) > .ValueMatrix(Row, 5) And .ValueMatrix(Row, 5) > 0 Then
                            MsgBox "The maximum value for " & .TextMatrix(Row, 0) & " is '" & format(.ValueMatrix(Row, 5), "0") & "'. The number you have entered is too large.", vbInformation, App.ProductName
                            Cancel = True
                        End If
                        If Val(.EditText) < .ValueMatrix(Row, 4) Then
                            MsgBox "The minimum value for " & .TextMatrix(Row, 0) & " is '" & format(.ValueMatrix(Row, 4), "0") & "'. The number you have entered is too small.", vbInformation, App.ProductName
                            Cancel = True
                        End If
                    End If
                    .EditText = Val(.EditText)
                End If
                If Not Cancel Then
                    On Error Resume Next
                    .Col = 1
                    .Row = Row
                End If
        End Select
    End With
End Sub


Private Sub GetVariableValues()
    Dim i As Long
    Dim s As String
    Dim rs As Recordset

    Dim ConditionName As String
    Dim ConditionType As String
    Dim ConditionUOM As String
    Dim VariableUOM As String


If OnScreenConnection Is Nothing Then Exit Sub
If OnScreenConnection.State <> adStateOpen Then Exit Sub

    With IIf(gVariables.Visible, gVariables, gHFVariables)
        For i = 1 To .Rows - 1


            'if reading from TL assembly then use different column
            VariableUOM = .TextMatrix(i, .ColIndex(IIf(gVariables.Visible, "ConditionTakeoffUOM", "UOM")))

            
            ConditionName = .TextMatrix(i, .ColIndex("ConditionName"))
            ConditionType = .TextMatrix(i, .ColIndex("ConditionType"))
            ConditionUOM = .TextMatrix(i, .ColIndex("ConditionUOM"))


            s = ""
            s = s & "select c.name" & vbCrLf
            s = s & "      ,q.description as type" & vbCrLf
            s = s & "      ,round(sum(t.Quantity1) * u.conversion * x.factor,0) as qty" & vbCrLf
            s = s & "      ,x.tounit as uom" & vbCrLf
            s = s & "  from bids b" & vbCrLf
            s = s & "      ,bidconditions c" & vbCrLf
            s = s & "      ,bidtakeofftotals t " & vbCrLf
            s = s & "      ,HF_UnitCodes u" & vbCrLf
            s = s & "      ,HF_QtyCodes  q" & vbCrLf
            s = s & "      ,HF_Conversions x" & vbCrLf
            s = s & " where b.uid=c.biduid" & vbCrLf
            s = s & "   and b.uid=t.biduid" & vbCrLf
            s = s & "   and c.uid=t.bidconditionuid" & vbCrLf
            s = s & "   and c.uom1=u.code" & vbCrLf
            s = s & "   and c.quantity1=q.code" & vbCrLf
            s = s & "   and u.description=x.fromunit" & vbCrLf
            s = s & "   and b.uid=" & DbQuote(Num, mOnScreenBidID) & vbCrLf
            s = s & "   and c.name=" & DbQuote(str, ConditionName) & vbCrLf
            s = s & "   and q.description=" & DbQuote(str, ConditionType) & vbCrLf
            If VariableUOM <> "" Then s = s & "   and x.tounit=" & DbQuote(str, VariableUOM) & vbCrLf
            s = s & "group by c.name,q.description,u.conversion,x.factor,x.tounit" & vbCrLf
            s = s & "UNION" & vbCrLf
            s = s & "select c.name" & vbCrLf
            s = s & "      ,q.description as type" & vbCrLf
            s = s & "      ,round(sum(t.Quantity2) * u.conversion * x.factor,0) as qty" & vbCrLf
            s = s & "      ,x.tounit as uom" & vbCrLf
            s = s & "  from bids b" & vbCrLf
            s = s & "      ,bidconditions c" & vbCrLf
            s = s & "      ,bidtakeofftotals t " & vbCrLf
            s = s & "      ,HF_UnitCodes u" & vbCrLf
            s = s & "      ,HF_QtyCodes  q" & vbCrLf
            s = s & "      ,HF_Conversions x" & vbCrLf
            s = s & " where b.uid=c.biduid" & vbCrLf
            s = s & "   and b.uid=t.biduid" & vbCrLf
            s = s & "   and c.uid=t.bidconditionuid" & vbCrLf
            s = s & "   and c.uom2=u.code" & vbCrLf
            s = s & "   and c.quantity2=q.code" & vbCrLf
            s = s & "   and u.description=x.fromunit" & vbCrLf
            s = s & "   and b.uid=" & DbQuote(Num, mOnScreenBidID) & vbCrLf
            s = s & "   and c.name=" & DbQuote(str, ConditionName) & vbCrLf
            s = s & "   and q.description=" & DbQuote(str, ConditionType) & vbCrLf
            If VariableUOM <> "" Then s = s & "   and x.tounit=" & DbQuote(str, VariableUOM) & vbCrLf
            s = s & "group by c.name,q.description,u.conversion,x.factor,x.tounit" & vbCrLf
            s = s & "UNION" & vbCrLf
            s = s & "select c.name" & vbCrLf
            s = s & "      ,q.description as type" & vbCrLf
            s = s & "      ,round(sum(t.Quantity3) * u.conversion * x.factor,0) as qty" & vbCrLf
            s = s & "      ,x.tounit as uom" & vbCrLf
            s = s & "  from bids b" & vbCrLf
            s = s & "      ,bidconditions c" & vbCrLf
            s = s & "      ,bidtakeofftotals t " & vbCrLf
            s = s & "      ,HF_UnitCodes u" & vbCrLf
            s = s & "      ,HF_QtyCodes  q" & vbCrLf
            s = s & "      ,HF_Conversions x" & vbCrLf
            s = s & " where b.uid=c.biduid" & vbCrLf
            s = s & "   and b.uid=t.biduid" & vbCrLf
            s = s & "   and c.uid=t.bidconditionuid" & vbCrLf
            s = s & "   and c.uom3=u.code" & vbCrLf
            s = s & "   and c.quantity3=q.code" & vbCrLf
            s = s & "   and u.description=x.fromunit" & vbCrLf
            s = s & "   and b.uid=" & DbQuote(Num, mOnScreenBidID) & vbCrLf
            s = s & "   and c.name=" & DbQuote(str, ConditionName) & vbCrLf
            s = s & "   and q.description=" & DbQuote(str, ConditionType) & vbCrLf
            If VariableUOM <> "" Then s = s & "   and x.tounit=" & DbQuote(str, VariableUOM) & vbCrLf
            s = s & "group by c.name,q.description,u.conversion,x.factor,x.tounit" & vbCrLf
            s = s & "UNION" & vbCrLf
            
            
            s = s & "select c.name" & vbCrLf
            s = s & "      ,'Height' as type" & vbCrLf
            
            
            If OnScreenConnection.Properties("DBMS Name") = "Microsoft SQL Server" Then
                s = s & "      ,round(c.height * x.factor / case when b.measurebase=1 then b.scalefactor1 else 1 end,0) as qty" & vbCrLf
            Else
                s = s & "      ,round(c.height * x.factor / iif(b.measurebase=1,b.scalefactor1,1),0) as qty" & vbCrLf
            End If
            
            s = s & "      ,x.tounit as uom" & vbCrLf
            s = s & "  from bids b" & vbCrLf
            s = s & "      ,bidconditions c" & vbCrLf
            s = s & "      ,HF_Conversions x" & vbCrLf
            s = s & " where b.uid=c.biduid" & vbCrLf
            's = s & "   and c.height/iif(b.measurebase=1,1,b.scalefactor2)<>0" & vbCrLf  --IF is not available in MSSQL
            s = s & "   and ((b.measurebase=1 and c.height<>0) or (b.measurebase<>1 and c.height/b.scalefactor2<>0))" & vbCrLf
            's = s & "   and iif(b.measurebase=0,'IN','mm')=x.fromunit" & vbCrLf  --IF is not available in MSSQL
            s = s & "   and ((b.measurebase=0  and 'IN'=x.fromunit) or (b.measurebase<>0  and 'mm'=x.fromunit))" & vbCrLf
            s = s & "   and b.uid=" & DbQuote(Num, mOnScreenBidID) & vbCrLf
            s = s & "   and c.name=" & DbQuote(str, ConditionName) & vbCrLf
            s = s & "   and 'Height'=" & DbQuote(str, ConditionType) & vbCrLf
            
            
            If VariableUOM <> "" Then s = s & "   and x.tounit=" & DbQuote(str, VariableUOM) & vbCrLf
            s = s & "UNION" & vbCrLf
            s = s & "select c.name" & vbCrLf
            s = s & "      ,'Thickness' as type" & vbCrLf
            If OnScreenConnection.Properties("DBMS Name") = "Microsoft SQL Server" Then
                s = s & "      ,round(c.thickness * x.factor / case when b.measurebase=1 then b.scalefactor1 else 1 end,0) as qty" & vbCrLf
            Else
                s = s & "      ,round(c.thickness * x.factor / iif(b.measurebase=1,b.scalefactor1,1),0) as qty" & vbCrLf
            End If
            s = s & "      ,x.tounit as uom" & vbCrLf
            s = s & "  from bids b" & vbCrLf
            s = s & "      ,bidconditions c" & vbCrLf
            s = s & "      ,HF_Conversions x" & vbCrLf
            s = s & " where b.uid=c.biduid" & vbCrLf
         's = s & "   and round(c.thickness/iif(b.measurebase=1,b.scalefactor1,1))<>0" & vbCrLf  --IF is not available in MSSQL
            s = s & "   and ((b.measurebase=1 and round(c.thickness/b.scalefactor1,0)<>0) or (b.measurebase<>1 and c.thickness<>0))" & vbCrLf
          's = s & "   and iif(b.measurebase=0,'IN','mm')=x.fromunit" & vbCrLf  --IF is not available in MSSQL
            s = s & "   and ((b.measurebase=0  and 'IN'=x.fromunit) or (b.measurebase<>0  and 'mm'=x.fromunit))" & vbCrLf
            s = s & "   and b.uid=" & DbQuote(Num, mOnScreenBidID) & vbCrLf
            s = s & "   and c.name=" & DbQuote(str, ConditionName) & vbCrLf
            s = s & "   and 'Thickness'=" & DbQuote(str, ConditionType) & vbCrLf
            If VariableUOM <> "" Then s = s & "   and x.tounit=" & DbQuote(str, VariableUOM) & vbCrLf
            s = s & "order by 1,2,3" & vbCrLf

            Set rs = OnScreenConnection.Execute(s)
            If Not rs.EOF Then
                .TextMatrix(i, .ColIndex("value")) = Val("" & rs("qty"))
                If gVariables.Visible Then
                   Call gVariables_ValidateEdit(i, gVariables.ColIndex("value"), False)
                End If
            End If

        Next
    End With





End Sub
Private Sub Slider_Move(Index As Integer)
    Call Form_Resize
End Sub
Private Sub Form_Resize()
On Error Resume Next

    Const margin = 60
    
    
    If Me.WindowState = 1 Then Exit Sub
    
    Slider(0).Min = 960
    Slider(0).Max = Me.ScaleHeight - 960
    Slider(1).Min = 1920
    Slider(1).Max = Me.ScaleWidth - 960
    Slider(2).Min = 840
    Slider(2).Max = Slider(0).Top - 750
    Slider(0).Move 0, Max(Slider(0).Min, Min(Slider(0).Max, Slider(0).Top)), Me.ScaleWidth
    Slider(1).Move Max(Slider(1).Min, Min(Slider(1).Max, Slider(1).Left)), IIf(frmOnScreen.Visible, frmOnScreen.Height, 0), Slider(1).Width, Slider(0).Top - IIf(frmOnScreen.Visible, frmOnScreen.Height, 0)
    Slider(2).Move Slider(1).Left + Slider(1).Width, Max(Slider(2).Min, Min(Slider(2).Max, Slider(2).Top)), Me.ScaleWidth - Slider(1).Left - Slider(1).Width
    Slider(2).ZOrder 0
    Slider(1).ZOrder 0
    Slider(0).ZOrder 0
    
    
    frmOnScreen.Move 0, 0, Me.ScaleWidth
    txtOnScreenProject.Width = Me.ScaleWidth - txtOnScreenProject.Left - cmdBrowse.Width - 30
    cmdBrowse.Move Me.ScaleWidth - cmdBrowse.Width, 30
    
    
    gList.Move 0, IIf(frmOnScreen.Visible, frmOnScreen.Height, 0), Slider(1).Left, Slider(0).Top - IIf(frmOnScreen.Visible, frmOnScreen.Height, 0)
    gVariables.Move Slider(2).Left, Slider(1).Top, Slider(2).Width, Slider(2).Top - Slider(1).Top
    gHFVariables.Move gVariables.Left, gVariables.Top, gVariables.Width, gVariables.Height
    txtHelp.Move Slider(2).Left, Slider(2).Top + Slider(2).Height, Slider(2).Width - cmdUndoPass.Width - (2 * Screen.TwipsPerPixelX), Slider(0).Top - Slider(2).Top - Slider(2).Height - Screen.TwipsPerPixelY
    
    cmdUndoPass.Move txtHelp.Left + txtHelp.Width + (2 * Screen.TwipsPerPixelX), Slider(0).Top - (2 * cmdUndoPass.Height) - (2 * Screen.TwipsPerPixelY)
    cmdAddPass.Move txtHelp.Left + txtHelp.Width + (2 * Screen.TwipsPerPixelX), Slider(0).Top - (1 * cmdUndoPass.Height) - (2 * Screen.TwipsPerPixelY)
    
    
    gItems.Move 0, Slider(0).Top + Slider(0).Height, Me.ScaleWidth, Me.ScaleHeight - Slider(0).Top - Slider(0).Height - cmdNav(0).Height - 2 * margin
    
    cmdNav(0).Move Me.ScaleWidth - (2 * (margin + cmdNav(0).Width)), Me.ScaleHeight - margin - cmdNav(0).Height - Screen.TwipsPerPixelY
    cmdNav(1).Move Me.ScaleWidth - (1 * (margin + cmdNav(0).Width)), Me.ScaleHeight - margin - cmdNav(0).Height - Screen.TwipsPerPixelY
    lblCategory.Top = gItems.Top + gItems.Height + margin
    cboCategory.Top = lblCategory.Top
    txtQty.Top = lblCategory.Top
    lblQty.Top = lblCategory.Top
    
    
    With gVariables
        Call .AutoSize(0, 1, 2)
        .ColWidth(1) = Max(.ColWidth(1), 1020)
        .ColWidth(0) = .Width - .ColWidth(1) - Max(.ColWidth(2), 700)
        mSorting = True
        .Col = 3
        .Sort = flexSortNumericAscending
        .Col = 1
        mSorting = False
    End With

    With gHFVariables
        Call .AutoSize(0, 1, 2)
        .ColWidth(1) = Max(.ColWidth(1), 1020)
        .ColWidth(0) = .Width - .ColWidth(1) - Max(.ColWidth(2), 700)
    End With
    
    
End Sub
Private Sub OpenTLEstimating()
    Static bLoaded As Boolean
    
    Dim i As Long
    Dim ExtensionIndex   As Long
    Dim EstPrefix As String
    
    If bLoaded Then Exit Sub
    
    If mViewIndex <> 6 And mViewIndex <> 7 Then Exit Sub
    
    
    'open estimating database
    On Error Resume Next
    mFilename = ""
    mFilename = HFApp.SqlExec("select est_db_path from tbllocality where area=" & DbQuote(str, mCommunity))(0)
    If mFilename = "" Then
        EstPrefix = ""
        mFilename = HFApp.Options(TimberlineEstimatingPath)
    Else
        EstPrefix = CleanFileName(mCommunity)
    End If
    On Error GoTo 0
    Set PEAccess = PECOMServer.NewAccess
    PEAccess.DBPath = mFilename
    PEAccess.DictionaryPath = HFApp.Options(TimberlineEstimatingDictionaryPath)
    PEAccess.OpenFlags = OPEN_FLAG_READ_ONLY
    Set PEDatabase = PECOMServer.Databases.Connect(PEAccess)
    
    
    'open temp estimate file
    On Error Resume Next
    i = 0
    Do
        Err.Clear
        i = i + 1
        mFilename = AppWorkingFolder & "\" & EstPrefix & "temp" & i & ".pee"
        Set PEEstimate = PECOMServer.Estimates.Create(mFilename, PEAccess)
    Loop While Err.Number <> 0 And i < 50
    If Err.Number <> 0 Then
        Call PECOMErrHandler(SRCFILE & "Takeoff", , mFilename)
        Exit Sub
    End If

    bLoaded = True
End Sub

Private Sub LoadWBSDescriptions()
    Dim i As Long
    Dim c As Long
    Dim s As String
    Dim rs As Recordset
    
    With gItems
    
        'add columns
        c = .Cols - 1
        .Cols = .Cols + 40
        For i = 1 To 40
            .ColKey(c + i) = "WBS" & format(i, "00")
            .ColHidden(c + i) = True
        Next
        
        'read column names
        s = "select Item,custom_description Description from customDescriptions where isnull(custom_description,'')<>'' and item like 'WBS__' order by item"
        Set rs = HFApp.SqlExec(s, dbHomeFront)
        While Not rs.EOF
            c = .ColIndex("" & rs(0))
            If c <> -1 Then
                .TextMatrix(0, c) = "" & rs(1)
                .ColHidden(c) = False
            End If
            rs.MoveNext
        Wend
    
    End With

End Sub

Private Sub txtQty_Validate(Cancel As Boolean)
    txtQty.Text = Val(txtQty.Text)
End Sub



Public Function GetItemTotal(TotalType As String, POIndex As String, CostCode As String, Phase As String, Item As String) As Double
On Error GoTo eh
    Dim r As Long
    Dim v As Double
    Dim c As String
    Dim s As String
    Dim rs As Recordset
    Dim Vendor As String
    Dim Rate As Double
    
    'determine which column
    Select Case TotalType
        Case "OrderQty", "Cost":   c = "PassOrderQty"
        Case "TakeoffQty":         c = "PassTakeoffQty"
    End Select
        
    'get values
    v = 0
    With gItems
        For r = 1 To .Rows - 1
            If ((POIndex = "" Or POIndex = .TextMatrix(r, .ColIndex("POIndex"))) And _
                (CostCode = "" Or CostCode = .TextMatrix(r, .ColIndex("JCCostcode"))) And _
                (Phase = "" Or Phase = .TextMatrix(r, .ColIndex("Phase"))) And _
                (Item = "" Or Item = .TextMatrix(r, .ColIndex("Item")))) Then
            
                
                If TotalType = "cost" Then
                
                
                    'get vendor
                    If mVendor = "" Then
                        Vendor = ""
                        Set rs = HFApp.SqlExec("SELECT dbo.Purch_GetCommunityVendor(" & DbQuote(str, mCommunity) & "," & DbQuote(str, POIndex) & "," & HFApp.DivisionID & ")")
                        If Not rs.EOF Then Vendor = "" & rs(0)
                    Else
                        Vendor = mVendor
                    End If
                    
                    'get cost rate
                    s = ""
                    s = s & "SELECT dbo.Purch_GetItemRate(0,0,"
                    s = s & DbQuote(str, mCommunity) & ","
                    s = s & DbQuote(str, mCommunityPhase) & ","
                    s = s & DbQuote(str, .TextMatrix(r, .ColIndex("Assembly"))) & ","
                    s = s & DbQuote(str, .TextMatrix(r, .ColIndex("Model"))) & ","
                    s = s & "'',"
                    s = s & DbQuote(str, .TextMatrix(r, .ColIndex("Phase"))) & ","
                    s = s & DbQuote(str, .TextMatrix(r, .ColIndex("Item"))) & ","
                    s = s & "0,"
                    s = s & DbQuote(str, Vendor) & ","
                    s = s & "GETDATE()," & HFApp.DivisionID & ")"
                    Set rs = HFApp.SqlExec(s)
                    If Not rs.EOF Then Rate = Val("" & rs(0))
                    
                    v = v + Rate * Val(.TextMatrix(r, .ColIndex(c)))
                        
                Else
                    v = v + Val(.TextMatrix(r, .ColIndex(c)))
                End If
                
            End If
        Next
    End With
    
    

    GetItemTotal = v
    
Exit Function
eh: Call errHandler(SRCFILE & "GetItemTotal")
End Function

Private Sub LookupCommunityStandards()
    Dim s As String
    Dim POIndex As String
    Dim Phase As String
    Dim Item As String
    Dim rs As Recordset
    Dim r As Long
    
    With gItems
    For r = 1 To .Rows - 1
        If Not .RowHidden(r) Then

            POIndex = .TextMatrix(r, .ColIndex("POIndex"))
            Phase = .TextMatrix(r, .ColIndex("Phase"))
            Item = .TextMatrix(r, .ColIndex("Item"))


            s = ""
            s = s & "SELECT 1,i.Phase,i.Item,i.Description,i.OrderUOM,i.TakeoffUOM,i.ConversionFactor,i.Roundto,i.RoundDir,i.WastePercent,isnull(nullif(i.JCCostCode,''),isnull(p.JCCostCode,'')) JCCostCode,i.JCCategory,i.AltJCCostCode,i.AltJCCategory,s.Notes Comments,i.Price" & vbCrLf
            s = s & "  FROM CommunityStandards s" & vbCrLf
            s = s & "       left outer join tblPhaseItem i on(s.phase=i.phase and s.item=i.item)" & vbCrLf
            s = s & "      ,tblPOIndex p" & vbCrLf
            s = s & " WHERE p.DivisionID = " & HFApp.DivisionID & " and p.poindex=" & DbQuote(str, POIndex) & vbCrLf
            s = s & "   AND s.community=" & DbQuote(str, mCommunity) & vbCrLf
            s = s & "   AND s.communityPhase=" & DbQuote(str, mCommunityPhase) & vbCrLf
            s = s & "   AND s.StdPhase=" & DbQuote(str, Phase) & vbCrLf
            s = s & "   AND s.stditem=" & DbQuote(str, Item) & vbCrLf
            s = s & "   AND i.DivisionID = " & HFApp.DivisionID & vbCrLf
            s = s & "UNION ALL" & vbCrLf
            s = s & "SELECT 2,i.Phase,i.Item,i.Description,i.OrderUOM,i.TakeoffUOM,i.ConversionFactor,i.Roundto,i.RoundDir,i.WastePercent,isnull(nullif(i.JCCostCode,''),isnull(p.JCCostCode,'')) JCCostCode,i.JCCategory,i.AltJCCostCode,i.AltJCCategory,s.Notes Comments,i.Price" & vbCrLf
            s = s & "  FROM CommunityStandards s" & vbCrLf
            s = s & "       left outer join tblPhaseItem i on(s.phase=i.phase and s.item=i.item)" & vbCrLf
            s = s & "      ,tblPOIndex p" & vbCrLf
            s = s & " WHERE p.DivisionID = " & HFApp.DivisionID & " and p.poindex=" & DbQuote(str, POIndex) & vbCrLf
            s = s & "   AND s.community=" & DbQuote(str, mCommunity) & vbCrLf
            s = s & "   AND s.communityPhase=''" & vbCrLf
            s = s & "   AND s.StdPhase=" & DbQuote(str, Phase) & vbCrLf
            s = s & "   AND s.stditem=" & DbQuote(str, Item) & vbCrLf
            s = s & "   AND i.DivisionID = " & HFApp.DivisionID & vbCrLf
            s = s & "UNION ALL" & vbCrLf
            s = s & "SELECT 3,i.Phase,i.Item,i.Description,i.OrderUOM,i.TakeoffUOM,i.ConversionFactor,i.Roundto,i.RoundDir,i.WastePercent,isnull(nullif(i.JCCostCode,''),isnull(p.JCCostCode,'')) JCCostCode,i.JCCategory,i.AltJCCostCode,i.AltJCCategory,i.Notes Comments,i.Price" & vbCrLf
            s = s & "  FROM tblPhaseItem i" & vbCrLf
            s = s & "      ,tblPOIndex p" & vbCrLf
            s = s & " WHERE p.DivisionID = " & HFApp.DivisionID & " and p.poindex=" & DbQuote(str, POIndex) & vbCrLf
            s = s & "   AND i.Phase=" & DbQuote(str, Phase) & vbCrLf
            s = s & "   AND i.item=" & DbQuote(str, Item) & vbCrLf
            s = s & "   AND i.DivisionID = " & HFApp.DivisionID & vbCrLf
            s = s & "ORDER BY 1" & vbCrLf
            Set rs = HFApp.SqlExec(s)
            If Not rs.EOF Then
                If "" <> "" & rs("Phase") & rs("Item") Then
                    .TextMatrix(r, .ColIndex("Phase")) = "" & rs("Phase")
                    .TextMatrix(r, .ColIndex("Item")) = "" & rs("Item")
                    .TextMatrix(r, .ColIndex("Description")) = "" & rs("Description")
                    .TextMatrix(r, .ColIndex("OrderUOM")) = "" & rs("OrderUOM")
                    .TextMatrix(r, .ColIndex("TakeoffUOM")) = "" & rs("TakeoffUOM")
                    .TextMatrix(r, .ColIndex("ConversionFactor")) = Val("" & rs("ConversionFactor"))
                    .TextMatrix(r, .ColIndex("RoundTo")) = Val("" & rs("RoundTo"))
                    .TextMatrix(r, .ColIndex("RoundDir")) = Val("" & rs("RoundDir"))
                    .TextMatrix(r, .ColIndex("WastePercent")) = Val("" & rs("WastePercent"))
                    
                    If .TextMatrix(r, .ColIndex("Notes")) = "" Then
                        .TextMatrix(r, .ColIndex("Notes")) = "" & rs("Comments")
                    End If
                    
                    If mAssemblyType = atModel Or (mUseAltCostCodesForCO And Not mIsChangeOrder) Then
                        .TextMatrix(r, .ColIndex("JCCostCode")) = "" & rs("JCCostCode")
                        .TextMatrix(r, .ColIndex("JCCategory")) = "" & rs("JCCategory")
                    Else
                        .TextMatrix(r, .ColIndex("JCCostCode")) = IIf("" & rs("AltJCCostCode") = "", "" & rs("JCCostCode"), "" & rs("AltJCCostCode"))
                        .TextMatrix(r, .ColIndex("JCCategory")) = IIf("" & rs("AltJCCategory") = "", "" & rs("JCCategory"), "" & rs("AltJCCategory"))
                    End If
                End If
            End If
        End If
    Next
    End With
    
End Sub
