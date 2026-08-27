VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Object = "{55473EAC-7715-4257-B5EF-6E14EBD6A5DD}#1.0#0"; "vbalProgBar6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FImportAssemblies 
   Caption         =   "Homefront Assembly Import"
   ClientHeight    =   6330
   ClientLeft      =   285
   ClientTop       =   1605
   ClientWidth     =   8925
   Icon            =   "FImportAssemblies.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6330
   ScaleWidth      =   8925
   Begin VB.CheckBox chkSaveCommunity 
      Caption         =   "Save Community on Assemblies?"
      Height          =   375
      Left            =   3360
      TabIndex        =   21
      Top             =   1080
      Value           =   1  'Checked
      Width           =   2895
   End
   Begin vbalProgBarLib6.vbalProgressBar ProgressBar 
      Height          =   315
      Left            =   60
      TabIndex        =   20
      Top             =   4560
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   556
      Picture         =   "FImportAssemblies.frx":000C
      ForeColor       =   0
      BorderStyle     =   0
      BarPicture      =   "FImportAssemblies.frx":0028
      ShowText        =   -1  'True
      TextAlignX      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Segments        =   -1  'True
      XpStyle         =   -1  'True
   End
   Begin VSFlex8Ctl.VSFlexGrid gCategories 
      Height          =   2055
      Left            =   3300
      TabIndex        =   18
      Top             =   2220
      Width           =   2415
      _cx             =   1968181412
      _cy             =   1968180777
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
   Begin VSFlex8Ctl.VSFlexGrid gErrors 
      Height          =   2055
      Left            =   180
      TabIndex        =   11
      Top             =   1740
      Width           =   2415
      _cx             =   1968181412
      _cy             =   1968180777
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
      FormatString    =   $"FImportAssemblies.frx":0044
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
      Begin VB.Image imgMsg 
         Height          =   240
         Index           =   2
         Left            =   480
         Picture         =   "FImportAssemblies.frx":0081
         Top             =   0
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgMsg 
         Height          =   240
         Index           =   0
         Left            =   0
         Picture         =   "FImportAssemblies.frx":060B
         Top             =   0
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgMsg 
         Height          =   240
         Index           =   1
         Left            =   240
         Picture         =   "FImportAssemblies.frx":0B95
         Top             =   0
         Visible         =   0   'False
         Width           =   240
      End
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   315
      Index           =   0
      Left            =   5220
      TabIndex        =   13
      Top             =   120
      Width           =   1035
   End
   Begin VB.CommandButton cmdNav 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   315
      Index           =   1
      Left            =   5220
      TabIndex        =   12
      Top             =   480
      Width           =   1035
   End
   Begin VSFlex8Ctl.VSFlexGrid gPhases 
      Height          =   2055
      Left            =   3000
      TabIndex        =   8
      Top             =   1980
      Width           =   2415
      _cx             =   1968181412
      _cy             =   1968180777
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
   Begin VSFlex8Ctl.VSFlexGrid gAssemblies 
      Height          =   2055
      Left            =   2700
      TabIndex        =   6
      Top             =   1740
      Width           =   2415
      _cx             =   1968181412
      _cy             =   1968180777
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
   Begin MSComctlLib.TabStrip TabStrip 
      Height          =   3135
      Left            =   60
      TabIndex        =   7
      Top             =   1380
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   5530
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Messages"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Assemblies"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Phases"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Categories"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   180
      TabIndex        =   19
      Top             =   1140
      Width           =   45
   End
   Begin VB.Image imgIconInfo 
      Height          =   480
      Left            =   300
      Picture         =   "FImportAssemblies.frx":111F
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgIconError 
      Height          =   480
      Left            =   300
      Picture         =   "FImportAssemblies.frx":19E9
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblErrorCountLabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Errors:"
      Height          =   255
      Left            =   2880
      TabIndex        =   17
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label lblErrorCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   4140
      TabIndex        =   16
      Top             =   720
      Width           =   75
   End
   Begin VB.Label lblWarningCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4140
      TabIndex        =   15
      Top             =   900
      Width           =   45
   End
   Begin VB.Label lblWarningCountLabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Warnings:"
      Height          =   255
      Left            =   3240
      TabIndex        =   14
      Top             =   900
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Phases:"
      Height          =   195
      Index           =   4
      Left            =   1185
      TabIndex        =   10
      Top             =   900
      Width           =   570
   End
   Begin VB.Label lblPhaseCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " "
      Height          =   195
      Left            =   1800
      TabIndex        =   9
      Top             =   900
      Width           =   45
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Assemblies:"
      Height          =   195
      Index           =   1
      Left            =   930
      TabIndex        =   1
      Top             =   540
      Width           =   825
   End
   Begin VB.Label lblItemCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " "
      Height          =   195
      Left            =   1800
      TabIndex        =   5
      Top             =   720
      Width           =   45
   End
   Begin VB.Label lblModelCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " "
      Height          =   195
      Left            =   1800
      TabIndex        =   4
      Top             =   540
      Width           =   45
   End
   Begin VB.Label lblFileName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " "
      Height          =   195
      Left            =   1800
      TabIndex        =   3
      Top             =   180
      Width           =   45
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Items:"
      Height          =   195
      Index           =   2
      Left            =   1335
      TabIndex        =   2
      Top             =   720
      Width           =   420
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "File Name:"
      Height          =   195
      Index           =   0
      Left            =   1005
      TabIndex        =   0
      Top             =   180
      Width           =   750
   End
   Begin VB.Image imgIconWarning 
      Height          =   480
      Left            =   300
      Picture         =   "FImportAssemblies.frx":22B3
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "FImportAssemblies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Const SRCFILE = "FImportAssemblies::"

Private mImportMode As String

'-- for errors and warnings --
Const ERROR = 0
Const WARN = 1
Const INFO = 2
Private mHasZEROs    As Boolean
Private mErrors      As Long
Private mWarnings    As Long
Private mInfoMsgs    As Long

'-- workbook properties --
Private mTemplateVersion   As String
Private mCommunity         As String
Private mCommunityDesc     As String
Private mCommunityPhase    As String

'-- workbook and template stuff --
Private mXlBook            As Object 'Excel.Workbook
Private mXlsFileName       As String
Private mAssemblyType      As AssemblyTypes

Private mColsPerAssembly   As Long
Private mIsShortForm       As Boolean
Private mFirstItemRow      As Long
Private mLastItemRow       As Long

'-- assembly item column indexes --
'-- see LoadAssemblies()
Private POIndex As Long
Private JCCostCode As Long
Private JCCategory As Long
Private EstPhase As Long
Private EstItem As Long
Private ItemDescription As Long
Private ItemUOM As Long
Private ItemNotes As Long
Private FirstAssemblyCol As Long

'-- assembly header row indexes --
'-- see LoadAssemblies()
Private AssemblyID As Long
Private AssemblyDescription As Long
Private ModelAssemblyID As Long
Private OptionAssemblyID As Long
Private OptionPrefix As Long
Private SalesCategory As Long
Private AssemblyNotes As Long
Private JCExtra As Long
Private SellingPretax As Long
Private COSellingPretax As Long
Private SellingPriceTaxIncl As Long
Private COSellingPriceTaxIncl As Long
Private ConstructionCutoff As Long
Private AssemblyUOM As Long
Private ModelPrefix As Long
Private series As Long
Private Elevation As Long
Private FloorArea As Long
Private Bedrooms As Long
Private Bathrooms As Long
Private Style As Long


Private AssemblyNoteField As String



Public Sub WriteSeries()
On Error GoTo eh

    Dim c As Long
    Dim s As String
    Dim sSeries As String
    
    With gAssemblies
        'for each assembly
        For c = 8 To .Cols - 1 Step mColsPerAssembly
            sSeries = .TextMatrix(series, c)
            If Trim(sSeries) <> "" Then
                s = ""
                s = s & "INSERT INTO tblSeries(DivisionID,Series,Description) "
                s = s & "VALUES(" & HFApp.DivisionID & "," & DbQuote(Str, sSeries)
                s = s & "," & DbQuote(Str, sSeries) & ")"
                Call HFApp.SqlExec(s)
            End If
        Next
    End With
    
Exit Sub
eh: If InStr(1, Err.Description, "duplicate", vbTextCompare) Then
        Resume Next
    Else
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    End If
End Sub

Private Sub chkSaveCommunity_Click()
 Call ValidateData
End Sub

Private Sub cmdNav_Click(Index As Integer)
On Error GoTo eh
    Dim t As Single
    Dim s As String
    t = Timer
    Select Case Index
    
        Case 0 'ok
            Call SaveData
            Unload Me
            
        Case 1 'cancel
            Unload Me
        
    End Select
    
Exit Sub
eh: Call errHandler(SRCFILE & "cmdNav_Click")
End Sub

Private Sub SaveData()
    
    Screen.MousePointer = vbHourglass
    Call WriteCommunity
    Call WriteSeries
    Call WriteCategories
    Call WritePhases
    If mImportMode <> "Assembly" Then Call WriteItems
    If mImportMode = "Assembly" Then Call WriteAssemblies
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Load()
On Error GoTo eh

    Dim stage As String
    
    mErrors = 0
    mWarnings = 0
    mHasZEROs = False

    Me.Caption = mImportMode & " Import"

    Call IniGetForm(Me)
    
    'load datafile
stage = "EXCEL"
    If VBGetOpenFileName(mXlsFileName, , , , , , "Excel files|*.xls*|All files|*.*", , , , , Me.hwnd) Then
        Set mXlBook = GetObject(mXlsFileName)
stage = ""
        
        TabStrip.Tabs(2).Selected = True
        lblFilename = mXlsFileName
        
        Call LoadAssemblies
        Call LoadPhases
        If mAssemblyType <> atModel Then Call LoadCategories
        Call mXlBook.Close(False)
        Set mXlBook = Nothing

        'display statistics
        lblModelCount = (gAssemblies.Cols - 8) / mColsPerAssembly
        lblItemCount = mLastItemRow - mFirstItemRow + 1
        lblPhaseCount = gPhases.Rows - 2
                        
        Call ValidateData
        If Val(lblErrorCount.Caption) <> 0 Then
            TabStrip.Tabs(1).Selected = True
        End If

    Else
        Unload Me
    End If
    
Exit Sub
eh:

    Select Case True
        Case Err.Number = -2147417846:         MsgBox "Microsoft Excel is busy. It may be waiting for a response from you.", vbExclamation, App.ProductName
        Case stage = "EXCEL":                  MsgBox "Unable to load file." & vbCrLf & vbCrLf & "Microsoft Excel reports:" & vbCrLf & Err.Description, vbCritical, App.ProductName
        Case Err.Number = vbObjectError + 88:  MsgBox "Invalid file format" & vbCrLf & vbCrLf & "The """ & Err.Description & """ worksheet could not be found.", vbCritical, App.ProductName
        Case Err.Number = vbObjectError + 89:  MsgBox "Invalid file format" & vbCrLf & vbCrLf & "Version """ & Err.Description & """ is not supported", vbCritical, App.ProductName
        Case Else:                             Call errHandler("FImportAssemblies::Form_Load")
    End Select
    Unload Me
End Sub



Private Sub LoadAssemblies()
    Dim i As Long
    '--------------------------------------------------------------------------------------------------------------
    ' version 2 added another row (at row 2) for the sales code. version 1's treat it same as asssembly id
    ' version 3&4 added another column (at col 1) for poindex. versions 1&2 did not write a poindex. it was assumed based on tblPhaseItem
    '--------------------------------------------------------------------------------------------------------------
    
    Call LoadExcelSheet(mXlBook, "Assemblies", gAssemblies)
    
    Me.Show
    Me.gAssemblies.Visible = True
    Call Me.gAssemblies.ZOrder(0)
    
    'read properties
    mTemplateVersion = Trim(UCase(gAssemblies.TextMatrix(1, 1)))
    mCommunity = Trim(gAssemblies.TextMatrix(4, 2))
    mCommunityDesc = Trim(gAssemblies.TextMatrix(5, 2))
    mCommunityPhase = Trim(gAssemblies.TextMatrix(6, 2))
    
    'load the format
    'item col indexes
    Select Case Parse(mTemplateVersion, 2, ".")
        Case "3", "4"
            POIndex = 1
            JCCostCode = 2
            JCCategory = 3
            EstPhase = 4
            EstItem = 5
            ItemDescription = 6
            ItemUOM = 7
            ItemNotes = 8
            FirstAssemblyCol = 9
        Case Else
            POIndex = -1 'poindex added in version 3&4 formats
            JCCostCode = 1
            JCCategory = 2
            EstPhase = 3
            EstItem = 4
            ItemDescription = 5
            ItemUOM = 6
            ItemNotes = 7
            FirstAssemblyCol = 8
    End Select
    
    'assembly row indexes
    AssemblyID = -1
    AssemblyDescription = -1
    ModelAssemblyID = -1
    OptionAssemblyID = -1
    OptionPrefix = -1
    ModelPrefix = -1
    SalesCategory = -1
    AssemblyNotes = -1
    JCExtra = -1
    SellingPretax = -1
    COSellingPretax = -1
    SellingPriceTaxIncl = -1
    COSellingPriceTaxIncl = -1
    ConstructionCutoff = -1
    AssemblyUOM = -1
    series = -1
    Elevation = -1
    FloorArea = -1
    Bedrooms = -1
    Bathrooms = -1
    Style = -1
    
    Select Case mTemplateVersion
        'Models
        Case "MS.1", "ML.1", "MS.3", "ML.3"
            mAssemblyType = atModel
            mColsPerAssembly = IIf(Mid(mTemplateVersion, 2, 1) = "S", 1, 3)
            mIsShortForm = mColsPerAssembly = 1
            mFirstItemRow = 14
            '-- assembly header row indexes --
            AssemblyID = 1
            ModelAssemblyID = 1
            AssemblyDescription = 2
            ModelPrefix = 3
            series = 4
            Elevation = 5
            AssemblyNotes = 6
            SellingPretax = 7
            SellingPriceTaxIncl = 8
            FloorArea = 9
            Bedrooms = 10
            Bathrooms = 11
            Style = 12
                   
        'Options
        Case "OS.1", "OL.1", "OS.3", "OL.3":
            mAssemblyType = atoption
            mColsPerAssembly = IIf(Mid(mTemplateVersion, 2, 1) = "S", 1, 3)
            mIsShortForm = mColsPerAssembly = 1
            mFirstItemRow = 16
            '-- assembly header row indexes --
            AssemblyID = 1
            OptionAssemblyID = 1
            AssemblyDescription = 2
            ModelAssemblyID = 3
            series = 4
            OptionPrefix = 5
            SalesCategory = 6
            AssemblyNotes = 7
            JCExtra = 8
            SellingPretax = 9
            COSellingPretax = 10
            SellingPriceTaxIncl = 11
            COSellingPriceTaxIncl = 12
            ConstructionCutoff = 13
            AssemblyUOM = 14
                                              
        'Global Options
        Case "GS.1", "GL.1", "GS.3", "GL.3"
            mAssemblyType = atGlobal
            mColsPerAssembly = IIf(Mid(mTemplateVersion, 2, 1) = "S", 1, 3)
            mIsShortForm = mColsPerAssembly = 1
            mFirstItemRow = 14
            '-- assembly header row indexes --
            AssemblyID = 1
            OptionAssemblyID = 1
            AssemblyDescription = 2
            OptionPrefix = 3
            SalesCategory = 4
            AssemblyNotes = 5
            JCExtra = 6
            SellingPretax = 7
            COSellingPretax = 8
            SellingPriceTaxIncl = 9
            COSellingPriceTaxIncl = 10
            ConstructionCutoff = 11
            AssemblyUOM = 12
            
        'Design Center Options
        Case "DS.1", "DL.1", "DS.3", "DL.3"
            mAssemblyType = atDesignCenter
            mColsPerAssembly = IIf(Mid(mTemplateVersion, 2, 1) = "S", 1, 3)
            mIsShortForm = mColsPerAssembly = 1
            mFirstItemRow = 14
            '-- assembly header row indexes --
            AssemblyID = 1
            OptionAssemblyID = 1
            AssemblyDescription = 2
            OptionPrefix = 3
            SalesCategory = 4
            AssemblyNotes = 5
            JCExtra = 6
            SellingPretax = 7
            COSellingPretax = 8
            SellingPriceTaxIncl = 9
            COSellingPriceTaxIncl = 10
            ConstructionCutoff = 11
            AssemblyUOM = 12
            
        '--------------------------------------------------------------------------------------------------------------
        ' version 2's include another row for the sales code. version 1's treat it same as asssembly id
        '--------------------------------------------------------------------------------------------------------------
        'models
        Case "MS.2", "ML.2", "MS.4", "ML.4"
            mAssemblyType = atModel
            mColsPerAssembly = IIf(Mid(mTemplateVersion, 2, 1) = "S", 1, 3)
            mIsShortForm = mColsPerAssembly = 1
            mFirstItemRow = 15
            '-- assembly header row indexes --
            AssemblyID = 1
            ModelAssemblyID = 2
            AssemblyDescription = 3
            ModelPrefix = 4
            series = 5
            Elevation = 6
            AssemblyNotes = 7
            SellingPretax = 8
            SellingPriceTaxIncl = 9
            FloorArea = 10
            Bedrooms = 11
            Bathrooms = 12
            Style = 13
            
        'Options
        Case "OS.2", "OL.2", "OS.4", "OL.4"
            mAssemblyType = atoption
            mColsPerAssembly = IIf(Mid(mTemplateVersion, 2, 1) = "S", 1, 3)
            mIsShortForm = mColsPerAssembly = 1
            mFirstItemRow = 17
            '-- assembly header row indexes --
            AssemblyID = 1
            OptionAssemblyID = 2
            AssemblyDescription = 3
            ModelAssemblyID = 4
            series = 5
            OptionPrefix = 6
            SalesCategory = 7
            AssemblyNotes = 8
            JCExtra = 9
            SellingPretax = 10
            COSellingPretax = 11
            SellingPriceTaxIncl = 12
            COSellingPriceTaxIncl = 13
            ConstructionCutoff = 14
            AssemblyUOM = 15
                       
        'Global Options
        Case "GS.2", "GL.2", "GS.4", "GL.4"
            mAssemblyType = atGlobal
            mColsPerAssembly = IIf(Mid(mTemplateVersion, 2, 1) = "S", 1, 3)
            mIsShortForm = mColsPerAssembly = 1
            mFirstItemRow = 15
            '-- assembly header row indexes --
            AssemblyID = 1
            OptionAssemblyID = 2
            AssemblyDescription = 3
            OptionPrefix = 4
            SalesCategory = 5
            AssemblyNotes = 6
            JCExtra = 7
            SellingPretax = 8
            COSellingPretax = 9
            SellingPriceTaxIncl = 10
            COSellingPriceTaxIncl = 11
            ConstructionCutoff = 12
            AssemblyUOM = 13
            
        'Design Center Options
        Case "DS.2", "DL.2", "DS.4", "DL.4"
            mAssemblyType = atDesignCenter
            mColsPerAssembly = IIf(Mid(mTemplateVersion, 2, 1) = "S", 1, 3)
            mIsShortForm = mColsPerAssembly = 1
            mFirstItemRow = 15
            '-- assembly header row indexes --
            AssemblyID = 1
            OptionAssemblyID = 2
            AssemblyDescription = 3
            OptionPrefix = 4
            SalesCategory = 5
            AssemblyNotes = 6
            JCExtra = 7
            SellingPretax = 8
            COSellingPretax = 9
            SellingPriceTaxIncl = 10
            COSellingPriceTaxIncl = 11
            ConstructionCutoff = 12
            AssemblyUOM = 13
            
        Case Else
            Call Err.Raise(vbObjectError + 89, , Trim(UCase(gAssemblies.TextMatrix(1, 1))))
            
    End Select
   
    'trim off trailing columns
    For i = gAssemblies.Cols - 1 To 0 Step -1
        If Trim(gAssemblies.TextMatrix(1, i)) <> "" Then
            Exit For
        End If
    Next
    gAssemblies.Cols = i + mColsPerAssembly
    
    
    'trim off trailing rows
    For i = gAssemblies.Rows - 1 To mFirstItemRow Step -1
        If Trim(gAssemblies.TextMatrix(i, EstPhase)) <> "" Or Trim(gAssemblies.TextMatrix(i, EstItem)) <> "" Then
            Exit For
        End If
    Next
    gAssemblies.Rows = i + 1
    
    'remove embedded comment rows - anything with no EstPhase and no EstItem
    For i = gAssemblies.Rows - 1 To mFirstItemRow Step -1
        If Trim(gAssemblies.TextMatrix(i, EstPhase)) = "" And Trim(gAssemblies.TextMatrix(i, EstItem)) = "" Then
            Call gAssemblies.RemoveItem(i)
        End If
    Next
    
    
    'add cost type M to any EstItem that has no cost type
    For i = gAssemblies.Rows - 1 To mFirstItemRow Step -1
        If Not IsIn(UCase(Right(Trim(gAssemblies.TextMatrix(i, EstItem)), 1)), "M", "S", "L", "E", "O") Then
            gAssemblies.TextMatrix(i, EstItem) = Trim(gAssemblies.TextMatrix(i, EstItem)) & "M"
        End If
    Next
    
    
    
    'if header label says "note" then write to estimatornotes else write to comments
    If gAssemblies.TextMatrix(AssemblyNotes, 8) = "Notes" Then
        AssemblyNoteField = "Notes"
    Else
        AssemblyNoteField = "Comments"
    End If
    
    
    'add an extra row and assign unused assembly headers to it.
    mLastItemRow = gAssemblies.Rows - 1
    gAssemblies.AddItem gAssemblies.ValueMatrix(mLastItemRow, 0) + 1
    
    
    If AssemblyID = -1 Then AssemblyID = gAssemblies.Rows - 1
    If AssemblyDescription = -1 Then AssemblyDescription = gAssemblies.Rows - 1
    If OptionPrefix = -1 Then OptionPrefix = gAssemblies.Rows - 1
    If ModelPrefix = -1 Then ModelPrefix = gAssemblies.Rows - 1
    If ModelAssemblyID = -1 Then ModelAssemblyID = gAssemblies.Rows - 1
    If OptionAssemblyID = -1 Then OptionAssemblyID = gAssemblies.Rows - 1
    If SalesCategory = -1 Then SalesCategory = gAssemblies.Rows - 1
    If AssemblyNotes = -1 Then AssemblyNotes = gAssemblies.Rows - 1
    If JCExtra = -1 Then JCExtra = gAssemblies.Rows - 1
    If SellingPretax = -1 Then SellingPretax = gAssemblies.Rows - 1
    If COSellingPretax = -1 Then COSellingPretax = gAssemblies.Rows - 1
    If SellingPriceTaxIncl = -1 Then SellingPriceTaxIncl = gAssemblies.Rows - 1
    If COSellingPriceTaxIncl = -1 Then COSellingPriceTaxIncl = gAssemblies.Rows - 1
    If ConstructionCutoff = -1 Then ConstructionCutoff = gAssemblies.Rows - 1
    If AssemblyUOM = -1 Then AssemblyUOM = gAssemblies.Rows - 1
    If series = -1 Then series = gAssemblies.Rows - 1
    If Elevation = -1 Then Elevation = gAssemblies.Rows - 1
    If FloorArea = -1 Then FloorArea = gAssemblies.Rows - 1
    If Bedrooms = -1 Then Bedrooms = gAssemblies.Rows - 1
    If Bathrooms = -1 Then Bathrooms = gAssemblies.Rows - 1
    If Style = -1 Then Style = gAssemblies.Rows - 1
    
        
    'color formatting
    gAssemblies.Cell(flexcpBackColor, 1, 1, mFirstItemRow - 1, ItemNotes) = vbCyan 'vbButtonFace
    gAssemblies.Cell(flexcpFontBold, 1, ItemNotes, mFirstItemRow - 1, ItemNotes) = True
    gAssemblies.Cell(flexcpAlignment, 1, ItemNotes, mFirstItemRow - 1, ItemNotes) = flexAlignRightCenter
    
    gAssemblies.Cell(flexcpBackColor, mFirstItemRow - 1, 1, mFirstItemRow - 1, gAssemblies.Cols - 1) = vbCyan 'vbButtonFace
    gAssemblies.Cell(flexcpFontBold, mFirstItemRow - 1, 1, mFirstItemRow - 1, gAssemblies.Cols - 1) = True
    
    gAssemblies.Cell(flexcpFontBold, 3, 1, 6, 1) = True
    gAssemblies.Cell(flexcpAlignment, 3, 1, 6, 1) = flexAlignRightCenter
    gAssemblies.Cell(flexcpBackColor, 3, 2, 6, 2) = vbWindowBackground
    

On Error Resume Next
    If mColsPerAssembly > 1 Then
        For i = FirstAssemblyCol To gAssemblies.Cols - 1 Step (2 * mColsPerAssembly)
            gAssemblies.Cell(flexcpBackColor, 1, i, gAssemblies.Rows - 1, i + mColsPerAssembly - 1) = vbButtonFace
            Me.gAssemblies.ColWidth(i) = 2700
            Me.gAssemblies.ColWidth(i + 3) = 2700
        Next
    End If
    
End Sub

Private Sub LoadPhases()
    Dim i As Long
    
    Call LoadExcelSheet(mXlBook, "Phases", gPhases)
    
    'trim off trailing rows
    For i = gPhases.Rows - 1 To 0 Step -1
        If Trim(gPhases.TextMatrix(i, 1)) <> "" Or Trim(gPhases.TextMatrix(i, 2)) <> "" Then
            Exit For
        End If
    Next
    gPhases.Rows = i + 1
            
    'trim off trailing columns
    gPhases.Cols = 4
        
    'remove embedded comment rows - anything with no "Group Phase" and no "Phase"
    For i = gPhases.Rows - 1 To 1 Step -1
        If gPhases.TextMatrix(i, 1) = "" And gPhases.TextMatrix(i, 2) = "" Then
            Call gPhases.RemoveItem(i)
        End If
    Next

    'formatting
    On Error Resume Next
    gPhases.Cell(flexcpBackColor, 1, 1, 1, gPhases.Cols - 1) = vbButtonFace
    gPhases.Cell(flexcpFontBold, 1, 1, 1, gPhases.Cols - 1) = True

End Sub


Private Sub LoadCategories()
    Dim i As Long
    
    Call LoadExcelSheet(mXlBook, "Categories", gCategories)
    
    'trim off trailing rows
    For i = gCategories.Rows - 1 To 0 Step -1
        If Trim(gCategories.TextMatrix(i, 1)) <> "" Then
            Exit For
        End If
    Next
    gCategories.Rows = i + 1
            
    'trim off trailing columns
    gCategories.Cols = 5
        
    
    'remove embedded comment rows - anything with no "Category"
    For i = gCategories.Rows - 1 To 1 Step -1
        If gCategories.TextMatrix(i, 1) = "" Then
            Call gCategories.RemoveItem(i)
        End If
    Next

    'formatting
    On Error Resume Next
    gCategories.Cell(flexcpBackColor, 1, 1, 1, gCategories.Cols - 1) = vbButtonFace
    gCategories.Cell(flexcpFontBold, 1, 1, 1, gCategories.Cols - 1) = True
    Call gCategories.AutoSize(1, 4)
End Sub


Private Sub Form_Resize()
On Error Resume Next
    Const margin = 60
    
    If Me.Height = 1935 Then
        cmdNav(2).Caption = "Details >>"
    Else
        cmdNav(2).Caption = "<< Details"
    End If
    
    cmdNav(0).Left = Me.ScaleWidth - cmdNav(0).Width - cmdNav(0).Top
    cmdNav(1).Left = cmdNav(0).Left
    cmdNav(2).Left = cmdNav(0).Left
    chkSaveCommunity.Left = cmdNav(1).Left - 1860
    
    TabStrip.Move margin, TabStrip.Top, Me.ScaleWidth - 2 * margin, Me.ScaleHeight - TabStrip.Top - 2 * margin - ProgressBar.Height
    ProgressBar.Move margin, Me.ScaleHeight - Me.ProgressBar.Height - margin, TabStrip.Width
    
    Call TabStrip_Click
End Sub

Private Sub ValidateData()
On Error GoTo eh
    Dim s As String
    Dim r As Long
    Dim c As Long
    Dim rs As Recordset
    Dim po As String
    Dim PhaseCSV     As String
    Dim AssemblyCSV  As String
    Dim bQtyDefaultsToOne As Boolean

    gErrors.Rows = 0
    mErrors = 0
    mWarnings = 0
    If mImportMode = "Assembly" Then

        'check that community exists when needed
        If (mAssemblyType = atModel And HFApp.Options(ModelByArea)) Or _
           (mAssemblyType = atoption And HFApp.Options(OptionByArea)) Or _
           (mAssemblyType = atGlobal And HFApp.Options(GlobalOptionByArea)) Or _
           (mAssemblyType = atDesignCenter And HFApp.Options(DCOptionByArea)) Then
                      
            If mCommunity <> "" Then
                s = "SELECT * FROM tblLocality WHERE Area=" & DbQuote(Str, mCommunity)
                Set rs = HFApp.SqlExec(s)
                If rs.EOF Then AddMsg ERROR, "Unknown Community."
            End If
        End If
    
        'check that community phase exists when needed
        If (mAssemblyType = atModel And HFApp.Options(ModelsByArea_Phase)) Or _
           (mAssemblyType = atoption And HFApp.Options(OptionByAreaPhase)) Or _
           (mAssemblyType = atGlobal And HFApp.Options(GlobalOptionByAreaPhase)) Or _
           (mAssemblyType = atDesignCenter And HFApp.Options(DCOptionByAreaPhase)) Then
           
            If mCommunityPhase <> "" Then
                s = "SELECT * FROM CommunityPhase WHERE Community=" & DbQuote(Str, mCommunity) & " AND CommunityPhase=" & DbQuote(Str, mCommunityPhase)
                Set rs = HFApp.SqlExec(s)
                If rs.EOF Then AddMsg ERROR, "Unknown Community Phase."
            End If
        End If
    
        'check that default vendors exist
        s = ""
        s = s & "select p.poindex" & vbCrLf
        s = s & "  from tblpoindex p" & vbCrLf
        s = s & "       left outer join poareavendor g on (p.DivisionID = g.DivisionID and p.poindex=g.poindex and isnull(g.area,'')='')" & vbCrLf
        s = s & "       left outer join poareavendor c on (p.DivisionID = c.DivisionID and p.poindex=c.poindex and c.area=" & DbQuote(Str, mCommunity) & ")" & vbCrLf
        s = s & " where g.vendor is null " & vbCrLf
        s = s & "   and c.vendor is null" & vbCrLf
        s = s & " and p.DivisionID = " & HFApp.DivisionID
        Set rs = HFApp.SqlExec(s)
        While Not rs.EOF
            AddMsg WARN, "No vendor for """ & rs(0) & """ has been specified. Vendor pricing cannot be written for items on this PO."
            rs.MoveNext
        Wend
    

        'check that all items exist and are assigned to a po index and that poindexes are valid
        With gAssemblies
            For r = mFirstItemRow To mLastItemRow
                
                'if poindex column is used
                If POIndex = -1 Then
                    po = ""
                Else
                    po = Trim(.TextMatrix(r, POIndex))
                End If
                
                
                s = "SELECT POIndex FROM tblPhaseItem WHERE DivisionID = " & HFApp.DivisionID & " and Phase=" & DbQuote(Str, .TextMatrix(r, EstPhase)) & " AND Item=" & DbQuote(Str, .TextMatrix(r, EstItem))
                Set rs = HFApp.SqlExec(s)
                If rs.EOF Then
                    AddMsg ERROR, CellAddress(r, EstItem) & " - unknown item " & vbQuote & .TextMatrix(r, EstPhase) & "/" & .TextMatrix(r, EstItem) & " " & .TextMatrix(r, ItemDescription) & vbQuote
                Else
                    If "" & rs(0) = "" And POIndex = -1 Then AddMsg WARN, vbQuote & .TextMatrix(r, EstPhase) & "/" & .TextMatrix(r, EstItem) & " " & .TextMatrix(r, ItemDescription) & vbQuote & " is not assigned to a PO. Vendor pricing can't be written."
                End If
            
            
                'validate given po
                If po <> "" Then
                    s = "SELECT POIndex FROM tblPOIndex WHERE DivisionID = " & HFApp.DivisionID & " and poindex=" & DbQuote(Str, po, , True)
                    Set rs = HFApp.SqlExec(s)
                    If rs.EOF Then
                        AddMsg ERROR, CellAddress(r, POIndex) & " - unknown poindex" & vbQuote & po & vbQuote
                    End If
                End If
            
            Next
        End With
        
    End If

    'for each phase
    PhaseCSV = ""
    If gPhases.Rows > 2 Then
        For r = 2 To gPhases.Rows - 1
           If Trim(gPhases.TextMatrix(r, 1)) <> "" And Trim(gPhases.TextMatrix(r, 2)) <> "" Then
                AddMsg ERROR, "Phases Row " & gPhases.TextMatrix(r, 0) & " - can't be both a phase and a group phase."
            End If
            'check for duplicate phases
            If InStr(1, PhaseCSV & Chr(1), Chr(1) & gPhases.TextMatrix(r, 1) + gPhases.TextMatrix(r, 2) & Chr(1)) Then
                AddMsg ERROR, "Phases must be unique - """ & gPhases.TextMatrix(r, 1) + gPhases.TextMatrix(r, 2) & """ used more than once."
            End If
            PhaseCSV = Chr(1) & gPhases.TextMatrix(r, 1) + gPhases.TextMatrix(r, 2)
        Next
    End If
    

    
    With gAssemblies
        
        '-- check for duplicate assemblies -------------------------------------------------------------------------------
        AssemblyCSV = ""
        For c = FirstAssemblyCol To .Cols - 1 Step mColsPerAssembly
        
            If InStr(1, AssemblyCSV & vbCrLf, vbCrLf & .TextMatrix(AssemblyID, c) & vbTab & .TextMatrix(ModelAssemblyID, c) & vbTab & .TextMatrix(OptionAssemblyID, c) & vbCrLf) Then
                AddMsg ERROR, "Assembly/Model/Option must be unique. " & .TextMatrix(AssemblyID, c) & " / " & .TextMatrix(ModelAssemblyID, c) & " / " & .TextMatrix(OptionAssemblyID, c) & " is specified multiple times."
            End If
            AssemblyCSV = AssemblyCSV & vbCrLf & .TextMatrix(AssemblyID, c) & vbTab & .TextMatrix(ModelAssemblyID, c) & vbTab & .TextMatrix(OptionAssemblyID, c)
        
            'now check database for duplicates
            s = ""
            s = s & "SELECT Community,Assembly,Model,OptionID" & vbCrLf
            s = s & "  FROM tblDBAssemblyMaster" & vbCrLf
            s = s & " WHERE AssemblyType=" & mAssemblyType & vbCrLf
            Select Case mAssemblyType
                Case atModel:      s = s & "   AND Community=" & DbQuote(Str, IIf(HFApp.Options(ModelByArea), IIf(Me.chkSaveCommunity, mCommunity, ""), "")) & vbCrLf
                Case atoption:     s = s & "   AND Community=" & DbQuote(Str, IIf(HFApp.Options(OptionByArea), IIf(Me.chkSaveCommunity, mCommunity, ""), "")) & vbCrLf
                Case atGlobal:     s = s & "   AND Community=" & DbQuote(Str, IIf(HFApp.Options(GlobalOptionByArea), IIf(Me.chkSaveCommunity, mCommunity, ""), "")) & vbCrLf
                Case atDesignCenter:   s = s & "   AND Community=" & DbQuote(Str, IIf(HFApp.Options(DCOptionByArea), IIf(Me.chkSaveCommunity, mCommunity, ""), "")) & vbCrLf
            End Select
            s = s & "   AND Assembly=" & DbQuote(Str, .TextMatrix(AssemblyID, c)) & vbCrLf
            s = s & "   AND Model=" & DbQuote(Str, .TextMatrix(ModelAssemblyID, c)) & vbCrLf
            s = s & "   AND OptionID=" & DbQuote(Str, .TextMatrix(OptionAssemblyID, c)) & vbCrLf
            s = s & "   AND DivisionID = " & HFApp.DivisionID
            Set rs = HFApp.SqlExec(s)
            If Not rs.EOF Then
                AddMsg WARN, CellAddress(AssemblyID, c) & " - Duplicate Assembly. The database already contains this assembly. It will be replaced."
            End If
            
        Next
        
        
        '-- check models for errors -------------------------------------------------------------------------------
        If mAssemblyType = atModel Then
            AssemblyCSV = ""
            For c = FirstAssemblyCol To .Cols - 1 Step mColsPerAssembly
    
    
                
                'is required
                r = series:              If Trim(.TextMatrix(r, c)) = "" Then AddMsg ERROR, CellAddress(r, c) & " - series is required."
            
                'is too big
                r = AssemblyID:          If Len(Trim(.TextMatrix(r, c))) > 20 Then AddMsg ERROR, CellAddress(r, c) & " - Assembly ID is too long (max 20)"
                r = AssemblyDescription: If Len(Trim(.TextMatrix(r, c))) > 50 Then AddMsg ERROR, CellAddress(r, c) & " - Model Description is too long (max 50)"
                r = ModelAssemblyID:     If Len(Trim(.TextMatrix(r, c))) > 20 Then AddMsg ERROR, CellAddress(r, c) & " - Model Assembly ID is too long (max 20)"
                r = series:              If Len(Trim(.TextMatrix(r, c))) > 10 Then AddMsg ERROR, CellAddress(r, c) & " - Series is too long (max 20)"
                r = Elevation:           If Len(Trim(.TextMatrix(r, c))) > 20 Then AddMsg ERROR, CellAddress(r, c) & " - Elevation is too long (max 20)"
                r = AssemblyNotes:       If Len(Trim(.TextMatrix(r, c))) > 4000 Then AddMsg ERROR, CellAddress(r, c) & " - Notes are too long (max 4000)"
                r = Style:               If Len(Trim(.TextMatrix(r, c))) > 30 Then AddMsg ERROR, CellAddress(r, c) & " - Style is too long (max 30)"
                
                'is number
                r = SellingPretax:         If Trim(.TextMatrix(r, c)) <> "" And Not IsNumeric(.TextMatrix(r, c)) Then AddMsg ERROR, CellAddress(r, c) & " - selling price (pretax) is not a number."
                r = SellingPriceTaxIncl:   If Trim(.TextMatrix(r, c)) <> "" And Not IsNumeric(.TextMatrix(r, c)) Then AddMsg ERROR, CellAddress(r, c) & " - selling price (incl tax) is not a number."
                r = FloorArea:             If Trim(.TextMatrix(r, c)) <> "" And Not IsNumeric(.TextMatrix(r, c)) Then AddMsg ERROR, CellAddress(r, c) & " - floor area is not a number."
                r = Bedrooms:              If Trim(.TextMatrix(r, c)) <> "" And Not IsNumeric(.TextMatrix(r, c)) Then AddMsg ERROR, CellAddress(r, c) & " - bedrooms is not a number."
                r = Bathrooms:             If Trim(.TextMatrix(r, c)) <> "" And Not IsNumeric(.TextMatrix(r, c)) Then AddMsg ERROR, CellAddress(r, c) & " - bathrooms is not a number."
                
                'is zero
                r = SellingPretax:         If Trim(.TextMatrix(r, c)) <> "" And .ValueMatrix(r, c) = 0 Then AddMsg WARN, CellAddress(r, c) & " - selling price (pretax) is zero."
                r = SellingPriceTaxIncl:   If Trim(.TextMatrix(r, c)) <> "" And .ValueMatrix(r, c) = 0 Then AddMsg WARN, CellAddress(r, c) & " - selling price (incl tax) is zero."
                r = FloorArea:             If Trim(.TextMatrix(r, c)) <> "" And .ValueMatrix(r, c) = 0 Then AddMsg WARN, CellAddress(r, c) & " - floor area is zero."
                r = Bedrooms:              If Trim(.TextMatrix(r, c)) <> "" And .ValueMatrix(r, c) = 0 Then AddMsg WARN, CellAddress(r, c) & " - bedrooms is zero."
                r = Bathrooms:             If Trim(.TextMatrix(r, c)) <> "" And .ValueMatrix(r, c) = 0 Then AddMsg WARN, CellAddress(r, c) & " - bathrooms is zero."
            
            Next
        End If
    
    
        '-- check options for errors -------------------------------------------------------------------------------
        If mAssemblyType = atoption Then
            AssemblyCSV = ""
            For c = FirstAssemblyCol To .Cols - 1 Step mColsPerAssembly
    
                'check const_cutoff
                If .ValueMatrix(ConstructionCutoff, c) <> 0 Then
                    Set rs = HFApp.SqlExec("SELECT * FROM tblJobStatus WHERE Job_Status=" & .ValueMatrix(ConstructionCutoff, c))
                    If rs.EOF Then AddMsg WARN, CellAddress(ConstructionCutoff, c) & " - Construction cutoff """ & .ValueMatrix(ConstructionCutoff, c) & """ not found. It will be created."
                End If
                
                'is too big
                r = AssemblyID:          If Len(Trim(.TextMatrix(r, c))) > 20 Then AddMsg ERROR, CellAddress(r, c) & " - Assembly ID is too long (max 20)"
                r = AssemblyDescription: If Len(Trim(.TextMatrix(r, c))) > 200 Then AddMsg ERROR, CellAddress(r, c) & " - Model Description is too long (max 50)"
                r = ModelAssemblyID:     If Len(Trim(.TextMatrix(r, c))) > 20 Then AddMsg ERROR, CellAddress(r, c) & " - Model ID is too long (max 20)"
                r = series:              If Len(Trim(.TextMatrix(r, c))) > 10 Then AddMsg ERROR, CellAddress(r, c) & " - Series is too long (max 20)"
                r = SalesCategory:       If Len(Trim(.TextMatrix(r, c))) > 15 Then AddMsg ERROR, CellAddress(r, c) & " - Category is too long (max 15)"
                r = AssemblyNotes:       If Len(Trim(.TextMatrix(r, c))) > 4000 Then AddMsg ERROR, CellAddress(r, c) & " - Notes are too long (max 4000)"
                r = JCExtra:             If Len(Trim(.TextMatrix(r, c))) > 10 Then AddMsg ERROR, CellAddress(r, c) & " - JC Extra is too long (max 10)"
                r = AssemblyUOM:         If Len(Trim(.TextMatrix(r, c))) > 10 Then AddMsg ERROR, CellAddress(r, c) & " - Sales UOM is too long (max 10)"
                
                'is number
                r = SellingPretax:         If Trim(.TextMatrix(r, c)) <> "" And Not IsNumeric(.TextMatrix(r, c)) Then AddMsg ERROR, CellAddress(r, c) & " - selling price (pretax) is not a number."
                r = SellingPriceTaxIncl:   If Trim(.TextMatrix(r, c)) <> "" And Not IsNumeric(.TextMatrix(r, c)) Then AddMsg ERROR, CellAddress(r, c) & " - selling price (incl tax) is not a number."
                r = COSellingPretax:       If Trim(.TextMatrix(r, c)) <> "" And Not IsNumeric(.TextMatrix(r, c)) Then AddMsg ERROR, CellAddress(r, c) & " - co selling price (pretax) is not a number."
                r = COSellingPriceTaxIncl: If Trim(.TextMatrix(r, c)) <> "" And Not IsNumeric(.TextMatrix(r, c)) Then AddMsg ERROR, CellAddress(r, c) & " - co selling price (incl tax) is not a number."
                r = ConstructionCutoff:    If Trim(.TextMatrix(r, c)) <> "" And Not IsNumeric(.TextMatrix(r, c)) Then AddMsg ERROR, CellAddress(r, c) & " - const cutoff is not a number."
                
                'is not zero
                r = SellingPretax:         If Trim(.TextMatrix(r, c)) <> "" And .ValueMatrix(r, c) = 0 Then AddMsg WARN, CellAddress(r, c) & " - selling price (pretax) is zero."
                r = SellingPriceTaxIncl:   If Trim(.TextMatrix(r, c)) <> "" And .ValueMatrix(r, c) = 0 Then AddMsg WARN, CellAddress(r, c) & " - selling price (incl tax) is zero."
                r = COSellingPretax:       If Trim(.TextMatrix(r, c)) <> "" And .ValueMatrix(r, c) = 0 Then AddMsg WARN, CellAddress(r, c) & " - co selling price (pretax) is zero."
                r = COSellingPriceTaxIncl: If Trim(.TextMatrix(r, c)) <> "" And .ValueMatrix(r, c) = 0 Then AddMsg WARN, CellAddress(r, c) & " - co selling price (incl tax) is zero."
                r = ConstructionCutoff:    If Trim(.TextMatrix(r, c)) <> "" And .ValueMatrix(r, c) = 0 Then AddMsg WARN, CellAddress(r, c) & " - const cutoff is zero."
            
            Next
        End If
        
        '-- check globals for errors -------------------------------------------------------------------------------
        If mAssemblyType = atGlobal Then
            AssemblyCSV = ""
            For c = FirstAssemblyCol To .Cols - 1 Step mColsPerAssembly
    
                'check const_cutoff
                If .ValueMatrix(ConstructionCutoff, c) <> 0 Then
                    Set rs = HFApp.SqlExec("SELECT * FROM tblJobStatus WHERE Job_Status=" & .ValueMatrix(ConstructionCutoff, c))
                    If rs.EOF Then AddMsg WARN, CellAddress(ConstructionCutoff, c) & " - Construction cutoff """ & .ValueMatrix(ConstructionCutoff, c) & """ not found. It will be created."
                End If
                
                'is too big
                r = AssemblyID:          If Len(Trim(.TextMatrix(r, c))) > 20 Then AddMsg ERROR, CellAddress(r, c) & " - Assembly ID is too long (max 20)"
                r = AssemblyDescription: If Len(Trim(.TextMatrix(r, c))) > 200 Then AddMsg ERROR, CellAddress(r, c) & " - Model Description is too long (max 50)"
                r = SalesCategory:       If Len(Trim(.TextMatrix(r, c))) > 15 Then AddMsg ERROR, CellAddress(r, c) & " - Category is too long (max 15)"
                r = AssemblyNotes:       If Len(Trim(.TextMatrix(r, c))) > 4000 Then AddMsg ERROR, CellAddress(r, c) & " - Notes are too long (max 4000)"
                r = JCExtra:             If Len(Trim(.TextMatrix(r, c))) > 10 Then AddMsg ERROR, CellAddress(r, c) & " - JC Extra is too long (max 10)"
                r = AssemblyUOM:         If Len(Trim(.TextMatrix(r, c))) > 10 Then AddMsg ERROR, CellAddress(r, c) & " - Sales UOM is too long (max 10)"
                
                'is number
                r = SellingPretax:         If Trim(.TextMatrix(r, c)) <> "" And Not IsNumeric(.TextMatrix(r, c)) Then AddMsg ERROR, CellAddress(r, c) & " - selling price (pretax) is not a number."
                r = SellingPriceTaxIncl:   If Trim(.TextMatrix(r, c)) <> "" And Not IsNumeric(.TextMatrix(r, c)) Then AddMsg ERROR, CellAddress(r, c) & " - selling price (incl tax) is not a number."
                r = COSellingPretax:       If Trim(.TextMatrix(r, c)) <> "" And Not IsNumeric(.TextMatrix(r, c)) Then AddMsg ERROR, CellAddress(r, c) & " - co selling price (pretax) is not a number."
                r = COSellingPriceTaxIncl: If Trim(.TextMatrix(r, c)) <> "" And Not IsNumeric(.TextMatrix(r, c)) Then AddMsg ERROR, CellAddress(r, c) & " - co selling price (incl tax) is not a number."
                r = ConstructionCutoff:    If Trim(.TextMatrix(r, c)) <> "" And Not IsNumeric(.TextMatrix(r, c)) Then AddMsg ERROR, CellAddress(r, c) & " - const cutoff is not a number."
                
                'is not zero
                r = SellingPretax:         If Trim(.TextMatrix(r, c)) <> "" And .ValueMatrix(r, c) = 0 Then AddMsg WARN, CellAddress(r, c) & " - selling price (pretax) is zero."
                r = SellingPriceTaxIncl:   If Trim(.TextMatrix(r, c)) <> "" And .ValueMatrix(r, c) = 0 Then AddMsg WARN, CellAddress(r, c) & " - selling price (incl tax) is zero."
                r = COSellingPretax:       If Trim(.TextMatrix(r, c)) <> "" And .ValueMatrix(r, c) = 0 Then AddMsg WARN, CellAddress(r, c) & " - co selling price (pretax) is zero."
                r = COSellingPriceTaxIncl: If Trim(.TextMatrix(r, c)) <> "" And .ValueMatrix(r, c) = 0 Then AddMsg WARN, CellAddress(r, c) & " - co selling price (incl tax) is zero."
                r = ConstructionCutoff:    If Trim(.TextMatrix(r, c)) <> "" And .ValueMatrix(r, c) = 0 Then AddMsg WARN, CellAddress(r, c) & " - const cutoff is zero."
            
            Next
        End If
        
        
        '-- check items for errors -------------------------------------------------------------------------------
        For r = mFirstItemRow To mLastItemRow
            
            'is required
            c = ItemUOM:          If Trim(.TextMatrix(r, c)) = "" Then AddMsg ERROR, CellAddress(r, c) & " - unit of measure is required."
            c = EstPhase:         If Trim(.TextMatrix(r, c)) = "" Then AddMsg ERROR, CellAddress(r, c) & " - phase is required."
            c = ItemDescription:  If Trim(.TextMatrix(r, c)) = "" Then AddMsg ERROR, CellAddress(r, c) & " - description is required."
            c = EstItem:          If Trim(.TextMatrix(r, c)) = "" Then AddMsg ERROR, CellAddress(r, c) & " - item is required."
            
            'is too big
            c = ItemUOM:          If Len(Trim(.TextMatrix(r, c))) > 10 Then AddMsg ERROR, CellAddress(r, c) & " - UOM is too long (max 10)."
            c = EstPhase:         If Len(Trim(.TextMatrix(r, c))) > 20 Then AddMsg ERROR, CellAddress(r, c) & " - phase is too long (max 20)."
            c = EstItem:          If Len(Trim(.TextMatrix(r, c))) > 16 Then AddMsg ERROR, CellAddress(r, c) & " - item is too long (max 16)."
            c = ItemDescription:  If Len(Trim(.TextMatrix(r, c))) > 200 Then AddMsg ERROR, CellAddress(r, c) & " - description is too long (max 200)."
            c = ItemNotes:        If Len(Trim(.TextMatrix(r, c))) > 4000 Then AddMsg ERROR, CellAddress(r, c) & " - description is too long (max 200)."
            c = JCCostCode:       If Len(Trim(.TextMatrix(r, c))) > 40 Then AddMsg ERROR, CellAddress(r, c) & " - cost code is too long (max 15)."
            c = JCCategory:       If Len(Trim(.TextMatrix(r, c))) > 40 Then AddMsg ERROR, CellAddress(r, c) & " - category is too long (max 3)."
            
            'Items that dont end with a valid alpha character
            Select Case UCase(Right(Trim(.TextMatrix(r, EstItem)), 1))
            Case "M", "S", "L", "E", "O"
            Case Else
             AddMsg ERROR, CellAddress(r, EstItem) & " - item code is invalid."
            End Select
        Next
    
    
        '-- check items in each assembly -------------------------------------------------------------------------------
        
        'sort by phase/item to find duplicates
        .Col = 3
        .ColSel = 4
        .ColSort(3) = flexSortNumericAscending
        .ColSort(4) = flexSortNumericAscending
        .Sort = flexSortNumericAscending
        'for each item
        For r = mFirstItemRow + 1 To mLastItemRow
            'check worksheet for duplicate phase/item
            If .TextMatrix(r, EstPhase) = .TextMatrix(r - 1, EstPhase) And .TextMatrix(r, EstItem) = .TextMatrix(r - 1, EstItem) Then
                AddMsg ERROR, CellAddress(r, EstPhase) & " - phase and item is not unique."
            End If
            'for each assembly
            For c = FirstAssemblyCol To .Cols - 1 Step mColsPerAssembly
                If mIsShortForm Then
                    If Trim(.TextMatrix(r, c + 0)) <> "" And Not IsNumeric(.TextMatrix(r, c + 0)) Then AddMsg ERROR, CellAddress(r, c + 0) & " - rate is not a number."
                Else
                    If Len(Trim(.TextMatrix(r, c + 0))) > 4000 Then AddMsg ERROR, CellAddress(r, c) & " - Notes are too long (max 4000)."
                    If Trim(.TextMatrix(r, c + 1)) <> "" And Not IsNumeric(.TextMatrix(r, c + 1)) Then AddMsg ERROR, CellAddress(r, c + 1) & " - qty is not a number."
                    If Trim(.TextMatrix(r, c + 2)) <> "" And Not IsNumeric(.TextMatrix(r, c + 2)) Then AddMsg ERROR, CellAddress(r, c + 2) & " - rate is not a number."
                    If .ValueMatrix(r, c + 1) > 0 And .ValueMatrix(r, c + 2) = 0 Then AddMsg WARN, CellAddress(r, c + 2) & " - item rate is zero."
                      
                    If .TextMatrix(r, c + 1) = "" And Not bQtyDefaultsToOne Then
                        AddMsg INFO, "Quantities not specified default to one."
                        bQtyDefaultsToOne = True
                    End If
                    
                End If
            Next
            
        Next
    
        .Col = 0
        .Sort = flexSortNumericAscending
    End With
    
    lblErrorCountLabel.Visible = mErrors > 0
    lblErrorCount.Visible = mErrors > 0
    lblErrorCount = mErrors

    lblWarningCountLabel.Visible = mWarnings > 0
    lblWarningCount.Visible = mWarnings > 0
    lblWarningCount = mWarnings

    gErrors.Col = 0
    gErrors.Sort = flexSortStringAscending
    cmdNav(0).Enabled = mErrors = 0
    If gErrors.Rows = 0 Then Call TabStrip.Tabs.Remove(1)

    Select Case True
        Case mErrors > 0:   imgIconError.Visible = True
        Case mWarnings > 0: imgIconWarning.Visible = True
        Case Else:          imgIconInfo.Visible = True
    End Select
    Exit Sub
eh: Call errHandler(SRCFILE & "Validate Data")
End Sub
Private Function CellAddress(Row As Long, Col As Long) As String
    CellAddress = gAssemblies.TextMatrix(0, Col) & gAssemblies.TextMatrix(Row, 0)
End Function
Public Sub WriteCommunity()
On Error Resume Next
    Dim s As String
    If Trim(mCommunity) <> "" Then
        s = ""
        s = s & "INSERT INTO tblLocality(Area,Description,Abr,Inactive)" & vbCrLf
        s = s & "VALUES(" & DbQuote(Str, mCommunity) & vbCrLf
        s = s & "," & DbQuote(Str, mCommunityDesc) & vbCrLf
        s = s & "," & DbQuote(Str, mCommunity) & vbCrLf
        s = s & ",0)"
        Call HFApp.SqlExec(s)
    End If
       
    If Trim(mCommunityPhase) <> "" Then
        s = ""
        s = s & "INSERT INTO CommunityPhase(Community,CommunityPhase,Description)" & vbCrLf
        s = s & "VALUES(" & DbQuote(Str, mCommunity) & vbCrLf
        s = s & "," & DbQuote(Str, mCommunityPhase) & vbCrLf
        s = s & "," & DbQuote(Str, mCommunityPhase) & ")"
        Call HFApp.SqlExec(s)
    End If
    
End Sub




Public Sub WriteCategories()
On Error GoTo eh

    Dim r As Long
    Dim s As String
    
    With gCategories
    For r = 2 To .Rows - 1
        If Trim(.TextMatrix(r, 3)) <> "" Then
            s = ""
            s = s & "INSERT INTO tblMajorGroups(Major_Group,Description)" & vbCrLf
            s = s & "VALUES(" & DbQuote(Str, .TextMatrix(r, 3))
            s = s & "," & DbQuote(Str, .TextMatrix(r, 4))
            s = s & ")"
        End If
        Call HFApp.SqlExec(s)
        
        s = ""
        s = s & "INSERT INTO tblCategories(Category,Description,Group_Code)" & vbCrLf
        s = s & "VALUES(" & DbQuote(Str, .TextMatrix(r, 1))
        s = s & "," & DbQuote(Str, .TextMatrix(r, 2))
        s = s & "," & DbQuote(Str, .TextMatrix(r, 3))
        s = s & ")"
        Call HFApp.SqlExec(s)
    Next
    End With
    Exit Sub
eh: If InStr(1, Err.Description, "duplicate", vbTextCompare) Then
        Resume Next
    Else
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    End If
End Sub


Public Sub WritePhases()
On Error GoTo eh

    Dim r As Long
    Dim s As String
    
    For r = 2 To gPhases.Rows - 1
        If Trim(gPhases.TextMatrix(r, 1)) <> "" Then
            s = ""
            s = s & "INSERT INTO tblEstPhases(DivisionID,Phase,SortOrder,Description,GroupPhase,UpdateEstimating,UStmp,TStmp)" & vbCrLf
            s = s & "VALUES(" & HFApp.DivisionID & "," & DbQuote(Str, gPhases.TextMatrix(r, 1), , True)
            s = s & "," & DbQuote(Num, gPhases.TextMatrix(r, 1))
            s = s & "," & DbQuote(Str, gPhases.TextMatrix(r, 3), , True)
            s = s & ",1,1," & DbQuote(Str, HFApp.LoginID) & ",getdate())"
        Else
            s = ""
            s = s & "INSERT INTO tblEstPhases(DivisionID,Phase,SortOrder,Description,GroupPhase,UpdateEstimating,UStmp,TStmp)" & vbCrLf
            s = s & "VALUES(" & HFApp.DivisionID & "," & DbQuote(Str, gPhases.TextMatrix(r, 2), , True)
            s = s & "," & DbQuote(Num, gPhases.TextMatrix(r, 2))
            s = s & "," & DbQuote(Str, gPhases.TextMatrix(r, 3), , True)
            s = s & ",0,1," & DbQuote(Str, HFApp.LoginID) & ",getdate())"
        End If
        Call HFApp.SqlExec(s)
    
    Next
    
    
    Call FixGroupPhaseValue
    
    Exit Sub
eh: If InStr(1, Err.Description, "duplicate", vbTextCompare) Then
        Resume Next
    Else
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    End If
End Sub

Public Sub WriteItems()
On Error GoTo eh
    Dim i As Long
    Dim r As Long
    Dim s As String
    
    For r = mFirstItemRow To mLastItemRow
        i = i + 1
        
        ProgressBar.value = i / Max(1, (mLastItemRow - mFirstItemRow) * 100)
        
        lblStatus.Caption = "Writing items row " & (r - mFirstItemRow) & " of " & (mLastItemRow - mFirstItemRow)
        lblStatus.Refresh
        
        
        s = ""
        s = s & "INSERT INTO tblPhaseItem(DivisionID,Phase,Item,PriceLink,Description,Notes,JCCostCode,JCCategory,TakeoffUOM,OrderUOM,ConversionFactor,UpdateEstimating,UStmp,TStmp,ItemNumber)" & vbCrLf
        s = s & "VALUES(" & HFApp.DivisionID & "," & DbQuote(Str, gAssemblies.TextMatrix(r, EstPhase), , True) & vbCrLf
        s = s & "," & DbQuote(Str, gAssemblies.TextMatrix(r, EstItem), , True) & vbCrLf
        s = s & ",0" & vbCrLf
        s = s & "," & DbQuote(Str, gAssemblies.TextMatrix(r, ItemDescription)) & vbCrLf
        s = s & "," & DbQuote(Str, gAssemblies.TextMatrix(r, ItemNotes)) & vbCrLf
        s = s & "," & DbQuote(Str, gAssemblies.TextMatrix(r, JCCostCode)) & vbCrLf
        s = s & "," & DbQuote(Str, gAssemblies.TextMatrix(r, JCCategory)) & vbCrLf
        s = s & "," & DbQuote(Str, gAssemblies.TextMatrix(r, ItemUOM)) & vbCrLf
        s = s & "," & DbQuote(Str, gAssemblies.TextMatrix(r, ItemUOM)) & vbCrLf
        s = s & ",1,1," & DbQuote(Str, HFApp.LoginID) & ",GETDATE()" & "," & DbQuote(Str, Left(Trim(gAssemblies.TextMatrix(r, EstItem)), Len(Trim(gAssemblies.TextMatrix(r, EstItem))) - 1), , , True) & " )"
        Call HFApp.SqlExec(s)
    
        s = ""
        s = s & "UPDATE tblPhaseItem" & vbCrLf
        s = s & "SET Description=" & DbQuote(Str, gAssemblies.TextMatrix(r, ItemDescription)) & vbCrLf
        s = s & "   ,Notes=" & DbQuote(Str, gAssemblies.TextMatrix(r, ItemNotes)) & vbCrLf
        s = s & "   ,JCCostCode=" & DbQuote(Str, gAssemblies.TextMatrix(r, JCCostCode)) & vbCrLf
        s = s & "   ,JCCategory=" & DbQuote(Str, gAssemblies.TextMatrix(r, JCCategory)) & vbCrLf
        s = s & "   ,TakeoffUOM=" & DbQuote(Str, gAssemblies.TextMatrix(r, ItemUOM)) & vbCrLf
        s = s & "   ,OrderUOM=" & DbQuote(Str, gAssemblies.TextMatrix(r, ItemUOM)) & vbCrLf
        s = s & "   ,ConversionFactor=1" & vbCrLf
        s = s & "   ,UpdateEstimating=1" & vbCrLf
        s = s & "   ,UStmp=" & DbQuote(Str, HFApp.LoginID) & vbCrLf
        s = s & "   ,TStmp=GETDATE()" & vbCrLf
        s = s & "WHERE DivisionID = " & HFApp.DivisionID & " and Phase=" & DbQuote(Str, gAssemblies.TextMatrix(r, EstPhase), , True) & vbCrLf
        s = s & "  AND Item=" & DbQuote(Str, gAssemblies.TextMatrix(r, EstItem), , True) & vbCrLf
      Call HFApp.SqlExec(s)
    Next

    Exit Sub
eh: If InStr(1, Err.Description, "duplicate", vbTextCompare) Then
        Resume Next
    Else
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    End If
End Sub


Public Sub WriteAssemblies()
On Error GoTo eh
    Dim rs As Recordset
    Dim r As Long
    Dim c As Long
    Dim s As String
    
    Dim assemblyCommunity As String
    
    Dim incentiveCost As Double
    Dim percentage    As Double
    Dim landCost      As Double
    
    Dim pretax       As Double
    Dim aftertax     As Double
    Dim tax          As Double
    Dim taxrate      As Double
    Dim includetax   As Boolean
    Dim copretax     As Double
    Dim coaftertax   As Double
    Dim cotax        As Double
    Dim markup       As Double
    Dim margin       As Double
    Dim itemqty      As Double
    Dim ItemRate     As Double
    Dim itemcost     As Double
    Dim totalcost    As Double
    Dim skiprow      As Boolean
    Dim Note         As String
    Dim po           As String
    
    
    Select Case True
        Case mAssemblyType = atModel And HFApp.Options(ModelByArea):            assemblyCommunity = IIf(Me.chkSaveCommunity, mCommunity, "")
        Case mAssemblyType = atoption And HFApp.Options(OptionByArea):          assemblyCommunity = IIf(Me.chkSaveCommunity, mCommunity, "")
        Case mAssemblyType = atGlobal And HFApp.Options(GlobalOptionByArea):    assemblyCommunity = IIf(Me.chkSaveCommunity, mCommunity, "")
        Case mAssemblyType = atDesignCenter And HFApp.Options(DCOptionByArea):  assemblyCommunity = IIf(Me.chkSaveCommunity, mCommunity, "")
        Case Else:                                                              assemblyCommunity = ""
    End Select
        
        
    percentage = HFApp.Options(IncentiveCostPercent) / 100
    If percentage = 0 Then percentage = 1
    
    
    With gAssemblies
        'for each assembly
        For c = FirstAssemblyCol To .Cols - 1 Step mColsPerAssembly
            
            lblStatus.Caption = "Writing assemblies " & (c - 7) & " of " & (.Cols - 8) / mColsPerAssembly
            lblStatus.Refresh
            ProgressBar.value = (c - 8) / (.Cols - 8) * 100
            Me.Refresh
            
            If .ValueMatrix(ConstructionCutoff, c) <> 0 Then
                Call HFApp.SqlExec("INSERT INTO tblJobStatus(Job_Status,Description) VALUES(" & .ValueMatrix(ConstructionCutoff, c) & ",'Status " & .ValueMatrix(ConstructionCutoff, c) & "')")
            End If
            
            If .ValueMatrix(ConstructionCutoff, c) <> 0 Then
                Call HFApp.SqlExec("INSERT INTO tblSeries(DivisionID,Series,Description) VALUES(" & HFApp.DivisionID & "," & DbQuote(Str, .TextMatrix(series, c)) & "," & DbQuote(Str, .TextMatrix(series, c)) & ")")
            End If
            
            pretax = Round(.ValueMatrix(SellingPretax, c), 2)
            aftertax = Round(.ValueMatrix(SellingPriceTaxIncl, c), 2)
            includetax = aftertax <> 0
            If mAssemblyType = atModel Then taxrate = HFApp.Options(ModelTaxRate)
            If mAssemblyType = atoption Then taxrate = HFApp.Options(MSOPTTaxRate)
            If mAssemblyType = atGlobal Then taxrate = HFApp.Options(GLOPTTaxRate)
            If mAssemblyType = atDesignCenter Then taxrate = HFApp.Options(GLOPTTaxRate)
            
            Select Case True
                Case aftertax <> 0 And pretax <> 0:  taxrate = (aftertax * 100 / pretax) - 100
                'Case aftertax <> 0:                  pretax = Round(aftertax / (taxrate + 100) * 100, 2)
                'Case pretax <> 0:                    aftertax = Round(pretax * (taxrate + 100) / 100, 2)
                Case aftertax <> 0:                  pretax = CalcPreTax(mAssemblyType, aftertax)
                Case pretax <> 0:                    aftertax = pretax + CalcTax(mAssemblyType, pretax)
            End Select
            tax = aftertax - pretax
            
            
            copretax = Round(.ValueMatrix(COSellingPretax, c), 2)
            coaftertax = Round(.ValueMatrix(COSellingPriceTaxIncl, c), 2)
            Select Case True
                Case coaftertax <> 0:                  copretax = CalcPreTax(mAssemblyType, coaftertax)
                Case copretax <> 0:                    coaftertax = copretax + CalcTax(mAssemblyType, copretax)
            End Select
            cotax = coaftertax - copretax
            
            
            
            s = ""
            s = s & "INSERT INTO tblDBAssemblyMaster(Community,DivisionID,Assembly,Model,OptionID,FloorArea,Bedrooms,Bathrooms,ConstCutoff,AssemblyType,Markup,Margin,Roundto,Pretax,Tax,Cost,COPretax,COTax,TaxRate,UseNormalSalesQtyFactors,UStmp,TStmp)" & vbCrLf
            s = s & "VALUES(" & DbQuote(Str, assemblyCommunity) & vbCrLf
            s = s & "," & HFApp.DivisionID & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(AssemblyID, c)) & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(ModelAssemblyID, c)) & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(OptionAssemblyID, c)) & vbCrLf
            s = s & "      ,0,0,0,1,1,0,0,0,0,0,0,0,0,0,1," & DbQuote(Str, HFApp.LoginID) & ",getdate())"
            Call HFApp.SqlExec(s)
            
            s = ""
            s = s & "UPDATE tblDBAssemblyMaster" & vbCrLf
            s = s & "SET AssemblyType=" & DbQuote(Num, mAssemblyType) & vbCrLf
            s = s & "   ,Pretax=" & DbQuote(Num, pretax) & vbCrLf
            s = s & "   ,Tax=" & DbQuote(Num, tax) & vbCrLf
            s = s & "   ,COPretax=" & DbQuote(Num, copretax) & vbCrLf
            s = s & "   ,COTax=" & DbQuote(Num, cotax) & vbCrLf
            s = s & "   ,IncludeTax=" & IIf(includetax, 1, 0) & vbCrLf
            s = s & "   ,TaxRate=" & DbQuote(Num, taxrate) & vbCrLf
            s = s & "   ,Elevation=" & DbQuote(Str, .TextMatrix(Elevation, c)) & vbCrLf
            s = s & "   ,Series=" & DbQuote(Str, .TextMatrix(series, c)) & vbCrLf
            s = s & "   ,ConstCutoff=" & DbQuote(Num, .TextMatrix(ConstructionCutoff, c)) & vbCrLf
            s = s & "   ,FloorArea=" & DbQuote(Num, .TextMatrix(FloorArea, c)) & vbCrLf
            s = s & "   ,Bedrooms=" & DbQuote(Num, .TextMatrix(Bedrooms, c)) & vbCrLf
            s = s & "   ,Bathrooms=" & DbQuote(Num, .TextMatrix(Bathrooms, c)) & vbCrLf
            s = s & "   ,Style=" & DbQuote(Str, .TextMatrix(Style, c)) & vbCrLf
            s = s & "   ,Description=" & DbQuote(Str, .TextMatrix(AssemblyDescription, c)) & vbCrLf
            s = s & "   ," & AssemblyNoteField & "=" & DbQuote(Str, .TextMatrix(AssemblyNotes, c)) & vbCrLf
            s = s & "   ,Category=" & DbQuote(Str, .TextMatrix(SalesCategory, c)) & vbCrLf
            s = s & "   ,AssemblyUOM=" & DbQuote(Str, .TextMatrix(AssemblyUOM, c)) & vbCrLf
            
            s = s & "   ,JCExtra=" & DbQuote(Str, .TextMatrix(JCExtra, c)) & vbCrLf
            s = s & "   ,UStmp=" & DbQuote(Str, HFApp.LoginID) & vbCrLf
            s = s & "   ,TStmp=getdate()" & vbCrLf
            s = s & "WHERE Community=" & DbQuote(Str, assemblyCommunity) & vbCrLf
            s = s & "  AND Assembly=" & DbQuote(Str, .TextMatrix(AssemblyID, c)) & vbCrLf
            s = s & "  AND Model=" & DbQuote(Str, .TextMatrix(ModelAssemblyID, c)) & vbCrLf
            s = s & "  AND OptionID=" & DbQuote(Str, .TextMatrix(OptionAssemblyID, c)) & vbCrLf
            s = s & " and DivisionID = " & HFApp.DivisionID
            Call HFApp.SqlExec(s)
        
            
            
            s = ""
            s = s & "DELETE FROM tblDBAssemblyDetails" & vbCrLf
            s = s & "WHERE Community=" & DbQuote(Str, assemblyCommunity) & vbCrLf
            s = s & "  AND Assembly=" & DbQuote(Str, .TextMatrix(AssemblyID, c)) & vbCrLf
            s = s & "  AND Model=" & DbQuote(Str, .TextMatrix(ModelAssemblyID, c)) & vbCrLf
            s = s & "  AND OptionID=" & DbQuote(Str, .TextMatrix(OptionAssemblyID, c)) & vbCrLf
            s = s & " and DivisionID = " & HFApp.DivisionID
            Call HFApp.SqlExec(s)
            
            incentiveCost = 0
            totalcost = 0
            landCost = 0
            
            'for each item in assembly
            For r = mFirstItemRow To mLastItemRow
                
                skiprow = False
                If mIsShortForm Then
                    If .ValueMatrix(r, c) = 0 Then
                        'skip row if rate=0
                        skiprow = True
                    Else
                        itemqty = 1
                        ItemRate = .ValueMatrix(r, c)
                        itemcost = itemqty * ItemRate
                        totalcost = totalcost + itemcost
                        Note = .TextMatrix(r, ItemNotes)
                        If POIndex = -1 Then
                            po = ""
                        Else
                            po = Trim(.TextMatrix(r, POIndex))
                        End If
                    End If
                Else
                    If .ValueMatrix(r, c + 1) = 0 Then
                        'skip row if qty=0
                        skiprow = True
                    Else
                        itemqty = IIf(Trim(.TextMatrix(r, c + 1)) = "", 1, .ValueMatrix(r, c + 1))
                        ItemRate = .ValueMatrix(r, c + 2)
                        itemcost = itemqty * ItemRate
                        totalcost = totalcost + itemcost
                        Note = IIf(Trim(.TextMatrix(r, c)) = "", .TextMatrix(r, ItemNotes), .TextMatrix(r, c))
                        If POIndex <> -1 Then po = .TextMatrix(r, POIndex)
                    End If
                End If
                
                
                If Not skiprow Then
                    
                    If .TextMatrix(r, EstPhase) = HFApp.Options(IncentivePhase) And .TextMatrix(r, EstItem) = HFApp.Options(IncentiveItem) Then
                       incentiveCost = incentiveCost + itemcost
                    End If
                    If .TextMatrix(r, EstPhase) = HFApp.Options(LandPhase) And .TextMatrix(r, EstItem) = HFApp.Options(LandItem) Then
                       landCost = landCost + itemcost
                    End If
                    
                    s = ""
                    s = s & "INSERT INTO tblDBAssemblyDetails(Community,DivisionID,Assembly,Model,OptionID,Phase,Item,TakeoffQty,OrderQty,Rate,Cost,Notes,POIndex,UStmp,TStmp)" & vbCrLf
                    s = s & "VALUES(" & DbQuote(Str, assemblyCommunity) & vbCrLf
                    s = s & "," & HFApp.DivisionID & vbCrLf
                    s = s & "      ," & DbQuote(Str, .TextMatrix(AssemblyID, c)) & vbCrLf
                    s = s & "      ," & DbQuote(Str, .TextMatrix(ModelAssemblyID, c)) & vbCrLf
                    s = s & "      ," & DbQuote(Str, .TextMatrix(OptionAssemblyID, c)) & vbCrLf
                    s = s & "      ," & DbQuote(Str, .TextMatrix(r, EstPhase), , True) & vbCrLf
                    s = s & "      ," & DbQuote(Str, .TextMatrix(r, EstItem), , True) & vbCrLf
                    s = s & "      ," & DbQuote(Num, itemqty) & vbCrLf
                    s = s & "      ," & DbQuote(Num, itemqty) & vbCrLf
                    s = s & "      ," & DbQuote(Num, ItemRate) & vbCrLf
                    s = s & "      ," & DbQuote(Num, itemcost) & vbCrLf
                    s = s & "      ," & DbQuote(Str, Note) & vbCrLf
                    s = s & "      ," & DbQuote(Str, po) & vbCrLf
                    s = s & "      ," & DbQuote(Str, HFApp.LoginID) & vbCrLf
                    s = s & "      ,getdate())" & vbCrLf
                    Call HFApp.SqlExec(s)
                    
                End If
                
            Next
        
            s = ""
            s = s & "UPDATE tblDBAssemblyDetails" & vbCrLf
            s = s & "   SET OrderQty=TakeoffQty*ISNULL(tblphaseitem.ConversionFactor,1)" & vbCrLf
            s = s & "      ,POIndex=case when tbldbAssemblyDetails.POIndex='' then tblphaseitem.POIndex else tbldbAssemblyDetails.POIndex end" & vbCrLf
            s = s & "  FROM tbldbAssemblyDetails LEFT OUTER JOIN tblphaseitem ON (tbldbassemblydetails.divisionid = tblphaseitem.DivisionID and tbldbAssemblyDetails.phase=tblphaseitem.phase and tbldbAssemblyDetails.item=tblphaseitem.item)" & vbCrLf
            s = s & " WHERE Community=" & DbQuote(Str, assemblyCommunity) & vbCrLf
            s = s & "   AND Assembly=" & DbQuote(Str, .TextMatrix(AssemblyID, c)) & vbCrLf
            s = s & "   AND Model=" & DbQuote(Str, .TextMatrix(ModelAssemblyID, c)) & vbCrLf
            s = s & "   AND tbldbAssemblyDetails.OptionID=" & DbQuote(Str, .TextMatrix(OptionAssemblyID, c)) & vbCrLf
            s = s & " and tbldbAssemblyDetails.DivisionID = " & HFApp.DivisionID
            Call HFApp.SqlExec(s)

        
            markup = 0
            margin = 0
            If totalcost <> 0 Then markup = (pretax - totalcost) / totalcost * 100
            If pretax <> 0 Then margin = (pretax - totalcost) / pretax * 100
            
            s = ""
            s = s & "UPDATE tblDBAssemblyMaster" & vbCrLf
            s = s & "SET Cost=" & DbQuote(Num, totalcost) & vbCrLf
            s = s & "   ,Markup=" & DbQuote(Num, markup) & vbCrLf
            s = s & "   ,Margin=" & DbQuote(Num, margin) & vbCrLf
            s = s & "   ,IncentiveRetail=" & DbQuote(Num, incentiveCost / percentage) & vbCrLf
            s = s & "   ,UseNormalSalesQtyFactors = 1" & vbCrLf
            s = s & "WHERE Community=" & DbQuote(Str, assemblyCommunity) & vbCrLf
            s = s & "  AND Assembly=" & DbQuote(Str, .TextMatrix(AssemblyID, c)) & vbCrLf
            s = s & "  AND Model=" & DbQuote(Str, .TextMatrix(ModelAssemblyID, c)) & vbCrLf
            s = s & "  AND OptionID=" & DbQuote(Str, .TextMatrix(OptionAssemblyID, c)) & vbCrLf
            s = s & " and DivisionID = " & HFApp.DivisionID
            Call HFApp.SqlExec(s)
            
            
            
            s = ""
            s = s & "INSERT tblVendorCost(Community,DivisionID,CommunityPhase,Assembly,Model,Phase,Item,PriceLink,Vendor" & vbCrLf
            s = s & "                    ,Current_Cost,Next_Cost1,Next_Cost2,Last_Cost1,Last_Cost2,Last_Cost3" & vbCrLf
            s = s & "                    ,Forecast1,Forecast2,Forecast3,Forecast4,Forecast5,Forecast6,Forecast7,Forecast8,Forecast9,Forecast10,Forecast11,Forecast12)" & vbCrLf
            s = s & "SELECT d.Community" & vbCrLf
            s = s & "," & HFApp.DivisionID & vbCrLf
            s = s & "      ," & DbQuote(Str, mCommunityPhase) & vbCrLf
            s = s & "      ,case when i.PriceLink=0 then d.Assembly else '' end" & vbCrLf
            s = s & "      ,d.Model" & vbCrLf
            s = s & "      ,d.Phase" & vbCrLf
            s = s & "      ,d.Item" & vbCrLf
            s = s & "      ,i.PriceLink" & vbCrLf
            s = s & "      ,dbo.Purch_GetCommunityVendor(" & DbQuote(Str, mCommunity) & ",i.POIndex," & HFApp.DivisionID & ") Vendor" & vbCrLf
            s = s & "      ,d.Rate,d.Rate,d.Rate,d.Rate,d.Rate,d.Rate,d.Rate,d.Rate,d.Rate,d.Rate,d.Rate,d.Rate,d.Rate,d.Rate,d.Rate,d.Rate,d.Rate,d.Rate" & vbCrLf
            s = s & "  FROM tblDBAssemblyDetails d" & vbCrLf
            s = s & "       JOIN tblPhaseItem i ON(d.DivisionID = i.DivisionID and d.phase=i.phase AND d.Item=i.Item)" & vbCrLf
            s = s & "       LEFT OUTER JOIN tblVendorCost c ON(d.DivisionID = c.DivisionID and d.Community=c.Community AND c.CommunityPhase=" & DbQuote(Str, mCommunityPhase) & " AND d.Assembly=c.Assembly AND d.Phase=c.Phase AND c.model=d.model AND c.Item=d.Item  AND c.vendor=dbo.Purch_GetCommunityVendor(d.Community,i.POIndex," & HFApp.DivisionID & "))" & vbCrLf
            s = s & " WHERE c.Phase IS NULL" & vbCrLf
            s = s & "   AND NOT dbo.Purch_GetCommunityVendor(" & DbQuote(Str, mCommunity) & ",i.POIndex," & HFApp.DivisionID & ") IS NULL" & vbCrLf
            s = s & "   AND d.Community=" & DbQuote(Str, IIf(Me.chkSaveCommunity, mCommunity, "")) & vbCrLf
            s = s & "   AND d.Assembly=" & DbQuote(Str, .TextMatrix(AssemblyID, c)) & vbCrLf
            s = s & "   AND d.Model=" & DbQuote(Str, .TextMatrix(ModelAssemblyID, c)) & vbCrLf
            s = s & "   AND d.OptionID=" & DbQuote(Str, .TextMatrix(OptionAssemblyID, c)) & vbCrLf
            s = s & "   AND isnull(d.rate,0)<>0" & vbCrLf
            s = s & "   AND d.DivisionID = " & HFApp.DivisionID
            Call HFApp.SqlExec(s)
        
        
                        
            
            s = ""
            s = s & "INSERT INTO tblDBAssemblyPrices(Community,CommunityPhase,Assembly,Model,OptionID)" & vbCrLf
            s = s & "VALUES(" & DbQuote(Str, assemblyCommunity) & vbCrLf
            s = s & "      ," & DbQuote(Str, mCommunityPhase) & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(AssemblyID, c)) & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(ModelAssemblyID, c)) & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(OptionAssemblyID, c)) & ")"
            Call HFApp.SqlExec(s)
            s = ""
            s = s & "UPDATE tblDBAssemblyPrices" & vbCrLf
            s = s & "   SET Margin=m.Margin" & vbCrLf
            s = s & "      ,Markup=m.Markup" & vbCrLf
            s = s & "      ,RoundTo=m.RoundTo" & vbCrLf
            s = s & "      ,Pretax=m.Pretax" & vbCrLf
            s = s & "      ,Tax=m.Tax" & vbCrLf
            s = s & "      ,Cost=m.Cost" & vbCrLf
            s = s & "      ,COPretax=m.COPretax" & vbCrLf
            s = s & "      ,COTax=m.COTax" & vbCrLf
            s = s & "      ,IncludeTax=m.IncludeTax" & vbCrLf
            s = s & "      ,TaxRate=m.TaxRate" & vbCrLf
            s = s & "      ,IncentiveCost=" & DbQuote(Num, incentiveCost) & vbCrLf
            s = s & "      ,IncentiveRetail=m.IncentiveRetail" & vbCrLf
            s = s & "      ,LandCost=" & DbQuote(Num, landCost) & vbCrLf
            s = s & "  FROM tblDBAssemblyMaster m" & vbCrLf
            s = s & "      ,tblDBAssemblyPrices p " & vbCrLf
            s = s & " WHERE p.Community=m.Community " & vbCrLf
            s = s & "   AND p.Assembly=m.Assembly" & vbCrLf
            s = s & "   AND p.Model=m.Model" & vbCrLf
            s = s & "   AND p.OptionID=m.OptionID" & vbCrLf
            s = s & "   AND m.Community=" & DbQuote(Str, assemblyCommunity) & vbCrLf
            s = s & "   AND p.CommunityPhase=" & DbQuote(Str, mCommunityPhase) & vbCrLf
            s = s & "   AND m.Assembly=" & DbQuote(Str, .TextMatrix(AssemblyID, c)) & vbCrLf
            s = s & "   AND m.Model=" & DbQuote(Str, .TextMatrix(ModelAssemblyID, c)) & vbCrLf
            s = s & "   AND m.OptionID=" & DbQuote(Str, .TextMatrix(OptionAssemblyID, c)) & vbCrLf
            s = s & "   AND m.DivisionID = " & HFApp.DivisionID
            Call HFApp.SqlExec(s)
        Next
    End With
    

    Exit Sub
eh: If InStr(1, Err.Description, "duplicate", vbTextCompare) Then
        Resume Next
    Else

        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call IniPutForm(Me)
End Sub

Private Sub TabStrip_Click()
On Error Resume Next
    gErrors.Move TabStrip.ClientLeft, TabStrip.ClientTop, TabStrip.ClientWidth, TabStrip.ClientHeight
    gAssemblies.Move TabStrip.ClientLeft, TabStrip.ClientTop, TabStrip.ClientWidth, TabStrip.ClientHeight
    gPhases.Move TabStrip.ClientLeft, TabStrip.ClientTop, TabStrip.ClientWidth, TabStrip.ClientHeight
    gCategories.Move TabStrip.ClientLeft, TabStrip.ClientTop, TabStrip.ClientWidth, TabStrip.ClientHeight
    gErrors.Visible = TabStrip.SelectedItem.Caption = "Messages"
    gAssemblies.Visible = TabStrip.SelectedItem.Caption = "Assemblies"
    gPhases.Visible = TabStrip.SelectedItem.Caption = "Phases"
    gCategories.Visible = TabStrip.SelectedItem.Caption = "Categories"
End Sub

Private Sub AddMsg(Level As Integer, Text As String)
    Select Case Level
        Case 0: mErrors = mErrors + 1
        Case 1: mWarnings = mWarnings + 1
        Case 2: mInfoMsgs = mInfoMsgs + 1
    End Select
    gErrors.AddItem Level & vbTab & Text
    gErrors.Cell(flexcpPicture, gErrors.Rows - 1, 1) = imgMsg(Level).Picture
End Sub

Public Property Let ImportMode(RHS As String)
    mImportMode = RHS
End Property
