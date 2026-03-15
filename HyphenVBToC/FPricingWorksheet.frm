VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FPricingWorksheet 
   Caption         =   "Sales Pricing Worksheet"
   ClientHeight    =   8745
   ClientLeft      =   2700
   ClientTop       =   2115
   ClientWidth     =   12960
   Icon            =   "FPricingWorksheet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8745
   ScaleWidth      =   12960
   Begin VB.Frame HeadingFrame 
      BorderStyle     =   0  'None
      Height          =   1275
      Left            =   120
      TabIndex        =   16
      Top             =   600
      Width           =   15195
      Begin VB.CheckBox chkArchive 
         Caption         =   "Archive this sheet"
         Height          =   195
         Left            =   11685
         TabIndex        =   26
         Top             =   60
         Width           =   1650
      End
      Begin VSFlex8Ctl.VSFlexGrid gLegend 
         Height          =   975
         Left            =   9480
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   360
         Width           =   4635
         _cx             =   1978146800
         _cy             =   1978140344
         Appearance      =   1
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
         BackColor       =   -2147483639
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483639
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
         FocusRect       =   1
         HighLight       =   0
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   4
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FPricingWorksheet.frx":000C
         ScrollTrack     =   0   'False
         ScrollBars      =   0
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
      Begin VB.TextBox txtPercentDecrease 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   220
         Left            =   8340
         TabIndex        =   13
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtIncentiveCostPercent 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   220
         Left            =   5520
         TabIndex        =   7
         Text            =   "0"
         Top             =   540
         Width           =   555
      End
      Begin VB.TextBox txtRoundTo 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   220
         Left            =   5520
         TabIndex        =   5
         Text            =   "0"
         Top             =   300
         Width           =   555
      End
      Begin VB.TextBox txtDollarChange 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   220
         Left            =   8340
         TabIndex        =   9
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox txtPercentIncrease 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   220
         Left            =   8340
         TabIndex        =   11
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txtSalesEffectiveDate 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   220
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Text            =   " "
         Top             =   960
         Width           =   2355
      End
      Begin VB.TextBox txtCostsEffectiveDate 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   220
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   720
         Width           =   2355
      End
      Begin HFEst.VBCombo cboCostBasis 
         Height          =   240
         Left            =   1440
         TabIndex        =   1
         Top             =   435
         Width           =   1755
         _ExtentX        =   1746
         _ExtentY        =   423
         Style           =   2
      End
      Begin VB.TextBox txtDescription 
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   0
         Top             =   180
         Width           =   2355
      End
      Begin HFEst.VBCombo cboLayouts 
         Height          =   240
         Left            =   4440
         TabIndex        =   24
         Top             =   900
         Width           =   2115
         _ExtentX        =   2117
         _ExtentY        =   423
         Style           =   2
      End
      Begin VB.Image cmdLayouts 
         Height          =   240
         Left            =   6555
         Picture         =   "FPricingWorksheet.frx":006E
         Top             =   900
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Layout"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   8
         Left            =   3780
         TabIndex        =   25
         Top             =   930
         Width           =   585
      End
      Begin VB.Label lblLegend 
         AutoSize        =   -1  'True
         Caption         =   "Legend"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   9480
         TabIndex        =   22
         Top             =   120
         Width           =   645
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Price Change "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   4
         Left            =   7560
         TabIndex        =   21
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Percent Decrease"
         Height          =   195
         Index           =   2
         Left            =   6990
         TabIndex        =   12
         Top             =   840
         Width           =   1290
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Round all prices to"
         Height          =   195
         Index           =   3
         Left            =   4140
         TabIndex        =   4
         Top             =   300
         Width           =   1320
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "By Dollar Value"
         Height          =   195
         Index           =   5
         Left            =   7200
         TabIndex        =   8
         Top             =   360
         Width           =   1080
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Percent Increase"
         Height          =   195
         Index           =   6
         Left            =   7065
         TabIndex        =   10
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblIncentiveCostPercent 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Incentive Cost%"
         Height          =   195
         Left            =   4320
         TabIndex        =   6
         Top             =   540
         Width           =   1140
      End
      Begin VB.Label lblSalesEffectiveDate 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Published to Sales"
         Height          =   195
         Left            =   90
         TabIndex        =   20
         Top             =   960
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Costs Effective"
         Height          =   195
         Index           =   9
         Left            =   330
         TabIndex        =   19
         Top             =   720
         Width           =   1065
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cost Basis"
         Height          =   195
         Index           =   1
         Left            =   660
         TabIndex        =   18
         Top             =   435
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Description"
         Height          =   195
         Index           =   0
         Left            =   600
         TabIndex        =   17
         Top             =   180
         Width           =   795
      End
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   15
      Top             =   0
      Width           =   12960
      _ExtentX        =   22860
      _ExtentY        =   1058
      ButtonWidth     =   2355
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Add"
            Key             =   "Add"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Remove"
            Key             =   "Delete"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "s1"
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Preview"
            Key             =   "Preview"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Costs"
            Key             =   "AssemblyCosts"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save"
            Key             =   "Save"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save As"
            Key             =   "SaveAs"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Publish Prices"
            Key             =   "Publish"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "s2"
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Rfrsh Costs"
            Key             =   "RePrice"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Rfrsh Published"
            Key             =   "LookupPublished"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VSFlex8Ctl.VSFlexGrid gData 
      Height          =   1275
      Left            =   300
      TabIndex        =   14
      Top             =   1980
      Width           =   11055
      _cx             =   1978158124
      _cy             =   1978140873
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
      SheetBorder     =   -2147483633
      FocusRect       =   1
      HighLight       =   2
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   3
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   191
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FPricingWorksheet.frx":01B8
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
      OutlineBar      =   5
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   7
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
      FrozenRows      =   1
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   -2147483624
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
End
Attribute VB_Name = "FPricingWorksheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Const SRCFILE = "FPricingWorksheet::"

Public Enum WorksheetTypes
    wsgeneral
    wsDesignCenter = 1
End Enum

Private mWorksheet          As Long
Private mWorksheetView      As String
Private mWorksheetType      As WorksheetTypes
Private mCommunity          As String
Private mCommunityPhase     As String
Private mCostEffectiveDate  As Date
Private mDirty              As Boolean
Private mReadOnly           As Boolean
      
Const DisabledText = vbGrayText
Const NewColumns = "Pretax,Tax,Total,Markup,Margin,Profit,COPretax,COTax,COTotal,COMarkup,COMargin,InSpec"
Const PublishedColumns = "PublishedPretax,PublishedTax,PublishedPrice,PublishedCOPrice,PublishedIncentive,PublishedProfit,PublishedMarkup,PublishedMargin," & _
                         "PublishedPretax2,PublishedTax2,PublishedPrice2,PublishedProfit2,PublishedMarkup2,PublishedMargin2," & _
                         "PublishedPretax3,PublishedTax3,PublishedPrice3,PublishedProfit3,PublishedMarkup3,PublishedMargin3," & _
                         "PublishedPretax4,PublishedTax4,PublishedPrice4,PublishedProfit4,PublishedMarkup4,PublishedMargin4," & _
                         "PublishedPretax5,PublishedTax5,PublishedPrice5,PublishedProfit5,PublishedMarkup5,PublishedMargin5," & _
                         "PublishedPretax6,PublishedTax6,PublishedPrice6,PublishedProfit6,PublishedMarkup6,PublishedMargin6," & _
                         "PublishedPretax7,PublishedTax7,PublishedPrice7,PublishedProfit7,PublishedMarkup7,PublishedMargin7," & _
                         "PublishedPretax8,PublishedTax8,PublishedPrice8,PublishedProfit8,PublishedMarkup8,PublishedMargin8," & _
                         "PublishedPretax9,PublishedTax9,PublishedPrice9,PublishedProfit9,PublishedMarkup9,PublishedMargin9," & _
                         "PublishedPretax10,PublishedTax10,PublishedPrice10,PublishedProfit10,PublishedMarkup10,PublishedMargin10"

Private MaxSpec As Long
Private VisibleMarketingColumns  As String
Private EnabledMarketingColumns  As String
Private VisibleEstimatingColumns As String
Private EnabledEstimatingColumns As String
Private MaxVendorPricing As Boolean


Private Property Get ColumnColor(ColumnClass As String) As Long
    Select Case ColumnClass
        Case "New":       ColumnColor = 16768477
        Case "Published": ColumnColor = 12381397
        Case "Spec2":     ColumnColor = 14076899
        Case "Spec3":     ColumnColor = 11456762
        Case "Spec4":     ColumnColor = 15527080
        Case "Spec5":     ColumnColor = 13752001
        Case "Spec6":     ColumnColor = 14024703
        Case "Spec7":     ColumnColor = 14076899
        Case "Spec8":     ColumnColor = 11456762
        Case "Spec9":     ColumnColor = 15527080
        Case "Spec10":    ColumnColor = 13752001
        Case Else:        ColumnColor = vbWindowBackground
    End Select
End Property







Private Sub chkArchive_Click()
    Dim s As String
    s = "update tblsalessheetmaster set inactive=" & DbQuote(Bit, chkArchive.value = vbChecked) & " where worksheet=" & DbQuote(Num, mWorksheet)
    Call HFApp.SqlExec(s)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case True
        Case KeyCode = vbKeyS And Shift > 1: Call Toolbar_ButtonClick(Toolbar.Buttons("Save"))
        Case KeyCode = vbKeyD And Shift > 1: Call Toolbar_ButtonClick(Toolbar.Buttons("Add"))
    End Select
End Sub


Private Sub gData_DblClick()
    Call Toolbar_ButtonClick(Toolbar.Buttons("AssemblyCosts"))
End Sub

Private Sub gData_RowColChange()
On Error Resume Next
    Dim r As Long
    Dim i As Long
    Dim c As Long
    Dim s As String
    
    With gData
        
        
        
        .Redraw = flexRDNone
        
        'disable all columns
        .Cell(flexcpForeColor, 2, 0, .Rows - 1, .Cols - 1) = DisabledText
            
            
        'color spec columns
        
            For i = 2 To 10
                .Cell(flexcpBackColor, 2, .ColIndex("Profit" & i), .Rows - 1, .ColIndex("Profit" & i)) = ColumnColor("Spec" & i)
                .Cell(flexcpBackColor, 2, .ColIndex("Total" & i), .Rows - 1, .ColIndex("Total" & i)) = ColumnColor("Spec" & i)
                .Cell(flexcpBackColor, 2, .ColIndex("Pretax" & i), .Rows - 1, .ColIndex("Pretax" & i)) = ColumnColor("Spec" & i)
                .Cell(flexcpBackColor, 2, .ColIndex("Tax" & i), .Rows - 1, .ColIndex("Tax" & i)) = ColumnColor("Spec" & i)
                .Cell(flexcpBackColor, 2, .ColIndex("Margin" & i), .Rows - 1, .ColIndex("Margin" & i)) = ColumnColor("Spec" & i)
                .Cell(flexcpBackColor, 2, .ColIndex("Markup" & i), .Rows - 1, .ColIndex("Markup" & i)) = ColumnColor("Spec" & i)
                .Cell(flexcpBackColor, 2, .ColIndex("InSpec" & i), .Rows - 1, .ColIndex("InSpec" & i)) = ColumnColor("Spec" & i)
            Next
        
            
            
        'color new price columns
        s = NewColumns
        
        For r = 2 To 10
            For i = 1 To Parse(s)
                c = .ColIndex(Parse(s, i))
                .Cell(flexcpBackColor, 2, c, .Rows - 1, c) = ColumnColor("New")
            Next
        Next
        
        
        
        
        'color Published columns
        s = PublishedColumns
        For i = 1 To Parse(s)
            c = .ColIndex(Parse(s, i))
            .Cell(flexcpBackColor, 2, c, .Rows - 1, c) = ColumnColor("Published")
        Next
        
        'enable editable columns
        s = EnabledColumns
        For i = 1 To Parse(s)
            c = .ColIndex(Parse(s, i))
            .Cell(flexcpForeColor, 2, c, .Rows - 1, c) = vbWindowText
        Next
        
        'clear some cells depending on assembly type
        r = .Row
        'For r = 1 To .Rows - 1
            Select Case .ValueMatrix(r, .ColIndex("AssemblyType"))
                Case atDesignCenter
                    .TextMatrix(r, .ColIndex("IncentiveCost")) = ""
                    .TextMatrix(r, .ColIndex("IncentiveRetail")) = ""
                    .TextMatrix(r, .ColIndex("LandCost")) = ""
                    .TextMatrix(r, .ColIndex("FloorArea")) = ""
                    .TextMatrix(r, .ColIndex("Bedrooms")) = ""
                    .TextMatrix(r, .ColIndex("Bathrooms")) = ""
                    .TextMatrix(r, .ColIndex("Style")) = ""
                    
                
                Case atModel
                    .TextMatrix(r, .ColIndex("ConstCutOff")) = ""
                    .TextMatrix(r, .ColIndex("Category")) = ""
                    .TextMatrix(r, .ColIndex("CategoryDesc")) = ""
                    .TextMatrix(r, .ColIndex("JCExtra")) = ""
                    .TextMatrix(r, .ColIndex("COTotal")) = ""
                    .TextMatrix(r, .ColIndex("COPretax")) = ""
                    .TextMatrix(r, .ColIndex("COTax")) = ""
                    .TextMatrix(r, .ColIndex("COMargin")) = ""
                    .TextMatrix(r, .ColIndex("COMarkup")) = ""
                    
                        For i = 2 To 10
                            .TextMatrix(r, .ColIndex("Pretax" & i)) = ""
                            .TextMatrix(r, .ColIndex("Tax" & i)) = ""
                            .TextMatrix(r, .ColIndex("Markup" & i)) = ""
                            .TextMatrix(r, .ColIndex("Margin" & i)) = ""
                            .TextMatrix(r, .ColIndex("Profit" & i)) = ""
                            .TextMatrix(r, .ColIndex("Total" & i)) = ""
                            .Cell(flexcpChecked, r, .ColIndex("InSpec" & i)) = flexNoCheckbox
                        Next
                    
                    .Cell(flexcpChecked, r, .ColIndex("InSpec")) = flexNoCheckbox
                    
                Case Else
                    .TextMatrix(r, .ColIndex("IncentiveCost")) = ""
                    .TextMatrix(r, .ColIndex("IncentiveRetail")) = ""
                    .TextMatrix(r, .ColIndex("LandCost")) = ""
                    .TextMatrix(r, .ColIndex("FloorArea")) = ""
                    .TextMatrix(r, .ColIndex("Bedrooms")) = ""
                    .TextMatrix(r, .ColIndex("Bathrooms")) = ""
                    .TextMatrix(r, .ColIndex("Style")) = ""
                    
                    For i = 2 To 10
                        .TextMatrix(r, .ColIndex("Pretax" & i)) = ""
                        .TextMatrix(r, .ColIndex("Tax" & i)) = ""
                        .TextMatrix(r, .ColIndex("Markup" & i)) = ""
                        .TextMatrix(r, .ColIndex("Margin" & i)) = ""
                        .TextMatrix(r, .ColIndex("Profit" & i)) = ""
                        .TextMatrix(r, .ColIndex("Total" & i)) = ""
                        .Cell(flexcpChecked, r, .ColIndex("InSpec" & i)) = flexNoCheckbox
                    Next
                    
                    .Cell(flexcpChecked, r, .ColIndex("InSpec")) = flexNoCheckbox
                    
            End Select
        'Next
        
        .Redraw = flexRDBuffered
    End With
End Sub





Private Sub LoadSeries()
    Dim i As Long
    Dim rs As Recordset
    
    
    'hide dc only stuff
    If mWorksheetType = wsDesignCenter Then
        '-------------------
        'load the legend
        '-------------------
        With gLegend
            'clear all
            .Cell(flexcpBackColor, 0, 0, 2, 3) = vbButtonFace
            'now load the 10 spec levels
            Set rs = HFApp.SqlExec("select reference_item, max(Description) Description from tblseries where DivisionID = " & HFApp.DivisionID & " and isnull(reference_item,'') <>'' Group By Reference_item")
            While Not rs.EOF
                i = i + 1
                Select Case i
                    Case 1:    .Cell(flexcpText, 0, 0) = "" & rs(1):    .Cell(flexcpBackColor, 0, 0) = ColumnColor("New")
                    
                    Case 2:    .Cell(flexcpText, 1, 0) = "" & rs(1):    .Cell(flexcpBackColor, 1, 0) = ColumnColor("Spec" & i)
                    Case 3:    .Cell(flexcpText, 2, 0) = "" & rs(1):    .Cell(flexcpBackColor, 2, 0) = ColumnColor("Spec" & i)
                    Case 4:    .Cell(flexcpText, 0, 1) = "" & rs(1):    .Cell(flexcpBackColor, 0, 1) = ColumnColor("Spec" & i)
                    Case 5:    .Cell(flexcpText, 1, 1) = "" & rs(1):    .Cell(flexcpBackColor, 1, 1) = ColumnColor("Spec" & i)
                    Case 6:    .Cell(flexcpText, 2, 1) = "" & rs(1):    .Cell(flexcpBackColor, 2, 1) = ColumnColor("Spec" & i)
                    Case 7:    .Cell(flexcpText, 0, 2) = "" & rs(1):    .Cell(flexcpBackColor, 0, 2) = ColumnColor("Spec" & i)
                    Case 8:    .Cell(flexcpText, 1, 2) = "" & rs(1):    .Cell(flexcpBackColor, 1, 2) = ColumnColor("Spec" & i)
                    Case 9:    .Cell(flexcpText, 2, 2) = "" & rs(1):    .Cell(flexcpBackColor, 2, 2) = ColumnColor("Spec" & i)
                    Case 10:   .Cell(flexcpText, 0, 3) = "" & rs(1):    .Cell(flexcpBackColor, 0, 3) = ColumnColor("Spec" & i)
                End Select
                rs.MoveNext
            Wend
            Call .AutoSize(0, 3, True)
        End With
        
        '-------------------
        'load the visible/enabled columns
        '-------------------
        On Error Resume Next
        MaxSpec = Val(HFApp.SqlExec("select max(substring(reference_item,5,2)) from tblseries where DivisionID = " & HFApp.DivisionID & " and isnull(reference_item,'') <>''")(0))
        If MaxSpec > 0 Then
            VisibleMarketingColumns = "LandCost,Community,CommunityDesc,CommunityPhase,Assembly,Description,PublishedPretax,PublishedTax,PublishedPrice,PublishedCOPrice,PublishedIncentive,PublishedProfit,PublishedMarkup,PublishedMargin,CategoryDesc,Pretax,Tax,Total,IncentiveRetail,Profit,Markup,Margin,Roundto,AssemblyUOM,FloorArea,Color,Location,Qty,IncludedOption,Series"
            EnabledMarketingColumns = "Total,Pretax,IncentiveRetail,Profit,Markup,Margin,Roundto,AssemblyUOM,FloorArea,Color,Location,Qty"
            
            VisibleEstimatingColumns = "Model,ModelDescription,Elevation,Assembly,OptionID,Description,PublishedPretax,PublishedTax,PublishedPrice,PublishedCOPrice,PublishedIncentive,PublishedProfit,PublishedMarkup,PublishedMargin,Notes,Comments,JCExtra,Series,FloorArea,Bedrooms,Bathrooms,Style,ConstCutOff,Category,CategoryDesc,Pretax,Tax,Total,IncentiveRetail,Profit,Markup,Margin,Roundto,IncentiveCost,LandCost,ConstructionCost,Cost,AssemblyUOM,InSpec,Color,Location,Qty,IncludedOption,Series"
            EnabledEstimatingColumns = "Description,Notes,Comments,JCExtra,Series,FloorArea,Bedrooms,Bathrooms,Style,ConstCutOff,Category,CategoryDesc,Pretax,Total,IncentiveRetail,Profit,Markup,Margin,Roundto,LandCost,AssemblyUOM,InSpec,Color,Location,Qty,IncludedOption,Series"
            
            For i = 2 To MaxSpec
                VisibleMarketingColumns = VisibleMarketingColumns & ",Pretax" & i & ",Tax" & i & ",Markup" & i & ",Margin" & i & ",Profit" & i & ",Total" & i & ",PublishedPretax" & i & ",PublishedTax" & i & ",PublishedPrice" & i & ",PublishedProfit" & i & ",PublishedMarkup" & i & ",PublishedMargin" & i
                VisibleEstimatingColumns = VisibleEstimatingColumns & ",Pretax" & i & ",Tax" & i & ",Markup" & i & ",Margin" & i & ",InSpec" & i & ",Profit" & i & ",Total" & i & ",PublishedPretax" & i & ",PublishedTax" & i & ",PublishedPrice" & i & ",PublishedProfit" & i & ",PublishedMarkup" & i & ",PublishedMargin" & i
                EnabledMarketingColumns = EnabledMarketingColumns & ",Pretax" & i & ",Markup" & i & ",Margin" & i & ",Profit" & i & ",Total" & i
                EnabledEstimatingColumns = EnabledEstimatingColumns & ",Pretax" & i & ",Markup" & i & ",Margin" & i & ",InSpec" & i & ",Profit" & i & ",Total" & i
            Next
        End If
    Else
        '-------------------
        'hide the legend
        '-------------------
        gLegend.Visible = False
        lblLegend.Visible = False
        '-------------------
        'load the visible/enabled columns
        '-------------------
        VisibleMarketingColumns = "PublishedPrice,PublishedCOPrice,Total,COTotal,LandCost,Community,CommunityDesc,CommunityPhase,Assembly,Description,PublishedPretax,PublishedTax,PublishedIncentive,PublishedProfit,PublishedMarkup,PublishedMargin,CategoryDesc,Pretax,Tax,IncentiveRetail,Profit,Markup,Margin,COMarkup,COMargin,Roundto,COPretax,COTax,AssemblyUOM,FloorArea,Color,Location,Qty,IncludedOption,CostPerFt,PricePerFt"
        EnabledMarketingColumns = "Total,Pretax,IncentiveRetail,Profit,Markup,Margin,COMarkup,COMargin,Roundto,COPretax,COTotal,AssemblyUOM,FloorArea,Color,Location,Qty"
        VisibleEstimatingColumns = "Inactive,Model,ModelDescription,Elevation,Assembly,OptionID,Description,PublishedPretax,PublishedTax,PublishedPrice,PublishedCOPrice,PublishedIncentive,PublishedProfit,PublishedMarkup,PublishedMargin,Notes,Comments,JCExtra,Series,FloorArea,Bedrooms,Bathrooms,Style,ConstCutOff,Category,CategoryDesc,Pretax,Tax,Total,IncentiveRetail,Profit,Markup,Margin,COMarkup,COMargin,Roundto,IncentiveCost,LandCost,Cost,ConstructionCost,COPretax,COTax,COTotal,AssemblyUOM,Color,Location,Qty,IncludedOption,CostPerFt,PricePerFt"
        EnabledEstimatingColumns = "Inactive,Description,Notes,Comments,JCExtra,Series,FloorArea,Bedrooms,Bathrooms,Style,ConstCutOff,Category,CategoryDesc,Pretax,Total,IncentiveRetail,Profit,Markup,Margin,COMarkup,COMargin,Roundto,LandCost,COPretax,COTotal,AssemblyUOM,Color,Location,Qty,IncludedOption"
        
    End If


    
    
    
End Sub

Private Function VisibleColumns() As String
    Dim s As String
    Select Case mWorksheetView
        Case "Marketing":  s = VisibleMarketingColumns
        Case Else:         s = VisibleEstimatingColumns
    End Select
    s = s & ",IncludeTax"
    s = s & ",Community,CommunityDesc"
    s = s & ",CommunityPhase"
    
    VisibleColumns = s
End Function
Private Function EnabledColumns() As String
    Dim s As String
    Select Case mWorksheetView
        Case "Marketing":  s = EnabledMarketingColumns
        Case Else:         s = EnabledEstimatingColumns
    End Select
    s = s & ",IncludeTax"
    s = s & ",Community,CommunityDesc"
    s = s & ",CommunityPhase"
    EnabledColumns = s
End Function




Private Sub gLegend_GotFocus()
On Error Resume Next
    gData.SetFocus
End Sub

Private Sub txtDollarChange_GotFocus()
    SelectAll txtDollarChange
End Sub

Private Sub txtPercentDecrease_GotFocus()
    SelectAll txtPercentDecrease
End Sub

Private Sub txtPercentIncrease_GotFocus()
    SelectAll txtPercentIncrease
End Sub

Private Sub txtRoundTo_GotFocus()
    SelectAll txtRoundTo
End Sub

Private Sub txtRoundTo_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    Dim b As Boolean
    If KeyCode = vbKeyReturn Then Call txtRoundTo_Validate(b)
End Sub
Private Sub txtRoundTo_Validate(Cancel As Boolean)
    If Trim(txtRoundTo.Text) = "" Then Exit Sub
    If Not IsNumeric(txtRoundTo.Text) Then
        Cancel = True
    Else
        If gData.Rows > 2 Then gData.Cell(flexcpText, 2, gData.ColIndex("Roundto"), gData.Rows - 1, gData.ColIndex("Roundto")) = Val(txtRoundTo.Text)
        Call CalcData(True, "Roundto")
        txtRoundTo.Text = ""
    End If
End Sub



Private Sub txtDollarChange_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    Dim b As Boolean
    If KeyCode = vbKeyReturn Then Call txtDollarChange_Validate(b)
End Sub
Private Sub txtDollarChange_Validate(Cancel As Boolean)
    Dim i As Long
    Dim Field As String
    
    With gData
    If Trim(txtDollarChange.Text) = "" Then Exit Sub
    If Not IsNumeric(txtDollarChange.Text) Then
        Cancel = True
    Else
        For i = 2 To .Rows - 1
            Field = IIf(.Cell(flexcpChecked, i, .ColIndex("IncludeTax")) = flexChecked, "Total", "Pretax")
            .TextMatrix(i, .ColIndex(Field)) = .ValueMatrix(i, gData.ColIndex(Field)) + Val(txtDollarChange.Text)
            Call CalcData(False, Field, i)
        Next
        txtDollarChange.Text = ""
    End If
    End With
    
End Sub

Private Sub txtPercentIncrease_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    Dim b As Boolean
    If KeyCode = vbKeyReturn Then Call txtPercentIncrease_Validate(b)
End Sub
Private Sub txtPercentIncrease_Validate(Cancel As Boolean)
    Dim i As Long
    Dim Field As String
    
    With gData
    If Trim(txtPercentIncrease.Text) = "" Then Exit Sub
    If Not IsNumeric(txtPercentIncrease.Text) Then
        Cancel = True
    Else
        For i = 2 To .Rows - 1
            Field = IIf(.Cell(flexcpChecked, i, .ColIndex("IncludeTax")), "Total", "Pretax")
            .TextMatrix(i, gData.ColIndex(Field)) = .ValueMatrix(i, gData.ColIndex(Field)) * (100 + Val(txtPercentIncrease.Text)) / 100
        Next
        Call CalcData(True, Field)
        txtPercentIncrease.Text = ""
    End If
    End With
End Sub

Private Sub txtPercentDecrease_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    Dim b As Boolean
    If KeyCode = vbKeyReturn Then Call txtPercentDecrease_Validate(b)
End Sub
Private Sub txtPercentDecrease_Validate(Cancel As Boolean)
    Dim i As Long
    Dim Field As String
    
    With gData
    If Trim(txtPercentDecrease.Text) = "" Then Exit Sub
    If Not IsNumeric(txtPercentDecrease.Text) Then
        Cancel = True
    Else
        For i = 2 To .Rows - 1
            Field = IIf(.Cell(flexcpChecked, i, .ColIndex("IncludeTax")), "Total", "Pretax")
            .TextMatrix(i, gData.ColIndex(Field)) = .ValueMatrix(i, gData.ColIndex(Field)) * (100 - Val(txtPercentDecrease.Text)) / 100
        Next
        Call CalcData(True, Field)
        txtPercentDecrease.Text = ""
    End If
    End With
End Sub


Private Property Get ReadOnly() As Boolean
    ReadOnly = mReadOnly
End Property

Private Property Let ReadOnly(RHS As Boolean)
Dim i As Long
    mReadOnly = RHS
    
    For i = 1 To Toolbar.Buttons.Count
        Toolbar.Buttons(i).Enabled = Not mReadOnly
    Next
    Toolbar.Buttons("SaveAs").Enabled = True
    Toolbar.Buttons("Preview").Enabled = True
    
    txtDescription.Enabled = Not mReadOnly
    cboCostBasis.Enabled = Not mReadOnly
    txtIncentiveCostPercent.Enabled = Not mReadOnly
    txtCostsEffectiveDate.Enabled = Not mReadOnly
    txtSalesEffectiveDate.Enabled = Not mReadOnly
    
    txtDollarChange.Enabled = Not mReadOnly
    txtPercentDecrease.Enabled = Not mReadOnly
    txtPercentIncrease.Enabled = Not mReadOnly
    txtRoundTo.Enabled = Not mReadOnly
    
End Property




Private Sub cboCostBasis_Click()
    mDirty = True
End Sub

Private Sub cboCostBasis_GotFocus()
    SelectAll cboCostBasis
End Sub

Public Sub OpenWorksheet(WorksheetType As WorksheetTypes, Worksheet As Long, WorksheetView As String, community As String, CommunityPhase As String)
        
    mWorksheet = Worksheet
    mWorksheetView = WorksheetView
    mWorksheetType = WorksheetType
    mCommunity = community
    mCommunityPhase = CommunityPhase
    
    
    
    Me.Show
End Sub



Private Sub Form_Unload(Cancel As Integer)
    
    If Not SaveData(True) Then
        Cancel = True
        Exit Sub
    End If
    Call HFApp.UnLockRecord("tblSalesSheetMaster", mWorksheet)
    Call IniPutForm(Me, , mWorksheetView)
    Call IniPutGrid(Me, gData, , mWorksheetView & cboLayouts.Text)
    Call SaveGridLayouts
    
    'just to be safe.
    Unload FComments
End Sub

Private Sub Form_Resize()
On Error Resume Next
    HeadingFrame.Move 0, Toolbar.Height, Me.ScaleWidth
    gData.Move 0, HeadingFrame.Top + HeadingFrame.Height, Me.ScaleWidth, Me.ScaleHeight - Toolbar.Height - HeadingFrame.Height
End Sub

Private Sub LoadData()
    Dim s As String
    Dim rs As Recordset
    Dim r As Long
    Dim i As Long
    Dim TaxIncl As Boolean
    ReadOnly = False
    
    
    
    
    
    
    s = ""
    s = s & "SELECT *" & vbCrLf
    s = s & "  FROM tblSalesSheetMaster" & vbCrLf
    s = s & " WHERE Worksheet=" & DbQuote(Num, mWorksheet) & vbCrLf
    Set rs = HFApp.SqlExec(s)
    If mWorksheet <> 0 Or Not rs.EOF Then
        chkArchive.value = IIf("" & rs("Inactive") = "True", vbChecked, vbUnchecked)
        txtDescription.Text = "" & rs("Description")
        cboCostBasis.ListIndex = Val("" & rs("CostBasisType"))
        txtIncentiveCostPercent.Text = Val("" & rs("IncentiveCostPercent"))
        CostEffectiveDate = rs("CostsEffectiveDate")
        txtSalesEffectiveDate.Text = format("" & rs("SalesEffectiveDate"), "mmmm d, yyyy")
    Else
        chkArchive.value = vbUnchecked
        txtDescription.Text = "New Worksheet"
        cboCostBasis.ListIndex = 0
        txtIncentiveCostPercent.Text = Val("" & HFApp.Options(IncentiveCostPercent))
        CostEffectiveDate = Now
        txtSalesEffectiveDate.Text = ""
    End If
    txtIncentiveCostPercent.Enabled = Trim(HFApp.Options(IncentiveItem)) <> ""
    ReadOnly = Trim(txtSalesEffectiveDate.Text) <> "" And HFApp.Options(LockPostedSalesWorksheets)


    With gData
    
        .TextMatrix(0, .ColIndex("COMargin")) = ""
        .ColHidden(.ColIndex("COMargin")) = True
        
        .Redraw = flexRDNone
        .Rows = 2
        r = r + 1
        s = ""
        s = s & "SELECT l.Description CommunityDesc,c.Description CategoryDesc,dm.Description ModelDescription, d.*" & vbCrLf
        s = s & "  FROM tblSalesSheetDetails d" & vbCrLf
        s = s & "  join tblSalesSheetMaster m on m.Worksheet = d.Worksheet" & vbCrLf
        s = s & "       LEFT OUTER JOIN tblLocality l ON(d.Community=l.Area)" & vbCrLf
        If HFApp.DivisionID <> "" Then
            s = s & " left outer join DivisionCommunities d1 on d1.Community = l.Area" & vbCrLf
        End If
        s = s & "       LEFT OUTER JOIN tblCategories c ON(d.Category=c.Category)" & vbCrLf
        s = s & "       left outer join distinctmodelsbyDivision dm on(dm.DivisionID = m.DivisionID and d.model=dm.model)" & vbCrLf
        s = s & " WHERE d.Worksheet=" & DbQuote(Num, mWorksheet) & vbCrLf
        If HFApp.DivisionID <> "" Then
            s = s & " and (d1.DivisionID is null or d1.Divisionid = " & HFApp.DivisionID & ")" & vbCrLf
        End If
        s = s & "ORDER BY d.Model,d.Assembly" & vbCrLf
        Set rs = HFApp.SqlExec(s)
        .TextMatrix(0, .ColIndex("Community")) = FMain.CD_Community
        .TextMatrix(0, .ColIndex("CommunityPhase")) = FMain.CD_Community & " Phase"
        .TextMatrix(0, .ColIndex("CommunityDesc")) = FMain.CD_Community & " Description"
        While Not rs.EOF
            r = r + 1
            .AddItem ""
             If rs("AssemblyType") = atoption Then
                .TextMatrix(r, .ColIndex("ModelDescription")) = "" & rs("ModelDescription")
             End If
             TaxIncl = IIf("" & rs("IncludeTax") = "True", True, False)
            .TextMatrix(r, .ColIndex("SourceCommunity")) = "" & rs("SourceCommunity")
            .TextMatrix(r, .ColIndex("Community")) = "" & rs("Community")
            .TextMatrix(r, .ColIndex("CommunityPhase")) = "" & rs("CommunityPhase")
            .TextMatrix(r, .ColIndex("CommunityDesc")) = "" & rs("CommunityDesc")
            
            
            .TextMatrix(r, .ColIndex("Assembly")) = "" & rs("Assembly")
            .TextMatrix(r, .ColIndex("AssemblyUOM")) = "" & rs("AssemblyUOM")
            .TextMatrix(r, .ColIndex("Model")) = "" & rs("Model")
            .TextMatrix(r, .ColIndex("Elevation")) = "" & rs("Elevation")
            .TextMatrix(r, .ColIndex("OptionID")) = "" & rs("OptionID")
            
            
            .TextMatrix(r, .ColIndex("DCSalesOnly")) = "" & rs("DesignCenterSalesOnly")
            .TextMatrix(r, .ColIndex("SelectByRoom")) = "" & rs("SelectByRoom")
            .TextMatrix(r, .ColIndex("DisplayTotalOnly")) = "" & rs("DisplayTotalOnly")
            
            .TextMatrix(r, .ColIndex("JCExtra")) = "" & rs("JCExtra")
            .TextMatrix(r, .ColIndex("Series")) = "" & rs("Series")
            .TextMatrix(r, .ColIndex("FloorArea")) = Val("" & rs("FloorArea"))
            .TextMatrix(r, .ColIndex("Bedrooms")) = "" & rs("Bedrooms")
            .TextMatrix(r, .ColIndex("Bathrooms")) = "" & rs("Bathrooms")
            .TextMatrix(r, .ColIndex("Style")) = "" & rs("Style")
            
            .TextMatrix(r, .ColIndex("ConstCutOff")) = Val("" & rs("ConstCutOff"))
            .TextMatrix(r, .ColIndex("SpecDocument")) = "" & rs("SpecDocument")
            .TextMatrix(r, .ColIndex("GraphicPath")) = "" & rs("GraphicPath")
            .TextMatrix(r, .ColIndex("MaxWidth")) = Val("" & rs("MaxWidth"))
            .TextMatrix(r, .ColIndex("MaxLength")) = Val("" & rs("MaxLength"))
            .Cell(flexcpChecked, r, .ColIndex("Inactive")) = IIf("" & rs("Inactive") = "True", flexChecked, flexUnchecked)

            
            .TextMatrix(r, .ColIndex("Color")) = "" & rs("Color")
            .TextMatrix(r, .ColIndex("Location")) = "" & rs("Location")
            .TextMatrix(r, .ColIndex("Qty")) = "" & rs("Qty")
            If .TextMatrix(r, .ColIndex("Qty")) = "0" Then .TextMatrix(r, .ColIndex("Qty")) = "1"
            .TextMatrix(r, .ColIndex("IncludedOption")) = "" & rs("IncludedOption")
            .Cell(flexcpChecked, r, .ColIndex("ReadyToPublish")) = IIf("" & rs("ReadyToPublish") = "True", flexChecked, flexUnchecked)
                       
                       
            .TextMatrix(r, .ColIndex("ColorListID")) = "" & rs("ColorListID")
            .TextMatrix(r, .ColIndex("StyleListID")) = "" & rs("StyleListID")
            .TextMatrix(r, .ColIndex("FinishListID")) = "" & rs("FinishListID")
            .TextMatrix(r, .ColIndex("OtherListID")) = "" & rs("OtherListID")
            .TextMatrix(r, .ColIndex("StyleValue")) = "" & rs("StyleValue")
            .TextMatrix(r, .ColIndex("FinishValue")) = "" & rs("FinishValue")
            .TextMatrix(r, .ColIndex("OtherValue")) = "" & rs("OtherValue")
                       
            
            .TextMatrix(r, .ColIndex("Description")) = "" & rs("Description")
            .TextMatrix(r, .ColIndex("Notes")) = "" & rs("Notes")
            .TextMatrix(r, .ColIndex("Comments")) = "" & rs("Comments")
            .TextMatrix(r, .ColIndex("Assemblytype")) = "" & rs("AssemblyType")
            .TextMatrix(r, .ColIndex("Category")) = "" & rs("Category")
            .TextMatrix(r, .ColIndex("CategoryDesc")) = "" & rs("CategoryDesc")
            
            .TextMatrix(r, .ColIndex("Markup")) = Val("" & rs("Markup"))
            .TextMatrix(r, .ColIndex("Margin")) = Val("" & rs("Margin"))
            
            .TextMatrix(r, .ColIndex("COMarkup")) = Val("" & rs("COMarkup"))
            .TextMatrix(r, .ColIndex("COMargin")) = Val("" & rs("COMargin"))
            
            .TextMatrix(r, .ColIndex("Roundto")) = Val("" & rs("Roundto"))
            
            .TextMatrix(r, .ColIndex("LandCost")) = Val("" & rs("LandCost"))
            .TextMatrix(r, .ColIndex("IncentiveRetail")) = Val("" & rs("IncentiveRetail"))
            .TextMatrix(r, .ColIndex("IncentiveCost")) = Val("" & rs("IncentiveCost"))
            .TextMatrix(r, .ColIndex("ConstructionCost")) = Val("" & rs("ConstructionCost"))
            .TextMatrix(r, .ColIndex("Cost")) = Val("" & rs("Cost"))
            If Not TaxIncl Then
                .TextMatrix(r, .ColIndex("Pretax")) = RoundToPrice(Val("" & rs("Pretax")), Val("" & rs("Roundto")))
            Else
                .TextMatrix(r, .ColIndex("Pretax")) = Val("" & rs("Pretax"))
            End If
            .TextMatrix(r, .ColIndex("Tax")) = Val("" & rs("Tax"))
            .TextMatrix(r, .ColIndex("Total")) = Val("" & rs("Pretax")) + Val("" & rs("Tax"))
            If Not TaxIncl Then
                .TextMatrix(r, .ColIndex("COPretax")) = RoundToPrice(Val("" & rs("COPretax")), Val("" & rs("Roundto")))
            Else
                .TextMatrix(r, .ColIndex("COPretax")) = Val("" & rs("COPretax"))
            End If
            .TextMatrix(r, .ColIndex("COTax")) = Val("" & rs("COTax"))
            .TextMatrix(r, .ColIndex("COTotal")) = Val("" & rs("COPretax")) + Val("" & rs("COTax"))
            
            .Cell(flexcpChecked, r, .ColIndex("InSpec")) = IIf("" & rs("IncludedInSpec") = "True", flexChecked, flexUnchecked)
            For i = 2 To 10
            If Not TaxIncl Then
                .TextMatrix(r, .ColIndex("Pretax" & i)) = RoundToPrice(Val("" & rs("Pretax" & i)), Val("" & rs("Roundto")))
            Else
                .TextMatrix(r, .ColIndex("Pretax" & i)) = Val("" & rs("Pretax" & i))
            End If
            .TextMatrix(r, .ColIndex("Tax" & i)) = Val("" & rs("Tax" & i))
            .TextMatrix(r, .ColIndex("Markup" & i)) = Val("" & rs("Markup" & i))
            .TextMatrix(r, .ColIndex("Margin" & i)) = Val("" & rs("Margin" & i))
            .Cell(flexcpChecked, r, .ColIndex("InSpec" & i)) = IIf("" & rs("IncludedInSpec" & i) = "True", flexChecked, flexUnchecked)
            .TextMatrix(r, .ColIndex("Total" & i)) = .ValueMatrix(r, .ColIndex("Pretax" & i)) - .ValueMatrix(r, .ColIndex("Cost"))
            .TextMatrix(r, .ColIndex("Profit" & i)) = .ValueMatrix(r, .ColIndex("Pretax" & i)) - .ValueMatrix(r, .ColIndex("Cost"))
            Next
            
            .Cell(flexcpChecked, r, .ColIndex("IncludeTax")) = IIf("" & rs("IncludeTax") = "True", flexChecked, flexUnchecked)
            
            
            .TextMatrix(r, .ColIndex("Profit")) = .ValueMatrix(r, .ColIndex("Pretax")) - .ValueMatrix(r, .ColIndex("Cost"))

            
            If .ValueMatrix(r, .ColIndex("FloorArea")) = 0 Then
                .TextMatrix(r, .ColIndex("CostPerFt")) = ""
                .TextMatrix(r, .ColIndex("PricePerFt")) = ""
            Else
                .TextMatrix(r, .ColIndex("CostPerFt")) = .ValueMatrix(r, .ColIndex("ConstructionCost")) / .ValueMatrix(r, .ColIndex("FloorArea"))
                .TextMatrix(r, .ColIndex("PricePerFt")) = .ValueMatrix(r, .ColIndex("Pretax")) / .ValueMatrix(r, .ColIndex("FloorArea"))
            End If
            
            rs.MoveNext
        Wend
        .Redraw = flexRDBuffered
    End With
    Call gData_RowColChange 'this sets the column coloring
    'Call RefreshPublished(False, True)
    Call gData_RowColChange
    mDirty = False
    
    

End Sub

Private Sub gData_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
    Dim r As Long
    With gData
        If Button = vbRightButton Then
        If .MouseRow < 1 Or ReadOnly Then
            Cancel = True
            Call FMain.ShowColumnMenu(gData, True, , , False)
        Else
            FMain.mnuSalesSheetSub(0).Enabled = mWorksheetView = "Estimating"
            FMain.mnuSalesSheetSub(4).Enabled = mWorksheetView = "Estimating"
            Set MouseCtrl = gData
            MouseCol = gData.MouseCol
            PopupMenu FMain.mnuSalesSheet
        End If
        End If
    End With
End Sub

Private Sub gData_BeforeSort(ByVal Col As Long, Order As Integer)
    gData.FixedRows = 2
End Sub
Private Sub gData_AfterSort(ByVal Col As Long, Order As Integer)
    gData.FixedRows = 1
End Sub

Private Sub gData_AfterSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long)
On Error Resume Next
'change selection to only single column
    Static bInHere As Boolean
    If bInHere Then Exit Sub
    bInHere = True
    If gData.Row = 1 And gData.RowSel <> 1 Then gData.RowSel = 1
    If gData.RowSel = 1 And gData.Row <> 1 Then gData.RowSel = 2
    gData.ColSel = gData.Col
    bInHere = False
End Sub


Private Sub ApplyFilter()
' apply filter to hide stuff that doesn't match
    Dim r As Long
    Dim c As Long

    Screen.MousePointer = vbHourglass
    With gData
        .Redraw = flexRDNone
        For r = 2 To .Rows - 1
            .RowHidden(r) = False
            For c = 0 To .Cols - 1
                If Not (UCase(.Cell(flexcpTextDisplay, r, c)) Like "*" & UCase(.Cell(flexcpTextDisplay, 1, c)) & "*") Then
                    .RowHidden(r) = True
                    Exit For
                End If
            Next
        Next
        .Redraw = flexRDBuffered
    
    End With
    Screen.MousePointer = vbDefault

End Sub

Private Sub gData_KeyDown(KeyCode As Integer, Shift As Integer)
        
    Dim r As Long
    Dim r1 As Long
    Dim r2 As Long
    Dim c As Long
    Dim Cancel As Boolean
    Dim s As String
    Dim clip As String
    Dim i As Long
    
    With gData
    Select Case True
        
        
        Case KeyCode = vbKeyG And Shift = vbCtrlMask
            r = Val(InputBox("Go to row:", App.ProductName))
            If r >= 1 And r < .Rows Then
                .Row = GridVisibleRowIndex(gData, r, 2)
                .SetFocus
            End If
        
        Case KeyCode = vbKeyF And Shift = vbCtrlMask
            Call FFind.ShowForm(gData)
        
        Case KeyCode = vbKeyDelete And Shift = vbCtrlMask And Not ReadOnly
            If mWorksheetView = "Estimating" Then
                mDirty = True
                For r = Max(.Row, .RowSel) To Max(2, Min(.Row, .RowSel)) Step -1
                If Not gData.RowHidden(r) Then
                    Call gData.RemoveItem(r)
                End If
                Next
                
                .RowSel = .Row
                Call CalcData(False, "just reset all tax included checkbox")
            End If
            
        
        Case KeyCode = vbKeyDelete And Not ReadOnly
            If mWorksheetView = "Estimating" Then
                mDirty = True
                For r = Max(.Row, .RowSel) To Max(2, Min(.Row, .RowSel)) Step -1
                    If Not gData.RowHidden(r) Then
                        For c = Min(.Col, .ColSel) To Max(.Col, .ColSel)
                            If IsIn(.ColKey(c), "Community", "CommunityDesc") Then
                                .TextMatrix(r, .ColIndex("Community")) = ""
                                .TextMatrix(r, .ColIndex("CommunityDesc")) = ""
                                .TextMatrix(r, .ColIndex("CommunityPhase")) = ""
                            End If
                            If IsIn(.ColKey(c), "CommunityPhase") Then
                                .TextMatrix(r, c) = ""
                            End If
                        Next
                    End If
                Next
                
                .RowSel = .Row
                Call CalcData(False, "just reset all tax included checkbox")
            End If
        
        Case Shift = vbCtrlMask And KeyCode = vbKeyC
            Clipboard.Clear
            Clipboard.SetText .clip
                        
        Case Shift = vbCtrlMask And KeyCode = vbKeyV And Not ReadOnly
            If .Row < 2 Then Exit Sub
                
            'get and sanitize clipboard
            clip = Clipboard.GetText
            
            '  don't do this it removes blank rows from the selection
'            clip = Replace(clip, vbCr, Chr(1))
'            clip = Replace(clip, vbLf, Chr(1))
'            clip = Replace(clip, Chr(1) & Chr(1), Chr(1))
'            clip = Replace(clip, Chr(1), vbCr)
'            clip = Trim(clip)
'            While Left(clip, 1) = vbCr
'                clip = Mid(clip, 2)
'            Wend
            
            c = .Col
            r1 = Min(.Row, .RowSel)
            r2 = r1
            r = r1
            For i = 1 To Parse(clip, , vbCr)
                s = Parse(clip, i, vbCr) 'get row
                s = Parse(s, 1, vbTab) 'get cell
                If s <> "" And r < .Rows Then
                    Call gData_BeforeEdit(r, c, Cancel)
                    If Cancel Or .ComboList = "..." Then
                        'skip this row
                    Else
                        .Row = r
                        Call gData_MYValidateEdit(r, c, Cancel, s)
                        .TextMatrix(r, c) = s
                        Call CalcData(, , r)
                        r2 = r
                    End If
                End If
                r = r + 1
            Next
            .Row = r1
            .RowSel = r2
            
    End Select
    End With
    
End Sub


Private Sub CalcData(Optional DoAllRows As Boolean = False, Optional EffectiveColumn As String, Optional EffectiveRow As Long = -1)
On Error GoTo eh
       
    Dim origR  As Long
    Dim startR As Long
    Dim stopR  As Long
    Dim i As Long
    Dim s As String
    Dim r As Long
    Dim TaxIncl As Boolean
    Dim CalcSelling As Boolean
    Dim CalcIncentive As Boolean
    Dim Multiple As Double
    Dim AssemblyType As AssemblyTypes
    
    Dim r1 As Long
    Dim r2 As Long
    Dim c1 As Long
    Dim c2 As Long
    
    
    With gData
        'no data rows
        If .Rows < 3 Then Exit Sub
        
        
        origR = .Row
        mDirty = True
        
        If EffectiveColumn = "" Then
            EffectiveColumn = .ColKey(.Col)
        End If
        
        If DoAllRows Then
            startR = 2
            stopR = .Rows - 1
        Else
            If EffectiveRow <> -1 Then
                startR = EffectiveRow
                stopR = EffectiveRow
            Else
                startR = Min(.Row, .RowSel)
                stopR = Max(.Row, .RowSel)
                
            End If
        End If
        If startR < 2 Then startR = 2
        
        CalcSelling = HFApp.Options(AdjustSellPriceWithCosts)
        CalcIncentive = HFApp.Options.ValueByName("AdjustIncentiveRetailWithCosts") <> "False"
        AssemblyType = .ValueMatrix(r, .ColIndex("AssemblyType"))
        
        For r = startR To stopR
If Not .RowHidden(r) Then
            
            'ensure taxincluded flag is set
'            If HFApp.Options(includetax) Then .Cell(flexcpChecked, r, .ColIndex("IncludeTax")) = flexChecked
            
            
            'ensure everything is numeric
            If .ValueMatrix(r, .ColIndex("AssemblyType")) = AssemblyTypes.atModel Then
                .TextMatrix(r, .ColIndex("FloorArea")) = .ValueMatrix(r, .ColIndex("FloorArea"))
                .TextMatrix(r, .ColIndex("IncentiveRetail")) = .ValueMatrix(r, .ColIndex("IncentiveRetail"))
                .TextMatrix(r, .ColIndex("IncentiveCost")) = .ValueMatrix(r, .ColIndex("IncentiveCost"))
            End If
            .TextMatrix(r, .ColIndex("Pretax")) = .ValueMatrix(r, .ColIndex("Pretax"))
            .TextMatrix(r, .ColIndex("Tax")) = .ValueMatrix(r, .ColIndex("Tax"))
            .TextMatrix(r, .ColIndex("Total")) = .ValueMatrix(r, .ColIndex("Total"))
            .TextMatrix(r, .ColIndex("LandCost")) = .ValueMatrix(r, .ColIndex("LandCost"))
            .TextMatrix(r, .ColIndex("Cost")) = .ValueMatrix(r, .ColIndex("Cost"))
            .TextMatrix(r, .ColIndex("COPretax")) = .ValueMatrix(r, .ColIndex("COPretax"))
            .TextMatrix(r, .ColIndex("COTax")) = .ValueMatrix(r, .ColIndex("COTax"))
            .TextMatrix(r, .ColIndex("COTotal")) = .ValueMatrix(r, .ColIndex("COTotal"))
            .TextMatrix(r, .ColIndex("Markup")) = .ValueMatrix(r, .ColIndex("Markup"))
            .TextMatrix(r, .ColIndex("Margin")) = .ValueMatrix(r, .ColIndex("Margin"))
            .Cell(flexcpChecked, r, .ColIndex("ReadyToPublish")) = flexChecked
            
            If .ValueMatrix(r, .ColIndex("AssemblyType")) = 4 Then
                For i = 2 To 10
                    .TextMatrix(r, .ColIndex("Pretax" & s)) = .ValueMatrix(r, .ColIndex("Pretax" & s))
                    .TextMatrix(r, .ColIndex("Tax" & s)) = .ValueMatrix(r, .ColIndex("Tax" & s))
                    .TextMatrix(r, .ColIndex("Markup" & s)) = .ValueMatrix(r, .ColIndex("Markup" & s))
                    .TextMatrix(r, .ColIndex("Margin" & s)) = .ValueMatrix(r, .ColIndex("Margin" & s))
                    If .ValueMatrix(r, .ColIndex("Margin" & s)) = 100 Then .TextMatrix(r, .ColIndex("Margin" & s)) = 99.99
                Next
            End If
            
            If .ValueMatrix(r, .ColIndex("Margin")) = 100 Then .TextMatrix(r, .ColIndex("Margin")) = 99.99
            
            .TextMatrix(r, .ColIndex("Roundto")) = .ValueMatrix(r, .ColIndex("Roundto"))
            
            TaxIncl = .Cell(flexcpChecked, r, .ColIndex("IncludeTax")) = flexChecked
            Multiple = .ValueMatrix(r, .ColIndex("Roundto"))

            
            
            'now recalc fields depending on which one they just changed
            Select Case EffectiveColumn
                                        
                Case "Cost", "LandCost", "IncentiveRetail", "IncentiveCost", "ConstructionCost"
                
                    If CalcIncentive Then
                        .TextMatrix(r, .ColIndex("IncentiveRetail")) = Round(.ValueMatrix(r, .ColIndex("IncentiveCost")) / IIf(Val(Me.txtIncentiveCostPercent.Text) = 0, 1, Val(Me.txtIncentiveCostPercent.Text)) * 100, 2)
                    Else
                        .TextMatrix(r, .ColIndex("IncentiveCost")) = Round(.ValueMatrix(r, .ColIndex("IncentiveRetail")) * Val(txtIncentiveCostPercent) / 100, 2)
                    End If
                    
                    .TextMatrix(r, .ColIndex("Cost")) = Round(.ValueMatrix(r, .ColIndex("ConstructionCost")), 2) + Round(.ValueMatrix(r, .ColIndex("LandCost")), 2) + Round(.ValueMatrix(r, .ColIndex("IncentiveCost")), 2)
                    

                    If CalcSelling Then
                        If TaxIncl Then
                            If (1 - (.ValueMatrix(r, .ColIndex("Margin")) / 100)) <> 0 Then
                                .TextMatrix(r, .ColIndex("Pretax")) = Round(.ValueMatrix(r, .ColIndex("Cost")) * (1 + .ValueMatrix(r, .ColIndex("Markup")) / 100), 2)
                                .TextMatrix(r, .ColIndex("Tax")) = CalcTax(.ValueMatrix(r, .ColIndex("AssemblyType")), .ValueMatrix(r, .ColIndex("Pretax")))
                                .TextMatrix(r, .ColIndex("Total")) = RoundToPrice(Round(.ValueMatrix(r, .ColIndex("Pretax")) + .ValueMatrix(r, .ColIndex("Tax")), 2), Multiple)
                                .TextMatrix(r, .ColIndex("Pretax")) = CalcPreTax(AssemblyType, Round(.ValueMatrix(r, .ColIndex("Total")), 2), False)
                                .TextMatrix(r, .ColIndex("Tax")) = Round(.ValueMatrix(r, .ColIndex("Total")) - .ValueMatrix(r, .ColIndex("Pretax")), 2)
                                
                                .TextMatrix(r, .ColIndex("COPretax")) = Round(.ValueMatrix(r, .ColIndex("Pretax")) * (1 + .ValueMatrix(r, .ColIndex("COMarkup")) / 100), 2)
                                .TextMatrix(r, .ColIndex("COPretax")) = Max(.ValueMatrix(r, .ColIndex("Pretax")), .ValueMatrix(r, .ColIndex("COPretax")))
                                
                                .TextMatrix(r, .ColIndex("COTax")) = CalcTax(.ValueMatrix(r, .ColIndex("AssemblyType")), .ValueMatrix(r, .ColIndex("COPretax")))
                                .TextMatrix(r, .ColIndex("COTotal")) = RoundToPrice(Round(.ValueMatrix(r, .ColIndex("COPretax")) + .ValueMatrix(r, .ColIndex("COTax")), 2), Multiple)
                                .TextMatrix(r, .ColIndex("COPretax")) = CalcPreTax(AssemblyType, .ValueMatrix(r, .ColIndex("COTotal")))
                                .TextMatrix(r, .ColIndex("COTax")) = Round(.ValueMatrix(r, .ColIndex("COTotal")) - .ValueMatrix(r, .ColIndex("COPretax")), 2)
                                If .ValueMatrix(r, .ColIndex("AssemblyType")) = 4 Then
                                    For i = 2 To 10
                                        If .ValueMatrix(r, .ColIndex("Cost")) <> 0 Then .TextMatrix(r, .ColIndex("Pretax" & s)) = Round(.ValueMatrix(r, .ColIndex("Cost")) * (1 + .ValueMatrix(r, .ColIndex("Markup" & s)) / 100), 2)
                                        .TextMatrix(r, .ColIndex("Tax" & s)) = CalcTax(.ValueMatrix(r, .ColIndex("AssemblyType")), .ValueMatrix(r, .ColIndex("Pretax" & s)))
                                        .TextMatrix(r, .ColIndex("Total" & s)) = RoundToPrice(Round(.ValueMatrix(r, .ColIndex("Pretax" & s)) + .ValueMatrix(r, .ColIndex("Tax" & s)), 2), Multiple)
                                        .TextMatrix(r, .ColIndex("Pretax" & s)) = CalcPreTax(AssemblyType, Round(.ValueMatrix(r, .ColIndex("Total" & s)), 2), False)
                                        .TextMatrix(r, .ColIndex("Tax" & s)) = Round(.ValueMatrix(r, .ColIndex("Total" & s)) - .ValueMatrix(r, .ColIndex("Pretax" & s)), 2)
                                    Next
                                End If
                            End If
                        Else
                            If .ValueMatrix(r, .ColIndex("Cost")) <> 0 Then .TextMatrix(r, .ColIndex("Pretax")) = RoundToPrice(.ValueMatrix(r, .ColIndex("Cost")) * (1 + .ValueMatrix(r, .ColIndex("Markup")) / 100), Multiple)
                            .TextMatrix(r, .ColIndex("Tax")) = CalcTax(.ValueMatrix(r, .ColIndex("AssemblyType")), .ValueMatrix(r, .ColIndex("Pretax")))
                            .TextMatrix(r, .ColIndex("Total")) = Round(.ValueMatrix(r, .ColIndex("Pretax")) + .ValueMatrix(r, .ColIndex("Tax")), 2)
                           
                            .TextMatrix(r, .ColIndex("COPretax")) = RoundToPrice(.ValueMatrix(r, .ColIndex("Pretax")) * (1 + .ValueMatrix(r, .ColIndex("COMarkup")) / 100), Multiple)
                            .TextMatrix(r, .ColIndex("COPretax")) = Max(.ValueMatrix(r, .ColIndex("Pretax")), .ValueMatrix(r, .ColIndex("COPretax")))

                            .TextMatrix(r, .ColIndex("COTax")) = CalcTax(.ValueMatrix(r, .ColIndex("AssemblyType")), .ValueMatrix(r, .ColIndex("COPretax")))
                            .TextMatrix(r, .ColIndex("COTotal")) = Round(.ValueMatrix(r, .ColIndex("COPretax")) + .ValueMatrix(r, .ColIndex("COTax")), 2)
                            If .ValueMatrix(r, .ColIndex("AssemblyType")) = 4 Then
                                For i = 2 To 10
                                    .TextMatrix(r, .ColIndex("Tax" & s)) = CalcTax(.ValueMatrix(r, .ColIndex("AssemblyType")), .ValueMatrix(r, .ColIndex("Pretax" & s)))
                                    .TextMatrix(r, .ColIndex("Total" & s)) = Round(.ValueMatrix(r, .ColIndex("Pretax" & s)) + .ValueMatrix(r, .ColIndex("Tax" & s)), 2)
                                Next
                            End If
                        End If
                        
                    End If
                    Call CalcMarkup(r, True, True, True, True, True)
                    If .ValueMatrix(r, .ColIndex("AssemblyType")) = 4 Then
                        For i = 2 To 10 ' used to be 2 to 9
                            Call CalcMarkup(r, , , , , , i, i)
                        Next
                    End If
                Case "Profit"
                    If TaxIncl Then
                        .TextMatrix(r, .ColIndex("Pretax")) = Round(.ValueMatrix(r, .ColIndex("Cost")) + .ValueMatrix(r, .ColIndex("Profit")), 2)
                        .TextMatrix(r, .ColIndex("Tax")) = CalcTax(.ValueMatrix(r, .ColIndex("AssemblyType")), .ValueMatrix(r, .ColIndex("Pretax")))
                        .TextMatrix(r, .ColIndex("Total")) = RoundToPrice(Round(.ValueMatrix(r, .ColIndex("Pretax")) + .ValueMatrix(r, .ColIndex("Tax")), 2), Multiple)
                        .TextMatrix(r, .ColIndex("Pretax")) = CalcPreTax(AssemblyType, Round(.ValueMatrix(r, .ColIndex("Total")), 2), False)
                        .TextMatrix(r, .ColIndex("Tax")) = Round(.ValueMatrix(r, .ColIndex("Total")) - .ValueMatrix(r, .ColIndex("Pretax")), 2)
                    Else
                        .TextMatrix(r, .ColIndex("Pretax")) = RoundToPrice(.ValueMatrix(r, .ColIndex("Cost")) + .ValueMatrix(r, .ColIndex("Profit")), Multiple)
                        .TextMatrix(r, .ColIndex("Tax")) = CalcTax(.ValueMatrix(r, .ColIndex("AssemblyType")), .ValueMatrix(r, .ColIndex("Pretax")))
                        .TextMatrix(r, .ColIndex("Total")) = Round(.ValueMatrix(r, .ColIndex("Pretax")) + .ValueMatrix(r, .ColIndex("Tax")), 2)
                    End If
                    .TextMatrix(r, .ColIndex("COPretax")) = Round(.ValueMatrix(r, .ColIndex("Pretax")) * (1 + .ValueMatrix(r, .ColIndex("COMarkup")) / 100), 2)
                    .TextMatrix(r, .ColIndex("COPretax")) = Max(.ValueMatrix(r, .ColIndex("Pretax")), .ValueMatrix(r, .ColIndex("COPretax")))
                    .TextMatrix(r, .ColIndex("COTax")) = CalcTax(.ValueMatrix(r, .ColIndex("AssemblyType")), .ValueMatrix(r, .ColIndex("COPretax")))
                    .TextMatrix(r, .ColIndex("COTotal")) = Round(.ValueMatrix(r, .ColIndex("COPretax")) + .ValueMatrix(r, .ColIndex("COTax")), 2)
                    Call CalcMarkup(r, , True, True, True, True)
                    If .ValueMatrix(r, .ColIndex("AssemblyType")) = 4 Then
                        For i = 2 To 10
                            Call CalcMarkup(r, , , , , , i, i)
                        Next
                    End If
                    
                Case "Markup", "Markup2", "Markup3", "Markup4", "Markup5", "Markup6", "Markup7", "Markup8", "Markup9", "Markup10"
                    s = Mid(EffectiveColumn, 7)
                    If TaxIncl Then
                        .TextMatrix(r, .ColIndex("Pretax" & s)) = Round(.ValueMatrix(r, .ColIndex("Cost")) * (1 + .ValueMatrix(r, .ColIndex("Markup")) / 100), 2)
                        .TextMatrix(r, .ColIndex("Tax" & s)) = CalcTax(.ValueMatrix(r, .ColIndex("AssemblyType")), .ValueMatrix(r, .ColIndex("Pretax" & s)))
                        .TextMatrix(r, .ColIndex("Total" & s)) = RoundToPrice(Round(.ValueMatrix(r, .ColIndex("Pretax" & s)) + .ValueMatrix(r, .ColIndex("Tax" & s)), 2), Multiple)
                        .TextMatrix(r, .ColIndex("Pretax" & s)) = CalcPreTax(AssemblyType, Round(.ValueMatrix(r, .ColIndex("Total" & s)), 2), False)
                        .TextMatrix(r, .ColIndex("Tax" & s)) = Round(.ValueMatrix(r, .ColIndex("Total" & s)) - .ValueMatrix(r, .ColIndex("Pretax" & s)), 2)
                    Else
                        .TextMatrix(r, .ColIndex("Pretax" & s)) = RoundToPrice(Round(.ValueMatrix(r, .ColIndex("Cost")) * (1 + .ValueMatrix(r, .ColIndex("Markup" & s)) / 100), 2), Multiple)
                        .TextMatrix(r, .ColIndex("Tax" & s)) = CalcTax(.ValueMatrix(r, .ColIndex("AssemblyType")), .ValueMatrix(r, .ColIndex("Pretax" & s)))
                        .TextMatrix(r, .ColIndex("Total" & s)) = Round(.ValueMatrix(r, .ColIndex("Pretax" & s)) + .ValueMatrix(r, .ColIndex("Tax" & s)), 2)
                    End If
                    .TextMatrix(r, .ColIndex("COPretax")) = Round(.ValueMatrix(r, .ColIndex("Pretax")) * (1 + .ValueMatrix(r, .ColIndex("COMarkup")) / 100), 2)
                    .TextMatrix(r, .ColIndex("COPretax")) = Max(.ValueMatrix(r, .ColIndex("Pretax")), .ValueMatrix(r, .ColIndex("COPretax")))
                    .TextMatrix(r, .ColIndex("COTax")) = CalcTax(.ValueMatrix(r, .ColIndex("AssemblyType")), .ValueMatrix(r, .ColIndex("COPretax")))
                    .TextMatrix(r, .ColIndex("COTotal")) = Round(.ValueMatrix(r, .ColIndex("COPretax")) + .ValueMatrix(r, .ColIndex("COTax")), 2)
                    If s = "" Then
                        Call CalcMarkup(r, True, , True)
                    Else
                        Call CalcMarkup(r, True, , , , , , Val(s))
                    End If
                    
                    
                Case "IncludeTax"
                    If TaxIncl Then
                        If (1 - (.ValueMatrix(r, .ColIndex("Margin")) / 100)) <> 0 Then
                            If .ValueMatrix(r, .ColIndex("Cost")) <> 0 Then
                                .TextMatrix(r, .ColIndex("Pretax")) = Round(.ValueMatrix(r, .ColIndex("Cost")) * (1 + .ValueMatrix(r, .ColIndex("Markup")) / 100), 2)
                            End If
                            .TextMatrix(r, .ColIndex("Tax")) = CalcTax(.ValueMatrix(r, .ColIndex("AssemblyType")), .ValueMatrix(r, .ColIndex("Pretax")))
                            .TextMatrix(r, .ColIndex("Total")) = RoundToPrice(Round(.ValueMatrix(r, .ColIndex("Pretax")) + .ValueMatrix(r, .ColIndex("Tax")), 2), Multiple)
                            .TextMatrix(r, .ColIndex("Pretax")) = CalcPreTax(AssemblyType, Round(.ValueMatrix(r, .ColIndex("Total")), 2), False)
                            .TextMatrix(r, .ColIndex("Tax")) = Round(.ValueMatrix(r, .ColIndex("Total")) - .ValueMatrix(r, .ColIndex("Pretax")), 2)
                            
                            .TextMatrix(r, .ColIndex("COPretax")) = Round(.ValueMatrix(r, .ColIndex("Pretax")) * (1 + .ValueMatrix(r, .ColIndex("COMarkup")) / 100), 2)
                            .TextMatrix(r, .ColIndex("COPretax")) = Max(.ValueMatrix(r, .ColIndex("Pretax")), .ValueMatrix(r, .ColIndex("COPretax")))
                            
                            .TextMatrix(r, .ColIndex("COTax")) = CalcTax(.ValueMatrix(r, .ColIndex("AssemblyType")), .ValueMatrix(r, .ColIndex("COPretax")))
                            .TextMatrix(r, .ColIndex("COTotal")) = RoundToPrice(Round(.ValueMatrix(r, .ColIndex("COPretax")) + .ValueMatrix(r, .ColIndex("COTax")), 2), Multiple)
                            .TextMatrix(r, .ColIndex("COPretax")) = CalcPreTax(AssemblyType, .ValueMatrix(r, .ColIndex("COTotal")))
                            .TextMatrix(r, .ColIndex("COTax")) = Round(.ValueMatrix(r, .ColIndex("COTotal")) - .ValueMatrix(r, .ColIndex("COPretax")), 2)
                            If .ValueMatrix(r, .ColIndex("AssemblyType")) = 4 Then
                                For i = 2 To 10
                                    If .ValueMatrix(r, .ColIndex("Cost")) <> 0 Then
                                        .TextMatrix(r, .ColIndex("Pretax" & s)) = Round(.ValueMatrix(r, .ColIndex("Cost")) * (1 + .ValueMatrix(r, .ColIndex("Markup" & s)) / 100), 2)
                                    End If
                                    .TextMatrix(r, .ColIndex("Tax" & s)) = CalcTax(.ValueMatrix(r, .ColIndex("AssemblyType")), .ValueMatrix(r, .ColIndex("Pretax" & s)))
                                    .TextMatrix(r, .ColIndex("Total" & s)) = RoundToPrice(Round(.ValueMatrix(r, .ColIndex("Pretax" & s)) + .ValueMatrix(r, .ColIndex("Tax" & s)), 2), Multiple)
                                    .TextMatrix(r, .ColIndex("Pretax" & s)) = CalcPreTax(AssemblyType, Round(.ValueMatrix(r, .ColIndex("Total" & s)), 2), False)
                                    .TextMatrix(r, .ColIndex("Tax" & s)) = Round(.ValueMatrix(r, .ColIndex("Total" & s)) - .ValueMatrix(r, .ColIndex("Pretax" & s)), 2)
                                Next
                            End If
                        End If
                    Else
                        If .ValueMatrix(r, .ColIndex("Cost")) <> 0 Then
                            .TextMatrix(r, .ColIndex("Pretax")) = RoundToPrice(Round(.ValueMatrix(r, .ColIndex("Cost")) * (1 + .ValueMatrix(r, .ColIndex("Markup")) / 100), 2), Multiple)
                        End If
                        .TextMatrix(r, .ColIndex("Tax")) = CalcTax(.ValueMatrix(r, .ColIndex("AssemblyType")), .ValueMatrix(r, .ColIndex("Pretax")))
                        .TextMatrix(r, .ColIndex("Total")) = Round(.ValueMatrix(r, .ColIndex("Pretax")) + .ValueMatrix(r, .ColIndex("Tax")), 2)
                        
                        .TextMatrix(r, .ColIndex("COPretax")) = RoundToPrice(Round(.ValueMatrix(r, .ColIndex("Pretax")) * (1 + .ValueMatrix(r, .ColIndex("COMarkup")) / 100), 2), Multiple)
                        .TextMatrix(r, .ColIndex("COPretax")) = Max(.ValueMatrix(r, .ColIndex("Pretax")), .ValueMatrix(r, .ColIndex("COPretax")))
                        .TextMatrix(r, .ColIndex("COTax")) = CalcTax(.ValueMatrix(r, .ColIndex("AssemblyType")), .ValueMatrix(r, .ColIndex("COPretax")))
                        .TextMatrix(r, .ColIndex("COTotal")) = Round(.ValueMatrix(r, .ColIndex("COPretax")) + .ValueMatrix(r, .ColIndex("COTax")), 2)
                        If .ValueMatrix(r, .ColIndex("AssemblyType")) = 4 Then
                            For i = 2 To 10
                                If .ValueMatrix(r, .ColIndex("Cost")) <> 0 Then
                                    .TextMatrix(r, .ColIndex("Pretax" & s)) = RoundToPrice(Round(.ValueMatrix(r, .ColIndex("Cost")) * (1 + .ValueMatrix(r, .ColIndex("Markup")) / 100), 2), Multiple)
                                End If
                                .TextMatrix(r, .ColIndex("Tax" & s)) = CalcTax(.ValueMatrix(r, .ColIndex("AssemblyType")), .ValueMatrix(r, .ColIndex("Pretax" & s)))
                                .TextMatrix(r, .ColIndex("Total" & s)) = Round(.ValueMatrix(r, .ColIndex("Pretax" & s)) + .ValueMatrix(r, .ColIndex("Tax" & s)), 2)
                            Next
                        End If
                    End If
                    Call CalcMarkup(r, True, True, True, True, True)
                    If .ValueMatrix(r, .ColIndex("AssemblyType")) = 4 Then
                        For i = 2 To 10 'Used to be 2 to 9
                            Call CalcMarkup(r, , , , , , i, i)
                        Next
                    End If
                    
                    
                    
                Case "Roundto"
                     s = ""
                     If TaxIncl Then
                         If .ValueMatrix(r, .ColIndex("Cost")) <> 0 Then .TextMatrix(r, .ColIndex("Pretax")) = Round(.ValueMatrix(r, .ColIndex("Cost")) * (1 + .ValueMatrix(r, .ColIndex("Markup")) / 100), 2)
                         
                         .TextMatrix(r, .ColIndex("Tax")) = CalcTax(.ValueMatrix(r, .ColIndex("AssemblyType")), .ValueMatrix(r, .ColIndex("Pretax")))
                         .TextMatrix(r, .ColIndex("Total")) = RoundToPrice(Round(.ValueMatrix(r, .ColIndex("Pretax")) + .ValueMatrix(r, .ColIndex("Tax")), 2), Multiple)
                         .TextMatrix(r, .ColIndex("Pretax")) = CalcPreTax(AssemblyType, Round(.ValueMatrix(r, .ColIndex("Total")), 2), False)
                         .TextMatrix(r, .ColIndex("Tax")) = Round(.ValueMatrix(r, .ColIndex("Total")) - .ValueMatrix(r, .ColIndex("Pretax")), 2)
                         
                         .TextMatrix(r, .ColIndex("COPretax")) = Round(.ValueMatrix(r, .ColIndex("Pretax")) * (1 + .ValueMatrix(r, .ColIndex("COMarkup")) / 100), 2)
                         .TextMatrix(r, .ColIndex("COPretax")) = Max(.ValueMatrix(r, .ColIndex("Pretax")), .ValueMatrix(r, .ColIndex("COPretax")))
                         
                         .TextMatrix(r, .ColIndex("COTax")) = CalcTax(.ValueMatrix(r, .ColIndex("AssemblyType")), .ValueMatrix(r, .ColIndex("COPretax")))
                         .TextMatrix(r, .ColIndex("COTotal")) = RoundToPrice(Round(.ValueMatrix(r, .ColIndex("COPretax")) + .ValueMatrix(r, .ColIndex("COTax")), 2), Multiple)
                         .TextMatrix(r, .ColIndex("COPretax")) = CalcPreTax(AssemblyType, .ValueMatrix(r, .ColIndex("COTotal")))
                         .TextMatrix(r, .ColIndex("COTax")) = Round(.ValueMatrix(r, .ColIndex("COTotal")) - .ValueMatrix(r, .ColIndex("COPretax")), 2)
                         If .ValueMatrix(r, .ColIndex("AssemblyType")) = 4 Then
                            For i = 2 To 10
                                If .ValueMatrix(r, .ColIndex("Cost")) <> 0 Then .TextMatrix(r, .ColIndex("Pretax" & s)) = Round(.ValueMatrix(r, .ColIndex("Cost")) * (1 + .ValueMatrix(r, .ColIndex("Markup" & s)) / 100), 2)
                                .TextMatrix(r, .ColIndex("Tax" & s)) = CalcTax(.ValueMatrix(r, .ColIndex("AssemblyType")), .ValueMatrix(r, .ColIndex("Pretax" & s)))
                                .TextMatrix(r, .ColIndex("Total" & s)) = RoundToPrice(.ValueMatrix(r, .ColIndex("Pretax" & s)) + .ValueMatrix(r, .ColIndex("Tax" & s)), Multiple)
                                .TextMatrix(r, .ColIndex("Pretax" & s)) = CalcPreTax(AssemblyType, Round(.ValueMatrix(r, .ColIndex("Total" & s)), 2), False)
                                .TextMatrix(r, .ColIndex("Tax" & s)) = Round(.ValueMatrix(r, .ColIndex("Total" & s)) - .ValueMatrix(r, .ColIndex("Pretax" & s)), 2)
                            Next
                         End If
                     Else
                         If .ValueMatrix(r, .ColIndex("Cost")) <> 0 Then
                            .TextMatrix(r, .ColIndex("Pretax")) = RoundToPrice(.ValueMatrix(r, .ColIndex("Cost")) * (1 + .ValueMatrix(r, .ColIndex("Markup")) / 100), Multiple)
                         Else
                            .TextMatrix(r, .ColIndex("Pretax")) = RoundToPrice(.ValueMatrix(r, .ColIndex("Pretax")), Multiple)
                         End If
                         .TextMatrix(r, .ColIndex("Tax")) = CalcTax(.ValueMatrix(r, .ColIndex("AssemblyType")), .ValueMatrix(r, .ColIndex("Pretax")))
                         .TextMatrix(r, .ColIndex("Total")) = Round(.ValueMatrix(r, .ColIndex("Pretax")) + .ValueMatrix(r, .ColIndex("Tax")), 2)
                     
                         .TextMatrix(r, .ColIndex("COPretax")) = RoundToPrice(.ValueMatrix(r, .ColIndex("Pretax")) * (1 + .ValueMatrix(r, .ColIndex("COMarkup")) / 100), Multiple)
                         .TextMatrix(r, .ColIndex("COPretax")) = Max(.ValueMatrix(r, .ColIndex("Pretax")), .ValueMatrix(r, .ColIndex("COPretax")))

                         .TextMatrix(r, .ColIndex("COTax")) = CalcTax(.ValueMatrix(r, .ColIndex("AssemblyType")), .ValueMatrix(r, .ColIndex("COPretax")))
                         .TextMatrix(r, .ColIndex("COTotal")) = Round(.ValueMatrix(r, .ColIndex("COPretax")) + .ValueMatrix(r, .ColIndex("COTax")), 2)
                         If .ValueMatrix(r, .ColIndex("AssemblyType")) = 4 Then
                            For i = 2 To 10
                                If .ValueMatrix(r, .ColIndex("Cost")) <> 0 Then .TextMatrix(r, .ColIndex("Pretax" & s)) = RoundToPrice(.ValueMatrix(r, .ColIndex("Cost")) * (1 + .ValueMatrix(r, .ColIndex("Markup")) / 100), Multiple)
                                .TextMatrix(r, .ColIndex("Tax" & s)) = CalcTax(.ValueMatrix(r, .ColIndex("AssemblyType")), .ValueMatrix(r, .ColIndex("Pretax" & s)))
                                .TextMatrix(r, .ColIndex("Total" & s)) = Round(.ValueMatrix(r, .ColIndex("Pretax" & s)) + .ValueMatrix(r, .ColIndex("Tax" & s)), 2)
                            Next
                         End If
                     End If
                     If s = "" Then
                         Call CalcMarkup(r, True, False, False)
                     Else
                         Call CalcMarkup(r, True, False, False, False, False, Val(s), Val(s))
                     End If
            
                    
                
                    
                Case "Margin", "Margin2", "Margin3", "Margin4", "Margin5", "Margin6", "Margin7", "Margin8", "Margin9", "Margin10"
                    s = Mid(EffectiveColumn, 7)
                    If TaxIncl Then
                        .TextMatrix(r, .ColIndex("Pretax" & s)) = Round((100 * (.ValueMatrix(r, .ColIndex("Cost")))) / (100 - .ValueMatrix(r, .ColIndex("Margin" & s))), 2)
                        .TextMatrix(r, .ColIndex("Tax" & s)) = CalcTax(.ValueMatrix(r, .ColIndex("AssemblyType")), .ValueMatrix(r, .ColIndex("Pretax" & s)))
                        .TextMatrix(r, .ColIndex("Total" & s)) = RoundToPrice(.ValueMatrix(r, .ColIndex("Pretax" & s)) + .ValueMatrix(r, .ColIndex("Tax" & s)), Multiple)
                        .TextMatrix(r, .ColIndex("Pretax" & s)) = CalcPreTax(AssemblyType, Round(.ValueMatrix(r, .ColIndex("Total" & s)), 2))
                        .TextMatrix(r, .ColIndex("Tax" & s)) = Round(.ValueMatrix(r, .ColIndex("Total" & s)) - .ValueMatrix(r, .ColIndex("Pretax" & s)), 2)
                    Else
                        .TextMatrix(r, .ColIndex("Pretax" & s)) = RoundToPrice((100 * (.ValueMatrix(r, .ColIndex("Cost")))) / (100 - .ValueMatrix(r, .ColIndex("Margin" & s))), Multiple)
                        .TextMatrix(r, .ColIndex("Tax" & s)) = CalcTax(.ValueMatrix(r, .ColIndex("AssemblyType")), .ValueMatrix(r, .ColIndex("Pretax" & s)))
                        .TextMatrix(r, .ColIndex("Total" & s)) = Round(.ValueMatrix(r, .ColIndex("Pretax" & s)) + .ValueMatrix(r, .ColIndex("Tax" & s)), 2)
                    End If
                    .TextMatrix(r, .ColIndex("COPretax")) = Round(.ValueMatrix(r, .ColIndex("Pretax")) * (1 + .ValueMatrix(r, .ColIndex("COMarkup")) / 100), 2)
                    .TextMatrix(r, .ColIndex("COPretax")) = Max(.ValueMatrix(r, .ColIndex("Pretax")), .ValueMatrix(r, .ColIndex("COPretax")))
                    .TextMatrix(r, .ColIndex("COTax")) = CalcTax(.ValueMatrix(r, .ColIndex("AssemblyType")), .ValueMatrix(r, .ColIndex("COPretax")))
                    .TextMatrix(r, .ColIndex("COTotal")) = Round(.ValueMatrix(r, .ColIndex("COPretax")) + .ValueMatrix(r, .ColIndex("COTax")), 2)
                    If s = "" Then
                        Call CalcMarkup(r, True, True)
                    Else
                        Call CalcMarkup(r, True, , , , , Val(s))
                    End If
                
                                    
                Case "Pretax", "Pretax2", "Pretax3", "Pretax4", "Pretax5", "Pretax6", "Pretax7", "Pretax8", "Pretax9", "Pretax10"
                    s = Mid(EffectiveColumn, 7)
                    .TextMatrix(r, .ColIndex("Tax" & s)) = CalcTax(.ValueMatrix(r, .ColIndex("AssemblyType")), .ValueMatrix(r, .ColIndex("Pretax" & s)))
                    .TextMatrix(r, .ColIndex("Total" & s)) = Round(.ValueMatrix(r, .ColIndex("Pretax" & s)) + .ValueMatrix(r, .ColIndex("Tax" & s)), 2)
                    .TextMatrix(r, .ColIndex("COPretax")) = Round(.ValueMatrix(r, .ColIndex("Pretax")) * (1 + .ValueMatrix(r, .ColIndex("COMarkup")) / 100), 2)
                    .TextMatrix(r, .ColIndex("COPretax")) = Max(.ValueMatrix(r, .ColIndex("Pretax")), .ValueMatrix(r, .ColIndex("COPretax")))
                    .TextMatrix(r, .ColIndex("COTax")) = CalcTax(.ValueMatrix(r, .ColIndex("AssemblyType")), .ValueMatrix(r, .ColIndex("COPretax")))
                    .TextMatrix(r, .ColIndex("COTotal")) = Round(.ValueMatrix(r, .ColIndex("COPretax")) + .ValueMatrix(r, .ColIndex("COTax")), 2)
                    If s = "" Then
                        Call CalcMarkup(r, True, True, True)
                    Else
                        Call CalcMarkup(r, True, , , , , Val(s), Val(s))
                    End If
                                    
                
                Case "Total", "Total2", "Total3", "Total4", "Total5", "Total6", "Total7", "Total8", "Total9", "Total10"
                    s = Mid(EffectiveColumn, 7)
                    .TextMatrix(r, .ColIndex("Pretax" & s)) = CalcPreTax(AssemblyType, Round(.ValueMatrix(r, .ColIndex("Total" & s)), 2))
                    .TextMatrix(r, .ColIndex("Tax" & s)) = Round(.ValueMatrix(r, .ColIndex("Total" & s)) - .ValueMatrix(r, .ColIndex("Pretax" & s)), 2)
                    .TextMatrix(r, .ColIndex("COPretax")) = Round(.ValueMatrix(r, .ColIndex("Pretax")) * (1 + .ValueMatrix(r, .ColIndex("COMarkup")) / 100), 2)
                    .TextMatrix(r, .ColIndex("COPretax")) = Max(.ValueMatrix(r, .ColIndex("Pretax")), .ValueMatrix(r, .ColIndex("COPretax")))
                    .TextMatrix(r, .ColIndex("COTax")) = CalcTax(.ValueMatrix(r, .ColIndex("AssemblyType")), .ValueMatrix(r, .ColIndex("COPretax")))
                    .TextMatrix(r, .ColIndex("COTotal")) = Round(.ValueMatrix(r, .ColIndex("COPretax")) + .ValueMatrix(r, .ColIndex("COTax")), 2)
                  
                    
                    If s = "" Then
                        Call CalcMarkup(r, True, True, True)
                    Else
                        Call CalcMarkup(r, True, , , , , Val(s), Val(s))
                    End If
                
                
                Case "Category", "CategoryDesc", "COMarkup"
                    If TaxIncl Then
                        .TextMatrix(r, .ColIndex("COPretax")) = Round(.ValueMatrix(r, .ColIndex("Pretax")) * (1 + .ValueMatrix(r, .ColIndex("COMarkup")) / 100), 2)
                        .TextMatrix(r, .ColIndex("COPretax")) = Max(.ValueMatrix(r, .ColIndex("Pretax")), .ValueMatrix(r, .ColIndex("COPretax")))
                        .TextMatrix(r, .ColIndex("COTax")) = CalcTax(.ValueMatrix(r, .ColIndex("AssemblyType")), .ValueMatrix(r, .ColIndex("COPretax")))
                        .TextMatrix(r, .ColIndex("COTotal")) = RoundToPrice(Round(.ValueMatrix(r, .ColIndex("COPretax")) + .ValueMatrix(r, .ColIndex("COTax")), 2), Multiple)
                        .TextMatrix(r, .ColIndex("COPretax")) = CalcPreTax(AssemblyType, Round(.ValueMatrix(r, .ColIndex("COTotal")), 2), False)
                        .TextMatrix(r, .ColIndex("COTax")) = Round(.ValueMatrix(r, .ColIndex("COTotal")) - .ValueMatrix(r, .ColIndex("COPretax")), 2)
                    Else
                        .TextMatrix(r, .ColIndex("COPretax")) = Round(.ValueMatrix(r, .ColIndex("Pretax")) * (1 + .ValueMatrix(r, .ColIndex("COMarkup")) / 100), 2)
                        .TextMatrix(r, .ColIndex("COPretax")) = Max(.ValueMatrix(r, .ColIndex("Pretax")), .ValueMatrix(r, .ColIndex("COPretax")))
                        .TextMatrix(r, .ColIndex("COTax")) = CalcTax(.ValueMatrix(r, .ColIndex("AssemblyType")), .ValueMatrix(r, .ColIndex("COPretax")))
                        .TextMatrix(r, .ColIndex("COTotal")) = Round(.ValueMatrix(r, .ColIndex("COPretax")) + .ValueMatrix(r, .ColIndex("COTax")), 2)
                    End If
                    Call CalcMarkup(r, , , , , True)
                                
                Case "COMargin"
                    If TaxIncl Then
                        .TextMatrix(r, .ColIndex("COPretax")) = Round((100 * (.ValueMatrix(r, .ColIndex("Pretax")))) / (100 - .ValueMatrix(r, .ColIndex("COMargin"))), 2)
                        .TextMatrix(r, .ColIndex("COPretax")) = Max(.ValueMatrix(r, .ColIndex("Pretax")), .ValueMatrix(r, .ColIndex("COPretax")))
                        .TextMatrix(r, .ColIndex("COTax")) = CalcTax(.ValueMatrix(r, .ColIndex("AssemblyType")), .ValueMatrix(r, .ColIndex("COPretax")))
                        .TextMatrix(r, .ColIndex("COTotal")) = RoundToPrice(Round(.ValueMatrix(r, .ColIndex("COPretax")) + .ValueMatrix(r, .ColIndex("COTax")), 2), Multiple)
                        .TextMatrix(r, .ColIndex("COPretax")) = CalcPreTax(AssemblyType, Round(.ValueMatrix(r, .ColIndex("COTotal")), 2))
                        .TextMatrix(r, .ColIndex("COTax")) = Round(.ValueMatrix(r, .ColIndex("COTotal")) - .ValueMatrix(r, .ColIndex("COPretax")), 2)
                    Else
                        .TextMatrix(r, .ColIndex("COPretax")) = RoundToPrice(Round((100 * (.ValueMatrix(r, .ColIndex("Pretax")))) / (100 - .ValueMatrix(r, .ColIndex("COMargin"))), 2), Multiple)
                        .TextMatrix(r, .ColIndex("COPretax")) = Max(.ValueMatrix(r, .ColIndex("Pretax")), .ValueMatrix(r, .ColIndex("COPretax")))
                        .TextMatrix(r, .ColIndex("COTax")) = CalcTax(.ValueMatrix(r, .ColIndex("AssemblyType")), .ValueMatrix(r, .ColIndex("COPretax")))
                        .TextMatrix(r, .ColIndex("COTotal")) = Round(.ValueMatrix(r, .ColIndex("COPretax")) + .ValueMatrix(r, .ColIndex("Tax")), 2)
                    End If
                    Call CalcMarkup(r, , , , True)
                
                Case "COPretax"
                    .TextMatrix(r, .ColIndex("COTax")) = CalcTax(.ValueMatrix(r, .ColIndex("AssemblyType")), .ValueMatrix(r, .ColIndex("COPretax")))
                    .TextMatrix(r, .ColIndex("COTotal")) = Round(.ValueMatrix(r, .ColIndex("COPretax")) + .ValueMatrix(r, .ColIndex("COTax")), 2)
                    Call CalcMarkup(r, , , , True, True)
                
                Case "COTotal"
                    .TextMatrix(r, .ColIndex("COPretax")) = CalcPreTax(AssemblyType, .ValueMatrix(r, .ColIndex("Total")))
                    .TextMatrix(r, .ColIndex("COTax")) = Round(.ValueMatrix(r, .ColIndex("Total")) - .ValueMatrix(r, .ColIndex("COPretax")), 2)
                    Call CalcMarkup(r, , , , True, True)
                
                
                Case "Series"
'                    Call RefreshPublished(False, False)
                    
            End Select
        End If
        
        If .ValueMatrix(r, .ColIndex("FloorArea")) = 0 Then
            .TextMatrix(r, .ColIndex("CostPerFt")) = ""
            .TextMatrix(r, .ColIndex("PricePerFt")) = ""
        Else
            .TextMatrix(r, .ColIndex("CostPerFt")) = .ValueMatrix(r, .ColIndex("ConstructionCost")) / .ValueMatrix(r, .ColIndex("FloorArea"))
            .TextMatrix(r, .ColIndex("PricePerFt")) = .ValueMatrix(r, .ColIndex("Pretax")) / .ValueMatrix(r, .ColIndex("FloorArea"))
        End If
        
    Next
        
        
        
        
        
    End With
    Call gData_RowColChange
    On Error Resume Next
    gData.Row = origR
    
Exit Sub
eh:  Call errHandler(SRCFILE & "CalcData()")
End Sub


Private Sub CalcMarkup(r As Long, _
                       Optional CalcProfit As Boolean, _
                       Optional CalcMarkup As Boolean, Optional CalcMargin As Boolean, _
                       Optional CalcCOMarkup As Boolean, Optional CalcCOMargin As Boolean, _
                       Optional CalcMarkupX As Long, Optional CalcMarginX As Long)
    With gData
    
        '--------------------------------------------------------------------------------------------------
        If CalcProfit Then
            .TextMatrix(r, .ColIndex("Profit")) = .ValueMatrix(r, .ColIndex("Pretax")) - .ValueMatrix(r, .ColIndex("Cost"))
        End If
        '--------------------------------------------------------------------------------------------------
        If CalcMarkup Then
            If .ValueMatrix(r, .ColIndex("Cost")) = 0 Then
                .TextMatrix(r, .ColIndex("Markup")) = 0
            Else
                .TextMatrix(r, .ColIndex("Markup")) = (.ValueMatrix(r, .ColIndex("Pretax")) - .ValueMatrix(r, .ColIndex("Cost"))) / .ValueMatrix(r, .ColIndex("Cost")) * 100
            End If
        End If
        '--------------------------------------------------------------------------------------------------
        If CalcMargin Then
            If .ValueMatrix(r, .ColIndex("Pretax")) = 0 Then
                .TextMatrix(r, .ColIndex("Margin")) = 0
            Else
                .TextMatrix(r, .ColIndex("Margin")) = ((.ValueMatrix(r, .ColIndex("Pretax")) - .ValueMatrix(r, .ColIndex("Cost"))) / .ValueMatrix(r, .ColIndex("Pretax"))) * 100
            End If
        End If
'        '--------------------------------------------------------------------------------------------------
'        If CalcCOMarkup Then
'            If .ValueMatrix(r, .ColIndex("Cost")) = 0 Then
'                .TextMatrix(r, .ColIndex("COMarkup")) = 0
'            Else
'                .TextMatrix(r, .ColIndex("COMarkup")) = (.ValueMatrix(r, .ColIndex("COPretax")) - .ValueMatrix(r, .ColIndex("Cost"))) / .ValueMatrix(r, .ColIndex("Cost")) * 100
'            End If
'        End If
'        '--------------------------------------------------------------------------------------------------
'        If CalcCOMargin Then
'            If .ValueMatrix(r, .ColIndex("COPretax")) = 0 Then
'                .TextMatrix(r, .ColIndex("COMargin")) = 0
'            Else
'                .TextMatrix(r, .ColIndex("COMargin")) = ((.ValueMatrix(r, .ColIndex("COPretax")) - .ValueMatrix(r, .ColIndex("Cost"))) / .ValueMatrix(r, .ColIndex("COPretax"))) * 100
'            End If
'        End If
        '--------------------------------------------------------------------------------------------------
        If CalcMarkupX <> 0 Then
            If .ValueMatrix(r, .ColIndex("Cost")) = 0 Then
                .TextMatrix(r, .ColIndex("Markup" & CalcMarkupX)) = 0
            Else
                .TextMatrix(r, .ColIndex("Markup" & CalcMarkupX)) = (.ValueMatrix(r, .ColIndex("Pretax" & CalcMarkupX)) - .ValueMatrix(r, .ColIndex("Cost"))) / .ValueMatrix(r, .ColIndex("Cost")) * 100
            End If
        End If
        '--------------------------------------------------------------------------------------------------
        If CalcMarginX <> 0 Then
            If .ValueMatrix(r, .ColIndex("Pretax" & CalcMarginX)) = 0 Then
                .TextMatrix(r, .ColIndex("Margin" & CalcMarginX)) = 0
            Else
                .TextMatrix(r, .ColIndex("Margin" & CalcMarginX)) = ((.ValueMatrix(r, .ColIndex("Pretax" & CalcMarginX)) - .ValueMatrix(r, .ColIndex("Cost"))) / .ValueMatrix(r, .ColIndex("Pretax" & CalcMarginX))) * 100
            End If
        End If
        
   End With
    
End Sub



Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo eh
    Dim i As Long
    Dim s As String
    
    Dim bHasGlobal     As Boolean
    Dim bNeedCommunity As Boolean
    Dim bNeedPhase     As Boolean
    

    Select Case Button.Key
    
        Case "AssemblyCosts"
            With gData
                If .TextMatrix(.Row, .ColIndex("Assembly")) <> "" And mWorksheetView <> "Marketing" And HFApp.Options(RecordSalesSheetCosts) Then
                    Call FAssemblyCosts.ShowForm(mWorksheet, .TextMatrix(.Row, .ColIndex("Community")), .TextMatrix(.Row, .ColIndex("CommunityPhase")), .TextMatrix(.Row, .ColIndex("Model")), .TextMatrix(.Row, .ColIndex("OptionID")), .TextMatrix(.Row, .ColIndex("Assembly")))
                End If
            End With
        
        Case "Publish"
            If vbYes = MsgBox("Publishing this worksheet will change the cost and selling" & vbCrLf & "prices in your Profit Builder database." & vbCrLf & "Are you sure this is what you want to do?", vbYesNo + vbExclamation, App.ProductName) Then
                If SaveData(False) Then
                    Call PublishPricingWorksheet(mWorksheet)
                    If HFApp.Options(LockPostedSalesWorksheets) Then ReadOnly = True
                    'dont do this. its too slow
                    'Call RefreshPublished(False, True)
                    txtSalesEffectiveDate.Text = format(Now(), "mmmm d, yyyy")
                End If
            End If
    
        Case "Add"
            Call AddAssembly
        
        Case "LookupPublished"
            Call RefreshPublished(False, True)
            
        Case "RePrice"
            CostEffectiveDate = Now
            Call RefreshCosts(True)
            Call RefreshLandCosts
            
        Case "Delete"
            Call gData_KeyDown(vbKeyDelete, vbCtrlMask)
            
        Case "Preview"
            If mDirty Then
                If vbOK = MsgBox("You must save your changes before opening the print preview window.", vbOKCancel + vbQuestion, App.ProductName) Then
                     If Not SaveData(False) Then Exit Sub
                Else
                    Exit Sub
                End If
            End If
            If HFApp.Options(RecordSalesSheetCosts) Then
                If HFApp.DivisionID = "" Then
                    s = PathAppend(HFApp.SystemFolder, "System\Reports\Estimating\SalesSheetCosts.rpt")
                Else
                    s = PathAppend(HFApp.SystemFolder, "System\Reports\Estimating\SalesSheetCosts Division version.rpt")
                End If
            Else
                If HFApp.DivisionID = "" Then
                    s = PathAppend(HFApp.SystemFolder, "System\Reports\Estimating\SalesSheetCurrentCosts.rpt")
                Else
                    s = PathAppend(HFApp.SystemFolder, "System\Reports\Estimating\SalesSheetCurrentCosts Division version.rpt")
                End If
            End If
            'Call FRptViewer.ShowReport(s, True, True, "Worksheet", mWorksheet)
            Dim c As New ZybUtil.Crystal
            Call c.LoadODBCReport(s, HFApp.LoginDSN, HFApp.LoginDBUID, HFApp.LoginDBPWD)
            On Error Resume Next
            Call c.ParameterValue("DivisionID", HFApp.DivisionID)
            Call c.ParameterValue("Worksheet", mWorksheet)
            On Error GoTo eh
            Call c.PrintPreview("Print Preview")
            
        Case "Save"
            Call SaveData(False)
            Call HFApp.LockRecord("tblSalesSheetMaster", mWorksheet)
                
        Case "SaveAs"
            s = InputBox(vbCrLf & vbCrLf & vbCrLf & "Enter a description for the new worksheet", "Save Worksheet As...")
            If s <> "" Then
                Call HFApp.UnLockRecord("tblSalesSheetMaster", mWorksheet)
                mWorksheet = 0
                txtDescription.Text = s
                mDirty = True
                Call SaveData(False)
                Call HFApp.LockRecord("tblSalesSheetMaster", mWorksheet)
                Me.Tag = mWorksheet
                Me.Caption = mWorksheetView & " Worksheet (number " & mWorksheet & ")"
                ReadOnly = False
            End If
            


    End Select
Exit Sub
eh: Call errHandler(SRCFILE & "Toolbar_ButtonClick", s)
End Sub

Private Sub AddAssemblyToGridOLD(SourceCommunity As String, community As String, CommunityPhase As String, Assembly As String, Model As String, OptionID As String, series As String)
    Dim i As Long
    Dim s As String
    Dim rs As Recordset
    Dim r As Long
    Dim AssemblyType As AssemblyTypes

Dim X As Single
Dim Y As String
X = Timer()
    s = ""
    s = s & "SELECT m.Model,m.Elevation,m.Description,dm.Description ModelDescription,m.Notes,m.Comments,m.AssemblyType,m.Qty,m.Color,m.Location,m.IncludedOption" & vbCrLf
    s = s & "      ,m.Category,c.Description CategoryDesc" & vbCrLf
    s = s & "      ,l.area community,l.Description CommunityDesc,m.Assembly,m.OptionID,m.AssemblyUOM" & vbCrLf
    s = s & "      ,m.Series,m.Style,m.JCExtra,m.FloorArea,m.Bedrooms,m.Bathrooms,m.ConstCutOff" & vbCrLf
    s = s & "      ,CASE isnull(p.assembly,'') when '' THEN m.IncentiveRetail ELSE p.IncentiveRetail END IncentiveRetail" & vbCrLf
    s = s & "      ,CASE isnull(p.assembly,'') when '' THEN m.IncludeTax ELSE p.IncludeTax END IncludeTax" & vbCrLf
    s = s & "      ,CASE isnull(p.assembly,'') when '' THEN m.Markup ELSE p.Markup END Markup" & vbCrLf
    s = s & "      ,CASE isnull(p.assembly,'') when '' THEN c.COMarkup ELSE p.COMarkup END COMarkup" & vbCrLf
    s = s & "      ,CASE isnull(p.assembly,'') when '' THEN m.Margin ELSE p.Margin END Margin" & vbCrLf
    s = s & "      ,CASE isnull(p.assembly,'') when '' THEN m.Roundto ELSE p.Roundto END Roundto" & vbCrLf
    s = s & "      ,CASE isnull(p.assembly,'') when '' THEN m.Pretax ELSE p.Pretax END Pretax" & vbCrLf
    s = s & "      ,CASE isnull(p.assembly,'') when '' THEN m.Tax ELSE p.Tax END Tax" & vbCrLf
    s = s & "      ,CASE isnull(p.assembly,'') when '' THEN m.COPretax ELSE p.COPretax END COPretax" & vbCrLf
    s = s & "      ,CASE isnull(p.assembly,'') when '' THEN m.COTax ELSE p.COTax END COTax      " & vbCrLf
    s = s & "      ,m.IncludedInSpec" & vbCrLf
    For i = 2 To 10
        s = s & "      ,m.Pretax" & i & vbCrLf
        s = s & "      ,m.Tax" & i & vbCrLf
        s = s & "      ,m.IncludedInSpec" & i & vbCrLf
    Next
    s = s & "      ,m.ColorListID,m.StyleListID,m.FinishListID,m.OtherListID,m.StyleValue,m.FinishValue,m.OtherValue,m.graphicpath,m.specdocument,m.maxwidth,m.maxlength,m.constcutoff" & vbCrLf
    s = s & ",m.DesignCenterSalesOnly,m.SelectByRoom,m.DisplayTotalOnly" & vbCrLf
    s = s & "  FROM tblDBAssemblyMaster m " & vbCrLf
    s = s & "       LEFT OUTER JOIN tblDBAssemblyPrices p ON(p.Community=" & DbQuote(Str, community) & " AND p.CommunityPhase=" & DbQuote(Str, CommunityPhase) & " AND m.Assembly=p.Assembly AND m.Model=p.Model AND m.OptionID=p.OptionID)" & vbCrLf
    s = s & "       LEFT OUTER JOIN tblCategories c ON m.Category=c.Category" & vbCrLf
    s = s & "       left outer join distinctmodels dm on(m.model=dm.model)" & vbCrLf
    s = s & "       LEFT OUTER JOIN tblLocality l ON l.area=" & DbQuote(Str, community) & vbCrLf
    s = s & " WHERE isnull(m.Community,'')=" & DbQuote(Str, SourceCommunity) & vbCrLf
    s = s & "   AND isnull(m.Assembly,'')=" & DbQuote(Str, Assembly) & vbCrLf
    s = s & "   AND isnull(m.Model,'')=" & DbQuote(Str, Model) & vbCrLf
    s = s & "   AND isnull(m.OptionID,'')=" & DbQuote(Str, OptionID) & vbCrLf
    s = s & "   AND m.DivisionID = " & HFApp.DivisionID & vbCrLf
    s = s & "union all" & vbCrLf
    s = s & "SELECT '' Model,'' Elevation,m.Description,'' ModelDescription,m.Notes,m.Notes Comments,3 AssemblyType,1 Qty,m.Color,m.Location,0 IncludedOption" & vbCrLf
    s = s & "      ,m.OptionCategory Category,c.Description CategoryDesc" & vbCrLf
    s = s & "      ,l.area community,l.Description CommunityDesc,'' Assembly,m.OptionID,m.OrderUOM AssemblyUOM" & vbCrLf
    s = s & "      ,'' Series,'' Style,'' JCExtra,0 FloorArea,0 Bedrooms,0 Bathrooms,0 ConstCutOff" & vbCrLf
    s = s & "      ,p.IncentiveRetail" & vbCrLf
    s = s & "      ,p.IncludeTax" & vbCrLf
    s = s & "      ,p.Markup" & vbCrLf
    s = s & "      ,p.COMarkup" & vbCrLf
    s = s & "      ,p.Margin" & vbCrLf
    s = s & "      ,p.Roundto" & vbCrLf
    s = s & "      ,CASE isnull(p.assembly,'') when '' THEN m.RetailPretax ELSE p.Pretax END Pretax" & vbCrLf
    s = s & "      ,p.Tax" & vbCrLf
    s = s & "      ,CASE isnull(p.assembly,'') when '' THEN m.RetailPretax ELSE p.Pretax END COPretax" & vbCrLf
    s = s & "      ,p.COTax" & vbCrLf
    s = s & "      ,0 IncludedInSpec" & vbCrLf
    For i = 2 To 10
        s = s & "      ,0 Pretax" & i & vbCrLf
        s = s & "      ,0 Tax" & i & vbCrLf
        s = s & "      ,0 IncludedInSpec" & i & vbCrLf
    Next
    s = s & "      ,0,0,0,0,'','','',null,null,null,null,null" & vbCrLf
    s = s & ",0 DesignCenterSalesOnly,0 SelectByRoom,0 DisplayTotalOnly" & vbCrLf
    s = s & "FROM tblphaseitem m" & vbCrLf
    s = s & "  LEFT OUTER JOIN tblDBAssemblyPrices p ON(p.Community=" & DbQuote(Str, community) & " AND p.CommunityPhase=" & DbQuote(Str, CommunityPhase) & " AND p.Assembly='' AND p.Model='' AND m.OptionID=p.OptionID)" & vbCrLf
    s = s & "  LEFT OUTER JOIN tblCategories c ON m.OptionCategory=c.Category" & vbCrLf
    s = s & "  LEFT OUTER JOIN tblLocality l ON l.area=" & DbQuote(Str, community) & vbCrLf
    s = s & "WHERE m.DivisionID = " & HFApp.DivisionID & " and isnull(m.OptionID,'')=" & DbQuote(Str, OptionID) & vbCrLf
    Set rs = HFApp.SqlExec(s)
Y = Y & "query " & format(Timer() - X, "0.000") & vbCrLf: X = Timer()

    If rs.EOF Then Exit Sub
    
    mDirty = True
    With gData
    
        AssemblyType = Val("" & rs("AssemblyType"))
        r = .Rows
        .AddItem ""
        .Cell(flexcpChecked, r, .ColIndex("ReadyToPublish")) = flexChecked
        
        Select Case True
            Case AssemblyType = atModel
                .TextMatrix(r, .ColIndex("Community")) = IIf(HFApp.Options(ModelByArea), "" & rs("Community"), "")
                .TextMatrix(r, .ColIndex("CommunityDesc")) = IIf(HFApp.Options(ModelByArea), "" & rs("CommunityDesc"), "")
                .TextMatrix(r, .ColIndex("CommunityPhase")) = IIf(HFApp.Options(ModelsByArea_Phase), CommunityPhase, "")
            Case AssemblyType = atoption
                .TextMatrix(r, .ColIndex("Community")) = IIf(HFApp.Options(OptionByArea), "" & rs("Community"), "")
                .TextMatrix(r, .ColIndex("CommunityDesc")) = IIf(HFApp.Options(OptionByArea), "" & rs("CommunityDesc"), "")
                .TextMatrix(r, .ColIndex("CommunityPhase")) = IIf(HFApp.Options(OptionByAreaPhase), CommunityPhase, "")
            Case AssemblyType = atGlobal
                .TextMatrix(r, .ColIndex("Community")) = IIf(HFApp.Options(GlobalOptionByArea), "" & rs("Community"), "")
                .TextMatrix(r, .ColIndex("CommunityDesc")) = IIf(HFApp.Options(GlobalOptionByArea), "" & rs("CommunityDesc"), "")
                .TextMatrix(r, .ColIndex("CommunityPhase")) = IIf(HFApp.Options(GlobalOptionByAreaPhase), CommunityPhase, "")
            Case AssemblyType = atDesignCenter
                .TextMatrix(r, .ColIndex("Community")) = IIf(HFApp.Options(DCOptionByArea), "" & rs("Community"), "")
                .TextMatrix(r, .ColIndex("CommunityDesc")) = IIf(HFApp.Options(DCOptionByArea), "" & rs("CommunityDesc"), "")
                .TextMatrix(r, .ColIndex("CommunityPhase")) = IIf(HFApp.Options(DCOptionByAreaPhase), CommunityPhase, "")
        End Select
        
        .TextMatrix(r, .ColIndex("SourceCommunity")) = SourceCommunity
        
        .TextMatrix(r, .ColIndex("Model")) = "" & rs("Model")
        If AssemblyType = atoption Then
            .TextMatrix(r, .ColIndex("ModelDescription")) = "" & rs("ModelDescription")
        End If
        .TextMatrix(r, .ColIndex("Elevation")) = "" & rs("Elevation")
        .TextMatrix(r, .ColIndex("Description")) = "" & rs("Description")
        .TextMatrix(r, .ColIndex("Notes")) = "" & rs("Notes")
        .TextMatrix(r, .ColIndex("Comments")) = "" & rs("Comments")
        .TextMatrix(r, .ColIndex("AssemblyType")) = AssemblyType
        .TextMatrix(r, .ColIndex("Category")) = "" & rs("Category")
        .TextMatrix(r, .ColIndex("CategoryDesc")) = "" & rs("CategoryDesc")
        .TextMatrix(r, .ColIndex("JCExtra")) = "" & rs("JCExtra")
        .TextMatrix(r, .ColIndex("Series")) = series
        .TextMatrix(r, .ColIndex("FloorArea")) = Val("" & rs("FloorArea"))
        .TextMatrix(r, .ColIndex("Bedrooms")) = "" & rs("Bedrooms")
        .TextMatrix(r, .ColIndex("Bathrooms")) = "" & rs("Bathrooms")
        .TextMatrix(r, .ColIndex("Style")) = "" & rs("Style")
        
        .TextMatrix(r, .ColIndex("DCSalesOnly")) = "" & rs("DesignCenterSalesOnly")
        .TextMatrix(r, .ColIndex("SelectByRoom")) = "" & rs("SelectByRoom")
        .TextMatrix(r, .ColIndex("DisplayTotalOnly")) = "" & rs("DisplayTotalOnly")
        
        .TextMatrix(r, .ColIndex("Color")) = "" & rs("Color")
        .TextMatrix(r, .ColIndex("Location")) = "" & rs("Location")
        .TextMatrix(r, .ColIndex("Qty")) = Val("" & rs("Qty"))
        If .TextMatrix(r, .ColIndex("Qty")) = "0" Then .TextMatrix(r, .ColIndex("Qty")) = "1"
        .TextMatrix(r, .ColIndex("IncludedOption")) = "" & rs("IncludedOption")
        
        .TextMatrix(r, .ColIndex("ConstCutOff")) = Val("" & rs("ConstCutOff"))
        .TextMatrix(r, .ColIndex("SpecDocument")) = "" & rs("SpecDocument")
        .TextMatrix(r, .ColIndex("GraphicPath")) = "" & rs("GraphicPath")
        .TextMatrix(r, .ColIndex("MaxWidth")) = Val("" & rs("MaxWidth"))
        .TextMatrix(r, .ColIndex("MaxLength")) = Val("" & rs("MaxLength"))
        .Cell(flexcpChecked, r, .ColIndex("Inactive")) = flexUnchecked
        
        
        .TextMatrix(r, .ColIndex("Markup")) = Val("" & rs("Markup"))
        .TextMatrix(r, .ColIndex("Margin")) = Val("" & rs("Margin"))
        .TextMatrix(r, .ColIndex("Roundto")) = Val("" & rs("Roundto"))
        .TextMatrix(r, .ColIndex("IncludeTax")) = "" & rs("IncludeTax")
        
        If Val("" & rs("IncludeTax")) = 1 Then
            .TextMatrix(r, .ColIndex("Pretax")) = Val("" & rs("Pretax"))
        Else
            .TextMatrix(r, .ColIndex("Pretax")) = RoundToPrice(Val("" & rs("Pretax")), Val("" & rs("Roundto")))
        End If
        .TextMatrix(r, .ColIndex("Assembly")) = "" & rs("Assembly")
        .TextMatrix(r, .ColIndex("AssemblyUOM")) = "" & rs("AssemblyUOM")
        .TextMatrix(r, .ColIndex("OptionID")) = "" & rs("OptionID")

        
        .TextMatrix(r, .ColIndex("InSpec")) = "" & rs("IncludedInSpec")
        For i = 2 To 10
            .TextMatrix(r, .ColIndex("Pretax" & i)) = Val("" & rs("Pretax" & i))
            .TextMatrix(r, .ColIndex("Tax" & i)) = Val("" & rs("Tax" & i))
            .TextMatrix(r, .ColIndex("InSpec" & i)) = Val("" & rs("IncludedInSpec" & i))
        Next
        
        .TextMatrix(r, .ColIndex("Tax")) = CalcTax(.ValueMatrix(r, .ColIndex("AssemblyType")), .ValueMatrix(r, .ColIndex("Pretax")))
        .TextMatrix(r, .ColIndex("Total")) = Round(.ValueMatrix(r, .ColIndex("Pretax")) + .ValueMatrix(r, .ColIndex("Tax")), 2)
        
        'get assembly costs
        .TextMatrix(r, .ColIndex("Cost")) = 0
        .TextMatrix(r, .ColIndex("LandCost")) = IIf(AssemblyType = atModel, GetAssemblyCost(r, HFApp.Options(LandPhase), HFApp.Options(LandItem)), 0)
        .TextMatrix(r, .ColIndex("IncentiveCost")) = IIf(AssemblyType = atModel, GetAssemblyCost(r, HFApp.Options(IncentivePhase), HFApp.Options(IncentiveItem)), 0)
        .TextMatrix(r, .ColIndex("ConstructionCost")) = .ValueMatrix(r, .ColIndex("Cost")) - .ValueMatrix(r, .ColIndex("LandCost")) - .ValueMatrix(r, .ColIndex("IncentiveCost"))
        .TextMatrix(r, .ColIndex("IncentiveRetail")) = "" & rs("IncentiveRetail")
        
        'replace incentive cost with the cost percent on this worksheet
        .TextMatrix(r, .ColIndex("IncentiveCost")) = Round(.ValueMatrix(r, .ColIndex("IncentiveRetail")) * Val(txtIncentiveCostPercent) / 100, 2)
        .TextMatrix(r, .ColIndex("Cost")) = .ValueMatrix(r, .ColIndex("ConstructionCost")) + .ValueMatrix(r, .ColIndex("LandCost")) + .ValueMatrix(r, .ColIndex("IncentiveCost"))

        .TextMatrix(r, .ColIndex("ColorListID")) = "" & rs("ColorListID")
        .TextMatrix(r, .ColIndex("StyleListID")) = "" & rs("StyleListID")
        .TextMatrix(r, .ColIndex("FinishListID")) = "" & rs("FinishListID")
        .TextMatrix(r, .ColIndex("OtherListID")) = "" & rs("OtherListID")
        .TextMatrix(r, .ColIndex("StyleValue")) = "" & rs("StyleValue")
        .TextMatrix(r, .ColIndex("FinishValue")) = "" & rs("FinishValue")
        .TextMatrix(r, .ColIndex("OtherValue")) = "" & rs("OtherValue")
        If AssemblyType = atDesignCenter Then
        'calc markup & margins for dc options
            For i = 2 To 10
            
                If .ValueMatrix(r, .ColIndex("Cost")) = 0 Then
                    .TextMatrix(r, .ColIndex("Markup" & i)) = "0"
                Else
                    .TextMatrix(r, .ColIndex("Markup" & i)) = (.ValueMatrix(r, .ColIndex("Pretax" & i)) - .ValueMatrix(r, .ColIndex("Cost"))) / .ValueMatrix(r, .ColIndex("Cost")) * 100
                End If
                
                
                If .ValueMatrix(r, .ColIndex("Pretax" & i)) = 0 Then
                    .TextMatrix(r, .ColIndex("Margin" & i)) = "0"
                Else
                    .TextMatrix(r, .ColIndex("Margin" & i)) = ((.ValueMatrix(r, .ColIndex("Pretax" & i)) - (.ValueMatrix(r, .ColIndex("Cost")))) / .ValueMatrix(r, .ColIndex("Pretax" & i))) * 100
                End If
            Next
        End If
        
        
        .TextMatrix(r, .ColIndex("COMarkup")) = Val("" & rs("Markup"))
        .TextMatrix(r, .ColIndex("COMargin")) = Val("" & rs("Margin"))
        .TextMatrix(r, .ColIndex("COPretax")) = Val("" & rs("COPretax"))
        .TextMatrix(r, .ColIndex("COTax")) = Round(.ValueMatrix(r, .ColIndex("COPretax")) * HFApp.Options(GST_Rate) / 100 * HFApp.Options(PST_Rate) / 100, 2)
        .TextMatrix(r, .ColIndex("COTotal")) = Round(.ValueMatrix(r, .ColIndex("COPretax")) + .ValueMatrix(r, .ColIndex("COTax")), 2)
        .Col = .ColIndex("Roundto")
        
        .Row = r
'Y = Y & "add row " & format(Timer() - X, "0.000") & vbCrLf: X = Timer()

        
'MsgBox y
        
    End With
    
End Sub

Private Function GetAssemblyCost(Row As Long, Optional EstPhase As String, Optional EstItem As String, Optional UseVendorPrices As Boolean = False) As Double
On Error Resume Next
    Dim s As String
    With gData
        
        s = ""
        If .TextMatrix(Row, .ColIndex("Assembly")) = "" Then
            If UseVendorPrices Then
                s = s & "select isnull(dbo.Purch_GetItemRate(" & DbQuote(Num, cboCostBasis.ListIndex + 1) & vbCrLf
                s = s & "                                   ,0" & vbCrLf
                s = s & "                                   ," & DbQuote(Str, .TextMatrix(Row, .ColIndex("Community"))) & vbCrLf
                s = s & "                                   ," & DbQuote(Str, .TextMatrix(Row, .ColIndex("CommunityPhase"))) & vbCrLf
                s = s & "                                   ,'','','',i.Phase,i.Item,0,'',GETDATE()," & HFApp.DivisionID & "),0)" & vbCrLf
                s = s & " from tblphaseitem i where i.DivisionID = " & HFApp.DivisionID & " and i.optionid=" & DbQuote(Str, .TextMatrix(Row, .ColIndex("OptionID")))
            Else
                s = "select max(price) from tblphaseitem where DivisionID = " & HFApp.DivisionID & " and optionid=" & DbQuote(Str, .TextMatrix(Row, .ColIndex("OptionID")))
            End If
        Else
            If UseVendorPrices Then
                s = ""
                s = s & "select round(sum(isnull(d.OrderQty,0)*" & vbCrLf
                If Not MaxVendorPricing Then
                    s = s & "                 isnull(dbo.Purch_GetItemRate(" & DbQuote(Num, cboCostBasis.ListIndex + 1) & vbCrLf
                    s = s & "                                             ,0" & vbCrLf
                    s = s & "                                             ," & DbQuote(Str, .TextMatrix(Row, .ColIndex("Community"))) & vbCrLf
                    s = s & "                                             ," & DbQuote(Str, .TextMatrix(Row, .ColIndex("CommunityPhase"))) & vbCrLf
                    s = s & "                                             ,d.Assembly,d.Model,d.OptionID,i.Phase,i.Item,d.sequence,'',GETDATE()," & HFApp.DivisionID & "),0)*" & vbCrLf
                Else
                    s = s & "                 isnull(dbo.Purch_GetMaxItemRate(" & DbQuote(Num, cboCostBasis.ListIndex + 1) & vbCrLf
                    s = s & "                                             ,0" & vbCrLf
                    s = s & "                                             ," & DbQuote(Str, .TextMatrix(Row, .ColIndex("Community"))) & vbCrLf
                    s = s & "                                             ," & DbQuote(Str, .TextMatrix(Row, .ColIndex("CommunityPhase"))) & vbCrLf
                    s = s & "                                             ,d.Assembly,d.Model,d.OptionID,i.Phase,i.Item,d.sequence,'',GETDATE()," & HFApp.DivisionID & "),0)*" & vbCrLf
                End If
                s = s & "                 (100+isnull(t.jcrate,0))/100)" & vbCrLf
                s = s & "            ,2) cost" & vbCrLf
            Else
                s = ""
                s = s & "select round(sum(isnull(d.OrderQty,0)*isnull(d.Rate,0)*(100+isnull(t.jcrate,0))/100),2) cost" & vbCrLf
            End If
            s = s & "from tblDBAssemblyDetails d" & vbCrLf
            s = s & "left outer join CommunityStandards s" & vbCrLf
            s = s & "    ON(s.Community=" & DbQuote(Str, .TextMatrix(Row, .ColIndex("Community")))
            s = s & "   AND s.CommunityPhase=dbo.Purch_GetCommunityStandardPhase(" & DbQuote(Str, .TextMatrix(Row, .ColIndex("Community"))) & "," & DbQuote(Str, .TextMatrix(Row, .ColIndex("CommunityPhase"))) & ", d.phase , d.Item)" & vbCrLf
            s = s & "   AND s.StdPhase=d.Phase AND s.StdItem=d.Item)" & vbCrLf
            s = s & "join tblphaseitem i on(i.DivisionID = d.DivisionID and i.phase=isnull(s.phase,d.phase) and i.item=isnull(s.item,d.item))" & vbCrLf
            s = s & "left outer join taxgroups t on(t.DivisionID = " & HFApp.DivisionID & " and t.taxgroup=dbo.Purch_GetDefaultTaxGroup(''," & DbQuote(Str, .TextMatrix(Row, .ColIndex("Community"))) & "," & DbQuote(Str, .TextMatrix(Row, .ColIndex("CommunityPhase"))) & ",d.Model,d.Assembly,i.Phase,i.Item,dbo.Purch_GetCommunityVendor(" & DbQuote(Str, .TextMatrix(Row, .ColIndex("Community"))) & ", i.POIndex," & HFApp.DivisionID & "),i.JCCategory,d.DivisionID))" & vbCrLf
            s = s & " WHERE d.DivisionID = " & HFApp.DivisionID & " and d.Community=" & DbQuote(Str, .TextMatrix(Row, .ColIndex("SourceCommunity"))) & vbCrLf
            s = s & "   AND d.Assembly=" & DbQuote(Str, .TextMatrix(Row, .ColIndex("Assembly"))) & vbCrLf
            s = s & "   AND d.Model=" & DbQuote(Str, .TextMatrix(Row, .ColIndex("Model"))) & vbCrLf
            s = s & "   AND d.OptionID=" & DbQuote(Str, .TextMatrix(Row, .ColIndex("OptionID"))) & vbCrLf
            If EstPhase <> "" Then s = s & "   AND d.Phase=" & DbQuote(Str, EstPhase) & vbCrLf
            If EstItem <> "" Then s = s & "   AND d.Item=" & DbQuote(Str, EstItem) & vbCrLf
        End If
        
        
        
        
        
        GetAssemblyCost = Val("" & HFApp.SqlExec(s)(0))
    End With
End Function


Private Sub RefreshLandCosts()
On Error Resume Next
    Dim s As String
    Dim Row As Long
    
    With gData
        For Row = 2 To gData.Rows - 1
            
            If .ValueMatrix(Row, .ColIndex("AssemblyType")) = atModel Then
        
                s = ""
                s = s & "SELECT dbo.Purch_GetItemRate(" & vbCrLf
                s = s & " " & DbQuote(Num, cboCostBasis.ListIndex + 1) & vbCrLf
                s = s & ",0" & vbCrLf
                s = s & "," & DbQuote(Str, .TextMatrix(Row, .ColIndex("Community"))) & vbCrLf
                s = s & "," & DbQuote(Str, .TextMatrix(Row, .ColIndex("CommunityPhase"))) & vbCrLf
                s = s & "," & DbQuote(Str, .TextMatrix(Row, .ColIndex("Assembly"))) & vbCrLf
                s = s & "," & DbQuote(Str, .TextMatrix(Row, .ColIndex("Model"))) & vbCrLf
                s = s & ",''" & vbCrLf
                s = s & "," & DbQuote(Str, HFApp.Options(LandPhase)) & vbCrLf
                s = s & "," & DbQuote(Str, HFApp.Options(LandItem)) & vbCrLf
                s = s & ",0" & vbCrLf
                s = s & ",dbo.purch_getcommunityvendor(" & DbQuote(Str, .TextMatrix(Row, .ColIndex("Community"))) & ",POIndex," & HFApp.DivisionID & ")" & vbCrLf
                s = s & ",GETDATE()," & HFApp.DivisionID & ")" & vbCrLf
                s = s & "FROM tblphaseitem" & vbCrLf
                s = s & "WHERE DivisionID = " & HFApp.DivisionID & " and Phase=" & DbQuote(Str, HFApp.Options(LandPhase)) & vbCrLf
                s = s & "  AND Item=" & DbQuote(Str, HFApp.Options(LandItem)) & vbCrLf
                
                .TextMatrix(Row, .ColIndex("LandCost")) = Val("" & HFApp.SqlExec(s)(0))
                
            End If
        Next
    End With
End Sub

Private Sub RefreshCosts(Optional DoAllRows As Boolean = False, Optional Row As Long)
On Error Resume Next
    
    Dim startR As Long
    Dim stopR  As Long
    Dim r    As Long
    Dim s As String
    Dim p As Double
    Dim bUpdateAttributes As Boolean
    Dim rs As ADODB.Recordset
    
    If DoAllRows Then
        If Not SaveData(False) Then
            Exit Sub
        Else
            bUpdateAttributes = MsgBox("Refresh attributes from Assemblies?", vbYesNo, "Refresh Attributes") = vbYes
        End If
    End If
    Screen.MousePointer = vbHourglass
'    If DoAllRows Then
'        s = "Purch_GetAssemblySalesSheetCost " & DbQuote(Num, mWorksheet) & ",'','','','',''"
'        Call HFApp.SqlExec(s)
'        Call LoadData
'    Else
        With gData
            If DoAllRows Then
                startR = 2
                stopR = .Rows - 1
                s = ""
                s = s & "exec dbo.Purch_GetAssemblySalesSheetCost"
                s = s & " " & DbQuote(Num, mWorksheet)
                s = s & "," & DbQuote(Str, "")
                s = s & "," & DbQuote(Str, "")
                s = s & "," & DbQuote(Str, "")
                s = s & "," & DbQuote(Str, "")
                s = s & "," & DbQuote(Str, "")
                Call HFApp.SqlExec(s)
            Else
                If Row > 1 Then
                    startR = Row
                    stopR = Row
                Else
                    startR = Min(.Row, .RowSel)
                    stopR = Max(.Row, .RowSel)
                End If
                If startR < 2 Then startR = 2
            End If
            For r = startR To stopR
                'update costs
                If Not DoAllRows Then
                    s = ""
                    s = s & "exec dbo.Purch_GetAssemblySalesSheetCost"
                    s = s & " " & DbQuote(Num, mWorksheet)
                    s = s & "," & DbQuote(Str, .TextMatrix(r, .ColIndex("Assembly")))
                    s = s & "," & DbQuote(Str, .TextMatrix(r, .ColIndex("Model")))
                    s = s & "," & DbQuote(Str, .TextMatrix(r, .ColIndex("OptionID")))
                    s = s & "," & DbQuote(Str, .TextMatrix(r, .ColIndex("Community")))
                    s = s & "," & DbQuote(Str, .TextMatrix(r, .ColIndex("CommunityPhase")))
                    Call HFApp.SqlExec(s)
                End If
                 
                If bUpdateAttributes Then
                    s = ""
                    s = s & "select m.style ModelStyle, m.constcutoff,m.location,m.bedrooms,m.bathrooms,m.floorarea,m.Description,m.notes,m.comments,m.ColorListID,m.StyleListID"
                    s = s & ",m.FinishListID,m.OtherListID,m.Color,m.StyleValue,m.FinishValue,m.OtherValue,m.MaxLength,m.MaxWidth,m.DesignCenterSalesOnly"
                    s = s & ",m.SelectByRoom,m.DisplayTotalOnly,m.category,c.description CategoryDesc" & vbCrLf
                    s = s & "from tbldbassemblymaster m" & vbCrLf
                    s = s & "left join tblcategories c on m.category=c.category" & vbCrLf
                    s = s & " where m.Assembly = " & DbQuote(Str, .TextMatrix(r, .ColIndex("Assembly"))) & vbCrLf
                    s = s & "   and isnull(m.model,'')=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Model"))) & vbCrLf
                    s = s & "   and isnull(m.optionid,'')=" & DbQuote(Str, .TextMatrix(r, .ColIndex("OptionID"))) & vbCrLf
                    s = s & "   and isnull(m.community,'')=" & DbQuote(Str, .TextMatrix(r, .ColIndex("SourceCommunity"))) & vbCrLf
                    s = s & "   AND m.DivisionID = " & HFApp.DivisionID & vbCrLf
                    Set rs = HFApp.SqlExec(s)
                    If Not rs.EOF Then
                    
                         .TextMatrix(r, .ColIndex("ConstCutOff")) = "" & rs("ConstCutOff")
                         .TextMatrix(r, .ColIndex("Location")) = "" & rs("Location")
                         .TextMatrix(r, .ColIndex("Bedrooms")) = "" & rs("Bedrooms")
                         .TextMatrix(r, .ColIndex("Bathrooms")) = "" & rs("Bathrooms")
                         .TextMatrix(r, .ColIndex("FloorArea")) = Val("" & rs("FloorArea"))
                    
                         .TextMatrix(r, .ColIndex("Notes")) = "" & rs("Notes")
                         .TextMatrix(r, .ColIndex("Comments")) = "" & rs("Comments")
                         .TextMatrix(r, .ColIndex("Style")) = "" & rs("ModelStyle")
                         
                         .TextMatrix(r, .ColIndex("ColorListID")) = "" & rs("ColorListID")
                         .TextMatrix(r, .ColIndex("StyleListID")) = "" & rs("StyleListID")
                         .TextMatrix(r, .ColIndex("FinishListID")) = "" & rs("FinishListID")
                         .TextMatrix(r, .ColIndex("OtherListID")) = "" & rs("OtherListID")
                         .TextMatrix(r, .ColIndex("Color")) = "" & rs("Color")
                         .TextMatrix(r, .ColIndex("StyleValue")) = "" & rs("StyleValue")
                         .TextMatrix(r, .ColIndex("FinishValue")) = "" & rs("FinishValue")
                         .TextMatrix(r, .ColIndex("OtherValue")) = "" & rs("OtherValue")
                         
                         
                         .TextMatrix(r, .ColIndex("Category")) = "" & rs("Category")
                         .TextMatrix(r, .ColIndex("CategoryDesc")) = "" & rs("CategoryDesc")
                         
                         
                         .TextMatrix(r, .ColIndex("DCSalesOnly")) = "" & rs("DesignCenterSalesOnly")
                         .TextMatrix(r, .ColIndex("SelectByRoom")) = "" & rs("SelectByRoom")
                         .TextMatrix(r, .ColIndex("DisplayTotalOnly")) = "" & rs("DisplayTotalOnly")
                         
                         .TextMatrix(r, .ColIndex("MaxWidth")) = "" & rs("MaxWidth")
                         .TextMatrix(r, .ColIndex("MaxLength")) = "" & rs("MaxLength")
                         
                         .TextMatrix(r, .ColIndex("Description")) = "" & rs("Description")
                         
                    End If
                End If
                
                'get total cost
                s = ""
                s = s & "select cost from tblsalessheetdetails" & vbCrLf
                s = s & "where worksheet=" & DbQuote(Num, mWorksheet) & vbCrLf
                s = s & "  and assembly=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Assembly"))) & vbCrLf
                s = s & "  and model=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Model"))) & vbCrLf
                s = s & "  and optionid=" & DbQuote(Str, .TextMatrix(r, .ColIndex("OptionID"))) & vbCrLf
                s = s & "  and community=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Community"))) & vbCrLf
                s = s & "  and communityphase=" & DbQuote(Str, .TextMatrix(r, .ColIndex("CommunityPhase"))) & vbCrLf
                .TextMatrix(r, .ColIndex("Cost")) = Val("" & HFApp.SqlExec(s)(0))
                
                
                'get land and incentive costs
                If HFApp.Options(LandPhase) = "" Or HFApp.Options(LandItem) = "" Or .TextMatrix(r, .ColIndex("Assembly")) = "" Or .ValueMatrix(r, .ColIndex("AssemblyType")) <> atModel Then
                    .TextMatrix(r, .ColIndex("LandCost")) = 0
                Else
                    .TextMatrix(r, .ColIndex("LandCost")) = GetAssemblyCost(r, HFApp.Options(LandPhase), HFApp.Options(LandItem), True)
                End If
                If HFApp.Options(IncentivePhase) = "" Or HFApp.Options(IncentiveItem) = "" Or .TextMatrix(r, .ColIndex("Assembly")) = "" Then
                    .TextMatrix(r, .ColIndex("IncentiveCost")) = 0
                Else
                    .TextMatrix(r, .ColIndex("IncentiveCost")) = GetAssemblyCost(r, HFApp.Options(IncentivePhase), HFApp.Options(IncentiveItem), True)
                End If
                        
                'construction costs = total cost - land and incentive
                .TextMatrix(r, .ColIndex("ConstructionCost")) = .ValueMatrix(r, .ColIndex("Cost")) - .ValueMatrix(r, .ColIndex("LandCost")) - .ValueMatrix(r, .ColIndex("IncentiveCost"))
                
                
                Call CalcData(False, "Cost", r)
            
            Next
        End With
    'End If
    Screen.MousePointer = vbDefault
End Sub



Public Function SaveData(prompt As Boolean) As Boolean
On Error GoTo eh:

    Dim rc As Long
    If Not mDirty Then
        SaveData = True
        Exit Function
    End If
    rc = vbYes
    If prompt Then rc = MsgBox("This data has changed." & vbCrLf & vbCrLf & "Do you want to save the changes?", vbExclamation + vbYesNoCancel, Me.Caption)
    Select Case rc
        Case vbNo:     SaveData = True:     Exit Function
        Case vbCancel: Exit Function
    End Select
    If Not ValidateForm() Then Exit Function
    
    Screen.MousePointer = vbHourglass
    Dim s As String
    Dim r As Long
    Dim i As Long
    Dim rs As Recordset
    
    If mWorksheet = 0 Then
        s = ""
        s = s & "INSERT INTO tblSalesSheetMaster(DivisionID,WorksheetType,Description,CostsEffectiveDate,CostBasisType,CostBasisName,IncentiveCostPercent,Community,CommunityPhase,UStmp,TStmp)" & vbCrLf
        s = s & "VALUES(" & HFApp.DivisionID & "," & DbQuote(Num, mWorksheetType) & vbCrLf
        s = s & "      ," & DbQuote(Str, txtDescription.Text) & vbCrLf
        s = s & "      ," & DbQuote(Date, CostEffectiveDate) & vbCrLf
        s = s & "      ," & DbQuote(Num, cboCostBasis.ListIndex) & vbCrLf
        s = s & "      ," & DbQuote(Str, cboCostBasis.Text) & vbCrLf
        s = s & "      ," & DbQuote(Num, txtIncentiveCostPercent.Text) & vbCrLf
        s = s & "      ," & DbQuote(Str, mCommunity) & vbCrLf
        s = s & "      ," & DbQuote(Str, mCommunityPhase) & vbCrLf
        s = s & "      ," & DbQuote(Str, HFApp.LoginID) & vbCrLf
        s = s & "      ,GETDATE())" & vbCrLf
        Call HFApp.SqlExec(s)
        mWorksheet = HFApp.SqlIdentity("tblSalesSheetMaster")
        Me.Caption = mWorksheetView & " Worksheet (number " & mWorksheet & ")"
    Else
        s = ""
        s = s & "UPDATE tblSalesSheetMaster" & vbCrLf
        s = s & "SET Description=" & DbQuote(Str, txtDescription.Text) & vbCrLf
        s = s & "   ,CostsEffectiveDate=" & DbQuote(Date, CostEffectiveDate) & vbCrLf
        s = s & "   ,CostBasisType=" & DbQuote(Num, cboCostBasis.ListIndex) & vbCrLf
        s = s & "   ,CostbasisName=" & DbQuote(Str, cboCostBasis.Text) & vbCrLf
        s = s & "   ,IncentiveCostPercent=" & DbQuote(Num, txtIncentiveCostPercent.Text) & vbCrLf
        s = s & "   ,UStmp=" & DbQuote(Str, HFApp.LoginID) & vbCrLf
        s = s & "   ,TStmp=GETDATE()" & vbCrLf
        s = s & "WHERE Worksheet=" & DbQuote(Num, mWorksheet) & vbCrLf
        Call HFApp.SqlExec(s)
    End If
    Call HFApp.SqlExec("DELETE FROM tblSalesSheetCosts WHERE Worksheet=" & mWorksheet)
    Call HFApp.SqlExec("DELETE FROM tblSalesSheetDetails WHERE Worksheet=" & mWorksheet)
    
    'rewrite details
    With gData
        For r = 2 To .Rows - 1
            s = ""
            s = s & "INSERT INTO tblSalesSheetDetails(Worksheet,SourceCommunity,Community,CommunityPhase,Assembly,OptionID,JCExtra,Series,FloorArea,Bedrooms,Bathrooms,SpecDocument,GraphicPath,MaxWidth,MaxLength,Inactive,Style,ConstCutOff,Description,AssemblyUOM,Notes,comments,AssemblyType,Category,Markup,Margin,COMarkup,COMargin,Roundto,LandCost,IncentiveRetail,IncentiveCost,ConstructionCost,Cost,Pretax,Tax,COPretax,COTax,IncludeTax,Model,Elevation,Color,Qty,Location,IncludedOption"
            If .ValueMatrix(r, .ColIndex("AssemblyType")) = 4 Then
            
                For i = 2 To 10
                    s = s & ",Pretax" & i & ",Tax" & i & ",Markup" & i & ",Margin" & i & ",IncludedInSpec" & i
                Next
            End If
            s = s & ",ColorListID,StyleListID,FinishListID,OtherListID,StyleValue,FinishValue,OtherValue,IncludedInSpec,ReadytoPublish,DesignCenterSalesOnly,SelectByRoom,DisplayTotalOnly)" & vbCrLf
            s = s & "VALUES(" & DbQuote(Num, mWorksheet) & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("SourceCommunity"))) & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("Community"))) & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("CommunityPhase"))) & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("Assembly"))) & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("OptionID"))) & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("JCExtra"))) & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("Series"))) & vbCrLf
            s = s & "      ," & DbQuote(Num, .TextMatrix(r, .ColIndex("FloorArea"))) & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("Bedrooms"))) & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("Bathrooms"))) & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("SpecDocument"))) & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("GraphicPath"))) & vbCrLf
            s = s & "      ," & DbQuote(Num, .TextMatrix(r, .ColIndex("MaxWidth"))) & vbCrLf
            s = s & "      ," & DbQuote(Num, .TextMatrix(r, .ColIndex("MaxLength"))) & vbCrLf
            s = s & "      ," & IIf(.Cell(flexcpChecked, r, .ColIndex("Inactive")) = flexChecked, 1, 0) & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("Style"))) & vbCrLf
            s = s & "      ," & DbQuote(Num, .TextMatrix(r, .ColIndex("ConstCutOff"))) & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("Description"))) & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("AssemblyUOM"))) & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("Notes"))) & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("Comments"))) & vbCrLf
            s = s & "      ," & DbQuote(Num, .ValueMatrix(r, .ColIndex("AssemblyType"))) & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("Category"))) & vbCrLf
            s = s & "      ," & DbQuote(Num, .ValueMatrix(r, .ColIndex("Markup"))) & vbCrLf
            s = s & "      ," & DbQuote(Num, .ValueMatrix(r, .ColIndex("Margin"))) & vbCrLf
            s = s & "      ," & DbQuote(Num, .ValueMatrix(r, .ColIndex("COMarkup"))) & vbCrLf
            s = s & "      ," & DbQuote(Num, .ValueMatrix(r, .ColIndex("COMargin"))) & vbCrLf
            s = s & "      ," & DbQuote(Num, .ValueMatrix(r, .ColIndex("Roundto"))) & vbCrLf
            
            s = s & "      ," & DbQuote(Num, .ValueMatrix(r, .ColIndex("LandCost"))) & vbCrLf
            s = s & "      ," & DbQuote(Num, .ValueMatrix(r, .ColIndex("IncentiveRetail"))) & vbCrLf
            s = s & "      ," & DbQuote(Num, .ValueMatrix(r, .ColIndex("IncentiveCost"))) & vbCrLf
            s = s & "      ," & DbQuote(Num, .ValueMatrix(r, .ColIndex("ConstructionCost"))) & vbCrLf
            s = s & "      ," & DbQuote(Num, .ValueMatrix(r, .ColIndex("Cost"))) & vbCrLf
            
            s = s & "      ," & DbQuote(Num, .ValueMatrix(r, .ColIndex("Pretax"))) & vbCrLf
            s = s & "      ," & DbQuote(Num, .ValueMatrix(r, .ColIndex("Tax"))) & vbCrLf
            s = s & "      ," & DbQuote(Num, .ValueMatrix(r, .ColIndex("COPretax"))) & vbCrLf
            s = s & "      ," & DbQuote(Num, .ValueMatrix(r, .ColIndex("COTax"))) & vbCrLf
            s = s & "      ," & IIf(.Cell(flexcpChecked, r, .ColIndex("IncludeTax")) = flexChecked, 1, 0) & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("Model"))) & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("Elevation"))) & vbCrLf
            
            s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("Color"))) & vbCrLf
            s = s & "      ," & DbQuote(Num, .TextMatrix(r, .ColIndex("Qty"))) & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("Location"))) & vbCrLf
            s = s & "      ," & DbQuote(Bit, .TextMatrix(r, .ColIndex("IncludedOption"))) & vbCrLf
            
            If .ValueMatrix(r, .ColIndex("AssemblyType")) = 4 Then
            For i = 2 To 10
            s = s & "      ," & DbQuote(Num, .ValueMatrix(r, .ColIndex("Pretax" & i))) & vbCrLf
            s = s & "      ," & DbQuote(Num, .ValueMatrix(r, .ColIndex("Tax" & i))) & vbCrLf
            s = s & "      ," & DbQuote(Num, .ValueMatrix(r, .ColIndex("Markup" & i))) & vbCrLf
            s = s & "      ," & DbQuote(Num, .ValueMatrix(r, .ColIndex("Margin" & i))) & vbCrLf
            s = s & "      ," & IIf(.Cell(flexcpChecked, r, .ColIndex("InSpec" & i)) = flexChecked, 1, 0) & vbCrLf
            Next
            End If
            
            s = s & "      ," & DbQuote(Num, .TextMatrix(r, .ColIndex("ColorListID"))) & vbCrLf
            s = s & "      ," & DbQuote(Num, .TextMatrix(r, .ColIndex("StyleListID"))) & vbCrLf
            s = s & "      ," & DbQuote(Num, .TextMatrix(r, .ColIndex("FinishListID"))) & vbCrLf
            s = s & "      ," & DbQuote(Num, .TextMatrix(r, .ColIndex("OtherListID"))) & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("StyleValue"))) & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("FinishValue"))) & vbCrLf
            s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("OtherValue"))) & vbCrLf
            
            s = s & "      ," & IIf(.Cell(flexcpChecked, r, .ColIndex("InSpec")) = flexChecked, 1, 0) & vbCrLf
            s = s & "      ,1" & vbCrLf
            s = s & "      ," & DbQuote(Bit, .TextMatrix(r, .ColIndex("DCSalesOnly"))) & vbCrLf
            s = s & "      ," & DbQuote(Bit, .TextMatrix(r, .ColIndex("SelectByRoom"))) & vbCrLf
            s = s & "      ," & DbQuote(Bit, .TextMatrix(r, .ColIndex("DisplayTotalOnly"))) & ")"
            Call HFApp.SqlExec(s)
            

            'now update default vendor costs to show new land & incentive costs
'            If HFApp.Options(RecordSalesSheetCosts) Then
'                If .ValueMatrix(r, .ColIndex("AssemblyType")) = atModel Then
'                    Call SaveVendorCost(.TextMatrix(r, .ColIndex("Community")), .TextMatrix(r, .ColIndex("CommunityPhase")), .TextMatrix(r, .ColIndex("Model")), .TextMatrix(r, .ColIndex("Assembly")), HFApp.Options(LandPhase), HFApp.Options(LandItem), .ValueMatrix(r, .ColIndex("LandCost")))
'                    Call SaveVendorCost(.TextMatrix(r, .ColIndex("Community")), .TextMatrix(r, .ColIndex("CommunityPhase")), .TextMatrix(r, .ColIndex("Model")), .TextMatrix(r, .ColIndex("Assembly")), HFApp.Options(IncentivePhase), HFApp.Options(IncentiveItem), .ValueMatrix(r, .ColIndex("IncentiveCost")))
'                End If
'            End If
        Next
    End With
    
    
    If HFApp.Options(RecordSalesSheetCosts) Then
        'rewrite costs
        s = ""
        s = s & "INSERT INTO tblSalesSheetCosts(Worksheet,Community,CommunityPhase,Assembly,Model,OptionID,Sequence,Phase,Item,TakeoffQty,OrderQty,ConversionFactor,Vendor,Rate,POIndex,Location" & vbCrLf & "      "
        For i = 1 To 20
            s = s & ",WBS" & format(i, "00")
        Next
        s = s & vbCrLf & "      "
        For i = 21 To 40
            s = s & ",WBS" & format(i, "00")
        Next
        s = s & ")" & vbCrLf
        s = s & "--  ASSEMBLY ITEMS  --------------------  --------------------  --------------------  --------------------" & vbCrLf
        s = s & "SELECT sd.Worksheet" & vbCrLf
        s = s & "      ,sd.Community" & vbCrLf
        s = s & "      ,sd.CommunityPhase" & vbCrLf
        s = s & "      ,ad.Assembly" & vbCrLf
        s = s & "      ,ad.Model" & vbCrLf
        s = s & "      ,ad.OptionID" & vbCrLf
        s = s & "      ,ad.Sequence" & vbCrLf
        s = s & "      ,i.Phase" & vbCrLf
        s = s & "      ,i.Item" & vbCrLf
        s = s & "      ,ad.TakeoffQty" & vbCrLf
        s = s & "      ,ad.OrderQty" & vbCrLf
        s = s & "      ,i.ConversionFactor" & vbCrLf
        s = s & "      ,dbo.Purch_GetCommunityVendor(sd.Community,i.POIndex," & HFApp.DivisionID & ") Vendor" & vbCrLf
        s = s & "      ,dbo.Purch_GetItemRate(" & DbQuote(Num, cboCostBasis.ListIndex + 1) & "," & DbQuote(Num, mWorksheet) & ",sd.Community,sd.CommunityPhase,sd.Assembly,sd.Model,sd.OptionID,ad.Phase,ad.Item,ad.Sequence,dbo.Purch_GetCommunityVendor(sd.Community,i.POIndex," & HFApp.DivisionID & "),sm.CostsEffectiveDate," & HFApp.DivisionID & ") Rate" & vbCrLf
        s = s & "      ,isnull(nullif(ad.poindex,''),i.poindex)" & vbCrLf
        s = s & "      ,isnull(nullif(ad.Location,''),i.Location)" & vbCrLf & "      "
        For i = 1 To 20
            s = s & ",isnull(nullif(ad.WBS" & format(i, "00") & ",''),i.WBS" & format(i, "00") & ")"
        Next
        s = s & vbCrLf & "      "
        For i = 21 To 40
            s = s & ",isnull(nullif(ad.WBS" & format(i, "00") & ",''),i.WBS" & format(i, "00") & ")"
        Next
        s = s & vbCrLf
        s = s & "  FROM tblSalesSheetMaster sm " & vbCrLf
        s = s & "       INNER JOIN tblSalesSheetDetails sd ON(sm.Worksheet=sd.Worksheet)" & vbCrLf
        
        If HFApp.Options.ValueByName("GetAssemblyCommunity") = "True" Then
            s = s & "       INNER JOIN tblDBAssemblydetails ad ON(ad.DivisionID = sm.DivisionID and ad.Community=dbo.Purch_GetAssemblyCommunity(sd.assembly,sd.model,sd.optionid,sd.community) AND sd.Assembly=ad.Assembly AND sd.Model=ad.Model AND sd.OptionID=ad.OptionID)" & vbCrLf
        Else
            s = s & "       INNER JOIN tblDBAssemblydetails ad ON(ad.DivisionID = sm.DivisionID and ad.Community=sd.SourceCommunity AND sd.Assembly=ad.Assembly AND sd.Model=ad.Model AND sd.OptionID=ad.OptionID)" & vbCrLf
        End If
        s = s & "       join tblphaseitem i on(i.DivisionID = ad.DivisionID and i.phase=ad.phase and i.item=ad.item)" & vbCrLf
        s = s & " WHERE sm.Worksheet=" & DbQuote(Num, mWorksheet) & vbCrLf
        s = s & "UNION ALL" & vbCrLf
        s = s & "--  COMPONENT ITEMS  --------------------  --------------------  --------------------  --------------------" & vbCrLf
        s = s & "SELECT sd.Worksheet" & vbCrLf
        s = s & "      ,sd.Community" & vbCrLf
        s = s & "      ,sd.CommunityPhase" & vbCrLf
        s = s & "      ,ad.Assembly" & vbCrLf
        s = s & "      ,ad.Model" & vbCrLf
        s = s & "      ,ad.OptionID" & vbCrLf
        s = s & "      ,ad.Sequence" & vbCrLf
        s = s & "      ,i.Phase" & vbCrLf
        s = s & "      ,i.Item" & vbCrLf
        s = s & "      ,ad.TakeoffQty" & vbCrLf
        s = s & "      ,ad.OrderQty" & vbCrLf
        s = s & "      ,i.ConversionFactor" & vbCrLf
        s = s & "      ,dbo.Purch_GetCommunityVendor(sd.Community,i.POIndex," & HFApp.DivisionID & ") Vendor" & vbCrLf
        s = s & "      ,dbo.Purch_GetItemRate(" & DbQuote(Num, cboCostBasis.ListIndex + 1) & "," & DbQuote(Num, mWorksheet) & ",sd.Community,sd.CommunityPhase,sd.Assembly,sd.Model,sd.OptionID,ad.Phase,ad.Item,ad.Sequence,dbo.Purch_GetCommunityVendor(sd.Community,i.POIndex," & HFApp.DivisionID & "),sm.CostsEffectiveDate," & HFApp.DivisionID & ") Rate" & vbCrLf
        s = s & "      ,isnull(nullif(ad.poindex,''),i.poindex)" & vbCrLf
        s = s & "      ,isnull(nullif(ad.Location,''),i.Location)" & vbCrLf & "      "
        For i = 1 To 20
            s = s & ",isnull(nullif(ad.WBS" & format(i, "00") & ",''),i.WBS" & format(i, "00") & ")"
        Next
        s = s & vbCrLf & "      "
        For i = 21 To 40
            s = s & ",isnull(nullif(ad.WBS" & format(i, "00") & ",''),i.WBS" & format(i, "00") & ")"
        Next
        s = s & vbCrLf
        s = s & "  FROM tblSalesSheetMaster sm " & vbCrLf
        s = s & "       join tblSalesSheetDetails sd ON(sm.Worksheet=sd.Worksheet)" & vbCrLf
        s = s & "       join tblDBAssemblyMaster m ON(m.DivisionID = sm.DivisionID and m.Community=sd.SourceCommunity AND sd.Assembly=m.Assembly AND m.Model=sd.Model AND sd.OptionID=m.OptionID)" & vbCrLf
        s = s & "       join tbldbassemblycomponents ac on m.assemblyid=ac.parentassemblyID" & vbCrLf
        s = s & "       join tbldbassemblydetails ad on ac.componentassemblyid=ad.assemblyid" & vbCrLf
        s = s & "       join tblphaseitem i on(i.DivisionID = ad.DivisionID and i.phase=ad.phase and i.item=ad.item)" & vbCrLf
        s = s & " WHERE sm.Worksheet=" & DbQuote(Num, mWorksheet) & vbCrLf
        s = s & "UNION ALL" & vbCrLf
        s = s & "--  GLOBAL ITEMS  --------------------  --------------------  --------------------  --------------------" & vbCrLf
        s = s & "SELECT sd.Worksheet" & vbCrLf
        s = s & "      ,sd.Community" & vbCrLf
        s = s & "      ,sd.CommunityPhase" & vbCrLf
        s = s & "      ,sd.Assembly" & vbCrLf
        s = s & "      ,sd.Model" & vbCrLf
        s = s & "      ,sd.OptionID" & vbCrLf
        s = s & "      ,0 Sequence" & vbCrLf
        s = s & "      ,i.Phase" & vbCrLf
        s = s & "      ,i.Item" & vbCrLf
        s = s & "      ,1/i.ConversionFactor TakeoffQty" & vbCrLf
        s = s & "      ,1 OrderQty" & vbCrLf
        s = s & "      ,i.ConversionFactor" & vbCrLf
        s = s & "      ,dbo.Purch_GetCommunityVendor(sd.Community,i.POIndex," & HFApp.DivisionID & ") Vendor" & vbCrLf
        s = s & "      ,dbo.Purch_GetItemRate(1,1,sd.Community,sd.CommunityPhase,'','',sd.OptionID,i.Phase,i.Item,0,'',sm.CostsEffectiveDate," & HFApp.DivisionID & ") Rate" & vbCrLf
        s = s & "      ,i.poindex" & vbCrLf
        s = s & "      ,i.Location" & vbCrLf & "      "
        For i = 1 To 20
            s = s & ",i.WBS" & format(i, "00")
        Next
        s = s & vbCrLf & "      "
        For i = 21 To 40
            s = s & ",i.WBS" & format(i, "00")
        Next
        s = s & vbCrLf
        s = s & "  FROM tblSalesSheetMaster sm " & vbCrLf
        s = s & "       INNER JOIN tblSalesSheetDetails sd ON(sm.Worksheet=sd.Worksheet)" & vbCrLf
        s = s & "       join tblphaseitem i on(i.DivisionID = sm.DivisionID and i.optionid=sd.optionid and isnull(sd.assembly,'')='')" & vbCrLf
        s = s & " WHERE sm.Worksheet=" & DbQuote(Num, mWorksheet) & vbCrLf
        
        Call HFApp.SqlExec(s)
    End If
    
 
 
    
    
    
 
    mDirty = False
    SaveData = True
    
    Screen.MousePointer = vbDefault
    Exit Function
eh: If InStr(1, Err.Description, "duplicate", vbTextCompare) Then
        Resume Next
    Else
        Call errHandler(SRCFILE & "SaveData", s)
    End If
End Function


Private Sub gData_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim t As AssemblyTypes
    
    
    If gData.Col <> Col Then gData.Col = Col
    If gData.Row <> Row Then gData.Row = Row
    
    If ReadOnly Or gData.Cell(flexcpForeColor, Row, Col) = DisabledText Then
        gData.AutoSearch = flexSearchFromCursor
        Cancel = True
        Exit Sub
    End If
    gData.AutoSearch = flexSearchNone
    
    If Row = 1 Then
        gData.ComboList = ""
        Cancel = gData.ColDataType(Col) = flexDTBoolean
    Else
        With gData
            .ComboList = ""
            
            t = .ValueMatrix(Row, .ColIndex("AssemblyType"))
            
            Select Case .ColKey(Col)
                Case "Community"
                    .ComboList = "|..."
                    
                    
                Case "CommunityDesc"
                    .ComboList = "..."
                    
                    
                Case "CommunityPhase"
                    .ComboList = "|..."
                
                
                Case "Location"
                    Cancel = t = atModel
                    .ComboList = "|..."
                    .EditMaxLength = 50
                
                Case "Color"
                    .EditMaxLength = 30
                
                
                Case "AssemblyUOM"
                    .ComboList = "|..."
                    .EditMaxLength = 10
                    
                Case "Comments", "Notes"
                    .ComboList = "|..."
                    .EditMaxLength = 4000
                    
                Case "Description"
                    .ComboList = "|..."
                    .EditMaxLength = 200
                    
                Case "CategoryDesc"
                    Cancel = t = atModel
                    .ComboList = "..."
                   
                Case "Category"
                    Cancel = t = atModel
                    .ComboList = "|..."
                
                Case "COPretax", "COTotal", "COMargin", "COMarkup"
                    Cancel = t = atModel
                    
                Case "IncentiveRetail"
                    Cancel = t <> atModel Or Trim(HFApp.Options(IncentiveItem)) = ""
                    
                Case "LandCost"
                    Cancel = t <> atModel Or Trim(HFApp.Options(LandItem)) = ""
                
                               
                Case "ConstCutOff"
                    Cancel = t = atModel
                    .ComboList = "|..."
                    
                Case "JCExtra"
                    Cancel = t = atModel
                    .EditMaxLength = 10
                    
                Case "FloorArea"
                    Cancel = t <> atModel
                    
                    
                Case "Style"
                    Cancel = t <> atModel
                    
                    If HFApp.Options(SalesSystem) = SalesSystems.asBuilder1440 Then
                        .ComboList = "Single-Family|Condo|Townhouse"
                    Else
                        .ComboList = "|..."
                    End If
                    .EditMaxLength = 30
                
                Case "Bedrooms", "Bathrooms"
                    Cancel = t <> atModel
                    .EditMaxLength = 30
                    
                
                Case "Profit", "Margin", "Markup", "RoundTo", "IncludeTax"
                    
                Case "InSpec", "Pretax2", "Margin2", "Markup2", "InSpec2", _
                     "Pretax3", "Margin3", "Markup3", "InSpec3", _
                     "Pretax4", "Margin4", "Markup4", "InSpec4", _
                     "Pretax5", "Margin5", "Markup5", "InSpec5", _
                     "Pretax6", "Margin6", "Markup6", "InSpec6", _
                     "Pretax7", "Margin7", "Markup7", "InSpec7", _
                     "Pretax8", "Margin8", "Markup8", "InSpec8", _
                     "Pretax9", "Margin9", "Markup9", "InSpec9", _
                     "Pretax10", "Margin10", "Markup10", "InSpec10"
                    Cancel = t <> atDesignCenter
                    
                Case "Pretax", "Total", "IncludedOption", "Qty", "Inactive"
                Case "COPretax", "COTotal"
                    Cancel = t <> atModel
                
                Case Else
                    Cancel = True
           End Select
        End With
    End If
End Sub


Private Sub gData_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim s As String
    Dim l As String
    Dim i As Long
    Dim rs As Recordset
    
    With gData
        .ComboList = ""
        Select Case .ColKey(Col)
            Case "Location"
                s = .Text
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Room Location", "select Area,Description from tblAreas", s) Then
                    For i = .Row To .RowSel
                        If .RowHidden(i) = False Then
                            .Cell(flexcpText, i, .ColIndex("Location")) = FPickList.SelectedItem("Area")
                        End If
                    Next
                End If
                
            Case "Community", "CommunityDesc"
                'is global available?
                s = ""
                s = s & "select * " & vbCrLf
                s = s & "  from tbldbassemblymaster " & vbCrLf
                s = s & " where isnull(inactive,0)=0 and isnull(community,'')='' " & vbCrLf
                s = s & "   and assembly=" & DbQuote(Str, .TextMatrix(.Row, .ColIndex("Assembly"))) & vbCrLf
                s = s & "   and model=" & DbQuote(Str, .TextMatrix(.Row, .ColIndex("Model"))) & vbCrLf
                s = s & "   and optionid=" & DbQuote(Str, .TextMatrix(.Row, .ColIndex("OptionID"))) & vbCrLf
                s = s & "   and divisionid = " & HFApp.DivisionID
                Set rs = HFApp.SqlExec(s, dbHomefront)
                If rs.EOF Then
                    'no global so list available communities
                    s = ""
                    s = s & "select l.area community, l.description" & vbCrLf
                    s = s & "  from tbldbassemblymaster a" & vbCrLf
                    s = s & "       join tbllocality l on (l.area=a.community)" & vbCrLf
                    If HFApp.DivisionID <> "" Then
                        s = s & " left outer join DivisionCommunities d on d.Community = l.Area" & vbCrLf
                    End If
                    s = s & " where isnull(a.inactive,0)=0 and isnull(l.Inactive,0)=0" & vbCrLf
                    s = s & "   and a.assembly=" & DbQuote(Str, .TextMatrix(.Row, .ColIndex("Assembly"))) & vbCrLf
                    s = s & "   and a.model=" & DbQuote(Str, .TextMatrix(.Row, .ColIndex("Model"))) & vbCrLf
                    s = s & "   and a.optionid=" & DbQuote(Str, .TextMatrix(.Row, .ColIndex("OptionID"))) & vbCrLf
                    s = s & "   and a.DivisionID = " & HFApp.DivisionID & vbCrLf
                    If HFApp.DivisionID <> "" Then
                        s = s & " and (d.DivisionID is null or d.DivisionID = " & HFApp.DivisionID & ")"
                    End If
                    
                Else
                    'all active communities
                    s = ""
                    s = s & "SELECT area Community,Description " & vbCrLf
                    s = s & "  FROM tbllocality " & vbCrLf
                    If HFApp.DivisionID <> "" Then
                        s = s & " left outer join DivisionCommunities d on d.Community = tbllocality.Area" & vbCrLf
                    End If
                    s = s & " WHERE isnull(Inactive,0)=0" & vbCrLf
                    If HFApp.DivisionID <> "" Then
                        s = s & " and (d.DivisionID is null or d.DivisionID = " & HFApp.DivisionID & ")"
                    End If
                End If
                
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Community", s, .TextMatrix(.Row, .ColIndex("Community")), , , , IIf(.ColKey(Col) = "CommunityDesc", "Community", "")) Then
                    For i = .Row To .RowSel
                        If .RowHidden(i) = False Then
                            .Cell(flexcpText, i, .ColIndex("Community")) = FPickList.SelectedItem("Community")
                            .Cell(flexcpText, i, .ColIndex("CommunityDesc")) = FPickList.SelectedItem("Description")
                            .Cell(flexcpText, i, .ColIndex("CommunityPhase")) = ""
                        End If
                    Next
'                    Call RefreshPublished(False, False)
                End If
            
            Case "CommunityPhase"
                s = .Text
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Category", "select CommunityPhase Phase,Description from communityphase where Community=" & DbQuote(Str, .TextMatrix(Row, .ColIndex("Community"))), s) Then
                    For i = .Row To .RowSel
                        If .RowHidden(i) = False Then
                            .Cell(flexcpText, i, .ColIndex("CommunityPhase")) = FPickList.SelectedItem("Phase")
                        End If
                    Next
                End If
        
            Case "Notes", "Comments"
                s = .Text
                If FComments.Edit(s, gData, , "Notes", 4000) Then
                    .Text = s
                    Call gData_AfterEdit(Row, .ColIndex("Notes"))
                End If
                
            Case "Description"
                s = .Text
                If FComments.Edit(s, gData, , "Description", 200) Then
                    .Text = s
                    Call gData_AfterEdit(Row, .ColIndex("Description"))
                End If
            
            Case "Category"
                s = .Text
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Category", "SELECT Category,Description,COMarkup FROM tblCategories", s) Then
                    For i = .Row To .RowSel
                        If .RowHidden(i) = False Then
                            .Cell(flexcpText, i, .ColIndex("Category")) = FPickList.SelectedItem("Category")
                            .Cell(flexcpText, i, .ColIndex("CategoryDesc")) = FPickList.SelectedItem("Description")
                            .Cell(flexcpText, i, .ColIndex("COMarkup")) = Val("" & FPickList.SelectedItem("COMarkup"))
                        End If
                    Next
                    Call CalcData
                End If
                
            Case "CategoryDesc"
                s = .Text
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Category", "SELECT Description,Category,COMarkup FROM tblCategories", s) Then
                    For i = .Row To .RowSel
                        If .RowHidden(i) = False Then
                            .Cell(flexcpText, i, .ColIndex("Category")) = FPickList.SelectedItem("Category")
                            .Cell(flexcpText, i, .ColIndex("CategoryDesc")) = FPickList.SelectedItem("Description")
                            .Cell(flexcpText, i, .ColIndex("COMarkup")) = Val("" & FPickList.SelectedItem("COMarkup"))
                        End If
                    Next
                    Call CalcData
                End If
                
            Case "ConstCutOff"
                s = .Text
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Status", "SELECT Job_Status Status,Description FROM tblJobStatus", s) Then
                    For i = .Row To .RowSel
                        If .RowHidden(i) = False Then
                            .Cell(flexcpText, i, .ColIndex("ConstCutOff")) = FPickList.SelectedItem("Status")
                        End If
                    Next
                    Call CalcData
                End If
                
            Case "Series"
                s = .Text
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Series", "SELECT Series,Description FROM tblSeries where DivisionID = " & HFApp.DivisionID, s) Then
                    For i = .Row To .RowSel
                        If .RowHidden(i) = False Then
                            .Cell(flexcpText, i, .ColIndex("Series")) = FPickList.SelectedItem("Series")
                        End If
                    Next
                    Call CalcData
                End If
                
            Case "AssemblyUOM"
                s = .Text
                If FPickList.Choose(HFApp.Databases(dbHomefront), "UOM", "SELECT DISTINCT ISNULL(AssemblyUOM,'') UOM FROM tblDBAssemblyMaster", s) Then
                    For i = .Row To .RowSel
                        If .RowHidden(i) = False Then
                            .Cell(flexcpText, i, .ColIndex("AssemblyUOM")) = FPickList.SelectedItem("UOM")
                        End If
                    Next
                End If
                
            Case "Style"
                s = .Text
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Style", l, s) Then
                    For i = .Row To .RowSel
                        If .RowHidden(i) = False Then
                            .Cell(flexcpText, i, .ColIndex("Style")) = FPickList.SelectedItem("Style")
                        End If
                    Next
                    Call CalcData
                End If
    
                
       End Select
       mDirty = True
    End With
End Sub
Private Sub gData_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    
    
    If Row = 1 Then
        Call ApplyFilter
        Exit Sub
    End If
    
    
    If Row > 1 Then
        Call CalcData
'        If gData.ColKey(Col) = "Community" Then Call RefreshPublished(False, False)
    End If
    
End Sub
Private Sub gData_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call gData_MYValidateEdit(Row, Col, Cancel, gData.EditText)
End Sub

Private Sub gData_MYValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean, EditText As String)
    Dim s As String
    Dim d As String
    Dim rs As Recordset
    Dim i As Long
    
    
    With gData
        If Row = 1 Then Exit Sub
        s = EditText
        Select Case .ColKey(Col)
            
            Case "Qty"
                s = Val(.EditText)
            
            Case "IncentiveRetail"
                s = Val(.EditText)
                For i = .Row To .RowSel
                If .RowHidden(i) = False Then
                    .TextMatrix(i, .ColIndex("IncentiveCost")) = Val(.EditText) * Val(txtIncentiveCostPercent) / 100
                End If
                Next
            
            Case "ConstCutOff"
                Set rs = HFApp.SqlExec("SELECT Job_Status Status FROM tblJobStatus WHERE Job_Status=" & DbQuote(Num, s))
                If rs.EOF Then
                    d = InputBox(vbCrLf & "Status not found. Do you want to add it?" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Description", App.ProductName, "Status " & s)
                    Cancel = d = ""
                    If Not Cancel Then Call HFApp.SqlExec("INSERT INTO tblJobStatus(Job_Status,Description) VALUES(" & DbQuote(Num, s) & "," & DbQuote(Str, d) & ")")
                Else
                    s = "" & rs(0)
                End If
                
            Case "Location"
                If Not ValidateField(gData, s, "Location not found", "SELECT area,description FROM tblAreas WHERE area=" & DbQuote(Str, s)) Then
                    Cancel = True
                End If
                
            Case "Community"
                .Row = Row
                If Not ValidateField(gData, s, "Community not found", "SELECT area,description CommunityDesc,'' CommunityPhase FROM tbllocality WHERE area=" & DbQuote(Str, s), "CommunityDesc", "CommunityPhase") Then
                    Cancel = True
                Else
                    'check that we have a global assembly or one in this community
                    d = ""
                    d = d & "select * " & vbCrLf
                    d = d & "  from tbldbassemblymaster a" & vbCrLf
                    d = d & " where isnull(a.inactive,0)=0 and (isnull(community,'')='' or isnull(community,'')=" & DbQuote(Str, .EditText) & ")" & vbCrLf
                    d = d & "   and assembly=" & DbQuote(Str, .TextMatrix(Row, .ColIndex("Assembly"))) & vbCrLf
                    d = d & "   and model=" & DbQuote(Str, .TextMatrix(Row, .ColIndex("Model"))) & vbCrLf
                    d = d & "   and optionid=" & DbQuote(Str, .TextMatrix(Row, .ColIndex("OptionID"))) & vbCrLf
                    Set rs = HFApp.SqlExec(d, dbHomefront)
                    If rs.EOF Then
                        Cancel = vbCancel = MsgBox("No global or community specific assembly could be found." & vbCrLf & vbCrLf & "Are you sure you want to continue?", vbInformation + vbOKCancel, App.ProductName)
                    End If
            
                End If
                
            Case "CommunityPhase":   Cancel = Not ValidateField(gData, s, "Phase not found", "SELECT CommunityPhase FROM CommunityPhase WHERE community=" & DbQuote(Str, .TextMatrix(Row, .ColIndex("Community"))) & " and communityphase=" & DbQuote(Str, s))
            Case "Series":           Cancel = Not ValidateField(gData, s, "Series not found", "SELECT Series FROM tblSeries WHERE DivisionID = " & HFApp.DivisionID & " and Series=" & DbQuote(Str, s))
            
            Case "Category"
                Cancel = Not ValidateField(gData, s, "Category not found", "SELECT Category,Description,COMarkup FROM tblCategories WHERE Category=" & DbQuote(Str, s), "CategoryDesc", "COMarkup")
                For i = .Row To .RowSel
                If .RowHidden(i) = False Then
                    .TextMatrix(i, .ColIndex("CategoryDesc")) = .TextMatrix(Row, .ColIndex("CategoryDesc"))
                    .TextMatrix(i, .ColIndex("COMarkup")) = .TextMatrix(Row, .ColIndex("COMarkup"))
                End If
                Next
            
        End Select
       
       
        For i = .Row To .RowSel
        If Cancel = False And .RowHidden(i) = False Then
            .TextMatrix(i, Col) = s
        End If
        Next
       
       
        .EditText = s
        mDirty = True
    End With

End Sub

Private Sub txtCostsEffectiveDate_GotFocus()
    SelectAll txtCostsEffectiveDate
End Sub

Private Sub txtDescription_Change()
    mDirty = True
End Sub

Private Sub txtDescription_GotFocus()
    SelectAll txtDescription
End Sub

Private Sub txtSalesEffectiveDate_GotFocus()
    SelectAll txtSalesEffectiveDate
End Sub

Private Property Get CostEffectiveDate() As Date
    CostEffectiveDate = mCostEffectiveDate
End Property
Private Property Let CostEffectiveDate(RHS As Date)
    mCostEffectiveDate = RHS
    txtCostsEffectiveDate.Text = format(mCostEffectiveDate, "mmmm d, yyyy")
End Property


Public Sub mnuSalesSheetSub_Click(Index As Integer)
On Error GoTo eh
    Dim r As Long
    Dim c As Long
    Dim startRow As Long
    Dim EndRow As Long
    Dim percent As Double
    Dim Dollars As Double
    
    Dim NewRow As Long
    Dim community As String
    
    Select Case Index
    
        Case mcPRICE_ADJUSTPRICES
            If FAdjustPrices.ShowForm(percent, Dollars) Then
                With gData
                
                    startRow = Min(.Row, .RowSel)
                    EndRow = Max(.Row, .RowSel)
                    If percent <> 0 Then
                        For r = startRow To EndRow
                            c = .ColIndex(IIf(.ValueMatrix(r, .ColIndex("IncludeTax")), "Total", "Pretax"))
                            .TextMatrix(r, c) = Val("" & .TextMatrix(r, c)) * (100 + percent) / 100
                            .Row = r
                            Call CalcData(, IIf(.ValueMatrix(r, .ColIndex("IncludeTax")), "Total", "Pretax"))
                        Next
                    End If
                    
                    If Dollars <> 0 Then
                        For r = startRow To EndRow
                            c = .ColIndex(IIf(.ValueMatrix(r, .ColIndex("IncludeTax")), "Total", "Pretax"))
                            .TextMatrix(r, c) = .ValueMatrix(r, c) + Dollars
                            .Row = r
                            Call CalcData(, IIf(.ValueMatrix(r, .ColIndex("IncludeTax")), "Total", "Pretax"))
                        Next
                    End If
                    
                End With
            End If
    
'        Case mcPRICE_COPY2COMMUNITY
'            If FPickList.Choose(HFApp.Databases(dbHomefront), "Community", "SELECT Description,Area Community FROM tblLocality", , , , FMain.SmallIcons.ListImages("area").Picture) Then
'                With gData
'                    For r = .Row To .RowSel
'                        If Not .RowHidden(r) Then
'                            newrow = .Rows
'                            .AddItem "", newrow
'                            For c = 0 To .Cols - 1
'                                .TextMatrix(newrow, c) = .TextMatrix(r, c)
'                            Next
'                            c = .ColIndex("Community")
'                            .TextMatrix(newrow, c) = FPickList.SelectedItem("Community")
'                            .TextMatrix(newrow, .ColIndex("CommunityDesc")) = FPickList.SelectedItem("Description")
'
'                            Call RefreshCosts(, newrow)
'                            Call CalcData(, "Cost")
'
'                        End If
'                    Next
'                End With
'            End If
                
        Case mcPRICE_REFRESHCOSTS
            Call RefreshCosts
'            Call CalcData(, "Cost")
            
        Case mcPRICE_REMOVEASSEMBLY
            Call gData_KeyDown(vbKeyDelete, vbCtrlMask)
        
    End Select
Exit Sub
eh: Call errHandler(SRCFILE & "mnuPricingWorksheetSub_Click")
End Sub

Private Function ValidateForm() As Boolean
On Error Resume Next
    Dim i      As Long
    Dim r      As Long
    Dim Item   As String
    Dim items  As New Collection
    Dim errors As String

    With gData
        For r = 2 To .Rows - 1
            If Not .RowHidden(r) Then
                i = i + 1
                
                'check for duplicates
                Err.Clear
                Item = "k" & .TextMatrix(r, .ColIndex("Community")) & Chr(1) & .TextMatrix(r, .ColIndex("CommunityPhase")) & Chr(1) & .TextMatrix(r, .ColIndex("Model")) & Chr(1) & .TextMatrix(r, .ColIndex("Series")) & Chr(1) & .TextMatrix(r, .ColIndex("OptionID"))
                items.Add "row " & i, Item
                If Err.Number <> 0 Then
                    errors = errors & vbBullet & " The assembly defined on row " & i & " is a duplicate of " & items(Item) & vbCrLf
                End If
                
            End If
            
        Next
    End With
    
    If errors = "" Then
        ValidateForm = True
    Else
        ValidateForm = False
        MsgBox "Unable to save this worksheet." & vbCrLf & vbCrLf & errors, vbExclamation, App.ProductName
    End If
    
End Function











Private Sub txtIncentiveCostPercent_GotFocus()
    SelectAll txtIncentiveCostPercent
End Sub

Private Sub txtIncentiveCostPercent_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim c As Boolean
    If KeyCode = vbKeyReturn Then
        Call txtIncentiveCostPercent_Validate(c)
    End If
End Sub

Private Sub txtIncentiveCostPercent_Validate(Cancel As Boolean)
    Dim Row As Long
    
    txtIncentiveCostPercent = Val(txtIncentiveCostPercent)
    With gData
        For Row = 2 To .Rows - 1
            'subtract current incentive costs
            .TextMatrix(Row, .ColIndex("Cost")) = .ValueMatrix(Row, .ColIndex("Cost")) - .ValueMatrix(Row, .ColIndex("IncentiveCost"))
            'calc new incentive cost
            .TextMatrix(Row, .ColIndex("IncentiveCost")) = .ValueMatrix(Row, .ColIndex("IncentiveRetail")) * Val(txtIncentiveCostPercent) / 100
            'add new incentive cost
            .TextMatrix(Row, .ColIndex("Cost")) = .ValueMatrix(Row, .ColIndex("Cost")) + .ValueMatrix(Row, .ColIndex("IncentiveCost"))
        Next
    End With
    Call CalcData(True, "IncentiveRetail")
    gData.SetFocus
    
End Sub





Private Sub RefreshPublished(MakeProposed As Boolean, DoAllRows As Boolean)
On Error Resume Next
    
    Dim startRow As Long
    Dim EndRow   As Long
    Dim r        As Long
    Dim s        As String
    Dim rs       As Recordset
    Dim i        As Long
    
    With gData
        Screen.MousePointer = vbHourglass
        
        If DoAllRows Then
            startRow = 2
            EndRow = .Rows - 1
        Else
            startRow = Min(.Row, .RowSel)
            EndRow = Max(.Row, .RowSel)
        End If
        
        For r = startRow To EndRow
            Select Case .ValueMatrix(r, .ColIndex("AssemblyType"))
            
                Case atDesignCenter
                    s = ""
                    s = s & "SELECT Item1  Pretax" & vbCrLf
                    s = s & "      ,NetTax1 Tax" & vbCrLf
                    s = s & "      ,null COPreTax" & vbCrLf
                    s = s & "      ,null COTax" & vbCrLf
                    s = s & "      ,0 Incentive" & vbCrLf
                    s = s & "      ,Margin1 Margin" & vbCrLf
                    s = s & "      ,Markup1 Markup" & vbCrLf
                    s = s & "      ,Item1-Item1_Cost Profit" & vbCrLf
                    s = s & "      ,Item2  Pretax2" & vbCrLf
                    s = s & "      ,NetTax2 Tax2" & vbCrLf
                    s = s & "      ,Margin2" & vbCrLf
                    s = s & "      ,Markup2" & vbCrLf
                    s = s & "      ,Item2-Item2_Cost Profit2" & vbCrLf
                    s = s & "      ,Item3  Pretax3" & vbCrLf
                    s = s & "      ,NetTax3 Tax3" & vbCrLf
                    s = s & "      ,Margin3" & vbCrLf
                    s = s & "      ,Markup3" & vbCrLf
                    s = s & "      ,Item3-Item3_Cost Profit3" & vbCrLf
                    s = s & "      ,Item4  Pretax4" & vbCrLf
                    s = s & "      ,NetTax4 Tax4" & vbCrLf
                    s = s & "      ,Margin4" & vbCrLf
                    s = s & "      ,Markup4" & vbCrLf
                    s = s & "      ,Item4-Item4_Cost Profit4" & vbCrLf
                    s = s & "      ,Item5  Pretax5" & vbCrLf
                    s = s & "      ,NetTax5 Tax5" & vbCrLf
                    s = s & "      ,Margin5" & vbCrLf
                    s = s & "      ,Markup5" & vbCrLf
                    s = s & "      ,Item5-Item5_Cost Profit5" & vbCrLf
                    s = s & "      ,Item6  Pretax6" & vbCrLf
                    s = s & "      ,NetTax6 Tax6" & vbCrLf
                    s = s & "      ,Margin6" & vbCrLf
                    s = s & "      ,Markup6" & vbCrLf
                    s = s & "      ,Item6-Item6_Cost Profit6" & vbCrLf
                    s = s & "      ,Item7  Pretax7" & vbCrLf
                    s = s & "      ,NetTax7 Tax7" & vbCrLf
                    s = s & "      ,Margin7" & vbCrLf
                    s = s & "      ,Markup7" & vbCrLf
                    s = s & "      ,Item7-Item7_Cost Profit7" & vbCrLf
                    s = s & "      ,Item8  Pretax8" & vbCrLf
                    s = s & "      ,NetTax8 Tax8" & vbCrLf
                    s = s & "      ,Margin8" & vbCrLf
                    s = s & "      ,Markup8" & vbCrLf
                    s = s & "      ,Item8-Item8_Cost Profit8" & vbCrLf
                    s = s & "      ,Item9  Pretax9" & vbCrLf
                    s = s & "      ,NetTax9 Tax9" & vbCrLf
                    s = s & "      ,Margin9" & vbCrLf
                    s = s & "      ,Markup9" & vbCrLf
                    s = s & "      ,TotalAmount9-Item9_Cost Profit9" & vbCrLf
                    s = s & "      ,Item10  Pretax10" & vbCrLf
                    s = s & "      ,NetTax10 Tax10" & vbCrLf
                    s = s & "      ,Margin10" & vbCrLf
                    s = s & "      ,Markup10" & vbCrLf
                    s = s & "      ,Item10-Item10_Cost Profit10" & vbCrLf
                    s = s & "  FROM tblDCOptions" & vbCrLf
                    s = s & " WHERE ISNULL(Community,'')=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Community"))) & vbCrLf
                    s = s & "   AND ISNULL(CommunityPhase,'')=" & DbQuote(Str, .TextMatrix(r, .ColIndex("CommunityPhase"))) & vbCrLf
                    s = s & "   AND ISNULL(Opt,'')=" & DbQuote(Str, .TextMatrix(r, .ColIndex("OptionID"))) & vbCrLf
                    s = s & "   AND DivisionID = " & HFApp.DivisionID
            
                Case atModel
                    s = ""
                    s = s & "SELECT Base_House Pretax" & vbCrLf
                    s = s & "      ,NetTax Tax" & vbCrLf
                    s = s & "      ,null COPreTax" & vbCrLf
                    s = s & "      ,null COTax" & vbCrLf
                    s = s & "      ,IncentiveRetail Incentive" & vbCrLf
                    s = s & "      ,Margin" & vbCrLf
                    s = s & "      ,Markup" & vbCrLf
                    s = s & "      ,Base_House-Cost_Amount Profit" & vbCrLf
                    s = s & "  FROM tblModels" & vbCrLf
                    s = s & " WHERE ISNULL(Area,'')=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Community"))) & vbCrLf
                    s = s & "   AND ISNULL(CommunityPhase,'')=" & DbQuote(Str, .TextMatrix(r, .ColIndex("CommunityPhase"))) & vbCrLf
                    s = s & "   AND ISNULL(Model,'')=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Model"))) & vbCrLf
                    s = s & "   AND ISNULL(Series,'')=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Series"))) & vbCrLf
                    s = s & "   AND DivisionID = " & HFApp.DivisionID
                    
                Case atoption
                    s = ""
                    s = s & "SELECT Price Pretax" & vbCrLf
                    s = s & "      ,NetTax Tax" & vbCrLf
                    s = s & "      ,co_Price COPreTax" & vbCrLf
                    s = s & "      ,netcotax COTax" & vbCrLf
                    s = s & "      ,0 Incentive" & vbCrLf
                    s = s & "      ,Margin" & vbCrLf
                    s = s & "      ,Markup" & vbCrLf
                    s = s & "      ,Price-Cost_Amount Profit" & vbCrLf
                    s = s & "  FROM tblOptions" & vbCrLf
                    s = s & " WHERE ISNULL(Area,'')=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Community"))) & vbCrLf
                    s = s & "   AND ISNULL(CommunityPhase,'')=" & DbQuote(Str, .TextMatrix(r, .ColIndex("CommunityPhase"))) & vbCrLf
                    s = s & "   AND ISNULL(Model,'')=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Model"))) & vbCrLf
                    s = s & "   AND ISNULL(Series,'')=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Series"))) & vbCrLf
                    s = s & "   AND ISNULL(Opt,'')=" & DbQuote(Str, .TextMatrix(r, .ColIndex("OptionID"))) & vbCrLf
                    s = s & "   AND DivisionID = " & HFApp.DivisionID
                
                Case atGlobal
                    s = ""
                    s = s & "SELECT Price Pretax" & vbCrLf
                    s = s & "      ,NetTax Tax" & vbCrLf
                    s = s & "      ,co_Price COPreTax" & vbCrLf
                    s = s & "      ,netcotax COTax" & vbCrLf
                    s = s & "      ,0 Incentive" & vbCrLf
                    s = s & "      ,Margin" & vbCrLf
                    s = s & "      ,Markup" & vbCrLf
                    s = s & "      ,Price-Cost_Amount Profit" & vbCrLf
                    s = s & "  FROM tblGlobalOptions" & vbCrLf
                    s = s & " WHERE ISNULL(Community,'')=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Community"))) & vbCrLf
                    s = s & "   AND ISNULL(CommunityPhase,'')=" & DbQuote(Str, .TextMatrix(r, .ColIndex("CommunityPhase"))) & vbCrLf
                    s = s & "   AND ISNULL(Opt,'')=" & DbQuote(Str, .TextMatrix(r, .ColIndex("OptionID"))) & vbCrLf
                    s = s & "   AND DivisionID = " & HFApp.DivisionID
                    
            End Select
            
            Set rs = HFApp.SqlExec(s)
            If rs.EOF Then
                .TextMatrix(r, .ColIndex("PublishedPretax")) = ""
                .TextMatrix(r, .ColIndex("PublishedTax")) = ""
                .TextMatrix(r, .ColIndex("PublishedPrice")) = ""
                .TextMatrix(r, .ColIndex("PublishedCOPrice")) = ""
                .TextMatrix(r, .ColIndex("PublishedIncentive")) = ""
                .TextMatrix(r, .ColIndex("PublishedMarkup")) = ""
                .TextMatrix(r, .ColIndex("PublishedMargin")) = ""
                .TextMatrix(r, .ColIndex("PublishedProfit")) = ""
                For i = 2 To 10
                    .TextMatrix(r, .ColIndex("PublishedPretax" & i)) = ""
                    .TextMatrix(r, .ColIndex("PublishedTax" & i)) = ""
                    .TextMatrix(r, .ColIndex("PublishedPrice" & i)) = ""
                    .TextMatrix(r, .ColIndex("PublishedMarkup" & i)) = ""
                    .TextMatrix(r, .ColIndex("PublishedMargin" & i)) = ""
                    .TextMatrix(r, .ColIndex("PublishedProfit" & i)) = ""
                Next
            Else
                .Cell(flexcpChecked, r, .ColIndex("ReadyToPublish")) = flexUnchecked
                .TextMatrix(r, .ColIndex("PublishedPretax")) = Val("" & rs("Pretax"))
                .TextMatrix(r, .ColIndex("PublishedTax")) = Val("" & rs("Tax"))
                .TextMatrix(r, .ColIndex("PublishedPrice")) = Val("" & rs("Pretax")) + Val("" & rs("Tax"))
                .TextMatrix(r, .ColIndex("PublishedCOPrice")) = Val("" & rs("COPretax")) + Val("" & rs("COTax"))
                
                .TextMatrix(r, .ColIndex("PublishedIncentive")) = Val("" & rs("Incentive"))
                .TextMatrix(r, .ColIndex("PublishedMarkup")) = Val("" & rs("Markup"))
                .TextMatrix(r, .ColIndex("PublishedMargin")) = Val("" & rs("Margin"))
                .TextMatrix(r, .ColIndex("PublishedProfit")) = Val("" & rs("Profit"))
                If MakeProposed Then .TextMatrix(r, .ColIndex("Pretax")) = .TextMatrix(r, .ColIndex("PublishedPretax"))
                
                On Error Resume Next
                For i = 2 To 10
                    .TextMatrix(r, .ColIndex("PublishedPretax" & i)) = Val("" & rs("Pretax" & i))
                    .TextMatrix(r, .ColIndex("PublishedTax" & i)) = Val("" & rs("Tax" & i))
                    .TextMatrix(r, .ColIndex("PublishedPrice" & i)) = Val("" & rs("Pretax" & i)) + Val("" & rs("Tax" & i))
                    .TextMatrix(r, .ColIndex("PublishedMarkup" & i)) = Val("" & rs("Markup" & i))
                    .TextMatrix(r, .ColIndex("PublishedMargin" & i)) = Val("" & rs("Margin" & i))
                    .TextMatrix(r, .ColIndex("PublishedProfit" & i)) = Val("" & rs("Profit" & i))
                Next
                On Error GoTo 0
            End If
        
        Next
        
        If MakeProposed Then Call CalcData(DoAllRows, "Pretax")
        
        Screen.MousePointer = vbDefault
        
    End With
End Sub





Private Sub Form_Load()
    Dim s  As String
    Dim i As Long
    
    Me.Tag = mWorksheet
    Me.Caption = mWorksheetView & " Worksheet (number " & mWorksheet & ")"
    
    If mWorksheet <> 0 Then Call HFApp.LockRecord("tblSalesSheetMaster", mWorksheet)
    MaxVendorPricing = HFApp.SqlExec("select top 1 isnull(MaxVendorPricing,0) from system_setup where id=" & HFApp.DivisionID)(0)
    
    Call SetToolbarIcons(Toolbar, FMain.LargeIcons)
    Call IniGetForm(Me, , mWorksheetView)
    Call LoadGridLayouts
    
    
    
    
    Toolbar.Buttons("RePrice").Visible = True ' mWorksheetView = "Estimating" Or Not HFApp.Options(MarketingWorkSheetAutoRefreshCosts)
    Toolbar.Buttons("AssemblyCosts").Visible = mWorksheetView <> "Marketing"
    
    Toolbar.Buttons("s2").Visible = Toolbar.Buttons("RePrice").Visible
    Toolbar.Buttons("Add").Visible = mWorksheetView = "Estimating"
    Toolbar.Buttons("Delete").Visible = mWorksheetView = "Estimating"
    Toolbar.Buttons("s1").Visible = mWorksheetView = "Estimating"
    Toolbar.Buttons("SaveAs").Visible = mWorksheetView = "Estimating"
    
    Toolbar.Buttons("AssemblyCosts").Visible = HFApp.Options(RecordSalesSheetCosts)
    
    
    lblSalesEffectiveDate.Visible = mWorksheetView = "Estimating"
    txtSalesEffectiveDate.Visible = mWorksheetView = "Estimating"
    lblIncentiveCostPercent.Visible = mWorksheetView = "Estimating"
    txtIncentiveCostPercent.Visible = mWorksheetView = "Estimating"
    
    Call LoadSeries
    
    With gData
        .Redraw = flexRDNone
        
        .TextMatrix(0, .ColIndex("COMargin")) = ""
        .ColHidden(.ColIndex("COMargin")) = True
                
        Call LoadLayout
            
        'ensure showable spec columns have column labels so they can be selected
        For i = 2 To MaxSpec - 1
            
             If .TextMatrix(0, .ColIndex("Pretax" & i)) = "" Then .TextMatrix(0, .ColIndex("Pretax" & i)) = "Pretax"
             If .TextMatrix(0, .ColIndex("Tax" & i)) = "" Then .TextMatrix(0, .ColIndex("Tax" & i)) = "Tax"
             If .TextMatrix(0, .ColIndex("Markup" & i)) = "" Then .TextMatrix(0, .ColIndex("Markup" & i)) = "Markup"
             If .TextMatrix(0, .ColIndex("Margin" & i)) = "" Then .TextMatrix(0, .ColIndex("Margin" & i)) = "Margin"
             If .TextMatrix(0, .ColIndex("Profit" & i)) = "" Then .TextMatrix(0, .ColIndex("Profit" & i)) = "Profit"
             If .TextMatrix(0, .ColIndex("Total" & i)) = "" Then .TextMatrix(0, .ColIndex("Total" & i)) = "Total"
             
             If .TextMatrix(0, .ColIndex("PublishedPretax" & i)) = "" Then .TextMatrix(0, .ColIndex("PublishedPretax" & i)) = "PublishedPretax"
             If .TextMatrix(0, .ColIndex("PublishedTax" & i)) = "" Then .TextMatrix(0, .ColIndex("PublishedTax" & i)) = "PublishedTax"
             If .TextMatrix(0, .ColIndex("PublishedPrice" & i)) = "" Then .TextMatrix(0, .ColIndex("PublishedPrice" & i)) = "PublishedPrice"
             If .TextMatrix(0, .ColIndex("PublishedProfit" & i)) = "" Then .TextMatrix(0, .ColIndex("PublishedProfit" & i)) = "PublishedProfit"
             If .TextMatrix(0, .ColIndex("PublishedMarkup" & i)) = "" Then .TextMatrix(0, .ColIndex("PublishedMarkup" & i)) = "PublishedMarkup"
             If .TextMatrix(0, .ColIndex("PublishedMargin" & i)) = "" Then .TextMatrix(0, .ColIndex("PublishedMargin" & i)) = "PublishedMargin"
             
        Next
        
        'hide invisible columns
        s = VisibleColumns
        For i = 0 To .Cols - 1
            If Not ItemInList(.ColKey(i), s) Then
                .ColHidden(i) = True
                .TextMatrix(0, i) = ""
            End If
        Next
        
        .Redraw = flexRDBuffered
    End With
    
    Call LoadCostTypes(cboCostBasis)
    Me.Show
    Call LoadData
    
    If mWorksheetView = "Marketing" And HFApp.Options(MarketingWorkSheetAutoRefreshCosts) And Not ReadOnly Then
        Call RefreshCosts(True)
        Call CalcData(True, "Cost")
        mDirty = True
    End If
    
    Call Form_Resize
End Sub

Private Sub SaveVendorCost(community As String, CommunityPhase As String, Model As String, Assembly As String, Phase As String, Item As String, rate As Double)
On Error GoTo eh
    Dim s As String
    Dim Vendor As String
    
    On Error Resume Next
    Vendor = HFApp.SqlExec("select dbo.purch_getcommunityvendor(" & DbQuote(Str, community) & ",(select poindex from tblPhaseItem where DivisionID = " & HFApp.DivisionID & " and phase=" & DbQuote(Str, Phase) & " and item=" & DbQuote(Str, Item) & ")," & HFApp.DivisionID & ")", dbHomefront)(0)
    On Error GoTo eh
    If Vendor = "" Then Exit Sub
    
    s = ""
    s = s & "INSERT INTO tblVendorCost(DivisionID,community,communityphase,Model,assembly,phase,item,vendor,current_cost,next_cost1,next_cost2,last_cost1,last_cost2,last_cost3,forecast1,forecast2,forecast3,forecast4,forecast5,forecast6,forecast7,forecast8,forecast9,forecast10,forecast11,forecast12)" & vbCrLf
    s = s & "VALUES(" & HFApp.DivisionID & "," & DbQuote(Str, community) & vbCrLf
    s = s & "      ," & DbQuote(Str, CommunityPhase) & vbCrLf
    s = s & "      ," & DbQuote(Str, Model) & vbCrLf
    s = s & "      ," & DbQuote(Str, Assembly) & vbCrLf
    s = s & "      ," & DbQuote(Str, Phase) & vbCrLf
    s = s & "      ," & DbQuote(Str, Item) & vbCrLf
    s = s & "      ," & DbQuote(Str, Vendor) & vbCrLf
    s = s & "      ," & DbQuote(Num, rate) & vbCrLf
    s = s & "      ," & DbQuote(Num, rate) & vbCrLf
    s = s & "      ," & DbQuote(Num, rate) & vbCrLf
    s = s & "      ," & DbQuote(Num, rate) & vbCrLf
    s = s & "      ," & DbQuote(Num, rate) & vbCrLf
    s = s & "      ," & DbQuote(Num, rate) & vbCrLf
    s = s & "      ," & DbQuote(Num, rate) & vbCrLf
    s = s & "      ," & DbQuote(Num, rate) & vbCrLf
    s = s & "      ," & DbQuote(Num, rate) & vbCrLf
    s = s & "      ," & DbQuote(Num, rate) & vbCrLf
    s = s & "      ," & DbQuote(Num, rate) & vbCrLf
    s = s & "      ," & DbQuote(Num, rate) & vbCrLf
    s = s & "      ," & DbQuote(Num, rate) & vbCrLf
    s = s & "      ," & DbQuote(Num, rate) & vbCrLf
    s = s & "      ," & DbQuote(Num, rate) & vbCrLf
    s = s & "      ," & DbQuote(Num, rate) & vbCrLf
    s = s & "      ," & DbQuote(Num, rate) & vbCrLf
    s = s & "      ," & DbQuote(Num, rate) & ")"
    
    HFApp.SqlExec s
    
    s = ""
    s = s & "UPDATE tblVendorCost" & vbCrLf
    s = s & "   SET " & Choose(cboCostBasis.ListIndex + 1, "current_cost", "next_cost1", "next_cost2", "current_cost", "current_cost", "current_cost", "forecast1", "forecast2", "forecast3", "forecast4", "forecast5", "forecast6", "forecast7", "forecast8", "forecast9", "forecast10", "forecast11", "forecast12")
    s = s & "=" & DbQuote(Num, rate) & vbCrLf
    s = s & "WHERE Community=" & DbQuote(Str, community) & vbCrLf
    s = s & "  AND CommunityPhase=" & DbQuote(Str, CommunityPhase) & vbCrLf
    s = s & "  AND Model=" & DbQuote(Str, Model) & vbCrLf
    s = s & "  AND Assembly=" & DbQuote(Str, Assembly) & vbCrLf
    s = s & "  AND Phase=" & DbQuote(Str, Phase) & vbCrLf
    s = s & "  AND Item=" & DbQuote(Str, Item) & vbCrLf
    s = s & "  AND Vendor=" & DbQuote(Str, Vendor) & vbCrLf
    s = s & "  AND DivisionID = " & HFApp.DivisionID
    HFApp.SqlExec s
    
Exit Sub
eh: If InStr(1, Err.Description, "duplicate", vbTextCompare) Then
        Resume Next
    Else
        Call errHandler(SRCFILE & "SaveVendorCost", s)
    End If
End Sub


Private Sub LoadGridLayouts()
On Error Resume Next
    Dim i As Long
    Dim s As String
    cboLayouts.Clear
    s = IniGet(AppIni, "PricingWorksheet", mWorksheetView & mWorksheetType & "Layouts", "General")
    For i = 1 To Parse(s, , Chr(1))
        If Trim(Parse(s, i, Chr(1))) <> "" Then
            Call cboLayouts.AddItem(Parse(s, i, Chr(1)))
        End If
    Next
    cboLayouts.ListIndex = Val(IniGet(AppIni, "PricingWorksheet", mWorksheetView & mWorksheetType & "CurrentLayout", ""))
    If cboLayouts.ListIndex < 0 Then cboLayouts.ListIndex = 0
End Sub
    
Private Sub SaveGridLayouts()
On Error Resume Next
    Dim i As Long
    Dim s As String
    s = ""
    For i = 0 To cboLayouts.ListCount - 1
        s = s & Chr(1) & cboLayouts.List(i)
    Next
    Call IniPut(AppIni, "PricingWorksheet", mWorksheetView & mWorksheetType & "Layouts", Mid(s, 2))
    Call IniPut(AppIni, "PricingWorksheet", mWorksheetView & mWorksheetType & "CurrentLayout", cboLayouts.ListIndex)
End Sub

Private Sub cmdLayouts_Click()
On Error Resume Next
    Dim lastindex As Long
    Dim i As Long
    Dim s As String
    s = ""
    For i = 0 To cboLayouts.ListCount - 1
        s = s & cboLayouts.List(i) & vbCrLf
    Next
    
    If FPricingWorksheetLayouts.ShowForm(s) Then
        lastindex = cboLayouts.ListIndex
        cboLayouts.Clear
        For i = 1 To Parse(s, , vbCrLf)
            If Trim(Parse(s, i, vbCrLf)) <> "" Then
                Call cboLayouts.AddItem(Trim(Parse(s, i, vbCrLf)))
            End If
        Next
        If cboLayouts.ListCount = 0 Then Call cboLayouts.AddItem("General")
        cboLayouts.ListIndex = IIf(lastindex > cboLayouts.ListCount Or lastindex < 0, 0, lastindex)
    End If
    
End Sub

Private Sub cboLayouts_Click()
Static oldlayout As String

    Screen.MousePointer = vbHourglass
    gData.Redraw = flexRDNone
    If oldlayout <> "" Then
        Call IniPutGrid(Me, gData, , mWorksheetView & oldlayout)
    End If
    oldlayout = cboLayouts.Text
    gData.TextMatrix(0, gData.ColIndex("COMargin")) = ""
    gData.ColHidden(gData.ColIndex("COMargin")) = True
    Call LoadLayout
    gData.Redraw = flexRDBuffered
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub LoadLayout()
    
    Call IniGetGrid(Me, gData, , , mWorksheetView & cboLayouts.Text)
    If gData.TextMatrix(0, gData.ColIndex("Inactive")) = "" Then gData.TextMatrix(0, gData.ColIndex("Inactive")) = "Inactive"
    
    If HFApp.Options.value(includetax) Then
        If gData.TextMatrix(0, gData.ColIndex("IncludeTax")) = "" Then
            gData.TextMatrix(0, gData.ColIndex("IncludeTax")) = "Tax Included"
        End If
    Else
        gData.ColHidden(gData.ColIndex("IncludeTax")) = True
        gData.TextMatrix(0, gData.ColIndex("IncludeTax")) = ""
    End If

End Sub

Private Sub AddAssemblyOLD()
    Static iView As Long
    Dim s As String
    Dim i As Long
    Dim CommPhase      As String
    Dim community As String
    Dim Model As String
    Dim NewRow As Long
    
    If mWorksheetType = wsDesignCenter Then
        s = ""
        If HFApp.Options(DCOptionByAreaPhase) Then
            s = s & "    select c.area AreaID,c.description " & FMain.CD_Community & ",cp.communityphase Phase,isnull(k.description,a.Category) Category,a.Model modelid,a.OptionID,a.Assembly,a.Description,a.assemblytype,a.community SourceCommunity " & vbCrLf
            s = s & "      from tbldbassemblymaster a" & vbCrLf
            s = s & "      join tbllocality c on(a.community=c.area or isnull(a.community,'')='')" & vbCrLf
            If HFApp.DivisionID <> "" Then
                s = s & "  join DivisionCommunities d on d.Community = c.Area" & vbCrLf
            End If
            s = s & "      left outer join CommunityPhase cp on(c.area=cp.community)" & vbCrLf
            s = s & "      left outer join tblcategories k on a.category=k.category" & vbCrLf
            s = s & "     where isnull(a.IsBaseAssembly,0)=0 and isnull(c.inactive,0)=0 and isnull(a.inactive,0)=0  and a.assemblytype=4" & vbCrLf
            If HFApp.DivisionID <> "" Then
                s = s & " and (d.DivisionID is null or d.DivisionID = " & HFApp.DivisionID & ")"
            End If
            s = s & "    order by 1,2,3,4,5,6,7,8" & Chr(0)
        ElseIf HFApp.Options(DCOptionByArea) Then
            s = s & "    select c.area AreaID,c.description " & FMain.CD_Community & ",isnull(k.description,a.Category) Category,a.Model modelid,a.OptionID,a.Assembly,a.Description,a.assemblytype,a.community SourceCommunity " & vbCrLf
            s = s & "      from tbldbassemblymaster a" & vbCrLf
            s = s & "      join tbllocality c on(a.community=c.area or isnull(a.community,'')='')" & vbCrLf
            If HFApp.DivisionID <> "" Then
                s = s & "  join DivisionCommunities d on d.Community = c.Area" & vbCrLf
            End If
            s = s & "      left outer join tblcategories k on a.category=k.category" & vbCrLf
            s = s & "     where isnull(a.IsBaseAssembly,0)=0 and isnull(c.inactive,0)=0 and isnull(a.inactive,0)=0 and a.assemblytype=4" & vbCrLf
            If HFApp.DivisionID <> "" Then
                s = s & " and (d.DivisionID is null or d.DivisionID = " & HFApp.DivisionID & ")"
            End If
            s = s & "    order by 1,2,3,4,5,6,7,8" & Chr(0)
        Else
            s = s & "    select isnull(k.description,a.Category) Category,a.Model modelid,a.OptionID,a.Assembly,a.Description,a.assemblytype,a.community SourceCommunity,'' AreaID " & vbCrLf
            s = s & "      from tbldbassemblymaster a" & vbCrLf
            s = s & "      left outer join tblcategories k on a.category=k.category" & vbCrLf
            s = s & "     where isnull(a.IsBaseAssembly,0)=0 and isnull(a.inactive,0)=0 and a.assemblytype=4" & vbCrLf
            s = s & "       and isnull(a.community,'')=''" & vbCrLf
            s = s & "    order by 1,2,3,4,5,6,7" & Chr(0)
        End If
        If s <> "" Then s = Left(s, Len(s) - 1)
    Else
        s = ""
        s = s & "Models" & Chr(1) & vbCrLf
        If HFApp.Options(ModelsByArea_Phase) Then
            s = s & "    select c.area AreaID,c.description " & FMain.CD_Community & ",cp.communityphase Phase,a.Model,a.Model modelid,a.OptionID,a.Assembly,a.Description,s.Series,a.assemblytype,a.community SourceCommunity " & vbCrLf
            s = s & "      from tbldbassemblymaster a" & vbCrLf
            s = s & "      left outer join tbllocality c on(a.community=c.area or isnull(a.community,'')='')" & vbCrLf
            s = s & "  join DivisionCommunities d on d.Community = c.Area" & vbCrLf
            s = s & "      left outer join CommunityPhase cp on(c.area=cp.community and 'True'=" & DbQuote(Str, HFApp.Options(ModelsByArea_Phase)) & ")" & vbCrLf
            s = s & "      join tblseries s on ((isnull(a.series,'')='' or a.series=s.series) and s.DivisionID = " & HFApp.DivisionID & ")" & vbCrLf
            s = s & "     where isnull(a.IsBaseAssembly,0)=0 and isnull(c.inactive,0)=0 and isnull(a.inactive,0)=0 and a.assemblytype=0" & vbCrLf
            s = s & " and a.DivisionID = " & HFApp.DivisionID & vbCrLf
            s = s & " and (d.DivisionID is null or d.DivisionID = " & HFApp.DivisionID & ")"
            s = s & "    order by 1,2,3,4,5,6,7,8" & Chr(0)
        ElseIf HFApp.Options(ModelByArea) Then
            s = s & "    select c.area AreaID,c.description " & FMain.CD_Community & ",a.Model,a.Model modelid,a.OptionID,a.Assembly,a.Description,s.Series,a.assemblytype,a.community SourceCommunity " & vbCrLf
            s = s & "      from tbldbassemblymaster a" & vbCrLf
            s = s & "      left outer join tbllocality c on(a.community=c.area or isnull(a.community,'')='')" & vbCrLf
            s = s & "  join DivisionCommunities d on d.Community = c.Area" & vbCrLf
            s = s & "      join tblseries s on s.DivisionID = " & HFApp.DivisionID & " and (isnull(a.series,'')='' or a.series=s.series)" & vbCrLf
            s = s & "     where isnull(a.IsBaseAssembly,0)=0 and isnull(c.inactive,0)=0 and isnull(a.inactive,0)=0 and a.assemblytype=0" & vbCrLf
            s = s & " and a.DivisionID = " & HFApp.DivisionID & vbCrLf
            s = s & " and (d.DivisionID is null or d.DivisionID = " & HFApp.DivisionID & ")"
            s = s & "    order by 1,2,3,4,5,6,7,8" & Chr(0)
        Else
            s = s & "    select a.Model,a.Model modelid,a.OptionID,a.Assembly,a.Description,s.Series,a.assemblytype,a.community SourceCommunity,'' AreaID " & vbCrLf
            s = s & "      from tbldbassemblymaster a" & vbCrLf
            s = s & "      join tblseries s on s.DivisionID =" & HFApp.DivisionID & " and (isnull(a.series,'')='' or a.series=s.series)" & vbCrLf
            s = s & "     where isnull(a.IsBaseAssembly,0)=0 and isnull(a.inactive,0)=0 and a.assemblytype=0" & vbCrLf
            s = s & "       and isnull(a.community,'')=''" & vbCrLf
            s = s & " and a.DivisionID = " & HFApp.DivisionID & vbCrLf
            s = s & "    order by 1,2,3,4,5,6" & Chr(0)
        End If
        
        s = s & "Model Options" & Chr(1) & vbCrLf
        If HFApp.Options(OptionByAreaPhase) Then
            s = s & "    select c.area AreaID,c.description " & FMain.CD_Community & ",cp.communityphase Phase,isnull(k.Description,a.Category) Category,a.Model,a.Model modelid,dm.description ModelDescription,dm.Series,a.OptionID,a.OptionID [Option],a.Assembly,a.Description,a.assemblytype,a.community SourceCommunity" & vbCrLf
            s = s & "      from tbldbassemblymaster a " & vbCrLf
            s = s & "      left outer join tbllocality c on(a.community=c.area or isnull(a.community,'')='') " & vbCrLf
            s = s & "      left outer join CommunityPhase cp on(c.area=cp.community)" & vbCrLf
            s = s & "  join DivisionCommunities d on d.Community = c.Area" & vbCrLf
            s = s & "      join DistinctDivisionModelSeries dm on(a.DivisionID = dm.DivisionID and a.model=dm.model and (isnull(nullif(a.series,''),dm.series)=dm.series) )" & vbCrLf
            s = s & "      left outer join tblcategories k on a.category=k.category" & vbCrLf
            s = s & "     where isnull(a.IsBaseAssembly,0)=0 and isnull(c.inactive,0)=0 and isnull(dm.inactive,0)=0 and isnull(a.inactive,0)=0 and a.assemblytype=2" & vbCrLf
            s = s & " and a.DivisionID = " & HFApp.DivisionID & vbCrLf
            s = s & " and (d.DivisionID is null or d.DivisionID = " & HFApp.DivisionID & ")"
            s = s & "    order by 1,2,3,4,5,6,7,8" & Chr(0)
        ElseIf HFApp.Options(OptionByArea) Then
            ' MODEL OPTION VIEW CAN HAVE HUNDREDS OF THOUSANDS OF ITEMS. PICKLIST IS TOO SLOW SO CHANGE INITIAL QUERY TO BE A LIST OF MODELS,
            ' THEN THEY CAN PICK FROM A SHORT LIST OF OPTIONS
            If HFApp.Options(ModelsByArea_Phase) Then
                s = s & "select c.area AreaID,c.description " & FMain.CD_Community & ",cp.communityphase Phase,a.Model,a.description ModelDescription,cast(count(*) as varchar) + ' options available'" & vbCrLf
                s = s & "from tbldbassemblymaster a" & vbCrLf
                s = s & "left outer join tbllocality c on(a.community=c.area or isnull(a.community,'')='')" & vbCrLf
                s = s & "join DivisionCommunities d on d.Community = c.Area" & vbCrLf
                s = s & "left outer join CommunityPhase cp on(c.area=cp.community)" & vbCrLf
                s = s & "join tblseries s on ((isnull(a.series,'')='' or a.series=s.series) and s.DivisionID = d.divisionid)" & vbCrLf
                s = s & "where isnull(a.IsBaseAssembly,0)=0 and isnull(c.inactive,0)=0 and isnull(a.inactive,0)=0 and a.assemblytype=0" & vbCrLf
                s = s & "and a.DivisionID = " & HFApp.DivisionID & vbCrLf
                s = s & "group by c.area,c.description ,cp.communityphase,a.Model,a.description" & vbCrLf
                s = s & "order by 1,2,3,4,5,6" & Chr(0)
            ElseIf HFApp.Options(ModelByArea) Then
                s = s & "select c.area AreaID,c.description " & FMain.CD_Community & ",'' Phase,a.Model,a.description ModelDescription,cast(count(*) as varchar) + ' options available'" & vbCrLf
                s = s & "from tbldbassemblymaster a" & vbCrLf
                s = s & "left outer join tbllocality c on(a.community=c.area or isnull(a.community,'')='')" & vbCrLf
                s = s & "join DivisionCommunities d on d.Community = c.Area" & vbCrLf
                s = s & "join tblseries s on s.DivisionID = a.divisionid and (isnull(a.series,'')='' or a.series=s.series)" & vbCrLf
                s = s & "where isnull(a.IsBaseAssembly,0)=0 and isnull(c.inactive,0)=0 and isnull(a.inactive,0)=0 and a.assemblytype=0" & vbCrLf
                s = s & "and a.DivisionID = " & HFApp.DivisionID & vbCrLf
                s = s & "group by c.area,c.description,a.Model,a.description " & vbCrLf
                s = s & "order by 1,2,3,4,5,6" & Chr(0)
            Else
                s = s & "select '' AreaID,'' " & FMain.CD_Community & ",'' Phase,a.Model,a.description ModelDescription,cast(count(*) as varchar) + ' options available'" & vbCrLf
                s = s & "from tbldbassemblymaster a" & vbCrLf
                s = s & "join tblseries s on s.DivisionID =a.divisionid and (isnull(a.series,'')='' or a.series=s.series)" & vbCrLf
                s = s & "where isnull(a.IsBaseAssembly,0)=0 and isnull(a.inactive,0)=0 and a.assemblytype=0" & vbCrLf
                s = s & "and isnull(a.community,'')=''" & vbCrLf
                s = s & "and a.DivisionID = " & HFApp.DivisionID & vbCrLf
                s = s & "group by a.Model,a.Model,a.description" & vbCrLf
                s = s & "order by 1,2,3,4,5,6" & Chr(0)
            End If
        
        
        Else
            s = s & "    select isnull(k.description,a.Category) Category,a.Model,a.Model modelid,dm.Series,dm.description ModelDescription,a.OptionID,a.OptionID [Option],a.Assembly,a.Description,a.assemblytype,a.community SourceCommunity,'' AreaID" & vbCrLf
            s = s & "      from tbldbassemblymaster a " & vbCrLf
            s = s & "      join DistinctDivisionModelSeries dm on(a.DivisionID = dm.DivisionID and a.model=dm.model and (isnull(nullif(a.series,''),dm.series)=dm.series) )" & vbCrLf
            s = s & "      left outer join tblcategories k on a.category=k.category" & vbCrLf
            s = s & "     where isnull(a.IsBaseAssembly,0)=0 and isnull(dm.inactive,0)=0 and isnull(a.inactive,0)=0 and a.assemblytype=2" & vbCrLf
            s = s & "       and a.DivisionID = " & HFApp.DivisionID & vbCrLf
            s = s & "       and isnull(a.community,'')=''" & vbCrLf
            s = s & "    order by 1,2,3,4,5,6,7,8" & Chr(0)
        End If
                        
        s = s & "Global Options" & Chr(1) & vbCrLf
        If HFApp.Options(GlobalOptionByAreaPhase) Then
            s = s & "    select c.area AreaID,c.description " & FMain.CD_Community & ",cp.communityphase Phase, isnull(k.description,a.category) Category,a.OptionID,a.OptionID [Option],a.Assembly,a.Description,a.assemblytype,a.community SourceCommunity " & vbCrLf
            s = s & "      from tbldbassemblymaster a " & vbCrLf
            s = s & "           left outer join tbllocality c on(a.community=c.area or isnull(a.community,'')='') " & vbCrLf
            If HFApp.DivisionID <> "" Then
                s = s & "  join DivisionCommunities d on d.Community = c.Area" & vbCrLf
            End If
            s = s & "           left outer join CommunityPhase cp on(c.area=cp.community and 'True'=" & DbQuote(Str, HFApp.Options(GlobalOptionByAreaPhase)) & ")" & vbCrLf
            s = s & "      left outer join tblcategories k on a.category=k.category" & vbCrLf
            s = s & "     where isnull(a.IsBaseAssembly,0)=0 and isnull(c.inactive,0)=0 and isnull(a.inactive,0)=0 and a.assemblytype=3" & vbCrLf
            s = s & "       and a.DivisionID = " & HFApp.DivisionID & vbCrLf
            If HFApp.DivisionID <> "" Then
                s = s & " and (d.DivisionID is null or d.DivisionID = " & HFApp.DivisionID & ")"
            End If
            s = s & "    union all" & vbCrLf
            s = s & "    select c.area AreaID,c.description " & FMain.CD_Community & ",cp.communityphase Phase,isnull(k.description,a.OptionCategory) Category,a.OptionID,a.OptionID [Option],'' Assembly,a.Description,3 assemblytype,'' SourceCommunity " & vbCrLf
            s = s & "      from tblphaseitem a " & vbCrLf
            s = s & "           left outer join tbllocality c on(1=1) " & vbCrLf
            If HFApp.DivisionID <> "" Then
                s = s & "  join DivisionCommunities d on d.Community = c.Area" & vbCrLf
            End If
            s = s & "           left outer join CommunityPhase cp on(c.area=cp.community )" & vbCrLf
            s = s & "      left outer join tblcategories k on a.optioncategory=k.category" & vbCrLf
            s = s & "     where a.DivisionID = " & HFApp.DivisionID & " and isnull(c.inactive,0)=0 and isnull(a.optionid,'')<>''" & vbCrLf
            If HFApp.DivisionID <> "" Then
                s = s & " and (d.DivisionID is null or d.DivisionID = " & HFApp.DivisionID & ")"
            End If
            s = s & "    order by 1,2,3,4,5,6,7,8" & Chr(0)
        ElseIf HFApp.Options(GlobalOptionByArea) Then
            s = s & "    select c.area AreaID,c.description " & FMain.CD_Community & ", isnull(k.description,a.category) Category,a.OptionID,a.OptionID [Option],a.Assembly,a.Description,a.assemblytype,a.community SourceCommunity " & vbCrLf
            s = s & "      from tbldbassemblymaster a " & vbCrLf
            s = s & "           left outer join tbllocality c on(a.community=c.area or isnull(a.community,'')='') " & vbCrLf
            If HFApp.DivisionID <> "" Then
                s = s & "  join DivisionCommunities d on d.Community = c.Area" & vbCrLf
            End If
            s = s & "      left outer join tblcategories k on a.category=k.category" & vbCrLf
            s = s & "     where isnull(a.IsBaseAssembly,0)=0 and isnull(c.inactive,0)=0 and isnull(a.inactive,0)=0 and a.assemblytype=3" & vbCrLf
            s = s & " and a.DivisionID = " & HFApp.DivisionID & vbCrLf
            If HFApp.DivisionID <> "" Then
                s = s & " and (d.DivisionID is null or d.DivisionID = " & HFApp.DivisionID & ")"
            End If
            s = s & "    union all" & vbCrLf
            s = s & "    select c.area AreaID,c.description " & FMain.CD_Community & ",isnull(k.description,a.OptionCategory) Category,a.OptionID,a.OptionID [Option],'' Assembly,a.Description,3 assemblytype,'' SourceCommunity " & vbCrLf
            s = s & "      from tblphaseitem a " & vbCrLf
            s = s & "           left outer join tbllocality c on(1=1) " & vbCrLf
            If HFApp.DivisionID <> "" Then
                s = s & "  join DivisionCommunities d on d.Community = c.Area" & vbCrLf
            End If
            s = s & "           left outer join CommunityPhase cp on(c.area=cp.community )" & vbCrLf
            s = s & "      left outer join tblcategories k on a.optioncategory=k.category" & vbCrLf
            s = s & "     where a.DivisionID = " & HFApp.DivisionID & " and isnull(c.inactive,0)=0 and isnull(a.optionid,'')<>''" & vbCrLf
            If HFApp.DivisionID <> "" Then
                s = s & " and (d.DivisionID is null or d.DivisionID = " & HFApp.DivisionID & ")"
            End If
            s = s & "    order by 1,2,3,4,5,6,7,8" & Chr(0)
        Else
            s = s & "    select isnull(k.description,a.category) Category,a.OptionID,a.OptionID [Option],a.Assembly,a.Description,a.assemblytype,a.community SourceCommunity,'' AreaID " & vbCrLf
            s = s & "      from tbldbassemblymaster a " & vbCrLf
            s = s & "      left outer join tblcategories k on a.category=k.category" & vbCrLf
            s = s & "     where isnull(a.IsBaseAssembly,0)=0 and isnull(a.inactive,0)=0 and a.assemblytype=3" & vbCrLf
            s = s & "       and a.DivisionID = " & HFApp.DivisionID & vbCrLf
            s = s & "    union all" & vbCrLf
            s = s & "    select isnull(k.description,a.OptionCategory) Category,a.OptionID,a.OptionID [Option],'' Assembly,a.Description,3 assemblytype,'' SourceCommunity,'' AreaID " & vbCrLf
            s = s & "      from tblphaseitem a " & vbCrLf
            s = s & "      left outer join tblcategories k on a.optioncategory=k.category" & vbCrLf
            s = s & "     where a.DivisionID = " & HFApp.DivisionID & " and isnull(a.optionid,'')<>''" & vbCrLf
            s = s & "    order by 1,2,3,4,5,6,7" & Chr(0)
        End If
        If s <> "" Then s = Left(s, Len(s) - 1)
    
    End If
    If FPickList.Choose(HFApp.Databases(dbHomefront), "Assembly", s, , , , , "areaID,modelid,optionid,assemblytype,SourceCommunity", True, iView) Then
        
        'save this for next time
        iView = FPickList.SelectedView

        ' MODEL OPTION VIEW CAN HAVE MANY THOUSANDS OF ITEMS. PICKLIST IS TOO SLOW SO CHANGE INITIAL QUERY TO BE A LIST OF MODELS, THEN THEY CAN PICK FROM A SHORT LIST OF OPTIONS
        If mWorksheetType = wsgeneral And FPickList.SelectedView = 2 Then
            community = FPickList.SelectedItem("areaID")
            CommPhase = FPickList.SelectedItem("phase")
            
            s = ""
            For i = 1 To FPickList.SelectedItems
                s = s & "," & DbQuote(Str, FPickList.SelectedItem("model", i))
            Next
            Model = Mid(s, 2)
            
            If HFApp.Options(OptionByAreaPhase) Then
                'not yet implemented
            ElseIf HFApp.Options(OptionByArea) Then
                s = ""
                s = s & "select c.area AreaID,c.description " & FMain.CD_Community & ",isnull(k.description,a.Category) Category,a.Model,a.Model modelid,dm.description ModelDescription,dm.Series,a.OptionID,a.OptionID [Option],a.Assembly,a.Description,a.assemblytype,a.community SourceCommunity" & vbCrLf
                s = s & "from tbldbassemblymaster a " & vbCrLf
                s = s & "left outer join tbllocality c on(a.community=c.area or isnull(a.community,'')='') " & vbCrLf
                s = s & "join DivisionCommunities d on d.Community = c.Area" & vbCrLf
                s = s & "left outer join CommunityPhase cp on(c.area=cp.community) and isnull(cp.communityphase,'')=" & DbQuote(Str, CommPhase) & vbCrLf
                s = s & "join DistinctDivisionModelSeries dm on(a.DivisionID = dm.DivisionID and a.model=dm.model and (isnull(nullif(a.series,''),dm.series)=dm.series) )" & vbCrLf
                s = s & "left outer join tblcategories k on a.category=k.category" & vbCrLf
                s = s & "where isnull(c.inactive,0)=0 and isnull(dm.inactive,0)=0 and isnull(a.inactive,0)=0 and a.assemblytype=2" & vbCrLf
                s = s & "and a.DivisionID = " & HFApp.DivisionID & vbCrLf
                s = s & "and c.area=" & DbQuote(Str, community) & vbCrLf
                s = s & "and dm.model in(" & Model & ")" & vbCrLf
                s = s & "order by 1,2,3,4,5,6,7,8"
                
                
                If Not FPickList.Choose(HFApp.Databases(dbHomefront), "Assembly", s, , , , , "areaID,modelid,optionid,assemblytype,SourceCommunity," & FMain.CD_Community, True) Then
                    Exit Sub
                End If
            End If
        End If

        'remember where you started adding
        NewRow = gData.Rows

        'add new ones
        For i = 1 To FPickList.SelectedItems
            Call AddAssemblyToGridOLD(FPickList.SelectedItem("SourceCommunity", i), _
                                   FPickList.SelectedItem("areaID", i), _
                                   FPickList.SelectedItem("phase", i), _
                                   FPickList.SelectedItem("assembly", i), _
                                   FPickList.SelectedItem("modelid", i), _
                                   FPickList.SelectedItem("optionid", i), _
                                   FPickList.SelectedItem("series", i))
        Next
        
        Call SaveData(False)
        For i = NewRow To gData.Rows - 1
            gData.Row = i
            Call RefreshCosts(, gData.Row)
            Call CalcData(, "Cost", gData.Row)
        Next
        
    End If

End Sub

Private Sub AddAssembly()
    Static iView As Long
    
    Dim s As String
    Dim i As Long
    Dim a As Long
    Dim c As Long
    
    Dim NewRow As Long
    
    If mWorksheetType = wsDesignCenter Then
        s = ""
        If HFApp.Options(DCOptionByAreaPhase) Then
            s = s & "    select c.area AreaID,c.description " & FMain.CD_Community & ",cp.communityphase Phase,isnull(k.description,a.Category) Category,a.Model,a.OptionID [Option],a.Assembly,a.Description,a.assemblytype,a.community SourceCommunity " & vbCrLf
            s = s & "      from tbldbassemblymaster a" & vbCrLf
            s = s & "      join tbllocality c on(a.community=c.area or isnull(a.community,'')='')" & vbCrLf
            If HFApp.DivisionID <> "" Then
                s = s & "  join DivisionCommunities d on d.Community = c.Area" & vbCrLf
            End If
            s = s & "      left outer join CommunityPhase cp on(c.area=cp.community)" & vbCrLf
            s = s & "      left outer join tblcategories k on a.category=k.category" & vbCrLf
            s = s & "     where a.divisionid=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
            s = s & "       and isnull(a.IsBaseAssembly,0)=0 and isnull(c.inactive,0)=0 and isnull(a.inactive,0)=0  and a.assemblytype=4" & vbCrLf
            If HFApp.DivisionID <> "" Then
                s = s & " and (d.DivisionID is null or d.DivisionID = " & HFApp.DivisionID & ")"
            End If
            s = s & "    order by 1,2,3,4,5,6,7,8" & Chr(0)
        ElseIf HFApp.Options(DCOptionByArea) Then
            s = s & "    select c.area AreaID,c.description " & FMain.CD_Community & ",isnull(k.description,a.Category) Category,a.Model,a.OptionID [Option],a.Assembly,a.Description,a.assemblytype,a.community SourceCommunity " & vbCrLf
            s = s & "      from tbldbassemblymaster a" & vbCrLf
            s = s & "      join tbllocality c on(a.community=c.area or isnull(a.community,'')='')" & vbCrLf
            If HFApp.DivisionID <> "" Then
                s = s & "  join DivisionCommunities d on d.Community = c.Area" & vbCrLf
            End If
            s = s & "      left outer join tblcategories k on a.category=k.category" & vbCrLf
            s = s & "     where a.divisionid=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
            s = s & "       and isnull(a.IsBaseAssembly,0)=0 and isnull(c.inactive,0)=0 and isnull(a.inactive,0)=0 and a.assemblytype=4" & vbCrLf
            If HFApp.DivisionID <> "" Then
                s = s & " and (d.DivisionID is null or d.DivisionID = " & HFApp.DivisionID & ")"
            End If
            s = s & "    order by 1,2,3,4,5,6,7,8" & Chr(0)
        Else
            s = s & "    select isnull(k.description,a.Category) Category,a.Model,a.OptionID [Option],a.Assembly,a.Description,a.assemblytype,a.community SourceCommunity,'' AreaID " & vbCrLf
            s = s & "      from tbldbassemblymaster a" & vbCrLf
            s = s & "      left outer join tblcategories k on a.category=k.category" & vbCrLf
            s = s & "     where a.divisionid=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
            s = s & "       and isnull(a.IsBaseAssembly,0)=0 and isnull(a.inactive,0)=0 and a.assemblytype=4" & vbCrLf
            s = s & "       and isnull(a.community,'')=''" & vbCrLf
            s = s & "    order by 1,2,3,4,5,6,7" & Chr(0)
        End If
        If s <> "" Then s = Left(s, Len(s) - 1)
    Else
    
    
        s = ""
        s = s & "Models" & Chr(1) & vbCrLf
        s = s & "select distinct m.Model,s.Series,m.Assembly,m.Description" & vbCrLf
        s = s & "from tbldbassemblymaster m" & vbCrLf
        s = s & "left join tblseries s on s.DivisionID = " & HFApp.DivisionID & " and (isnull(m.series,'')='' or m.series=s.series)" & vbCrLf
        s = s & "where m.divisionid=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
        s = s & "and isnull(m.community,'')=''" & vbCrLf
        s = s & "and isnull(m.inactive,0)=0" & vbCrLf
        s = s & "and m.assemblytype=0" & vbCrLf
        s = s & "and isnull(m.isbaseassembly,0)=0" & vbCrLf
        s = s & "order by 1,2,3,4" & vbCrLf
        s = s & Chr(0) & vbCrLf
        
        s = s & "Community Models" & Chr(1) & vbCrLf
        s = s & "select distinct m.Model,s.Series,m.Assembly,m.Description,m.community SourceCommunity,c.description Community" & vbCrLf
        s = s & "from tbldbassemblymaster m" & vbCrLf
        s = s & "join tbllocality c on c.area=m.community" & vbCrLf
        s = s & "left join divisioncommunities d on m.community=d.community and m.divisionid=d.divisionid" & vbCrLf
        s = s & "left join tblseries s on s.DivisionID = " & HFApp.DivisionID & " and (isnull(m.series,'')='' or m.series=s.series)" & vbCrLf
        s = s & "where m.divisionid=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
        s = s & "and (d.community is not null or isnull(m.community,'')='')" & vbCrLf
        s = s & "and isnull(m.inactive,0)=0" & vbCrLf
        s = s & "and m.assemblytype=0" & vbCrLf
        s = s & "and isnull(m.isbaseassembly,0)=0" & vbCrLf
        s = s & "order by 1,2,3,4" & vbCrLf
        s = s & Chr(0) & vbCrLf
        
        s = s & "Model Options" & Chr(1) & vbCrLf
        s = s & "select distinct k.Description Category,m.Model,dm.description ModelDesc,s.Series,m.OptionID [Option],m.Assembly,m.Description" & vbCrLf
        s = s & "from tbldbassemblymaster m" & vbCrLf
        s = s & "left join distinctmodelsbydivision dm on m.model=dm.model and m.divisionid=dm.divisionid" & vbCrLf
        s = s & "left outer join tblcategories k on m.category=k.category" & vbCrLf
        s = s & "left join tblseries s on s.DivisionID = " & HFApp.DivisionID & " and (isnull(m.series,'')='' or m.series=s.series)" & vbCrLf
        s = s & "where m.divisionid=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
        s = s & "and isnull(m.community,'')=''" & vbCrLf
        s = s & "and isnull(m.inactive,0)=0" & vbCrLf
        s = s & "and m.assemblytype=2" & vbCrLf
        s = s & "and isnull(m.isbaseassembly,0)=0" & vbCrLf
        s = s & "order by 1,2,3,4" & vbCrLf
        s = s & Chr(0) & vbCrLf
        
        s = s & "Community Model Options" & Chr(1) & vbCrLf
        s = s & "select distinct k.Description Category,m.Model,dm.description ModelDesc,s.Series,m.OptionID [Option],m.Assembly,m.Description,m.community SourceCommunity,c.description Community" & vbCrLf
        s = s & "from tbldbassemblymaster m" & vbCrLf
        s = s & "left join distinctmodelsbydivision dm on m.model=dm.model and m.divisionid=dm.divisionid" & vbCrLf
        s = s & "join tbllocality c on c.area=m.community" & vbCrLf
        s = s & "left join divisioncommunities d on m.community=d.community and m.divisionid=d.divisionid" & vbCrLf
        s = s & "left outer join tblcategories k on m.category=k.category" & vbCrLf
        s = s & "left join tblseries s on s.DivisionID = " & HFApp.DivisionID & " and (isnull(m.series,'')='' or m.series=s.series)" & vbCrLf
        s = s & "where m.divisionid=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
        s = s & "and (d.community is not null or isnull(m.community,'')='')" & vbCrLf
        s = s & "and isnull(m.inactive,0)=0" & vbCrLf
        s = s & "and m.assemblytype=2" & vbCrLf
        s = s & "and isnull(m.isbaseassembly,0)=0" & vbCrLf
        s = s & "order by 1,2,3,4" & vbCrLf
        s = s & Chr(0) & vbCrLf
        
        s = s & "Global Options" & Chr(1) & vbCrLf
        s = s & "select distinct k.Description Category,m.OptionID [Option],m.Assembly,m.Description" & vbCrLf
        s = s & "from tbldbassemblymaster m" & vbCrLf
        s = s & "left join distinctmodelsbydivision dm on m.model=dm.model and m.divisionid=dm.divisionid" & vbCrLf
        s = s & "left outer join tblcategories k on m.category=k.category" & vbCrLf
        s = s & "where m.divisionid=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
        s = s & "and isnull(m.community,'')=''" & vbCrLf
        s = s & "and isnull(m.inactive,0)=0" & vbCrLf
        s = s & "and isnull(m.isbaseassembly,0)=0" & vbCrLf
        s = s & "and m.assemblytype=3" & vbCrLf
        s = s & "union" & vbCrLf
        s = s & "select distinct k.Description Category,a.OptionID [Option],'' Assembly,a.Description" & vbCrLf
        s = s & "from tblphaseitem a " & vbCrLf
        s = s & "left outer join tblcategories k on a.optioncategory=k.category" & vbCrLf
        s = s & "where a.divisionid=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
        s = s & "and isnull(a.optionid,'')<>''" & vbCrLf
        s = s & "order by 1,2,3,4" & vbCrLf
        s = s & Chr(0) & vbCrLf
        
        s = s & "Community Global Options" & Chr(1) & vbCrLf
        s = s & "select distinct k.Description Category,m.OptionID [Option],m.Assembly,m.Description,m.community SourceCommunity,c.description Community" & vbCrLf
        s = s & "from tbldbassemblymaster m" & vbCrLf
        s = s & "join tbllocality c on c.area=m.community" & vbCrLf
        s = s & "left join divisioncommunities d on m.community=d.community and m.divisionid=d.divisionid" & vbCrLf
        s = s & "left outer join tblcategories k on m.category=k.category" & vbCrLf
        s = s & "where m.divisionid=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
        s = s & "and (d.community is not null or isnull(m.community,'')='')" & vbCrLf
        s = s & "and isnull(m.inactive,0)=0" & vbCrLf
        s = s & "and m.assemblytype=3" & vbCrLf
        s = s & "and isnull(m.isbaseassembly,0)=0" & vbCrLf
        s = s & "order by 1,2,3,4" & vbCrLf
    
    End If
        
    Dim AssemblyPL As New FPickList
    If Not AssemblyPL.Choose(HFApp.Databases(dbHomefront), "Assembly", s, , , , , "areaID,modelid,optionid,assemblytype,SourceCommunity", True, iView) Then Exit Sub
    iView = AssemblyPL.SelectedView


    If IsIn(iView, 2, 4, 6) Then
        'community specific assemblies
        
        'remember where you started adding
        NewRow = gData.Rows
    
        'add new ones
        For a = 1 To AssemblyPL.SelectedItems
            Call AddAssemblyToGrid(AssemblyPL.SelectedItem("SourceCommunity", a), _
                                   "", _
                                   AssemblyPL.SelectedItem("Assembly", a), _
                                   AssemblyPL.SelectedItem("Model", a), _
                                   AssemblyPL.SelectedItem("Option", a), _
                                   AssemblyPL.SelectedItem("series", a))
        Next
    Else
        'global assemblies

        Dim CommunityPL As New FPickList
        
        s = ""
        s = s & "select '** All Commmnities **' Description,'' Community,'' Phase,'' Inactive,-1 SortOrder" & vbCrLf
        s = s & "union" & vbCrLf
        s = s & "select c.Description,c.area Community,'' Phase,case when c.Inactive=1 then 'Inactive' else '' end ,c.inactive SortOrder" & vbCrLf
        s = s & "from tbllocality c " & vbCrLf
        s = s & " left outer join DivisionCommunities d on d.Community = c.Area" & vbCrLf
        s = s & "where d.DivisionID=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
        s = s & "union" & vbCrLf
        s = s & "select c.description + isnull(' ' + p.description,''),c.area Community, isnull(p.communityphase,'') Phase,case when c.Inactive=1 then 'Inactive' else '' end ,c.inactive SortOrder" & vbCrLf
        s = s & "from tbllocality c left outer join communityphase p on(c.area=p.community)" & vbCrLf
        s = s & " left outer join DivisionCommunities d on d.Community = c.Area" & vbCrLf
        s = s & "where d.DivisionID=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
        s = s & "order by 5,1" & vbCrLf
        If Not CommunityPL.Choose(HFApp.Databases(dbHomefront), "Selling Communities", s, , , , , "SortOrder", True, , , "Select selling communities") Then Exit Sub
    
    
        'remember where you started adding
        NewRow = gData.Rows
    
        'add new ones
        For a = 1 To AssemblyPL.SelectedItems
            For c = 1 To CommunityPL.SelectedItems
                Call AddAssemblyToGrid(CommunityPL.SelectedItem("Community", c), _
                                       CommunityPL.SelectedItem("Phase", c), _
                                       AssemblyPL.SelectedItem("Assembly", a), _
                                       AssemblyPL.SelectedItem("Model", a), _
                                       AssemblyPL.SelectedItem("Option", a), _
                                       AssemblyPL.SelectedItem("series", a))
            
            Next
        Next
    End If
    
    
    Call SaveData(False)
    For i = NewRow To gData.Rows - 1
        gData.Row = i
        Call RefreshCosts(, gData.Row)
        Call CalcData(, "Cost", gData.Row)
    Next
        

End Sub


Private Sub AddAssemblyToGrid(community As String, CommunityPhase As String, Assembly As String, Model As String, OptionID As String, series As String)
    Dim i As Long
    Dim s As String
    Dim rs As Recordset
    Dim r As Long
    Dim AssemblyType As AssemblyTypes

Dim X As Single
Dim Y As String
X = Timer()
    s = ""
    s = s & "SELECT m.Community SourceCommunity,m.Model,m.Elevation,m.Description,dm.Description ModelDescription,m.Notes,m.Comments,m.AssemblyType,m.Qty,m.Color,m.Location,m.IncludedOption" & vbCrLf
    s = s & "      ,m.Category,c.Description CategoryDesc" & vbCrLf
    s = s & "      ,l.area community,l.Description CommunityDesc,m.Assembly,m.OptionID,m.AssemblyUOM" & vbCrLf
    s = s & "      ,m.Series,m.Style,m.JCExtra,m.FloorArea,m.Bedrooms,m.Bathrooms,m.ConstCutOff" & vbCrLf
    s = s & "      ,CASE isnull(p.assembly,'') when '' THEN m.IncentiveRetail ELSE p.IncentiveRetail END IncentiveRetail" & vbCrLf
    s = s & "      ,CASE isnull(p.assembly,'') when '' THEN m.IncludeTax ELSE p.IncludeTax END IncludeTax" & vbCrLf
    s = s & "      ,CASE isnull(p.assembly,'') when '' THEN m.Markup ELSE p.Markup END Markup" & vbCrLf
    s = s & "      ,CASE isnull(p.assembly,'') when '' THEN c.COMarkup ELSE p.COMarkup END COMarkup" & vbCrLf
    s = s & "      ,CASE isnull(p.assembly,'') when '' THEN m.Margin ELSE p.Margin END Margin" & vbCrLf
    s = s & "      ,CASE isnull(p.assembly,'') when '' THEN m.Roundto ELSE p.Roundto END Roundto" & vbCrLf
    s = s & "      ,CASE isnull(p.assembly,'') when '' THEN m.Pretax ELSE p.Pretax END Pretax" & vbCrLf
    s = s & "      ,CASE isnull(p.assembly,'') when '' THEN m.Tax ELSE p.Tax END Tax" & vbCrLf
    s = s & "      ,CASE isnull(p.assembly,'') when '' THEN m.COPretax ELSE p.COPretax END COPretax" & vbCrLf
    s = s & "      ,CASE isnull(p.assembly,'') when '' THEN m.COTax ELSE p.COTax END COTax      " & vbCrLf
    s = s & "      ,m.IncludedInSpec" & vbCrLf
    For i = 2 To 10
        s = s & "      ,m.Pretax" & i & vbCrLf
        s = s & "      ,m.Tax" & i & vbCrLf
        s = s & "      ,m.IncludedInSpec" & i & vbCrLf
    Next
    s = s & "      ,m.ColorListID,m.StyleListID,m.FinishListID,m.OtherListID,m.StyleValue,m.FinishValue,m.OtherValue,m.graphicpath,m.specdocument,m.maxwidth,m.maxlength,m.constcutoff" & vbCrLf
    s = s & ",m.DesignCenterSalesOnly,m.SelectByRoom,m.DisplayTotalOnly" & vbCrLf
    s = s & "  FROM tblDBAssemblyMaster m " & vbCrLf
    s = s & "       LEFT OUTER JOIN tblDBAssemblyPrices p ON(p.Community=" & DbQuote(Str, community) & " AND p.CommunityPhase=" & DbQuote(Str, CommunityPhase) & " AND m.Assembly=p.Assembly AND m.Model=p.Model AND m.OptionID=p.OptionID)" & vbCrLf
    s = s & "       LEFT OUTER JOIN tblCategories c ON m.Category=c.Category" & vbCrLf
    s = s & "       left join distinctmodelsbydivision dm on m.model=dm.model and m.divisionid=dm.divisionid" & vbCrLf
    s = s & "       LEFT OUTER JOIN tblLocality l ON l.area=" & DbQuote(Str, community) & vbCrLf
    s = s & " WHERE (isnull(m.Community,'')='' or isnull(m.Community,'')=" & DbQuote(Str, community) & ")" & vbCrLf
    s = s & "   AND isnull(m.Assembly,'')=" & DbQuote(Str, Assembly) & vbCrLf
    s = s & "   AND isnull(m.Model,'')=" & DbQuote(Str, Model) & vbCrLf
    s = s & "   AND isnull(m.OptionID,'')=" & DbQuote(Str, OptionID) & vbCrLf
    s = s & "   AND m.DivisionID = " & HFApp.DivisionID & vbCrLf
    If OptionID <> "" Then
        s = s & "union all" & vbCrLf
        s = s & "SELECT '' SourceCommunity,'' Model,'' Elevation,m.Description,'' ModelDescription,m.Notes,m.Notes Comments,3 AssemblyType,1 Qty,m.Color,m.Location,0 IncludedOption" & vbCrLf
        s = s & "      ,m.OptionCategory Category,c.Description CategoryDesc" & vbCrLf
        s = s & "      ,l.area community,l.Description CommunityDesc,'' Assembly,m.OptionID,m.OrderUOM AssemblyUOM" & vbCrLf
        s = s & "      ,'' Series,'' Style,'' JCExtra,0 FloorArea,0 Bedrooms,0 Bathrooms,0 ConstCutOff" & vbCrLf
        s = s & "      ,p.IncentiveRetail" & vbCrLf
        s = s & "      ,p.IncludeTax" & vbCrLf
        s = s & "      ,p.Markup" & vbCrLf
        s = s & "      ,p.COMarkup" & vbCrLf
        s = s & "      ,p.Margin" & vbCrLf
        s = s & "      ,p.Roundto" & vbCrLf
        s = s & "      ,CASE isnull(p.assembly,'') when '' THEN m.RetailPretax ELSE p.Pretax END Pretax" & vbCrLf
        s = s & "      ,p.Tax" & vbCrLf
        s = s & "      ,CASE isnull(p.assembly,'') when '' THEN m.RetailPretax ELSE p.Pretax END COPretax" & vbCrLf
        s = s & "      ,p.COTax" & vbCrLf
        s = s & "      ,0 IncludedInSpec" & vbCrLf
        For i = 2 To 10
            s = s & "      ,0 Pretax" & i & vbCrLf
            s = s & "      ,0 Tax" & i & vbCrLf
            s = s & "      ,0 IncludedInSpec" & i & vbCrLf
        Next
        s = s & "      ,0,0,0,0,'','','',null,null,null,null,null" & vbCrLf
        s = s & ",0 DesignCenterSalesOnly,0 SelectByRoom,0 DisplayTotalOnly" & vbCrLf
        s = s & "FROM tblphaseitem m" & vbCrLf
        s = s & "  LEFT OUTER JOIN tblDBAssemblyPrices p ON(p.Community=" & DbQuote(Str, community) & " AND p.CommunityPhase=" & DbQuote(Str, CommunityPhase) & " AND p.Assembly='' AND p.Model='' AND m.OptionID=p.OptionID)" & vbCrLf
        s = s & "  LEFT OUTER JOIN tblCategories c ON m.OptionCategory=c.Category" & vbCrLf
        s = s & "  LEFT OUTER JOIN tblLocality l ON l.area=" & DbQuote(Str, community) & vbCrLf
        s = s & "WHERE m.DivisionID = " & HFApp.DivisionID & " and isnull(m.OptionID,'')=" & DbQuote(Str, OptionID) & vbCrLf
    End If
    s = s & "order by 1 desc"
    Set rs = HFApp.SqlExec(s)
Y = Y & "query " & format(Timer() - X, "0.000") & vbCrLf: X = Timer()

    If rs.EOF Then Exit Sub
    
    mDirty = True
    With gData
    
        AssemblyType = Val("" & rs("AssemblyType"))
        r = .Rows
        .AddItem ""
        .Cell(flexcpChecked, r, .ColIndex("ReadyToPublish")) = flexChecked
        
        .TextMatrix(r, .ColIndex("Community")) = "" & rs("Community")
        .TextMatrix(r, .ColIndex("CommunityDesc")) = "" & rs("CommunityDesc")
        .TextMatrix(r, .ColIndex("CommunityPhase")) = CommunityPhase
        
        .TextMatrix(r, .ColIndex("SourceCommunity")) = "" & rs("SourceCommunity")
        
        .TextMatrix(r, .ColIndex("Model")) = "" & rs("Model")
        If AssemblyType = atoption Then
            .TextMatrix(r, .ColIndex("ModelDescription")) = "" & rs("ModelDescription")
        End If
        .TextMatrix(r, .ColIndex("Elevation")) = "" & rs("Elevation")
        .TextMatrix(r, .ColIndex("Description")) = "" & rs("Description")
        .TextMatrix(r, .ColIndex("Notes")) = "" & rs("Notes")
        .TextMatrix(r, .ColIndex("Comments")) = "" & rs("Comments")
        .TextMatrix(r, .ColIndex("AssemblyType")) = AssemblyType
        .TextMatrix(r, .ColIndex("Category")) = "" & rs("Category")
        .TextMatrix(r, .ColIndex("CategoryDesc")) = "" & rs("CategoryDesc")
        .TextMatrix(r, .ColIndex("JCExtra")) = "" & rs("JCExtra")
        .TextMatrix(r, .ColIndex("Series")) = series
        .TextMatrix(r, .ColIndex("FloorArea")) = Val("" & rs("FloorArea"))
        .TextMatrix(r, .ColIndex("Bedrooms")) = "" & rs("Bedrooms")
        .TextMatrix(r, .ColIndex("Bathrooms")) = "" & rs("Bathrooms")
        .TextMatrix(r, .ColIndex("Style")) = "" & rs("Style")
        
        .TextMatrix(r, .ColIndex("DCSalesOnly")) = "" & rs("DesignCenterSalesOnly")
        .TextMatrix(r, .ColIndex("SelectByRoom")) = "" & rs("SelectByRoom")
        .TextMatrix(r, .ColIndex("DisplayTotalOnly")) = "" & rs("DisplayTotalOnly")
        
        .TextMatrix(r, .ColIndex("Color")) = "" & rs("Color")
        .TextMatrix(r, .ColIndex("Location")) = "" & rs("Location")
        .TextMatrix(r, .ColIndex("Qty")) = Val("" & rs("Qty"))
        If .TextMatrix(r, .ColIndex("Qty")) = "0" Then .TextMatrix(r, .ColIndex("Qty")) = "1"
        .TextMatrix(r, .ColIndex("IncludedOption")) = "" & rs("IncludedOption")
        
        .TextMatrix(r, .ColIndex("ConstCutOff")) = Val("" & rs("ConstCutOff"))
        .TextMatrix(r, .ColIndex("SpecDocument")) = "" & rs("SpecDocument")
        .TextMatrix(r, .ColIndex("GraphicPath")) = "" & rs("GraphicPath")
        .TextMatrix(r, .ColIndex("MaxWidth")) = Val("" & rs("MaxWidth"))
        .TextMatrix(r, .ColIndex("MaxLength")) = Val("" & rs("MaxLength"))
        .Cell(flexcpChecked, r, .ColIndex("Inactive")) = flexUnchecked
        
        
        .TextMatrix(r, .ColIndex("Markup")) = Val("" & rs("Markup"))
        .TextMatrix(r, .ColIndex("Margin")) = Val("" & rs("Margin"))
        .TextMatrix(r, .ColIndex("Roundto")) = Val("" & rs("Roundto"))
        .TextMatrix(r, .ColIndex("IncludeTax")) = "" & rs("IncludeTax")
        
        If Val("" & rs("IncludeTax")) = 1 Then
            .TextMatrix(r, .ColIndex("Pretax")) = Val("" & rs("Pretax"))
        Else
            .TextMatrix(r, .ColIndex("Pretax")) = RoundToPrice(Val("" & rs("Pretax")), Val("" & rs("Roundto")))
        End If
        .TextMatrix(r, .ColIndex("Assembly")) = "" & rs("Assembly")
        .TextMatrix(r, .ColIndex("AssemblyUOM")) = "" & rs("AssemblyUOM")
        .TextMatrix(r, .ColIndex("OptionID")) = "" & rs("OptionID")

        
        .TextMatrix(r, .ColIndex("InSpec")) = "" & rs("IncludedInSpec")
        For i = 2 To 10
            .TextMatrix(r, .ColIndex("Pretax" & i)) = Val("" & rs("Pretax" & i))
            .TextMatrix(r, .ColIndex("Tax" & i)) = Val("" & rs("Tax" & i))
            .TextMatrix(r, .ColIndex("InSpec" & i)) = Val("" & rs("IncludedInSpec" & i))
        Next
        
        .TextMatrix(r, .ColIndex("Tax")) = CalcTax(.ValueMatrix(r, .ColIndex("AssemblyType")), .ValueMatrix(r, .ColIndex("Pretax")))
        .TextMatrix(r, .ColIndex("Total")) = Round(.ValueMatrix(r, .ColIndex("Pretax")) + .ValueMatrix(r, .ColIndex("Tax")), 2)
        
        'get assembly costs
        .TextMatrix(r, .ColIndex("Cost")) = 0
        .TextMatrix(r, .ColIndex("LandCost")) = IIf(AssemblyType = atModel, GetAssemblyCost(r, HFApp.Options(LandPhase), HFApp.Options(LandItem)), 0)
        .TextMatrix(r, .ColIndex("IncentiveCost")) = IIf(AssemblyType = atModel, GetAssemblyCost(r, HFApp.Options(IncentivePhase), HFApp.Options(IncentiveItem)), 0)
        .TextMatrix(r, .ColIndex("ConstructionCost")) = .ValueMatrix(r, .ColIndex("Cost")) - .ValueMatrix(r, .ColIndex("LandCost")) - .ValueMatrix(r, .ColIndex("IncentiveCost"))
        .TextMatrix(r, .ColIndex("IncentiveRetail")) = "" & rs("IncentiveRetail")
        
        'replace incentive cost with the cost percent on this worksheet
        .TextMatrix(r, .ColIndex("IncentiveCost")) = Round(.ValueMatrix(r, .ColIndex("IncentiveRetail")) * Val(txtIncentiveCostPercent) / 100, 2)
        .TextMatrix(r, .ColIndex("Cost")) = .ValueMatrix(r, .ColIndex("ConstructionCost")) + .ValueMatrix(r, .ColIndex("LandCost")) + .ValueMatrix(r, .ColIndex("IncentiveCost"))

        .TextMatrix(r, .ColIndex("ColorListID")) = "" & rs("ColorListID")
        .TextMatrix(r, .ColIndex("StyleListID")) = "" & rs("StyleListID")
        .TextMatrix(r, .ColIndex("FinishListID")) = "" & rs("FinishListID")
        .TextMatrix(r, .ColIndex("OtherListID")) = "" & rs("OtherListID")
        .TextMatrix(r, .ColIndex("StyleValue")) = "" & rs("StyleValue")
        .TextMatrix(r, .ColIndex("FinishValue")) = "" & rs("FinishValue")
        .TextMatrix(r, .ColIndex("OtherValue")) = "" & rs("OtherValue")
        If AssemblyType = atDesignCenter Then
        'calc markup & margins for dc options
            For i = 2 To 10
            
                If .ValueMatrix(r, .ColIndex("Cost")) = 0 Then
                    .TextMatrix(r, .ColIndex("Markup" & i)) = "0"
                Else
                    .TextMatrix(r, .ColIndex("Markup" & i)) = (.ValueMatrix(r, .ColIndex("Pretax" & i)) - .ValueMatrix(r, .ColIndex("Cost"))) / .ValueMatrix(r, .ColIndex("Cost")) * 100
                End If
                
                
                If .ValueMatrix(r, .ColIndex("Pretax" & i)) = 0 Then
                    .TextMatrix(r, .ColIndex("Margin" & i)) = "0"
                Else
                    .TextMatrix(r, .ColIndex("Margin" & i)) = ((.ValueMatrix(r, .ColIndex("Pretax" & i)) - (.ValueMatrix(r, .ColIndex("Cost")))) / .ValueMatrix(r, .ColIndex("Pretax" & i))) * 100
                End If
            Next
        End If
        
        
        .TextMatrix(r, .ColIndex("COMarkup")) = Val("" & rs("COMarkup"))
        .TextMatrix(r, .ColIndex("COMargin")) = Val("" & rs("Margin"))
        .TextMatrix(r, .ColIndex("COPretax")) = Val("" & rs("COPretax"))
        .TextMatrix(r, .ColIndex("COTax")) = Round(.ValueMatrix(r, .ColIndex("COPretax")) * HFApp.Options(GST_Rate) / 100 * HFApp.Options(PST_Rate) / 100, 2)
        .TextMatrix(r, .ColIndex("COTotal")) = Round(.ValueMatrix(r, .ColIndex("COPretax")) + .ValueMatrix(r, .ColIndex("COTax")), 2)
        .Col = .ColIndex("Roundto")
        
        .Row = r
'Y = Y & "add row " & format(Timer() - X, "0.000") & vbCrLf: X = Timer()

        
'MsgBox y
        
    End With

End Sub

