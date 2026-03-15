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



